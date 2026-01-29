using Microsoft.VisualStudio.Shell;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace SSMS_EnvTabs
{
    internal sealed class ColorByRegexConfigWriter
    {
        private const string BeginMarker = "// SSMS EnvTabs: BEGIN generated";
        private const string EndMarker = "// SSMS EnvTabs: END generated";
        private const int ResolveRetryMax = 6;
        private const int ResolveRetryDelayMs = 500;
        private const double CreationSkewSeconds = 2.0;
        private const double CreationMaxWindowSeconds = 30.0;

        private string resolvedConfigPath;
        private DateTime? firstDocSeenUtc;
        private bool resolveRetryScheduled;
        private int resolveRetryCount;
        private List<RdtEventManager.OpenDocumentInfo> lastDocsSnapshot;
        private IReadOnlyList<TabRuleMatcher.CompiledRule> lastRulesSnapshot;
        private IReadOnlyList<TabRuleMatcher.CompiledManualRule> lastManualRulesSnapshot;

        public void UpdateFromSnapshot(IEnumerable<RdtEventManager.OpenDocumentInfo> docs, IReadOnlyList<TabRuleMatcher.CompiledRule> rules, IReadOnlyList<TabRuleMatcher.CompiledManualRule> manualRules)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (docs == null)
            {
                return;
            }

            lastDocsSnapshot = docs as List<RdtEventManager.OpenDocumentInfo> ?? docs.ToList();
            lastRulesSnapshot = rules;
            lastManualRulesSnapshot = manualRules;

            if (!firstDocSeenUtc.HasValue && lastDocsSnapshot.Count > 0)
            {
                firstDocSeenUtc = DateTime.UtcNow;
            }

            var safeRules = rules ?? Array.Empty<TabRuleMatcher.CompiledRule>();
            bool hasManualRules = manualRules != null && manualRules.Count > 0;
            EnvTabsLog.Info($"ColorByRegexConfigWriter.cs::UpdateFromSnapshot - Rules={safeRules.Count}, ManualRules={(manualRules?.Count ?? 0)}, Docs={lastDocsSnapshot.Count}");
            if (safeRules.Count == 0 && !hasManualRules)
            {
                EnvTabsLog.Info("ColorByRegexConfigWriter.cs::UpdateFromSnapshot - No rules to write. Skipping.");
                return;
            }

            string configPath = ResolveConfigPath(lastDocsSnapshot.Select(d => d?.Moniker));
            if (string.IsNullOrWhiteSpace(configPath))
            {
                EnvTabsLog.Info("ColorByRegexConfigWriter.cs::UpdateFromSnapshot - Config path not resolved. Skipping.");
                ScheduleResolveRetryIfNeeded();
                return;
            }
            resolveRetryCount = 0;
            resolveRetryScheduled = false;
            EnvTabsLog.Info($"ColorByRegexConfigWriter.cs::UpdateFromSnapshot - ConfigPath='{configPath}'");

            var groupToPaths = safeRules.Count > 0 ? BuildGroupPathMap(lastDocsSnapshot, safeRules) : null;
            EnvTabsLog.Info($"ColorByRegexConfigWriter.cs::UpdateFromSnapshot - GroupToPaths={(groupToPaths?.Count ?? 0)}");
            string newContent = BuildConfigContent(configPath, groupToPaths, safeRules, manualRules);
            WriteIfChanged(configPath, newContent);
        }

        private struct RegexEntry
        {
            public string Pattern;
            public int Priority;
        }

        private string BuildConfigContent(string existingPath, Dictionary<string, SortedSet<string>> groupToPaths, IReadOnlyList<TabRuleMatcher.CompiledRule> rules, IReadOnlyList<TabRuleMatcher.CompiledManualRule> manualRules)
        {
            var entries = new List<RegexEntry>();

            // Manual rules
            if (manualRules != null)
            {
                foreach (var m in manualRules)
                {
                    string line = m.OriginalPattern;
                    if (m.ColorIndex.HasValue)
                    {
                        try
                        {
                            line = TabGroupColorSolver.SolveForColor(m.OriginalPattern, m.ColorIndex.Value);
                        }
                        catch
                        {
                            line = m.OriginalPattern;
                        }
                    }
                    entries.Add(new RegexEntry { Pattern = line, Priority = m.Priority });
                }
            }

            // Generated rules
            if (groupToPaths != null)
            {
                foreach (var rule in rules)
                {
                    if (groupToPaths.TryGetValue(rule.GroupName, out var paths) && paths.Count > 0)
                    {
                        string line = BuildResolvedRegexLine(paths, rule);
                        entries.Add(new RegexEntry { Pattern = line, Priority = rule.Priority });
                    }
                }
            }

            var sortedEntries = entries.OrderBy(e => e.Priority).ToList();

            // Preserve content outside markers.
            string existing = File.Exists(existingPath) ? File.ReadAllText(existingPath) : string.Empty;
            var generatedLines = new List<string>();

            generatedLines.Add(BeginMarker);
            foreach(var e in sortedEntries)
            {
                generatedLines.Add(e.Pattern);
            }
            generatedLines.Add(EndMarker);

            return ReplaceOrAppendBlock(existing, generatedLines);
        }        

        internal string BuildGeneratedBlockPreview(IEnumerable<RdtEventManager.OpenDocumentInfo> docs, IReadOnlyList<TabRuleMatcher.CompiledRule> rules)
        {
            if (docs == null || rules == null || rules.Count == 0)
            {
                return string.Empty;
            }

            var groupToPaths = BuildGroupPathMap(docs, rules);
            var lines = BuildGeneratedBlock(groupToPaths, rules);
            return string.Join("\n", lines);
        }

        private Dictionary<string, SortedSet<string>> BuildGroupPathMap(IEnumerable<RdtEventManager.OpenDocumentInfo> docs, IReadOnlyList<TabRuleMatcher.CompiledRule> rules)
        {
            var map = new Dictionary<string, SortedSet<string>>(StringComparer.OrdinalIgnoreCase);
            foreach (var rule in rules)
            {
                if (!map.ContainsKey(rule.GroupName))
                {
                    map[rule.GroupName] = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
                }
            }

            foreach (var doc in docs)
            {
                if (doc == null)
                {
                    continue;
                }

                string group = TabRuleMatcher.MatchGroup(rules, doc.Server, doc.Database);
                if (string.IsNullOrWhiteSpace(group))
                {
                    continue;
                }

                if (string.IsNullOrWhiteSpace(doc.Moniker) || !Path.IsPathRooted(doc.Moniker))
                {
                    continue;
                }

                if (!map.TryGetValue(group, out var set))
                {
                    set = new SortedSet<string>(StringComparer.OrdinalIgnoreCase);
                    map[group] = set;
                }

                // Use only the filename so colors follow file moves.
                try 
                {
                    string fileName = Path.GetFileName(doc.Moniker);
                    if (!string.IsNullOrEmpty(fileName))
                    {
                        set.Add(fileName);
                    }
                }
                catch
                {
                    // If path is invalid, skip
                }
            }

            return map;
        }

        private static List<string> BuildGeneratedBlock(Dictionary<string, SortedSet<string>> groupToPaths, IReadOnlyList<TabRuleMatcher.CompiledRule> rules)
        {
            var lines = new List<string> { BeginMarker };

            foreach (var rule in rules.OrderBy(r => r.Priority).ThenBy(r => r.GroupName, StringComparer.OrdinalIgnoreCase))
            {
                groupToPaths.TryGetValue(rule.GroupName, out var paths);
                lines.Add(BuildResolvedRegexLine(paths, rule));
            }

            lines.Add(EndMarker);
            return lines;
        }

        private static string BuildResolvedRegexLine(SortedSet<string> paths, TabRuleMatcher.CompiledRule rule)
        {
            if (paths == null || paths.Count == 0)
            {
                return "(?!)";
            }

            var pathList = paths.ToList();
            bool allSql = pathList.All(path => path.EndsWith(".sql", StringComparison.OrdinalIgnoreCase));
            IEnumerable<string> escaped;
            string suffix = string.Empty;
            if (allSql)
            {
                escaped = pathList.Select(path => Regex.Escape(Path.GetFileNameWithoutExtension(path)));
                suffix = "\\.sql";
            }
            else
            {
                escaped = pathList.Select(Regex.Escape);
            }

            // Match path ending with the filename.
            string baseRegex = $"(?:^|[\\\\/])(?:{string.Join("|", escaped)}){suffix}$";

            // If the rule specifies a color index, solve for a salt.
            string salt = null;
            if (rule.ColorIndex >= 0)
            {
                string currentFull = baseRegex;
                int currentHash = TabGroupColorSolver.GetSsmsStableHashCode(currentFull);
                int currentColor = TabGroupColorSolver.GetColorIndex(currentHash);

                if (currentColor != rule.ColorIndex)
                {
                    string newSalt = TabGroupColorSolver.Solve(baseRegex, rule.ColorIndex);
                    if (newSalt != null)
                    {
                        salt = newSalt;
                    }
                }
            }

            string finalRegex = baseRegex;
            if (!string.IsNullOrWhiteSpace(salt)) 
            {
                finalRegex += $"(?#salt:{salt})";
            }
            
            return finalRegex;
        }

        private static string ReplaceOrAppendBlock(string existing, List<string> generatedLines)
        {
            var lines = SplitLines(existing);
            int begin = lines.FindIndex(line => string.Equals(line?.TrimEnd(), BeginMarker, StringComparison.Ordinal));
            int end = lines.FindIndex(line => string.Equals(line?.TrimEnd(), EndMarker, StringComparison.Ordinal));

            var result = new List<string>();
            if (begin >= 0 && end >= begin)
            {
                result.AddRange(lines.Take(begin));
                result.AddRange(generatedLines);
                result.AddRange(lines.Skip(end + 1));
            }
            else
            {
                result.AddRange(lines);
                if (result.Count > 0 && !string.IsNullOrWhiteSpace(result[result.Count - 1]))
                {
                    result.Add(string.Empty);
                }

                result.AddRange(generatedLines);
            }

            return string.Join(Environment.NewLine, result);
        }

        private static List<string> SplitLines(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return new List<string>();
            }

            return text
                .Replace("\r\n", "\n")
                .Replace("\r", "\n")
                .Split('\n')
                .ToList();
        }

        private static void WriteIfChanged(string path, string content)
        {
            string existing = File.Exists(path) ? File.ReadAllText(path) : string.Empty;
            if (string.Equals(NormalizeNewlines(existing), NormalizeNewlines(content ?? string.Empty), StringComparison.Ordinal))
            {
                EnvTabsLog.Info($"ColorByRegexConfigWriter.cs::WriteIfChanged - No changes detected for '{path}'.");
                return;
            }

            string dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            string tmp = path + ".tmp";
            EnvTabsLog.Info($"ColorByRegexConfigWriter.cs::WriteIfChanged - Writing file '{path}'.");
            File.WriteAllText(tmp, content ?? string.Empty, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

            if (File.Exists(path))
            {
                File.Replace(tmp, path, null, ignoreMetadataErrors: true);
            }
            else
            {
                File.Move(tmp, path);
            }
        }

        private static string NormalizeNewlines(string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return string.Empty;
            }

            return text.Replace("\r\n", "\n").Replace("\r", "\n");
        }

        private void ScheduleResolveRetryIfNeeded()
        {
            if (resolveRetryScheduled)
            {
                return;
            }

            if (resolveRetryCount >= ResolveRetryMax)
            {
                return;
            }

            resolveRetryScheduled = true;
            int attempt = ++resolveRetryCount;
            EnvTabsLog.Info($"ColorByRegexConfigWriter.cs::UpdateFromSnapshot - Resolve retry scheduled ({attempt}/{ResolveRetryMax}) in {ResolveRetryDelayMs}ms.");

            _ = ThreadHelper.JoinableTaskFactory.RunAsync(async () =>
            {
                try
                {
                    await System.Threading.Tasks.Task.Delay(ResolveRetryDelayMs).ConfigureAwait(true);
                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();

                    resolveRetryScheduled = false;

                    if (lastDocsSnapshot != null)
                    {
                        UpdateFromSnapshot(lastDocsSnapshot, lastRulesSnapshot, lastManualRulesSnapshot);
                    }
                }
                catch
                {
                    resolveRetryScheduled = false;
                }
            });
        }

        private string ResolveConfigPath(IEnumerable<string> monikers)
        {
            if (!string.IsNullOrWhiteSpace(resolvedConfigPath) && File.Exists(resolvedConfigPath))
            {
                return resolvedConfigPath;
            }

            // Try from open documents first.
            foreach (var moniker in monikers ?? Enumerable.Empty<string>())
            {
                string candidate = TryGetConfigPathFromMoniker(moniker);
                if (!string.IsNullOrWhiteSpace(candidate))
                {
                    resolvedConfigPath = candidate;
                    return candidate;
                }
            }

            // Scan Temp until the file appears.
            string fallback = TryScanTempForConfig(firstDocSeenUtc);
            if (!string.IsNullOrWhiteSpace(fallback))
            {
                resolvedConfigPath = fallback;
                return fallback;
            }

            return null;
        }

        private static string TryGetConfigPathFromMoniker(string moniker)
        {
            if (string.IsNullOrWhiteSpace(moniker) || !Path.IsPathRooted(moniker))
            {
                return null;
            }

            try
            {
                string dir = Path.GetDirectoryName(moniker);
                if (string.IsNullOrWhiteSpace(dir))
                {
                    return null;
                }

                var current = new DirectoryInfo(dir);

                string guidRoot = TryGetTempGuidRoot(current);
                if (!string.IsNullOrWhiteSpace(guidRoot))
                {
                    EnvTabsLog.Info($"ColorByRegexConfigWriter.cs::TryGetConfigPathFromMoniker - Temp GUID root='{guidRoot}'");
                    string candidate = Path.Combine(guidRoot, "ColorByRegexConfig.txt");
                    if (File.Exists(candidate))
                    {
                        return candidate;
                    }
                }

                var walker = current;
                for (int i = 0; i < 4 && walker != null; i++)
                {
                    string candidate = Path.Combine(walker.FullName, "ColorByRegexConfig.txt");
                    if (File.Exists(candidate))
                    {
                        return candidate;
                    }

                    walker = walker.Parent;
                }

                return null;
            }
            catch
            {
                return null;
            }
        }

        private static string TryScanTempForConfig(DateTime? referenceUtc)
        {
            try
            {
                if (!referenceUtc.HasValue)
                {
                    EnvTabsLog.Info("ColorByRegexConfigWriter.cs::TryScanTempForConfig - No firstDocSeenUtc; skipping temp scan.");
                    return null;
                }

                string temp = Path.GetTempPath();
                if (string.IsNullOrWhiteSpace(temp) || !Directory.Exists(temp))
                {
                    return null;
                }

                // SSMS places the file in a GUID-named Temp subdirectory.

                var dirs = Directory.GetDirectories(temp)
                    .Where(d => IsGuidDirectoryName(Path.GetFileName(d)))
                    .Select(d => new { Dir = d, CreatedUtc = Directory.GetCreationTimeUtc(d) })
                    .OrderBy(d => d.CreatedUtc)
                    .ToList();

                foreach (var entry in dirs)
                {
                    try
                    {
                        string candidate = Path.Combine(entry.Dir, "ColorByRegexConfig.txt");
                        bool fileExists = File.Exists(candidate);

                        DateTime dirCreateUtc = entry.CreatedUtc;
                        double deltaSeconds = (dirCreateUtc - referenceUtc.Value).TotalSeconds;

                        // Skip folders outside the creation window.
                        if (deltaSeconds < -CreationSkewSeconds)
                        {
                            continue;
                        }

                        if (deltaSeconds > CreationMaxWindowSeconds)
                        {
                            continue;
                        }

                        EnvTabsLog.Info($"ColorByRegexConfigWriter.cs::TryScanTempForConfig - FirstDocSeenUtc={referenceUtc:O}, RegexFolderCreateUtc={dirCreateUtc:O}, DeltaSeconds={deltaSeconds:0.###}, FileExists={fileExists}, Folder='{entry.Dir}'");
                        if (fileExists)
                        {
                            return candidate;
                        }
                    }
                    catch
                    {
                        // Ignore access errors
                    }
                }
                EnvTabsLog.Info($"ColorByRegexConfigWriter.cs::TryScanTempForConfig - No candidate found. FirstDocSeenUtc={referenceUtc:O}");
                return null;
            }
            catch
            {
                return null;
            }
        }

        private static string TryGetTempGuidRoot(DirectoryInfo start)
        {
            if (start == null)
            {
                return null;
            }

            string tempRoot = Path.GetTempPath();
            var current = start;
            while (current != null)
            {
                if (IsGuidDirectoryName(current.Name))
                {
                    if (string.IsNullOrWhiteSpace(tempRoot))
                    {
                        return current.FullName;
                    }

                    if (current.FullName.StartsWith(tempRoot, StringComparison.OrdinalIgnoreCase))
                    {
                        return current.FullName;
                    }
                }

                current = current.Parent;
            }

            return null;
        }

        private static bool IsGuidDirectoryName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return false;
            }

            return Guid.TryParse(name, out _);
        }
    }
}
