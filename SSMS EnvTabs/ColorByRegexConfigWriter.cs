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

        private string resolvedConfigPath;
        private DateTime? firstDocSeenUtc;

        public void UpdateFromSnapshot(IEnumerable<RdtEventManager.OpenDocumentInfo> docs, IReadOnlyList<TabRuleMatcher.CompiledRule> rules, IReadOnlyList<TabRuleMatcher.CompiledManualRule> manualRules)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (docs == null)
            {
                return;
            }

            if (!firstDocSeenUtc.HasValue && docs.Any())
            {
                firstDocSeenUtc = DateTime.UtcNow;
            }

            var safeRules = rules ?? Array.Empty<TabRuleMatcher.CompiledRule>();
            bool hasManualRules = manualRules != null && manualRules.Count > 0;
            EnvTabsLog.Info($"ColorByRegexConfigWriter.cs::UpdateFromSnapshot - Rules={safeRules.Count}, ManualRules={(manualRules?.Count ?? 0)}, Docs={docs.Count()}");
            if (safeRules.Count == 0 && !hasManualRules)
            {
                EnvTabsLog.Info("ColorByRegexConfigWriter.cs::UpdateFromSnapshot - No rules to write. Skipping.");
                return;
            }

            string configPath = ResolveConfigPath(docs.Select(d => d?.Moniker));
            if (string.IsNullOrWhiteSpace(configPath))
            {
                EnvTabsLog.Info("ColorByRegexConfigWriter.cs::UpdateFromSnapshot - Config path not resolved. Skipping.");
                return;
            }
            EnvTabsLog.Info($"ColorByRegexConfigWriter.cs::UpdateFromSnapshot - ConfigPath='{configPath}'");

            var groupToPaths = safeRules.Count > 0 ? BuildGroupPathMap(docs, safeRules) : null;
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

            // 1. Process Manual Rules
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
                            // fallback
                            line = m.OriginalPattern;
                        }
                    }
                    entries.Add(new RegexEntry { Pattern = line, Priority = m.Priority });
                }
            }

            // 2. Process Connection Rules (Generated)
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

            // 3. Sort
            var sortedEntries = entries.OrderBy(e => e.Priority).ToList();

            // 4. Build Content
            // We read the existing file to preserve anything OUTSIDE our markers (if any).
            
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

                // Use only the filename for the regex to avoid path dependency issues
                // and to support files moving between folders while keeping color.
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

        private string BuildConfigContent(string configPath, Dictionary<string, SortedSet<string>> groupToPaths, IReadOnlyList<TabRuleMatcher.CompiledRule> rules)
        {
            var entries = new List<RegexEntry>();
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
             var generatedLines = new List<string>();
            generatedLines.Add(BeginMarker);
            foreach(var e in sortedEntries)
            {
                generatedLines.Add(e.Pattern);
            }
            generatedLines.Add(EndMarker);

            string existing = File.Exists(configPath) ? File.ReadAllText(configPath) : string.Empty;
            return ReplaceOrAppendBlock(existing, generatedLines);
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

            // Match filepath ending with one of the filenames.
            string baseRegex = $"(?:^|[\\\\/])(?:{string.Join("|", escaped)}){suffix}$";

            // If the rule specifically asks for a specific ColorIndex (0-15), solve for salt
            // Note: 0 is Lavender
            string salt = null;
            if (rule.ColorIndex >= 0)
            {
                // Attempt to see if current salt (if any) already produces the right color
                string currentFull = baseRegex;
                int currentHash = TabGroupColorSolver.GetSsmsStableHashCode(currentFull);
                int currentColor = Math.Abs(currentHash) % 16;

                if (currentColor != rule.ColorIndex)
                {
                    // Need to find a new salt
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

        private string ResolveConfigPath(IEnumerable<string> monikers)
        {
            if (!string.IsNullOrWhiteSpace(resolvedConfigPath) && File.Exists(resolvedConfigPath))
            {
                return resolvedConfigPath;
            }

            // 1. Try to deduce from open documents (most reliable if file is near them)
            foreach (var moniker in monikers ?? Enumerable.Empty<string>())
            {
                string candidate = TryGetConfigPathFromMoniker(moniker);
                if (!string.IsNullOrWhiteSpace(candidate))
                {
                    resolvedConfigPath = candidate;
                    return candidate;
                }
            }

            // 2. Scan Temp folder. 
            // We do this every time if not found yet, because the file might be created later by SSMS.
            // But we must ensure the scan is fast and doesn't crash.
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
                if (current.Exists && Regex.IsMatch(current.Name, "^v\\d+$", RegexOptions.IgnoreCase))
                {
                    var tempGuidDir = current.Parent;
                    if (tempGuidDir != null)
                    {
                        string candidate = Path.Combine(tempGuidDir.FullName, "ColorByRegexConfig.txt");
                        if (File.Exists(candidate))
                        {
                            return candidate;
                        }
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
                string temp = Path.GetTempPath();
                if (string.IsNullOrWhiteSpace(temp) || !Directory.Exists(temp))
                {
                    return null;
                }

                // SSMS typically places the file in a subdirectory of Temp.
                // Searching recursively (AllDirectories) is dangerous due to permissions.
                // We iterate top-level directories and look inside them.
                
                string bestCandidate = null;
                double bestScore = double.MinValue;

                var dirs = Directory.GetDirectories(temp);
                foreach (var dir in dirs)
                {
                    try
                    {
                        string candidate = Path.Combine(dir, "ColorByRegexConfig.txt");
                        if (!File.Exists(candidate))
                        {
                            continue;
                        }

                        bool hasVersionFolder = Directory.GetDirectories(dir)
                            .Select(d => Path.GetFileName(d))
                            .Any(name => Regex.IsMatch(name ?? string.Empty, "^v\\d+$", RegexOptions.IgnoreCase));

                        DateTime dirCreateUtc = Directory.GetCreationTimeUtc(dir);
                        double secondsFromReference = referenceUtc.HasValue
                            ? Math.Abs((dirCreateUtc - referenceUtc.Value).TotalSeconds)
                            : double.PositiveInfinity;

                        bool withinTimeWindow = !referenceUtc.HasValue || secondsFromReference <= 60;

                        // Prefer dirs with version folder and creation time near the first doc open.
                        // If no reference time, rely on write time and version folder presence.
                        double writeUtcTicks = File.GetLastWriteTimeUtc(candidate).Ticks;
                        double score = 0;
                        score += hasVersionFolder ? 1000000 : 0;
                        score += withinTimeWindow ? 100000 : 0;
                        score -= referenceUtc.HasValue ? secondsFromReference : 0;
                        score += writeUtcTicks / 10000000.0; // scaled

                        if (score > bestScore)
                        {
                            bestScore = score;
                            bestCandidate = candidate;
                        }
                    }
                    catch
                    {
                        // Ignore access errors
                    }
                }

                return bestCandidate;
            }
            catch
            {
                return null;
            }
        }
    }
}
