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
        private bool fallbackScanAttempted;

        public void UpdateFromSnapshot(IEnumerable<RdtEventManager.OpenDocumentInfo> docs, IReadOnlyList<TabRuleMatcher.CompiledRule> rules)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (docs == null || rules == null || rules.Count == 0)
            {
                return;
            }

            string configPath = ResolveConfigPath(docs.Select(d => d?.Moniker));
            if (string.IsNullOrWhiteSpace(configPath))
            {
                return;
            }

            var groupToPaths = BuildGroupPathMap(docs, rules);
            string newContent = BuildConfigContent(configPath, groupToPaths, rules);
            WriteIfChanged(configPath, newContent);
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

                set.Add(doc.Moniker);
            }

            return map;
        }

        private string BuildConfigContent(string configPath, Dictionary<string, SortedSet<string>> groupToPaths, IReadOnlyList<TabRuleMatcher.CompiledRule> rules)
        {
            string existing = File.Exists(configPath) ? File.ReadAllText(configPath) : string.Empty;
            var generatedLines = BuildGeneratedBlock(groupToPaths, rules);
            return ReplaceOrAppendBlock(existing, generatedLines);
        }

        private static List<string> BuildGeneratedBlock(Dictionary<string, SortedSet<string>> groupToPaths, IReadOnlyList<TabRuleMatcher.CompiledRule> rules)
        {
            var lines = new List<string> { BeginMarker };

            foreach (var rule in rules.OrderBy(r => r.Priority).ThenBy(r => r.GroupName, StringComparer.OrdinalIgnoreCase))
            {
                groupToPaths.TryGetValue(rule.GroupName, out var paths);
                lines.Add(BuildRegexLine(paths));
            }

            lines.Add(EndMarker);
            return lines;
        }

        private static string BuildRegexLine(SortedSet<string> paths)
        {
            if (paths == null || paths.Count == 0)
            {
                return "(?!)";
            }

            var escaped = paths.Select(Regex.Escape);
            return $"^(?:{string.Join("|", escaped)})$";
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
            if (string.Equals(existing, content ?? string.Empty, StringComparison.Ordinal))
            {
                return;
            }

            string dir = Path.GetDirectoryName(path);
            if (!string.IsNullOrWhiteSpace(dir) && !Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }

            string tmp = path + ".tmp";
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

        private string ResolveConfigPath(IEnumerable<string> monikers)
        {
            if (!string.IsNullOrWhiteSpace(resolvedConfigPath) && File.Exists(resolvedConfigPath))
            {
                return resolvedConfigPath;
            }

            foreach (var moniker in monikers ?? Enumerable.Empty<string>())
            {
                string candidate = TryGetConfigPathFromMoniker(moniker);
                if (!string.IsNullOrWhiteSpace(candidate))
                {
                    resolvedConfigPath = candidate;
                    return candidate;
                }
            }

            if (!fallbackScanAttempted)
            {
                fallbackScanAttempted = true;
                string fallback = TryScanTempForConfig();
                if (!string.IsNullOrWhiteSpace(fallback))
                {
                    resolvedConfigPath = fallback;
                    return fallback;
                }
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

        private static string TryScanTempForConfig()
        {
            try
            {
                string temp = Path.GetTempPath();
                if (string.IsNullOrWhiteSpace(temp) || !Directory.Exists(temp))
                {
                    return null;
                }

                string newest = null;
                DateTime newestWriteUtc = DateTime.MinValue;

                foreach (var file in Directory.EnumerateFiles(temp, "ColorByRegexConfig.txt", SearchOption.AllDirectories))
                {
                    try
                    {
                        DateTime writeUtc = File.GetLastWriteTimeUtc(file);
                        if (writeUtc > newestWriteUtc)
                        {
                            newestWriteUtc = writeUtc;
                            newest = file;
                        }
                    }
                    catch
                    {
                        // Ignore this path.
                    }
                }

                return newest;
            }
            catch
            {
                return null;
            }
        }
    }
}
