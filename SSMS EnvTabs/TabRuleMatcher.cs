using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace SSMS_EnvTabs
{
    internal static class TabRuleMatcher
    {
        internal sealed class CompiledRule
        {
            public string GroupName { get; }
            public int Priority { get; }
            public int ColorIndex { get; }
            public string Server { get; }
            public string Database { get; }
            public Regex ServerRegex { get; }
            public Regex DatabaseRegex { get; }

            public CompiledRule(string groupName, int priority, int colorIndex, string server, string database, Regex serverRegex, Regex databaseRegex)
            {
                GroupName = groupName;
                Priority = priority;
                ColorIndex = colorIndex;
                Server = server;
                Database = database;
                ServerRegex = serverRegex;
                DatabaseRegex = databaseRegex;
            }
        }

        internal sealed class CompiledManualRule
        {
            public string GroupName { get; }
            public int Priority { get; }
            public int? ColorIndex { get; }
            public Regex FileRegex { get; }
            public string OriginalPattern { get; }

            public CompiledManualRule(string groupName, string pattern, int priority, int? colorIndex)
            {
                GroupName = groupName;
                OriginalPattern = pattern;
                Priority = priority;
                ColorIndex = colorIndex;
                try
                {
                    FileRegex = new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
                }
                catch
                {
                    // Invalid regex
                }
            }
        }

        public static List<CompiledManualRule> CompileManualRules(TabGroupConfig config)
        {
            var list = new List<CompiledManualRule>();
            if (config?.ManualRegexLines == null) return list;

            foreach (var m in config.ManualRegexLines)
            {
                if (string.IsNullOrWhiteSpace(m.Pattern)) continue;
                list.Add(new CompiledManualRule(m.GroupName, m.Pattern, m.Priority, m.ColorIndex));
            }

            return list.OrderBy(x => x.Priority).ToList();
        }

        public static CompiledManualRule MatchManual(IReadOnlyList<CompiledManualRule> manuals, string moniker)
        {
            if (manuals == null || string.IsNullOrWhiteSpace(moniker)) return null;

            foreach (var m in manuals)
            {
                if (m.FileRegex != null && m.FileRegex.IsMatch(moniker))
                {
                    return m;
                }
            }
            return null;
        }

        public static List<CompiledRule> CompileRules(TabGroupConfig config)
        {
            var rules = new List<CompiledRule>();
            if (config?.ConnectionGroups == null)
            {
                return rules;
            }

            foreach (var rule in config.ConnectionGroups)
            {
                if (string.IsNullOrWhiteSpace(rule?.GroupName))
                {
                    continue;
                }

                string server = string.IsNullOrWhiteSpace(rule.Server) ? null : rule.Server.Trim();
                string database = string.IsNullOrWhiteSpace(rule.Database) ? null : rule.Database.Trim();

                if (string.IsNullOrWhiteSpace(server) && string.IsNullOrWhiteSpace(database))
                {
                    continue;
                }

                Regex serverRegex = CreateLikeRegexOrNull(server);
                Regex databaseRegex = CreateLikeRegexOrNull(database);

                rules.Add(new CompiledRule(rule.GroupName.Trim(), rule.Priority, rule.ColorIndex, server, database, serverRegex, databaseRegex));
            }
            
            return rules
                .OrderBy(r => r.Priority)
                .ThenBy(r => r.GroupName, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        public static string MatchGroup(IReadOnlyList<CompiledRule> rules, string server, string database)
        {
            if (rules == null || rules.Count == 0)
            {
                return null;
            }

            foreach (var rule in rules)
            {
                if (!string.IsNullOrWhiteSpace(rule.Server)
                    && !Matches(rule.Server, rule.ServerRegex, server))
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(rule.Database)
                    && !Matches(rule.Database, rule.DatabaseRegex, database))
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(rule.Server) || !string.IsNullOrWhiteSpace(rule.Database))
                {
                    return rule.GroupName;
                }
            }

            return null;
        }

        private static Regex CreateLikeRegexOrNull(string pattern)
        {
            if (string.IsNullOrWhiteSpace(pattern) || !pattern.Contains("%"))
            {
                return null;
            }

            string regex = "^" + Regex.Escape(pattern).Replace("%", ".*") + "$";
            return new Regex(regex, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);
        }

        private static bool Matches(string pattern, Regex compiledRegex, string value)
        {
            if (string.IsNullOrWhiteSpace(pattern)) return true;
            if (string.IsNullOrWhiteSpace(value)) return false;

            if (compiledRegex != null)
            {
                return compiledRegex.IsMatch(value);
            }

            return string.Equals(pattern, value, StringComparison.OrdinalIgnoreCase);
        }
    }
}
