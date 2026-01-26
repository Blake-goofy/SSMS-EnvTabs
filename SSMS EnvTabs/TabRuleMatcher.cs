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

            public CompiledRule(string groupName, int priority, int colorIndex, string server, string database)
            {
                GroupName = groupName;
                Priority = priority;
                ColorIndex = colorIndex;
                Server = server;
                Database = database;
            }
        }

        public static List<CompiledRule> CompileRules(TabGroupConfig config)
        {
            var rules = new List<CompiledRule>();
            if (config?.Groups == null)
            {
                return rules;
            }

            foreach (var rule in config.Groups)
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

                rules.Add(new CompiledRule(rule.GroupName.Trim(), rule.Priority, rule.ColorIndex, server, database));
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
                    && !MatchesLike(rule.Server, server))
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(rule.Database)
                    && !MatchesLike(rule.Database, database))
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

        private static bool MatchesLike(string pattern, string value)
        {
            if (string.IsNullOrWhiteSpace(pattern)) return true;
            if (string.IsNullOrWhiteSpace(value)) return false;

            if (!pattern.Contains("%"))
            {
                return string.Equals(pattern, value, StringComparison.OrdinalIgnoreCase);
            }

            string regex = "^" + Regex.Escape(pattern).Replace("%", ".*") + "$";
            return Regex.IsMatch(value, regex, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        }
    }
}
