using System;
using System.Text.RegularExpressions;

namespace SSMS_EnvTabs
{
    internal static class TabCaptionFormatter
    {
        private const string ExecutingToken = "Executing...";

        internal static string GetPureName(string rawCaption, string ssmsSuffix, bool enableRemoveDotSql)
        {
            string caption = StripExecutionPrefix(StripDirtyIndicators(rawCaption));
            if (string.IsNullOrEmpty(caption)) return string.Empty;

            if (!string.IsNullOrEmpty(ssmsSuffix))
            {
                while (caption.EndsWith(ssmsSuffix, StringComparison.OrdinalIgnoreCase))
                {
                    caption = caption.Substring(0, caption.Length - ssmsSuffix.Length).TrimEnd();
                }
            }

            // Dirty markers can be left just before the SSMS suffix (e.g. "name.sql* ⬤ - SERVER").
            // Strip again after suffix removal so extension/token cleanup can proceed correctly.
            caption = StripDirtyIndicators(caption);

            caption = StripSqlExtension(caption, enableRemoveDotSql);
            return caption;
        }

        internal static bool CaptionsEquivalent(string actualCaption, string expectedCaption, string ssmsSuffix, bool enableRemoveDotSql)
        {
            string actual = NormalizeForComparison(actualCaption, ssmsSuffix, enableRemoveDotSql);
            string expected = NormalizeForComparison(expectedCaption, ssmsSuffix, enableRemoveDotSql);

            return !string.IsNullOrWhiteSpace(actual)
                && !string.IsNullOrWhiteSpace(expected)
                && string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase);
        }

        internal static string NormalizeForComparison(string rawCaption, string ssmsSuffix, bool enableRemoveDotSql)
        {
            string caption = GetPureName(rawCaption, ssmsSuffix, enableRemoveDotSql);
            return caption?.Trim() ?? string.Empty;
        }

        internal static string StripDirtyIndicators(string name)
        {
            if (string.IsNullOrEmpty(name)) return name ?? string.Empty;

            string s = name.Trim();
            bool changed;
            do
            {
                changed = false;

                while (s.EndsWith("*"))
                {
                    s = s.TrimEnd('*').TrimEnd();
                    changed = true;
                }

                while (s.EndsWith("\u2B24"))
                {
                    s = s.Substring(0, s.Length - 1).TrimEnd();
                    changed = true;
                }
            } while (changed);

            return s;
        }

        internal static string StripExecutionPrefix(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return name ?? string.Empty;
            }

            string s = name.Trim();
            bool changed;
            do
            {
                changed = false;

                if (s.StartsWith(ExecutingToken, StringComparison.OrdinalIgnoreCase))
                {
                    s = s.Substring(ExecutingToken.Length).TrimStart();
                    changed = true;
                }

                string withoutSuffix = Regex.Replace(s, @"\s+-\s*Executing\.\.\.$", string.Empty, RegexOptions.IgnoreCase);
                if (!string.Equals(withoutSuffix, s, StringComparison.Ordinal))
                {
                    s = withoutSuffix.TrimEnd();
                    changed = true;
                }
            } while (changed);

            return s;
        }

        internal static bool HasExecutionMarker(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return false;
            }

            string caption = StripDirtyIndicators(name);
            if (string.IsNullOrWhiteSpace(caption))
            {
                return false;
            }

            return caption.StartsWith(ExecutingToken, StringComparison.OrdinalIgnoreCase)
                || Regex.IsMatch(caption, @"\s+-\s*Executing\.\.\.$", RegexOptions.IgnoreCase);
        }

        internal static string StripSqlExtension(string name, bool enabled)
        {
            if (!enabled || string.IsNullOrEmpty(name))
            {
                return name;
            }

            return Regex.Replace(name, @"\.sql$", string.Empty, RegexOptions.IgnoreCase);
        }

        internal static string GetFilenameToken(string moniker, bool enableRemoveDotSql)
        {
            if (string.IsNullOrWhiteSpace(moniker))
            {
                return string.Empty;
            }

            return enableRemoveDotSql
                ? System.IO.Path.GetFileNameWithoutExtension(moniker)
                : System.IO.Path.GetFileName(moniker);
        }

        internal static string BuildSavedStyleCaption(string savedStyle, string filenameToken, string groupName, string server, string serverAlias, string database)
        {
            return BuildStyleCaption(savedStyle, filenameToken, groupName, server, serverAlias, database, null);
        }

        internal static string BuildStyleCaption(string style, string filenameToken, string groupName, string server, string serverAlias, string database, int? index)
        {
            return (style ?? string.Empty)
                .Replace("[filename]", filenameToken ?? string.Empty)
                .Replace("[groupName]", groupName ?? string.Empty)
                .Replace("[server]", server ?? string.Empty)
                .Replace("[serverAlias]", serverAlias ?? server ?? string.Empty)
                .Replace("[db]", database ?? string.Empty)
                .Replace("[#]", index?.ToString() ?? string.Empty);
        }
    }
}
