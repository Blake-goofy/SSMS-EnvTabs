using System;
using System.Text.RegularExpressions;

namespace SSMS_EnvTabs
{
    internal static class TabCaptionFormatter
    {
        internal static string GetPureName(string rawCaption, string ssmsSuffix, bool enableRemoveDotSql)
        {
            string caption = StripDirtyIndicators(rawCaption);
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

        internal static string StripSqlExtension(string name, bool enabled)
        {
            if (!enabled || string.IsNullOrEmpty(name))
            {
                return name;
            }

            return Regex.Replace(name, @"\.sql$", string.Empty, RegexOptions.IgnoreCase);
        }

        internal static string BuildSavedStyleCaption(string savedStyle, string filenameToken, string groupName, string server, string serverAlias, string database)
        {
            return (savedStyle ?? string.Empty)
                .Replace("[filename]", filenameToken ?? string.Empty)
                .Replace("[groupName]", groupName ?? string.Empty)
                .Replace("[server]", server ?? string.Empty)
                .Replace("[serverAlias]", serverAlias ?? server ?? string.Empty)
                .Replace("[db]", database ?? string.Empty);
        }
    }
}
