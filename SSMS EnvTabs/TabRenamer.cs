using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;

namespace SSMS_EnvTabs
{
    internal static class TabRenamer
    {
        private static readonly Dictionary<uint, (string GroupName, int Index)> CookieToAssignment =
            new Dictionary<uint, (string GroupName, int Index)>();

        public static void ForgetCookie(uint cookie)
        {
            if (cookie == 0) return;
            CookieToAssignment.Remove(cookie);
        }

        public static int ApplyRenamesOrThrow(IEnumerable<(uint cookie, IVsWindowFrame frame, string server, string database, string frameCaption)> tabs, IReadOnlyList<TabRuleMatcher.CompiledRule> rules)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (rules == null) throw new ArgumentNullException(nameof(rules));
            if (rules.Count == 0) return 0;

            var nextIndexByGroup = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            foreach (var assignment in CookieToAssignment.Values)
            {
                if (!nextIndexByGroup.TryGetValue(assignment.GroupName, out int next))
                {
                    next = 1;
                }

                nextIndexByGroup[assignment.GroupName] = Math.Max(next, assignment.Index + 1);
            }

            int renamed = 0;

            foreach (var (cookie, frame, server, database, frameCaption) in tabs)
            {
                if (frame == null) continue;

                string group = TabRuleMatcher.MatchGroup(rules, server, database);
                if (string.IsNullOrWhiteSpace(group))
                {
                    continue;
                }

                if (!CookieToAssignment.TryGetValue(cookie, out var assignment) || !string.Equals(assignment.GroupName, group, StringComparison.OrdinalIgnoreCase))
                {
                    if (!nextIndexByGroup.TryGetValue(group, out int next))
                    {
                        next = 1;
                    }

                    assignment = (group, next);
                    CookieToAssignment[cookie] = assignment;
                    nextIndexByGroup[group] = next + 1;
                }

                string newCaption = $"{assignment.GroupName}{assignment.Index}";

                if (!string.IsNullOrEmpty(frameCaption))
                {
                    string expectedPrefix = newCaption + " -";
                    if (frameCaption.StartsWith(expectedPrefix, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                }

                int hr = TrySetTabCaption(frame, newCaption, out string propertyNameUsed);
                if (ErrorHandler.Succeeded(hr))
                {
                    renamed++;
                    EnvTabsLog.Info($"Renamed ({propertyNameUsed}): cookie={cookie}, '{frameCaption}' -> '{newCaption}'");
                }
                else
                {
                    EnvTabsLog.Info($"Rename failed (hr=0x{hr:X8}): cookie={cookie}, caption='{frameCaption}', target='{newCaption}'");
                }
            }

            return renamed;
        }

        private static int TrySetTabCaption(IVsWindowFrame frame, string newCaption, out string propertyNameUsed)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            propertyNameUsed = null;

            int hr = frame.SetProperty((int)__VSFPROPID.VSFPROPID_OwnerCaption, newCaption);
            if (ErrorHandler.Succeeded(hr))
            {
                propertyNameUsed = nameof(__VSFPROPID.VSFPROPID_OwnerCaption);
                return hr;
            }

            int hr2 = frame.SetProperty((int)__VSFPROPID.VSFPROPID_Caption, newCaption);
            if (ErrorHandler.Succeeded(hr2))
            {
                propertyNameUsed = nameof(__VSFPROPID.VSFPROPID_Caption);
                return hr2;
            }

            int hr3 = frame.SetProperty((int)__VSFPROPID.VSFPROPID_EditorCaption, newCaption);
            if (ErrorHandler.Succeeded(hr3))
            {
                propertyNameUsed = nameof(__VSFPROPID.VSFPROPID_EditorCaption);
                return hr3;
            }

            return hr;
        }
    }
}
