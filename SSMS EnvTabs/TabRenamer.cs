using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;

namespace SSMS_EnvTabs
{
    internal class TabRenameContext
    {
        public uint Cookie { get; set; }
        public IVsWindowFrame Frame { get; set; }
        public string Server { get; set; }
        public string Database { get; set; }
        public string FrameCaption { get; set; }
        public string Moniker { get; set; }
    }

    internal static class TabRenamer
    {
        private static readonly Dictionary<uint, (string GroupName, int Index)> CookieToAssignment =
            new Dictionary<uint, (string GroupName, int Index)>();

        public static void ForgetCookie(uint cookie)
        {
            if (cookie == 0) return;
            CookieToAssignment.Remove(cookie);
        }

        public static int ApplyRenamesOrThrow(IEnumerable<TabRenameContext> tabs, IReadOnlyList<TabRuleMatcher.CompiledRule> rules, IReadOnlyList<TabRuleMatcher.CompiledManualRule> manualRules, string renameStyle = null, string savedFileRenameStyle = null)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (rules == null) throw new ArgumentNullException(nameof(rules));
            
            // Default style if null
            if (string.IsNullOrWhiteSpace(renameStyle))
            {
                renameStyle = "[groupName][#]";
            }

            // Default saved file style if null (fallback to legacy behavior)
            string effectiveSavedStyle = savedFileRenameStyle;
            if (string.IsNullOrWhiteSpace(effectiveSavedStyle))
            {
                effectiveSavedStyle = "[filename] [groupName]";
            }

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

            foreach (var tab in tabs)
            {
                if (tab.Frame == null) continue;

                var manualMatch = TabRuleMatcher.MatchManual(manualRules, tab.Moniker);
                string group = manualMatch?.GroupName ?? TabRuleMatcher.MatchGroup(rules, tab.Server, tab.Database);

                if (string.IsNullOrWhiteSpace(group))
                {
                    continue;
                }

                if (!CookieToAssignment.TryGetValue(tab.Cookie, out var assignment) || !string.Equals(assignment.GroupName, group, StringComparison.OrdinalIgnoreCase))
                {
                    if (!nextIndexByGroup.TryGetValue(group, out int next))
                    {
                        next = 1;
                    }

                    assignment = (group, next);
                    CookieToAssignment[tab.Cookie] = assignment;
                    nextIndexByGroup[group] = next + 1;
                }

                string newCaption;
                if (manualMatch != null)
                {
                    // Manual match implies overwriting caption with GroupName (User Request)
                    newCaption = assignment.GroupName;
                }
                else if (RdtEventManager.IsTempFile(tab.Moniker))
                {
                    // Case 1: Temp File (New Query) -> Use User Configured Style
                    // Replace [groupName] and [#]
                    // We support "#" as a placeholder too if user uses it, but prefer [#] to be explicit.
                    // Implementation: Simple replace.
                    newCaption = renameStyle
                        .Replace("[groupName]", assignment.GroupName)
                        .Replace("[#]", assignment.Index.ToString())
                        .Replace("#", assignment.Index.ToString()); // Support raw # as requested by user
                }
                else
                {
                    // Case 2: Saved File -> Use Saved File Configured Style
                    // Replace [filename], [groupName], [server], [db]
                    string fileName = System.IO.Path.GetFileNameWithoutExtension(tab.Moniker);
                    string server = tab.Server ?? "";
                    string database = tab.Database ?? "";

                    newCaption = effectiveSavedStyle
                        .Replace("[filename]", fileName)
                        .Replace("[groupName]", assignment.GroupName)
                        .Replace("[server]", server)
                        .Replace("[db]", database);
                }
                
                // Only skip if currently matches exact target
                if (!string.IsNullOrEmpty(tab.FrameCaption) && string.Equals(tab.FrameCaption, newCaption, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                int hr = TrySetTabCaption(tab.Frame, newCaption, out string propertyNameUsed);
                if (ErrorHandler.Succeeded(hr))
                {
                    renamed++;
                    EnvTabsLog.Info($"Renamed ({propertyNameUsed}): cookie={tab.Cookie}, '{tab.FrameCaption}' -> '{newCaption}'");
                }
                else
                {
                    EnvTabsLog.Info($"Rename failed (hr=0x{hr:X8}): cookie={tab.Cookie}, caption='{tab.FrameCaption}', target='{newCaption}'");
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
