using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Windows.Forms;
using Microsoft.VisualStudio.Shell;

namespace SSMS_EnvTabs
{
    internal static class AutoConfigurationService
    {
        private static readonly HashSet<string> suppressedConnections = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        internal static event Action<DialogClosedInfo> DialogClosed;

        internal sealed class DialogClosedInfo
        {
            public DialogResult Result { get; set; }
            public string Server { get; set; }
            public string Database { get; set; }
            public bool ChangesApplied { get; set; }
        }

        public static void ClearSuppressed()
        {
            suppressedConnections.Clear();
        }

        public static void ProposeNewRule(TabGroupConfig config, string server, string database)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (string.IsNullOrWhiteSpace(server)) return;
            if (config == null || config.Settings == null) return;

            string mode = config.Settings.AutoConfigure;
            if (string.IsNullOrWhiteSpace(mode)) return;

            // Normalize mode
            bool useDb = mode.IndexOf("db", System.StringComparison.OrdinalIgnoreCase) >= 0;

            string connectionKey = useDb ? $"{server}::{database}" : server;
            
            // Check if already configured (in strict memory) to avoid duplicate prompt if re-triggered quickly
            // Though RdtEventManager should handle this via "hasMatchingRule" check.

            if (suppressedConnections.Contains(connectionKey)) 
            {
                EnvTabsLog.Info($"AutoConfig: Suppressed key {connectionKey}");
                return;
            }

            // Prepare suggested values
            int nextColor = FindNextColorValues(config);
            string suggestedName = useDb && !string.IsNullOrWhiteSpace(database) ? $"{server}.{database}" : server;

            EnvTabsLog.Info($"AutoConfig: Proposing new rule for {connectionKey}. Color={nextColor}, Name={suggestedName}, Prompt={config.Settings.EnableConfigurePrompt}");
            
            // Step 1: ALWAYS Create and Save the rule immediately (Requirement: "rule should be written... right away")
            var newRule = AddRuleAndSave(config, server, database, useDb, suggestedName, nextColor);
            
            // Mark as suppressed immediately so we don't re-enter if RDT events fire again while dialog is open or after
            suppressedConnections.Add(connectionKey);

            // Silent mode enabled? Or Prompt?
            if (config.Settings.EnableConfigurePrompt)
            {
                try
                {
                    EnvTabsLog.Info("AutoConfig: Opening prompt dialog...");
                    // Step 2: Prompt to edit the JUST CREATED rule
                    // We use newRule.ColorIndex because AddRuleAndSave might have adjusted it if we add more logic, 
                    // but currently it just uses what we passed.
                    using (var dlg = new NewRuleDialog(server, database, suggestedName, nextColor))
                    {
                        // To avoid "Color goes away" on Cancel:
                        // We must ensure that we DO NOT trigger a save/write if canceled.
                        // And we should ensure the initial save was enough.
                        
                        var result = dlg.ShowDialog();
                        EnvTabsLog.Info($"AutoConfig: Dialog result = {result}");
                        bool changesApplied = false;
                        if (result == DialogResult.OK || result == DialogResult.Yes) 
                        {
                            string updatedName = string.IsNullOrWhiteSpace(dlg.RuleName) ? newRule.GroupName : dlg.RuleName;
                            int updatedColor = dlg.SelectedColorIndex;

                            // User wants to change it
                            if (updatedName != newRule.GroupName || updatedColor != newRule.ColorIndex)
                            {
                                newRule.GroupName = updatedName;
                                newRule.ColorIndex = updatedColor;
                                SaveConfig(config);
                                changesApplied = true;
                            }
                             
                            if (result == DialogResult.Yes)
                            {
                                OpenConfigInEditor();
                            }
                        }

                        DialogClosed?.Invoke(new DialogClosedInfo
                        {
                            Result = result,
                            Server = server,
                            Database = database,
                            ChangesApplied = changesApplied
                        });
                    }
                }
                catch (System.Exception ex)
                {
                    EnvTabsLog.Error($"AutoConfig: Error showing prompt: {ex}");
                }
            }
        }

        private static TabGroupRule AddRuleAndSave(TabGroupConfig config, string server, string database, bool useDb, string groupName, int colorIndex)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Check if we should remove the default example rule
            var exampleRule = config.Groups.FirstOrDefault(g => g.GroupName == "Example: Exact Match");
            if (exampleRule != null)
            {
                config.Groups.Remove(exampleRule);
            }
            
            // Renumber priorities: existing rules move to 20, 30, 40...
            var sorted = config.Groups.OrderBy(x => x.Priority).ToList();
            int currentBase = 20;
            foreach (var rule in sorted)
            {
                rule.Priority = currentBase;
                currentBase += 10;
            }

            // Create new rule at 10
            var newRule = new TabGroupRule
            {
                GroupName = groupName,
                Server = server,
                Database = useDb ? (database ?? "%") : "%",
                Priority = 10,
                ColorIndex = colorIndex
            };

            config.Groups.Add(newRule);
            
            // Save
            SaveConfig(config);
            return newRule;
        }

        private static void OpenConfigInEditor()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string path = TabGroupConfigLoader.GetUserConfigPath();
            VsShellUtilities.OpenDocument(ServiceProvider.GlobalProvider, path);
        }

        private static int FindNextColorValues(TabGroupConfig config)
        {
            // Requirement 1: "First available" - find a color not currently used.
            // Requirement 2: Strict "Fill the gap" logic without jumping.
            
            var used = new HashSet<int>(config.Groups.Select(x => x.ColorIndex));
            
            // Find first hole in 0-15
            for (int i = 0; i <= 15; i++)
            {
                if (!used.Contains(i)) return i;
            }

            // If all used, return 0
            return 0;
        }

        private static void SaveConfig(TabGroupConfig config)
        {
            try
            {
                string path = TabGroupConfigLoader.GetUserConfigPath();
                var serializer = new DataContractJsonSerializer(typeof(TabGroupConfig), new DataContractJsonSerializerSettings
                {
                    UseSimpleDictionaryFormat = true
                });

                using (var stream = new MemoryStream())
                {
                    using (var writer = JsonReaderWriterFactory.CreateJsonWriter(stream, System.Text.Encoding.UTF8, true, true, "  "))
                    {
                        serializer.WriteObject(writer, config);
                        writer.Flush();
                    }
                    
                    // JsonReaderWriterFactory/DataContractJsonSerializer escapes forward slashes as \/.
                    // We decode to string, replace them for readability, and write back.
                    string json = System.Text.Encoding.UTF8.GetString(stream.ToArray());
                    json = json.Replace("\\/", "/");
                    
                    File.WriteAllText(path, json, new System.Text.UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
                }
            }
            catch (System.Exception ex)
            {
                EnvTabsLog.Info($"AutoConfig save failed: {ex.Message}");
            }
        }
    }
}
