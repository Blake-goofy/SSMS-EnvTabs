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
        private static readonly HashSet<string> suppressedConnections = new HashSet<string>();

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
            if (suppressedConnections.Contains(connectionKey)) return;

            // Apply immediately (per user requirement "regardless it should write the json")
            AddRuleAndSave(config, server, database, useDb);

            // Prompt to edit
            if (config.Settings.EnableConfigurePrompt)
            {
                string msg = $"SSMS EnvTabs Auto-Configuration\n\nA new tab grouping rule has been added for:\nServer: {server}\nDatabase: {(useDb ? database : "(any)")}\n\nDo you want to edit the configuration file now?";
                DialogResult result = MessageBox.Show(msg, "EnvTabs Configuration Updated", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                if (result == DialogResult.Yes)
                {
                    OpenConfigInEditor();
                }
            }

            suppressedConnections.Add(connectionKey);
        }

        private static void AddRuleAndSave(TabGroupConfig config, string server, string database, bool useDb)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Calculate new values
            int nextColor = FindNextColorValues(config);
            string groupName = useDb && !string.IsNullOrWhiteSpace(database) ? $"{server}.{database}" : server;
            
            // Renumber priorities: existing rules move to 20, 30, 40...
            // Sort existing by priority
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
                ColorIndex = nextColor
            };

            config.Groups.Add(newRule);
            
            // Save
            SaveConfig(config);
        }

        private static void OpenConfigInEditor()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string path = TabGroupConfigLoader.GetUserConfigPath();
            VsShellUtilities.OpenDocument(ServiceProvider.GlobalProvider, path);
        }

        private static int FindNextColorValues(TabGroupConfig config)
        {
            // Simple heuristic: round robin 0-15 based on usage count, or just first unused.
            // Requirement: "first color that doesn't appear in the color index (skipping -1, just 0-16)"
            var used = new HashSet<int>(config.Groups.Select(x => x.ColorIndex));
            
            for (int i = 0; i <= 15; i++)
            {
                if (!used.Contains(i)) return i;
            }

            // If all used, pick random or 0
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
                using (var writer = JsonReaderWriterFactory.CreateJsonWriter(stream, System.Text.Encoding.UTF8, true, true, "  "))
                {
                    serializer.WriteObject(writer, config);
                    writer.Flush();
                    File.WriteAllBytes(path, stream.ToArray());
                }
            }
            catch (System.Exception ex)
            {
                EnvTabsLog.Info($"AutoConfig save failed: {ex.Message}");
            }
        }
    }
}
