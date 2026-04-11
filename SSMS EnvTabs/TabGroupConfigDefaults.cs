using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Runtime.Serialization;

namespace SSMS_EnvTabs
{
    internal static class TabGroupConfigDefaults
    {
        private sealed class SettingMember
        {
            public string JsonName { get; set; }

            public PropertyInfo Property { get; set; }
        }

        private static readonly SettingMember[] SettingsMembers = typeof(TabGroupSettings)
            .GetProperties(BindingFlags.Public | BindingFlags.Instance)
            .Select(property => new SettingMember
            {
                JsonName = property.GetCustomAttribute<DataMemberAttribute>()?.Name,
                Property = property
            })
            .Where(member => !string.IsNullOrWhiteSpace(member.JsonName) && member.Property.CanRead && member.Property.CanWrite)
            .ToArray();

        internal static IReadOnlyList<string> ApplyMissingSettingDefaults(TabGroupConfig config, string existingJsonText, TabGroupConfig defaultConfig)
        {
            List<string> migratedSettings = new List<string>();
            if (config == null)
            {
                return migratedSettings;
            }

            TabGroupSettings defaultSettings = defaultConfig?.Settings ?? CreateFallbackDefaultConfig().Settings;
            if (defaultSettings == null)
            {
                return migratedSettings;
            }

            if (!TryExtractObjectJson(existingJsonText, "settings", out string settingsJson))
            {
                config.Settings = CloneSettings(defaultSettings);
                migratedSettings.AddRange(SettingsMembers.Select(member => member.JsonName));
                return migratedSettings;
            }

            if (config.Settings == null)
            {
                config.Settings = new TabGroupSettings();
            }

            foreach (SettingMember setting in SettingsMembers)
            {
                if (JsonContainsProperty(settingsJson, setting.JsonName))
                {
                    continue;
                }

                object defaultValue = setting.Property.GetValue(defaultSettings, null);
                setting.Property.SetValue(config.Settings, defaultValue, null);
                migratedSettings.Add(setting.JsonName);
            }

            return migratedSettings;
        }

        internal static TabGroupConfig CreateFallbackDefaultConfig()
        {
            return new TabGroupConfig
            {
                Settings = new TabGroupSettings
                {
                    EnableLogging = false,
                    EnableVerboseLogging = false,
                    EnableAutoRename = true,
                    EnableAutoColor = true,
                    EnableConfigurePrompt = true,
                    EnableConnectionPolling = true,
                    EnableColorWarning = true,
                    EnableServerAliasPrompt = true,
                    EnableUpdateChecks = true,
                    AutoConfigure = "server db",
                    SuggestedGroupNameStyle = "[serverAlias] [db]",
                    NewQueryRenameStyle = "[#]. [groupName]",
                    SavedFileRenameStyle = "[filename]",
                    EnableRemoveDotSql = true,
                    EnableLineIndicatorColor = true,
                    EnableStatusBarColor = true
                }
            };
        }

        private static TabGroupSettings CloneSettings(TabGroupSettings settings)
        {
            if (settings == null)
            {
                return new TabGroupSettings();
            }

            return new TabGroupSettings
            {
                EnableLogging = settings.EnableLogging,
                EnableVerboseLogging = settings.EnableVerboseLogging,
                EnableAutoRename = settings.EnableAutoRename,
                EnableAutoColor = settings.EnableAutoColor,
                EnableConfigurePrompt = settings.EnableConfigurePrompt,
                EnableConnectionPolling = settings.EnableConnectionPolling,
                EnableColorWarning = settings.EnableColorWarning,
                EnableServerAliasPrompt = settings.EnableServerAliasPrompt,
                EnableUpdateChecks = settings.EnableUpdateChecks,
                AutoConfigure = settings.AutoConfigure,
                NewQueryRenameStyle = settings.NewQueryRenameStyle,
                SuggestedGroupNameStyle = settings.SuggestedGroupNameStyle,
                SavedFileRenameStyle = settings.SavedFileRenameStyle,
                EnableRemoveDotSql = settings.EnableRemoveDotSql,
                EnableLineIndicatorColor = settings.EnableLineIndicatorColor,
                EnableStatusBarColor = settings.EnableStatusBarColor
            };
        }

        private static bool JsonContainsProperty(string jsonText, string propertyName)
        {
            if (string.IsNullOrWhiteSpace(jsonText) || string.IsNullOrWhiteSpace(propertyName))
            {
                return false;
            }

            return jsonText.IndexOf("\"" + propertyName + "\"", StringComparison.Ordinal) >= 0;
        }

        private static bool TryExtractObjectJson(string jsonText, string propertyName, out string objectJson)
        {
            objectJson = null;
            if (string.IsNullOrWhiteSpace(jsonText) || string.IsNullOrWhiteSpace(propertyName))
            {
                return false;
            }

            string propertyToken = "\"" + propertyName + "\"";
            for (int i = 0; i < jsonText.Length; i++)
            {
                if (jsonText[i] != '"')
                {
                    continue;
                }

                int tokenStart = i;
                i++;
                bool escaped = false;
                while (i < jsonText.Length)
                {
                    char current = jsonText[i];
                    if (escaped)
                    {
                        escaped = false;
                        i++;
                        continue;
                    }

                    if (current == '\\')
                    {
                        escaped = true;
                        i++;
                        continue;
                    }

                    if (current == '"')
                    {
                        break;
                    }

                    i++;
                }

                if (i >= jsonText.Length)
                {
                    return false;
                }

                string token = jsonText.Substring(tokenStart, (i - tokenStart) + 1);
                if (!string.Equals(token, propertyToken, StringComparison.Ordinal))
                {
                    continue;
                }

                int colonIndex = SkipWhitespace(jsonText, i + 1);
                if (colonIndex >= jsonText.Length || jsonText[colonIndex] != ':')
                {
                    continue;
                }

                int valueIndex = SkipWhitespace(jsonText, colonIndex + 1);
                if (valueIndex >= jsonText.Length || jsonText[valueIndex] != '{')
                {
                    return false;
                }

                int objectEnd = FindMatchingBrace(jsonText, valueIndex);
                if (objectEnd < 0)
                {
                    return false;
                }

                objectJson = jsonText.Substring(valueIndex, (objectEnd - valueIndex) + 1);
                return true;
            }

            return false;
        }

        private static int SkipWhitespace(string text, int startIndex)
        {
            int index = startIndex;
            while (index < text.Length && char.IsWhiteSpace(text[index]))
            {
                index++;
            }

            return index;
        }

        private static int FindMatchingBrace(string text, int openingBraceIndex)
        {
            int depth = 0;
            bool inString = false;
            bool escaped = false;

            for (int i = openingBraceIndex; i < text.Length; i++)
            {
                char current = text[i];
                if (inString)
                {
                    if (escaped)
                    {
                        escaped = false;
                        continue;
                    }

                    if (current == '\\')
                    {
                        escaped = true;
                        continue;
                    }

                    if (current == '"')
                    {
                        inString = false;
                    }

                    continue;
                }

                if (current == '"')
                {
                    inString = true;
                    continue;
                }

                if (current == '{')
                {
                    depth++;
                    continue;
                }

                if (current == '}')
                {
                    depth--;
                    if (depth == 0)
                    {
                        return i;
                    }
                }
            }

            return -1;
        }
    }
}