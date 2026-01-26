using System;
using System.IO;
using System.Reflection;
using System.Runtime.Serialization.Json;
using System.Text;

namespace SSMS_EnvTabs
{
    internal static class TabGroupConfigLoader
    {
        private const string DefaultConfigResourceName = "SSMS_EnvTabs.DefaultTabGroupConfig.json";

        public static string GetUserConfigPath()
        {
            string docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            return Path.Combine(docs, "SSMS EnvTabs", "TabGroupConfig.json");
        }

        public static void EnsureDefaultConfigExists()
        {
            try
            {
                string path = GetUserConfigPath();
                string dir = Path.GetDirectoryName(path);
                if (!Directory.Exists(dir))
                {
                    Directory.CreateDirectory(dir);
                }

                if (!File.Exists(path))
                {
                    string defaultJson = ReadDefaultConfigJsonOrNull();
                    if (string.IsNullOrWhiteSpace(defaultJson))
                    {
                        defaultJson = "{\n  \"groups\": [],\n  \"settings\": { \"enableAutoRename\": true, \"enableAutoColor\": false }\n}\n";
                    }

                    File.WriteAllText(path, defaultJson, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
                }
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"EnsureDefaultConfigExists failed: {ex.Message}");
            }
        }

        private static string ReadDefaultConfigJsonOrNull()
        {
            try
            {
                Assembly asm = typeof(TabGroupConfigLoader).Assembly;
                using (Stream s = asm.GetManifestResourceStream(DefaultConfigResourceName))
                {
                    if (s == null)
                    {
                        EnvTabsLog.Info($"Default config resource not found: '{DefaultConfigResourceName}'");
                        return null;
                    }

                    using (var reader = new StreamReader(s, Encoding.UTF8, detectEncodingFromByteOrderMarks: true))
                    {
                        return reader.ReadToEnd();
                    }
                }
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"ReadDefaultConfigJson failed: {ex.Message}");
                return null;
            }
        }

        public static TabGroupConfig LoadOrNull()
        {
            try
            {
                string path = GetUserConfigPath();
                if (!File.Exists(path))
                {
                    return null;
                }

                string jsonText;
                using (var reader = new StreamReader(path, Encoding.UTF8, detectEncodingFromByteOrderMarks: true))
                {
                    jsonText = reader.ReadToEnd();
                }

                if (string.IsNullOrWhiteSpace(jsonText))
                {
                    return null;
                }

                jsonText = jsonText.TrimStart('\uFEFF', '\u200B', '\u0000', ' ', '\t', '\r', '\n');

                byte[] utf8Bytes = Encoding.UTF8.GetBytes(jsonText);
                using (var stream = new MemoryStream(utf8Bytes))
                {
                    var serializer = new DataContractJsonSerializer(typeof(TabGroupConfig), new DataContractJsonSerializerSettings
                    {
                        UseSimpleDictionaryFormat = true
                    });
                    return serializer.ReadObject(stream) as TabGroupConfig;
                }
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"Config load failed: {ex.Message}");
                return null;
            }
        }
    }
}
