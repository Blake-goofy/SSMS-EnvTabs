using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SSMS_EnvTabs.Tests
{
    [TestClass]
    public class TabGroupConfigDefaultsTests
    {
        [TestMethod]
        public void ApplyMissingSettingDefaults_RestoresMissingNewSettingsFromDefaults()
        {
            var config = new TabGroupConfig
            {
                Settings = new TabGroupSettings
                {
                    InitialLineIndicatorColor = false,
                    InitialStatusBarColor = false
                }
            };

            var defaultConfig = TabGroupConfigDefaults.CreateFallbackDefaultConfig();
            string json = "{\"settings\":{\"enableLogging\":false,\"enableRemoveDotSql\":true}}";

            var migrated = TabGroupConfigDefaults.ApplyMissingSettingDefaults(config, json, defaultConfig);

            CollectionAssert.AreEquivalent(
                new[] { "enableVerboseLogging", "enableAutoRename", "enableAutoColor", "enableConfigurePrompt", "enableConnectionPolling", "enableColorWarning", "enableServerAliasPrompt", "enableUpdateChecks", "autoConfigure", "newQueryRenameStyle", "suggestedGroupNameStyle", "savedFileRenameStyle", "initialLineIndicatorColor", "initialStatusBarColor", "initialEditorTint", "editorTintStrength" },
                migrated.ToArray());
            Assert.IsTrue(config.Settings.InitialLineIndicatorColor);
            Assert.IsTrue(config.Settings.InitialStatusBarColor);
            Assert.IsFalse(config.Settings.InitialEditorTint);
        }

        [TestMethod]
        public void ApplyMissingSettingDefaults_PreservesExplicitUserValues()
        {
            var config = new TabGroupConfig
            {
                Settings = new TabGroupSettings
                {
                    EnableLogging = false,
                    InitialLineIndicatorColor = false,
                    InitialStatusBarColor = false,
                    InitialEditorTint = true
                }
            };

            var defaultConfig = TabGroupConfigDefaults.CreateFallbackDefaultConfig();
            string json = "{\"settings\":{\"enableLogging\":false,\"initialLineIndicatorColor\":false,\"initialStatusBarColor\":false,\"initialEditorTint\":true}}";

            var migrated = TabGroupConfigDefaults.ApplyMissingSettingDefaults(config, json, defaultConfig);

            Assert.IsFalse(migrated.Contains("initialLineIndicatorColor"));
            Assert.IsFalse(migrated.Contains("initialStatusBarColor"));
            Assert.IsFalse(config.Settings.InitialLineIndicatorColor);
            Assert.IsFalse(config.Settings.InitialStatusBarColor);
            Assert.IsFalse(migrated.Contains("initialEditorTint"));
            Assert.IsTrue(config.Settings.InitialEditorTint);
        }

        [TestMethod]
        public void ApplyMissingSettingDefaults_RestoresEntireSettingsObjectWhenMissing()
        {
            var config = new TabGroupConfig
            {
                Settings = null
            };

            var defaultConfig = TabGroupConfigDefaults.CreateFallbackDefaultConfig();
            string json = "{\"connectionGroups\":[]}";

            var migrated = TabGroupConfigDefaults.ApplyMissingSettingDefaults(config, json, defaultConfig);

            Assert.AreEqual(18, migrated.Count);
            Assert.IsNotNull(config.Settings);
            Assert.IsTrue(config.Settings.EnableAutoColor);
            Assert.IsTrue(config.Settings.InitialLineIndicatorColor);
            Assert.IsTrue(config.Settings.InitialStatusBarColor);
            Assert.IsFalse(config.Settings.InitialEditorTint);
        }

        [TestMethod]
        public void EditorTint_UpgradeDefaultsOffAndPreservesGroupOverrides()
        {
            const string json = "{\"settings\":{\"enableAutoColor\":true},\"connectionGroups\":[{\"server\":\"existing\"},{\"server\":\"opt-in\",\"enableEditorTint\":true},{\"server\":\"opt-out\",\"enableEditorTint\":false}]}";
            var serializer = new DataContractJsonSerializer(typeof(TabGroupConfig));
            using (var stream = new MemoryStream(Encoding.UTF8.GetBytes(json)))
            {
                var config = (TabGroupConfig)serializer.ReadObject(stream);
                TabGroupConfigDefaults.ApplyMissingSettingDefaults(config, json, TabGroupConfigDefaults.CreateFallbackDefaultConfig());

                Assert.IsFalse(new TabGroupSettings().InitialEditorTint);
                Assert.IsFalse(config.Settings.InitialEditorTint);
                Assert.IsFalse(config.ConnectionGroups[0].EnableEditorTint ?? config.Settings.InitialEditorTint);
                Assert.AreEqual(true, config.ConnectionGroups[1].EnableEditorTint);
                Assert.AreEqual(false, config.ConnectionGroups[2].EnableEditorTint);
            }
        }
    }
}
