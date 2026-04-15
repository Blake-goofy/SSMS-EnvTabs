using System.Linq;
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
                new[] { "enableVerboseLogging", "enableAutoRename", "enableAutoColor", "enableConfigurePrompt", "enableConnectionPolling", "enableColorWarning", "enableServerAliasPrompt", "enableUpdateChecks", "autoConfigure", "newQueryRenameStyle", "suggestedGroupNameStyle", "savedFileRenameStyle", "initialLineIndicatorColor", "initialStatusBarColor" },
                migrated.ToArray());
            Assert.IsTrue(config.Settings.InitialLineIndicatorColor);
            Assert.IsTrue(config.Settings.InitialStatusBarColor);
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
                    InitialStatusBarColor = false
                }
            };

            var defaultConfig = TabGroupConfigDefaults.CreateFallbackDefaultConfig();
            string json = "{\"settings\":{\"enableLogging\":false,\"initialLineIndicatorColor\":false,\"initialStatusBarColor\":false}}";

            var migrated = TabGroupConfigDefaults.ApplyMissingSettingDefaults(config, json, defaultConfig);

            Assert.IsFalse(migrated.Contains("initialLineIndicatorColor"));
            Assert.IsFalse(migrated.Contains("initialStatusBarColor"));
            Assert.IsFalse(config.Settings.InitialLineIndicatorColor);
            Assert.IsFalse(config.Settings.InitialStatusBarColor);
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

            Assert.AreEqual(16, migrated.Count);
            Assert.IsNotNull(config.Settings);
            Assert.IsTrue(config.Settings.EnableAutoColor);
            Assert.IsTrue(config.Settings.InitialLineIndicatorColor);
            Assert.IsTrue(config.Settings.InitialStatusBarColor);
        }
    }
}