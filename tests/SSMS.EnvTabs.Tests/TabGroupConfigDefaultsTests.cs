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
                    EnableLineIndicatorColor = false,
                    EnableStatusBarColor = false
                }
            };

            var defaultConfig = TabGroupConfigDefaults.CreateFallbackDefaultConfig();
            string json = "{\"settings\":{\"enableLogging\":false,\"enableRemoveDotSql\":true}}";

            var migrated = TabGroupConfigDefaults.ApplyMissingSettingDefaults(config, json, defaultConfig);

            CollectionAssert.AreEquivalent(
                new[] { "enableVerboseLogging", "enableAutoRename", "enableAutoColor", "enableConfigurePrompt", "enableConnectionPolling", "enableColorWarning", "enableServerAliasPrompt", "enableUpdateChecks", "autoConfigure", "newQueryRenameStyle", "suggestedGroupNameStyle", "savedFileRenameStyle", "enableLineIndicatorColor", "enableStatusBarColor" },
                migrated.ToArray());
            Assert.IsTrue(config.Settings.EnableLineIndicatorColor);
            Assert.IsTrue(config.Settings.EnableStatusBarColor);
        }

        [TestMethod]
        public void ApplyMissingSettingDefaults_PreservesExplicitUserValues()
        {
            var config = new TabGroupConfig
            {
                Settings = new TabGroupSettings
                {
                    EnableLogging = false,
                    EnableLineIndicatorColor = false,
                    EnableStatusBarColor = false
                }
            };

            var defaultConfig = TabGroupConfigDefaults.CreateFallbackDefaultConfig();
            string json = "{\"settings\":{\"enableLogging\":false,\"enableLineIndicatorColor\":false,\"enableStatusBarColor\":false}}";

            var migrated = TabGroupConfigDefaults.ApplyMissingSettingDefaults(config, json, defaultConfig);

            Assert.IsFalse(migrated.Contains("enableLineIndicatorColor"));
            Assert.IsFalse(migrated.Contains("enableStatusBarColor"));
            Assert.IsFalse(config.Settings.EnableLineIndicatorColor);
            Assert.IsFalse(config.Settings.EnableStatusBarColor);
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
            Assert.IsTrue(config.Settings.EnableLineIndicatorColor);
            Assert.IsTrue(config.Settings.EnableStatusBarColor);
        }
    }
}