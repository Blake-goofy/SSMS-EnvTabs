using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SSMS_EnvTabs.Tests
{
    [TestClass]
    public class TabRuleMatcherTests
    {
        [TestMethod]
        public void MatchRule_UsesPriorityOrderForMoreSpecificRule()
        {
            var config = new TabGroupConfig
            {
                ConnectionGroups = new List<TabGroupRule>
                {
                    new TabGroupRule { GroupName = "ProdFallback", Server = "PROD%", Database = "%", Priority = 20, ColorIndex = 1 },
                    new TabGroupRule { GroupName = "ProdOrders", Server = "PROD%", Database = "Orders", Priority = 10, ColorIndex = 2 }
                }
            };

            var rules = TabRuleMatcher.CompileRules(config);
            var match = TabRuleMatcher.MatchRule(rules, "Prod-East", "Orders");

            Assert.IsNotNull(match);
            Assert.AreEqual("ProdOrders", match.GroupName);
            Assert.AreEqual("ProdOrders", TabRuleMatcher.MatchGroup(rules, "Prod-East", "Orders"));
        }

        [TestMethod]
        public void MatchRule_NullGroupRuleSuppressesRenameButStillMatches()
        {
            var config = new TabGroupConfig
            {
                ConnectionGroups = new List<TabGroupRule>
                {
                    new TabGroupRule { GroupName = null, Server = "LegacyServer", Database = "%", Priority = 10, ColorIndex = null }
                }
            };

            var rules = TabRuleMatcher.CompileRules(config);
            var match = TabRuleMatcher.MatchRule(rules, "LegacyServer", "master");

            Assert.IsNotNull(match);
            Assert.IsNull(match.GroupName);
            Assert.IsNull(TabRuleMatcher.MatchGroup(rules, "LegacyServer", "master"));
        }

        [TestMethod]
        public void CompileManualRules_SortsByPriorityAscending()
        {
            var config = new TabGroupConfig
            {
                ManualRegexLines = new List<ManualRegexEntry>
                {
                    new ManualRegexEntry { GroupName = "Late", Pattern = ".*late.*", Priority = 50, ColorIndex = 1 },
                    new ManualRegexEntry { GroupName = "Early", Pattern = ".*early.*", Priority = 10, ColorIndex = 2 }
                }
            };

            var manuals = TabRuleMatcher.CompileManualRules(config);

            Assert.AreEqual(2, manuals.Count);
            Assert.AreEqual("Early", manuals[0].GroupName);
            Assert.AreEqual("Late", manuals[1].GroupName);
        }
    }
}
