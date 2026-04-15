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
        public void CompileRules_NullGroupRuleIncludedInCompiledList()
        {
            var config = new TabGroupConfig
            {
                ConnectionGroups = new List<TabGroupRule>
                {
                    new TabGroupRule { GroupName = "Prod", Server = "PROD-01", Database = "%", Priority = 10, ColorIndex = 5 },
                    new TabGroupRule { GroupName = null, Server = "(localdb)\\MSSQLLocalDB", Database = "tempdb", Priority = 10, ColorIndex = null }
                }
            };

            var rules = TabRuleMatcher.CompileRules(config);

            Assert.AreEqual(2, rules.Count);
            var nullGroupRule = rules.Find(r => r.GroupName == null);
            Assert.IsNotNull(nullGroupRule);
            Assert.IsNull(nullGroupRule.ColorIndex);
        }

        [TestMethod]
        public void CompileRules_SkipsRulesWithNeitherServerNorDatabase()
        {
            var config = new TabGroupConfig
            {
                ConnectionGroups = new List<TabGroupRule>
                {
                    new TabGroupRule { GroupName = "Empty", Server = null, Database = null, Priority = 10, ColorIndex = 1 },
                    new TabGroupRule { GroupName = "Valid", Server = "SRV", Database = null, Priority = 10, ColorIndex = 2 }
                }
            };

            var rules = TabRuleMatcher.CompileRules(config);

            Assert.AreEqual(1, rules.Count);
            Assert.AreEqual("Valid", rules[0].GroupName);
        }

        [TestMethod]
        public void CompileRules_SkipsNullRuleEntries()
        {
            var config = new TabGroupConfig
            {
                ConnectionGroups = new List<TabGroupRule>
                {
                    null,
                    new TabGroupRule { GroupName = "OK", Server = "SRV", Priority = 10, ColorIndex = 0 },
                    null
                }
            };

            var rules = TabRuleMatcher.CompileRules(config);

            Assert.AreEqual(1, rules.Count);
            Assert.AreEqual("OK", rules[0].GroupName);
        }

        [TestMethod]
        public void CompileRules_NullConfigReturnsEmptyList()
        {
            var rules = TabRuleMatcher.CompileRules(null);

            Assert.AreEqual(0, rules.Count);
        }

        [TestMethod]
        public void CompileRules_NullConnectionGroupsReturnsEmptyList()
        {
            var config = new TabGroupConfig { ConnectionGroups = null };

            var rules = TabRuleMatcher.CompileRules(config);

            Assert.AreEqual(0, rules.Count);
        }

        [TestMethod]
        public void MatchRule_NullRulesListReturnsNull()
        {
            Assert.IsNull(TabRuleMatcher.MatchRule(null, "SRV", "DB"));
        }

        [TestMethod]
        public void MatchRule_EmptyRulesListReturnsNull()
        {
            Assert.IsNull(TabRuleMatcher.MatchRule(new List<TabRuleMatcher.CompiledRule>(), "SRV", "DB"));
        }

        [TestMethod]
        public void MatchRule_ExactServerAndDatabaseMatch()
        {
            var config = new TabGroupConfig
            {
                ConnectionGroups = new List<TabGroupRule>
                {
                    new TabGroupRule { GroupName = "Exact", Server = "MyServer", Database = "MyDB", Priority = 10, ColorIndex = 3 }
                }
            };

            var rules = TabRuleMatcher.CompileRules(config);

            Assert.IsNotNull(TabRuleMatcher.MatchRule(rules, "MyServer", "MyDB"));
            Assert.IsNull(TabRuleMatcher.MatchRule(rules, "MyServer", "OtherDB"));
            Assert.IsNull(TabRuleMatcher.MatchRule(rules, "OtherServer", "MyDB"));
        }

        [TestMethod]
        public void MatchRule_ServerOnlyRuleMatchesAnyDatabase()
        {
            var config = new TabGroupConfig
            {
                ConnectionGroups = new List<TabGroupRule>
                {
                    new TabGroupRule { GroupName = "ServerOnly", Server = "SRV", Database = null, Priority = 10, ColorIndex = 1 }
                }
            };

            var rules = TabRuleMatcher.CompileRules(config);

            Assert.IsNotNull(TabRuleMatcher.MatchRule(rules, "SRV", "AnyDB"));
            Assert.IsNotNull(TabRuleMatcher.MatchRule(rules, "SRV", null));
            Assert.IsNull(TabRuleMatcher.MatchRule(rules, "OTHER", "AnyDB"));
        }

        [TestMethod]
        public void MatchRule_WildcardServerMatchesCaseInsensitive()
        {
            var config = new TabGroupConfig
            {
                ConnectionGroups = new List<TabGroupRule>
                {
                    new TabGroupRule { GroupName = "Wild", Server = "PROD%", Database = "%", Priority = 10, ColorIndex = 4 }
                }
            };

            var rules = TabRuleMatcher.CompileRules(config);

            Assert.IsNotNull(TabRuleMatcher.MatchRule(rules, "prod-east", "sales"));
            Assert.IsNotNull(TabRuleMatcher.MatchRule(rules, "PROD-WEST", "ORDERS"));
            Assert.IsNull(TabRuleMatcher.MatchRule(rules, "dev-east", "sales"));
        }

        [TestMethod]
        public void MatchGroup_ReturnsNullForNullGroupRule()
        {
            var config = new TabGroupConfig
            {
                ConnectionGroups = new List<TabGroupRule>
                {
                    new TabGroupRule { GroupName = null, Server = "SRV", Database = "%", Priority = 10, ColorIndex = null }
                }
            };

            var rules = TabRuleMatcher.CompileRules(config);

            Assert.IsNull(TabRuleMatcher.MatchGroup(rules, "SRV", "master"));
        }

        [TestMethod]
        public void MatchGroup_ReturnsNullWhenNoMatch()
        {
            var config = new TabGroupConfig
            {
                ConnectionGroups = new List<TabGroupRule>
                {
                    new TabGroupRule { GroupName = "Prod", Server = "PROD", Database = "%", Priority = 10, ColorIndex = 1 }
                }
            };

            var rules = TabRuleMatcher.CompileRules(config);

            Assert.IsNull(TabRuleMatcher.MatchGroup(rules, "DEV", "master"));
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

        [TestMethod]
        public void CompileManualRules_SkipsBlankPatterns()
        {
            var config = new TabGroupConfig
            {
                ManualRegexLines = new List<ManualRegexEntry>
                {
                    new ManualRegexEntry { GroupName = "Blank", Pattern = "", Priority = 10, ColorIndex = 1 },
                    new ManualRegexEntry { GroupName = "Null", Pattern = null, Priority = 10, ColorIndex = 1 },
                    new ManualRegexEntry { GroupName = "Ok", Pattern = ".*ok.*", Priority = 10, ColorIndex = 2 }
                }
            };

            var manuals = TabRuleMatcher.CompileManualRules(config);

            Assert.AreEqual(1, manuals.Count);
            Assert.AreEqual("Ok", manuals[0].GroupName);
        }

        [TestMethod]
        public void CompileManualRules_NullConfigReturnsEmptyList()
        {
            Assert.AreEqual(0, TabRuleMatcher.CompileManualRules(null).Count);
        }

        [TestMethod]
        public void MatchManual_NullMonikerReturnsNull()
        {
            var config = new TabGroupConfig
            {
                ManualRegexLines = new List<ManualRegexEntry>
                {
                    new ManualRegexEntry { GroupName = "Test", Pattern = ".*", Priority = 10, ColorIndex = 0 }
                }
            };

            var manuals = TabRuleMatcher.CompileManualRules(config);

            Assert.IsNull(TabRuleMatcher.MatchManual(manuals, null));
            Assert.IsNull(TabRuleMatcher.MatchManual(manuals, ""));
            Assert.IsNull(TabRuleMatcher.MatchManual(manuals, "   "));
        }

        [TestMethod]
        public void MatchManual_InvalidRegexDoesNotThrow()
        {
            var config = new TabGroupConfig
            {
                ManualRegexLines = new List<ManualRegexEntry>
                {
                    new ManualRegexEntry { GroupName = "Bad", Pattern = "[invalid", Priority = 10, ColorIndex = 0 },
                    new ManualRegexEntry { GroupName = "Good", Pattern = ".*good.*", Priority = 20, ColorIndex = 1 }
                }
            };

            var manuals = TabRuleMatcher.CompileManualRules(config);

            Assert.AreEqual(2, manuals.Count);
            Assert.IsNull(manuals[0].FileRegex);
            var match = TabRuleMatcher.MatchManual(manuals, @"C:\temp\good_file.sql");
            Assert.IsNotNull(match);
            Assert.AreEqual("Good", match.GroupName);
        }

        [TestMethod]
        public void MatchRule_NullGroupRuleCoexistsWithNamedRules()
        {
            var config = new TabGroupConfig
            {
                ConnectionGroups = new List<TabGroupRule>
                {
                    new TabGroupRule { GroupName = null, Server = "(localdb)\\MSSQLLocalDB", Database = "tempdb", Priority = 10, ColorIndex = null },
                    new TabGroupRule { GroupName = "Production", Server = "PROD%", Database = "%", Priority = 20, ColorIndex = 5 }
                }
            };

            var rules = TabRuleMatcher.CompileRules(config);

            var localMatch = TabRuleMatcher.MatchRule(rules, "(localdb)\\MSSQLLocalDB", "tempdb");
            Assert.IsNotNull(localMatch);
            Assert.IsNull(localMatch.GroupName);

            var prodMatch = TabRuleMatcher.MatchRule(rules, "PROD-EAST", "Sales");
            Assert.IsNotNull(prodMatch);
            Assert.AreEqual("Production", prodMatch.GroupName);
            Assert.AreEqual(5, prodMatch.ColorIndex);
        }
    }
}
