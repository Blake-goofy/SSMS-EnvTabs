using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SSMS_EnvTabs.Tests
{
    [TestClass]
    public class TabCaptionFormatterTests
    {
        [TestMethod]
        public void GetPureName_RemovesDirtyIndicatorsSuffixAndDotSql()
        {
            string caption = "report.sql* \u2B24 - PROD";
            string pure = TabCaptionFormatter.GetPureName(caption, " - PROD", enableRemoveDotSql: true);

            Assert.AreEqual("report", pure);
        }

        [TestMethod]
        public void GetPureName_PreservesDotSqlWhenDisabled()
        {
            string pure = TabCaptionFormatter.GetPureName("report.sql - SRV", " - SRV", enableRemoveDotSql: false);

            Assert.AreEqual("report.sql", pure);
        }

        [TestMethod]
        public void GetPureName_NullCaptionReturnsEmpty()
        {
            Assert.AreEqual(string.Empty, TabCaptionFormatter.GetPureName(null, " - SRV", true));
        }

        [TestMethod]
        public void GetPureName_EmptyCaptionReturnsEmpty()
        {
            Assert.AreEqual(string.Empty, TabCaptionFormatter.GetPureName("", " - SRV", true));
        }

        [TestMethod]
        public void GetPureName_NullSuffixStripsOnlyDotSql()
        {
            string pure = TabCaptionFormatter.GetPureName("script.sql", null, enableRemoveDotSql: true);

            Assert.AreEqual("script", pure);
        }

        [TestMethod]
        public void GetPureName_ConsecutiveDirtyIndicators()
        {
            string pure = TabCaptionFormatter.GetPureName("test** \u2B24 \u2B24", "", enableRemoveDotSql: true);

            Assert.AreEqual("test", pure);
        }

        [TestMethod]
        public void GetPureName_CaptionIsOnlySuffix()
        {
            string pure = TabCaptionFormatter.GetPureName("query - SERVER", " - SERVER", enableRemoveDotSql: true);

            Assert.AreEqual("query", pure);
        }

        [TestMethod]
        public void StripDirtyIndicators_NullReturnsEmpty()
        {
            Assert.AreEqual(string.Empty, TabCaptionFormatter.StripDirtyIndicators(null));
        }

        [TestMethod]
        public void StripDirtyIndicators_NoIndicatorsReturnsTrimmed()
        {
            Assert.AreEqual("clean", TabCaptionFormatter.StripDirtyIndicators("  clean  "));
        }

        [TestMethod]
        public void StripSqlExtension_CaseInsensitive()
        {
            Assert.AreEqual("file", TabCaptionFormatter.StripSqlExtension("file.SQL", true));
            Assert.AreEqual("file", TabCaptionFormatter.StripSqlExtension("file.Sql", true));
        }

        [TestMethod]
        public void StripSqlExtension_DisabledPreservesExtension()
        {
            Assert.AreEqual("file.sql", TabCaptionFormatter.StripSqlExtension("file.sql", false));
        }

        [TestMethod]
        public void BuildSavedStyleCaption_ReplacesExpectedTokens()
        {
            string style = "[filename] [groupName] [serverAlias] [db]";
            string caption = TabCaptionFormatter.BuildSavedStyleCaption(
                style,
                "daily_report",
                "Prod",
                "ProdServer01",
                "Prod",
                "SalesDb");

            Assert.AreEqual("daily_report Prod Prod SalesDb", caption);
        }

        [TestMethod]
        public void BuildSavedStyleCaption_NullTokensReplaceWithEmpty()
        {
            string style = "[filename]-[groupName]-[server]-[serverAlias]-[db]";
            string result = TabCaptionFormatter.BuildSavedStyleCaption(style, null, null, null, null, null);

            Assert.AreEqual("----", result);
        }

        [TestMethod]
        public void BuildSavedStyleCaption_NullStyleReturnsEmpty()
        {
            string result = TabCaptionFormatter.BuildSavedStyleCaption(null, "file", "group", "srv", "alias", "db");

            Assert.AreEqual(string.Empty, result);
        }

        [TestMethod]
        public void BuildSavedStyleCaption_ServerAliasFallsBackToServer()
        {
            string style = "[serverAlias]";
            string result = TabCaptionFormatter.BuildSavedStyleCaption(style, "f", "g", "PROD-01", null, "db");

            Assert.AreEqual("PROD-01", result);
        }
    }
}
