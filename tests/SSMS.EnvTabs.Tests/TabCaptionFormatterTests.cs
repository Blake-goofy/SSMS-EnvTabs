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
        public void GetPureName_RemovesExecutingPrefix()
        {
            string pure = TabCaptionFormatter.GetPureName("Executing... report.sql* \u2B24 - PROD", " - PROD", enableRemoveDotSql: true);

            Assert.AreEqual("report", pure);
        }

        [TestMethod]
        public void GetPureName_RemovesExecutingSuffix()
        {
            string pure = TabCaptionFormatter.GetPureName("Prod db -  Executing... \u2B24", "", enableRemoveDotSql: true);

            Assert.AreEqual("Prod db", pure);
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
        public void StripDirtyIndicators_RemovesRepeatedCircleMarkers()
        {
            Assert.AreEqual("xidjdmdj..sql", TabCaptionFormatter.StripDirtyIndicators("xidjdmdj..sql \u2B24 \u2B24"));
        }

        [TestMethod]
        public void StripDirtyIndicators_RemovesRepeatedStarMarkers()
        {
            Assert.AreEqual("xidjdmdj..sql", TabCaptionFormatter.StripDirtyIndicators("xidjdmdj..sql**"));
        }

        [TestMethod]
        public void StripDirtyIndicators_RemovesMixedDirtyMarkers()
        {
            Assert.AreEqual("xidjdmdj..sql", TabCaptionFormatter.StripDirtyIndicators("xidjdmdj..sql* \u2B24 *"));
        }

        [TestMethod]
        public void StripExecutionPrefix_RemovesExecutingPrefix()
        {
            Assert.AreEqual("report.sql", TabCaptionFormatter.StripExecutionPrefix("Executing... report.sql"));
        }

        [TestMethod]
        public void StripExecutionPrefix_RemovesExecutingSuffix()
        {
            Assert.AreEqual("report.sql", TabCaptionFormatter.StripExecutionPrefix("report.sql -  Executing..."));
        }

        [TestMethod]
        public void HasExecutionMarker_DetectsSuffixForm()
        {
            Assert.IsTrue(TabCaptionFormatter.HasExecutionMarker("SQLQuery13.sql -  Executing... \u2B24"));
        }

        [TestMethod]
        public void CaptionsEquivalent_IgnoresDirtyIndicatorsAndExecutingPrefix()
        {
            bool equivalent = TabCaptionFormatter.CaptionsEquivalent(
                "Executing... report.sql* \u2B24 - PROD",
                "report",
                " - PROD",
                enableRemoveDotSql: true);

            Assert.IsTrue(equivalent);
        }

        [TestMethod]
        public void CaptionsEquivalent_IgnoresExecutingSuffix()
        {
            bool equivalent = TabCaptionFormatter.CaptionsEquivalent(
                "Prod db -  Executing... \u2B24",
                "Prod db",
                string.Empty,
                enableRemoveDotSql: true);

            Assert.IsTrue(equivalent);
        }

        [TestMethod]
        public void CaptionsEquivalent_DistinguishesCustomCaptionFromGeneratedCaption()
        {
            bool equivalent = TabCaptionFormatter.CaptionsEquivalent(
                "Customer Investigation",
                "1. PROD",
                "",
                enableRemoveDotSql: true);

            Assert.IsFalse(equivalent);
        }

        [TestMethod]
        public void SelectRenameSourceCaption_UsesObservedCaptionForCaptionPoll()
        {
            string selected = TabCaptionFormatter.SelectRenameSourceCaption(
                "Local model",
                "1xf1xih1..sql \u2B24",
                "CaptionPoll");

            Assert.AreEqual("1xf1xih1..sql \u2B24", selected);
        }

        [TestMethod]
        public void SelectRenameSourceCaption_KeepsFrameCaptionForOtherReasons()
        {
            string selected = TabCaptionFormatter.SelectRenameSourceCaption(
                "Local model",
                "1xf1xih1..sql \u2B24",
                "ConnectionPoll");

            Assert.AreEqual("Local model", selected);
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

        [TestMethod]
        public void BuildStyleCaption_ReplacesAllTokensForNewQueryStyle()
        {
            string style = "[serverAlias] - [db] - [filename] - [groupName] - [#]";
            string result = TabCaptionFormatter.BuildStyleCaption(
                style,
                "SQLQuery1",
                "Prod",
                "ProdServer01",
                "ProdAlias",
                "SalesDb",
                2);

            Assert.AreEqual("ProdAlias - SalesDb - SQLQuery1 - Prod - 2", result);
        }

        [TestMethod]
        public void BuildStyleCaption_ReplacesIndexTokenInSavedFileStyle()
        {
            string result = TabCaptionFormatter.BuildStyleCaption("[filename] [#]", "daily_report", "Prod", "Srv", null, "Db", 3);

            Assert.AreEqual("daily_report 3", result);
        }

        [TestMethod]
        public void BuildStyleCaption_NullIndexRemovesIndexToken()
        {
            string result = TabCaptionFormatter.BuildStyleCaption("[filename] [#]", "daily_report", "Prod", "Srv", null, "Db", null);

            Assert.AreEqual("daily_report ", result);
        }

        [TestMethod]
        public void GetFilenameToken_RemovesSqlExtensionWhenEnabled()
        {
            string result = TabCaptionFormatter.GetFilenameToken(@"C:\Queries\daily_report.sql", enableRemoveDotSql: true);

            Assert.AreEqual("daily_report", result);
        }

        [TestMethod]
        public void GetFilenameToken_PreservesSqlExtensionWhenDisabled()
        {
            string result = TabCaptionFormatter.GetFilenameToken(@"C:\Queries\daily_report.sql", enableRemoveDotSql: false);

            Assert.AreEqual("daily_report.sql", result);
        }
    }
}
