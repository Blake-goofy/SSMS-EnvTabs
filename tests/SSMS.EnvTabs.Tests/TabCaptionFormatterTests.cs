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
    }
}
