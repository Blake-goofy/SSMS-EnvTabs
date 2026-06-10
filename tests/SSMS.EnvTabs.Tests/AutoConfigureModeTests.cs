using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SSMS_EnvTabs.Tests
{
    [TestClass]
    public class AutoConfigureModeTests
    {
        [TestMethod]
        public void IsEnabled_ReturnsExpectedValues()
        {
            Assert.IsTrue(AutoConfigureMode.IsEnabled("server"));
            Assert.IsTrue(AutoConfigureMode.IsEnabled("server db"));
            Assert.IsFalse(AutoConfigureMode.IsEnabled("off"));
            Assert.IsFalse(AutoConfigureMode.IsEnabled(" off "));
            Assert.IsTrue(AutoConfigureMode.IsEnabled(null));
            Assert.IsTrue(AutoConfigureMode.IsEnabled(string.Empty));
            Assert.IsTrue(AutoConfigureMode.IsEnabled("unknown"));
        }

        [TestMethod]
        public void UsesDatabase_ReturnsExpectedValues()
        {
            Assert.IsTrue(AutoConfigureMode.UsesDatabase("server db"));
            Assert.IsFalse(AutoConfigureMode.UsesDatabase("server"));
            Assert.IsFalse(AutoConfigureMode.UsesDatabase("off"));
        }

        [TestMethod]
        public void Normalize_ReturnsKnownModesAndDefault()
        {
            Assert.AreEqual(AutoConfigureMode.Server, AutoConfigureMode.Normalize("server"));
            Assert.AreEqual(AutoConfigureMode.ServerDatabase, AutoConfigureMode.Normalize("server db"));
            Assert.AreEqual(AutoConfigureMode.Off, AutoConfigureMode.Normalize("off"));
            Assert.AreEqual(AutoConfigureMode.Off, AutoConfigureMode.Normalize(" off "));
            Assert.AreEqual(AutoConfigureMode.ServerDatabase, AutoConfigureMode.Normalize("unknown"));
    }
}
