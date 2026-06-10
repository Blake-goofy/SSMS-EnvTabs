using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SSMS_EnvTabs.Tests
{
    [TestClass]
    public class AutoConfigureModeTests
    {
        [TestMethod]
        public void IsEnabled_ReturnsFalse_WhenModeIsOff()
        {
            Assert.IsFalse(AutoConfigureMode.IsEnabled("off"));
        }

        [TestMethod]
        public void UsesDatabase_ReturnsFalse_WhenModeIsOff()
        {
            Assert.IsFalse(AutoConfigureMode.UsesDatabase("off"));
        }

        [TestMethod]
        public void Normalize_PreservesKnownModes()
        {
            Assert.AreEqual(AutoConfigureMode.Server, AutoConfigureMode.Normalize("server"));
            Assert.AreEqual(AutoConfigureMode.ServerDatabase, AutoConfigureMode.Normalize("server db"));
            Assert.AreEqual(AutoConfigureMode.Off, AutoConfigureMode.Normalize("off"));
        }
    }
}
