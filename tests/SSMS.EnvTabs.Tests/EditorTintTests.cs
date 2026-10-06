using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SSMS_EnvTabs.Tests
{
    [TestClass]
    public class EditorTintTests
    {
        [TestMethod]
        public void Blend_DarkThemeBackground_ProducesSubtleTint()
        {
            // Burgundy (#cf6468) at 8% over the dark editor background (#1e1e1e).
            int blended = EditorTint.Blend(0x1E1E1E, 0xCF6468, 8);

            Assert.AreEqual(0x2C2424, blended);
        }

        [TestMethod]
        public void Blend_LightThemeBackground_ProducesSubtleTint()
        {
            int blended = EditorTint.Blend(0xFFFFFF, 0xCF6468, 8);

            Assert.AreEqual(0xFBF3F3, blended);
        }

        [TestMethod]
        public void Blend_SameColor_IsUnchanged()
        {
            Assert.AreEqual(0x30B1CD, EditorTint.Blend(0x30B1CD, 0x30B1CD, 20));
        }

        [TestMethod]
        public void ClampStrength_KeepsValueInRange()
        {
            Assert.AreEqual(EditorTint.MinStrength, EditorTint.ClampStrength(0));
            Assert.AreEqual(EditorTint.MinStrength, EditorTint.ClampStrength(-5));
            Assert.AreEqual(12, EditorTint.ClampStrength(12));
            Assert.AreEqual(EditorTint.MaxStrength, EditorTint.ClampStrength(100));
        }
    }
}
