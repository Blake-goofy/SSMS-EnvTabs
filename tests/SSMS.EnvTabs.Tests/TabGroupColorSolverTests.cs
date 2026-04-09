using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SSMS_EnvTabs.Tests
{
    [TestClass]
    public class TabGroupColorSolverTests
    {
        [TestMethod]
        public void Solve_HonorsForbiddenHashesAndStillFindsTargetColor()
        {
            const string baseRegex = "(?:^|[\\\\/])(?:orders)\\.sql$";
            const int targetColor = 7;

            string firstSalt = TabGroupColorSolver.Solve(baseRegex, targetColor);
            Assert.IsFalse(string.IsNullOrWhiteSpace(firstSalt), "Expected first salt candidate.");

            string firstRegex = baseRegex + "(?#salt:" + firstSalt + ")";
            int firstHash = TabGroupColorSolver.GetSsmsStableHashCode(firstRegex);

            var forbidden = new HashSet<int> { firstHash };
            string secondSalt = TabGroupColorSolver.Solve(baseRegex, targetColor, forbidden);

            Assert.IsFalse(string.IsNullOrWhiteSpace(secondSalt), "Expected second salt candidate.");
            string secondRegex = baseRegex + "(?#salt:" + secondSalt + ")";
            int secondHash = TabGroupColorSolver.GetSsmsStableHashCode(secondRegex);

            Assert.AreNotEqual(firstHash, secondHash, "Forbidden hash should be avoided.");
            Assert.AreEqual(targetColor, Math.Abs(secondHash) % 16, "Salt should still map to requested color index.");
        }

        [TestMethod]
        public void Solve_ReturnsNullForInvalidTargetColor()
        {
            Assert.IsNull(TabGroupColorSolver.Solve("abc", -1));
            Assert.IsNull(TabGroupColorSolver.Solve("abc", 16));
        }
    }
}
