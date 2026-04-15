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

        [TestMethod]
        public void Solve_FindsSaltForAllSixteenColorIndices()
        {
            const string baseRegex = "(?:^|[\\\\/])(?:test)\\.sql$";

            for (int target = 0; target <= 15; target++)
            {
                string salt = TabGroupColorSolver.Solve(baseRegex, target);
                Assert.IsNotNull(salt, $"Expected salt for color index {target}.");

                string fullRegex = baseRegex + "(?#salt:" + salt + ")";
                int hash = TabGroupColorSolver.GetSsmsStableHashCode(fullRegex);
                Assert.AreEqual(target, Math.Abs(hash) % 16, $"Salt did not produce expected color index {target}.");
            }
        }

        [TestMethod]
        public void SolveForColor_ReturnsRegexWithSaltAppended()
        {
            const string baseRegex = "(?:^|[\\\\/])(?:demo)\\.sql$";
            const int target = 3;

            string result = TabGroupColorSolver.SolveForColor(baseRegex, target);

            Assert.IsTrue(result.StartsWith(baseRegex), "Result should start with base regex.");
            Assert.IsTrue(result.Contains("(?#salt:"), "Result should contain salt comment.");

            int hash = TabGroupColorSolver.GetSsmsStableHashCode(result);
            Assert.AreEqual(target, Math.Abs(hash) % 16);
        }

        [TestMethod]
        public void SolveForColor_ReturnsBaseRegexWhenSolveReturnsNull()
        {
            string result = TabGroupColorSolver.SolveForColor("abc", -1);

            Assert.AreEqual("abc", result);
        }

        [TestMethod]
        public void GetSsmsStableHashCode_DeterministicForSameInput()
        {
            const string input = "(?:^|[\\\\/])(?:query1|query2)\\.sql$";

            int hash1 = TabGroupColorSolver.GetSsmsStableHashCode(input);
            int hash2 = TabGroupColorSolver.GetSsmsStableHashCode(input);

            Assert.AreEqual(hash1, hash2);
        }

        [TestMethod]
        public void GetSsmsStableHashCode_DifferentInputsDifferentHashes()
        {
            int hash1 = TabGroupColorSolver.GetSsmsStableHashCode("alpha");
            int hash2 = TabGroupColorSolver.GetSsmsStableHashCode("beta");

            Assert.AreNotEqual(hash1, hash2);
        }

        [TestMethod]
        public void GetSsmsStableHashCode_EmptyStringDoesNotThrow()
        {
            int hash = TabGroupColorSolver.GetSsmsStableHashCode("");
            Assert.AreEqual(hash, TabGroupColorSolver.GetSsmsStableHashCode(""));
        }

        [TestMethod]
        public void GetSsmsStableHashCode_OddLengthStringDoesNotThrow()
        {
            int hash = TabGroupColorSolver.GetSsmsStableHashCode("abc");
            Assert.AreEqual(hash, TabGroupColorSolver.GetSsmsStableHashCode("abc"));
        }
    }
}
