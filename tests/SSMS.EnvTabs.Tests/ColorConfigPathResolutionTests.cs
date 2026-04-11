using System;
using System.Collections.Generic;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SSMS_EnvTabs.Tests
{
    [TestClass]
    public class ColorConfigPathResolutionTests
    {
        [TestMethod]
        public void OrderTempScanCandidates_PrioritizesWindowDistanceLegacyAndWriteTime()
        {
            var candidates = new List<ColorConfigPathResolution.TempScanCandidate>
            {
                new ColorConfigPathResolution.TempScanCandidate
                {
                    Path = @"C:\temp\d\ColorByRegexConfig.txt",
                    GuidName = "dddd",
                    InPreferredWindow = false,
                    AbsDeltaSeconds = 1,
                    HasLegacyVersionFolder = true,
                    LastWriteUtc = new DateTime(2026, 4, 10, 12, 0, 4, DateTimeKind.Utc)
                },
                new ColorConfigPathResolution.TempScanCandidate
                {
                    Path = @"C:\temp\c\ColorByRegexConfig.txt",
                    GuidName = "cccc",
                    InPreferredWindow = true,
                    AbsDeltaSeconds = 10,
                    HasLegacyVersionFolder = false,
                    LastWriteUtc = new DateTime(2026, 4, 10, 12, 0, 3, DateTimeKind.Utc)
                },
                new ColorConfigPathResolution.TempScanCandidate
                {
                    Path = @"C:\temp\b\ColorByRegexConfig.txt",
                    GuidName = "bbbb",
                    InPreferredWindow = true,
                    AbsDeltaSeconds = 5,
                    HasLegacyVersionFolder = true,
                    LastWriteUtc = new DateTime(2026, 4, 10, 12, 0, 2, DateTimeKind.Utc)
                },
                new ColorConfigPathResolution.TempScanCandidate
                {
                    Path = @"C:\temp\a\ColorByRegexConfig.txt",
                    GuidName = "aaaa",
                    InPreferredWindow = true,
                    AbsDeltaSeconds = 5,
                    HasLegacyVersionFolder = false,
                    LastWriteUtc = new DateTime(2026, 4, 10, 12, 0, 5, DateTimeKind.Utc)
                }
            };

            List<string> ordered = ColorConfigPathResolution.OrderTempScanCandidates(candidates);

            CollectionAssert.AreEqual(
                new[]
                {
                    @"C:\temp\b\ColorByRegexConfig.txt",
                    @"C:\temp\a\ColorByRegexConfig.txt",
                    @"C:\temp\c\ColorByRegexConfig.txt",
                    @"C:\temp\d\ColorByRegexConfig.txt"
                },
                ordered);
        }

        [TestMethod]
        public void BuildFanoutTargets_IncludesPrimaryAndTruncatesDistinctTargets()
        {
            List<string> targets = ColorConfigPathResolution.BuildFanoutTargets(
                primaryPath: @"C:\temp\primary\ColorByRegexConfig.txt",
                fallbackPaths: new[]
                {
                    @"C:\temp\primary\ColorByRegexConfig.txt",
                    @"C:\temp\one\ColorByRegexConfig.txt",
                    @"C:\temp\two\ColorByRegexConfig.txt",
                    @"C:\temp\three\ColorByRegexConfig.txt"
                },
                maxTargets: 3,
                isValidPath: _ => true);

            CollectionAssert.AreEqual(
                new[]
                {
                    @"C:\temp\primary\ColorByRegexConfig.txt",
                    @"C:\temp\one\ColorByRegexConfig.txt",
                    @"C:\temp\two\ColorByRegexConfig.txt"
                },
                targets);
        }

        [TestMethod]
        public void BuildWriteTargetPaths_ReturnsPrimaryForAuthoritativeSource()
        {
            List<string> targets = ColorConfigPathResolution.BuildWriteTargetPaths(
                source: "command-capture",
                primaryPath: @"C:\temp\primary\ColorByRegexConfig.txt",
                fallbackPaths: new[]
                {
                    @"C:\temp\primary\ColorByRegexConfig.txt",
                    @"C:\temp\other\ColorByRegexConfig.txt"
                },
                maxTargets: 8,
                isValidPath: _ => true);

            CollectionAssert.AreEqual(
                new[]
                {
                    @"C:\temp\primary\ColorByRegexConfig.txt"
                },
                targets);
        }

        [TestMethod]
        public void BuildWriteTargetPaths_ReturnsPrimaryForAuthoritativeSourceWithoutFallbacks()
        {
            List<string> targets = ColorConfigPathResolution.BuildWriteTargetPaths(
                source: "command-capture",
                primaryPath: @"C:\temp\primary\ColorByRegexConfig.txt",
                fallbackPaths: null,
                maxTargets: 8,
                isValidPath: _ => true);

            CollectionAssert.AreEqual(
                new[]
                {
                    @"C:\temp\primary\ColorByRegexConfig.txt"
                },
                targets);
        }

        [TestMethod]
        public void ShouldAcceptTempScanResults_FalseWhenAuthoritativePathAlreadyCaptured()
        {
            bool shouldAccept = ColorConfigPathResolution.ShouldAcceptTempScanResults(
                new ColorConfigPathResolution.PathState
                {
                    CommandCapturedPath = @"C:\temp\live\ColorByRegexConfig.txt",
                    ResolvedPath = @"C:\temp\live\ColorByRegexConfig.txt",
                    Source = "command-capture"
                });

            Assert.IsFalse(shouldAccept);
        }

        [TestMethod]
        public void ShouldAcceptTempScanResults_TrueBeforeAuthoritativeCapture()
        {
            bool shouldAccept = ColorConfigPathResolution.ShouldAcceptTempScanResults(
                new ColorConfigPathResolution.PathState
                {
                    ResolvedPath = @"C:\temp\guess\ColorByRegexConfig.txt",
                    Source = "temp-scan"
                });

            Assert.IsTrue(shouldAccept);
        }

        [TestMethod]
        public void PromoteObservedPath_ClearsFallbackAndMarksTempCandidate()
        {
            var result = ColorConfigPathResolution.PromoteObservedPath(
                new ColorConfigPathResolution.PathState
                {
                    ResolvedPath = @"C:\temp\primary\ColorByRegexConfig.txt",
                    Source = "temp-scan",
                    FallbackPaths = new List<string>
                    {
                        @"C:\temp\primary\ColorByRegexConfig.txt",
                        @"C:\temp\observed\ColorByRegexConfig.txt"
                    }
                },
                @"C:\temp\observed\ColorByRegexConfig.txt",
                "command-capture");

            Assert.AreEqual("temp-scan", result.PreviousSource);
            Assert.AreEqual(@"C:\temp\primary\ColorByRegexConfig.txt", result.PreviousPath);
            Assert.IsTrue(result.WasTempCandidate);
            Assert.AreEqual(2, result.PreviousTempCandidateCount);
            Assert.AreEqual("command-capture", result.NewState.Source);
            Assert.AreEqual(@"C:\temp\observed\ColorByRegexConfig.txt", result.NewState.ResolvedPath);
            Assert.IsNull(result.NewState.FallbackPaths);
        }

        [TestMethod]
        public void PromoteObservedPath_HandlesPathOutsideInitialTopNFanout()
        {
            var fallbackPaths = new List<string>
            {
                @"C:\temp\first\ColorByRegexConfig.txt",
                @"C:\temp\second\ColorByRegexConfig.txt",
                @"C:\temp\observed\ColorByRegexConfig.txt"
            };

            List<string> initialFanout = ColorConfigPathResolution.BuildFanoutTargets(
                primaryPath: fallbackPaths[0],
                fallbackPaths: fallbackPaths,
                maxTargets: 2,
                isValidPath: _ => true);

            CollectionAssert.DoesNotContain(initialFanout, @"C:\temp\observed\ColorByRegexConfig.txt");

            var result = ColorConfigPathResolution.PromoteObservedPath(
                new ColorConfigPathResolution.PathState
                {
                    ResolvedPath = fallbackPaths[0],
                    Source = "temp-scan",
                    FallbackPaths = fallbackPaths
                },
                @"C:\temp\observed\ColorByRegexConfig.txt",
                "command-capture");

            Assert.IsTrue(result.WasTempCandidate);
            Assert.AreEqual(@"C:\temp\observed\ColorByRegexConfig.txt", result.NewState.ResolvedPath);
            Assert.IsNull(result.NewState.FallbackPaths);
        }
    }
}