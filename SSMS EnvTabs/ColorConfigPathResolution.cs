using System;
using System.Collections.Generic;
using System.Linq;

namespace SSMS_EnvTabs
{
    internal static class ColorConfigPathResolution
    {
        internal sealed class TempScanCandidate
        {
            public string Path { get; set; }
            public string GuidName { get; set; }
            public DateTime LastWriteUtc { get; set; }
            public bool InPreferredWindow { get; set; }
            public double AbsDeltaSeconds { get; set; }
            public bool HasLegacyVersionFolder { get; set; }
        }

        internal sealed class PathState
        {
            public string ResolvedPath { get; set; }
            public string Source { get; set; }
            public List<string> FallbackPaths { get; set; }
            public string CommandCapturedPath { get; set; }
        }

        internal sealed class PromotionResult
        {
            public string PreviousPath { get; set; }
            public string PreviousSource { get; set; }
            public bool WasTempCandidate { get; set; }
            public int PreviousTempCandidateCount { get; set; }
            public PathState NewState { get; set; }
        }

        internal static bool ShouldAcceptTempScanResults(PathState state)
        {
            if (state == null)
            {
                return true;
            }

            if (!string.IsNullOrWhiteSpace(state.CommandCapturedPath))
            {
                return false;
            }

            return !string.Equals(state.Source, "command-capture", StringComparison.OrdinalIgnoreCase);
        }

        internal static List<string> OrderTempScanCandidates(IEnumerable<TempScanCandidate> candidates)
        {
            return (candidates ?? Enumerable.Empty<TempScanCandidate>())
                .Where(candidate => candidate != null && !string.IsNullOrWhiteSpace(candidate.Path))
                .OrderByDescending(candidate => candidate.InPreferredWindow)
                .ThenBy(candidate => candidate.AbsDeltaSeconds)
                .ThenByDescending(candidate => candidate.HasLegacyVersionFolder)
                .ThenByDescending(candidate => candidate.LastWriteUtc)
                .ThenBy(candidate => candidate.GuidName, StringComparer.OrdinalIgnoreCase)
                .Select(candidate => candidate.Path)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        internal static List<string> BuildFanoutTargets(string primaryPath, IEnumerable<string> fallbackPaths, int maxTargets, Func<string, bool> isValidPath)
        {
            var result = new List<string>();
            if (isValidPath == null)
            {
                throw new ArgumentNullException(nameof(isValidPath));
            }

            if (!string.IsNullOrWhiteSpace(primaryPath) && isValidPath(primaryPath))
            {
                result.Add(primaryPath);
            }

            if (fallbackPaths == null || maxTargets <= 0)
            {
                return result;
            }

            foreach (string candidate in fallbackPaths)
            {
                if (result.Count >= maxTargets)
                {
                    break;
                }

                if (string.IsNullOrWhiteSpace(candidate) || !isValidPath(candidate))
                {
                    continue;
                }

                if (!result.Contains(candidate, StringComparer.OrdinalIgnoreCase))
                {
                    result.Add(candidate);
                }
            }

            return result;
        }

        internal static List<string> BuildWriteTargetPaths(string source, string primaryPath, IEnumerable<string> fallbackPaths, int maxTargets, Func<string, bool> isValidPath)
        {
            if (isValidPath == null)
            {
                throw new ArgumentNullException(nameof(isValidPath));
            }

            if (string.IsNullOrWhiteSpace(primaryPath) || !isValidPath(primaryPath))
            {
                return new List<string>();
            }

            bool shouldFanout = string.Equals(source, "temp-scan", StringComparison.OrdinalIgnoreCase)
                && fallbackPaths != null
                && fallbackPaths.Skip(1).Any();

            if (!shouldFanout)
            {
                return new List<string> { primaryPath };
            }

            return BuildFanoutTargets(primaryPath, fallbackPaths, maxTargets, isValidPath);
        }

        internal static PromotionResult PromoteObservedPath(PathState state, string observedPath, string promotedSource)
        {
            if (string.IsNullOrWhiteSpace(observedPath))
            {
                throw new ArgumentException("Observed path is required.", nameof(observedPath));
            }

            var previousFallbackPaths = (state?.FallbackPaths ?? new List<string>())
                .Where(path => !string.IsNullOrWhiteSpace(path))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            return new PromotionResult
            {
                PreviousPath = state?.ResolvedPath,
                PreviousSource = state?.Source,
                WasTempCandidate = string.Equals(state?.Source, "temp-scan", StringComparison.OrdinalIgnoreCase)
                    && previousFallbackPaths.Contains(observedPath, StringComparer.OrdinalIgnoreCase),
                PreviousTempCandidateCount = previousFallbackPaths.Count,
                NewState = new PathState
                {
                    CommandCapturedPath = observedPath,
                    ResolvedPath = observedPath,
                    Source = promotedSource,
                    FallbackPaths = null
                }
            };
        }
    }
}