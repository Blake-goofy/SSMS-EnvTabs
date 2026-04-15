using System;
using System.IO;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace SSMS_EnvTabs.Tests
{
    [TestClass]
    public class VersionConsistencyTests
    {
        [TestMethod]
        public void AssemblyVersion_MatchesVsixManifestVersion()
        {
            string repoRoot = GetRepositoryRoot();
            string assemblyInfoPath = Path.Combine(repoRoot, "SSMS EnvTabs", "Properties", "AssemblyInfo.cs");
            string manifestPath = Path.Combine(repoRoot, "SSMS EnvTabs", "source.extension.vsixmanifest");

            string assemblyInfoText = File.ReadAllText(assemblyInfoPath);
            string assemblyVersion = GetAssemblyVersion(assemblyInfoText);
            string manifestVersion = GetManifestVersion(manifestPath);

            Assert.AreEqual(manifestVersion, assemblyVersion,
                $"AssemblyVersion in '{assemblyInfoPath}' must match VSIX manifest version in '{manifestPath}'.");
        }

        private static string GetRepositoryRoot()
        {
            string currentDirectory = AppDomain.CurrentDomain.BaseDirectory;

            while (!string.IsNullOrWhiteSpace(currentDirectory))
            {
                string solutionPath = Path.Combine(currentDirectory, "SSMS EnvTabs.sln");
                if (File.Exists(solutionPath))
                {
                    return currentDirectory;
                }

                DirectoryInfo parent = Directory.GetParent(currentDirectory);
                currentDirectory = parent?.FullName;
            }

            Assert.Fail("Could not locate repository root from test base directory.");
            return null;
        }

        private static string GetAssemblyVersion(string assemblyInfoText)
        {
            Match match = Regex.Match(
                assemblyInfoText,
                "^\\s*\\[assembly:\\s*AssemblyVersion\\(\\\"(?<version>[^\\\"]+)\\\"\\)\\]",
                RegexOptions.Multiline);

            Assert.IsTrue(match.Success, "AssemblyVersion attribute not found in AssemblyInfo.cs.");
            return match.Groups["version"].Value;
        }

        private static string GetManifestVersion(string manifestPath)
        {
            XDocument document = XDocument.Load(manifestPath);
            XNamespace ns = "http://schemas.microsoft.com/developer/vsx-schema/2011";
            XElement metadataElement = document.Root?.Element(ns + "Metadata");
            XElement identityElement = metadataElement?.Element(ns + "Identity");
            string version = identityElement?.Attribute("Version")?.Value;

            Assert.IsFalse(string.IsNullOrWhiteSpace(version), "VSIX manifest Identity Version attribute not found.");
            return version;
        }
    }
}