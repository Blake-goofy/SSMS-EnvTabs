using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Reflection;
using System.Runtime.Serialization;
using System.Runtime.Serialization.Json;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace SSMS_EnvTabs
{
    internal static class UpdateChecker
    {
        private const string ReleasesApiUrl = "https://api.github.com/repos/Blake-goofy/SSMS-EnvTabs/releases/latest";
        private const string VsixId = "SSMS_EnvTabs";
        private const string LegacyVsixId = "SSMS_EnvTabs.20d4f774-2a12-403b-a25d-1ce263e878d7";
        private static int pendingStartupCheck;
        private static readonly object diagnosticsLock = new object();
        private static string lastUpdateResult = "No update check has run yet.";
        internal static event Action LastUpdateResultChanged;

        internal static string LastUpdateResult
        {
            get
            {
                lock (diagnosticsLock)
                {
                    return lastUpdateResult;
                }
            }
        }

        private static void SetLastUpdateResult(string message)
        {
            if (string.IsNullOrWhiteSpace(message))
            {
                return;
            }

            string stamped = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}";
            lock (diagnosticsLock)
            {
                lastUpdateResult = stamped;
            }

            try
            {
                LastUpdateResultChanged?.Invoke();
            }
            catch
            {
                // Best-effort diagnostics notification.
            }
        }

        public static void ScheduleCheck(AsyncPackage package, TabGroupSettings settings)
        {
            if (settings?.EnableUpdateChecks == false)
            {
                EnvTabsLog.Info("Update check skipped: disabled by settings.");
                SetLastUpdateResult("Update check skipped because it is disabled in settings.");
                return;
            }

            Interlocked.Exchange(ref pendingStartupCheck, 1);
            EnvTabsLog.Info("Update check deferred until first valid connected document is opened.");
            SetLastUpdateResult("Startup update check deferred until first connected query tab.");
        }

        internal static void NotifyConnectedDocument(AsyncPackage package, TabGroupSettings settings, string server, string database)
        {
            if (package == null || string.IsNullOrWhiteSpace(server))
            {
                return;
            }

            if (settings?.EnableUpdateChecks == false)
            {
                Interlocked.Exchange(ref pendingStartupCheck, 0);
                SetLastUpdateResult("Startup update check canceled because update checks are disabled.");
                return;
            }

            if (Interlocked.CompareExchange(ref pendingStartupCheck, 0, 1) != 1)
            {
                return;
            }

            _ = package.JoinableTaskFactory.RunAsync(async () =>
            {
                try
                {
                    var token = package.DisposalToken;
                    await Task.Delay(TimeSpan.FromMilliseconds(900), token);
                    EnvTabsLog.Info($"Update check resumed after first connected document. Server='{server}', Database='{database}'.");
                    SetLastUpdateResult("Running deferred startup update check.");
                    await CheckForUpdatesAsync(package, token, showUpToDate: false);
                }
                catch (OperationCanceledException)
                {
                    // Ignored
                }
                catch (Exception ex)
                {
                    EnvTabsLog.Info($"Deferred update check failed: {ex.Message}");
                    SetLastUpdateResult($"Deferred update check failed: {ex.Message}");
                }
            });
        }

        public static void CheckNow(AsyncPackage package, TabGroupSettings settings, bool ignoreSettings = true)
        {
            if (!ignoreSettings && settings?.EnableUpdateChecks == false)
            {
                EnvTabsLog.Info("Manual update check skipped: disabled by settings.");
                SetLastUpdateResult("Manual update check skipped because it is disabled in settings.");
                return;
            }

            _ = package.JoinableTaskFactory.RunAsync(async () =>
            {
                try
                {
                    EnvTabsLog.Info("Manual update check started.");
                    SetLastUpdateResult("Manual update check started.");
                    await CheckForUpdatesAsync(package, package.DisposalToken, showUpToDate: true);
                }
                catch (OperationCanceledException)
                {
                    // Ignored
                }
                catch (Exception ex)
                {
                    EnvTabsLog.Info($"Manual update check failed: {ex.Message}");
                    SetLastUpdateResult($"Manual update check failed: {ex.Message}");
                }
            });
        }

        private static async Task CheckForUpdatesAsync(AsyncPackage package, CancellationToken token, bool showUpToDate)
        {
            var currentVersion = GetCurrentVersion();
            if (currentVersion == null)
            {
                EnvTabsLog.Info("Update check failed: current version unavailable.");
                SetLastUpdateResult("Update check failed because current version could not be determined.");
                return;
            }

            var release = await GetLatestReleaseAsync(token);
            if (release == null || release.Draft)
            {
                EnvTabsLog.Info("Update check failed: release info unavailable.");
                SetLastUpdateResult("Update check failed because latest release info was unavailable.");
                return;
            }

            var latestVersion = ParseVersion(release.TagName);
            if (latestVersion == null)
            {
                EnvTabsLog.Info("Update check failed: latest version parse failed.");
                SetLastUpdateResult("Update check failed because latest release version could not be parsed.");
                return;
            }

            if (latestVersion <= currentVersion)
            {
                EnvTabsLog.Info($"Update check: already on latest ({currentVersion}).");
                SetLastUpdateResult($"Up to date ({FormatVersion(currentVersion)}).");
                if (showUpToDate)
                {
                    await package.JoinableTaskFactory.SwitchToMainThreadAsync(token);
                    ShowUpToDatePrompt(package, currentVersion);
                }
                return;
            }

            await package.JoinableTaskFactory.SwitchToMainThreadAsync(token);

            EnvTabsLog.Info($"Update available: {latestVersion} (current {currentVersion}).");
            SetLastUpdateResult($"Update available: {FormatVersion(currentVersion)} -> {FormatVersion(latestVersion)}.");
            ShowUpdatePrompt(package, release, latestVersion, currentVersion);
        }

        private static async Task<GitHubRelease> GetLatestReleaseAsync(CancellationToken token)
        {
            try
            {
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.UserAgent.ParseAdd("SSMS-EnvTabs");
                    client.DefaultRequestHeaders.Accept.ParseAdd("application/vnd.github+json");

                    using (var response = await client.GetAsync(ReleasesApiUrl, token))
                    {
                        if (!response.IsSuccessStatusCode)
                        {
                            EnvTabsLog.Info($"Update check HTTP {response.StatusCode}");
                            return null;
                        }

                        using (var stream = await response.Content.ReadAsStreamAsync())
                        {
                            var serializer = new DataContractJsonSerializer(typeof(GitHubRelease));
                            return serializer.ReadObject(stream) as GitHubRelease;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"Update check fetch failed: {ex.Message}");
                return null;
            }
        }

        internal static Version GetCurrentVersion()
        {
            try
            {
                var assemblyVersion = Assembly.GetExecutingAssembly().GetName().Version;
                if (assemblyVersion != null)
                {
                    EnvTabsLog.Info($"Update check current version (assembly): {assemblyVersion}");
                    EnvTabsLog.Info("Update check: manifest version is ignored by design.");
                    return assemblyVersion;
                }
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"Update check version parse failed: {ex.Message}");
            }

            EnvTabsLog.Info("Update check current version not found in assembly. Defaulting to 0.0.0.");
            return new Version(0, 0, 0, 0);
        }


        private static Version ParseVersion(string text)
        {
            if (string.IsNullOrWhiteSpace(text))
            {
                return null;
            }

            string value = text.Trim();
            if (value.StartsWith("v", StringComparison.OrdinalIgnoreCase))
            {
                value = value.Substring(1);
            }

            if (Version.TryParse(value, out Version parsed))
            {
                return parsed;
            }

            return null;
        }

        private static void ShowUpdatePrompt(AsyncPackage package, GitHubRelease release, Version latestVersion, Version currentVersion)
        {
            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                string latestDisplay = FormatVersion(latestVersion);
                string currentDisplay = FormatVersion(currentVersion);
                using (var dialog = new UpdatePromptDialog(
                    latestDisplay,
                    currentDisplay,
                    () => OpenUrl(release?.HtmlUrl),
                    () => OpenConfig(package)))
                {
                    var result = dialog.ShowDialog();
                    if (result == System.Windows.Forms.DialogResult.Yes)
                    {
                        InstallRelease(release, currentVersion);
                    }
                }
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"Update prompt failed: {ex.Message}");
            }
        }

        private static void ShowUpToDatePrompt(AsyncPackage package, Version currentVersion)
        {
            try
            {
                ThreadHelper.ThrowIfNotOnUIThread();
                VsShellUtilities.ShowMessageBox(
                    package,
                    $"SSMS EnvTabs is up to date ({FormatVersion(currentVersion)}).",
                    "Information",
                    OLEMSGICON.OLEMSGICON_INFO,
                    OLEMSGBUTTON.OLEMSGBUTTON_OK,
                    OLEMSGDEFBUTTON.OLEMSGDEFBUTTON_FIRST);
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"Up-to-date prompt failed: {ex.Message}");
            }
        }

        private static void OpenConfig(AsyncPackage package)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                OpenConfigCommand.OpenSettingsWindow(
                    OpenConfigCommand.TargetTabSettings,
                    highlightUpdateChecks: true,
                    forceReload: false,
                    package);
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"OpenConfig from update prompt failed: {ex.Message}");
            }
        }


        private static void InstallRelease(GitHubRelease release, Version currentVersion)
        {
            if (release == null)
            {
                return;
            }

            GitHubAsset vsixAsset = GetVsixAsset(release);
            if (vsixAsset == null || string.IsNullOrWhiteSpace(vsixAsset.DownloadUrl))
            {
                OpenUrl(release.HtmlUrl);
                return;
            }

            _ = Task.Run(async () =>
            {
                try
                {
                    SetLastUpdateResult("Preparing update package download.");
                    string expectedSha256 = await GetExpectedSha256Async(release, vsixAsset);
                    if (string.IsNullOrWhiteSpace(expectedSha256))
                    {
                        EnvTabsLog.Info("Update install aborted: .sha256 release asset missing or invalid.");
                        SetLastUpdateResult("Update install aborted: missing or invalid .sha256 release asset.");
                        OpenUrl(release.HtmlUrl);
                        return;
                    }

                    string tempPath = Path.Combine(Path.GetTempPath(), $"SSMS-EnvTabs-{Guid.NewGuid():N}.vsix");
                    await DownloadFileAsync(vsixAsset.DownloadUrl, tempPath);

                    string actualSha256 = ComputeSha256Hex(tempPath);
                    if (!string.Equals(actualSha256, expectedSha256, StringComparison.OrdinalIgnoreCase))
                    {
                        EnvTabsLog.Info($"Update install aborted: checksum mismatch. Expected={expectedSha256}, Actual={actualSha256}");
                        SetLastUpdateResult("Update install aborted: checksum verification failed.");
                        OpenUrl(release.HtmlUrl);
                        return;
                    }

                    EnvTabsLog.Info("Update package checksum verified successfully.");
                    SetLastUpdateResult("Update package verified. Launching installer.");

                    if (!LaunchUpdateScript(tempPath, release, currentVersion))
                    {
                        SetLastUpdateResult("Could not launch installer automatically; opened release page instead.");
                        OpenUrl(release.HtmlUrl);
                    }
                }
                catch (Exception ex)
                {
                    EnvTabsLog.Info($"Update install failed: {ex.Message}");
                    SetLastUpdateResult($"Update install failed: {ex.Message}");
                    OpenUrl(release.HtmlUrl);
                }
            });
        }

        private static GitHubAsset GetVsixAsset(GitHubRelease release)
        {
            if (release?.Assets == null)
            {
                return null;
            }

            foreach (var asset in release.Assets)
            {
                if (!string.IsNullOrWhiteSpace(asset?.Name) && asset.Name.EndsWith(".vsix", StringComparison.OrdinalIgnoreCase))
                {
                    return asset;
                }
            }

            return null;
        }

        private static GitHubAsset GetChecksumAsset(GitHubRelease release, string vsixAssetName)
        {
            if (release?.Assets == null)
            {
                return null;
            }

            string expectedName = string.IsNullOrWhiteSpace(vsixAssetName) ? null : vsixAssetName + ".sha256";
            if (!string.IsNullOrWhiteSpace(expectedName))
            {
                foreach (var asset in release.Assets)
                {
                    if (asset == null || string.IsNullOrWhiteSpace(asset.Name))
                    {
                        continue;
                    }

                    if (string.Equals(asset.Name, expectedName, StringComparison.OrdinalIgnoreCase))
                    {
                        return asset;
                    }
                }
            }

            foreach (var asset in release.Assets)
            {
                if (asset == null || string.IsNullOrWhiteSpace(asset.Name))
                {
                    continue;
                }

                if (asset.Name.EndsWith(".sha256", StringComparison.OrdinalIgnoreCase))
                {
                    return asset;
                }
            }

            return null;
        }

        private static async Task<string> GetExpectedSha256Async(GitHubRelease release, GitHubAsset vsixAsset)
        {
            var checksumAsset = GetChecksumAsset(release, vsixAsset?.Name);
            if (checksumAsset == null || string.IsNullOrWhiteSpace(checksumAsset.DownloadUrl))
            {
                return null;
            }

            try
            {
                using (var client = new HttpClient())
                {
                    client.DefaultRequestHeaders.UserAgent.ParseAdd("SSMS-EnvTabs");
                    string checksumText = await client.GetStringAsync(checksumAsset.DownloadUrl);
                    return ParseSha256FromText(checksumText, vsixAsset?.Name);
                }
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"Checksum download failed: {ex.Message}");
                return null;
            }
        }

        private static string ParseSha256FromText(string checksumText, string vsixAssetName)
        {
            if (string.IsNullOrWhiteSpace(checksumText))
            {
                return null;
            }

            string[] lines = checksumText
                .Replace("\r\n", "\n")
                .Replace("\r", "\n")
                .Split('\n');

            if (!string.IsNullOrWhiteSpace(vsixAssetName))
            {
                foreach (string rawLine in lines)
                {
                    string line = rawLine?.Trim();
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }

                    // Common format: <hash>  <fileName>
                    var match = Regex.Match(line, @"^([A-Fa-f0-9]{64})\s+\*?(.+)$");
                    if (match.Success)
                    {
                        string fileToken = match.Groups[2].Value.Trim();
                        if (fileToken.Equals(vsixAssetName, StringComparison.OrdinalIgnoreCase))
                        {
                            return match.Groups[1].Value.ToLowerInvariant();
                        }
                    }

                    // Alternate format: SHA256(fileName)= <hash>
                    var alt = Regex.Match(line, @"^SHA256\((.+)\)\s*=\s*([A-Fa-f0-9]{64})$", RegexOptions.IgnoreCase);
                    if (alt.Success)
                    {
                        string fileToken = alt.Groups[1].Value.Trim();
                        if (fileToken.Equals(vsixAssetName, StringComparison.OrdinalIgnoreCase))
                        {
                            return alt.Groups[2].Value.ToLowerInvariant();
                        }
                    }
                }
            }

            // Fallback: accept the first 64-hex token from the file.
            var firstHash = Regex.Match(checksumText, @"\b([A-Fa-f0-9]{64})\b");
            if (firstHash.Success)
            {
                return firstHash.Groups[1].Value.ToLowerInvariant();
            }

            return null;
        }

        private static string ComputeSha256Hex(string filePath)
        {
            using (var stream = File.OpenRead(filePath))
            using (var sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(stream);
                var builder = new StringBuilder(hash.Length * 2);
                foreach (byte b in hash)
                {
                    builder.Append(b.ToString("x2"));
                }

                return builder.ToString();
            }
        }

        private static async Task DownloadFileAsync(string url, string path)
        {
            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.UserAgent.ParseAdd("SSMS-EnvTabs");
                var bytes = await client.GetByteArrayAsync(url);
                File.WriteAllBytes(path, bytes);
            }
        }

        private static bool LaunchUpdateScript(string vsixPath, GitHubRelease release, Version currentVersion)
        {
            try
            {
                string installerPath = GetVsixInstallerPath();
                if (string.IsNullOrWhiteSpace(installerPath) || !File.Exists(installerPath))
                {
                    EnvTabsLog.Info("VSIXInstaller.exe not found; falling back to browser.");
                    return false;
                }

                string scriptPath = Path.Combine(Path.GetTempPath(), $"SSMS-EnvTabs-Update-{Guid.NewGuid():N}.bat");
                string releaseUrl = release?.HtmlUrl ?? string.Empty;
                string releaseTag = release?.TagName ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(releaseTag) && releaseTag.StartsWith("v", StringComparison.OrdinalIgnoreCase))
                {
                    releaseTag = releaseTag.Substring(1);
                }
                var releaseParsed = ParseVersion(releaseTag);
                if (releaseParsed != null)
                {
                    releaseTag = FormatVersion(releaseParsed);
                }
                if (string.IsNullOrWhiteSpace(releaseTag))
                {
                    releaseTag = "(new version)";
                }
                string currentTag = FormatVersion(currentVersion);
                string uninstallLabel = $"Uninstalling v{currentTag}...";
                string uninstallDone = $"Uninstalling v{currentTag}... done!";
                string installLabel = $"Installing v{releaseTag}...";
                string installDone = $"Installing v{releaseTag}... done!";

                string uninstallPs =
                    "$spinner='|/-\\';$i=0;[Console]::CursorVisible=$false;" +
                    "Write-Host \"" + uninstallLabel + "\";" +
                    "$line=[Console]::CursorTop-1;" +
                    "while(Get-Process VSIXInstaller -ErrorAction SilentlyContinue){" +
                    "$ch=$spinner[$i++%4];" +
                    "[Console]::SetCursorPosition(0,$line);" +
                    "Write-Host -NoNewline (\"" + uninstallLabel + " \"+$ch+\"   \" );" +
                    "[Console]::SetCursorPosition(0,$line+1);" +
                    "Start-Sleep -Milliseconds 250};" +
                    "[Console]::SetCursorPosition(0,$line);" +
                    "Write-Host \"" + uninstallDone + "   \";" +
                    "[Console]::SetCursorPosition(0,$line+1);" +
                    "[Console]::CursorVisible=$true";

                string installPs =
                    "$spinner='|/-\\';$i=0;[Console]::CursorVisible=$false;" +
                    "Write-Host \"" + installLabel + "\";" +
                    "$line=[Console]::CursorTop-1;" +
                    "while(Get-Process VSIXInstaller -ErrorAction SilentlyContinue){" +
                    "$ch=$spinner[$i++%4];" +
                    "[Console]::SetCursorPosition(0,$line);" +
                    "Write-Host -NoNewline (\"" + installLabel + " \"+$ch+\"   \" );" +
                    "[Console]::SetCursorPosition(0,$line+1);" +
                    "Start-Sleep -Milliseconds 250};" +
                    "[Console]::SetCursorPosition(0,$line);" +
                    "Write-Host \"" + installDone + "   \";" +
                    "[Console]::SetCursorPosition(0,$line+1);" +
                    "[Console]::CursorVisible=$true";

                string uninstallPsEncoded = Convert.ToBase64String(Encoding.Unicode.GetBytes(uninstallPs));
                string installPsEncoded = Convert.ToBase64String(Encoding.Unicode.GetBytes(installPs));
                string script =
                    "@echo off" + Environment.NewLine +
                    "setlocal EnableExtensions EnableDelayedExpansion" + Environment.NewLine +
                    "set installer=\"" + installerPath + "\"" + Environment.NewLine +
                    "set vsix=\"" + vsixPath + "\"" + Environment.NewLine +
                    "if not exist %installer% goto fallback" + Environment.NewLine +
                    "echo ===Updating SSMS EnvTabs===" + Environment.NewLine +
                    "echo This will take around 1 minute..." + Environment.NewLine +
                    "taskkill /IM Ssms.exe /F >nul 2>&1" + Environment.NewLine +
                    "timeout /t 2 /nobreak >nul" + Environment.NewLine +
                    "echo." + Environment.NewLine +
                    "start \"\" %installer% /quiet /uninstall:" + VsixId + Environment.NewLine +
                    "powershell -NoProfile -EncodedCommand " + uninstallPsEncoded + Environment.NewLine +
                    "start \"\" %installer% /quiet /uninstall:" + LegacyVsixId + Environment.NewLine +
                    "powershell -NoProfile -EncodedCommand " + uninstallPsEncoded + Environment.NewLine +
                    "start \"\" %installer% /quiet %vsix%" + Environment.NewLine +
                    "powershell -NoProfile -EncodedCommand " + installPsEncoded + Environment.NewLine +
                    "echo." + Environment.NewLine +
                    "echo Update complete!" + Environment.NewLine +
                    "echo You can launch SSMS now" + Environment.NewLine +
                    "pause" + Environment.NewLine +
                    "goto done" + Environment.NewLine +
                    ":fallback" + Environment.NewLine +
                    (string.IsNullOrWhiteSpace(releaseUrl) ? string.Empty : "start \"\" \"" + releaseUrl + "\"" + Environment.NewLine) +
                    ":done" + Environment.NewLine +
                    "endlocal" + Environment.NewLine +
                    "del \"%~f0\"" + Environment.NewLine;

                File.WriteAllText(scriptPath, script, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));

                var info = new ProcessStartInfo
                {
                    FileName = "cmd.exe",
                    Arguments = $"/c \"\"{scriptPath}\"\"",
                    UseShellExecute = true,
                    CreateNoWindow = false
                };

                Process.Start(info);
                return true;
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"Launch update script failed: {ex.Message}");
                return false;
            }
        }

        private static string GetVsixInstallerPath()
        {
            try
            {
                string exePath = Process.GetCurrentProcess()?.MainModule?.FileName;
                if (!string.IsNullOrWhiteSpace(exePath))
                {
                    string exeDir = Path.GetDirectoryName(exePath);
                    string candidate = Path.Combine(exeDir ?? string.Empty, "VSIXInstaller.exe");
                    if (File.Exists(candidate))
                    {
                        return candidate;
                    }
                }
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"Resolve VSIXInstaller from process failed: {ex.Message}");
            }

            try
            {
                string baseDir = AppDomain.CurrentDomain.BaseDirectory;
                string candidate = Path.Combine(baseDir ?? string.Empty, "VSIXInstaller.exe");
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"Resolve VSIXInstaller from AppDomain failed: {ex.Message}");
            }

            return null;
        }

        internal static string FormatVersion(Version version)
        {
            if (version == null)
            {
                return "0.0.0";
            }

            if (version.Revision > 0)
            {
                return version.ToString(4);
            }

            if (version.Build > 0)
            {
                return version.ToString(3);
            }

            return version.ToString(2);
        }

        private static void OpenUrl(string url)
        {
            if (string.IsNullOrWhiteSpace(url))
            {
                return;
            }

            try
            {
                Process.Start(new ProcessStartInfo
                {
                    FileName = url,
                    UseShellExecute = true
                });
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"Open URL failed: {ex.Message}");
            }
        }

        [DataContract]
        private sealed class GitHubRelease
        {
            [DataMember(Name = "tag_name")]
            public string TagName { get; set; }

            [DataMember(Name = "html_url")]
            public string HtmlUrl { get; set; }

            [DataMember(Name = "draft")]
            public bool Draft { get; set; }

            [DataMember(Name = "prerelease")]
            public bool Prerelease { get; set; }

            [DataMember(Name = "assets")]
            public List<GitHubAsset> Assets { get; set; }
        }

        [DataContract]
        private sealed class GitHubAsset
        {
            [DataMember(Name = "name")]
            public string Name { get; set; }

            [DataMember(Name = "browser_download_url")]
            public string DownloadUrl { get; set; }
        }

    }
}
