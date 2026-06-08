using System;
using System.Diagnostics;
using System.IO;
using System.Text;
using EnvDTE;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace SSMS_EnvTabs
{
    internal static class UninstallCoordinator
    {
        private const string VsixInstallerPath = @"C:\Program Files\Microsoft SQL Server Management Studio 22\Release\Common7\IDE\VSIXInstaller.exe";
        private const string VsixUninstallArguments = "/u:SSMS_EnvTabs";
        private const int UninstallLaunchTimeoutSeconds = 120;

        internal static bool IsVsixInstallerAvailable()
        {
            return File.Exists(VsixInstallerPath);
        }

        internal static bool RequestNormalShutdown()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                var dte = Package.GetGlobalService(typeof(SDTE)) as DTE;
                if (dte == null)
                {
                    EnvTabsLog.Info("Deferred uninstall shutdown failed: DTE service unavailable.");
                    return false;
                }

                int currentProcessId = System.Diagnostics.Process.GetCurrentProcess().Id;
                if (!LaunchUninstallHelper(currentProcessId))
                {
                    EnvTabsLog.Info("Deferred uninstall shutdown failed: helper process could not be started.");
                    return false;
                }

                dte.Quit();
                EnvTabsLog.Info("Requested normal SSMS shutdown for uninstall.");
                return true;
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"Deferred uninstall shutdown failed: {ex.Message}");
                return false;
            }
        }

        private static bool LaunchUninstallHelper(int processId)
        {
            try
            {
                string escapedInstallerPath = EscapePowerShellSingleQuotedString(VsixInstallerPath);
                string escapedArguments = EscapePowerShellSingleQuotedString(VsixUninstallArguments);
                string helperScript =
                    "$ErrorActionPreference = 'SilentlyContinue'; " +
                    "$deadline = (Get-Date).AddSeconds(" + UninstallLaunchTimeoutSeconds + "); " +
                    "while ((Get-Date) -lt $deadline) { " +
                    "if (-not (Get-Process -Id " + processId + " -ErrorAction SilentlyContinue)) { " +
                    "Start-Process -FilePath '" + escapedInstallerPath + "' -ArgumentList '" + escapedArguments + "'; " +
                    "exit 0; " +
                    "} " +
                    "Start-Sleep -Milliseconds 500; " +
                    "} " +
                    "exit 1;";

                string encodedScript = Convert.ToBase64String(Encoding.Unicode.GetBytes(helperScript));

                System.Diagnostics.Process.Start(new ProcessStartInfo
                {
                    FileName = "powershell.exe",
                    Arguments = "-NoProfile -WindowStyle Hidden -ExecutionPolicy Bypass -EncodedCommand " + encodedScript,
                    UseShellExecute = true,
                    CreateNoWindow = true,
                    WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden
                });

                EnvTabsLog.Info($"Started uninstall helper for SSMS process {processId}.");
                return true;
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"Start uninstall helper failed: {ex.Message}");
                return false;
            }
        }

        private static string EscapePowerShellSingleQuotedString(string value)
        {
            return string.IsNullOrEmpty(value)
                ? string.Empty
                : value.Replace("'", "''");
        }
    }
}
