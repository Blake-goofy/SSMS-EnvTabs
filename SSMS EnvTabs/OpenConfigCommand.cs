using System;
using System.ComponentModel.Design;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio;
using Task = System.Threading.Tasks.Task;

namespace SSMS_EnvTabs
{
    internal sealed class OpenConfigCommand
    {
        internal const int TargetTabSettings = 0;
        internal const int TargetTabStyleTemplates = 1;
        internal const int TargetTabConnectionGroups = 2;

        public const int CommandId = 0x0103;
        public static readonly Guid CommandSet = SSMS_EnvTabsPackage.PackageCmdSetGuid;

        private readonly AsyncPackage package;

        private OpenConfigCommand(AsyncPackage package, OleMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            var menuCommandID = new CommandID(CommandSet, CommandId);
            var menuItem = new MenuCommand(this.Execute, menuCommandID);
            commandService.AddCommand(menuItem);
        }

        public static async Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            OleMenuCommandService commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as OleMenuCommandService;
            new OpenConfigCommand(package, commandService);
        }

        private void Execute(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try 
            {
                OpenSettingsWindow(TargetTabSettings, highlightUpdateChecks: false, forceReload: false, this.package);
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"OpenConfigCommand failed: {ex.Message}");
            }
        }

        internal static void OpenSettingsWindow(int targetTab, bool highlightUpdateChecks, bool forceReload, AsyncPackage package = null)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            AsyncPackage effectivePackage = package
                ?? SSMS_EnvTabsPackage.Instance;

            if (effectivePackage == null)
            {
                EnvTabsLog.Info("OpenSettingsWindow failed: package unavailable.");
                return;
            }

            ToolWindowPane window = effectivePackage.FindToolWindow(typeof(SettingsToolWindow), 0, true);
            if (window?.Frame is IVsWindowFrame windowFrame)
            {
                windowFrame.SetProperty((int)__VSFPROPID.VSFPROPID_FrameMode, VSFRAMEMODE.VSFM_MdiChild);
                ErrorHandler.ThrowOnFailure(windowFrame.Show());
            }

            SettingsToolWindowControl control = window?.Content as SettingsToolWindowControl;
            control?.NavigateTo(targetTab, highlightUpdateChecks, forceReload);
        }
    }
}
