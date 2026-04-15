using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Imaging;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;

namespace SSMS_EnvTabs
{
    internal sealed class UpdateInfoBar : IVsInfoBarUIEvents
    {
        private const string ActionUpdateNow = "update_now";
        private const string ActionUpdateOnClose = "update_on_close";
        private const string ActionReleaseNotes = "release_notes";

        private readonly AsyncPackage package;
        private readonly string latestVersion;
        private readonly string releaseUrl;
        private readonly Action<string> onUpdateRequested;

        private IVsInfoBarUIElement currentElement;
        private uint adviseCookie;

        internal UpdateInfoBar(AsyncPackage package, string latestVersion, string releaseUrl, Action<string> onUpdateRequested)
        {
            this.package = package;
            this.latestVersion = latestVersion;
            this.releaseUrl = releaseUrl;
            this.onUpdateRequested = onUpdateRequested;
        }

        internal bool TryShow()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                var shell = package.GetService<SVsShell, IVsShell>();
                if (shell == null)
                {
                    EnvTabsLog.Info("UpdateInfoBar: IVsShell service unavailable.");
                    return false;
                }

                if (ErrorHandler.Failed(shell.GetProperty((int)__VSSPROPID7.VSSPROPID_MainWindowInfoBarHost, out object hostObj))
                    || !(hostObj is IVsInfoBarHost host))
                {
                    EnvTabsLog.Info("UpdateInfoBar: MainWindowInfoBarHost not available.");
                    return false;
                }

                var factory = package.GetService<SVsInfoBarUIFactory, IVsInfoBarUIFactory>();
                if (factory == null)
                {
                    EnvTabsLog.Info("UpdateInfoBar: IVsInfoBarUIFactory service unavailable.");
                    return false;
                }

                var model = new InfoBarModel(
                    new IVsInfoBarTextSpan[]
                    {
                        new InfoBarTextSpan($"EnvTabs v{latestVersion} update available.  ")
                    },
                    new IVsInfoBarActionItem[]
                    {
                        new InfoBarButton("Update Now", ActionUpdateNow),
                        new InfoBarButton("Update on Close", ActionUpdateOnClose),
                        new InfoBarHyperlink("Release Notes", ActionReleaseNotes)
                    },
                    KnownMonikers.StatusInformation,
                    isCloseButtonVisible: true);

                currentElement = factory.CreateInfoBar(model);
                if (currentElement == null)
                {
                    EnvTabsLog.Info("UpdateInfoBar: CreateInfoBar returned null.");
                    return false;
                }

                currentElement.Advise(this, out adviseCookie);
                host.AddInfoBar(currentElement);
                EnvTabsLog.Info("UpdateInfoBar: Shown successfully.");
                return true;
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"UpdateInfoBar: Failed to show: {ex.Message}");
                return false;
            }
        }

        internal void Dismiss()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                if (currentElement != null)
                {
                    currentElement.Unadvise(adviseCookie);
                    currentElement.Close();
                    currentElement = null;
                }
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"UpdateInfoBar: Dismiss failed: {ex.Message}");
                currentElement = null;
            }
        }

        public void OnActionItemClicked(IVsInfoBarUIElement infoBarUIElement, IVsInfoBarActionItem actionItem)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            string context = actionItem.ActionContext as string;
            switch (context)
            {
                case ActionUpdateNow:
                    EnvTabsLog.Info("UpdateInfoBar: User clicked Update Now.");
                    Dismiss();
                    onUpdateRequested?.Invoke(ActionUpdateNow);
                    break;

                case ActionUpdateOnClose:
                    EnvTabsLog.Info("UpdateInfoBar: User clicked Update on Close.");
                    Dismiss();
                    onUpdateRequested?.Invoke(ActionUpdateOnClose);
                    break;

                case ActionReleaseNotes:
                    EnvTabsLog.Info("UpdateInfoBar: User clicked Release Notes.");
                    if (!string.IsNullOrWhiteSpace(releaseUrl))
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = releaseUrl,
                            UseShellExecute = true
                        });
                    }
                    break;

            }
        }

        public void OnClosed(IVsInfoBarUIElement infoBarUIElement)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                if (currentElement != null)
                {
                    currentElement.Unadvise(adviseCookie);
                    currentElement = null;
                }
            }
            catch
            {
                currentElement = null;
            }
        }
    }
}
