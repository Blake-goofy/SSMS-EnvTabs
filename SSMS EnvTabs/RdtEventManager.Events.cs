using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;

namespace SSMS_EnvTabs
{
    internal sealed partial class RdtEventManager
    {
        // --- RDT events ---
        public int OnAfterFirstDocumentLock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining)
        {
            EnvTabsLog.Info($"RdtEventManager.Events.cs::OnAfterFirstDocumentLock - Cookie={docCookie}");
            ThreadHelper.ThrowIfNotOnUIThread(); // Ensure UI thread

            // Try to hook into this event since OnAfterDocumentWindowShow is unreliable in SSMS
            string moniker = TryGetMonikerFromCookie(docCookie);
            EnvTabsLog.Info($"RdtEventManager.Events.cs::OnAfterFirstDocumentLock - Moniker='{moniker}'");

            if (!string.IsNullOrWhiteSpace(moniker))
            {
                IVsWindowFrame frame = TryGetFrameFromMoniker(moniker);
                if (frame != null)
                {
                    EnvTabsLog.Info($"OnAfterFirstDocumentLock: Frame found for cookie {docCookie}. Triggering HandlePotentialChange.");
                    bool done = HandlePotentialChange(docCookie, frame, reason: "FirstDocumentLock");
                    if (!done)
                    {
                        EnvTabsLog.Info($"OnAfterFirstDocumentLock: Connection info missing. Scheduling retry for cookie {docCookie}.");
                        ScheduleRenameRetry(docCookie, "FirstDocumentLock");
                    }
                }
                else
                {
                    EnvTabsLog.Info($"OnAfterFirstDocumentLock: Frame NOT found for cookie {docCookie} via moniker.");
                }
            }

            return VSConstants.S_OK;
        }

        public int OnAfterDocumentWindowShow(uint docCookie, int fFirstShow, IVsWindowFrame pFrame)
        {
            EnvTabsLog.Info($"OnAfterDocumentWindowShow: Cookie={docCookie}, FirstShow={fFirstShow}");
            ThreadHelper.ThrowIfNotOnUIThread();
            
            if (pFrame != null)
            {
                string caption = TryReadFrameCaption(pFrame);
                EnvTabsLog.Info($"OnAfterDocumentWindowShow: Frame Caption='{caption}'");

                bool done = HandlePotentialChange(docCookie, pFrame, reason: "DocumentWindowShow");
                if (!done)
                {
                    ScheduleRenameRetry(docCookie, "DocumentWindowShow");
                }
            }
            else 
            {
                EnvTabsLog.Info($"OnAfterDocumentWindowShow: Frame is NULL");
            }

            return VSConstants.S_OK;
        }

        public int OnAfterAttributeChangeEx(uint docCookie, uint grfAttribs, IVsHierarchy pHierOld, uint itemidOld, string pszMkDocumentOld, IVsHierarchy pHierNew, uint itemidNew, string pszMkDocumentNew)
        {
            EnvTabsLog.Info($"RdtEventManager.Events.cs::OnAfterAttributeChangeEx - Entered. Cookie={docCookie}, Attribs={grfAttribs}, PsOld='{pszMkDocumentOld}', PsNew='{pszMkDocumentNew}'");
            ThreadHelper.ThrowIfNotOnUIThread();
            
            string moniker = pszMkDocumentNew;
            if (string.IsNullOrWhiteSpace(moniker))
            {
                moniker = pszMkDocumentOld;
            }
            
            if (string.IsNullOrWhiteSpace(moniker))
            {
                // Fallback: If moniker args are empty (common for attribute-only changes like Dirty/Reload),
                // fetch it from the RDT using the cookie.
                moniker = TryGetMonikerFromCookie(docCookie);
                EnvTabsLog.Info($"OnAfterAttributeChangeEx: Moniker fetched from cookie='{moniker}'");
            }
            else 
            {
                EnvTabsLog.Info($"OnAfterAttributeChangeEx: Moniker from args='{moniker}'");
            }

            if (!string.IsNullOrWhiteSpace(moniker))
            {
                IVsWindowFrame frame = TryGetFrameFromMoniker(moniker);
                if (frame != null)
                {
                    EnvTabsLog.Info($"OnAfterAttributeChangeEx: Frame found. Triggering HandlePotentialChange.");
                    bool done = HandlePotentialChange(docCookie, frame, reason: "AttributeChangeEx");

                    // Force a retry if it's an AttributeChange, often the DocView is not yet updated with new connection info
                    if (!done)
                    {
                        ScheduleRenameRetry(docCookie, "AttributeChangeEx");
                    }
                }
                else
                {
                    EnvTabsLog.Info($"OnAfterAttributeChangeEx: Frame NOT found for moniker.");
                }
            }

            return VSConstants.S_OK;
        }

        public int OnBeforeLastDocumentUnlock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining)
        {
            EnvTabsLog.Info($"OnBeforeLastDocumentUnlock: Cookie={docCookie}");
            ThreadHelper.ThrowIfNotOnUIThread();

            if (dwReadLocksRemaining == 0 && dwEditLocksRemaining == 0)
            {
                TabRenamer.ForgetCookie(docCookie);
                renameRetryCounts.Remove(docCookie);
                lastConnectionByCookie.Remove(docCookie);

                UpdateColorOnly("LastDocumentUnlock");
            }

            return VSConstants.S_OK;
        }

        public int OnAfterDocumentWindowHide(uint docCookie, IVsWindowFrame pFrame)
        {
            EnvTabsLog.Info($"OnAfterDocumentWindowHide: Cookie={docCookie}");
            return VSConstants.S_OK;
        }

        public int OnAfterSave(uint docCookie)
        {
            EnvTabsLog.Info($"OnAfterSave: Cookie={docCookie}");
            ThreadHelper.ThrowIfNotOnUIThread();
            UpdateColorOnly("AfterSave");
            return VSConstants.S_OK;
        }

        public int OnBeforeSave(uint docCookie) => VSConstants.S_OK;

        public int OnAfterAttributeChange(uint docCookie, uint grfAttribs)
        {
            EnvTabsLog.Info($"RdtEventManager.Events.cs::OnAfterAttributeChange - Entered. Cookie={docCookie}, Attribs={grfAttribs}");
            ThreadHelper.ThrowIfNotOnUIThread();

            string moniker = TryGetMonikerFromCookie(docCookie);
            if (!string.IsNullOrWhiteSpace(moniker))
            {
                IVsWindowFrame frame = TryGetFrameFromMoniker(moniker);
                if (frame != null)
                {
                    bool done = HandlePotentialChange(docCookie, frame, reason: "AttributeChange");

                    // Force a retry if it's an AttributeChange, because the caption is likely stale.
                    if (!done)
                    {
                        ScheduleRenameRetry(docCookie, "AttributeChange");
                    }
                    else
                    {
                        ScheduleRenameRetry(docCookie, "AttributeChange:Force");
                    }
                }
            }

            return VSConstants.S_OK;
        }

        public int OnBeforeDocumentWindowShow(uint docCookie, int fFirstShow, IVsWindowFrame pFrame) => VSConstants.S_OK;
    }
}
