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
            return VSConstants.S_OK;
        }

        public int OnAfterDocumentWindowShow(uint docCookie, int fFirstShow, IVsWindowFrame pFrame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (pFrame != null)
            {
                bool done = HandlePotentialChange(docCookie, pFrame, reason: "DocumentWindowShow");
                if (!done)
                {
                    ScheduleRenameRetry(docCookie, "DocumentWindowShow");
                }
            }

            return VSConstants.S_OK;
        }

        public int OnAfterAttributeChangeEx(uint docCookie, uint grfAttribs, IVsHierarchy pHierOld, uint itemidOld, string pszMkDocumentOld, IVsHierarchy pHierNew, uint itemidNew, string pszMkDocumentNew)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            string moniker = pszMkDocumentNew;
            if (string.IsNullOrWhiteSpace(moniker))
            {
                moniker = pszMkDocumentOld;
            }

            if (!string.IsNullOrWhiteSpace(moniker))
            {
                IVsWindowFrame frame = TryGetFrameFromMoniker(moniker);
                if (frame != null)
                {
                    HandlePotentialChange(docCookie, frame, reason: "AttributeChangeEx");
                }
            }

            return VSConstants.S_OK;
        }

        public int OnBeforeLastDocumentUnlock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (dwReadLocksRemaining == 0 && dwEditLocksRemaining == 0)
            {
                TabRenamer.ForgetCookie(docCookie);
                renameRetryCounts.Remove(docCookie);

                UpdateColorOnly("LastDocumentUnlock");
            }

            return VSConstants.S_OK;
        }

        public int OnAfterDocumentWindowHide(uint docCookie, IVsWindowFrame pFrame) => VSConstants.S_OK;

        public int OnAfterSave(uint docCookie)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            UpdateColorOnly("AfterSave");
            return VSConstants.S_OK;
        }

        public int OnBeforeSave(uint docCookie) => VSConstants.S_OK;

        public int OnAfterAttributeChange(uint docCookie, uint grfAttribs)
        {
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

        // --- Selection events ---

        public int OnElementValueChanged(uint elementid, object varValueOld, object varValueNew)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (elementid == SeidDocumentFrame || elementid == SeidWindowFrame)
            {
                if (varValueNew is IVsWindowFrame frame)
                {
                    uint cookie = 0;
                    string moniker = null;
                    if (TryGetMonikerFromFrame(frame, out moniker))
                    {
                        TryGetCookieFromMoniker(moniker, out cookie);
                    }

                    if (cookie == 0)
                    {
                        // Found frame but no cookie
                    }

                    bool done = HandlePotentialChange(cookie, frame, reason: "ActiveFrameChanged");
                    if (!done && cookie != 0)
                    {
                        ScheduleRenameRetry(cookie, "ActiveFrameChanged");
                    }
                }
            }

            return VSConstants.S_OK;
        }

        public int OnCmdUIContextChanged(uint dwCmdUICookie, int fActive) => VSConstants.S_OK;

        public int OnSelectionChanged(IVsHierarchy pHierOld, uint itemidOld, IVsMultiItemSelect pMISOld, ISelectionContainer pSCOld,
            IVsHierarchy pHierNew, uint itemidNew, IVsMultiItemSelect pMISNew, ISelectionContainer pSCNew) => VSConstants.S_OK;
    }
}
