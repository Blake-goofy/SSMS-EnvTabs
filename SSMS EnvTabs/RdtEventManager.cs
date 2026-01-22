using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;

namespace SSMS_EnvTabs
{
    internal sealed class RdtEventManager : IVsRunningDocTableEvents3, IVsSelectionEvents, IDisposable
    {
        private const uint SeidWindowFrame = 1;
        private const uint SeidDocumentFrame = 2;
        private const int RenameRetryCount = 2;
        private const int RenameRetryDelayMs = 250;

        private readonly AsyncPackage package;
        private readonly IVsRunningDocumentTable rdt;
        private readonly IVsUIShellOpenDocument shellOpenDoc;
        private readonly IVsMonitorSelection monitorSelection;
        private readonly ColorByRegexConfigWriter colorWriter;

        private uint rdtEventsCookie;
        private uint selectionEventsCookie;

        private readonly Dictionary<uint, int> renameRetryCounts = new Dictionary<uint, int>();

        private TabGroupConfig cachedConfig;
        private DateTime cachedConfigLastWriteUtc;
        private List<TabRuleMatcher.CompiledRule> cachedRules;

        internal sealed class OpenDocumentInfo
        {
            public uint Cookie { get; set; }
            public IVsWindowFrame Frame { get; set; }
            public string Caption { get; set; }
            public string Moniker { get; set; }
            public string Server { get; set; }
            public string Database { get; set; }
        }

        private RdtEventManager(AsyncPackage package, IVsRunningDocumentTable rdt, IVsUIShellOpenDocument shellOpenDoc, IVsMonitorSelection monitorSelection)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            this.rdt = rdt ?? throw new ArgumentNullException(nameof(rdt));
            this.shellOpenDoc = shellOpenDoc;
            this.monitorSelection = monitorSelection;
            this.colorWriter = new ColorByRegexConfigWriter();
        }

        public static async Task<RdtEventManager> CreateAndStartAsync(AsyncPackage package, CancellationToken cancellationToken)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(cancellationToken);

            var rdt = await package.GetServiceAsync(typeof(SVsRunningDocumentTable)) as IVsRunningDocumentTable;
            var shellOpenDoc = await package.GetServiceAsync(typeof(SVsUIShellOpenDocument)) as IVsUIShellOpenDocument;
            var monitorSelection = await package.GetServiceAsync(typeof(SVsShellMonitorSelection)) as IVsMonitorSelection;

            if (rdt == null)
            {
                EnvTabsLog.Error("RdtEventManager: SVsRunningDocumentTable missing; EnvTabs disabled.");
                return null;
            }

            var mgr = new RdtEventManager(package, rdt, shellOpenDoc, monitorSelection);
            mgr.Start();
            return mgr;
        }

        private void Start()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                rdt.AdviseRunningDocTableEvents(this, out rdtEventsCookie);
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"RdtEventManager: failed to subscribe RDT events: {ex.Message}");
            }

            if (monitorSelection != null)
            {
                try
                {
                    monitorSelection.AdviseSelectionEvents(this, out selectionEventsCookie);
                }
                catch (Exception ex)
                {
                    EnvTabsLog.Info($"RdtEventManager: failed to subscribe selection events: {ex.Message}");
                }
            }
        }

        public void Dispose()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                if (selectionEventsCookie != 0 && monitorSelection != null)
                {
                    monitorSelection.UnadviseSelectionEvents(selectionEventsCookie);
                    selectionEventsCookie = 0;
                }
            }
            catch
            {
                // best-effort
            }

            try
            {
                if (rdtEventsCookie != 0)
                {
                    rdt.UnadviseRunningDocTableEvents(rdtEventsCookie);
                    rdtEventsCookie = 0;
                }
            }
            catch
            {
                // best-effort
            }
        }

        private TabGroupConfig LoadConfigOrNull()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                string path = TabGroupConfigLoader.GetUserConfigPath();
                if (!File.Exists(path))
                {
                    cachedConfig = null;
                    cachedRules = null;
                    cachedConfigLastWriteUtc = default;
                    return null;
                }

                DateTime lastWriteUtc = File.GetLastWriteTimeUtc(path);
                if (cachedConfig != null && lastWriteUtc == cachedConfigLastWriteUtc)
                {
                    return cachedConfig;
                }

                var loaded = TabGroupConfigLoader.LoadOrNull();
                cachedConfig = loaded;
                cachedConfigLastWriteUtc = lastWriteUtc;
                cachedRules = TabRuleMatcher.CompileRules(loaded);
                return loaded;
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"RdtEventManager: config reload failed: {ex.Message}");
                cachedConfig = null;
                cachedRules = null;
                cachedConfigLastWriteUtc = default;
                return null;
            }
        }

        private bool HandlePotentialChange(uint docCookie, IVsWindowFrame frame, string reason)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var config = LoadConfigOrNull();
            if (config == null)
            {
                return true;
            }

            var rules = cachedRules ?? new List<TabRuleMatcher.CompiledRule>();
            if (rules.Count == 0)
            {
                return true;
            }

            bool needsRetry = false;

            if (config.Settings?.EnableAutoRename != false && frame != null)
            {
                string caption = TryReadFrameCaption(frame);
                if (IsRenameEligible(caption))
                {
                    if (TryGetConnectionInfo(frame, out string server, out string database))
                    {
                        try
                        {
                            TabRenamer.ApplyRenamesOrThrow(new[] { (docCookie, frame, server, database, caption) }, rules);
                        }
                        catch (Exception ex)
                        {
                            EnvTabsLog.Info($"Rename failed ({reason}) cookie={docCookie}: {ex.Message}");
                        }
                    }
                    else
                    {
                        needsRetry = true;
                    }
                }
            }

            if (config.Settings?.EnableAutoColor == true)
            {
                try
                {
                    var docs = GetOpenDocumentsSnapshot();
                    colorWriter.UpdateFromSnapshot(docs, rules);
                }
                catch (Exception ex)
                {
                    EnvTabsLog.Info($"ColorByRegex update failed ({reason}): {ex.Message}");
                }
            }

            return !needsRetry;
        }

        private static bool TryGetConnectionInfo(IVsWindowFrame frame, out string server, out string database)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            server = null;
            database = null;

            if (frame == null)
            {
                return false;
            }

            try
            {
                TryPopulateFromCaptions(frame, ref server, ref database);
                return !string.IsNullOrWhiteSpace(server) || !string.IsNullOrWhiteSpace(database);
            }
            catch
            {
                return false;
            }
        }

        private static void TryPopulateFromCaptions(IVsWindowFrame frame, ref string server, ref string database)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            if (frame == null) return;

            string caption = TryReadFrameCaption(frame);
            if (TryParseServerDatabaseFromCaption(caption, out string s1, out string d1))
            {
                if (string.IsNullOrWhiteSpace(server)) server = s1;
                if (string.IsNullOrWhiteSpace(database)) database = d1;
                return;
            }

            try
            {
                if (frame.GetProperty((int)__VSFPROPID.VSFPROPID_EditorCaption, out object editorCaptionObj) == VSConstants.S_OK)
                {
                    string editorCaption = editorCaptionObj as string;
                    if (TryParseServerDatabaseFromCaption(editorCaption, out string s2, out string d2))
                    {
                        if (string.IsNullOrWhiteSpace(server)) server = s2;
                        if (string.IsNullOrWhiteSpace(database)) database = d2;
                    }
                }
            }
            catch
            {
                // ignore
            }
        }

        private static bool TryParseServerDatabaseFromCaption(string caption, out string server, out string database)
        {
            server = null;
            database = null;

            if (string.IsNullOrWhiteSpace(caption)) return false;

            try
            {
                int dash = caption.IndexOf(" - ", StringComparison.Ordinal);
                if (dash < 0)
                {
                    return false;
                }

                string tail = caption.Substring(dash + 3);
                int paren = tail.IndexOf(" (", StringComparison.Ordinal);
                if (paren >= 0)
                {
                    tail = tail.Substring(0, paren);
                }

                tail = tail.Trim();
                if (string.IsNullOrWhiteSpace(tail)) return false;

                int dot = tail.IndexOf('.');
                if (dot > 0 && dot < tail.Length - 1)
                {
                    server = tail.Substring(0, dot).Trim();
                    database = tail.Substring(dot + 1).Trim();
                }
                else
                {
                    server = tail;
                }

                return !string.IsNullOrWhiteSpace(server) || !string.IsNullOrWhiteSpace(database);
            }
            catch
            {
                return false;
            }
        }

        private void UpdateColorOnly(string reason)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var config = LoadConfigOrNull();
            if (config?.Settings?.EnableAutoColor != true)
            {
                return;
            }

            var rules = cachedRules ?? new List<TabRuleMatcher.CompiledRule>();
            if (rules.Count == 0)
            {
                return;
            }

            try
            {
                var docs = GetOpenDocumentsSnapshot();
                colorWriter.UpdateFromSnapshot(docs, rules);
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"ColorByRegex update failed ({reason}): {ex.Message}");
            }
        }

        private List<OpenDocumentInfo> GetOpenDocumentsSnapshot()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var list = new List<OpenDocumentInfo>();

            rdt.GetRunningDocumentsEnum(out IEnumRunningDocuments enumDocs);
            if (enumDocs == null)
            {
                return list;
            }

            uint[] cookies = new uint[1];
            while (enumDocs.Next(1, cookies, out uint fetched) == VSConstants.S_OK && fetched == 1)
            {
                uint cookie = cookies[0];
                string moniker = TryGetMonikerFromCookie(cookie);
                if (string.IsNullOrWhiteSpace(moniker))
                {
                    continue;
                }

                IVsWindowFrame frame = TryGetFrameFromMoniker(moniker);
                if (frame == null)
                {
                    continue;
                }

                string caption = TryReadFrameCaption(frame);
                TryGetConnectionInfo(frame, out string server, out string database);
                list.Add(new OpenDocumentInfo
                {
                    Cookie = cookie,
                    Frame = frame,
                    Caption = caption,
                    Moniker = moniker,
                    Server = server,
                    Database = database
                });
            }

            return list;
        }

        private string TryGetMonikerFromCookie(uint docCookie)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (docCookie == 0) return null;

            try
            {
                rdt.GetDocumentInfo(
                    docCookie,
                    out uint _,
                    out uint _,
                    out uint _,
                    out string moniker,
                    out IVsHierarchy _,
                    out uint _,
                    out IntPtr _);

                return moniker;
            }
            catch
            {
                return null;
            }
        }

        private static string TryReadFrameCaption(IVsWindowFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (frame == null) return null;

            try
            {
                if (frame.GetProperty((int)__VSFPROPID.VSFPROPID_Caption, out object caption) == VSConstants.S_OK)
                {
                    return caption as string;
                }
            }
            catch
            {
                // ignore
            }

            return null;
        }

        private static string TryReadFrameEditorCaption(IVsWindowFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (frame == null) return null;

            try
            {
                if (frame.GetProperty((int)__VSFPROPID.VSFPROPID_EditorCaption, out object caption) == VSConstants.S_OK)
                {
                    return caption as string;
                }
            }
            catch
            {
                // ignore
            }

            return null;
        }

        private void LogFrameCaptions(IVsWindowFrame frame, string reason)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (frame == null) return;

            string caption = TryReadFrameCaption(frame) ?? string.Empty;
            string editorCaption = TryReadFrameEditorCaption(frame) ?? string.Empty;
            EnvTabsLog.Info($"Frame caption ({reason}): '{caption}' | editor='{editorCaption}'");
        }

        private void ScheduleRenameRetry(uint docCookie, string reason)
        {
            if (docCookie == 0) return;

            if (renameRetryCounts.ContainsKey(docCookie))
            {
                return;
            }

            renameRetryCounts[docCookie] = 0;

            package.JoinableTaskFactory.RunAsync(async () =>
            {
                for (int i = 0; i < RenameRetryCount; i++)
                {
                    await Task.Delay(RenameRetryDelayMs).ConfigureAwait(true);
                    await package.JoinableTaskFactory.SwitchToMainThreadAsync();

                    if (!renameRetryCounts.ContainsKey(docCookie))
                    {
                        return;
                    }

                    renameRetryCounts[docCookie] = i + 1;

                    string moniker = TryGetMonikerFromCookie(docCookie);
                    if (string.IsNullOrWhiteSpace(moniker))
                    {
                        continue;
                    }

                    IVsWindowFrame frame = TryGetFrameFromMoniker(moniker);
                    if (frame == null)
                    {
                        continue;
                    }

                    string attemptReason = $"{reason}:Retry#{i + 1}";
                    if (i == 0)
                    {
                        LogFrameCaptions(frame, attemptReason);
                    }
                    bool done = HandlePotentialChange(docCookie, frame, attemptReason);
                    if (done)
                    {
                        renameRetryCounts.Remove(docCookie);
                        return;
                    }
                }

                renameRetryCounts.Remove(docCookie);
            });
        }

        private bool TryGetMonikerFromFrame(IVsWindowFrame frame, out string moniker)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            moniker = null;
            if (frame == null) return false;

            try
            {
                if (frame.GetProperty((int)__VSFPROPID.VSFPROPID_pszMkDocument, out object mk) == VSConstants.S_OK)
                {
                    moniker = mk as string;
                    return !string.IsNullOrWhiteSpace(moniker);
                }
            }
            catch
            {
                // ignore
            }

            return false;
        }

        private bool TryGetCookieFromMoniker(string moniker, out uint cookie)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            cookie = 0;
            if (string.IsNullOrWhiteSpace(moniker)) return false;

            try
            {
                rdt.FindAndLockDocument((uint)_VSRDTFLAGS.RDT_NoLock, moniker, out IVsHierarchy _, out uint _, out IntPtr _, out cookie);
                return cookie != 0;
            }
            catch
            {
                return false;
            }
        }

        private IVsWindowFrame TryGetFrameFromMoniker(string moniker)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (shellOpenDoc == null || string.IsNullOrWhiteSpace(moniker)) return null;

            try
            {
                Guid logicalView = Guid.Empty;
                shellOpenDoc.OpenDocumentViaProject(
                    moniker,
                    ref logicalView,
                    out Microsoft.VisualStudio.OLE.Interop.IServiceProvider _,
                    out IVsUIHierarchy _,
                    out uint _,
                    out IVsWindowFrame frame);

                return frame;
            }
            catch
            {
                return null;
            }
        }

        // --- RDT events ---

        public int OnAfterFirstDocumentLock(uint docCookie, uint dwRDTLockType, uint dwReadLocksRemaining, uint dwEditLocksRemaining)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            string moniker = TryGetMonikerFromCookie(docCookie);
            if (!string.IsNullOrWhiteSpace(moniker))
            {
                IVsWindowFrame frame = TryGetFrameFromMoniker(moniker);
                if (frame != null)
                {
                    LogFrameCaptions(frame, "FirstDocumentLock");
                    bool done = HandlePotentialChange(docCookie, frame, reason: "FirstDocumentLock");
                    if (!done)
                    {
                        ScheduleRenameRetry(docCookie, "FirstDocumentLock");
                    }
                }
            }

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
        public int OnAfterSave(uint docCookie) => VSConstants.S_OK;
        public int OnBeforeSave(uint docCookie) => VSConstants.S_OK;
        public int OnAfterAttributeChange(uint docCookie, uint grfAttribs) => VSConstants.S_OK;
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
                    if (TryGetMonikerFromFrame(frame, out string moniker))
                    {
                        TryGetCookieFromMoniker(moniker, out cookie);
                    }

                    HandlePotentialChange(cookie, frame, reason: "ActiveFrameChanged");
                }
            }

            return VSConstants.S_OK;
        }

        public int OnCmdUIContextChanged(uint dwCmdUICookie, int fActive) => VSConstants.S_OK;

        public int OnSelectionChanged(IVsHierarchy pHierOld, uint itemidOld, IVsMultiItemSelect pMISOld, ISelectionContainer pSCOld,
            IVsHierarchy pHierNew, uint itemidNew, IVsMultiItemSelect pMISNew, ISelectionContainer pSCNew) => VSConstants.S_OK;

        private static bool IsRenameEligible(string caption)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(caption))
                {
                    return false;
                }

                return System.Text.RegularExpressions.Regex.IsMatch(
                    caption,
                    @"^SQLQuery\d+\.sql\b",
                    System.Text.RegularExpressions.RegexOptions.IgnoreCase | System.Text.RegularExpressions.RegexOptions.CultureInvariant);
            }
            catch
            {
                return false;
            }
        }
    }
}
