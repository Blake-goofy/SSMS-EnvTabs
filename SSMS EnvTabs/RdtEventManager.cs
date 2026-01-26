using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Linq.Expressions;
using System.Threading;
using System.Threading.Tasks;

namespace SSMS_EnvTabs
{
    internal sealed class RdtEventManager : IVsRunningDocTableEvents3, IVsSelectionEvents, IDisposable
    {
        private const uint SeidWindowFrame = 1;
        private const uint SeidDocumentFrame = 2;
        private const int RenameRetryCount = 20; // Increased to handle slow connection changes (up to 5s)
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

        private sealed class ReflectionEventSubscription : IDisposable
        {
            private readonly object target;
            private readonly EventInfo evt;
            private readonly Delegate handler;

            public ReflectionEventSubscription(object target, EventInfo evt, Delegate handler)
            {
                this.target = target;
                this.evt = evt;
                this.handler = handler;
            }

            public void Dispose()
            {
                try
                {
                    evt?.RemoveEventHandler(target, handler);
                }
                catch
                {
                    // best-effort
                }
            }
        }

        private readonly Dictionary<uint, List<ReflectionEventSubscription>> docViewSubscriptionsByCookie =
            new Dictionary<uint, List<ReflectionEventSubscription>>();
        
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
                foreach (var kvp in docViewSubscriptionsByCookie)
                {
                    foreach (var sub in kvp.Value)
                    {
                        try { sub.Dispose(); } catch { }
                    }
                }
                docViewSubscriptionsByCookie.Clear();
            }
            catch { }

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
            
            // Too verbose: EnvTabsLog.Info($"HandlePotentialChange: cookie={docCookie} reason={reason}");

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
            int renamedCount = 0;

            if (config.Settings?.EnableAutoRename != false && frame != null)
            {
                string caption = TryReadFrameCaption(frame);
                
                // If the reason is AttributeChange (connection change), we should attempt rename even if it doesn't look eligible anymore (e.g. if it was already renamed)
                // because the user is changing the connection to something new.
                // Or if it reverted to "SQLQuery1.sql".
                bool forceCheck = reason != null && (
                    reason.StartsWith("AttributeChange", StringComparison.OrdinalIgnoreCase) ||
                    reason.IndexOf("AttributeChange", StringComparison.OrdinalIgnoreCase) >= 0 ||
                    reason.IndexOf("DocViewEvent", StringComparison.OrdinalIgnoreCase) >= 0
                );

                if (forceCheck || IsRenameEligible(caption))
                {
                    if (TryGetConnectionInfo(frame, out string server, out string database))
                    {
                        try
                        {
                            renamedCount = TabRenamer.ApplyRenamesOrThrow(new[] { (docCookie, frame, server, database, caption) }, rules);
                            
                            // Auto-Configure logic if no rule matched and not renamed
                            if (renamedCount == 0 && !string.IsNullOrWhiteSpace(config.Settings?.AutoConfigure))
                            {
                                // Check if this server/db actually matched any existing rule?
                                // TabRenamer returns 0 if no match found.
                                // But we should verify we haven't already processed it.
                                // We need to be careful not to spam prompts on initial load.
                                // Only prompt if this is a "valid" connection (has text).
                                if (!string.IsNullOrWhiteSpace(server))
                                {
                                    // Dispatch to UI thread later to avoid blocking RDT event
                                    _ = package.JoinableTaskFactory.RunAsync(async () =>
                                    {
                                        await package.JoinableTaskFactory.SwitchToMainThreadAsync();
                                        AutoConfigurationService.ProposeNewRule(config, server, database);
                                    });
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            EnvTabsLog.Info($"Rename failed ({reason}) cookie={docCookie}: {ex.Message}");
                        }
                    }
                    else
                    {
                        // Needs retry if connection info not found yet
                        needsRetry = true;
                    }
                }
            }

            bool shouldUpdateColor = renamedCount > 0 
                || (reason != null && reason.IndexOf("FirstDocumentLock", StringComparison.OrdinalIgnoreCase) >= 0)
                || (reason != null && reason.IndexOf("DocumentWindowShow", StringComparison.OrdinalIgnoreCase) >= 0);

            if (config.Settings?.EnableAutoColor == true && shouldUpdateColor)
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

        private void TryHookDocViewConnectionEvents(uint docCookie, IVsWindowFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (docCookie == 0 || frame == null)
            {
                return;
            }

            if (docViewSubscriptionsByCookie.ContainsKey(docCookie))
            {
                return; // already hooked
            }

            object docView = null;
            try
            {
                if (frame.GetProperty((int)__VSFPROPID.VSFPROPID_DocView, out object dv) == VSConstants.S_OK)
                {
                    docView = dv;
                }

                if (docView == null && frame.GetProperty((int)__VSFPROPID.VSFPROPID_ViewHelper, out object vh) == VSConstants.S_OK)
                {
                    docView = vh;
                }
            }
            catch
            {
                docView = null;
            }

            if (docView == null)
            {
                return;
            }

            var type = docView.GetType();

            // Names observed in logs: add_ConnectionChanged, add_ConnectionDisconnected, add_NewConnectionForScript
            string[] eventNames = new[]
            {
                "ConnectionChanged",
                "ConnectionDisconnected",
                "NewConnectionForScript",
            };

            var subs = new List<ReflectionEventSubscription>();

            foreach (string eventName in eventNames)
            {
                EventInfo evt = null;
                try
                {
                    evt = type.GetEvent(eventName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                }
                catch
                {
                    evt = null;
                }

                if (evt?.EventHandlerType == null)
                {
                    continue;
                }

                try
                {
                    var handler = CreateEventHandlerDelegate(evt.EventHandlerType, docCookie, eventName);
                    evt.AddEventHandler(docView, handler);
                    subs.Add(new ReflectionEventSubscription(docView, evt, handler));
                    EnvTabsLog.Info($"Hooked DocView event '{eventName}' for cookie={docCookie}");
                }
                catch (Exception ex)
                {
                    EnvTabsLog.Info($"Failed to hook DocView event '{eventName}' for cookie={docCookie}: {ex.Message}");
                }
            }

            if (subs.Count > 0)
            {
                docViewSubscriptionsByCookie[docCookie] = subs;
            }
        }

        private Delegate CreateEventHandlerDelegate(Type handlerType, uint docCookie, string eventName)
        {
            // Build a delegate matching *any* event handler signature; ignore its parameters.
            MethodInfo invoke = handlerType.GetMethod("Invoke");
            if (invoke == null)
            {
                throw new InvalidOperationException("Handler delegate has no Invoke method.");
            }

            var parameters = invoke.GetParameters()
                .Select(p => Expression.Parameter(p.ParameterType, p.Name))
                .ToArray();

            var methodInfo = typeof(RdtEventManager).GetMethod(nameof(OnDocViewEventFired), BindingFlags.Instance | BindingFlags.NonPublic);
            var call = Expression.Call(
                Expression.Constant(this),
                methodInfo,
                Expression.Constant(docCookie),
                Expression.Constant(eventName));

            Expression body;
            if (invoke.ReturnType == typeof(void))
            {
                body = call;
            }
            else
            {
                body = Expression.Block(call, Expression.Default(invoke.ReturnType));
            }

            return Expression.Lambda(handlerType, body, parameters).Compile();
        }

        private void OnDocViewEventFired(uint docCookie, string eventName)
        {
            // Don't assume event thread; schedule onto UI thread.
            _ = package.JoinableTaskFactory.RunAsync(async () =>
            {
                try
                {
                    await Task.Delay(200); // Small delay to let SSMS update internal state
                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                    
                    // Force a retry sequence
                    ScheduleRenameRetry(docCookie, $"DocViewEvent:{eventName}");
                }
                catch (Exception ex)
                {
                    await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                    EnvTabsLog.Info($"DocView event handler failed ({eventName}) cookie={docCookie}: {ex.Message}");
                }
            });
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

            _ = package.JoinableTaskFactory.RunAsync(async () =>
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
                    TryHookDocViewConnectionEvents(docCookie, frame);
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
                TryHookDocViewConnectionEvents(docCookie, pFrame);
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

            // EnvTabsLog.Info($"OnAfterAttributeChangeEx: cookie={docCookie} attribs={grfAttribs} old={pszMkDocumentOld} new={pszMkDocumentNew}");

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

                if (docViewSubscriptionsByCookie.TryGetValue(docCookie, out var subs))
                {
                    foreach (var sub in subs)
                    {
                        try { sub.Dispose(); } catch { }
                    }
                    docViewSubscriptionsByCookie.Remove(docCookie);
                }

                UpdateColorOnly("LastDocumentUnlock");
            }

            return VSConstants.S_OK;
        }

        public int OnAfterDocumentWindowHide(uint docCookie, IVsWindowFrame pFrame) => VSConstants.S_OK;
        public int OnAfterSave(uint docCookie) => VSConstants.S_OK;
        public int OnBeforeSave(uint docCookie) => VSConstants.S_OK;
        public int OnAfterAttributeChange(uint docCookie, uint grfAttribs)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // EnvTabsLog.Info($"OnAfterAttributeChange: cookie={docCookie} attribs={grfAttribs}");

            string moniker = TryGetMonikerFromCookie(docCookie);
            if (!string.IsNullOrWhiteSpace(moniker))
            {
                IVsWindowFrame frame = TryGetFrameFromMoniker(moniker);
                if (frame != null)
                {
                    // Pass "force=true" semantics via the reason string or just rely on ScheduleRenameRetry
                    // The issue is that HandlePotentialChange calls IsRenameEligible which expects "SQLQueryX.sql" format OR already renamed format.
                    // But if SSMS resets to "SQLQuery1.sql" (no connection info), IsRenameEligible=true, but TryGetConnectionInfo=false -> needsRetry=true.
                    
                    bool done = HandlePotentialChange(docCookie, frame, reason: "AttributeChange");
                    
                    // If not successfully renamed (e.g. because connection info missing), schedule retries
                    if (!done)
                    {
                        ScheduleRenameRetry(docCookie, "AttributeChange");
                    }
                    else
                    {
                        // Even if "done" is true (meaning we processed it), if we just renamed it back to "SQLQuery1.sql" because connection info was missing...
                        // Wait, HandlePotentialChange only returns true if it *successfully* renamed (ApplyRenamesOrThrow) OR if it decided no rename needed.
                        // If connection info is missing, it returns false (needsRetry).
                        // So logic holds.
                        
                        // BUT: What if IsRenameEligible returns false? (e.g. caption="Previously Renamed Tab")
                        // If user changes connection, caption might momentarily stay "Previously Renamed Tab".
                        // IsRenameEligible checks for "SQLQuery...". It returns FALSE for "Previously Renamed Tab".
                        // So HandlePotentialChange returns TRUE (early exit).
                        // And we never retry.
                        
                        // FIX: We must force a retry if it's an AttributeChange, because the caption is likely STALE.
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
                         // EnvTabsLog.Info($"ActiveFrameChanged: Found frame but no cookie. Moniker='{moniker}'");
                    }
                    else
                    {
                        TryHookDocViewConnectionEvents(cookie, frame);
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
