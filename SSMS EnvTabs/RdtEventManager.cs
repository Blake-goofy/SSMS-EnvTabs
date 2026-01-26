using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Linq.Expressions;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace SSMS_EnvTabs
{
    internal sealed partial class RdtEventManager : IVsRunningDocTableEvents3, IVsSelectionEvents, IDisposable
    {
        private const uint SeidWindowFrame = 1;
        private const uint SeidDocumentFrame = 2;
        private const int RenameRetryCount = 20; // Increased to handle slow connection changes (up to 5s)
        private const int RenameRetryDelayMs = 250;
        private static readonly Regex RenameEligibleRegex = new Regex(
            @"^SQLQuery\d+\.sql\b",
            RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

        private readonly AsyncPackage package;
        private readonly IVsRunningDocumentTable rdt;
        private readonly IVsUIShellOpenDocument shellOpenDoc;
        private readonly IVsMonitorSelection monitorSelection;
        private readonly ColorByRegexConfigWriter colorWriter;

        private uint rdtEventsCookie;
        private uint selectionEventsCookie;

        private readonly Dictionary<uint, int> renameRetryCounts = new Dictionary<uint, int>();
        private FileSystemWatcher configWatcher;
        private CancellationTokenSource debounceCts;
        private readonly object debounceLock = new object();

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

            try
            {
                string configPath = TabGroupConfigLoader.GetUserConfigPath();
                string configDir = Path.GetDirectoryName(configPath);
                if (Directory.Exists(configDir))
                {
                    configWatcher = new FileSystemWatcher(configDir, "TabGroupConfig.json");
                    configWatcher.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.CreationTime | NotifyFilters.Size | NotifyFilters.FileName;
                    configWatcher.Changed += OnConfigChanged;
                    configWatcher.Created += OnConfigChanged;
                    configWatcher.Renamed += OnConfigRenamed;
                    configWatcher.EnableRaisingEvents = true;
                }
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"RdtEventManager: failed to start config watcher: {ex.Message}");
            }
        }

        public void Dispose()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                configWatcher?.Dispose();
                debounceCts?.Cancel();
                debounceCts?.Dispose();
            }
            catch { }

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

        private bool HandlePotentialChange(uint docCookie, IVsWindowFrame frame, string reason)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            
            var config = LoadConfigOrNull();
            if (config == null)
            {
                return true;
            }

            var rules = cachedRules ?? new List<TabRuleMatcher.CompiledRule>();
            // If rules.Count == 0, we still proceed because we might need to AutoConfigure/ProposeNewRule

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
                            
                            // Check for match explicitly to avoid duplicate rules if rename fails
                            string matchedGroup = TabRuleMatcher.MatchGroup(rules, server, database);
                            bool hasMatchingRule = !string.IsNullOrWhiteSpace(matchedGroup);

                            // Auto-Configure logic if no rule matched
                            if (!hasMatchingRule && !string.IsNullOrWhiteSpace(config.Settings?.AutoConfigure))
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

            bool isConnectionEvent = reason != null && (
                reason.IndexOf("DocumentWindowShow", StringComparison.OrdinalIgnoreCase) >= 0 ||
                reason.IndexOf("AttributeChange", StringComparison.OrdinalIgnoreCase) >= 0 ||
                reason.IndexOf("AttributeChangeEx", StringComparison.OrdinalIgnoreCase) >= 0 ||
                reason.IndexOf("ActiveFrameChanged", StringComparison.OrdinalIgnoreCase) >= 0 ||
                reason.IndexOf("DocViewEvent", StringComparison.OrdinalIgnoreCase) >= 0
            );

            bool shouldUpdateColor = renamedCount > 0 || isConnectionEvent;

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

    }
}
