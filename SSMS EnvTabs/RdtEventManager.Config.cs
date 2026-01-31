using Microsoft.VisualStudio.Shell;
using System.Diagnostics.CodeAnalysis;
using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace SSMS_EnvTabs
{
    internal sealed partial class RdtEventManager
    {
        private void OnAutoConfigDialogClosed(AutoConfigurationService.DialogClosedInfo info)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            string reason = info == null
                ? "DialogClosed"
                : $"Result={info.Result}, Server={info.Server}, Database={info.Database}, ChangesApplied={info.ChangesApplied}";

            LogColorSnapshot(reason);

            if (info != null && !info.ChangesApplied)
            {
                UpdateColorOnly("DialogClosed", force: true);
                // Avoid immediately overwriting regex with a partial snapshot after dialog close.
                suppressColorUpdatesUntilUtc = DateTime.UtcNow.AddSeconds(2);
            }
        }

        private bool IsColorUpdateSuppressed()
        {
            return DateTime.UtcNow < suppressColorUpdatesUntilUtc;
        }

        private void LogColorSnapshot(string reason)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var config = LoadConfigOrNull();
            var rules = cachedRules ?? new System.Collections.Generic.List<TabRuleMatcher.CompiledRule>();
            var docs = GetOpenDocumentsSnapshot();

            EnvTabsLog.Info($"Snapshot ({reason}) - Rules={rules.Count}, Tabs={docs.Count}, AutoColor={(config?.Settings?.EnableAutoColor == true)}");

            foreach (var rule in rules)
            {
                EnvTabsLog.Info($"Rule: Name='{rule.GroupName}', Server='{rule.Server}', Database='{rule.Database}', Priority={rule.Priority}, ColorIndex={rule.ColorIndex}");
            }

            foreach (var doc in docs)
            {
                string fileName = null;
                try { fileName = System.IO.Path.GetFileName(doc.Moniker); }
                catch (Exception ex)
                {
                    EnvTabsLog.Verbose($"LogColorSnapshot - File name parse failed: {ex.Message}");
                }
                string group = TabRuleMatcher.MatchGroup(rules, doc.Server, doc.Database);
                EnvTabsLog.Info($"Tab: Cookie={doc.Cookie}, Server='{doc.Server}', Database='{doc.Database}', Group='{group}', File='{fileName}', Moniker='{doc.Moniker}'");
            }

            string block = colorWriter.BuildGeneratedBlockPreview(docs, rules);
            if (!string.IsNullOrWhiteSpace(block))
            {
                EnvTabsLog.Info("Regex Preview:\n" + block);
            }
        }
        private void OnConfigRenamed(object sender, RenamedEventArgs e)
        {
            OnConfigChanged(sender, e);
        }

        [SuppressMessage("Usage", "VSTHRD010", Justification = "FileSystemWatcher callbacks are off-thread; this method marshals as needed.")]
        private void OnConfigChanged(object sender, FileSystemEventArgs e)
        {
            try
            {
                // EnvTabsLog.Info($"Config file system event: {e.ChangeType} - {e.FullPath}"); 

                // Debounce logic
                CancellationToken token;
                lock (debounceLock)
                {
                    debounceCts?.Cancel();
                    debounceCts = new CancellationTokenSource();
                    token = debounceCts.Token;
                }

                _ = package.JoinableTaskFactory.RunAsync(async () =>
                {
                    try
                    {
                        // Log inside async task to allow switching to UI thread if needed (though we just write to file off-thread primarily)
                        // EnvTabsLog.Info handles CheckAccess internally now.
                        EnvTabsLog.Info($"Config file system event: {e.ChangeType} - {e.FullPath}");

                        await Task.Delay(500, token);
                        await package.JoinableTaskFactory.SwitchToMainThreadAsync(token);
                        ReloadAndApplyConfig();
                    }
                    catch (OperationCanceledException)
                    {
                        // Ignored
                    }
                    catch (Exception ex)
                    {
                        EnvTabsLog.Info($"Config reload task failed: {ex.Message}");
                    }
                });
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"OnConfigChanged error: {ex.Message}");
            }
        }

        private void ReloadAndApplyConfig()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            EnvTabsLog.Info("Configuration changed, reloading...");

            // Invalidate cache
            cachedConfig = null;
            // Clear suppressed "do not ask again" connections so that if user removed a rule, we effectively reset the "ignore" state
            // and will prompt again if they connect to it.
            AutoConfigurationService.ClearSuppressed();

            var config = LoadConfigOrNull();
            var rules = cachedRules; // LoadConfigOrNull updates cachedRules
            var manualRules = cachedManualRules;

            if (config == null) return;

            var docs = GetOpenDocumentsSnapshot();

            // Rename
            int renamedCount = 0;
            if (config.Settings?.EnableAutoRename != false)
            {
                try
                {
                    var renameCandidates = docs
                        .Where(doc => !string.IsNullOrWhiteSpace(doc?.Moniker)) // Filter empty?
                        .Select(doc => new TabRenameContext
                        {
                            Cookie = doc.Cookie,
                            Frame = doc.Frame,
                            Server = doc.Server,
                            Database = doc.Database,
                            FrameCaption = doc.Caption,
                            Moniker = doc.Moniker
                        })
                        .ToList();

                    if (renameCandidates.Count > 0)
                    {
                        renamedCount = TabRenamer.ApplyRenamesOrThrow(renameCandidates, rules, manualRules, config.Settings?.NewQueryRenameStyle);
                    }
                }
                catch (Exception ex)
                {
                    EnvTabsLog.Info($"Reload rename failed: {ex.Message}");
                }
            }

            // Color
            if (config.Settings?.EnableAutoColor == true)
            {
                try
                {
                    if (!IsColorUpdateSuppressed())
                    {
                        colorWriter.UpdateFromSnapshot(docs, rules, manualRules);
                    }
                    else
                    {
                        EnvTabsLog.Info("Color update suppressed (ConfigReload)");
                    }
                }
                catch (Exception ex)
                {
                    EnvTabsLog.Info($"Reload color update failed: {ex.Message}");
                }
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
                    cachedManualRules = null;
                    cachedConfigLastWriteUtc = default;
                    return null;
                }

                DateTime lastWriteUtc = File.GetLastWriteTimeUtc(path);
                if (cachedConfig != null && lastWriteUtc == cachedConfigLastWriteUtc)
                {
                    if ((cachedRules == null || cachedManualRules == null) && cachedConfig != null)
                    {
                        cachedRules = TabRuleMatcher.CompileRules(cachedConfig);
                        cachedManualRules = TabRuleMatcher.CompileManualRules(cachedConfig);
                        EnvTabsLog.Info($"RdtEventManager.Config.cs::LoadConfigOrNull - Rebuilt cached rules. Rules={cachedRules?.Count ?? 0}, ManualRules={cachedManualRules?.Count ?? 0}");
                    }
                    return cachedConfig;
                }

                var loaded = TabGroupConfigLoader.LoadOrNull();

                if (loaded != null && loaded.Settings != null)
                {
                    EnvTabsLog.Enabled = loaded.Settings.EnableLogging;
                }

                cachedConfig = loaded;
                cachedConfigLastWriteUtc = lastWriteUtc;
                cachedRules = TabRuleMatcher.CompileRules(loaded);
                cachedManualRules = TabRuleMatcher.CompileManualRules(loaded);
                EnvTabsLog.Info($"RdtEventManager.Config.cs::LoadConfigOrNull - Loaded. Rules={cachedRules?.Count ?? 0}, ManualRules={cachedManualRules?.Count ?? 0}, AutoColor={loaded?.Settings?.EnableAutoColor}, AutoRename={loaded?.Settings?.EnableAutoRename}, Polling={loaded?.Settings?.EnableConnectionPolling}");
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

        private void UpdateColorOnly(string reason, bool force = false)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!force && IsColorUpdateSuppressed())
            {
                EnvTabsLog.Info($"Color update suppressed ({reason})");
                return;
            }

            var config = LoadConfigOrNull();
            if (config?.Settings?.EnableAutoColor != true)
            {
                return;
            }

            var rules = cachedRules ?? TabRuleMatcher.CompileRules(config);
            var manualRules = cachedManualRules ?? TabRuleMatcher.CompileManualRules(config);

            if (rules.Count == 0 && manualRules.Count == 0)
            {
                return;
            }

            try
            {
                var docs = GetOpenDocumentsSnapshot();
                colorWriter.UpdateFromSnapshot(docs, rules, manualRules);
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"ColorByRegex update failed ({reason}): {ex.Message}");
            }
        }
    }
}
