using Microsoft.VisualStudio.Shell;
using System.Diagnostics.CodeAnalysis;
using System;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

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

        private void OnColorConfigPathResolved(string configPath)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (string.IsNullOrWhiteSpace(configPath))
            {
                return;
            }

            lastColorConfigPath = configPath;

            string dir = Path.GetDirectoryName(configPath);
            if (string.IsNullOrWhiteSpace(dir) || !Directory.Exists(dir))
            {
                return;
            }

            if (string.Equals(dir, groupColorWatcherDir, StringComparison.OrdinalIgnoreCase) && groupColorWatcher != null)
            {
                return;
            }

            try
            {
                groupColorWatcher?.Dispose();
                groupColorWatcher = new FileSystemWatcher(dir, "customized-groupid-color-*.json");
                groupColorWatcher.NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.CreationTime | NotifyFilters.Size | NotifyFilters.FileName;
                groupColorWatcher.Changed += OnGroupColorMapChanged;
                groupColorWatcher.Created += OnGroupColorMapChanged;
                groupColorWatcher.Renamed += OnGroupColorMapChanged;
                groupColorWatcher.EnableRaisingEvents = true;
                groupColorWatcherDir = dir;

                EnvTabsLog.Verbose($"RdtEventManager.Config.cs::OnColorConfigPathResolved - Watching '{dir}' for group color overrides.");
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"Group color watcher setup failed: {ex.Message}");
            }
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
        private void OnGroupColorMapChanged(object sender, FileSystemEventArgs e)
        {
            try
            {
                CancellationToken token;
                lock (groupColorDebounceLock)
                {
                    groupColorDebounceCts?.Cancel();
                    groupColorDebounceCts = new CancellationTokenSource();
                    token = groupColorDebounceCts.Token;
                }

                _ = package.JoinableTaskFactory.RunAsync(async () =>
                {
                    try
                    {
                        EnvTabsLog.Info($"Group color override event: {e.ChangeType} - {e.FullPath}");
                        await Task.Delay(500, token);
                        await package.JoinableTaskFactory.SwitchToMainThreadAsync(token);
                        ApplyGroupColorOverridesToConfig();
                        UpdateColorOnly("GroupColorMapChanged", force: true);
                    }
                    catch (OperationCanceledException)
                    {
                        // Ignored
                    }
                    catch (Exception ex)
                    {
                        EnvTabsLog.Info($"Group color override reload failed: {ex.Message}");
                    }
                });
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"OnGroupColorMapChanged error: {ex.Message}");
            }
        }

        private void ApplyGroupColorOverridesToConfig()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (string.IsNullOrWhiteSpace(lastColorConfigPath))
            {
                return;
            }

            var config = LoadConfigOrNull();
            if (config?.ConnectionGroups == null || config.ConnectionGroups.Count == 0)
            {
                return;
            }

            var rules = cachedRules ?? TabRuleMatcher.CompileRules(config);
            var docs = GetOpenDocumentsSnapshot();

            var overrideByBaseRegex = colorWriter.TryGetOverrideColorsByBaseRegex(lastColorConfigPath);
            if (overrideByBaseRegex == null || overrideByBaseRegex.Count == 0)
            {
                return;
            }

            var groupToBaseRegex = colorWriter.BuildGroupNameToBaseRegex(docs, rules);
            if (groupToBaseRegex == null || groupToBaseRegex.Count == 0)
            {
                return;
            }

            bool updated = false;
            bool enableColorWarning = config.Settings?.EnableColorWarning != false;

            foreach (var rule in config.ConnectionGroups)
            {
                if (rule == null || string.IsNullOrWhiteSpace(rule.GroupName))
                {
                    continue;
                }

                if (groupToBaseRegex.TryGetValue(rule.GroupName, out string baseRegex)
                    && overrideByBaseRegex.TryGetValue(baseRegex, out int newColor)
                    && rule.ColorIndex != newColor)
                {
                    if (enableColorWarning && IsColorUsedByOtherRule(config, rule, newColor))
                    {
                        var choice = MessageBox.Show(
                            "This color is already assigned to another connection group. Open the config to resolve?",
                            "SSMS EnvTabs",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Warning);

                        if (choice == DialogResult.Yes)
                        {
                            OpenConfigInEditor();
                        }
                    }

                    EnvTabsLog.Info($"Updating colorIndex for group '{rule.GroupName}' from {rule.ColorIndex} to {newColor}");
                    rule.ColorIndex = newColor;
                    updated = true;
                }
            }

            if (updated)
            {
                SaveConfig(config);
                EnvTabsLog.Info("Group color overrides saved. Awaiting config watcher reload.");
            }
        }

        private static bool IsColorUsedByOtherRule(TabGroupConfig config, TabGroupRule currentRule, int colorIndex)
        {
            if (config?.ConnectionGroups == null) return false;

            foreach (var rule in config.ConnectionGroups)
            {
                if (!ReferenceEquals(rule, currentRule) && rule?.ColorIndex == colorIndex)
                {
                    return true;
                }
            }

            return false;
        }

        private static void OpenConfigInEditor()
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            string path = TabGroupConfigLoader.GetUserConfigPath();
            VsShellUtilities.OpenDocument(ServiceProvider.GlobalProvider, path);
        }

        private static void SaveConfig(TabGroupConfig config)
        {
            TabGroupConfigLoader.SaveConfig(config);
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
