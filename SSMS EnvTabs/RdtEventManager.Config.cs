using Microsoft.VisualStudio.Shell;
using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace SSMS_EnvTabs
{
    internal sealed partial class RdtEventManager
    {
        private void OnConfigRenamed(object sender, RenamedEventArgs e)
        {
            OnConfigChanged(sender, e);
        }

        private void OnConfigChanged(object sender, FileSystemEventArgs e)
        {
            try
            {
                EnvTabsLog.Info($"Config file system event: {e.ChangeType} - {e.FullPath}");

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

            if (config == null) return;

            var docs = GetOpenDocumentsSnapshot();

            // Rename
            int renamedCount = 0;
            try
            {
                var renameCandidates = docs
                    .Where(doc => !string.IsNullOrWhiteSpace(doc?.Server))
                    .Select(doc => (doc.Cookie, doc.Frame, doc.Server, doc.Database, doc.Caption))
                    .ToList();

                if (renameCandidates.Count > 0)
                {
                    renamedCount = TabRenamer.ApplyRenamesOrThrow(renameCandidates, rules);
                }
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"Reload rename failed: {ex.Message}");
            }

            // Color
            if (config.Settings?.EnableAutoColor == true)
            {
                try
                {
                    colorWriter.UpdateFromSnapshot(docs, rules);
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
                    cachedConfigLastWriteUtc = default;
                    return null;
                }

                DateTime lastWriteUtc = File.GetLastWriteTimeUtc(path);
                if (cachedConfig != null && lastWriteUtc == cachedConfigLastWriteUtc)
                {
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

        private void UpdateColorOnly(string reason)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var config = LoadConfigOrNull();
            if (config?.Settings?.EnableAutoColor != true)
            {
                return;
            }

            var rules = cachedRules ?? new System.Collections.Generic.List<TabRuleMatcher.CompiledRule>();
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
    }
}
