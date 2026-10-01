using Microsoft.VisualStudio;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using System;
using System.Collections.Generic;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Media;
using System.Windows.Threading;

namespace SSMS_EnvTabs
{
    internal sealed partial class RdtEventManager
    {
        private static readonly ConditionalWeakTable<IWpfTextView, EditorTintController> editorTintControllers = new ConditionalWeakTable<IWpfTextView, EditorTintController>();
        private readonly HashSet<uint> editorTintedCookies = new HashSet<uint>();

        /// <summary>
        /// Blends a small amount of the group color into the query editor background.
        /// Returns false when the editor is not ready yet so the caller can retry.
        /// </summary>
        private bool TryApplyEditorTint(uint docCookie, IVsWindowFrame frame, string moniker, IReadOnlyList<TabRuleMatcher.CompiledRule> rules, IReadOnlyList<TabRuleMatcher.CompiledManualRule> manualRules, TabGroupSettings settings)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (frame == null)
            {
                return true;
            }

            if (settings?.EnableAutoColor != true || SystemParameters.HighContrast)
            {
                TryRestoreEditorTint(docCookie, frame);
                return true;
            }

            if (!TryGetConnectionInfo(frame, out string server, out string database) || string.IsNullOrWhiteSpace(server))
            {
                TryRestoreEditorTint(docCookie, frame);
                return true;
            }

            var manualMatch = TabRuleMatcher.MatchManual(manualRules, moniker);
            var matchedRule = TabRuleMatcher.MatchRule(rules, server, database);

            // Per-rule setting overrides global; null means use global default.
            bool tintEnabled = matchedRule?.EnableEditorTint ?? settings.InitialEditorTint;
            int? colorIndex = manualMatch?.ColorIndex ?? matchedRule?.ColorIndex;
            if (!tintEnabled || !colorIndex.HasValue || colorIndex.Value < 0 || colorIndex.Value >= ColorPalette.Hex.Length)
            {
                TryRestoreEditorTint(docCookie, frame);
                return true;
            }

            List<IWpfTextView> views = GetFrameWpfTextViews(frame);
            if (views.Count == 0)
            {
                EnvTabsLog.Info($"EditorTint: text view not ready cookie={docCookie} server='{server}' db='{database}'");
                return false;
            }

            var tint = (Color)ColorConverter.ConvertFromString(ColorPalette.Hex[colorIndex.Value]);
            int strength = EditorTint.ClampStrength(settings.EditorTintStrength);

            int appliedCount = 0;
            foreach (IWpfTextView view in views)
            {
                if (editorTintControllers.GetValue(view, v => new EditorTintController(v)).Apply(tint, strength))
                {
                    appliedCount++;
                }
            }

            if (appliedCount > 0 && docCookie != 0)
            {
                editorTintedCookies.Add(docCookie);
            }

            EnvTabsLog.Info($"EditorTint: {(appliedCount > 0 ? "applied" : "failed")} cookie={docCookie} server='{server}' db='{database}' colorIndex={colorIndex.Value} strength={strength}% views={views.Count} updated={appliedCount}");
            return appliedCount == views.Count;
        }

        private void TryRestoreEditorTint(uint docCookie, IVsWindowFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            // Skip the text view lookup for documents that were never tinted (the common case when the feature is off).
            if (docCookie == 0 || !editorTintedCookies.Remove(docCookie))
            {
                return;
            }

            foreach (IWpfTextView view in GetFrameWpfTextViews(frame))
            {
                if (editorTintControllers.TryGetValue(view, out EditorTintController controller))
                {
                    controller.Restore();
                }
            }
        }

        /// <summary>
        /// Returns every WPF text view of the document, including the second pane of a split editor.
        /// </summary>
        private List<IWpfTextView> GetFrameWpfTextViews(IVsWindowFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var views = new List<IWpfTextView>();
            if (frame == null)
            {
                return views;
            }

            var vsViews = new List<IVsTextView>();
            if (TryGetFrameCodeWindow(frame, out IVsCodeWindow codeWindow))
            {
                try
                {
                    if (ErrorHandler.Succeeded(codeWindow.GetPrimaryView(out IVsTextView primary)) && primary != null)
                    {
                        vsViews.Add(primary);
                    }

                    if (ErrorHandler.Succeeded(codeWindow.GetSecondaryView(out IVsTextView secondary)) && secondary != null)
                    {
                        vsViews.Add(secondary);
                    }
                }
                catch
                {
                    // Ignore and fall back to the active view.
                }
            }

            if (vsViews.Count == 0 && TryGetActiveVsTextView(frame, out IVsTextView active, out _))
            {
                vsViews.Add(active);
            }

            foreach (IVsTextView vsView in vsViews)
            {
                if (TryResolveWpfTextView(vsView, out _, out IWpfTextView wpfView, out _)
                    && wpfView != null
                    && !wpfView.IsClosed
                    && !views.Contains(wpfView))
                {
                    views.Add(wpfView);
                }
            }

            return views;
        }

        private bool TryGetFrameCodeWindow(IVsWindowFrame frame, out IVsCodeWindow codeWindow)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            codeWindow = null;
            foreach (int propertyId in new[] { (int)__VSFPROPID.VSFPROPID_DocView, (int)__VSFPROPID.VSFPROPID_DocData })
            {
                try
                {
                    if (frame.GetProperty(propertyId, out object value) != VSConstants.S_OK || value == null)
                    {
                        continue;
                    }

                    if (value is IVsCodeWindow direct)
                    {
                        codeWindow = direct;
                        return true;
                    }

                    if (TryFindObjectGraphInstance(value, out codeWindow, maxDepth: 5, maxNodes: 400))
                    {
                        return true;
                    }
                }
                catch
                {
                    // Ignore and try the next property.
                }
            }

            return false;
        }

        /// <summary>
        /// Owns the tint for one text view: remembers the theme background, re-applies the blend when
        /// the editor resets its background (theme or Fonts and Colors change), and restores on request.
        /// </summary>
        private sealed class EditorTintController
        {
            private readonly IWpfTextView view;
            private SolidColorBrush baseBrush;
            private Brush appliedBrush;
            private Color? tintColor;
            private int strength;
            private bool isApplying;
            private bool reapplyQueued;
            private bool attached;

            public EditorTintController(IWpfTextView view)
            {
                this.view = view;
            }

            public bool Apply(Color tint, int strengthPercent)
            {
                if (view.IsClosed)
                {
                    return false;
                }

                if (baseBrush == null)
                {
                    if (!(view.Background is SolidColorBrush themeBrush))
                    {
                        return false;
                    }

                    baseBrush = themeBrush;
                }

                Attach();
                tintColor = tint;
                strength = strengthPercent;
                return ApplyDesired();
            }

            public void Restore()
            {
                tintColor = null;
                Detach();

                if (!view.IsClosed && baseBrush != null && IsAppliedBrush(view.Background))
                {
                    isApplying = true;
                    try
                    {
                        view.Background = baseBrush;
                    }
                    finally
                    {
                        isApplying = false;
                    }
                }

                baseBrush = null;
                appliedBrush = null;
            }

            private bool ApplyDesired()
            {
                if (!tintColor.HasValue || baseBrush == null || view.IsClosed)
                {
                    return false;
                }

                Color baseColor = baseBrush.Color;
                Color tint = tintColor.Value;
                int blended = EditorTint.Blend(
                    (baseColor.R << 16) | (baseColor.G << 8) | baseColor.B,
                    (tint.R << 16) | (tint.G << 8) | tint.B,
                    strength);
                Color target = Color.FromArgb(baseColor.A, (byte)(blended >> 16), (byte)(blended >> 8), (byte)blended);

                if (appliedBrush is SolidColorBrush current && current.Color == target && IsAppliedBrush(view.Background))
                {
                    return true;
                }

                var brush = new SolidColorBrush(target);
                brush.Freeze();

                isApplying = true;
                try
                {
                    appliedBrush = brush;
                    view.Background = brush;
                }
                finally
                {
                    isApplying = false;
                }

                return true;
            }

            private void OnBackgroundBrushChanged(object sender, BackgroundBrushChangedEventArgs e)
            {
                if (isApplying || !tintColor.HasValue || IsAppliedBrush(e.NewBackgroundBrush))
                {
                    return;
                }

                // The editor pushed a new background (theme switch or Fonts and Colors change): blend over that instead.
                if (e.NewBackgroundBrush is SolidColorBrush newBase)
                {
                    baseBrush = newBase;
                }

                if (reapplyQueued)
                {
                    return;
                }

                reapplyQueued = true;
                view.VisualElement.Dispatcher.BeginInvoke(new Action(() =>
                {
                    reapplyQueued = false;
                    ApplyDesired();
                }), DispatcherPriority.Background);
            }

            private bool IsAppliedBrush(Brush brush)
            {
                if (brush == null || appliedBrush == null)
                {
                    return false;
                }

                return ReferenceEquals(brush, appliedBrush)
                    || (brush is SolidColorBrush solid && appliedBrush is SolidColorBrush applied && solid.Color == applied.Color);
            }

            private void OnClosed(object sender, EventArgs e)
            {
                Detach();
                editorTintControllers.Remove(view);
            }

            private void Attach()
            {
                if (attached)
                {
                    return;
                }

                view.BackgroundBrushChanged += OnBackgroundBrushChanged;
                view.Closed += OnClosed;
                attached = true;
            }

            private void Detach()
            {
                if (!attached)
                {
                    return;
                }

                view.BackgroundBrushChanged -= OnBackgroundBrushChanged;
                view.Closed -= OnClosed;
                attached = false;
            }
        }
    }
}
