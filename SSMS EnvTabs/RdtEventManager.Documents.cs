using Microsoft.VisualStudio;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace SSMS_EnvTabs
{
    internal sealed partial class RdtEventManager
    {
        private const string IndicatorTypeName = "Indicator";
        private const string IndicatorTypeFullName = "Microsoft.VisualStudio.Shell.Controls.Indicator";
        private static readonly ConditionalWeakTable<System.Windows.Forms.StatusStrip, StatusStripColorController> statusStripColorControllers = new ConditionalWeakTable<System.Windows.Forms.StatusStrip, StatusStripColorController>();
        private static readonly ConditionalWeakTable<DependencyObject, IndicatorBrushSnapshot> originalIndicatorBrushState = new ConditionalWeakTable<DependencyObject, IndicatorBrushSnapshot>();

        private IComponentModel componentModel;
        private IVsEditorAdaptersFactoryService editorAdaptersFactoryService;

        private bool TryApplyLineIndicatorColor(uint docCookie, IVsWindowFrame frame, string moniker, IReadOnlyList<TabRuleMatcher.CompiledRule> rules, IReadOnlyList<TabRuleMatcher.CompiledManualRule> manualRules, TabGroupSettings settings)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (frame == null)
            {
                return true;
            }

            EnvTabsLog.Info($"LineIndicator: attempt cookie={docCookie} moniker='{moniker}' enabled={settings?.InitialLineIndicatorColor == true}");

            if (settings?.EnableAutoColor != true)
            {
                // Auto-color globally disabled: do not modify any indicator visuals.
                return true;
            }

            if (!TryGetConnectionInfo(frame, out string server, out string database) || string.IsNullOrWhiteSpace(server))
            {
                EnvTabsLog.Info($"LineIndicator: skipped (no connection info) cookie={docCookie} moniker='{moniker}'");
                return true;
            }

            var manualMatch = TabRuleMatcher.MatchManual(manualRules, moniker);
            var matchedRule = TabRuleMatcher.MatchRule(rules, server, database);

            // Per-rule setting overrides global; null means use global default.
            bool lineIndicatorEnabled = matchedRule?.EnableLineIndicatorColor ?? settings.InitialLineIndicatorColor;
            if (!lineIndicatorEnabled)
            {
                EnvTabsLog.Info($"LineIndicator: disabled for rule cookie={docCookie} server='{server}' db='{database}'");
                TryRestoreLineIndicatorColor(frame);
                return true;
            }

            int? colorIndex = manualMatch?.ColorIndex ?? matchedRule?.ColorIndex;
            if (!colorIndex.HasValue || colorIndex.Value < 0 || colorIndex.Value >= ColorPalette.Hex.Length)
            {
                EnvTabsLog.Info($"LineIndicator: skipped (no valid color index) cookie={docCookie} server='{server}' db='{database}'");
                TryRestoreLineIndicatorColor(frame);
                return true;
            }

            if (!TryFindLineIndicatorElements(docCookie, frame, out List<DependencyObject> indicators) || indicators.Count == 0)
            {
                EnvTabsLog.Info($"LineIndicator: editor indicator not ready cookie={docCookie} server='{server}' db='{database}' colorIndex={colorIndex.Value}");
                return false;
            }

            Brush brush = BuildIndicatorBrush(colorIndex.Value);
            int appliedCount = 0;
            foreach (DependencyObject indicator in indicators)
            {
                if (TrySetIndicatorBackground(indicator, brush))
                {
                    appliedCount++;
                }
            }

            EnvTabsLog.Info($"LineIndicator: target type='{DescribeIndicator(indicators[0])}'");
            EnvTabsLog.Info($"LineIndicator: {(appliedCount > 0 ? "applied" : "failed")} cookie={docCookie} server='{server}' db='{database}' colorIndex={colorIndex.Value} indicators={indicators.Count} updated={appliedCount}");
            return appliedCount > 0;
        }

        private void TryApplyStatusBarColor(uint docCookie, IVsWindowFrame frame, string moniker, IReadOnlyList<TabRuleMatcher.CompiledRule> rules, IReadOnlyList<TabRuleMatcher.CompiledManualRule> manualRules, TabGroupSettings settings)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (frame == null || settings?.EnableAutoColor != true)
            {
                return;
            }

            if (!TryGetConnectionInfo(frame, out string server, out string database) || string.IsNullOrWhiteSpace(server))
            {
                return;
            }

            var manualMatch = TabRuleMatcher.MatchManual(manualRules, moniker);
            var matchedRule = TabRuleMatcher.MatchRule(rules, server, database);

            // Per-rule setting overrides global; null means use global default.
            bool statusBarEnabled = matchedRule?.EnableStatusBarColor ?? settings.InitialStatusBarColor;
            if (!statusBarEnabled)
            {
                TryRestoreQueryEditorStatusBarColor(frame);
                return;
            }

            int? colorIndex = manualMatch?.ColorIndex ?? matchedRule?.ColorIndex;
            if (!colorIndex.HasValue || colorIndex.Value < 0 || colorIndex.Value >= ColorPalette.Hex.Length)
            {
                TryRestoreQueryEditorStatusBarColor(frame);
                return;
            }

            var color = System.Drawing.ColorTranslator.FromHtml(ColorPalette.Hex[colorIndex.Value]);
            if (TrySetQueryEditorStatusBarColor(frame, color, out string docViewTarget))
            {
                EnvTabsLog.Info($"StatusBarColor: applied cookie={docCookie} server='{server}' db='{database}' colorIndex={colorIndex.Value} mode=statusbar-manager target='{docViewTarget}'");
                return;
            }

            EnvTabsLog.Info($"StatusBarColor: target not found cookie={docCookie} server='{server}' db='{database}'");
        }

        private static bool TrySetQueryEditorStatusBarColor(IVsWindowFrame frame, System.Drawing.Color color, out string targetDescription)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            targetDescription = null;
            if (frame == null)
            {
                return false;
            }

            if (frame.GetProperty((int)__VSFPROPID.VSFPROPID_DocView, out object docView) != VSConstants.S_OK || docView == null)
            {
                return false;
            }

            if (!TryGetReflectionMemberValue(docView, "StatusBarManager", out object statusBarManager)
                && !TryGetReflectionMemberValue(docView, "m_statusBarManager", out statusBarManager)
                && !TryGetReflectionMemberValue(docView, "statusBarManager", out statusBarManager))
            {
                return false;
            }

            bool nativeApplied = TrySetNativeServerBackground(statusBarManager, color);
            bool stripApplied = false;
            string statusStripTarget = "<not-found>";
            if (TryGetStatusStrip(statusBarManager, out System.Windows.Forms.StatusStrip statusStrip))
            {
                GetStatusStripColorController(statusStrip).ApplyColor(color);
                stripApplied = true;
                statusStripTarget = statusStrip.GetType().FullName;
            }

            if (!nativeApplied && !stripApplied)
            {
                return false;
            }

            targetDescription = "DocView.StatusBarManager; nativeSetServerBackground=" + nativeApplied
                + "; statusStrip=" + statusStripTarget;
            return true;
        }

        private static void TryRestoreQueryEditorStatusBarColor(IVsWindowFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!TryGetStatusBarManager(frame, out object statusBarManager))
            {
                return;
            }

            if (!TryGetStatusStrip(statusBarManager, out System.Windows.Forms.StatusStrip statusStrip))
            {
                return;
            }

            if (statusStripColorControllers.TryGetValue(statusStrip, out StatusStripColorController controller))
            {
                controller.Restore();
                statusStripColorControllers.Remove(statusStrip);
            }
        }

        private static bool TryGetStatusBarManager(IVsWindowFrame frame, out object statusBarManager)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            statusBarManager = null;
            if (frame == null)
            {
                return false;
            }

            if (frame.GetProperty((int)__VSFPROPID.VSFPROPID_DocView, out object docView) != VSConstants.S_OK || docView == null)
            {
                return false;
            }

            return TryGetReflectionMemberValue(docView, "StatusBarManager", out statusBarManager)
                || TryGetReflectionMemberValue(docView, "m_statusBarManager", out statusBarManager)
                || TryGetReflectionMemberValue(docView, "statusBarManager", out statusBarManager);
        }

        private static bool TrySetNativeServerBackground(object statusBarManager, System.Drawing.Color color)
        {
            if (statusBarManager == null)
            {
                return false;
            }

            try
            {
                MethodInfo method = statusBarManager.GetType().GetMethod(
                    "SetServerBackground",
                    BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic,
                    null,
                    new[] { typeof(System.Drawing.Color) },
                    null);

                if (method == null)
                {
                    return false;
                }

                method.Invoke(statusBarManager, new object[] { color });
                return true;
            }
            catch (Exception ex)
            {
                EnvTabsLog.Verbose($"StatusBarColor: native SetServerBackground failed: {ex.Message}");
                return false;
            }
        }

        private static bool TryGetStatusStrip(object statusBarManager, out System.Windows.Forms.StatusStrip statusStrip)
        {
            statusStrip = null;

            if (statusBarManager == null)
            {
                return false;
            }

            if (!TryGetReflectionMemberValue(statusBarManager, "statusStrip", out object statusStripObject)
                && !TryGetReflectionMemberValue(statusBarManager, "StatusStrip", out statusStripObject))
            {
                return false;
            }

            statusStrip = statusStripObject as System.Windows.Forms.StatusStrip;
            return statusStrip != null && !statusStrip.IsDisposed;
        }

        private static StatusStripColorController GetStatusStripColorController(System.Windows.Forms.StatusStrip statusStrip)
        {
            if (!statusStripColorControllers.TryGetValue(statusStrip, out StatusStripColorController controller))
            {
                controller = new StatusStripColorController(statusStrip);
                statusStripColorControllers.Add(statusStrip, controller);
            }

            return controller;
        }

        private static bool TryGetReflectionMemberValue(object source, string memberName, out object value)
        {
            value = null;
            if (source == null || string.IsNullOrWhiteSpace(memberName))
            {
                return false;
            }

            const BindingFlags Flags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
            Type type = source.GetType();

            FieldInfo field = type.GetField(memberName, Flags);
            if (field != null)
            {
                value = field.GetValue(source);
                return value != null;
            }

            PropertyInfo property = type.GetProperty(memberName, Flags);
            if (property != null && property.CanRead && property.GetIndexParameters().Length == 0)
            {
                value = property.GetValue(source, null);
                return value != null;
            }

            return false;
        }

        private static IEnumerable<object> ReflectChildren(object source)
        {
            if (source == null)
            {
                yield break;
            }

            foreach ((string _, object value) in EnumerateProbeMembers(source))
            {
                if (value != null)
                {
                    yield return value;
                }
            }
        }

        private sealed class StatusStripColorController
        {
            private readonly System.Windows.Forms.StatusStrip statusStrip;
            private readonly Dictionary<System.Windows.Forms.ToolStripItem, ToolStripItemColorSnapshot> itemSnapshots = new Dictionary<System.Windows.Forms.ToolStripItem, ToolStripItemColorSnapshot>();
            private System.Drawing.Color originalBackColor;
            private System.Drawing.Color originalForeColor;
            private bool hasSnapshot;
            private bool isApplying;
            private System.Drawing.Color? desiredColor;

            public StatusStripColorController(System.Windows.Forms.StatusStrip statusStrip)
            {
                this.statusStrip = statusStrip;

                if (statusStrip != null)
                {
                    statusStrip.BackColorChanged += OnBackColorChanged;
                    statusStrip.ItemAdded += OnItemAdded;
                    statusStrip.Paint += OnPaint;
                    statusStrip.Disposed += OnDisposed;
                }
            }

            public void ApplyColor(System.Drawing.Color backColor)
            {
                if (statusStrip == null || statusStrip.IsDisposed)
                {
                    return;
                }

                CaptureSnapshotIfNeeded();
                desiredColor = backColor;
                ApplyDesiredColor();
            }

            public void Restore()
            {
                desiredColor = null;

                if (statusStrip == null || statusStrip.IsDisposed)
                {
                    return;
                }

                if (!hasSnapshot)
                {
                    Dispose();
                    return;
                }

                isApplying = true;
                statusStrip.SuspendLayout();

                try
                {
                    statusStrip.BackColor = originalBackColor;
                    statusStrip.ForeColor = originalForeColor;

                    foreach (var snapshot in itemSnapshots)
                    {
                        if (snapshot.Key == null || snapshot.Key.IsDisposed)
                        {
                            continue;
                        }

                        snapshot.Key.BackColor = snapshot.Value.BackColor;
                        snapshot.Key.ForeColor = snapshot.Value.ForeColor;
                    }
                }
                finally
                {
                    statusStrip.ResumeLayout(true);
                    isApplying = false;
                }

                Dispose();
                statusStrip.Invalidate(true);
                statusStrip.Refresh();
            }

            private void ApplyDesiredColor()
            {
                if (!desiredColor.HasValue || statusStrip == null || statusStrip.IsDisposed)
                {
                    return;
                }

                System.Drawing.Color backColor = desiredColor.Value;
                System.Drawing.Color foreColor = GetReadableStatusTextColor(backColor);

                isApplying = true;
                statusStrip.SuspendLayout();

                try
                {
                    statusStrip.BackColor = backColor;
                    statusStrip.ForeColor = foreColor;

                    foreach (System.Windows.Forms.ToolStripItem item in statusStrip.Items)
                    {
                        ApplyItemColor(item, backColor, foreColor);
                    }
                }
                finally
                {
                    statusStrip.ResumeLayout(true);
                    isApplying = false;
                }

                statusStrip.Invalidate(true);
                statusStrip.Refresh();
            }

            private void ApplyItemColor(System.Windows.Forms.ToolStripItem item, System.Drawing.Color backColor, System.Drawing.Color foreColor)
            {
                if (item == null || item.IsDisposed)
                {
                    return;
                }

                CaptureItemSnapshotIfNeeded(item);
                item.BackColor = backColor;

                if (!IsSubtleStatusStripSeparator(item))
                {
                    item.ForeColor = foreColor;
                }
            }

            private void CaptureSnapshotIfNeeded()
            {
                if (hasSnapshot || statusStrip == null || statusStrip.IsDisposed)
                {
                    return;
                }

                originalBackColor = statusStrip.BackColor;
                originalForeColor = statusStrip.ForeColor;
                itemSnapshots.Clear();

                foreach (System.Windows.Forms.ToolStripItem item in statusStrip.Items)
                {
                    CaptureItemSnapshotIfNeeded(item);
                }

                hasSnapshot = true;
            }

            private void CaptureItemSnapshotIfNeeded(System.Windows.Forms.ToolStripItem item)
            {
                if (item == null || itemSnapshots.ContainsKey(item))
                {
                    return;
                }

                itemSnapshots[item] = new ToolStripItemColorSnapshot
                {
                    BackColor = item.BackColor,
                    ForeColor = item.ForeColor
                };
            }

            private void OnBackColorChanged(object sender, EventArgs e)
            {
                if (isApplying || !desiredColor.HasValue || statusStrip == null || statusStrip.IsDisposed)
                {
                    return;
                }

                if (statusStrip.BackColor.ToArgb() == desiredColor.Value.ToArgb())
                {
                    return;
                }

                QueueApplyDesiredColor();
            }

            private void OnItemAdded(object sender, System.Windows.Forms.ToolStripItemEventArgs e)
            {
                if (!desiredColor.HasValue || e?.Item == null)
                {
                    return;
                }

                ApplyItemColor(e.Item, desiredColor.Value, GetReadableStatusTextColor(desiredColor.Value));
            }

            private void OnPaint(object sender, System.Windows.Forms.PaintEventArgs e)
            {
                if (isApplying || !desiredColor.HasValue || statusStrip == null || statusStrip.IsDisposed)
                {
                    return;
                }

                if (!StatusStripColorMatchesDesired())
                {
                    QueueApplyDesiredColor();
                }
            }

            private void QueueApplyDesiredColor()
            {
                try
                {
                    if (statusStrip.IsHandleCreated)
                    {
                        statusStrip.BeginInvoke((System.Windows.Forms.MethodInvoker)ApplyDesiredColor);
                    }
                    else
                    {
                        ApplyDesiredColor();
                    }
                }
                catch
                {
                    ApplyDesiredColor();
                }
            }

            private void OnDisposed(object sender, EventArgs e)
            {
                Dispose();
            }

            private void Dispose()
            {
                if (statusStrip == null)
                {
                    return;
                }

                statusStrip.BackColorChanged -= OnBackColorChanged;
                statusStrip.ItemAdded -= OnItemAdded;
                statusStrip.Paint -= OnPaint;
                statusStrip.Disposed -= OnDisposed;
            }

            private bool StatusStripColorMatchesDesired()
            {
                if (!desiredColor.HasValue || statusStrip.BackColor.ToArgb() != desiredColor.Value.ToArgb())
                {
                    return false;
                }

                foreach (System.Windows.Forms.ToolStripItem item in statusStrip.Items)
                {
                    if (item != null && !item.IsDisposed && item.BackColor.ToArgb() != desiredColor.Value.ToArgb())
                    {
                        return false;
                    }
                }

                return true;
            }
        }

        private sealed class ToolStripItemColorSnapshot
        {
            public System.Drawing.Color BackColor { get; set; }
            public System.Drawing.Color ForeColor { get; set; }
        }

        private static System.Drawing.Color GetReadableStatusTextColor(System.Drawing.Color backColor)
        {
            double contrastWithBlack = GetContrastRatio(backColor, System.Drawing.Color.Black);
            double contrastWithWhite = GetContrastRatio(backColor, System.Drawing.Color.White);
            return contrastWithBlack >= contrastWithWhite ? System.Drawing.Color.Black : System.Drawing.Color.White;
        }

        private static double GetContrastRatio(System.Drawing.Color a, System.Drawing.Color b)
        {
            double luminanceA = GetRelativeLuminance(a);
            double luminanceB = GetRelativeLuminance(b);
            double lighter = Math.Max(luminanceA, luminanceB);
            double darker = Math.Min(luminanceA, luminanceB);
            return (lighter + 0.05) / (darker + 0.05);
        }

        private static double GetRelativeLuminance(System.Drawing.Color color)
        {
            double r = GetLinearChannel(color.R / 255.0);
            double g = GetLinearChannel(color.G / 255.0);
            double b = GetLinearChannel(color.B / 255.0);
            return (0.2126 * r) + (0.7152 * g) + (0.0722 * b);
        }

        private static double GetLinearChannel(double channel)
        {
            return channel <= 0.03928
                ? channel / 12.92
                : Math.Pow((channel + 0.055) / 1.055, 2.4);
        }

        private static bool IsSubtleStatusStripSeparator(System.Windows.Forms.ToolStripItem item)
        {
            if (item == null)
            {
                return false;
            }

            if (item is System.Windows.Forms.ToolStripSeparator)
            {
                return true;
            }

            string probe = ((item.Name ?? string.Empty)
                + " "
                + (item.AccessibleName ?? string.Empty)
                + " "
                + (item.AccessibleDescription ?? string.Empty)
                + " "
                + (item.GetType().FullName ?? string.Empty)).ToLowerInvariant();

            if (probe.IndexOf("separator", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            string text = (item.Text ?? string.Empty).Trim();
            return text.Length == 1 && IsSeparatorGlyph(text[0]);
        }

        private static bool IsSeparatorGlyph(char value)
        {
            switch (value)
            {
                case '|':
                case '\u00A6':
                case '\u2502':
                case '\u2503':
                case '\u2506':
                case '\u250A':
                    return true;
                default:
                    return false;
            }
        }

        private sealed class IndicatorBrushSnapshot
        {
            private readonly List<IndicatorBrushPropertySnapshot> properties = new List<IndicatorBrushPropertySnapshot>();

            public void CaptureFrom(DependencyObject target)
            {
                properties.Clear();

                foreach (string propertyName in new[] { "Background", "BorderBrush", "Foreground", "Fill", "Stroke" })
                {
                    DependencyProperty dependencyProperty = FindDependencyProperty(target.GetType(), propertyName);
                    if (dependencyProperty != null)
                    {
                        properties.Add(new IndicatorBrushPropertySnapshot
                        {
                            DependencyProperty = dependencyProperty,
                            OriginalLocalValue = target.ReadLocalValue(dependencyProperty)
                        });
                    }
                }
            }

            public void RestoreTo(DependencyObject target)
            {
                foreach (IndicatorBrushPropertySnapshot property in properties)
                {
                    if (property.DependencyProperty == null)
                    {
                        continue;
                    }

                    if (property.OriginalLocalValue == DependencyProperty.UnsetValue)
                    {
                        target.ClearValue(property.DependencyProperty);
                    }
                    else
                    {
                        target.SetValue(property.DependencyProperty, property.OriginalLocalValue);
                    }
                }
            }
        }

        private sealed class IndicatorBrushPropertySnapshot
        {
            public DependencyProperty DependencyProperty { get; set; }
            public object OriginalLocalValue { get; set; }
        }

        private static DependencyProperty FindDependencyProperty(Type type, string propertyName)
        {
            if (type == null || string.IsNullOrWhiteSpace(propertyName))
            {
                return null;
            }

            for (Type current = type; current != null; current = current.BaseType)
            {
                FieldInfo field = current.GetField(
                    propertyName + "Property",
                    BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy);

                if (field?.GetValue(null) is DependencyProperty dependencyProperty)
                {
                    return dependencyProperty;
                }
            }

            return null;
        }

        private sealed class ReferenceEqualityComparer : IEqualityComparer<object>
        {
            public static readonly ReferenceEqualityComparer Instance = new ReferenceEqualityComparer();
            private ReferenceEqualityComparer() { }
            public new bool Equals(object x, object y) => ReferenceEquals(x, y);
            public int GetHashCode(object obj) => obj == null ? 0 : RuntimeHelpers.GetHashCode(obj);
        }

        private static Brush BuildIndicatorBrush(int colorIndex)
        {
            if (colorIndex < 0 || colorIndex >= ColorPalette.Hex.Length)
            {
                return null;
            }

            var brush = (SolidColorBrush)new BrushConverter().ConvertFromString(ColorPalette.Hex[colorIndex]);
            if (brush != null && brush.CanFreeze)
            {
                brush.Freeze();
            }

            return brush;
        }

        private static bool TrySetIndicatorBackground(DependencyObject indicator, Brush brush)
        {
            if (indicator == null || brush == null)
            {
                return false;
            }

            CaptureIndicatorBrushState(indicator);

            bool changed = false;

            if (indicator is Control control)
            {
                control.Background = brush;
                changed = true;
            }

            if (indicator is Border border)
            {
                border.Background = brush;
                changed = true;
            }

            if (indicator is Panel panel)
            {
                panel.Background = brush;
                changed = true;
            }

            if (indicator is System.Windows.Shapes.Shape shape)
            {
                shape.Fill = brush;
                shape.Stroke = brush;
                changed = true;
            }

            if (TrySetBrushProperty(indicator, "Background", brush))
            {
                changed = true;
            }

            if (TrySetBrushProperty(indicator, "BorderBrush", brush))
            {
                changed = true;
            }

            if (TrySetBrushProperty(indicator, "Foreground", brush))
            {
                changed = true;
            }

            if (TrySetBrushProperty(indicator, "Fill", brush))
            {
                changed = true;
            }

            if (TrySetBrushProperty(indicator, "Stroke", brush))
            {
                changed = true;
            }

            // Some templates paint through a parent element, not the Indicator itself.
            DependencyObject parent = null;
            try
            {
                parent = VisualTreeHelper.GetParent(indicator);
            }
            catch
            {
                parent = null;
            }

            if (parent != null)
            {
                CaptureIndicatorBrushState(parent);

                if (TrySetBrushProperty(parent, "Background", brush)
                    || TrySetBrushProperty(parent, "BorderBrush", brush)
                    || TrySetBrushProperty(parent, "Fill", brush)
                    || TrySetBrushProperty(parent, "Stroke", brush))
                {
                    changed = true;
                }
            }

            return changed;
        }

        private void TryRestoreLineIndicatorColor(IVsWindowFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!TryFindLineIndicatorElements(0, frame, out List<DependencyObject> indicators) || indicators.Count == 0)
            {
                return;
            }

            foreach (DependencyObject indicator in indicators)
            {
                TryRestoreIndicatorBackground(indicator);
            }
        }

        private static void TryRestoreIndicatorBackground(DependencyObject indicator)
        {
            RestoreIndicatorBrushState(indicator);

            DependencyObject parent = null;
            try
            {
                parent = VisualTreeHelper.GetParent(indicator);
            }
            catch
            {
                parent = null;
            }

            if (parent != null)
            {
                RestoreIndicatorBrushState(parent);
            }
        }

        private static void CaptureIndicatorBrushState(DependencyObject target)
        {
            if (target == null || originalIndicatorBrushState.TryGetValue(target, out _))
            {
                return;
            }

            var snapshot = new IndicatorBrushSnapshot();
            snapshot.CaptureFrom(target);
            originalIndicatorBrushState.Add(target, snapshot);
        }

        private static void RestoreIndicatorBrushState(DependencyObject target)
        {
            if (target == null || !originalIndicatorBrushState.TryGetValue(target, out IndicatorBrushSnapshot snapshot))
            {
                return;
            }

            snapshot.RestoreTo(target);
            originalIndicatorBrushState.Remove(target);
        }

        private static bool TrySetBrushProperty(object target, string propertyName, Brush brush)
        {
            if (target == null || brush == null || string.IsNullOrWhiteSpace(propertyName))
            {
                return false;
            }

            try
            {
                var prop = target.GetType().GetProperty(propertyName);
                if (prop != null && prop.CanWrite && typeof(Brush).IsAssignableFrom(prop.PropertyType))
                {
                    prop.SetValue(target, brush, null);
                    return true;
                }
            }
            catch
            {
                // Ignore reflection setter failures.
            }

            return false;
        }

        private static string DescribeIndicator(DependencyObject indicator)
        {
            if (!(indicator is FrameworkElement fe))
            {
                return indicator?.GetType().FullName ?? "<null>";
            }

            string type = fe.GetType().FullName ?? fe.GetType().Name;
            string name = fe.Name ?? string.Empty;
            return $"{type} name='{name}' w={fe.ActualWidth:0.##} h={fe.ActualHeight:0.##}";
        }

        private static bool IsIndicatorControl(DependencyObject candidate)
        {
            if (candidate == null)
            {
                return false;
            }

            string typeFullName = candidate.GetType().FullName ?? string.Empty;
            string typeName = candidate.GetType().Name ?? string.Empty;
            return string.Equals(typeFullName, IndicatorTypeFullName, StringComparison.OrdinalIgnoreCase)
                || string.Equals(typeName, IndicatorTypeName, StringComparison.OrdinalIgnoreCase);
        }

        private bool TryFindLineIndicatorElements(uint docCookie, IVsWindowFrame frame, out List<DependencyObject> indicators)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            indicators = new List<DependencyObject>();
            if (frame == null)
            {
                return false;
            }

            try
            {
                if (docCookie != 0
                    && lineIndicatorPreferEditorScopeByCookie.TryGetValue(docCookie, out bool preferEditorScope)
                    && preferEditorScope)
                {
                    if (!TryResolveEditorScopedSearchRoots(frame, out List<DependencyObject> cachedRoots, out string cachedResolutionSource))
                    {
                        return false;
                    }

                    EnvTabsLog.Info($"LineIndicator: editor scope resolved via cached-fallback({cachedResolutionSource}) roots={cachedRoots.Count}");
                    CollectIndicatorControls(cachedRoots, indicators);
                    if (indicators.Count == 0)
                    {
                        lineIndicatorPreferEditorScopeByCookie.Remove(docCookie);
                    }
                    else
                    {
                        EnvTabsLog.Info($"LineIndicator: editor scope discovered {indicators.Count} gutter candidate(s).");
                        return true;
                    }
                }

                if (!TryResolveIndicatorSearchRoots(frame, out List<DependencyObject> roots, out string resolutionSource))
                {
                    return false;
                }

                EnvTabsLog.Info($"LineIndicator: editor scope resolved via {resolutionSource} roots={roots.Count}");

                CollectIndicatorControls(roots, indicators);

                if (indicators.Count == 0
                    && TryResolveEditorScopedSearchRoots(frame, out List<DependencyObject> fallbackRoots, out string fallbackResolutionSource))
                {
                    CollectIndicatorControls(fallbackRoots, indicators);
                    if (indicators.Count > 0)
                    {
                        if (docCookie != 0)
                        {
                            lineIndicatorPreferEditorScopeByCookie[docCookie] = true;
                        }

                        resolutionSource = resolutionSource + " -> fallback(" + fallbackResolutionSource + ")";
                        EnvTabsLog.Info($"LineIndicator: exact indicator margin was empty; fallback editor scope found candidates via {fallbackResolutionSource} roots={fallbackRoots.Count}");
                    }
                }

                if (indicators.Count > 0 && docCookie != 0)
                {
                    lineIndicatorPreferEditorScopeByCookie.Remove(docCookie);
                }

                if (indicators.Count == 0)
                {
                    EnvTabsLog.Info($"LineIndicator: editor scope resolved but no gutter candidates were found via {resolutionSource}.");
                    return false;
                }

                EnvTabsLog.Info($"LineIndicator: editor scope discovered {indicators.Count} gutter candidate(s).");
                return true;
            }
            catch
            {
                return false;
            }
        }

        private bool TryResolveEditorScopedSearchRoots(IVsWindowFrame frame, out List<DependencyObject> roots, out string resolutionSource)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            roots = new List<DependencyObject>();
            resolutionSource = null;

            if (!TryGetActiveVsTextView(frame, out IVsTextView vsTextView, out string textViewSource))
            {
                resolutionSource = textViewSource ?? "no-active-vstextview";
                return false;
            }

            if (!TryResolveWpfTextView(vsTextView, out IWpfTextViewHost textViewHost, out IWpfTextView wpfTextView, out string wpfSource))
            {
                resolutionSource = $"{textViewSource}; {wpfSource}";
                return false;
            }

            TryAddEditorScopeRoots(textViewHost, wpfTextView, roots);
            if (roots.Count == 0)
            {
                resolutionSource = $"{textViewSource}; {wpfSource}; no-editor-roots";
                return false;
            }

            resolutionSource = $"{textViewSource}; {wpfSource}";
            return true;
        }

        private bool TryGetActiveVsTextView(IVsWindowFrame frame, out IVsTextView textView, out string source)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            textView = null;
            source = null;

            if (frame == null)
            {
                source = "frame-null";
                return false;
            }

            if (TryGetVsTextViewFromFrameProperty(frame, (int)__VSFPROPID.VSFPROPID_DocView, "DocView", out textView, out source))
            {
                return true;
            }

            if (TryGetVsTextViewFromFrameProperty(frame, (int)__VSFPROPID.VSFPROPID_DocData, "DocData", out textView, out source))
            {
                return true;
            }

            source = "no-vstextview-from-frame";
            return false;
        }

        private bool TryResolveIndicatorSearchRoots(IVsWindowFrame frame, out List<DependencyObject> roots, out string resolutionSource)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            roots = new List<DependencyObject>();
            resolutionSource = null;

            if (!TryGetActiveVsTextView(frame, out IVsTextView vsTextView, out string textViewSource))
            {
                resolutionSource = textViewSource ?? "no-active-vstextview";
                return false;
            }

            if (!TryResolveWpfTextView(vsTextView, out IWpfTextViewHost textViewHost, out IWpfTextView wpfTextView, out string wpfSource))
            {
                resolutionSource = $"{textViewSource}; {wpfSource}";
                return false;
            }

            if (textViewHost != null)
            {
                try
                {
                    if (textViewHost.GetTextViewMargin("Indicator") is IWpfTextViewMargin indicatorMargin)
                    {
                        TryAddUniqueRoot(indicatorMargin.VisualElement, roots);
                    }
                }
                catch
                {
                    // Ignore margin access failures and fall back to text view visuals.
                }
            }

            if (roots.Count == 0 && wpfTextView != null)
            {
                TryAddUniqueRoot(wpfTextView.VisualElement, roots);
            }

            if (roots.Count == 0)
            {
                resolutionSource = $"{textViewSource}; {wpfSource}; no-indicator-roots";
                return false;
            }

            resolutionSource = $"{textViewSource}; {wpfSource}; exact-indicator-margin";
            return true;
        }

        private bool TryGetVsTextViewFromFrameProperty(IVsWindowFrame frame, int propertyId, string propertyLabel, out IVsTextView textView, out string source)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            textView = null;
            source = null;

            try
            {
                if (frame.GetProperty(propertyId, out object propertyValue) != VSConstants.S_OK || propertyValue == null)
                {
                    source = propertyLabel + ":missing";
                    return false;
                }

                if (TryExtractVsTextView(propertyValue, out textView, out string extractionSource))
                {
                    source = propertyLabel + ":" + extractionSource;
                    return true;
                }

                source = propertyLabel + ":unsupported-type=" + propertyValue.GetType().FullName;
                return false;
            }
            catch (Exception ex)
            {
                source = propertyLabel + ":exception=" + ex.Message;
                return false;
            }
        }

        private bool TryExtractVsTextView(object candidate, out IVsTextView textView, out string source)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            textView = null;
            source = null;

            if (candidate == null)
            {
                source = "candidate-null";
                return false;
            }

            if (candidate is IVsTextView directView)
            {
                textView = directView;
                source = "direct-ivstextview";
                return true;
            }

            if (candidate is IVsCodeWindow directCodeWindow && TryGetTextViewFromCodeWindow(directCodeWindow, out textView))
            {
                source = "direct-ivscodewindow";
                return true;
            }

            if (TryFindObjectGraphInstance(candidate, out IVsTextView graphView, maxDepth: 5, maxNodes: 400))
            {
                textView = graphView;
                source = "graph-ivstextview";
                return true;
            }

            if (TryFindObjectGraphInstance(candidate, out IVsCodeWindow graphCodeWindow, maxDepth: 5, maxNodes: 400)
                && TryGetTextViewFromCodeWindow(graphCodeWindow, out textView))
            {
                source = "graph-ivscodewindow";
                return true;
            }

            source = "no-textview-in-graph";
            return false;
        }

        private static bool TryGetTextViewFromCodeWindow(IVsCodeWindow codeWindow, out IVsTextView textView)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            textView = null;
            if (codeWindow == null)
            {
                return false;
            }

            try
            {
                if (ErrorHandler.Succeeded(codeWindow.GetLastActiveView(out textView)) && textView != null)
                {
                    return true;
                }
            }
            catch
            {
                // Ignore and fall back.
            }

            try
            {
                if (ErrorHandler.Succeeded(codeWindow.GetPrimaryView(out textView)) && textView != null)
                {
                    return true;
                }
            }
            catch
            {
                // Ignore and fall back.
            }

            try
            {
                if (ErrorHandler.Succeeded(codeWindow.GetSecondaryView(out textView)) && textView != null)
                {
                    return true;
                }
            }
            catch
            {
                // Ignore failures.
            }

            return false;
        }

        private bool TryResolveWpfTextView(IVsTextView vsTextView, out IWpfTextViewHost textViewHost, out IWpfTextView wpfTextView, out string source)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            textViewHost = null;
            wpfTextView = null;
            source = null;

            if (vsTextView == null)
            {
                source = "vsTextView-null";
                return false;
            }

            if (vsTextView is IVsUserData userData)
            {
                try
                {
                    Guid hostGuid = Microsoft.VisualStudio.Editor.DefGuidList.guidIWpfTextViewHost;
                    if (ErrorHandler.Succeeded(userData.GetData(ref hostGuid, out object hostObject)))
                    {
                        textViewHost = hostObject as IWpfTextViewHost;
                        wpfTextView = textViewHost?.TextView;
                        if (textViewHost != null || wpfTextView != null)
                        {
                            source = "ivuserdata-host";
                            return true;
                        }
                    }
                }
                catch
                {
                    // Ignore and fall back to editor adapters.
                }
            }

            IVsEditorAdaptersFactoryService adaptersFactory = GetEditorAdaptersFactoryService();
            if (adaptersFactory != null)
            {
                try
                {
                    wpfTextView = adaptersFactory.GetWpfTextView(vsTextView);
                    if (wpfTextView != null)
                    {
                        source = "editor-adapters";
                        return true;
                    }
                }
                catch
                {
                    // Ignore adapter failures.
                }
            }

            source = "no-wpf-textview";
            return false;
        }

        private IVsEditorAdaptersFactoryService GetEditorAdaptersFactoryService()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (editorAdaptersFactoryService != null)
            {
                return editorAdaptersFactoryService;
            }

            componentModel = componentModel ?? Microsoft.VisualStudio.Shell.Package.GetGlobalService(typeof(SComponentModel)) as IComponentModel;
            editorAdaptersFactoryService = componentModel?.GetService<IVsEditorAdaptersFactoryService>();
            return editorAdaptersFactoryService;
        }

        private static void TryAddEditorScopeRoots(IWpfTextViewHost textViewHost, IWpfTextView wpfTextView, List<DependencyObject> roots)
        {
            if (roots == null)
            {
                return;
            }

            if (textViewHost != null)
            {
                TryAddUniqueRoot(textViewHost.HostControl, roots);

                foreach (string marginName in new[] { "Glyph", "LineNumber", "LeftSelection", "Indicator", "SpacerMargin" })
                {
                    try
                    {
                        if (textViewHost.GetTextViewMargin(marginName) is IWpfTextViewMargin margin)
                        {
                            TryAddUniqueRoot(margin.VisualElement, roots);
                        }
                    }
                    catch
                    {
                        // Ignore margin access failures.
                    }
                }
            }

            if (wpfTextView != null)
            {
                TryAddUniqueRoot(wpfTextView.VisualElement, roots);
            }
        }

        private static void TryAddUniqueRoot(DependencyObject root, List<DependencyObject> roots)
        {
            if (root != null && !roots.Contains(root))
            {
                roots.Add(root);
            }
        }

        private static void CollectIndicatorControls(IEnumerable<DependencyObject> roots, List<DependencyObject> indicators)
        {
            if (roots == null || indicators == null)
            {
                return;
            }

            var seen = new HashSet<DependencyObject>(indicators);
            foreach (DependencyObject root in roots)
            {
                foreach (DependencyObject found in FindIndicatorControls(root))
                {
                    if (found != null && seen.Add(found))
                    {
                        indicators.Add(found);
                    }
                }
            }
        }

        private static IEnumerable<DependencyObject> FindIndicatorControls(DependencyObject root)
        {
            var results = new List<DependencyObject>();
            if (root == null)
            {
                return results;
            }

            foreach (DependencyObject candidate in EnumerateVisualDescendants(root, maxNodes: 4000))
            {
                if (IsIndicatorControl(candidate) && !results.Contains(candidate))
                {
                    results.Add(candidate);
                }
            }

            return results;
        }

        private static bool TryFindObjectGraphInstance<T>(object rootObject, out T match, int maxDepth, int maxNodes) where T : class
        {
            match = null;
            if (rootObject == null || maxDepth < 0 || maxNodes <= 0)
            {
                return false;
            }

            var visited = new HashSet<object>(ReferenceEqualityComparer.Instance);
            var queue = new Queue<(object Obj, int Depth)>();
            queue.Enqueue((rootObject, 0));
            visited.Add(rootObject);

            int processed = 0;
            while (queue.Count > 0 && processed < maxNodes)
            {
                var current = queue.Dequeue();
                processed++;

                if (current.Obj is T typed)
                {
                    match = typed;
                    return true;
                }

                if (current.Depth >= maxDepth || current.Obj == null)
                {
                    continue;
                }

                if (current.Obj is string || current.Obj.GetType().IsPrimitive || current.Obj.GetType().IsEnum)
                {
                    continue;
                }

                foreach (object child in ReflectChildren(current.Obj))
                {
                    if (child == null || !visited.Add(child))
                    {
                        continue;
                    }

                    queue.Enqueue((child, current.Depth + 1));
                    if (queue.Count + processed >= maxNodes)
                    {
                        break;
                    }
                }
            }

            return false;
        }

        private static IEnumerable<DependencyObject> EnumerateVisualDescendants(DependencyObject root, int maxNodes)
        {
            if (root == null || maxNodes <= 0)
            {
                yield break;
            }

            var queue = new Queue<DependencyObject>();
            queue.Enqueue(root);
            int visited = 0;

            while (queue.Count > 0 && visited < maxNodes)
            {
                DependencyObject current = queue.Dequeue();
                if (current == null)
                {
                    continue;
                }

                visited++;
                yield return current;

                int childCount;
                try
                {
                    childCount = VisualTreeHelper.GetChildrenCount(current);
                }
                catch
                {
                    continue;
                }

                for (int i = 0; i < childCount; i++)
                {
                    DependencyObject child = null;
                    try
                    {
                        child = VisualTreeHelper.GetChild(current, i);
                    }
                    catch
                    {
                        child = null;
                    }

                    if (child != null)
                    {
                        queue.Enqueue(child);
                    }
                }
            }
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
                // 1. DocView/Reflection Method (Main attempt)
                if (TryPopulateFromDocView(frame, out string docViewServer, out string docViewDatabase))
                {
                    server = docViewServer;
                    database = docViewDatabase;
                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                EnvTabsLog.Error($"TryGetConnectionInfo exception: {ex.Message}");
                return false;
            }
        }

        private static bool TryPopulateFromDocView(IVsWindowFrame frame, out string server, out string database)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            server = null;
            database = null;

            try
            {
                if (frame.GetProperty((int)__VSFPROPID.VSFPROPID_DocView, out object docView) == VSConstants.S_OK && docView != null)
                {
                    // Inspect private field "m_connection" on the DocView (SqlScriptEditorControl)
                    // This is more reliable than tooltips which users can disable.
                    var flags = System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance;
                    var field = docView.GetType().GetField("m_connection", flags);
                    if (field != null)
                    {
                        var connection = field.GetValue(docView);
                        if (connection != null)
                        {
                            var connType = connection.GetType();
                            // properties: DataSource (=Server), Database
                            var propDataSource = connType.GetProperty("DataSource");
                            var propDatabase = connType.GetProperty("Database");
                            
                            string s = propDataSource?.GetValue(connection) as string;
                            string d = propDatabase?.GetValue(connection) as string;

                            if (!string.IsNullOrWhiteSpace(s))
                            {
                                server = s;
                                database = d;
                                return true;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Fail silently, fallback to other methods
                EnvTabsLog.Error($"Error accessing DocView connection info: {ex.Message}");
            }
            return false;
        }

        private static IEnumerable<(string Name, object Value)> EnumerateProbeMembers(object obj)
        {
            if (obj == null)
            {
                yield break;
            }

            const BindingFlags Flags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
            Type type = obj.GetType();

            FieldInfo[] fields;
            try
            {
                fields = type.GetFields(Flags);
            }
            catch
            {
                fields = Array.Empty<FieldInfo>();
            }

            foreach (FieldInfo field in fields)
            {
                object value = null;
                try
                {
                    value = field.GetValue(obj);
                }
                catch
                {
                    value = null;
                }

                if (value != null)
                {
                    yield return (field.Name, value);
                }
            }

            PropertyInfo[] props;
            try
            {
                props = type.GetProperties(Flags);
            }
            catch
            {
                props = Array.Empty<PropertyInfo>();
            }

            foreach (PropertyInfo prop in props)
            {
                if (!prop.CanRead || prop.GetIndexParameters().Length != 0)
                {
                    continue;
                }

                object value = null;
                try
                {
                    value = prop.GetValue(obj, null);
                }
                catch
                {
                    value = null;
                }

                if (value != null)
                {
                    yield return (prop.Name, value);
                }
            }
        }

        private static bool IsSimpleProbeType(Type type)
        {
            if (type == null)
            {
                return true;
            }

            if (type.IsPrimitive || type.IsEnum)
            {
                return true;
            }

            return type == typeof(string)
                || type == typeof(decimal)
                || type == typeof(DateTime)
                || type == typeof(TimeSpan)
                || type == typeof(Guid)
                || type == typeof(IntPtr)
                || type == typeof(UIntPtr)
                || type == typeof(System.Drawing.Color)
                || typeof(Brush).IsAssignableFrom(type);
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

                // Only consider SQL query documents
                if (!moniker.EndsWith(".sql", StringComparison.OrdinalIgnoreCase))
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

        private IVsWindowFrame TryGetFrameFromMoniker(string moniker)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (shellOpenDoc == null || string.IsNullOrWhiteSpace(moniker)) return null;

            // Main attempt: IsDocumentOpen
            try
            {
                Guid logicalView = Guid.Empty;
                uint[] itemid = new uint[1];
                if (shellOpenDoc.IsDocumentOpen(null, 0, moniker, ref logicalView, 0, out IVsUIHierarchy _, itemid, out IVsWindowFrame frame, out int isOpen) == VSConstants.S_OK && isOpen != 0)
                {
                    return frame;
                }
            }
            catch (Exception ex)
            {
                EnvTabsLog.Verbose($"TryGetFrameFromMoniker - Error: {ex.Message}");
                return null;
            }

            return null;
        }

        private static bool IsRenameEligible(string moniker, string caption)
        {
            if (string.IsNullOrWhiteSpace(moniker)) return false;

            // Check if it's a Temp file (New Query)
            if (IsTempFile(moniker))
            {
                // Verify it's a SQL file
                return moniker.EndsWith(".sql", StringComparison.OrdinalIgnoreCase);
            }

            // For saved files, we are eligible ONLY if the current caption matches the filename exactly.
            // This prevents overwriting our own renames or user-customized renames.
            try
            {
                string fileName = System.IO.Path.GetFileName(moniker);
                if (string.Equals(caption, fileName, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                EnvTabsLog.Verbose($"IsRenameEligible - File name parse failed: {ex.Message}");
            }

            return false;
        }

        public static bool IsTempFile(string path)
        {
            try
            {
                string tempPath = System.IO.Path.GetTempPath();
                if (path != null && path.StartsWith(tempPath, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            catch (Exception ex)
            {
                EnvTabsLog.Verbose($"IsTempFile - Temp path check failed: {ex.Message}");
            }
            return false;
        }

        private static bool IsSqlDocumentMoniker(string moniker)
        {
            if (string.IsNullOrWhiteSpace(moniker))
            {
                return false;
            }

            if (!Path.IsPathRooted(moniker))
            {
                return false;
            }

            return moniker.EndsWith(".sql", StringComparison.OrdinalIgnoreCase);
        }
    }
}
