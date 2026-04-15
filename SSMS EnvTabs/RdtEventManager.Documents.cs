using Microsoft.VisualStudio;
using Microsoft.VisualStudio.ComponentModelHost;
using Microsoft.VisualStudio.Editor;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using Microsoft.VisualStudio.Text.Editor;
using Microsoft.VisualStudio.TextManager.Interop;
using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Automation;
using System.Windows.Controls;
using System.Windows.Interop;
using System.Windows.Media;

namespace SSMS_EnvTabs
{
    internal sealed partial class RdtEventManager
    {
        private const string IndicatorTypeName = "Indicator";
        private const string IndicatorTypeFullName = "Microsoft.VisualStudio.Shell.Controls.Indicator";
        private const string AdornmentLayerTypeName = "AdornmentLayer";
        private const string AdornmentLayerTypeFullName = "Microsoft.VisualStudio.Text.Editor.Implementation.AdornmentLayer";
        private const uint CommonControlSetBkColorMessage = 0x2001;
        private static readonly string[] PreferredTextViewMarginNames = new[]
        {
            "Glyph",
            "LineNumber",
            "LeftSelection",
            "Indicator",
            "SpacerMargin"
        };
        private const string StatusBarLeftContainerName = "PART_StatusBarLeftFrameControlContainer";

        private static readonly object statusBarProbeLock = new object();
        private static readonly HashSet<string> loggedStatusBarProbeKeys = new HashSet<string>(StringComparer.Ordinal);
        private static readonly HashSet<string> loggedStatusStripInventoryKeys = new HashSet<string>(StringComparer.Ordinal);
        private static readonly string[] IndicatorBrushPropertyNames = new[] { "Background", "BorderBrush", "Foreground", "Fill", "Stroke" };
        private static readonly ConditionalWeakTable<System.Windows.Forms.StatusStrip, StatusStripSnapshot> originalStatusStripState = new ConditionalWeakTable<System.Windows.Forms.StatusStrip, StatusStripSnapshot>();
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
                // Restore default status bar appearance if we previously colored it.
                TryResetDocViewStatusBarColor(frame);
                return;
            }

            int? colorIndex = manualMatch?.ColorIndex ?? matchedRule?.ColorIndex;
            if (!colorIndex.HasValue || colorIndex.Value < 0 || colorIndex.Value >= ColorPalette.Hex.Length)
            {
                return;
            }

            if (TryApplyDocViewStatusBarColor(frame, colorIndex.Value, out string docViewTarget))
            {
                EnvTabsLog.Info($"StatusBarColor: applied cookie={docCookie} server='{server}' db='{database}' colorIndex={colorIndex.Value} mode=docview target='{docViewTarget}'");
                return;
            }

            EnvTabsLog.Info($"StatusBarColor: target not found cookie={docCookie} server='{server}' db='{database}'");
        }

        private bool TryApplyDocViewStatusBarColor(IVsWindowFrame frame, int colorIndex, out string targetDescription)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            targetDescription = null;
            if (frame == null || colorIndex < 0 || colorIndex >= ColorPalette.Hex.Length)
            {
                return false;
            }

            if (frame.GetProperty((int)__VSFPROPID.VSFPROPID_DocView, out object docView) != VSConstants.S_OK || docView == null)
            {
                return false;
            }

            if ((TryGetReflectionMemberValue(docView, "statusBar", out object statusBarObject)
                    || TryGetReflectionMemberValue(docView, "StatusBar", out statusBarObject))
                && statusBarObject is System.Windows.Forms.StatusStrip statusStrip
                && TryApplyWinFormsStatusStripColor(statusStrip, colorIndex, out string textColorName))
            {
                targetDescription = "DocView.statusBar; control=" + statusStrip.GetType().FullName + "; text=" + textColorName;
                return true;
            }

            return false;
        }

        private void TryResetDocViewStatusBarColor(IVsWindowFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (frame == null)
            {
                return;
            }

            if (frame.GetProperty((int)__VSFPROPID.VSFPROPID_DocView, out object docView) != VSConstants.S_OK || docView == null)
            {
                return;
            }

            if ((TryGetReflectionMemberValue(docView, "statusBar", out object statusBarObject)
                    || TryGetReflectionMemberValue(docView, "StatusBar", out statusBarObject))
                && statusBarObject is System.Windows.Forms.StatusStrip statusStrip)
            {
                TryResetWinFormsStatusStrip(statusStrip);
            }
        }

        private static void TryResetWinFormsStatusStrip(System.Windows.Forms.StatusStrip statusStrip)
        {
            if (statusStrip == null || statusStrip.IsDisposed)
            {
                return;
            }

            if (!originalStatusStripState.TryGetValue(statusStrip, out StatusStripSnapshot snapshot))
            {
                // No snapshot means we never painted this strip — nothing to restore.
                return;
            }

            statusStrip.SuspendLayout();
            try
            {
                RestoreStatusStripRenderer(statusStrip, snapshot);
                statusStrip.BackColor = snapshot.BackColor;
                statusStrip.ForeColor = snapshot.ForeColor;

                foreach (StatusStripItemSnapshot itemSnapshot in snapshot.ItemSnapshots)
                {
                    if (itemSnapshot.Item == null || itemSnapshot.Item.IsDisposed)
                    {
                        continue;
                    }

                    itemSnapshot.Item.BackColor = itemSnapshot.BackColor;
                    itemSnapshot.Item.ForeColor = itemSnapshot.ForeColor;
                }
            }
            finally
            {
                statusStrip.ResumeLayout(true);
            }

            // Remove snapshot so a future paint will re-capture the (now restored) state.
            originalStatusStripState.Remove(statusStrip);

            statusStrip.Invalidate(true);
            statusStrip.Refresh();
        }

        private static void RestoreStatusStripRenderer(System.Windows.Forms.StatusStrip statusStrip, StatusStripSnapshot snapshot)
        {
            if (snapshot.RenderMode == System.Windows.Forms.ToolStripRenderMode.Custom)
            {
                if (snapshot.Renderer != null)
                {
                    statusStrip.Renderer = snapshot.Renderer;
                }

                return;
            }

            statusStrip.Renderer = null;
            statusStrip.RenderMode = snapshot.RenderMode;
        }

        private static bool TryApplyWinFormsStatusStripColor(System.Windows.Forms.StatusStrip statusStrip, int colorIndex, out string textColorName)
        {
            textColorName = null;
            if (statusStrip == null || statusStrip.IsDisposed || colorIndex < 0 || colorIndex >= ColorPalette.Hex.Length)
            {
                return false;
            }

            var backColor = System.Drawing.ColorTranslator.FromHtml(ColorPalette.Hex[colorIndex]);
            var foreColor = GetReadableStatusTextColor(backColor);
            var renderer = new SolidStatusStripRenderer(backColor, foreColor);
            textColorName = IsNearBlack(foreColor) ? "black" : "white";

            LogStatusStripItemsOnce(statusStrip);

            // Snapshot original state before first paint so we can restore later.
            if (!originalStatusStripState.TryGetValue(statusStrip, out _))
            {
                originalStatusStripState.GetOrCreateValue(statusStrip).CaptureFrom(statusStrip);
            }

            statusStrip.SuspendLayout();
            try
            {
                statusStrip.RenderMode = System.Windows.Forms.ToolStripRenderMode.Professional;
                statusStrip.Renderer = renderer;
                statusStrip.BackColor = backColor;
                statusStrip.ForeColor = foreColor;

                ApplyStatusStripItemColors(statusStrip.Items, backColor, foreColor);
            }
            finally
            {
                statusStrip.ResumeLayout(true);
            }

            statusStrip.Invalidate(true);
            statusStrip.Refresh();
            return true;
        }

        private static void ApplyStatusStripItemColors(System.Windows.Forms.ToolStripItemCollection items, System.Drawing.Color backColor, System.Drawing.Color foreColor)
        {
            if (items == null)
            {
                return;
            }

            foreach (System.Windows.Forms.ToolStripItem item in items)
            {
                item.BackColor = backColor;

                if (ShouldOverrideStatusStripItemTextColor(item))
                {
                    item.ForeColor = foreColor;
                }

                if (item is System.Windows.Forms.ToolStripDropDownItem dropDownItem)
                {
                    ApplyStatusStripItemColors(dropDownItem.DropDownItems, backColor, foreColor);
                }
            }
        }

        private static bool ShouldOverrideStatusStripItemTextColor(System.Windows.Forms.ToolStripItem item)
        {
            return !IsSubtleStatusStripSeparator(item);
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

            string probe = string.Join(" ", new[]
            {
                item.Name,
                item.AccessibleName,
                item.AccessibleDescription,
                item.GetType().FullName
            }.Where(value => !string.IsNullOrWhiteSpace(value))).ToLowerInvariant();

            if (probe.Contains("separator"))
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

        private static void LogStatusStripItemsOnce(System.Windows.Forms.StatusStrip statusStrip)
        {
            if (statusStrip == null)
            {
                return;
            }

            string key = (statusStrip.GetType().FullName ?? statusStrip.GetType().Name)
                + "#"
                + RuntimeHelpers.GetHashCode(statusStrip);

            lock (statusBarProbeLock)
            {
                if (!loggedStatusStripInventoryKeys.Add(key))
                {
                    return;
                }
            }

            var items = (statusStrip.Items ?? new System.Windows.Forms.ToolStripItemCollection(statusStrip, null))
                .Cast<System.Windows.Forms.ToolStripItem>()
                .Select(DescribeStatusStripItemForLog)
                .ToList();

            EnvTabsLog.Info("StatusBarColor: item inventory -> " + (items.Count == 0 ? "(none)" : string.Join(" | ", items)));
        }

        private static string DescribeStatusStripItemForLog(System.Windows.Forms.ToolStripItem item)
        {
            if (item == null)
            {
                return "<null>";
            }

            string text = item.Text ?? string.Empty;
            string codes = text.Length == 0
                ? "none"
                : string.Join(",", text.Select(ch => ((int)ch).ToString("X4")));

            return "type='" + (item.GetType().FullName ?? item.GetType().Name)
                + "' name='" + (item.Name ?? string.Empty)
                + "' text='" + text.Replace("'", "''")
                + "' codes=[" + codes + "]"
                + " preserveTextColor=" + IsSubtleStatusStripSeparator(item);
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

        private static bool IsNearBlack(System.Drawing.Color color)
        {
            return color.R < 32 && color.G < 32 && color.B < 32;
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

        private static IEnumerable<object> EnumerateProbeObjectGraph(object root, int maxDepth, int maxNodes)
        {
            if (root == null || maxDepth < 0 || maxNodes <= 0)
            {
                yield break;
            }

            var visited = new HashSet<object>(ReferenceEqualityComparer.Instance);
            var queue = new Queue<(object Node, int Depth)>();
            visited.Add(root);
            queue.Enqueue((root, 0));

            int visitedCount = 0;
            while (queue.Count > 0 && visitedCount < maxNodes)
            {
                var current = queue.Dequeue();
                visitedCount++;
                yield return current.Node;

                if (current.Depth >= maxDepth)
                {
                    continue;
                }

                foreach ((string _, object child) in EnumerateProbeMembers(current.Node))
                {
                    if (child == null || IsSimpleProbeType(child.GetType()) || !visited.Add(child))
                    {
                        continue;
                    }

                    queue.Enqueue((child, current.Depth + 1));
                }
            }
        }

        private sealed class StatusStripSnapshot
        {
            public System.Windows.Forms.ToolStripRenderMode RenderMode { get; set; }
            public System.Windows.Forms.ToolStripRenderer Renderer { get; set; }
            public System.Drawing.Color BackColor { get; set; }
            public System.Drawing.Color ForeColor { get; set; }
            public List<StatusStripItemSnapshot> ItemSnapshots { get; } = new List<StatusStripItemSnapshot>();

            public void CaptureFrom(System.Windows.Forms.StatusStrip strip)
            {
                RenderMode = strip.RenderMode;
                Renderer = strip.Renderer;
                BackColor = strip.BackColor;
                ForeColor = strip.ForeColor;
                ItemSnapshots.Clear();

                foreach (System.Windows.Forms.ToolStripItem item in strip.Items)
                {
                    ItemSnapshots.Add(StatusStripItemSnapshot.FromItem(item));

                    if (item is System.Windows.Forms.ToolStripDropDownItem dropDownItem)
                    {
                        foreach (System.Windows.Forms.ToolStripItem subItem in dropDownItem.DropDownItems)
                        {
                            ItemSnapshots.Add(StatusStripItemSnapshot.FromItem(subItem));
                        }
                    }
                }
            }
        }

        private sealed class IndicatorBrushSnapshot
        {
            private readonly List<IndicatorBrushPropertySnapshot> properties = new List<IndicatorBrushPropertySnapshot>();

            public void CaptureFrom(DependencyObject target)
            {
                properties.Clear();

                foreach (string propertyName in IndicatorBrushPropertyNames)
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

        private sealed class StatusStripItemSnapshot
        {
            public System.Windows.Forms.ToolStripItem Item { get; private set; }
            public System.Drawing.Color BackColor { get; private set; }
            public System.Drawing.Color ForeColor { get; private set; }

            public static StatusStripItemSnapshot FromItem(System.Windows.Forms.ToolStripItem item)
            {
                return new StatusStripItemSnapshot
                {
                    Item = item,
                    BackColor = item?.BackColor ?? System.Drawing.SystemColors.Control,
                    ForeColor = item?.ForeColor ?? System.Drawing.SystemColors.ControlText
                };
            }
        }

        private sealed class SolidStatusStripRenderer : System.Windows.Forms.ToolStripProfessionalRenderer
        {
            public SolidStatusStripRenderer(System.Drawing.Color backColor, System.Drawing.Color foreColor)
                : base(new SolidStatusStripColorTable(backColor, foreColor))
            {
                RoundedEdges = false;
            }
        }

        private sealed class ReferenceEqualityComparer : IEqualityComparer<object>
        {
            public static readonly ReferenceEqualityComparer Instance = new ReferenceEqualityComparer();

            private ReferenceEqualityComparer()
            {
            }

            public new bool Equals(object x, object y)
            {
                return ReferenceEquals(x, y);
            }

            public int GetHashCode(object obj)
            {
                return obj == null ? 0 : RuntimeHelpers.GetHashCode(obj);
            }
        }

        private sealed class SolidStatusStripColorTable : System.Windows.Forms.ProfessionalColorTable
        {
            private readonly System.Drawing.Color backColor;
            private readonly System.Drawing.Color foreColor;
            private readonly System.Drawing.Color separatorDarkColor;
            private readonly System.Drawing.Color separatorLightColor;

            public SolidStatusStripColorTable(System.Drawing.Color backColor, System.Drawing.Color foreColor)
            {
                this.backColor = backColor;
                this.foreColor = foreColor;
                separatorDarkColor = BlendStatusStripColor(backColor, foreColor, 0.24);
                separatorLightColor = BlendStatusStripColor(backColor, foreColor, 0.12);
            }

            public override System.Drawing.Color StatusStripGradientBegin => backColor;

            public override System.Drawing.Color StatusStripGradientEnd => backColor;

            public override System.Drawing.Color ToolStripGradientBegin => backColor;

            public override System.Drawing.Color ToolStripGradientMiddle => backColor;

            public override System.Drawing.Color ToolStripGradientEnd => backColor;

            public override System.Drawing.Color ToolStripBorder => backColor;

            public override System.Drawing.Color ImageMarginGradientBegin => backColor;

            public override System.Drawing.Color ImageMarginGradientMiddle => backColor;

            public override System.Drawing.Color ImageMarginGradientEnd => backColor;

            public override System.Drawing.Color MenuItemBorder => backColor;

            public override System.Drawing.Color MenuItemSelected => backColor;

            public override System.Drawing.Color MenuItemSelectedGradientBegin => backColor;

            public override System.Drawing.Color MenuItemSelectedGradientEnd => backColor;

            public override System.Drawing.Color ButtonSelectedBorder => backColor;

            public override System.Drawing.Color ButtonSelectedGradientBegin => backColor;

            public override System.Drawing.Color ButtonSelectedGradientMiddle => backColor;

            public override System.Drawing.Color ButtonSelectedGradientEnd => backColor;

            public override System.Drawing.Color ButtonPressedBorder => backColor;

            public override System.Drawing.Color ButtonPressedGradientBegin => backColor;

            public override System.Drawing.Color ButtonPressedGradientMiddle => backColor;

            public override System.Drawing.Color ButtonPressedGradientEnd => backColor;

            public override System.Drawing.Color ButtonCheckedGradientBegin => backColor;

            public override System.Drawing.Color ButtonCheckedGradientMiddle => backColor;

            public override System.Drawing.Color ButtonCheckedGradientEnd => backColor;

            public override System.Drawing.Color ButtonCheckedHighlight => backColor;

            public override System.Drawing.Color ButtonCheckedHighlightBorder => backColor;

            public override System.Drawing.Color CheckBackground => backColor;

            public override System.Drawing.Color CheckPressedBackground => backColor;

            public override System.Drawing.Color CheckSelectedBackground => backColor;

            public override System.Drawing.Color GripDark => separatorDarkColor;

            public override System.Drawing.Color GripLight => separatorLightColor;

            public override System.Drawing.Color SeparatorDark => separatorDarkColor;

            public override System.Drawing.Color SeparatorLight => separatorLightColor;
        }

        private static System.Drawing.Color BlendStatusStripColor(System.Drawing.Color background, System.Drawing.Color foreground, double foregroundWeight)
        {
            foregroundWeight = Math.Max(0.0, Math.Min(1.0, foregroundWeight));
            double backgroundWeight = 1.0 - foregroundWeight;

            int r = (int)Math.Round((background.R * backgroundWeight) + (foreground.R * foregroundWeight));
            int g = (int)Math.Round((background.G * backgroundWeight) + (foreground.G * foregroundWeight));
            int b = (int)Math.Round((background.B * backgroundWeight) + (foreground.B * foregroundWeight));

            return System.Drawing.Color.FromArgb(255, r, g, b);
        }

        private static bool TrySetStatusBarBackground(DependencyObject target, Brush brush)
        {
            if (target == null || brush == null)
            {
                return false;
            }

            bool changed = false;

            if (target is Control control)
            {
                control.Background = brush;
                changed = true;
            }

            if (target is Border border)
            {
                border.Background = brush;
                changed = true;
            }

            if (target is Panel panel)
            {
                panel.Background = brush;
                changed = true;
            }

            if (target is System.Windows.Shapes.Shape shape)
            {
                shape.Fill = brush;
                shape.Stroke = brush;
                changed = true;
            }

            if (TrySetBrushProperty(target, "Background", brush))
            {
                changed = true;
            }

            if (TrySetBrushProperty(target, "BorderBrush", brush))
            {
                changed = true;
            }

            if (TrySetBrushProperty(target, "Fill", brush))
            {
                changed = true;
            }

            if (TrySetBrushProperty(target, "Stroke", brush))
            {
                changed = true;
            }

            return changed;
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

        private static void TryClearBrushProperty(object target, string propertyName)
        {
            if (target == null || string.IsNullOrWhiteSpace(propertyName))
            {
                return;
            }

            try
            {
                var prop = target.GetType().GetProperty(propertyName);
                if (prop != null && prop.CanWrite)
                {
                    prop.SetValue(target, null, null);
                }
            }
            catch
            {
                // Ignore reflection clear failures.
            }
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

                foreach (string marginName in PreferredTextViewMarginNames)
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

        private static void TryLogStatusBarOwnerCandidates(object docView, object connection)
        {
            if (docView == null)
            {
                return;
            }

            string key = (docView.GetType().FullName ?? docView.GetType().Name)
                + "#"
                + RuntimeHelpers.GetHashCode(docView)
                + "|"
                + (connection?.GetType().FullName ?? "<null>")
                + "#"
                + (connection != null ? RuntimeHelpers.GetHashCode(connection).ToString() : "null");

            lock (statusBarProbeLock)
            {
                if (!loggedStatusBarProbeKeys.Add(key))
                {
                    return;
                }
            }

            EnvTabsLog.Info($"StatusBarProbe: begin docView='{docView.GetType().FullName}' connection='{connection?.GetType().FullName ?? "<null>"}'");

            int emitted = 0;
            foreach (string candidate in EnumerateStatusBarOwnerCandidates(docView, connection))
            {
                EnvTabsLog.Info("StatusBarProbe: " + candidate);
                emitted++;
                if (emitted >= 40)
                {
                    break;
                }
            }

            if (emitted == 0)
            {
                EnvTabsLog.Info("StatusBarProbe: no candidate members found near DocView or connection object.");
            }
        }

        private static IEnumerable<string> EnumerateStatusBarOwnerCandidates(object docView, object connection)
        {
            var visited = new HashSet<object>(ReferenceEqualityComparer.Instance);
            var queue = new Queue<(object Node, string Path, int Depth)>();

            EnqueueProbeNode(docView, "DocView", 0, queue, visited);
            EnqueueProbeNode(connection, "Connection", 0, queue, visited);

            int visitedCount = 0;
            while (queue.Count > 0 && visitedCount < 160)
            {
                var current = queue.Dequeue();
                visitedCount++;

                foreach (var member in EnumerateProbeMembers(current.Node))
                {
                    if (member.Value == null)
                    {
                        continue;
                    }

                    Type valueType = member.Value.GetType();
                    string path = current.Path + "." + member.Name;
                    string probe = ((member.Name ?? string.Empty) + " " + (valueType.FullName ?? valueType.Name ?? string.Empty)).ToLowerInvariant();

                    if (LooksLikeStatusBarOwnerProbe(probe))
                    {
                        yield return path + " => type='" + (valueType.FullName ?? valueType.Name) + "' value='" + FormatProbeValue(member.Value) + "'";
                    }

                    if (current.Depth >= 3 || IsSimpleProbeType(valueType))
                    {
                        continue;
                    }

                    if (visited.Add(member.Value))
                    {
                        queue.Enqueue((member.Value, path, current.Depth + 1));
                    }
                }
            }
        }

        private static void EnqueueProbeNode(object value, string path, int depth, Queue<(object Node, string Path, int Depth)> queue, HashSet<object> visited)
        {
            if (value == null || queue == null || visited == null)
            {
                return;
            }

            if (visited.Add(value))
            {
                queue.Enqueue((value, path, depth));
            }
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

        private static bool LooksLikeStatusBarOwnerProbe(string probe)
        {
            if (string.IsNullOrWhiteSpace(probe))
            {
                return false;
            }

            return probe.Contains("status")
                || probe.Contains("color")
                || probe.Contains("pane")
                || probe.Contains("host")
                || probe.Contains("hwn")
                || probe.Contains("viewpresenter")
                || probe.Contains("generic")
                || probe.Contains("client")
                || probe.Contains("strip")
                || probe.Contains("connection");
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

        private static string FormatProbeValue(object value)
        {
            if (value == null)
            {
                return "<null>";
            }

            if (value is string s)
            {
                return s;
            }

            if (value is System.Drawing.Color drawingColor)
            {
                return drawingColor.ToArgb().ToString("X8");
            }

            if (value is Brush brush)
            {
                return brush.ToString();
            }

            Type type = value.GetType();
            if (type.IsPrimitive || type.IsEnum || value is decimal || value is Guid || value is DateTime || value is TimeSpan || value is IntPtr || value is UIntPtr)
            {
                return Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture);
            }

            return type.FullName ?? type.Name;
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
