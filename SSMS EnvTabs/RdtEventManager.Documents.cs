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
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;

namespace SSMS_EnvTabs
{
    internal sealed partial class RdtEventManager
    {
        private const string IndicatorTypeName = "Indicator";
        private const string IndicatorTypeFullName = "Microsoft.VisualStudio.Shell.Controls.Indicator";
        private const string AdornmentLayerTypeName = "AdornmentLayer";
        private const string AdornmentLayerTypeFullName = "Microsoft.VisualStudio.Text.Editor.Implementation.AdornmentLayer";
        private static readonly string[] PreferredTextViewMarginNames = new[]
        {
            "Glyph",
            "LineNumber",
            "LeftSelection",
            "Indicator",
            "SpacerMargin"
        };

        private static readonly string[] IndicatorPaletteHex = new[]
        {
            "#9083ef", "#d0b132", "#30b1cd", "#cf6468",
            "#6ba12a", "#bc8f6f", "#5bb2fa", "#d67441",
            "#bdbcbc", "#cbcc38", "#2aa0a4", "#d957a7",
            "#6bc6a5", "#946a5b", "#6a8ec6", "#e0a3a5"
        };

        private IComponentModel componentModel;
        private IVsEditorAdaptersFactoryService editorAdaptersFactoryService;

        private bool TryApplyLineIndicatorColor(uint docCookie, IVsWindowFrame frame, string moniker, IReadOnlyList<TabRuleMatcher.CompiledRule> rules, IReadOnlyList<TabRuleMatcher.CompiledManualRule> manualRules, TabGroupSettings settings)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (frame == null)
            {
                return true;
            }

            EnvTabsLog.Info($"LineIndicator: attempt cookie={docCookie} moniker='{moniker}' enabled={settings?.EnableLineIndicatorColor == true}");

            if (settings?.EnableLineIndicatorColor != true || settings.EnableAutoColor != true)
            {
                TryClearLineIndicatorColor(frame);
                return true;
            }

            if (!TryGetConnectionInfo(frame, out string server, out string database) || string.IsNullOrWhiteSpace(server))
            {
                EnvTabsLog.Info($"LineIndicator: skipped (no connection info) cookie={docCookie} moniker='{moniker}'");
                return true;
            }

            var manualMatch = TabRuleMatcher.MatchManual(manualRules, moniker);
            var matchedRule = TabRuleMatcher.MatchRule(rules, server, database);
            int? colorIndex = manualMatch?.ColorIndex ?? matchedRule?.ColorIndex;
            if (!colorIndex.HasValue || colorIndex.Value < 0 || colorIndex.Value >= IndicatorPaletteHex.Length)
            {
                EnvTabsLog.Info($"LineIndicator: skipped (no valid color index) cookie={docCookie} server='{server}' db='{database}'");
                TryClearLineIndicatorColor(frame);
                return true;
            }

            if (!TryFindLineIndicatorElements(frame, out List<DependencyObject> indicators) || indicators.Count == 0)
            {
                EnvTabsLog.Info($"LineIndicator: editor indicator not ready cookie={docCookie} server='{server}' db='{database}' colorIndex={colorIndex.Value}");
                return false;
            }

            var targetIndicators = SelectLikelyLineIndicators(indicators);
            if (targetIndicators.Count == 0)
            {
                EnvTabsLog.Info($"LineIndicator: no line-indicator candidates matched cookie={docCookie} server='{server}' db='{database}' totalIndicators={indicators.Count}");
                return false;
            }

            Brush brush = BuildIndicatorBrush(colorIndex.Value);
            int appliedCount = 0;
            foreach (DependencyObject indicator in targetIndicators)
            {
                if (TrySetIndicatorBackground(indicator, brush))
                {
                    appliedCount++;
                }
            }

            EnvTabsLog.Info($"LineIndicator: selected target type='{DescribeIndicator(targetIndicators[0])}'");
            EnvTabsLog.Info($"LineIndicator: {(appliedCount > 0 ? "applied" : "failed")} cookie={docCookie} server='{server}' db='{database}' colorIndex={colorIndex.Value} indicators={indicators.Count} selected={targetIndicators.Count} updated={appliedCount}");
            return appliedCount > 0;
        }

        private static Brush BuildIndicatorBrush(int colorIndex)
        {
            if (colorIndex < 0 || colorIndex >= IndicatorPaletteHex.Length)
            {
                return null;
            }

            var brush = (SolidColorBrush)new BrushConverter().ConvertFromString(IndicatorPaletteHex[colorIndex]);
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

        private void TryClearLineIndicatorColor(IVsWindowFrame frame)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!TryFindLineIndicatorElements(frame, out List<DependencyObject> indicators) || indicators.Count == 0)
            {
                return;
            }

            var targetIndicators = SelectLikelyLineIndicators(indicators);
            foreach (DependencyObject indicator in targetIndicators)
            {
                TryClearIndicatorBackground(indicator);
            }
        }

        private static void TryClearIndicatorBackground(DependencyObject indicator)
        {
            if (indicator is Control control)
            {
                control.ClearValue(Control.BackgroundProperty);
            }

            if (indicator is Border border)
            {
                border.ClearValue(Border.BackgroundProperty);
            }

            if (indicator is Panel panel)
            {
                panel.ClearValue(Panel.BackgroundProperty);
            }

            if (indicator is System.Windows.Shapes.Shape shape)
            {
                shape.ClearValue(System.Windows.Shapes.Shape.FillProperty);
                shape.ClearValue(System.Windows.Shapes.Shape.StrokeProperty);
            }

            TryClearBrushProperty(indicator, "Background");
            TryClearBrushProperty(indicator, "BorderBrush");
            TryClearBrushProperty(indicator, "Foreground");
            TryClearBrushProperty(indicator, "Fill");
            TryClearBrushProperty(indicator, "Stroke");
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

        private static List<DependencyObject> SelectLikelyLineIndicators(List<DependencyObject> indicators)
        {
            var selected = new List<DependencyObject>();
            if (indicators == null || indicators.Count == 0)
            {
                return selected;
            }

            var strictIndicators = indicators
                .Where(IsIndicatorControl)
                .Where(indicator => !IsRejectedLineIndicatorCandidate(indicator))
                .ToList();

            if (strictIndicators.Count == 0)
            {
                EnvTabsLog.Info("LineIndicator: strict indicator selection found 0 Indicator controls in editor scope.");
                return selected;
            }

            var scored = new List<(DependencyObject Indicator, int Score, string Reason)>();
            foreach (DependencyObject indicator in strictIndicators)
            {
                int score = ScoreLineIndicatorCandidate(indicator, out string reason);
                scored.Add((indicator, score, reason));
            }

            if (scored.Count == 0)
            {
                return selected;
            }

            var ordered = scored
                .OrderByDescending(c => c.Score)
                .ThenBy(c => DescribeIndicator(c.Indicator), StringComparer.Ordinal)
                .ToList();

            EnvTabsLog.Info("LineIndicator: candidate scores -> " + string.Join(" | ", ordered.Select(c => $"{c.Score}:{DescribeIndicator(c.Indicator)}:{c.Reason}")));

            var best = ordered.FirstOrDefault();
            if (best.Indicator == null)
            {
                return selected;
            }

            if (best.Score < 2)
            {
                return selected;
            }

            if (!selected.Contains(best.Indicator))
            {
                selected.Add(best.Indicator);
            }

            return selected;
        }

        private static int ScoreLineIndicatorCandidate(DependencyObject indicator, out string reason)
        {
            reason = string.Empty;
            if (!(indicator is FrameworkElement element))
            {
                reason = "not-framework-element";
                return -100;
            }

            if (element.ActualWidth <= 0 || element.ActualHeight <= 0)
            {
                reason = "no-size";
                return -100;
            }

            int score = 0;
            var reasons = new List<string>();
            string name = (element.Name ?? string.Empty).ToLowerInvariant();

            if (name.Contains("accentrect"))
            {
                score -= 20;
                reasons.Add("accentrect");
            }

            if (name.Contains("autohidetabselectedindicator"))
            {
                score -= 40;
                reasons.Add("autohide-indicator");
            }

            if (name.Contains("indicator"))
            {
                score += 1;
                reasons.Add("name-indicator");
            }

            if (element.ActualWidth <= 6)
            {
                score += 5;
                reasons.Add("thin");
            }
            else if (element.ActualWidth <= 12)
            {
                score += 1;
                reasons.Add("narrow");
            }
            else
            {
                score -= 5;
                reasons.Add("wide");
            }

            if (element.ActualHeight >= 12 && element.ActualHeight <= 120)
            {
                score += 1;
                reasons.Add("height-range");
            }

            if (element.ActualHeight >= 20)
            {
                score += 1;
                reasons.Add("tall-ish");
            }

            if (element.ActualHeight >= 40 && element.ActualHeight <= 800)
            {
                score += 4;
                reasons.Add("tall-marker-range");
            }

            if (TryGetElementLocation(element, out Point location))
            {
                if (location.Y >= 120)
                {
                    score += 8;
                    reasons.Add("below-tab-strip");
                }
                else
                {
                    score -= 10;
                    reasons.Add("near-top");
                }

                if (location.X <= 80)
                {
                    score += 3;
                    reasons.Add("left-edge");
                }
                else if (location.X >= 260)
                {
                    score -= 2;
                    reasons.Add("far-right");
                }

                if (TryIsNearBottomOfWindow(element, location, out bool nearBottom) && nearBottom)
                {
                    score -= 20;
                    reasons.Add("near-bottom");
                }
            }

            bool hasEditorAncestor = HasAncestorToken(element, "editor", "textview", "adornment", "margin", "scroll", "code");
            bool hasTabAncestor = HasAncestorToken(element, "tab", "documentwell", "title", "header", "well");
            bool hasToolWindowAncestor = HasAncestorToken(element, "autohide", "toolwindow", "objectexplorer");

            if (element is System.Windows.Shapes.Rectangle)
            {
                score -= 6;
                reasons.Add("rectangle-overlay");
            }

            if (hasEditorAncestor)
            {
                score += 4;
                reasons.Add("editor-ancestor");
            }

            if (hasTabAncestor)
            {
                score -= 5;
                reasons.Add("tab-ancestor");
            }

            if (hasToolWindowAncestor)
            {
                score -= 20;
                reasons.Add("toolwindow-ancestor");
            }

            reason = string.Join(",", reasons);
            return score;
        }

        private static bool IsRejectedLineIndicatorCandidate(DependencyObject indicator)
        {
            if (!(indicator is FrameworkElement element))
            {
                return true;
            }

            string name = (element.Name ?? string.Empty).ToLowerInvariant();
            string typeName = (element.GetType().FullName ?? element.GetType().Name ?? string.Empty).ToLowerInvariant();
            if (name.Contains("autohidetabselectedindicator") || name.Contains("accentrect"))
            {
                return true;
            }

            string probe = typeName + " " + name;
            if (probe.Contains("statusbar") || probe.Contains("taskstatus") || probe.Contains("compartment"))
            {
                return true;
            }

            if (typeName.Contains("vstoolbarthumb") || typeName.Contains("vsbutton") || typeName.Contains("separator"))
            {
                return true;
            }

            if (name.Contains("thumbborder"))
            {
                return true;
            }

            if (TryGetElementLocation(element, out Point location) && location.Y < 100)
            {
                return true;
            }

            if (element.ActualWidth > 12)
            {
                return true;
            }

            if (TryGetElementLocation(element, out Point location2)
                && TryIsNearBottomOfWindow(element, location2, out bool nearBottom)
                && nearBottom)
            {
                return true;
            }

            if (HasAncestorToken(element, "autohide", "toolwindow", "objectexplorer"))
            {
                return true;
            }

            return false;
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

        private static bool TryIsNearBottomOfWindow(FrameworkElement element, Point location, out bool isNearBottom)
        {
            isNearBottom = false;
            Window window = Window.GetWindow(element);
            if (window == null || window.ActualHeight <= 0)
            {
                return false;
            }

            isNearBottom = location.Y >= (window.ActualHeight - 140);
            return true;
        }

        private static bool HasAncestorToken(DependencyObject element, params string[] tokens)
        {
            DependencyObject current = element;
            for (int depth = 0; depth < 16 && current != null; depth++)
            {
                string typeName = current.GetType().Name ?? string.Empty;
                string name = (current as FrameworkElement)?.Name ?? string.Empty;
                string probe = (typeName + " " + name).ToLowerInvariant();

                for (int i = 0; i < (tokens?.Length ?? 0); i++)
                {
                    string token = tokens[i];
                    if (!string.IsNullOrWhiteSpace(token) && probe.Contains(token.ToLowerInvariant()))
                    {
                        return true;
                    }
                }

                try
                {
                    current = VisualTreeHelper.GetParent(current);
                }
                catch
                {
                    break;
                }
            }

            return false;
        }

        private static bool TryGetElementLocation(FrameworkElement element, out Point location)
        {
            location = new Point(double.NaN, double.NaN);
            if (element == null)
            {
                return false;
            }

            Window window = Window.GetWindow(element);
            if (window == null)
            {
                return false;
            }

            try
            {
                GeneralTransform transform = element.TransformToAncestor(window);
                location = transform.Transform(new Point(0, 0));
                return true;
            }
            catch
            {
                return false;
            }
        }

        private bool TryFindLineIndicatorElements(IVsWindowFrame frame, out List<DependencyObject> indicators)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            indicators = new List<DependencyObject>();
            if (frame == null)
            {
                return false;
            }

            try
            {
                if (!TryResolveEditorScopedSearchRoots(frame, out List<DependencyObject> roots, out string resolutionSource))
                {
                    return false;
                }

                EnvTabsLog.Info($"LineIndicator: editor scope resolved via {resolutionSource} roots={roots.Count}");

                var seen = new HashSet<DependencyObject>();
                foreach (DependencyObject root in roots)
                {
                    foreach (DependencyObject found in FindScopedLineIndicatorCandidates(root))
                    {
                        if (found != null && seen.Add(found))
                        {
                            indicators.Add(found);
                        }
                    }
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

        private static IEnumerable<DependencyObject> FindScopedLineIndicatorCandidates(DependencyObject root)
        {
            var results = new List<DependencyObject>();
            if (!(root is FrameworkElement rootElement))
            {
                return results;
            }

            foreach (DependencyObject candidate in EnumerateVisualDescendants(root, maxNodes: 4000))
            {
                if (!(candidate is FrameworkElement element))
                {
                    continue;
                }

                if (element.ActualWidth <= 0 || element.ActualHeight <= 0)
                {
                    continue;
                }

                if (element.ActualWidth > 20 || element.ActualHeight < 10 || element.ActualHeight > 600)
                {
                    continue;
                }

                string typeName = (element.GetType().FullName ?? element.GetType().Name ?? string.Empty).ToLowerInvariant();
                string name = (element.Name ?? string.Empty).ToLowerInvariant();
                string probe = typeName + " " + name;
                if (probe.Contains("button")
                    || probe.Contains("thumb")
                    || probe.Contains("separator")
                    || probe.Contains("image")
                    || probe.Contains("textblock")
                    || probe.Contains("scrollbar")
                    || probe.Contains("status")
                    || probe.Contains("compartment")
                    || probe.Contains("overflow"))
                {
                    continue;
                }

                if (!TryGetLocationRelativeToRoot(element, rootElement, out Point relativeLocation))
                {
                    continue;
                }

                double xThreshold = Math.Max(40, Math.Min(80, rootElement.ActualWidth * 0.35));
                if (relativeLocation.X < -2 || relativeLocation.X > xThreshold)
                {
                    continue;
                }

                if (relativeLocation.Y < -2 || relativeLocation.Y > rootElement.ActualHeight + 2)
                {
                    continue;
                }

                bool looksLikeMarker = element.ActualWidth <= 8
                    || probe.Contains("indicator")
                    || probe.Contains("glyph")
                    || probe.Contains("margin")
                    || probe.Contains("adornment")
                    || element is Border
                    || element is Panel
                    || element is System.Windows.Shapes.Shape;

                if (!looksLikeMarker)
                {
                    continue;
                }

                if (!results.Contains(candidate))
                {
                    results.Add(candidate);
                }
            }

            return results;
        }

        private static bool TryGetLocationRelativeToRoot(FrameworkElement element, FrameworkElement root, out Point location)
        {
            location = new Point(double.NaN, double.NaN);
            if (element == null || root == null)
            {
                return false;
            }

            try
            {
                GeneralTransform transform = element.TransformToAncestor(root);
                location = transform.Transform(new Point(0, 0));
                return true;
            }
            catch
            {
                return false;
            }
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

        private static List<DependencyObject> GetIndicatorSearchRoots(object docView)
        {
            var roots = new List<DependencyObject>();
            if (docView == null)
            {
                return roots;
            }

            TryAddRootCandidate(docView, roots);

            string[] knownMembers = new[]
            {
                "TextViewHost", "m_textViewHost", "_textViewHost",
                "WpfTextViewHost", "m_wpfTextViewHost",
                "ViewHost", "m_viewHost",
                "TextView", "m_textView"
            };

            foreach (string member in knownMembers)
            {
                TryAddMemberRoot(docView, member, roots);
            }

            TryAddLikelyViewMembers(docView, roots);
            TryAddEnumerableItems(docView, roots);
            TryAddObjectGraphRoots(docView, roots, maxDepth: 4, maxNodes: 350);

            if (roots.Count > 0)
            {
                EnvTabsLog.Info($"LineIndicator: discovered {roots.Count} visual root candidate(s) from DocView graph.");
            }

            return roots;
        }

        private static void TryAddLikelyViewMembers(object owner, List<DependencyObject> roots)
        {
            if (owner == null)
            {
                return;
            }

            const BindingFlags Flags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;

            try
            {
                foreach (FieldInfo field in owner.GetType().GetFields(Flags))
                {
                    if (!IsLikelyVisualMemberName(field?.Name))
                    {
                        continue;
                    }

                    TryAddRootCandidate(field.GetValue(owner), roots);
                }

                foreach (PropertyInfo prop in owner.GetType().GetProperties(Flags))
                {
                    if (prop == null || !prop.CanRead || prop.GetIndexParameters().Length != 0)
                    {
                        continue;
                    }

                    if (!IsLikelyVisualMemberName(prop.Name))
                    {
                        continue;
                    }

                    TryAddRootCandidate(prop.GetValue(owner, null), roots);
                }
            }
            catch
            {
                // Ignore reflection failures.
            }
        }

        private static void TryAddEnumerableItems(object owner, List<DependencyObject> roots)
        {
            if (owner is IEnumerable enumerable)
            {
                foreach (object item in enumerable)
                {
                    TryAddRootCandidate(item, roots);
                }
            }
        }

        private static void TryAddObjectGraphRoots(object rootObject, List<DependencyObject> roots, int maxDepth, int maxNodes)
        {
            if (rootObject == null || maxDepth < 0 || maxNodes <= 0)
            {
                return;
            }

            var comparer = ReferenceEqualityComparer.Instance;
            var visited = new HashSet<object>(comparer);
            var queue = new Queue<(object Obj, int Depth)>();
            queue.Enqueue((rootObject, 0));
            visited.Add(rootObject);

            int processed = 0;
            while (queue.Count > 0 && processed < maxNodes)
            {
                var current = queue.Dequeue();
                object obj = current.Obj;
                int depth = current.Depth;
                processed++;

                TryAddRootCandidate(obj, roots);

                if (depth >= maxDepth || obj == null)
                {
                    continue;
                }

                if (obj is string || obj.GetType().IsPrimitive || obj.GetType().IsEnum)
                {
                    continue;
                }

                foreach (object child in ReflectChildren(obj))
                {
                    if (child == null)
                    {
                        continue;
                    }

                    if (!visited.Add(child))
                    {
                        continue;
                    }

                    queue.Enqueue((child, depth + 1));
                    if (queue.Count + processed >= maxNodes)
                    {
                        break;
                    }
                }
            }
        }

        private static IEnumerable<object> ReflectChildren(object obj)
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
                    // Ignore inaccessible fields.
                }

                if (value == null)
                {
                    continue;
                }

                if (value is IEnumerable enumerable && !(value is string))
                {
                    foreach (object item in enumerable)
                    {
                        if (item != null)
                        {
                            yield return item;
                        }
                    }
                }

                yield return value;
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
                    // Ignore getter exceptions.
                }

                if (value == null)
                {
                    continue;
                }

                if (value is IEnumerable enumerable && !(value is string))
                {
                    foreach (object item in enumerable)
                    {
                        if (item != null)
                        {
                            yield return item;
                        }
                    }
                }

                yield return value;
            }
        }

        private sealed class ReferenceEqualityComparer : IEqualityComparer<object>
        {
            internal static readonly ReferenceEqualityComparer Instance = new ReferenceEqualityComparer();

            public new bool Equals(object x, object y)
            {
                return ReferenceEquals(x, y);
            }

            public int GetHashCode(object obj)
            {
                return RuntimeHelpers.GetHashCode(obj);
            }
        }

        private static bool IsLikelyVisualMemberName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
            {
                return false;
            }

            string n = name.ToLowerInvariant();
            return n.Contains("view") || n.Contains("host") || n.Contains("editor") || n.Contains("text");
        }

        private static void TryAddMemberRoot(object owner, string memberName, List<DependencyObject> roots)
        {
            if (owner == null || string.IsNullOrWhiteSpace(memberName))
            {
                return;
            }

            const BindingFlags Flags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;

            try
            {
                PropertyInfo prop = owner.GetType().GetProperty(memberName, Flags);
                if (prop != null)
                {
                    TryAddRootCandidate(prop.GetValue(owner, null), roots);
                    return;
                }

                FieldInfo field = owner.GetType().GetField(memberName, Flags);
                if (field != null)
                {
                    TryAddRootCandidate(field.GetValue(owner), roots);
                }
            }
            catch
            {
                // Ignore reflection failures.
            }
        }

        private static void TryAddRootCandidate(object candidate, List<DependencyObject> roots)
        {
            if (candidate == null)
            {
                return;
            }

            if (candidate is DependencyObject direct)
            {
                if (!roots.Contains(direct))
                {
                    roots.Add(direct);
                }
                return;
            }

            const BindingFlags Flags = BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic;
            try
            {
                PropertyInfo visualElement = candidate.GetType().GetProperty("VisualElement", Flags);
                if (visualElement?.GetValue(candidate, null) is DependencyObject visualRoot && !roots.Contains(visualRoot))
                {
                    roots.Add(visualRoot);
                }
            }
            catch
            {
                // Ignore reflection failures.
            }
        }

        private static List<DependencyObject> FindVisualByTypeNameOrName(DependencyObject root, string targetTypeName, string targetTypeFullName, string targetName)
        {
            var results = new List<DependencyObject>();
            if (root == null || (string.IsNullOrWhiteSpace(targetTypeName) && string.IsNullOrWhiteSpace(targetTypeFullName) && string.IsNullOrWhiteSpace(targetName)))
            {
                return results;
            }

            string typeName = root.GetType().Name;
            string fullTypeName = root.GetType().FullName;

            bool typeMatch = (!string.IsNullOrWhiteSpace(targetTypeName) && string.Equals(typeName, targetTypeName, StringComparison.OrdinalIgnoreCase))
                || (!string.IsNullOrWhiteSpace(targetTypeFullName) && string.Equals(fullTypeName, targetTypeFullName, StringComparison.OrdinalIgnoreCase));

            if (typeMatch)
            {
                results.Add(root);
            }

            if (!string.IsNullOrWhiteSpace(targetName)
                && root is FrameworkElement fe
                && string.Equals(fe.Name, targetName, StringComparison.OrdinalIgnoreCase))
            {
                if (!results.Contains(root))
                {
                    results.Add(root);
                }
            }

            int childCount;
            try
            {
                childCount = VisualTreeHelper.GetChildrenCount(root);
            }
            catch
            {
                return results;
            }

            for (int i = 0; i < childCount; i++)
            {
                DependencyObject child = VisualTreeHelper.GetChild(root, i);
                List<DependencyObject> found = FindVisualByTypeNameOrName(child, targetTypeName, targetTypeFullName, targetName);
                if (found.Count > 0)
                {
                    foreach (DependencyObject item in found)
                    {
                        if (!results.Contains(item))
                        {
                            results.Add(item);
                        }
                    }
                }
            }

            return results;
        }

        private static List<DependencyObject> FindIndicatorsUnderAdornmentLayers(DependencyObject root)
        {
            var indicators = new List<DependencyObject>();
            if (root == null)
            {
                return indicators;
            }

            List<DependencyObject> layers = FindVisualByTypeNameOrName(root, AdornmentLayerTypeName, AdornmentLayerTypeFullName, null);
            foreach (DependencyObject layer in layers)
            {
                List<DependencyObject> foundIndicators = FindVisualByTypeNameOrName(layer, IndicatorTypeName, IndicatorTypeFullName, IndicatorTypeName);
                foreach (DependencyObject indicator in foundIndicators)
                {
                    if (!indicators.Contains(indicator))
                    {
                        indicators.Add(indicator);
                    }
                }
            }

            return indicators;
        }

        private static List<DependencyObject> FindIndicatorsFromApplicationWindows()
        {
            var indicators = new List<DependencyObject>();

            Application app = Application.Current;
            if (app == null)
            {
                return indicators;
            }

            foreach (Window window in app.Windows)
            {
                if (window == null || !window.IsVisible)
                {
                    continue;
                }

                foreach (DependencyObject found in FindVisualByTypeNameOrName(window, IndicatorTypeName, IndicatorTypeFullName, IndicatorTypeName))
                {
                    if (found != null && !indicators.Contains(found))
                    {
                        indicators.Add(found);
                    }
                }
            }

            return indicators;
        }

        private static List<DependencyObject> FindLikelyEditorMarkerCandidatesFromApplicationWindows()
        {
            var candidates = new List<DependencyObject>();

            Application app = Application.Current;
            if (app == null)
            {
                return candidates;
            }

            foreach (Window window in app.Windows)
            {
                if (window == null || !window.IsVisible)
                {
                    continue;
                }

                foreach (DependencyObject node in EnumerateVisualDescendants(window, maxNodes: 12000))
                {
                    if (!(node is FrameworkElement fe))
                    {
                        continue;
                    }

                    if (fe.ActualWidth <= 0 || fe.ActualHeight <= 0)
                    {
                        continue;
                    }

                    if (fe.ActualWidth > 30 || fe.ActualHeight < 8 || fe.ActualHeight > 2200)
                    {
                        continue;
                    }

                    string feName = (fe.Name ?? string.Empty).ToLowerInvariant();
                    if (feName.Contains("accentrect") || feName.Contains("autohidetabselectedindicator"))
                    {
                        continue;
                    }

                    if (HasAncestorToken(fe, "autohide", "toolwindow", "objectexplorer"))
                    {
                        continue;
                    }

                    if (!TryGetElementLocation(fe, out Point location))
                    {
                        continue;
                    }

                    if (location.X > 500 || location.Y < 20)
                    {
                        continue;
                    }

                    if (!candidates.Contains(fe))
                    {
                        candidates.Add(fe);
                    }
                }
            }

            return candidates;
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
                    else
                    {
                        // ignore
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
