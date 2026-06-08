using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Documents;
using System.Windows.Media;
using Microsoft.VisualStudio.PlatformUI;
using WinFormsDialogResult = System.Windows.Forms.DialogResult;

namespace SSMS_EnvTabs
{
    internal sealed class UninstallPromptDialog : DialogWindow, IDisposable
    {
        private WinFormsDialogResult dialogResult = WinFormsDialogResult.Cancel;
        private Button uninstallButton;
        private Button cancelButton;

        public UninstallPromptDialog()
        {
            Title = "Uninstall EnvTabs";
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            ShowInTaskbar = false;
            SizeToContent = SizeToContent.Height;
            Width = 500;
            MinWidth = 500;

            BuildUi();
            ApplyThemeResources();
        }

        public new WinFormsDialogResult ShowDialog()
        {
            base.ShowDialog();
            return dialogResult;
        }

        public void Dispose()
        {
            if (IsVisible)
            {
                Close();
            }
        }

        private void BuildUi()
        {
            var root = new Grid();
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            root.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var contentPanel = new StackPanel
            {
                Margin = new Thickness(16, 14, 16, 10),
                Orientation = Orientation.Vertical
            };

            var header = new TextBlock
            {
                Text = "Uninstall EnvTabs?",
                FontWeight = FontWeights.Bold,
                FontSize = SystemFonts.MessageFontSize + 6,
                Margin = new Thickness(0, 0, 0, 10)
            };
            contentPanel.Children.Add(header);

            var warning = new TextBlock
            {
                Text = "Proceeding will close SSMS. SSMS may prompt you to save any unsaved queries before it exits.",
                TextWrapping = TextWrapping.Wrap,
                FontWeight = FontWeights.SemiBold,
                Margin = new Thickness(0, 2, 0, 8)
            };
            contentPanel.Children.Add(warning);

            var details = new TextBlock
            {
                Text = "After SSMS closes, the EnvTabs uninstaller will start automatically.",
                TextWrapping = TextWrapping.Wrap,
                Margin = new Thickness(0, 0, 0, 2)
            };
            contentPanel.Children.Add(details);

            var buttonPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Center,
                Margin = new Thickness(0, 6, 0, 8)
            };

            uninstallButton = new Button
            {
                Content = "Close + _Uninstall",
                MinWidth = 150,
                Height = 26,
                Margin = new Thickness(6, 10, 6, 10),
                IsDefault = true
            };
            ApplyControlBrushes(uninstallButton);
            uninstallButton.Click += (s, e) =>
            {
                dialogResult = WinFormsDialogResult.Yes;
                DialogResult = true;
            };

            cancelButton = new Button
            {
                Content = "_Cancel",
                MinWidth = 110,
                Height = 26,
                Margin = new Thickness(6, 10, 6, 10),
                IsCancel = true
            };
            ApplyControlBrushes(cancelButton);
            cancelButton.Click += (s, e) =>
            {
                dialogResult = WinFormsDialogResult.Cancel;
                DialogResult = false;
            };

            buttonPanel.Children.Add(uninstallButton);
            buttonPanel.Children.Add(cancelButton);

            Grid.SetRow(contentPanel, 0);
            Grid.SetRow(buttonPanel, 1);

            root.Children.Add(contentPanel);
            root.Children.Add(buttonPanel);

            Content = root;
        }

        private void ApplyThemeResources()
        {
            SetResourceReference(BackgroundProperty, EnvironmentColors.ToolWindowBackgroundBrushKey);
            SetResourceReference(ForegroundProperty, EnvironmentColors.ToolWindowTextBrushKey);

            if (Content is FrameworkElement root)
            {
                root.SetResourceReference(BackgroundProperty, EnvironmentColors.ToolWindowBackgroundBrushKey);
                root.SetResourceReference(TextElement.ForegroundProperty, EnvironmentColors.ToolWindowTextBrushKey);
            }

            ApplyDialogButtonStyles();
        }

        private void ApplyDialogButtonStyles()
        {
            var baseBackground = TryFindResource(EnvironmentColors.ToolWindowBackgroundBrushKey) as Brush
                ?? TryFindResource(SystemColors.ControlBrushKey) as Brush
                ?? Brushes.Transparent;

            var baseForeground = TryFindResource(EnvironmentColors.ToolWindowTextBrushKey) as Brush
                ?? TryFindResource(SystemColors.ControlTextBrushKey) as Brush
                ?? Brushes.Black;

            bool isLightTheme = GetRelativeLuminance(GetBrushColor(baseBackground, Colors.White)) > 0.6;

            var primaryBaseColor = (Color)ColorConverter.ConvertFromString(isLightTheme ? "#5649B0" : "#9184EE");
            var primaryHoverColor = (Color)ColorConverter.ConvertFromString(isLightTheme ? "#665bb7" : "#867bda");

            var accentBrush = new SolidColorBrush(primaryBaseColor);
            var primaryHover = new SolidColorBrush(primaryHoverColor);
            var primaryPressed = BlendBrush(accentBrush, Colors.Black, 0.12f);

            var accentForeground = isLightTheme ? Brushes.White : Brushes.Black;
            var secondaryForeground = isLightTheme ? Brushes.Black : Brushes.White;

            var baseBorder = WithOpacity(secondaryForeground, 0.28);
            var secondaryHover = isLightTheme
                ? BlendBrush(baseBackground, Colors.Black, 0.06f)
                : BlendBrush(baseBackground, Colors.White, 0.06f);
            var secondaryPressed = isLightTheme
                ? BlendBrush(baseBackground, Colors.Black, 0.12f)
                : BlendBrush(baseBackground, Colors.Black, 0.08f);

            var primaryStyle = CreateButtonStyle(accentBrush, baseBorder, accentForeground, primaryHover, primaryPressed);
            var secondaryStyle = CreateButtonStyle(baseBackground, baseBorder, secondaryForeground, secondaryHover, secondaryPressed);

            ApplyButtonStyle(uninstallButton, primaryStyle);
            ApplyButtonStyle(cancelButton, secondaryStyle);
        }

        private Style CreateButtonStyle(Brush background, Brush border, Brush foreground, Brush hoverBackground, Brush pressedBackground)
        {
            var style = new Style(typeof(Button));
            style.Setters.Add(new Setter(BackgroundProperty, background));
            style.Setters.Add(new Setter(BorderBrushProperty, border));
            style.Setters.Add(new Setter(ForegroundProperty, foreground));
            style.Setters.Add(new Setter(BorderThicknessProperty, new Thickness(1)));
            style.Setters.Add(new Setter(HorizontalContentAlignmentProperty, HorizontalAlignment.Center));
            style.Setters.Add(new Setter(VerticalContentAlignmentProperty, VerticalAlignment.Center));

            var template = new ControlTemplate(typeof(Button));
            var borderFactory = new FrameworkElementFactory(typeof(Border));
            borderFactory.Name = "Bd";
            borderFactory.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(BackgroundProperty));
            borderFactory.SetValue(Border.BorderBrushProperty, new TemplateBindingExtension(BorderBrushProperty));
            borderFactory.SetValue(Border.BorderThicknessProperty, new TemplateBindingExtension(BorderThicknessProperty));
            borderFactory.SetValue(Border.CornerRadiusProperty, new CornerRadius(3));

            var contentPresenter = new FrameworkElementFactory(typeof(ContentPresenter));
            contentPresenter.SetValue(ContentPresenter.ContentProperty, new TemplateBindingExtension(ContentControl.ContentProperty));
            contentPresenter.SetValue(ContentPresenter.ContentTemplateProperty, new TemplateBindingExtension(ContentControl.ContentTemplateProperty));
            contentPresenter.SetValue(ContentPresenter.ContentStringFormatProperty, new TemplateBindingExtension(ContentControl.ContentStringFormatProperty));
            contentPresenter.SetValue(ContentPresenter.MarginProperty, new TemplateBindingExtension(Control.PaddingProperty));
            contentPresenter.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Center);
            contentPresenter.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            contentPresenter.SetValue(ContentPresenter.RecognizesAccessKeyProperty, true);

            borderFactory.AppendChild(contentPresenter);
            template.VisualTree = borderFactory;

            var hoverTrigger = new Trigger { Property = IsMouseOverProperty, Value = true };
            hoverTrigger.Setters.Add(new Setter(Border.BackgroundProperty, hoverBackground, "Bd"));

            var pressedTrigger = new Trigger { Property = Button.IsPressedProperty, Value = true };
            pressedTrigger.Setters.Add(new Setter(Border.BackgroundProperty, pressedBackground, "Bd"));

            var disabledTrigger = new Trigger { Property = IsEnabledProperty, Value = false };
            disabledTrigger.Setters.Add(new Setter(OpacityProperty, 0.55));

            template.Triggers.Add(hoverTrigger);
            template.Triggers.Add(pressedTrigger);
            template.Triggers.Add(disabledTrigger);

            style.Setters.Add(new Setter(TemplateProperty, template));

            return style;
        }

        private static void ApplyButtonStyle(Button button, Style style)
        {
            if (button == null || style == null)
            {
                return;
            }

            button.ClearValue(BackgroundProperty);
            button.ClearValue(ForegroundProperty);
            button.ClearValue(BorderBrushProperty);
            button.Style = style;
            button.MinHeight = 24;
            button.Padding = new Thickness(12, 4, 12, 4);
        }

        private static void ApplyControlBrushes(Control control)
        {
            if (control == null)
            {
                return;
            }

            control.SetResourceReference(BackgroundProperty, EnvironmentColors.ToolWindowBackgroundBrushKey);
            control.SetResourceReference(ForegroundProperty, EnvironmentColors.ToolWindowTextBrushKey);
        }

        private static Color GetBrushColor(Brush brush, Color fallback)
        {
            if (brush is SolidColorBrush solidBrush)
            {
                return solidBrush.Color;
            }

            return fallback;
        }

        private static double GetRelativeLuminance(Color color)
        {
            return (0.2126 * color.R + 0.7152 * color.G + 0.0722 * color.B) / 255.0;
        }

        private static Brush WithOpacity(Brush brush, double opacity)
        {
            if (brush is SolidColorBrush solid)
            {
                var color = solid.Color;
                var alpha = (byte)Math.Max(0, Math.Min(255, opacity * 255));
                return new SolidColorBrush(Color.FromArgb(alpha, color.R, color.G, color.B));
            }

            return brush;
        }

        private static Brush BlendBrush(Brush brush, Color blendColor, float amount)
        {
            if (brush is SolidColorBrush solid)
            {
                var blended = BlendColor(solid.Color, blendColor, amount);
                return new SolidColorBrush(blended);
            }

            return brush;
        }

        private static Color BlendColor(Color baseColor, Color blendColor, float blendAmount)
        {
            blendAmount = Math.Max(0, Math.Min(1, blendAmount));
            byte r = (byte)(baseColor.R + (blendColor.R - baseColor.R) * blendAmount);
            byte g = (byte)(baseColor.G + (blendColor.G - baseColor.G) * blendAmount);
            byte b = (byte)(baseColor.B + (blendColor.B - baseColor.B) * blendAmount);
            return Color.FromArgb(baseColor.A, r, g, b);
        }
    }
}
