using System;
using System.Collections.Generic;
using System.Reflection;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Animation;
using System.Windows.Shapes;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.PlatformUI;
using IOPath = System.IO.Path;
using ShapePath = System.Windows.Shapes.Path;

namespace SSMS_EnvTabs
{
    public partial class SettingsToolWindowControl : UserControl
    {
        private bool suppressAutoSave;

        private Button openJsonButton;
        private Button resetDefaultsButton;
        private Button reloadButton;
        private Button updateEnvTabsButton;
        private Button copyLogPathButton;
        private TabControl settingsTabControl;
        private TabItem generalSettingsTab;
        private TabItem styleTemplatesTab;
        private TabItem connectionGroupsTab;
        private ComboBox autoConfigureCombo;
        private CheckBox loggingToggle;
        private CheckBox autoRenameToggle;
        private CheckBox autoColorToggle;
        private CheckBox configurePromptToggle;
        private CheckBox serverAliasPromptToggle;
        private CheckBox updateChecksToggle;
        private CheckBox removeDotSqlToggle;
        private CheckBox lineIndicatorColorToggle;
        private CheckBox statusBarColorToggle;
        private TextBox suggestedGroupNameStyleTextBox;
        private TextBox newQueryRenameStyleTextBox;
        private TextBox savedFileRenameStyleTextBox;
        private Button suggestedGroupNameEditButton;
        private Button suggestedGroupNameSaveButton;
        private Button suggestedGroupNameCancelButton;
        private Button newQueryRenameEditButton;
        private Button newQueryRenameSaveButton;
        private Button newQueryRenameCancelButton;
        private Button savedFileRenameEditButton;
        private Button savedFileRenameSaveButton;
        private Button savedFileRenameCancelButton;
        private StackPanel suggestedGroupNameTokenPanel;
        private StackPanel newQueryRenameTokenPanel;
        private StackPanel savedFileRenameTokenPanel;
        private TextBlock suggestedGroupNameTokenHelpText;
        private TextBlock newQueryRenameTokenHelpText;
        private TextBlock savedFileRenameTokenHelpText;
        private StackPanel serverAliasSectionPanel;
        private Grid serverAliasColumnHeaderGrid;
        private StackPanel serverAliasRowsPanel;
        private StackPanel connectionGroupCardsPanel;
        private TextBlock connectionGroupsDescriptionText;
        private Grid checkUpdatesRowGrid;
        private Button addServerAliasButton;
        private Button addConnectionGroupButton;
        private TextBlock statusText;
        private EditableStyleField activeStyleEditField = EditableStyleField.None;
        private string styleEditSnapshot;
        private List<ServerAliasRowState> serverAliasRows = new List<ServerAliasRowState>();
        private List<ConnectionGroupCardState> connectionGroupCards = new List<ConnectionGroupCardState>();
        private bool showServerAliasSection = true;
        private bool showGroupNameInConnectionCards = true;
        private string currentAutoConfigureMode = DefaultAutoConfigureValue;
        private int nextInlineRowId = 1;
        private bool isThemeEventSubscribed;
        private ServerAliasRowState activeAliasEditRow;
        private TextBox activeAliasServerTextBox;
        private TextBox activeAliasAliasTextBox;
        private ConnectionGroupCardState activeConnectionGroupEditCard;
        private TextBox activeConnectionGroupNameTextBox;
        private TextBox activeConnectionServerTextBox;
        private TextBox activeConnectionDatabaseTextBox;
        private TextBox activeConnectionPriorityTextBox;
        private ComboBox activeConnectionColorCombo;

        private enum EditableStyleField
        {
            None,
            SuggestedGroupName,
            NewQueryRename,
            SavedFileRename
        }

        private sealed class AutoConfigureItem
        {
            public string DisplayName { get; set; }
            public string Value { get; set; }
        }

        private sealed class ServerAliasRowState
        {
            public int Id { get; set; }
            public string Server { get; set; }
            public string Alias { get; set; }
            public bool IsEditing { get; set; }
            public bool IsNew { get; set; }
            public string SnapshotServer { get; set; }
            public string SnapshotAlias { get; set; }
        }

        private sealed class ConnectionGroupCardState
        {
            public int Id { get; set; }
            public string GroupName { get; set; }
            public string Server { get; set; }
            public string Database { get; set; }
            public int Priority { get; set; }
            public int? ColorIndex { get; set; }
            public bool IsEditing { get; set; }
            public bool IsNew { get; set; }
            public string SnapshotGroupName { get; set; }
            public string SnapshotServer { get; set; }
            public string SnapshotDatabase { get; set; }
            public int SnapshotPriority { get; set; }
            public int? SnapshotColorIndex { get; set; }
        }

        private sealed class InlineColorChoice
        {
            public int? Index { get; set; }
            public string DisplayName { get; set; }
            public Brush SwatchBrush { get; set; }
            public Brush SwatchBorderBrush { get; set; }
        }

        private sealed class ColorPaletteItem
        {
            public int Index { get; set; }
            public string Name { get; set; }
            public Color Color { get; set; }
        }

        private static readonly string DefaultAutoConfigureValue = "server db";
        private const bool DefaultEnableLogging = false;
        private const bool DefaultEnableVerboseLogging = false;
        private const bool DefaultEnableAutoRename = true;
        private const bool DefaultEnableAutoColor = true;
        private const bool DefaultEnableConfigurePrompt = true;
        private const bool DefaultEnableServerAliasPrompt = true;
        private const bool DefaultEnableUpdateChecks = true;
        private const bool DefaultEnableRemoveDotSql = true;
        private const bool DefaultEnableLineIndicatorColor = true;
        private const bool DefaultEnableStatusBarColor = true;
        private const string DefaultSuggestedGroupNameStyle = "[serverAlias] [db]";
        private const string DefaultNewQueryRenameStyle = "[#]. [groupName]";
        private const string DefaultSavedFileRenameStyle = "[filename]";
        private const string DefaultConfigResourceName = "SSMS_EnvTabs.DefaultTabGroupConfig.json";

        private static readonly IReadOnlyList<ColorPaletteItem> ColorPalette =
            new List<ColorPaletteItem>
            {
                new ColorPaletteItem { Index = 0, Name = "Lavender", Color = (Color)ColorConverter.ConvertFromString("#9083ef") },
                new ColorPaletteItem { Index = 1, Name = "Gold", Color = (Color)ColorConverter.ConvertFromString("#d0b132") },
                new ColorPaletteItem { Index = 2, Name = "Cyan", Color = (Color)ColorConverter.ConvertFromString("#30b1cd") },
                new ColorPaletteItem { Index = 3, Name = "Burgundy", Color = (Color)ColorConverter.ConvertFromString("#cf6468") },
                new ColorPaletteItem { Index = 4, Name = "Green", Color = (Color)ColorConverter.ConvertFromString("#6ba12a") },
                new ColorPaletteItem { Index = 5, Name = "Brown", Color = (Color)ColorConverter.ConvertFromString("#bc8f6f") },
                new ColorPaletteItem { Index = 6, Name = "Royal Blue", Color = (Color)ColorConverter.ConvertFromString("#5bb2fa") },
                new ColorPaletteItem { Index = 7, Name = "Pumpkin", Color = (Color)ColorConverter.ConvertFromString("#d67441") },
                new ColorPaletteItem { Index = 8, Name = "Gray", Color = (Color)ColorConverter.ConvertFromString("#bdbcbc") },
                new ColorPaletteItem { Index = 9, Name = "Volt", Color = (Color)ColorConverter.ConvertFromString("#cbcc38") },
                new ColorPaletteItem { Index = 10, Name = "Teal", Color = (Color)ColorConverter.ConvertFromString("#2aa0a4") },
                new ColorPaletteItem { Index = 11, Name = "Magenta", Color = (Color)ColorConverter.ConvertFromString("#d957a7") },
                new ColorPaletteItem { Index = 12, Name = "Mint", Color = (Color)ColorConverter.ConvertFromString("#6bc6a5") },
                new ColorPaletteItem { Index = 13, Name = "Dark Brown", Color = (Color)ColorConverter.ConvertFromString("#946a5b") },
                new ColorPaletteItem { Index = 14, Name = "Blue", Color = (Color)ColorConverter.ConvertFromString("#6a8ec6") },
                new ColorPaletteItem { Index = 15, Name = "Pink", Color = (Color)ColorConverter.ConvertFromString("#e0a3a5") }
            };

        public SettingsToolWindowControl()
        {
            LoadView();
            Loaded += OnLoaded;
            IsVisibleChanged += OnIsVisibleChanged;
            Unloaded += OnUnloaded;
        }

        private void LoadView()
        {
            var viewUri = new Uri("/SSMS EnvTabs;component/SettingsToolWindowControl.xaml", UriKind.Relative);
            Application.LoadComponent(this, viewUri);

            openJsonButton = FindRequiredControl<Button>("OpenJsonButton");
            resetDefaultsButton = FindRequiredControl<Button>("ResetDefaultsButton");
            reloadButton = FindRequiredControl<Button>("ReloadButton");
            updateEnvTabsButton = FindRequiredControl<Button>("UpdateEnvTabsButton");
            copyLogPathButton = FindRequiredControl<Button>("CopyLogPathButton");
            settingsTabControl = FindRequiredControl<TabControl>("SettingsTabControl");
            generalSettingsTab = FindRequiredControl<TabItem>("GeneralSettingsTab");
            styleTemplatesTab = FindRequiredControl<TabItem>("StyleTemplatesTab");
            connectionGroupsTab = FindRequiredControl<TabItem>("ConnectionGroupsTab");
            autoConfigureCombo = FindRequiredControl<ComboBox>("AutoConfigureCombo");
            loggingToggle = FindRequiredControl<CheckBox>("LoggingToggle");
            autoRenameToggle = FindRequiredControl<CheckBox>("AutoRenameToggle");
            autoColorToggle = FindRequiredControl<CheckBox>("AutoColorToggle");
            configurePromptToggle = FindRequiredControl<CheckBox>("ConfigurePromptToggle");
            serverAliasPromptToggle = FindRequiredControl<CheckBox>("ServerAliasPromptToggle");
            updateChecksToggle = FindRequiredControl<CheckBox>("UpdateChecksToggle");
            removeDotSqlToggle = FindRequiredControl<CheckBox>("RemoveDotSqlToggle");
            lineIndicatorColorToggle = FindRequiredControl<CheckBox>("LineIndicatorColorToggle");
            statusBarColorToggle = FindRequiredControl<CheckBox>("StatusBarColorToggle");
            suggestedGroupNameStyleTextBox = FindRequiredControl<TextBox>("SuggestedGroupNameStyleTextBox");
            newQueryRenameStyleTextBox = FindRequiredControl<TextBox>("NewQueryRenameStyleTextBox");
            savedFileRenameStyleTextBox = FindRequiredControl<TextBox>("SavedFileRenameStyleTextBox");
            suggestedGroupNameEditButton = FindRequiredControl<Button>("SuggestedGroupNameEditButton");
            suggestedGroupNameSaveButton = FindRequiredControl<Button>("SuggestedGroupNameSaveButton");
            suggestedGroupNameCancelButton = FindRequiredControl<Button>("SuggestedGroupNameCancelButton");
            newQueryRenameEditButton = FindRequiredControl<Button>("NewQueryRenameEditButton");
            newQueryRenameSaveButton = FindRequiredControl<Button>("NewQueryRenameSaveButton");
            newQueryRenameCancelButton = FindRequiredControl<Button>("NewQueryRenameCancelButton");
            savedFileRenameEditButton = FindRequiredControl<Button>("SavedFileRenameEditButton");
            savedFileRenameSaveButton = FindRequiredControl<Button>("SavedFileRenameSaveButton");
            savedFileRenameCancelButton = FindRequiredControl<Button>("SavedFileRenameCancelButton");
            suggestedGroupNameTokenPanel = FindRequiredControl<StackPanel>("SuggestedGroupNameTokenPanel");
            newQueryRenameTokenPanel = FindRequiredControl<StackPanel>("NewQueryRenameTokenPanel");
            savedFileRenameTokenPanel = FindRequiredControl<StackPanel>("SavedFileRenameTokenPanel");
            suggestedGroupNameTokenHelpText = FindRequiredControl<TextBlock>("SuggestedGroupNameTokenHelpText");
            newQueryRenameTokenHelpText = FindRequiredControl<TextBlock>("NewQueryRenameTokenHelpText");
            savedFileRenameTokenHelpText = FindRequiredControl<TextBlock>("SavedFileRenameTokenHelpText");
            serverAliasSectionPanel = FindRequiredControl<StackPanel>("ServerAliasSectionPanel");
            serverAliasColumnHeaderGrid = FindRequiredControl<Grid>("ServerAliasColumnHeaderGrid");
            serverAliasRowsPanel = FindRequiredControl<StackPanel>("ServerAliasRowsPanel");
            connectionGroupCardsPanel = FindRequiredControl<StackPanel>("ConnectionGroupCardsPanel");
            connectionGroupsDescriptionText = FindRequiredControl<TextBlock>("ConnectionGroupsDescriptionText");
            checkUpdatesRowGrid = FindRequiredControl<Grid>("CheckUpdatesRowGrid");
            addServerAliasButton = FindRequiredControl<Button>("AddServerAliasButton");
            addConnectionGroupButton = FindRequiredControl<Button>("AddConnectionGroupButton");
            statusText = FindRequiredControl<TextBlock>("StatusText");

            openJsonButton.Click += OpenJsonButton_Click;
            resetDefaultsButton.Click += ResetDefaultsButton_Click;
            reloadButton.Click += ReloadButton_Click;
            updateEnvTabsButton.Click += UpdateEnvTabsButton_Click;
            copyLogPathButton.Click += CopyLogPathButton_Click;
            addServerAliasButton.Click += AddServerAliasButton_Click;
            addConnectionGroupButton.Click += AddConnectionGroupButton_Click;
            autoConfigureCombo.SelectionChanged += AutoConfigureCombo_SelectionChanged;
            settingsTabControl.SelectionChanged += SettingsTabControl_SelectionChanged;

            WireToggle(loggingToggle);
            WireToggle(autoRenameToggle);
            WireToggle(autoColorToggle);
            WireToggle(configurePromptToggle);
            WireToggle(serverAliasPromptToggle);
            WireToggle(updateChecksToggle);
            WireToggle(removeDotSqlToggle);
            WireToggle(lineIndicatorColorToggle);
            WireToggle(statusBarColorToggle);
            PreviewKeyDown += SettingsToolWindowControl_PreviewKeyDown;
            UpdateStyleEditUi();
        }

        private void SettingsToolWindowControl_PreviewKeyDown(object sender, KeyEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (e.Key != Key.Enter && e.Key != Key.Escape)
            {
                return;
            }

            bool handled = e.Key == Key.Enter
                ? TryInvokePrimaryActionFromKeyboard()
                : TryInvokeCancelActionFromKeyboard();

            if (handled)
            {
                e.Handled = true;
            }
        }

        private bool TryInvokePrimaryActionFromKeyboard()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (activeStyleEditField != EditableStyleField.None)
            {
                StyleSaveButton_Click(GetSaveButton(activeStyleEditField), new RoutedEventArgs());
                return true;
            }

            if (activeAliasEditRow != null)
            {
                SaveServerAliasRow(activeAliasEditRow, activeAliasServerTextBox?.Text, activeAliasAliasTextBox?.Text);
                return true;
            }

            if (activeConnectionGroupEditCard != null)
            {
                SaveConnectionGroupCard(
                    activeConnectionGroupEditCard,
                    ShouldShowGroupNameInConnectionCards() ? activeConnectionGroupNameTextBox?.Text : activeConnectionGroupEditCard.GroupName,
                    activeConnectionServerTextBox?.Text,
                    activeConnectionDatabaseTextBox?.Text,
                    activeConnectionPriorityTextBox?.Text,
                    activeConnectionColorCombo?.SelectedValue);
                return true;
            }

            return false;
        }

        private bool TryInvokeCancelActionFromKeyboard()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (activeStyleEditField != EditableStyleField.None)
            {
                StyleCancelButton_Click(GetCancelButton(activeStyleEditField), new RoutedEventArgs());
                return true;
            }

            if (activeAliasEditRow != null)
            {
                CancelServerAliasEdit(activeAliasEditRow);
                return true;
            }

            if (activeConnectionGroupEditCard != null)
            {
                CancelConnectionGroupEdit(activeConnectionGroupEditCard);
                return true;
            }

            return false;
        }

        private void WireToggle(CheckBox toggle)
        {
            if (toggle == null)
            {
                return;
            }

            toggle.Checked += OptionToggleChanged;
            toggle.Unchecked += OptionToggleChanged;
        }

        private T FindRequiredControl<T>(string name) where T : class
        {
            var element = FindName(name) as T;
            if (element == null)
            {
                throw new InvalidOperationException($"Missing required control '{name}' in SettingsToolWindowControl.xaml.");
            }

            return element;
        }

        private void OnLoaded(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            Loaded -= OnLoaded;
            SubscribeToThemeChanges();
            ApplyCurrentThemeAndRefreshUi(rebuildDynamicRows: false);
            InitializeAutoConfigureOptions();
            ApplyComboBoxPopupTheme(autoConfigureCombo);
            ApplyComboBoxTemplate(autoConfigureCombo);
            LoadFromConfig();
        }

        private void OnUnloaded(object sender, RoutedEventArgs e)
        {
            UnsubscribeFromThemeChanges();
        }

        private void OnIsVisibleChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (IsVisible)
            {
                SubscribeToThemeChanges();
                ApplyCurrentThemeAndRefreshUi(rebuildDynamicRows: true);
            }
        }

        private void SubscribeToThemeChanges()
        {
            if (isThemeEventSubscribed)
            {
                return;
            }

            VSColorTheme.ThemeChanged += OnVsThemeChanged;
            isThemeEventSubscribed = true;
        }

        private void UnsubscribeFromThemeChanges()
        {
            if (!isThemeEventSubscribed)
            {
                return;
            }

            VSColorTheme.ThemeChanged -= OnVsThemeChanged;
            isThemeEventSubscribed = false;
        }

        private void OnVsThemeChanged(ThemeChangedEventArgs e)
        {
            ThreadHelper.JoinableTaskFactory.Run(async delegate
            {
                await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync();
                ApplyCurrentThemeAndRefreshUi(rebuildDynamicRows: true);
            });
        }

        private void ApplyCurrentThemeAndRefreshUi(bool rebuildDynamicRows)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            ApplyThemeBrushResources();
            ApplyComboBoxPopupTheme(autoConfigureCombo);
            ApplyComboBoxTemplate(autoConfigureCombo);

            if (rebuildDynamicRows)
            {
                RebuildServerAliasRowsUi();
                RebuildConnectionGroupCardsUi();
            }
        }

        private void ApplyThemeBrushResources()
        {
            Brush bg = ResolveBrush(EnvironmentColors.ToolWindowBackgroundBrushKey)
                ?? SystemColors.WindowBrush;
            Brush fg = ResolveBrush(EnvironmentColors.ToolWindowTextBrushKey)
                ?? SystemColors.WindowTextBrush;
            Brush border = ResolveBrush(EnvironmentColors.ToolWindowBorderBrushKey)
                ?? SystemColors.ActiveBorderBrush;
            Brush link = ResolveBrush(EnvironmentColors.ControlLinkTextBrushKey)
                ?? SystemColors.HighlightBrush;
            Brush toggleOn = ResolveToggleAccentBrush(border, out string toggleAccentSource);

            Color bgColor = GetBrushColor(bg, Colors.White);
            Color fgColor = GetBrushColor(fg, Colors.Black);
            bool isLightTheme = GetRelativeLuminance(bgColor) > 0.6;

            Color neutralBaseColor = isLightTheme
                ? BlendColors(bgColor, Colors.Black, 0.06)
                : BlendColors(bgColor, Colors.White, 0.10);
            Color neutralHoverColor = isLightTheme
                ? BlendColors(neutralBaseColor, Colors.Black, 0.06)
                : BlendColors(neutralBaseColor, Colors.White, 0.18);
            Color neutralPressedColor = isLightTheme
                ? BlendColors(neutralBaseColor, Colors.Black, 0.12)
                : BlendColors(neutralBaseColor, Colors.White, 0.28);

            Color accentColor = GetBrushColor(toggleOn, GetBrushColor(border, fgColor));
            Color primaryHoverColor = BlendColors(accentColor, Colors.Black, 0.08);
            Color primaryPressedColor = BlendColors(accentColor, Colors.Black, 0.16);
            Color primaryTextColor = GetRelativeLuminance(accentColor) > 0.58 ? Colors.Black : Colors.White;

            Color dangerBaseColor = GetBrushColor(
                TryFindResource("EnvTabsButtonDangerBrush") as Brush,
                (Color)ColorConverter.ConvertFromString("#B42318"));
            Color dangerHoverColor = BlendColors(dangerBaseColor, Colors.Black, 0.08);
            Color dangerPressedColor = BlendColors(dangerBaseColor, Colors.Black, 0.16);

            Color tokenBaseColor = neutralBaseColor;
            Color tokenHoverColor = neutralHoverColor;
            Color tokenPressedColor = neutralPressedColor;

            Color tabHeaderColor = isLightTheme
                ? BlendColors(bgColor, Colors.Black, 0.05)
                : BlendColors(bgColor, Colors.White, 0.06);
            Color tabHoverColor = isLightTheme
                ? BlendColors(bgColor, Colors.Black, 0.10)
                : BlendColors(bgColor, Colors.White, 0.12);
            Color tabSelectedColor = isLightTheme
                ? BlendColors(bgColor, Colors.Black, 0.16)
                : BlendColors(bgColor, Colors.White, 0.18);

            Resources["EnvTabsBackgroundBrush"] = bg;
            Resources["EnvTabsForegroundBrush"] = fg;
            Resources["EnvTabsBorderBrush"] = border;
            Resources["EnvTabsLinkBrush"] = link;
            Resources["EnvTabsToggleOnBrush"] = toggleOn;
            Resources["EnvTabsButtonNeutralBrush"] = new SolidColorBrush(neutralBaseColor);
            Resources["EnvTabsButtonNeutralHoverBrush"] = new SolidColorBrush(neutralHoverColor);
            Resources["EnvTabsButtonNeutralPressedBrush"] = new SolidColorBrush(neutralPressedColor);
            Resources["EnvTabsButtonPrimaryHoverBrush"] = new SolidColorBrush(primaryHoverColor);
            Resources["EnvTabsButtonPrimaryPressedBrush"] = new SolidColorBrush(primaryPressedColor);
            Resources["EnvTabsButtonPrimaryTextBrush"] = new SolidColorBrush(primaryTextColor);
            Resources["EnvTabsButtonDangerHoverBrush"] = new SolidColorBrush(dangerHoverColor);
            Resources["EnvTabsButtonDangerPressedBrush"] = new SolidColorBrush(dangerPressedColor);
            Resources["EnvTabsButtonDangerTextBrush"] = new SolidColorBrush(dangerBaseColor);
            Resources["EnvTabsTokenButtonBackgroundBrush"] = new SolidColorBrush(tokenBaseColor);
            Resources["EnvTabsTokenButtonHoverBrush"] = new SolidColorBrush(tokenHoverColor);
            Resources["EnvTabsTokenButtonPressedBrush"] = new SolidColorBrush(tokenPressedColor);
            Resources["EnvTabsButtonHoverBrush"] = new SolidColorBrush(neutralHoverColor);
            Resources["EnvTabsButtonHoverTextBrush"] = new SolidColorBrush(fgColor);
            Resources["EnvTabsButtonPressedBrush"] = new SolidColorBrush(neutralPressedColor);
            Resources["EnvTabsTabHeaderBackgroundBrush"] = new SolidColorBrush(tabHeaderColor);
            Resources["EnvTabsTabHeaderForegroundBrush"] = new SolidColorBrush(fgColor);
            Resources["EnvTabsTabHeaderHoverBrush"] = new SolidColorBrush(tabHoverColor);
            Resources["EnvTabsTabHeaderSelectedBrush"] = new SolidColorBrush(tabSelectedColor);

            Color toggleColor = GetBrushColor(toggleOn, Colors.Transparent);
            EnvTabsLog.Info($"SettingsToolWindowControl toggle accent source={toggleAccentSource}, color=#{toggleColor.R:X2}{toggleColor.G:X2}{toggleColor.B:X2}");
        }

        private Brush ResolveBrush(object resourceKey)
        {
            if (resourceKey == null)
            {
                return null;
            }

            return TryFindResource(resourceKey) as Brush
                ?? Application.Current?.TryFindResource(resourceKey) as Brush;
        }

        private Brush ResolveEnvironmentBrushByName(string keyName)
        {
            PropertyInfo property = typeof(EnvironmentColors).GetProperty(keyName, BindingFlags.Public | BindingFlags.Static);
            object key = property?.GetValue(null);
            return ResolveBrush(key);
        }

        private Brush ResolveToggleAccentBrush(Brush fallback, out string source)
        {
            Brush mainWindowAccent = ResolveEnvironmentBrushByName("MainWindowActiveDefaultBorderBrushKey");
            if (mainWindowAccent != null)
            {
                source = "MainWindowActiveDefaultBorderBrushKey";
                return mainWindowAccent;
            }

            source = "fallback";
            return fallback;
        }

        private void InitializeAutoConfigureOptions()
        {
            autoConfigureCombo.ItemsSource = new[]
            {
                new AutoConfigureItem { DisplayName = "Server", Value = "server" },
                new AutoConfigureItem { DisplayName = "Server + Database", Value = "server db" },
                new AutoConfigureItem { DisplayName = "Off", Value = "off" }
            };

            autoConfigureCombo.DisplayMemberPath = nameof(AutoConfigureItem.DisplayName);
            autoConfigureCombo.SelectedValuePath = nameof(AutoConfigureItem.Value);
        }

        private void LoadFromConfig()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                suppressAutoSave = true;

                TabGroupConfigLoader.EnsureDefaultConfigExists();
                TabGroupConfig config = TabGroupConfigLoader.LoadOrNull() ?? new TabGroupConfig();
                TabGroupSettings settings = config.Settings ?? new TabGroupSettings();

                loggingToggle.IsChecked = settings.EnableLogging || settings.EnableVerboseLogging;
                autoRenameToggle.IsChecked = settings.EnableAutoRename;
                autoColorToggle.IsChecked = settings.EnableAutoColor;
                configurePromptToggle.IsChecked = settings.EnableConfigurePrompt;
                serverAliasPromptToggle.IsChecked = settings.EnableServerAliasPrompt;
                updateChecksToggle.IsChecked = settings.EnableUpdateChecks;
                removeDotSqlToggle.IsChecked = settings.EnableRemoveDotSql;
                lineIndicatorColorToggle.IsChecked = settings.EnableLineIndicatorColor;
                statusBarColorToggle.IsChecked = settings.EnableStatusBarColor;
                suggestedGroupNameStyleTextBox.Text = NormalizeStyleValue(settings.SuggestedGroupNameStyle, DefaultSuggestedGroupNameStyle);
                newQueryRenameStyleTextBox.Text = NormalizeStyleValue(settings.NewQueryRenameStyle, DefaultNewQueryRenameStyle);
                savedFileRenameStyleTextBox.Text = NormalizeStyleValue(settings.SavedFileRenameStyle, DefaultSavedFileRenameStyle);

                autoConfigureCombo.SelectedValue = NormalizeAutoConfigure(settings.AutoConfigure);
                currentAutoConfigureMode = NormalizeAutoConfigure(settings.AutoConfigure);
                LoadConnectionGroupsTabState(config);
                ExitStyleEditMode(restoreSnapshot: false, setStatus: false);
                statusText.Text = "Settings loaded from TabGroupConfig.json";
            }
            catch (Exception ex)
            {
                statusText.Text = "Failed to load settings.";
                EnvTabsLog.Info($"SettingsToolWindowControl load failed: {ex.Message}");
            }
            finally
            {
                suppressAutoSave = false;
            }
        }

        private static string NormalizeAutoConfigure(string value)
        {
            if (string.Equals(value, "server", StringComparison.OrdinalIgnoreCase))
            {
                return "server";
            }

            if (string.Equals(value, "server db", StringComparison.OrdinalIgnoreCase))
            {
                return "server db";
            }

            if (string.Equals(value, "off", StringComparison.OrdinalIgnoreCase))
            {
                return "off";
            }

            return DefaultAutoConfigureValue;
        }

        private void SaveSettingsTabFromUi(string statusMessage)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                TabGroupConfigLoader.EnsureDefaultConfigExists();

                TabGroupConfig config = TabGroupConfigLoader.LoadOrNull() ?? new TabGroupConfig();
                if (config.Settings == null)
                {
                    config.Settings = new TabGroupSettings();
                }

                bool loggingEnabled = loggingToggle.IsChecked == true;

                config.Settings.EnableLogging = loggingEnabled;
                config.Settings.EnableVerboseLogging = loggingEnabled;
                config.Settings.EnableAutoRename = autoRenameToggle.IsChecked == true;
                config.Settings.EnableAutoColor = autoColorToggle.IsChecked == true;
                config.Settings.EnableConfigurePrompt = configurePromptToggle.IsChecked == true;
                config.Settings.EnableServerAliasPrompt = serverAliasPromptToggle.IsChecked == true;
                config.Settings.EnableUpdateChecks = updateChecksToggle.IsChecked == true;
                config.Settings.EnableRemoveDotSql = removeDotSqlToggle.IsChecked == true;
                config.Settings.EnableLineIndicatorColor = lineIndicatorColorToggle.IsChecked == true;
                config.Settings.EnableStatusBarColor = statusBarColorToggle.IsChecked == true;

                string selectedAutoConfigure = autoConfigureCombo.SelectedValue as string;
                config.Settings.AutoConfigure = NormalizeAutoConfigure(selectedAutoConfigure);

                TabGroupConfigLoader.SaveConfig(config);

                EnvTabsLog.Enabled = config.Settings.EnableLogging;
                EnvTabsLog.VerboseEnabled = config.Settings.EnableVerboseLogging;
                SsmsSettingsUpdater.EnsureRegexTabColorizationEnabled(config.Settings.EnableAutoColor);

                statusText.Text = statusMessage;
            }
            catch (Exception ex)
            {
                statusText.Text = "Failed to save settings.";
                EnvTabsLog.Info($"SettingsToolWindowControl save failed: {ex.Message}");
            }
        }

        private void SaveStyleTabFromUi(string statusMessage)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                TabGroupConfigLoader.EnsureDefaultConfigExists();

                TabGroupConfig config = TabGroupConfigLoader.LoadOrNull() ?? new TabGroupConfig();
                if (config.Settings == null)
                {
                    config.Settings = new TabGroupSettings();
                }

                config.Settings.SuggestedGroupNameStyle = GetPersistedStyleValue(EditableStyleField.SuggestedGroupName);
                config.Settings.NewQueryRenameStyle = GetPersistedStyleValue(EditableStyleField.NewQueryRename);
                config.Settings.SavedFileRenameStyle = GetPersistedStyleValue(EditableStyleField.SavedFileRename);

                TabGroupConfigLoader.SaveConfig(config);
                statusText.Text = statusMessage;
            }
            catch (Exception ex)
            {
                statusText.Text = "Failed to save style settings.";
                EnvTabsLog.Info($"SettingsToolWindowControl style save failed: {ex.Message}");
            }
        }

        private void OptionToggleChanged(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (suppressAutoSave)
            {
                return;
            }

            SaveSettingsTabFromUi("Settings saved.");

            if (sender == serverAliasPromptToggle || sender == autoRenameToggle)
            {
                ReloadConnectionGroupsTabFromConfig();
            }
        }

        private void AutoConfigureCombo_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (suppressAutoSave)
            {
                return;
            }

            SaveSettingsTabFromUi("Settings saved.");
            currentAutoConfigureMode = NormalizeAutoConfigure(autoConfigureCombo.SelectedValue as string);
            RebuildConnectionGroupCardsUi();
        }

        private void SettingsTabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (e.OriginalSource != settingsTabControl)
            {
                return;
            }

            if (activeStyleEditField != EditableStyleField.None && !IsStyleTabFocused())
            {
                ExitStyleEditMode(restoreSnapshot: true, setStatus: false);
            }

            if (!IsConnectionGroupsTabFocused())
            {
                CancelAllConnectionTabEdits(setStatus: false);
            }
        }

        private void ReloadButton_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            ExitStyleEditMode(restoreSnapshot: true, setStatus: false);
            LoadFromConfig();
        }

        private void UpdateEnvTabsButton_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                var config = TabGroupConfigLoader.LoadOrNull();
                UpdateChecker.CheckNow(SSMS_EnvTabsPackage.Instance, config?.Settings, ignoreSettings: true);
            }
            catch (Exception ex)
            {
                statusText.Text = "Failed to check updates.";
                EnvTabsLog.Info($"SettingsToolWindowControl update check failed: {ex.Message}");
            }
        }

        private void ResetDefaultsButton_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            MessageBoxResult confirmation = MessageBox.Show(
                IsStyleTabFocused()
                    ? "This will reset style templates in the current tab to default values. Continue?"
                    : IsConnectionGroupsTabFocused()
                        ? showServerAliasSection
                            ? "This will reset connection groups and server aliases in the current tab to default values. Continue?"
                            : "This will reset connection groups in the current tab to default values. Continue?"
                    : "This will reset settings in the current tab to default values. Continue?",
                "Reset settings to defaults",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (confirmation != MessageBoxResult.Yes)
            {
                return;
            }

            try
            {
                suppressAutoSave = true;

                if (IsStyleTabFocused())
                {
                    ExitStyleEditMode(restoreSnapshot: false, setStatus: false);
                    suggestedGroupNameStyleTextBox.Text = DefaultSuggestedGroupNameStyle;
                    newQueryRenameStyleTextBox.Text = DefaultNewQueryRenameStyle;
                    savedFileRenameStyleTextBox.Text = DefaultSavedFileRenameStyle;
                }
                else if (IsConnectionGroupsTabFocused())
                {
                    ResetConnectionGroupsTabToDefaults();
                    return;
                }
                else
                {
                    loggingToggle.IsChecked = DefaultEnableLogging || DefaultEnableVerboseLogging;
                    autoRenameToggle.IsChecked = DefaultEnableAutoRename;
                    autoColorToggle.IsChecked = DefaultEnableAutoColor;
                    configurePromptToggle.IsChecked = DefaultEnableConfigurePrompt;
                    serverAliasPromptToggle.IsChecked = DefaultEnableServerAliasPrompt;
                    updateChecksToggle.IsChecked = DefaultEnableUpdateChecks;
                    removeDotSqlToggle.IsChecked = DefaultEnableRemoveDotSql;
                    lineIndicatorColorToggle.IsChecked = DefaultEnableLineIndicatorColor;
                    statusBarColorToggle.IsChecked = DefaultEnableStatusBarColor;
                    autoConfigureCombo.SelectedValue = DefaultAutoConfigureValue;
                }
            }
            finally
            {
                suppressAutoSave = false;
            }

            if (IsStyleTabFocused())
            {
                SaveStyleTabFromUi("Style templates reset to defaults.");
            }
            else if (IsConnectionGroupsTabFocused())
            {
                // ResetConnectionGroupsTabToDefaults handles save and status.
            }
            else
            {
                SaveSettingsTabFromUi("Settings reset to defaults.");
            }
        }

        private bool IsStyleTabFocused()
        {
            return settingsTabControl?.SelectedItem == styleTemplatesTab;
        }

        private bool IsConnectionGroupsTabFocused()
        {
            return settingsTabControl?.SelectedItem == connectionGroupsTab;
        }

        private void CopyLogPathButton_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                string logPath = ResolveRuntimeLogPath();
                Clipboard.SetText(logPath);
                statusText.Text = "Runtime log path copied.";
            }
            catch (Exception ex)
            {
                statusText.Text = "Failed to copy runtime log path.";
                EnvTabsLog.Info($"SettingsToolWindowControl copy log path failed: {ex.Message}");
            }
        }

        private static string ResolveRuntimeLogPath()
        {
            string baseDir = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            if (string.IsNullOrWhiteSpace(baseDir))
            {
                baseDir = IOPath.GetTempPath();
            }

            return IOPath.Combine(baseDir, "SSMS EnvTabs", "runtime.log");
        }

        private void ApplyComboBoxPopupTheme(ComboBox combo)
        {
            if (combo == null)
            {
                return;
            }

            Brush background = ResolveBrush(EnvironmentColors.ToolWindowBackgroundBrushKey) ?? SystemColors.WindowBrush;
            Brush foreground = ResolveBrush(EnvironmentColors.ToolWindowTextBrushKey) ?? SystemColors.WindowTextBrush;
            Brush highlight = ResolveBrush(SystemColors.HighlightBrushKey) ?? SystemColors.HighlightBrush;
            Brush highlightText = ResolveBrush(SystemColors.HighlightTextBrushKey) ?? SystemColors.HighlightTextBrush;
            Brush inactiveHighlight = ResolveBrush(SystemColors.InactiveSelectionHighlightBrushKey) ?? background;
            Brush inactiveHighlightText = ResolveBrush(SystemColors.InactiveSelectionHighlightTextBrushKey) ?? foreground;

            combo.Resources[SystemColors.WindowBrushKey] = background;
            combo.Resources[SystemColors.ControlBrushKey] = background;
            combo.Resources[SystemColors.WindowTextBrushKey] = foreground;
            combo.Resources[SystemColors.ControlTextBrushKey] = foreground;
            combo.Resources[SystemColors.HighlightBrushKey] = highlight;
            combo.Resources[SystemColors.HighlightTextBrushKey] = highlightText;
            combo.Resources[SystemColors.InactiveSelectionHighlightBrushKey] = inactiveHighlight;
            combo.Resources[SystemColors.InactiveSelectionHighlightTextBrushKey] = inactiveHighlightText;

            Color bgColor = GetBrushColor(background, Colors.White);
            bool isLightTheme = GetRelativeLuminance(bgColor) > 0.6;

            if (isLightTheme)
            {
                Color hoverColor = BlendColors(bgColor, Colors.Black, 0.06);
                Color selectedColor = BlendColors(bgColor, Colors.Black, 0.12);
                combo.Resources["EnvTabsComboHoverBrush"] = new SolidColorBrush(hoverColor);
                combo.Resources["EnvTabsComboSelectedBrush"] = new SolidColorBrush(selectedColor);
                combo.Resources["EnvTabsComboHoverTextBrush"] = new SolidColorBrush(Colors.Black);
                combo.Resources["EnvTabsComboSelectedTextBrush"] = new SolidColorBrush(Colors.Black);
                combo.Resources["EnvTabsComboButtonHoverBrush"] = new SolidColorBrush(BlendColors(bgColor, Colors.Black, 0.08));
                combo.Resources["EnvTabsComboButtonPressedBrush"] = new SolidColorBrush(BlendColors(bgColor, Colors.Black, 0.12));
            }
            else
            {
                Color selectedColor = ApplyAlpha(GetBrushColor(highlight, BlendColors(bgColor, Colors.White, 0.18)), 0.78);
                Color selectedTextColor = GetBrushColor(highlightText, Colors.White);
                Color hoverColor = BlendColors(bgColor, Colors.White, 0.12);
                Color hoverTextColor = GetBrushColor(foreground, Colors.White);

                combo.Resources["EnvTabsComboHoverBrush"] = new SolidColorBrush(hoverColor);
                combo.Resources["EnvTabsComboSelectedBrush"] = new SolidColorBrush(selectedColor);
                combo.Resources["EnvTabsComboHoverTextBrush"] = new SolidColorBrush(hoverTextColor);
                combo.Resources["EnvTabsComboSelectedTextBrush"] = new SolidColorBrush(selectedTextColor);
                combo.Resources["EnvTabsComboButtonHoverBrush"] = new SolidColorBrush(BlendColors(bgColor, Colors.White, 0.06));
                combo.Resources["EnvTabsComboButtonPressedBrush"] = new SolidColorBrush(BlendColors(bgColor, Colors.White, 0.10));
            }

            combo.SetResourceReference(Control.BorderBrushProperty, EnvironmentColors.ToolWindowBorderBrushKey);
            combo.BorderThickness = new Thickness(1);
            combo.Foreground = foreground;
            combo.Background = background;
        }

        private void ApplyComboBoxTemplate(ComboBox combo)
        {
            if (combo == null)
            {
                return;
            }

            combo.Template = CreateThemedComboBoxTemplate();
        }

        private static ControlTemplate CreateThemedComboBoxTemplate()
        {
            var template = new ControlTemplate(typeof(ComboBox));

            var root = new FrameworkElementFactory(typeof(Grid));
            template.VisualTree = root;

            var toggle = new FrameworkElementFactory(typeof(ToggleButton));
            toggle.Name = "ToggleButton";
            toggle.SetValue(ToggleButton.FocusableProperty, false);
            toggle.SetValue(ToggleButton.ClickModeProperty, ClickMode.Press);
            toggle.SetBinding(Control.BackgroundProperty, new Binding("Background")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent)
            });
            toggle.SetBinding(Control.BorderBrushProperty, new Binding("BorderBrush")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent)
            });
            toggle.SetBinding(Control.BorderThicknessProperty, new Binding("BorderThickness")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent)
            });
            toggle.SetBinding(Control.ForegroundProperty, new Binding("Foreground")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent)
            });
            toggle.SetBinding(ToggleButton.IsCheckedProperty, new Binding("IsDropDownOpen")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent)
            });

            var toggleTemplate = new ControlTemplate(typeof(ToggleButton));

            var border = new FrameworkElementFactory(typeof(Border));
            border.Name = "Border";
            border.SetValue(Border.CornerRadiusProperty, new CornerRadius(2));
            border.SetValue(Border.BackgroundProperty, new TemplateBindingExtension(Control.BackgroundProperty));
            border.SetValue(Border.BorderBrushProperty, new TemplateBindingExtension(Control.BorderBrushProperty));
            border.SetValue(Border.BorderThicknessProperty, new TemplateBindingExtension(Control.BorderThicknessProperty));

            var contentDock = new FrameworkElementFactory(typeof(DockPanel));
            contentDock.SetValue(DockPanel.LastChildFillProperty, true);

            var arrow = new FrameworkElementFactory(typeof(ShapePath));
            arrow.Name = "Arrow";
            arrow.SetValue(FrameworkElement.WidthProperty, 8.0);
            arrow.SetValue(FrameworkElement.HeightProperty, 5.0);
            arrow.SetValue(FrameworkElement.MarginProperty, new Thickness(0, 0, 8, 0));
            arrow.SetValue(FrameworkElement.VerticalAlignmentProperty, VerticalAlignment.Center);
            arrow.SetValue(FrameworkElement.HorizontalAlignmentProperty, HorizontalAlignment.Right);
            arrow.SetValue(DockPanel.DockProperty, Dock.Right);
            arrow.SetValue(ShapePath.DataProperty, Geometry.Parse("M 0 0 L 4 4 L 8 0 Z"));
            arrow.SetValue(Shape.FillProperty, new DynamicResourceExtension(SystemColors.WindowTextBrushKey));

            var contentPresenter = new FrameworkElementFactory(typeof(ContentPresenter));
            contentPresenter.Name = "ContentSite";
            contentPresenter.SetValue(ContentPresenter.MarginProperty, new Thickness(6, 2, 4, 2));
            contentPresenter.SetValue(ContentPresenter.VerticalAlignmentProperty, VerticalAlignment.Center);
            contentPresenter.SetValue(ContentPresenter.HorizontalAlignmentProperty, HorizontalAlignment.Left);
            contentPresenter.SetValue(ContentPresenter.RecognizesAccessKeyProperty, true);
            contentPresenter.SetBinding(ContentPresenter.ContentProperty, new Binding("SelectedItem")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(ComboBox), 1)
            });
            contentPresenter.SetBinding(ContentPresenter.ContentTemplateProperty, new Binding("ItemTemplate")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(ComboBox), 1)
            });
            contentPresenter.SetBinding(ContentPresenter.ContentTemplateSelectorProperty, new Binding("ItemTemplateSelector")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(ComboBox), 1)
            });
            contentPresenter.SetBinding(TextElement.ForegroundProperty, new Binding("Foreground")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor, typeof(ComboBox), 1)
            });

            contentDock.AppendChild(arrow);
            contentDock.AppendChild(contentPresenter);
            border.AppendChild(contentDock);
            toggleTemplate.VisualTree = border;

            var toggleHoverTrigger = new Trigger { Property = UIElement.IsMouseOverProperty, Value = true };
            toggleHoverTrigger.Setters.Add(new Setter(Border.BackgroundProperty, new DynamicResourceExtension("EnvTabsComboButtonHoverBrush"), "Border"));
            toggleTemplate.Triggers.Add(toggleHoverTrigger);

            var toggleCheckedTrigger = new Trigger { Property = ToggleButton.IsCheckedProperty, Value = true };
            toggleCheckedTrigger.Setters.Add(new Setter(Border.BackgroundProperty, new DynamicResourceExtension("EnvTabsComboButtonPressedBrush"), "Border"));
            toggleTemplate.Triggers.Add(toggleCheckedTrigger);

            toggle.SetValue(Control.TemplateProperty, toggleTemplate);
            root.AppendChild(toggle);

            var popup = new FrameworkElementFactory(typeof(Popup));
            popup.Name = "PART_Popup";
            popup.SetValue(Popup.PlacementProperty, PlacementMode.Bottom);
            popup.SetValue(Popup.AllowsTransparencyProperty, true);
            popup.SetValue(Popup.FocusableProperty, false);
            popup.SetBinding(Popup.IsOpenProperty, new Binding("IsDropDownOpen")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent)
            });

            var popupBorder = new FrameworkElementFactory(typeof(Border));
            popupBorder.SetValue(Border.BackgroundProperty, new DynamicResourceExtension(SystemColors.WindowBrushKey));
            popupBorder.SetValue(Border.BorderBrushProperty, new TemplateBindingExtension(Control.BorderBrushProperty));
            popupBorder.SetValue(Border.BorderThicknessProperty, new TemplateBindingExtension(Control.BorderThicknessProperty));
            popupBorder.SetValue(Border.SnapsToDevicePixelsProperty, true);

            var scrollViewer = new FrameworkElementFactory(typeof(ScrollViewer));
            scrollViewer.SetValue(ScrollViewer.SnapsToDevicePixelsProperty, true);
            scrollViewer.SetValue(ScrollViewer.HorizontalScrollBarVisibilityProperty, ScrollBarVisibility.Auto);
            scrollViewer.SetValue(ScrollViewer.VerticalScrollBarVisibilityProperty, ScrollBarVisibility.Auto);
            scrollViewer.SetBinding(ScrollViewer.CanContentScrollProperty, new Binding("CanContentScroll")
            {
                RelativeSource = new RelativeSource(RelativeSourceMode.TemplatedParent)
            });

            var itemsPresenter = new FrameworkElementFactory(typeof(ItemsPresenter));
            scrollViewer.AppendChild(itemsPresenter);
            popupBorder.AppendChild(scrollViewer);
            popup.AppendChild(popupBorder);
            root.AppendChild(popup);

            var disabledTrigger = new Trigger { Property = UIElement.IsEnabledProperty, Value = false };
            disabledTrigger.Setters.Add(new Setter(Control.ForegroundProperty, new DynamicResourceExtension(SystemColors.GrayTextBrushKey)));
            template.Triggers.Add(disabledTrigger);

            return template;
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

        private static Color BlendColors(Color baseColor, Color blendColor, double blendAmount)
        {
            blendAmount = Math.Max(0, Math.Min(1, blendAmount));
            byte r = (byte)(baseColor.R + (blendColor.R - baseColor.R) * blendAmount);
            byte g = (byte)(baseColor.G + (blendColor.G - baseColor.G) * blendAmount);
            byte b = (byte)(baseColor.B + (blendColor.B - baseColor.B) * blendAmount);
            return Color.FromArgb(baseColor.A, r, g, b);
        }

        private static Color ApplyAlpha(Color color, double alphaFactor)
        {
            alphaFactor = Math.Max(0, Math.Min(1, alphaFactor));
            byte a = (byte)(color.A * alphaFactor);
            return Color.FromArgb(a, color.R, color.G, color.B);
        }

        private static string NormalizeStyleValue(string configuredValue, string defaultValue)
        {
            return string.IsNullOrWhiteSpace(configuredValue) ? defaultValue : configuredValue;
        }

        private string GetPersistedStyleValue(EditableStyleField field)
        {
            string currentValue;
            if (activeStyleEditField == field && styleEditSnapshot != null)
            {
                currentValue = styleEditSnapshot;
            }
            else
            {
                currentValue = GetStyleTextBox(field)?.Text;
            }

            return NormalizeStyleValue(currentValue, GetStyleDefault(field));
        }

        private static string GetStyleDefault(EditableStyleField field)
        {
            switch (field)
            {
                case EditableStyleField.SuggestedGroupName:
                    return DefaultSuggestedGroupNameStyle;
                case EditableStyleField.NewQueryRename:
                    return DefaultNewQueryRenameStyle;
                case EditableStyleField.SavedFileRename:
                    return DefaultSavedFileRenameStyle;
                default:
                    return string.Empty;
            }
        }

        private bool TryGetStyleFieldFromSender(object sender, out EditableStyleField field)
        {
            field = EditableStyleField.None;
            FrameworkElement element = sender as FrameworkElement;
            if (element == null)
            {
                return false;
            }

            string name = element.Name;
            if (name.StartsWith("SuggestedGroupName", StringComparison.Ordinal))
            {
                field = EditableStyleField.SuggestedGroupName;
                return true;
            }

            if (name.StartsWith("NewQueryRename", StringComparison.Ordinal))
            {
                field = EditableStyleField.NewQueryRename;
                return true;
            }

            if (name.StartsWith("SavedFileRename", StringComparison.Ordinal))
            {
                field = EditableStyleField.SavedFileRename;
                return true;
            }

            return false;
        }

        private TextBox GetStyleTextBox(EditableStyleField field)
        {
            switch (field)
            {
                case EditableStyleField.SuggestedGroupName:
                    return suggestedGroupNameStyleTextBox;
                case EditableStyleField.NewQueryRename:
                    return newQueryRenameStyleTextBox;
                case EditableStyleField.SavedFileRename:
                    return savedFileRenameStyleTextBox;
                default:
                    return null;
            }
        }

        private Button GetEditButton(EditableStyleField field)
        {
            switch (field)
            {
                case EditableStyleField.SuggestedGroupName:
                    return suggestedGroupNameEditButton;
                case EditableStyleField.NewQueryRename:
                    return newQueryRenameEditButton;
                case EditableStyleField.SavedFileRename:
                    return savedFileRenameEditButton;
                default:
                    return null;
            }
        }

        private Button GetSaveButton(EditableStyleField field)
        {
            switch (field)
            {
                case EditableStyleField.SuggestedGroupName:
                    return suggestedGroupNameSaveButton;
                case EditableStyleField.NewQueryRename:
                    return newQueryRenameSaveButton;
                case EditableStyleField.SavedFileRename:
                    return savedFileRenameSaveButton;
                default:
                    return null;
            }
        }

        private Button GetCancelButton(EditableStyleField field)
        {
            switch (field)
            {
                case EditableStyleField.SuggestedGroupName:
                    return suggestedGroupNameCancelButton;
                case EditableStyleField.NewQueryRename:
                    return newQueryRenameCancelButton;
                case EditableStyleField.SavedFileRename:
                    return savedFileRenameCancelButton;
                default:
                    return null;
            }
        }

        private StackPanel GetTokenPanel(EditableStyleField field)
        {
            switch (field)
            {
                case EditableStyleField.SuggestedGroupName:
                    return suggestedGroupNameTokenPanel;
                case EditableStyleField.NewQueryRename:
                    return newQueryRenameTokenPanel;
                case EditableStyleField.SavedFileRename:
                    return savedFileRenameTokenPanel;
                default:
                    return null;
            }
        }

        private TextBlock GetTokenHelpText(EditableStyleField field)
        {
            switch (field)
            {
                case EditableStyleField.SuggestedGroupName:
                    return suggestedGroupNameTokenHelpText;
                case EditableStyleField.NewQueryRename:
                    return newQueryRenameTokenHelpText;
                case EditableStyleField.SavedFileRename:
                    return savedFileRenameTokenHelpText;
                default:
                    return null;
            }
        }

        private void UpdateStyleEditUi()
        {
            UpdateStyleRowUi(EditableStyleField.SuggestedGroupName);
            UpdateStyleRowUi(EditableStyleField.NewQueryRename);
            UpdateStyleRowUi(EditableStyleField.SavedFileRename);
        }

        private void UpdateStyleRowUi(EditableStyleField field)
        {
            bool hasActiveEdit = activeStyleEditField != EditableStyleField.None;
            bool isActive = activeStyleEditField == field;

            TextBox textBox = GetStyleTextBox(field);
            Button editButton = GetEditButton(field);
            Button saveButton = GetSaveButton(field);
            Button cancelButton = GetCancelButton(field);
            StackPanel tokenPanel = GetTokenPanel(field);
            TextBlock tokenHelpText = GetTokenHelpText(field);

            if (textBox != null)
            {
                textBox.IsReadOnly = !isActive;
                textBox.IsHitTestVisible = isActive;
                textBox.Focusable = isActive;
                textBox.IsTabStop = isActive;
                textBox.Cursor = isActive ? Cursors.IBeam : Cursors.Arrow;
            }

            if (editButton != null)
            {
                editButton.Visibility = isActive ? Visibility.Collapsed : Visibility.Visible;
                editButton.IsEnabled = !hasActiveEdit;
            }

            if (saveButton != null)
            {
                saveButton.Visibility = isActive ? Visibility.Visible : Visibility.Collapsed;
            }

            if (cancelButton != null)
            {
                cancelButton.Visibility = isActive ? Visibility.Visible : Visibility.Collapsed;
            }

            if (tokenPanel != null)
            {
                tokenPanel.Visibility = isActive ? Visibility.Visible : Visibility.Collapsed;
            }

            if (tokenHelpText != null)
            {
                tokenHelpText.Visibility = isActive ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        private void EnterStyleEditMode(EditableStyleField field)
        {
            if (field == EditableStyleField.None)
            {
                return;
            }

            TextBox textBox = GetStyleTextBox(field);
            if (textBox == null)
            {
                return;
            }

            activeStyleEditField = field;
            styleEditSnapshot = textBox.Text;
            UpdateStyleEditUi();

            textBox.Focus();
            textBox.CaretIndex = textBox.Text?.Length ?? 0;
            statusText.Text = "Editing style. Save or Cancel before editing another style.";
        }

        private void ExitStyleEditMode(bool restoreSnapshot, bool setStatus)
        {
            if (activeStyleEditField == EditableStyleField.None)
            {
                return;
            }

            if (restoreSnapshot)
            {
                TextBox textBox = GetStyleTextBox(activeStyleEditField);
                if (textBox != null && styleEditSnapshot != null)
                {
                    textBox.Text = styleEditSnapshot;
                }
            }

            activeStyleEditField = EditableStyleField.None;
            styleEditSnapshot = null;
            UpdateStyleEditUi();

            if (setStatus)
            {
                statusText.Text = "Style edit canceled.";
            }
        }

        private void StyleEditButton_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!TryGetStyleFieldFromSender(sender, out EditableStyleField field))
            {
                return;
            }

            if (activeStyleEditField != EditableStyleField.None)
            {
                statusText.Text = "Save or Cancel the current style edit before editing another style.";
                return;
            }

            EnterStyleEditMode(field);
        }

        private void StyleSaveButton_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!TryGetStyleFieldFromSender(sender, out EditableStyleField field) || activeStyleEditField != field)
            {
                return;
            }

            styleEditSnapshot = null;
            SaveStyleTabFromUi("Style saved.");
            ExitStyleEditMode(restoreSnapshot: false, setStatus: false);
        }

        private void StyleCancelButton_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!TryGetStyleFieldFromSender(sender, out EditableStyleField field) || activeStyleEditField != field)
            {
                return;
            }

            ExitStyleEditMode(restoreSnapshot: true, setStatus: true);
        }

        private void TokenButton_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (activeStyleEditField == EditableStyleField.None)
            {
                statusText.Text = "Click Edit before inserting tokens.";
                return;
            }

            FrameworkElement tokenSource = sender as FrameworkElement;
            string token = tokenSource?.Tag as string;
            if (string.IsNullOrWhiteSpace(token))
            {
                return;
            }

            TextBox textBox = GetStyleTextBox(activeStyleEditField);
            if (textBox == null)
            {
                return;
            }

            string current = textBox.Text ?? string.Empty;
            int selectionStart = Math.Max(0, Math.Min(current.Length, textBox.SelectionStart));
            int selectionLength = Math.Max(0, Math.Min(current.Length - selectionStart, textBox.SelectionLength));

            textBox.Text = current.Remove(selectionStart, selectionLength).Insert(selectionStart, token);
            textBox.SelectionStart = selectionStart + token.Length;
            textBox.SelectionLength = 0;
            textBox.Focus();

            statusText.Text = "Token inserted.";
        }

        private void ReloadConnectionGroupsTabFromConfig()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                TabGroupConfig config = TabGroupConfigLoader.LoadOrNull() ?? new TabGroupConfig();
                LoadConnectionGroupsTabState(config);
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"SettingsToolWindowControl reload connection tab failed: {ex.Message}");
            }
        }

        private void LoadConnectionGroupsTabState(TabGroupConfig config)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            var settings = config?.Settings ?? new TabGroupSettings();
            showGroupNameInConnectionCards = settings.EnableAutoRename;
            showServerAliasSection = settings.EnableServerAliasPrompt && settings.EnableAutoRename;
            currentAutoConfigureMode = NormalizeAutoConfigure(settings.AutoConfigure);

            serverAliasSectionPanel.Visibility = showServerAliasSection ? Visibility.Visible : Visibility.Collapsed;
            UpdateConnectionGroupsDescriptionText();

            var aliasMap = config?.ServerAliases ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            serverAliasRows = aliasMap
                .OrderBy(kvp => kvp.Key ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .Select(kvp => new ServerAliasRowState
                {
                    Id = nextInlineRowId++,
                    Server = kvp.Key,
                    Alias = kvp.Value
                })
                .ToList();

            var groups = config?.ConnectionGroups ?? new List<TabGroupRule>();
            connectionGroupCards = groups
                .Select(rule => new ConnectionGroupCardState
                {
                    Id = nextInlineRowId++,
                    GroupName = rule?.GroupName,
                    Server = rule?.Server,
                    Database = rule?.Database,
                    Priority = rule?.Priority ?? 0,
                    ColorIndex = rule?.ColorIndex
                })
                .ToList();

            CancelAllConnectionTabEdits(setStatus: false);
            RebuildServerAliasRowsUi();
            RebuildConnectionGroupCardsUi();
        }

        private void CancelAllConnectionTabEdits(bool setStatus)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            bool changed = false;

            for (int i = serverAliasRows.Count - 1; i >= 0; i--)
            {
                ServerAliasRowState row = serverAliasRows[i];
                if (!row.IsEditing)
                {
                    continue;
                }

                if (row.IsNew)
                {
                    serverAliasRows.RemoveAt(i);
                }
                else
                {
                    row.Server = row.SnapshotServer;
                    row.Alias = row.SnapshotAlias;
                    row.IsEditing = false;
                    row.SnapshotServer = null;
                    row.SnapshotAlias = null;
                }

                changed = true;
            }

            for (int i = connectionGroupCards.Count - 1; i >= 0; i--)
            {
                ConnectionGroupCardState card = connectionGroupCards[i];
                if (!card.IsEditing)
                {
                    continue;
                }

                if (card.IsNew)
                {
                    connectionGroupCards.RemoveAt(i);
                }
                else
                {
                    card.GroupName = card.SnapshotGroupName;
                    card.Server = card.SnapshotServer;
                    card.Database = card.SnapshotDatabase;
                    card.Priority = card.SnapshotPriority;
                    card.ColorIndex = card.SnapshotColorIndex;
                    card.IsEditing = false;
                    card.SnapshotGroupName = null;
                    card.SnapshotServer = null;
                    card.SnapshotDatabase = null;
                }

                changed = true;
            }

            if (changed)
            {
                RebuildServerAliasRowsUi();
                RebuildConnectionGroupCardsUi();
            }

            if (setStatus)
            {
                statusText.Text = "Edit canceled.";
            }
        }

        private bool HasActiveConnectionTabEdit()
        {
            return serverAliasRows.Any(row => row.IsEditing)
                || connectionGroupCards.Any(card => card.IsEditing);
        }

        private void UpdateConnectionTabEditInteractionState()
        {
            bool hasActiveEdit = HasActiveConnectionTabEdit();
            bool allowGlobalActions = !hasActiveEdit;

            openJsonButton.IsEnabled = allowGlobalActions;
            resetDefaultsButton.IsEnabled = allowGlobalActions;
            reloadButton.IsEnabled = allowGlobalActions;
            updateEnvTabsButton.IsEnabled = allowGlobalActions;
            copyLogPathButton.IsEnabled = allowGlobalActions;
            addServerAliasButton.IsEnabled = allowGlobalActions;
            addConnectionGroupButton.IsEnabled = allowGlobalActions;

            suggestedGroupNameEditButton.IsEnabled = allowGlobalActions;
            suggestedGroupNameSaveButton.IsEnabled = allowGlobalActions;
            suggestedGroupNameCancelButton.IsEnabled = allowGlobalActions;
            newQueryRenameEditButton.IsEnabled = allowGlobalActions;
            newQueryRenameSaveButton.IsEnabled = allowGlobalActions;
            newQueryRenameCancelButton.IsEnabled = allowGlobalActions;
            savedFileRenameEditButton.IsEnabled = allowGlobalActions;
            savedFileRenameSaveButton.IsEnabled = allowGlobalActions;
            savedFileRenameCancelButton.IsEnabled = allowGlobalActions;

            suggestedGroupNameTokenPanel.IsEnabled = allowGlobalActions;
            newQueryRenameTokenPanel.IsEnabled = allowGlobalActions;
            savedFileRenameTokenPanel.IsEnabled = allowGlobalActions;
        }

        private bool CanStartConnectionTabEdit(string blockedMessage)
        {
            if (activeStyleEditField != EditableStyleField.None)
            {
                statusText.Text = "Save or Cancel the current style edit first.";
                return false;
            }

            if (HasActiveConnectionTabEdit())
            {
                statusText.Text = blockedMessage;
                return false;
            }

            return true;
        }

        private void AddServerAliasButton_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!CanStartConnectionTabEdit("Save or Cancel the current connection edit before adding another alias."))
            {
                return;
            }

            var row = new ServerAliasRowState
            {
                Id = nextInlineRowId++,
                IsEditing = true,
                IsNew = true,
                Server = string.Empty,
                Alias = string.Empty
            };

            serverAliasRows.Insert(0, row);
            RebuildServerAliasRowsUi();
            RebuildConnectionGroupCardsUi();
            statusText.Text = "Enter alias values, then click Save.";
        }

        private void AddConnectionGroupButton_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!CanStartConnectionTabEdit("Save or Cancel the current connection edit before adding another group."))
            {
                return;
            }

            var card = new ConnectionGroupCardState
            {
                Id = nextInlineRowId++,
                IsEditing = true,
                IsNew = true,
                GroupName = string.Empty,
                Server = string.Empty,
                Database = string.Empty,
                Priority = 0,
                ColorIndex = null
            };

            connectionGroupCards.Insert(0, card);
            RebuildConnectionGroupCardsUi();
            RebuildServerAliasRowsUi();
            statusText.Text = "Enter group values, then click Save.";
        }

        private void RebuildServerAliasRowsUi()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            activeAliasEditRow = null;
            activeAliasServerTextBox = null;
            activeAliasAliasTextBox = null;
            serverAliasRowsPanel.Children.Clear();
            bool hasActiveEdit = HasActiveConnectionTabEdit();

            if (!showServerAliasSection)
            {
                serverAliasColumnHeaderGrid.Visibility = Visibility.Collapsed;
                UpdateConnectionTabEditInteractionState();
                return;
            }

            if (serverAliasRows.Count == 0)
            {
                serverAliasColumnHeaderGrid.Visibility = Visibility.Collapsed;
                serverAliasRowsPanel.Children.Add(new TextBlock
                {
                    Text = "No server aliases configured.",
                    Opacity = 0.72,
                    Margin = new Thickness(16, 6, 0, 8)
                });
                UpdateConnectionTabEditInteractionState();
                return;
            }

            serverAliasColumnHeaderGrid.Visibility = Visibility.Visible;

            foreach (ServerAliasRowState row in serverAliasRows)
            {
                Border cardBorder = new Border
                {
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(4),
                    Padding = new Thickness(8, 6, 8, 6),
                    Margin = new Thickness(0, 0, 0, 6)
                };
                cardBorder.SetResourceReference(Border.BorderBrushProperty, "EnvTabsBorderBrush");
                cardBorder.SetResourceReference(Border.BackgroundProperty, "EnvTabsBackgroundBrush");

                Grid grid = new Grid();
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                StackPanel fieldsPanel = new StackPanel();
                if (row.IsEditing)
                {
                    Grid editorGrid = new Grid();
                    editorGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
                    editorGrid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

                    Grid inputsGrid = new Grid { Margin = new Thickness(0, 0, 0, 8) };
                    inputsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                    inputsGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                    StackPanel serverPanel = new StackPanel { Margin = new Thickness(0, 0, 6, 0) };
                    TextBox serverTextBox = new TextBox { Text = row.Server ?? string.Empty, MinWidth = 250, Padding = new Thickness(6, 3, 6, 3) };
                    ApplyEditorTextBoxTheme(serverTextBox);
                    serverPanel.Children.Add(serverTextBox);
                    Grid.SetColumn(serverPanel, 0);
                    inputsGrid.Children.Add(serverPanel);

                    StackPanel aliasPanel = new StackPanel { Margin = new Thickness(6, 0, 0, 0) };
                    TextBox aliasTextBox = new TextBox { Text = row.Alias ?? string.Empty, MinWidth = 250, Padding = new Thickness(6, 3, 6, 3) };
                    ApplyEditorTextBoxTheme(aliasTextBox);
                    aliasPanel.Children.Add(aliasTextBox);
                    Grid.SetColumn(aliasPanel, 1);
                    inputsGrid.Children.Add(aliasPanel);
                    editorGrid.Children.Add(inputsGrid);

                    StackPanel actionPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(0, 0, 0, 0) };

                    Button saveButton = new Button { Content = "Save", Style = ResolveButtonStyle("PrimaryCompactCardButtonStyle") };
                    saveButton.Click += (s, e) => SaveServerAliasRow(row, serverTextBox.Text, aliasTextBox.Text);
                    actionPanel.Children.Add(saveButton);

                    Button cancelButton = new Button { Content = "Cancel", Style = ResolveButtonStyle("CompactCardButtonStyle") };
                    cancelButton.Click += (s, e) => CancelServerAliasEdit(row);
                    actionPanel.Children.Add(cancelButton);

                    Grid.SetRow(actionPanel, 1);
                    editorGrid.Children.Add(actionPanel);

                    fieldsPanel.Children.Add(editorGrid);

                    Grid.SetColumn(fieldsPanel, 0);
                    Grid.SetColumnSpan(fieldsPanel, 2);
                    grid.Children.Add(fieldsPanel);

                    activeAliasEditRow = row;
                    activeAliasServerTextBox = serverTextBox;
                    activeAliasAliasTextBox = aliasTextBox;
                }
                else
                {
                    Grid summaryRowGrid = new Grid();
                    summaryRowGrid.MinHeight = 32;

                    Grid summaryGrid = new Grid();
                    summaryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                    summaryGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                    TextBlock serverValueText = new TextBlock
                    {
                        Text = string.IsNullOrWhiteSpace(row.Server) ? "(none)" : row.Server,
                        FontWeight = FontWeights.SemiBold,
                        VerticalAlignment = VerticalAlignment.Center
                    };
                    Grid.SetColumn(serverValueText, 0);
                    summaryGrid.Children.Add(serverValueText);

                    TextBlock aliasValueText = new TextBlock
                    {
                        Text = string.IsNullOrWhiteSpace(row.Alias) ? "(none)" : row.Alias,
                        FontWeight = FontWeights.SemiBold,
                        VerticalAlignment = VerticalAlignment.Center
                    };
                    Grid.SetColumn(aliasValueText, 1);
                    summaryGrid.Children.Add(aliasValueText);

                    summaryRowGrid.Children.Add(summaryGrid);

                    StackPanel actionPanel = new StackPanel
                    {
                        Orientation = Orientation.Horizontal,
                        HorizontalAlignment = HorizontalAlignment.Right,
                        VerticalAlignment = VerticalAlignment.Center,
                        Margin = new Thickness(10, 0, 0, 0)
                    };
                    if (hasActiveEdit)
                    {
                        actionPanel.Visibility = Visibility.Collapsed;
                    }
                    else
                    {
                        actionPanel.Style = TryFindResource("HoverRevealReservedSpaceActionPanelStyle") as Style;
                    }

                    Button editButton = new Button { Content = "Edit", Style = ResolveButtonStyle("CompactCardButtonStyle") };
                    editButton.IsEnabled = !hasActiveEdit;
                    editButton.Click += (s, e) => BeginServerAliasEdit(row);
                    actionPanel.Children.Add(editButton);

                    Button deleteButton = new Button { Content = "Delete", Style = ResolveButtonStyle("DangerCompactCardButtonStyle") };
                    deleteButton.IsEnabled = !hasActiveEdit;
                    deleteButton.Click += (s, e) => DeleteServerAliasRow(row);
                    actionPanel.Children.Add(deleteButton);

                    summaryRowGrid.Children.Add(actionPanel);
                    fieldsPanel.Children.Add(summaryRowGrid);

                    Grid.SetColumn(fieldsPanel, 0);
                    Grid.SetColumnSpan(fieldsPanel, 2);
                    grid.Children.Add(fieldsPanel);
                }

                cardBorder.Child = grid;
                serverAliasRowsPanel.Children.Add(cardBorder);
            }

            UpdateConnectionTabEditInteractionState();
        }

        private void BeginServerAliasEdit(ServerAliasRowState row)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!CanStartConnectionTabEdit("Save or Cancel the current connection edit before editing another alias."))
            {
                return;
            }

            row.IsEditing = true;
            row.IsNew = false;
            row.SnapshotServer = row.Server;
            row.SnapshotAlias = row.Alias;
            RebuildServerAliasRowsUi();
            RebuildConnectionGroupCardsUi();
            statusText.Text = "Editing alias. Click Save or Cancel.";
        }

        private void SaveServerAliasRow(ServerAliasRowState row, string serverValue, string aliasValue)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            string trimmedServer = (serverValue ?? string.Empty).Trim();
            string trimmedAlias = (aliasValue ?? string.Empty).Trim();

            if (string.IsNullOrWhiteSpace(trimmedServer) || string.IsNullOrWhiteSpace(trimmedAlias))
            {
                statusText.Text = "Server and Alias are required.";
                return;
            }

            bool duplicate = serverAliasRows.Any(other => other.Id != row.Id && string.Equals(other.Server, trimmedServer, StringComparison.OrdinalIgnoreCase));
            if (duplicate)
            {
                statusText.Text = "A server alias for that server already exists.";
                return;
            }

            row.Server = trimmedServer;
            row.Alias = trimmedAlias;
            row.IsEditing = false;
            row.IsNew = false;
            row.SnapshotServer = null;
            row.SnapshotAlias = null;

            SaveConnectionGroupsTabFromUi("Server alias saved.", persistAliases: true);
            RebuildServerAliasRowsUi();
            RebuildConnectionGroupCardsUi();
        }

        private void CancelServerAliasEdit(ServerAliasRowState row)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (row.IsNew)
            {
                serverAliasRows.Remove(row);
            }
            else
            {
                row.Server = row.SnapshotServer;
                row.Alias = row.SnapshotAlias;
                row.IsEditing = false;
                row.SnapshotServer = null;
                row.SnapshotAlias = null;
            }

            RebuildServerAliasRowsUi();
            RebuildConnectionGroupCardsUi();
            statusText.Text = "Alias edit canceled.";
        }

        private void DeleteServerAliasRow(ServerAliasRowState row)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            MessageBoxResult result = MessageBox.Show(
                $"Delete alias mapping for '{row.Server}'?",
                "Delete server alias",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result != MessageBoxResult.Yes)
            {
                return;
            }

            serverAliasRows.Remove(row);
            SaveConnectionGroupsTabFromUi("Server alias deleted.", persistAliases: true);
            RebuildServerAliasRowsUi();
            RebuildConnectionGroupCardsUi();
        }

        private void RebuildConnectionGroupCardsUi()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            activeConnectionGroupEditCard = null;
            activeConnectionGroupNameTextBox = null;
            activeConnectionServerTextBox = null;
            activeConnectionDatabaseTextBox = null;
            activeConnectionPriorityTextBox = null;
            activeConnectionColorCombo = null;
            connectionGroupCardsPanel.Children.Clear();
            bool hasActiveEdit = HasActiveConnectionTabEdit();
            bool showDatabaseField = ShouldShowDatabaseInCards();
            bool showGroupNameField = ShouldShowGroupNameInConnectionCards();

            if (connectionGroupCards.Count == 0)
            {
                connectionGroupCardsPanel.Children.Add(new TextBlock
                {
                    Text = "No connection groups configured.",
                    Opacity = 0.72,
                    Margin = new Thickness(16, 6, 0, 8)
                });
                UpdateConnectionTabEditInteractionState();
                return;
            }

            foreach (ConnectionGroupCardState card in connectionGroupCards
                .OrderBy(c => c.Priority)
                .ThenBy(c => c.GroupName ?? string.Empty, StringComparer.OrdinalIgnoreCase))
            {
                Brush cardSurface = TryFindResource("EnvTabsBackgroundBrush") as Brush ?? Brushes.White;
                Brush cardForeground = ResolveReadableForegroundBrush(cardSurface);
                Brush swatchFill = ResolveColorSwatchFillBrush(card.ColorIndex);
                Brush swatchBorder = ResolveColorSwatchBorderBrush(card.ColorIndex);

                Border cardBorder = new Border
                {
                    BorderThickness = new Thickness(1),
                    CornerRadius = new CornerRadius(4),
                    Padding = new Thickness(8, 6, 8, 6),
                    Margin = new Thickness(0, 0, 0, 6)
                };
                cardBorder.SetResourceReference(Border.BorderBrushProperty, "EnvTabsBorderBrush");
                cardBorder.SetResourceReference(Border.BackgroundProperty, "EnvTabsBackgroundBrush");

                Grid grid = new Grid();
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                grid.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

                Border swatch = new Border
                {
                    Width = 16,
                    Height = 16,
                    Margin = new Thickness(0, 0, 10, 0),
                    VerticalAlignment = VerticalAlignment.Center,
                    Visibility = card.IsEditing ? Visibility.Collapsed : Visibility.Visible,
                    CornerRadius = new CornerRadius(1),
                    BorderThickness = new Thickness(1),
                    Background = swatchFill,
                    BorderBrush = swatchBorder
                };
                Grid.SetColumn(swatch, 0);
                grid.Children.Add(swatch);

                StackPanel fieldsPanel = new StackPanel();
                if (card.IsEditing)
                {
                    TextBox groupTextBox = null;
                    if (showGroupNameField)
                    {
                        fieldsPanel.Children.Add(new TextBlock { Text = "Group name", FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 3), Foreground = cardForeground });
                        groupTextBox = new TextBox { Text = card.GroupName ?? string.Empty, MinWidth = 260, Margin = new Thickness(0, 0, 0, 6), Padding = new Thickness(6, 3, 6, 3) };
                        ApplyEditorTextBoxTheme(groupTextBox);
                        fieldsPanel.Children.Add(groupTextBox);
                    }

                    Grid serverDbGrid = new Grid { Margin = new Thickness(0, 0, 0, 6) };
                    serverDbGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
                    serverDbGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });

                    StackPanel serverPanel = new StackPanel { Margin = new Thickness(0, 0, 6, 0) };
                    serverPanel.Children.Add(new TextBlock { Text = "Server", FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 3), Foreground = cardForeground });
                    TextBox serverTextBox = new TextBox { Text = card.Server ?? string.Empty, Padding = new Thickness(6, 3, 6, 3) };
                    ApplyEditorTextBoxTheme(serverTextBox);
                    serverPanel.Children.Add(serverTextBox);
                    Grid.SetColumn(serverPanel, 0);
                    serverDbGrid.Children.Add(serverPanel);

                    StackPanel databasePanel = new StackPanel { Margin = new Thickness(6, 0, 0, 0) };
                    databasePanel.Children.Add(new TextBlock { Text = "Database", FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 3), Foreground = cardForeground });
                    TextBox databaseTextBox = new TextBox { Text = card.Database ?? string.Empty, Padding = new Thickness(6, 3, 6, 3) };
                    ApplyEditorTextBoxTheme(databaseTextBox);
                    databasePanel.Children.Add(databaseTextBox);
                    Grid.SetColumn(databasePanel, 1);
                    serverDbGrid.Children.Add(databasePanel);

                    fieldsPanel.Children.Add(serverDbGrid);

                    Grid rowGrid = new Grid { Margin = new Thickness(0, 0, 0, 0) };
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(140) });
                    rowGrid.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(220) });

                    StackPanel priorityPanel = new StackPanel();
                    priorityPanel.Children.Add(new TextBlock { Text = "Priority", FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 3), Foreground = cardForeground });
                    TextBox priorityTextBox = new TextBox { Text = card.Priority.ToString(), Padding = new Thickness(6, 3, 6, 3), Margin = new Thickness(0, 0, 8, 0) };
                    ApplyEditorTextBoxTheme(priorityTextBox);
                    priorityPanel.Children.Add(priorityTextBox);
                    Grid.SetColumn(priorityPanel, 0);
                    rowGrid.Children.Add(priorityPanel);

                    StackPanel colorPanel = new StackPanel();
                    colorPanel.Children.Add(new TextBlock { Text = "Color", FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 3), Foreground = cardForeground });
                    ComboBox colorCombo = new ComboBox
                    {
                        MinWidth = 200,
                        Padding = new Thickness(6, 3, 6, 3),
                        ItemsSource = BuildColorChoices(card),
                        SelectedValuePath = nameof(InlineColorChoice.Index)
                    };
                    colorCombo.ItemTemplate = CreateInlineColorChoiceTemplate();
                    ApplyComboBoxPopupTheme(colorCombo);
                    ApplyComboBoxTemplate(colorCombo);
                    colorCombo.SelectedValue = card.ColorIndex;
                    colorPanel.Children.Add(colorCombo);
                    Grid.SetColumn(colorPanel, 1);
                    rowGrid.Children.Add(colorPanel);

                    fieldsPanel.Children.Add(rowGrid);

                    StackPanel actionPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(10, 0, 0, 0) };

                    Button saveButton = new Button { Content = "Save", Style = ResolveButtonStyle("PrimaryCompactCardButtonStyle") };
                    saveButton.Click += (s, e) => SaveConnectionGroupCard(card, showGroupNameField ? groupTextBox?.Text : card.GroupName, serverTextBox.Text, databaseTextBox.Text, priorityTextBox.Text, colorCombo.SelectedValue);
                    actionPanel.Children.Add(saveButton);

                    Button cancelButton = new Button { Content = "Cancel", Style = ResolveButtonStyle("CompactCardButtonStyle") };
                    cancelButton.Click += (s, e) => CancelConnectionGroupEdit(card);
                    actionPanel.Children.Add(cancelButton);

                    Grid.SetColumn(fieldsPanel, 1);
                    Grid.SetColumn(actionPanel, 2);
                    grid.Children.Add(fieldsPanel);
                    grid.Children.Add(actionPanel);

                    activeConnectionGroupEditCard = card;
                    activeConnectionGroupNameTextBox = groupTextBox;
                    activeConnectionServerTextBox = serverTextBox;
                    activeConnectionDatabaseTextBox = databaseTextBox;
                    activeConnectionPriorityTextBox = priorityTextBox;
                    activeConnectionColorCombo = colorCombo;
                }
                else
                {
                    string groupNameText = string.IsNullOrWhiteSpace(card.GroupName) ? "(unnamed)" : card.GroupName;
                    string colorText = GetColorDisplayName(card.ColorIndex);
                    string headingText = showGroupNameField ? (groupNameText + " | " + colorText) : colorText;
                    fieldsPanel.Children.Add(new TextBlock
                    {
                        Text = headingText,
                        FontWeight = FontWeights.Bold,
                        FontSize = 13,
                        Foreground = cardForeground
                    });

                    string serverText = string.IsNullOrWhiteSpace(card.Server) ? "(any server)" : card.Server;
                    string databaseText = string.IsNullOrWhiteSpace(card.Database) ? "(any database)" : card.Database;
                    string serverDbSummary = showDatabaseField
                        ? (serverText + ", " + databaseText)
                        : serverText;
                    fieldsPanel.Children.Add(new TextBlock
                    {
                        Text = serverDbSummary,
                        Margin = new Thickness(0, 3, 0, 0),
                        Opacity = 0.82,
                        FontSize = 12,
                        Foreground = cardForeground
                    });

                    StackPanel actionPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right, Margin = new Thickness(10, 0, 0, 0) };
                    if (hasActiveEdit)
                    {
                        actionPanel.Visibility = Visibility.Collapsed;
                    }
                    else
                    {
                        actionPanel.Style = TryFindResource("HoverRevealActionPanelStyle") as Style;
                    }

                    Button editButton = new Button { Content = "Edit", Style = ResolveButtonStyle("CompactCardButtonStyle") };
                    editButton.IsEnabled = !hasActiveEdit;
                    editButton.Click += (s, e) => BeginConnectionGroupEdit(card);
                    actionPanel.Children.Add(editButton);

                    Button deleteButton = new Button { Content = "Delete", Style = ResolveButtonStyle("DangerCompactCardButtonStyle") };
                    deleteButton.IsEnabled = !hasActiveEdit;
                    deleteButton.Click += (s, e) => DeleteConnectionGroupCard(card);
                    actionPanel.Children.Add(deleteButton);

                    Grid.SetColumn(fieldsPanel, 1);
                    Grid.SetColumn(actionPanel, 2);
                    grid.Children.Add(fieldsPanel);
                    grid.Children.Add(actionPanel);
                }

                cardBorder.Child = grid;
                connectionGroupCardsPanel.Children.Add(cardBorder);
            }

            UpdateConnectionTabEditInteractionState();
        }

        private void BeginConnectionGroupEdit(ConnectionGroupCardState card)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (!CanStartConnectionTabEdit("Save or Cancel the current connection edit before editing another group."))
            {
                return;
            }

            card.IsEditing = true;
            card.IsNew = false;
            card.SnapshotGroupName = card.GroupName;
            card.SnapshotServer = card.Server;
            card.SnapshotDatabase = card.Database;
            card.SnapshotPriority = card.Priority;
            card.SnapshotColorIndex = card.ColorIndex;
            RebuildConnectionGroupCardsUi();
            RebuildServerAliasRowsUi();
            statusText.Text = "Editing group. Click Save or Cancel.";
        }

        private void SaveConnectionGroupCard(ConnectionGroupCardState card, string groupName, string server, string database, string priorityText, object selectedColorValue)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            string trimmedGroupName = (groupName ?? string.Empty).Trim();
            string trimmedServer = (server ?? string.Empty).Trim();
            string trimmedDatabase = (database ?? string.Empty).Trim();
            bool showGroupNameField = ShouldShowGroupNameInConnectionCards();

            if (showGroupNameField && string.IsNullOrWhiteSpace(trimmedGroupName))
            {
                statusText.Text = "Group name is required.";
                return;
            }

            if (string.IsNullOrWhiteSpace(trimmedServer) && string.IsNullOrWhiteSpace(trimmedDatabase))
            {
                statusText.Text = "Provide at least Server or Database.";
                return;
            }

            if (!int.TryParse((priorityText ?? string.Empty).Trim(), out int parsedPriority))
            {
                statusText.Text = "Priority must be a whole number.";
                return;
            }

            int? selectedColorIndex = null;
            if (selectedColorValue is int)
            {
                selectedColorIndex = (int)selectedColorValue;
            }
            else if (selectedColorValue is int?)
            {
                selectedColorIndex = (int?)selectedColorValue;
            }

            if (selectedColorIndex.HasValue && (selectedColorIndex.Value < 0 || selectedColorIndex.Value > 15))
            {
                statusText.Text = "Color index must be between 0 and 15.";
                return;
            }

            if (!showGroupNameField)
            {
                if (string.IsNullOrWhiteSpace(trimmedGroupName))
                {
                    trimmedGroupName = !string.IsNullOrWhiteSpace(card.GroupName)
                        ? card.GroupName
                        : BuildFallbackGroupName(trimmedServer, trimmedDatabase);
                }
            }

            card.GroupName = trimmedGroupName;
            card.Server = string.IsNullOrWhiteSpace(trimmedServer) ? null : trimmedServer;
            card.Database = string.IsNullOrWhiteSpace(trimmedDatabase) ? null : trimmedDatabase;
            card.Priority = parsedPriority;
            card.ColorIndex = selectedColorIndex;
            card.IsEditing = false;
            card.IsNew = false;
            card.SnapshotGroupName = null;
            card.SnapshotServer = null;
            card.SnapshotDatabase = null;

            SaveConnectionGroupsTabFromUi("Connection group saved.", persistAliases: showServerAliasSection);
            RebuildConnectionGroupCardsUi();
            RebuildServerAliasRowsUi();
        }

        private void CancelConnectionGroupEdit(ConnectionGroupCardState card)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (card.IsNew)
            {
                connectionGroupCards.Remove(card);
            }
            else
            {
                card.GroupName = card.SnapshotGroupName;
                card.Server = card.SnapshotServer;
                card.Database = card.SnapshotDatabase;
                card.Priority = card.SnapshotPriority;
                card.ColorIndex = card.SnapshotColorIndex;
                card.IsEditing = false;
                card.SnapshotGroupName = null;
                card.SnapshotServer = null;
                card.SnapshotDatabase = null;
            }

            RebuildConnectionGroupCardsUi();
            RebuildServerAliasRowsUi();
            statusText.Text = "Group edit canceled.";
        }

        private void DeleteConnectionGroupCard(ConnectionGroupCardState card)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            MessageBoxResult result = MessageBox.Show(
                $"Delete connection group '{card.GroupName}'?",
                "Delete connection group",
                MessageBoxButton.YesNo,
                MessageBoxImage.Warning);

            if (result != MessageBoxResult.Yes)
            {
                return;
            }

            connectionGroupCards.Remove(card);
            SaveConnectionGroupsTabFromUi("Connection group deleted.", persistAliases: showServerAliasSection);
            RebuildConnectionGroupCardsUi();
            RebuildServerAliasRowsUi();
        }

        private List<InlineColorChoice> BuildColorChoices(ConnectionGroupCardState card)
        {
            var usedIndexes = new HashSet<int>(
                connectionGroupCards
                    .Where(other => other.Id != card.Id && other.ColorIndex.HasValue)
                    .Select(other => other.ColorIndex.Value));

            var choices = new List<InlineColorChoice>
            {
                new InlineColorChoice
                {
                    Index = null,
                    DisplayName = "None",
                    SwatchBrush = Brushes.Transparent,
                    SwatchBorderBrush = TryFindResource("EnvTabsBorderBrush") as Brush ?? Brushes.Gray
                }
            };

            foreach (ColorPaletteItem color in ColorPalette)
            {
                bool isUsed = usedIndexes.Contains(color.Index);
                string displayName = isUsed ? $"{color.Name} (used)" : color.Name;
                choices.Add(new InlineColorChoice
                {
                    Index = color.Index,
                    DisplayName = displayName,
                    SwatchBrush = new SolidColorBrush(color.Color),
                    SwatchBorderBrush = Brushes.Transparent
                });
            }

            return choices;
        }

        private static DataTemplate CreateInlineColorChoiceTemplate()
        {
            var root = new FrameworkElementFactory(typeof(StackPanel));
            root.SetValue(StackPanel.OrientationProperty, Orientation.Horizontal);

            var swatch = new FrameworkElementFactory(typeof(Border));
            swatch.SetValue(Border.WidthProperty, 14.0);
            swatch.SetValue(Border.HeightProperty, 14.0);
            swatch.SetValue(Border.MarginProperty, new Thickness(0, 0, 8, 0));
            swatch.SetBinding(Border.BackgroundProperty, new Binding(nameof(InlineColorChoice.SwatchBrush)));
            swatch.SetBinding(Border.BorderBrushProperty, new Binding(nameof(InlineColorChoice.SwatchBorderBrush)));
            swatch.SetValue(Border.BorderThicknessProperty, new Thickness(1));
            swatch.SetValue(Border.CornerRadiusProperty, new CornerRadius(2));

            var text = new FrameworkElementFactory(typeof(TextBlock));
            text.SetBinding(TextBlock.TextProperty, new Binding(nameof(InlineColorChoice.DisplayName)));
            text.SetValue(TextBlock.VerticalAlignmentProperty, VerticalAlignment.Center);

            root.AppendChild(swatch);
            root.AppendChild(text);

            return new DataTemplate { VisualTree = root };
        }

        private void ApplyEditorTextBoxTheme(TextBox textBox)
        {
            if (textBox == null)
            {
                return;
            }

            textBox.SetResourceReference(Control.BackgroundProperty, "EnvTabsBackgroundBrush");
            textBox.SetResourceReference(Control.ForegroundProperty, "EnvTabsForegroundBrush");
            textBox.SetResourceReference(Control.BorderBrushProperty, "EnvTabsBorderBrush");
        }

        private Style ResolveButtonStyle(string preferredKey)
        {
            Style style = TryFindResource(preferredKey) as Style;
            if (style != null)
            {
                return style;
            }

            return TryFindResource("StyleEditButtonStyle") as Style;
        }

        private string GetColorDisplayName(int? colorIndex)
        {
            if (!colorIndex.HasValue)
            {
                return "None";
            }

            foreach (ColorPaletteItem color in ColorPalette)
            {
                if (color.Index == colorIndex.Value)
                {
                    return color.Name;
                }
            }

            return colorIndex.Value.ToString();
        }

        private bool ShouldShowDatabaseInCards()
        {
            return !string.Equals(currentAutoConfigureMode, "server", StringComparison.OrdinalIgnoreCase);
        }

        private bool ShouldShowGroupNameInConnectionCards()
        {
            return showGroupNameInConnectionCards;
        }

        private void UpdateConnectionGroupsDescriptionText()
        {
            if (connectionGroupsDescriptionText == null)
            {
                return;
            }

            connectionGroupsDescriptionText.Text = showGroupNameInConnectionCards
                ? "Each record maps server and optional database patterns to a group name and tab color."
                : "Each record maps server and optional database patterns to a tab color.";
        }

        private static string BuildFallbackGroupName(string server, string database)
        {
            if (!string.IsNullOrWhiteSpace(server) && !string.IsNullOrWhiteSpace(database))
            {
                return server + " " + database;
            }

            if (!string.IsNullOrWhiteSpace(server))
            {
                return server;
            }

            if (!string.IsNullOrWhiteSpace(database))
            {
                return database;
            }

            return "Group";
        }

        private Brush ResolveColorSwatchFillBrush(int? colorIndex)
        {
            if (!colorIndex.HasValue)
            {
                return Brushes.Transparent;
            }

            foreach (ColorPaletteItem color in ColorPalette)
            {
                if (color.Index == colorIndex.Value)
                {
                    return new SolidColorBrush(color.Color);
                }
            }

            return Brushes.Transparent;
        }

        private Brush ResolveColorSwatchBorderBrush(int? colorIndex)
        {
            if (colorIndex.HasValue)
            {
                return Brushes.Transparent;
            }

            return TryFindResource("EnvTabsBorderBrush") as Brush ?? Brushes.Gray;
        }

        private static Brush ResolveReadableForegroundBrush(Brush background)
        {
            Color backgroundColor = GetBrushColor(background, Colors.White);
            return GetRelativeLuminance(backgroundColor) > 0.58
                ? Brushes.Black
                : Brushes.White;
        }

        private void SaveConnectionGroupsTabFromUi(string statusMessage, bool persistAliases)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                TabGroupConfigLoader.EnsureDefaultConfigExists();
                TabGroupConfig config = TabGroupConfigLoader.LoadOrNull() ?? new TabGroupConfig();

                if (persistAliases)
                {
                    var serverAliasMap = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    foreach (ServerAliasRowState row in serverAliasRows)
                    {
                        if (string.IsNullOrWhiteSpace(row.Server) || string.IsNullOrWhiteSpace(row.Alias))
                        {
                            continue;
                        }

                        serverAliasMap[row.Server.Trim()] = row.Alias.Trim();
                    }

                    config.ServerAliases = serverAliasMap;
                }

                config.ConnectionGroups = connectionGroupCards
                    .Select(card => new TabGroupRule
                    {
                        GroupName = card.GroupName,
                        Server = card.Server,
                        Database = card.Database,
                        Priority = card.Priority,
                        ColorIndex = card.ColorIndex
                    })
                    .ToList();

                TabGroupConfigLoader.SaveConfig(config);
                statusText.Text = statusMessage;
            }
            catch (Exception ex)
            {
                statusText.Text = "Failed to save connection groups.";
                EnvTabsLog.Info($"SettingsToolWindowControl connection tab save failed: {ex.Message}");
            }
        }

        private void ResetConnectionGroupsTabToDefaults()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            TabGroupConfig defaultConfig = LoadDefaultConfigOrNull() ?? new TabGroupConfig();

            if (showServerAliasSection)
            {
                var defaultAliases = defaultConfig.ServerAliases ?? new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                serverAliasRows = defaultAliases
                    .OrderBy(kvp => kvp.Key ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                    .Select(kvp => new ServerAliasRowState
                    {
                        Id = nextInlineRowId++,
                        Server = kvp.Key,
                        Alias = kvp.Value
                    })
                    .ToList();
            }

            var defaultGroups = defaultConfig.ConnectionGroups ?? new List<TabGroupRule>();
            connectionGroupCards = defaultGroups
                .Select(rule => new ConnectionGroupCardState
                {
                    Id = nextInlineRowId++,
                    GroupName = rule?.GroupName,
                    Server = rule?.Server,
                    Database = rule?.Database,
                    Priority = rule?.Priority ?? 0,
                    ColorIndex = rule?.ColorIndex
                })
                .ToList();

            RebuildServerAliasRowsUi();
            RebuildConnectionGroupCardsUi();
            SaveConnectionGroupsTabFromUi("Connection groups reset to defaults.", persistAliases: showServerAliasSection);
        }

        private TabGroupConfig LoadDefaultConfigOrNull()
        {
            try
            {
                Assembly assembly = typeof(SettingsToolWindowControl).Assembly;
                using (Stream stream = assembly.GetManifestResourceStream(DefaultConfigResourceName))
                {
                    if (stream == null)
                    {
                        return null;
                    }

                    using (var reader = new StreamReader(stream, Encoding.UTF8, detectEncodingFromByteOrderMarks: true))
                    {
                        string jsonText = reader.ReadToEnd();
                        if (string.IsNullOrWhiteSpace(jsonText))
                        {
                            return null;
                        }

                        byte[] utf8 = Encoding.UTF8.GetBytes(jsonText.TrimStart('\uFEFF', '\u200B', '\u0000', ' ', '\t', '\r', '\n'));
                        using (var memoryStream = new MemoryStream(utf8))
                        {
                            var serializer = new DataContractJsonSerializer(typeof(TabGroupConfig), new DataContractJsonSerializerSettings
                            {
                                UseSimpleDictionaryFormat = true
                            });

                            return serializer.ReadObject(memoryStream) as TabGroupConfig;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"SettingsToolWindowControl failed to load defaults: {ex.Message}");
                return null;
            }
        }

        private void OpenJsonButton_Click(object sender, RoutedEventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            try
            {
                TabGroupConfigLoader.EnsureDefaultConfigExists();
                string path = TabGroupConfigLoader.GetUserConfigPath();
                VsShellUtilities.OpenDocument(ServiceProvider.GlobalProvider, path);
            }
            catch (Exception ex)
            {
                statusText.Text = "Failed to open JSON config.";
                EnvTabsLog.Info($"SettingsToolWindowControl Open JSON failed: {ex.Message}");
            }
        }

        internal void NavigateTo(int targetTab, bool highlightUpdateChecks, bool forceReload)
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (forceReload)
            {
                LoadFromConfig();
            }

            switch (targetTab)
            {
                case OpenConfigCommand.TargetTabConnectionGroups:
                    settingsTabControl.SelectedItem = connectionGroupsTab;
                    break;
                case OpenConfigCommand.TargetTabStyleTemplates:
                    settingsTabControl.SelectedItem = styleTemplatesTab;
                    break;
                default:
                    settingsTabControl.SelectedItem = generalSettingsTab;
                    break;
            }

            if (highlightUpdateChecks)
            {
                HighlightCheckUpdatesRow();
            }
        }

        private void HighlightCheckUpdatesRow()
        {
            ThreadHelper.ThrowIfNotOnUIThread();

            if (checkUpdatesRowGrid == null)
            {
                return;
            }

            Color baseColor = GetBrushColor(
                TryFindResource("EnvTabsTabHeaderSelectedBrush") as Brush,
                Colors.Gold);

            var highlightBrush = new SolidColorBrush(Color.FromArgb(130, baseColor.R, baseColor.G, baseColor.B));
            checkUpdatesRowGrid.Background = highlightBrush;

            var fade = new DoubleAnimation
            {
                From = 1.0,
                To = 0.0,
                BeginTime = TimeSpan.FromMilliseconds(700),
                Duration = TimeSpan.FromMilliseconds(1300)
            };

            fade.Completed += (s, e) =>
            {
                checkUpdatesRowGrid.ClearValue(Panel.BackgroundProperty);
            };

            highlightBrush.BeginAnimation(SolidColorBrush.OpacityProperty, fade);
        }
    }
}