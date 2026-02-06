using System;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.PlatformUI;
using System.Runtime.InteropServices;
using System.Drawing.Drawing2D;

namespace SSMS_EnvTabs
{
    public class NewRuleDialog : Form
    {
        public string RuleName { get; private set; }
        public int SelectedColorIndex { get; private set; }
        public bool OpenConfigRequested { get; private set; }
        
        // New Alias Property
        public string ServerAlias { get; private set; }

        private TextBox txtName;
        private ComboBox cmbColor;
        private Button btnSave;
        private Button btnCancel;
        private Button btnOpenConfig;
        
        // Alias Step Controls
        private Panel panelAlias;
        private Panel panelRule;
        private TextBox txtAlias;
        private Button btnNext;
        private Button btnCancelAlias;

        private Panel panelRuleButtons;
        private Panel panelAliasButtons;
        
        private Label lblHeader;
        private Label lblServerLabel;
        private Label lblServerValue;
        private Label lblDatabaseLabel;
        private Label lblDatabaseValue;
        private Label lblName;
        private Label lblColor;
        private Label lblAliasHeader;
        private Font boldFont;

        private Color primaryButtonBack = Color.Empty;
        private Color primaryButtonHoverBack = Color.Empty;
        private Color primaryButtonFore = Color.Empty;
        private Color secondaryButtonBack = Color.Empty;
        private Color secondaryButtonHoverBack = Color.Empty;
        private Color secondaryButtonFore = Color.Empty;
        private Color focusButtonBorder = Color.Empty;
        private readonly HashSet<Button> hoveredButtons = new HashSet<Button>();

        private Color comboBackColor = Color.Empty;
        private Color comboTextColor = Color.Empty;
        private Color comboHighlightBack = Color.Empty;
        private Color comboHighlightFore = Color.Empty;
        private Color comboArrowColor = Color.Empty;
        private Color comboBorderColor = Color.Empty;

        private readonly string serverName;
        private readonly string databaseName;
        private readonly string existingAlias;
        private readonly string suggestedName;
        private readonly string suggestedGroupNameStyle;
        private readonly bool hideAliasStep;
        private readonly bool hideGroupNameRow;
        private readonly HashSet<int> usedColorIndexes;

        private class ColorItem
        {
            public int Index { get; set; }
            public string Name { get; set; }
            public Color Color { get; set; }

            public override string ToString() => Name;
        }

        private static readonly List<ColorItem> ColorList = new List<ColorItem>
        {
            new ColorItem { Index = 0, Name = "Lavender", Color = ColorTranslator.FromHtml("#9083ef") },
            new ColorItem { Index = 1, Name = "Gold", Color = ColorTranslator.FromHtml("#d0b132") },
            new ColorItem { Index = 2, Name = "Cyan", Color = ColorTranslator.FromHtml("#30b1cd") },
            new ColorItem { Index = 3, Name = "Burgundy", Color = ColorTranslator.FromHtml("#cf6468") },
            new ColorItem { Index = 4, Name = "Green", Color = ColorTranslator.FromHtml("#6ba12a") },
            new ColorItem { Index = 5, Name = "Brown", Color = ColorTranslator.FromHtml("#bc8f6f") },
            new ColorItem { Index = 6, Name = "Royal Blue", Color = ColorTranslator.FromHtml("#5bb2fa") },
            new ColorItem { Index = 7, Name = "Pumpkin", Color = ColorTranslator.FromHtml("#d67441") },
            new ColorItem { Index = 8, Name = "Gray", Color = ColorTranslator.FromHtml("#bdbcbc") },
            new ColorItem { Index = 9, Name = "Volt", Color = ColorTranslator.FromHtml("#cbcc38") },
            new ColorItem { Index = 10, Name = "Teal", Color = ColorTranslator.FromHtml("#2aa0a4") },
            new ColorItem { Index = 11, Name = "Magenta", Color = ColorTranslator.FromHtml("#d957a7") },
            new ColorItem { Index = 12, Name = "Mint", Color = ColorTranslator.FromHtml("#6bc6a5") },
            new ColorItem { Index = 13, Name = "Dark Brown", Color = ColorTranslator.FromHtml("#946a5b") },
            new ColorItem { Index = 14, Name = "Blue", Color = ColorTranslator.FromHtml("#6a8ec6") },
            new ColorItem { Index = 15, Name = "Pink", Color = ColorTranslator.FromHtml("#e0a3a5") },
        };

        public class NewRuleDialogOptions
        {
            public string Server { get; set; }
            public string Database { get; set; }
            public string SuggestedName { get; set; }
            public string SuggestedGroupNameStyle { get; set; }
            public int SuggestedColorIndex { get; set; }
            public string ExistingAlias { get; set; }
            public bool HideDatabaseRow { get; set; }
            public bool HideAliasStep { get; set; }
            public bool HideGroupNameRow { get; set; }
            public IEnumerable<int> UsedColorIndexes { get; set; }
        }

        public NewRuleDialog(NewRuleDialogOptions options)
        {
            if (options == null) throw new ArgumentNullException(nameof(options));

            this.serverName = options.Server;
            this.databaseName = options.Database;
            this.existingAlias = options.ExistingAlias;
            this.suggestedName = options.SuggestedName;
            this.suggestedGroupNameStyle = options.SuggestedGroupNameStyle;
            this.hideAliasStep = options.HideAliasStep;
            this.hideGroupNameRow = options.HideGroupNameRow;
            this.usedColorIndexes = new HashSet<int>(options.UsedColorIndexes ?? Enumerable.Empty<int>());

            InitializeComponent(options.Server, options.Database, options.SuggestedColorIndex, options.HideDatabaseRow);
            ApplyModernStyling();
            
            // If we have an existing alias, use it; otherwise, default alias is the server name.
            this.ServerAlias = options.ExistingAlias ?? options.Server;

            // Start at the correct step
            if (!hideAliasStep && string.IsNullOrWhiteSpace(options.ExistingAlias))
            {
                ShowAliasStep();
                this.txtAlias.Text = options.Server; // Pre-fill with server name as default alias
            }
            else
            {
                ShowRuleStep();
            }

            ApplyVsTheme();
            
            if (txtName != null)
            {
                txtName.Text = options.SuggestedName;
            }
            RuleName = options.SuggestedName;
            SelectedColorIndex = options.SuggestedColorIndex;
        }

        protected override bool ShowKeyboardCues => true;

        protected override void OnHandleCreated(EventArgs e)
        {
            base.OnHandleCreated(e);
            TryEnableRoundedCorners();
            ShowKeyboardCuesAlways();
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            ShowKeyboardCuesAlways();
        }

        private void ShowAliasStep()
        {
            this.Text = "SSMS EnvTabs - Assign Server Alias";
            panelAlias.Visible = true;
            panelRule.Visible = false;
            this.AcceptButton = btnNext;
            this.CancelButton = btnCancelAlias;
            txtAlias.Select();
            ShowKeyboardCuesAlways();
        }

        private void ShowRuleStep()
        {
            this.Text = "SSMS EnvTabs - New Rule";
            panelAlias.Visible = false;
            panelRule.Visible = true;
            
            if (!hideGroupNameRow && txtName != null)
            {
                string expectedWithServer = BuildSuggestedGroupName(serverName, databaseName);
                string expectedWithExistingAlias = BuildSuggestedGroupName(string.IsNullOrWhiteSpace(existingAlias) ? serverName : existingAlias, databaseName);
                string expectedWithAlias = BuildSuggestedGroupName(string.IsNullOrWhiteSpace(ServerAlias) ? serverName : ServerAlias, databaseName);

                bool matchesSuggested = string.IsNullOrWhiteSpace(txtName.Text)
                    || string.Equals(txtName.Text, expectedWithServer, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(txtName.Text, expectedWithExistingAlias, StringComparison.OrdinalIgnoreCase)
                    || string.Equals(txtName.Text, suggestedName, StringComparison.OrdinalIgnoreCase);

                if (matchesSuggested)
                {
                    txtName.Text = expectedWithAlias;
                }
            }

            this.AcceptButton = btnSave;
            this.CancelButton = btnCancel;
            if (txtName != null)
            {
                txtName.Select();
            }
            else
            {
                cmbColor?.Select();
            }
            ShowKeyboardCuesAlways();
        }

        private string BuildSuggestedGroupName(string serverValue, string databaseValue)
        {
            string serverPart = serverValue ?? string.Empty;
            string dbPart = databaseValue ?? string.Empty;

            if (!string.IsNullOrWhiteSpace(suggestedGroupNameStyle))
            {
                return suggestedGroupNameStyle
                    .Replace("[server]", serverPart)
                    .Replace("[db]", dbPart);
            }

            if (string.IsNullOrWhiteSpace(dbPart))
            {
                return serverPart;
            }

            return $"{serverPart} {dbPart}";
        }

        private void InitializeComponent(string server, string database, int suggestedColorIndex, bool hideDatabaseRow)
        {
            Font baseFont = SystemFonts.MessageBoxFont;
            Font scaledFont = new Font(baseFont.FontFamily, baseFont.Size + 1f);
            boldFont = new Font(scaledFont, FontStyle.Bold);

            int yShift = hideDatabaseRow ? 32 : 0;
            int nameRowOffset = hideGroupNameRow ? 0 : 12;
            int inputLeft = 124;
            this.Size = new Size(450, 300 - yShift + nameRowOffset);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Font = scaledFont;
            this.AutoScaleMode = AutoScaleMode.Dpi;
            this.Padding = new Padding(0);

            // --- Panel 2: Rule (Existing) ---
            panelRule = new Panel { Dock = DockStyle.Fill, Visible = false };
            this.Controls.Add(panelRule);

            panelRuleButtons = new Panel { Dock = DockStyle.Bottom, Height = 48, Margin = Padding.Empty, Padding = Padding.Empty };
            panelRule.Controls.Add(panelRuleButtons);

            lblHeader = new Label
            {
                Text = "New connection detected",
                Location = new Point(16, 14),
                AutoSize = true
            };
            panelRule.Controls.Add(lblHeader);

            lblServerLabel = new Label { Text = "Server:", Location = new Point(16, 48), AutoSize = true };
            panelRule.Controls.Add(lblServerLabel);

            lblServerValue = new Label
            {
                Text = server,
                Location = new Point(inputLeft, 48),
                AutoSize = true,
                Font = boldFont
            };
            panelRule.Controls.Add(lblServerValue);

            if (!hideDatabaseRow)
            {
                lblDatabaseLabel = new Label { Text = "Database:", Location = new Point(16, 70), AutoSize = true };
                panelRule.Controls.Add(lblDatabaseLabel);

                lblDatabaseValue = new Label
                {
                    Text = database ?? "(any)",
                    Location = new Point(inputLeft, 70),
                    AutoSize = true,
                    Font = boldFont
                };
                panelRule.Controls.Add(lblDatabaseValue);
            }

            // Layout adjustment when the Group Name row is visible

            if (!hideGroupNameRow)
            {
                lblName = new Label { Text = "Group Name:", Location = new Point(16, 108 - yShift), AutoSize = true };
                panelRule.Controls.Add(lblName);

                txtName = new TextBox { Location = new Point(inputLeft, 110 - yShift), Size = new Size(270, 26) };
                panelRule.Controls.Add(txtName);
            }

            lblColor = new Label { Text = "Color:", Location = new Point(16, 146 - yShift + nameRowOffset), AutoSize = true };
            panelRule.Controls.Add(lblColor);

            cmbColor = new ModernComboBox 
            { 
                Location = new Point(inputLeft, 148 - yShift + nameRowOffset), 
                Size = new Size(270, 26),
                DropDownStyle = ComboBoxStyle.DropDownList,
                DrawMode = DrawMode.OwnerDrawFixed,
                FlatStyle = FlatStyle.Flat,
                ItemHeight = 22
            };
            panelRule.Controls.Add(cmbColor);

            // Prepare ordered list with suggested color FIRST
            var orderedList = new List<ColorItem>();
            var suggestedItem = ColorList.FirstOrDefault(c => c.Index == suggestedColorIndex);
            
            if (suggestedItem != null)
            {
                orderedList.Add(suggestedItem);
                orderedList.AddRange(ColorList.Where(c => c.Index != suggestedColorIndex));
            }
            else
            {
                orderedList.AddRange(ColorList);
            }

            cmbColor.DrawItem += CmbColor_DrawItem;
            cmbColor.DataSource = orderedList;
            
            btnSave = new Button { Text = "&Save", Location = new Point(45, 10), Size = new Size(110, 26), DialogResult = DialogResult.OK };
            btnSave.Click += (s, e) => {
                if (txtName != null)
                {
                    RuleName = txtName.Text;
                }
                else
                {
                    RuleName = suggestedName;
                }
                SelectedColorIndex = ((ColorItem)cmbColor.SelectedItem).Index;
            };
            panelRuleButtons.Controls.Add(btnSave);

            btnCancel = new Button { Text = "&Cancel", Location = new Point(167, 10), Size = new Size(110, 26), DialogResult = DialogResult.Cancel };
            panelRuleButtons.Controls.Add(btnCancel);

            btnOpenConfig = new Button { Text = "&Open Config", Location = new Point(289, 10), Size = new Size(120, 26), DialogResult = DialogResult.Yes };
            btnOpenConfig.Click += (s, e) => { 
                RuleName = txtName != null ? txtName.Text : suggestedName; 
                SelectedColorIndex = ((ColorItem)cmbColor.SelectedItem).Index; 
                OpenConfigRequested = true; 
            };
            panelRuleButtons.Controls.Add(btnOpenConfig);


            // --- Panel 1: Alias (New) ---
            panelAlias = new Panel { Dock = DockStyle.Fill, Visible = false };
            this.Controls.Add(panelAlias);

            panelAliasButtons = new Panel { Dock = DockStyle.Bottom, Height = 48, Margin = Padding.Empty, Padding = Padding.Empty };
            panelAlias.Controls.Add(panelAliasButtons);

            lblAliasHeader = new Label
            {
                Text = "Assign an alias for this server",
                Location = new Point(16, 14),
                AutoSize = true,
                Font = boldFont
            };
            panelAlias.Controls.Add(lblAliasHeader);

            var lblAliasServerLabel = new Label { Text = "Server:", Location = new Point(16, 50), AutoSize = true };
            panelAlias.Controls.Add(lblAliasServerLabel);

            var lblAliasServerValue = new Label
            {
                Text = server,
                Location = new Point(inputLeft, 50),
                AutoSize = true,
                Font = boldFont
            };
            panelAlias.Controls.Add(lblAliasServerValue);

            var lblAliasPrompt = new Label { Text = "Alias:", Location = new Point(16, 90), AutoSize = true };
            panelAlias.Controls.Add(lblAliasPrompt);

            txtAlias = new TextBox { Location = new Point(inputLeft, 86), Size = new Size(270, 26) };
            panelAlias.Controls.Add(txtAlias);

            btnNext = new Button { Text = "&Next >", Location = new Point(109, 10), Size = new Size(110, 26) };
            btnNext.Click += (s, e) => {
                if (string.IsNullOrWhiteSpace(txtAlias.Text))
                {
                    ServerAlias = serverName;
                }
                else
                {
                    ServerAlias = txtAlias.Text.Trim();
                }
                ShowRuleStep();
            };
            panelAliasButtons.Controls.Add(btnNext);

            btnCancelAlias = new Button { Text = "&Cancel", Location = new Point(231, 10), Size = new Size(110, 26) };
            // If they cancel the alias dialog, we proceed to rule creation BUT alias = ServerName.
            btnCancelAlias.Click += (s, e) => {
                ServerAlias = serverName; // Default to server name
                ShowRuleStep();
            };
            panelAliasButtons.Controls.Add(btnCancelAlias);
            

            this.FormClosing += (s, e) =>
            {
                if (this.DialogResult == DialogResult.OK || this.DialogResult == DialogResult.Yes)
                {
                    if (txtName != null)
                    {
                        RuleName = txtName.Text;
                    }
                    else
                    {
                        RuleName = suggestedName;
                    }
                    if (cmbColor.SelectedItem is ColorItem item)
                    {
                        SelectedColorIndex = item.Index;
                    }
                }
            };
        }

        private void ApplyModernStyling()
        {
            // Emphasize headers similar to SSMS dialogs
            if (lblHeader != null)
            {
                lblHeader.Font = new Font(Font.FontFamily, Font.Size + 3.0f, FontStyle.Bold);
            }

            if (lblAliasHeader != null)
            {
                lblAliasHeader.Font = new Font(Font.FontFamily, Font.Size + 3.0f, FontStyle.Bold);
            }

            // Soften layout with padding and consistent control sizing
            panelRule.Padding = new Padding(0, 6, 0, 0);
            panelAlias.Padding = new Padding(0, 6, 0, 0);

            ConfigureButton(btnSave, isPrimary: true);
            ConfigureButton(btnNext, isPrimary: true);
            ConfigureButton(btnCancel, isPrimary: false);
            ConfigureButton(btnOpenConfig, isPrimary: false);
            ConfigureButton(btnCancelAlias, isPrimary: false);

            if (txtName != null)
            {
                txtName.BorderStyle = BorderStyle.FixedSingle;
            }

            if (txtAlias != null)
            {
                txtAlias.BorderStyle = BorderStyle.FixedSingle;
            }

            if (cmbColor != null)
            {
                cmbColor.FlatStyle = FlatStyle.Flat;
            }
        }

        private void ApplyVsTheme()
        {
            try
            {
                // Attempt to get VS colors. 
                // Note: VSColorTheme.GetThemedColor returns System.Drawing.Color in VSSDK.
                Color bgColor = ColorTranslator.FromHtml("#2c2c2c");
                Color fgColor = ColorTranslator.FromHtml("#f0f0f0");
                Color txtBg = ColorTranslator.FromHtml("#333333");
                Color txtFg = fgColor;
                Color accentBg = ColorTranslator.FromHtml("#9184ee");
                Color accentFg = ColorTranslator.FromHtml("#000000");

                BackColor = bgColor;
                ForeColor = fgColor;

                // Labels inherit parent usually
                lblHeader.ForeColor = fgColor;
                lblServerLabel.ForeColor = fgColor;
                if (lblName != null)
                {
                    lblName.ForeColor = fgColor;
                }
                lblColor.ForeColor = fgColor;
                lblServerValue.ForeColor = fgColor;

                if (lblDatabaseLabel != null)
                {
                    lblDatabaseLabel.ForeColor = fgColor;
                }

                if (lblDatabaseValue != null)
                {
                    lblDatabaseValue.ForeColor = fgColor;
                }
                
                // Labels in Panel 1
                foreach(Control c in panelAlias.Controls)
                {
                     if(c is Label) c.ForeColor = fgColor;
                }
                
                // TextBoxes
                if (txtName != null)
                {
                    txtName.BackColor = txtBg;
                    txtName.ForeColor = txtFg;
                }

                txtAlias.BackColor = txtBg;
                txtAlias.ForeColor = txtFg;

                // ComboBox
                cmbColor.BackColor = txtBg;
                cmbColor.ForeColor = txtFg;

                comboBackColor = txtBg;
                comboTextColor = txtFg;
                comboHighlightBack = BlendColors(txtBg, Color.White, 0.08f);
                comboHighlightFore = txtFg;
                comboArrowColor = BlendColors(txtFg, txtBg, 0.35f);
                comboBorderColor = BlendColors(txtBg, txtFg, 0.25f);

                if (cmbColor is ModernComboBox modernCombo)
                {
                    modernCombo.ArrowColor = comboArrowColor;
                    modernCombo.ButtonBackColor = txtBg;
                    modernCombo.BorderColor = comboBorderColor;
                }

                primaryButtonBack = accentBg;
                primaryButtonHoverBack = ColorTranslator.FromHtml("#867bda");
                primaryButtonFore = accentFg;
                secondaryButtonBack = ColorTranslator.FromHtml("#353535");
                secondaryButtonHoverBack = ColorTranslator.FromHtml("#3a3a3a");
                secondaryButtonFore = fgColor;
                focusButtonBorder = fgColor;

                if (panelRuleButtons != null)
                {
                    panelRuleButtons.BackColor = ColorTranslator.FromHtml("#282828");
                }

                if (panelAliasButtons != null)
                {
                    panelAliasButtons.BackColor = ColorTranslator.FromHtml("#282828");
                }

                if (panelRule != null)
                {
                    panelRule.BackColor = bgColor;
                }

                if (panelAlias != null)
                {
                    panelAlias.BackColor = bgColor;
                }

                RefreshButtons();
            }
            catch
            {
                // Fallback to standard windows theme if VS service fails
            }
        }

        private void TryEnableRoundedCorners()
        {
            if (!IsHandleCreated)
            {
                return;
            }

            if (!IsWindows10OrGreater())
            {
                return;
            }

            try
            {
                int preference = 2; // DWMWCP_ROUND
                DwmSetWindowAttribute(Handle, DWMWA_WINDOW_CORNER_PREFERENCE, ref preference, sizeof(int));
            }
            catch
            {
                // Ignore if unsupported
            }
        }

        private static bool IsWindows10OrGreater()
        {
            Version version = Environment.OSVersion.Version;
            return version.Major >= 10;
        }

        private const int DWMWA_WINDOW_CORNER_PREFERENCE = 33;

        private const int WM_UPDATEUISTATE = 0x0128;
        private const int UIS_CLEAR = 2;
        private const int UISF_HIDEACCEL = 0x2;

        [DllImport("dwmapi.dll", PreserveSig = true)]
        private static extern int DwmSetWindowAttribute(IntPtr hwnd, int dwAttribute, ref int pvAttribute, int cbAttribute);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

        private void ShowKeyboardCuesAlways()
        {
            if (!IsHandleCreated)
            {
                return;
            }

            int wParam = (UISF_HIDEACCEL << 16) | UIS_CLEAR;
            SendMessage(Handle, WM_UPDATEUISTATE, (IntPtr)wParam, IntPtr.Zero);
        }


        private void CmbColor_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;

            var item = (ColorItem)cmbColor.Items[e.Index];
            string displayName = usedColorIndexes.Contains(item.Index) ? $"{item.Name} (used)" : item.Name;

            e.DrawBackground();

            bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
            Color backColor = comboBackColor != Color.Empty ? comboBackColor : e.BackColor;
            Color textColor = comboTextColor != Color.Empty ? comboTextColor : e.ForeColor;
            if (isSelected)
            {
                backColor = comboHighlightBack != Color.Empty ? comboHighlightBack : backColor;
                textColor = comboHighlightFore != Color.Empty ? comboHighlightFore : textColor;
            }

            using (var backBrush = new SolidBrush(backColor))
            {
                e.Graphics.FillRectangle(backBrush, e.Bounds);
            }

            // Draw Swatch (rounded, no outline)
            var swatchRect = new Rectangle(e.Bounds.Left + 6, e.Bounds.Top + 3, 18, e.Bounds.Height - 6);
            using (var brush = new SolidBrush(item.Color))
            using (var path = CreateRoundedRectanglePath(swatchRect, 3))
            {
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
                e.Graphics.FillPath(brush, path);
                e.Graphics.SmoothingMode = SmoothingMode.Default;
            }

            // Draw Text
            using (var brush = new SolidBrush(textColor))
            {
                e.Graphics.DrawString(displayName, e.Font, brush, e.Bounds.Left + 30, e.Bounds.Top + 1);
            }
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                boldFont?.Dispose();
            }

            base.Dispose(disposing);
        }

        private static GraphicsPath CreateRoundedRectanglePath(RectangleF rect, float radius)
        {
            float diameter = radius * 2;
            var path = new GraphicsPath();

            if (diameter <= 0)
            {
                path.AddRectangle(rect);
                return path;
            }

            var arc = new RectangleF(rect.Location, new SizeF(diameter, diameter));

            path.AddArc(arc, 180, 90);
            arc.X = rect.Right - diameter;
            path.AddArc(arc, 270, 90);
            arc.Y = rect.Bottom - diameter;
            path.AddArc(arc, 0, 90);
            arc.X = rect.Left;
            path.AddArc(arc, 90, 90);
            path.CloseFigure();

            return path;
        }

        private void ConfigureButton(Button button, bool isPrimary)
        {
            if (button == null)
            {
                return;
            }

            button.AutoSize = false;
            button.Height = 26;
            button.FlatStyle = FlatStyle.Flat;
            button.FlatAppearance.BorderSize = 0;
            button.UseVisualStyleBackColor = false;
            button.Tag = isPrimary;

            button.Font = new Font(Font.FontFamily, Font.Size - 1.0f);

            button.Paint -= Button_Paint;
            button.Paint += Button_Paint;
            button.GotFocus -= Button_FocusChanged;
            button.LostFocus -= Button_FocusChanged;
            button.GotFocus += Button_FocusChanged;
            button.LostFocus += Button_FocusChanged;

            button.MouseEnter -= Button_MouseEnter;
            button.MouseLeave -= Button_MouseLeave;
            button.MouseEnter += Button_MouseEnter;
            button.MouseLeave += Button_MouseLeave;
        }

        private void Button_MouseEnter(object sender, EventArgs e)
        {
            if (sender is Button button)
            {
                hoveredButtons.Add(button);
                button.Invalidate();
            }
        }

        private void Button_MouseLeave(object sender, EventArgs e)
        {
            if (sender is Button button)
            {
                hoveredButtons.Remove(button);
                button.Invalidate();
            }
        }

        private void Button_FocusChanged(object sender, EventArgs e)
        {
            if (sender is Control control)
            {
                control.Invalidate();
            }
        }

        private void Button_Paint(object sender, PaintEventArgs e)
        {
            if (!(sender is Button button))
            {
                return;
            }

            bool isPrimary = button.Tag is bool tag && tag;
            bool isHovered = hoveredButtons.Contains(button);
            Color back = isPrimary && primaryButtonBack != Color.Empty ? primaryButtonBack : secondaryButtonBack;
            if (isHovered)
            {
                back = isPrimary && primaryButtonHoverBack != Color.Empty ? primaryButtonHoverBack : secondaryButtonHoverBack;
            }
            Color fore = isPrimary && primaryButtonFore != Color.Empty ? primaryButtonFore : secondaryButtonFore;
            bool showOutline = button.Focused;
            Color parentBack = button.Parent?.BackColor ?? back;

            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            var rect = new RectangleF(0, 0, button.Width, button.Height);
            e.Graphics.Clear(parentBack);

            // Draw Background
            using (var path = CreateRoundedRectanglePath(rect, 8))
            using (var backBrush = new SolidBrush(back))
            {
                e.Graphics.FillPath(backBrush, path);
            }

            // Draw Focus Ring
            if (showOutline && focusButtonBorder != Color.Empty)
            {
                // Draw Outer Border (2px)
                // Center the pen at 1px inset so it draws from 0px to 2px
                var borderRect = RectangleF.Inflate(rect, -1f, -1f);
                using (var pen = new Pen(focusButtonBorder, 2f)) 
                using (var path = CreateRoundedRectanglePath(borderRect, 8))
                {
                    e.Graphics.DrawPath(pen, path);
                }

                // Draw Gap (1px, Parent Background) to create double-border effect
                // The outer border ends at 2px. The gap should occupy 2px to 3px.
                // So center of gap pen is at 2.5px.
                var gapRect = RectangleF.Inflate(rect, -2.5f, -2.5f);
                using (var gapPen = new Pen(parentBack, 1f))
                using (var gapPath = CreateRoundedRectanglePath(gapRect, 7f)) // Slightly tighter radius
                {
                    e.Graphics.DrawPath(gapPen, gapPath);
                }
            }

            TextRenderer.DrawText(
                e.Graphics,
                button.Text,
                button.Font,
                new Rectangle(0, 0, button.Width, button.Height),
                fore,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.NoPadding
            );

            e.Graphics.SmoothingMode = SmoothingMode.Default;
        }

        private void RefreshButtons()
        {
            foreach (var btn in new[] { btnSave, btnCancel, btnOpenConfig, btnNext, btnCancelAlias })
            {
                btn?.Invalidate();
            }
        }

        private static Color BlendColors(Color baseColor, Color blendColor, float blendAmount)
        {
            blendAmount = Math.Max(0, Math.Min(1, blendAmount));
            int r = (int)(baseColor.R + (blendColor.R - baseColor.R) * blendAmount);
            int g = (int)(baseColor.G + (blendColor.G - baseColor.G) * blendAmount);
            int b = (int)(baseColor.B + (blendColor.B - baseColor.B) * blendAmount);
            return Color.FromArgb(baseColor.A, r, g, b);
        }

        private sealed class ModernComboBox : ComboBox
        {
            public Color ArrowColor { get; set; } = SystemColors.ControlText;
            public Color ButtonBackColor { get; set; } = Color.Empty;
            public Color BorderColor { get; set; } = Color.Empty;

            public ModernComboBox()
            {
                // Enable UserPaint and double buffering to prevent flickering
                SetStyle(ControlStyles.UserPaint | 
                         ControlStyles.ResizeRedraw |
                         ControlStyles.OptimizedDoubleBuffer | 
                         ControlStyles.AllPaintingInWmPaint, true);
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                // Do NOT call base.OnPaint(e) to avoid interference/ghosting

                // 1. Draw Background
                using (var brush = new SolidBrush(BackColor))
                {
                    e.Graphics.FillRectangle(brush, ClientRectangle);
                }

                // 2. Draw Selected Item Content
                if (SelectedIndex >= 0)
                {
                    // Center the item vertically so the swatch size (derived from height) 
                    // matches the one in the dropdown list consistently.
                    // This prevents the "elongated" look or duplicate artifacts.
                    int itemHeight = ItemHeight > 0 ? ItemHeight : Height;
                    int yOffset = (Height - itemHeight) / 2;
                    if (yOffset < 0) yOffset = 0;

                    // Manually invoke the DrawItem event to render the text/swatch using the parent's logic.
                    // Reduce width to exclude the arrow button area.
                    Rectangle textRect = new Rectangle(0, yOffset, Width - 18, itemHeight);
                    
                    var args = new DrawItemEventArgs(
                        e.Graphics,
                        Font,
                        textRect,
                        SelectedIndex,
                        DrawItemState.None,
                        ForeColor,
                        BackColor);

                    OnDrawItem(args);
                }

                // 3. Draw Arrow and Border Overlay
                int buttonWidth = 18;
                var rect = new Rectangle(Width - buttonWidth - 2, 2, buttonWidth, Height - 4);
                
                e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

                int cx = rect.Left + rect.Width / 2;
                int cy = rect.Top + rect.Height / 2 - 1;
                using (var pen = new Pen(ArrowColor, 1.6f))
                {
                    e.Graphics.DrawLine(pen, cx - 4, cy - 1, cx, cy + 3);
                    e.Graphics.DrawLine(pen, cx, cy + 3, cx + 4, cy - 1);
                }

                if (BorderColor != Color.Empty)
                {
                    using (var pen = new Pen(BorderColor, 1f))
                    {
                        e.Graphics.DrawRectangle(pen, 0, 0, Width - 1, Height - 1);
                    }
                }
                
                e.Graphics.SmoothingMode = SmoothingMode.Default;
            }
        }
    }
}
