using System;
using System.Drawing;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using Microsoft.VisualStudio.PlatformUI;

namespace SSMS_EnvTabs
{
    public class NewRuleDialog : Form
    {
        public string RuleName { get; private set; }
        public int SelectedColorIndex { get; private set; }
        public bool OpenConfigRequested { get; private set; }

        private TextBox txtName;
        private ComboBox cmbColor;
        private Button btnSave;
        private Button btnCancel;
        private Button btnOpenConfig;
        private Label lblInfo;
        private Label lblName;
        private Label lblColor;

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

        public NewRuleDialog(string server, string database, string suggestedName, int suggestedColorIndex)
        {
            InitializeComponent(server, database, suggestedColorIndex);
            ApplyVsTheme();
            
            txtName.Text = suggestedName;
        }

        private void InitializeComponent(string server, string database, int suggestedColorIndex)
        {
            this.Text = "SSMS EnvTabs - New Rule";
            this.Size = new Size(400, 240);
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.MaximizeBox = false;
            this.MinimizeBox = false;

            lblInfo = new Label
            {
                Text = $"New connection detected:\nServer: {server}\nDatabase: {database ?? "(any)"}",
                Location = new Point(15, 15),
                Size = new Size(360, 40),
                AutoSize = false
            };
            this.Controls.Add(lblInfo);

            lblName = new Label { Text = "Group Name:", Location = new Point(15, 65), AutoSize = true };
            this.Controls.Add(lblName);

            txtName = new TextBox { Location = new Point(100, 62), Size = new Size(260, 23) };
            this.Controls.Add(txtName);

            lblColor = new Label { Text = "Color:", Location = new Point(15, 95), AutoSize = true };
            this.Controls.Add(lblColor);

            cmbColor = new ComboBox 
            { 
                Location = new Point(100, 92), 
                Size = new Size(260, 23),
                DropDownStyle = ComboBoxStyle.DropDownList,
                DrawMode = DrawMode.OwnerDrawFixed
            };
            this.Controls.Add(cmbColor);

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
            
            // Set DataSource after adding to controls to ensure BindingContext is ready
            cmbColor.DataSource = orderedList;
            
            // Note: DataSource automatically sets SelectedIndex to 0 if list is not empty.
            // We rely on that behavior since orderedList[0] is our desired selection.

            btnSave = new Button { Text = "Save", Location = new Point(90, 150), DialogResult = DialogResult.OK };
            btnSave.Click += (s, e) => { RuleName = txtName.Text; SelectedColorIndex = ((ColorItem)cmbColor.SelectedItem).Index; };
            this.Controls.Add(btnSave);

            btnCancel = new Button { Text = "Cancel", Location = new Point(170, 150), DialogResult = DialogResult.Cancel };
            this.Controls.Add(btnCancel);

            btnOpenConfig = new Button { Text = "Open Config", Location = new Point(250, 150), Size = new Size(120, 23), DialogResult = DialogResult.Yes };
            btnOpenConfig.Click += (s, e) => { 
                RuleName = txtName.Text; 
                SelectedColorIndex = ((ColorItem)cmbColor.SelectedItem).Index; 
                OpenConfigRequested = true; 
            };
            this.Controls.Add(btnOpenConfig);
                        
            btnSave.Location = new Point(50, 150);
            btnOpenConfig.Location = new Point(130, 150);
            btnCancel.Location = new Point(260, 150);
            
            this.AcceptButton = btnSave;
            this.CancelButton = btnCancel;
        }

        private void ApplyVsTheme()
        {
            try
            {
                // Attempt to get VS colors. 
                // Note: VSColorTheme.GetThemedColor returns System.Drawing.Color in VSSDK.
                Color bgColor = VSColorTheme.GetThemedColor(EnvironmentColors.ToolWindowBackgroundColorKey);
                Color fgColor = VSColorTheme.GetThemedColor(EnvironmentColors.ToolWindowTextColorKey);
                Color btnBg = VSColorTheme.GetThemedColor(EnvironmentColors.CommandBarGradientBeginColorKey); // Just a guess for button
                Color txtBg = VSColorTheme.GetThemedColor(EnvironmentColors.ComboBoxBackgroundColorKey);
                Color txtFg = VSColorTheme.GetThemedColor(EnvironmentColors.ComboBoxTextColorKey);

                this.BackColor = bgColor;
                this.ForeColor = fgColor;

                // Labels inherit parent usually
                lblInfo.ForeColor = fgColor;
                lblName.ForeColor = fgColor;
                lblColor.ForeColor = fgColor;

                // TextBoxes
                txtName.BackColor = txtBg;
                txtName.ForeColor = txtFg;
                
                // ComboBox
                cmbColor.BackColor = txtBg;
                cmbColor.ForeColor = txtFg;

                // Buttons
                // WinForms Buttons are hard to style perfectly flat without custom painting,
                // but we can try to make them blend in.
                foreach (var btn in new[] { btnSave, btnCancel, btnOpenConfig })
                {
                    btn.FlatStyle = FlatStyle.Flat;
                    btn.BackColor = bgColor; 
                    btn.ForeColor = fgColor;
                    btn.FlatAppearance.BorderColor = fgColor;
                }
            }
            catch
            {
                // Fallback to standard windows theme if VS service fails
            }
        }


        private void CmbColor_DrawItem(object sender, DrawItemEventArgs e)
        {
            if (e.Index < 0) return;

            var item = (ColorItem)cmbColor.Items[e.Index];

            e.DrawBackground();

            // Draw Swatch
            using (var brush = new SolidBrush(item.Color))
            {
                e.Graphics.FillRectangle(brush, e.Bounds.Left + 2, e.Bounds.Top + 2, 20, e.Bounds.Height - 4);
            }
            e.Graphics.DrawRectangle(SystemPens.ControlDark, e.Bounds.Left + 2, e.Bounds.Top + 2, 20, e.Bounds.Height - 4);

            // Draw Text
            using (var brush = new SolidBrush(e.ForeColor))
            {
                e.Graphics.DrawString(item.Name, e.Font, brush, e.Bounds.Left + 30, e.Bounds.Top + 1);
            }

            e.DrawFocusRectangle();
        }
    }
}
