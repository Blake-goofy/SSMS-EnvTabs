using System;
using System.Drawing;
using System.Windows.Forms;
using System.Runtime.InteropServices;

namespace SSMS_EnvTabs
{
    internal sealed class UpdatePromptDialog : Form
    {
        private readonly Action releaseNotesRequested;
        private readonly Action openConfigRequested;

        private static readonly Color TopPanelColor = ColorTranslator.FromHtml("#2c2c2c");
        private static readonly Color BottomPanelColor = ColorTranslator.FromHtml("#282828");
        private static readonly Color PrimaryButtonColor = ColorTranslator.FromHtml("#9184ee");
        private static readonly Color PrimaryButtonHoverColor = ColorTranslator.FromHtml("#867bda");
        private static readonly Color SecondaryButtonColor = ColorTranslator.FromHtml("#353535");
        private static readonly Color SecondaryButtonHoverColor = ColorTranslator.FromHtml("#3a3a3a");
        private static readonly Color TextColor = ColorTranslator.FromHtml("#f0f0f0");

        public UpdatePromptDialog(string latestVersion, string currentVersion, Action releaseNotesRequested, Action openConfigRequested)
        {
            this.releaseNotesRequested = releaseNotesRequested;
            this.openConfigRequested = openConfigRequested;
            Text = "SSMS EnvTabs Update";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoSize = false;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(0);
            ClientSize = new Size(460, 240);

            BackColor = TopPanelColor;
            ForeColor = TextColor;

            var header = new Label
            {
                AutoSize = true,
                MaximumSize = new Size(420, 0),
                Text = "EnvTabs update available!",
                Font = new Font(SystemFonts.MessageBoxFont.FontFamily, SystemFonts.MessageBoxFont.Size + 5f, FontStyle.Bold)
            };

            var versionTable = new TableLayoutPanel
            {
                ColumnCount = 2,
                RowCount = 2,
                AutoSize = true,
                Margin = Padding.Empty,
                Padding = Padding.Empty
            };
            versionTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
            versionTable.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

            var currentLabel = new Label
            {
                AutoSize = true,
                Text = "Current:",
                Font = new Font(SystemFonts.MessageBoxFont.FontFamily, SystemFonts.MessageBoxFont.Size + 1f, FontStyle.Regular)
            };
            var currentValue = new Label
            {
                AutoSize = true,
                Text = currentVersion,
                Font = new Font(SystemFonts.MessageBoxFont.FontFamily, SystemFonts.MessageBoxFont.Size + 2f, FontStyle.Bold),
                ForeColor = TextColor
            };
            var availableLabel = new Label
            {
                AutoSize = true,
                Text = "Available:",
                Font = new Font(SystemFonts.MessageBoxFont.FontFamily, SystemFonts.MessageBoxFont.Size + 1f, FontStyle.Regular)
            };
            var availableValue = new Label
            {
                AutoSize = true,
                Text = latestVersion,
                Font = new Font(SystemFonts.MessageBoxFont.FontFamily, SystemFonts.MessageBoxFont.Size + 2f, FontStyle.Bold),
                ForeColor = TextColor
            };

            versionTable.Controls.Add(currentLabel, 0, 0);
            versionTable.Controls.Add(currentValue, 1, 0);
            versionTable.Controls.Add(availableLabel, 0, 1);
            versionTable.Controls.Add(availableValue, 1, 1);

            var releaseNotesLink = new LinkLabel
            {
                AutoSize = true,
                MaximumSize = new Size(420, 0),
                Text = "Full release notes here.",
                LinkColor = ColorTranslator.FromHtml("#bcbcbc"),
                ActiveLinkColor = ColorTranslator.FromHtml("#ffffff"),
                VisitedLinkColor = ColorTranslator.FromHtml("#bcbcbc")
            };
            releaseNotesLink.LinkClicked += (sender, args) =>
            {
                releaseNotesRequested?.Invoke();
            };

            var configNote = new Label
            {
                AutoSize = true,
                MaximumSize = new Size(420, 0),
                Text = "To disable update checking, press \"Open Config\" below.",
                ForeColor = ColorTranslator.FromHtml("#bcbcbc")
            };

            var installButton = new RoundedButton
            {
                Text = "&Install",
                DialogResult = DialogResult.Yes,
                AutoSize = false,
                Size = new Size(110, 26),
                TabIndex = 0,
                IsPrimary = true,
                UseShieldIcon = true
            };

            var openConfigButton = new RoundedButton
            {
                Text = "&Open Config",
                AutoSize = false,
                Size = new Size(120, 26),
                TabIndex = 1,
                IsPrimary = false,
                DialogResult = DialogResult.OK
            };
            openConfigButton.Click += (sender, args) =>
            {
                openConfigRequested?.Invoke();
                Close();
            };

            var laterButton = new RoundedButton
            {
                Text = "&Later",
                DialogResult = DialogResult.Cancel,
                AutoSize = false,
                Size = new Size(110, 26),
                TabIndex = 2,
                IsPrimary = false
            };

            AcceptButton = null;
            CancelButton = laterButton;

            var topPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BackColor = TopPanelColor,
                Padding = new Padding(16, 14, 16, 10)
            };

            var bottomPanel = new Panel
            {
                Dock = DockStyle.Bottom,
                Height = 48,
                BackColor = BottomPanelColor
            };

            var mainLayout = new TableLayoutPanel
            {
                ColumnCount = 1,
                RowCount = 3,
                Dock = DockStyle.Fill,
                Padding = Padding.Empty,
                Margin = Padding.Empty
            };
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
            mainLayout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

            var topBlock = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = true,
                Dock = DockStyle.Fill
            };
            topBlock.Controls.Add(header);
            topBlock.Controls.Add(new Label { Height = 6, AutoSize = false });
            topBlock.Controls.Add(versionTable);

            var bottomBlock = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoSize = true,
                Dock = DockStyle.Fill
            };
            bottomBlock.Controls.Add(new Label { Height = 18, AutoSize = false });
            bottomBlock.Controls.Add(releaseNotesLink);
            bottomBlock.Controls.Add(new Label { Height = 6, AutoSize = false });
            bottomBlock.Controls.Add(configNote);

            mainLayout.Controls.Add(topBlock, 0, 0);
            mainLayout.Controls.Add(new Panel { Dock = DockStyle.Fill }, 0, 1);
            mainLayout.Controls.Add(bottomBlock, 0, 2);

            topPanel.Controls.Add(mainLayout);

            var buttonFlow = new FlowLayoutPanel
            {
                AutoSize = true,
                FlowDirection = FlowDirection.LeftToRight,
                WrapContents = false,
                BackColor = BottomPanelColor,
                Margin = Padding.Empty,
                Padding = Padding.Empty
            };

            installButton.Margin = new Padding(6, 10, 6, 10);
            openConfigButton.Margin = new Padding(6, 10, 6, 10);
            laterButton.Margin = new Padding(6, 10, 6, 10);

            buttonFlow.Controls.Add(installButton);
            buttonFlow.Controls.Add(openConfigButton);
            buttonFlow.Controls.Add(laterButton);

            bottomPanel.Controls.Add(buttonFlow);
            bottomPanel.Resize += (s, e) => CenterFlow(bottomPanel, buttonFlow);

            Controls.Add(topPanel);
            Controls.Add(bottomPanel);

            installButton.ApplyColors(PrimaryButtonColor, PrimaryButtonHoverColor, Color.Black, TextColor);
            openConfigButton.ApplyColors(SecondaryButtonColor, SecondaryButtonHoverColor, TextColor, TextColor);
            laterButton.ApplyColors(SecondaryButtonColor, SecondaryButtonHoverColor, TextColor, TextColor);
            ShowKeyboardCuesAlways();

            ActiveControl = releaseNotesLink;
        }

        protected override bool ShowKeyboardCues => true;

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            ShowKeyboardCuesAlways();
        }

        private static void CenterFlow(Panel panel, Control flow)
        {
            int x = Math.Max(0, (panel.Width - flow.Width) / 2);
            int y = Math.Max(0, (panel.Height - flow.Height) / 2);
            flow.Location = new Point(x, y);
        }

        private const int WM_UPDATEUISTATE = 0x0128;
        private const int UIS_CLEAR = 2;
        private const int UISF_HIDEACCEL = 0x2;

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

        private sealed class RoundedButton : Button
        {
            private Color backColor = SecondaryButtonColor;
            private Color hoverBackColor = SecondaryButtonColor;
            private Color foreColor = TextColor;
            private Color focusBorder = TextColor;
            private Image shieldIcon;
            private bool isHovered;

            public bool IsPrimary { get; set; }
            public bool UseShieldIcon { get; set; }

            public RoundedButton()
            {
                SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint, true);
                FlatStyle = FlatStyle.Flat;
                FlatAppearance.BorderSize = 0;
                Font = new Font(SystemFonts.MessageBoxFont.FontFamily, SystemFonts.MessageBoxFont.Size - 1f);
            }

            public void ApplyColors(Color back, Color hoverBack, Color fore, Color focus)
            {
                backColor = back;
                hoverBackColor = hoverBack;
                foreColor = fore;
                focusBorder = focus;
                shieldIcon = SystemIcons.Shield.ToBitmap();
                Invalidate();
            }

            protected override void OnMouseEnter(EventArgs e)
            {
                base.OnMouseEnter(e);
                isHovered = true;
                Invalidate();
            }

            protected override void OnMouseLeave(EventArgs e)
            {
                base.OnMouseLeave(e);
                isHovered = false;
                Invalidate();
            }

            protected override void OnPaint(PaintEventArgs e)
            {
                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
                var rect = new RectangleF(0, 0, Width, Height);
                e.Graphics.Clear(Parent?.BackColor ?? BackColor);

                Color fillColor = isHovered ? hoverBackColor : backColor;
                using (var path = CreateRoundedRectanglePath(rect, 8))
                using (var brush = new SolidBrush(fillColor))
                {
                    e.Graphics.FillPath(brush, path);
                }

                if (Focused)
                {
                    // Draw Outer Border (2px)
                    var borderRect = RectangleF.Inflate(rect, -1f, -1f);
                    using (var pen = new Pen(focusBorder, 2f))
                    using (var path = CreateRoundedRectanglePath(borderRect, 8))
                    {
                        e.Graphics.DrawPath(pen, path);
                    }

                    // Draw Gap (1px, Parent Background)
                    var gapRect = RectangleF.Inflate(rect, -2.5f, -2.5f);
                    using (var gapPen = new Pen(Parent?.BackColor ?? BackColor, 1f))
                    using (var gapPath = CreateRoundedRectanglePath(gapRect, 7f))
                    {
                        e.Graphics.DrawPath(gapPen, gapPath);
                    }
                }

                if (UseShieldIcon && shieldIcon != null)
                {
                    int iconSize = 16;
                    int gap = 6;
                    int textWidth = TextRenderer.MeasureText(Text, Font).Width;
                    int totalWidth = iconSize + gap + textWidth;
                    int startX = Math.Max(0, (Width - totalWidth) / 2);
                    int iconY = (Height - iconSize) / 2;
                    e.Graphics.DrawImage(shieldIcon, new Rectangle(startX, iconY, iconSize, iconSize));

                    var textRect = new Rectangle(startX + iconSize + gap, 0, textWidth, Height);
                    TextRenderer.DrawText(
                        e.Graphics,
                        Text,
                        Font,
                        textRect,
                        foreColor,
                        TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine
                    );
                }
                else
                {
                    TextRenderer.DrawText(
                        e.Graphics,
                        Text,
                        Font,
                        new Rectangle(0, 0, Width, Height),
                        foreColor,
                        TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine
                    );
                }

                e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.Default;
            }

            private static System.Drawing.Drawing2D.GraphicsPath CreateRoundedRectanglePath(RectangleF rect, float radius)
            {
                float diameter = radius * 2;
                var path = new System.Drawing.Drawing2D.GraphicsPath();
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
        }
    }
}
