using System;
using System.Drawing;
using System.Windows.Forms;

namespace SSMS_EnvTabs
{
    internal sealed class UpdatePromptDialog : Form
    {
        private readonly Action releaseNotesRequested;

        public UpdatePromptDialog(string latestVersion, string currentVersion, Action releaseNotesRequested)
        {
            this.releaseNotesRequested = releaseNotesRequested;
            Text = "SSMS EnvTabs Update";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(16);

            var message = new Label
            {
                AutoSize = true,
                MaximumSize = new Size(480, 0),
                Text = $"A new SSMS EnvTabs version ({latestVersion}) is available.\n\nCurrent version: {currentVersion}"
            };

            var installButton = new Button
            {
                Text = "Install",
                DialogResult = DialogResult.Yes,
                AutoSize = true,
                TabIndex = 0
            };

            var releaseNotesButton = new Button
            {
                Text = "Release notes",
                AutoSize = true,
                TabIndex = 1
            };

            releaseNotesButton.Click += (sender, args) =>
            {
                releaseNotesRequested?.Invoke();
            };

            var laterButton = new Button
            {
                Text = "Later",
                DialogResult = DialogResult.Cancel,
                AutoSize = true,
                TabIndex = 2
            };

            AcceptButton = installButton;
            CancelButton = laterButton;

            var buttonPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.LeftToRight,
                Dock = DockStyle.Bottom,
                AutoSize = true,
                Padding = new Padding(0, 12, 0, 0)
            };

            buttonPanel.Controls.Add(installButton);
            buttonPanel.Controls.Add(releaseNotesButton);
            buttonPanel.Controls.Add(laterButton);

            var layout = new TableLayoutPanel
            {
                AutoSize = true,
                ColumnCount = 1,
                RowCount = 2,
                Dock = DockStyle.Fill
            };

            layout.Controls.Add(message, 0, 0);
            layout.Controls.Add(buttonPanel, 0, 1);

            Controls.Add(layout);

            NativeMethods.SetShield(installButton, true);
        }

        private static class NativeMethods
        {
            private const int BCM_SETSHIELD = 0x160C;

            [System.Runtime.InteropServices.DllImport("user32.dll", CharSet = System.Runtime.InteropServices.CharSet.Unicode)]
            private static extern IntPtr SendMessage(IntPtr hWnd, int msg, IntPtr wParam, IntPtr lParam);

            public static void SetShield(Button button, bool show)
            {
                if (button == null || button.IsDisposed) return;
                button.FlatStyle = FlatStyle.System;
                SendMessage(button.Handle, BCM_SETSHIELD, IntPtr.Zero, show ? new IntPtr(1) : IntPtr.Zero);
            }
        }
    }
}
