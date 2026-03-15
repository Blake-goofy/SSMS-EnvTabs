using System;
using System.Runtime.InteropServices;
using Microsoft.VisualStudio.Shell;

namespace SSMS_EnvTabs
{
    [Guid("D9334B9A-6A80-4308-99C2-EA45E9B28E9E")]
    public sealed class SettingsToolWindow : ToolWindowPane
    {
        public SettingsToolWindow() : base(null)
        {
            Caption = "SSMS EnvTabs | Settings";
            Content = new SettingsToolWindowControl();
        }
    }
}