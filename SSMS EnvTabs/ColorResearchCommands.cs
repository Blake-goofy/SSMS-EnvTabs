using System;
using System.ComponentModel.Design;
using System.IO;
using System.Linq;
using System.Runtime.Serialization.Json;
using System.Text;
using System.Windows.Forms;
using Microsoft.VisualStudio.Shell;
using Microsoft.VisualStudio.Shell.Interop;
using SSMS_EnvTabs.ResearchModels;

namespace SSMS_EnvTabs
{
    // Simple models for reading the SSMS temp json file
    namespace ResearchModels
    {
        [System.Runtime.Serialization.DataContract]
        public class SsmsColorFile
        {
            [System.Runtime.Serialization.DataMember(Name = "ColorMap")]
            public System.Collections.Generic.Dictionary<string, ColorMapEntry> ColorMap { get; set; }
        }

        [System.Runtime.Serialization.DataContract]
        public class ColorMapEntry
        {
            [System.Runtime.Serialization.DataMember(Name = "GroupId")]
            public long GroupId { get; set; }

            [System.Runtime.Serialization.DataMember(Name = "ColorIndex")]
            public int ColorIndex { get; set; }
        }
    }

    internal sealed class ColorResearchCommands
    {
        private readonly AsyncPackage package;
        private readonly IMenuCommandService commandService;
        private string logFilePath;

        // Path to the dev/research.log file
        private string GetLogPath()
        {
            if (logFilePath != null) return logFilePath;
            
            // Try to find the solution or repo root
            // For now, hardcoding based on user context or falling back to temp
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            string repoPath = Path.Combine(userProfile, "source", "repos", "SSMS EnvTabs");
            if (Directory.Exists(repoPath))
            {
                logFilePath = Path.Combine(repoPath, "dev", "research.log");
            }
            else
            {
                logFilePath = Path.Combine(Path.GetTempPath(), "SSMS_EnvTabs_Research.log");
            }
            return logFilePath;
        }

        public static async System.Threading.Tasks.Task InitializeAsync(AsyncPackage package)
        {
            await ThreadHelper.JoinableTaskFactory.SwitchToMainThreadAsync(package.DisposalToken);
            var commandService = await package.GetServiceAsync(typeof(IMenuCommandService)) as IMenuCommandService;
            new ColorResearchCommands(package, commandService);
        }

        private ColorResearchCommands(AsyncPackage package, IMenuCommandService commandService)
        {
            this.package = package ?? throw new ArgumentNullException(nameof(package));
            this.commandService = commandService ?? throw new ArgumentNullException(nameof(commandService));

            RegisterCommand(SSMS_EnvTabsPackage.cmdidCaptureData, OnCaptureData);
        }

        private void RegisterCommand(int id, EventHandler handler)
        {
            var commandId = new CommandID(SSMS_EnvTabsPackage.PackageCmdSetGuid, id);
            var menuItem = new MenuCommand(handler, commandId);
            this.commandService.AddCommand(menuItem);
        }

        private static System.Collections.Generic.HashSet<long> _seenGroupIds = new System.Collections.Generic.HashSet<long>();

        private void OnCaptureData(object sender, EventArgs e)
        {
            ThreadHelper.ThrowIfNotOnUIThread();
            try
            {
                var file = FindLatestSsmsColorFile();
                if (file == null)
                {
                    MessageBox.Show("No SSMS color file found in Temp.", "Error");
                    return;
                }

                var data = ParseSsmsColorFile(file);
                if (data?.ColorMap == null || data.ColorMap.Count == 0) 
                {
                    MessageBox.Show("Color map empty or invalid.", "Error");
                    return;
                }

                // Find groups we haven't seen yet
                var newGroups = data.ColorMap.Values
                    .Where(g => !_seenGroupIds.Contains(g.GroupId))
                    .ToList();

                ColorMapEntry targetEntry = null;

                if (newGroups.Count == 1)
                {
                    targetEntry = newGroups[0];
                }
                else if (newGroups.Count > 1)
                {
                    // If multiple new, use the last one (based on typical JSON serialization order or just last in list)
                    targetEntry = newGroups.Last();
                    MessageBox.Show($"Found {newGroups.Count} new groups. Using the last one (ID: {targetEntry.GroupId})", "Note");
                }
                else
                {
                    // No new groups found. Fallback to existing logic or prompt?
                    // User says: "expect a new group to be made after each time".
                    // So if we don't find a new one, maybe the file hasn't updated yet?
                    var res = MessageBox.Show("No *new* groups found since last capture. Use the very last group in the file anyway?", "Warning", MessageBoxButtons.YesNo);
                    if (res == DialogResult.Yes)
                    {
                        targetEntry = data.ColorMap.Values.Last();
                    }
                    else
                    {
                        return;
                    }
                }

                // Mark as seen
                foreach (var g in data.ColorMap.Values)
                {
                    _seenGroupIds.Add(g.GroupId);
                }

                // Prompt for Salt
                string salt = PromptForSalt();
                if (salt == null) return; // User cancelled

                string filename = "myQuery1.sql"; // Hardcoded for this specific test plan
                string computedRegex = $"(?:^|[\\\\/])(?:{System.Text.RegularExpressions.Regex.Escape(filename)})$";
                
                if (!string.IsNullOrWhiteSpace(salt))
                {
                    computedRegex += $"(?#salt:{salt})";
                }

                string logLine = $"{computedRegex} = ColorIndex: {targetEntry.ColorIndex} (GroupId: {targetEntry.GroupId})";
                
                // Read existing lines to avoid duplicates if possible, or just append
                File.AppendAllText(GetLogPath(), logLine + Environment.NewLine);
                
                MessageBox.Show($"Captured!\nSalt: {salt}\nIndex: {targetEntry.ColorIndex}\nGroupId: {targetEntry.GroupId}", "Success");

            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error capturing: {ex.Message}");
            }
        }

        private string PromptForSalt()
        {
            Form prompt = new Form()
            {
                Width = 400,
                Height = 150,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = "Enter Salt",
                StartPosition = FormStartPosition.CenterScreen
            };
            Label textLabel = new Label() { Left = 20, Top = 20, Text = "Salt:" };
            TextBox textBox = new TextBox() { Left = 20, Top = 50, Width = 340 };
            Button confirmation = new Button() { Text = "Ok", Left = 280, Width = 80, Top = 80, DialogResult = DialogResult.OK };
            confirmation.Click += (sender, e) => { prompt.Close(); };
            prompt.Controls.Add(textBox);
            prompt.Controls.Add(confirmation);
            prompt.Controls.Add(textLabel);
            prompt.AcceptButton = confirmation;

            return prompt.ShowDialog() == DialogResult.OK ? textBox.Text : null;
        }

        private string FindLatestSsmsColorFile()
        {
            string temp = Path.GetTempPath();
            var allCandidates = new System.Collections.Generic.List<FileInfo>();

            try
            {
                // Instead of recursive search which hits permission errors,
                // we iterate top-level directories (where the SSMS GUID folders live).
                foreach (var dir in Directory.GetDirectories(temp))
                {
                    try
                    {
                        var info = new DirectoryInfo(dir);
                        var files = info.GetFiles("customized-groupid-color-*.json", SearchOption.TopDirectoryOnly);
                        allCandidates.AddRange(files);
                    }
                    catch
                    {
                        // Ignore directories we can't read (WinSAT, etc.)
                    }
                }
            }
            catch (Exception ex)
            {
                EnvTabsLog.Info($"Error searching for SSMS files: {ex.Message}");
            }

            return allCandidates
                .OrderByDescending(f => f.LastWriteTime)
                .FirstOrDefault()?.FullName;
        }

        private SsmsColorFile ParseSsmsColorFile(string path)
        {
            try
            {
                using (var ms = new MemoryStream(Encoding.UTF8.GetBytes(File.ReadAllText(path))))
                {
                    // SSMS uses the "Simple" Dictionary format {"Key": {Obj}, "Key": {Obj}}
                    // NOT the default DataContract format [{"Key":..., "Value":...}]
                    var settings = new DataContractJsonSerializerSettings
                    {
                        UseSimpleDictionaryFormat = true
                    };
                    var ser = new DataContractJsonSerializer(typeof(SsmsColorFile), settings);
                    return ser.ReadObject(ms) as SsmsColorFile;
                }
            }
            catch(Exception ex)
            {
                // Log needed for debugging
                EnvTabsLog.Error($"JSON Parse error: {ex.Message}");
                return null;
            }
        }
    }
}
