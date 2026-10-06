$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
# Run after a Debug build: powershell.exe -NoProfile -Sta -File tests/CheckConnectionGroupTint.ps1
$vsDirectory = [IO.DirectoryInfo](Split-Path (Get-Command MSBuild.exe -ErrorAction Stop).Source)
while ($vsDirectory -and !(Test-Path (Join-Path $vsDirectory.FullName 'Common7/IDE'))) { $vsDirectory = $vsDirectory.Parent }
if (!$vsDirectory) { throw 'Could not locate Visual Studio from MSBuild.exe' }
$ide = Join-Path $vsDirectory.FullName 'Common7/IDE'
[AppDomain]::CurrentDomain.add_AssemblyResolve({
    param($sender, $eventArgs)
    $name = ([Reflection.AssemblyName]$eventArgs.Name).Name + '.dll'
    foreach ($directory in @("$ide\PublicAssemblies", "$ide\PrivateAssemblies")) {
        $path = Join-Path $directory $name
        if (Test-Path -LiteralPath $path) { return [Reflection.Assembly]::LoadFrom($path) }
    }
    return $null
})
$assembly = [Reflection.Assembly]::LoadFrom((Join-Path $PSScriptRoot '..\SSMS EnvTabs\bin\Debug\SSMS EnvTabs.dll'))
$type = $assembly.GetType('SSMS_EnvTabs.SettingsToolWindowControl', $true)
$control = [Activator]::CreateInstance($type)
$flags = [Reflection.BindingFlags]'Instance,NonPublic'
$colorToggle = New-Object Windows.Controls.CheckBox
$colorToggle.IsChecked = $true
$type.GetField('autoColorToggle', $flags).SetValue($control, $colorToggle)
$method = $type.GetMethod('ResolveConnectionGroupBackground', $flags)
function Check($enabled, $index, $expected) {
    $colorIndex = if ($null -eq $index) { $null } else { [int]$index }
    $brush = $method.Invoke($control, [object[]]@([bool]$enabled, $colorIndex))
    if ($brush.Color.ToString() -ne $expected) { throw "Expected $expected, got $($brush.Color)" }
}
$control.Resources['EnvTabsBackgroundBrush'] = [Windows.Media.Brushes]::White
Check $false 3 '#FFFFFFFF'
Check $true $null '#FFFFFFFF'
Check $true -1 '#FFFFFFFF'
Check $true 16 '#FFFFFFFF'
Check $true 3 '#FFFBF3F3'
$control.Resources['EnvTabsBackgroundBrush'] = [Windows.Media.SolidColorBrush]::new([Windows.Media.Color]::FromRgb(30,30,30))
Check $true 3 '#FF2C2424'
$type.GetField('editorTintStrength', $flags).SetValue($control, 30)
Check $true 3 '#FF533334'
$colorToggle.IsChecked = $false
Check $true 3 '#FF1E1E1E'
Write-Output 'Compiled WPF preview check passed: light/dark themes, strength, disabled state, invalid and missing colors.'


$colorToggle.IsChecked = $true
$editorToggle = [Windows.Controls.CheckBox]::new()
$type.GetField('editorTintToggle', $flags).SetValue($control, $editorToggle)
$type.GetField('editorTintStrength', $flags).SetValue($control, 8)
$cardType = $type.GetNestedType('ConnectionGroupCardState', [Reflection.BindingFlags]'NonPublic')
$cards = $type.GetField('connectionGroupCards', $flags).GetValue($control)
foreach ($entry in @(@('Tint enabled', $true), @('Tint disabled', $false), @('Default (off)', $null))) {
    $card = [Activator]::CreateInstance($cardType)
    $cardType.GetProperty('GroupName').SetValue($card, $entry[0])
    $cardType.GetProperty('ColorIndex').SetValue($card, [int]3)
    $cardType.GetProperty('EnableEditorTint').SetValue($card, $entry[1])
    $cards.Add($card)
}
$threadHelper = [Reflection.Assembly]::LoadFrom("$ide\PublicAssemblies\Microsoft.VisualStudio.Shell.15.0.dll").GetType('Microsoft.VisualStudio.Shell.ThreadHelper')
$threadHelper.GetMethod('SetUIThread', [Reflection.BindingFlags]'Static,NonPublic,Public').Invoke($null, @()) | Out-Null
$type.GetMethod('RebuildConnectionGroupCardsUi', $flags).Invoke($control, @()) | Out-Null
$panel = $type.GetField('connectionGroupCardsPanel', $flags).GetValue($control)
if ($panel.Children[0].Background.Color.ToString() -ne '#FF1E1E1E' -or $panel.Children[1].Background.Color.ToString() -ne '#FF1E1E1E' -or $panel.Children[2].Background.Color.ToString() -ne '#FF2C2424') { throw 'Saved card tint mismatch' }
Write-Output 'Actual saved group cards passed.'

$editorToggle.IsChecked = $true
$type.GetMethod('RebuildConnectionGroupCardsUi', $flags).Invoke($control, @()) | Out-Null
if ($panel.Children[0].Background.Color.ToString() -ne '#FF2C2424' -or $panel.Children[1].Background.Color.ToString() -ne '#FF1E1E1E') { throw 'Global tint inheritance or explicit opt-out failed' }
$editorToggle.IsChecked = $false


$type.GetMethod('BeginConnectionGroupEdit', $flags).Invoke($control, [object[]]@($cards[0])) | Out-Null
$checkbox = $type.GetField('activeConnectionEditorTintCheckBox', $flags).GetValue($control)
$combo = $type.GetField('activeConnectionColorCombo', $flags).GetValue($control)
$editingBorder = $panel.Children[2]
$checkbox.IsChecked = $false
if ($editingBorder.Background.Color.ToString() -ne '#FF1E1E1E') { throw 'Turning tint off did not update the card' }
$checkbox.IsChecked = $true
$combo.SelectedIndex = 3
if ($editingBorder.Background.Color.ToString() -ne '#FF1F2A2C') { throw "Color preview mismatch: $($editingBorder.Background.Color)" }
$type.GetMethod('CancelConnectionGroupEdit', $flags).Invoke($control, [object[]]@($cards[0])) | Out-Null
if ($panel.Children[2].Background.Color.ToString() -ne '#FF2C2424') { throw 'Cancel did not restore the saved tint' }
Write-Output 'Actual editing cards passed: immediate toggle/color previews and Cancel restore.'
