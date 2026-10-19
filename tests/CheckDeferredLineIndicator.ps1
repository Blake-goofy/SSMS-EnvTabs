$ErrorActionPreference = 'Stop'
# Run after a Debug build: powershell.exe -NoProfile -Sta -File tests/CheckDeferredLineIndicator.ps1
Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
$vsDirectory = [IO.DirectoryInfo](Split-Path (Get-Command MSBuild.exe -ErrorAction Stop).Source)
while ($vsDirectory -and !(Test-Path (Join-Path $vsDirectory.FullName 'Common7/IDE'))) { $vsDirectory = $vsDirectory.Parent }
if (!$vsDirectory) { throw 'Could not locate Visual Studio from MSBuild.exe' }
$ide = Join-Path $vsDirectory.FullName 'Common7/IDE'
[AppDomain]::CurrentDomain.add_AssemblyResolve({
    param($sender, $eventArgs)
    $name = ([Reflection.AssemblyName]$eventArgs.Name).Name + '.dll'
    foreach ($directory in @("$ide\PublicAssemblies", "$ide\PrivateAssemblies", "$ide\CommonExtensions\Microsoft\Editor")) {
        $path = Join-Path $directory $name
        if (Test-Path -LiteralPath $path) { return [Reflection.Assembly]::LoadFrom($path) }
    }
    return $null
})
# The proxy supplies just editor properties/events; the controller and WPF visual tree are real.
Add-Type -ReferencedAssemblies PresentationFramework, PresentationCore, WindowsBase, System.Xaml -TypeDefinition @'
using System;
using System.Collections.Generic;
using System.Runtime.Remoting.Messaging;
using System.Runtime.Remoting.Proxies;
using System.Runtime.Serialization;
public class Indicator : System.Windows.Controls.Border { }
public class EditorEventProxy : RealProxy
{
    public readonly Dictionary<string, object> Values = new Dictionary<string, object>();
    public readonly Dictionary<string, Delegate> Events = new Dictionary<string, Delegate>();
    public EditorEventProxy(Type contract) : base(contract) { }
    public override IMessage Invoke(IMessage message)
    {
        var call = (IMethodCallMessage)message;
        string name = call.MethodName;
        object value = null;
        if (name.StartsWith("get_")) Values.TryGetValue(name.Substring(4), out value);
        else if (name.StartsWith("add_") || name.StartsWith("remove_"))
        {
            string key = name.Substring(name.IndexOf('_') + 1);
            Delegate previous;
            Events.TryGetValue(key, out previous);
            Events[key] = name.StartsWith("add_") ? Delegate.Combine(previous, (Delegate)call.Args[0]) : Delegate.Remove(previous, (Delegate)call.Args[0]);
        }
        else if (name == "GetHashCode") value = GetHashCode();
        else if (name == "GetType") value = GetProxiedType();
        else if (name == "ToString") value = GetProxiedType().Name;
        else if (name == "Equals") value = Object.ReferenceEquals(GetTransparentProxy(), call.Args[0]);
        else if (name == "GetTextViewMargin") Values.TryGetValue("Margin:" + call.Args[0], out value);
        else throw new NotSupportedException(name);
        return new ReturnMessage(value, null, 0, call.LogicalCallContext, call);
    }
    public void Raise(string name)
    {
        Delegate handler;
        if (!Events.TryGetValue(name, out handler) || handler == null) return;
        Type argsType = handler.GetType().GetMethod("Invoke").GetParameters()[1].ParameterType;
        handler.DynamicInvoke(GetTransparentProxy(), FormatterServices.GetUninitializedObject(argsType));
    }
    public int SubscriberCount(string name) { return Events.ContainsKey(name) && Events[name] != null ? Events[name].GetInvocationList().Length : 0; }
}
'@
$assembly = [Reflection.Assembly]::LoadFrom((Join-Path $PSScriptRoot '..\SSMS EnvTabs\bin\Debug\SSMS EnvTabs.dll'))
$managerType = $assembly.GetType('SSMS_EnvTabs.RdtEventManager', $true)
$controllerType = $managerType.GetNestedType('LineIndicatorColorController', [Reflection.BindingFlags]'NonPublic')
$wpf = [Reflection.Assembly]::LoadFrom("$ide\CommonExtensions\Microsoft\Editor\Microsoft.VisualStudio.Text.UI.Wpf.dll")
$ui = [Reflection.Assembly]::LoadFrom("$ide\CommonExtensions\Microsoft\Editor\Microsoft.VisualStudio.Text.UI.dll")
$view = [EditorEventProxy]::new($wpf.GetType('Microsoft.VisualStudio.Text.Editor.IWpfTextView', $true))
$caret = [EditorEventProxy]::new($ui.GetType('Microsoft.VisualStudio.Text.Editor.ITextCaret', $true))
$root = [Windows.Controls.Grid]::new()
$marginRoot = [Windows.Controls.Grid]::new()
$hostProxy = [EditorEventProxy]::new($wpf.GetType('Microsoft.VisualStudio.Text.Editor.IWpfTextViewHost', $true))
$marginProxy = [EditorEventProxy]::new($wpf.GetType('Microsoft.VisualStudio.Text.Editor.IWpfTextViewMargin', $true))
$hostProxy.Values['HostControl'] = [Windows.Controls.ContentControl]::new()
$hostProxy.Values['Margin:Indicator'] = $marginProxy.GetTransparentProxy()
$marginProxy.Values['VisualElement'] = $marginRoot
$view.Values['VisualElement'] = $root
$view.Values['Caret'] = $caret.GetTransparentProxy()
$view.Values['IsClosed'] = $false
$manager = [Runtime.Serialization.FormatterServices]::GetUninitializedObject($managerType)
$controllersField = $managerType.GetField('lineIndicatorControllers', [Reflection.BindingFlags]'Instance,NonPublic')
$controllers = [Activator]::CreateInstance($controllersField.FieldType)
$controllersField.SetValue($manager, $controllers)
$controller = $controllerType.GetConstructors()[0].Invoke([object[]]@($manager, $view.GetTransparentProxy()))
$controllersField.FieldType.GetMethod('Add').Invoke($controllers, [object[]]@($view.GetTransparentProxy(), $controller))
function Apply($brush) { $controllerType.GetMethod('Apply').Invoke($controller, [object[]]@($hostProxy.GetTransparentProxy(), $brush)) }
function Pump {
    $frame = [Windows.Threading.DispatcherFrame]::new()
    [Windows.Threading.Dispatcher]::CurrentDispatcher.BeginInvoke([Windows.Threading.DispatcherPriority]::ApplicationIdle, [Action]{ $frame.Continue = $false }) | Out-Null
    [Windows.Threading.Dispatcher]::PushFrame($frame)
}
function CheckColor($indicator, $brush) {
    if ($indicator.Background -ne $brush) { throw "Expected $brush, got $($indicator.Background)" }
}
if ((Apply ([Windows.Media.Brushes]::Red))) { throw 'Off-screen caret should have no indicator yet' }
Pump
# Let the existing three-second retry window expire before realizing the indicator.
Start-Sleep -Milliseconds 3100
$view.Raise('LayoutChanged')
$indicator = [Indicator]::new()
$indicator.Background = [Windows.Media.Brushes]::Gray
$root.Children.Add($indicator) | Out-Null
Pump
CheckColor $indicator ([Windows.Media.Brushes]::Red)
$root.Children.Clear()
$marginRoot.Children.Add($indicator) | Out-Null
$view.Raise('LayoutChanged')
Pump
CheckColor $indicator ([Windows.Media.Brushes]::Red)
# Caret movement can replace the visual without changing the visible text layout.
$marginRoot.Children.Clear()
$indicator = [Indicator]::new()
$indicator.Background = [Windows.Media.Brushes]::Gray
$marginRoot.Children.Add($indicator) | Out-Null
$caret.Raise('PositionChanged')
Pump
CheckColor $indicator ([Windows.Media.Brushes]::Red)
(Apply ([Windows.Media.Brushes]::Blue)) | Out-Null
Pump
CheckColor $indicator ([Windows.Media.Brushes]::Blue)
if ($view.SubscriberCount('LayoutChanged') -ne 1 -or $caret.SubscriberCount('PositionChanged') -ne 1) { throw 'Repeated Apply duplicated subscriptions' }
# Disabling while a refresh is queued must restore the original and prevent repainting.
$view.Raise('LayoutChanged')
$controller.Restore()
Pump
CheckColor $indicator ([Windows.Media.Brushes]::Gray)
if ($view.SubscriberCount('LayoutChanged') -ne 0 -or $caret.SubscriberCount('PositionChanged') -ne 0) { throw 'Restore retained subscriptions' }
(Apply ([Windows.Media.Brushes]::Red)) | Out-Null
$view.Values['IsClosed'] = $true
$view.Raise('Closed')
$indicator.Background = [Windows.Media.Brushes]::Gray
Pump
CheckColor $indicator ([Windows.Media.Brushes]::Gray)
if ($controllers.Count -ne 0 -or $view.SubscriberCount('Closed') -ne 0) { throw 'Close retained controller or subscriptions' }
Write-Output 'Deferred line indicator passed: late realization, exact margin and editor fallback, caret replacement, color changes, disable and close with queued work.'
