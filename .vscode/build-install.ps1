Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$solutionPath = Join-Path $repoRoot "SSMS EnvTabs.sln"
$installer = "C:\Program Files\Microsoft SQL Server Management Studio 22\Release\Common7\IDE\VSIXInstaller.exe"
$vsix = Join-Path $repoRoot "SSMS EnvTabs\bin\Release\SSMS EnvTabs.vsix"
$vsixLogDir = Join-Path $env:TEMP "SSMS_EnvTabs_logs"
$installLog = Join-Path $vsixLogDir "vsix-install.log"
$uninstallLog = Join-Path $vsixLogDir "vsix-uninstall.log"
$extensionId = "SSMS_EnvTabs"

function Find-MSBuild {
    $direct = Get-Command msbuild.exe -ErrorAction SilentlyContinue
    if ($null -ne $direct) {
        return $direct.Source
    }

    $vswhere = Join-Path ${env:ProgramFiles(x86)} "Microsoft Visual Studio\Installer\vswhere.exe"
    if (Test-Path $vswhere) {
        $path = & $vswhere -latest -requires Microsoft.Component.MSBuild -find "MSBuild\**\Bin\MSBuild.exe" | Select-Object -First 1
        if (-not [string]::IsNullOrWhiteSpace($path) -and (Test-Path $path)) {
            return $path
        }
    }

    throw "MSBuild.exe not found. Install Visual Studio Build Tools or Visual Studio with MSBuild components."
}

function Write-LogTailIfPresent {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [Parameter(Mandatory = $true)]
        [string]$Title
    )

    if (Test-Path $Path) {
        Write-Host ("{0}: {1}" -f $Title, $Path)
        Get-Content -Path $Path -Tail 40 | ForEach-Object { Write-Host $_ }
    }
}

function Invoke-VsixInstaller {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$ArgumentList
    )

    & $installer @ArgumentList
    return $LASTEXITCODE
}

if (-not (Test-Path $solutionPath)) {
    throw "Solution file not found: $solutionPath"
}

if (-not (Test-Path $installer)) {
    throw "VSIXInstaller not found: $installer"
}

if (-not (Test-Path $vsixLogDir)) {
    New-Item -Path $vsixLogDir -ItemType Directory -Force | Out-Null
}

$msbuild = Find-MSBuild
Write-Host "Using MSBuild: $msbuild"
Write-Host "Building solution: $solutionPath"

& $msbuild $solutionPath /restore /m /p:Configuration=Release /p:Platform="Any CPU" /p:DeployExtension=false /verbosity:minimal
if ($LASTEXITCODE -ne 0) {
    throw "Build failed with exit code $LASTEXITCODE"
}

if (-not (Test-Path $vsix)) {
    throw "VSIX not found after build: $vsix"
}

Write-Host "Installing VSIX with /force..."
$installExitCode = Invoke-VsixInstaller -ArgumentList @("/quiet", "/force", "/logFile:$installLog", $vsix)
if ($installExitCode -eq 0) {
    Write-Host "Build and install completed."
    exit 0
}

Write-Host "Forced install failed with exit code $installExitCode. Falling back to uninstall/install."
Write-LogTailIfPresent -Path $installLog -Title "Install log"

$uninstallExitCode = Invoke-VsixInstaller -ArgumentList @("/quiet", "/uninstall:$extensionId", "/logFile:$uninstallLog")
if ($uninstallExitCode -ne 0 -and $uninstallExitCode -ne 1002) {
    Write-LogTailIfPresent -Path $uninstallLog -Title "Uninstall log"
    throw "Uninstall failed with exit code $uninstallExitCode"
}

$reinstallExitCode = Invoke-VsixInstaller -ArgumentList @("/quiet", "/logFile:$installLog", $vsix)
if ($reinstallExitCode -ne 0) {
    Write-LogTailIfPresent -Path $installLog -Title "Install log"
    throw "Install failed with exit code $reinstallExitCode"
}

Write-Host "Build and install completed."
