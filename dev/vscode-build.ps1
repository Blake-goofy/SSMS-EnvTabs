param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug",

    [ValidateSet("Any CPU", "x86", "arm64")]
    [string]$Platform = "Any CPU"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$solutionPath = Join-Path $repoRoot "SSMS EnvTabs.sln"

if (-not (Test-Path $solutionPath)) {
    throw "Solution file not found: $solutionPath"
}

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

$msbuild = Find-MSBuild
Write-Host "Using MSBuild: $msbuild"
Write-Host "Building solution: $solutionPath"
Write-Host "Configuration=$Configuration, Platform=$Platform"

& $msbuild $solutionPath /restore /m /p:Configuration=$Configuration /p:Platform="$Platform" /verbosity:minimal
if ($LASTEXITCODE -ne 0) {
    throw "Build failed with exit code $LASTEXITCODE"
}

Write-Host "Build completed successfully."
