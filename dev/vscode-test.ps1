param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Debug"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"

$repoRoot = Split-Path -Parent $PSScriptRoot
$testProject = Join-Path $repoRoot "tests\SSMS.EnvTabs.Tests\SSMS.EnvTabs.Tests.csproj"

if (-not (Test-Path $testProject)) {
    throw "Test project not found: $testProject"
}

$dotnet = Get-Command dotnet.exe -ErrorAction SilentlyContinue
if ($null -eq $dotnet) {
    throw "dotnet.exe not found. Install .NET SDK (for test execution in VS Code)."
}

Write-Host "Using dotnet: $($dotnet.Source)"
Write-Host "Running tests: $testProject"
Write-Host "Configuration=$Configuration"

& $dotnet.Source test $testProject --configuration $Configuration --nologo --verbosity minimal
if ($LASTEXITCODE -ne 0) {
    throw "Tests failed with exit code $LASTEXITCODE"
}

Write-Host "All tests passed."
