param(
    [ValidateSet("Debug", "Release")]
    [string]$Configuration = "Release"
)

Set-StrictMode -Version Latest
$ErrorActionPreference = "Stop"
$startTime = Get-Date
$finalStatus = "FAILED"

$repoRoot = Split-Path -Parent $PSScriptRoot
$buildScript = Join-Path $PSScriptRoot "vscode-build.ps1"
$testScript = Join-Path $PSScriptRoot "vscode-test.ps1"
$installer = "C:\Program Files\Microsoft SQL Server Management Studio 22\Release\Common7\IDE\VSIXInstaller.exe"
$vsix = Join-Path $repoRoot "SSMS EnvTabs\bin\$Configuration\SSMS EnvTabs.vsix"
$vsixLogDir = Join-Path $env:TEMP "SSMS_EnvTabs_logs"
$id = "SSMS_EnvTabs"
$ssms = "C:\ProgramData\Microsoft\Windows\Start Menu\Programs\SSMS.lnk"

if (-not (Test-Path $vsixLogDir)) {
    New-Item -Path $vsixLogDir -ItemType Directory -Force | Out-Null
}

function Show-SpinnerForProcesses {
    param(
        [Parameter(Mandatory = $true)]
        [System.Diagnostics.Process[]]$Processes,
        [Parameter(Mandatory = $true)]
        [string]$Label,
        [Parameter(Mandatory = $true)]
        [string]$DoneLabel,
        [Parameter(Mandatory = $true)]
        [string]$FailedLabel,
        [int[]]$IgnoredExitCodes = @()
    )

    $spinner = '|/-\'
    $i = 0
    [Console]::CursorVisible = $false
    Write-Host $Label
    $line = [Console]::CursorTop - 1

    try {
        while ($Processes | Where-Object { -not $_.HasExited }) {
            $ch = $spinner[$i++ % 4]
            [Console]::SetCursorPosition(0, $line)
            Write-Host -NoNewline ("$Label $ch   ")
            [Console]::SetCursorPosition(0, $line + 1)
            Start-Sleep -Milliseconds 250
        }

        $failed = $Processes | Where-Object {
            ($_.ExitCode -ne 0) -and ($IgnoredExitCodes -notcontains $_.ExitCode)
        }
        if ($failed) {
            [Console]::SetCursorPosition(0, $line)
            Write-Host "$FailedLabel   "
            [Console]::SetCursorPosition(0, $line + 1)
            $codes = ($failed | ForEach-Object { $_.ExitCode }) -join ", "
            throw "$FailedLabel Exit code(s): $codes"
        }

        $ignored = $Processes | Where-Object {
            ($_.ExitCode -ne 0) -and ($IgnoredExitCodes -contains $_.ExitCode)
        }
        [Console]::SetCursorPosition(0, $line)
        Write-Host "$DoneLabel   "
        [Console]::SetCursorPosition(0, $line + 1)

        if ($ignored) {
            $ignoredCodes = ($ignored | ForEach-Object { $_.ExitCode } | Select-Object -Unique) -join ", "
            Write-Host "Note: Non-fatal exit code(s): $ignoredCodes"
        }
    }
    finally {
        [Console]::CursorVisible = $true
    }
}

function Write-LogTailIfPresent {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,
        [string]$Title = "VSIXInstaller log"
    )

    if (Test-Path $Path) {
        Write-Host ("{0}: {1}" -f $Title, $Path)
        Write-Host "--- log tail ---"
        Get-Content -Path $Path -Tail 40 | ForEach-Object { Write-Host $_ }
        Write-Host "--- end log tail ---"
    }
}

function Start-ProcessStepWithSpinner {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilePath,
        [string[]]$ArgumentList,
        [Parameter(Mandatory = $true)]
        [string]$StepName
    )

    $proc = Start-Process -FilePath $FilePath -ArgumentList $ArgumentList -PassThru
    Show-SpinnerForProcesses `
        -Processes @($proc) `
        -Label "$StepName..." `
        -DoneLabel "$StepName... done!" `
        -FailedLabel "$StepName... failed!"
}

function Invoke-ScriptStep {
    param(
        [Parameter(Mandatory = $true)]
        [string]$StepName,
        [Parameter(Mandatory = $true)]
        [scriptblock]$Action
    )

    Write-Host "$StepName..."
    try {
        & $Action
        Write-Host "$StepName... done!"
    }
    catch {
        Write-Host "$StepName... failed!"
        throw
    }
}

if (-not (Test-Path $buildScript)) {
    throw "Build script not found: $buildScript"
}

if (-not (Test-Path $testScript)) {
    throw "Test script not found: $testScript"
}

if (-not (Test-Path $installer)) {
    throw "VSIXInstaller not found: $installer"
}

Write-Host "Starting pipeline with Configuration=$Configuration"

# 1) Build first.
Invoke-ScriptStep -StepName "Building solution" -Action {
    & $buildScript -Configuration $Configuration -Platform "Any CPU"
}

# 2) Run tests and stop if anything fails.
Invoke-ScriptStep -StepName "Running tests" -Action {
    & $testScript -Configuration $Configuration
}

if (-not (Test-Path $vsix)) {
    throw "VSIX not found after build: $vsix"
}

# 3) Uninstall currently installed extension versions.
$uninstallCurrentLog = Join-Path $vsixLogDir "vsix-uninstall-current.log"
$uninstallCurrent = Start-Process -FilePath $installer -ArgumentList @("/quiet", "/uninstall:$id", "/logFile:$uninstallCurrentLog") -PassThru
Show-SpinnerForProcesses `
    -Processes @($uninstallCurrent) `
    -Label "Uninstalling..." `
    -DoneLabel "Uninstalling... done!" `
    -FailedLabel "Uninstalling... failed!" `
    -IgnoredExitCodes @(1002)

# 4) Install the newly built VSIX.
$installLog = Join-Path $vsixLogDir "vsix-install.log"
try {
    Start-ProcessStepWithSpinner `
        -FilePath $installer `
    -ArgumentList @("/quiet", "/logFile:$installLog", "`"$vsix`"") `
        -StepName "Installing"
}
catch {
    Write-LogTailIfPresent -Path $installLog -Title "Install log"
    throw
}

# 5) Launch SSMS.
Start-Process -FilePath $ssms
Write-Host "SSMS launched."
$finalStatus = "SUCCESS"

$elapsed = (Get-Date) - $startTime
Write-Host ("Final status: {0}" -f $finalStatus)
$elapsedText = "{0:00}:{1:00}:{2:00}" -f [int]$elapsed.TotalHours, $elapsed.Minutes, $elapsed.Seconds
Write-Host ("Elapsed: {0}" -f $elapsedText)
Read-Host "Press Enter to exit"
