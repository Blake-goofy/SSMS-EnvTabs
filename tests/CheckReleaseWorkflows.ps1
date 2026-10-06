# Run after a Release build: pwsh -NoProfile -File tests/CheckReleaseWorkflows.ps1
$ErrorActionPreference = 'Stop'
$repoRoot = Split-Path $PSScriptRoot
foreach ($file in @('pre-release.yml', 'publish-gallery.yml')) {
    $yaml = Get-Content (Join-Path $repoRoot ".github/workflows/$file") -Raw
    $events = [regex]::Match($yaml, '(?ms)^on:\r?\n(?<events>.*?)(?=^\S)').Groups['events'].Value
    if (([regex]::Matches($events, '(?m)^  \S[^\r\n]*').Value -join "`n") -cne '  workflow_dispatch:') { throw "$file must only run manually." }
}
if ((Get-Content (Join-Path $repoRoot '.github/workflows/build.yml') -Raw) -match 'publish-vsixgallery') { throw 'Build must not publish to the gallery.' }
function ReadWorkflowScript($file) {
    $yaml = Get-Content (Join-Path $repoRoot ".github/workflows/$file") -Raw
    $match = [regex]::Match($yaml, '(?m)^        run: \|\r?\n(?<script>(?:          .*\r?\n|\r?\n)+)')
    if (!$match.Success) { throw "Workflow script not found: $file" }
    [scriptblock]::Create([regex]::Replace($match.Groups['script'].Value, '(?m)^          ', ''))
}
$publish = ReadWorkflowScript 'publish-gallery.yml'
$preRelease = ReadWorkflowScript 'pre-release.yml'
[xml]$manifest = Get-Content (Join-Path $repoRoot 'SSMS EnvTabs/source.extension.vsixmanifest')
$version = $manifest.PackageManifest.Metadata.Identity.Version
$assetName = "SSMS_EnvTabs_$version.vsix"
$package = Join-Path $repoRoot 'SSMS EnvTabs/bin/Release/SSMS EnvTabs.vsix'
if (!(Test-Path -LiteralPath $package)) { throw 'Build Release first.' }
$scratch = Join-Path ([IO.Path]::GetTempPath()) ([IO.Path]::GetRandomFileName())
New-Item -ItemType Directory -Path $scratch | Out-Null
$oldOutput = $env:GITHUB_OUTPUT
$oldTag = $env:RELEASE_TAG
$oldVersion = $env:VERSION
$oldAsset = $env:ASSET_NAME
$oldCommit = $env:RELEASE_COMMIT
function gh {
    $global:LASTEXITCODE = 0
    if ($args[1] -eq 'view') { return ($script:release | ConvertTo-Json -Depth 5 -Compress) }
    if ($args[1] -eq 'download') {
        New-Item -ItemType Directory -Path artifacts -Force | Out-Null
        Copy-Item -LiteralPath $package -Destination "artifacts/$($script:release.assets[0].name)" -Force
    } elseif ($args[1] -eq 'create') { $script:createArgs = $args }
    else { throw "Unexpected gh command: $args" }
}
Push-Location $scratch
try {
    $env:GITHUB_OUTPUT = Join-Path $scratch 'outputs'
    $env:RELEASE_TAG = "v$version"
    $script:release = @{ isDraft = $false; isPrerelease = $false; assets = @(@{ name = $assetName }) }
    & $publish
    if ((Get-Content $env:GITHUB_OUTPUT) -ne "asset-name=$assetName") { throw 'Wrong publish filename.' }
    foreach ($case in @('draft', 'pre-release', 'missing-asset', 'wrong-version', 'invalid-tag')) {
        $script:release.isDraft = $case -eq 'draft'
        $script:release.isPrerelease = $case -eq 'pre-release'
        $env:RELEASE_TAG = if ($case -eq 'wrong-version') { 'v99.99.99' } elseif ($case -eq 'invalid-tag') { '--latest' } else { "v$version" }
        $script:release.assets[0].name = if ($case -eq 'missing-asset') { 'missing.vsix' } elseif ($case -eq 'wrong-version') { 'SSMS_EnvTabs_99.99.99.vsix' } else { $assetName }
        $rejected = $false
        try { & $publish } catch { $rejected = $true }
        if (!$rejected) { throw "Publisher accepted $case." }
    }
    $env:VERSION = $version
    $env:ASSET_NAME = $assetName
    $env:RELEASE_COMMIT = '0123456789012345678901234567890123456789'
    & $preRelease
    foreach ($expected in @("v$version", "artifacts/$assetName", $env:RELEASE_COMMIT, '--prerelease', '--latest=false', '--generate-notes')) {
        if ($expected -notin $script:createArgs) { throw "Pre-release argument missing: $expected" }
    }
    Write-Output 'Release workflow checks passed: versioned asset, pre-release creation, and publishing rejection checks.'
} finally {
    Pop-Location
    $env:GITHUB_OUTPUT = $oldOutput
    $env:RELEASE_TAG = $oldTag
    $env:VERSION = $oldVersion
    $env:ASSET_NAME = $oldAsset
    $env:RELEASE_COMMIT = $oldCommit
    $resolvedScratch = [IO.Path]::GetFullPath($scratch)
    $tempRoot = [IO.Path]::GetFullPath([IO.Path]::GetTempPath()).TrimEnd('\') + '\'
    if (!$resolvedScratch.StartsWith($tempRoot, [StringComparison]::OrdinalIgnoreCase)) { throw 'Cleanup path is outside TEMP.' }
    Remove-Item -LiteralPath $resolvedScratch -Recurse -Force
}
