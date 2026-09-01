<#
.SYNOPSIS
  Build version, plugin manifest and user-facing release names must agree.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$propsPath = Join-Path $repo 'Directory.Build.props'
$manifestPath = Join-Path $repo '.codex-plugin\plugin.json'
$hostPath = Join-Path $repo 'src\DocBridge.Core\Services\DocBridgeHost.cs'
$readmePath = Join-Path $repo 'README.md'
$installPath = Join-Path $repo 'INSTALL.md'
$guideCandidates = @(Get-ChildItem -LiteralPath (Join-Path $repo 'tools\deployment') -Filter '*.html' -File)
if ($guideCandidates.Count -ne 1) { throw "Expected exactly one deployment HTML guide, found $($guideCandidates.Count)." }
$guidePath = $guideCandidates[0].FullName

if (-not (Test-Path -LiteralPath $propsPath -PathType Leaf)) {
    throw "Missing version source: $propsPath"
}
[xml]$props = Get-Content -LiteralPath $propsPath -Raw -Encoding UTF8
$version = [string]$props.Project.PropertyGroup.VersionPrefix
if ([string]::IsNullOrWhiteSpace($version)) { throw 'Directory.Build.props VersionPrefix is empty.' }

$manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$pluginBase = ([string]$manifest.version).Split('+')[0]
$readmeText = Get-Content -LiteralPath $readmePath -Raw -Encoding UTF8
$installText = Get-Content -LiteralPath $installPath -Raw -Encoding UTF8
$guideText = Get-Content -LiteralPath $guidePath -Raw -Encoding UTF8
$versionCode = '`' + $version + '`'
$checks = [ordered]@{
    'plugin manifest base version' = ($pluginBase -eq $version)
    'runtime host version' = ((Get-Content -LiteralPath $hostPath -Raw -Encoding UTF8) -match
        ('public const string Version = "' + [regex]::Escape($version) + '";'))
    'README heading' = ($readmeText -match
        ('(?m)^# DocBridge ' + [regex]::Escape($version) + '$'))
    'INSTALL archive name' = ($installText -match
        ([regex]::Escape("DocBridge-$version-win-x64.zip")))
    'INSTALL expected runtime version' = ($installText -match
        ([regex]::Escape($versionCode) + '.*doc-bridge-hwp-worker\.exe'))
    'deployment HTML metadata version' = ($guideText -match
        ('<meta\s+name="docbridge-version"\s+content="' + [regex]::Escape($version) + '"\s*/?>'))
    'deployment HTML footer version' = ($guideText -match
        ('<footer>DocBridge ' + [regex]::Escape($version) + '\b'))
    'deployment HTML current ZIP guidance' = ($guideText -match
        ([regex]::Escape("$version ZIP") + '.*2-EXCEL-LIVE-TEST\.cmd'))
}

$failed = @($checks.GetEnumerator() | Where-Object { -not $_.Value })
foreach ($check in $checks.GetEnumerator()) {
    Write-Host ("[{0}] {1}" -f $(if ($check.Value) { 'OK' } else { 'FAIL' }), $check.Key) `
        -ForegroundColor $(if ($check.Value) { 'Green' } else { 'Red' })
}
if ($failed.Count -gt 0) {
    throw "DocBridge version mismatch. Expected $version in: $($failed.Name -join ', ')"
}

Write-Host "DocBridge version sources agree: $version" -ForegroundColor Green
