<#
.SYNOPSIS
  Builds a self-contained offline DocBridge installer for Windows x64.
#>
[CmdletBinding()]
param(
    [switch]$SkipTests,
    [switch]$KeepExpanded
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$outputRoot = Join-Path $repo 'releases'
$manifestPath = Join-Path $repo '.codex-plugin\plugin.json'
$manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
$baseVersion = ([string]$manifest.version).Split('+')[0]
$releaseName = "DocBridge-$baseVersion-win-x64"
$stage = Join-Path $outputRoot $releaseName
$zip = Join-Path $outputRoot ($releaseName + '.zip')
$zipHashFile = $zip + '.sha256'

& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-DocBridgeVersion.ps1')
if ($LASTEXITCODE -ne 0) { throw "Version consistency check failed (exit $LASTEXITCODE)." }

function Assert-ReleasePath([string]$Path) {
    $resolved = [System.IO.Path]::GetFullPath($Path).TrimEnd('\')
    $root = [System.IO.Path]::GetFullPath($outputRoot).TrimEnd('\')
    if (-not $resolved.StartsWith($root + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing unsafe release path: $resolved"
    }
    return $resolved
}

Write-Host 'Publishing self-contained win-x64 binaries...' -ForegroundColor Cyan
$publishArgs = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File', (Join-Path $PSScriptRoot 'publish.ps1'), '-SelfContained')
if ($SkipTests) { $publishArgs += '-SkipTests' }
& powershell.exe @publishArgs
if ($LASTEXITCODE -ne 0) { throw "Self-contained publish failed (exit $LASTEXITCODE)." }

$dist = Join-Path $repo 'dist'
foreach ($required in @('doc-bridge-mcp.exe', 'doc-bridge-cli.exe', 'doc-bridge-hwp-worker.exe', 'coreclr.dll', 'hwp-security\FilePathCheckerModuleExample.dll', 'install-hwp-security.ps1', 'uninstall-hwp-security.ps1')) {
    $path = Join-Path $dist $required
    if (-not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "Publish output is incomplete: $path" }
}
$versionStartInfo = New-Object System.Diagnostics.ProcessStartInfo
$versionStartInfo.FileName = Join-Path $dist 'doc-bridge-mcp.exe'
$versionStartInfo.Arguments = '--version'
$versionStartInfo.UseShellExecute = $false
$versionStartInfo.CreateNoWindow = $true
$versionStartInfo.RedirectStandardOutput = $true
$versionStartInfo.RedirectStandardError = $true
$versionProcess = [System.Diagnostics.Process]::Start($versionStartInfo)
$versionStdOut = $versionProcess.StandardOutput.ReadToEnd()
$versionStdErr = $versionProcess.StandardError.ReadToEnd()
$versionProcess.WaitForExit()
$runtimeVersion = ($versionStdOut + $versionStdErr).Trim()
if ($versionProcess.ExitCode -ne 0 -or $runtimeVersion -ne $baseVersion) {
    throw "Release version mismatch: plugin=$baseVersion executable=$runtimeVersion"
}

New-Item -ItemType Directory -Path $outputRoot -Force | Out-Null
if (Test-Path -LiteralPath $stage) {
    Remove-Item -LiteralPath (Assert-ReleasePath $stage) -Recurse -Force
}
if (Test-Path -LiteralPath $zip) {
    Remove-Item -LiteralPath (Assert-ReleasePath $zip) -Force
}
if (Test-Path -LiteralPath $zipHashFile) {
    Remove-Item -LiteralPath (Assert-ReleasePath $zipHashFile) -Force
}

$pluginRoot = Join-Path $stage 'payload\codex-marketplace\plugins\doc-bridge'
$marketplaceDir = Join-Path $stage 'payload\codex-marketplace\.agents\plugins'
$supportDir = Join-Path $stage 'support'
New-Item -ItemType Directory -Path $pluginRoot -Force | Out-Null
New-Item -ItemType Directory -Path $marketplaceDir -Force | Out-Null
New-Item -ItemType Directory -Path $supportDir -Force | Out-Null

foreach ($directory in @('.codex-plugin', 'skills', 'dist', 'docs', 'examples', 'clients')) {
    Copy-Item -LiteralPath (Join-Path $repo $directory) -Destination $pluginRoot -Recurse -Force
}
foreach ($file in @('README.md', 'INSTALL.md')) {
    Copy-Item -LiteralPath (Join-Path $repo $file) -Destination $pluginRoot -Force
}
# publish.ps1 emits developer-machine absolute paths under dist\clients.  Those
# files are useful only in the source checkout and must never be shipped to a
# different PC.  Install-DocBridge.ps1 regenerates equivalent files under
# generated-configs with the target PC's actual installation path.
$packagedDistClients = Join-Path $pluginRoot 'dist\clients'
if (Test-Path -LiteralPath $packagedDistClients -PathType Container) {
    Remove-Item -LiteralPath $packagedDistClients -Recurse -Force
}

$marketplace = [ordered]@{
    name = 'docbridge-local'
    interface = [ordered]@{ displayName = 'DocBridge Local' }
    plugins = @(
        [ordered]@{
            name = 'doc-bridge'
            source = [ordered]@{ source = 'local'; path = './plugins/doc-bridge' }
            # Current Codex CLI registers local marketplaces but no longer exposes
            # top-level `plugin add/list`. Install the bundled plugin when the
            # marketplace is registered so a full Codex restart is sufficient.
            policy = [ordered]@{ installation = 'INSTALLED_BY_DEFAULT'; authentication = 'ON_INSTALL' }
            category = 'Productivity'
        }
    )
} | ConvertTo-Json -Depth 10
[System.IO.File]::WriteAllText((Join-Path $marketplaceDir 'marketplace.json'), $marketplace, (New-Object System.Text.UTF8Encoding($false)))

foreach ($file in @('Install-DocBridge.ps1', 'Test-DocBridge.ps1', 'Uninstall-DocBridge.ps1')) {
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot $file) -Destination $stage -Force
}
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'deployment\DocBridge.Deployment.psm1') -Destination $supportDir -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'verify-mcp.ps1') -Destination $supportDir -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'deployment\Test-PackageIntegrity.ps1') -Destination $supportDir -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'Convert-DocBridgeTextEncoding.ps1') -Destination $supportDir -Force
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'deployment\README-INSTALL.md') -Destination (Join-Path $stage 'README-INSTALL.md') -Force
$guideCandidates = @(Get-ChildItem -LiteralPath (Join-Path $PSScriptRoot 'deployment') -Filter '*.html' -File)
if ($guideCandidates.Count -ne 1) {
    throw "Expected exactly one deployment HTML guide, found $($guideCandidates.Count)."
}
$guideFile = $guideCandidates[0].Name
foreach ($file in @($guideFile, 'START-HERE.txt', '0-VERIFY.cmd', '1-INSTALL.cmd', '2-TEST.cmd', '2-EXCEL-LIVE-TEST.cmd', '3-UNINSTALL.cmd')) {
    Copy-Item -LiteralPath (Join-Path $PSScriptRoot ('deployment\' + $file)) -Destination $stage -Force
}

$buildInfo = [ordered]@{
    product = 'DocBridge'
    pluginVersion = [string]$manifest.version
    packageVersion = $baseVersion
    runtime = 'win-x64 self-contained'
    beginnerGuide = $guideFile
    builtAt = (Get-Date).ToString('o')
} | ConvertTo-Json -Depth 6
[System.IO.File]::WriteAllText((Join-Path $stage 'build-info.json'), $buildInfo, (New-Object System.Text.UTF8Encoding($false)))

# Windows PowerShell 5.1 guesses non-BOM UTF-8 using the legacy ANSI code page.
# Normalize every shipped PowerShell source before checksums are calculated.
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Convert-DocBridgeTextEncoding.ps1') `
    -Mode NormalizePowerShell -Path $stage
if ($LASTEXITCODE -ne 0) { throw "PowerShell encoding normalization failed (exit $LASTEXITCODE)." }

$hashLines = Get-ChildItem -LiteralPath $stage -Recurse -File | Sort-Object FullName | ForEach-Object {
    $relative = $_.FullName.Substring($stage.Length + 1).Replace('\', '/')
    $hash = (Get-FileHash -LiteralPath $_.FullName -Algorithm SHA256).Hash.ToLowerInvariant()
    "$hash  $relative"
}
[System.IO.File]::WriteAllLines((Join-Path $stage 'SHA256SUMS.txt'), [string[]]$hashLines, (New-Object System.Text.UTF8Encoding($false)))

Write-Host 'Verifying exact package manifest coverage...' -ForegroundColor Cyan
& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $stage 'support\Test-PackageIntegrity.ps1') -PackageRoot $stage
if ($LASTEXITCODE -ne 0) { throw "Package integrity verification failed (exit $LASTEXITCODE)." }

Write-Host "Creating archive: $zip" -ForegroundColor Cyan
Compress-Archive -LiteralPath $stage -DestinationPath $zip -CompressionLevel Optimal
$zipHash = (Get-FileHash -LiteralPath $zip -Algorithm SHA256).Hash.ToLowerInvariant()
$zipHashLine = "$zipHash  $(Split-Path -Leaf $zip)`r`n"
[System.IO.File]::WriteAllText($zipHashFile, $zipHashLine, (New-Object System.Text.UTF8Encoding($false)))

if (-not $KeepExpanded) {
    Remove-Item -LiteralPath (Assert-ReleasePath $stage) -Recurse -Force
}

Write-Host 'Release package completed.' -ForegroundColor Green
if ($KeepExpanded) {
    Write-Host "Folder: $stage"
} else {
    Write-Host 'Expanded folder: removed after archiving (use -KeepExpanded for installer development tests)'
}
Write-Host "ZIP:    $zip"
Write-Host "SHA256: $zipHash"
Write-Host "Hash:   $zipHashFile"
