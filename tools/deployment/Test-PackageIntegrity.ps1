[CmdletBinding()]
param([string]$PackageRoot)

$ErrorActionPreference = 'Stop'
if ([string]::IsNullOrWhiteSpace($PackageRoot)) {
    $PackageRoot = Split-Path -Parent $PSScriptRoot
}
$PackageRoot = [System.IO.Path]::GetFullPath($PackageRoot).TrimEnd('\')
$checksumFile = Join-Path $PackageRoot 'SHA256SUMS.txt'
if (-not (Test-Path -LiteralPath $checksumFile -PathType Leaf)) {
    throw "Checksum list was not found: $checksumFile"
}

$checked = 0
$failures = New-Object System.Collections.Generic.List[string]
$manifestEntries = New-Object 'System.Collections.Generic.Dictionary[string,string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($line in Get-Content -LiteralPath $checksumFile -Encoding UTF8) {
    if ([string]::IsNullOrWhiteSpace($line)) { continue }
    $parts = $line -split '  ', 2
    if ($parts.Count -ne 2 -or $parts[0] -notmatch '^[0-9a-fA-F]{64}$') {
        [void]$failures.Add("Invalid checksum entry: $line")
        continue
    }
    $relative = $parts[1].Replace('/', '\')
    if ($relative -eq 'SHA256SUMS.txt') {
        [void]$failures.Add('SHA256SUMS.txt must not contain a self-referential checksum entry.')
        continue
    }
    if ($manifestEntries.ContainsKey($relative)) {
        [void]$failures.Add("Duplicate checksum path: $relative")
        continue
    }
    $manifestEntries.Add($relative, $parts[0])
    $candidate = [System.IO.Path]::GetFullPath((Join-Path $PackageRoot $relative))
    if (-not $candidate.StartsWith($PackageRoot + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
        [void]$failures.Add("Unsafe checksum path: $relative")
        continue
    }
    if (-not (Test-Path -LiteralPath $candidate -PathType Leaf)) {
        [void]$failures.Add("Missing file: $relative")
        continue
    }
    $actual = (Get-FileHash -LiteralPath $candidate -Algorithm SHA256).Hash
    if (-not $actual.Equals($parts[0], [System.StringComparison]::OrdinalIgnoreCase)) {
        [void]$failures.Add("Hash mismatch: $relative")
        continue
    }
    $checked++
}

$actualFiles = @(Get-ChildItem -LiteralPath $PackageRoot -Recurse -File | Where-Object {
    -not $_.FullName.Equals($checksumFile, [System.StringComparison]::OrdinalIgnoreCase)
})
$actualRelativePaths = New-Object 'System.Collections.Generic.HashSet[string]' ([System.StringComparer]::OrdinalIgnoreCase)
foreach ($file in $actualFiles) {
    $relative = $file.FullName.Substring($PackageRoot.Length + 1)
    [void]$actualRelativePaths.Add($relative)
    if (-not $manifestEntries.ContainsKey($relative)) {
        [void]$failures.Add("File is not covered by SHA256SUMS.txt: $relative")
    }
}
foreach ($relative in $manifestEntries.Keys) {
    if (-not $actualRelativePaths.Contains($relative)) {
        [void]$failures.Add("Checksum entry has no package file: $relative")
    }
}

$powerShellFiles = @(Get-ChildItem -LiteralPath $PackageRoot -Recurse -File | Where-Object {
    $_.Extension -in @('.ps1', '.psm1', '.psd1')
})
foreach ($file in $powerShellFiles) {
    $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
    if ($bytes.Length -lt 3 -or $bytes[0] -ne 0xEF -or $bytes[1] -ne 0xBB -or $bytes[2] -ne 0xBF) {
        $relative = $file.FullName.Substring($PackageRoot.Length + 1)
        [void]$failures.Add("PowerShell file is not UTF-8 BOM: $relative")
    }
}

$requiredCursorFiles = @(
    'payload\codex-marketplace\plugins\doc-bridge\clients\cursor\mcp.example.json',
    'payload\codex-marketplace\plugins\doc-bridge\clients\cursor\CURSOR_USAGE.md',
    'payload\codex-marketplace\plugins\doc-bridge\clients\cursor\docbridge-user-rule.txt',
    'payload\codex-marketplace\plugins\doc-bridge\clients\cursor\rules\docbridge-safe-automation.mdc'
)
foreach ($relative in $requiredCursorFiles) {
    if (-not (Test-Path -LiteralPath (Join-Path $PackageRoot $relative) -PathType Leaf)) {
        [void]$failures.Add("Missing Cursor integration file: $relative")
    }
}

$requiredDeploymentFiles = @(
    '0-VERIFY.cmd',
    '1-INSTALL.cmd',
    '2-TEST.cmd',
    '2-EXCEL-LIVE-TEST.cmd',
    '3-UNINSTALL.cmd',
    'Install-DocBridge.ps1',
    'Test-DocBridge.ps1',
    'Uninstall-DocBridge.ps1',
    'support\verify-mcp.ps1'
)
foreach ($relative in $requiredDeploymentFiles) {
    if (-not (Test-Path -LiteralPath (Join-Path $PackageRoot $relative) -PathType Leaf)) {
        [void]$failures.Add("Missing deployment file: $relative")
    }
}

Write-Host "Checked files: $checked"
Write-Host "Manifest coverage: $($manifestEntries.Count) checksums / $($actualFiles.Count) package files"
Write-Host "Checked PowerShell UTF-8 BOM files: $($powerShellFiles.Count)"
if ($failures.Count -gt 0) {
    $failures | ForEach-Object { Write-Host "[FAIL] $_" -ForegroundColor Red }
    exit 1
}
Write-Host 'All package checksums match.' -ForegroundColor Green
exit 0
