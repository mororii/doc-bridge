<#
.SYNOPSIS
  오래된 DocBridge 배포 압축 해제 폴더만 안전하게 정리한다.

.DESCRIPTION
  releases 아래의 DocBridge-*-win-x64 디렉터리만 대상으로 하며 ZIP과 SHA256 파일은
  삭제하지 않는다. 기본 실행은 미리보기이고 -Apply를 지정해야 실제로 제거한다.

.EXAMPLE
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Clear-DocBridgeReleaseCache.ps1
  powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\Clear-DocBridgeReleaseCache.ps1 -KeepLatest 2 -Apply
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [ValidateRange(0, 100)]
    [int]$KeepLatest = 2,
    [switch]$Apply
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$releaseRoot = [System.IO.Path]::GetFullPath((Join-Path $repo 'releases')).TrimEnd('\')
if (-not (Test-Path -LiteralPath $releaseRoot -PathType Container)) {
    Write-Host "Release cache does not exist: $releaseRoot"
    exit 0
}

function Assert-ReleaseCacheDirectory([string]$Path) {
    $resolved = [System.IO.Path]::GetFullPath($Path).TrimEnd('\')
    if (-not $resolved.StartsWith($releaseRoot + '\', [System.StringComparison]::OrdinalIgnoreCase)) {
        throw "Refusing path outside release cache: $resolved"
    }
    if ([System.IO.Path]::GetFileName($resolved) -notmatch '^DocBridge-.+-win-x64$') {
        throw "Refusing unexpected release directory: $resolved"
    }
    return $resolved
}

$expanded = @(Get-ChildItem -LiteralPath $releaseRoot -Directory |
    Where-Object { $_.Name -match '^DocBridge-.+-win-x64$' } |
    Sort-Object LastWriteTimeUtc -Descending)
$targets = @($expanded | Select-Object -Skip $KeepLatest)

Write-Host "Expanded release folders: $($expanded.Count); keep latest: $KeepLatest; candidates: $($targets.Count)"
foreach ($target in $targets) {
    $validated = Assert-ReleaseCacheDirectory $target.FullName
    $size = (Get-ChildItem -LiteralPath $validated -Recurse -File -ErrorAction SilentlyContinue |
        Measure-Object -Property Length -Sum).Sum
    Write-Host ("  {0}  {1:N1} MB" -f $target.Name, ($size / 1MB))
    if ($Apply -and $PSCmdlet.ShouldProcess($validated, 'Remove expanded release cache directory')) {
        Remove-Item -LiteralPath $validated -Recurse -Force
    }
}

if (-not $Apply -and $targets.Count -gt 0) {
    Write-Host 'Preview only. Re-run with -Apply after checking the exact list above.' -ForegroundColor Yellow
}
