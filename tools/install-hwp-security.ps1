param(
    [string]$ModulePath
)

$ErrorActionPreference = 'Stop'
$valueName = 'DocBridgeFilePathChecker'
$registryPath = 'HKCU\SOFTWARE\HNC\HwpAutomation\Modules'

if ([string]::IsNullOrWhiteSpace($ModulePath)) {
    $candidates = @(
        (Join-Path $PSScriptRoot 'hwp-security\FilePathCheckerModuleExample.dll'),
        (Join-Path (Split-Path $PSScriptRoot -Parent) 'assets\hwp-security\FilePathCheckerModuleExample.dll')
    )
    $ModulePath = $candidates | Where-Object { Test-Path -LiteralPath $_ } | Select-Object -First 1
}

if ([string]::IsNullOrWhiteSpace($ModulePath) -or -not (Test-Path -LiteralPath $ModulePath)) {
    throw '한컴 자동화 보안 모듈 FilePathCheckerModuleExample.dll을 찾을 수 없습니다.'
}

$resolved = (Resolve-Path -LiteralPath $ModulePath).Path
& reg.exe add $registryPath /v $valueName /t REG_SZ /d $resolved /f /reg:32 | Out-Null
if ($LASTEXITCODE -ne 0) { throw "한글 자동화 보안 모듈 레지스트리 등록에 실패했습니다: $registryPath" }

Write-Host "Registered: $registryPath\$valueName"
Write-Host "Module: $resolved"
