$ErrorActionPreference = 'Stop'
$valueName = 'DocBridgeFilePathChecker'
$registryPath = 'HKCU\SOFTWARE\HNC\HwpAutomation\Modules'

& reg.exe query $registryPath /v $valueName /reg:32 *> $null
if ($LASTEXITCODE -ne 0) {
    Write-Host "Not registered: $registryPath\$valueName"
    exit 0
}

& reg.exe delete $registryPath /v $valueName /f /reg:32 | Out-Null
if ($LASTEXITCODE -ne 0) { throw "한글 자동화 보안 모듈 등록 해제에 실패했습니다: $registryPath" }
Write-Host "Removed: $registryPath\$valueName"
