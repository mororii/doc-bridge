# doc-bridge-cli.ps1 — CLI 래퍼 (MCP를 못 붙이는 환경용 fallback, 명령서 §8.3)
# 사용법:
#   powershell.exe -ExecutionPolicy Bypass -File clients/kimi/doc-bridge-cli.ps1 core_ping
#   powershell.exe -ExecutionPolicy Bypass -File clients/kimi/doc-bridge-cli.ps1 excel_get_active_context
#   powershell.exe -ExecutionPolicy Bypass -File clients/kimi/doc-bridge-cli.ps1 hwp_apply_ops --ops ops.json --dry-run
#   powershell.exe -ExecutionPolicy Bypass -File clients/kimi/doc-bridge-cli.ps1 hwp_apply_ops --ops ops.json --confirm-token conf_...
#
# 우선순위: dist\ (publish.ps1 결과) → Release 빌드 → Debug 빌드 → 없으면 즉석 빌드.
# 진단 메시지는 stderr 로만 보낸다 — stdout 은 결과 JSON 전용이어야 하기 때문.

$ErrorActionPreference = "Stop"
$repo = Split-Path -Parent (Split-Path -Parent (Split-Path -Parent $MyInvocation.MyCommand.Path))

$candidates = @(
    (Join-Path $repo "dist\doc-bridge-cli.exe")
    (Join-Path $repo "src\DocBridge.Cli\bin\Release\net8.0-windows\doc-bridge-cli.exe")
    (Join-Path $repo "src\DocBridge.Cli\bin\Debug\net8.0-windows\doc-bridge-cli.exe")
)
$exe = $candidates | Where-Object { Test-Path $_ } | Select-Object -First 1

if (-not $exe) {
    $exe = $candidates[2]
    [Console]::Error.WriteLine("[doc-bridge-cli] exe가 없어 빌드합니다: $exe")
    dotnet build (Join-Path $repo "src\DocBridge.Cli") -v q -nodeReuse:false 2>&1 | ForEach-Object { [Console]::Error.WriteLine($_) }
    if (-not (Test-Path $exe)) { [Console]::Error.WriteLine("[doc-bridge-cli] 빌드 실패"); exit 2 }
}

& $exe @args
exit $LASTEXITCODE

