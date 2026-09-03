<#
.SYNOPSIS
  doc-bridge를 dist\ 에 발행하고, Claude / Codex / Kimi / Cursor 용 설정 파일을 실제 경로로 생성한다.

.DESCRIPTION
  MCP 클라이언트가 매번 `dotnet run` 을 하면 복원/빌드 때문에 기동이 느려
  Codex의 startup_timeout_sec 을 넘겨 연결이 끊길 수 있다.
  이 스크립트로 exe를 한 번 발행해 두고, 클라이언트는 그 exe만 직접 실행한다.

.PARAMETER SelfContained
  .NET 8 런타임이 없는 PC에서도 돌아가도록 런타임 포함(약 90MB)으로 발행한다.

.PARAMETER SkipTests
  발행 전 단위 테스트를 건너뛴다.

.EXAMPLE
  powershell.exe -ExecutionPolicy Bypass -File tools\publish.ps1
  powershell.exe -ExecutionPolicy Bypass -File tools\publish.ps1 -SelfContained
#>
[CmdletBinding()]
param(
    [switch]$SelfContained,
    [switch]$SkipTests
)

$ErrorActionPreference = 'Stop'
$repo = Split-Path -Parent $PSScriptRoot
$dist = Join-Path $repo 'dist'

function Step($msg) { Write-Host "`n=== $msg" -ForegroundColor Cyan }

# ---------- 0. dotnet 확인 ----------
Step 'dotnet SDK 확인'
$dotnet = Get-Command dotnet -ErrorAction SilentlyContinue
if (-not $dotnet) {
    throw "dotnet SDK를 찾을 수 없습니다. https://dotnet.microsoft.com/download/dotnet/8.0 에서 .NET 8 SDK를 설치하세요."
}
& dotnet --version | ForEach-Object { Write-Host "  dotnet $_" }

& powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $PSScriptRoot 'Test-DocBridgeVersion.ps1')
if ($LASTEXITCODE -ne 0) { throw "버전 일관성 검사 실패 — 발행을 중단합니다." }

# ---------- 1. 테스트 ----------
if (-not $SkipTests) {
    Step '단위 테스트 (실제 Office 없이 실행되는 것만)'
    Remove-Item Env:DOCBRIDGE_E2E -ErrorAction SilentlyContinue
    & dotnet test (Join-Path $repo 'DocBridge.sln') -c Release --nologo -v minimal `
        --filter 'FullyQualifiedName!~E2ETests&Category!=E2E'
    if ($LASTEXITCODE -ne 0) { throw "테스트 실패 — 발행을 중단합니다." }
}

# ---------- 2. 발행 ----------
Step "발행 → $dist"
if (Test-Path $dist) { Remove-Item $dist -Recurse -Force }
New-Item -ItemType Directory -Path $dist | Out-Null

# Public packages must not embed the builder's user/profile path in DLL CodeView
# records, PDB documents or CallerFilePath strings. Local test builds keep symbols.
$publishArgs = @('-c', 'Release', '-o', $dist, '--nologo',
    '-p:DebugType=None', '-p:DebugSymbols=false', "-p:PathMap=$repo=/_/docbridge")
if ($SelfContained) {
    $publishArgs += @('-r', 'win-x64', '--self-contained', 'true')
} else {
    $publishArgs += @('--self-contained', 'false')
}

& dotnet publish (Join-Path $repo 'src\DocBridge.Mcp') @publishArgs
if ($LASTEXITCODE -ne 0) { throw "DocBridge.Mcp 발행 실패" }

& dotnet publish (Join-Path $repo 'src\DocBridge.HwpWorker') @publishArgs
if ($LASTEXITCODE -ne 0) { throw "DocBridge.HwpWorker 발행 실패" }

& dotnet publish (Join-Path $repo 'src\DocBridge.Cli') @publishArgs
if ($LASTEXITCODE -ne 0) { throw "DocBridge.Cli 발행 실패" }

# ---------- 3. ops 리소스 동봉 ----------
# PolicyEngine / CadAdapter 는 exe 위치에서 위로 올라가며 ops\ 를 찾는다.
Step 'ops\ (정책·스키마·스크립트 템플릿) 복사'
Copy-Item (Join-Path $repo 'ops') (Join-Path $dist 'ops') -Recurse -Force
Copy-Item (Join-Path $repo 'tools\install-hwp-security.ps1') $dist -Force
Copy-Item (Join-Path $repo 'tools\uninstall-hwp-security.ps1') $dist -Force

$mcpExe = Join-Path $dist 'doc-bridge-mcp.exe'
$cliExe = Join-Path $dist 'doc-bridge-cli.exe'
$hwpWorkerExe = Join-Path $dist 'doc-bridge-hwp-worker.exe'
foreach ($exe in @($mcpExe, $cliExe, $hwpWorkerExe)) {
    if (-not (Test-Path $exe)) { throw "발행 결과에 $exe 가 없습니다." }
    Write-Host "  OK  $exe"
}

# ---------- 4. 클라이언트 설정 생성 ----------
Step '클라이언트 설정 생성 → dist\clients'
$clients = Join-Path $dist 'clients'
New-Item -ItemType Directory -Path $clients -Force | Out-Null

$mcpExeToml = $mcpExe -replace '\\', '/'       # TOML 은 슬래시가 안전

# Claude Desktop/Code, Kimi CLI, Cursor가 함께 읽는 표준 mcpServers JSON.
# ConvertTo-Json을 사용해 Windows 역슬래시를 정확히 이스케이프한다.
$serverMap = [ordered]@{}
$serverMap['doc-bridge'] = [ordered]@{
    command = $mcpExe
    args = @('--stdio')
}
$mcpConfig = [ordered]@{ mcpServers = $serverMap }
$mcpJson = $mcpConfig | ConvertTo-Json -Depth 8
$mcpJson | Set-Content (Join-Path $clients 'claude_desktop_config.json') -Encoding UTF8
$mcpJson | Set-Content (Join-Path $clients 'claude-code.mcp.json') -Encoding UTF8
$mcpJson | Set-Content (Join-Path $clients 'kimi-mcp.json') -Encoding UTF8
$mcpJson | Set-Content (Join-Path $clients 'cursor-mcp.json') -Encoding UTF8

# Codex  (~/.codex/config.toml)
@"
[mcp_servers.doc-bridge]
command = "$mcpExeToml"
args = ["--stdio"]
startup_timeout_sec = 30
tool_timeout_sec = 300

[mcp_servers.doc-bridge.tools.excel_apply_ops]
approval_mode = "approve"

[mcp_servers.doc-bridge.tools.hwp_apply_ops]
approval_mode = "approve"

[mcp_servers.doc-bridge.tools.cad_apply_ops]
approval_mode = "approve"

[mcp_servers.doc-bridge.tools.core_restore_snapshot]
approval_mode = "approve"
"@ | Set-Content (Join-Path $clients 'codex-config.toml') -Encoding UTF8

Get-ChildItem $clients | ForEach-Object { Write-Host "  생성  $($_.FullName)" }

# ---------- 5. 연결 스모크 테스트 ----------
Step '연결 확인 (initialize → tools/list → core_ping)'
& (Join-Path $PSScriptRoot 'verify-mcp.ps1') -Exe $mcpExe
if ($LASTEXITCODE -ne 0) { throw "연결 확인 실패" }

Step '완료'
Write-Host @"
다음 단계 — dist\clients 의 파일 내용을 각 클라이언트 설정에 병합하세요.

  Claude Desktop : %APPDATA%\Claude\claude_desktop_config.json   (병합 후 Claude 재시작)
  Kimi CLI       : ~\.kimi\mcp.json                              (또는: kimi mcp add --transport stdio doc-bridge -- "$mcpExe" --stdio)
  Cursor         : %USERPROFILE%\.cursor\mcp.json                (병합 후 Cursor 완전 재시작)
  Codex          : Codex 플러그인으로 설치하거나 ~\.codex\config.toml 에 병합

자세한 절차는 INSTALL.md 를 보세요.
"@ -ForegroundColor Green

