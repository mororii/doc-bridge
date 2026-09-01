<#
.SYNOPSIS
  doc-bridge MCP 서버가 실제로 핸드셰이크에 성공하는지 확인한다.
  (클라이언트에 등록하기 전에 이걸로 먼저 걸러낸다.)

.DESCRIPTION
  stdio 모드: exe를 파이프로 띄워 initialize → notifications/initialized → tools/list
             → tools/call core_ping 을 보내고 응답을 검증한다.
             -RequireExcelRuntime을 주면 이미 실행 중인 Excel 프로세스를 먼저 확인한 뒤
             excel_get_active_context를 같은 MCP 세션에서 2회 호출한다. 기본 검사에서는
             Excel을 새로 실행하지 않으며 live Excel 검증을 SKIP으로 표시한다.
             stdout에 JSON-RPC 이외의 줄(로그/배너)이 섞이면 실패로 본다.
  http 모드 : POST /mcp 로 같은 순서를 보낸다.

.EXAMPLE
  powershell.exe -ExecutionPolicy Bypass -File tools\verify-mcp.ps1
  powershell.exe -ExecutionPolicy Bypass -File tools\verify-mcp.ps1 -Exe C:\path\to\doc-bridge-mcp.exe
  powershell.exe -ExecutionPolicy Bypass -File tools\verify-mcp.ps1 -RequireExcelRuntime
  powershell.exe -ExecutionPolicy Bypass -File tools\verify-mcp.ps1 -Http -Port 5177
#>
[CmdletBinding()]
param(
    [string]$Exe,
    [switch]$Http,
    [int]$Port = 5177,
    [int]$TimeoutSec = 60,
    [switch]$RequireExcelRuntime
)

$ErrorActionPreference = 'Stop'
$utf8NoBom = New-Object System.Text.UTF8Encoding($false)
try { [Console]::OutputEncoding = $utf8NoBom } catch { }
$OutputEncoding = $utf8NoBom
$repo = Split-Path -Parent $PSScriptRoot

$script:fail = 0
function Check($label, [bool]$ok, $detail = '') {
    if ($ok) {
        Write-Host ("  [OK]   {0}" -f $label) -ForegroundColor Green
    } else {
        Write-Host ("  [FAIL] {0} {1}" -f $label, $detail) -ForegroundColor Red
        $script:fail++
    }
}

function Skip($label, $detail = '') {
    Write-Host ("  [SKIP] {0} {1}" -f $label, $detail) -ForegroundColor Yellow
}

function Get-ToolPayload($response) {
    if ($null -eq $response -or $null -eq $response.result) { return $null }
    if ($null -ne $response.result.structuredContent) { return $response.result.structuredContent }
    if ($null -ne $response.result.content -and $response.result.content.Count -gt 0) {
        try { return $response.result.content[0].text | ConvertFrom-Json }
        catch { return $null }
    }
    return $null
}

function Test-ExcelUnavailablePayload($payload) {
    if ($null -eq $payload -or [bool]$payload.ok) { return $false }
    $errorText = (@($payload.errors) -join ' ')
    if ([string]::IsNullOrWhiteSpace($errorText) -and $null -ne $payload.error) {
        $errorText = [string]$payload.error
    }
    return $errorText -match '(?i)Excel.*(not running|실행 중이지)|no active workbook|활성 통합문서'
}

function Test-ExcelContextPayload($payload) {
    if ($null -eq $payload -or -not [bool]$payload.ok) { return $false }
    if (@($payload.errors).Count -ne 0) { return $false }
    $summary = $payload.summary
    if ($null -eq $summary) { return $false }
    $required = @('workbook', 'sheets', 'activeSheet', 'usedRange', 'saved', 'openWorkbooks')
    foreach ($name in $required) {
        if ($null -eq $summary.PSObject.Properties[$name]) { return $false }
    }
    if ([string]::IsNullOrWhiteSpace([string]$summary.workbook) -or
        [string]::IsNullOrWhiteSpace([string]$summary.activeSheet) -or
        [string]::IsNullOrWhiteSpace([string]$summary.usedRange)) { return $false }
    if (@($summary.sheets).Count -lt 1 -or @($summary.openWorkbooks).Count -lt 1) { return $false }
    if ($null -eq $payload.selection -or
        [string]::IsNullOrWhiteSpace([string]$payload.selection.ref)) { return $false }
    return $true
}

$runExcelRuntime = $false
$excelPrecondition = ''
if ($RequireExcelRuntime) {
    try {
        $runExcelRuntime = @(Get-Process -Name EXCEL -ErrorAction SilentlyContinue).Count -gt 0
    } catch { $runExcelRuntime = $false }
    if (-not $runExcelRuntime) {
        $excelPrecondition = 'Excel 프로세스가 없습니다. 통합문서를 먼저 연 뒤 다시 실행하세요.'
    }
}

$requests = @(
    '{"jsonrpc":"2.0","id":1,"method":"initialize","params":{"protocolVersion":"2025-06-18","capabilities":{},"clientInfo":{"name":"verify-mcp","version":"1.0"}}}'
    '{"jsonrpc":"2.0","method":"notifications/initialized"}'
    '{"jsonrpc":"2.0","id":2,"method":"tools/list"}'
    '{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"core_ping","arguments":{}}}'
)
if ($runExcelRuntime) {
    $requests += '{"jsonrpc":"2.0","id":4,"method":"tools/call","params":{"name":"excel_get_active_context","arguments":{}}}'
    $requests += '{"jsonrpc":"2.0","id":5,"method":"tools/call","params":{"name":"excel_get_active_context","arguments":{}}}'
}

# ---------------------------------------------------------------- HTTP 모드
if ($Http) {
    Write-Host "`n[doc-bridge] HTTP 모드 확인 — http://127.0.0.1:$Port/mcp" -ForegroundColor Cyan
    $responses = @()
    foreach ($r in $requests) {
        try {
            $res = Invoke-WebRequest -Uri "http://127.0.0.1:$Port/mcp" -Method Post `
                -ContentType 'application/json' -Headers @{ Accept = 'application/json, text/event-stream' } `
                -Body ([System.Text.Encoding]::UTF8.GetBytes($r)) -TimeoutSec $TimeoutSec -UseBasicParsing
        } catch {
            Check 'HTTP 요청' $false $_.Exception.Message
            exit 1
        }
        if ($res.StatusCode -eq 202) { continue }   # notification
        $responses += ($res.Content | ConvertFrom-Json)
    }
}
# --------------------------------------------------------------- stdio 모드
else {
    if (-not $Exe) {
        $candidate = Join-Path $repo 'dist\doc-bridge-mcp.exe'
        if (Test-Path $candidate) {
            $Exe = $candidate
        } else {
            $candidate = Join-Path $repo 'src\DocBridge.Mcp\bin\Release\net8.0-windows\doc-bridge-mcp.exe'
            if (-not (Test-Path $candidate)) {
                $candidate = Join-Path $repo 'src\DocBridge.Mcp\bin\Debug\net8.0-windows\doc-bridge-mcp.exe'
            }
            if (-not (Test-Path $candidate)) {
                throw "doc-bridge-mcp.exe 를 찾을 수 없습니다. 먼저 tools\publish.ps1 을 실행하세요."
            }
            $Exe = $candidate
        }
    }
    Write-Host "`n[doc-bridge] stdio 모드 확인 — $Exe" -ForegroundColor Cyan

    $psi = [System.Diagnostics.ProcessStartInfo]::new()
    $psi.FileName = $Exe
    $psi.Arguments = '--stdio'
    $psi.UseShellExecute = $false
    $psi.RedirectStandardInput = $true
    $psi.RedirectStandardOutput = $true
    $psi.RedirectStandardError = $true
    # Windows PowerShell 5.1의 .NET Framework ProcessStartInfo에는 일부 Encoding 속성이 없다.
    $utf8NoBom = New-Object System.Text.UTF8Encoding($false)
    if ($psi.PSObject.Properties.Name -contains 'StandardInputEncoding') { $psi.StandardInputEncoding = $utf8NoBom }
    if ($psi.PSObject.Properties.Name -contains 'StandardOutputEncoding') { $psi.StandardOutputEncoding = $utf8NoBom }
    if ($psi.PSObject.Properties.Name -contains 'StandardErrorEncoding') { $psi.StandardErrorEncoding = $utf8NoBom }
    $psi.CreateNoWindow = $true

    $proc = [System.Diagnostics.Process]::Start($psi)

    # 파이프 버퍼가 차서 교착되지 않도록 읽기를 먼저 시작한다.
    $stdoutTask = $proc.StandardOutput.ReadToEndAsync()
    $stderrTask = $proc.StandardError.ReadToEndAsync()

    # StreamWriter 기본 인코딩은 부모 프로세스 상태에 따라 첫 줄에 BOM을 넣을 수 있다.
    # (publish.ps1 안에서 호출될 때 initialize가 parse error가 되던 원인.)
    # UTF-8 no-BOM 바이트를 직접 써서 독립 실행/중첩 실행 결과를 동일하게 만든다.
    $inputBytes = $utf8NoBom.GetBytes(($requests -join "`n") + "`n")
    $inputStream = $proc.StandardInput.BaseStream
    $inputStream.Write($inputBytes, 0, $inputBytes.Length)
    $inputStream.Flush()
    $inputStream.Close()      # EOF → 서버가 정상 종료

    if (-not $proc.WaitForExit($TimeoutSec * 1000)) {
        try { $proc.Kill() } catch { }
        Check "서버가 ${TimeoutSec}s 안에 종료(EOF 처리)" $false '입력 종료 후에도 응답이 없습니다'
        exit 1
    }

    $stdout = $stdoutTask.GetAwaiter().GetResult()
    $stderr = $stderrTask.GetAwaiter().GetResult()

    $lines = @($stdout -split "`r?`n" | Where-Object { $_.Trim() -ne '' })
    $expectedResponseCount = if ($runExcelRuntime) { 5 } else { 3 }
    Check "stdout 응답 ${expectedResponseCount}줄 (notification 제외)" `
        ($lines.Count -eq $expectedResponseCount) "실제 $($lines.Count)줄"

    $responses = @()
    foreach ($line in $lines) {
        try {
            $obj = $line | ConvertFrom-Json
        } catch {
            Check 'stdout 순수성 (JSON-RPC 외 출력 없음)' $false "비-JSON 줄: $line"
            continue
        }
        if ($obj.jsonrpc -ne '2.0') {
            Check 'stdout 순수성 (jsonrpc 2.0)' $false $line
        }
        $responses += $obj
    }
    if ($stderr) { Write-Verbose "stderr: $stderr" }
}

# ---------------------------------------------------------------- 응답 검증
$init  = $responses | Where-Object { $_.id -eq 1 } | Select-Object -First 1
$list  = $responses | Where-Object { $_.id -eq 2 } | Select-Object -First 1
$ping  = $responses | Where-Object { $_.id -eq 3 } | Select-Object -First 1
$excelContext1 = $responses | Where-Object { $_.id -eq 4 } | Select-Object -First 1
$excelContext2 = $responses | Where-Object { $_.id -eq 5 } | Select-Object -First 1

$responseIds = @($responses | ForEach-Object { [string]$_.id }) -join ','
$firstResponse = if ($responses.Count -gt 0) { $responses[0] | ConvertTo-Json -Compress -Depth 8 } else { '<none>' }
$request0 = [string]$requests[0]
$initDetail = if ($init.error.message) { $init.error.message } else { "response ids=$responseIds first=$firstResponse request0=$request0" }
Check 'initialize 응답' ($null -ne $init.result) $initDetail
Check 'protocolVersion 협상 (2025-06-18 echo)' ($init.result.protocolVersion -eq '2025-06-18') $init.result.protocolVersion
Check 'serverInfo.name = doc-bridge' ($init.result.serverInfo.name -eq 'doc-bridge') $init.result.serverInfo.name
Check 'capabilities.tools 선언' ($null -ne $init.result.capabilities.tools)

$toolNames = @($list.result.tools | ForEach-Object { $_.name })
$expectedToolNames = @(
    'core_ping','core_get_status','core_get_capabilities','core_disconnect','core_create_snapshot','core_list_snapshots','core_restore_snapshot',
    'excel_get_active_context','excel_read_range','excel_inspect','excel_apply_ops','excel_disconnect',
    'hwp_plan_creation','hwp_launch','hwp_get_active_context','hwp_doctor','hwp_repair_typelib','hwp_read_text','hwp_apply_ops','hwp_submit_ops','hwp_get_job',
    'cad_launch','cad_get_active_context','cad_query_entities','cad_apply_ops'
)
$missingTools = @($expectedToolNames | Where-Object { $toolNames -notcontains $_ })
$extraTools = @($toolNames | Where-Object { $expectedToolNames -notcontains $_ })
Check 'tools/list 정확한 공개 도구 집합' (($missingTools.Count -eq 0) -and ($extraTools.Count -eq 0)) `
    "actual=$($toolNames.Count), missing=$($missingTools -join ','), extra=$($extraTools -join ',')"
foreach ($want in @('core_ping','core_disconnect','excel_get_active_context','excel_inspect','excel_apply_ops','excel_disconnect',
                    'hwp_plan_creation','hwp_launch','hwp_get_active_context','hwp_doctor','hwp_repair_typelib','hwp_apply_ops','hwp_submit_ops','hwp_get_job',
                    'cad_launch','cad_get_active_context','cad_apply_ops')) {
    Check "tool 존재: $want" ($toolNames -contains $want)
}
Check 'tool 이름에 점(.) 없음' (-not ($toolNames | Where-Object { $_ -like '*.*' }))
Check '모든 tool에 inputSchema' (@($list.result.tools | Where-Object { $null -eq $_.inputSchema }).Count -eq 0)

Check 'core_ping isError=false' ($ping.result.isError -eq $false)
$payload = $null
if ($ping.result.content) { $payload = $ping.result.content[0].text | ConvertFrom-Json }
Check 'core_ping ok=true' ($payload.ok -eq $true)
Check 'core_ping 어댑터 excel/hwp/cad 등록' (
    ($payload.adapters -contains 'excel') -and
    ($payload.adapters -contains 'hwp') -and
    ($payload.adapters -contains 'cad')) ($payload.adapters -join ',')
Check 'structuredContent 동봉' ($null -ne $ping.result.structuredContent)

$excelPayload1 = Get-ToolPayload $excelContext1
$excelPayload2 = Get-ToolPayload $excelContext2
if (-not $RequireExcelRuntime) {
    Skip 'Excel live RCW context' '기본 진단은 Excel을 실행하지 않습니다. 통합문서를 연 뒤 2-TEST.cmd -RequireExcelRuntime으로 확인하세요.'
} elseif (-not $runExcelRuntime) {
    Check 'Excel live RCW context' $false $excelPrecondition
} elseif ((Test-ExcelUnavailablePayload $excelPayload1) -and
    (Test-ExcelUnavailablePayload $excelPayload2)) {
    Check 'Excel live RCW context' $false 'Excel은 실행 중이지만 활성 통합문서가 없습니다.'
} else {
    $context1Ok = $excelContext1.result.isError -eq $false -and (Test-ExcelContextPayload $excelPayload1)
    $context2Ok = $excelContext2.result.isError -eq $false -and (Test-ExcelContextPayload $excelPayload2)
    $contextDetail = "first=$($excelPayload1 | ConvertTo-Json -Compress -Depth 5); second=$($excelPayload2 | ConvertTo-Json -Compress -Depth 5)"
    Check 'Excel live RCW context first read' $context1Ok $contextDetail
    Check 'Excel live RCW context repeated read' $context2Ok $contextDetail
    if ($context1Ok -and $context2Ok) {
        Check 'Excel context document remains stable' `
            ([string]$excelPayload1.documentRef -eq [string]$excelPayload2.documentRef) `
            "first=$([string]$excelPayload1.documentRef); second=$([string]$excelPayload2.documentRef)"
    }
}

Write-Host ''
if ($script:fail -eq 0) {
    Write-Host "모두 통과 — 클라이언트에 등록해도 됩니다." -ForegroundColor Green
    exit 0
} else {
    Write-Host "$($script:fail)건 실패" -ForegroundColor Red
    exit 1
}

