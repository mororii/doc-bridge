<#
.SYNOPSIS
  실제 Excel / 한글(HWP) / AutoCAD 를 띄워 doc-bridge 전 경로를 검증한다.

.DESCRIPTION
  각 프로그램마다:
    읽기 → dry-run(diff+토큰) → apply → COM으로 직접 결과 확인 → 스냅샷 복원 → 원복 확인
  까지 수행한다. 원본 demo 파일은 건드리지 않고 작업용 사본으로만 테스트한다.

  설치돼 있지 않은 프로그램 구간은 SKIP 으로 넘어간다.
  결과는 화면과 로그 파일에 동시에 기록된다.

.PARAMETER Only
  excel / hwp / cad 중 하나만 실행.

.PARAMETER SkipUnitTests
  dotnet test 단계를 건너뛴다.

.EXAMPLE
  powershell.exe -ExecutionPolicy Bypass -File tools\run-e2e.ps1
  powershell.exe -ExecutionPolicy Bypass -File tools\run-e2e.ps1 -Only excel
#>
[CmdletBinding()]
param(
    [ValidateSet('excel', 'hwp', 'cad', 'all')]
    [string]$Only = 'all',
    [switch]$SkipUnitTests
)

$ErrorActionPreference = 'Continue'
$repo = Split-Path -Parent $PSScriptRoot
$workspace = Split-Path -Parent $repo
$logPath = Join-Path $workspace 'e2e-result.log'
$scratch = Join-Path $env:TEMP ("docbridge-e2e-" + (Get-Date -Format 'yyyyMMdd-HHmmss'))
New-Item -ItemType Directory -Path $scratch -Force | Out-Null

"" | Set-Content $logPath -Encoding UTF8
$script:pass = 0
$script:fail = 0
$script:skip = 0

function Log($msg, $color = 'Gray') {
    Write-Host $msg -ForegroundColor $color
    Add-Content -Path $logPath -Value $msg -Encoding UTF8
}
function Section($t) { Log ""; Log ("=" * 70); Log "  $t"; Log ("=" * 70) }
function Ok($t, $d = '')     { $script:pass++; Log ("  [OK]   $t" + $(if ($d) { "  -> $d" })) 'Green' }
function Bad($t, $d = '')    { $script:fail++; Log ("  [FAIL] $t" + $(if ($d) { "  -> $d" })) 'Red' }
function Skip($t, $d = '')   { $script:skip++; Log ("  [SKIP] $t" + $(if ($d) { "  -> $d" })) 'Yellow' }
function Check($t, [bool]$c, $d = '') { if ($c) { Ok $t $d } else { Bad $t $d } }
function Info($msg)          { Log "         $msg" 'DarkGray' }

Log "doc-bridge 실제 프로그램 E2E — $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
Log "repo:    $repo"
Log "scratch: $scratch"
Log "log:     $logPath"

# ---------------------------------------------------------------- 0. 준비
Section '0. 준비 — dotnet / CLI 빌드'

$dotnet = Get-Command dotnet -ErrorAction SilentlyContinue
if (-not $dotnet) {
    Bad 'dotnet SDK' '설치되어 있지 않습니다. https://dotnet.microsoft.com/download/dotnet/8.0'
    Log "`n중단합니다."
    exit 1
}
Ok 'dotnet SDK' (& dotnet --version)

$cliExe = Join-Path $repo 'src\DocBridge.Cli\bin\Release\net8.0-windows\doc-bridge-cli.exe'
Info '현재 소스로 CLI를 Release 빌드합니다...'
& dotnet build (Join-Path $repo 'src\DocBridge.Cli') -c Release --nologo -v quiet 2>&1 |
    ForEach-Object { Info $_ }
if (-not (Test-Path $cliExe)) { Bad 'CLI 빌드' $cliExe; exit 1 }
if ($LASTEXITCODE -ne 0) { Bad 'CLI 빌드' "exit=$LASTEXITCODE"; exit 1 }
Ok 'CLI 준비' $cliExe

# CLI 호출 헬퍼 — stdout 마지막 JSON 줄을 객체로 돌려준다
function Invoke-DocBridgeCli {
    param([Parameter(ValueFromRemainingArguments = $true)][string[]]$CliArgs)
    $raw = & $cliExe @CliArgs 2>$null
    $line = @($raw | Where-Object { $_ -and $_.Trim().StartsWith('{') }) | Select-Object -Last 1
    if (-not $line) { return [pscustomobject]@{ ok = $false; errors = @("CLI 응답 없음: $($CliArgs -join ' ')") } }
    try { return $line | ConvertFrom-Json }
    catch { return [pscustomobject]@{ ok = $false; errors = @("CLI 응답 파싱 실패: $line") } }
}

function WriteOps($path, $ops) {
    [ordered]@{ ops = @($ops) } | ConvertTo-Json -Depth 12 | Set-Content $path -Encoding UTF8
}

function Invoke-DocBridgeJson($tool, $value) {
    $path = Join-Path $scratch ("args-" + [guid]::NewGuid().ToString('N') + '.json')
    $value | ConvertTo-Json -Depth 12 | Set-Content $path -Encoding UTF8
    Invoke-DocBridgeCli $tool --json-file $path
}

# 프로그램 설치 여부 (ProgID 등록 확인)
function ProgIdExists($progId) {
    try { return $null -ne [Type]::GetTypeFromProgID($progId) } catch { return $false }
}

if (-not $SkipUnitTests) {
    Section '0-1. 단위 테스트 (Office 불필요)'
    Remove-Item Env:DOCBRIDGE_E2E -ErrorAction SilentlyContinue
    $t = & dotnet test (Join-Path $repo 'DocBridge.sln') -c Release --nologo -v minimal `
        --filter 'FullyQualifiedName!~E2ETests&Category!=E2E' 2>&1
    $t | ForEach-Object { Info $_ }
    Check '단위 테스트 통과' ($LASTEXITCODE -eq 0) "exit=$LASTEXITCODE"
}

# ================================================================ 1. EXCEL
if ($Only -in @('all', 'excel')) {
    Section '1. Excel — 실제 프로그램 조작'
    $xl = $null
    $wb = $null
    $ownsExcel = $false
    try {
        if (-not (ProgIdExists 'Excel.Application')) {
            Skip 'Excel' '설치되어 있지 않습니다 (Excel.Application 미등록)'
        } else {
            $src = Join-Path $workspace 'demo-excel.xlsx'
            $work = Join-Path $scratch 'excel-test.xlsx'
            if (Test-Path $src) { Copy-Item $src $work } else { $work = Join-Path $scratch 'excel-new.xlsx' }

            Info 'Excel 실행 중...'
            try {
                $xl = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
                Info '기존 Excel 인스턴스에 테스트 사본을 엽니다.'
            } catch {
                $xl = New-Object -ComObject Excel.Application
                $ownsExcel = $true
            }
            $xl.Visible = $true
            $xl.DisplayAlerts = $false
            if (Test-Path $work) { $wb = $xl.Workbooks.Open($work) }
            else { $wb = $xl.Workbooks.Add(); $wb.SaveAs($work) }
            $ws = $wb.Worksheets.Item(1)
            $wb.Activate()
            $ws.Activate()
            [void]$ws.Range('A1').Select()

            # 결정적인 테스트 값 심기
            $marker = 'DOCBRIDGE-원본'
            $ws.Range('A1').Value2 = $marker
            $wb.Save()
            Ok 'Excel 실행 + 테스트 문서 열기' (Split-Path $work -Leaf)
            Info "A1 초기값: '$($ws.Range('A1').Value2)'"

            # 1) 읽기
            $ctx = Invoke-DocBridgeCli excel_get_active_context
            Check 'excel_get_active_context' ([bool]$ctx.ok) ($ctx.errors -join '; ')
            if ($ctx.ok) { Info "workbook=$($ctx.summary.workbook)  sheet=$($ctx.summary.activeSheet)  selection=$($ctx.selection.ref)" }
            if (-not $ctx.ok -or -not [string]::Equals([string]$ctx.documentRef, [string]$work, [StringComparison]::OrdinalIgnoreCase)) {
                throw "doc-bridge가 테스트 사본이 아닌 workbook에 연결되었습니다. expected='$work', actual='$($ctx.documentRef)'. 안전을 위해 쓰기 테스트를 중단합니다."
            }

            $rd = Invoke-DocBridgeJson excel_read_range @{ range = 'A1:B3' }
            Check 'excel_read_range A1:B3' ([bool]$rd.ok) ($rd.errors -join '; ')

            # 2) dry-run
            $opsPath = Join-Path $scratch 'excel-ops.json'
            WriteOps $opsPath @(@{
                op = 'find_replace'
                target = @{ scope = 'workbook' }
                find = $marker
                replace = 'DOCBRIDGE-변경됨'
            })
            $dry = Invoke-DocBridgeCli excel_apply_ops --ops $opsPath --dry-run
            Check 'dry-run 성공' ([bool]$dry.ok) ($dry.errors -join '; ')
            Check 'confirmToken 발급' ([bool]$dry.confirmToken)
            Check 'snapshotId 발급' ([bool]$dry.snapshotId)
            Check 'dry-run 이후 문서 미변경' ($ws.Range('A1').Value2 -eq $marker) $ws.Range('A1').Value2
            if ($dry.diff) { Info "diff: $($dry.diff | ConvertTo-Json -Compress -Depth 4)" }

            if ($dry.ok -and $dry.confirmToken) {
                # 3) apply
                $ap = Invoke-DocBridgeCli excel_apply_ops --ops $opsPath --confirm-token $dry.confirmToken
                Check 'apply 성공' ([bool]$ap.ok) ($ap.errors -join '; ')
                Check 'readback.verified' ([bool]$ap.readback.verified) ($ap.readback | ConvertTo-Json -Compress)

                # 4) COM 으로 직접 확인 — 이게 진짜 증거
                $actual = $ws.Range('A1').Value2
                Check '★ Excel 문서가 실제로 바뀜' ($actual -eq 'DOCBRIDGE-변경됨') "A1='$actual'"

                # 5) 스냅샷 복원 (2단계)
                $r1 = Invoke-DocBridgeJson core_restore_snapshot @{ snapshotId = $dry.snapshotId }
                if ($r1.confirmToken) {
                    $r2 = Invoke-DocBridgeJson core_restore_snapshot @{ snapshotId = $dry.snapshotId; confirmToken = $r1.confirmToken }
                    Check '스냅샷 복원 성공' ([bool]$r2.ok) ($r2.errors -join '; ')
                    Check '★ 원본 값으로 되돌아옴' ($ws.Range('A1').Value2 -eq $marker) "A1='$($ws.Range('A1').Value2)'"
                } else {
                    Bad '복원 1단계 토큰 발급' ($r1.errors -join '; ')
                }
            }
        }
    } catch {
        Bad 'Excel 구간 예외' $_.Exception.Message
    } finally {
        try { if ($wb) { $wb.Close($false) } } catch { }
        try { if ($xl -and $ownsExcel) { $xl.Quit() } } catch { }
        try { if ($xl) { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($xl) } } catch { }
    }
}

# ================================================================ 2. 한글
if ($Only -in @('all', 'hwp')) {
    Section '2. 한글(HWP) — 파일 기반 조작'
    try {
        if (-not (ProgIdExists 'HWPFrame.HwpObject')) {
            Skip '한글' '설치되어 있지 않습니다 (HWPFrame.HwpObject 미등록)'
        } else {
            $src = Join-Path $workspace 'demo-hwp.hwp'
            if (-not (Test-Path $src)) {
                Skip '한글' "테스트 문서가 없습니다: $src"
            } else {
                $work = Join-Path $scratch 'hwp-test.hwp'
                Copy-Item $src $work
                $workJson = $work -replace '\\', '/'
                Ok '테스트 문서 준비' (Split-Path $work -Leaf)

                # 1) 읽기
                $rd = Invoke-DocBridgeJson hwp_read_text @{ file = $workJson; scope = 'document' }
                Check 'hwp_read_text (문서 전체)' ([bool]$rd.ok) ($rd.errors -join '; ')
                $original = "$($rd.text)"
                Info "원문 앞부분: $($original.Substring(0, [Math]::Min(60, $original.Length)))"

                # 문서에 실제로 있는 낱말을 골라 치환 대상으로 삼는다
                $word = ($original -split '[\s\r\n\.,]+' | Where-Object { $_.Length -ge 2 } | Select-Object -First 1)
                if (-not $word) {
                    Skip '한글 치환' '문서에서 치환할 낱말을 찾지 못했습니다'
                } else {
                    $replacement = "$word(수정됨)"
                    Info "치환 대상: '$word' -> '$replacement'"

                    $opsPath = Join-Path $scratch 'hwp-ops.json'
                    WriteOps $opsPath @(@{
                        op = 'find_replace'
                        file = $workJson
                        find = $word
                        replace = $replacement
                    })

                    # 2) dry-run
                    $dry = Invoke-DocBridgeCli hwp_apply_ops --ops $opsPath --dry-run
                    Check 'dry-run 성공' ([bool]$dry.ok) ($dry.errors -join '; ')
                    Check 'confirmToken 발급' ([bool]$dry.confirmToken)
                    Check 'snapshotId 발급' ([bool]$dry.snapshotId)

                    $afterDry = "$((Invoke-DocBridgeJson hwp_read_text @{ file = $workJson; scope = 'document' }).text)"
                    Check 'dry-run 이후 파일 미변경' ($afterDry -eq $original)

                    if ($dry.ok -and $dry.confirmToken) {
                        # 3) apply
                        $ap = Invoke-DocBridgeCli hwp_apply_ops --ops $opsPath --confirm-token $dry.confirmToken
                        Check 'apply 성공' ([bool]$ap.ok) ($ap.errors -join '; ')
                        Check 'readback.verified' ([bool]$ap.readback.verified) ($ap.readback | ConvertTo-Json -Compress)
                        if ($ap.warnings) { Info "warnings: $($ap.warnings -join ' / ')" }

                        # 4) 파일을 다시 읽어 확인
                        $after = "$((Invoke-DocBridgeJson hwp_read_text @{ file = $workJson; scope = 'document' }).text)"
                        Check '★ 한글 문서가 실제로 바뀜' ($after -like "*$replacement*") "'$replacement' 포함 여부"

                        # 5) 복원
                        $r1 = Invoke-DocBridgeJson core_restore_snapshot @{ snapshotId = $dry.snapshotId }
                        if ($r1.confirmToken) {
                            $r2 = Invoke-DocBridgeJson core_restore_snapshot @{ snapshotId = $dry.snapshotId; confirmToken = $r1.confirmToken }
                            Check '스냅샷 복원 성공' ([bool]$r2.ok) ($r2.errors -join '; ')
                            $restored = "$((Invoke-DocBridgeJson hwp_read_text @{ file = $workJson; scope = 'document' }).text)"
                            Check '★ 원본 내용으로 되돌아옴' ($restored -notlike "*$replacement*")
                        } else {
                            Bad '복원 1단계 토큰 발급' ($r1.errors -join '; ')
                        }
                    }
                }
            }
        }
    } catch {
        Bad '한글 구간 예외' $_.Exception.Message
    }
}

# ================================================================ 3. CAD
if ($Only -in @('all', 'cad')) {
    Section '3. AutoCAD — 실제 프로그램 조작'
    $acad = $null
    $doc = $null
    $ownsAcad = $false
    try {
        $progId = @('AutoCAD.Application', 'AutoCAD.Application.26', 'AutoCAD.Application.25', 'AutoCAD.Application.24') |
            Where-Object { ProgIdExists $_ } | Select-Object -First 1
        if (-not $progId) {
            Skip 'AutoCAD' '설치되어 있지 않습니다 (AutoCAD.Application 미등록)'
        } else {
            $src = Join-Path $workspace 'demo-cad.dwg'
            if (-not (Test-Path $src)) {
                Skip 'AutoCAD' "테스트 도면이 없습니다: $src"
            } else {
                $work = Join-Path $scratch 'cad-test.dwg'
                Copy-Item $src $work

                Info "AutoCAD 실행 중 (ProgID=$progId, 최초 기동은 1~2분 걸릴 수 있음)..."
                try {
                    $acad = [Runtime.InteropServices.Marshal]::GetActiveObject($progId)
                    Info '이미 실행 중인 AutoCAD 에 연결했습니다.'
                } catch {
                    $acad = New-Object -ComObject $progId
                    $ownsAcad = $true
                }
                $acad.Visible = $true
                $doc = $acad.Documents.Open($work)
                Start-Sleep -Seconds 2
                Ok 'AutoCAD 실행 + 도면 열기' (Split-Path $work -Leaf)

                # 1) 읽기
                $ctx = Invoke-DocBridgeCli cad_get_active_context
                Check 'cad_get_active_context' ([bool]$ctx.ok) ($ctx.errors -join '; ')
                if ($ctx.ok) { Info "drawing=$($ctx.summary.drawing)  layers=$($ctx.summary.layers.Count)  units=$($ctx.summary.insunits)" }

                $q = Invoke-DocBridgeJson cad_query_entities @{ limit = 10 }
                Check 'cad_query_entities' ([bool]$q.ok) ($q.errors -join '; ')
                if ($q.ok) { Info "entities=$($q.count)" }

                # 대상 레이어 고르기
                $layerName = $null
                foreach ($l in $doc.Layers) { if ($l.Name -ne '0') { $layerName = $l.Name; break } }
                if (-not $layerName) { $layerName = '0' }
                $beforeColor = ($doc.Layers.Item($layerName)).Color
                Info "대상 레이어: '$layerName' (현재 색=$beforeColor)"
                $targetColor = if ($beforeColor -eq 1) { 3 } else { 1 }

                # 2) dry-run
                $opsPath = Join-Path $scratch 'cad-ops.json'
                WriteOps $opsPath @(@{
                    op = 'set_layer_color'
                    layer = $layerName
                    color = $targetColor
                })
                $dry = Invoke-DocBridgeCli cad_apply_ops --ops $opsPath --dry-run
                Check 'dry-run 성공' ([bool]$dry.ok) ($dry.errors -join '; ')
                Check 'confirmToken 발급' ([bool]$dry.confirmToken)
                Check 'dry-run 이후 도면 미변경' (($doc.Layers.Item($layerName)).Color -eq $beforeColor)

                if ($dry.ok -and $dry.confirmToken) {
                    # 3) apply
                    $ap = Invoke-DocBridgeCli cad_apply_ops --ops $opsPath --confirm-token $dry.confirmToken
                    Check 'apply 성공' ([bool]$ap.ok) ($ap.errors -join '; ')
                    Check 'readback.verified' ([bool]$ap.readback.verified) ($ap.readback | ConvertTo-Json -Compress)

                    # 4) COM 으로 직접 확인
                    $actualColor = ($doc.Layers.Item($layerName)).Color
                    Check '★ AutoCAD 도면이 실제로 바뀜' ($actualColor -eq $targetColor) "레이어 '$layerName' 색: $beforeColor -> $actualColor"

                    # 5) 복원
                    $r1 = Invoke-DocBridgeJson core_restore_snapshot @{ snapshotId = $dry.snapshotId }
                    if ($r1.confirmToken) {
                        $r2 = Invoke-DocBridgeJson core_restore_snapshot @{ snapshotId = $dry.snapshotId; confirmToken = $r1.confirmToken }
                        Check '스냅샷 복원 성공' ([bool]$r2.ok) ($r2.errors -join '; ')
                        Check '★ 원래 색으로 되돌아옴' ((($doc.Layers.Item($layerName)).Color) -eq $beforeColor) "색=$(($doc.Layers.Item($layerName)).Color)"
                    } else {
                        Bad '복원 1단계 토큰 발급' ($r1.errors -join '; ')
                    }
                }

                # 6) 고위험 op 차단 확인 (실제로 지우지는 않는다)
                $hrPath = Join-Path $scratch 'cad-highrisk.json'
                WriteOps $hrPath @(@{ op = 'delete_entities'; handles = @('FFFFFFFF') })
                $hrDry = Invoke-DocBridgeCli cad_apply_ops --ops $hrPath --dry-run
                if ($hrDry.confirmToken) {
                    $hrNo = Invoke-DocBridgeCli cad_apply_ops --ops $hrPath --confirm-token $hrDry.confirmToken
                    Check '고위험 op는 highRiskConfirm 없으면 거부' (-not $hrNo.ok) ($hrNo.errors -join '; ')
                }
            }
        }
    } catch {
        Bad 'AutoCAD 구간 예외' $_.Exception.Message
    } finally {
        # 주의: acad.exe 를 강제종료하면 라이센싱 구성요소가 손상된다. 반드시 COM 으로 닫는다.
        try { if ($doc) { $doc.Close($false) } } catch { }
        try { if ($acad -and $ownsAcad) { $acad.Quit() } } catch { }
    }
}

# ================================================================ 요약
Section '요약'
Log "  통과 $script:pass / 실패 $script:fail / 건너뜀 $script:skip"
Log ""
Log "  작업 사본: $scratch  (원본 demo 파일은 건드리지 않았습니다)"
Log "  감사 로그: $env:LOCALAPPDATA\DocBridge\logs"
Log "  결과 파일: $logPath"

if ($script:fail -eq 0) { Log "`n  실제 프로그램 검증 통과." 'Green' }
else { Log "`n  실패 $script:fail 건 — 위 [FAIL] 항목을 확인하세요." 'Red' }

exit $(if ($script:fail -eq 0) { 0 } else { 1 })

