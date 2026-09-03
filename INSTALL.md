# DocBridge 설치와 클라이언트 연결

대상은 Windows의 Kimi CLI/Kimi Code, Cursor, Claude Code/Claude Desktop, Codex입니다.

## 0. 새 PC 통합 설치 패키지

새 컴퓨터에는 저장소 전체나 .NET SDK를 복사할 필요가 없습니다. 개발 PC에서 다음 명령으로 자체 포함 오프라인 ZIP을 만듭니다.

```powershell
cd C:\Tools\DocBridge
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\package-release.ps1
```

기본적으로 최종 ZIP과 SHA-256만 남깁니다. 설치기 개발 테스트를 위해 압축 해제 폴더도
유지하려면 `-KeepExpanded`를 추가합니다. 과거 펼친 폴더는
`tools\Clear-DocBridgeReleaseCache.ps1`을 먼저 미리보기로 실행한 뒤 `-Apply`로 정리할 수 있으며,
ZIP과 SHA-256 파일은 이 정리 대상에 포함되지 않습니다.

생성 결과는 `releases\DocBridge-0.4.19-win-x64.zip`입니다. ZIP에는 .NET 8 `win-x64` 런타임, 한글 COM 격리 worker, Codex 플러그인/로컬 marketplace, Claude/Kimi/Cursor MCP 설정 병합기, 한글 보안 모듈 등록기, 진단기와 제거기가 모두 들어갑니다.

새 PC에서 ZIP을 압축 해제하고 AI 클라이언트와 Office/CAD 프로그램을 종료한 뒤 실행합니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-DocBridge.ps1
```

기본 설치 위치는 `%LOCALAPPDATA%\DocBridge`입니다. 설치기는 기존 Claude Desktop, Kimi, Cursor 사용자 전역 JSON을 먼저 백업하고 `mcpServers.doc-bridge`만 병합합니다. Cursor의 프로젝트별 `.cursor\mcp.json`은 변경하지 않습니다. Codex는 PATH, 사용자 `config.toml`의 `CODEX_CLI_PATH`, 앱 내부 `%LOCALAPPDATA%\OpenAI\Codex\bin`, 앱 패키지 순서로 CLI를 탐색합니다. Codex가 설치된 PC에서는 플러그인이 실제 `installed, enabled` 상태가 아니면 설치를 실패 처리하며, `generated-configs` 아래에 절대 경로 기반 수동 명령을 남깁니다.

설치 확인과 제거는 다음과 같습니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Test-DocBridge.ps1
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Uninstall-DocBridge.ps1 -RemoveHwpSecurity
```

기본 진단은 Excel을 새로 실행하지 않습니다. Excel RCW 경로까지 실제로 확인하려면 대상 통합문서를
먼저 연 뒤 `2-EXCEL-LIVE-TEST.cmd`를 실행하거나 다음 명령을 사용합니다. 같은 MCP 프로세스에서
활성 컨텍스트를 두 번 읽어 `errors=[]`, 필수 summary/selection, 동일 문서 유지 여부를 검사합니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Test-DocBridge.ps1 -RequireExcelRuntime
```

설치기는 같은 이름의 기존 `doc-bridge` 항목을 `installation.json`에 기록합니다. 제거 시 현재 항목이 설치 당시 command/args와 정확히 일치할 때만 제거하거나 이전 항목을 복원하며, 사용자가 설치 뒤 바꾼 항목은 경고와 함께 그대로 둡니다. `-Clients Cursor`처럼 일부 클라이언트만 제거하면 공유 실행 파일과 한글 보안 등록은 남은 클라이언트를 위해 유지되고, `installation.json`의 관리 대상만 갱신됩니다. 마지막 관리 클라이언트를 제거할 때 공유 payload와 활성 설치 기록이 백업 폴더로 이동합니다.

제거기는 다른 MCP 설정을 보존하고 `doc-bridge` 항목만 지웁니다. 배포 파일은 기본적으로 백업 폴더로 이동하며, 백업과 로그까지 삭제하려는 경우에만 `-RemoveData`를 추가합니다.

## 1. 배포본 확인

현재 저장소에는 .NET 런타임을 포함한 `win-x64` 배포본이 이미 있습니다.

```powershell
cd C:\Tools\DocBridge
.\dist\doc-bridge-mcp.exe --version
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\verify-mcp.ps1
```

예상 버전은 `0.4.19`이며 검증 끝에 `모두 통과`가 나와야 합니다. 배포 무결성 검사에는 `doc-bridge-hwp-worker.exe`도 포함됩니다.

GitHub 기본 CI는 Office가 없어도 재현되는 테스트만 실행하고, Excel·한글·AutoCAD 실제 COM 검증은 전용 Windows self-hosted workflow에서 수동 실행합니다. 준비와 안전 규칙은 [실제 앱 E2E 운영 안내](docs/REAL-APP-E2E.md)를 참고하세요.

소스에서 다시 만들려면 .NET 8 SDK가 필요합니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\publish.ps1 -SelfContained
```

`publish.ps1`은 단위 테스트, self-contained 발행, 개발 PC용 Claude·Kimi·Cursor·Codex 설정 생성, MCP 핸드셰이크 검증을 한 번에 수행합니다. 이 개발 경로가 들어간 `dist\clients` 파일은 배포 ZIP에 포함되지 않습니다. 다른 PC용 설정은 설치기가 `%LOCALAPPDATA%\DocBridge\generated-configs`에 실제 설치 경로로 다시 생성합니다. Windows PowerShell 5.1과 PowerShell 7에서 동작합니다.

## 2. Kimi

현재 Kimi CLI의 사용자 MCP 파일은 `~\.kimi\mcp.json`입니다. 명령으로 등록하는 방법이 가장 단순합니다.

```powershell
kimi mcp add --transport stdio doc-bridge -- "C:\Tools\DocBridge\dist\doc-bridge-mcp.exe" --stdio
kimi mcp list
kimi mcp test doc-bridge
```

또는 설치 후 `%LOCALAPPDATA%\DocBridge\generated-configs\kimi-mcp.json`의 `mcpServers.doc-bridge` 항목을 기존 `~\.kimi\mcp.json`에 병합합니다. 이 파일은 대상 PC의 실제 설치 경로로 생성됩니다. 임시로만 로드하려면:

```powershell
kimi --mcp-config-file "$env:LOCALAPPDATA\DocBridge\generated-configs\kimi-mcp.json"
```

구형 Kimi Code를 쓰는 경우 프로젝트의 `.kimi-code\mcp.json`에도 같은 `mcpServers` 형식을 사용할 수 있습니다.

공식 문서: [Kimi MCP configuration](https://moonshotai.github.io/kimi-cli/en/customization/mcp.html)

## 3. Cursor

통합 설치기는 다음 사용자 전역 설정을 먼저 백업하고 `mcpServers.doc-bridge`만 병합합니다.

```text
%USERPROFILE%\.cursor\mcp.json
```

이 전역 파일은 모든 Cursor 프로젝트에서 사용됩니다. 반면 `<프로젝트>\.cursor\mcp.json`은 해당 프로젝트 전용 설정이므로 설치기가 만들거나 변경하지 않습니다. 특별한 이유가 없다면 같은 이름의 서버를 전역과 프로젝트 양쪽에 중복 등록하지 마십시오. 수동 설정이 필요하면 설치 후 `%LOCALAPPDATA%\DocBridge\generated-configs\cursor-mcp.json`의 `doc-bridge` 항목만 기존 JSON에 병합합니다.

설치 후 Cursor를 완전히 종료했다가 다시 실행하고 Cursor의 MCP 설정 화면에서 `doc-bridge`가 활성 상태인지 확인합니다. 첫 요청은 다음처럼 읽기 전용으로 시작합니다.

```text
doc-bridge로 core_ping과 core_get_status를 실행하고 Excel, 한글, AutoCAD 연결 상태만 보여줘. 아직 수정하지 마.
```

배포본은 `%LOCALAPPDATA%\DocBridge\generated-configs\cursor`에 다음 파일을 설치합니다.

- `docbridge-user-rule.txt`: Cursor Settings → Rules의 사용자 규칙에 붙여 넣는 일반 텍스트
- `rules\docbridge-safe-automation.mdc`: 프로젝트의 `.cursor\rules`에 선택적으로 복사하는 규칙
- `CURSOR_USAGE.md`: 전역/프로젝트 설정, 승인 작업, 대형 CAD 후속조회 안내

쓰기 작업은 반드시 `dryRun=true`의 diff와 confirmToken을 사용자에게 보여 주고 승인을 받은 뒤 정확히 같은 ops로 적용하고 readback합니다. confirmToken은 5분 동안 한 번만 유효하므로 ops가 바뀌거나 사용자가 같은 문서를 수정했으면 재사용하지 않습니다. Cursor와 Codex·Claude·Kimi가 동시에 연결되어 있어도 같은 문서를 동시에 편집하면 안 됩니다.

공식 문서: [Cursor MCP](https://docs.cursor.com/context/model-context-protocol), [Cursor Rules](https://docs.cursor.com/context/rules)

## 4. Claude

### Claude Code

프로젝트 루트의 `.mcp.json`으로 공유하거나 사용자 scope로 등록할 수 있습니다. 대상 PC의 실제 경로가 들어간 예시는 설치 후 `%LOCALAPPDATA%\DocBridge\generated-configs\claude-code.mcp.json`에 생성됩니다.

```powershell
claude mcp add --scope user doc-bridge -- "C:\Tools\DocBridge\dist\doc-bridge-mcp.exe" --stdio
claude mcp list
```

프로젝트 설정을 사용할 때는 생성된 JSON을 프로젝트의 `.mcp.json`으로 복사합니다. Claude Code는 프로젝트 scope MCP를 처음 사용할 때 신뢰 승인을 요청합니다.

### Claude Desktop

설치 후 `%LOCALAPPDATA%\DocBridge\generated-configs\claude_desktop_config.json`의 `mcpServers.doc-bridge`를 다음 파일의 기존 `mcpServers` 객체에 병합합니다.

```text
%APPDATA%\Claude\claude_desktop_config.json
```

기존 설정을 통째로 덮어쓰지 마세요. 병합 후 Claude Desktop을 트레이 아이콘까지 완전히 종료한 다음 다시 실행합니다. 최신 Claude Desktop에서는 로컬 MCP를 Desktop Extension(DXT)로 배포할 수도 있지만, 이 저장소는 개발/검증이 쉬운 표준 stdio 설정을 제공합니다.

공식 문서: [Anthropic MCP](https://docs.anthropic.com/en/docs/mcp)

## 5. Codex

### 통합 설치 ZIP의 로컬 플러그인

새 PC에서는 `1-INSTALL.cmd`가 `%LOCALAPPDATA%\DocBridge\codex-marketplace`를 등록합니다. 현재 Codex CLI는 예전의 최상위 `plugin add/list` 명령을 제공하지 않으므로, 배포 marketplace의 `INSTALLED_BY_DEFAULT` 정책으로 DocBridge가 활성화됩니다. Codex를 완전히 재시작한 뒤 `2-TEST.cmd` 결과를 확인합니다.

```text
Codex CLI discovery                  OK
Codex plugin CLI mode                MarketplaceOnly
Codex marketplace registered         OK
Codex MCP visible and enabled        OK
```

설치기는 일반 PATH에 `codex`가 없어도 Codex 앱 내부 CLI를 찾습니다. 자동 등록이 실패하면 `%LOCALAPPDATA%\DocBridge\generated-configs\install-codex.txt`에 발견된 CLI 절대 경로를 사용하는 명령이 생성됩니다.

설치 뒤 Codex를 완전히 재시작하고 새 작업을 만들어야 bundled skill과 MCP tool이 로드됩니다. 기존 작업은 설치 전 도구 목록을 유지할 수 있습니다. 새 작업에도 도구가 없다면 AI가 프로젝트 파일이나 대체 스크립트를 수정하도록 두지 말고 `2-TEST.cmd`를 다시 실행합니다.

### 개발 PC의 개인 플러그인

이 PC에는 다음 위치로 개인 플러그인과 marketplace 항목을 만들어 두었습니다.

```text
%USERPROFILE%\plugins\doc-bridge
%USERPROFILE%\.agents\plugins\marketplace.json
```

marketplace 이름은 `personal`, 플러그인 이름은 `doc-bridge`입니다. Codex의 Plugins 화면에서 Personal → DocBridge를 설치하거나, marketplace가 아직 등록되지 않았다면 일반 Codex CLI에서 다음을 실행한 뒤 Codex를 완전히 재시작합니다.

```powershell
codex plugin marketplace add "$env:USERPROFILE"
```

플러그인 매니페스트는 `./dist/doc-bridge-mcp.exe` 상대 경로를 사용하며, 배포된 개인 플러그인 복사본에서도 버전 실행과 공식 validator를 통과해야 합니다.

### config.toml 직접 등록

플러그인을 사용하지 않을 때는 설치 후 대상 PC 경로로 생성된 `%LOCALAPPDATA%\DocBridge\generated-configs\codex-config.toml`을 `~\.codex\config.toml`에 병합합니다.

```toml
[mcp_servers.doc-bridge]
command = "C:/Tools/DocBridge/dist/doc-bridge-mcp.exe"
args = ["--stdio"]
startup_timeout_sec = 30
tool_timeout_sec = 300
```

쓰기 도구와 snapshot 복원은 생성된 예시처럼 `approval_mode = "approve"`로 유지하는 것을 권장합니다.

## 6. 첫 작업

1. Excel, 한글 또는 AutoCAD에서 대상 문서를 엽니다. Excel 쓰기는 대상 통합문서가 이미 열려 있어야 합니다. 한글 파일을 창 없이 처리하려면 절대 파일 경로를 준비합니다.
2. AI에게 `core_get_status`로 먼저 상태를 읽게 합니다. Excel은 `connected:true`이고 `document`가 비어 있지 않을 때만 컨텍스트를 읽고, 한글은 `hwp_doctor`, 복잡한 Excel은 `excel_inspect(scope="diagnostics")`도 실행합니다.
3. AI가 `*_apply_ops`를 `dryRun: true`로 호출합니다.
4. diff, affected, warnings를 확인하고 승인합니다.
5. AI가 정확히 같은 ops와 confirmToken으로 실제 적용합니다.
6. `readback.verified`와 mismatches를 확인합니다.

Excel 쓰기는 활성 시트를 추정하지 않습니다. AI가 각 op에 정확한 시트를 넣도록 요청합니다.

Excel 경로만 전달해도 닫힌 파일을 자동 실행하지 않습니다. 사용자가 닫힌 파일을 열어 읽으라고 명시한 경우에만 `workbook` 절대 경로와 `allowOpenFile:true`를 함께 사용합니다. 쓰기에는 이 옵션을 쓰지 않습니다. DocBridge 실패를 `openpyxl`, `pywin32`/직접 COM, PowerShell COM 또는 `Start-Process`로 우회하도록 지시하지 마십시오.

```json
{
  "ops": [
    {
      "op": "set_values",
      "target": { "sheet": "고색-1구간" },
      "range": "B2",
      "values": [[1500]]
    }
  ],
  "dryRun": true
}
```

`range`를 `"'고색-1구간'!B2"`처럼 시트 한정 형식으로 써도 됩니다. `target.sheet`와 시트 한정 범위를 함께 쓰면 두 시트명이 반드시 같아야 합니다. 숫자는 Excel COM 규격에 맞춰 `Double`로 전달되므로 `Value2 = Int32` 캐스팅 오류가 발생하지 않습니다. 15자리보다 긴 번호·코드의 모든 자릿수를 유지하려면 JSON 문자열로 입력합니다.

병합과 숨김 상태는 수정 전에 `excel_read_range`의 `includeLayout:true`로 확인합니다.

```json
{
  "sheet": "고색-1구간",
  "range": "A1:H30",
  "includeFormulas": true,
  "includeLayout": true
}
```

기본 편집 op는 `merge_cells`, `unmerge_cells`, `set_rows_hidden`, `set_cols_hidden`,
`set_sheet_visibility`입니다. 예를 들어 12~14행을 숨길 때는 다음 batch를 먼저
`dryRun:true`로 실행하고, diff 승인 후 정확히 같은 ops와 `confirmToken`으로 적용합니다.

```json
{
  "ops": [
    {
      "op": "set_rows_hidden",
      "target": { "sheet": "고색-1구간" },
      "row": 12,
      "count": 3,
      "hidden": true
    }
  ],
  "dryRun": true
}
```

행·열을 다시 표시할 때는 `hidden:false`, 시트를 표시할 때는
`visibility:"visible"`을 사용합니다. 활성 시트와 마지막 표시 시트는 숨길 수 없습니다.
병합/병합 해제는 한 batch에서 단독으로 실행하고, visibility op도 값·서식·복사·병합 op와
별도 batch로 나눠야 작업 범위 스냅샷이 정확히 복구됩니다. 전체 예제와 제한은
[Excel 기본 편집 operations](docs/EXCEL-OPERATIONS.md)를 참고하세요.

권장 첫 요청:

```text
현재 연결 상태와 활성 문서를 읽어 줘. 아직 수정하지 말고 대상 documentRef를 먼저 보여 줘.
```

한글 예시:

```text
지금 열려 있는 한글 문서 전체 텍스트를 읽고, '1000원'을 '1200원'으로 바꾸는 dry-run만 보여 줘.
```

한글 창이나 탭이 여러 개일 때:

```text
hwp_get_active_context.summary.openDocuments로 열린 한글 문서를 전부 표로 보여 줘.
내가 지정한 문서의 documentRef를 읽기와 모든 쓰기 op에 계속 사용하고 현재 앞에 보이는 창이라고 추측하지 마.
경로 문자가 깨지거나 같은 파일이 중복되어 있으면 해당 항목의 instanceRef를 documentRef 값으로 사용해.
```

기존 양식 중간 삽입 예시:

```text
현재 열린 한글 문서에서 '관련근거 : 현장확인' 문구를 기준으로 바로 다음 새 문단에
'검토 결과: 추가 보완사항 없음'을 넣어 줘. 같은 문구가 여러 개면 먼저 개수를 세고
정확한 occurrence를 정한 뒤 dry-run → 적용 → 재열기 검증을 해 줘. 주변 문단 서식을 유지해 줘.
```

기존 창 작업은 선택한 `documentRef`를 사용하고 `file`을 생략합니다. 특정 파일을 백그라운드에서 처리할 때만 모든 op에 같은 절대 `file` 경로를 지정합니다. `documentRef`와 `file`은 동시에 쓰지 않으며 한 배치에 서로 다른 문서를 섞지 않습니다. 새 문서를 만들거나 전체 본문을 다시 작성할 때는 `replace_document_text`, 문서 끝 추가에는 `append_text`, 기존 양식 중간에는 `insert_before_text`/`insert_after_text` op를 사용합니다. 중간 삽입은 고유한 `anchor`를 쓰고 반복 문구면 1부터 시작하는 `occurrence`를 지정합니다. 배포 폴더의 `hwp-security\FilePathCheckerModuleExample.dll`은 한컴 공식 자동화 보안 모듈입니다. 설치 후 `powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\install-hwp-security.ps1`을 한 번 실행하면 `DocBridgeFilePathChecker` 전용 등록값이 유지됩니다. 제거할 때는 같은 폴더의 `uninstall-hwp-security.ps1`을 실행합니다. 한컴의 다른 등록값은 변경하지 않습니다.

## 7. 문제 해결

| 증상 | 확인 |
|---|---|
| tool이 보이지 않음 | 생성 JSON/TOML을 기존 설정에 올바르게 병합했는지 확인하고 클라이언트 재시작 |
| Cursor에서 `doc-bridge`가 보이지 않음 | `%USERPROFILE%\.cursor\mcp.json`의 `mcpServers.doc-bridge.command`가 현재 설치 EXE 절대 경로인지 확인하고 Cursor를 완전히 재시작합니다. `2-TEST.cmd`의 `Cursor global config`와 `MCP handshake`가 [OK]인지 확인합니다. |
| Cursor 프로젝트에서 다른 설정이 적용됨 | 프로젝트의 `.cursor\mcp.json`에 같은 이름의 서버가 중복됐는지 확인합니다. 설치기는 사용자 전역 파일만 병합하고 프로젝트 파일은 변경하지 않습니다. 한 위치의 구성만 사용하십시오. |
| Cursor에서 CAD 레이어가 비거나 상태가 안 보임 | 기본 `basic`의 `layers:[]`는 조회 생략이며 `layerSummaryStatus:"omitted"`로 표시합니다. `cad_query_entities(scope="layers")`로 전체 목록을 페이지 조회하십시오. `current`는 현재 작업, `on`은 켜짐, `freeze`는 동결, `locked`는 잠금입니다. `modelVisible`은 켜지고 동결되지 않은 상태이며 뷰포트별 동결·객체 투명도는 별도입니다. `null`은 조회 불가입니다. |
| CAD 문자가 마우스를 올릴 때만 보임 | 0.4.19는 편집 후 직접 ActiveX `Regen(acAllViewports)`를 수행합니다. `readback.displayRefresh.status`가 `failed`이면 좌표·축척 작업을 반복하지 말고 대상 도면 확인 후 `regen_document`만 dry-run → apply하십시오. 객체 색상·표시·투명도는 `includeGeometry:true`로 확인할 수 있습니다. API 갱신 성공은 육안 배치 검수가 아닙니다. |
| startup timeout | `dotnet run` 대신 `dist\doc-bridge-mcp.exe` 절대 경로 사용 |
| Excel 대상이 다름 | `excel_get_active_context.documentRef` 확인. 제한된 보기·모달 여부는 `excel_inspect(scope="diagnostics")`로 확인 |
| Cursor 작업 뒤 빈 회색 Excel 창이 남음 | 구버전이 workbook 없는 context를 실행 probe로 호출했거나 Cursor가 DocBridge 실패 뒤 `openpyxl`/`DispatchEx`/PowerShell COM으로 우회한 증상입니다. 최신판 설치 후 Cursor를 완전히 재시작합니다. 최신판은 status/context/dry-run으로 Excel을 만들지 않고, 명시적 `allowOpenFile:true` 읽기만 파일 열기를 허용하며 대체 COM·파일 덮어쓰기 우회를 금지합니다. |
| Excel이 활성 시트에만 씀 | 최신 배포본은 쓰기에서 활성 시트를 사용하지 않습니다. 각 op에 `target.sheet`를 넣거나 `range`를 `'시트 이름'!A1` 형식으로 지정하십시오 |
| `Value2 = Int32` 캐스팅 오류 | 최신 배포본은 모든 JSON 숫자를 Excel COM `Double`로 정규화합니다. 구버전을 제거한 뒤 최신 ZIP을 다시 설치하고 AI를 완전히 재시작하십시오 |
| Excel 창을 닫아도 `EXCEL.EXE`가 남고 추가기능이 사라짐 | 최신 배포본은 Excel COM을 전용 worker가 소유하고, AI가 강제 종료돼도 파이프 EOF에서 자신이 만든 Excel만 정상 회수합니다. 새 ZIP 설치 뒤 Codex·Claude·Kimi·Cursor를 완전히 종료했다 다시 시작하십시오. 기존 사용자 Excel에는 `Quit()`하지 않습니다. |
| 한글 실행 창에 연결 안 됨 | 먼저 `hwp_doctor`의 `state`를 확인. 기존 문서는 한글에서 표시하고 새 문서는 `hwp_launch({"newDocument":true})`를 한 번만 호출. 특정 파일은 절대 `file` 경로 사용 |
| 설치 검사에서 `HWP runtime update recommended` 경고 | DocBridge·TypeLib·자동화 환경은 정상이지만 한글 실행 파일이 권장 패치보다 오래된 상태입니다. 자동화 차단이 아니라 안정성 권고이며 `ownedAutomationBlocked:false`이면 계속 사용할 수 있습니다. 가능할 때 한컴 자동 업데이트 후 `2-TEST.cmd`를 다시 실행하십시오. |
| `PopupBorderImpl`/`TourPopup` `TypeInitializationException` | 오류 창에서 **아니요(N)**를 눌러 문서를 유지합니다. 먼저 0.4.10 이상을 설치하고 `hwp_doctor`의 `automationWindowsDirectory`, `automationEnvironmentPolicy`, `ownedAutomationBlocked`를 확인합니다. 0.4.10은 실제 오류창이 감지될 때만 동일 호출 재시도를 중단합니다. 환경 복구 후에도 오류가 남으면 [한컴 자동 업데이트 2024](https://help.hancom.com/hoffice130/ko-KR/HCell/introduction/update.htm)를 실행하십시오. |
| `MS.Internal.FontCache.Util`/`CultureFontManager` 오류 | 이벤트 로그의 실제 내부 예외는 `System.UriFormatException`입니다. WPF가 `windir + "\\Fonts\\"`로 절대 URI를 만들므로 축소된 AI 자식 환경에서 `windir`가 없으면 한글 UI 전체 초기화가 실패합니다. 0.4.10은 worker와 COM 활성화 순간에 검증된 `windir`/`SystemRoot`를 주입하고 한글 설치 `Bin`을 작업 폴더로 고정합니다. |
| 동일 한글 파일 중복 오류 | `summary.openDocuments`에서 원하는 창의 `instanceRef`를 `documentRef` 값으로 지정하거나 중복 창/탭을 닫고 다시 시도 |
| 한글 COM 시간 초과 | 한글 팝업을 닫습니다. 격리 worker는 자동 교체되므로 문서를 다시 읽고 새 위치 기준으로 재계획 |
| CAD RPC 거부 | AutoCAD가 명령/저장 중인지 확인 후 잠시 뒤 재시도 |
| PowerShell 스크립트 문자 오류 | 반드시 `-ExecutionPolicy Bypass -File` 방식 사용. 배포 패키지 무결성 검사는 모든 `.ps1/.psm1/.psd1`의 UTF-8 BOM을 확인합니다 |
| 외부 VBA `.bas` 한글이 깨짐 | 일반 DocBridge 편집은 VBA를 사용하지 않습니다. VBA 교환이 꼭 필요할 때만 `powershell.exe -ExecutionPolicy Bypass -File .\support\Convert-DocBridgeTextEncoding.ps1 -Mode ConvertBasToCp949 -Path .\모듈.bas`를 실행해 CP949/CRLF로 변환합니다 |
| 감사 기록 위치 | `%LOCALAPPDATA%\DocBridge\logs` |
| 스냅샷 위치 | `%LOCALAPPDATA%\DocBridge\snapshots` |

Kimi와 Claude 실행 파일은 이 PC에 현재 설치되어 있지 않아 실제 로그인 세션 테스트는 생략했습니다. 대신 양쪽 공식 stdio JSON 형식, 경로 escaping, MCP stdio/HTTP 프로토콜을 로컬에서 검증했습니다. Cursor는 2026-08-18 사용자 실환경에서 Excel·한글·AutoCAD 연결·읽기와 쓰기 dry-run이 확인됐으며 설치 테스트는 기존 전역 설정 보존과 프로젝트 설정 불변까지 검증합니다.
