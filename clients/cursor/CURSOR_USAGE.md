# Cursor에서 DocBridge 사용하기

## 자동 설치

배포 ZIP의 `1-INSTALL.cmd`는 다음 사용자 전역 설정에 `mcpServers.doc-bridge`만 병합합니다.

```text
%USERPROFILE%\.cursor\mcp.json
```

기존 최상위 설정과 다른 MCP 서버는 그대로 유지하고 원본은 `%LOCALAPPDATA%\DocBridge\backups`에 백업합니다. 설치 후 Cursor를 완전히 종료했다가 다시 실행하고 MCP 설정 화면에서 `doc-bridge`가 활성 상태인지 확인합니다.

첫 점검 요청:

```text
doc-bridge로 core_ping과 core_get_status를 실행하고 Excel, 한글, AutoCAD 연결 상태만 보여줘. 아직 수정하지 마.
```

## 빈 Excel 창 방지

Excel은 `core_get_status` 응답의 `apps.excel.connected`와 `apps.excel.document`를 먼저 판정합니다.

- `connected=true`이고 `document`가 비어 있지 않을 때만 `excel_get_active_context`를 한 번 호출합니다.
- 연결이 없거나 `document`가 비어 있으면 `excel_get_active_context`를 실행 확인용으로 호출하지 않습니다. 같은 실패를 반복 재시도하지 말고, 사용자가 통합문서를 연 뒤 새 `core_get_status`에서 문서가 보일 때만 다시 시도합니다.
- `allowOpenFile`의 기본값은 `false`입니다. 사용자가 닫힌 파일 읽기를 명시적으로 요청하고 기존 파일의 absolute workbook path를 제공한 경우에만 지원되는 읽기 도구에 `workbook`과 `allowOpenFile:true`를 함께 전달합니다. 일반 Excel 요청에서 이를 추론하거나 쓰기에 사용하거나 경로를 추측하지 않습니다. 쓰기 도구가 열린 workbook을 요구하면 빈 Excel을 띄우지 말고 먼저 해당 파일을 Excel에서 열어야 한다고 안내합니다.
- DocBridge Excel 호출이 실패하거나 기능 제약을 반환하면 original DocBridge error를 그대로 보여 주고 중단합니다. `openpyxl`, `pywin32` 또는 `DispatchEx("Excel.Application")`, PowerShell Excel COM, `Start-Process`, shell/UI automation으로 우회하거나 원본 workbook을 디스크에서 반복 overwrite하지 않습니다. 실제 상태 변화가 있거나 사용자가 지원되는 DocBridge 경로를 명시적으로 선택한 경우에만 다시 진행합니다.
- 경로 기반 작업 뒤 또는 `detail`이 `DocBridge가 생성한 인스턴스`이면 저장되지 않은 변경이 없음을 확인하고 `excel_disconnect`를 호출합니다. `detail`이 `사용자가 열어 둔 엑셀 창에 연결됨`인 경우 Excel을 종료하지 않습니다.

Cursor에는 다음처럼 요청할 수 있습니다.

```text
core_get_status의 apps.excel.connected와 apps.excel.document를 먼저 확인해줘.
둘 다 연결된 workbook을 가리킬 때만 excel_get_active_context를 한 번 호출하고,
아니면 빈 Excel을 실행하거나 같은 호출을 반복하지 말고 상태만 알려줘.
```

통합문서가 열린 것처럼 보이다가 회색 빈 Excel 창만 남으면 작업 완료가 아닙니다. 추가 호출을 중단하고 `core_get_status` 결과와 Cursor의 DocBridge 오류를 확인한 뒤, DocBridge가 만든 저장되지 않은 내용 없는 인스턴스만 `excel_disconnect`로 정리합니다.

## 전역 설정과 프로젝트 설정

- `%USERPROFILE%\.cursor\mcp.json`: 모든 Cursor 프로젝트에서 사용하는 사용자 전역 설정입니다. 설치기는 이 파일만 자동 병합합니다.
- `<프로젝트>\.cursor\mcp.json`: 해당 프로젝트에서만 사용하는 설정입니다. 설치기는 이 파일을 만들거나 변경하지 않습니다.

프로젝트 설정에 같은 이름의 `doc-bridge`가 있으면 어느 구성이 적용되는지 혼동될 수 있습니다. 특별한 이유가 없다면 전역 설정 하나만 사용하십시오. 프로젝트별 구성이 꼭 필요하면 [mcp.example.json](mcp.example.json)의 `command`를 현재 PC의 절대 설치 경로로 바꿔 프로젝트 설정에 직접 병합합니다.

Cursor 공식 문서: [MCP 설정](https://docs.cursor.com/context/model-context-protocol)

## 안전 규칙 적용

Cursor 사용자 규칙은 `Cursor Settings > Rules`에서 관리하는 일반 텍스트입니다. [docbridge-user-rule.txt](docbridge-user-rule.txt)를 복사해 사용자 규칙으로 넣으면 모든 프로젝트에 적용할 수 있습니다.

프로젝트와 함께 공유하려면 [docbridge-safe-automation.mdc](rules/docbridge-safe-automation.mdc)를 다음 위치에 복사합니다.

```text
<프로젝트>\.cursor\rules\docbridge-safe-automation.mdc
```

사용자 규칙과 프로젝트 규칙은 역할이 다릅니다. 설치기는 프로젝트 파일을 임의로 바꾸지 않으므로 규칙 복사는 선택 사항입니다. Cursor 공식 문서: [Rules](https://docs.cursor.com/context/rules)

설치 후 같은 템플릿은 다음 위치에도 복사됩니다.

```text
%LOCALAPPDATA%\DocBridge\generated-configs\cursor
```

## 반드시 지킬 쓰기 절차

1. 정확한 문서와 위치를 읽습니다.
2. `*_apply_ops`를 `dryRun=true`로 호출합니다.
3. diff와 영향 범위, confirmToken을 사용자에게 보여 주고 승인을 기다립니다.
4. 같은 ops와 confirmToken으로 `dryRun=false`를 실행합니다.
5. 변경 위치를 다시 읽어 검증합니다.

confirmToken은 5분 동안 한 번만 유효하고 정확한 대상과 ops에 묶입니다. 토큰 발급 뒤 ops를 바꾸거나 사용자가 같은 문서를 수정했으면 다시 읽고 새 dry-run을 만듭니다. Cursor와 Codex·Claude·Kimi가 동시에 연결되어 있어도 같은 문서를 동시에 편집하면 안 됩니다.

## 앱별 대상 지정

- Excel: 쓰기 op마다 `target.sheet` 또는 `'시트명'!A1`을 명시합니다. 활성 시트를 추측하지 않습니다. 병합·행/열 숨김 전에는 `excel_read_range(includeLayout=true)`로 현재 `mergedAreas`, `rowStates`, `columnStates`, `sheetVisibility`를 읽습니다.
- 한글: `hwp_get_active_context.summary.openDocuments`에서 대상의 `documentRef`/`instanceRef`를 선택합니다. `file`과 `documentRef`를 동시에 사용하지 않습니다.
- AutoCAD: `interaction.interrupted` 또는 `userActivityDetected`가 true이면 도면을 다시 읽고 남은 작업만 새 dry-run으로 만듭니다.

## Excel 병합·숨김 요청

Cursor에 다음처럼 요청하면 현재 레이아웃을 확인한 뒤 안전한 dry-run을 만들 수 있습니다.

```text
Excel의 정확한 documentRef와 시트 이름을 먼저 확인해줘.
공정표!A1:H30을 includeLayout=true로 읽고 병합 영역과 행·열 숨김 상태를 보여줘.
그 다음 B2:E2 병합과 12~14행 숨김은 서로 별도 dry-run으로 계획하고 아직 적용하지 마.
```

사용 가능한 기본 op는 다음과 같습니다.

- `merge_cells`, `unmerge_cells`: 한 batch에 하나만 사용합니다. 좌상단 외 값·수식이 없어야 병합할 수 있습니다.
- `set_rows_hidden`, `set_cols_hidden`: `hidden:true`로 숨기고 `false`로 다시 표시합니다.
- `set_sheet_visibility`: `visibility:"hidden"|"visible"`만 사용합니다. 활성 시트, 마지막 표시 시트, `veryHidden` 신규 설정은 차단됩니다.

visibility op끼리는 같은 batch에 묶을 수 있지만 값·수식·서식·복사·병합 op와 섞지 않습니다.
적용 후 같은 범위를 `includeLayout=true`로 다시 읽고 `readback.verified`와 함께 검증합니다.
세부 payload와 제한은 [Excel 기본 편집 operations](../../docs/EXCEL-OPERATIONS.md)를 참고하십시오.

## 대형 CAD 후속조회

`cad_get_active_context`는 대형 도면의 지연을 줄이기 위해 기본 `detailLevel=basic`에서 문서·배치·전체 개수 같은 가벼운 상태만 반환하고 레이어/엔티티를 순회하지 않습니다. 표본 통계가 필요할 때만 `detailLevel=summary`를 요청하고, 응답의 `nextActions`가 안내하는 `cad_query_entities` 후속 범위를 사용합니다.

Cursor에 다음처럼 요청할 수 있습니다.

```text
cad_get_active_context(detailLevel=basic)로 도면 상태와 nextActions를 확인해줘.
레이어 표본이 필요할 때만 detailLevel=summary로 다시 읽고,
전체 레이어는 cad_query_entities(scope=layers)로 조회한 다음
관로 문자 레이어와 신설 관로 객체 종류를 좁힌 다음 해당 레이어만 후속 조회해줘.
```

```text
도곽 세 개의 좌표 범위를 먼저 읽은 뒤 cad_query_entities(scope=regions)로 한 번에 비교해줘.
각 영역의 문자, 블록, 선 개수와 bbox를 표로 보여주고 아직 수정하지 마.
```

```text
이 좌표창 안의 관로만 cad_query_entities(scope=window)로 조회하고
레이어와 entityType을 함께 제한해줘. 결과가 잘렸으면 다음 페이지를 이어서 조회해줘.
```

## 제거

`3-UNINSTALL.cmd` 또는 다음 명령은 전역 Cursor 설정에서 `doc-bridge` 항목만 제거하고 다른 설정은 보존합니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Uninstall-DocBridge.ps1 -Clients Cursor
```
