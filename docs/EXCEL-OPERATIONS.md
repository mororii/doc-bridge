# Excel 기본 편집 operations

DocBridge는 Excel 화면을 클릭하거나 VBA를 실행하지 않고, 실행 중인 Excel의 ActiveX COM
객체를 통해 workbook을 직접 읽고 수정한다. 이 문서는 `excel_apply_ops`의 현재 작성·레이아웃·수명주기·데이터 객체 계약과
`excel_read_range(includeLayout:true)`의 안전 계약을 설명한다.

## 연결 사전검사와 파일 열기

먼저 `core_get_status({"app":"excel"})`의 `apps.excel`을 확인한다. `connected:true`이고
`document`가 비어 있지 않을 때만 `excel_get_active_context`를 한 번 호출한다. Excel이 닫혔거나
열린 workbook이 없으면 상태조회·컨텍스트·쓰기는 새 Excel을 만들지 않으며, 상태 변화 없이
같은 호출을 반복하지 않는다. ping·전체 앱 status·반복 context·전체 시트 읽기를 강제하지 않는다.

### 지정 창 고정 (instance pin)

여러 Excel이 열려 있으면 자동 선택이 엇갈릴 수 있다. 사용자가 지목한 창에서만
작업하려면 `excel_launch`로 핀을 박는다. 핀은 `%LOCALAPPDATA%\DocBridge`의
`excel-instance-pin.json`에 저장되어 CLI와 MCP가 공유한다.

- `{"processId": 1234}` 또는 `{"hwnd": 5678}`, 둘의 쌍으로 지정한다.
  `excel_get_active_context`의 `openWorkbooks[].processId`/`excelHwnd`를 그대로 쓴다.
- `{"activeWindow": true}`는 현재 포그라운드 Excel 창을 지정한다.
  (inline-ai의 포그라운드/HWND 추적과 같은 원칙. 최소화된 창은 hwnd가
  비어 보일 수 있어 복원 후 지정한다.)
- `{"clearPin": true}`는 핀을 지운다. 지정자·`dedicatedInstance`와 함께 쓸 수 없다.
- 핀이 있으면 이후 모든 호출이 그 창에만 붙는다. 핀 창이 사라지면 다른 창을
  몰래 쓰지 않고 `[EXCEL_PINNED_INSTANCE_GONE]`으로 중단한다.
- 핀 없이 호출하면 종전 자동 선택 그대로다. `dedicatedInstance`는 핀을 건드리지 않는다.

`excel_read_range`와 `excel_inspect`도 workbook 경로만으로 닫힌 파일을 자동으로 열지 않는다.
사용자가 닫힌 기존 파일을 열어 읽으라고 명시한 경우에만 존재하는 절대 경로와
`allowOpenFile:true`를 함께 쓴다. 이 옵션은 읽기 전용이며 Excel 쓰기에는 사용할 수 없다.
쓰기 대상은 먼저 Excel에서 열어야 한다.

## Formula trace (read-only)

`excel_inspect`의 `scope:"formula_trace"`는 전체 workbook을 스캔하지 않고 지정한 수식 셀에서
상류 의존 셀을 추적한다. 이 scope에는 기존 workbook routing의 명시 `workbook`, 명시 `sheet`,
그리고 연속 A1 `range`가 모두 필요하다. `maxDepth`(기본 4, 최대 20)와 `maxCells`(기본 500,
최대 5,000)는 항상 적용된다.

반환된 `nodes`는 실제 읽은 workbook/sheet/address/value/formula와 Formula2 읽기 상태를 담고,
`edges`는 실제로 읽은 대상 셀에만 연결한다. `coverage.complete:false`, `frontier`,
`truncationReasons`, `unresolved`는 제한·미지원 참조가 남았음을 뜻하며 완전한 영향 분석이라고
과장하지 않는다. 외부 workbook은 따라가지 않으며, `INDIRECT`/`OFFSET`, 3D, 표 구조 참조,
깨진 참조, 상수·표현식 이름은 reason과 함께 unresolved로 남긴다. 상세 계약은
[EXCEL-FORMULA-TRACE.md](EXCEL-FORMULA-TRACE.md)를 따른다.

DocBridge 오류나 제약을 `openpyxl`, `pywin32`/직접 Excel COM, PowerShell Excel COM,
`Start-Process` 또는 UI 자동화로 우회하지 않는다. 도구 오류와 필요한 사용자 조치를 보고하고
중단한다. 우회 파일 재작성은 기존 서식·매크로를 손상할 수 있고, 직접 COM 인스턴스는 열린
통합문서가 없는 회색 Excel 창이나 잔류 `EXCEL.EXE`를 만들 수 있다.

## 지원 범위

아래는 현재 `excel_apply_ops` 가족이다. 기본 5개 op만 있는 제품이 아니다.
`core_get_capabilities({"app":"excel"})`의 `writeOps`가 권위 목록이다.

| 가족 | 대표 op | 비고 |
| --- | --- | --- |
| 값/수식 | `set_values`, `set_formulas`, `clear_range`, `fill_range`, `auto_fill` | JSON 타입 유지. Formula2는 opt-in. `""`는 빈 칸 |
| 병합 | `merge_cells`, `unmerge_cells` | 같은 종류만, 비겹침 batch. 31개월 머리글은 merge-only 한 batch |
| 서식/레이아웃 | `format_range`, `set_row_heights`, `set_column_widths`, `freeze_panes`, `set_page_setup`, `set_page_breaks`, `set_view` | `set_view`는 대상 시트를 활성화해 읽고 원래 활성 시트·선택·스크롤을 되돌린다 |
| 표시 | `set_rows_hidden`, `set_cols_hidden`, `set_sheet_visibility` | visibility-only batch; `veryHidden` 포함, 마지막 표시·활성 시트 보호 |
| 구조 | `insert_rows`/`cols`, `delete_rows`/`cols`, `add_sheet`, `copy_sheet`, `move_sheet`, `rename_sheet`, `delete_sheet` | `delete_sheet`는 다른 시트 수식·이름·차트·피벗 의존을 캡처. 표 참조·같은 책 `[owned.xlsx]Sheet`·3D는 유지. 못 캡처하면 삭제 전 거절 |
| 수명주기 | `create_workbook`, `open_workbook`, `save_workbook`, `close_workbook`, `export_pdf` | 소유 Application 자동 Quit는 빈 컬렉션만 |
| 데이터 객체 | 표, 이름, 유효성, 조건부 서식, 차트, 그림, 메모, 링크, `create_pivot`/`update_pivot`/`refresh_pivot`/`delete_pivot` | 같은 통합문서 피벗은 지원. 피벗 **캐시/외부 연결**만 범위 밖 |
| 통합문서·고급 개체 | 외부 링크 3종, 계산 모드, 값 고정, 선택 붙여넣기, 목표값 찾기, 통합문서 보호 2종, 창 분할, 스파크라인 3종, 슬라이서 2종, 셀 스타일 | [EXCEL-WORKBOOK-OPS.md](EXCEL-WORKBOOK-OPS.md); VBA·매크로는 범위 밖 |

모든 Excel 쓰기는 활성 시트를 추정하지 않는다. `target.sheet`와 시트 한정 범위를 함께
사용하면 두 시트명이 정확히 같아야 한다. 여러 workbook이 열려 있으면 먼저
`excel_get_active_context`의 `documentRef`를 확인하고, 필요할 때 op에
`target.workbook` 또는 `targetWorkbook`을 명시한다.

## 적용 전 레이아웃 읽기

병합·숨김 상태를 바꾸기 전에는 대상 범위를 레이아웃과 함께 읽는다.

```json
{
  "workbook": "현장일보.xlsx",
  "sheet": "공정표",
  "range": "A1:H20",
  "includeFormulas": true,
  "includeLayout": true
}
```

`includeLayout:true`이면 `layout`에 다음 정보가 추가된다.

- `sheetVisibility`: `visible`, `hidden`, `veryHidden` 중 현재 상태
- `rowStates`: 요청 범위에 포함된 행 번호와 `hidden` 상태
- `columnStates`: 열 문자·열 번호와 `hidden` 상태
- `mergedAreas`: 요청 범위에서 확인한 병합 영역 주소
- `coverage`: 요청·반환 행/열 수와 병합 영역 스캔의 완전성

읽기는 `maxReadCells` 한도 안에서 반환된다. `coverage.complete:false`이면 반환된 일부 상태만
보고 전체 범위라고 단정하지 말고 더 작은 범위로 나누어 다시 읽는다.

## 범위 읽기 페이지

`excel_read_range`는 요청 `range` 전체를 먼저 `Value2`/`Formula` 배열로 만들지 않는다. 각 호출은
기본 최대 10,000 셀(더 작은 `maxCells` 지정 가능)의 직사각형 페이지 하나만 COM에서 읽는다.
`rowOffset`과 `columnOffset`은 요청 range 안의 0-based 위치이며, `maxRows`/`maxColumns`로 더 작은
페이지를 정할 수 있다. 응답의 `range`와 `requestedRange`는 원래 요청, `returnedRange`는 실제 읽은
주소다. `coverage.complete`는 이 응답 하나가 원래 요청 전체를 포함할 때만 true다.
`coverage.hasMore`는 다음 페이지 존재 여부이고, `coverage.continuation`이 있으면 다음 요청에 그대로
합친다. `rowOffset`/`columnOffset`도 coverage에 반환한다. 가로 타일은 한 행씩 진행해 마지막 좁은
타일에서도 셀을 빠뜨리거나 중복하지 않는다.

```json
{
  "sheet": "공정표",
  "range": "A1:XFD500",
  "maxCells": 10000,
  "includeFormulas": true,
  "formulaMode": "formula2"
}
```

`formulaMode`는 `includeFormulas:true`일 때만 사용한다. 기본 `formula`는 기존 `Range.Formula`와
호환되고, 동적 배열 의미론은 명시적으로 `formula2`를 요청한다. `formula2`를 요청한 COM 읽기가
실패하면 `Formula`로 조용히 대체하지 않고 오류를 반환한다. 여러 area로 이루어진 비연속 range는
지원하지 않으며 명확히 거절한다.

## 쓰기 경로

일반적인 사용자 편집 요청은 그 범위의 승인이다. 클라이언트 권한 UI와 서버 사전 검사,
사람의 범위 승인을 구분한다. 모든 쓰기마다 재승인 질문을 요구하지 않는다.
`highRiskConfirm`은 권한 UI를 대체하거나 사람 승인을 증명하지 않는다.

`set_values`, `set_formulas`, `format_range`는 `autoExecuteOps`다. 저장된 통합문서의
절대 경로가 있으면 한 호출로 적용한다. 현재 어댑터는 인스턴스 바인딩 참조를 발급하지 않는다.

```json
{
  "ops": [
    {
      "op": "set_values",
      "target": { "sheet": "공정표" },
      "range": "B2",
      "values": [[1500]]
    }
  ],
  "executionMode": "execute",
  "requestId": "11111111-1111-1111-1111-111111111111",
  "expectedDocumentRef": "C:\\작업\\현장일보.xlsx"
}
```

`requestId`는 호출마다 새 UUID다. 같은 작업의 통신 재시도만 같은 UUID를 쓴다.
같은 UUID·앱·문서·ops는 저장된 결과를 재전송하고 다른 payload는 거절한다.
`dryRun`/`confirmToken`/`highRiskConfirm`은 false라도 넣지 않는다.
`Book1` 같은 미저장 이름은 경로를 만들려고 자동 저장하지 않고 아래 토큰 경로를 쓴다.
`outcomeUnknown`이면 새 UUID로 재실행하지 말고 문서를 먼저 확인한다.

미리보기, 고위험, 병합·숨김·복사·구조 변경은 기존처럼 원하는 op를 `dryRun:true`로 호출한다.

```json
{
  "ops": [
    {
      "op": "set_rows_hidden",
      "target": { "sheet": "공정표", "workbook": "현장일보.xlsx" },
      "row": 12,
      "count": 3,
      "hidden": true
    }
  ],
  "dryRun": true
}
```

응답의 `diff`, `affected`, `warnings`, `snapshotId`를 확인한다. 승인된 경우에만 op를 한 글자도
바꾸지 않고 반환된 토큰으로 적용한다.

```json
{
  "ops": [
    {
      "op": "set_rows_hidden",
      "target": { "sheet": "공정표", "workbook": "현장일보.xlsx" },
      "row": 12,
      "count": 3,
      "hidden": true
    }
  ],
  "dryRun": false,
  "confirmToken": "conf_직전_dry_run_토큰"
}
```

적용 뒤 `readback.verified`와 `rollback.verified`를 확인한다. `readback.verified:true`일 때만
같은 범위를 `includeLayout:true`로 다시 읽어 완료로 본다. 보호된 시트에서 Bold 등 COM
쓰기가 거절되면 자동 롤백도 같은 보호 때문에 쓰지 못하고 `rollback.verified`는 false다.
모든 적용 실패가 검증된 롤백으로 끝난다고 보지 않는다.
문서나 op가 달라졌거나 토큰의 5분 유효시간이 지났으면 기존 토큰을 재사용하지 않는다.

확인 토큰은 객체 키 순서가 달라도 같은 의미의 ops이면 동일하게 바인딩한다. JSON 숫자는
어휘 계수·지수로 정규화한다. `1`/`1.0`/`1e0`은 같고, `1e-100`은 `0`과 다르며, 긴 소수와
거대 지수는 반올림하거나 0으로 접히지 않는다. 관리 객체의 NaN/Infinity는 거부한다.
파서가 `1e100`처럼 큰 유한 숫자를 float/double로 넘치게 변환한 경우는 거부가 아니다.
이 계약은 RFC 8785 준수를 주장하지 않는다.

적용 실패는 빈 `errors`와 batch-fallback만으로 끝내지 않는다. 각 op의
`operationResults`에 단계, 경과 시간, 예외 형식, HRESULT를 남기고, 실패한 op 이후는
`stage:skipped`로 표시한다. `readback.mismatches`가 있으면 `errors`에도 복사한다.
COM 데이터 readback 성공과 화면에서 보이는 조건부 서식·인쇄 배치는 다른 증거다.

같은 PC의 다른 DocBridge 프로세스가 Excel/한글/CAD를 쓰는 동안에는 전역 자동화 잠금으로
직렬화된다. 이는 동시 편집을 허용하는 기능이 아니다.

Excel worker의 상태/컨텍스트/범위 읽기(discovery)는 45초다. `validatePreviewReuse`는
대상 서식 전체를 다시 읽어 지문하므로 apply와 같은 150초 예산이다. STA COM 한도는
120초로 그대로 둔다.

v3 format-only는 `styleScope=written-properties`다. 실제로 쓰는 속성과 색 커플링만
스냅샷·지문·복구한다. Bold-only는 다른 색·tint를 읽거나 복구하지 않는다. 균일
bool/글꼴 크기/`NumberFormat`은 범위 빠른 경로를 쓰고, 혼합·색은 필요한 셀 상태를
확보한 뒤 복구 후 확인한다. 기존 v2 전체 서식 스냅샷도 복구한다. 오래된 format
preview 토큰은 새 dry-run이 필요할 수 있다. `MaxFormatSnapshotCells`는 100,000,
STA COM 한도는 120초로 그대로다.

검사 도구 표본과 한계는 [PERFORMANCE.md](PERFORMANCE.md)에 있다. 균일 Bold와
혼합 색/서식은 다른 작업이며 만능 속도 배수를 만들지 않는다. 보호된 시트에서 Bold
COM이 거절되면 `rollback.verified=false`를 숨기지 않는다.

## 1. 셀 병합

```json
{
  "ops": [
    {
      "op": "merge_cells",
      "target": { "sheet": "공정표" },
      "range": "B2:E2"
    }
  ],
  "dryRun": true
}
```

Excel은 병합 범위의 좌상단 셀 값만 유지한다. DocBridge는 데이터 손실을 막기 위해 좌상단 외
셀에 값이나 수식이 하나라도 있으면 `[EXCEL_MERGE_WOULD_DELETE_CONTENT]`로 거부한다. 기존 병합 영역과
부분적으로 겹치는 범위도 거부하며, 셀별 서식까지 정확히 스냅샷·복원할 수 있도록 한 번에 최대
2,000셀까지만 분석한다.

읽기·ClearContents·읽기 실패는 `[EXCEL_MERGE_UNPROVEN_CELL]`로 Merge 전에 중단한다.
이미 진짜 빈 칸(Value2가 null/DBNull)은 건드리지 않는다. 좌상단 밖 상수 빈 문자열만
ClearContents한 뒤, 그 Value2가 진짜 빈 칸인지 다시 본다. Formula는 `""`여도 된다.
ClearContents가 no-op이라 Value2가 `""`로 남으면 Merge하지 않는다. 이 검사는
`set_values`의 `""`≈빈 칸 동등과 다르다. 네이티브가 빈 문자열을 남기는 경우가 있어
후자를 쓰면 확인된 병합 경고가 다시 날 수 있다.

`merge_cells`와 `unmerge_cells`는 서로 섞거나 값/서식과 한 batch에 넣을 수 없다.
같은 종류의 비겹침 범위는 한 batch에서 최대 400개까지 허용한다. 겹치면
`[EXCEL_MERGE_BATCH_OVERLAP]`으로 배치 전체가 거부된다. 31개월 머리글 쌍
(`G5:H5`, `I5:J5`, … 총 31개 인접 두 열)은 하나의 merge-only batch다. 스냅샷
`restoreMode`는 `merge-state`다. 한 개 op는 v1 봉투를 유지하고, 여러 개는 v2
`entries[]`를 역순 복구한다.

## 2. 병합 해제

```json
{
  "ops": [
    {
      "op": "unmerge_cells",
      "range": "'공정표'!B2:E2"
    }
  ],
  "dryRun": true
}
```

대상은 한 병합 영역 안의 단일 셀이거나 해제할 병합 영역 전체를 포함해야 한다. 여러 병합 영역을
일부만 걸치는 다중 셀 범위는 `[EXCEL_UNMERGE_PARTIAL_OVERLAP]`으로 거부한다. 적용 후 병합 영역과
좌상단 값·수식을 다시 읽어 검증하며, 실패하면 `merge-state` 스냅샷으로 자동 복구한다.

## 3. 행 숨김과 표시

행 12~14를 숨기는 예:

```json
{
  "ops": [
    {
      "op": "set_rows_hidden",
      "target": { "sheet": "공정표" },
      "row": 12,
      "count": 3,
      "hidden": true
    }
  ],
  "dryRun": true
}
```

다시 표시할 때는 같은 범위에 `hidden:false`를 사용한다. `row`와 `count`는 Excel 행
1~1,048,576 안에 있어야 한다.

## 4. 열 숨김과 표시

열 D~F를 숨기는 예:

```json
{
  "ops": [
    {
      "op": "set_cols_hidden",
      "target": { "sheet": "공정표" },
      "col": "D",
      "count": 3,
      "hidden": true
    }
  ],
  "dryRun": true
}
```

`col`은 `D` 같은 열 문자 또는 1부터 시작하는 열 번호를 받는다. `count`를 포함한 최종 열은
Excel의 마지막 열 XFD(16,384)를 넘을 수 없다. 다시 표시할 때는 `hidden:false`를 사용한다.

## 5. 시트 숨김과 표시

```json
{
  "ops": [
    {
      "op": "set_sheet_visibility",
      "target": { "sheet": "검토용" },
      "visibility": "hidden"
    }
  ],
  "dryRun": true
}
```

다시 표시할 때는 `visibility:"visible"`을 사용한다. 다음 경우는 적용 전에 차단한다.

- 현재 활성 시트를 숨기는 작업
- workbook의 마지막 표시 시트를 숨기는 작업
- workbook 구조가 보호된 상태에서 시트 표시 상태를 바꾸는 작업
- `veryHidden`을 새로 설정하는 작업

`veryHidden`은 일반 Excel UI에서 사용자가 직접 해제할 수 없으므로 이 계약에서는 생성하지
않는다. 다만 기존 문서에 있던 `veryHidden` 상태는 읽기와 스냅샷 복구에서 원래 값으로 보존한다.

## 6. 범위 서식 (`format_range`)

```json
{
  "ops": [
    {
      "op": "format_range",
      "target": { "sheet": "공정표" },
      "range": "B2:B4",
      "style": { "bold": true, "fillColor": "#FFFF00" }
    }
  ],
  "dryRun": true
}
```

쓰기·읽기 별칭을 받는다. `bold`/`fontBold`, `italic`/`fontItalic`,
`fillColor`/`interiorColor`/`fill`은 같은 속성이다. `fontSize`, `numberFormat`,
`fontColor`도 허용한다. 별칭이 서로 다른 값을 가리키거나, 미지원 키·null·잘못된 색이
있으면 snapshot/토큰 전에 거부한다. 색은 OLE 0..16777215 또는 `#RRGGBB`이다.

dry-run `diff.before`는 `"current"` 문자열이 아니라 대상 범위의 현재 서식 요약이다.
`excel_read_range(includeStyles:true)`는 쓰기 키와 읽기 별칭을 함께 반환하고
`fillPattern`/`fillColorIndex`를 포함한다. 채움 없음은 `Pattern=none` 상태이며
`Color=0`(검정)과 같지 않다.

`format_range`만 있는 배치는 대상의 쓰는 서식만 스냅샷한다. 무관한 `UsedRange` 값/수식을
훑지 않는다. v3 지문은 스코프에 들어온 속성만 해지하므로, Bold-only preview는 채움만
바뀌어도 그대로일 수 있다. 스코프 안 서식이 바뀌면 기존 토큰을 재사용하지 않고
거부한다(`freshPreviewAllowed:false`). 값·수식·구조가 섞인 배치는 전 문서 지문을
쓰지 않고 새 preview를 받는다(`freshPreviewAllowed:true`).
혼합 배치에 대해 재사용 가능하다고 가정하지 않는다.

다음 경우는 쓰기 전에 차단한다.

- 대상 밖으로 뻗는 부분 병합 (`[EXCEL_FORMAT_PARTIAL_MERGE]`)
- 셀 Font/Interior가 `null`/`DBNull`인 혼합·리치 텍스트 (`[EXCEL_FORMAT_MIXED_RICHTEXT]`)
- ThemeColor 조회가 Excel 1004(테마 미적용)가 아닌 COM 단절/사용 중/알 수 없는 실패
  (`[EXCEL_FORMAT_THEME_READ_FAILED]`)
- 테마가 아닌 font/fill/pattern RGB에 0이 아닌 tint가 있는 셀
  (`[EXCEL_FORMAT_UNSUPPORTED_RGB_TINT]`). Color 후 Tint를 다시 주면 음영이 이중 적용된다.
  테마에 연결된 tint는 허용한다. RGB tint를 역산하거나 네이티브 복사로 우회하지 않는다.

테마·자동색의 실측 복원은 COM 속성과 화면 렌더를 따로 기록해야 하며,
데이터 readback만으로 육안 품질을 보장하지 않는다.

보호된 시트에서 `Font.Bold` 설정이 COM으로 거절되면(`Font 클래스 중 Bold 속성을 설정할 수
없습니다`) 스냅샷 롤백도 같은 보호 셀에 쓰지 못한다. 이 경우의 통과는 오류를 숨기지 않고
`rollback.verified=false`를 보고한 것이지, 보호 시트에서도 롤백이 성공했다는 뜻이 아니다.

DocBridge가 만든 저장된 통합문서는 `excel_disconnect`와 worker 파이프 정상 종료에서
회수된다. 호스트 크래시 뒤에는 빈 인스턴스와 저장된 소유 통합문서 모두에서 `EXCEL.EXE`가
남는 것이 관측되었다. 잔류 프로세스가 해결되었다고 보지 않는다. 사용자가 연 Excel은
강제 종료하지 않는다.

## 스냅샷 보조 workbook 사본

스냅샷 디렉터리에는 `state.json`과 별도로 **보조 workbook 사본**이 남는다. 자동 롤백은 `state.json`의 operation-scoped 상태가 담당하며, 이 사본을 읽어 자동 복원하는 경로는 없다. 운영자가 사후에 참고하는 증거 파일이다.

기본값은 마지막으로 저장된 디스크 파일 복사다. 저장 전 변경이 있는 workbook이라면 그 변경은 사본에 없다. 사본이 최신인지는 추측하지 말고 metadata를 읽는다.

| metadata 키 | 뜻 |
|---|---|
| `workbookBackupSource` | `last-saved-file`(기본) / `current-memory-savecopyas`(opt-in) / `none` |
| `workbookBackupFresh` | 저장 전 메모리 상태를 담은 사본일 때만 `true` |
| `workbookBackupAvailable` | 사본 파일이 실제로 생겼는지 |
| `workbookBackupSavedFlag` | 캡처 시점 `Saved`. 읽지 못하면 `null`이며 `false`로 바꾸지 않는다 |
| `workbookBackupReason` | 이 출처를 고른 이유 문장 |
| `workbookBackupFreshCopyEnabled` | opt-in 환경변수가 켜져 있었는지 |
| `workbookBackupFileFormat` | opt-in이 켜졌을 때만 기록한다. 키가 없으면 "읽지 않음"이며, `null`인 "읽었지만 읽히지 않음"과 다르다 |

`DOCBRIDGE_EXCEL_FRESH_WORKBOOK_BACKUP=1`을 설정하면(정확히 `1`만 활성) 일반 `.xlsx`(`FileFormat` 51)이고 `Saved`가 읽히며 false인 workbook에 한해 `SaveCopyAs`로 저장 전 상태를 복사한다. 이때 workbook 전체가 직렬화되므로 편집 크기와 무관하게 workbook 크기에 비례하는 비용이 든다. 켜기 전에 [PERFORMANCE.md](PERFORMANCE.md)의 비용과 안전 조건을 확인한다.

opt-in을 켠 경우에도 워크북을 저장·닫기·활성화하지 않고 `Saved`나 경로를 되돌려 쓰지 않으며 `DisplayAlerts`를 건드리지 않는다. 복사 전후의 `FullName`, `Saved`, 소유 인스턴스의 열린 workbook 수를 읽어 비교한다. 복사 전에 상태를 모두 읽지 못하면 최신 사본을 시도하지 않고 마지막 저장 파일을 복사하며 그 이유를 기록한다. 복사를 시도한 뒤 상태가 보존됐음을 확인하지 못하면 스냅샷을 거절해 편집을 막는다. 이미 있는 대상 파일은 어느 경우에도 덮어쓰지 않는다.

## 지연 서식 체크포인트 (opt-in, 실험적 확대 중)

기본값은 꺼져 있다. `DOCBRIDGE_EXCEL_DEFERRED_FORMAT_SNAPSHOT=1`(정확히 `1`만 활성)일 때, 그리고 **execute 경로**일 때만 후보가 된다. dry-run·confirm 재사용·`core_create_snapshot`은 host가 execute 컨텍스트를 찍지 않으므로 절대 지연 경로로 가지 않는다. 이는 dry-run이 지연 스냅샷을 만들면 이후 confirm이 `freshPreviewAllowed=false`로 거절되기 때문이며, 환경변수만으로 추론하지 않는다.

무엇을 바꾸는가: 적용 전 서식 상태를 **어디서 읽는지**만 바꾼다. 무엇을 복원하는지, 무엇을 검증된 롤백으로 볼지, 기본 동작과 dry-run 의미는 그대로다. 비싼 셀별 COM 읽기를 캡처 시점이 아니라 **롤백이 실제로 필요할 때**로 미루고, 그 대신 캡처 시점에 `SaveCopyAs`로 저장 전 상태의 체크포인트 사본을 남긴다.

### 자격 조건

요청 단계(COM 없음)와 워크북 단계(라이브 COM), 그리고 체크포인트 **바이트** 단계를 모두 통과해야 한다. 자격 거절은 세 단계 모두 apply보다 앞에 있으므로 **워크북에 아무것도 쓰기 전에** 기존 eager format-only 경로로 되돌아가며, 부분 적용 뒤에 경로가 바뀌는 일은 없다. 거절 사유는 스냅샷 `metadata.json`의 `deferredEligibility`(`eligible=false`, `used=false`, `code`, `reason`)에 남는다.

다만 **모든 실패가 대체 경로로 가는 것은 아니다.** 기준은 `SaveCopyAs`를 **시도했는지**다.

- 복사 **전에** `FullName`·`Saved`·열린 워크북 수를 다 읽지 못하면 `identity` 자격 거절이다. 사본을 만들지 않은 채 기본 eager 경로로 되돌아간다. 검증할 수 없는 사본을 만들지 않기 위한 보수적 선택이며, 여기까지는 평범한 대체 경로다.
- 복사를 **시도한 뒤**에는 같은 값을 다시 읽어 이전과 비교한다. 이 비교로 워크북이 그대로임을 확인하지 못하면 — 값이 달라졌든, 이번에는 읽히지 않든 — 대체 경로로 내려가지 않고 스냅샷 자체를 거절해 그 편집이 진행되지 않게 한다. `SaveCopyAs`가 실패했거나 파일을 만들지 못한 경우도 마찬가지로, 확인이 되지 않으면 거절이다. 신원을 확인하지 못한 워크북에 쓰기를 이어 가지 않기 위해서다.
- 비교가 "그대로"임을 확인했는데 `SaveCopyAs`만 실패했거나 파일이 없으면, 그때는 `identity` 자격 거절로 기본 경로로 되돌아간다.

| 단계 | 거절 code |
|---|---|
| 요청 | `policy`, `host-context`, `ops`, `sheet-range`, `style`, `cell-count` |
| 워크북 | `path`, `file-format`, `file-size`, `macros`, `protection`, `password`, `links`, `identity` |
| 체크포인트 바이트 | `encrypted-ooxml`, `ole2-compound-file`, `xlsb-binary`, `unknown-container`, `bad-zip`, `duplicate-package-part`, `strict-ooxml`, `macro-enabled-main`, `vba-project`, `xlm-macrosheet`, `external-links`, `connections`, `ole-or-activex`, `formulas`, `defined-names`, `conditional-formatting`, `data-validation`, `merge-cells`, `rich-text`, `rgb-tint`, `full-calc-on-load`, `complex-part`, `missing-content-types`, `missing-workbook-xml`, `relationship-target`, `dtd-prohibited`, `malformed-xml`, `uncompressed-xml`, `file-size` |

- 단일 `format_range` 배치, 단일 직사각형 A1 대상, `target.sheet` 명시, `fillColor` 포함이어야 한다. Bold 전용처럼 `fillColor`가 없는 배치는 기존 균일 빠른 경로가 이미 싸므로 지연 대상이 아니다.
- 대상 셀 수는 1,000..5,000, 원본과 체크포인트는 16MiB(16,777,216바이트) 이하여야 한다. 최적 교차점 주장이 아니라 보수적 롤아웃 한계다.
- 로컬 `.xlsx`(`FileFormat` 51)만이며 UNC/URI, 매크로, 보호, 암호, 외부 링크는 거절한다. 암호가 걸린 워크북은 `SaveCopyAs` 사본도 암호가 걸리고, 그것을 사용자의 Excel에서 다시 열면 **모달 암호 창**이 뜬다. 그래서 바이트 단계에서 먼저 막는다.
- **수식은 전부 거절한다.** 대상 범위 안이든 밖이든, 다른 시트에 있든, 값이 캐시되어 있든 상관없다. `<f>` 요소를 만나면 `formulas`이고, `xl/calcChain.xml`·`xl/tables/`·`xl/volatileDependencies.xml`처럼 수식을 전제하는 파트가 있어도 `formulas`다. 워크북에 수식이 하나라도 있으면 이 경로는 후보가 아니며, 서식 편집 자체가 수식과 무관해도 마찬가지다.
- **plain-data 파트 어휘 밖의 파트도 전부 거절한다**(`complex-part`). 허용 목록은 `[Content_Types].xml`, `_rels/.rels`, `docProps/core|app|custom.xml`, `xl/workbook.xml`, `xl/_rels/workbook.xml.rels`, `xl/styles.xml`, `xl/sharedStrings.xml`, `xl/theme/theme1.xml`, `xl/worksheets/sheetN.xml`(+ 그 `_rels`), `xl/printerSettings/printerSettingsN.bin`이며, **이 목록에 없는 것은 모두 거절**이다. 파트 개수 상한을 넘겨도 같은 `complex-part`다.
- 특히 웹 추가 기능(Office 추가 기능) 파트 `xl/webextensions/taskpanes.xml`과 `xl/webextensions/webextensionN.xml`은 허용 목록에 **없으므로 거절 대상**이다. 추가 기능이 **화면에 보이지 않거나 숨겨져 있어도**, 사용자가 그것을 쓰지 않아도, 파트가 패키지에 들어 있으면 거절이다.
- 새 워크북에 이 파트가 붙는지는 **Excel 시작 상태에 따라 달라진다**. 같은 장비에서 `Workbooks.Add`가 webextension 파트 3개를 넣은 실행과 하나도 넣지 않은 실행이 모두 관측되었다. 그래서 "새 워크북이면 통과한다"고 가정하지 않고, 판단은 항상 실제 체크포인트 바이트로 한다.
- v1에서는 병합·서식 있는 텍스트·미지원 RGB+tint·이름 정의·조건부 서식·데이터 유효성을 포함한 체크포인트도 각자의 code로 거절한다.
- 셀 한도 100,000과 STA 120초는 올리지 않았다.

### 구버전 호환 봉투(envelope)

`state.json`은 `restoreMode=copy-sheet-topology`, `snapshotVersion=4`, `payloadMode=deferred-format-only`의 **정확한 3-튜플**로 기록한다. 새 디코더는 이 조합을 일반 topology 복원보다 **먼저** 정확히 매칭해 라우팅하고, 3-튜플이 불완전하면 일반 topology로 흘려보내지 않고 거절한다.

이 조합인 이유는 하나다. 설치본 0.4.18에는 format-only 디코더 자체가 없어서 `restoreMode=format-only`로 적으면 legacy 분기로 떨어지고 0셀 복원을 성공으로 보고할 수 있다. 반면 0.4.18과 0.4.20 모두 copy-sheet-topology는 디코드하며, 쓰기 전에 `snapshotVersion`을 먼저 비교하므로 4를 **변경 없이 거절**한다. 이 거절은 설치본 0.4.18, 로컬 동결 후보, GitHub에 게시된 v0.4.20 자산 세 바이너리에서 실제로 실행해 확인했다. 근거는 아래 「검증 상태」에 있다.

### 복원 시 동작

체크포인트 경로는 `state.json`에 적힌 경로를 따라가지 않고 `snapshotDir`에서 다시 조립한다. 기록된 파일 이름이 다르면 **거절**이지 재지정이 아니다. 해시를 `FileShare.Read` 핸들을 쥔 채 확인하고, 읽기 전용(`UpdateLinks=0`, `AddToMru=false`, `Notify=false`)으로 연 뒤 기존 `CaptureFormatOnlyState`로 서식을 읽고, 체크포인트만 닫고 **열기 전에 활성이던 워크북**과 워크북 수를 되돌린 다음 기존 scoped 복원을 라이브 워크북에 적용한다. `DisplayAlerts`는 건드리지 않는다.

추출이 실패하면 `ok=false`이고 `readback.verified`가 없으며 체크포인트 파일은 남긴다. `deferredExtraction`으로 "롤백이 돌았지만 불일치"와 "롤백을 재구성조차 못함"을 구분한다. host의 `IsVerifiedRestore`에는 지연 전용 예외를 추가하지 않는다.

### 한계 (문서화된 사실)

- **크래시 복구가 아니다.** 체크포인트는 파일이라 오래 남지만, 지연 추출은 Excel이 살아 있고 원본 워크북이 기록된 `documentRef`로 여전히 열려 있어야 동작한다. host 프로세스는 새로 떠도 되지만 Excel이 다시 뜬 상황은 복원할 수 없다. `deferredRestoreRequiresLiveSession=true`가 이를 명시한다.
- **롤백이 느려진다.** 캡처가 싸지는 대신 실제 롤백은 추출+복원 비용을 낸다. 롤백이 드물다는 데 거는 선택이다.
- **롤백 페이로드가 캡처 시점에 완성되어 있지 않다.** 추출이 실패하면 "적용 실패 + 롤백 없음"이 될 수 있고, 그 경우 host는 기존 문구로 검증 실패를 보고한다.

### 검증 상태 (2026-09-10)

측정 대상 `DocBridge.Core.dll`은 SHA-256 `357D981F2F11F0F41D424657D04F645200E4B80EC3485AD4E35A2DAE1664E248`이다. CLI 출력 디렉터리와 테스트 출력 디렉터리에 같은 해시를 가진 **동일한 사본**이 들어 있다는 뜻이며, CLI 실행 파일과 테스트 어셈블리는 각자 다른 해시를 가진 별개 파일이다. 이 빌드의 제품 회귀는 Core 631/631, MCP 20/20이다. 아래는 root가 실제 Excel에서 실행해 통과한 결과이며, 릴리스나 설치를 뜻하지 않는다. 수치는 [PERFORMANCE.md](PERFORMANCE.md)에 있다.

- 켠 상태 execute와 명시 복원(1,000셀): `native-1000-20260910-100904`. 대상 1,000셀 서식 전수 비교 불일치 0이고, 체크포인트 **이후에** 넣은 값·수식·숫자 서식, 다른 시트의 표식, 저장되지 않은 dirty 상태와 세션 설정이 보존되었다. 같은 확인을 CLI 경로에서도 했다(`native-1000-20260910-101029`).
- 5,000셀(`native-5000-20260910-102039`)에서 실행한 것은 **짝 비교의 서식 복원과 부분 실패 롤백**이다. 5,000셀 서식은 전수 비교 불일치 0이지만, 체크포인트 이후 값·수식·숫자 서식 보존 검사는 이 규모에서 실행하지 않았다. 그 검사는 위 1,000셀 두 사례의 결과다.
- 실제 부분 실패의 자동 롤백: 1,000셀 `native-1000-20260910-101029`, 5,000셀 `native-5000-20260910-102039`. 주입한 실패 전에 실제 셀 쓰기가 일어났고, 롤백은 `verified=true`로 대상 셀 서식이 전수 일치했다.
- CLI worker 왕복: `native-1000-20260910-101029`. 별도 프로세스가 만든 스냅샷 metadata가 온전히 남고, 또 다른 새 프로세스가 그 스냅샷으로 복원했다.
- 손상된 체크포인트 거절: `native-1000-20260910-101029`. 해시 불일치로 거절했고 라이브 워크북은 그대로, 체크포인트 파일은 증거로 남았다.
- 수식으로 자격에서 거절된 워크북의 eager 대체 경로: `native-1000-20260910-101750`. 제품이 `formulas`로 거절하고 기존 eager format-only 스냅샷을 만든 뒤, 명시 복원이 대상 서식을 되돌리고 심어 둔 수식·다른 시트 표식·세션 상태를 보존했다.
- 구버전 거절: `native-1000-20260910-102852`. 설치본 0.4.18, 로컬에 동결해 둔 이전 0.4.20 기반 후보(Core `AE19909C…`), 그리고 **GitHub에 게시된 v0.4.20 자산**(릴리스 ZIP SHA-256 `1B41A71557368C94D57EAC2833D61FB5EC045EB6A86B0BF2E6BA97FB16C7CB04`, 그 안의 Core `CD5BD267D88D5712BF5AF032FFEA26DAE563AE90C3B947891D73AE129982FCDF`) 세 바이너리가 모두 dry-run 토큰까지 받은 뒤 `snapshotVersion` 4를 지원하지 않는다며 **쓰기 전에** 거절했고, 이어진 native 재검사에서 모든 셀이 그대로였다. 공개 자산 실행 결과는 `old-cli-doc-bridge-78b2d71c-restore.json`이고 provenance에 세 Core 해시가 함께 기록되어 있다.

남은 한계 하나는 그대로다. **추가 기능 파트(`complex-part`)를 통한 대체 경로는 native로 확인하지 못했다.** 그 사례를 시도한 실행의 Excel 인스턴스가 `Workbooks.Add`에서 webextension 파트를 하나도 만들지 않아 전제 조건에서 실패했다. 앞선 실행들에서는 같은 방식으로 만든 파일에 파트 3개가 실제로 들어 있었고 제품은 그것을 정확히 거절했지만, 시작 상태에 따라 달라지는 이상 재현 가능한 native fixture가 아니다. 그래서 결정적인 수식 기반 사례로 교체했으며, 위 `formulas` 거절을 `complex-part` 거절의 native 증거로 읽지 않는다.

## Visibility batch 계약

`set_rows_hidden`, `set_cols_hidden`, `set_sheet_visibility`는 같은 batch에 함께 넣을 수 있다.
DocBridge는 각 행·열의 원래 hidden 값과 시트의 표시 상태, 원래 활성 시트를
`visibility-state` 스냅샷에 저장한다. 일부 행만 이미 숨겨진 혼합 상태도 항목별로 복구한다.

정확한 자동 복구를 위해 visibility op와 값·수식·서식·복사·병합 op는 같은 batch에 섞을 수
없다. 먼저 visibility batch를 완료하고 readback한 뒤, 후속 편집을 새 dry-run으로 실행한다.

## 5. 레이아웃·작성·수명주기

같은 가족끼리만 한 batch에 넣는다. `executionMode=execute` allowlist는 기존처럼
`set_values`/`set_formulas`/`format_range`만이다. 아래 op는 dry-run + confirmToken이다.

| op | 입력 | 스냅샷 |
| --- | --- | --- |
| `set_row_heights` | `rows:[{row, count?, heightPoints\|autoFit}]`. 단위는 Excel 포인트 0.1–409.5 | `sheet-layout-state` |
| `set_column_widths` | `columns:[{col, count?, widthChars\|autoFit}]`. 단위는 문자 너비 0–255 | `sheet-layout-state` |
| `freeze_panes` | `cell:"G6"` 또는 `rows`/`columns` 또는 `unfreeze:true`. G6 → xSplit=6, ySplit=5 | `sheet-layout-state` |
| `set_page_setup` | `page.scale`이 fit보다 우선. A3=8. 여백은 mm | `sheet-layout-state` |
| `rename_sheet` | `target.sheet`, `newName` | `rename-state` |
| `clear_range` | `what`: all\|contents\|formats\|formulas | `range-edit-state` |
| `copy_range` | `destRange`, `mode`: all\|values\|formulas\|formats | `range-edit-state` |
| `delete_rows` / `delete_cols` | 고위험. 삭제 띠의 값/수식/서식/크기를 스냅샷 | `delete-strip-state` |
| `add_sheet` / `move_sheet` | `copy_sheet`는 계속 단독 batch | `sheet-structure-state` |
| `delete_sheet` | 고위험. 쓰기 전 현재 workbook 사본. 다른 시트 수식·이름·차트·피벗 의존을 캡처. 표/`[owned.xlsx]Sheet`/3D는 유지. 못 캡처하면 삭제 전 거절 | `sheet-structure-state` + `workbook-copy-sheet` |
| `set_view` | 대상 시트를 활성화해 Zoom/눈금/보기 모드를 읽고 쓴다. 끝나면 원래 활성 시트·선택·스크롤을 되돌린다 | `sheet-layout-state` |
| `fill_range` / `auto_fill` | `auto_fill`은 채운 띠의 수열/상대 수식을 검사 | `extended-ops` |
| `import_csv` / `export_csv` | 같은 `delimiter`. import는 파싱 중 셀 수 상한 | `extended-ops` |
| `set_outline` | `axis` row\|column. `show:false`는 접기이며 ClearOutline이 아님 | `extended-ops` |
| `set_page_breaks` | 기본 `orientation=row`는 전체 행(`Rows(n).PageBreak`). `column`은 전체 열. 셀 `Range.PageBreak`는 쓰지 않음. readback은 H/VPageBreaks Location | `extended-ops` |
| `set_formulas` | 기본 `engine=formula`. `formula2`는 명시 opt-in | 수식 배치 또는 legacy used-range |
| `protect_sheet` / `unprotect_sheet` | 암호 필드 거부. UI-only 보호 | `protect-state` |
| `create_workbook` / `open_workbook` | 각각 단독. 실행 중인 Excel에만 연결하며 사용자 창을 닫지 않음 | `lifecycle-state` |
| `save_workbook` / `export_pdf` | 고위험. 보호된 원본 경로 거부. 기존 파일은 `overwrite:true` | `lifecycle-state` |
| `close_workbook` | 명시 `target.workbook`만. 활성 창 close 금지 | `lifecycle-state` |

owned Application을 만들었다는 것은 그 뒤에 열린 모든 통합문서의 소유가 아니다.
`Saved=true`만으로 `Application.Quit`하지 않는다. 암시적 dispose·연결 해제는
비어 있지 않은 Workbooks 컬렉션을 보존한 채 detach한다. auto-Quit는 workbook이
0개인 소유 인스턴스에만 허용한다. 특정 파일을 닫는 경로는 `close_workbook`이다.

lifecycle rollback은 신원을 실제로 확인했을 때만 `verified`/`complete`다.
확인 0건은 `unproven`이다. 기존 통합문서가 사라지면 `incomplete`이며 재생성하지 않는다.
ROT/창 탐색이 어댑터가 이미 가진 Application RCW와 같으면 그 alias를 FinalRelease하지 않는다.

## ExtendedOps와 Formula2 한계

이 절은 지원을 과장하지 않기 위한 현재 계약이다. 실물 문서 검증과 로컬 플러그인
반성이 끝나기 전에는 Excel 완성을 주장하지 않는다.

`delete_sheet` 복구는 소유한 **현재** workbook 사본(`SaveCopyAs` current-memory 또는
`Saved=true`인 last-saved-file)에서 시트 전체를 다시 복사한 뒤, 다른 시트의 수식·이름·차트
시리즈·피벗 원본을 다시 쓴다. `=Data!A1+SUM(Items[Amount])`, 같은 책
`'[owned.xlsx]Sheet'!A1`, 3D `Data:Summary!A1`은 캡처한다. 실제 다른 파일/백업 토큰만
제외한다. `[`가 하나 있다고 전부 버리지는 않는다. 닫히지 않은 `[`나 알 수 없는 대괄호는
`[EXCEL_DELETE_SHEET_DEPENDENCY_UNCAPTURED]`로 **삭제 전에** 거절한다. used-range
값·NumberFormat·탭 색만 되살리는 경로는 전체 복구로 인정하지 않으며, 사본이 없으면
쓰기 전에 거절한다.

`auto_fill` readback은 원본 블록 보존에 더해 채운 영역의 내용, 숫자 수열, 상대 수식 연속을
검사한다. `copy_range`의 formulas/all 모드는 원본 수식 문자열과 목적지를 그대로 비교하지
않고 A1 상대 참조를 이동한 뒤 비교한다.

`set_values`는 JSON 타입을 유지한다. 문자열 `"2026-09-10"`은 텍스트이며 일련번호 46275가 아니다.
`"123"`과 숫자 `123`은 다르다. 문자열은 잠시 `@`로 쓴 뒤 **원래 NumberFormat을 되돌린다**.
혼합 NumberFormat은 범위에서 null이므로 셀마다 캡처하고, 캡처에 실패하면 `@`를 쓰지 않는다.
공개 fixture의 `General`은 그대로 둔다. 한국어 Excel이 `NumberFormat="General"`에서
`0x800A03EC`를 내면 `Application.International(26)` 이름(`G/표준` 등)으로
`NumberFormatLocal`만 보조 기록한다. 요청 값을 빼거나 다른 서식으로 바꾸지 않는다.
복구는 Value2 예외를 포함해 finally에서 하며, 유지된 형식을 다시 읽는다. `""`는 빈 칸으로
정규화될 수 있다. 값 readback은 한 번의 `Value2` 행렬과 타입 비교다.
수식은 `set_formulas`만 쓴다.

`calculate`는 `Workbook.Calculate`나 `Application.Calculate`를 호출하지 않는다. 범위가 있으면
그 Range만, 없으면 대상 workbook의 각 Worksheet에서 `Calculate`하고 CalculationState를 확인한다.

`set_formulas`의 기본 엔진은 legacy `Range.Formula`다. Formula2는 `engine`/`formulaEngine`/
`formula2:true`로만 켠다. FILTER/SORT/UNIQUE는 `spillRows`를 요구하지 않는다. 쓰기 후 실제
spill이 dest에 붙어 있으면 성공으로 둔다. 네이티브 속성은 `SpillingToRange`이며, 조회
실패는 dest 크기 scalar로 취급하지 않는다. 롤백은 dest와 새로 소유한 spill을 **먼저**
지운 뒤 이전 Formula2와 원래 spill을 복구한다. 복구 증거는 복원된 수식과 셀 값이다.
쓰기 횟수는 증거가 아니다. 막힌 spill(`hasSpill` false/null)은 이웃 사용자 데이터를
지우지 않는다. 상수 SEQUENCE 또는 명시한 spill 힌트만 캡처보다 큰 spill을 거절한다.
주소는 `$`를 제거한 뒤 비교한다. 한 batch에 formula와 formula2가 같이 있으면 op마다
엔진을 유지한다.

`set_view`와 레이아웃 스냅샷은 대상 시트를 활성화해 그 창의 Zoom/눈금/보기 모드를 읽거나
쓴 다음, 원래 활성 시트·선택·스크롤을 되돌린다. 숨긴 시트의 창 설정은 읽지 않는다.
행 높이 픽셀 매핑은 대상 통합문서 창을 쓰고, 다른 책이 활성이거나 Zoom이 100이 아니면
그 창을 다시 잰다.

`set_page_setup`의 `printArea`/`printTitleRows`/`printTitleColumns`는 `$`를 뺀 범위
문법만 비교한다. `1:2`와 Excel `$1:$2`는 같다. 머리글·바닥글 문자열은 그대로 비교한다.
불일치는 필드별 expected/actual로 남긴다.

`import_csv`는 파서가 필드를 추가할 때 셀 수를 세고 `MaxImportCells`를 넘기면 테이블을
다 만들기 전에 거절한다. export도 같은 delimiter로 다시 읽어 값 동등을 확인한다.

자동화 락은 `DOCBRIDGE_HOME`/RootDir마다 다른 named mutex를 쓴다. 집을 나눈 프로세스는
서로 60초씩 기다리지 않는다. 같은 집을 쓰는 프로세스는 여전히 직렬화된다.

`format_range.style`는 `fontName`, `horizontalAlign`, `verticalAlign`, `wrapText`, `borders`를
쓴다. `left/right/top/bottom`은 **각 셀** 변, `outline`은 범위 둘레,
`insideHorizontal`/`insideVertical`은 내부 격자, `all`은 셀 변+내부다. `medium`은 굵기다.
새 키는 deferred-format 자격에서 제외된다. v2/v3 written-properties 복구는 유지한다.

인접 범위는 변 객체를 공유한다. `A16:E18` 지우기는 `F16:H18` 아웃라인의 왼쪽 변도
함께 지운다. 같은 배치에서 공유 변에 그리기 뒤 지우기가 오면
`[EXCEL_BORDER_ORDER]`로 거절한다(지우기→그리기 순서나 배치 분리로 해결).
다른 배치의 간섭은 dry-run 경고 `[EXCEL_BORDER_SHARED_EDGE]`로 알린다.
readback은 배치 자신의 op만 검증하므로, 테두리 작업 후에는 기존 둘레를 엣지
단위로 재확인한다.

데이터/보고 op(표, 정렬/필터, 이름, 유효성, 조건부 서식, 차트, 그림, 메모, 하이퍼링크,
`create_pivot`/`update_pivot`/`refresh_pivot`/`delete_pivot`)는 별도 data-only batch이며
`restoreMode=data-objects`다. execute allowlist에 넣지 않는다. 피벗 **캐시/외부 연결**과
매크로·암호 시트는 범위 밖이다. 같은 통합문서 안의 시트 피벗 만들기는 지원한다.

## 도형·텍스트 상자 실무 서식

기존 `insert_shape`, `update_shape`, `insert_textbox`, `update_textbox`는 `fillColor`/`text`/`position`에
더해 다음 선택 필드를 지원한다. 생략한 필드는 기존 값을 보존하며, `null`은 생략이 아니므로 거절한다.

- `lineColor`: 기존 OLE/`#RRGGBB` 색상 계약
- `lineWeight`: 0 초과 20 이하의 유한 point 수
- `lineVisible`: boolean (`false`는 선을 숨김)
- `rotation`: 정수 degree `0..360`; `360`은 readback에서 `0`으로 정규화
- `font`: 전체 도형/텍스트 상자 텍스트에 적용하는 부분 객체 `{name,size,bold,italic,color}`. 지정한
  font 속성만 바꾸며, run별 rich text 편집은 이 계약에 포함되지 않는다.

`excel_inspect`와 data-object apply readback은 `lineColor`, `lineWeight`, `lineVisible`, `rotation`, `font`를
반환하려 시도한다. COM이 개별 속성을 읽지 못하면 `...Unreadable`/`fontUnreadableFields`를 명시하며,
요청한 필드를 읽지 못한 apply는 성공으로 간주하지 않는다.

```json
{
  "ops": [{
    "op": "update_shape",
    "target": { "sheet": "월간보고" },
    "name": "HighlightBox",
    "lineColor": "#112233",
    "lineWeight": 1.5,
    "lineVisible": true,
    "rotation": 90,
    "font": { "name": "Arial", "size": 11, "bold": true, "color": "#FFFFFF" }
  }]
}
```

## Microsoft 공식 근거

초보·실무 교육 주제를 기능 범위로 정할 때 참고한 Microsoft Support 자료:

- [Hide or unhide worksheets](https://support.microsoft.com/en-us/excel/hide-or-unhide-worksheets)
- [Hide or show rows or columns](https://support.microsoft.com/en-US/Excel/get-started/hide-or-show-rows-or-columns)
- [Merge and unmerge cells in Excel](https://support.microsoft.com/en-US/Excel/get-started/merge-and-unmerge-cells-in-excel)
- [Enter and format data](https://support.microsoft.com/en-us/excel/enter-and-format-data)

DocBridge의 직접 COM 구현과 readback 계약을 확인할 때 참고한 Microsoft Learn API 자료:

- [Range.Merge method](https://learn.microsoft.com/en-us/office/vba/api/excel.range.merge): 범위를 병합하며 병합 셀 값은 좌상단 셀에 유지된다.
- [Range.UnMerge method](https://learn.microsoft.com/en-us/office/vba/api/excel.range.unmerge): 병합 영역을 개별 셀로 분리한다.
- [Range.MergeArea property](https://learn.microsoft.com/en-us/office/vba/api/excel.range.mergearea): 셀이 속한 병합 영역을 확인한다.
- [Range.Hidden property](https://learn.microsoft.com/en-us/office/vba/api/excel.range.hidden): 전체 행 또는 전체 열의 숨김 상태를 읽고 설정한다.
- [Worksheet.Visible property](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.visible): 워크시트 표시 상태를 읽고 설정한다.
- [XlSheetVisibility enumeration](https://learn.microsoft.com/en-us/office/vba/api/excel.xlsheetvisibility): `xlSheetVisible`, `xlSheetHidden`, `xlSheetVeryHidden`의 의미와 값을 정의한다.
- [Shape.Rotation property](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.rotation): rotation은 degree이며 Excel은 가장 가까운 정수로 반올림한다.
- [Shape.Line property](https://learn.microsoft.com/en-us/office/vba/api/excel.shape.line) 및 [LineFormat](https://learn.microsoft.com/en-us/office/vba/api/excel.lineformat): 선 색/두께/표시 상태를 제공한다.
- [TextFrame.Characters method](https://learn.microsoft.com/en-us/office/vba/api/excel.textframe.characters): start/length 생략은 전체 텍스트를 선택하며 `Characters.Font`를 쓸 수 있다.

## 후속 단계 매트릭스

아래 항목은 현재 지원을 과장하지 않기 위한 계획 구분이며 일정 확약이 아니다.

| 구분 | 기능군 | 상태 |
| --- | --- | --- |
| 지원 | 병합 batch, 서식, 행·열 크기, 고정 창, 페이지 설정, 이름 변경, 범위 지우기/복사, 행·열 삭제, 시트 추가/이동, 보호, 수명주기/저장/PDF, 표/차트/이름/유효성/조건부 서식, 같은 통합문서 피벗 | 제품 코드·정책·스키마·단위 테스트. 실물 COM 완료는 별도 검증 |
| 범위 밖 | 피벗 캐시/외부 연결, 매크로, 암호 시트 | 별도 보안·환경 |

새 기능은 `core_get_capabilities({"app":"excel"})`의 `writeOps`, `limits`, `safety`에 노출되고,
정책 allowlist·MCP 스키마·단위 테스트·실제 Excel E2E가 함께 통과한 뒤에만 지원 완료로 표시한다.
