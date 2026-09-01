# Excel 기본 편집 operations

DocBridge는 Excel 화면을 클릭하거나 VBA를 실행하지 않고, 실행 중인 Excel의 ActiveX COM
객체를 통해 workbook을 직접 읽고 수정한다. 이 문서는 `excel_apply_ops`에 추가된 기본 편집
5종과 `excel_read_range(includeLayout:true)`의 안전 계약을 설명한다.

## 연결 사전검사와 파일 열기

먼저 `core_get_status`의 `apps.excel`을 확인한다. `connected:true`이고 `document`가 비어 있지
않을 때만 `excel_get_active_context`를 한 번 호출한다. Excel이 닫혔거나 열린 workbook이 없으면
상태조회·컨텍스트·dry-run은 새 Excel을 만들지 않으며, 상태 변화 없이 같은 호출을 반복하지 않는다.

`excel_read_range`와 `excel_inspect`도 workbook 경로만으로 닫힌 파일을 자동으로 열지 않는다.
사용자가 닫힌 기존 파일을 열어 읽으라고 명시한 경우에만 존재하는 절대 경로와
`allowOpenFile:true`를 함께 쓴다. 이 옵션은 읽기 전용이며 Excel 쓰기에는 사용할 수 없다.
쓰기 대상은 먼저 Excel에서 열어야 한다.

DocBridge 오류나 제약을 `openpyxl`, `pywin32`/직접 Excel COM, PowerShell Excel COM,
`Start-Process` 또는 UI 자동화로 우회하지 않는다. 도구 오류와 필요한 사용자 조치를 보고하고
중단한다. 우회 파일 재작성은 기존 서식·매크로를 손상할 수 있고, 직접 COM 인스턴스는 열린
통합문서가 없는 회색 Excel 창이나 잔류 `EXCEL.EXE`를 만들 수 있다.

## 지원 범위

| op | 용도 | 필수 입력 |
| --- | --- | --- |
| `merge_cells` | 직사각 범위를 하나의 셀로 병합 | `range`; `target.sheet` 또는 시트 한정 `range` |
| `unmerge_cells` | 지정 범위 안의 병합 영역을 해제 | `range`; `target.sheet` 또는 시트 한정 `range` |
| `set_rows_hidden` | 연속된 행을 숨기거나 다시 표시 | `target.sheet`, `row`, `count`, `hidden` |
| `set_cols_hidden` | 연속된 열을 숨기거나 다시 표시 | `target.sheet`, `col`, `count`, `hidden` |
| `set_sheet_visibility` | 워크시트를 일반 숨김 또는 표시 | `target.sheet`, `visibility` (`visible` 또는 `hidden`) |

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

## 공통 dry-run과 적용

먼저 원하는 op를 `dryRun:true`로 호출한다.

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

적용 뒤 `readback.verified:true`를 확인하고 같은 범위를 `includeLayout:true`로 다시 읽는다.
문서나 op가 달라졌거나 토큰의 5분 유효시간이 지났으면 기존 토큰을 재사용하지 않는다.

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

`merge_cells`와 `unmerge_cells`는 정확한 작업 범위 복구를 위해 한 batch에서 단독 op로만
실행한다. 값 입력이나 서식 변경은 별도 dry-run batch로 나눈다.

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

## Visibility batch 계약

`set_rows_hidden`, `set_cols_hidden`, `set_sheet_visibility`는 같은 batch에 함께 넣을 수 있다.
DocBridge는 각 행·열의 원래 hidden 값과 시트의 표시 상태, 원래 활성 시트를
`visibility-state` 스냅샷에 저장한다. 일부 행만 이미 숨겨진 혼합 상태도 항목별로 복구한다.

정확한 자동 복구를 위해 visibility op와 값·수식·서식·복사·병합 op는 같은 batch에 섞을 수
없다. 먼저 visibility batch를 완료하고 readback한 뒤, 후속 편집을 새 dry-run으로 실행한다.

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

## 후속 단계 매트릭스

아래 항목은 현재 지원을 과장하지 않기 위한 계획 구분이며 일정 확약이 아니다.

| 단계 | 기능군 | 상태 | 구현 전 필수 검증 |
| --- | --- | --- | --- |
| 1차 | 병합/병합 해제, 행·열 숨김/표시, 시트 일반 숨김/표시, `includeLayout` 읽기 | 구현·정책·스냅샷·readback 제공 | 실제 Excel E2E, 자동 롤백, 잔류 `EXCEL.EXE` 없음 |
| 2차 | 행 높이, 열 너비, AutoFit, 줄 바꿈, 정렬, 테두리 | 후보 | 혼합 셀 상태의 정확한 스냅샷과 단위/자동맞춤 readback |
| 3차 | 행·열 삽입/삭제 확장, 고정 창, 그룹/윤곽, 정렬·필터 | 후보 | 필터 숨김과 수동 숨김 구분, 구조 변경 후 주소 재계산 |
| 4차 | 표(ListObject), 이름 정의, 데이터 유효성, 조건부 서식 | 후보 | 수식·이름 범위·테이블 참조 보존과 operation-scoped 복구 |
| 5차 | 차트·피벗 수정, 페이지 설정·인쇄 영역·PDF 출력 | 후보 | 캐시/외부 연결, 출력 파일 교체 승인, 실제 렌더 검증 |

새 기능은 `core_get_capabilities({"app":"excel"})`의 `writeOps`, `limits`, `safety`에 노출되고,
정책 allowlist·MCP 스키마·단위 테스트·실제 Excel E2E가 함께 통과한 뒤에만 지원 완료로 표시한다.
