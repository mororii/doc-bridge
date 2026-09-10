# Excel 기본 편집 operations

DocBridge는 Excel 화면을 클릭하거나 VBA를 실행하지 않고, 실행 중인 Excel의 ActiveX COM
객체를 통해 workbook을 직접 읽고 수정한다. 이 문서는 `excel_apply_ops`에 추가된 기본 편집
5종과 `excel_read_range(includeLayout:true)`의 안전 계약을 설명한다.

## 연결 사전검사와 파일 열기

먼저 `core_get_status({"app":"excel"})`의 `apps.excel`을 확인한다. `connected:true`이고
`document`가 비어 있지 않을 때만 `excel_get_active_context`를 한 번 호출한다. Excel이 닫혔거나
열린 workbook이 없으면 상태조회·컨텍스트·쓰기는 새 Excel을 만들지 않으며, 상태 변화 없이
같은 호출을 반복하지 않는다. ping·전체 앱 status·반복 context·전체 시트 읽기를 강제하지 않는다.

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
| `format_range` | 지정 범위의 글꼴·채우기 등 기본 서식 | `range`, `style`; `target.sheet` 또는 시트 한정 `range` |

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

같은 검사 도구의 표본은 [PERFORMANCE.md](PERFORMANCE.md)에 있다.
1,000셀 균일 Bold는 이전 안정화 후보와 개선 후보(legacy·execute)가 전수 mismatch 0으로
통과했고, 7,000셀 균일 Bold execute도 전수 mismatch 0으로 통과했다(제품 합계 0.459초).
1,000셀 혼합의 제품 합계는 legacy 20.903초, execute 11.182초다. 실제 COM 장애 뒤 원복,
UUID 재전송 challenge/conflict, 작은 회귀 9건도 통과했다. 보호시트 건은
`rollback.verified=false`를 숨기지 않으며 원복 성공이 아니다. 7,000셀 혼합 execute는
전수 7,000셀 검증이 통과했지만 execute 전체 52.154초(그중 snapshot 51.864초, 실제
apply 0.169초), restore 30.655초, 합 82.809초다. 별도 native 81.523초는 제품 합계에
넣지 않는다. 균일 0.459초와 혼합 82.809초는 다른 작업이다. 만능 속도 배수를 만들지
않는다. 각 행은 단일 표본이며 동일 워크스테이션의 세션/캐시/부하 차이가 있다.

다른 시트의 부풀린 UsedRange는 가용성 회복이지 일반 속도 개선이 아니다. 당시 후보가
대상만 스냅샷해 통과한 기록은 역사적 가용성 증거이며 현재 소스의 최종 시간이 아니다.

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
| 1차 | 병합/병합 해제, 행·열 숨김/표시, 시트 일반 숨김/표시, `includeLayout` 읽기, `format_range` 별칭·v3 written-properties 스냅샷, execute allowlist | 구현·정책·스냅샷·readback 제공. 균일/혼합 1,000·7,000셀과 COM 원복·UUID 재전송은 [PERFORMANCE.md](PERFORMANCE.md). 균일 0.459초와 혼합 82.809초는 구분 | 실제 Excel E2E, `rollback.verified` 확인, 정상 disconnect/파이프 회수. 크래시 잔류는 미해결. 부풀림 통과는 가용성이지 속도 개선이 아님 |
| 2차 | 행 높이, 열 너비, AutoFit, 줄 바꿈, 정렬, 테두리 | 후보 | 혼합 셀 상태의 정확한 스냅샷과 단위/자동맞춤 readback |
| 3차 | 행·열 삽입/삭제 확장, 고정 창, 그룹/윤곽, 정렬·필터 | 후보 | 필터 숨김과 수동 숨김 구분, 구조 변경 후 주소 재계산 |
| 4차 | 표(ListObject), 이름 정의, 데이터 유효성, 조건부 서식 | 후보 | 수식·이름 범위·테이블 참조 보존과 operation-scoped 복구 |
| 5차 | 차트·피벗 수정, 페이지 설정·인쇄 영역·PDF 출력 | 후보 | 캐시/외부 연결, 출력 파일 교체 승인, 실제 렌더 검증 |

새 기능은 `core_get_capabilities({"app":"excel"})`의 `writeOps`, `limits`, `safety`에 노출되고,
정책 allowlist·MCP 스키마·단위 테스트·실제 Excel E2E가 함께 통과한 뒤에만 지원 완료로 표시한다.
