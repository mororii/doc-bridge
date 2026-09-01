# HWP 작업 명세

DocBridge 0.4의 한글 작업은 화면 좌표나 스크립트 매크로를 사용하지 않는다. 한컴이 공개한 HWP Automation `HAction`, `HParameterSet`, `InsertPicture`, `SaveAs` API만 사용한다. 모든 쓰기 작업은 `dryRun`으로 스냅샷과 확인 토큰을 만든 뒤 적용하며, 작업 중 하나라도 실패하면 적용 직전 네이티브 HWP 스냅샷으로 자동 복원한다.

## 공통 호출

- 새 문서 경로 결정: 파일이나 창을 만들기 전에 `hwp_plan_creation`을 호출한다. 새 일반 문서는 `docx-first`, 기존 HWP/HWPX·한글 템플릿·필드·복잡한 병합표·한글 전용 개체·원본 배치 보존은 `native-hwp`를 사용한다.
- 새 일반 문서 작성: 먼저 DOCX를 렌더 검수하고 `hwp_launch`의 `creationMode:"docx-first"`와 `sourceFile` 모드로 새 HWPX/HWP를 만든다. 기존 출력은 거부하며 원본 DOCX는 저장 대상으로 사용하지 않는다. Word COM은 실행하지 않는다.
- HWP 전용 새 문서 작성: 작업 시작에 `hwp_launch({"creationMode":"native-hwp","newDocument":true})`를 정확히 한 번 호출하고, 이후 같은 열린 문서에 단계별 쓰기 작업을 적용한다.
- 사용자가 열어 둔 한글 문서: `hwp_get_active_context.summary.openDocuments`에서 대상을 확인한다. 문서가 둘 이상이면 읽기와 모든 op에 선택한 `documentRef`를 넣고 `file`은 생략한다.
- 파일 기반 편집: 배치의 모든 op에 같은 절대 `file` 경로를 넣는다.
- `documentRef`는 저장 문서의 정규화 절대 경로 또는 저장 전 문서의 `untitled-<PID>-<문서ID>`이다. 경로 인코딩이나 같은 파일의 중복 창을 구분해야 할 때는 해당 항목의 `instanceRef`(`hwp:<PID>:<문서ID>`)를 `documentRef` 값으로 사용할 수 있다.
- 한 배치에서는 모든 op가 같은 `documentRef`를 사용하거나 모두 생략해야 한다. `file`과 `documentRef`를 동시에 지정하거나 서로 다른 문서를 한 배치에 섞으면 적용 전에 거부한다.

DOCX 우선 새 문서 예시:

```json
{
  "creationMode": "docx-first",
  "sourceFile": "C:\\작업\\일일안전교육일지.docx",
  "outputFile": "C:\\작업\\일일안전교육일지.hwpx",
  "closeAfterImport": false,
  "expectedPageCount": 1,
  "expectedTableCount": 4,
  "requiredText": ["일일안전교육 일지", "교육내용", "참석자"]
}
```

한글은 `FileOpen`의 `OOXML` 형식으로 DOCX를 가져온 뒤 `FileSaveAs`로 새 네이티브 파일을 저장한다. 응답은 원본·출력 SHA-256, 파일 크기, OOXML 열기/네이티브 저장 시간, 표 수, 쪽 수와 필수 문구 검증을 포함한다. 기대 쪽 수보다 정확히 1쪽 많고 마지막 문단이 비어 있으면 OOXML 가져오기가 만든 끝 문단만 최소 높이로 축소하며 결과를 `summary.compatibilityAdjustment`에 기록한다. `summary.verification.passed:false`이면 출력이 만들어졌어도 완료가 아니며, DOCX를 조정해 새 출력 이름으로 다시 변환한다.
- `file`을 생략한 호출은 사용자가 연 창만 탐색하며 빈 한글을 자동 실행하지 않는다. 활성 문서가 없으면 문서를 연 뒤 다시 호출한다.
- 먼저 `hwp_doctor`가 `CHECK_PASSED`인지 확인하고 `core_get_capabilities({"app":"hwp"})`, `hwp_get_active_context`를 호출한다. `automationWorkingDirectory`는 설치된 한글 `Bin`, `automationWindowsDirectory`는 실제 Windows 폴더여야 한다. `automationEnvironmentRepairNeeded:true`이면 AI 런처의 process-level `windir`/`SystemRoot`가 잘못됐지만 DocBridge가 worker 및 COM 자식 환경에 정상값을 주입한다. `updateRecommended:true`는 한글 패치 권고이지 자동 차단 조건이 아니다. `ownedAutomationBlocked:true`이면 실제 오류창 또는 복구 불가능한 Windows 환경이 있으므로 새 인스턴스를 만들지 않는다. 여러 단계 편집은 `hwp_read_text({"scope":"document_map"})`의 `lineId`와 coverage를 기준으로 삼는다. 본문·문단 지도·구조가 함께 필요하면 `scope:"bundle", sections:["text","document_map","structure"]`로 한 번에 읽는다. 필요한 경우에만 `fields`·`tables` section을 추가하고 기존 표 서식은 `includeStyles:true`로 읽는다.
- 적용 뒤에는 `readback.postEditReread`와 `readback.session`을 확인한다. 실패 전 위치나 occurrence를 재사용하지 않고 갱신된 문단 지도로 다음 배치를 만든다.
- 적용은 같은 `ops`로 1) `dryRun:true`, 2) 반환된 `confirmToken`을 넣어 `dryRun:false` 순서로 수행한다.
- 10개를 넘는 op, 큰 표/그림/PDF 같은 긴 작업은 같은 payload를 `hwp_submit_ops`로 제출하고 반환된 `jobId`를 `hwp_get_job`으로 조회한다. 클라이언트가 기다리다 timeout되어도 같은 배치를 다시 제출하지 않는다. `succeeded`의 `result`가 원래 `hwp_apply_ops` 결과다.
- `dryRun`은 같은 배치의 앞 op가 만든 본문·anchor·표·행·열 상태를 뒤 op가 이어받아 순차 시뮬레이션한다. 미리보기에서 후속 op가 앞 op 결과를 찾지 못하면 적용하지 말고 결함으로 취급한다.
- 적용 결과의 `timings.previewReused`와 단계별 ms를 확인한다. HWP fingerprint가 dry-run 뒤 달라졌다는 오류는 재시도하지 말고 문서를 다시 읽은 뒤 새 dry-run을 만든다.

## 편집 작업

### `insert_text` / `append_text` / `insert_before_text` / `insert_after_text` / `replace_document_text`

- `insert_text`: 현재 커서 또는 선택 영역에 입력한다.
- `append_text`: 문서 끝으로 이동해 내용을 붙인다. `startNewParagraph` 기본값은 `true`이다.
- `insert_before_text` / `insert_after_text`: `anchor`로 찾은 기준 문구의 앞/뒤에 삽입한다. `mode:"paragraph"`(기본)는 기준 문단과 바로 위·아래 문단의 글자/문단 서식을 비교한다. 위·아래가 같으면 그 합의를, 충돌하면 기준 문단을 사용한다. `mode:"inline"`은 기준 문구 바로 옆에 삽입한다. 같은 기준 문구가 둘 이상이면 문서 앞에서 1로 시작하는 `occurrence`를 지정해야 하며, 생략 시 dry-run이 중단된다. `matchCase` 기본값은 `true`이다.
- `replace_document_text`: 기존 본문 전체를 교체한다.

모든 삽입/교체 op는 기본 `preserveStyle:true`로 글자 모양과 문단 모양을 함께 보존한다. 기존 글자 교체는 대상 서식을 최우선으로 한다. `style`은 문맥 서식보다 우선하고, 후보가 충돌할 때만 `styleSource`의 `text+occurrence` 또는 `tableIndex+cellIndex`를 지정한다. 의도적으로 보존을 끌 때만 `preserveStyle:false`를 쓴다. `text`의 줄바꿈은 실제 한글 문단으로 입력한다. 여러 문단도 한 op로 처리하며, 내용 길이나 문단 수를 이유로 PowerShell/Python COM 제어로 전환하지 않는다. 중간 삽입은 `GetSelectedPosBySet`과 `SetPosBySet`으로 선택 시작/끝 위치를 정확히 고정하고, 적용 후 삽입문이 지정 occurrence의 바로 앞/뒤 영역에 있는지와 기준 문구 개수가 변하지 않았는지를 확인한다. HWP TEXT readback이 `−`(U+2212)를 `&#8722;`처럼 직렬화하면 공개 읽기·길이·검색 전에 원문 Unicode로 디코딩한다. 호환문자/NFKC·등가 대시 변환은 검증 비교에만 적용하고 실제 입력 원문은 바꾸지 않는다. 엔터티 글자 자체를 찾으려는 특수 상황만 `options.literalEntities:true`를 사용한다.

### `find_replace`

최상위 `find`, `replace`는 필수다. `occurrence`를 생략하면 범위 안의 모든 항목, 지정하면 범위 안에서 1부터 시작하는 해당 항목 하나만 바꾼다. 범위 생략은 전체 문서이며, `scope:{"startParagraph":2,"endParagraph":4}`는 문단 범위, `scope:{"tableIndex":0,"row":1,"col":2}` 또는 `cellIndex`는 표 셀 하나다. dry-run과 apply가 같은 디코딩·대소문자·범위 규칙을 사용하고 적용 뒤 전체 텍스트 또는 셀 텍스트를 정확히 재읽는다.

### `set_paragraph_style_basic`

`style`은 `fontName`, `fontSize`, `bold`, `italic`, `textColor`, `shadeColor`, `underline`, `underlineColor`, `strikeout`, `strikeoutColor`, `letterSpacing`(-50~50), `widthRatio`(50~200), `offset`(-100~100), `superscript`, `subscript`, `align`을 지원한다. 색상은 `#RRGGBB`이다. `target.scope`은 `selection|document`, `target.text`는 정확한 대상 문구이다.

### `set_paragraph_format`

`style`은 `align`, `leftMarginMm`, `rightMarginMm`, `firstLineIndentMm`, `spaceBeforePt`, `spaceAfterPt`, `lineSpacingPercent`(50~500), `widowOrphan`, `keepWithNext`, `keepLinesTogether`, `pageBreakBefore`를 지원한다. `target.scope`은 `selection|paragraph|document`이다.

### `set_page_setup`

`page`에 `widthMm`, `heightMm`, `orientation`(`portrait|landscape`), 네 방향 `*MarginMm`, `headerMm`, `footerMm`, `gutterMm`을 넣는다. `applyTo`는 `selection|current-section|document|new-section`이다.

### `insert_break`

`type`은 `line|paragraph|page|section|column`이다.

### `insert_table` / `table_cell_set_text` / `table_set_cells` / 표 구조

`insert_table`은 직사각형 `rows`, `header`, `columnWidths`, `headerFill`, `firstColumnFill`, `fontSize`, `hideAllBorders`, `cellStyles`, `mergeCells`를 지원한다. `table_cell_set_text`는 0부터 시작하는 `tableIndex`와 `row`+`col` 또는 `cellIndex`, 교체할 `text`를 받는다. 여러 셀은 `table_set_cells`의 `cells` 배열(최대 500개)에 묶는다. 표 컨트롤과 수식 위치를 한 번만 읽어 순차 적용하므로 사진 뒤 표에서도 빠르며 각 셀은 정확히 재읽는다. 단순 값 채우기는 `preserveStyle:false`, 기존 양식 채우기는 기본 `true`를 사용한다. 병합 표에서는 `row`/`col`이 모호하므로 `scope:"tables"`의 실제 셀 이동 순서인 `cellIndex`를 사용한다.

빈 셀은 한 점으로 표현되는 모호한 "대상 위치"만 보지 않는다. 1) 앞선 반복 표의 동일 라벨 값 셀, 2) 같은 역할을 가진 위·아래 후보의 합의, 3) 위·아래에서 반복되는 다수 서식, 4) 빈 셀 자체의 기본 서식 순으로 글자/문단 서식을 찾는다. 양쪽이 다르면 빈 셀 기본 서식과 일치하는 후보를 보조 기준으로 사용한다. 기존 값이 있으면 그 값의 서식이 최우선이다. `hwp_read_text`의 `scope:"tables"`는 `HeadCtrl` 표 목록을 한 번 고정해 중복/누락 없이 읽고 각 표의 `controlRef`, 각 셀의 `cellIndex`, `text`, `hasFormula`를 반환한다. `includeStyles:true`일 때만 셀별 글자·문단 서식을 추가로 읽는다. `hasFormula:true` 셀을 일반 텍스트로 덮어쓰려 하면 `%fmu` 수식 셀 전용 오류로 중단한다.

- `table_insert_rows` / `table_insert_columns`: 기준 `row`, `col`, `count`(1~20), `position`(`before|after`).
- `table_delete_rows` / `table_delete_columns`: 대상 `row`, `col`, `count`(기본 1, 최대 20). 삽입·삭제 모두 한글의 다중 Count 동작에 의존하지 않고 대상 셀을 다시 지정해 한 줄씩 반복 실행한다. 각 회차 뒤 실제 전체 셀 수 증감이 같은 양인지 검사하고, 가능한 빌드에서는 행·열 수 증감도 함께 확인한다. 구조를 검증하지 못하면 실행 성공 반환만으로 성공 처리하지 않는다. 삭제는 내용과 서식이 사라지므로 high-risk 확인이 필요하다.
- `table_merge_cells`: `startRow`, `startCol`, `endRow`, `endCol`의 직사각형 범위를 병합한다.
- `table_set_row_height`: `tableIndex`, 0부터 시작하는 `row`, 목표 최소 `heightMm`(4~50)을 받는다. 행 전체를 블록 선택한 뒤 한글의 `TablePropertyDialog`와 `ShapeTableCell.Height`로 mm 단위 높이를 지정하고 같은 행의 실제 높이를 다시 읽는다. 내용 때문에 한글이 더 크게 보정하면 그 실측값을 반환하며, 목표보다 작을 때만 실패한다. 글자 크기나 빈 줄로 높이를 흉내 내지 않는다.

셀 나누기는 한글 2024 자동화에서 숨은 대화상자를 발생시켜 공개 명령에서 제외했다. 필요한 문서라면 병합 전 표를 다시 구성하는 방식으로 처리한다.

### 기존 양식 필드

`hwp_read_text({"scope":"fields","maxFields":100,"includeValues":true})`는 기존 필드를 제한된 개수로 읽는다. `set_field_text`는 템플릿에 이미 존재하는 `name`의 내용을 바꾸고 `GetFieldText`로 정확히 확인한다. 새 필드·북마크·하이퍼링크 삽입은 이 버전의 한글에서 숨은 대화상자를 일으켜 공개 명령에 포함하지 않는다.

### `insert_picture`

필수 `path`와 선택적인 `embedded`, `sizeOption`(`real|specific|cell|cell-ratio`), `widthMm`, `heightMm`, `reverse`, `watermark`, `effect`(`original|grayscale|black-white`)를 지원한다. `specific`에는 양수 너비와 높이가 필요하다. 표 셀 안에 넣을 때는 `tableIndex`와 `row`+`col` 또는 병합표의 `cellIndex`를 지정한다. 이 경우 기본 `sizeOption`은 `cell-ratio`이며 `clearCell:false`가 기본이라 기존 셀 내용은 보존한다. 검증은 구버전 `$pic`과 최신 한글의 `gso`를 합산해 문서 전체 그림 컨트롤이 정확히 1개 늘었는지 확인한다.

### `insert_page_number`

`position`은 `top|bottom`과 `left|center|right|inside|outside` 조합 또는 `none`이다. `format`은 `arabic|circled|roman-upper|roman-lower|alpha-upper|hangul|chinese`, `startNumber`는 1 이상이다.

### `set_header_footer_text`

`kind`은 `header|footer`, `pages`는 `both|even|odd`, `text`는 넣을 내용이다.

### `export_pdf`

`output`에 절대 `.pdf` 경로를 지정한다. 기존 파일이 있으면 교체될 수 있어 `highRiskConfirm:true`가 필요하다. 출력 파일 존재와 크기를 적용 후 확인한다.

## 검증 기준

- 오프라인 배포 게이트: Core 비-E2E 198건 + MCP 19건.
- 이슈 #1 전용 실제 한글: Unicode 숫자 엔터티 읽기/찾기, occurrence·문단 범위, 9×44+후속 표의 안정 인벤토리, 4셀 batch, 행 4개 삽입·삭제의 단계별 실제 셀 수 증감을 한 전용 문서에서 검증한다.
- 실제 다중 문서: 한글 프로세스 2개와 탭 4개를 한 번에 열거하고, 저장 전 문서와 저장 문서를 서로 다른 `documentRef`/`instanceRef`로 지정해 읽기가 분리되는 것을 검증한다.
- 실제 한글 2024(13.0.0.866): 기본 편집, 문서 끝 다문단 추가, 고유/반복 기준 문구 앞뒤 중간 삽입, 글자/문단/쪽 설정, 표 생성·기존값 서식 보존·빈 반복양식 서식 상속·셀 교체·행/열 추가/삭제·병합, 표 셀/서식 비파괴 읽기, 그림, 쪽 번호, 꼬리말, PDF를 검증한다.
- 실제 17MB 양식 복사본: 표의 오른쪽 본문 셀에서 중간 문단 삽입 후 재열기, 기준 문구 직후 위치, 표·그림 control 개수 보존, PDF 시각 배치를 검증한다.
- 전용 파일 자동화 인스턴스는 창을 정상 종료한 뒤 잔류하는 경우 생성 시 기록한 PID만 정리한다. 사용자가 연 한글 창은 종료하지 않는다.
- 시각 검수물: `reports/hwp-e2e/hwp-production-e2e.pdf`와 렌더링 PNG.
