---
name: hwp-production-workflows
description: 한컴 한글 HWP/HWPX에서 계획서·보고서·표·조직도·양식 문서를 만들거나 기존 문서를 서식 보존 편집하고 PDF로 검증할 때 사용한다. 새 일반 문서는 검수한 DOCX를 OOXML로 가져오는 DOCX 우선 경로를 기본으로 하고, 기존 HWP/HWPX·한글 고유 기능은 직접 편집한다. 열린 한글 창 편집, 표 셀/행/열 수정, 쪽 설정, 그림, 머리말·꼬리말·쪽번호, 템플릿 필드 변경 요청에도 적용한다.
---

# 한글 실무 문서 제작

한글 열기·변환·편집은 DocBridge의 한글 Automation만 사용한다. 화면 클릭, 키 입력 흉내, 매크로 파일, 셸 파일 연결은 사용하지 않는다. 새 일반 문서의 DOCX 원본 생성·렌더 검수에는 사용 가능한 DOCX 문서 제작 도구를 사용할 수 있다.

## 도구 가용성 게이트

1. 작업 전에 `core_get_status`, `hwp_plan_creation`, `hwp_launch`, `hwp_doctor`, `hwp_get_active_context`, `hwp_read_text`, `hwp_apply_ops`, `hwp_submit_ops`, `hwp_get_job`이 실제 도구 목록에 있는지 확인한다.
2. 하나라도 없거나 호출되지 않으면 즉시 중단하고 DocBridge가 로드되지 않았다고 알린다.
3. 이때 프로젝트 파일, PowerShell/Python 스크립트, 한글 매크로를 새로 만들거나 수정해서 우회하지 않는다. 컴퓨터 화면 조작이나 HWP 파일 내부 편집으로 자동 전환하지 않는다.
4. `2-TEST.cmd` 실행, AI 프로그램 완전 종료·재실행, 새 작업 시작을 안내하고 도구가 보인 뒤에만 계속한다.
5. 텍스트가 길거나 여러 문단이라는 이유로 PowerShell/Python COM으로 전환하지 않는다. 필요하면 DocBridge 배치를 나누되 끝까지 DocBridge 도구만 사용한다.
6. 10개 초과 op, 큰 표·사진·PDF 등 60초를 넘을 수 있는 작업은 `hwp_submit_ops`로 한 번만 제출하고 `hwp_get_job`으로 조회한다. 클라이언트 timeout 뒤 같은 payload를 다시 제출하지 않는다.

## 시작과 대상 고정

1. 첫 한글 작업에서 `hwp_doctor`를 호출한다. `state:"CHECK_PASSED"`일 때 계속하며 `automationWorkingDirectory`가 설치된 한글의 `Bin`, `automationWindowsDirectory`가 실제 Windows 폴더인지 확인한다. `automationEnvironmentRepairNeeded:true`는 AI 런처가 `windir`/`SystemRoot`를 누락·오염시켰지만 DocBridge가 worker와 COM 자식 환경에서 복구한다는 뜻이므로 오류가 아니다. TypeLib 누락·버전 불일치는 편집을 시작하지 말고 원인을 알린다. `hwp_repair_typelib`은 사용자가 레지스트리 재등록과 UAC를 명시 승인한 경우에만 `confirm:true`로 실행하고 한글과 AI 클라이언트를 완전히 재시작한다. `PopupBorderImpl`/`TourPopup`/`MS.Internal.FontCache.Util`/`CultureFontManager`, `HWP_UI_INITIALIZATION_FAILED`가 실제로 보이면 같은 실행을 반복하지 않는다. 오류 창에서 `아니요(N)`를 눌러 문서를 유지한다. `ownedAutomationBlocked:true`이면 먼저 오류창을 닫거나 Windows 자동화 환경을 복구하고, 환경이 정상인데도 재현될 때 한컴 자동 업데이트를 실행한다.
2. 기존 문서 편집은 `core_get_capabilities({"app":"hwp"})`, `hwp_get_active_context`를 호출한다. `summary.openDocuments`의 모든 표시 창과 탭을 확인하고, 요청한 파일명·경로·내용과 일치하는 한 항목을 선택한다. 새 문서 작성은 먼저 `hwp_plan_creation`을 호출한다. `mode:"docx-first"`이면 DOCX를 먼저 만들고 `hwp_launch`로 가져오며, `mode:"native-hwp"`일 때만 `hwp_launch({"creationMode":"native-hwp","newDocument":true})`를 작업 시작에 정확히 한 번 호출해 반환된 `documentRef`를 고정한다.
3. 열린 문서가 하나뿐이어도 가능하면 `hwp_read_text`와 모든 쓰기 op에 선택한 `documentRef`를 넣는다. 문서가 둘 이상이면 반드시 넣는다. 한 배치의 모든 op는 같은 `documentRef`를 사용해야 하고 `file`과 함께 쓰지 않는다. 저장 문서는 경로가 `documentRef`이며, 저장 전 문서는 `untitled-<PID>-<문서ID>`다. 한글 경로가 셸에서 깨지거나 동일 파일이 중복 열린 경우에는 `instanceRef`(`hwp:<PID>:<문서ID>`)를 `documentRef` 값으로 사용한다.
4. 명시적인 디스크 파일 작업만 모든 op에 같은 절대 `file`을 넣는다. 경로를 지정하면 DocBridge는 모든 표시 한글 창과 탭을 조사하고 해당 문서를 활성화한다. `HWP_DUPLICATE_LOCAL_PATH`가 나오면 `openDocuments`의 고유 `instanceRef`를 사용하거나 중복 창을 닫으며 임의 선택하지 않는다. 실시간 호출은 빈 한글 창을 자동 실행하지 않는다. 열린 문서가 없으면 사용자가 문서를 열도록 안내하고 중단한다.
5. 편집 전 `hwp_read_text`로 문서를 읽는다. 본문·문단 지도·구조가 함께 필요하면 `scope:"bundle", sections:["text","document_map","structure"]` 한 번을 우선 사용한다. 표·필드가 실제로 필요할 때만 `tables`·`fields` section을 추가하고 기존 표 서식은 `includeStyles:true`로 읽는다. 긴 문서는 `coverage.complete` 또는 후속 `nextStartParagraph`를 확인한다.
6. 대상 문구·표 번호·셀 좌표를 확인하지 못했거나 읽기 범위가 잘렸으면 추측해서 쓰지 않는다.
7. DocBridge는 다른 앱에서 작업 중인 사용자의 전경 창을 유지한다. 대상 한글 탭은 Automation 제약상 내부적으로 잠깐 활성화할 수 있지만 호출이 끝나면 원래 탭·커서·선택 영역을 복원한다. 화면 클릭이나 창 활성화로 이를 보조하지 않는다.
8. 사용자가 다른 프로그램에서 계속 일하는 것은 허용한다. 사용자가 같은 한글 창을 동시에 조작해 `interaction.userActivityDetected:true`, `interaction.interrupted:true`, 또는 `APP_USER_ACTIVITY_DETECTED`가 반환되면 남은 op를 완료됐다고 간주하지 않는다. 사용자가 한글 작업을 마친 뒤 문서를 다시 읽고 미실행 단계만 새 dry-run으로 계획한다. `foregroundPreserved:false`나 `originalStateRestored:false`도 재읽기 전 추가 쓰기를 금지한다.

## 제작 방식 선택

- 새 문서를 만들기 전에 `hwp_plan_creation`을 반드시 호출한다. 새 일반 문서는 `documentState:"new"`로 평가하면 기본적으로 `mode:"docx-first"`가 된다. 기존 HWP/HWPX, 기존 한글 템플릿, 한글 필드·누름틀, 한글 전용 개체, 복잡한 병합표, 원본 배치 동일성 중 하나라도 필요하면 `mode:"native-hwp"`가 된다. AI가 선호에 따라 이 결정을 바꾸지 않는다.
- 새 일반 문서의 기본 경로: 문단·표·그림·단순 머리말/꼬리말로 구성되고 기존 HWP 템플릿 필드나 한글 전용 개체가 필요하지 않으면 DOCX 우선 경로를 사용한다. DOCX 제작 도구로 A4와 표 너비·행 높이·글꼴·문단 간격을 완성하고 PDF/PNG로 모든 쪽을 먼저 검수한 다음 `hwp_launch`에 `creationMode:"docx-first"`, `sourceFile`, `outputFile`을 지정해 가져온다. 탐색기 더블클릭, Word/한워드 셸 연결, PowerShell/Python HWP COM은 사용하지 않는다. 운영 변환 경로에는 Word COM이 필요하지 않다.
- 검증된 성능 기준: 대표 A4 문서에서 DOCX 생성 평균 0.264초, 한글 OOXML 가져오기·저장 평균 0.668초로 핵심 작업은 약 0.93초였다. 직접 HWP 복합 배치 21건의 적용 중앙값은 2.746초였고, 표·서식을 여러 단계로 구성한 빈 문서 제작은 약 7.9초 이상이었다. 서로 다른 문서 로그이므로 절대 배율을 보장하지는 않지만 새 일반 문서의 기본값을 DOCX 우선으로 정할 근거로 사용한다.
- DOCX 우선 품질 게이트: 렌더한 DOCX의 쪽 수와 표 수를 `expectedPageCount`·`expectedTableCount`, 제목·이름·마지막 셀 같은 핵심 문구를 `requiredText`로 전달한다. `summary.sourceUnchanged:true`, `summary.verification.passed:true`, 빈 `warnings`를 모두 확인한다. 변환 파일이 생성됐다는 이유만으로 완료로 보고하지 않는다. 출력 경로가 이미 있으면 새 이름을 사용하며 덮어쓰기 우회는 하지 않는다.
- DOCX 우선 검증: 변환 후 `scope:"bundle"`, `sections:["text","document_map","structure","tables"]`, `includePageCount:true`로 다시 읽는다. 기대 쪽 수보다 정확히 1쪽 많고 마지막 문단이 비어 있으면 DocBridge가 OOXML 호환성 빈 문단만 최소 높이로 축소하며, `summary.compatibilityAdjustment.applied:true`와 최종 쪽 수를 확인한다. 그래도 페이지 수가 다르면 DOCX에 10~15% 세로 여유를 두고 표 높이·문단 간격을 줄여 재생성한다. PDF 시각 검수는 별도 `export_pdf` 고위험 승인을 받은 뒤 수행한다.
- HWP 직접 경로: 기존 HWP/HWPX 양식의 중간 수정, 템플릿 필드, 복잡한 병합표, 한글 전용 개체나 원본 서식과 동일해야 하는 작업은 처음부터 기존 HWP 직접 편집 경로를 유지한다. DOCX 변환으로 우회하지 않는다. 새 HWP 전용 양식은 `creationMode:"native-hwp"`를 명시한다.
- 새 문서나 전체 재작성: `replace_document_text`로 문단 골격을 먼저 만들고 서식·표·그림을 뒤따라 적용한다.
- 새 HWP 전용 빈 문서 작성: `hwp_plan_creation`이 `native-hwp`를 반환한 경우에만 `hwp_launch({"creationMode":"native-hwp","newDocument":true})`로 문서를 한 번 만든다. 이후 문단마다 다시 실행하지 말고 같은 문서에 단계별 배치를 적용한다.
- 기존 문서 끝에 내용 추가: `append_text`를 사용한다. `text`에 줄바꿈을 넣으면 실제 한글 문단으로 보존되므로 여러 문단도 한 op로 처리한다. 기본값 `startNewParagraph:true`는 기존 본문 뒤에서 새 문단을 시작한다.
- 기존 문서 중간에 새 내용 삽입: 고유한 주변 문구를 `anchor`로 삼아 `insert_before_text` 또는 `insert_after_text`를 사용한다. `mode:"paragraph"`(기본)는 기준 문단과 바로 위·아래 문단의 글자/문단 서식을 비교한다. 위·아래가 같으면 그 합의를, 충돌하면 기준 문단을 사용한다. 같은 문구가 둘 이상이면 문서 앞에서부터 1로 시작하는 `occurrence`를 반드시 지정한다. 기준 문구 바로 옆에 붙이는 경우에만 `mode:"inline"`을 사용한다.
- 기존 문서 부분 수정: 정확한 `target.text`, 선택 영역, `find_replace`, `insert_before_text`, `insert_after_text`, `table_cell_set_text`, `set_field_text`를 사용한다. `find_replace`는 전체 문서, `occurrence`, 문단 범위 또는 표 셀 scope를 명시할 수 있다. 중간 삽입 후에는 기준 문구와 삽입문의 순서, 기준 문구 개수, 표·그림 control inventory를 다시 읽고 필요하면 PDF로 주변 배치를 확인한다.
- 표 중심 문서: 표 구조와 셀 병합·채움·테두리를 먼저 만든다. 신규 `insert_table`에는 내용 역할에 맞는 `columnWidths`를 기본으로 명시한다. 여러 셀은 `table_set_cells.cells` 한 op로 묶고 단순 값 채우기는 `preserveStyle:false`, 기존 양식은 기본 `true`를 사용한다. 여러 논리 행 높이는 `table_set_row_heights.rows:[{"row":0,"heightMm":9}, ...]` 한 op로 적용한다. 한 행만 바꿀 때만 `table_set_row_height`를 사용한다. 행·열 변경은 별도 배치에서 검증한다.
- A4 한 장 양식의 세로 배치를 맞출 때는 빈 줄이나 글자 크기로 표 높이를 흉내 내지 않는다. 여러 행은 `table_set_row_heights`, 한 행은 `table_set_row_height`로 mm 단위 지정하고, `scope:"structure", includePageCount:true`로 정확히 1쪽인지 확인한다. 2쪽이 되면 해당 배치를 자동 롤백하거나 더 작은 높이로 다시 계획한다.
- 참고 이미지가 있으면 글자 내용을 그대로 추정하지 말고, 사용자가 제공한 값과 이미지의 구조·색·배치만 반영한다.

## 안전한 실행 순서

1. 한 번의 배치에는 같은 목적의 op만 넣는다.
2. 동일한 `ops`로 `hwp_apply_ops(dryRun:true)`를 실행한다.
3. 오류, 경고, 대상, diff를 검사한다. 저위험 작업은 사용자의 원 요청이 승인이다.
   - 같은 배치의 뒤 op는 앞 op가 만든 본문·anchor·표 구조를 dry-run에서 이어받아야 한다. 0 occurrence나 표 없음이 나오면 그대로 apply하지 않는다.
4. 반환된 `confirmToken`과 정확히 같은 `ops`로 `dryRun:false`를 실행한다.
5. `ok`뿐 아니라 `readback.verified`, `mismatches`, `operationResults`, `readback.session`, `timings`를 확인한다. 한 단계가 실패하면 뒤 단계가 실행되지 않았는지 `failedStep`과 `stoppedEarly`로 확인한다. fingerprint 변경 오류는 오래된 토큰으로 반복하지 말고 갱신된 문서를 다시 읽어 새 dry-run을 만든다.
6. 다음 배치는 `readback.postEditReread`의 갱신된 본문·문단 지도·문서 ID를 기준으로 다시 계획한다. 실패 이전의 anchor 순번이나 위치를 그대로 재사용하지 않는다.
7. 표 행/열 삭제와 `export_pdf`에는 명시적 승인 후 `highRiskConfirm:true`를 넣는다.
8. 실패 시 자동 롤백 결과를 확인한다. 수동 복원은 `core_restore_snapshot`의 두 단계 흐름만 사용한다. `HWP_COM_TIMEOUT` 또는 `HWP_CIRCUIT_OPEN`이면 즉시 자동 재시도하지 않고 `retryPolicy.mode:"after-delay"`와 `retryAfterMs`를 지킨다. 이 보호 시간에는 새 worker나 빈 한글 창을 반복 실행하지 않는다. 팝업을 닫고 지연 뒤 문서를 다시 읽어 새 dry-run을 만든다. `HWP_UI_INITIALIZATION_FAILED`이면 자동·수동 재시도를 모두 중단하고 오류창과 `hwp_doctor`를 확인한다. `HWP_AUTOMATION_ENVIRONMENT_INVALID`이면 `windir`/`SystemRoot` 및 Windows 설치 폴더를 복구하기 전에는 새 인스턴스를 만들지 않는다.

## 실무 서식 원칙

- 모든 텍스트 쓰기는 `preserveStyle:true`가 기본이다. 기존 글자 교체는 기존 값의 글자/문단 서식을 유지한다. 빈 표 셀은 한 점의 "대상 위치"로 판단하지 않고 앞선 반복 양식의 같은 라벨 값 셀과 같은 역할의 위·아래 후보를 함께 비교한다. 양쪽 합의, 주변 다수 서식, 빈 셀 기본 서식과의 일치 순으로 해석한다.
- `style`은 자동으로 찾은 문맥 서식보다 우선한다. 주변 문맥이 서로 충돌할 때만 `styleSource`의 `text+occurrence` 또는 `tableIndex+cellIndex`로 복사 원본을 명시한다. 의도적으로 기본 서식을 쓰려는 경우에만 `preserveStyle:false`를 사용한다.
- 제목·본문·표 머리글을 역할별로 나눈다. 같은 대상에 글자·문단 서식을 모두 적용하거나 대상이 여러 개면 `format_paragraphs.items`로 묶어 대상 검색을 한 번만 수행한다. 단일 한 종류 서식만 바꿀 때는 `set_paragraph_style_basic` 또는 `set_paragraph_format`을 사용한다.
- 용지·여백·방향은 내용 작성 초기에 `set_page_setup`으로 확정한다.
- 정식 양식의 표 크기 조절은 선택 사항이 아니다. 열 비율은 `insert_table.columnWidths`, 여러 행 높이는 `table_set_row_heights.rows`, 한 행 높이는 `table_set_row_height.heightMm`로 명시한다. 일반적인 시작값은 머리글 9~10mm, 한 줄 본문 10~12mm, 설명·서명 행 12~16mm이며 실제 글자 수와 A4 가용 높이에 따라 조정한다. 적용 결과의 행별 실측 높이와 최종 페이지 수를 반드시 확인한다.
- 문서 전체 줄간격을 먼저 정한 뒤 제목·도입문·첨부·날짜·서명 등 역할별 `set_paragraph_format`으로 앞뒤 간격과 줄간격을 재정의한다. 빈 문단을 반복해 세로 위치를 맞추지 않는다.
- 표 좌표는 0부터 시작한다. 직사각 표는 `row`+`col`, 병합 표는 `scope:"tables"`에서 확인한 실제 이동 순서 `cellIndex`를 쓴다. `controlRef`가 표마다 고유한지 확인하고 `hasFormula:true` 셀은 일반 텍스트로 덮어쓰지 않는다. `table_cell_set_text`와 `table_set_cells`는 셀 내용을 추가가 아니라 정확히 교체한다.
- 표 입력 뒤 `scope:"tables", includeStyles:true`를 다시 읽어 텍스트뿐 아니라 글꼴, 크기, 굵기, 정렬, 여백, 들여쓰기, 문단 간격, 줄 간격이 원본 역할 셀과 같은지 확인한다.
- 행·열 추가/삭제 후에는 문서 텍스트와 `scope:"structure"`를 다시 읽어 표 수와 핵심 셀 문자열을 확인한다.
- 행·열을 여러 개 추가·삭제할 때는 `count`를 실제 필요한 수로 지정하고 적용 결과의 `operationResults`에서 `completed/count`를 확인한다. `scope:"tables"`가 `rowCount`/`columnCount`를 반환하는 빌드는 함께 대조하고, 반환하지 않는 빌드는 새 마지막 행·열과 삭제 경계 다음 셀에 고유 표식을 넣어 readback한다. 한 번 호출했다는 이유만으로 요청 개수가 반영됐다고 간주하지 않는다.
- 머리말/꼬리말, 쪽번호, 그림은 별도 구조 제어이므로 텍스트 readback과 control inventory를 함께 확인한다.
- 표 셀 사진은 `insert_picture`에 `tableIndex`와 `row`+`col` 또는 `cellIndex`를 지정한다. 기존 셀 내용을 지우라는 요청이 없으면 `clearCell:false`를 유지하고 `sizeOption:"cell-ratio"`로 셀 안에 맞춘다.
- PDF 납품물은 `export_pdf` 후 파일 크기뿐 아니라 렌더링된 페이지를 시각 검수한다.

## 지원하지 않는 무인 작업

한글 2024에서 셀 나누기, 새 필드, 북마크, 하이퍼링크 삽입은 숨은 대화상자를 일으켰으므로 실행하지 않는다. 기존 템플릿 필드의 값 변경은 `set_field_text`로 지원한다. 필요한 경우 문서 구조를 다시 구성하되 자동화되지 않은 기능을 지원된 것처럼 주장하지 않는다.

## 작업별 레시피

계획서·보고서·조직도·템플릿 필드·PDF 예시는 [references/recipes.md](references/recipes.md)를 읽는다.
