# 한글 실무 레시피

## 새 일반 문서 DOCX 우선 생성

기존 한글 양식을 수정하는 작업이 아니고 문단·표·그림·단순 머리말/꼬리말로 구성된 새 문서라면 DOCX에서 먼저 완성하고 한글로 가져온다.

1. 먼저 `hwp_plan_creation({"documentState":"new"})`을 호출한다. 기존 한글 템플릿·한글 필드·복잡한 병합표·한글 전용 개체·원본 배치 보존이 없으면 `mode:"docx-first"`가 반환되어야 한다.
2. DOCX 제작 도구로 A4 용지, 고정 표 너비, 글꼴, 행 높이, 문단 간격을 명시해 `.docx`를 만든다.
3. DOCX를 PDF/PNG로 렌더하고 모든 페이지의 겹침·잘림·빈 페이지를 확인한다.
4. DOCX 표 수와 핵심 문구를 센다. 한글 변환에서 재검증할 값을 준비한다.
5. 다음처럼 새 출력 경로로 변환한다. 기존 출력은 덮어쓰지 않는다.

```json
{
  "creationMode": "docx-first",
  "sourceFile": "C:\\작업\\현장기술인_변경계.docx",
  "outputFile": "C:\\작업\\현장기술인_변경계.hwpx",
  "expectedPageCount": 1,
  "expectedTableCount": 6,
  "requiredText": ["현장기술인 변경계", "홍길동", "발주처 담당자"],
  "closeAfterImport": false
}
```

6. 응답의 `creationMode:"docx-first"`, `sourceFormat:"OOXML"`, `sourceUnchanged:true`, `verification.passed:true`를 확인한다. `pageCountMatches:false`, `tableCountMatches:false`, 필수 문구 누락 또는 경고가 하나라도 있으면 완료로 보고하지 않는다.
7. 변환된 `documentRef` 또는 `outputFile`을 `scope:"bundle"`, `sections:["text","document_map","structure","tables"]`, `includePageCount:true`로 다시 읽는다. PDF 검수는 사용자가 `export_pdf`를 명시 승인한 뒤 수행한다.

Word 1쪽이 한글에서 2쪽이 되는 경우가 있다. 기대 쪽 수보다 정확히 1쪽 많고 마지막 문단이 비어 있으면 DocBridge가 한글 OOXML 가져오기에서 생긴 끝 문단만 최소 높이로 축소한다. `summary.compatibilityAdjustment.applied:true`와 최종 `verification.passed:true`를 확인한다. 이 조건이 아니거나 여전히 쪽 수가 다르면 DOCX를 페이지 끝까지 채우지 말고 10~15% 세로 여유를 둔다. 기존 HWP 템플릿 필드·복잡한 병합표·한글 전용 개체가 필요하면 이 경로를 쓰지 않고 HWP 직접 편집을 사용한다.

## 여러 문단 이어 쓰기

문서 끝에 여러 문단을 붙일 때는 커서 이동이나 외부 COM 스크립트를 만들지 말고 한 번의 `append_text`를 사용한다.

```json
{
  "op": "append_text",
  "text": "첫째 문단\n둘째 문단\n셋째 문단",
  "startNewParagraph": true
}
```

줄바꿈은 실제 한글 문단으로 입력된다. 도구 호출 크기 제한이 있으면 같은 순서를 유지해 여러 `append_text` 배치로 나누고, 각 배치마다 dry-run → apply → readback을 완료한다.

## 기존 양식 중간에 문단 삽입

기존 양식의 특정 항목 주변에 새 문단을 넣을 때는 커서 위치를 추측하지 말고, 문서에서 고유하게 식별되는 기준 문구를 사용한다.

```json
{
  "op": "insert_after_text",
  "anchor": "관련근거 : 현장확인",
  "text": "검토 결과: 추가 보완사항 없음",
  "mode": "paragraph",
  "matchCase": true
}
```

같은 문구가 두 번 이상이면 dry-run이 `occurrence`를 요구한다. 예를 들어 두 번째 항목 뒤에 붙이려면 `"occurrence":2`를 추가한다. 기준 문구 바로 뒤에 같은 줄로 붙여야 할 때만 `"mode":"inline"`을 사용한다. 적용 후 문서 텍스트를 다시 읽어 `anchor → 삽입문` 순서를 확인하고, 표·그림이 있는 양식은 `scope:"structure"`의 control 개수가 유지되는지 확인한다. 중요한 문서는 PDF로 해당 쪽을 렌더링해 셀 경계, 줄바꿈, 글꼴과 정렬을 눈으로 확인한다.

## 기존 양식 표 셀 채우기

셀을 추측해 쓰기 전에 표 내용과 서식을 읽는다.

```json
{"scope":"tables","tableIndex":3,"maxCells":100,"includeStyles":true}
```

병합이 없는 표는 `row`와 `col`, 병합 표는 읽기 결과의 `cellIndex`로 쓴다.

```json
{
  "op":"table_cell_set_text",
  "tableIndex":3,
  "cellIndex":7,
  "text":"홍길동",
  "preserveStyle":true
}
```

여러 셀은 한 번에 묶는다.

```json
{
  "op":"table_set_cells",
  "tableIndex":0,
  "preserveStyle":false,
  "cells":[
    {"row":1,"col":0,"text":"2026-08-31"},
    {"row":1,"col":1,"text":"홍길동"}
  ]
}
```

큰 표·그림·PDF처럼 60초를 넘을 수 있는 배치는 `hwp_submit_ops`로 한 번만 제출하고 `hwp_get_job({"jobId":"..."})`의 `terminal:true`를 기다린다. timeout이 나도 같은 배치를 재제출하지 않는다.

기존 글자가 있으면 그 글자와 문단 서식을 유지한다. 빈 셀은 한 점의 위치만 보지 않고 바로 왼쪽 라벨과 앞선 반복 표의 같은 라벨 값 셀을 우선 확인한다. 그 다음 같은 역할의 위·아래 후보를 함께 모아 양쪽 합의와 주변 다수 서식을 선택하며, 충돌 시 빈 셀 기본 서식과 일치하는 쪽을 보조 기준으로 삼는다. 그래도 의도가 불명확할 때만 `"styleSource":{"tableIndex":1,"cellIndex":7}`처럼 원본을 지정한다. 적용 뒤 `scope:"tables"`를 다시 읽어 글꼴·크기·굵기·정렬·여백·들여쓰기·문단 간격·줄 간격을 전후 비교한다.

## 일일·주간 계획서

1. A4 세로, 업무 성격에 맞는 여백을 `set_page_setup`으로 지정한다.
2. 문서 제목, 현장명, 날짜, 담당자, 시간대별 계획의 텍스트 골격을 만든다.
3. 담당자처럼 템플릿 필드가 있으면 `scope:"fields"`로 이름을 확인하고 `set_field_text`를 쓴다. 필드가 없으면 정확한 기존 문구를 치환한다.
4. 시간표는 실제 `insert_table`로 만들고 머리글 채움, 열 비율, 가운데 정렬을 지정한다.
5. 저녁 9시까지 확장할 때 기존 마지막 시간과 중복되지 않게 행을 하나씩 센다.
6. 제목·표 머리글·본문을 서로 다른 크기와 굵기로 지정하고 꼬리말과 쪽번호를 넣는다.

## 보고서·안전문서

- 장 제목은 `target.text`로 정확히 선택해 굵기·크기·간격을 지정한다.
- 본문은 양쪽 정렬, 문단 앞뒤 간격, 줄 간격, 첫 줄 들여쓰기를 함께 설정한다.
- 사진은 `insert_picture`의 `specific` 크기로 넣어 폭주를 막는다.
- 표 행/열 삭제는 high-risk 배치로 분리한다.
- 납품 전 `export_pdf`, Poppler 렌더링, 페이지별 육안 검수를 수행한다.

## 표 셀에 사진 넣기

표를 먼저 `scope:"tables"`로 읽어 대상 `tableIndex`와 셀 위치를 확인한다. 병합 표는 `cellIndex`를 사용한다.

```json
{
  "op":"insert_picture",
  "path":"C:\\현장사진\\사진-01.jpg",
  "tableIndex":2,
  "row":1,
  "col":0,
  "sizeOption":"cell-ratio",
  "embedded":true,
  "clearCell":false
}
```

기존 셀 제목이나 설명을 유지할 때는 `clearCell:false`를 사용한다. 사진만 남겨야 한다는 명시적 요청이 있을 때만 `clearCell:true`로 기존 셀 내용을 지운다. 적용 후 `scope:"structure"`에서 그림 control 증가를 확인하고 PDF로 셀 경계 안의 크기와 잘림을 검수한다.

## A4 한 장 공문·변경계 표 크기

1. `set_page_setup`으로 A4 세로와 여백을 먼저 확정한다.
2. `insert_table.columnWidths`에 모든 열의 상대 비율을 명시한다. 긴 자격·사유 열은 넓게, 구분 열은 좁게 계획한다.
3. 표 삽입 직후 모든 행을 하나씩 세되 적용은 `table_set_row_heights.rows` 한 op로 묶는다. 시작 기준은 머리글 9~10mm, 단문 본문 10~12mm, 설명 행 12~16mm이다.
4. 문서 기본 줄간격은 140~150%에서 시작하고 제목 125~135%, 표 셀 145~160%, 첨부 130~140%처럼 역할별로 조정한다.
5. 첨부 문단은 통상 제출일자·서명란보다 위에 배치한다. 문단 이동 후 `document_map`에서 `첨부 → 일자 → 수급인/서명` 순서를 확인한다.
6. 적용 응답의 행별 실제 높이와 `scope:"structure", includePageCount:true`의 `pageCount:1`을 모두 확인한다.

## 표형 조직도

1. 페이지 전체를 큰 표 한 개로 억지로 만들지 않는다. 연결선과 상자가 필요한 레이아웃은 셀 병합과 숨은 테두리를 계획해 표 격자로 구성한다.
2. 빨간 점선은 한글 편집 화면의 숨은 표 선 표시이므로 출력선으로 오해하지 않는다.
3. 상자 셀에는 필요한 바깥 실선만 지정하고 연결 통로 셀의 테두리는 숨긴다.
4. 안전총괄·관리자·협의체 등 역할별 채움색을 지정하되 인쇄 대비를 확인한다.
5. 각 셀 텍스트와 병합 범위를 하나씩 대조하고 PDF에서 선 연결이 끊기지 않았는지 확인한다.

## 기존 양식 필드

```json
{"scope":"fields","maxFields":100,"includeValues":true}
```

필드 확인 후 쓰기 op:

```json
{"op":"set_field_text","name":"담당자","text":"홍길동"}
```

필드가 없으면 새 필드를 만들지 말고 텍스트 위치를 확인한 뒤 `find_replace` 또는 셀 교체를 사용한다.
