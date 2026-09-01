---
name: cad-profile-sheet-pipeline
description: AutoCAD에서 관망도와 종단면 자료를 분석해 빈 종평면도 도곽에 평면도·종단면도·제목·도곽 타이틀·키맵을 배치하고 검증한다. DocBridge의 직접 ActiveX COM 편집만 사용하며 AutoLISP, LSP 생성, WBLOCK 우회, 컴퓨터 화면 조작은 사용하지 않는다. 하수관로 종평면도 작성, 관망도 노선 추출, 빈 도곽 채우기, 키맵 갱신 요청에 사용한다.
---

# CAD 종평면도 직접 편집 파이프라인

DocBridge MCP로 열린 AutoCAD 문서를 직접 읽고 편집한다. 화면 좌표를 클릭하거나 LISP를 생성하지 말고, 객체의 실제 좌표·문자·레이어·경계를 근거로 작업한다.

## 도구 가용성 게이트

1. 시작 전에 `core_get_status`, `cad_get_active_context`, `cad_query_entities`, `cad_apply_ops`가 실제로 제공되는지 확인한다.
2. 필요한 도구가 없으면 종평면도 작업을 시작하지 않는다. 프로젝트의 스크립트·플러그인·도면 파일을 임의 수정해 기능을 재구현하지 않는다.
3. DocBridge 미로드 상태를 보고하고 `2-TEST.cmd` 실행, Codex/Claude/Kimi 완전 재시작, 새 작업 시작을 안내한다.
4. 도구가 다시 보이는 것이 확인된 뒤에만 이 파이프라인의 분석 단계로 진행한다.

## 절대 규칙

1. `cad_get_active_context({"detailLevel":"basic"})`으로 활성 문서와 열린 DWG 목록을 먼저 확인한다. 기본 응답의 빈 `layers`는 성능을 위해 생략한 것이므로 레이어가 없다고 판단하지 않는다.
2. 사용자가 지정한 대상 DWG가 열린 문서 목록에 있는지 확인하고, 대상이 불명확하면 쓰지 않는다.
3. 모든 `cad_apply_ops` 배치의 첫 연산은 `{"op":"activate_document","document":"대상문서패턴"}`으로 두어 다른 DWG 탭이 활성화돼도 잘못된 문서를 수정하지 않게 한다.
4. 읽기는 `cad_query_entities`, 쓰기는 `cad_apply_ops`만 사용한다.
5. AutoLISP 코드를 만들거나 실행하지 않는다. `.lsp`, `(load ...)`, `run_script_template`, LISP 기반 `WBLOCK`, `OOPS`를 사용하지 않는다.
6. 프로그램 화면 클릭이나 computer-use로 도면을 편집하지 않는다.
7. XREF 경계 자르기에 한해 DocBridge가 내부적으로 호출하는 AutoCAD 기본 `XCLIP` 명령을 허용한다. 이것은 AutoLISP가 아니다.
8. 모든 쓰기는 동일한 ops로 `dryRun: true`를 먼저 실행하고, 반환된 `confirmToken`으로 `dryRun: false`를 실행한다. 적용 결과의 `verified`와 `mismatches`를 확인한다.
9. 적용 전 스냅샷 ID와 대상 문서 경로를 기록한다. 사용자가 저장을 명시하지 않으면 DWG를 강제 저장하지 않는다.
10. 기존 종단면도·평면도·제목이 있으면 중복 배치하지 않는다. 삭제·교체는 사용자의 범위 승인과 정확한 핸들 또는 증분 인덱스가 있을 때만 수행한다.
11. 도면 전체를 감으로 배열하지 않는다. 노선, 체인리지, 객체 수, 블록 모듈, 키맵 칸을 하나씩 대조한다.

## 필요한 참조

- 프로젝트 고유 도곽 좌표와 노선 정보는 [references/project-data.md](references/project-data.md)를 읽는다.
- 직접 COM 작업과 검증 규칙은 [references/cad-com-rules.md](references/cad-com-rules.md)를 읽는다.
- 키맵 칸 좌표는 [references/keymap-tiles.csv](references/keymap-tiles.csv)를 읽는다.

## 1. 문서와 작업 대상 확정

1. `cad_get_active_context({"detailLevel":"basic"})`를 호출한다. `summary.openDocuments`, `entityCount`, `layerCount`로 대상만 확정하고 객체 판단은 하지 않는다.
2. 반환된 활성 문서 경로, 열린 문서 목록, ModelSpace 개수를 기록한다.
3. 대상 종평면도, 관망도, 종단면 원본을 이름과 경로로 구분한다.
4. `cad_query_entities`의 `document`에 문서명 또는 경로 패턴을 넣어 각 문서를 읽는다.
5. 대량 도면은 먼저 `scope:"layers"`와 `scope:"xrefs"`로 호스트/XREF 레이어와 XCLIP 상태를 확인하고, 좁은 좌표 구간은 `scope:"window"`로 읽는다. `countOnly: true` 또는 레이어·문자·경계 필터를 함께 사용하며 전체 엔티티를 무조건 열거하지 않는다.
6. 쓰기 전에는 활성 탭을 가정하지 말고 `activate_document`를 첫 op로 넣는다. `cad_apply_ops`의 건식 검증에서 대상 문서명이 반환되는지 확인한다.

## 2. 빈 도곽과 기존 내용 판정

1. 프로젝트 도곽 격자식으로 목표 Frame의 `X0`, `Y0`를 계산한다.
2. 평면도 박스, 종단면 박스, 상단 제목, 하단 도면번호, 키맵 영역을 `cad_query_entities(scope:"regions")` 한 호출의 별도 region으로 검사한다. 각 region에 `minCount`/`maxCount`, 유형·레이어 조건을 넣어 누락과 중복을 동시에 찾는다. `sampleCoverage.truncated:true`인 영역은 반환된 `nextActions`의 `scope:"window"` 조회까지 수행한다.
3. `boundsMode: inside`와 `intersect`를 필요에 맞게 구분한다.
4. `textContains`로 노선명, 도면번호, 분구명, 축척 문자를 조회한다.
5. 박스 안에 목표 노선의 종단면 또는 평면 XREF가 이미 있으면 해당 부분은 건너뛴다.

## 3. 종단면도 직접 배치

1. 원본 문서에서 목표 노선과 체인리지 범위를 문자와 지오메트리로 찾는다.
2. `sourceBounds`가 모든 필요한 객체를 포함하는지 객체 수와 합산 bbox로 확인한다.
3. 기준점 `sourceOrigin`과 대상 `targetOrigin`을 정하고 이동량을 계산한다. 임의 눈대중 이동은 금지한다.
4. `copy_entities_between_documents`를 사용한다. 가능한 경우 `sourceDocument`, `targetDocument`, `sourceBounds`, `sourceOrigin`, `targetOrigin`, `scale`, `rotationDeg`를 명시한다.
5. 같은 문서 안 복사는 엔티티별 ActiveX `Copy()`로 직접 처리된다. 문서 간 복사는 AutoCAD 설치본의 `AcadEntity[]` COM 배열과 공식 ActiveX `CopyObjects`를 우선 사용해 해치·블록 종속성을 포함한 원본 객체를 깊은 복사한다. 이 경로가 실패할 때만 지원 엔티티를 대상 ModelSpace의 `Add*` 메서드로 직접 재생성한다.
6. 문서 간 복사에서 지원되지 않는 엔티티가 발견되면 작업을 중지하고 오류를 보고한다. LISP·WBLOCK로 우회하지 않는다.
7. 적용 후 시작 ModelSpace 인덱스부터 `cad_query_entities(startIndex=...)`로 증분 조회해 객체 수와 대상 bbox를 검증한다. `CopyObjects`가 종속 객체를 추가해 예상 수보다 많아지면 원좌표에 남은 객체를 핸들로 식별하고, 실제 이동된 객체가 별도로 존재하는지 확인한 뒤 해당 잔여만 건식 검증 후 삭제한다.

## 4. 관망도 평면 구간 계산과 배치

1. 관망도에서 목표 관로번호 문자열과 관로 중심선/맨홀 좌표를 조회한다.
2. 목표 노선의 시작·끝 체인리지와 연결 순서를 확인한다. 인접 노선을 섞지 않는다.
3. 노선 중심선의 꺾임점을 실제 좌표와 연결 길이로 하나씩 센다. 방향이 달라지는 구간은 각 직선 계열별로 분리하고, 단일 전역 회전으로 억지로 맞추지 않는다.
4. 각 구간의 원본 좌표 두 점 `A`, `B`와 대상 평면 박스의 대응점 `a`, `b`를 사용해 회전각과 삽입점을 계산한다. 분할 구간은 좌→우 체인리지 순서로 별도 XREF를 삽입하고 각 패널을 따로 XCLIP한다.
5. 스케일은 프로젝트 기준이 달리 지정되지 않으면 1:1로 유지한다.
6. `insert_xref`로 관망도 DWG를 삽입한다. `sourceFile`, `insertionPoint`, `rotationDeg`, `scale`, `layer`, `name`, `clipBounds`를 명시한다. 완성 Frame과 같은 XREF 동결 레이어를 유지해야 하면 `reuseExistingDefinition: true`와 `existingDefinition`을 사용한다.
7. `clipBounds`는 해당 분할 평면 패널의 표시 박스와 정확히 같게 둔다. DocBridge가 기본 `XCLIP`으로 자른다.
8. 삽입 후 각 XREF의 이름·경로·삽입점·회전·스케일·패널 bbox를 조회해 계산값과 비교한다. 새 이름이 붙으면 종속 레이어 상태가 달라질 수 있으므로 이름도 반드시 확인한다.
9. 사용자가 꺾인 관로를 수평 전개하라고 지정하면 XREF 배경만 회전해 끝내지 않는다. 각 관로 길이를 시작점부터 하나씩 누적해 대상 기준선의 X 좌표를 만들고, 패널 경계의 동일 맨홀은 양쪽 패널에 표시한다. 그 위에 직접 관로선·맨홀·관로번호·제원·지시선을 별도 소배치로 그린다.
10. 관로번호와 제원은 구간 수와 각각 1:1이어야 한다. 예를 들어 9구간이면 관로선 9개, 번호 9개, 제원 9개, 지시선 9개를 확인하고 합계만 맞춘 채 순서를 추정하지 않는다.

## 5. 제목·도곽 타이틀·키맵 수정

1. 기존 도곽의 문자 스타일, 높이, 정렬, 레이어를 인접 완성 Frame에서 읽는다.
2. 도곽이 속성 블록이면 `set_block_attributes`로 TITLE·SHEET·SCALE 같은 태그를 갱신한다. 일반 문자일 때만 `set_text_value`, 새 문자는 `draw_entities`의 `text`를 사용한다.
3. 제목 장식이나 심벌은 완성 Frame의 정확한 핸들을 조회한 뒤 같은 문서 내 `copy_entities_between_documents`로 직접 복사한다. 같은 블록을 여러 위치에 반복할 때는 `draw_entities`의 `block` 타입을 사용해 한 번에 삽입한다.
4. Frame 번호, 전체 매수, 노선명, 분구명, 수평/수직 축척을 서로 대조한다.
5. 키맵은 `keymap-tiles.csv`에서 실제 노선과 각 분할 평면 패널이 통과하는 도엽을 하나씩 대조한다. 통과 도엽이 한 칸이면 한 칸, 두 칸 이상이면 확인된 칸만 각각 닫힌 `lwpolyline`으로 표시하고 기존 레이어·색 규칙을 따른다.
6. 키맵 칸 수와 위치를 추정하지 말고 인접 완성 Frame의 도엽 표시와 원본 노선 좌표로 검증한다.

## 6. 검증과 사용자 확인

1. 쓰기 직전 ModelSpace 개수를 기록한다.
2. 적용 후 `cad_query_entities`의 `startIndex`와 `endIndex`로 새 객체만 조회한다.
3. `scope:"regions"`로 종단·평면·제목·도면번호·키맵을 한 번에 다시 검사하고, 노선/체인리지, XREF 회전·클립, 블록 속성, 타일 bbox와 객체 수를 모두 검증한다.
4. `zoom_window`로 완성 Frame 전체가 보이게 한다. 화면은 사용자 육안 확인용이고 좌표 검증을 대신하지 않는다.
5. 새 객체를 만든 직후 `scope:"window"` 결과가 적으면 `zoom_window`로 해당 Frame을 표시해 AutoCAD REGEN을 유도한 뒤 다시 조회한다.
6. 적용 실패와 자동 롤백이 발생하면 즉시 같은 배치를 반복하지 않는다. REGEN 후 대상 영역을 재조회하고 동일 좌표·레이어의 잔여 핸들을 확인한다. 잔여가 있으면 정확한 핸들만 새 스냅샷과 고위험 확인으로 삭제한다.
7. 검증 실패 시 새로 추가한 객체의 정확한 핸들 또는 시작 인덱스로만 되돌린다.
8. 결과 보고에는 대상 Frame, 노선, 추가/수정 객체 수, 스냅샷 ID, 저장 여부, 검증 결과를 포함한다.

## 금지된 레거시 방식

과거 PowerShell 스크립트의 LISP·WBLOCK 우회를 실행하지 않는다. 필요한 기능이 DocBridge에 없으면 플러그인의 직접 COM 연산을 먼저 보완하고 테스트한 다음 사용한다.
