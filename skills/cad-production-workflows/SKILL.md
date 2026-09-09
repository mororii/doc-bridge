---
name: cad-production-workflows
description: AutoCAD 또는 GstarCAD DWG/DXF를 DocBridge 직접 ActiveX COM으로 분석·편집할 때 사용한다. AutoCAD는 도형·해치·블록·레이어·배치·출력, GstarCAD는 별도 도구로 기본 도형·문자·레이어 편집을 지원한다. AutoLISP와 화면 조작은 사용하지 않는다.
---

# AutoCAD / GstarCAD 실무 제작

## 제품을 먼저 구분한다

- AutoCAD는 `cad_*`, GstarCAD는 `gstarcad_*` 도구만 사용한다. GstarCAD 작업에서 아래 공통 읽기/기본 편집 지침의 `cad_*`는 `gstarcad_*`, `app:"cad"`는 `app:"gstarcad"`로 선택한다. 두 제품을 자동 대체 연결하지 않는다.
- GstarCAD는 `core_get_capabilities({"app":"gstarcad"})`의 지원 목록을 먼저 확인한다. 해치, 문서 간 복사, XREF, 블록 삽입, 배치/뷰포트 편집, PDF 출력, 등록 스크립트, RGB는 아직 지원하지 않는다. 아래 AutoCAD 고급 기능 지침을 GstarCAD에 적용하지 않는다. ACI 색상과 동일 도면 내 기본 편집을 사용한다.
- GstarCAD에서 필요한 전용 도구가 없으면 구버전일 수 있다. 원본 AutoCAD 도구나 외부 COM 스크립트로 우회하지 말고 업데이트 후 새 작업에서 확인한다.

모든 도면 판단은 좌표·핸들·레이어·객체 유형·bbox를 근거로 한다. 화면은 최종 확인용이며 편집 수단이 아니다.

## 도구 가용성 게이트

1. 작업 전에 `core_get_status`, `cad_get_active_context`, `cad_query_entities`, `cad_apply_ops`가 실제 도구 목록에 있는지 확인한다.
2. 하나라도 없거나 호출되지 않으면 즉시 중단하고 DocBridge가 로드되지 않았다고 알린다.
3. 프로젝트에 대체 COM 래퍼, PowerShell/Python 스크립트, SCR/LISP를 만들거나 수정해서 우회하지 않는다. computer-use나 DWG 파일 내부 편집으로 자동 전환하지 않는다.
4. `2-TEST.cmd` 실행, AI 프로그램 완전 종료·재실행, 새 작업 시작을 안내하고 도구가 보인 뒤에만 계속한다.

## 절대 규칙

1. 한 앱 작업은 `core_get_status({"app":"cad"})`로 시작한다. ping·전체 앱 조회를 반복하지 않는다. 지원 목록이 필요할 때만 `core_get_capabilities({"app":"cad"})`를 보고, `cad_get_active_context({"detailLevel":"basic"})`으로 열린 문서와 대상 DWG를 확인한다. `basic`은 대형 도면의 COM 엔티티·레이어를 순회하지 않으므로 `layers`가 빈 배열인 것이 정상이다. 레이어 미리보기와 최대 500개 유형 표본이 실제로 필요할 때만 `detailLevel:"summary"`를 사용한다.
2. 쓰기 배치 첫 op에 `activate_document`를 넣어 탭 전환 오작업을 막는다.
3. AutoLISP/LSP, WBLOCK 우회, computer-use, SendKeys를 사용하지 않는다.
4. XREF 자르기만 DocBridge 내부의 AutoCAD 기본 XCLIP을 허용한다.
5. 저장된 도면의 `set_text_value`/`set_layer_visibility`/`set_layer_color`/`regen_document`는 `executionMode=execute`와 새 UUID `requestId`, 저장된 절대 경로 `expectedDocumentRef`로 한 번에 적용한다. 현재 어댑터는 인스턴스 바인딩 참조를 발급하지 않는다. `dryRun`/`confirmToken`/`highRiskConfirm`을 넣지 않는다. 도형·이동·복사·줌·삭제·저장·내보내기·`activate_document`는 기존 dry-run → 같은 ops/`confirmToken`이다. dirty 광범위 작업이나 저장 지문이 맞지 않으면 거절한다. `Drawing1`은 저장해 경로를 만들지 않으며 지문 거절을 우회하지 않는다. 문자 핸들·레이어 스냅샷은 그 범위만 복구하며 전체 도면 복구가 아니다. Geometry 스냅샷은 drawing-backup과 부분 상태이며 자동 복구를 보장하지 않는다. 적용 후 응답 `readback`을 확인하고, 필요한 영역/핸들만 다시 조회한다.
6. 삭제·저장·PDF 출력·등록 스크립트는 토큰 경로에서 명시적 승인과 `highRiskConfirm:true`가 필요하다. `highRiskConfirm`은 권한 UI를 대체하거나 사람 승인을 증명하지 않는다. 이미 요청한 저장·PDF에 새 질문을 무조건 요구하지 않는다.
7. DocBridge는 다른 앱에서 작업 중인 사용자의 전경 창을 유지한다. AutoCAD API가 대상 도면이나 배치를 내부 활성화하더라도 배치가 끝나면 원래 도면·레이아웃·모델/종이 공간·보기 중심과 크기를 복원한다. 화면 클릭이나 강제 창 활성화로 보조하지 않는다.
8. 사용자가 다른 프로그램에서 계속 일하는 것은 허용하지만 같은 AutoCAD 창의 동시 조작은 허용하지 않는다. `interaction.userActivityDetected:true`, `interaction.interrupted:true`, 또는 `APP_USER_ACTIVITY_DETECTED`가 반환되면 남은 op를 완료됐다고 보고하지 말고 도면을 다시 조회한 뒤 미실행 단계만 새 dry-run으로 만든다. `foregroundPreserved:false`나 `originalStateRestored:false`이면 추가 쓰기 전에 활성 문서·레이아웃을 재확인한다.

## 분석을 먼저 한다

- `cad_get_active_context`의 `coverage`, `entitySummaryStatus`, `nextActions`를 먼저 확인한다. `basic`의 `entitySummaryStatus:"omitted"` 또는 빈 `layers`를 도면에 객체/레이어가 없다는 뜻으로 해석하지 않는다.
- 레이어 켜짐/꺼짐은 `cad_query_entities({"scope":"layers"})`의 `on`, 동결은 `freeze`, 잠금은 `locked`로 확인한다. `current`는 현재 작업 레이어이며 표시 여부와 다르다. `modelVisible`은 `on && !freeze`이고 뷰포트별 동결/객체 투명도/가림까지 보장하지 않는다. `null`은 조회 불가다. `layerSummaryStatus:"omitted"`이면 목록을 조회하고, `truncated`이면 끝까지 페이지를 읽는다.
- 큰 도면은 컨텍스트 표본으로 결론내리지 않고 `countOnly`, 레이어·유형·문자·bounds 필터와 `nextStartIndex`를 사용한다. `truncated:true`이면 응답의 실행 가능한 `nextActions[].arguments`로 계속 조회한다.
- 도곽이 여러 개면 `scope:"regions"`에 최대 100개 영역을 넣어 한 번에 객체수·유형·실제 bbox를 구한다.
- `scope:"regions"`의 각 영역은 핸들 표본을 최대 20개만 반환한다. `sampleCoverage.truncated:true`이면 그 영역의 `nextActions`가 제시하는 `scope:"window"` 조회로 실제 객체를 가져온다.
- 배치/뷰포트 작업 전 `scope:"layouts"`로 이름, 출력 설정, Target, CustomScale을 읽는다.
- 대상이 이미 존재하면 중복 작성하지 않는다. 교체는 정확한 핸들 또는 검증된 증분 시작 인덱스로만 한다.

## 제작과 수정

- 도형은 `draw_entities`의 12종 타입을 사용한다. 해치는 typed `AcadEntity[]` 경계로 직접 만들어지며 LISP를 쓰지 않는다.
- 같은 도면 사본은 `copy_entities`, 축척은 `scale_entities`, 대칭 사본은 `mirror_entities`, 평행선/곡선은 `offset_entities`를 쓴다.
- 레이어·색·선종류·선종류 축척·선가중치·표시는 `set_entity_properties`로 한 번에 변경한다.
- 도곽 블록의 제목·도면번호·축척은 가능하면 `set_block_attributes`로 태그를 정확히 지정한다. 일반 문자일 때만 `set_text_value`를 쓴다.
- 문서 간 객체는 `copy_entities_between_documents`의 typed ActiveX `CopyObjects` 경로를 쓴다.

## 편집 후 표시 문제

- 문자가 마우스를 올릴 때만 보이면 위치/크기/색상을 임의로 다시 바꾸지 않는다. 핸들 재조회(`includeGeometry:true`)로 실제 문자·좌표·색상·visible·transparency와 레이어 상태를 먼저 확인한다.
- 쓰기 후 `readback.displayRefresh.status`를 확인한다. 자동 `Regen(acAllViewports)`가 실패하면 이미 적용된 move/scale을 반복하지 말고 `regen_document`만 적용한다. 저장된 도면이면 execute, 미저장 이름은 토큰 경로다. `activate_document`는 execute가 아니다. 재생성은 좌표·문자 내용·색상을 변경하지 않는다.
- `readback.verified`는 데이터 검증, `displayRefresh:completed`는 API 갱신 완료다. 둘 다 눈으로 확인한 배치 품질/겹침 없음의 증거가 아니다. PDF 검수나 사용자의 화면 확인을 구분해서 보고한다.

## 배치와 출력

1. `configure_layout`으로 배치와 PDF 장치를 설정한다.
2. `create_viewport`에 종이공간 크기와 모델 `viewCenter/viewHeight`를 명시하고 잠근다.
3. `scope:"layouts"`로 실제 Target·CustomScale·핸들을 확인한다.
4. 사용자가 저장을 요청한 경우에만 `save_document`; PDF는 `plot_pdf`로 동기 출력한다.
5. PDF를 렌더링해 회전, 용지, 선, 문자, 여백을 시각 검수한다.

## 시공도·반복 모듈

보도블록·유도블록·패널처럼 배열이 중요한 도면은 점 몇 개를 추정해 반복하지 않는다. 실제 모듈 크기, 행·열 개수, 방향 전환, 부분 블록을 하나씩 센 명시적 좌표표를 만든 뒤 생성한다. 이미지의 외곽은 스케일 기준점으로 변환하고, 블록 배열은 규격 치수의 정수 격자에 맞춘다.

## 작업별 레시피

도형·수정·도곽 검증·배치/출력 예시는 [references/recipes.md](references/recipes.md)를 읽는다. 하수관로 종평면도는 별도 `cad-profile-sheet-pipeline` 스킬을 함께 사용한다.
