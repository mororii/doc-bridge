# DocBridge 0.4.19 — CAD 화면 재생성·레이어 상태 조회

2026-09-03 · Windows x64 · Excel / 한글 / AutoCAD 공통 MCP

## 이번 수정

- 문자 이동·축척 뒤 마우스를 올려야 보이던 사례에서 `REGENALL` 직후 표시가 회복됐습니다. 기존 일반 편집 배치의 종료 경로에는 재생성이 없었습니다. 데이터 삭제나 투명도 변경으로 단정하지 않고 화면 갱신 문제로 처리합니다.
- 변경 도면마다 배치 종료 시 직접 ActiveX `Regen(1 /* acAllViewports */)`을 추가했습니다. 원래 보기 복원 뒤, 전경 보호가 끝나기 전에 실행합니다. 기존 복사/작성 helper의 필요한 중간 재생성은 유지하며 이동/축척마다 전체 도면을 재생성하지 않습니다.
- `readback.displayRefresh`에 완료/실패, 문서별 결과와 소요 시간을 반환합니다. 재생성 오류는 경고로 표시하며 이미 완료한 편집을 롤백하거나 반복하지 않습니다. `regen_document`는 동일 dry-run → snapshot → confirmToken → apply 경로에서 화면만 갱신합니다.
- CAD 적용 중 COM 호출 거부가 생겼을 때 전체 배치를 처음부터 자동 반복하지 않습니다. 원래 보기와 같으면 레이아웃·공간·확대를 불필요하게 다시 설정하지 않습니다.
- `cad_query_entities(scope="layers")`와 summary는 `current`(현재 작업), `on`(켜짐), `freeze`(동결), `locked`(잠금), `plottable`, `color`, `linetype`, `modelVisible`을 반환합니다. basic은 계속 빠르게 목록을 생략하되 `layerSummaryStatus:"omitted"`와 `currentLayer`를 명시합니다.
- 레이어 속성 조회 실패는 누락/false 대신 `null`과 `unavailableProperties`로 표시합니다. `modelVisible`은 `on && !freeze`이며 뷰포트별 동결, 객체 숨김·투명도, 가림, 화면 상태까지 보장하지 않습니다.
- 엔티티 `includeGeometry:true`에는 `visible`, ACI `color`, `transparency`와 미지원 속성 목록을 추가했습니다. ByLayer/ByBlock 값은 실제 화면 색으로 임의 변환하지 않습니다.
- CAD 스킬과 초보자 설명서에 표시 문제 진단 순서를 반영했습니다. 표시 오류만으로 문자 위치·크기·색상을 다시 바꾸지 않습니다.
- 공개 ZIP은 디버그 심볼을 제외하고 빌드 소스 경로를 정규화합니다. 바이너리까지 UTF-8/UTF-16으로 검사하여 빌드 PC의 실제 사용자/작업 경로가 포함되면 패키징을 중단합니다.

## 앞서 승인한 최적화도 포함

스키마 오류의 기대 형태/필수·선택 필드를 함께 안내하고 여러 오류를 모아 반환합니다. 같은 문서·같은 ops의 반복 dry-run은 지문 검증을 통과한 경우에만 snapshot/preview를 재사용합니다. 한 호출 내 중복 상태 조회를 줄이고 HMAC 초기화와 한글 행 높이 묶음 작업 안내를 보강했습니다. dry-run 생략, 문서 identity 검사 제거, HMAC 바인딩 약화는 없습니다.

CAD 미저장 도면의 작업 대상 지문은 지원되는 문자·레이어·재생성 작업에 한정합니다. 영역/외부 복사/배치/일반 곡선 등 범위를 충분히 검증할 수 없는 경우 재사용하지 않으며, 대상 도면을 저장한 뒤 새 dry-run이 필요합니다. 조회 실패한 대상은 같은 값으로 간주하지 않습니다.

## 검증과 한계

- 비-E2E: Core 218개 + MCP 19개 통과, 빌드 경고/오류 0개.
- 실제 AutoCAD: 별도 임시 DWG 1회(약 4초). 문자 이동·2배 축척 readback, 자동 재생성, 명시적 재생성 후 데이터 불변, 레이어 켜짐 변경, 현재/꺼짐/동결/잠금 상태 조회 통과.
- 임시 도면은 종료했고 기존 도면들의 경로·저장 상태·객체 수가 유지됨을 검사했습니다. 사용자 작업 도면의 문자 배치는 추가 변경하지 않았습니다.
- `readback.verified`는 데이터 검증, `displayRefresh.status:"completed"`는 갱신 API 완료입니다. 도면 전체의 육안 품질/겹침 없음이나 모든 GPU 환경의 표시 문제 해결을 보장하지 않습니다.

## 업데이트

`DocBridge-0.4.19-win-x64.zip`을 새 폴더에 풀고 기존 작업을 저장한 뒤 AI 프로그램을 완전히 종료하십시오. `0-VERIFY.cmd` → `1-INSTALL.cmd` → `2-TEST.cmd`를 실행하고 AI를 재시작하여 새 대화에서 `core_ping` 버전 `0.4.19`를 확인합니다. 실행 중인 이전 MCP에는 새 코드가 자동 반영되지 않습니다.

레이어 확인 요청 예시: “DocBridge로 현재 도면의 레이어 전체를 끝까지 조회해서 현재 작업/켜짐/동결/잠금 상태를 구분해 보여줘. 수정하지 마.”

## 기술 근거

Autodesk는 ActiveX로 변경한 도형이 즉시 화면에 나타나지 않을 수 있어 Update/Regen을 사용하도록 설명합니다. [Updating Geometry](https://help.autodesk.com/cloudhelp/2019/ENU/AutoCAD-ActiveX/files/GUID-8D9140F0-3676-41D7-B618-27D94AB9DE7E.htm)

잠금은 표시 여부와 별개이며 켜지고 해동된 잠금 레이어는 계속 보입니다. [Lock Property](https://help.autodesk.com/cloudhelp/2024/ENU/AutoCAD-ActiveX-Reference/files/GUID-49CA344E-0F8C-4AB2-8336-9E696F8BD5D7.htm)
