# DocBridge 0.4.22 릴리즈 노트

기준일: 2026-09-17. 이전 공개 릴리즈 v0.4.21 이후의 변경을 묶었다.
VBA·매크로는 범위 밖이다(`run_macro` 금지 유지).

## 새로 추가된 Excel 쓰기 op 16종

외부 링크 3종(`update_external_links`, `change_link_source`,
`break_external_link`), 계산 모드(`set_calculation_mode`),
값 고정(`freeze_values`), 선택 붙여넣기(`paste_special`, 클립보드 미사용),
목표값 찾기(`goal_seek`), 통합문서 보호/해제(`protect_workbook`,
`unprotect_workbook`, 무암호), 창 분할(`set_split_panes`), 스파크라인 3종,
슬라이서 2종(`create_slicer`는 `name` 필수), 셀 스타일(`apply_cell_style`).

## 새로 추가된 조회 scope 4종

`links`, `sparklines`, `slicers`, `cellStyles`.
`set_sheet_visibility`는 `veryHidden`을 추가로 받는다.

## 지정 창 고정 (instance pin)

여러 Excel이 열려 있을 때 자동 선택이 엇갈린다는 현장 지적을 반영했다.
`excel_launch`에 `processId`/`hwnd` 쌍·단독 또는 `activeWindow:true`로 창을
지정하면 `%LOCALAPPDATA%\DocBridge\excel-instance-pin.json`에 저장되고,
이후 모든 호출이 그 창에만 붙는다. 핀 창이 사라지면 다른 창을 몰래 쓰지
않고 `[EXCEL_PINNED_INSTANCE_GONE]`으로 중단한다. Excel이 최상위 창 핸들을
재생성하면 프로세스 시작 시각으로 동일 프로세스를 확인해 핀을 자가치유한다.
`excel_launch`는 워커 경로(`--excel-worker`)로도 연결되므로 CLI·MCP 모두
동작한다. 핀 없이 호출하면 종전 자동 선택 그대로다.

## 테두리 공유 변 간섭 방지

인접 범위는 변 객체를 공유하므로, 지우기가 먼저 그린 아웃라인을 지울 수
있다. 같은 배치에서 공유 변에 그리기 뒤 지우기가 오면
`[EXCEL_BORDER_ORDER]`로 거절한다(지우기→그리기 순서는 허용). 다른 배치의
간섭은 dry-run 경고 `[EXCEL_BORDER_SHARED_EDGE]`로 알린다.

## 검증 근거 (로컬, 공개 전)

- 통합 Release `-warnaserror` 경고 0·오류 0.
- 비-E2E Core 994·MCP 20 (필터 `FullyQualifiedName!~E2ETests&Category!=E2E`).
- 버전 일치·공개 소스·초보자 가이드 검사 통과.
- 격리 Excel COM 실검증: 주간 작업계획서(생성→입력→재계산→저장→재열기→PDF,
  전 batch readback verified), 자재검수요청서(A4 세로 1쪽, PDF目视),
  4시트 검측 세트(갑지/검측리스트/공사참여자명부/사진대지, 4쪽),
  실제 업무 파일의 다음 순번 작성(저장 없음).
- `winLoss` 스파크라인·암호 통합문서·고정 너비 텍스트 나누기는 fail-closed로
  거절한다.

## 남은 한계

- Power Query/외부 데이터 연결(링크 새로고침 제외), 타임라인,
  시나리오/데이터 테이블, 테마 편집, 머리말·바닥글 그림, 폼 컨트롤은 미구현.
- 실제 COM E2E CI와 설치 후 클라이언트 재시작은 릴리즈 뒤에 확인한다.
