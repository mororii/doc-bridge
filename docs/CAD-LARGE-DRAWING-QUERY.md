# AutoCAD 대형 도면 조회와 검증

DocBridge 0.4.7은 수만 개 객체가 있는 관망도와 종평면도에서 전체 순차 조회를 줄이기 위해 다음 조회 범위를 지원합니다.

- `scope: "layers"`: 레이어의 켜짐, 동결, 잠금, 출력, 색상, 선종류를 조회합니다.
- `scope: "xrefs"`: XREF 경로와 참조 수, 종속 레이어 상태, XCLIP 공간 필터를 조회합니다.
- `scope: "window"`: 지정한 좌표 창을 AutoCAD 네이티브 Window/Crossing 선택으로 조회합니다. 비활성 원본 도면은 일시 활성화하고 ModelSpace로 전환한 뒤 원래 상태로 복원합니다.
- `scope: "regions"`: 여러 검증 영역을 ModelSpace 한 번의 순회로 집계합니다.

## 권장 순서

1. `cad_get_active_context({"detailLevel":"basic"})`으로 대상 문서와 열린 원본 문서를 확인합니다. `basic`은 레이어와 ModelSpace 엔티티를 전혀 순회하지 않으므로 `summary.layers:[]`, `entitySummaryStatus:"omitted"`가 정상입니다. `layerCount`와 `entityCount`는 전체 개수입니다.
2. 레이어 최대 50개와 객체 유형 최대 500개 표본이 꼭 필요할 때만 `detailLevel:"summary"`를 사용합니다. 표본은 전체 집계가 아니며 `coverage.complete:false`이면 결론에 사용하지 않습니다.
3. 컨텍스트의 `nextActions`, `scope: "layers"`와 `scope: "xrefs"`로 필요한 관로·문자·기호 레이어와 XCLIP을 확인합니다. 페이지 조회의 `truncated:true`는 `nextStartIndex` 및 `nextActions[].arguments`로 끝까지 이어갑니다.
4. 원본 관로는 `scope: "window"`에 실제 좌표와 레이어·객체 형식을 함께 지정해 찾습니다. 반환 한도에 걸리면 `nextActions`의 `countOnly:true` 또는 확장된 `limit` 조회를 사용합니다.
5. 여러 도곽을 `scope:"regions"`로 센 결과의 핸들 샘플은 영역당 최대 20개입니다. `sampleCoverage.truncated:true`이면 함께 반환된 `scope:"window"` 후속 조회로 실제 목록을 확인합니다.
6. 쓰기 배치의 첫 연산은 항상 `activate_document`로 둡니다.
7. 작은 배치로 적용하고 각 배치의 readback을 확인합니다.
8. 새 객체가 공간 조회에 바로 나타나지 않으면 `zoom_window`로 Frame을 표시해 REGEN을 유도한 뒤 다시 조회합니다.
9. 적용 실패 후에는 같은 배치를 바로 반복하지 않습니다. 대상 영역을 재조회해 메모리에 남은 동일 좌표 객체가 있는지 확인합니다.
10. 중복이 있으면 자동 백업 후 정확한 핸들만 삭제합니다.

## 종평면도 수평 전개 검증

꺾인 관로를 수평으로 전개할 때는 각 구간의 실측 길이를 순서대로 누적해 X 좌표를 계산합니다. 9개 구간이면 다음 항목도 각각 9개여야 합니다.

- 관로 중심선
- 관로번호
- 관경·연장 제원
- 지시선
- 신설 기호

맨홀은 시작·끝을 포함하므로 기본 10개이며, 패널 경계에서 같은 맨홀을 양쪽 패널에 표시하면 화면 객체 수가 늘어날 수 있습니다. XREF 배경, 패널 외곽선, 제목, 도면번호, 축척, 키맵은 별도 영역으로 검증합니다.
