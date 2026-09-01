# AutoCAD 작업 명세

DocBridge 0.4의 CAD 경로는 AutoCAD ActiveX COM을 직접 사용한다. 화면 좌표 조작과 AutoLISP/LSP 생성·로드는 사용하지 않는다. XREF 경계가 필요한 경우에만 AutoCAD 기본 `XCLIP` 명령을 사용한다. 쓰기는 항상 dry-run 스냅샷, 확인 토큰, op별 readback, 실패 시 자동 롤백 순서이다.

## 읽기와 일괄 검증

- `cad_get_active_context`: 기본 `detailLevel:"basic"`은 열린 도면·전체 레이어/엔티티 개수·단위만 읽고 레이어와 ModelSpace를 순회하지 않는다. 따라서 `layers:[]`는 "없음"이 아니라 "생략"이다. `summary`는 레이어 50개와 엔티티 500개 유형 표본만 읽는다. 두 모드 모두 `coverage`, `entitySummaryStatus`, `nextActions`로 완전성을 표시한다.
- `cad_query_entities`: `document`, `layer`, `entityType`, `textContains`, `blockName`, `bounds`, `boundsMode`, `startIndex`, `endIndex`, `limit`, `includeGeometry`를 지원한다. `truncated`, `coverage`, `nextStartIndex`, `nextActions`로 계속 읽는다.
- `scope:"layouts"`: 배치, 출력 장치/용지, 뷰포트 핸들·크기·Target·CustomScale을 읽는다.
- `scope:"regions"`: 최대 100개 영역을 ModelSpace 한 번의 순회로 집계한다. 각 항목은 `name`, `bounds`와 선택적인 `boundsMode`, `entityTypes`, `layer`, `textContains`, `minCount`, `maxCount`를 받으며 객체수·유형별 수·실제 bbox·샘플 핸들을 반환한다. 샘플은 최대 20개이며 `sampleCoverage.truncated:true`일 때 영역별 `nextActions`로 `scope:"window"` 실제 목록을 조회한다. 종평면도 도곽, 평면 박스, 종단 박스, 제목, 키맵 검증에 사용한다.

## 도형 생성

`draw_entities.entities`에는 다음 유형을 섞어서 최대 1,000개까지 넣는다.

- `lwpolyline`: `points`, 선택 `closed`, `color`, `layer`.
- `circle`: `center`, 양수 `radius`.
- `hatch`: 닫힌 `loop.points`, 선택 `loop.bulges`. 임시 경계와 typed `AcadEntity[]` SAFEARRAY로 솔리드 해치를 직접 만들고 임시 경계는 삭제한다.
- `block`: `point`, 기존 블록 `name`, 축척과 회전.
- `text`: `point`, `text`, `height`, 선택 글꼴 스타일·정렬·회전.
- `line`: `start`, `end`.
- `arc`: `center`, `radius`, `startAngleDeg`, `endAngleDeg`.
- `ellipse`: `center`, 중심에서 장축 끝으로 향하는 `majorAxis`, `radiusRatio`(0~1].
- `point`: `point`.
- `mtext`: `point`, `width`, `text`, 선택 `height`, `rotationDeg`, `attachmentPoint`.
- `dim_aligned`: `start`, `end`, `textPoint`.
- `dim_rotated`: `start`, `end`, `dimensionLinePoint`, `rotationDeg`.

공통으로 `layer`, `color`(`aci` 또는 `rgb`), `linetype`, `linetypeScale`, `lineweight`, `visible`을 지정할 수 있다.

## 수정·블록·문서 간 작업

- `copy_entities`: `handles`, `dx`, `dy`로 같은 도면에 복사한다.
- `scale_entities`: `handles`, `basePoint`, 양수 `factor`.
- `mirror_entities`: `handles`, `axisStart`, `axisEnd`; 원본을 보존하고 대칭 사본을 만든다.
- `offset_entities`: `handles`, 0이 아닌 `distance`.
- `set_entity_properties`: `handles`와 공통 `properties`.
- `set_block_attributes`: 블록 참조 `handle`과 `{태그: 값}`의 `attributes`; 변경 후 조회하면 속성값도 반환된다.
- `copy_entities_between_documents`: 열린 원본/대상, 핸들·레이어·유형·bounds 선택, 원점/축척/회전을 받아 ActiveX `CopyObjects` typed 배열로 깊은 복사한다.
- `insert_xref`: XREF 정의 재사용과 배치 변환을 지원한다.

## 배치·뷰포트·출력

- `configure_layout`: `name`, 선택 `create`, `configName`, `canonicalMediaName`, `plotRotation`, `plotType`, `centerPlot`, `standardScale`, `useStandardScale`.
- `create_viewport`: `layout`, 종이공간 `center`, `width`, `height`, 모델 `viewCenter`, `viewHeight`, 선택 `twistAngleDeg`, `displayLocked`. AutoCAD 2027 호환을 위해 `Target`과 `CustomScale=height/viewHeight`를 사용한다.
- `save_document`: `output`이 있으면 SaveAs, 없으면 현재 도면 저장. 원본/기존 파일에 영향을 줄 수 있어 high-risk 확인이 필요하다.
- `plot_pdf`: 절대 `.pdf` `output`, 선택 `configName`. `BACKGROUNDPLOT=0`으로 동기 출력하고 비어 있지 않은 파일을 확인한다. high-risk 확인이 필요하다.

## 검증 결과

- 일반 회귀: Core 70건 + MCP 17건.
- 실제 AutoCAD 2027: 해치 포함 12종 도형, 복사·축척·대칭·offset·속성, 블록 속성, 두 영역 일괄 검증, 배치·뷰포트, PDF 출력과 새 DWG 저장을 임시 도면에서 통과했다.
- PDF 증거: `reports/cad-e2e/cad-production-e2e.pdf`와 렌더링 PNG.
