# DocBridge CAD 직접 COM 규칙

과거의 AutoLISP, WBLOCK, OOPS, 임의 PowerShell COM 스크립트보다 이 규칙이 우선한다.

## 허용 경로

- 읽기: `cad_get_active_context({"detailLevel":"basic"})`, `cad_query_entities`. 기본 컨텍스트의 빈 `layers`와 `entitySummaryStatus:"omitted"`는 미조회 상태이며 객체 없음이 아니다.
- 쓰기: `cad_apply_ops`
- 직접 편집 op: `copy_entities_between_documents`, `insert_xref`, 12종 `draw_entities`, `copy/scale/mirror/offset_entities`, `set_entity_properties`, `set_text_value`, `set_block_attributes`, `move_entities`, `rotate_entities`, `configure_layout`, `create_viewport`, `zoom_window`
- XREF 자르기: `insert_xref.clipBounds`를 통해 DocBridge가 AutoCAD 기본 `XCLIP`을 호출하는 경우만 허용

## 금지 경로

- `.lsp` 파일 생성 또는 로드
- `(command ...)`, `(load ...)`, `ssget`, `entget` 등 AutoLISP 실행
- LISP 또는 명령문 기반 WBLOCK/INSERT/EXPLODE 우회
- `run_script_template`
- computer-use, SendKeys, 화면 클릭으로 객체 작성
- AutoCAD 프로세스 강제 종료

## 안전 순서

1. 활성 문서와 열린 문서를 읽고 대상 경로와 ModelSpace 개수를 기록한다.
2. 범위·레이어·문자 필터로 소스 객체를 확정한다.
3. 같은 ops를 `dryRun: true`로 보내 diff, snapshotId, confirmToken을 받는다.
4. diff의 대상 문서, 좌표, 객체 수를 확인한다.
5. 동일 ops와 confirmToken으로 적용한다.
6. `verified: true`이고 `mismatches`가 비어 있는지 확인한다.
7. 시작 인덱스부터 증분 조회해 실제 bbox·문자·XREF 속성을 확인한다.

## 조회와 복사

- 큰 도면은 컨텍스트 표본 대신 `countOnly`, `layer`, `entityType`, `textContains`, `blockName`, `bounds`를 먼저 사용한다. `truncated:true`이면 `nextActions`로 이어 읽고, 여러 도곽은 `scope:regions`로 최대 100개를 한 번에 센다. 영역 샘플이 잘리면 `sampleCoverage`와 영역별 `nextActions`를 따른다.
- `inside`는 객체 전체가 박스 안이어야 할 때, `intersect`는 관로·도곽 선이 경계를 가로지를 때 사용한다.
- 같은 문서 내 복사는 직접 `Copy()`가 검증된 경로다.
- 문서 간 복사는 AutoCAD 설치본의 `AcadEntity[]` COM 배열과 공식 ActiveX `CopyObjects`를 우선 사용한다. 실패할 때만 지원 엔티티를 대상 ModelSpace의 `Add*` 메서드로 직접 재생성하며, 두 경로 모두 dry-run과 실제 증분 검증을 거친다.
- `CopyObjects`는 해치·블록의 종속 객체를 함께 깊은 복사할 수 있다. 적용 후 대상 ModelSpace 증가량이 주 객체 수보다 크면 원좌표 잔여를 핸들로 식별하고, 이동된 주 객체의 존재를 확인한 뒤 잔여만 별도 dry-run/apply로 삭제한다.
- 문서 간 복사가 실패하면 중단하고 LISP나 WBLOCK로 우회하지 않는다.

## XREF와 저장

- `insert_xref`는 기본적으로 `AttachExternalReference` 직접 COM 경로다. `reuseExistingDefinition: true`이면 기존 XREF 블록 정의를 `InsertBlock`으로 재사용한다.
- sourceFile, insertionPoint, scale, rotationDeg, layer, name을 명시한다.
- clipBounds가 있으면 AutoCAD 기본 XCLIP만 사용된다.
- 완성 Frame과 같은 동결 레이어 구성이 필요하면 기존 정의를 재사용하고, 삽입 후 XREF handle, 이름, 경로, 삽입점, 회전, 축척을 읽어 확인한다.
- 사용자가 저장을 명시하지 않으면 열린 DWG는 수정된 상태로 두고 저장하지 않는다. 저장과 `plot_pdf`는 high-risk 확인이 필요하다.
