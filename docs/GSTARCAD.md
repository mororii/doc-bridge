# GstarCAD 별도 연결 (0.4.20)

AutoCAD는 기존 `cad_*`, GstarCAD는 새 `gstarcad_*` 도구를 사용합니다. 같은 이름의 DWG가 두 제품에 열려 있어도 서로 대신 연결하지 않습니다. COM 연결, 승인 토큰의 `apply:gstarcad` 범위, `snapshots/gstarcad` 백업은 AutoCAD와 분리됩니다. 안전을 위해 서버 간 COM 작업 직렬화는 유지합니다.

## 사용

1. 업데이트 ZIP을 설치하고 AI 클라이언트를 완전히 재시작한 후 새 작업을 엽니다.
2. GstarCAD에서 도면을 엽니다. 읽기/상태 조회는 새 CAD 창을 만들지 않습니다.
3. 다음처럼 요청합니다.

> DocBridge의 GstarCAD 전용 도구로 열린 도면과 레이어 상태를 확인해 줘. AutoCAD는 건드리지 마.

> GstarCAD의 [저장된 도면 절대 경로]에서 [레이어 이름]의 문자만 높이를 2배로 키워 줘. 먼저 대상 핸들과 개수를 확인하고 적용한 뒤 응답 readback과 필요한 범위만 다시 읽어 검증해 줘. Drawing1처럼 미저장 이름은 저장해 경로를 만들지 말고 토큰 경로를 써 줘.

도구는 `gstarcad_get_active_context`, `gstarcad_query_entities`, `gstarcad_apply_ops`, `gstarcad_launch`입니다. 한 앱 작업은 `core_get_status({"app":"gstarcad"})`로 시작하고 ping·전체 앱 조회를 반복하지 않습니다. `core_get_capabilities({"app":"gstarcad"})`로 현재 설치본의 정확한 지원 범위와 `autoExecuteOps`를 확인합니다. `core_get_status.apps.gstarcad`와 `apps.cad`는 별개입니다. 설치된 GstarCAD의 `Gcad.Application` COM 등록을 사용하며, 다른 버전/에디션에서 COM 미등록·권한 불일치가 있으면 연결 실패를 보고합니다.

execute는 `set_text_value`, `set_layer_visibility`, `set_layer_color`, `regen_document`만입니다. UUID `requestId`와 저장된 도면의 절대 경로 `expectedDocumentRef`가 필요하고 `dryRun`/`confirmToken`/`highRiskConfirm`을 넣지 않습니다. 현재 어댑터는 인스턴스 바인딩 참조를 발급하지 않습니다. 도형 작성·이동·복사·줌·삭제·저장·내보내기·문서 활성화는 기존 dry-run → confirmToken입니다. dirty 광범위 작업이나 저장 지문이 맞지 않으면 거절합니다. `Drawing1`은 경로를 만들려고 자동 저장하지 않으며, 지문 거절을 우회하지 않습니다.

## 기본 지원 범위와 제한

- 읽기: 열린 도면, 엔티티·문자·좌표, 레이어 켜짐/동결/잠금/현재 상태, 영역별 조회, 배치/XREF 정보.
- 편집: 문자 변경, 이동·회전·축척, 같은 도면의 복사·대칭·오프셋, 기본 객체 속성, 레이어 표시·ACI 색상, 기존 블록 속성 값.
- 작성: 선, 원, 폴리선, TEXT/MTEXT, 호, 타원, 점, 정렬/회전 치수. 모든 타입을 모든 제품 버전에서 검증한 것은 아닙니다.
- 명시적 고위험 승인: 객체 삭제, DWG 저장/다른 이름으로 저장.
- **지원하지 않는 기능:** 해치, 문서 간 복사, 새 블록 삽입, XREF 삽입/자르기, 배치·뷰포트 편집, PDF 출력, 등록 스크립트, RGB 색상. 미지원 op가 배치 뒤에 있어도 전체 배치를 COM 편집 전에 거부합니다.

현재 소스의 geometry production은 GstarCAD 25.382초, AutoCAD 4.204초의 단일 표본입니다. 선/원/폴리선, 한글·Unicode TEXT/MTEXT, ACI, 이동/축척/복사/대칭/offset, 문자 변경, 레이어, regen, 깨끗하게 저장된 소유 도면 SaveAs/재조회를 확인했습니다. 기존 사용자 도면 경로·저장 플래그·객체 수는 보존했습니다. execute 최종 fixture 2행은 GstarCAD 23.390초, AutoCAD 4.281초로 통과했습니다. 문자·레이어·regen, native 롤백, UUID 재전송, 도형 execute 거절, dirty SaveAs 거절(출력 없음), 깨끗한 소유 도면 SaveAs 성공을 확인했습니다. 두 시간은 다른 표본이며 합산하지 않습니다. 해치·문서 간 복사·XREF·배치·PDF·스크립트·RGB는 미지원이며 검증했다고 쓰지 않습니다. 당시 실측(GstarCAD 34.231초, AutoCAD launch 재시도 5.316초)은 역사적 게이트이며 「확인된 실측」과 합산하지 않습니다. AutoCAD와 GstarCAD는 도구·토큰·스냅샷이 분리됩니다. CAD는 데이터/화면 갱신 검증이며 육안 품질 보장이 아닙니다.

비활성 창 작업의 안전장치는 유지합니다. 같은 GstarCAD 창에서 사용자 입력이 겹치면 중단할 수 있습니다. 다른 앱에서 작업하고, 중단된 이동/축척을 그대로 반복하지 마세요. `readback.verified`는 데이터 검증, `displayRefresh`는 API 화면 갱신이며 육안 배치 검수를 뜻하지 않습니다. 요청한 문자 핸들·레이어 속성 스냅샷은 그 범위만 완전 복구하며 전체 도면 복구가 아닙니다. Geometry 스냅샷은 drawing-backup과 부분 상태이며 자동 복구 성공을 보장하지 않습니다.

GstarCAD COM 제공 및 이식의 근거: [공식 개발자 안내](https://www.gstarcad.net/developer/), [공식 SDK 문서](https://www.gstarcad.com/sdk_doc/). AutoCAD 호환이라는 설명만으로 모든 인터페이스를 검증 완료로 취급하지 않습니다.
