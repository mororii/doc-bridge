# DocBridge 0.4.20 — 단일 호출 편집·서식 속도와 GstarCAD 독립 연결

2026-09-09 · Windows x64 · Excel / 한글 / AutoCAD / GstarCAD

버전은 **0.4.20**입니다. 이 문서는 공개 동작 계약과 확인된 실측을 정리합니다.
아래 「기록된 검증」은 당시 게이트이며 「확인된 실측」과 합산하지 않습니다.
확인된 범위의 숫자와 한계는 「확인된 실측」과 [PERFORMANCE.md](PERFORMANCE.md)에 있습니다.
전체 소스 릴리즈 완료나 미지원 Gstar 기능 검증을 주장하지 않습니다. 배포 ZIP과 격리 설치는
패키지 생성 단계의 범위이며, 이 문서에 파일 해시나 내부 경로를 넣지 않습니다.

## 쓰기 경로와 Excel v3 서식 스냅샷

- 일반적인 사용자 편집 요청은 해당 범위의 승인이다. 클라이언트 권한 UI, 서버 사전 검사, 사람의 범위 승인을 구분한다. 모든 쓰기마다 재승인 질문을 요구하지 않는다. `highRiskConfirm`은 권한 UI를 무력화하거나 사람 승인을 인증하지 않는다.
- 미리보기·고위험은 기존 dry-run → 같은 ops/`confirmToken` apply다.
- 선택 경로 `executionMode=execute`는 UUID `requestId`와 `expectedDocumentRef`가 필수다. `dryRun`/`confirmToken`/`highRiskConfirm`은 false라도 포함하면 거절한다. 앱별 `autoExecuteOps`만 한 호출에서 사전검사·snapshot·apply·readback·실패 시 복구를 수행한다.
- Excel execute: `set_values`, `set_formulas`, `format_range`. CAD/GstarCAD execute: `set_text_value`, `set_layer_visibility`, `set_layer_color`, `regen_document`. 한글 allowlist는 `default.policy`와 `core_get_capabilities`를 따른다. `insert_text`는 execute 대상이 아니다.
- Excel/CAD/GstarCAD execute의 `expectedDocumentRef`는 현재 저장된 절대 경로만 안내한다. 어댑터는 인스턴스 바인딩 참조를 발급하지 않는다. `Book1`/`Drawing1`은 기존 토큰 경로이며 경로를 만들려고 자동 저장하지 않는다. 한글은 반환된 `documentRef`/`instanceRef`를 그대로 사용한다.
- `requestId`는 호출마다 새 UUID다. 같은 작업의 통신 재시도만 같은 UUID를 쓴다. 같은 UUID/앱/문서/ops는 저장된 결과를 재전송하고 다른 payload는 거절한다. `outcomeUnknown`이면 새 UUID로 재실행하지 말고 문서를 먼저 확인한다.
- Excel v3 format-only는 `styleScope=written-properties`다. Bold-only는 다른 색·tint를 건드리지 않는다. 균일 bool/글꼴 크기/`NumberFormat`은 범위 빠른 경로, 혼합·색은 셀 상태를 확보한 뒤 복구 후 확인한다. 기존 v2 스냅샷도 복구한다. 오래된 format preview 토큰은 새 dry-run이 필요할 수 있다. `MaxFormatSnapshotCells` 100,000과 STA 120초는 올리지 않았다.
- CAD/GstarCAD 문자 핸들·레이어 속성 스냅샷은 그 범위만 완전 복구한다. 전체 도면 복구가 아니다. Geometry 스냅샷은 drawing-backup과 부분 상태이며 자동 복구 성공을 보장하지 않는다. dirty 광범위 작업이나 저장 지문이 맞지 않으면 거절한다. 경로를 만들려고 자동 저장하거나 지문 거절을 우회하지 않는다. 깨끗하게 저장된 소유 도면의 명시적 SaveAs만 기존 고위험 토큰 경로다.

## 2026-09-09 Excel·확인 토큰

- 확인 토큰의 JSON 정규화는 객체 키를 재귀 정렬합니다. 배열 순서, 누락과 null, 문자열/Unicode는
  그대로 둡니다. 숫자는 어휘 계수·지수로 맞춥니다. `1`/`1.0`/`1e0`은 같고 `1e-100`은 `0`과 다르며,
  긴 소수와 거대 지수는 반올림하거나 0으로 접히지 않습니다. 관리 객체의 NaN/Infinity는 거부합니다.
  파서가 유한 JSON(`1e100` 등)을 float/double Infinity로 넘치게 변환한 경우는 거부가 아닙니다.
  RFC 8785 준수를 주장하지 않습니다.
- Excel `format_range`는 `bold`/`fontBold`, `italic`/`fontItalic`,
  `fillColor`/`interiorColor`/`fill` 별칭을 받고, 충돌·미지원 키·null·잘못된 색은
  snapshot/토큰 전에 거부합니다.
- dry-run 서식 `before`는 실제 셀 요약입니다. 읽기 `includeStyles`는 쓰기 키와 읽기 별칭,
  `fillPattern`을 함께 줍니다. 채움 없음은 `Pattern=none`이며 `Color=0`과 구분합니다.
- 적용 실패는 op별 단계·시간·예외·HRESULT를 남기고, 이후 op는 `skipped`입니다. mismatch는
  빈 `errors`로 숨기지 않습니다. COM 데이터 검증과 화면 육안 품질은 별개입니다.
- `format_range`만 있는 배치는 대상 서식만 스냅샷·지문합니다. 부분 병합, 혼합 리치 텍스트,
  ThemeColor의 단절/사용 중 COM 실패, 그리고 테마가 아닌 RGB+nonzero tint
  (`[EXCEL_FORMAT_UNSUPPORTED_RGB_TINT]`)는 쓰기 전에 거절합니다. 테마 tint는 유지합니다.
  값·수식·구조가 섞인 Excel 배치는 전 문서 지문을 쓰지 않고 새 preview를 받습니다.
- 전역 자동화 잠금은 프로세스 간 직렬화일 뿐 같은 창의 동시 편집을 허용하지 않습니다.
- Excel worker `validatePreviewReuse`는 apply와 같은 150초입니다. status/context/read discovery는
  45초를 유지합니다. STA COM 제한은 120초입니다. 큰 대상의 병목은 서식 쓰기가 아니라
  snapshot/fingerprint/restore이며, 정상 대량 서식은 이전 경로보다 느릴 수 있습니다.

## 2026-09-03 GstarCAD 독립 연결·기본 편집

- 기존 AutoCAD `cad_*` 도구는 그대로 유지하고 `gstarcad_launch`, `gstarcad_get_active_context`, `gstarcad_query_entities`, `gstarcad_apply_ops`를 추가했습니다. 공통 MCP 공개 도구는 29개입니다.
- `Gcad.Application`에 기존 인스턴스 우선으로 연결합니다. 상태/읽기 호출은 프로그램을 생성하지 않습니다. 명시적 launch만 해당 제품을 실행할 수 있으며 AutoCAD와 상호 대체하지 않습니다.
- 제품별 어댑터·승인 토큰·스냅샷을 분리했습니다. 조회 결과와 후속 조회 도구도 해당 제품을 유지합니다. DXF 파일 읽기도 제품 표기를 보존합니다.
- 죽은 CAD COM 참조는 연결됨으로 표시하지 않고 해제 후 다음 연결에서 다시 획득합니다. 사용자 CAD에 Quit/강제 종료를 호출하지 않습니다.
- GstarCAD 기본 작성/수정은 기존 ActiveX 편집 파이프라인을 재사용하며 dry-run, HMAC 바인딩, 문서 identity, snapshot, 쓰기 readback, 전경 보존, 자동 Regen을 유지합니다. 원/호의 반지름도 geometry 조회에 포함됩니다.
- 미지원 GstarCAD 배치는 쓰기 전에 차단합니다. AutoCAD 전용 typed entity-array와 명령을 GstarCAD로 잘못 전달하지 않습니다.
- CAD/문서 자동화 스킬, Cursor 규칙, 초기 MCP 지침과 초보자 HTML 설명서에 제품 선택·미지원 범위를 반영했습니다.

## 지원 한계

GstarCAD 해치, 문서 간 복사, 블록 삽입, XREF, 배치·뷰포트 편집, PDF 출력, 스크립트, RGB 색상은 아직 지원하지 않습니다. 호·타원·점·치수·회전·일반 속성/블록 속성·삭제는 기본 API 경로가 제공되지만 이번 실앱 실행의 검증 범위에는 포함되지 않았습니다. 필요한 기능은 `core_get_capabilities(app="gstarcad")`로 확인하십시오.

Excel `format_range`는 기본 글꼴·채우기 별칭과 format-only 대상 스냅샷만 계약합니다. 혼합 리치 텍스트, 부분 병합, 테마 조회 COM 실패, 비테마 RGB+tint는 쓰기 전에 거절합니다. 행 높이·열 너비·테두리·조건부 서식은 이 버전의 지원 완료 항목이 아닙니다. 보호된 시트에서 Bold 등 COM 쓰기가 거절되면 자동 롤백도 같은 보호 때문에 쓰지 못하고 `rollback.verified`는 false입니다. 모든 적용 실패가 검증된 롤백으로 끝난다고 보지 마십시오.

균일 Bold와 혼합 서식 1,000셀·7,000셀 전수 검사는 「확인된 실측」과 [PERFORMANCE.md](PERFORMANCE.md)를 따른다. 균일 7,000셀 0.459초와 혼합 7,000셀 82.809초는 다른 작업이며 만능 속도 배수를 만들지 않는다.

다른 시트의 부풀린 UsedRange는 가용성 회복이지 일반 속도 개선이 아닙니다. 소유 Bloat 시트의 UsedRange는 양쪽 모두 131×16,384=2,146,304셀입니다. 기준선은 `Target!A1` dry-run을 Bloat 스냅샷이 1,000,000셀을 넘어 거절했고 쓰지 않았습니다. 후보는 같은 대상의 dry-run/apply/restore가 모두 통과했고, 완료 951ms이며 스냅샷은 대상만 담았습니다.

Excel 전용 worker가 저장된 소유 통합문서를 `excel_disconnect` 또는 파이프 정상 종료하면 종료 코드 0으로 회수됩니다. 호스트가 비정상 종료된 뒤에는 빈 인스턴스와 저장된 소유 통합문서 모두에서 watchdog가 연결(`bind-ok`)·부모 소실(`parent-gone`)·Quit 호출(`quit-ok`)까지 성공해도, 종료 코드 6(`process-still-running`)으로 `EXCEL.EXE`가 남을 수 있습니다. 저장되지 않은 시작 통합문서가 원인이라는 증거는 없습니다. 잔류 Excel 프로세스가 해결되었다고 보지 마십시오. 사용자가 연 Excel은 강제 종료하지 않습니다. 이는 미해결 기준선 한계입니다.

비활성 창 작업은 지원하지만 동일 대상 창의 동시 사용자 조작은 지원하지 않습니다. `readback.verified`와 Regen 성공은 전체 문서의 육안 품질 검수와 다릅니다. 기하 변경의 완전한 복구가 필요하면 스냅샷의 원본 DWG 사본을 사용해야 할 수 있습니다.

## 확인된 실측

같은 검사 도구가 대상 셀을 적용 전·적용 후·복구 후 COM으로 읽어 Bold, fill, Pattern을 검사했다. 200셀 단위이며 unreadable과 mismatch는 모두 0이다. 별도 전수 검사 시간은 제품 호출 시간에 넣지 않는다. 기존 Excel 인스턴스는 보존했고 각 시험에서 소유한 새 Excel만 정상 종료했다.

| 대상/경로 | dry-run ms | apply/execute ms | restore ms | 제품 합계 ms | snapshot bytes | 별도 전수 검사 ms |
|---|---:|---:|---:|---:|---:|---:|
| 이전 안정화 후보, 1000셀 legacy | 22955 | 24353 | 72391 | 119699 | 579493 | 27130 |
| 개선 후보, 1000셀 legacy | 186 | 422 | 151 | 759 | 21485 | 25729 |
| 개선 후보, 1000셀 execute | 해당 없음 | 522 | 581 | 1103 | 20284 | 23085 |
| 개선 후보, 7000셀 execute | 해당 없음 | 308 | 151 | 459 | 57663 | 94255 |

각 행은 단일 표본이며 동일 워크스테이션의 세션/캐시/부하 차이가 있다. 단일 표본으로 일반 속도 배수, p50/p95, 7000셀이 1000셀보다 빠르다는 결론을 내리지 않는다. 1000셀 비교의 전후/복구 digest는 기준선과 후보가 일치한다. 위 표의 7000셀 0.459초는 균일 Bold execute다. 혼합 색의 성능을 대표하지 않는다. execute는 도구 호출 왕복을 줄이는 계약이지만 이 표에서 legacy보다 항상 빠르다는 주장은 할 수 없다.

1,000셀 혼합 색의 제품 호출 합계는 legacy 20.903초, execute 11.182초다. 별도 전수 검사는 18.206초/10.460초다. 실제 COM 장애 뒤 원복, UUID 재전송 challenge/conflict는 통과했다. 7,000셀 혼합 execute는 전수 7,000셀 검증이 통과했다. execute 전체 52.154초(그중 snapshot 51.864초, 실제 apply 0.169초), restore 30.655초, 합 82.809초, snapshot 62,579 bytes다. 별도 전수 native 검사는 81.523초이며 제품 합계에 넣지 않는다. 균일 7,000셀 0.459초와 혼합 7,000셀 82.809초는 다른 작업이다. 만능 속도 배수를 만들지 않는다.

작은 회귀 9건은 실제 Excel에서 통과했다. 색 별칭, JSON 순서 canonical hash, 보호시트 실패 진단, 서식 지문 변경 차단, theme/no-fill 원복, RGB+tint 색 변경 차단, 다른 활성 workbook 보존, legacy와 execute 각각의 Bold false→true와 RGB/tint 보존이다. 보호시트 건은 `rollback.verified=false`를 숨기지 않은 성공이며 원복 성공이 아니다.

한글 일반 execute/UUID 재전송 1건은 31.269초, 제작 format/table/page/picture/PDF 1건은 10.833초다. PDF 1쪽 596×840pt이며 완성 문서 디자인 검수로 주장하지 않는다.

CAD/GstarCAD geometry production은 GstarCAD 25.382초, AutoCAD 4.204초다. execute 최종 fixture 2행은 GstarCAD 23.390초, AutoCAD 4.281초로 통과했다. 문자·레이어·regen, native 롤백, UUID 재전송, 도형 execute 거절, dirty SaveAs 거절(출력 없음), 깨끗한 소유 도면 SaveAs 성공을 확인했다. 두 시간은 다른 표본이며 합산하지 않는다. 미지원 Gstar 기능은 검증 범위가 아니다. dirty 광범위 작업/저장 지문 불일치는 거절하며 자동 저장으로 우회하지 않는다.

통합 Release 빌드(`-warnaserror`)는 경고 0, 오류 0이다. Core 일반 테스트 421/421, MCP 20/20이며 E2E 클래스는 필터로 제외했다. `core_get_status` app 필터 3표본은 [PERFORMANCE.md](PERFORMANCE.md)에 있다. 불필요한 앱 조회는 줄였으나 전체 status 속도 개선은 입증되지 않았다.

「기록된 검증」과 합산하지 않는다. 이전 Excel owner 강제 종료 시험에서 watchdog Quit 이후에도 Excel 프로세스가 남았다. 원인/해결은 미확정이다. 정상 disconnect/worker 종료 성공을 이 문제의 해결 증거로 사용하지 않는다.

## 기록된 검증

이미 기록된 결과와 현재 소스의 최종 게이트를 섞지 않습니다. 아래 시간과 건수는
당시 기록이며 현재 소스 실측이 아닙니다. 모든 항목이 첫 시도에 통과했다고 보지 마십시오.

- **안정화 수정 전 0.4.20 소스:** 임시 대상에서 Excel, AutoCAD, GstarCAD 실앱 흐름이 통과로 기록되었습니다. 그 기록은 현재 소스의 완료 증거가 아닙니다. CAD는 데이터/화면 갱신 검증이며 육안 품질 보장이 아닙니다.
- **2026-09-03 GstarCAD:** 실행 중인 GstarCAD에서 새 임시 도면만 편집·저장하고 기존 사용자 도면을 보존한 실측이 있습니다. 한 PC의 소형 임시 도면 측정이며 대형 도면 성능 보장이 아닙니다.
- **2026-09-09 Excel 생명주기 프로브:** 파이프 worker + 저장된 소유 통합문서 + 정상 종료는 통과입니다. 크래시 경로(빈 인스턴스, 저장된 소유 통합문서)는 둘 다 watchdog `bind-ok` / `parent-gone` / `quit-ok` 이후 종료 코드 6, `process-still-running`입니다. 저장되지 않은 시작 통합문서가 원인이라는 증거는 없습니다. 사용자 Excel은 강제 종료하지 않습니다.
- **2026-09-09 Excel format-only 100셀:** 별칭, JSON 재정렬, 보호 시트 COM, 과도한 글꼴 크기 COM, dry-run 후 서식 변경, 테마/채움없음 복원, 비테마 RGB+tint 거절, 다른 활성 통합문서 보존, 100셀 규모 등 9건이 진단 및 지원되는 복원 계약으로 통과했습니다. 보호 시트 건의 통과는 Bold COM 거절과 롤백 불가(`rollback.verified=false`)를 숨기지 않은 것이며, 보호 시트에서도 롤백이 성공했다는 뜻이 아닙니다.
- **2026-09-09 이전 format-only 1,000셀 기록:** 당시 단일 표본이며 「확인된 실측」의 같은 검사 도구 1,000/7,000셀과 합산하지 않는다.
- **2026-09-09 Excel format-only 부풀림(v3):** 위 지원 한계와 같습니다. 기준선은 Bloat UsedRange 2,146,304셀 때문에 `Target!A1` dry-run을 거절했습니다. 후보는 대상만 스냅샷해 dry-run/apply/restore가 통과했고 완료 951ms입니다. 가용성 회복이며 일반 속도 개선이 아닙니다.
- **2026-09-09 이전 독립 Release 게이트:** 당시 비-E2E는 Core 303과 MCP 19였다. 현재 확인 건수는 「확인된 실측」의 Core 421·MCP 20이다.
- **2026-09-09 당시 실앱 게이트(현재 소스 최종이 아님):** `Excel_full_flow` 27.404초 통과. 비생성 discovery 1.420초 통과. Excel 재연결의 첫 COM 생성은 `0x80010001`로 실패했고, 격리 재시도는 16.041초 통과했습니다. `Hwp_full_flow` 10.176초 통과, 저장 산출물 있음. 이 한글 기본 흐름은 아래 이전 제작 fixture와 다른 게이트입니다. GstarCAD 실측 34.231초 통과. AutoCAD는 처음에 실행 중이 아니었고 조회가 인스턴스를 만들지 않았습니다. 명시적 launch 뒤 격리 재시도 5.316초 통과. 비활성인 추가 CAD 이론 행은 집계하지 않았습니다. CAD는 데이터/화면 갱신 검증이며 육안 품질 보장이 아닙니다.
- **이전 한글 제작 fixture(중간 Core 단계, 최종 한글 기본과 구분):** production format/table/page/PDF 1건 114.100초 통과. 1쪽 PDF를 렌더링해 육안으로 확인했습니다. 최종 `Hwp_full_flow` 10.176초와 합산하거나 같은 게이트로 보지 마십시오.

GitHub 이슈는 이 문서에서 닫지 않습니다.

## 패키지

배포 ZIP 생성과 격리 설치·제거 증거는 패키지 생성 단계의 범위다. 이 문서에 ZIP 해시나 추측 건수를 넣지 않으며, 패키지 완료를 주장하지 않는다.

## 적용

`DocBridge-0.4.20-win-x64.zip`을 새 폴더에 풀고 기존 작업을 저장하십시오. AI 클라이언트들을 완전히 종료한 뒤 `0-VERIFY.cmd` → `1-INSTALL.cmd` → `2-TEST.cmd`를 실행하고 AI를 재시작해 새 작업에서 `core_ping`의 `0.4.20`과 `gstarcad_*` 4개 도구를 확인합니다. 실행 중인 이전 서버를 강제로 종료하거나 설치 잠금 검사를 우회하지 마세요. ZIP 생성은 현재 사용 중인 AI 세션의 자동 업데이트를 뜻하지 않습니다.

사용 예시와 공식 개발자 자료는 [GstarCAD 안내](GSTARCAD.md), [Excel operations](EXCEL-OPERATIONS.md), [HWP 작업 명세](HWP-OPERATIONS.md)를 참고하세요.
