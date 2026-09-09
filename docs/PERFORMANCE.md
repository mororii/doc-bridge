# DocBridge 성능 설계와 측정

DocBridge는 안전 장치(미리보기 스냅샷, 확인 토큰 또는 execute 저널, readback, 자동 롤백)를 제거하지 않고 중복 분석과 큰 응답을 줄인다. 자동 롤백은 보호 시트 등 COM이 쓰기를 거절하면 `verified=false`가 될 수 있다.

일반 편집(`autoExecuteOps`)은 `executionMode=execute`로 한 잠금 안에서 사전검사·snapshot·apply·readback·실패 시 복구를 한 번에 수행한다. 이 경로는 dry-run 토큰과 fingerprint 재사용이 없다. 미리보기·고위험은 기존 dry-run → 같은 ops/`confirmToken`이다.

Excel format-only v3는 쓰는 속성만 스냅샷한다. Bold-only는 다른 색·tint COM을 읽거나 복구하지 않는다. 균일 bool/글꼴 크기/`NumberFormat`은 범위 빠른 경로를 쓰고, 혼합·색은 필요한 셀 상태를 확보한다. 셀 한도 100,000과 STA 120초는 올리지 않았다. 다른 시트의 부풀린 UsedRange를 대상만 스냅샷하는 것은 가용성 회복이며 일반 속도 개선이 아니다.

아래 「확인된 실측」은 현재 소스의 단일 표본이다. 별도 전수 native 검사는 제품 호출 시간에 넣지 않는다. p50/p95·만능 배율·전체 소스 릴리즈 완료를 주장하지 않는다. 배포 ZIP 생성은 이 문서 밖이다.

## 적용된 빠른 경로

1. 모든 `*_apply_ops` 응답은 `timings`를 반환한다.
   - `validationMs`, `lockWaitMs`, `statusMs`
   - `previewMs`, `snapshotMs`, `tokenMs`
   - apply 시 `tokenValidationMs`, `snapshotLookupMs`, `documentIdentityMs`
   - HWP 빠른 경로의 `fingerprintValidationMs`, `previewReused`, `fingerprintMethod`
   - `applyMs`, 실패한 경우 `rollbackMs`, 전체 `totalMs`
2. dry-run preview는 스냅샷 `metadata.json`에 보존한다. HWP와 테스트 adapter는 전체 fingerprint가 일치할 때만 실제 적용에서 이를 재사용한다.
3. HWP preview는 한 batch에서 `GetTextFile("TEXT")`를 한 번만 읽고 표 control 개수를 캐시한다.
4. `scope:"bundle"`은 요청한 HWP 읽기 section을 한 COM 연결에서 처리한다. 기본 section은 `text`, `document_map`, `structure`다.
5. `postEditReread`는 한 번 읽은 본문으로 hash·미리보기·문단 지도를 모두 만든다.
6. CAD 활성 컨텍스트는 레이어 50개만 미리 보여 준다. 전체 레이어는 `cad_query_entities({"scope":"layers"})`로 명시적으로 읽는다.
7. 새 HWP 프로세스의 status 조회도 실행 중인 ROT 창에 먼저 연결한다. 따라서 CLI처럼 dry-run과 apply가 별도 프로세스여도 저장되지 않은 문서의 `untitled-*` 식별자가 비지 않는다.
8. HWP worker 제한시간은 공통 135초가 아니라 작업별이다. 상태 15초, 컨텍스트 20초, 일반 읽기 30초, 복합 읽기·preview 45~60초, apply·DOCX/PDF·복원 45~90초 범위로 제한한다.
9. `format_paragraphs`는 같은 대상의 글자·문단 서식을 한 번의 찾기 순회에서 적용한다. `table_set_row_heights`는 한 표 control을 재사용해 여러 행의 높이·readback을 한 op로 처리한다.
10. COM timeout이나 worker 실패 뒤에는 45초/15초 회로를 열어 새 worker·빈 한글 창 반복 생성을 막는다. 응답의 `automaticRetry:false`, `retryPolicy.mode:"after-delay"`, `retryAfterMs`를 따르며 자동 롤백용 `restoreSnapshot`, 진단, 명시적 복구만 보호 시간에도 허용한다.
11. 한글 표의 여러 셀은 최대 500개를 `table_set_cells` 한 op로 묶는다. 표 컨트롤과 수식 위치를 한 번만 읽고 각 셀을 정확히 재검증한다.
12. 클라이언트 제한시간을 넘길 수 있는 한글 쓰기는 `hwp_submit_ops`로 한 번만 제출하고 `hwp_get_job`을 조회한다. timeout은 제출 실패를 뜻하지 않으므로 같은 payload를 재제출하지 않는다.

## HWP fingerprint

비결정적인 네이티브 HWP 직렬화값을 preview 재사용 판단에 사용하지 않는다. 다음 결정적 상태를 SHA-256으로 묶는다.

- 문서 ID와 경로
- 정규화 본문 전체
- 선택 영역과 커서 위치
- control ID 순서
- 필드 목록
- 현재 커서의 문자·문단 서식

fingerprint가 바뀌면 apply를 중단하고 새 dry-run을 요구한다. 이 거부 단계에서는 기존 확인 토큰을 소비하지 않는다. 네이티브 HWP 전체 백업과 복원 검증은 기존 방식 그대로 유지한다.

## 확인된 실측

같은 검사 도구가 대상 셀을 적용 전·적용 후·복구 후 COM으로 읽어 Bold, fill, Pattern을 검사했다. 200셀 단위이며 unreadable과 mismatch는 모두 0이다. 별도 전수 검사 시간은 제품 호출 시간에 넣지 않는다. 기존 Excel 인스턴스는 보존했고 각 시험에서 소유한 새 Excel만 정상 종료했다.

| 대상/경로 | dry-run ms | apply/execute ms | restore ms | 제품 합계 ms | snapshot bytes | 별도 전수 검사 ms |
|---|---:|---:|---:|---:|---:|---:|
| 이전 안정화 후보, 1000셀 legacy | 22955 | 24353 | 72391 | 119699 | 579493 | 27130 |
| 개선 후보, 1000셀 legacy | 186 | 422 | 151 | 759 | 21485 | 25729 |
| 개선 후보, 1000셀 execute | 해당 없음 | 522 | 581 | 1103 | 20284 | 23085 |
| 개선 후보, 7000셀 execute | 해당 없음 | 308 | 151 | 459 | 57663 | 94255 |

각 행은 단일 표본이며 동일 워크스테이션의 세션/캐시/부하 차이가 있다. 단일 표본으로 일반 속도 배수, p50/p95, 7000셀이 1000셀보다 빠르다는 결론을 내리지 않는다. 1000셀 비교의 전후/복구 digest는 기준선과 후보가 일치한다. 위 표의 7000셀 0.459초는 균일 Bold execute다. 혼합 색의 성능을 대표하지 않는다. execute는 도구 호출 왕복을 줄이는 계약이지만 이 표에서 legacy보다 항상 빠르다는 주장은 할 수 없다.

### Excel 혼합 서식·복구·UUID·작은 회귀

1,000셀 혼합 색의 제품 호출 합계는 legacy 20.903초, execute 11.182초다. 별도 전수 검사는 18.206초/10.460초다. 실제 COM 장애(op0 성공/op1 COM 실패/op2 미실행) 뒤 원복, UUID 재전송 challenge/conflict 실제 호출은 통과했다. 이 행은 dry-run/restore 분해와 snapshot bytes를 공개하지 않는다.

7,000셀 혼합 execute는 대상 7,000셀 전수 unreadable/mismatch 0이다. execute 전체 52.154초(그중 snapshot 51.864초, 실제 apply 0.169초), restore 30.655초, 제품 합계 82.809초, snapshot 62,579 bytes다. 별도 전수 native 검사는 81.523초이며 제품 합계에 넣지 않는다. 큰 색 스냅샷 비용은 미해결 성능 한계다.

균일 7,000셀 execute 0.459초와 혼합 7,000셀 합 82.809초는 다른 작업이다. 만능 속도 배수를 만들지 않는다. 균일 서식의 짧은 시간은 요청 속성만 읽고 복구하는 범위 스냅샷 때문이다. 복잡한 혼합 색은 셀별 COM 읽기·검증 비용이 남는다.

작은 회귀 9건은 실제 Excel에서 통과했다. 색 별칭, JSON 순서 canonical hash, 보호시트 실패 진단, 서식 지문 변경 차단, theme/no-fill 원복, RGB+tint 색 변경 차단, 다른 활성 workbook 보존, legacy와 execute 각각의 Bold false→true와 RGB/tint 보존이다. 보호시트 건은 롤백 쓰기도 거절되므로 `rollback.verified=false`를 숨기지 않은 성공이며 원복 성공이 아니다.

별도 저장된 OOXML 검사에서 균일 1,000/7,000은 백업 파일과 불일치 0이다. 혼합 1,000/7,000은 메모리 초기 seed와 저장된 OOXML을 독립 비교해 불일치 0이다. 아직 저장되지 않은 혼합 seed를 디스크 workbook-backup과 비교한 초기 결과는 기준이 잘못되어 전부 다르게 나오며 복구 실패의 근거가 아니다. native 전후/복구 digest도 일치한다.

### 한글

일반 execute/UUID 재전송 1건은 31.269초 통과했다. 중복 삽입 1회와 대상 `documentRef`를 확인했다. 한글 프로세스 전체 종료 시험은 아니다.

제작 format/table/page/picture/PDF 1건은 10.833초 통과했다. PDF 1쪽 596×840pt. 렌더에서 한글, 파랑/빨강 글자서식, 표, 검은 시험 그림, 바닥글/쪽번호를 확인했다. 시험 이미지가 본문 중간에 있어 줄이 나뉘며 완성 문서 디자인 검수로 주장하지 않는다.

### CAD/GstarCAD

기존 geometry production fixture 2건은 GstarCAD 25.382초, AutoCAD 4.204초다. 선/원/폴리선, 한글·Unicode TEXT/MTEXT, ACI, 이동/축척/복사/대칭/offset, 문자 변경, 레이어, regen, 깨끗하게 저장된 소유 도면 SaveAs/재조회를 확인했다. 기존 사용자 도면 경로·저장 플래그·객체 수는 보존했다. 해치·문서 간 복사·XREF·배치·PDF·스크립트·RGB는 미지원이며 이번 실측의 검증 범위가 아니다.

execute 최종 fixture 2행은 GstarCAD 23.390초, AutoCAD 4.281초로 통과했다. 첫 문자 변경, 주입 Apply 장애 전 native 문자 확인, native 롤백, UUID 재전송 challenge, 주변 객체 보호, 문자·레이어·regen, 도형 execute 거절과 legacy 도형 성공, dirty legacy SaveAs 거절(출력 없음), 깨끗하게 저장된 소유 도면의 고위험 SaveAs 성공, 기존 사용자 도면 경로·저장 플래그·객체 수 보존을 확인했다. 도형은 execute가 아니며 기존 토큰 경로다. dirty 광범위 작업이나 저장 지문이 맞지 않으면 거절한다. 경로를 만들려고 자동 저장하거나 지문 거절을 우회하지 않는다. 이 시간은 geometry production 25.382초/4.204초와 다른 표본이며 합산하지 않는다. 미지원 Gstar 기능은 검증 범위가 아니다.

### status·스냅샷 조회

`core_get_status` 동일 앱 상태의 연속 3표본(서로 다른 프로세스의 cold/warm)은 이전 전체 1777.592/67.649/71.444ms, 새 전체 2166.565/90.917/68.133ms, 새 `app=excel` 302.116/16.967/17.555ms다. app 필터는 불필요한 앱 조회를 줄였으나 전체 status 자체의 개선은 입증되지 않았다.

snapshot synthetic history exact lookup은 preview-only이며 쓰기가 없다. 3표본은 1,000건 31.316/102.260/1.385ms, 5,000건 45.431/99.765/1.397ms다. 이전 5,000건은 1113.104/1673.622/1115.652ms다. 캐시/디스크 변동이 있어 3표본을 그대로 두고 p50/p95/보편 배율은 없다.

### 빌드와 생명주기

통합 Release 빌드(`-warnaserror`)는 경고 0, 오류 0이다. Core 일반 테스트 421/421, MCP 20/20이며 E2E 클래스는 필터로 제외했다.

배포 MCP 읽기 후 stdin EOF에서 worker는 0.360초 종료했다. root가 만든 MCP parent만 강제 종료하면 worker는 0.473초 종료했다. 기존 Excel PID는 보존했다. 이 결과는 소유 Excel 편집 중 parent crash 정리가 해결되었음을 뜻하지 않는다.

이전 Excel owner 강제 종료 시험에서 watchdog Quit 이후에도 Excel 프로세스가 남았다. 원인/해결은 미확정이다. 정상 disconnect/worker 종료 성공을 이 문제의 해결 증거로 사용하지 않는다.

## 운영 지침

- 관련된 저위험 변경은 하나의 논리적 batch로 묶어 snapshot과 재읽기 횟수를 줄인다.
- 10개를 넘는 한글 op 또는 큰 표·그림·PDF 작업은 비동기 job으로 제출하고 완료 상태를 조회한다.
- 서로 무관한 변경과 고위험 변경은 별도 batch로 유지한다.
- HWP 표·필드가 필요하지 않으면 bundle sections에 포함하지 않는다.
- `includePageCount`는 최종 쪽수 확인 때만 사용한다. 한글의 전체 pagination을 유발하기 때문이다.
- CAD 다중 도곽은 `scope:"regions"` 한 번으로 검증한다.
- 긴 결과는 잘림 표시를 확인하고 다음 범위를 이어 읽는다.
- `HWP_COM_TIMEOUT`·`HWP_CIRCUIT_OPEN`은 즉시 반복 호출하지 않는다. `retryAfterMs` 뒤 문서를 다시 읽고 새 execute 또는 dry-run을 만든다.
- execute `outcomeUnknown`이면 새 UUID로 재실행하지 말고 문서 상태를 먼저 확인한다.
- 한 앱 작업은 `core_get_status` app 필터와 필요한 범위 읽기만 하고 ping·전체 앱 조회·전체 재읽기를 반복하지 않는다.

## 2026-08-10 기준 실측 예

기존 감사 로그에서 실제 HWP apply는 서식 4개 243ms, 표 행 높이 19개 2.84초, 텍스트와 표 생성 8개 7.25초였다. 사용자 체감 지연의 큰 부분은 COM 자체보다 도구 왕복·중복 preview·재읽기·재계획이었다. 새 `timings`로 문서별 p50/p95를 수집한 뒤 다음 최적화 대상을 결정한다.
