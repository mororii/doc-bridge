# DocBridge 성능 설계와 측정

DocBridge는 안전 장치(dry-run, 전체 스냅샷, 확인 토큰, readback, 자동 롤백)를 제거하지 않고 중복 분석과 큰 응답을 줄인다.

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

## 운영 지침

- 관련된 저위험 변경은 하나의 논리적 batch로 묶어 snapshot과 재읽기 횟수를 줄인다.
- 10개를 넘는 한글 op 또는 큰 표·그림·PDF 작업은 비동기 job으로 제출하고 완료 상태를 조회한다.
- 서로 무관한 변경과 고위험 변경은 별도 batch로 유지한다.
- HWP 표·필드가 필요하지 않으면 bundle sections에 포함하지 않는다.
- `includePageCount`는 최종 쪽수 확인 때만 사용한다. 한글의 전체 pagination을 유발하기 때문이다.
- CAD 다중 도곽은 `scope:"regions"` 한 번으로 검증한다.
- 긴 결과는 잘림 표시를 확인하고 다음 범위를 이어 읽는다.
- `HWP_COM_TIMEOUT`·`HWP_CIRCUIT_OPEN`은 즉시 반복 호출하지 않는다. `retryAfterMs` 뒤 문서를 다시 읽고 새 dry-run을 만든다.

## 2026-08-10 기준 실측 예

기존 감사 로그에서 실제 HWP apply는 서식 4개 243ms, 표 행 높이 19개 2.84초, 텍스트와 표 생성 8개 7.25초였다. 사용자 체감 지연의 큰 부분은 COM 자체보다 도구 왕복·중복 preview·재읽기·재계획이었다. 새 `timings`로 문서별 p50/p95를 수집한 뒤 다음 최적화 대상을 결정한다.
