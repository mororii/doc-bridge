# Inline AI 0.4.5 한글·엑셀 구현 분석

분석일: 2026-08-10  
분석 대상: `%LOCALAPPDATA%\Programs\inline-ai`  
목적: Inline AI의 구현을 그대로 복제하지 않고, DocBridge의 한글·엑셀 직접 제어 품질과 안정성을 높이는 데 참고할 구조·규칙·복구 정책을 식별한다.

## 0. DocBridge 0.4.8 반영 결과

2026-08-10 기준 다음 항목을 실제 코드·도구·설치 진단에 반영했다.

- 한글: 모든 실행 중 HWP 인스턴스와 탭을 정규화된 절대 경로로 탐색하고, 동일 경로가 두 번 열리면 `HWP_DUPLICATE_LOCAL_PATH`로 중단한다.
- 한글: `hwp_doctor`가 설치 경로, ProgID, 제품 버전, TypeLib GUID와 32/64비트 등록 경로를 사전 점검한다. 명시 승인 시 `hwp_repair_typelib`이 설치된 `Hwp.exe /RegServer`로 복구한다.
- 한글: COM 제어를 `doc-bridge-hwp-worker.exe` 별도 프로세스에 격리하고, 읽기 전용 호출만 최대 3회 재시도한다. timeout·프로세스 종료·poison 상태는 안정적인 오류 코드와 사용자 조치로 반환한다.
- 한글: COM 서버 생성 순간의 작업 폴더를 설치된 `Hwp.exe`의 `Bin`으로 고정해 `CultureFontManager`/`MS.Internal.FontCache.Util`의 사설 글꼴 URI 초기화 오류를 방지한다.
- 한글: `hwp_read_text(scope="document_map")`에 내용 기반 `lineId`, 문단 ordinal, coverage와 다음 읽기 위치를 제공한다. 편집은 첫 실패에서 후속 step을 멈추며, 결과에 세션 상태와 `postEditReread`를 포함한다.
- Excel: `excel_inspect`로 workbook/sheet 사용범위, 표·차트·도형·피벗·이름, 수식 오류 셀과 Protected View·모달/편집 상태를 읽기 전용 진단한다.
- 배포: worker와 도구 22개를 설치 필수조건으로 검사하고, 설치 후 `hwp_doctor`까지 통과해야 정상 설치로 판정한다.

아직 Inline AI의 전체 CVD처럼 표 중첩 경로·그림·수식·각주를 모두 안정 ID로 표현하고 임의 구조 diff를 컴파일하는 단계까지 구현한 것은 아니다. 0.4.8의 `document_map`은 본문 문단 중심의 CVD-lite이며, 기존 DocBridge의 표 전용 읽기·연산 및 주변 문맥 서식 복제와 함께 사용한다.

## 1. 결론

Inline AI에서 가장 참고할 부분은 개별 COM 명령이 아니라 그 앞뒤를 둘러싼 **문서 식별, 편집 세션, 구조화된 읽기, 실패 후 재읽기, 프로세스 격리**이다.

- 한글은 단순 `InsertText` 자동화가 아니다. 문서를 줄·문단·표 셀 단위의 편집 가능한 중간 표현(CVD)으로 읽고, 안정적인 `line_id`와 표 경로를 부여한 다음, diff를 COM primitive로 변환하고 번역 결과를 검증한다.
- 한글 편집은 `begin → step 반복 → end` 세션으로 실행된다. 첫 단계 실패 시 뒤 작업을 중단하지만 `end`는 살려 두어 저장 결정을 확정하고 현재 문서를 다시 읽는다.
- 열린 문서는 “현재 활성 창”만으로 고르지 않고 **정규화된 절대 파일 경로**로 찾는다. 같은 경로가 두 번 열려 있으면 임의로 하나를 선택하지 않고 명시적 중복 오류를 낸다.
- Excel은 `xlwings + pywin32 COM`을 중심으로 하며, 읽기는 스캔·범위·개체·검색·오류 검사로 분리한다. 쓰기는 승인 후 제한된 Python 코드 실행으로 유연성을 얻되 import와 위험 호출을 차단한다.
- 전체 시스템 프롬프트와 세부 도구 정의는 설치 폴더에 완전한 파일로 들어 있지 않다. 세션 유형을 서버에 보내면 백엔드가 시스템 프롬프트와 도구 정의를 분기한다. 로컬에는 도구 허용목록, 입력 검증, 오류별 복구 지침 같은 강제 규칙이 남아 있다.

DocBridge는 이미 Inline AI보다 강한 부분도 있다. dry-run, 전체 스냅샷, confirm token, op별 readback, 자동 롤백, 열린 Excel 재연결, 한글 주변 문맥 서식 추론은 유지해야 한다. 특히 Excel 쓰기를 곧바로 임의 Python 실행으로 바꾸는 것은 권장하지 않는다.

## 2. 설치 구조와 기술

Inline AI 0.4.5는 Electron 데스크톱 앱이며 설치 크기는 약 1.19GB이다.

- `inline-ai.exe`: 서명된 Electron 본체
- `resources/app.asar`: UI, 에이전트 오케스트레이션, 도구 검증, 오류 정책
- `resources/runtimes/agent_document_tool_executor`: 실제 Excel/HWP/Word 읽기·편집 실행기
- `resources/runtimes/document_interaction_worker`: 열린 문서 감시, 문서 식별, 미리보기·선택 영역·상태 수집
- `local_file_tool_executor`, `local_search_index_worker`: 파일 읽기와 로컬 검색

Python 실행기는 Python 3.12 기반 Nuitka 빌드다. 원본 `.py`는 배포되지 않았지만, 모듈명·함수명·오류 문자열·계약 문자열은 실행 파일에 남아 있어 구조는 충분히 확인할 수 있었다.

## 3. 한글 구현에서 확인된 핵심

### 3.1 문서 식별이 작업의 출발점이다

모든 공개 HWP 도구는 `local_path`를 필수로 받는다. 내부에는 다음 역할이 분리되어 있다.

- `document_interaction_worker.utils.doc_identity`
- 실행 중 HWP dispatch 열거
- 파일 경로 정규화와 경로별 HWP proxy 바인딩
- HWP 창·프로세스 소유자 PID 확인
- 좀비 연결 정리
- 같은 경로가 여러 창에 열린 경우 `HWP_DUPLICATE_LOCAL_PATH`
- 파일 경로가 없는 새 문서는 `UNSAVED_HWP_DOCUMENT`

이 설계의 의미는 “지금 활성화된 창일 것”이라는 추정을 없애는 것이다. 사용자가 여러 한글 창을 켜 둔 실무 환경에서는 이 차이가 매우 크다.

### 3.2 읽기 결과가 편집 좌표계다

공개 읽기 도구는 다음 네 가지다.

- `hwp_scan_document(local_path)`
- `hwp_read_pages(local_path, range)`
- `hwp_read_image(local_path, line_id)`
- `hwp_search(local_path, text/style 조건)`

페이지 읽기는 단순 텍스트 덤프가 아니다. 모듈 구조상 문단·목록·표·셀·그림·수식·각주·미주·텍스트박스를 트리로 만들고, 각 편집 가능한 줄에 `line_id`와 다음 위치 정보를 붙인다.

- 본문 문단 ordinal
- 표의 중첩 경로, 행·열·셀 위치
- 셀 fragment와 페이지 분할 상태
- 이미지·수식·각주 같은 control anchor
- 해당 줄의 문자·문단·표 서식 보조정보(style sidecar)

따라서 다음 편집은 검색 문자열을 다시 어림잡는 대신 직전 읽기 결과의 안정적인 대상 ID를 기준으로 계획할 수 있다.

### 3.3 `patch_document`는 한글용 diff 컴파일러다

공개 편집 step은 네 종류로 제한되어 있다.

- `patch_document`
- `find_replace`
- `set_style`
- `set_image`

이 가운데 핵심은 `patch_document`다. 확인된 내부 모듈은 문단 교체뿐 아니라 다음을 각각 별도 lowering 단계로 다룬다.

- 문단 추가·교체·삭제와 문단 fragment
- 단일 셀 및 다중 셀 내용 교체
- 행 삽입·행 교체·다중 행 교체
- 표 전체 교체와 stream rewrite
- 병합 셀, 중첩표, parent cell
- 열너비·표 geometry·셀 배경·테두리·서식
- 페이지에 걸친 표 fragment와 반복 머리행
- 그림·수식·각주·미주·텍스트박스

컴파일 결과가 요청한 diff 의도를 실제로 구현하는지 `translation_validation`으로 검증하고, 불일치하면 `HWP_PATCH_TV_MISMATCH`로 중단한다. 이것이 표 구조를 대충 추측해 `TableRightCell`을 반복하는 방식과 가장 큰 차이다.

### 3.4 편집 세션은 원자적 작업 단위다

사용자에게는 `hwp_edit_by_steps` 하나로 보이지만 내부에서는 다음처럼 실행된다.

1. `hwp_edit_begin(local_path, auto_save)`
2. `hwp_edit_by_step(step)` 반복
3. `hwp_edit_end(save)`

주요 규칙:

- 쓰기는 직렬 큐에서만 실행한다.
- 이전 세션이 완전히 끝나야 다음 세션을 시작한다.
- 열린 구조 patch가 여러 스트림 조각으로 나뉘면 완성될 때까지 합친다.
- 첫 step 실패 시 후속 step을 큐에서 제거하고 LLM의 나머지 step 생성을 중단한다.
- 그러나 `edit_end`는 취소하지 않는다. 저장 여부와 실패 시점, 편집 후 재읽기 결과를 다음 턴에 전달해야 하기 때문이다.
- 자동 저장은 세션 시작 시점과 종료 시점에 상태가 기록된다.
- 취소 중에도 실제 저장이 이미 끝났다면 결과를 단순 “중단됨”으로 숨기지 않는다.

이 방식은 긴 문서 작업에서 “중간에 실패했는데 AI가 끝난 줄 아는 문제”를 줄이는 데 직접 효과가 있다.

### 3.5 편집 후 재읽기가 기본이다

`hwp_edit_end` 결과에는 `<post_edit_reread>`가 포함될 수 있다. 실패한 오래된 anchor를 그대로 재사용하지 말고, 갱신된 CVD에서 새 diff를 만들도록 오류 지침도 고정되어 있다.

DocBridge의 현재 readback은 op별 문자열 존재 여부와 일부 서식 비교에는 강하지만, 실패 후 **편집된 주변 구조를 다시 모델에 공급하는 계약**은 아직 약하다. 이 부분은 우선 반영 가치가 높다.

### 3.6 서식은 단일 폰트명이 아니라 문서 구조의 일부다

확인된 구성:

- `style_sidecar`
- `paragraph_text_format`
- `style_apply`
- `cell_fill_style`
- `font_family_aliases`
- `fonts.json`

`fonts.json`은 한글/라틴/한자/일본어/기타/기호/사용자 7개 script slot별 글꼴 이름과 유형을 보유한다. 예전 한글 글꼴 별칭을 실제 설치 글꼴로 정규화하는 용도다.

DocBridge는 현재 명시적 `fontName` 적용 시 7개 FaceName 슬롯을 모두 쓰고, 빈 셀 입력 시 주변 표·이전 반복 양식까지 비교한다. 이 방향은 맞다. 추가로 필요한 것은 **글꼴 별칭·대체 규칙과 서식 sidecar를 읽기 결과에 안정적으로 노출하는 것**이다.

### 3.7 HWP 환경 복구가 별도 기능이다

Inline AI는 HWP 2018 이상을 요구하고, 설치 버전과 TypeLib 레지스트리 경로를 비교한다.

- TypeLib GUID와 `HwpObject.tlb`의 등록 경로를 확인한다.
- 설치된 기본 한글 버전과 등록된 TypeLib 버전이 다르면 `VERSION_MISMATCH`로 분리한다.
- 사용자가 승인하면 한글 실행 파일의 `/RegServer`를 관리자 권한으로 실행한다.
- 등록 완료 후 앱 재시작이 필요하다는 상태를 별도로 반환한다.
- COM 미초기화, 열린 문서 없음, 문서 경로 미발견을 서로 다른 오류로 구분한다.

동봉된 `FilePathCheckerModule.dll`은 DocBridge의 `FilePathCheckerModuleExample.dll`과 SHA-256이 완전히 같다.

`9AC5B97C47AC8AED1E8BCA27A3EEF39411361D8F68C262509F0C40A8F9D21BB6`

즉, 보안 모듈 자체에는 Inline AI만의 특별한 코드가 없다. 참고할 것은 DLL이 아니라 **등록 경로 재검증, TypeLib 선검사, 버전 불일치 복구 흐름**이다.

## 4. Excel 구현에서 확인된 핵심

### 4.1 읽기와 쓰기를 분리한다

읽기는 병렬 실행 가능하며 다음처럼 목적별로 쪼개져 있다.

- 문서·시트·개체 요약 스캔
- 지정 범위 읽기
- 차트·표·피벗 같은 개체 읽기
- 텍스트, 표시 형식, 셀 색으로 검색
- 수식·참조·표시 오류 검사

쓰기는 `excel_run_python_code` 하나지만 승인과 직렬화를 강제한다.

### 4.2 Python 쓰기는 임의 셸이 아니다

실제 편집기는 `xlwings`와 pywin32 COM을 사용한다. 모듈이 차트, 수식·데이터, 필터, 서식, 이미지, 피벗, Power Query, 도형, 시트, 슬라이서, 표 편집기로 분리되어 있다.

전달된 Python 코드는 다음 제약을 받는다.

- import 문 금지
- 금지 함수 호출 차단
- 미리 로드된 편집 helper만 사용
- 활성 시트 추정 금지, sheet-qualified range 요구
- 대상 workbook을 `local_path`로 고정
- Protected View와 모달·수식 편집 모드를 별도 오류로 분류

이 구조는 넓은 Excel 기능을 빠르게 제공하기에는 좋지만, DocBridge의 현재 typed operation보다 감사·재현·롤백이 어렵다. 따라서 DocBridge에는 임의 Python 실행을 기본 경로로 추가하기보다, 필요한 helper를 명시적 op로 계속 확장하는 편이 안전하다.

### 4.3 DocBridge에 가져올 Excel 요소

- `scan_document`, `read_object`, `error_check` 수준의 읽기 도구 분리
- Protected View와 일반 COM 오류를 분리
- 모달/수식 편집 모드 감지 후 같은 호출을 맹목적으로 반복하지 않는 정책
- 여러 Excel 인스턴스에서 절대 경로로 workbook을 찾고 중복을 오류 처리
- 타임아웃 뒤 대기 큐 전체를 같은 원인으로 종료해 연쇄 hang 방지

DocBridge는 이미 Excel 재시작 뒤 끊어진 COM 참조를 다시 잡고, 여러 인스턴스에서 열린 workbook을 경로로 검색하며, 중복 시 임의 선택을 거부한다. 이 부분은 Inline AI와 동등하거나 더 명시적이다.

## 5. 코드 외 작업 지침의 위치

완전한 시스템 프롬프트나 “한글 작업 지침.md” 같은 파일은 설치본에 없다.

확인된 흐름:

- 앱이 `sessionProfile`을 백엔드에 전달한다.
- 백엔드가 일반 세션, Excel 편집, HWP 편집 등에 맞춰 시스템 프롬프트와 도구 정의를 분기한다.
- 데스크톱 앱은 같은 profile로 도구 허용목록을 한 번 더 검사한다.
- 로컬 오류 카탈로그가 오류별 “재시도 금지/사용자 조치/다른 도구 사용” 지침을 모델에 돌려준다.
- 긴 읽기 결과가 잘리면 읽은 페이지·시트·줄 범위를 명시하고, 발췌를 전체 문서처럼 취급하지 말라는 경고를 삽입한다.

따라서 참고할 “지침”은 한 개의 프롬프트가 아니라 다음 세 층이다.

1. 서버의 작업 계획·도구 선택 프롬프트
2. 로컬 도구 allowlist와 입력 schema
3. 실행기가 반환하는 오류·coverage·재읽기 계약

DocBridge는 2와 3을 MCP 도구 설명·스킬·오류 코드에 더 강하게 넣는 것이 적합하다.

## 6. 안정성·보안·개인정보 관찰

- Inline AI 본체와 문서 실행기는 Elements의 유효한 코드 서명이 있다.
- HWP FilePathChecker 예제 DLL은 서명되지 않았고 DocBridge가 가진 파일과 동일하다.
- 문서 COM 실행기는 Electron 메인 프로세스가 아니라 별도 프로세스다. 시작 ping, 요청 timeout, 최대 3회 재시작, 지수 backoff가 있다.
- COM timeout이 발생하면 남은 문서 작업 큐를 drain한다. 응답이 없는 STA 스레드에 계속 작업을 쌓지 않는다.
- `.env`에는 백엔드 API, Sentry, Mixpanel, 일반 로그 업로드 관련 설정 키가 있다. 값은 분석·기록하지 않았다.
- 이 앱은 완전 오프라인 구조가 아니다. 문서 COM 조작은 로컬 worker에서 일어나지만, 읽기 결과·편집 지시·도구 결과는 에이전트 컨텍스트로 백엔드 모델 호출에 사용된다. 이것이 문서 전체를 항상 업로드한다는 뜻은 아니지만, 민감 문서에서는 조직 정책 확인이 필요하다.
- 앱 라이선스는 Proprietary다. 복원한 소스·문자열을 DocBridge에 복사하지 않고, 공개 COM API와 독자 구현으로 구조적 아이디어만 반영해야 한다.

## 7. 현재 DocBridge와의 차이

| 항목 | DocBridge 현재 | Inline AI에서 참고할 점 | 우선순위 |
|---|---|---|---|
| 안전성 | dry-run, snapshot, confirm token, readback, rollback | 유지. Inline AI보다 명확한 장점 | 유지 |
| HWP 대상 선택 | 첫 번째 보이는 ROT 창에 연결 후 그 인스턴스 안에서 파일 확인 | 모든 보이는 HWP 인스턴스를 경로로 열거, 중복 거부 | 최우선 |
| COM hang | 같은 프로세스의 단일 STA thread에 120초 timeout | 별도 worker process, timeout 후 worker 재생성·queue drain | 최우선 |
| HWP 편집 단위 | 한 batch 안에서 op 반복 | begin/step/end, 첫 실패 뒤 후속 중단, 종료 저장 결정 | 높음 |
| HWP readback | 문자열·일부 표/서식·구조 검사 | 변경 주변을 다시 읽어 다음 수정의 기준으로 반환 | 높음 |
| HWP 좌표 | 문자열 occurrence, 표 index/cell index | 페이지별 line_id, 문단 ordinal, 중첩 table path | 높음 |
| 표 편집 | 직접 cell/row/column/merge op | 구조 diff compiler와 table geometry validation | 중장기 |
| 서식 | 주변 문단·반복표·근접 셀 추론, 7개 글꼴 slot | font alias map, style sidecar, 페이지 구조와 결합 | 중간 |
| TypeLib | ProgID 및 보안 모듈 중심 | 설치 버전/TypeLib 경로 비교, 승인 복구, 재시작 상태 | 높음 |
| Excel | typed op, 재연결, 경로 기반 workbook 검색 | object/error scan, Protected View·modal 상태 분리 | 중간 |
| 오류 계약 | 일반 예외 문자열이 여전히 많음 | 안정적인 오류 코드와 오류별 no-retry 지침 | 높음 |

## 8. 권장 적용 순서

### 1단계: 연결과 장애 격리

1. HWP도 Excel처럼 모든 ROT 인스턴스를 열거해 `documentRef` 절대 경로로 대상 선택
2. 같은 경로가 여러 번 열려 있으면 `HWP_DUPLICATE_DOCUMENT_REF` 반환
3. 새 문서는 `documentId`로만 다루고 파일 경로 작업과 명확히 분리
4. HWP/Excel/CAD COM adapter를 별도 worker process로 이동하거나, 최소한 HWP부터 격리
5. timeout 시 해당 worker 폐기·재시작, 같은 worker의 대기 큐 drain
6. TypeLib doctor: 설치 버전, ProgID, TypeLib GUID·경로, 보안 모듈 등록을 별도 진단

### 2단계: 구조화된 읽기와 재검증

1. `hwp_scan_document`
2. `hwp_read_pages(start,end)`
3. 문단·표 셀에 안정적 `lineId`와 `tablePath` 부여
4. 읽은 범위·전체 페이지 수·잘림 여부를 항상 반환
5. 모든 쓰기 결과에 `postEditReadback`을 포함
6. 실패한 op 주변 문맥과 새 target ID를 반환

### 3단계: 세션형 한글 편집

1. `hwp_edit_begin(documentRef, autoSave)`
2. `hwp_edit_step(sessionId, op)`
3. `hwp_edit_end(sessionId, save)`
4. 첫 실패 시 남은 step 취소, end는 반드시 실행
5. 시작·끝 document identity와 내용 fingerprint 비교
6. 저장/취소/미확정 상태를 명확한 enum으로 반환

### 4단계: 한글 구조 diff

전체 CVD compiler를 한 번에 복제하지 않는다.

1. 일반 문단 replace/insert/delete
2. 단순 표 셀 replace와 행 insert/replace
3. 병합 셀·열너비·셀 배경
4. 중첩표와 페이지 분할 표
5. 그림·수식·각주·텍스트박스

각 단계는 “요청 diff → 계획된 primitive → 실행 후 구조 재읽기” 검증기를 먼저 만든 뒤 확대한다.

### 5단계: Excel 보강

1. scan/object/error-check 읽기 API
2. Protected View·모달·수식 편집 상태 코드
3. 현재 typed op의 차트·피벗·표·필터 기능 확대
4. 필요한 경우에만 관리자용 제한 코드 실행을 별도 opt-in 기능으로 검토

## 9. 필수 회귀 테스트

- HWP 창 3개에서 서로 다른 문서가 열린 상태에서 정확한 경로만 편집
- 같은 HWP 파일을 두 창에 중복으로 열고 임의 편집이 차단되는지
- 저장되지 않은 새 문서와 저장된 문서가 함께 있을 때 식별 안정성
- HWP TypeLib 미등록·구버전 등록·설치 버전 불일치
- HWP 모달에서 timeout 후 다음 요청이 영구 정지하지 않는지
- 10개 step 중 4번째 실패 시 5~10번째가 실행되지 않고 end/readback은 실행되는지
- 빈 표 셀에 주변 역할·반복 양식 서식이 복사되는지
- 병합 셀, 중첩표, 여러 페이지 표에서 행 삽입 후 구조 보존
- 한글 파일 재열기 후 텍스트·서식·표 구조 유지
- Excel Protected View, 수식 편집 모드, 모달, 재시작 뒤 COM 재연결
- Excel 여러 인스턴스와 중복 workbook 이름에서 절대 경로 식별

## 10. 근거 파일

- Electron 에이전트 및 도구 계약: `tmp/inline-ai-analysis/asar-extracted/main/agent.js`
- 설치 환경·TypeLib 복구: `tmp/inline-ai-analysis/asar-extracted/main/index.js`
- Nuitka 모듈·함수·오류 문자열 분석: `tmp/inline-ai-analysis/agent-document-strings.json`
- 열린 문서 감시 worker 분석: `tmp/inline-ai-analysis/document-interaction-strings.json`
- 현재 DocBridge HWP 구현: `doc-bridge/src/DocBridge.Core/Adapters/HwpAdapter.cs`
- 현재 DocBridge 문맥 서식 구현: `doc-bridge/src/DocBridge.Core/Adapters/HwpAdapter.StyleContext.cs`
- 현재 DocBridge Excel 구현: `doc-bridge/src/DocBridge.Core/Adapters/ExcelAdapter.cs`
- 현재 COM STA/timeout 구현: `doc-bridge/src/DocBridge.Core/Services/StaThreadRunner.cs`, `ComAdapterBase.cs`

분석 중 Inline AI 설치 폴더는 수정하지 않았으며, 비밀값·토큰·DSN·URL 값은 수집하지 않았다.
