# DocBridge 0.4.24

Windows의 Microsoft Excel, 한컴 한글(HWP/HWPX), AutoCAD 및 GstarCAD를 Kimi·Claude·Codex·Cursor가 공통 MCP 도구로 읽고 수정하게 하는 로컬 브리지입니다.

쓰기는 서버가 두 경로 중 하나로 적용합니다. 일반적인 사용자 편집 요청은 그 범위의 승인입니다. 클라이언트 신뢰/도구 권한 UI, 서버 사전 검사, 사람의 작업 범위 승인을 같은 것으로 보지 마십시오. 이미 요청한 저장·PDF를 포함해 모든 쓰기마다 재승인 질문을 요구하지 않습니다. `highRiskConfirm`은 권한 UI를 무력화하거나 사람 승인 사실을 인증하지 않습니다.

```text
미리보기·고위험: read → dry-run(diff + snapshot + confirmToken)
                 → 동일 ops + confirmToken apply → readback → audit

일반 편집(앱별 autoExecuteOps): status(app) → 필요한 범위 읽기
                 → executionMode=execute + UUID requestId + expectedDocumentRef
                 → 한 호출에서 사전검사·snapshot·apply·readback·실패 시 복구
```

처음 설치한다면 [INSTALL.md](INSTALL.md)를 먼저 보세요.

## 현재 검증 상태

0.4.22는 Excel 쓰기 op 16종과 조회 scope 4종을 추가합니다. 외부 링크 목록·새로고침·원본 변경·끊기, 계산 모드(수동/자동/준자동 + 전체/전체재작성), 값 고정, 클립보드 없는 선택 붙여넣기(값/수식/서식/전체 + 전치·빈칸건너뜀·사칙연산), 목표값 찾기, 암호 없는 통합문서 보호/해제, 창 분할, 스파크라인 생성/수정/삭제, 슬라이서 생성/삭제, 셀 스타일 적용과 `links`/`sparklines`/`slicers`/`cellStyles` 조회입니다. `set_sheet_visibility`는 `veryHidden`을 지원합니다. VBA·매크로는 범위 밖입니다. 상세 계약은 [EXCEL-WORKBOOK-OPS.md](docs/EXCEL-WORKBOOK-OPS.md)입니다.

0.4.22는 이와 함께 사용자가 지목한 Excel 창에만 붙는 인스턴스 핀(`excel_launch`의 `processId`/`hwnd`/`activeWindow`, 이후 호출은 핀 창 전용, 소멸 시 명시 오류)과 테두리 공유 변 간섭 방지(같은 배치의 그리기 뒤 지우기는 `[EXCEL_BORDER_ORDER]` 거절, 다른 배치는 `[EXCEL_BORDER_SHARED_EDGE]` 경고)를 추가합니다. `excel_launch`는 워커 경로(`--excel-worker`)로도 연결됩니다.

0.4.21은 Excel **지연 서식 체크포인트**를 실제 execute 경로(호스트와 CLI worker)에 추가합니다. 기본값은 꺼져 있고 `DOCBRIDGE_EXCEL_DEFERRED_FORMAT_SNAPSHOT=1`(정확히 `1`만 활성)일 때만 후보를 검사합니다. 적용 전에 저장 전 메모리 상태를 체크포인트로 복사하고, 복구가 필요할 때 그 사본에서 **그 작업이 바꾼 서식만** 읽어 되돌립니다. 전체 워크북을 덮어쓰지 않으며 체크포인트 이후에 입력한 값·수식·대상 밖 서식은 유지합니다. 사용자 승인 단계를 새로 추가하지 않았습니다. 정상 적용이 빨라지는 대신 **복구는 느려집니다**. 후보 범위는 단일 `format_range`(`fillColor` 포함, 시트와 직사각형 A1 명시), 1,000~5,000셀, 로컬 일반 `.xlsx`, 원본과 체크포인트 각각 16MiB 이하입니다. **모든 수식**과 매크로·외부 연결·보호·암호·병합·웹 추가 기능 같은 복잡한 파트가 있으면 **변경 전에** 기존 스냅샷 경로로 되돌아갑니다. 다만 체크포인트 사본을 시도한 뒤 워크북이 그대로임을 확인하지 못하면 되돌아가지 않고 스냅샷을 거절해 그 편집이 진행되지 않게 합니다. 복구는 Excel과 대상 워크북이 살아 있는 세션을 요구하며 크래시 복구가 아닙니다. 계약은 [Excel 작업 안내](docs/EXCEL-OPERATIONS.md), 변경 요약과 사용 지침은 [RELEASE-0.4.21.md](docs/RELEASE-0.4.21.md)에 있습니다.

0.4.21은 이와 함께 보조 workbook 사본의 출처를 metadata에 명시하고 저장 전 메모리 사본을 선택적으로 켤 수 있게 했으며, CAD의 객체당 중복 COM 속성 읽기 7회를 지문 변경 없이 제거하고, 한글의 문단 지도 생성에 상한을 두고 전체 텍스트 읽기 실패를 성공으로 숨기지 않습니다.

0.4.21 구현 빌드에서 확인한 범위(버전을 올리기 전 구현 측정이며, 아래 로컬 릴리즈 검증과 합산하지 않는다):

- 비-E2E Core 631·MCP 20, 통합 Release `-warnaserror` 경고 0·오류 0.
- Excel 1,000셀 실제 앱 7종(정상 적용·복구, 실제 부분 실패 자동 복구, 손상 체크포인트 거절, CLI 새 프로세스 복구, 구버전 거절, 기존/새 방식 비교, 수식 파일의 기존 경로 전환)과 5,000셀 2종(비교, 부분 실패 자동 복구) 통과. 대상 셀 서식 전수 비교 불일치 0.
- 적용 1,000셀 7.235→0.113초, 5,000셀 24.206→0.124초. 복구는 9.021→15.736초, 39.175→60.805초. 각 규모마다 순서가 있는 단일 표본이며 p50/p95나 만능 속도 배율이 아니다. 비용이 사라진 것이 아니라 캡처에서 롤백으로 옮겨갔다.
- 체크포인트 이후 입력한 값·수식·표시 형식의 보존은 **1,000셀 정상 사례와 1,000셀 CLI 사례에서만** 확인했다. 5,000셀 시험은 서식 복원과 부분 실패 복구다.
- 구버전 거절은 설치본 0.4.18, 이전 로컬 후보 빌드, GitHub에서 내려받은 공개 0.4.20 바이너리에서 실제로 통과했다.
- 이전 단계의 AutoCAD/GstarCAD native 4건과 한글 native 3건은 그대로 유지된다.
- 웹 추가 기능 파트를 통한 기존 경로 전환은 실제 앱에서 확인하지 못했다. 생성 파일에 그 정보가 들어가는지가 실행마다 달라 재현 가능한 시험 파일이 아니며, 전환 자체는 수식 파일로 검증했다.

### 0.4.21 로컬 릴리즈 검증

버전을 0.4.21로 맞춘 뒤 이 PC에서 통과한 로컬 게이트다. 위 구현 빌드 수치와 합산하지 않는다. GitHub CI와 공개 ZIP은 소스 푸시 뒤에 확인하며, 이 절은 그 통과를 주장하지 않는다. 진행 상태는 [Actions](https://github.com/mororii/doc-bridge/actions)와 [Releases](https://github.com/mororii/doc-bridge/releases)에서 본다.

- 통합 Release `-warnaserror` 경고 0·오류 0.
- 비-E2E Core 659·MCP 20. 필터는 `FullyQualifiedName!~E2ETests&Category!=E2E`다. 구현 빌드 Core 631과의 28건 차이는 예전 더 넓은 E2E 이름 필터에 걸려 빠졌던 `HwpE2EOwnershipTests` 순수 단위 테스트이며, 제품·테스트 소스를 바꾼 결과가 아니다.
- 버전 일치와 초보자 가이드 검증 통과.
- 공개 소스 검사 222파일에서 자격 증명·개인 자료 없음, 추적 PowerShell 19파일 구문 통과.
- win-x64 자체 포함 패키지 생성, 체크섬 407개 전수 일치, 빌드 PC·프로필 개인 경로 검사 통과.
- 가짜 사용자 폴더에서 설치·진단·제거 수명주기 통과.
- 패키지 CLI로 1,000셀 실제 Excel 왕복 45.999초 통과. worker가 지연 서식 메타데이터를 전달했고, 다른 CLI 프로세스에서 복구했으며, 대상 1,000셀 원래 서식 전수 일치, 이후 입력한 셀 값·수식·표시 형식과 다른 시트 내용 유지, 원래 Excel 세션 설정 유지, 소유 Excel 종료 후 잔류 없음.

기본값은 꺼져 있고, 후보 조건·수식/복잡 파트 제외·살아 있는 세션 복구 제한은 그대로다. 웹 추가 기능 파트를 통한 기존 경로 전환의 실앱 확인과 호스트 크래시 뒤 Excel 잔류는 이 게이트로 해소되지 않았다.

0.4.20은 GstarCAD 전용 `gstarcad_*` 도구 4개를 추가합니다. 기존 `cad_*`는 AutoCAD 전용으로 유지합니다. 연결 인스턴스, 승인 토큰, 스냅샷은 제품별로 분리하며 자동 대체 연결하지 않습니다. 기본 편집만 지원하고 해치·문서 간 복사·XREF·배치 편집·PDF·스크립트·RGB 색상은 차단합니다. 미지원 기능을 검증했다고 쓰지 않습니다. [GstarCAD 사용·검증 안내](docs/GSTARCAD.md)를 참고하세요.

일반 편집은 앱별 `autoExecuteOps`에서 `executionMode=execute`로 한 번에 적용할 수 있습니다. Excel은 `set_values`/`set_formulas`/`format_range`, 한글은 `default.policy`와 `core_get_capabilities`의 현재 목록(커서 삽입 `insert_text`는 제외), CAD/GstarCAD는 `set_text_value`/`set_layer_visibility`/`set_layer_color`/`regen_document`만입니다. Excel/CAD/GstarCAD execute는 저장된 절대 경로만 안내합니다. 현재 어댑터는 인스턴스 바인딩 참조를 발급하지 않습니다. `Book1`/`Drawing1` 같은 미저장 이름은 기존 토큰 경로이며, 경로를 만들기 위해 자동 저장하지 않습니다. 한글은 반환된 `documentRef`/`instanceRef`를 그대로 사용합니다. 미리보기와 고위험·구조 변경·삭제·저장·내보내기는 기존 dry-run → 같은 ops/`confirmToken` 경로입니다. CAD dirty 광범위 작업이나 저장 지문이 맞지 않으면 거절하며, 자동 저장으로 우회하지 않습니다.

Excel 서식-only는 v3 `written-properties` 스냅샷을 씁니다. Bold만 바꾸면 다른 색·tint를 읽거나 복구하지 않습니다. 이전 v2 전체 서식 스냅샷도 복구합니다. 오래된 format preview 토큰은 새 dry-run이 필요할 수 있습니다. 셀 한도 100,000과 STA 120초는 올리지 않았습니다. 저장된 소유 통합문서의 정상 disconnect/파이프 정리는 지원하지만 호스트 크래시 뒤 Excel 잔류는 반증 전까지 미해결입니다.

확인된 범위의 숫자와 단일 표본 제한은 [PERFORMANCE.md](docs/PERFORMANCE.md), 최신 릴리즈 노트 [RELEASE-0.4.21.md](docs/RELEASE-0.4.21.md), 이전 [RELEASE-0.4.20.md](docs/RELEASE-0.4.20.md)에 있습니다.

0.4.20에서 확인한 범위:

- 통합 Release `-warnaserror` 경고 0·오류 0, 비-E2E Core 421·MCP 20.
- Excel 균일 Bold 1,000/7,000셀과 혼합 1,000/7,000셀 전수 mismatch 0. 균일 7,000셀 execute 0.459초와 혼합 7,000셀 82.809초는 다른 작업이다. 만능 속도 배수를 만들지 않는다.
- 실제 COM 장애 원복, UUID 재전송 challenge/conflict, 작은 회귀 9건. 보호시트 건은 `rollback.verified=false`를 숨기지 않으며 원복 성공이 아니다.
- 한글 일반 execute/UUID 1건(31.269초), 제작 format/table/page/picture/PDF 1건(10.833초). 완성 문서 디자인 검수로 주장하지 않는다.
- AutoCAD/GstarCAD geometry production 2건(Gstar 25.382초, AutoCAD 4.204초). execute 최종 fixture 2행은 Gstar 23.390초, AutoCAD 4.281초. 두 시간은 다른 표본이며 합산하지 않는다. 미지원 Gstar 기능은 검증 범위가 아니다.
- `core_get_status` app 필터는 불필요한 앱 조회를 줄였으나 전체 status 속도 개선은 입증되지 않았다.
- 배포 ZIP·격리 설치는 패키지 생성 단계의 범위다. 전체 소스 릴리즈 완료를 주장하지 않는다.

이전 버전의 검증 기록(2026-08-31 기준):

- 오프라인 배포 게이트: 217/217(Core 비-E2E 198 + MCP 19), 0.4.15 실Excel RCW·복사·구조 복구 E2E 1/1, 0.4.17 비생성 조회·경로 탐색 실Excel E2E 1/1. 0.4.18 이슈 #1 실한글 E2E는 정상 시작 실행 1/1을 기록했으며, 이 PC의 신선한 인스턴스 반복 검사는 구버전 한글 2024(13.0.0.866) TourPopup 초기화 오류로 op 진입 전에 차단되어 권장 패치 13.0.0.3870 이상에서 재실행해야 합니다.
- 기존 실제 프로그램 E2E: 37/37
  - Excel: 임시 workbook 읽기 → dry-run → 값 변경 → readback → 복원
  - 한글: 기존 GUI 창 직접 연결 → 라이브 HWP 네이티브 백업 → 치환 → readback → 복원, 파일 기반 fallback 14/14
  - AutoCAD 2027: 임시 DWG 조회 → 레이어 변경 → readback → 복원
- Excel 추가 E2E: 숫자·수식·서식 적용/복원, 행 삽입/구조 복원
- Excel 다중 인스턴스 실제 검증: 다른 프로세스의 시트 복사 → 수식·값·열 너비 1,057항목 readback 일치
- Excel 0.4.15 RCW 안정화: `ActiveWorkbook`/`ActiveSheet`와 컬렉션 `Item()`이 공유하는 RCW는 획득한 참조만 1회 균형 해제합니다. 반복 활성 컨텍스트 조회와 같은 통합문서의 명시적 `copy_sheet` dry-run/apply 후에도 살아 있는 별칭을 분리하지 않으며, 컨텍스트 중간 오류를 `ok:true` 부분 성공으로 반환하지 않습니다. 기술 근거와 수용 기준은 [Excel RCW 핫픽스](docs/2026-08-20-EXCEL-RCW-HOTFIX.md)를 참고하세요.
- Excel/Cursor 0.4.17 빈 창 방지: status/context/dry-run은 Excel을 신규 생성하지 않고, 닫힌 파일 읽기는 명시적 `allowOpenFile:true`로만 허용합니다. 파일 열기 실패 뒤 소유 빈 인스턴스를 즉시 정리하고 Cursor의 `openpyxl`·직접 COM 우회를 차단합니다. 재현 근거와 수용 기준은 [Cursor Excel 빈 회색 창 수정 보고](docs/2026-08-21-CURSOR-EXCEL-BLANK-WINDOW-FIX.md)를 참고하세요.
- Excel 기본 편집 1차: `merge_cells`/`unmerge_cells`, 행·열 숨김/표시, 시트 일반 숨김/표시를 직접 COM op로 제공하고 `excel_read_range(includeLayout:true)`로 병합 영역과 숨김 상태를 검증합니다. 병합 데이터 손실, 활성/마지막 표시 시트 숨김, workbook 구조 보호를 적용 전에 차단하고 작업 범위 스냅샷으로 자동 복구합니다.
- 한글 0.4.8: 모든 표시 창과 탭을 `openDocuments`로 열거하고, 안정적인 `documentRef`/`instanceRef`로 문서별 읽기·편집, 중복 경로 차단, TypeLib doctor, 설치 Bin 작업 폴더 고정, 외부 worker 장애 격리, 문단 지도와 편집 후 재읽기
- 한글 0.4.10: Windows 이벤트 로그와 한글/WPF IL을 대조해 FontCache 종료 원인을 누락·오염된 process-level `windir`/`SystemRoot`로 확인했습니다. MCP가 축소된 환경으로 시작되어도 worker와 한글 COM 자식 프로세스에 검증된 Windows 경로를 명시적으로 주입하고 한글 설치 `Bin`에서 활성화합니다. 실제 `PopupBorderImpl`/`TourPopup`/FontCache 오류 창이 생긴 경우에만 동일 호출 재시도와 빈 창 반복 생성을 즉시 중단합니다.
- 비활성 창 작업 0.4.11: Excel·한글·AutoCAD COM 작업이 사용자의 다른 전경 앱을 유지하고, 한글은 원래 문서·커서·선택 영역을, CAD는 원래 도면·레이아웃·모델/종이 공간을 복원합니다. 같은 대상 앱에서 사용자 입력이 겹치면 남은 op를 안전 중단하고 `interaction` 진단을 반환합니다.
- 한글 0.4.11 보강: 표 행·열 삽입/삭제의 `count`를 한 줄씩 정확히 실행·구조 검증하고, `−`·`㎜` 같은 Unicode readback 표현 차이를 오판하지 않습니다. `dryRun`은 같은 배치의 앞 op 결과를 뒤 op가 이어받아 순차 시뮬레이션합니다.
- 한글 0.4.18 이슈 #1: 숫자 엔터티를 실제 Unicode로 복원하고, 범위·발생 순서를 지정하는 `find_replace`, 최대 500셀 `table_set_cells`, 안정적인 대형 표 inventory, `$pic`/`gso` 그림 검증, 수식 셀 보호를 추가했습니다. 긴 작업은 `hwp_submit_ops` 후 `hwp_get_job`으로 조회하여 클라이언트 timeout 뒤 같은 작업을 중복 제출하지 않습니다. 항목별 근거는 [HWP 이슈 #1 수정 보고](docs/2026-08-31-HWP-ISSUE-1.md)를 참고하세요.
- 한글 0.4.12 DOCX 우선 생성: 검수한 `.docx`를 한글 `FileOpen(OOXML)`로 가져와 새 `.hwpx`/`.hwp`로 저장합니다. 원본 SHA-256 불변, 비덮어쓰기, 새 출력 해시·크기, 처리시간, 표/쪽 수와 필수 문구 품질 게이트를 한 응답에서 검증합니다. Word 1쪽이 한글에서 머리말·꼬리말만 있는 빈 2쪽으로 바뀌는 경우에는 기대 쪽 수보다 정확히 1쪽 많고 마지막 문단이 비어 있을 때만 그 문단을 최소 높이로 축소합니다. 대표 A4 표 문서 4회 연속 변환에서 1쪽·표 6개·필수 문구가 모두 보존됐고 경고와 HWP 잔류 프로세스는 0개였습니다.
- 한글 DOCX 우선 제작 경로 고정: 새 문서 전에 읽기 전용 `hwp_plan_creation`으로 경로를 결정합니다. 새 일반 문서는 `docx-first`, 기존 HWP/HWPX·기존 한글 템플릿·한글 필드·복잡한 병합표·한글 전용 개체·원본 배치 보존은 `native-hwp`로 고정합니다. 대표 측정에서 DOCX 생성+한글 변환 핵심 작업은 약 0.93초, 기존 복합 HWP 적용 21건의 중앙값은 2.746초, 다단계 빈 문서 제작은 약 7.9초 이상이었습니다. 자세한 기준은 [한글 새 문서 제작 경로 정책](docs/HWP-CREATION-POLICY.md)을 참고하세요.
- 한글 표 사진: `insert_picture`에 `tableIndex`와 `row`+`col` 또는 `cellIndex`를 주어 셀 비율에 맞춰 삽입하고 기존 셀 내용을 보존하거나 명시적으로 비울 수 있습니다.
- 한글 전경 보존: 탭 활성화 직후 복구와 30ms 전경 감시로 긴 COM 호출 중 창이 앞으로 나오는 시간을 최소화합니다.
- 한글 표 높이: `table_set_row_height`로 행 전체를 mm 단위 조절하고 적용 후 실제 높이를 재측정
- 성능 0.4.8: 모든 쓰기 결과에 단계별 `timings`를 반환하고, HWP 전체 fingerprint가 같을 때 dry-run preview artifact를 실제 적용에서 재사용
- 한글 통합 읽기: `hwp_read_text(scope="bundle")`로 본문·문단 지도·구조·필드·표를 한 COM 연결에서 선택적으로 함께 읽음
- 한글 CLI 안정성: 건식 검증과 실제 적용이 서로 다른 프로세스여도 실행 중인 한글 창을 status 단계에서 다시 식별해 저장되지 않은 문서의 `untitled-*` 참조를 유지
- CAD 컨텍스트 경량화: `cad_get_active_context` 기본 `detailLevel="basic"`은 레이어·엔티티 순회를 생략하고 `nextActions`를 반환합니다. 제한 표본은 `summary`, 전체·영역 조회는 `cad_query_entities`의 `layers`/`regions`/`window` scope로 분리합니다.
- CAD 0.4.19: 편집 배치 후 변경 도면을 자동 재생성하고 `readback.displayRefresh`로 결과를 구분합니다. 화면만 갱신하는 `regen_document`도 dry-run/토큰 경로로 지원합니다. `scope:"layers"`는 `current/on/freeze/locked/modelVisible`을 구분하며, 조회 생략은 `layerSummaryStatus:"omitted"`, 조회 불가는 `null`로 명시합니다. 객체의 색상·표시·투명도는 `includeGeometry:true`로 읽습니다.
- Excel 0.4.8: workbook/객체 스캔, 수식 오류 검사, Protected View·모달 상태 진단
- MCP stdio와 Streamable HTTP 연결 검증
- MCP 프로토콜 `2024-11-05`, `2025-03-26`, `2025-06-18` 협상
- Codex 플러그인 공식 validator 통과
- Windows PowerShell 5.1과 PowerShell 7 스크립트 호환
- Cursor MCP 실무 연결 검증: Excel·한글·약 26만 객체 AutoCAD 도면의 연결·읽기와 쓰기 dry-run 확인, 사용자 전역 설정 자동 병합·진단·제거 지원

실제 E2E 로그는 [e2e-result.log](../e2e-result.log)에 있습니다.
GitHub 기본 CI와 전용 Windows self-hosted 실기기 E2E의 역할, 대상 PC 준비와 수동 실행 절차는 [실제 앱 E2E 운영 안내](docs/REAL-APP-E2E.md)를 참고하세요.

## 바로 실행

이미 만들어 둔 `dist`는 .NET 런타임까지 포함한 `win-x64` self-contained 배포본입니다.

```powershell
cd C:\Tools\DocBridge
.\dist\doc-bridge-mcp.exe --version
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\verify-mcp.ps1
# Excel 문서를 먼저 연 뒤 RCW 반복 조회까지 검사
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\verify-mcp.ps1 -RequireExcelRuntime
```

소스를 고쳐 다시 발행할 때:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\publish.ps1 -SelfContained
```

## 구성

```text
.codex-plugin/plugin.json  Codex 플러그인 매니페스트
skills/document-automation Codex용 안전 작업 지침
src/DocBridge.Core         정책, 토큰, 스냅샷, 감사 로그, 앱 어댑터
src/DocBridge.Mcp          stdio/HTTP MCP 서버와 29개 tool
src/DocBridge.HwpWorker    한글 COM 장애 격리·자동 교체 worker
src/DocBridge.Cli          MCP를 못 쓰는 환경용 동일 기능 CLI
ops/                       allowlist 정책, 스키마, CAD 허용 템플릿
clients/                   Claude·Codex·Kimi·Cursor 설정과 안전 규칙 예시
tools/                     발행, MCP 검증, 실제 프로그램 E2E
dist/                      최종 self-contained 실행 파일과 생성 설정
tests/                     단위, 프로토콜, 실제 앱 E2E
```

상위 폴더의 `cad-work`는 특정 도면을 처리하던 실험/작업 산출물 279개입니다. 범용 MCP 런타임에는 포함하지 않았고 원본도 수정하지 않았습니다.

## Cursor 연결

통합 설치기는 `%USERPROFILE%\.cursor\mcp.json`의 기존 설정을 백업하고 `mcpServers.doc-bridge` 항목만 사용자 전역으로 병합합니다. 프로젝트별 `<프로젝트>\.cursor\mcp.json`은 팀이나 프로젝트의 의도와 충돌할 수 있으므로 자동으로 만들거나 변경하지 않습니다. 설치 후 Cursor를 완전히 재시작하고 MCP 설정 화면에서 `doc-bridge` 상태를 확인한 다음 `core_ping`과 `core_get_status`를 호출합니다.

Cursor용 안전 규칙과 상세 안내는 [Cursor 사용 안내](clients/cursor/CURSOR_USAGE.md)에 있습니다. 프로젝트 규칙은 [docbridge-safe-automation.mdc](clients/cursor/rules/docbridge-safe-automation.mdc)를 `<프로젝트>\.cursor\rules`에 선택적으로 복사하고, 모든 프로젝트에 적용할 사용자 규칙은 [docbridge-user-rule.txt](clients/cursor/docbridge-user-rule.txt)를 Cursor Settings → Rules에 붙여 넣습니다. 설치기는 두 템플릿을 `%LOCALAPPDATA%\DocBridge\generated-configs\cursor`에도 복사합니다.

Cursor와 Codex·Claude·Kimi를 함께 연결할 수는 있지만 같은 Excel·한글·AutoCAD 문서를 동시에 수정하면 안 됩니다. 한 클라이언트의 작업과 readback을 끝낸 뒤 다른 클라이언트에서 문서를 다시 읽고 새 execute 또는 dry-run을 만듭니다. 공식 설정 위치와 규칙 형식은 [Cursor MCP 문서](https://docs.cursor.com/context/model-context-protocol)와 [Cursor Rules 문서](https://docs.cursor.com/context/rules)를 참고하십시오.

## MCP tools

| 영역 | tools |
|---|---|
| core | `core_ping`, `core_get_status`, `core_get_capabilities`, `core_disconnect`, `core_create_snapshot`, `core_list_snapshots`, `core_restore_snapshot` |
| Excel | `excel_get_active_context`, `excel_read_range`, `excel_inspect`, `excel_apply_ops`, `excel_disconnect` |
| 한글 | `hwp_plan_creation`, `hwp_launch`, `hwp_get_active_context`, `hwp_doctor`, `hwp_repair_typelib`, `hwp_read_text`, `hwp_apply_ops`, `hwp_submit_ops`, `hwp_get_job` |
| CAD | `cad_launch`, `cad_get_active_context`, `cad_query_entities`, `cad_apply_ops` |
| GstarCAD | `gstarcad_launch`, `gstarcad_get_active_context`, `gstarcad_query_entities`, `gstarcad_apply_ops` |

허용 op와 `autoExecuteOps`는 [default.policy.json](ops/policies/default.policy.json)과 `core_get_capabilities`에서 확인합니다. 목록에 없는 op, 임의 매크로, 외부 스크립트는 차단됩니다. `export_pdf`처럼 기존 출력 파일을 교체할 수 있는 작업은 기존 dry-run 경로와 별도의 고위험 확인이 필요합니다. `core_get_status`는 `app`으로 한 앱만 조회할 수 있습니다. 한 앱 작업에 ping·전체 앱 status·반복 context·전체 읽기·별도 승인·전체 재조회를 강제하지 않습니다.

Excel 병합·숨김 5종의 정확한 payload, `includeLayout`, batch 분리 규칙과 후속 기능 단계는 [Excel 기본 편집 operations](docs/EXCEL-OPERATIONS.md)를 참고하세요.

MCP 2025 tool annotation으로 읽기/파괴 가능 도구를 표시하지만, 이 힌트는 서버의 토큰·승인 검사를 대신하지 않습니다.

## 앱별 동작

### Excel

- Excel COM은 전용 `--excel-worker` 프로세스가 소유합니다. DocBridge가 만든 저장된 통합문서는 `excel_disconnect`와 worker 파이프 정상 종료에서 회수됩니다. 호스트가 비정상 종료된 뒤에는 빈 인스턴스와 저장된 소유 통합문서 모두에서 `EXCEL.EXE`가 남는 것이 관측되었습니다. 잔류 프로세스가 해결되었다고 보지 마십시오. 사용자가 연 Excel은 강제 종료하지 않습니다.
- 실행 중인 Excel을 찾으면 `ownsInstance=false`로 연결하여 참조만 해제하고 절대 `Quit()`하지 않습니다. 실행 중인 Excel이 없을 때만 표시 상태(`Visible=true`)로 새 인스턴스를 만들고 `ownsInstance=true`로 추적합니다.
- Excel을 종료하고 다시 실행해도 살아 있는 창과 workbook을 재탐색해 같은 MCP 세션에서 COM 연결을 자동 복구합니다.
- 실행 중인 Excel의 활성 workbook에 연결하고, `copy_sheet`는 모든 Excel 인스턴스를 내부 COM으로 탐색합니다.
- `core_get_status({"app":"excel"})`는 Excel을 실행하지 않습니다. `apps.excel.connected:true`이고 `document`가 비어 있지 않을 때만 `excel_get_active_context`를 호출하며, 상태가 바뀌지 않은 실패를 반복하지 않습니다.
- workbook 경로만으로 닫힌 파일을 자동 실행하지 않습니다. 사용자가 닫힌 기존 파일을 열어 **읽으라고 명시한 경우에만** 절대 경로와 `allowOpenFile:true`를 함께 사용할 수 있고, 쓰기는 Excel에서 이미 열린 workbook만 대상으로 합니다.
- DocBridge 연결 오류를 `openpyxl`, `pywin32`/직접 Excel COM, PowerShell COM, `Start-Process` 또는 UI 자동화로 우회하지 않습니다. 이런 우회는 서식·매크로 손상이나 통합문서 없는 회색 Excel 인스턴스를 남길 수 있습니다.
- 문자열·논리값·수식을 보존하고, 숫자는 Excel `Range.Value2`의 실제 COM 형식인 `Double`로 정규화해 `Int32` SAFEARRAY 캐스팅 오류를 방지합니다. 15자리보다 긴 식별번호는 숫자가 아니라 문자열로 입력합니다.
- 모든 쓰기 op는 대상 시트를 명시해야 합니다. `target.sheet:"매출"` 또는 `range:"'매출'!B2"`를 사용하며, 두 값이 다르면 적용 전에 차단합니다. 읽기만 하는 경우에는 활성 시트를 사용할 수 있습니다.
- 찾기/바꾸기는 수식 셀을 건드리지 않습니다.
- `merge_cells`는 좌상단 외 셀의 값·수식 손실 가능성과 기존 병합 영역의 부분 겹침을 차단합니다. `unmerge_cells`는 한 병합 영역 안의 단일 셀 또는 병합 영역 전체를 대상으로 하며 두 op는 한 batch에서 단독으로 실행합니다.
- `set_rows_hidden`/`set_cols_hidden`은 `hidden:true|false`, `set_sheet_visibility`는 `visibility:"hidden"|"visible"`을 사용합니다. 활성 시트와 마지막 표시 시트는 숨기지 않으며 `veryHidden` 신규 설정은 지원하지 않습니다.
- 병합·숨김 상태를 확인하려면 `excel_read_range`에 `includeLayout:true`를 지정합니다. visibility op는 서로 한 batch에 묶을 수 있지만 값·서식·복사·병합 op와는 별도 dry-run으로 나눕니다.
- `format_range`는 `bold`/`fontBold`, `italic`/`fontItalic`, `fillColor`/`interiorColor`/`fill` 별칭을 받고 충돌·미지원 키는 snapshot 전에 거부합니다. dry-run before는 실제 셀 서식이며, 채움 없음은 `fillPattern`으로 검정(`Color=0`)과 구분합니다. v3 `styleScope=written-properties`는 실제로 쓰는 속성과 색 커플링만 스냅샷·지문·복구합니다. Bold-only는 다른 색·tint를 건드리지 않습니다. 균일 bool/글꼴 크기/`NumberFormat`은 범위 빠른 경로를 쓰고, 혼합·색은 필요한 셀 상태를 확보한 뒤 복구 후 확인합니다. 부분 병합·혼합 리치 텍스트·비테마 RGB+tint는 쓰기 전에 거절합니다. 테마 tint는 허용합니다. 값·수식·구조가 섞인 배치는 전 문서 지문을 쓰지 않고 새 preview를 받습니다. 이전 v2 전체 서식 스냅샷도 복구합니다.
- 셀 값·수식·`format_range` 서식, `insert_rows`/`insert_cols` 구조를 스냅샷에서 복원하려고 시도합니다. 보호된 시트에서 Bold 등 COM 쓰기가 거절되면 롤백도 같은 보호 때문에 쓰지 못하고 `rollback.verified`는 false입니다. 모든 적용 실패가 검증된 롤백으로 끝난다고 보지 마십시오. 적용 실패는 op별 HRESULT와 `skipped`를 남기고, COM readback과 화면 육안 품질을 같은 증거로 쓰지 않습니다.
- `copy_sheet`는 같은 프로세스뿐 아니라 서로 다른 Excel 프로세스 사이에서도 시트 서식·수식·열 너비를 보존합니다.
- `copy_sheet` 복구는 복사된 시트만 제거하고 원래 시트 순서·활성 시트를 확인합니다. 기존 셀 수식을 전체 재기록하지 않도록 다른 Excel op와 한 배치에 섞는 것은 거부하며, 복사 후 새 dry-run 배치로 후속 편집합니다.
- 쓰기 전에 execute `expectedDocumentRef` 또는 dry-run의 `documentRef`가 현재 활성 workbook과 같은지 다시 검사합니다.
- `excel_inspect`는 시트별 사용범위·표·차트·도형·피벗, 수식 오류 셀, Protected View와 모달/수식 편집 상태를 읽기 전용으로 진단합니다.

Excel 인스턴스가 여러 개면 `excel_get_active_context`의 저장된 절대 경로 `documentRef`를 확인합니다. E2E 도구도 대상 임시 workbook이 아니면 쓰기를 즉시 중단합니다.
연결만 즉시 놓으려면 `excel_disconnect` 또는 `core_disconnect({"app":"excel"})`를 호출합니다. 이 명령도 사용자 인스턴스는 종료하지 않습니다.

```json
{"ops":[{"op":"set_values","target":{"sheet":"매출"},"range":"B2","values":[[1500]]}],"executionMode":"execute","requestId":"11111111-1111-1111-1111-111111111111","expectedDocumentRef":"C:\\작업\\매출.xlsx"}
```

미리보기나 미저장 `Book1`은 기존처럼 `dryRun:true` 후 같은 ops와 `confirmToken`을 씁니다. execute 객체에 `dryRun`/`confirmToken`/`highRiskConfirm`을 false로라도 넣지 않습니다. `requestId`는 호출마다 새 UUID이고, 같은 작업의 통신 재시도만 같은 UUID를 재사용합니다. 같은 UUID·앱·문서·ops는 저장된 결과를 다시 주고, 다른 payload는 거절합니다. `outcomeUnknown`이면 새 UUID로 무작정 재실행하지 말고 문서 상태를 먼저 확인합니다.

배포되는 PowerShell 스크립트는 Windows PowerShell 5.1 호환 UTF-8 BOM으로 검사됩니다. DocBridge의 일반 Excel 편집은 VBA를 사용하지 않습니다. 외부 VBA 모듈 교환이 꼭 필요한 고급 사용자는 배포본의 `support\Convert-DocBridgeTextEncoding.ps1`로 `.bas`를 CP949/CRLF로 변환한 뒤 가져옵니다.

### 한글

한글 2024가 게시하는 `!HwpObject.*` ROT 항목을 탐색해 사용자가 이미 열어 둔 표시 창을 우선 연결합니다. 연결한 사용자 창은 doc-bridge 종료 시 닫지 않습니다.

- `hwp_get_active_context.summary.openDocuments`는 모든 표시 한글 프로세스의 모든 문서 탭을 Excel의 `openWorkbooks`처럼 반환합니다. 저장 문서는 정규화된 절대 경로, 저장 전 문서는 `untitled-<PID>-<문서ID>`를 `documentRef`로 사용합니다.
- 열린 문서가 여러 개면 `hwp_read_text`와 모든 쓰기 op에 선택한 `documentRef`를 넣습니다. `instanceRef`(`hwp:<PID>:<문서ID>`)도 같은 대상으로 쓸 수 있어 한글 경로가 셸에서 깨지거나 동일 파일이 중복 열린 경우에도 정확한 인스턴스를 지정할 수 있습니다. 한 배치에서는 모든 op가 같은 `documentRef`를 사용해야 하며 `file`과 함께 지정할 수 없습니다.
- 한글 COM은 `doc-bridge-hwp-worker.exe`에서 격리 실행합니다. 모달이나 COM hang이 발생하면 worker만 종료·교체하며 MCP의 Excel/CAD 세션은 유지합니다.
- `hwp_doctor`는 한글 실행 전에 ProgID, 설치 버전, TypeLib GUID와 등록 경로를 비교합니다. `hwp_repair_typelib`은 명시 승인 후 설치된 `Hwp.exe /RegServer`로 복구합니다.
- 절대 `file` 경로를 주면 모든 표시 한글 창과 탭을 조사해 정확한 문서를 활성화합니다. 같은 파일이 둘 이상 열렸으면 `HWP_DUPLICATE_LOCAL_PATH`로 임의 편집을 거부합니다.
- `hwp_read_text(scope="document_map")`는 내용 hash 기반 `lineId`와 읽기 coverage를 반환합니다. 적용 결과의 `postEditReread`와 세션 상태를 다음 편집의 기준으로 사용합니다.
- 본문·문단 지도·구조를 함께 볼 때는 `hwp_read_text(scope="bundle", sections=["text","document_map","structure"])`를 사용합니다. 표/필드는 필요할 때만 sections에 추가하고 표 서식은 `includeStyles:true`를 명시합니다.
- dry-run 스냅샷에는 preview artifact가 저장됩니다. apply 직전 본문·선택·커서·control/field 구조·현재 서식 fingerprint가 같을 때만 preview를 재사용하며, 달라졌으면 토큰을 소비하지 않고 새 dry-run을 요구합니다.
- `*_apply_ops`의 `timings`에는 검증·자동화 잠금 대기·상태 확인·preview·snapshot·fingerprint·apply·rollback 시간이 포함됩니다. 성능 분석은 [PERFORMANCE.md](docs/PERFORMANCE.md)를 참고하세요.

- 새 문서를 만들기 전에 `hwp_plan_creation`을 호출합니다. 새 일반 문서는 `docx-first`, 기존 한글 템플릿·네이티브 필드·복잡한 병합표·한글 전용 개체·원본 배치 보존은 `native-hwp`로 결정되며 AI가 임의로 바꾸지 않습니다.
- 일반적인 새 문서는 DOCX에서 A4·표·서식을 먼저 완성하고 렌더 검수한 뒤 `hwp_launch`에 `creationMode:"docx-first"`, 절대 `sourceFile`과 새 `outputFile`을 지정해 OOXML로 가져옵니다. `expectedPageCount`, `expectedTableCount`, `requiredText`를 함께 넣고 `summary.verification.passed:true`일 때만 완료합니다. 기존 출력 파일은 덮어쓰지 않습니다. 운영 경로는 Word COM을 실행하지 않습니다.
- HWP 전용 양식이나 기존 템플릿 필드가 필요한 새 문서는 `hwp_launch({"creationMode":"native-hwp","newDocument":true})`를 작업 시작에 한 번 호출합니다. 생성된 빈 문서는 계속 열린 상태로 유지되며 이후 단계별 편집에서 재사용합니다.
- 기존 창 작업은 `hwp_read_text`와 write op의 `file`을 생략합니다. `hwp_get_active_context.summary.connectionMode`가 `existing-window`인지 확인합니다.
- 파일 기반 백그라운드 작업은 절대 `.hwp`/`.hwpx` `file` 경로를 지정합니다. 한 배치에서는 모든 op의 `file`을 모두 생략하거나 모두 같은 파일로 지정합니다.
- `file`을 생략한 실시간 작업은 사용자가 연 한글 창만 연결하며, 창을 찾지 못해도 빈 한글을 자동 실행하지 않습니다.
- 새 문서 작성이나 전체 본문 교체에는 `replace_document_text`를 사용하며, 입력한 줄바꿈을 실제 문단으로 보존합니다.
- 기존 문서 끝에 붙일 때는 `append_text`를 사용합니다. 여러 문단을 한 op로 입력할 수 있으며 PowerShell/Python COM으로 우회하지 않습니다.
- 기존 양식 중간에는 `insert_before_text`/`insert_after_text`를 사용합니다. 고유한 `anchor` 또는 반복 문구의 1부터 시작하는 `occurrence`로 위치를 고정하며, `mode=paragraph`는 주변 문단 서식을 상속하고 `mode=inline`은 기준 문구 바로 옆에 붙입니다. 적용 후 지정 occurrence 주변과 기준 문구 개수를 다시 검증합니다.
- `set_paragraph_style_basic`은 현재 선택뿐 아니라 `target.scope=document` 전체 문서와 `target.text` 특정 문구를 지원합니다. 글꼴·크기·색·음영·굵게·기울임·장평·자간·위/아래 첨자·밑줄·취소선과 문단 정렬을 적용할 수 있습니다.
- `set_paragraph_format`은 문단 정렬, 좌우 여백, 첫 줄 들여쓰기/내어쓰기, 문단 앞뒤 간격, 줄 간격, 외톨이줄/문단 보호를 적용합니다. `set_page_setup`, `insert_break`, `insert_page_number`, `set_header_footer_text`로 구역과 쪽 구성도 제어합니다.
- `insert_table`은 `rows` 2차원 문자열 배열로 실제 한글 표를 만들며, `columnWidths` 비율과 `headerFill`·`firstColumnFill` 색상, `fontSize`를 선택적으로 지정할 수 있습니다. `hideAllBorders`로 기본 선을 숨기고 `cellStyles`의 `borders`로 필요한 실선만 다시 그릴 수 있으며, `mergeCells`로 같은 행의 셀을 가로 병합할 수 있습니다.
- `table_cell_set_text`는 `tableIndex`, `row`, `col`(모두 0부터 시작)로 기존 표 셀의 내용을 정확히 교체하며, 여러 셀은 `table_set_cells` 한 op에 최대 500개까지 묶습니다. 표 행·열 추가/삭제는 요청한 `count`만큼 실제 셀 수 변화가 났는지 단계별 검증하고, 범위 병합과 기존 양식 필드 조회·값 변경도 지원합니다. `insert_picture`는 문서 포함/크기/효과를 지정하고 `$pic`과 `gso`를 모두 검증하며 `export_pdf`는 한컴 PDF 필터를 사용해 결과를 저장합니다.
- 한글 2024에서 숨은 대화상자를 일으킨 셀 나누기·새 필드·북마크·하이퍼링크 삽입은 무인 실행 안정성을 위해 공개 명령에서 제외했습니다.
- 전체 필드와 예시는 [HWP 작업 명세](docs/HWP-OPERATIONS.md)를 참고하세요.
- 공식 한컴 자동화 보안 모듈을 `hwp-security` 폴더에 동봉하고 `DocBridgeFilePathChecker` 전용 이름으로 사용자 레지스트리에 등록합니다. 설치는 `install-hwp-security.ps1`, 제거는 `uninstall-hwp-security.ps1`을 사용하며 한컴의 다른 등록값은 건드리지 않습니다.
- 라이브 창은 적용 전에 `GetTextFile("HWP")`로 저장 전 화면 상태까지 포함한 네이티브 전체 데이터를 백업하고 SHA-256으로 검증합니다. 파일 작업은 원본 바이너리를 그대로 복사합니다.
- 라이브 복원은 `SetTextFile("HWP")`와 텍스트 readback을, 파일 복원은 임시 파일·원자적 교체와 SHA-256 재검증을 사용합니다.

### AutoCAD

- `cad_launch`는 AutoCAD를 COM으로 직접 실행하거나 기존 인스턴스에 연결하고, 창을 표시한 뒤 활성 도면이 없으면 새 도면을 만듭니다.
- `draw_taegeukgi`는 AutoLISP나 `SendCommand` 없이 ActiveX `ModelSpace` 메서드로 태극기 객체를 직접 생성합니다.
- `draw_union_jack`은 같은 직접 COM 방식으로 영국 국기를 생성하며 `originX`, `originY`, `width`, `height`로 배치를 조절할 수 있습니다.
- `draw_block_wall_schematic`은 1:1 mm 시공도면으로 22,500×5,400mm 벽체와 정확한 600/400/200mm 블록 모듈, 확대 상세, 리더, 치수, 문자를 전용 레이어에 생성합니다.
- `cad_query_entities`는 열린 여러 DWG 또는 지정 DWG를 대상으로 경계·좌표·블록명·문자를 조회하고, `startIndex`/`endIndex`로 방금 추가된 객체만 빠르게 검증합니다.
- `copy_entities_between_documents`, `insert_xref`(기존 XREF 정의 재사용 지원), `draw_entities`, `zoom_window`는 도면 객체와 보기를 ActiveX COM으로 직접 편집합니다. 해치도 typed `AcadEntity[]` 경계 배열로 직접 생성하며 AutoLISP를 생성하거나 로드하지 않습니다. XREF 경계만 AutoCAD 기본 `XCLIP` 명령을 사용합니다.
- 선·호·타원·점·MText·정렬/회전 치수, 복사·축척·대칭·간격띄우기·공통 속성, 블록 속성값, 배치·뷰포트, DWG 저장·PDF 출력까지 지원합니다. `cad_query_entities(scope="regions")`는 최대 100개 도곽을 한 번의 ModelSpace 순회로 검증합니다.
- 전체 필드와 예시는 [CAD 작업 명세](docs/CAD-OPERATIONS.md)를 참고하세요.
- 실행 중인 AutoCAD ActiveX 인스턴스와 활성 DWG에 연결합니다.
- AutoCAD가 없을 때 `cad_query_entities(file=...)`로 DXF를 읽기 전용 분석할 수 있습니다.
- 임의 script는 실행하지 않으며 `ops/script-templates`의 등록 템플릿만 허용합니다.
- 템플릿 매개변수의 제어문자, 경로/명령 삽입, 미해결 placeholder를 차단합니다.

`delete_entities`와 `run_script_template`은 기존 dry-run 경로와 별도 `highRiskConfirm=true`가 필요합니다. execute는 저장된 절대 경로의 문자·레이어·regen만이며 도형·저장은 토큰 경로입니다. dirty 광범위 작업이나 저장 지문이 맞지 않으면 거절합니다. `Drawing1`을 저장해 경로를 만들거나 지문 거절을 우회하지 않습니다. 요청한 문자 핸들·레이어 속성 스냅샷은 그 범위만 완전 복구할 수 있으며 전체 도면 복구가 아닙니다. Geometry 스냅샷은 drawing-backup과 부분 상태이며 자동 복구 성공을 보장하지 않습니다. 이 경우 스냅샷의 `drawing-backup*.dwg`를 직접 열어 복구해야 합니다.

## 보안·운영

- confirmToken: HMAC-SHA256, 5분 TTL, 1회용, app/scope/정확한 ops에 바인딩
- execute `requestId`: UUID, 같은 앱/문서/ops만 재전송, 충돌 키(`dryRun`/`confirmToken`/`highRiskConfirm`) 포함 금지
- 토큰 발급 후 활성 문서가 바뀌면 apply 거부
- COM 호출: STA 스레드와 전역 named mutex로 직렬화
- 쓰기 전 스냅샷, 쓰기 후 실제 값 readback. execute는 한 잠금 안에서 사전검사부터 복구까지 수행
- 감사 로그: `%LOCALAPPDATA%\DocBridge\logs\audit-YYYYMMDD.jsonl`
- 스냅샷: `%LOCALAPPDATA%\DocBridge\snapshots\<app>\...`
- HTTP: 기본 `127.0.0.1`만 bind, 선택적 `DOCBRIDGE_HTTP_TOKEN` Bearer 인증

AutoCAD 프로세스를 작업 관리자에서 강제 종료하지 마세요. E2E는 사용자의 원본 파일이 아니라 `%TEMP%`의 작업 사본만 사용합니다.

## 개발 테스트

```powershell
dotnet test .\DocBridge.sln -c Release

# 실제 앱 임시 사본 E2E
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\run-e2e.ps1

# 앱 하나만
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\run-e2e.ps1 -Only excel
```

HTTP 검증:

```powershell
.\dist\doc-bridge-mcp.exe --http --port 5177
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\verify-mcp.ps1 -Http -Port 5177
```
