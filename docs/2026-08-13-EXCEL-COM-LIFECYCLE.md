# Excel COM 인스턴스 생명주기 수정 기록

## 사용자 피해

Excel 도구 호출 뒤 모든 창을 닫아도 DocBridge가 보관한 COM RCW 때문에 `EXCEL.EXE`가
남을 수 있었다. 이후 탐색기에서 `.xlsx`를 더블클릭하면 Windows DDE가 이 자동화
인스턴스에 문서를 붙일 수 있고, 그 인스턴스에는 사용자의 XLSTART, XLAM, COM
추가기능이 로드되지 않아 매크로와 추가기능이 모두 사라진 것처럼 보였다.

## 확인된 원인

1. Excel Application 루트 RCW를 MCP 서버 수명 동안 보관했다.
2. 마지막 통합문서가 닫힌 뒤 참조를 놓는 유휴 감시가 없었다.
3. `AccessibleObjectFromWindow`가 만든 Window RCW와 ROT 중복 후보가 해제되지 않았다.
4. Workbook/Worksheet/Range 같은 임시 RCW가 남아도 연결 해제 시 즉시 수거하지 않았다.
5. 테스트 정리 코드도 `app.Workbooks` 점 체이닝과 `foreach (dynamic ...)`를 사용해
   실제 누수를 숨길 수 있었다.
6. 프로세스를 즉시 강제 종료하면 종료 훅이 실행되지 않고, Excel을 처음 생성한 프로세스의
   고아 COM 참조는 다른 프로세스의 사후 `Quit()`만으로 회수되지 않았다.
7. `ActiveWorkbook`/`ActiveSheet`와 컬렉션 `Item()`이 같은 RCW를 돌려줄 수 있는데,
   열거 중인 별칭에 `FinalReleaseComObject`를 호출하면 계속 사용해야 할 다른 별칭까지
   native COM 객체에서 분리되어 `InvalidComObjectException`이 발생했다.

## 적용한 소유권 규칙

- ROT/실행 창에서 찾은 Excel: `ownsInstance=false`
  - 상태: `사용자가 열어 둔 엑셀 창에 연결됨`
  - 연결 해제와 서버 종료에서 `Application.Quit()`를 절대 호출하지 않는다.
- 실행 중 인스턴스가 없어 DocBridge가 만든 Excel: `ownsInstance=true`
  - `Application.Visible=true`
  - 상태: `DocBridge가 생성한 인스턴스`
  - 저장되지 않은 통합문서가 하나라도 있거나 저장 상태를 확인할 수 없으면 Quit하지
    않고 참조만 해제한다.
  - 모두 저장됐거나 통합문서가 0개일 때만 Quit하고 정상 프로세스 종료를 기다린다.
- `DisplayAlerts=false`를 종료 수단으로 사용하지 않으며 `taskkill`도 사용하지 않는다.

## 전용 Excel worker

- 운영 경로의 Excel COM 객체는 MCP/CLI 본체가 아니라 같은 배포본의 숨은 `--excel-worker`
  프로세스가 처음부터 끝까지 소유한다.
- 본체와 worker는 UTF-8 NDJSON 파이프로만 요청·응답한다. 본체가 정상 종료하면 stdin을 닫고,
  강제 종료되면 운영체제가 파이프를 닫는다.
- worker는 EOF를 받으면 자신의 `ExcelAdapter.Dispose()`를 정상 실행한다. `ownsInstance=true`이고
  모든 통합문서가 저장된 경우에만 `Application.Quit()`하며, 사용자 인스턴스나 미저장 문서는
  참조만 해제한다.
- 실행되지 않을 수 있는 `ProcessExit` 훅에만 의존하지 않으며, worker가 Excel을 만든 원래 COM
  소유 프로세스이므로 사후 연결 방식의 고아 RCW 문제를 피한다.

## 자동·명시적 정리

- 1초 간격 생명주기 감시가 마지막 통합문서가 닫힌 연결을 STA에서 해제한다.
- `core_disconnect({"app":"excel"})`와 `excel_disconnect`를 제공한다.
- stdio EOF, 정상 종료, Ctrl+C, SIGINT, SIGTERM, `ProcessExit`, 런타임 unload에서
  `DocBridgeHost.Dispose()`가 한 번만 실행되도록 종료 훅을 등록했다.
- COM 참조는 획득 방식과 소유권에 따라 구분해 해제한다.
  - DocBridge가 단독 소유하고 더 이상 다른 별칭이 없는 Application root와 고유 임시 RCW만
    최종 소유권 경계에서 `FinalReleaseComObject`한다.
  - `ActiveWorkbook`, `ActiveSheet`, `Workbooks.Item()`, `Worksheets.Item()`처럼 기존 RCW의
    별칭이 될 수 있는 반환값은 획득한 참조 수만큼 `ReleaseComObject`를 정확히 한 번 호출해
    균형 해제한다. 이런 별칭에는 `FinalReleaseComObject`를 호출하지 않는다.
  - `app.Workbooks.Item(...).Worksheets.Item(...)` 같은 점 체이닝을 피하고 각 COM 반환값을
    변수에 받아 역순으로 정리한다.
- 명시적 COM 해제 뒤에는 두 번의 GC/finalizer pass를 실행해 작업 메서드의 수명이 끝난
  숨은 임시 RCW까지 정리한다.

## 검증 결과

- 비-E2E/MCP 회귀: Core 147건 + MCP 17건 = 164건 통과.
- 실제 Microsoft 365 Excel 집중 회귀: Excel 관련 25건 통과.
- Excel이 없는 상태에서 CLI 진단이 만든 인스턴스: 명령 종료 직후 프로세스 0건.
- 사용자 Excel 모의 경로: PID 유지, 새 PID 0건, 통합문서 유지, 상태 detail 정확.
- 장시간 MCP 경로: 사용자가 Excel을 닫은 뒤 MCP는 계속 실행된 상태에서 약 3.23초
  안에 해당 `EXCEL.EXE`가 종료됨.
- 실제 `Excel_full_flow`: 테스트 전후 PID 집합이 동일하여 새 잔류 인스턴스가 없음.
- 전용 worker 강제 종료 모사: 부모 프로세스를 강제 종료한 뒤 worker가 파이프 EOF를 감지하여
  테스트 전용 Excel PID를 정상 종료함.
- 기존 사용자 Excel 연결 상태에서 MCP 강제 종료 모사: worker PID만 종료되고 기존 Excel PID
  집합과 열린 통합문서는 그대로 유지됨.

전체 실프로그램 묶음 실행에서는 Excel 외 CAD `ActiveLayout` 환경 상태 1건과 한글 사용자 활동
감지 1건이 안전 중단되었다. Excel 집중 테스트와 비-E2E/MCP 회귀에는 실패가 없었다.

## 운영 참고

업데이트 전 MCP 프로세스가 이미 Excel RCW를 보유했다면 새 바이너리를 설치하는 것만으로
그 구버전 프로세스의 메모리 상태가 바뀌지는 않는다. Codex, Claude, Kimi를 완전히 종료해
구버전 MCP를 끝낸 뒤 다시 실행해야 새 생명주기 정책이 적용된다.
