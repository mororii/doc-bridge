# Cursor Excel 빈 회색 창 수정 보고

## 현상

Cursor가 Excel 문서를 찾지 못한 뒤 통합문서가 없는 회색 Excel 창을 반복 실행하고,
작업이 끝난 뒤에도 `EXCEL.EXE` 또는 별도 자동화 인스턴스가 남았다.

## 확인된 원인

1. 구버전 `excel_get_active_context`와 일부 쓰기 준비 경로가 실행 중 workbook이 없어도
   Excel COM 인스턴스를 만들 수 있었다.
2. workbook 경로를 받은 읽기 호출은 별도 동의 없이 닫힌 파일을 자동으로 열 수 있었지만,
   쓰기는 열린 workbook만 지원해 읽기와 쓰기의 연결 계약이 달랐다.
3. Cursor는 DocBridge 오류 뒤 `openpyxl`로 원본 workbook을 다시 쓰고
   `pywin32.DispatchEx("Excel.Application")`로 별도 COM 인스턴스를 만드는 우회를 수행했다.
4. Excel worker가 COM 호출에 멈추면 표준입력 종료만으로는 worker가 EOF를 처리하지 못해
   RCW를 가진 하위 프로세스가 남을 수 있었다.

## 0.4.17 수정

- `core_get_status`, `excel_get_active_context`, diagnostics, dry-run, apply, snapshot은 Excel을
  신규 생성하지 않는다.
- `excel_read_range`와 `excel_inspect`는 기본적으로 열린 workbook만 읽는다.
- 사용자가 닫힌 파일 읽기를 명시한 경우에만 존재하는 절대 `workbook` 경로와
  `allowOpenFile:true`를 함께 허용한다. 쓰기에는 사용할 수 없다.
- 파일 열기 실패 또는 읽기 종료 뒤 DocBridge 소유 Excel의 workbook 수가 0이면 즉시
  `Quit()`과 COM 참조 정리를 수행한다. 사용자 소유 Excel에는 `Quit()`하지 않는다.
- status/context/read worker 제한시간을 일반적인 Cursor MCP 제한보다 짧게 두고, COM에
  고착된 경우 정확히 DocBridge가 시작한 하위 worker만 종료한다. Excel 프로세스나 프로세스
  트리를 강제 종료하지 않는다.
- MCP 서버 instructions와 Cursor 규칙은 동일 실패 반복, `openpyxl`, 직접 COM,
  PowerShell COM, `Start-Process`, 셸/UI 자동화, 원본 반복 덮어쓰기 우회를 금지한다.

## 수용 기준

1. Excel이 닫힌 상태에서 status/context/경로 probe/dry-run을 호출해도 새 `EXCEL.EXE`가
   생기지 않는다.
2. workbook 경로만 준 읽기는 파일을 열지 않으며 `allowOpenFile:true`가 있어야만 열린다.
3. 명시적 파일 열기 실패 뒤 DocBridge 소유 빈 Excel이 남지 않는다.
4. 사용자 Excel은 오류·disconnect·worker 정리 과정에서 종료되지 않는다.
5. Cursor는 DocBridge가 지원하지 않는 경로를 다른 Excel 라이브러리나 직접 COM으로
   우회하지 않고 원래 오류와 필요한 사용자 조치를 보고한다.
