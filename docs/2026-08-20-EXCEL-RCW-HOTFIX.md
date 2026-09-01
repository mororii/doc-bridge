# Excel RCW 수명주기 핫픽스

DocBridge 0.4.15는 Excel COM 객체의 공유 RCW(Runtime Callable Wrapper)를 과도하게 최종
해제하던 0.4.14 결함을 수정합니다. 이 문서는 개인 PC 경로와 업무 문서를 포함하지 않는 일반화된
기술 기록입니다.

## 증상

- `excel_get_active_context`가 `ok:true`이면서 `errors`에는 “COM object that has been separated
  from its underlying RCW”를 포함했습니다.
- `summary.workbook`과 `summary.sheets` 뒤의 `activeSheet`, `usedRange`, `openWorkbooks`,
  `selection`이 누락됐습니다.
- 같은 통합문서를 명시적 source/target으로 사용한 `copy_sheet` dry-run이 실패했습니다.

## 원인과 수정

Excel은 `ActiveWorkbook`/`ActiveSheet`와 `Workbooks.Item()`/`Worksheets.Item()`이 같은 RCW를
공유할 수 있습니다. 별칭 하나에 `Marshal.FinalReleaseComObject`를 호출하면 다른 지역 변수가
보유한 같은 RCW까지 분리됩니다.

다음 9개 별칭 가능 획득은 `Marshal.ReleaseComObject` 1회에 해당하는 균형 해제로 변경했습니다.

| 경로 | 변경 수 |
|---|---:|
| `OpenWorkbook.Dispose` | 2 |
| `WorkbookLease.Dispose` | 1 |
| `ResolveTargetWorkbook` | 2 |
| `GetActiveContext` 시트 열거 | 1 |
| `ListOpenWorkbooks` 시트·통합문서·활성 통합문서 | 3 |

DocBridge가 직접 생성해 소유하는 Excel `Application` 루트의 최종 해제와 안전한 `Quit` 정책은
그대로 유지합니다. 활성 컨텍스트는 모든 필수 필드를 채운 뒤에만 `ok:true`가 되며 중간 예외가
발생하면 반드시 `ok:false`입니다.

`copy_sheet`의 dry-run 복구도 작업 범위에 맞게 바꿨습니다. 복구 시 기존 통합문서 전체의
`UsedRange.Formula`를 다시 쓰지 않고, 새로 복사한 대상 시트만 역순으로 삭제한 뒤 원래 시트
이름·순서·개수와 활성 시트를 검증합니다. 따라서 기존 수식·값에는 쓰기 작업이 발생하지 않습니다.
`copy_sheet`는 다른 Excel 작업과 같은 배치에 섞을 수 없으며, 복사를 적용한 다음 새 dry-run
배치에서 후속 셀 편집을 수행해야 합니다.

## 검증 기준

- 같은 MCP 세션의 활성 컨텍스트 2회가 모두 `ok:true`, `errors=[]`
- `activeSheet`, `usedRange`, `saved`, `openWorkbooks`, `selection` 존재
- 같은 통합문서 `copy_sheet` dry-run/apply/readback 성공
- snapshot restore 후 복사 시트 제거, 원래 시트 순서와 기존 수식·값 불변
- 후속 read/preview 성공
- 사용자 Excel 연결 시 새 Excel PID를 만들지 않고 disconnect 후 사용자 Excel 유지
- DocBridge 소유 Excel은 disconnect/worker EOF 뒤 잔류 PID 없음

설치 후 Excel 통합문서를 연 상태에서 다음 명령으로 읽기 전용 반복 컨텍스트 검사를 실행합니다.

```powershell
.\2-EXCEL-LIVE-TEST.cmd
```
