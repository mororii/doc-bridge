# 실제 Excel·한글·AutoCAD E2E

일반 GitHub Actions는 Office·한글·AutoCAD가 없는 호스팅 환경에서 실행됩니다. 따라서 기본 CI는
실기기 테스트를 명시적으로 제외하고 단위/MCP 테스트만 실행합니다. `E2ETests`가 조용히 반환되어
통과처럼 보이지 않도록 한 조치입니다.

## 전용 테스트 PC

실제 COM 테스트는 업무용 PC가 아닌 별도 Windows PC에서만 실행합니다.

1. Excel, 한글 2024, AutoCAD를 설치하고 각각 한 번 정상 실행합니다.
2. 비공개 저장소용 GitHub self-hosted runner를 설치합니다.
3. runner에 `docbridge-e2e` 사용자 지정 레이블을 추가합니다.
4. GitHub Actions의 **Real Windows App E2E**를 수동 실행합니다.

워크플로는 필요한 세 ProgID를 먼저 검사하고 하나라도 없으면 실패합니다. 테스트 중에는 실제
프로그램과 COM 인스턴스를 생성·전환하므로 사람이 같은 PC에서 문서를 편집하면 안 됩니다.

## 로컬 실행

기존 작업 문서를 건드리지 않는 CLI 기반 전체 검증:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\tools\run-e2e.ps1
```

xUnit 실기기 세트만 실행하려면 다음 환경변수를 설정합니다.

```powershell
$env:DOCBRIDGE_E2E = '1'
$env:DOCBRIDGE_HWP_TABLE_COUNT_E2E = '1'
$env:DOCBRIDGE_HWP_UNICODE_E2E = '1'
$env:DOCBRIDGE_HWP_TABLE_PICTURE_E2E = '1'
$env:DOCBRIDGE_HWP_SEQUENTIAL_PREVIEW_E2E = '1'
$env:DOCBRIDGE_HWP_ISSUE1_E2E = '1'
dotnet test .\tests\DocBridge.Core.Tests\DocBridge.Core.Tests.csproj -c Release `
  --filter 'FullyQualifiedName~E2ETests|Category=E2E'
```

테스트가 끝나면 생성한 테스트 문서와 DocBridge 소유 COM 인스턴스를 정리하지만, 비정상 종료에
대비해 작업 관리자와 감사 로그도 함께 확인합니다. 사용자 소유 Excel·한글·AutoCAD 인스턴스는
강제 종료하지 않습니다.

한글 2024는 `hwp_doctor`의 `installedVersion`이 권장 안정 패치 이상인지 먼저 확인합니다. 구버전에서
TourPopup/FontCache `오류` 창이 뜨면 op 실패가 아니라 자동화 인스턴스 초기화 실패이므로 해당 실행은
수용 테스트 통과로 계산하지 않습니다. 한글을 업데이트하고 새 Windows 세션에서 다시 실행합니다.

## Excel RCW 0.4.15 수용 테스트

설치 상태만 확인하는 기본 `2-TEST.cmd`는 Excel을 새로 실행하지 않습니다. 실제 RCW 경로는 별도
테스트 PC에서 Excel 통합문서를 연 뒤 다음 명령으로 확인합니다.

```powershell
.\2-EXCEL-LIVE-TEST.cmd
# 또는
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Test-DocBridge.ps1 -RequireExcelRuntime
```

이 검사는 같은 MCP 세션에서 `excel_get_active_context`를 두 번 호출합니다. 두 응답 모두 `ok:true`,
`errors=[]`이어야 하고 `activeSheet`, `usedRange`, `saved`, `openWorkbooks`, `selection`과 동일한
`documentRef`가 유지되어야 합니다. 정식 실기기 세트는 같은 통합문서를 source/target으로 지정한
`copy_sheet` dry-run → apply → readback → snapshot restore → 후속 preview까지 검증합니다.
복구 후에는 복사 시트가 없어지고 원래 시트 이름·순서·활성 시트가 그대로여야 하며, 기존 셀의
Formula/Value2 setter가 호출되지 않아야 합니다. `copy_sheet`와 다른 Excel op의 혼합 배치는
검증 단계에서 거부되어야 합니다.
