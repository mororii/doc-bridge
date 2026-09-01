# DocBridge 통합 설치 패키지

이 패키지는 Windows 10/11 x64에서 Excel, 한글, AutoCAD를 제어하는 로컬 MCP를 Codex, Claude Code/Desktop, Kimi, Cursor에 연결합니다. .NET 런타임이 포함되어 있으므로 대상 PC에 .NET을 별도로 설치할 필요가 없습니다.

## 초보자 설치

1. ZIP 파일에서 마우스 오른쪽 버튼을 누르고 **모두 압축 풀기**를 선택합니다. ZIP 내부에서 바로 실행하면 안 됩니다.
2. 압축을 푼 폴더의 `사용설명서_보고따라하세요.html`을 더블클릭해 그림 형태의 안내서를 엽니다.
3. Excel, 한글, AutoCAD와 Codex/Claude/Kimi/Cursor를 모두 종료합니다.
4. 선택 사항으로 `0-VERIFY.cmd`를 더블클릭해 패키지 무결성을 확인합니다.
5. `1-INSTALL.cmd`를 더블클릭하고 반드시 `INSTALLATION SUCCESS`가 나올 때까지 기다립니다.
6. AI 클라이언트를 다시 실행한 뒤 `2-TEST.cmd`로 연결을 확인합니다. Cursor는 `Cursor global config`가 `[OK]`인지 확인합니다.
7. Excel 통합문서를 하나 연 뒤 `2-EXCEL-LIVE-TEST.cmd`를 실행해 반복 활성 컨텍스트와 RCW 상태를 확인합니다. 이 검사는 Excel 문서가 열려 있지 않으면 실패하며 새 Excel을 자동 실행하지 않습니다.

`2-TEST.cmd`는 설치 프로그램이 아닙니다. 설치 전에 실행하면 `TEST NOT STARTED`와 함께 `1-INSTALL.cmd`를 먼저 실행하라는 안내만 표시됩니다. 설치 창에 `INSTALLATION FAILED`가 표시되면 `2-TEST.cmd`를 실행하지 말고 그 설치 창의 오류 내용을 확인하십시오.

명령 입력에 익숙한 사용자는 다음 명령으로 같은 설치를 실행할 수 있습니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-DocBridge.ps1
```

설치 위치는 기본적으로 `%LOCALAPPDATA%\DocBridge`입니다. 설치 프로그램은 기존 Claude Desktop/Kimi JSON과 Cursor 사용자 전역 `%USERPROFILE%\.cursor\mcp.json`을 백업하고 `doc-bridge` 항목만 병합합니다. 프로젝트별 `.cursor\mcp.json`은 변경하지 않습니다. 한글 자동화 보안 모듈은 현재 사용자(HKCU)에 등록되며 재부팅 뒤에도 유지됩니다.

특정 클라이언트만 연결하려면:

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Install-DocBridge.ps1 -Clients Codex,ClaudeDesktop,Cursor
```

설치기는 PATH뿐 아니라 Codex 앱 내부의 CLI도 자동 탐색합니다. 구형 CLI의 `plugin add/list` 방식과 현재 CLI의 marketplace 전용 방식을 자동 판별합니다. 현재 CLI에서는 marketplace 등록을 확인하고, 배포 정책으로 플러그인을 기본 활성화한 뒤 Codex 재시작을 안내합니다. 수동 등록 명령은 `%LOCALAPPDATA%\DocBridge\generated-configs`에 생성되며, 발견된 Codex CLI의 절대 경로를 사용합니다. Claude Code가 설치되지 않은 경우에는 해당 등록만 건너뜁니다.

## 진단

초보자는 `2-TEST.cmd`를 더블클릭합니다. 명령 입력 방식은 다음과 같습니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Test-DocBridge.ps1
```

핵심 실행 파일, 자체 포함 런타임, MCP 핸드셰이크, 클라이언트 설정, 한글 보안 모듈 등록을 검사하고 `%LOCALAPPDATA%\DocBridge\doctor-report.json`을 만듭니다.

기본 진단은 Excel을 띄우지 않습니다. Excel 통합문서를 연 뒤 `2-EXCEL-LIVE-TEST.cmd`를 실행하면
같은 MCP 세션에서 활성 컨텍스트를 두 번 읽어 공유 RCW 분리와 부분 성공 응답을 추가 검사합니다.

Cursor는 사용자 전역 설정과 안전 규칙 템플릿을 함께 검사합니다. 설치된 안내와 규칙은 `%LOCALAPPDATA%\DocBridge\generated-configs\cursor`에 있습니다. Cursor Settings → Rules에는 `docbridge-user-rule.txt`를 붙여 넣고, 프로젝트 규칙이 필요할 때만 `rules\docbridge-safe-automation.mdc`를 프로젝트의 `.cursor\rules`에 복사합니다.

Codex를 설치한 PC에서는 `Codex CLI discovery`, `Codex plugin CLI mode`, `Codex marketplace registered`, `Codex MCP visible and enabled`가 모두 `[OK]`이어야 합니다. 구형 CLI에서는 `Codex plugin installed and enabled`가 대신 표시됩니다. 설치 직후 MCP 항목만 `[WARN]`이면 Codex를 완전히 재시작하고 새 작업을 연 뒤 `2-TEST.cmd`를 다시 실행합니다.

- `[OK]`: 정상
- `[WARN]`: 핵심 설치는 되었으나 해당 AI의 수동 등록이나 재시작이 필요함
- `[SKIP]`: 해당 AI 프로그램이 이 PC에 없어 검사를 건너뜀(오류 아님)
- `[FAIL]`: 핵심 설치나 연결에 실제 문제가 있음

## 제거

초보자는 `3-UNINSTALL.cmd`를 더블클릭하고 `Y`를 누릅니다. 명령 입력 방식은 다음과 같습니다.

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File .\Uninstall-DocBridge.ps1 -RemoveHwpSecurity
```

기존 설정의 다른 MCP는 유지하고 `doc-bridge` 항목만 제거합니다. 설치 파일은 즉시 삭제하지 않고 `backups` 아래로 옮깁니다. 백업과 로그까지 모두 지우려면 명시적으로 `-RemoveData`를 추가합니다.

같은 이름의 항목이 설치 전에 있었다면 설치 기록의 원본을 복원합니다. 설치 후 사용자가 command/args를 바꾼 항목은 제거하지 않습니다. `-Clients Cursor`처럼 일부만 제거하면 공유 실행 파일과 한글 보안 등록은 유지하며, 마지막 관리 클라이언트를 제거할 때만 payload와 활성 `installation.json`을 백업으로 옮깁니다.

## 사용 전 확인

- Codex, Claude Desktop, Cursor는 설치 후 완전히 종료했다가 다시 시작합니다.
- Codex는 재실행 후 반드시 새 작업을 시작합니다. 기존 작업은 설치 전 도구 목록을 계속 사용할 수 있습니다.
- 새 작업에서 DocBridge 도구가 없다면 프로젝트 파일이나 대체 스크립트를 수정하게 두지 말고 `2-TEST.cmd`부터 다시 확인합니다.
- 제어할 Excel/한글/AutoCAD 문서는 해당 Windows 사용자 세션에서 열어 둡니다.
- 한글 창이나 탭이 여러 개면 AI에게 `hwp_get_active_context.summary.openDocuments`를 먼저 표로 보여 달라고 하고, 선택한 항목의 `documentRef`를 읽기와 수정에 계속 사용하게 합니다. 한글 경로가 깨지면 `instanceRef`를 `documentRef` 값으로 사용합니다.
- AI가 먼저 활성 문서를 읽고, 변경안(dry-run)과 백업 위치를 확인한 뒤 적용하도록 합니다.
- Cursor·Codex·Claude·Kimi가 동시에 연결돼 있어도 같은 문서를 동시에 편집하지 않습니다. 앞 작업이 끝난 뒤 문서를 다시 읽고 새 dry-run을 만듭니다.
- 상세 사용법은 패키지 내부 플러그인의 `INSTALL.md`와 `README.md`를 참고합니다.
