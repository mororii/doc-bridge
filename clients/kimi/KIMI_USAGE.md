# Kimi에서 doc-bridge 사용하기

Kimi CLI(Kimi Code CLI)는 MCP 서버를 정식으로 지원한다. **1번(MCP 등록)이 기본**이고,
MCP를 붙일 수 없는 환경(웹 Kimi, 샌드박스 등)에서만 2번 CLI fallback을 쓴다.
둘 다 같은 `DocBridgeHost`를 거치므로 tool 이름·인자·안전 규칙이 완전히 같다.

---

## 1. MCP로 등록 (권장)

### 준비

```powershell
powershell.exe -NoProfile -ExecutionPolicy Bypass -File tools\publish.ps1 -SelfContained
```

### 등록 — 명령줄

```powershell
kimi mcp add --transport stdio doc-bridge -- "C:\Tools\DocBridge\dist\doc-bridge-mcp.exe" --stdio
```

### 등록 — 설정 파일 직접 편집

`~/.kimi/mcp.json` (프로젝트 한정이면 저장소의 `.kimi-code/mcp.json`):

```json
{
  "mcpServers": {
    "doc-bridge": {
      "command": "C:\\Path\\To\\DocBridge\\codex-marketplace\\plugins\\doc-bridge\\dist\\doc-bridge-mcp.exe",
      "args": ["--stdio"]
    }
  }
}
```

`clients/kimi/mcp.example.json`을 그대로 복사해도 된다.
`command`는 **절대경로**로 둔다 — PATH에 없기 때문이다.

### 확인

```powershell
kimi mcp list          # doc-bridge 가 보이는지
kimi mcp test doc-bridge
```

세션에서 `core_ping`을 호출해 `ok: true`와 `adapters: [excel, hwp, cad]`가 나오면 연결 완료다.

### 사용

읽기 tool은 그냥 호출하면 된다.

```
excel_get_active_context
excel_read_range   {"range":"A1:B10"}
hwp_read_text      {"file":"C:/path/문서.hwp","scope":"document"}
cad_get_active_context
cad_query_entities {"entityType":"Text","limit":20}
```

쓰기는 **반드시 두 번 호출**한다.

```
1) excel_apply_ops {"ops":[...], "dryRun":true}
   → diff, snapshotId, confirmToken 반환. 문서는 아직 안 바뀐다.
   → diff를 사용자에게 보여주고 승인을 받는다.

2) excel_apply_ops {"ops":[...동일...], "dryRun":false, "confirmToken":"conf_..."}
   → 적용 후 readback.verified 확인
```

---

## 2. CLI fallback (MCP를 못 붙일 때)

```powershell
$cli = (Resolve-Path ".\dist\doc-bridge-cli.exe").Path

# 상태 확인
& $cli core_ping
& $cli core_get_status

# 읽기
& $cli excel_get_active_context
@'
{
  "range": "A1:B10"
}'@ | Set-Content .\args.json -Encoding UTF8
& $cli excel_read_range --json-file .\args.json

@'
{
  "file": "C:/path/문서.hwp",
  "scope": "document"
}'@ | Set-Content .\args.json -Encoding UTF8
& $cli hwp_read_text --json-file .\args.json
```

쓰기 ops는 JSON 파일로 만든다 (배열 또는 `{ "ops": [...] }` 객체 모두 가능):

```json
[
  { "op": "find_replace", "target": { "scope": "workbook" },
    "find": "사과", "replace": "청사과", "options": { "matchCase": false } }
]
```

```powershell
# 1) dry-run: diff + confirmToken 발급 (문서 미변경, 스냅샷만 생성)
& $cli excel_apply_ops --ops ops.json --dry-run

# 2) 사용자가 승인하면 같은 ops 파일 + confirmToken 으로 apply
& $cli excel_apply_ops --ops ops.json --confirm-token conf_xxxx.yyyy

# 3) 되돌리기 (고위험: 2단계)
@'
{"snapshotId":"SNAP_ID"}
'@ | Set-Content restore-dry.json -Encoding UTF8
& $cli core_restore_snapshot --json-file restore-dry.json

# 첫 호출에서 받은 confirmToken을 아래 파일에 넣는다.
@'
{"snapshotId":"SNAP_ID","confirmToken":"conf_..."}
'@ | Set-Content restore-confirmed.json -Encoding UTF8
& $cli core_restore_snapshot --json-file restore-confirmed.json
```

출력 — stdout: 결과 JSON 한 줄 (`ok=true`면 exit 0, 아니면 1). stderr: 사용법/오류.

---

## 3. 사전 준비 (양쪽 공통)

- Excel / AutoCAD 는 사용자가 **문서를 직접 열어 둔다** (실행 중 인스턴스에 연결하는 방식).
- 한글(HWP)은 그 방식이 불가능하다 → 아래 절 참조.

---

## 4. 한글(HWP)은 파일 기반으로 작업한다

한글 COM 서버는 ROT에 등록되지 않고 COM 인스턴스가 헤드리스 전용이라,
Excel/AutoCAD처럼 "열어둔 창에 연결"이 안 된다. 대신 op에 `"file"` 인자로 파일을 지정하면
어댑터가 열기 → 수정 → 저장 → 닫기까지 수행한다.

```json
[
  { "op": "find_replace", "file": "C:/path/to/문서.hwp",
    "find": "청사과", "replace": "사과" }
]
```

apply가 끝나면 사용자가 한글에서 그 파일을 직접 열어 결과를 확인한다.
원본은 스냅샷(파일 백업)으로 보존되며 `core_restore_snapshot`으로 되돌릴 수 있다.

주의: `insert_text`의 `\n`은 한글이 제거한다(경고 표시됨).
Git Bash에서 CLI를 쓸 땐 JSON 안 경로를 `C:/path/...` 슬래시 형태로 쓰면 이스케이프 문제가 없다.

---

## 5. 규칙 요약

- `dryRun=false`는 confirmToken 없이 **항상 실패**한다.
- confirmToken은 5분 TTL, **1회용**, ops 내용에 바인딩된다 (ops를 바꾸면 무효).
- 고위험 op(`delete_entities`, `run_script_template`)는 `highRiskConfirm=true`
  (CLI는 `--high-risk-confirm`)가 추가로 필요하다.
- `run_script_template`은 repo `ops/script-templates/*.scr`에 등록된 템플릿만 실행된다.
  임의 스크립트/매크로는 차단.
- 모든 호출은 `%LOCALAPPDATA%\DocBridge\logs`에 JSONL로 기록된다.
- 한글 find_replace는 내부적으로 `ps.IgnoreMessage = 1`을 설정한다 (모달 대화상자로 인한 자동화 정지 방지).
- `hwp table_cell_set_text`는 미구현이라 allowlist에서 제거되어 정책 단계에서 거부된다.
- **AutoCAD(`acad.exe`)를 작업 관리자로 강제종료하지 말 것** — 라이센싱 구성요소가 손상되어
  이후 "라이센스 구성요소와 연결할 수 없습니다" 오류로 AutoCAD가 반복 종료될 수 있다.
