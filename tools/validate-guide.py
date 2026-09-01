#!/usr/bin/env python3
"""Structural checks for the offline beginner installation guide."""

from html.parser import HTMLParser
from pathlib import Path


class GuideParser(HTMLParser):
    def __init__(self) -> None:
        super().__init__()
        self.starts: dict[str, int] = {}

    def handle_starttag(self, tag: str, attrs: list[tuple[str, str | None]]) -> None:
        self.starts[tag] = self.starts.get(tag, 0) + 1


guide_candidates = list((Path(__file__).parent / "deployment").glob("*.html"))
assert len(guide_candidates) == 1, f"expected one HTML guide, found {len(guide_candidates)}"
guide = guide_candidates[0]
text = guide.read_text(encoding="utf-8")
parser = GuideParser()
parser.feed(text)
parser.close()

assert parser.starts.get("html") == 1
assert parser.starts.get("section", 0) >= 9
assert parser.starts.get("table", 0) >= 3
assert '<meta charset="utf-8">' in text
for required in (
    "0-VERIFY.cmd",
    "1-INSTALL.cmd",
    "2-TEST.cmd",
    "3-UNINSTALL.cmd",
    "%LOCALAPPDATA%\\DocBridge",
    "core_ping",
    "INSTALLATION SUCCESS",
    "TEST NOT STARTED",
    "[SKIP]",
    "Cursor global config",
    "%USERPROFILE%\\.cursor\\mcp.json",
    "docbridge-safe-automation.mdc",
    "detailLevel=basic",
):
    assert required in text, f"guide is missing: {required}"


repo_root = Path(__file__).resolve().parent.parent
cursor_docs = {
    "Cursor usage guide": repo_root / "clients" / "cursor" / "CURSOR_USAGE.md",
    "Cursor project rule": repo_root / "clients" / "cursor" / "rules" / "docbridge-safe-automation.mdc",
    "Cursor user rule": repo_root / "clients" / "cursor" / "docbridge-user-rule.txt",
    "document automation skill": repo_root / "skills" / "document-automation" / "SKILL.md",
}

excel_preflight_markers = (
    "core_get_status",
    "apps.excel.connected",
    "apps.excel.document",
    "excel_get_active_context",
    "absolute workbook path",
    "allowOpenFile:true",
    "excel_disconnect",
    "DocBridge가 생성한 인스턴스",
)

excel_no_bypass_markers = (
    "original DocBridge error",
    "openpyxl",
    "pywin32",
    "DispatchEx",
    "PowerShell Excel COM",
    "Start-Process",
    "UI automation",
    "overwrite",
)

for label, path in cursor_docs.items():
    doc_text = path.read_text(encoding="utf-8")
    for marker in excel_preflight_markers:
        assert marker in doc_text, f"{label} is missing Excel preflight marker: {marker}"
    for marker in excel_no_bypass_markers:
        assert marker.lower() in doc_text.lower(), (
            f"{label} is missing Excel no-bypass marker: {marker}"
        )
    assert doc_text.index("core_get_status") < doc_text.index("excel_get_active_context"), (
        f"{label} must check core_get_status before excel_get_active_context"
    )
    retry_language = doc_text.lower()
    assert "retry" in retry_language or "재시도" in doc_text, (
        f"{label} must state the no-repeat-retry rule"
    )
    allow_open_window = doc_text[doc_text.index("allowOpenFile") :]
    assert "false" in allow_open_window and "true" in allow_open_window, (
        f"{label} must keep allowOpenFile opt-in rather than default-on"
    )
    assert "explicit" in allow_open_window.lower() or "명시적" in allow_open_window, (
        f"{label} must require an explicit closed-file read request"
    )
    assert "write" in allow_open_window.lower() or "쓰기" in allow_open_window, (
        f"{label} must prohibit allowOpenFile for writes"
    )
    error_window = doc_text[doc_text.lower().index("original docbridge error") :]
    assert "state change" in error_window.lower() or "상태 변화" in error_window, (
        f"{label} must stop until an observable state change"
    )
    assert "explicit" in error_window.lower() or "명시적" in error_window, (
        f"{label} must require an explicit supported-route request"
    )

print(
    "Guide validation passed: "
    f"sections={parser.starts.get('section', 0)}, "
    f"tables={parser.starts.get('table', 0)}, characters={len(text)}, "
    f"cursorExcelPreflightDocs={len(cursor_docs)}"
)
