#!/usr/bin/env python3
"""Reject credentials and local personal data from tracked public source files."""

from __future__ import annotations

import pathlib
import re
import subprocess
import sys


ROOT = pathlib.Path(__file__).resolve().parents[1]
BINARY_SUFFIXES = {
    ".7z",
    ".dll",
    ".docx",
    ".dwg",
    ".dxf",
    ".exe",
    ".gif",
    ".ico",
    ".jpeg",
    ".jpg",
    ".nupkg",
    ".pdf",
    ".pdb",
    ".png",
    ".pyd",
    ".so",
    ".ttf",
    ".woff",
    ".woff2",
    ".xlsx",
    ".zip",
}

RULES = {
    "Windows user home path": re.compile(
        r"[A-Za-z]:[\\/]+Users[\\/]+[^\\/\s\"']+[\\/]+(?:Desktop|AppData|Documents)",
        re.IGNORECASE,
    ),
    "local storage drive path": re.compile(
        r"[A-Za-z]:[\\/]+(?:HDD|SSD)\d+(?:[\\/]|$)", re.IGNORECASE
    ),
    "GitHub token": re.compile(r"(?:github_pat_|gh[pousr]_)[A-Za-z0-9_]{20,}"),
    "OpenAI or Anthropic key": re.compile(r"(?:sk-ant-|sk-)[A-Za-z0-9_-]{20,}"),
    "Google API key": re.compile(r"AIza[0-9A-Za-z_-]{30,}"),
    "AWS access key": re.compile(r"AKIA[0-9A-Z]{16}"),
    "Slack token": re.compile(r"xox[baprs]-[0-9A-Za-z-]{10,}"),
    "private key": re.compile(
        r"-----BEGIN (?:RSA |EC |OPENSSH |DSA )?PRIVATE KEY-----"
    ),
    "Korean mobile number": re.compile(
        r"(?<!\d)01[016789][ -]?\d{3,4}[ -]?\d{4}(?!\d)"
    ),
    "Korean resident registration number": re.compile(
        r"(?<!\d)\d{6}-[1-4]\d{6}(?!\d)"
    ),
}


def tracked_files() -> list[pathlib.Path]:
    result = subprocess.run(
        ["git", "ls-files", "-z"],
        cwd=ROOT,
        check=True,
        stdout=subprocess.PIPE,
    )
    return [
        ROOT / entry.decode("utf-8", errors="strict")
        for entry in result.stdout.split(b"\0")
        if entry
    ]


def main() -> int:
    failures: list[str] = []
    scanned = 0
    for path in tracked_files():
        if path.suffix.lower() in BINARY_SUFFIXES:
            continue
        try:
            text = path.read_text(encoding="utf-8")
        except (UnicodeDecodeError, OSError):
            continue
        scanned += 1
        relative = path.relative_to(ROOT).as_posix()
        for line_number, line in enumerate(text.splitlines(), start=1):
            for label, pattern in RULES.items():
                if pattern.search(line):
                    failures.append(f"{relative}:{line_number}: {label}")

    if failures:
        for failure in failures:
            print(failure, file=sys.stderr)
        print(
            f"Public source safety failed with {len(failures)} finding(s).",
            file=sys.stderr,
        )
        return 1

    print(f"Public source safety passed for {scanned} tracked text files.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
