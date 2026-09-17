# DocBridge 0.4.23 릴리스 노트

기준일: 2026-09-17. 이전 공개 릴리스 v0.4.22 이후의 한글(HWP) 실무 기능 강화를 묶었다.

## 새로 추가된 HWP 쓰기 op 3종

- `insert_footnote` / `insert_endnote` — 주석 내용(`text` 필수) 입력. `target.text`가 있으면
  문서 처음부터의 첫 일치 문구에 달고(occurrence 미지원), 없으면 현재 커서/선택 위치에 단다.
  각주(`fn`)/미주(`en`) 컨트롤 수 +1로 검증하며 삽입 뒤 캐럿은 문서 시작으로 복귀한다.
- `table_set_repeat_header` — `tableIndex` + `repeat`(기본 true). 표 첫 행을 각 페이지에 반복한다.
  `HShapeObject.RepeatHeader`(ushort) 경로이며 적용 뒤 재선택·재조회로 검증한다.

## 기존 HWP op 확장

- `set_paragraph_style_basic` — `outline`, `shadow`+`shadowColor`, `emboss`/`engrave`(동시 true 거부),
  `smallCaps`, `kerning`.
- `set_paragraph_format` — `level` 0~9(0=본문, 개요 수준).
- `set_page_setup` — `lineNumbers`, `lineNumberStart`(1 이상).

신규 3종은 정책 `writeOps`에 등록(Allowed)했으며, 신규라 `autoExecuteOps`에는 넣지 않았다
(dry-run → confirmToken → apply 기본).

## 실측 근거 (로컬, 한글 2024 13.0.0.866)

- 통합 Release `-warnaserror` 경고 0·오류 0.
- COM-free Core 1012·MCP 20.
- 소유 탭 E2E 3종: 글자 효과·개요 수준·각주/미주(fn/en +1)·표제행 off→on 토글·줄번호 시작값,
  제품 경로 dry-run → apply → readback + 원시 재조회.
- 형광펜/다단/글머리·번호 적용/캡션/텍스트박스는 Execute false, 수식은 내용 입력 계약이 없어
  공개 명령에서 제외했다(HWP-OPERATIONS.md 기록).

## 이전 버전 대비

- Excel 16 ops + 4 scopes + veryHidden, instance pin, 테두리 순서 가드는 0.4.22 그대로 포함.
