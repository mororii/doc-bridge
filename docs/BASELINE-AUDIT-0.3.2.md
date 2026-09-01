# DocBridge 0.3.2 HWP/CAD baseline audit

Audited: 2026-08-05

Baseline verification: `dotnet test DocBridge.sln -c Release --no-restore` passed 68/68 tests (51 core, 17 MCP). The green baseline proves the existing narrow surface, not the production target in `OFFICIAL-CAPABILITY-MATRIX.md`.

## HWP findings

- Six usable write operations: insert, whole/selection replace, find/replace, basic text/alignment style, and table insertion.
- `table_cell_set_text` is declared by the validator and appears in switch statements but intentionally fails; policy does not allow it. This is a misleading dead surface.
- Character formatting only applies bold, italic and size. Font, color, underline, strike, spacing and width are absent.
- Paragraph formatting only applies alignment. Page, section, columns, header/footer, numbering, images, captions, bookmarks and fields are absent.
- Table insertion is detailed but cell addressing and post-insertion editing are absent; merging is horizontal only.
- `GetTextFile("TEXT")` is repeatedly called during batches and style matching. There is no bounded structure inventory.
- Live snapshots preserve native HWP content; file-target snapshots copy the complete source file. This is a strong rollback base, but the host does not automatically invoke it when apply partially fails.

## CAD findings

- Query already supports document/file, filters, bounds, start/end index, result limit and optional geometry. It does not return a continuation index and scans from the requested start each time.
- Drawing supports lightweight polyline, circle, block, single-line text and hatch. Line, arc, ellipse, point, MText and dimensions are missing even though read serialization handles several of them.
- Modify supports move, rotate, text update and delete. Copy, scale, mirror, offset, common entity properties and block attributes are missing.
- Direct typed `CopyObjects` and cross-document fallback are present. This is the correct production direction.
- Layouts, paper-space viewports, plot configuration and plot-to-file are absent.
- Specialized flag/wall drawing operations are useful demonstrations but should not substitute for a complete primitive/modify surface.
- Context scans up to the global entity cap every call. Long drawings need a fast mode and explicit full-summary mode.

## Shared findings

- Policy allowlists and validator rules are manually duplicated from adapter switches and tool descriptions, allowing drift.
- Apply results have only a batch readback. There is no per-operation status/duration and no overall elapsed time.
- A failed `ApplyExecution` leaves partial changes unless the adapter itself cleans them up. The pre-apply snapshot is available but not automatically restored.
- MCP exposes 15 tools and no capabilities/preflight endpoint.
- The JSON result schema does not describe rollback or timing metadata.
- Existing E2E tests cover one HWP text flow and one small CAD flow. They do not cover the official production matrix.

## Implementation order

1. Runtime capabilities and shared operation catalog.
2. Host-level automatic rollback and timing envelope.
3. Bounded/continuable CAD query and optional expensive HWP metrics.
4. HWP basic/production operations.
5. CAD primitive/modify/layout/output operations.
6. Workflow skills, E2E probes, installer and three-client verification.

