# Excel workbook-level and advanced-object operations

Product contract for the 2026-09-16 Excel reinforcement bundle (VBA/macros
excluded). All ops below are `excel_apply_ops` data/reporting ops: dry-run
preview, `data-objects` snapshot, confirm token, native COM apply, readback.
`core_get_capabilities({"app":"excel"})` `writeOps` is the authoritative list.

Implemented in `ExcelWorkbookOpsContract` (COM-free validation),
`ExcelAdapter.WorkbookOps.cs`, `ExcelAdapter.WorkbookObjects.cs`, and
`ExcelAdapter.WorkbookOpsSnapshot.cs`.

## External links (quantity-bridge core)

| op | effect |
| --- | --- |
| `update_external_links` | `Workbook.UpdateLink` for all (or `source`/`sourceContains`-filtered) Excel links. Bindings unchanged. |
| `change_link_source` | `Workbook.ChangeLink` old to new. `newSource` must be an absolute existing path. |
| `break_external_link` | `Workbook.BreakLink` (high-risk). Dependents keep cached values. |

- `source` accepts the full link path or a unique leaf/filename; ambiguous
  input is rejected with candidates instead of guessing.
- Links whose status is missing-file/missing-sheet/invalid-name are refused
  for update (`[EXCEL_LINK_UNREACHABLE]`); a modal prompt is never awaited.
- Break/change capture dependent formulas (bounded, 20,000 cells, fail-closed)
  so restore can relink: change re-points at the old source, break rewrites
  the captured formulas and recalculates.
- Update does not mutate bindings, so its restore only re-verifies the
  inventory. Refreshed values depend on the external files and are read back,
  not rolled back.
- New read scope: `links` (`source`, `leaf`, `status`, `statusName`).

## Calculation

`set_calculation_mode` sets `Application.Calculation` to
`manual|automatic|semiautomatic` with optional `recalculate`
(`none|full|fullRebuild`). `fullRebuild` matches the quantity-engine refresh
semantics (`CalculateFullRebuild`). The snapshot stores the previous mode and
restore sets it back. Readback checks the mode and `CalculationState=xlDone`.

## Freeze and paste-special (no clipboard)

- `freeze_values` replaces formulas with cached values in place (formats
  kept). Error cells cannot be written as values, so they keep their formulas
  and are listed in warnings; merged ranges are refused. At most 20,000 cells.
  Restore rewrites the captured formulas.
- `paste_special` transfers `sourceRange` to `destination` without touching
  the system clipboard: `values|formulas` by direct array transfer,
  `all|formats` by native `Range.Copy` (borders included), then values/formulas
  are rewritten from the pre-state for `formats`. Options: `transpose`
  (values/formulas only), `skipBlanks`, `operation`
  (`none|add|subtract|multiply|divide`, values/all only). A single-cell
  destination expands to the source shape; larger destinations must fit the
  (transposed) source. Readback verifies values, formulas, and number formats.

## Goal seek

`goal_seek` runs `Range.GoalSeek` on a formula `goalCell` driven by changing
`cell` (same worksheet, enforced). `goal` is finite; optional relative
`tolerance` (default 1e-4, floor 1e-9) bounds `|actual - goal|`. Both cells are
snapshotted; restore rewrites the goal surface and the changing cell.

## Workbook protection and views

- `protect_workbook` / `unprotect_workbook` manage structure/windows
  protection without passwords. Password workbooks are refused
  (`[EXCEL_WORKBOOK_PASSWORD]`), consistent with `protect_sheet` scope.
- `set_split_panes` sets window Split or FreezePanes anchored at `cell`
  (rows above / columns left stay fixed) or explicit
  `splitRows`/`splitColumns`; `remove` clears both. Hidden sheets are refused.
  The previous window (active sheet, selection, scroll) is restored after
  apply, snapshotted, and compared on readback.
- `set_sheet_visibility` now accepts `veryHidden` in addition to
  `visible`/`hidden`. Last-visible and active-sheet guards apply to both
  hidden states; unhiding a `veryHidden` sheet uses `visible`.

## Sparklines, slicers, cell styles

- `create_sparkline` / `update_sparkline` / `delete_sparkline` manage native
  `SparklineGroups` on one cell, row, or column. `type` is `line|column`;
  `winLoss` is rejected until a distinct native mapping is proven (same
  fail-closed policy as fixed-width text-to-columns). Options: `sourceData`
  (may be another sheet of the same workbook), `markers`, `lineColor`,
  `showHigh|showLow|showNegative|showFirst|showLast`. Type changes recreate the
  group. Groups are enumerated through the sheet-wide `Cells` range because
  `Worksheet.SparklineGroups` can report empty while live groups exist
  (observed on Excel 16.0). New read scope: `sparklines`.
- `create_slicer` / `delete_slicer` manage native slicers from a ListObject or
  PivotTable field. `name` is required (stable identity for delete/verify/
  restore) and must be unique workbook-wide, as must the cache name. The cache
  is removed when its last slicer goes. `delete_slicer` accepts optional
  `source`+`field` for exact restore; otherwise the slicer caption is used as
  the field. Slicers are resolved through `SlicerCaches` because
  `Worksheet.Slicers` can report empty while the slicer shape is live
  (observed on Excel 16.0). New read scope: `slicers`.
- `apply_cell_style` applies an existing workbook style (see `cellStyles`
  scope) to at most 5,000 cells. Previous per-cell style names are captured
  (flat row-major grid) and restored with run-length-encoded writes. New read
  scope: `cellStyles`.

## Limits and non-goals

- VBA, macros, form/ActiveX controls: out of scope (`run_macro` forbidden).
- Power Query / external data connections (beyond link refresh), timelines,
  scenarios/data-tables, theme editing, header/footer pictures: not in this
  bundle; tracked as follow-ups in EXCEL-PRACTICAL-SCOPE.md.
- `winLoss` sparklines, fixed-width text-to-columns: rejected until native
  mappings are proven.
- Password-protected workbooks: refused by protection ops.
- Verification: Release build warnings/errors 0, Core 973 + MCP 20 green
  (contract, schema, and fake-object tests). Native plan-document test below
  passed. Local install reflection is pending.

## Native plan-document test (2026-09-17, isolated)

Built `주간작업계획서_테스트.xlsx` through create, input, recalculation,
save, reopen, and PDF using real Excel COM: table, formulas, conditional
formatting, cell style, sparkline, calculation mode, freeze panes, page setup,
paste-special, freeze-values, goal seek (J1=10 reaching J3=11000), slicer
create/delete, split panes, veryHidden hide/show, workbook protect/unprotect,
save, PDF, close, and reopen — every apply readback-verified, with totals
435/195 and formulas intact after reopen. Batch-mixing violations are refused
at dry-run as designed. Bugs found and fixed by this test: re-adding parented
nodes in inspect readers, `Protect(Password, Structure, Windows)` argument
order, missing `MergeCells`/`BuiltIn`/`ProtectStructure`/`ProtectWindows` in
`TryComBool`, table style names read as raw COM objects, and the empty
`Worksheet.Slicers`/`SparklineGroups` collection traps.
