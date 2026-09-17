# Excel formula trace

`excel_inspect` has an on-demand `formula_trace` scope for explaining a calculation from real Excel cell evidence. It does not scan the workbook or evaluate user formulas.

## Request

`workbook`, `sheet`, and `range` are required for this scope. `range` is a contiguous A1 cell or rectangle; an optional sheet qualifier must agree with `sheet`.

```json
{
  "scope": "formula_trace",
  "workbook": "C:\\work\\cost.xlsx",
  "sheet": "요약",
  "range": "F12",
  "maxDepth": 4,
  "maxCells": 500
}
```

The existing workbook routing remains in effect: an open workbook can be selected by name or absolute path. A closed file is read only when `allowOpenFile:true` is explicitly supplied with an existing absolute path.

## Result evidence

`nodes[]` contains `workbook`, `sheet`, `address`, `value`, `formula`, and `formulaRead`; formula nodes and displayed Excel errors also include `displayText`, and errors include `errorDisplay`. Excel-error classification calls native `Application.WorksheetFunction.IsError` with the Range object, rather than inferring an error from marshalled `Value2` or displayed text. If that native classification fails, `errorClassification` is `unknown` and coverage is incomplete. `formulaRead` says `formula2` when `Range.Formula2` succeeded. If it did not, the node records the explicit `Formula` fallback or an unavailable state.

`edges[]` are emitted only after the destination cells were actually read from the same workbook. Each edge has cell-addressed `from` and `to` objects plus the formula token that led to it. `cycles[]` reports verified graph cycles.

`coverage` reports the configured limits, returned node and edge counts, a `frontier`, and truncation reasons. `complete:false` means a limit or unresolved reference prevented complete evidence; it is not a claim that unseen cells are irrelevant.

## Resolution boundaries

Supported references are direct same-sheet and cross-sheet A1 cells/ranges, including `$`, quoted Korean or space-containing sheet names, and doubled apostrophes. Spill anchors such as `A1#`, `Data!A1#`, and `'원가 표'!A1#` are resolved with native `Range.SpillingToRange`, never `Evaluate`; the resulting same-workbook, single-area rectangle is read through the normal bounded evidence path. Whole columns and rows (`A:A`, `$B:$D`, `1:3`, `$1:$3`, including qualified sheets) expand to Excel's actual worksheet limits and are bounded by remaining `maxCells`; truncation/frontier therefore means unread cells remain. Workbook- and sheet-scoped names are resolved through Excel native `Names` / `RefersToRange`; names that are expressions/constants, multi-area ranges, absent, or external are reported in `unresolved`.

External workbooks are never followed. Structured table references find a named table across the selected workbook's worksheets or use the source cell's table for local tokens. Selection planning uses native header, data-body and totals bounds with actual column indexes. Plain columns exclude headers/totals; empty optional parts produce no cell dependencies. Reads are bounded before cell access. Supported selectors include `#Data`, `#Headers`, `#Totals`, `#All`, column ranges, part unions, `Table1[@Amount]`, `[@Amount]`, `[@[Column Name]]` and `[[#This Row],[Amount]]`. Current-row selection must lie within the data-body rows. Korean/spaced names and apostrophe-escaped special column characters are parsed without converting them into selector keywords. Malformed/ambiguous selectors and missing native table context remain explicitly unresolved.

`INDIRECT`, `OFFSET`, quoted or unquoted 3D ranges, broken `#REF!`, missing sheets, and failed native range resolution are explicit `unresolved` entries. Text inside string literals and numeric literals (including scientific notation) is skipped, and function identifiers such as `LOG10(` are not interpreted as A1 references. The trace does not infer why an Excel error occurred; it supplies the source values, formulas, and displayed error evidence for the connected AI to reason about it.

## Implementation status

Within one trace invocation, an explicit workbook-scoped read context caches cell evidence. Unread adjacent cells are grouped into rectangles; `Value2` and `Formula2` are read once per rectangle, with an explicit Formula fallback. A row-major prefix is bounded before allocating native arrays, and vertical runs are coalesced so a tall column is not fetched one row at a time. Error classification still uses real cell references, with shared native helpers per rectangle. Array/scalar shape mismatches mark evidence incomplete. The cache is discarded between invocations; it is not reused after writes or across workbooks.

`compact:true` is an optional output mode. It omits repeated per-node workbook paths when they equal the root workbook, and encodes edge/cycle/unresolved cell endpoints as quoted `'Sheet Name'!A1` strings. Nodes retain their sheet/address, values, formulas and error evidence; coverage and frontier are unchanged. `compact:false` (default) keeps the original object endpoints. Consumers must read the root workbook context before interpreting compact citations.

The current implementation has passed a Release build and code-level checks. Real Excel validation of the latest spill, table, whole-axis and error-classification changes is deferred at the user's request. This document describes implemented behavior, not a completed native acceptance or Claude parity claim.
