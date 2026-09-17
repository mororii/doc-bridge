# Excel data/reporting operations

This is the product contract for ListObject tables, sort/filter, defined names,
data validation, conditional formatting, native charts, sheet pictures, cell
notes, hyperlinks, and the workbook-level/advanced-object bundle
([EXCEL-WORKBOOK-OPS.md](EXCEL-WORKBOOK-OPS.md)). It is implemented in:

- `ExcelDataOperationsContract` / `ExcelDataObjectCatalog` / `ExcelWorkbookOpsContract`
- `ExcelAdapter.DataOperations.cs`, `ExcelAdapter.DataReaders.cs`, `ExcelAdapter.DataSnapshot.cs`
- `ExcelAdapter.WorkbookOps.cs`, `ExcelAdapter.WorkbookObjects.cs`, `ExcelAdapter.WorkbookOpsSnapshot.cs`

Shared wiring connects Preview/Apply, `OperationValidator`, policy, ToolRegistry,
`GetCapabilities`, and snapshot restoreMode to `excel_apply_ops`. Public contract
and fake-object checks have passed; native Excel execution remains deferred.

## Rules

- Direct Excel ActiveX COM only. No UI, VBA, openpyxl, or offline authoring.
- Sheet-scoped writes require `target.sheet` or a sheet-qualified A1 range.
  The six workbook-level ops (`update_external_links`, `change_link_source`,
  `break_external_link`, `set_calculation_mode`, `protect_workbook`,
  `unprotect_workbook`) bind the workbook directly and must not take a sheet.
- Active sheet is never assumed. Conflicting sheet names are rejected.
- Modifications use dry-run preview, scoped `data-objects` snapshot, confirm
  token, native-object apply, and readback. None of these ops are auto-execute.
- Data ops cannot share a batch with `set_values`, `format_range`, merge,
  visibility, or `copy_sheet`.
- Excel pictures use `insert_sheet_picture`, not HWP `insert_picture`.
- Unsupported rule/chart/filter tokens fail validation or throw
  `[EXCEL_DATA_UNSUPPORTED]`. Readback never reports success without reading
  the live COM object.

## Write ops

- [Rich text cells](EXCEL-RICH-TEXT.md)
- [Individual conditional rules](EXCEL-CONDITIONAL-RULES.md)
- [Table rows](EXCEL-TABLE-ROWS.md)
- [Connectors](EXCEL-CONNECTORS.md)
- [Chart details](EXCEL-CHART-DETAILS.md)
- [Picture editing](EXCEL-PICTURE-EDITING.md)
- [Workbook links, calculation, freeze/paste, goal seek, protection, split, sparklines, slicers, cell styles](EXCEL-WORKBOOK-OPS.md)

See `output/docbridge-production-20260910/cursor-data/OP-SCHEMA.md` for the
field table. Host scenario payloads for the materials ledger and monthly chart
report are in `ExcelDataOperationsHostScenarios` and
`output/docbridge-production-20260910/cursor-data/HOST-SCENARIOS.md`.

## Native read

Inspect calls `ExcelAdapter.ReadDataObjects(workbook, args)`
with `scope` = `all|tables|charts|pictures|names|validations|conditionalFormats|notes|hyperlinks|filters|pivots|shapes|connectors|richText|links|sparklines|slicers|cellStyles`.

## Restore

`restoreMode=data-objects`, `snapshotVersion=2`. Created tables/charts/pictures
and defined names are deleted. Sort restores formulas of the exact range.
Validation/CF/notes/hyperlinks restore the captured native state on the exact
target. Filter restore calls `ShowAllData` and only turns `AutoFilterMode` off
when the snapshot recorded the filter as absent. New kinds: `externalLinks`
(relink), `calculationMode`, `freezeValues`/`pasteSpecial` (surface rewrite),
`goalSeek`, `workbookProtection`, `splitPanes`, `sparkline`, `slicer`,
`cellStyle`. Update-links restore is verify-only because bindings are
unchanged; refreshed values depend on external files.
