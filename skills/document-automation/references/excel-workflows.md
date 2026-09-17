# Excel business workflows

Use the installed tool schemas for exact operation fields and supported scopes. These are planning recipes over existing DocBridge operations, not evidence that a native workflow has passed. Respect an explicit deferral of real application execution. The main skill's workbook targeting and execution contracts apply to every recipe.

## Work in the user's designated window

When several Excels are open, never guess. Read `excel_get_active_context` (`openWorkbooks[].processId`/`excelHwnd`), ask which window the user means if it is ambiguous, and pin it with `excel_launch` (`processId`/`hwnd` pair or `activeWindow:true`). Afterwards every call sticks to that window; a vanished pin fails closed with `[EXCEL_PINNED_INSTANCE_GONE]`. Unpin with `excel_launch` `clearPin` when done. Never move, resize, minimize, or activate user windows to force targeting.

## Explain or change a calculation

Start with the requested output cells: `excel_inspect` / `formula_trace`, explicit workbook/sheet/range, bounded depth/cells, `compact:true` when exposed. Use the returned values and dependencies to identify relevant inputs; expand only material unresolved frontiers. Do not infer a denominator error from a zero in an unrelated cell or treat text beginning with `#` as an Excel error.

For an authorized input change, write the specified inputs with `set_values`, and use `set_formulas` only where a formula change is intended. Request Formula2 for dynamic-array formulas when supported. Use `calculate` when recalculation is part of the task, keeping it on its supported execution path. Read the specific outputs and formula footprints needed to check the requested result. Do not scan all sheets before or after a small change.

## Clean imported data

Use `import_csv` with explicit target/path/format fields from the installed schema. Preserve identifiers and leading-zero text unless the user requests conversion. Apply the requested `remove_duplicates`, `text_to_columns`, `sort_range`/`sort_table`, or filter operations to the actual imported range. Keep column headers, key columns and deduplication criteria explicit. Compare row counts and selected key/value examples; a displayed sample does not prove all rows were processed.

## Extend a table and update a pivot

Inspect the named table and pivot metadata in their available object scopes. Use stable names and explicit workbook/sheet targets. Extend the existing table with `resize_table`, write the added data, and update calculated columns only when required. Inspect/adjust the pivot source with `update_pivot` when it still uses a fixed old range; then `refresh_pivot`. A refresh of an unchanged old source does not include the new rows. Check the requested aggregate and source coverage rather than rereading every raw cell.

## Create or revise a chart report

Read the relevant data range and existing chart metadata first. Use `create_chart` for a new chart and `update_chart` for the requested changes to an existing chart. Make category/value ranges, series, chart type, placement and titles explicit from the task/schema. Keep unrelated report objects in place. Separate evidence of correct source ranges/series from visual layout acceptance.

## Refresh a linked workbook bridge (quantity pattern)

Inspect `links` scope first to list sources and statuses; unreachable links are refused, never retried blindly. Point copies at local files with `change_link_source` (absolute existing path), refresh cached values with `update_external_links` (optionally filtered), then set `set_calculation_mode` manual + `fullRebuild` before reading totals. Freeze handoff values with `freeze_values` only in value-fixed outputs; keep link-preserving outputs formula-live and separate. Breaking a link (`break_external_link`, high-risk) converts dependents to cached values with a captured relink path. Do not replace formulas with hardcoded zeros and do not reserialize package XML; calcChain and namespace preservation stay with the existing quantity tools.

## Order border writes around shared edges

Adjacent ranges share one edge object: clearing `A16:E18` also clears the left edge of an `F16:H18` outline drawn earlier, and readback only verifies each batch's own ops, so cross-batch interference passes silently. Draw outlines after adjacent clears in the same batch (file order), and after any later border/clear batch re-verify previously drawn perimeters edge-by-edge (LineStyle/Weight over live COM or a cropped PDF render), not just the batch's own readback.

## Minimize calls without losing evidence

Group dependent edits when the supported batch contract permits them; ordinary execute mode supports only its advertised allowlist. Other operations use the existing preview/confirmation-token flow, without inventing a new user approval step when authorization already exists. Reuse a fresh targeted read within a task; invalidate it after relevant writes. Track requested versus returned coverage and report unsupported operations or deferred acceptance explicitly.
