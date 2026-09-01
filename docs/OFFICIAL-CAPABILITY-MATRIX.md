# DocBridge Excel/HWP/CAD capability matrix

This matrix turns official beginner training topics and the official automation object models into a bounded implementation contract. It deliberately favors direct HWP Automation and AutoCAD ActiveX calls. UI driving and arbitrary AutoLISP/macros are outside the normal path.

## Design levels

- **Basic**: the operations a new user needs to create and edit a normal document/drawing.
- **Production**: reusable formatting, annotation, blocks, layouts, exports, and structured reads.
- **Workflow**: multi-step work such as formatted reports and longitudinal/profile sheet assembly.
- **Reliability**: capability preflight, bounded queries, dry-run snapshot, confirmation, readback, automatic rollback, and timings.

## Excel

| Area | Official automation basis | Existing baseline | Target in this work |
| --- | --- | --- | --- |
| Connection/workbook | `Application`, ROT, `Workbooks`, `Workbook`, `Worksheets` | existing-instance-first connection, explicit workbook/sheet targeting, safe disconnect | retain user-owned Excel, clean up only DocBridge-owned instances, expose connection type and supported ops |
| Values/formulas | `Range.Value2`, `Range.Formula` | bounded range read, values/formulas, find/replace, formatting | preserve formulas and numeric COM types; keep explicit sheet identity and exact readback |
| Cell structure | [`Range.Merge`](https://learn.microsoft.com/en-us/office/vba/api/excel.range.merge), [`Range.UnMerge`](https://learn.microsoft.com/en-us/office/vba/api/excel.range.unmerge), [`Range.MergeArea`](https://learn.microsoft.com/en-us/office/vba/api/excel.range.mergearea) | row/column insert and sheet copy | delivered: loss-blocking rectangular merge, bounded unmerge, operation-scoped merge snapshot and readback |
| Visibility | [`Range.Hidden`](https://learn.microsoft.com/en-us/office/vba/api/excel.range.hidden), [`Worksheet.Visible`](https://learn.microsoft.com/en-us/office/vba/api/excel.worksheet.visible), [`XlSheetVisibility`](https://learn.microsoft.com/en-us/office/vba/api/excel.xlsheetvisibility) | sheet visibility was inspect-only | delivered: row/column hide and unhide, normal sheet hide/show, last-visible/active-sheet protection, exact mixed-state rollback |
| Inspection | `UsedRange`, `MergeCells`, `Hidden`, `Visible`, workbook collections | scan, objects, formula errors, diagnostics | delivered: `includeLayout` returns merged areas, row/column hidden states, sheet visibility and bounded coverage |
| Next basic formatting | `RowHeight`, `ColumnWidth`, `AutoFit`, alignment, borders | partial `format_range` | planned only after per-property snapshot/readback and real-Excel verification |
| Production data features | `ListObject`, validation, names, conditional formats, filters, page setup | inspection only or absent | staged follow-up; do not advertise as writable until policy, rollback and E2E gates pass |

The detailed operation payloads, batching restrictions and follow-up matrix are in
[Excel basic editing operations](EXCEL-OPERATIONS.md).

## HWP

| Area | Official automation basis | Baseline 0.3.2 | Target in this work |
| --- | --- | --- | --- |
| Connection/document | `IHwpObject`, `XHwpDocuments`, `Open`, `Save`, `SaveAs`, ROT | Connect existing window; optional file open/save | document info, create/open/activate/save/export with explicit lifecycle |
| Text/search | `InsertText`, `FindReplace`, `GetSelectedPosBySet`, `SetPosBySet`, selection/move actions | insert, replace whole/selection, all replace | delivered: Unicode entity normalization, occurrence/document/paragraph/table-cell scoped search, page/paragraph/section/column breaks, exact before/after anchor insertion with adjacent style inheritance, existing template-field read/write |
| Character format | `CharShape` parameter set | bold, italic, size | font family, text color, underline/strike, spacing, width, superscript/subscript |
| Paragraph format | `ParaShape` and paragraph alignment actions | alignment only | margins, indentation, spacing, line spacing, keep/widow/orphan controls, named style application |
| Page/section | `PageSetup`, `SecDef`, break actions, columns | absent | page size/orientation/margins, page and section breaks, columns, page numbering |
| Tables | `TableCreate`, `CellBorderFill`, cell block/move/merge actions | insert, per-cell style, horizontal merge | delivered: exact cell replace, up to 500-cell batch replace, stable large-table inventory, exact-count row/column insert/delete, rectangular merge, formula-cell protection; split withheld because HWP 2024 opens a hidden modal |
| Objects | controls linked list, `ShapeObject`, image and header/footer actions | control count only | delivered: bounded control inventory, image insertion, page number, header/footer; modal-prone bookmark/hyperlink/new-field insertion withheld |
| Inspection | `GetTextFile`, control list, document properties | plain text/selection | pagination-safe document metrics and bounded control/table inventory |

Performance rule from the official object model: `PageCount` forces full pagination and can be very slow when repeatedly queried on long documents. DocBridge therefore keeps page-count inspection opt-in and never polls it during a batch.

## AutoCAD

| Area | Official ActiveX basis | Baseline 0.3.2 | Target in this work |
| --- | --- | --- | --- |
| Document/query | `Application.Documents`, `Document`, model/paper spaces | launch, activate, bounded model-space query | delivered: explicit Save/SaveAs, layouts/viewports, continuation cursor, up to 100 region checks in one scan |
| Geometry | `AddLine`, `AddLightWeightPolyline`, `AddCircle`, `AddHatch` | polyline, circle, block, text, hatch | delivered: line, arc, ellipse, point, MText, aligned/rotated dimensions; hatch uses typed AcadEntity[] without LISP |
| Modify | entity `Move`, `Rotate`, `Delete`; `CopyObjects` | move, rotate, text, delete, cross-document copy | delivered: copy vector, scale, mirror, offset, common property updates, block attribute edits |
| Organization | Layers, blocks, xrefs | layer on/color, insert block/xref | create/update layers, linetype/lineweight, block attributes, xref reload/detach/bind where ActiveX supports it |
| Layout/output | Layout, Block, PaperSpace, `AddPViewport`, Plot | absent | delivered and E2E-tested: layout create/activate, paper-space viewport, synchronous PDF plot, SaveAs |
| Production workflow | direct object selection/copy/transform | profile-sheet skill exists | deterministic sheet analyzer, frame-centering checks, title/keymap checks, resumable batches |

## Reliability contract

Every write path must satisfy all of the following:

1. `core_get_capabilities` reports support, availability, limits, and direct-COM status before a complex workflow starts.
2. Dry-run validates every operation and creates the full snapshot before returning a confirmation token.
3. Apply is bound to the same operation hash and document identity.
4. Each operation contributes an operation-level status and duration.
5. Any failed batch triggers best-effort automatic restoration from its pre-apply snapshot; the result reports whether rollback was verified.
6. Reads are bounded. CAD scans accept start/end index and return a continuation index; expensive HWP pagination is opt-in.
7. High-risk delete/script/overwrite/export actions remain policy-gated.

## Acceptance workflows

- Excel basic: read values/formulas plus `includeLayout`, merge and unmerge a safe range, hide/show rows and columns, hide/show a non-active sheet, then verify exact layout readback and snapshot restoration.
- Excel reliability: refuse content-losing merge, partial-overlap unmerge, active/last-visible sheet hiding and protected-structure changes; disconnect without leaving a DocBridge-owned `EXCEL.EXE` process.
- HWP basic: create a document, insert/format paragraphs, configure page, insert/edit/merge a table, save and reopen, verify text and control inventory.
- HWP production: build a styled daily plan with header/footer, page number, image and table; export PDF; verify text, controls and output existence.
- CAD basic: create all supported entity types on controlled layers, modify them, read them back by handle/bounds/type, save and reopen.
- CAD production: create a layout and viewport, update block attributes, plot a PDF, then verify layout, viewport and output file.
- CAD profile workflow: analyze source/target drawings, copy each longitudinal profile and matching plan, center inside frames, update title and keymap, and verify every requested line with no out-of-frame entities.
