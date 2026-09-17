namespace DocBridge.Development.ExcelProductionWorkflowProbe;

/// <summary>
/// Remaining advertised-op cycles. Create/update/delete of the same object
/// are separate batches. Fixture E3/E6/E8/E9 geometry is spliced by
/// ScenarioRunner; E9 import chain stays in the fixtures. E2 resume IDs stay.
/// </summary>
internal static class RemainingOpWorkflows
{
    public static IEnumerable<PlannedBatch> E3(string sheet, string pictureDir)
    {
        var scratchPng = TinyPng.WriteSample(pictureDir, "e3-scratch.png", 80, 80, 80);
        yield return ApplyPlanner.Batch("e3", "e3-scratch-picture-insert", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "insert_sheet_picture",
            ["target"] = JsonUtil.Target(sheet),
            ["path"] = scratchPng,
            ["name"] = "ScratchPhoto",
            ["position"] = new JsonObject { ["left"] = 420, ["top"] = 20, ["width"] = 36, ["height"] = 36 },
        }), "token", "Photo1/Photo2 stay; scratch is extra");
        yield return ApplyPlanner.Batch("e3", "e3-scratch-picture-update", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "update_picture",
            ["target"] = JsonUtil.Target(sheet),
            ["name"] = "ScratchPhoto",
            ["position"] = new JsonObject { ["left"] = 424, ["top"] = 24, ["width"] = 32, ["height"] = 32 },
        }), "token", "readback ScratchPhoto geometry before delete", capture: "object-evidence");
        yield return ApplyPlanner.Batch("e3", "e3-scratch-picture-delete", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "delete_picture",
            ["target"] = JsonUtil.Target(sheet),
            ["name"] = "ScratchPhoto",
            ["path"] = scratchPng,
        }), "token", "final pictures remain Photo1 and Photo2");
        yield return ApplyPlanner.Batch("e3", "e3-shape-insert", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "insert_shape",
            ["target"] = JsonUtil.Target(sheet),
            ["shapeType"] = "rectangle",
            ["name"] = "ScratchBox",
            ["fillColor"] = "#D6DCE4",
            ["position"] = new JsonObject { ["left"] = 420, ["top"] = 60, ["width"] = 40, ["height"] = 16 },
        }), "token");
        yield return ApplyPlanner.Batch("e3", "e3-shape-update", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "update_shape",
            ["target"] = JsonUtil.Target(sheet),
            ["name"] = "ScratchBox",
            ["text"] = "스크래치",
        }), "token", "keep ScratchBox text evidence before delete", capture: "object-evidence");
        yield return ApplyPlanner.Batch("e3", "e3-shape-delete", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "delete_shape",
            ["target"] = JsonUtil.Target(sheet),
            ["name"] = "ScratchBox",
        }), "token");
        yield return ApplyPlanner.Batch("e3", "e3-textbox-insert", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "insert_textbox",
            ["target"] = JsonUtil.Target(sheet),
            ["name"] = "ScratchText",
            ["text"] = "메모",
            ["position"] = new JsonObject { ["left"] = 420, ["top"] = 80, ["width"] = 60, ["height"] = 16 },
        }), "token");
        yield return ApplyPlanner.Batch("e3", "e3-textbox-update", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "update_textbox",
            ["target"] = JsonUtil.Target(sheet),
            ["name"] = "ScratchText",
            ["text"] = "메모2",
        }), "token", "keep ScratchText evidence before delete", capture: "object-evidence");
        yield return ApplyPlanner.Batch("e3", "e3-textbox-delete", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "delete_textbox",
            ["target"] = JsonUtil.Target(sheet),
            ["name"] = "ScratchText",
        }), "token");
    }

    public static IEnumerable<PlannedBatch> E4(string sheet, string xlsx)
    {
        _ = xlsx;
        const string revision = "개정";
        yield return ApplyPlanner.Batch("e4", "e4-table-style", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "style_table",
            ["target"] = JsonUtil.Target(sheet),
            ["name"] = "자재표",
            ["showRowStripes"] = true,
            ["showAutoFilter"] = true,
        }), "token");
        yield return ApplyPlanner.Batch("e4", "e4-table-totals", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_table_totals",
            ["target"] = JsonUtil.Target(sheet),
            ["name"] = "자재표",
            ["showTotals"] = true,
            ["columns"] = JsonUtil.Arr(
                new JsonObject { ["column"] = "입고", ["function"] = "sum" },
                new JsonObject { ["column"] = "출고", ["function"] = "sum" }),
        }), "token", "입고 200 / 출고 113 after TX-004; void row excluded");
        yield return ApplyPlanner.Batch("e4", "e4-add-calc-col", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "add_table_column",
            ["target"] = JsonUtil.Target(sheet),
            ["name"] = "자재표",
            ["columnName"] = "순입고",
            ["formula"] = "=[@입고]-[@출고]",
        }), "token", "expected J3:J6 = 80,7,−15,15 matching G remain formulas", capture: "read:자재대장!A2:J8");
        yield return ApplyPlanner.Batch("e4", "e4-table-totals-net", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_table_totals",
            ["target"] = JsonUtil.Target(sheet),
            ["name"] = "자재표",
            ["showTotals"] = true,
            ["columns"] = JsonUtil.Arr(
                new JsonObject { ["column"] = "입고", ["function"] = "sum" },
                new JsonObject { ["column"] = "출고", ["function"] = "sum" },
                new JsonObject { ["column"] = "순입고", ["function"] = "sum" }),
        }), "token", "입고 200 / 출고 113 / 순입고 87 after TX-004");
        yield return ApplyPlanner.Batch("e4", "e4-note-set", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_cell_note",
            ["target"] = JsonUtil.Target(sheet),
            ["cell"] = "C4",
            ["text"] = "현장 검수 후 입고 확정",
        }), "token", capture: "object-evidence");
        yield return ApplyPlanner.Batch("e4", "e4-add-revision-sheet", OpFamily.Structure, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "add_sheet",
            ["name"] = revision,
            ["afterSheet"] = sheet,
        }), "token", "owned copy sheet; 자재표 name stays on 자재대장");
        yield return ApplyPlanner.Batch("e4", "e4-copy-revision-range", OpFamily.RangeEdit, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "copy_range",
            ["target"] = JsonUtil.Target(sheet),
            ["range"] = "A1:I12",
            ["destRange"] = $"{revision}!A1",
            ["mode"] = "formulas",
        }), "token", "populated ledger copy A1:I12; VOID stays off the table on both sheets");
        yield return ApplyPlanner.Batch("e4", "e4-rev-insert-line", OpFamily.Structure, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "insert_rows",
            ["target"] = JsonUtil.Target(revision),
            ["row"] = 5,
            ["count"] = 1,
        }), "token", "insert inside populated TX-001..TX-004; TX-003 shifts to row 6");
        yield return ApplyPlanner.Batch("e4", "e4-rev-fill-line", OpFamily.Values, JsonUtil.Arr(
            new JsonObject
            {
                ["op"] = "set_values",
                ["target"] = JsonUtil.Target(revision),
                ["range"] = "A5:I5",
                ["values"] = Arr(Arr("TX-005", "2026-09-06", "PVC관", "D300", 30, 8, null, "최감독", "개정 추가")),
            },
            new JsonObject
            {
                ["op"] = "set_formulas",
                ["target"] = JsonUtil.Target(revision),
                ["range"] = "G5",
                ["formulas"] = Arr(Arr("=E5-F5")),
            }), "execute", "new line remain 22; independent insert total 입고 230 / 출고 121");
        yield return ApplyPlanner.Batch("e4", "e4-rev-calc", OpFamily.Values, JsonUtil.Arr(
            new JsonObject { ["op"] = "calculate" }), "token", "G5=22 before deleting obsolete TX-003", capture: "read:개정!A1:I8");
        yield return ApplyPlanner.Batch("e4", "e4-find-replace", OpFamily.RangeEdit, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "find_replace",
            ["target"] = new JsonObject { ["sheet"] = revision, ["scope"] = "sheet" },
            ["find"] = "야간",
            ["replace"] = "야간작업",
        }), "token", "TX-003 비고 still on the copy at row 6");
        yield return ApplyPlanner.Batch("e4", "e4-rev-delete-obsolete", OpFamily.Delete, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "delete_rows",
            ["target"] = JsonUtil.Target(revision),
            ["row"] = 6,
            ["count"] = 1,
        }), "token", "remove populated TX-003; TX-004 shifts to row 6; 자재대장 TX-003 stays");
        yield return ApplyPlanner.Batch("e4", "e4-rev-sort", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "sort_range",
            ["target"] = JsonUtil.Target(revision),
            ["range"] = "A2:I6",
            ["hasHeaders"] = true,
            ["keys"] = JsonUtil.Arr(
                new JsonObject { ["column"] = "C", ["order"] = "asc" },
                new JsonObject { ["column"] = "A", ["order"] = "asc" }),
        }), "token", "two keys 품목 then 거래ID; PVC관 is a duplicate primary key", capture: "read:개정!A1:I8");
        yield return ApplyPlanner.Batch("e4", "e4-rev-aux-col", OpFamily.Structure, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "insert_cols",
            ["target"] = JsonUtil.Target(revision),
            ["col"] = "J",
            ["count"] = 1,
        }), "token", "auxiliary remain mirror on the populated copy");
        yield return ApplyPlanner.Batch("e4", "e4-rev-aux-formula", OpFamily.Values, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_formulas",
            ["target"] = JsonUtil.Target(revision),
            ["range"] = "J3:J6",
            ["formulas"] = Arr(Arr("=G3"), Arr("=G4"), Arr("=G5"), Arr("=G6")),
        }), "execute", "J mirrors remain after sort; expected 7,80,15,22");
        yield return ApplyPlanner.Batch("e4", "e4-rev-aux-delete", OpFamily.Delete, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "delete_cols",
            ["target"] = JsonUtil.Target(revision),
            ["col"] = "J",
            ["count"] = 1,
        }), "token", "copy keeps TX-001/002/004/005 associations after the aux cycle");
        yield return ApplyPlanner.Batch("e4", "e4-fill", OpFamily.RangeEdit, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "fill_range",
            ["target"] = JsonUtil.Target(revision),
            ["range"] = "H6",
            ["value"] = "박하린",
        }), "token", "TX-004 담당 after obsolete-row delete");
        yield return ApplyPlanner.Batch("e4", "e4-auto-fill", OpFamily.RangeEdit, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "auto_fill",
            ["target"] = JsonUtil.Target(revision),
            ["range"] = "G3:G4",
            ["destRange"] = "G3:G6",
        }), "token", "continue remain formulas on the four remaining copy rows");
        yield return ApplyPlanner.Batch("e4", "e4-packed-values", OpFamily.Values, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_values",
            ["target"] = JsonUtil.Target(revision),
            ["range"] = "L3",
            ["values"] = Arr(Arr("PVC|D300")),
        }), "execute");
        yield return ApplyPlanner.Batch("e4", "e4-text-to-columns", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "text_to_columns",
            ["target"] = JsonUtil.Target(revision),
            ["range"] = "L3",
            ["destination"] = "M3",
            ["dataType"] = "delimited",
            ["other"] = true,
            ["otherChar"] = "|",
        }), "token");
        yield return ApplyPlanner.Batch("e4", "e4-scratch-cf", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "add_conditional_format",
            ["target"] = JsonUtil.Target(revision),
            ["range"] = "G3:G6",
            ["rule"] = new JsonObject { ["type"] = "cellValue", ["operator"] = "less", ["formula1"] = "0" },
            ["style"] = new JsonObject { ["fillColor"] = "#FFF2CC" },
        }), "token");
        yield return ApplyPlanner.Batch("e4", "e4-clear-cf", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "clear_conditional_formats",
            ["target"] = JsonUtil.Target(revision),
            ["range"] = "G3:G6",
        }), "token", "document CF on 자재대장 A3:I20 stays");
        yield return ApplyPlanner.Batch("e4", "e4-scratch-validation", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_data_validation",
            ["target"] = JsonUtil.Target(revision),
            ["range"] = "L3",
            ["type"] = "textLength",
            ["operator"] = "lessEqual",
            ["formula1"] = "20",
        }), "token");
        yield return ApplyPlanner.Batch("e4", "e4-clear-validation", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "clear_data_validation",
            ["target"] = JsonUtil.Target(revision),
            ["range"] = "L3",
        }), "token");
        yield return ApplyPlanner.Batch("e4", "e4-scratch-note", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_cell_note",
            ["target"] = JsonUtil.Target(revision),
            ["cell"] = "A1",
            ["text"] = "개정 노트",
        }), "token");
        yield return ApplyPlanner.Batch("e4", "e4-clear-note", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "clear_cell_note",
            ["target"] = JsonUtil.Target(revision),
            ["cell"] = "A1",
        }), "token", "C4 note on 자재대장 stays");
        yield return ApplyPlanner.Batch("e4", "e4-scratch-table", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "create_table",
            ["target"] = JsonUtil.Target(revision),
            ["name"] = "개정표",
            ["range"] = "A2:I6",
            ["hasHeaders"] = true,
        }), "token");
        yield return ApplyPlanner.Batch("e4", "e4-scratch-table-style", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "style_table",
            ["target"] = JsonUtil.Target(revision),
            ["name"] = "개정표",
            ["showRowStripes"] = true,
        }), "token", "evidence of 개정표 before delete", capture: "object-evidence");
        yield return ApplyPlanner.Batch("e4", "e4-delete-scratch-table", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "delete_table",
            ["target"] = JsonUtil.Target(revision),
            ["name"] = "개정표",
        }), "token", "do not delete 자재표; 개정 sheet stays");
        yield return ApplyPlanner.Batch("e4", "e4-merge-scratch", OpFamily.Merge, ApplyPlanner.Merges(revision, ["A20:B20"]), "token");
        yield return ApplyPlanner.Batch("e4", "e4-unmerge-scratch", OpFamily.Merge, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "unmerge_cells",
            ["target"] = JsonUtil.Target(revision),
            ["range"] = "A20:B20",
        }), "token");
        yield return ApplyPlanner.Batch("e4", "e4-hide-void", OpFamily.Visibility, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_rows_hidden",
            ["target"] = JsonUtil.Target(sheet),
            ["row"] = 12,
            ["count"] = 1,
            ["hidden"] = true,
        }), "token", "hide VOID on source ledger only");
        yield return ApplyPlanner.Batch("e4", "e4-hide-col", OpFamily.Visibility, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_cols_hidden",
            ["target"] = JsonUtil.Target(revision),
            ["col"] = "L",
            ["count"] = 1,
            ["hidden"] = true,
        }), "token");
        yield return ApplyPlanner.Batch("e4", "e4-unhide-col", OpFamily.Visibility, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_cols_hidden",
            ["target"] = JsonUtil.Target(revision),
            ["col"] = "L",
            ["count"] = 1,
            ["hidden"] = false,
        }), "token");
        yield return ApplyPlanner.Batch("e4", "e4-outline", OpFamily.RangeEdit, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_outline",
            ["target"] = JsonUtil.Target(revision),
            ["range"] = "A3:A6",
            ["level"] = 2,
            ["summaryBelow"] = true,
        }), "token");
        yield return ApplyPlanner.Batch("e4", "e4-view", OpFamily.SheetLayout, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_view",
            ["target"] = JsonUtil.Target(sheet),
            ["zoom"] = 100,
            ["view"] = "normal",
            ["displayGridlines"] = true,
        }), "token");
        yield return ApplyPlanner.Batch("e4", "e4-tab-color", OpFamily.Structure, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_tab_color",
            ["target"] = JsonUtil.Target(sheet),
            ["color"] = "#1F4E79",
        }), "token");
        yield return ApplyPlanner.Batch("e4", "e4-move-revision", OpFamily.Structure, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "move_sheet",
            ["target"] = JsonUtil.Target(revision),
            ["position"] = "last",
        }), "token");
        yield return ApplyPlanner.Batch("e4", "e4-sheet-hidden", OpFamily.Visibility, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_sheet_visibility",
            ["target"] = JsonUtil.Target(revision),
            ["visibility"] = "hidden",
        }), "token");
        yield return ApplyPlanner.Batch("e4", "e4-sheet-visible", OpFamily.Visibility, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_sheet_visibility",
            ["target"] = JsonUtil.Target(revision),
            ["visibility"] = "visible",
        }), "token");
        yield return ApplyPlanner.Batch("e4", "e4-shape-insert", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "insert_shape",
            ["target"] = JsonUtil.Target(revision),
            ["shapeType"] = "rectangle",
            ["name"] = "RevBox",
            ["position"] = new JsonObject { ["left"] = 480, ["top"] = 12, ["width"] = 36, ["height"] = 16 },
        }), "token");
        yield return ApplyPlanner.Batch("e4", "e4-shape-update", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "update_shape",
            ["target"] = JsonUtil.Target(revision),
            ["name"] = "RevBox",
            ["fillColor"] = "#D6DCE4",
        }), "token");
        yield return ApplyPlanner.Batch("e4", "e4-shape-delete", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "delete_shape",
            ["target"] = JsonUtil.Target(revision),
            ["name"] = "RevBox",
        }), "token");
        yield return ApplyPlanner.Batch("e4", "e4-textbox-insert", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "insert_textbox",
            ["target"] = JsonUtil.Target(revision),
            ["name"] = "RevText",
            ["text"] = "개정",
            ["position"] = new JsonObject { ["left"] = 480, ["top"] = 32, ["width"] = 64, ["height"] = 16 },
        }), "token");
        yield return ApplyPlanner.Batch("e4", "e4-textbox-update", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "update_textbox",
            ["target"] = JsonUtil.Target(revision),
            ["name"] = "RevText",
            ["text"] = "개정본",
        }), "token");
        yield return ApplyPlanner.Batch("e4", "e4-textbox-delete", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "delete_textbox",
            ["target"] = JsonUtil.Target(revision),
            ["name"] = "RevText",
        }), "token");
        yield return ApplyPlanner.Batch("e4", "e4-scratch-sheet", OpFamily.Structure, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "add_sheet",
            ["name"] = "스크래치삭제",
            ["afterSheet"] = revision,
        }), "token");
        yield return ApplyPlanner.Batch("e4", "e4-delete-scratch-sheet", OpFamily.Delete, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "delete_sheet",
            ["target"] = JsonUtil.Target("스크래치삭제"),
        }), "token", "개정 stays with TX-001/002/004/005; 자재대장 stays pristine");
    }

    public static IEnumerable<PlannedBatch> E5(string summary, string renamedInput, string xlsx)
    {
        yield return ApplyPlanner.Batch("e5", "e5-relative-name", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "define_name",
            ["name"] = "RelativeActiveCell",
            ["refersTo"] = "=B2",
            ["scope"] = "sheet",
            ["target"] = JsonUtil.Target(summary),
        }), "token", "separate relative-name coverage; ContractAmount/JanAmt stay quoted absolute $B$n");
        yield return ApplyPlanner.Batch("e5", "e5-scratch-rate", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "define_name",
            ["name"] = "ScratchRate",
            ["refersTo"] = "=0.1",
            ["scope"] = "workbook",
            ["replace"] = true,
        }), "token");
        yield return ApplyPlanner.Batch("e5", "e5-scratch-amount", OpFamily.Values, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_formulas",
            ["target"] = JsonUtil.Target(summary),
            ["range"] = "B6",
            ["formulas"] = Arr(Arr("=ScratchRate*100000000")),
        }), "execute", "initial 10,000,000 from 0.1 * 100M");
        yield return ApplyPlanner.Batch("e5", "e5-update-rate", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "update_name",
            ["name"] = "ScratchRate",
            ["refersTo"] = "=0.12",
            ["scope"] = "workbook",
        }), "token");
        yield return ApplyPlanner.Batch("e5", "e5-calc-rate", OpFamily.Values, JsonUtil.Arr(
            new JsonObject { ["op"] = "calculate" }), "token", "B6 must become 12,000,000 after rate 0.12");
        yield return ApplyPlanner.Batch("e5", "e5-delete-rate", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "delete_name",
            ["name"] = "ScratchRate",
            ["scope"] = "workbook",
        }), "token", "do not delete ContractAmount/JanAmt/FebAmt/MarAmt");
        yield return ApplyPlanner.Batch("e5", "e5-scratch-link", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_hyperlink",
            ["target"] = JsonUtil.Target(summary),
            ["range"] = "A6",
            ["subAddress"] = $"'{renamedInput}'!A1",
            ["textToDisplay"] = "스크래치 링크",
        }), "token");
        yield return ApplyPlanner.Batch("e5", "e5-clear-scratch-link", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "clear_hyperlink",
            ["target"] = JsonUtil.Target(summary),
            ["range"] = "A6",
        }), "token", "B5 월별기성!A1 acceptance link is not cleared");
        yield return ApplyPlanner.Batch("e5", "e5-copy-sheet", OpFamily.Structure, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "copy_sheet",
            ["sourceWorkbook"] = xlsx,
            ["sourceSheet"] = summary,
            ["targetSheet"] = "누계복사",
        }), "token");
        yield return ApplyPlanner.Batch("e5", "e5-tab-color", OpFamily.Structure, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_tab_color",
            ["target"] = JsonUtil.Target(renamedInput),
            ["color"] = "#305496",
        }), "token");
        yield return ApplyPlanner.Batch("e5", "e5-delete-copy-sheet", OpFamily.Delete, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "delete_sheet",
            ["target"] = JsonUtil.Target("누계복사"),
        }), "token", "copied sheet only; 월별기성/누계집계 stay");
    }

    public static IEnumerable<PlannedBatch> E6(string sheet)
    {
        yield return ApplyPlanner.Batch("e6", "e6-update-chart", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "update_chart",
            ["target"] = JsonUtil.Target(sheet),
            ["name"] = "PlanActualLine",
            ["sourceRange"] = "A2:C8",
            ["title"] = "계획/실적 누계율",
            ["series"] = JsonUtil.Arr(
                new JsonObject { ["values"] = "B2:B8", ["categories"] = "A2:A8", ["name"] = "계획누계율" },
                new JsonObject { ["values"] = "C2:C8", ["categories"] = "A2:A8", ["name"] = "실적누계율" }),
        }), "token", "retarget the real report chart after Mar 45 with applied series; geometry no-overlap is fixture-owned", capture: "object-evidence");
        yield return ApplyPlanner.Batch("e6", "e6-update-diff-chart", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "update_chart",
            ["target"] = JsonUtil.Target(sheet),
            ["name"] = "DiffColumn",
            ["sourceRange"] = "A2:A8,D2:D8",
            ["series"] = JsonUtil.Arr(
                new JsonObject { ["values"] = "D2:D8", ["categories"] = "A2:A8", ["name"] = "수량" }),
        }), "token", "data pin union + series; F2:G8 helper remains valid", capture: "object-evidence");
        yield return ApplyPlanner.Batch("e6", "e6-scratch-chart-create", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "create_chart",
            ["target"] = JsonUtil.Target(sheet),
            ["name"] = "ScratchChart",
            ["chartType"] = "columnClustered",
            ["sourceRange"] = "A2:B4",
            ["position"] = new JsonObject { ["left"] = 12, ["top"] = 20, ["width"] = 80, ["height"] = 48 },
        }), "token");
        yield return ApplyPlanner.Batch("e6", "e6-scratch-chart-update", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "update_chart",
            ["target"] = JsonUtil.Target(sheet),
            ["name"] = "ScratchChart",
            ["title"] = "스크래치",
        }), "token", "evidence before delete; PlanActualLine is the real retarget", capture: "object-evidence");
        yield return ApplyPlanner.Batch("e6", "e6-scratch-chart-delete", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "delete_chart",
            ["target"] = JsonUtil.Target(sheet),
            ["name"] = "ScratchChart",
        }), "token", "final charts remain PlanActualLine and DiffColumn");
        yield return ApplyPlanner.Batch("e6", "e6-set-view", OpFamily.SheetLayout, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_view",
            ["target"] = JsonUtil.Target(sheet),
            ["zoom"] = 100,
            ["view"] = "normal",
            ["displayGridlines"] = true,
        }), "token");
    }

    public static IEnumerable<PlannedBatch> E7(string sheet)
    {
        yield return ApplyPlanner.Batch("e7", "e7-style-table", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "style_table",
            ["target"] = JsonUtil.Target(sheet),
            ["name"] = "원장표",
            ["showRowStripes"] = true,
        }), "token");
        yield return ApplyPlanner.Batch("e7", "e7-scratch-pivot", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "create_pivot",
            ["target"] = JsonUtil.Target(sheet),
            ["name"] = "스크래치피벗",
            ["sourceRange"] = $"{sheet}!A1:C3",
            ["destination"] = $"{sheet}!E20",
            ["rows"] = JsonUtil.Arr("품목"),
            ["values"] = JsonUtil.Arr(new JsonObject { ["field"] = "금액", ["function"] = "sum" }),
        }), "token", "scratch only; 원장피벗 on 피벗!A3 stays");
        yield return ApplyPlanner.Batch("e7", "e7-delete-scratch-pivot", OpFamily.Data, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "delete_pivot",
            ["target"] = JsonUtil.Target(sheet),
            ["name"] = "스크래치피벗",
        }), "token");
    }

    public static IEnumerable<PlannedBatch> E8(string sheet)
    {
        yield return ApplyPlanner.Batch("e8", "e8-b1-unprotect", OpFamily.Protect, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "unprotect_sheet",
            ["target"] = JsonUtil.Target(sheet),
        }), "token", "Remaining review: unlocked B1 edit must recalc B3; fixture IDs e8-unprotect/e8-reprotect stay");
        yield return ApplyPlanner.Batch("e8", "e8-unlock-b1", OpFamily.Values, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_values",
            ["target"] = JsonUtil.Target(sheet),
            ["range"] = "B1",
            ["values"] = Arr(Arr(11)),
        }), "execute", "unlocked B1 10→11; B3 formula must become 11000");
        yield return ApplyPlanner.Batch("e8", "e8-calculate", OpFamily.Values, JsonUtil.Arr(
            new JsonObject { ["op"] = "calculate" }), "token", "B3=B1*B2 → 11000");
        yield return ApplyPlanner.Batch("e8", "e8-b1-reprotect", OpFamily.Protect, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "protect_sheet",
            ["target"] = JsonUtil.Target(sheet),
        }), "token", "final protected so locked B3 write is refused");
    }

    public static IEnumerable<PlannedBatch> E9(string sheet)
    {
        _ = sheet;
        yield break;
    }

    private static JsonArray Arr(params object?[] cells)
    {
        var a = new JsonArray();
        foreach (var c in cells) a.Add(JsonUtil.FromClr(c));
        return a;
    }
}
