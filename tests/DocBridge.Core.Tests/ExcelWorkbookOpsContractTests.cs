using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelWorkbookOpsContractTests
{
    [Fact]
    public void Workbook_op_catalog_is_registered()
    {
        foreach (var name in new[]
                 {
                     "update_external_links", "change_link_source", "break_external_link",
                     "set_calculation_mode", "freeze_values", "paste_special", "goal_seek",
                     "protect_workbook", "unprotect_workbook", "set_split_panes",
                     "create_sparkline", "update_sparkline", "delete_sparkline",
                     "create_slicer", "delete_slicer", "apply_cell_style",
                 })
        {
            Assert.True(ExcelDataOperationsContract.IsDataOperation(name), name);
            Assert.True(ExcelDataOperationsContract.RequiredFields.ContainsKey(name), name);
            Assert.True(ExcelDataOperationsContract.OptionalFields.ContainsKey(name), name);
        }
        foreach (var scope in new[] { "links", "sparklines", "slicers", "cellStyles" })
            Assert.True(ExcelDataOperationsContract.IsReadScope(scope), scope);
        foreach (var name in new[]
                 {
                     "update_external_links", "change_link_source", "break_external_link",
                     "set_calculation_mode", "protect_workbook", "unprotect_workbook",
                 })
            Assert.True(ExcelDataOperationsContract.IsWorkbookLevelOp(name), name);
        Assert.False(ExcelDataOperationsContract.IsWorkbookLevelOp("freeze_values"));
    }

    [Fact]
    public void Workbook_level_ops_do_not_require_a_sheet()
    {
        var update = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "update_external_links", """{ "op": "update_external_links" }"""));
        Assert.Empty(update);

        var change = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "change_link_source",
            """
            { "op": "change_link_source", "source": "a.xlsx", "newSource": "C:\\data\\b.xlsx" }
            """));
        Assert.DoesNotContain(change, error => error.Contains("target.sheet"));

        var calc = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "set_calculation_mode", """{ "op": "set_calculation_mode", "mode": "manual" }"""));
        Assert.Empty(calc);
    }

    [Fact]
    public void Change_link_source_requires_an_existing_absolute_target()
    {
        var missing = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "change_link_source",
            """
            { "op": "change_link_source", "source": "a.xlsx", "newSource": "C:\\no\\such\\file.xlsx" }
            """));
        Assert.Contains(missing, error => error.Contains("was not found on disk"));

        var relative = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "change_link_source",
            """
            { "op": "change_link_source", "source": "a.xlsx", "newSource": "relative\\b.xlsx" }
            """));
        Assert.Contains(relative, error => error.Contains("absolute"));
    }

    [Fact]
    public void Calculation_and_paste_tokens_are_strict()
    {
        var mode = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "set_calculation_mode", """{ "op": "set_calculation_mode", "mode": "sometimes" }"""));
        Assert.Contains(mode, error => error.Contains("mode"));

        var recalc = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "set_calculation_mode", """{ "op": "set_calculation_mode", "mode": "manual", "recalculate": "later" }"""));
        Assert.Contains(recalc, error => error.Contains("recalculate"));

        var paste = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "paste_special",
            """
            { "op": "paste_special", "target": { "sheet": "자재대장" },
              "sourceRange": "A1:B2", "destination": "D1", "paste": "essence" }
            """));
        Assert.Contains(paste, error => error.Contains("paste"));

        var operation = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "paste_special",
            """
            { "op": "paste_special", "target": { "sheet": "자재대장" },
              "sourceRange": "A1:B2", "destination": "D1", "paste": "formulas", "operation": "add" }
            """));
        Assert.Contains(operation, error => error.Contains("operation"));
    }

    [Fact]
    public void Goal_seek_cells_must_share_a_sheet()
    {
        var cross = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "goal_seek",
            """
            { "op": "goal_seek", "target": { "sheet": "원가" },
              "cell": "시트A!B1", "goalCell": "시트B!B3", "goal": 5 }
            """));
        Assert.Contains(cross, error => error.Contains("same worksheet"));

        var multi = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "goal_seek",
            """
            { "op": "goal_seek", "target": { "sheet": "원가" },
              "cell": "B1:B2", "goalCell": "B3", "goal": 5 }
            """));
        Assert.Contains(multi, error => error.Contains("single A1 cell"));

        var ok = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "goal_seek",
            """
            { "op": "goal_seek", "target": { "sheet": "원가" },
              "cell": "B1", "goalCell": "B3", "goal": 11000 }
            """));
        Assert.Empty(ok);
    }

    [Fact]
    public void Split_panes_rejects_mixed_anchors_and_remove_combinations()
    {
        var mixed = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "set_split_panes",
            """
            { "op": "set_split_panes", "target": { "sheet": "자재대장" }, "cell": "C5", "splitRows": 4 }
            """));
        Assert.Contains(mixed, error => error.Contains("mutually exclusive"));

        var remove = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "set_split_panes",
            """
            { "op": "set_split_panes", "target": { "sheet": "자재대장" }, "remove": true, "cell": "C5" }
            """));
        Assert.Contains(remove, error => error.Contains("remove"));

        var ok = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "set_split_panes",
            """
            { "op": "set_split_panes", "target": { "sheet": "자재대장" }, "cell": "C5", "freeze": true }
            """));
        Assert.Empty(ok);
    }

    [Fact]
    public void Sparkline_winloss_is_rejected_until_native_mapping_is_proven()
    {
        var winLoss = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "create_sparkline",
            """
            { "op": "create_sparkline", "target": { "sheet": "월간보고" },
              "location": "C2", "sourceData": "B2:B6", "type": "winLoss" }
            """));
        Assert.Contains(winLoss, error => error.Contains("winLoss"));

        var idle = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "update_sparkline",
            """
            { "op": "update_sparkline", "target": { "sheet": "월간보고" }, "location": "C2" }
            """));
        Assert.Contains(idle, error => error.Contains("updatable field"));

        var ok = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "create_sparkline",
            """
            { "op": "create_sparkline", "target": { "sheet": "월간보고" },
              "location": "C2", "sourceData": "B2:B6", "type": "line", "showHigh": true }
            """));
        Assert.Empty(ok);
    }

    [Fact]
    public void Slicer_names_are_required_and_positions_are_validated()
    {
        var nameless = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "create_slicer",
            """
            { "op": "create_slicer", "target": { "sheet": "월간보고" }, "source": "Materials", "field": "구분" }
            """));
        Assert.Contains(nameless, error => error.Contains("name"));

        var split = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "delete_slicer",
            """
            { "op": "delete_slicer", "target": { "sheet": "월간보고" },
              "name": "Slicer_구분", "source": "Materials" }
            """));
        Assert.Contains(split, error => error.Contains("required together"));
    }

    [Fact]
    public void Cell_style_and_freeze_ranges_are_bounded()
    {
        var wide = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "apply_cell_style",
            """
            { "op": "apply_cell_style", "target": { "sheet": "자재대장" },
              "range": "A1:Z500", "styleName": "Normal" }
            """));
        Assert.Contains(wide, error => error.Contains("5000"));

        var huge = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "freeze_values",
            """
            { "op": "freeze_values", "target": { "sheet": "자재대장" }, "range": "A1:XFD100" }
            """));
        Assert.Contains(huge, error => error.Contains("20000"));
    }

    [Fact]
    public void Calculation_mode_and_link_status_tokens_map()
    {
        Assert.True(ExcelWorkbookOpsContract.TryCalculationMode("manual", out var manual));
        Assert.Equal(-4135, manual);
        Assert.True(ExcelWorkbookOpsContract.TryCalculationMode("AUTOMATIC", out var auto));
        Assert.Equal(-4105, auto);
        Assert.False(ExcelWorkbookOpsContract.TryCalculationMode("sometimes", out _));
        Assert.Equal("ok", ExcelWorkbookOpsContract.LinkStatusName(0));
        Assert.Equal("missingFile", ExcelWorkbookOpsContract.LinkStatusName(1));
        Assert.True(ExcelWorkbookOpsContract.IsUpdatableLinkStatus(0));
        Assert.False(ExcelWorkbookOpsContract.IsUpdatableLinkStatus(1));
        Assert.True(ExcelWorkbookOpsContract.TryResolveSplit(
            Json.ParseObject("""{ "cell": "C5" }""")!, out var rows, out var cols, out _));
        Assert.Equal(4, rows);
        Assert.Equal(2, cols);
        Assert.False(ExcelWorkbookOpsContract.TryResolveSplit(
            Json.ParseObject("""{ "cell": "C5", "splitRows": 1 }""")!, out _, out _, out var error));
        Assert.Contains("mutually exclusive", error, StringComparison.OrdinalIgnoreCase);
        Assert.Equal("집계.xlsx", ExcelWorkbookOpsContract.LeafOf("C:\\data\\집계.xlsx"));
    }

    [Fact]
    public void New_ops_match_exactly_one_published_branch()
    {
        var items = PublishedItems();
        foreach (var (name, payload) in PublishedPayloads())
        {
            var errors = new List<string>();
            Assert.True(ExcelDataOperationSchema.TryValidatePublishedItems(items, payload, errors),
                name + ": " + string.Join("; ", errors));
            Assert.Equal(1, ExcelDataOperationSchema.CountPublishedMatches(items, payload));
            Assert.Empty(ExcelDataOperationsContract.ValidatePublicInput((JsonObject)payload.DeepClone()));
        }
    }

    private static Dictionary<string, JsonObject> PublishedPayloads() => new(StringComparer.OrdinalIgnoreCase)
    {
        ["update_external_links"] = Op("""{ "op": "update_external_links" }"""),
        ["change_link_source"] = ChangePayload(),
        ["break_external_link"] = Op("""{ "op": "break_external_link", "source": "a.xlsx" }"""),
        ["set_calculation_mode"] = Op("""{ "op": "set_calculation_mode", "mode": "manual", "recalculate": "full" }"""),
        ["freeze_values"] = Op("""{ "op": "freeze_values", "target": { "sheet": "자재대장" }, "range": "A1:G8" }"""),
        ["paste_special"] = Op("""{ "op": "paste_special", "target": { "sheet": "자재대장" }, "sourceRange": "'원가'!A1:B2", "destination": "D1", "paste": "values", "transpose": true }"""),
        ["goal_seek"] = Op("""{ "op": "goal_seek", "target": { "sheet": "원가" }, "cell": "B1", "goalCell": "B3", "goal": 7.5 }"""),
        ["protect_workbook"] = Op("""{ "op": "protect_workbook", "structure": true }"""),
        ["unprotect_workbook"] = Op("""{ "op": "unprotect_workbook" }"""),
        ["set_split_panes"] = Op("""{ "op": "set_split_panes", "target": { "sheet": "자재대장" }, "splitRows": 2, "splitColumns": 1 }"""),
        ["create_sparkline"] = Op("""{ "op": "create_sparkline", "target": { "sheet": "월간보고" }, "location": "C2:C6", "sourceData": "B2:B6", "type": "column" }"""),
        ["update_sparkline"] = Op("""{ "op": "update_sparkline", "target": { "sheet": "월간보고" }, "location": "C2:C6", "markers": true, "lineColor": "#FF0000" }"""),
        ["delete_sparkline"] = Op("""{ "op": "delete_sparkline", "target": { "sheet": "월간보고" }, "location": "C2:C6" }"""),
        ["create_slicer"] = Op("""{ "op": "create_slicer", "target": { "sheet": "월간보고" }, "source": "Materials", "field": "구분", "name": "Slicer_구분", "position": { "left": 8, "top": 8 } }"""),
        ["delete_slicer"] = Op("""{ "op": "delete_slicer", "target": { "sheet": "월간보고" }, "name": "Slicer_구분", "source": "Materials", "field": "구분" }"""),
        ["apply_cell_style"] = Op("""{ "op": "apply_cell_style", "target": { "sheet": "자재대장" }, "range": "A1:G1", "styleName": "Normal" }"""),
    };

    private static JsonObject ChangePayload()
    {
        var path = Path.Combine(Path.GetTempPath(), "docbridge-link-target.xlsx");
        try { File.WriteAllText(path, "probe"); } catch { /* validation reads existence only */ }
        return Json.ParseObject(
            """{ "op": "change_link_source", "source": "a.xlsx", "newSource": "__PATH__" }"""
                .Replace("__PATH__", path.Replace("\\", "\\\\", StringComparison.Ordinal), StringComparison.Ordinal))!;
    }

    private static JsonObject PublishedItems()
    {
        var items = new JsonObject
        {
            ["type"] = "object",
            ["required"] = new JsonArray("op"),
            ["properties"] = new JsonObject
            {
                ["op"] = new JsonObject
                {
                    ["type"] = "string",
                    ["enum"] = new JsonArray("set_values", "move_sheet"),
                },
            },
        };
        ExcelDataOperationSchema.AttachToApplyItems(items);
        return items;
    }

    private static JsonObject Op(string _, string json) => Op(json);

    private static JsonObject Op(string json) =>
        Json.ParseObject(json) ?? throw new InvalidOperationException(json);
}
