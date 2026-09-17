using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelPhase1ContractTests : IClassFixture<TestHome>
{
    private readonly OperationValidator _validator = new(new PolicyEngine());

    [Fact]
    public void Multi_merge_batch_is_accepted_when_ranges_do_not_overlap()
    {
        var batch = Json.ParseObject("""
        {
          "ops": [
            { "op": "merge_cells", "target": { "sheet": "예정공정표" }, "range": "A1:B1" },
            { "op": "merge_cells", "target": { "sheet": "예정공정표" }, "range": "C1:D1" }
          ],
          "dryRun": true
        }
        """);
        var errors = new List<string>();
        Assert.NotNull(_validator.Validate(batch, "excel", errors));
        Assert.Empty(errors);
    }

    [Fact]
    public void Overlapping_merge_batch_is_rejected()
    {
        var batch = Json.ParseObject("""
        {
          "ops": [
            { "op": "merge_cells", "target": { "sheet": "예정공정표" }, "range": "A1:C1" },
            { "op": "merge_cells", "target": { "sheet": "예정공정표" }, "range": "C1:E1" }
          ],
          "dryRun": true
        }
        """);
        var errors = new List<string>();
        Assert.Null(_validator.Validate(batch, "excel", errors));
        Assert.Contains(errors, error => error.Contains("EXCEL_MERGE_BATCH_OVERLAP", StringComparison.Ordinal));
    }

    [Fact]
    public void Merge_and_unmerge_cannot_share_a_batch()
    {
        var batch = Json.ParseObject("""
        {
          "ops": [
            { "op": "merge_cells", "target": { "sheet": "예정공정표" }, "range": "A1:B1" },
            { "op": "unmerge_cells", "target": { "sheet": "예정공정표" }, "range": "C1:D1" }
          ],
          "dryRun": true
        }
        """);
        var errors = new List<string>();
        Assert.Null(_validator.Validate(batch, "excel", errors));
        Assert.Contains(errors, error => error.Contains("cannot share a batch", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void New_format_keys_and_cell_bottom_borders_normalize()
    {
        var errors = new List<string>();
        Assert.True(ExcelStyleContract.TryNormalize(Json.ParseObject("""
        {
          "fontName": "돋움",
          "horizontalAlign": "center",
          "verticalAlign": "center",
          "wrapText": true,
          "borders": { "bottom": { "weight": "medium", "lineStyle": "continuous" } }
        }
        """), out var canonical, errors));
        Assert.Empty(errors);
        Assert.Equal("돋움", Json.GetString(canonical, "fontName"));
        Assert.Equal("center", Json.GetString(canonical, "horizontalAlign"));
        var edges = ExcelBorderContract.ReadEdges(Json.GetObj(canonical, "borders")!);
        Assert.Contains(edges, edge => edge.Name == "bottom" && edge.Scope == ExcelBorderContract.ScopeCell &&
                                      edge.Weight == ExcelBorderContract.WeightMedium);
    }

    [Fact]
    public void Published_edge_schema_rejects_banana_weight_and_accepts_true_none_or_object()
    {
        var all = Json.GetObj(Json.GetObj(ExcelBorderContract.DescribeSchema(), "properties"), "all");
        Assert.NotNull(all?["anyOf"]);
        var errors = new List<string>();
        Assert.False(ExcelDataOperationSchema.ValidateNode(
            JsonNode.Parse("""{"weight":"banana"}"""), all!, errors, "$.borders.all"));
        Assert.Contains(errors, error => error.Contains("enum", StringComparison.OrdinalIgnoreCase) ||
                                        error.Contains("anyOf", StringComparison.OrdinalIgnoreCase));
        errors.Clear();
        Assert.True(ExcelDataOperationSchema.ValidateNode(JsonValue.Create(true), all!, errors, "$.borders.all"));
        Assert.Empty(errors);
        Assert.True(ExcelDataOperationSchema.ValidateNode(JsonValue.Create("none"), all!, errors, "$.borders.all"));
        Assert.Empty(errors);
        Assert.True(ExcelDataOperationSchema.ValidateNode(
            Json.ParseObject("""{ "weight": "thin", "lineStyle": "continuous", "color": "#000000" }"""),
            all!, errors, "$.borders.all"));
        Assert.Empty(errors);
    }

    [Fact]
    public void Medium_is_a_border_weight_not_a_line_style()
    {
        var errors = new List<string>();
        Assert.False(ExcelBorderContract.TryNormalize(Json.ParseObject("""
        { "bottom": { "lineStyle": "medium" } }
        """), 1, errors, out _));
        Assert.Contains(errors, error => error.Contains("lineStyle", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Thirty_one_month_header_merge_pair_batch_is_accepted()
    {
        var ops = new JsonArray();
        foreach (var range in ExcelMergeBatchContract.MonthHeaderPairRanges())
        {
            ops.Add(new JsonObject
            {
                ["op"] = "merge_cells",
                ["target"] = new JsonObject { ["sheet"] = "예정공정표" },
                ["range"] = range,
            });
        }

        var batch = new JsonObject { ["ops"] = ops, ["dryRun"] = true };
        var errors = new List<string>();
        Assert.Equal(ExcelMergeBatchContract.MonthHeaderPairCount, ops.Count);
        Assert.Equal("G5:H5", ExcelMergeBatchContract.MonthHeaderPairRanges()[0]);
        Assert.NotNull(_validator.Validate(batch, "excel", errors));
        Assert.Empty(errors);
    }

    [Fact]
    public void Page_setup_scale_wins_over_fit_and_maps_a3()
    {
        var errors = new List<string>();
        Assert.True(ExcelPageSetupContract.TryNormalize(Json.ParseObject("""
        {
          "paperSize": "A3",
          "orientation": "landscape",
          "scale": 55,
          "fitToWidth": 1,
          "printArea": "A1:BQ96"
        }
        """), 1, errors, out var page));
        Assert.Empty(errors);
        Assert.Equal(ExcelPageSetupContract.XlPaperA3, Json.GetInt(page, "paperSize"));
        Assert.Equal(ExcelPageSetupContract.ModeScale, Json.GetString(page, "zoomMode"));
        Assert.Equal(55, Json.GetInt(page, "scale"));
        Assert.True(Json.GetBool(page, "fitIgnored"));
    }

    [Fact]
    public void Freeze_G6_and_row_column_units_are_validated()
    {
        var freeze = Json.ParseObject("""
        { "ops": [ { "op": "freeze_panes", "target": { "sheet": "예정공정표" }, "cell": "G6" } ], "dryRun": true }
        """);
        var freezeErrors = new List<string>();
        Assert.NotNull(_validator.Validate(freeze, "excel", freezeErrors));
        Assert.Empty(freezeErrors);
        Assert.True(ExcelA1Box.TryParseCell("G6", out var row, out var col));
        Assert.Equal(6, row);
        Assert.Equal(7, col);

        var sizes = Json.ParseObject("""
        {
          "ops": [
            { "op": "set_row_heights", "target": { "sheet": "예정공정표" }, "rows": [ { "row": 5, "heightPoints": 18 } ] },
            { "op": "set_column_widths", "target": { "sheet": "예정공정표" }, "columns": [ { "col": "A", "widthChars": 12 } ] }
          ],
          "dryRun": true
        }
        """);
        var sizeErrors = new List<string>();
        Assert.NotNull(_validator.Validate(sizes, "excel", sizeErrors));
        Assert.Empty(sizeErrors);
    }

    [Fact]
    public void Rename_rejects_illegal_names_and_protected_save_path()
    {
        var rename = Json.ParseObject("""
        { "ops": [ { "op": "rename_sheet", "target": { "sheet": "Sheet1" }, "newName": "bad:name" } ], "dryRun": true }
        """);
        var renameErrors = new List<string>();
        Assert.Null(_validator.Validate(rename, "excel", renameErrors));
        Assert.Contains(renameErrors, error => error.Contains("forbidden", StringComparison.OrdinalIgnoreCase));

        const string protectedPath = @"C:\DocBridgeTest\Desktop\AI작업\10.만회대책\수원시_하수관로정비_공사예정공정표(변경).xlsx";
        Environment.SetEnvironmentVariable(ExcelAuthoringPaths.ProtectedSourcesVariable, protectedPath);
        try
        {
            var save = Json.ParseObject("""
            {
              "ops": [
                { "op": "save_workbook", "output": "C:\\DocBridgeTest\\Desktop\\AI작업\\10.만회대책\\수원시_하수관로정비_공사예정공정표(변경).xlsx" }
              ],
              "dryRun": true
            }
            """);
            var saveErrors = new List<string>();
            Assert.Null(_validator.Validate(save, "excel", saveErrors));
            Assert.Contains(saveErrors, error => error.Contains("protected source", StringComparison.OrdinalIgnoreCase));
            Assert.True(ExcelAuthoringPaths.IsProtectedSource(protectedPath));
        }
        finally
        {
            Environment.SetEnvironmentVariable(ExcelAuthoringPaths.ProtectedSourcesVariable, null);
        }
    }

    [Fact]
    public void Execute_still_blocks_merge_and_lifecycle_and_keeps_auto_execute_allowlist()
    {
        var policy = new PolicyEngine();
        Assert.True(policy.IsAutoExecutable("excel", "set_values"));
        Assert.True(policy.IsAutoExecutable("excel", "set_formulas"));
        Assert.True(policy.IsAutoExecutable("excel", "format_range"));
        Assert.False(policy.IsAutoExecutable("excel", "merge_cells"));
        Assert.False(policy.IsAutoExecutable("excel", "save_workbook"));
        Assert.False(policy.IsAutoExecutable("excel", "create_table"));
        Assert.Equal(OpClass.HighRisk, policy.ClassifyOp("excel", "delete_rows"));
        Assert.Equal(OpClass.HighRisk, policy.ClassifyOp("excel", "export_pdf"));

        var execute = Json.ParseObject("""
        {
          "ops": [ { "op": "merge_cells", "target": { "sheet": "예정공정표" }, "range": "A1:B1" } ],
          "executionMode": "execute",
          "requestId": "11111111-1111-1111-1111-111111111111",
          "expectedDocumentRef": "C:\\tmp\\schedule.xlsx"
        }
        """);
        var errors = new List<string>();
        Assert.Null(_validator.Validate(execute, "excel", errors));
        Assert.Contains(errors, error => error.Contains("does not allow", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Data_ops_stay_isolated_and_are_not_unknown()
    {
        var ok = Json.ParseObject("""
        {
          "ops": [ { "op": "create_table", "target": { "sheet": "자재대장" }, "range": "A1:C3", "name": "Materials" } ],
          "dryRun": true
        }
        """);
        var okErrors = new List<string>();
        Assert.NotNull(_validator.Validate(ok, "excel", okErrors));
        Assert.Empty(okErrors);

        var mixed = Json.ParseObject("""
        {
          "ops": [
            { "op": "create_table", "target": { "sheet": "자재대장" }, "range": "A1:C3" },
            { "op": "set_values", "target": { "sheet": "자재대장" }, "range": "A1", "values": [[1]] }
          ],
          "dryRun": true
        }
        """);
        var mixedErrors = new List<string>();
        Assert.Null(_validator.Validate(mixed, "excel", mixedErrors));
        Assert.Contains(mixedErrors, error => error.Contains("data/reporting", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Capabilities_publish_phase1_and_data_ops_without_changing_auto_execute()
    {
        using var adapter = new ExcelAdapter(() => null);
        var capabilities = adapter.GetCapabilities();
        var writeOps = Json.GetArr(capabilities, "writeOps")!
            .Select(node => node!.GetValue<string>()).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var auto = Json.GetArr(capabilities, "autoExecuteOps")!
            .Select(node => node!.GetValue<string>()).ToHashSet(StringComparer.OrdinalIgnoreCase);
        Assert.Contains("set_row_heights", writeOps);
        Assert.Contains("freeze_panes", writeOps);
        Assert.Contains("create_workbook", writeOps);
        Assert.Contains("create_table", writeOps);
        Assert.True(auto.SetEquals(new[] { "set_values", "set_formulas", "format_range" }));
        Assert.Equal(400, Json.GetInt(Json.GetObj(capabilities, "limits"), "maxMergeBatchOperations"));
    }

    [Fact]
    public void Password_fields_are_rejected()
    {
        var batch = Json.ParseObject("""
        {
          "ops": [ { "op": "protect_sheet", "target": { "sheet": "예정공정표" }, "password": "secret" } ],
          "dryRun": true
        }
        """);
        var errors = new List<string>();
        Assert.Null(_validator.Validate(batch, "excel", errors));
        Assert.Contains(errors, error => error.Contains("password", StringComparison.OrdinalIgnoreCase));
    }
}
