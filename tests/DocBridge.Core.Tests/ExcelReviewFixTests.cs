using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelReviewFixTests
{
    [Fact]
    public void Lifecycle_restore_closes_only_owned_workbooks()
    {
        var owned = new JsonArray
        {
            new JsonObject { ["name"] = "Book2", ["fullName"] = "Book2" },
        };
        Assert.True(ExcelAuthoringSchema.ShouldCloseOnLifecycleRestore("Book2", "Book2", owned));
        Assert.False(ExcelAuthoringSchema.ShouldCloseOnLifecycleRestore("수원시.xlsx", @"C:\work\수원시.xlsx", owned));
        Assert.False(ExcelAuthoringSchema.ShouldCloseOnLifecycleRestore("Book3", "Book3", new JsonArray()));
    }

    [Fact]
    public void Same_path_save_does_not_require_overwrite()
    {
        var path = @"C:\out\schedule.xlsx";
        Assert.False(ExcelAuthoringSchema.SaveConflictsWithExistingFile(path, path, overwrite: false, targetExists: true));
        Assert.True(ExcelAuthoringSchema.SaveConflictsWithExistingFile(@"C:\out\other.xlsx", path, overwrite: false, targetExists: true));
        Assert.False(ExcelAuthoringSchema.SaveConflictsWithExistingFile(path, path, overwrite: true, targetExists: true));
    }

    [Fact]
    public void Page_setup_mm_capture_materializes_points_and_matches_margins_and_fit()
    {
        var captured = new JsonObject
        {
            ["leftMarginMm"] = 10d,
            ["headerMarginMm"] = 8d,
            ["zoomMode"] = ExcelPageSetupContract.ModeFit,
            ["fitToWidth"] = 1,
            ["fitToHeight"] = 1,
        };
        ExcelPageSetupContract.EnsurePointKeys(captured);
        Assert.True(captured.ContainsKey("leftMarginPoints"));
        Assert.True(captured.ContainsKey("headerMarginPoints"));

        var actual = new JsonObject
        {
            ["leftMarginMm"] = 10d,
            ["headerMarginMm"] = 8d,
            ["zoomMode"] = ExcelPageSetupContract.ModeFit,
            ["fitToWidth"] = 1,
            ["fitToHeight"] = 1,
        };
        Assert.True(ExcelPageSetupContract.MatchesRequested(actual, captured));
        actual["fitToWidth"] = 2;
        Assert.False(ExcelPageSetupContract.MatchesRequested(actual, captured));
        var why = ExcelPageSetupContract.DescribeMismatches(actual, captured);
        Assert.Contains(why, item => item.Contains("fitToWidth", StringComparison.Ordinal) &&
                                    item.Contains("expected", StringComparison.Ordinal));
    }

    [Fact]
    public void Page_setup_normalizes_title_ranges_but_keeps_header_footer_literals()
    {
        var requested = new JsonObject
        {
            ["printArea"] = "A1:F20",
            ["printTitleRows"] = "1:2",
            ["printTitleColumns"] = "A:B",
            ["leftHeader"] = "1:2",
            ["centerFooter"] = "Cost $1:$2",
        };
        var excel = new JsonObject
        {
            ["printArea"] = "$A$1:$F$20",
            ["printTitleRows"] = "$1:$2",
            ["printTitleColumns"] = "$A:$B",
            ["leftHeader"] = "1:2",
            ["centerFooter"] = "Cost $1:$2",
        };
        Assert.True(ExcelPageSetupContract.PrintRangesMatch("1:2", "$1:$2"));
        Assert.True(ExcelPageSetupContract.MatchesRequested(excel, requested));
        Assert.Empty(ExcelPageSetupContract.DescribeMismatches(excel, requested));

        excel["leftHeader"] = "$1:$2";
        Assert.False(ExcelPageSetupContract.HeaderFooterMatch("1:2", "$1:$2"));
        Assert.False(ExcelPageSetupContract.MatchesRequested(excel, requested));
        var why = ExcelPageSetupContract.DescribeMismatches(excel, requested);
        Assert.Contains(why, item => item.StartsWith("leftHeader ", StringComparison.Ordinal));
        Assert.DoesNotContain(why, item => item.StartsWith("printTitleRows ", StringComparison.Ordinal));
    }

    [Fact]
    public void Delete_restore_skips_unapplied_entries()
    {
        var state = new JsonObject
        {
            ["entries"] = new JsonArray
            {
                new JsonObject { ["applied"] = false, ["start"] = 2, ["count"] = 1 },
                new JsonObject { ["applied"] = true, ["start"] = 5, ["count"] = 1 },
            },
        };
        Assert.False(ExcelAuthoringSchema.ShouldRestoreDeleteEntry((JsonObject)Json.GetArr(state, "entries")![0]!));
        Assert.True(ExcelAuthoringSchema.MarkDeleteEntryApplied(state, 0));
        Assert.True(ExcelAuthoringSchema.ShouldRestoreDeleteEntry((JsonObject)Json.GetArr(state, "entries")![0]!));
    }

    [Fact]
    public void Inspect_normalizes_scope_and_objectKind()
    {
        var fromKind = ExcelAuthoringSchema.NormalizeInspectArgs(Json.ParseObject("""{ "scope": "objects", "objectKind": "tables" }""")!);
        Assert.Equal("objects", Json.GetString(fromKind, "scope"));
        Assert.Equal("tables", Json.GetString(fromKind, "objectKind"));

        var fromScope = ExcelAuthoringSchema.NormalizeInspectArgs(Json.ParseObject("""{ "scope": "names" }""")!);
        Assert.Equal("objects", Json.GetString(fromScope, "scope"));
        Assert.Equal("names", Json.GetString(fromScope, "objectKind"));
        Assert.True(ExcelAuthoringSchema.IsDataObjectInspect("names", null));
    }

    [Fact]
    public void New_format_keys_normalize_and_stay_out_of_deferred_set()
    {
        var errors = new List<string>();
        Assert.True(ExcelStyleContract.TryNormalize(Json.ParseObject("""
        {
          "shrinkToFit": true,
          "underline": "single",
          "strikethrough": true,
          "indent": 1,
          "orientation": 45,
          "locked": true,
          "noFill": true
        }
        """), out var canonical, errors));
        Assert.Empty(errors);
        Assert.True(Json.GetBool(canonical, ExcelStyleContract.NoFill));
        Assert.Equal("single", Json.GetString(canonical, ExcelStyleContract.Underline));

        var conflict = new List<string>();
        Assert.False(ExcelStyleContract.TryNormalize(Json.ParseObject("""{ "noFill": true, "fillColor": "#FF0000" }"""), out _, conflict));

        foreach (var key in new[]
                 {
                     ExcelStyleContract.ShrinkToFit, ExcelStyleContract.Underline, ExcelStyleContract.Strikethrough,
                     ExcelStyleContract.Indent, ExcelStyleContract.Orientation, ExcelStyleContract.Locked,
                     ExcelStyleContract.FillPattern, ExcelStyleContract.NoFill,
                 })
            Assert.False(ExcelAdapter.IsDeferredWrittenStyleKey(key));
        Assert.True(ExcelAdapter.IsDeferredWrittenStyleKey(ExcelStyleContract.Bold));
    }

    [Fact]
    public void Policy_wires_data_ops_and_delete_sheet_is_high_risk()
    {
        var policy = new PolicyEngine();
        Assert.Equal(OpClass.HighRisk, policy.ClassifyOp("excel", "delete_sheet"));
        Assert.Equal(OpClass.Allowed, policy.ClassifyOp("excel", "close_workbook"));
        Assert.Equal(OpClass.Allowed, policy.ClassifyOp("excel", "create_table"));
        Assert.True(ExcelDataOperationsContract.WriteOpNames.Count >= 39);
        Assert.Equal(2, ExcelDataOperationsContract.SnapshotVersion);
        Assert.Equal(OpClass.Allowed, policy.ClassifyOp("excel", "set_view"));
        Assert.False(policy.IsAutoExecutable("excel", "merge_cells"));
        Assert.True(policy.IsAutoExecutable("excel", "set_values"));
    }

    [Fact]
    public void Public_schema_publishes_rows_page_output_and_inspect_objectKind()
    {
        var props = ExcelAuthoringSchema.DescribeApplyOpProperties();
        Assert.NotNull(Json.GetObj(props, "rows"));
        Assert.NotNull(Json.GetObj(props, "page"));
        Assert.NotNull(Json.GetObj(props, "output"));
        Assert.NotNull(Json.GetObj(props, "path"));
        Assert.NotNull(Json.GetObj(props, "options"));
        Assert.NotNull(Json.GetObj(props, "view"));
        Assert.NotNull(Json.GetObj(props, "formula2"));
        var rowItem = Json.GetObj(Json.GetObj(props, "rows"), "items");
        Assert.Equal("object", Json.GetString(rowItem, "type"));
        Assert.NotNull(Json.GetObj(Json.GetObj(rowItem, "properties"), "row"));
        Assert.NotNull(Json.GetObj(Json.GetObj(rowItem, "properties"), "heightPoints"));
        var colItem = Json.GetObj(Json.GetObj(props, "columns"), "items");
        Assert.NotNull(Json.GetObj(Json.GetObj(colItem, "properties"), "col"));
        Assert.NotNull(Json.GetObj(Json.GetObj(colItem, "properties"), "widthChars"));
        var schema = ExcelDataOperationsContract.DescribeSchema();
        Assert.Equal(2, Json.GetInt(schema, "snapshotVersion"));
        Assert.True((Json.GetArr(schema, "writeOps")?.Count ?? 0) >= 39);
        var scopes = ExcelAuthoringSchema.InspectScopeEnum().Select(node => node!.GetValue<string>()).ToHashSet();
        Assert.Contains("tables", scopes);
        Assert.Contains("objects", scopes);
    }

    [Fact]
    public void Protected_source_is_not_hardcoded()
    {
        Environment.SetEnvironmentVariable(ExcelAuthoringPaths.ProtectedSourcesVariable, null);
        Assert.False(ExcelAuthoringPaths.IsProtectedSource(
            @"C:\DocBridgeTest\Desktop\AI작업\10.만회대책\수원시_하수관로정비_공사예정공정표(변경).xlsx"));
        Environment.SetEnvironmentVariable(ExcelAuthoringPaths.ProtectedSourcesVariable, @"C:\tmp\protected.xlsx");
        try
        {
            Assert.True(ExcelAuthoringPaths.IsProtectedSource(@"C:\tmp\protected.xlsx"));
        }
        finally
        {
            Environment.SetEnvironmentVariable(ExcelAuthoringPaths.ProtectedSourcesVariable, null);
        }
    }

    [Fact]
    public void Create_or_open_only_excludes_close_and_close_requires_explicit_target()
    {
        Assert.True(ExcelAdapter.IsCreateOrOpenOnly(new[] { new JsonObject { ["op"] = "create_workbook" } }));
        Assert.True(ExcelAdapter.IsCreateOrOpenOnly(new[] { new JsonObject { ["op"] = "open_workbook" } }));
        Assert.False(ExcelAdapter.IsCreateOrOpenOnly(new[] { new JsonObject { ["op"] = "close_workbook" } }));
        Assert.True(ExcelAdapter.IsCloseOnly(new[] { new JsonObject { ["op"] = "close_workbook" } }));

        var validator = new OperationValidator(new PolicyEngine());
        var missing = Json.ParseObject("""{ "ops": [ { "op": "close_workbook" } ], "dryRun": true }""");
        var missingErrors = new List<string>();
        Assert.Null(validator.Validate(missing, "excel", missingErrors));
        Assert.Contains(missingErrors, error => error.Contains("target.workbook", StringComparison.OrdinalIgnoreCase));

        var ok = Json.ParseObject("""
        { "ops": [ { "op": "close_workbook", "target": { "workbook": "C:\\\\out\\\\schedule.xlsx" } } ], "dryRun": true }
        """);
        var okErrors = new List<string>();
        Assert.NotNull(validator.Validate(ok, "excel", okErrors));
        Assert.Empty(okErrors);
    }

    [Fact]
    public void Owned_workbooks_append_exact_identities_not_inventory_diff()
    {
        var state = new JsonObject
        {
            ["preexistingWorkbooks"] = new JsonArray(
                new JsonObject { ["name"] = "수원시.xlsx", ["fullName"] = @"C:\work\수원시.xlsx" }),
            ["ownedWorkbooks"] = new JsonArray(),
        };
        ExcelAuthoringSchema.AppendOwnedWorkbook(state, "Book2", "Book2", ExcelAuthoringSchema.AddOwnershipSource);
        ExcelAuthoringSchema.AppendOwnedWorkbook(state, "Book2", "Book2", ExcelAuthoringSchema.AddOwnershipSource);
        var owned = Json.GetArr(state, "ownedWorkbooks")!;
        Assert.Single(owned);
        Assert.Equal("Workbooks.Add", Json.GetString((JsonObject)owned[0]!, "source"));
        Assert.False(ExcelAuthoringSchema.ShouldCloseOnLifecycleRestore("수원시.xlsx", @"C:\work\수원시.xlsx", owned));
        Assert.False(ExcelAuthoringSchema.ShouldCloseOnLifecycleRestore("통합 문서1", "통합 문서1", owned));
        Assert.True(ExcelAuthoringSchema.ShouldCloseOnLifecycleRestore("Book2", "Book2", owned));
    }

    [Fact]
    public void Launch_ownership_flags_are_honest_and_open_path_is_not_writable()
    {
        var attached = ExcelAuthoringSchema.DescribeLaunchOwnership(
            ownsInstance: false, createdWorkbook: true, dedicatedRequested: false);
        Assert.False(Json.GetBool(attached, "ownedInstance"));
        Assert.False(Json.GetBool(attached, "createdInstance"));
        Assert.True(Json.GetBool(attached, "attachedExisting"));
        Assert.True(Json.GetBool(attached, "createdWorkbook"));

        var dedicated = ExcelAuthoringSchema.DescribeLaunchOwnership(
            ownsInstance: true, createdWorkbook: true, dedicatedRequested: true);
        Assert.True(Json.GetBool(dedicated, "ownedInstance"));
        Assert.True(Json.GetBool(dedicated, "createdInstance"));
        Assert.False(Json.GetBool(dedicated, "attachedExisting"));

        var open = new JsonObject
        {
            ["op"] = "open_workbook",
            ["path"] = @"C:\work\source.xlsx",
            ["role"] = ExcelAuthoringSchema.OpenSourceRole,
            ["writable"] = false,
        };
        Assert.False(ExcelAuthoringSchema.IsWritableLifecycleOutput(open));
        Assert.True(ExcelAuthoringSchema.IsWritableLifecycleOutput(new JsonObject
        {
            ["op"] = "save_workbook",
            ["path"] = @"C:\out\schedule.xlsx",
            ["writable"] = true,
        }));
    }

    [Fact]
    public void Same_path_save_capture_and_save_helpers()
    {
        var path = @"C:\out\schedule.xlsx";
        Assert.True(ExcelAuthoringSchema.IsSamePathSave(path, path));
        Assert.False(ExcelAuthoringSchema.IsSamePathSave(path, "Book1"));
        Assert.Equal(path, ExcelAuthoringSchema.ResolveSaveCapturePath(null, path));
        Assert.Null(ExcelAuthoringSchema.ResolveSaveCapturePath(null, "Book1"));
        Assert.Equal(@"C:\out\other.xlsx", ExcelAuthoringSchema.ResolveSaveCapturePath(@"C:\out\other.xlsx", path));
    }

    [Fact]
    public void Apply_input_schema_attaches_data_anyOf_and_admits_table_chart_pivot_layout()
    {
        var items = new JsonObject
        {
            ["type"] = "object",
            ["properties"] = new JsonObject
            {
                ["op"] = new JsonObject
                {
                    ["type"] = "string",
                    ["enum"] = new JsonArray("set_values", "move_sheet", "set_row_heights", "create_table", "create_chart", "create_pivot"),
                },
                ["position"] = new JsonObject { ["type"] = "string" },
                ["rows"] = new JsonObject { ["type"] = "array" },
                ["columns"] = new JsonObject { ["type"] = "array" },
                ["values"] = new JsonObject { ["type"] = "array" },
                ["target"] = new JsonObject { ["type"] = "object" },
                ["range"] = new JsonObject { ["type"] = "string" },
            },
        };
        ExcelDataOperationSchema.AttachToApplyItems(items);
        Assert.Null(Json.GetObj(items, "properties"));
        Assert.True((Json.GetArr(items, "anyOf")?.Count ?? 0) >= 40);
        var leftover = Json.GetArr(items, "anyOf")!.OfType<JsonObject>().First(branch =>
            Json.GetArr(Json.GetObj(Json.GetObj(branch, "properties"), "op"), "enum") is not null);
        Assert.Equal("string", Json.GetString(Json.GetObj(Json.GetObj(leftover, "properties"), "position"), "type"));

        Assert.Empty(ExcelDataOperationSchema.Validate(Json.ParseObject("""
            { "op": "create_table", "range": "A1:D8", "name": "자재표", "hasHeaders": true, "target": { "sheet": "자재대장" } }
            """)!));
        Assert.Empty(ExcelDataOperationSchema.Validate(Json.ParseObject("""
            { "op": "create_chart", "sourceRange": "A1:D8", "chartType": "columnClustered",
              "position": { "left": 200, "top": 40, "width": 360, "height": 220 },
              "target": { "sheet": "자재대장" } }
            """)!));
        Assert.Empty(ExcelDataOperationSchema.Validate(Json.ParseObject("""
            { "op": "create_pivot", "sourceRange": "A1:D8", "destination": "F1", "name": "원가피벗",
              "rows": ["공종"], "columns": ["월"], "values": [{ "field": "금액", "function": "sum" }],
              "target": { "sheet": "원가" } }
            """)!));
    }

    [Fact]
    public void Set_view_and_protect_options_validate_without_password()
    {
        var validator = new OperationValidator(new PolicyEngine());
        var view = Json.ParseObject("""
        {
          "ops": [ { "op": "set_view", "target": { "sheet": "예정공정표" }, "zoom": 80, "displayGridlines": false } ],
          "dryRun": true
        }
        """);
        var viewErrors = new List<string>();
        Assert.NotNull(validator.Validate(view, "excel", viewErrors));
        Assert.Empty(viewErrors);

        var protect = Json.ParseObject("""
        {
          "ops": [ { "op": "protect_sheet", "target": { "sheet": "예정공정표" }, "options": { "allowFormattingCells": true, "allowFiltering": true } } ],
          "dryRun": true
        }
        """);
        var protectErrors = new List<string>();
        Assert.NotNull(validator.Validate(protect, "excel", protectErrors));
        Assert.Empty(protectErrors);
    }
}
