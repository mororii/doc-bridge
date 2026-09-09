using System.Collections;
using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelFormatOnlySafetyTests
{
    [Fact]
    public void Format_only_snapshot_captures_every_area_and_skips_used_range()
    {
        using var home = new TestHome();
        var workbook = new FormatWorkbook(@"C:\docbridge-fixtures\format-only-multi-area.xlsx", "Sheet1");
        var a1 = workbook.Sheet.Cell(1, 1);
        a1.Font.Bold = true;
        var c3 = workbook.Sheet.Cell(3, 3);
        c3.Font.Italic = true;
        using var adapter = CreateAdapter(workbook);

        adapter.CaptureSnapshot(CreateSnapshotDir(home), new JsonObject(), FormatOps("A1,C3"));

        var state = ReadState(home);
        Assert.Equal("format-only", Json.GetString(state, "restoreMode"));
        Assert.Equal(ExcelAdapter.FormatOnlyScopedSnapshotVersion, Json.GetInt(state, "snapshotVersion"));
        Assert.Equal(ExcelAdapter.WrittenStyleScope, Json.GetString(state, "styleScope"));
        Assert.Equal(0, workbook.UsedRangeAccessCount);
        var formatStates = Json.GetArr(state, "formatStates")!;
        Assert.Equal(2, formatStates.Count);
        Assert.Equal("A1", Json.GetString((JsonObject)formatStates[0]!, "range"));
        Assert.Equal("C3", Json.GetString((JsonObject)formatStates[1]!, "range"));
        var a1Style = FirstCapturedStyle((JsonObject)formatStates[0]!);
        var c3Style = FirstCapturedStyle((JsonObject)formatStates[1]!);
        Assert.True(Json.GetBool(a1Style, "bold"));
        Assert.False(Json.GetBool(c3Style, "bold"));
        Assert.False(a1Style.ContainsKey("italic"));
        Assert.False(c3Style.ContainsKey("italic"));
        Assert.False(a1Style.ContainsKey("fontTintAndShade"));
        Assert.False(a1Style.ContainsKey("fillColor"));
        Assert.Equal(2, Json.GetInt(Json.GetObj(state, "coverage"), "cellCount"));
        Assert.True(Json.GetBool(Json.GetObj(state, "coverage"), "complete"));
    }

    [Fact]
    public void Partial_merged_coverage_is_rejected_before_snapshot()
    {
        using var home = new TestHome();
        var workbook = new FormatWorkbook(@"C:\docbridge-fixtures\format-only-partial-merge.xlsx", "Sheet1");
        workbook.Sheet.Merge(1, 1, 2, 2);
        using var adapter = CreateAdapter(workbook);

        var ex = CaptureSnapshotError(adapter, home, "A1:B1");
        Assert.Contains("[EXCEL_FORMAT_PARTIAL_MERGE]", ex.Message);
        Assert.DoesNotContain("false", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(0, workbook.MultiCellMergeAreaReads);
    }

    [Fact]
    public void Multiple_in_target_merges_are_captured_and_an_overflowing_second_merge_is_rejected()
    {
        using var home = new TestHome();
        var inside = new FormatWorkbook(@"C:\docbridge-fixtures\format-only-multi-merge-ok.xlsx", "Sheet1");
        inside.Sheet.Merge(1, 1, 1, 2);
        inside.Sheet.Merge(2, 1, 1, 2);
        using (var adapter = CreateAdapter(inside))
        {
            adapter.CaptureSnapshot(CreateSnapshotDir(home), new JsonObject(), FormatOps("A1:B2"));
        }

        Assert.Equal(0, inside.MultiCellMergeAreaReads);
        var state = ReadState(home);
        Assert.Equal(4, Json.GetInt(Json.GetObj(state, "coverage"), "cellCount"));

        using var overflowHome = new TestHome();
        var overflow = new FormatWorkbook(@"C:\docbridge-fixtures\format-only-multi-merge-overflow.xlsx", "Sheet1");
        overflow.Sheet.Merge(1, 1, 1, 2);
        overflow.Sheet.Merge(2, 1, 1, 3);
        using var overflowAdapter = CreateAdapter(overflow);
        var ex = CaptureSnapshotError(overflowAdapter, overflowHome, "A1:B2");
        Assert.Contains("[EXCEL_FORMAT_PARTIAL_MERGE]", ex.Message);
        Assert.Equal(0, overflow.MultiCellMergeAreaReads);
    }

    [Fact]
    public void Mixed_rich_text_font_nulls_are_not_coerced()
    {
        using var home = new TestHome();
        var workbook = new FormatWorkbook(@"C:\docbridge-fixtures\format-only-rich-text.xlsx", "Sheet1");
        workbook.Sheet.Cell(1, 1).Font.Bold = DBNull.Value;
        using var adapter = CreateAdapter(workbook);

        var ex = CaptureSnapshotError(adapter, home, "A1");
        Assert.Contains("[EXCEL_FORMAT_MIXED_RICHTEXT]", ex.Message);
        Assert.Contains("Font.Bold", ex.Message);
    }

    [Fact]
    public void Restore_rejects_missing_and_malformed_format_state_instead_of_verifying_zero_cells()
    {
        using var home = new TestHome();
        var workbook = new FormatWorkbook(@"C:\docbridge-fixtures\format-only-malformed.xlsx", "Sheet1");
        using var adapter = CreateAdapter(workbook);
        var snapshotDir = CreateSnapshotDir(home);

        WriteState(snapshotDir, new JsonObject
        {
            ["snapshotVersion"] = 2,
            ["restoreMode"] = "format-only",
            ["documentRef"] = workbook.FullName,
            ["coverage"] = new JsonObject { ["complete"] = true, ["cellCount"] = 1 },
        });
        var missingStates = adapter.RestoreSnapshot(snapshotDir, new JsonObject());
        Assert.False(Json.GetBool(missingStates, "ok"));
        Assert.Contains("formatStates", string.Join(" ", Json.GetArr(missingStates, "errors")!.Select(n => n!.GetValue<string>())));

        WriteState(snapshotDir, new JsonObject
        {
            ["snapshotVersion"] = 2,
            ["restoreMode"] = "format-only",
            ["documentRef"] = workbook.FullName,
            ["formatStates"] = new JsonArray(),
            ["coverage"] = new JsonObject { ["complete"] = true, ["cellCount"] = 0 },
        });
        var emptyStates = adapter.RestoreSnapshot(snapshotDir, new JsonObject());
        Assert.False(Json.GetBool(emptyStates, "ok"));
        Assert.Contains("empty", string.Join(" ", Json.GetArr(emptyStates, "errors")!.Select(n => n!.GetValue<string>())));

        WriteState(snapshotDir, new JsonObject
        {
            ["snapshotVersion"] = 2,
            ["restoreMode"] = "format-only",
            ["documentRef"] = workbook.FullName,
            ["formatStates"] = new JsonArray(ValidFormatState()),
        });
        var missingCoverage = adapter.RestoreSnapshot(snapshotDir, new JsonObject());
        Assert.False(Json.GetBool(missingCoverage, "ok"));
        Assert.Contains("coverage", string.Join(" ", Json.GetArr(missingCoverage, "errors")!.Select(n => n!.GetValue<string>())));

        var badDimensions = ValidFormatState();
        badDimensions["rows"] = 2;
        WriteState(snapshotDir, new JsonObject
        {
            ["snapshotVersion"] = 2,
            ["restoreMode"] = "format-only",
            ["documentRef"] = workbook.FullName,
            ["formatStates"] = new JsonArray(badDimensions),
            ["coverage"] = new JsonObject { ["complete"] = true, ["cellCount"] = 1 },
        });
        var dimensionMismatch = adapter.RestoreSnapshot(snapshotDir, new JsonObject());
        Assert.False(Json.GetBool(dimensionMismatch, "ok"));
        Assert.Contains("styles rows", string.Join(" ", Json.GetArr(dimensionMismatch, "errors")!.Select(n => n!.GetValue<string>())));
        Assert.Null(Json.GetObj(dimensionMismatch, "readback"));

        WriteState(snapshotDir, new JsonObject
        {
            ["snapshotVersion"] = 2,
            ["restoreMode"] = "format-only",
            ["documentRef"] = workbook.FullName,
            ["formatStates"] = new JsonArray(ValidFormatState()),
            ["coverage"] = new JsonObject { ["complete"] = false, ["cellCount"] = 1 },
        });
        var incomplete = adapter.RestoreSnapshot(snapshotDir, new JsonObject());
        Assert.False(Json.GetBool(incomplete, "ok"));
        Assert.Contains("incomplete", string.Join(" ", Json.GetArr(incomplete, "errors")!.Select(n => n!.GetValue<string>())));
    }

    [Fact]
    public void Pattern_color_is_restored_and_verified()
    {
        using var home = new TestHome();
        var workbook = new FormatWorkbook(@"C:\docbridge-fixtures\format-only-pattern-color.xlsx", "Sheet1");
        var cell = workbook.Sheet.Cell(1, 1);
        cell.Interior.Pattern = 2;
        cell.Interior.PatternColor = 65535d;
        cell.Interior.PatternColorIndex = 6;
        using var adapter = CreateAdapter(workbook);
        var snapshotDir = CreateSnapshotDir(home);
        adapter.CaptureSnapshot(snapshotDir, new JsonObject(), FormatOps("A1", new JsonObject { ["fillColor"] = 255 }));

        var captured = FirstCapturedStyle((JsonObject)Json.GetArr(ReadState(home), "formatStates")![0]!);
        Assert.Equal(65535d, captured["fillPatternColor"]!.GetValue<double>());

        cell.Interior.PatternColor = 255d;
        cell.Interior.Pattern = 1;
        var restored = adapter.RestoreSnapshot(snapshotDir, new JsonObject());
        Assert.True(Json.GetBool(restored, "ok"), restored.ToJsonString());
        Assert.True(Json.GetBool(Json.GetObj(restored, "readback"), "verified"));
        Assert.Equal(65535d, cell.Interior.PatternColor);
        Assert.Equal(2, cell.Interior.Pattern);
    }

    [Fact]
    public void Theme_and_automatic_colors_are_preserved_instead_of_rgb_only()
    {
        using var home = new TestHome();
        var workbook = new FormatWorkbook(@"C:\docbridge-fixtures\format-only-theme-auto.xlsx", "Sheet1");
        var themed = workbook.Sheet.Cell(1, 1);
        themed.Font.ThemeColor = 5;
        themed.Font.TintAndShade = 0.25d;
        themed.Font.Color = 999d;
        var automatic = workbook.Sheet.Cell(1, 2);
        automatic.Font.ColorIndex = ExcelStyleContract.XlColorIndexAutomatic;
        automatic.Font.Color = 0d;
        using var adapter = CreateAdapter(workbook);
        var snapshotDir = CreateSnapshotDir(home);
        adapter.CaptureSnapshot(snapshotDir, new JsonObject(), FormatOps("A1:B1", new JsonObject { ["fontColor"] = 0 }));

        var states = Json.GetArr(ReadState(home), "formatStates")!;
        var captured = (JsonObject)states[0]!;
        Assert.Equal(ExcelAdapter.FormatStyleModeGroups, Json.GetString(captured, "styleMode"));
        var a1 = StyleAt(captured, 1, 1);
        var b1 = StyleAt(captured, 1, 2);
        Assert.Equal(5, a1["fontThemeColor"]!.GetValue<int>());
        Assert.Equal(0.25d, a1["fontTintAndShade"]!.GetValue<double>());
        Assert.Equal(ExcelStyleContract.XlColorIndexAutomatic, b1["fontColorIndex"]!.GetValue<int>());

        themed.Font.ThemeColor = 1;
        themed.Font.TintAndShade = 0d;
        themed.Font.Color = 0d;
        themed.Font.ColorAssignments = 0;
        automatic.Font.ColorIndex = 1;
        automatic.Font.Color = 255d;
        automatic.Font.ColorAssignments = 0;

        var restored = adapter.RestoreSnapshot(snapshotDir, new JsonObject());
        Assert.True(Json.GetBool(restored, "ok"), restored.ToJsonString());
        Assert.Equal(5, themed.Font.ThemeColor);
        Assert.Equal(0.25d, themed.Font.TintAndShade);
        Assert.Equal(0, themed.Font.ColorAssignments);
        Assert.Equal(ExcelStyleContract.XlColorIndexAutomatic, automatic.Font.ColorIndex);
        Assert.Equal(0, automatic.Font.ColorAssignments);
    }

    [Fact]
    public void Theme_color_disconnected_com_fails_snapshot_before_write()
    {
        using var home = new TestHome();
        var workbook = new FormatWorkbook(@"C:\docbridge-fixtures\format-theme-disconnected.xlsx", "Sheet1");
        workbook.Sheet.Cell(1, 1).Font.ThemeColorFault = new COMException(
            "The object invoked has disconnected from its clients.",
            unchecked((int)0x80010108));
        using var adapter = CreateAdapter(workbook);

        var ex = CaptureSnapshotError(adapter, home, "A1", new JsonObject { ["fontColor"] = 0 });
        Assert.Contains("[EXCEL_FORMAT_THEME_READ_FAILED]", ex.Message);
        Assert.Contains("0x80010108", ex.Message);
        Assert.DoesNotContain("false", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(Path.Combine(home.Dir, "snapshot", "state.json")));
    }

    [Fact]
    public void Non_theme_rgb_tint_fails_format_only_snapshot_before_write()
    {
        using var home = new TestHome();
        var workbook = new FormatWorkbook(@"C:\docbridge-fixtures\format-rgb-tint.xlsx", "Sheet1");
        var target = workbook.Sheet.Cell(3, 3);
        target.Interior.Color = 42662d;
        target.Interior.TintAndShade = -0.3499862666707358d;
        var bystander = workbook.Sheet.Cell(1, 1);
        bystander.Font.Bold = true;
        bystander.Interior.Color = 255d;
        using var adapter = CreateAdapter(workbook);

        var ex = CaptureSnapshotError(adapter, home, "C3", new JsonObject { ["fillColor"] = 42662 });
        Assert.Contains("[EXCEL_FORMAT_UNSUPPORTED_RGB_TINT]", ex.Message);
        Assert.Contains("C3", ex.Message);
        Assert.Contains("Interior.TintAndShade", ex.Message);
        Assert.False(File.Exists(Path.Combine(home.Dir, "snapshot", "state.json")));
        Assert.Equal(42662d, target.Interior.Color);
        Assert.Equal(-0.3499862666707358d, target.Interior.TintAndShade);
        Assert.True((bool)bystander.Font.Bold!);
        Assert.Equal(255d, bystander.Interior.Color);
        Assert.Equal(0, workbook.UsedRangeAccessCount);
    }

    [Fact]
    public void Format_only_fingerprint_is_reusable_until_target_style_changes()
    {
        using var home = new TestHome();
        var workbook = new FormatWorkbook(@"C:\docbridge-fixtures\format-only-fingerprint.xlsx", "Sheet1");
        workbook.Sheet.Cell(1, 1).Font.Bold = true;
        using var adapter = CreateAdapter(workbook);
        var snapshotDir = CreateSnapshotDir(home);
        var metadata = new JsonObject();
        var ops = FormatOps("A1");
        adapter.CaptureSnapshot(snapshotDir, metadata, ops);

        var matched = adapter.ValidatePreviewReuse(snapshotDir, metadata, ops);
        Assert.True(Json.GetBool(matched, "ok"));
        Assert.True(Json.GetBool(matched, "reusable"));
        Assert.False(Json.GetBool(matched, "freshPreviewAllowed"));
        Assert.Equal("excel-format-target-sha256", Json.GetString(matched, "fingerprintMethod"));

        workbook.Sheet.Cell(1, 1).Font.Bold = false;
        var changed = adapter.ValidatePreviewReuse(snapshotDir, metadata, ops);
        Assert.True(Json.GetBool(changed, "ok"));
        Assert.False(Json.GetBool(changed, "reusable"));
        Assert.False(Json.GetBool(changed, "freshPreviewAllowed"));
    }

    [Fact]
    public void Non_format_batches_request_a_fresh_preview_instead_of_a_fingerprint_deny()
    {
        using var home = new TestHome();
        var workbook = new FormatWorkbook(@"C:\docbridge-fixtures\format-only-nonformat.xlsx", "Sheet1");
        using var adapter = CreateAdapter(workbook);
        var snapshotDir = CreateSnapshotDir(home);
        adapter.CaptureSnapshot(snapshotDir, new JsonObject(), FormatOps("A1"));

        var validation = adapter.ValidatePreviewReuse(snapshotDir, new JsonObject(), new[]
        {
            new JsonObject
            {
                ["op"] = "set_values",
                ["target"] = new JsonObject { ["sheet"] = "Sheet1" },
                ["range"] = "A1",
                ["values"] = new JsonArray(new JsonArray("1")),
            },
        });
        Assert.False(Json.GetBool(validation, "reusable"));
        Assert.True(Json.GetBool(validation, "freshPreviewAllowed"));
        Assert.Equal("excel-not-fingerprinted", Json.GetString(validation, "fingerprintMethod"));
    }

    [Fact]
    public void Happy_path_restores_direct_rgb_without_touching_used_range()
    {
        using var home = new TestHome();
        var workbook = new FormatWorkbook(@"C:\docbridge-fixtures\format-only-happy.xlsx", "Sheet1");
        var cell = workbook.Sheet.Cell(2, 2);
        cell.Font.Bold = true;
        cell.Font.Color = 255d;
        cell.Interior.Color = 128d;
        using var adapter = CreateAdapter(workbook);
        var snapshotDir = CreateSnapshotDir(home);
        adapter.CaptureSnapshot(snapshotDir, new JsonObject(), FormatOps("B2", new JsonObject
        {
            ["bold"] = true,
            ["fontColor"] = 255,
            ["fillColor"] = 128,
        }));
        Assert.Equal(0, workbook.UsedRangeAccessCount);

        cell.Font.Bold = false;
        cell.Font.Color = 0d;
        cell.Interior.Color = 0d;
        var restored = adapter.RestoreSnapshot(snapshotDir, new JsonObject());
        Assert.True(Json.GetBool(restored, "ok"), restored.ToJsonString());
        Assert.True((bool)cell.Font.Bold!);
        Assert.Equal(255d, cell.Font.Color);
        Assert.Equal(128d, cell.Interior.Color);
        Assert.Equal(0, workbook.UsedRangeAccessCount);
    }

    [Fact]
    public void Apply_records_per_op_hresult_and_skips_remaining_ops()
    {
        var workbook = new FormatWorkbook(@"C:\docbridge-fixtures\format-apply-diagnostics.xlsx", "Sheet1");
        workbook.Sheet.RangeFaultAddress = "B1";
        workbook.Sheet.RangeFault = new COMException("Excel could not apply the format", unchecked((int)0x800A03EC));
        using var adapter = CreateAdapter(workbook);

        var exec = adapter.Apply(new[]
        {
            FormatOp("A1"),
            FormatOp("B1"),
            FormatOp("C1"),
        }, "snap-apply-diagnostics");

        Assert.False(exec.Ok);
        Assert.Equal(3, exec.OperationResults.Count);

        var first = exec.OperationResults[0]!.AsObject();
        Assert.Equal("format_range", Json.GetString(first, "op"));
        Assert.True(Json.GetBool(first, "ok"));
        Assert.Equal("apply", Json.GetString(first, "stage"));
        Assert.Equal("per-op", Json.GetString(first, "timingScope"));
        Assert.True((bool)workbook.Sheet.Cell(1, 1).Font.Bold!);

        var failed = exec.OperationResults[1]!.AsObject();
        Assert.False(Json.GetBool(failed, "ok"));
        Assert.Equal("apply", Json.GetString(failed, "stage"));
        Assert.Equal("0x800A03EC", Json.GetString(failed, "hresult"));
        Assert.Equal(unchecked((int)0x800A03EC), Json.GetInt(failed, "hresultValue"));
        Assert.Contains("COMException", Json.GetString(failed, "exceptionType"));
        Assert.Contains("0x800A03EC", string.Join(" ", exec.Errors));

        var skipped = exec.OperationResults[2]!.AsObject();
        Assert.False(Json.GetBool(skipped, "ok"));
        Assert.Equal("skipped", Json.GetString(skipped, "stage"));
        Assert.False((bool)workbook.Sheet.Cell(1, 3).Font.Bold!);
    }

    [Fact]
    public void Preview_format_range_uses_real_cell_style_not_current_string()
    {
        var workbook = new FormatWorkbook(@"C:\docbridge-fixtures\format-preview-before.xlsx", "Sheet1");
        var black = workbook.Sheet.Cell(1, 1);
        black.Font.Bold = true;
        black.Interior.Color = 0d;
        black.Interior.Pattern = ExcelStyleContract.XlPatternSolid;
        var empty = workbook.Sheet.Cell(1, 2);
        empty.Interior.Color = 0d;
        empty.Interior.Pattern = ExcelStyleContract.XlPatternNone;
        using var adapter = CreateAdapter(workbook);

        var preview = adapter.Preview(new[] { FormatOp("A1") });
        Assert.Empty(preview.Errors);
        Assert.Single(preview.Diff);
        var before = Assert.IsType<JsonObject>(preview.Diff[0].Before);
        Assert.NotEqual("current", before.ToJsonString());
        Assert.True(Json.GetBool(before, "bold"));
        Assert.True(Json.GetBool(before, "fontBold"));
        Assert.Equal(0d, before["fillColor"]!.GetValue<double>());
        Assert.Equal(0d, before["interiorColor"]!.GetValue<double>());
        Assert.Equal(ExcelStyleContract.XlPatternSolid, Json.GetInt(before, "fillPattern"));
        var after = Assert.IsType<JsonObject>(preview.Diff[0].After);
        Assert.True(Json.GetBool(after, "bold"));

        var noFill = adapter.Preview(new[] { FormatOp("B1") });
        var noFillBefore = Assert.IsType<JsonObject>(noFill.Diff[0].Before);
        Assert.Equal(0d, noFillBefore["fillColor"]!.GetValue<double>());
        Assert.Equal(ExcelStyleContract.XlPatternNone, Json.GetInt(noFillBefore, "fillPattern"));
        Assert.NotEqual(Json.GetInt(before, "fillPattern"), Json.GetInt(noFillBefore, "fillPattern"));
    }

    [Fact]
    public void Read_includeStyles_emits_both_alias_families_and_fill_pattern()
    {
        var workbook = new FormatWorkbook(@"C:\docbridge-fixtures\format-read-styles.xlsx", "Sheet1");
        var cell = workbook.Sheet.Cell(2, 2);
        cell.Font.Italic = true;
        cell.Interior.Color = 255d;
        cell.Interior.Pattern = ExcelStyleContract.XlPatternSolid;
        using var adapter = CreateAdapter(workbook);

        var read = adapter.Read(new JsonObject
        {
            ["sheet"] = "Sheet1",
            ["range"] = "B2",
            ["includeStyles"] = true,
        });
        Assert.True(Json.GetBool(read, "ok"), read.ToJsonString());
        var styles = Json.GetObj(read, "styles")!;
        Assert.True(Json.GetBool(styles, "italic"));
        Assert.True(Json.GetBool(styles, "fontItalic"));
        Assert.Equal(255d, styles["fillColor"]!.GetValue<double>());
        Assert.Equal(255d, styles["interiorColor"]!.GetValue<double>());
        Assert.Equal(255d, styles["fill"]!.GetValue<double>());
        Assert.Equal(ExcelStyleContract.XlPatternSolid, Json.GetInt(styles, "fillPattern"));
        Assert.Equal(1, Json.GetInt(styles, "fillColorIndex"));
    }

    private static Exception CaptureSnapshotError(ExcelAdapter adapter, TestHome home, string range, JsonObject? style = null)
    {
        var ex = Assert.ThrowsAny<Exception>(() =>
            adapter.CaptureSnapshot(CreateSnapshotDir(home), new JsonObject(), FormatOps(range, style)));
        return ex is AggregateException aggregate
            ? aggregate.Flatten().InnerExceptions[0]
            : ex;
    }

    private static ExcelAdapter CreateAdapter(FormatWorkbook workbook) =>
        new(() => new FormatExcelApplication(workbook));

    private static string CreateSnapshotDir(TestHome home)
    {
        var dir = Path.Combine(home.Dir, "snapshot");
        Directory.CreateDirectory(dir);
        return dir;
    }

    internal static IReadOnlyList<JsonObject> FormatOps(string range, JsonObject? style = null) =>
        new[] { FormatOp(range, style) };

    internal static JsonObject FormatOp(string range, JsonObject? style = null) =>
        new()
        {
            ["op"] = "format_range",
            ["target"] = new JsonObject { ["sheet"] = "Sheet1" },
            ["range"] = range,
            ["style"] = style ?? new JsonObject { ["bold"] = true },
        };

    internal static JsonObject FirstCapturedStyle(JsonObject formatState)
    {
        if (formatState["style"] is JsonObject uniform)
            return uniform;
        if (formatState["groups"] is JsonArray groups && groups.Count > 0)
            return ((JsonObject)groups[0]!).ContainsKey("style")
                ? (JsonObject)((JsonObject)groups[0]!)["style"]!
                : ((JsonObject)groups[0]!);
        return ((JsonArray)((JsonArray)formatState["styles"]!)[0]!)[0]!.AsObject();
    }

    internal static JsonObject StyleAt(JsonObject formatState, int row, int column)
    {
        if (formatState["groups"] is JsonArray groups)
        {
            foreach (var node in groups)
            {
                var group = (JsonObject)node!;
                var startRow = Json.GetInt(group, "row") ?? 0;
                var startCol = Json.GetInt(group, "column") ?? 0;
                var rows = Json.GetInt(group, "rows") ?? 0;
                var cols = Json.GetInt(group, "columns") ?? 0;
                if (row >= startRow && row < startRow + rows && column >= startCol && column < startCol + cols)
                    return (JsonObject)group["style"]!;
            }
        }

        if (formatState["style"] is JsonObject uniform)
            return uniform;
        var start = Json.GetInt(formatState, "row") ?? 1;
        var startColumn = Json.GetInt(formatState, "column") ?? 1;
        return ((JsonArray)((JsonArray)formatState["styles"]!)[row - start]!)[column - startColumn]!.AsObject();
    }

    private static JsonObject ReadState(TestHome home) =>
        JsonNode.Parse(File.ReadAllText(Path.Combine(home.Dir, "snapshot", "state.json")))!.AsObject();

    private static void WriteState(string snapshotDir, JsonObject state)
    {
        Directory.CreateDirectory(snapshotDir);
        File.WriteAllText(Path.Combine(snapshotDir, "state.json"), state.ToJsonString());
    }

    private static JsonObject ValidFormatState() => new()
    {
        ["sheet"] = "Sheet1",
        ["range"] = "A1",
        ["row"] = 1,
        ["column"] = 1,
        ["rows"] = 1,
        ["columns"] = 1,
        ["styles"] = new JsonArray(new JsonArray(new JsonObject
        {
            ["bold"] = false,
            ["italic"] = false,
            ["fontSize"] = 11,
            ["numberFormat"] = "General",
            ["fontColor"] = 0,
            ["fontColorIndex"] = 1,
            ["fontTintAndShade"] = 0,
            ["fillColor"] = 16777215,
            ["fillColorIndex"] = ExcelStyleContract.XlColorIndexNone,
            ["fillTintAndShade"] = 0,
            ["fillPattern"] = ExcelStyleContract.XlPatternNone,
            ["fillPatternColor"] = 0,
            ["fillPatternColorIndex"] = ExcelStyleContract.XlColorIndexNone,
            ["fillPatternTintAndShade"] = 0,
        })),
    };

    public sealed class FormatExcelApplication
    {
        public FormatExcelApplication(FormatWorkbook workbook)
        {
            ActiveWorkbook = workbook;
            Workbooks = new FormatWorkbooks(workbook);
            workbook.Application = this;
        }

        public FormatWorkbook ActiveWorkbook { get; }
        public FormatWorkbooks Workbooks { get; }
        public long Hwnd => 4242;
        public bool ScreenUpdating { get; set; } = true;
        public bool DisplayAlerts { get; set; } = true;
    }

    public sealed class FormatWorkbooks : IEnumerable<FormatWorkbook>
    {
        private readonly FormatWorkbook _workbook;
        public FormatWorkbooks(FormatWorkbook workbook) => _workbook = workbook;
        public int Count => 1;
        public FormatWorkbook Item(int index) => index == 1 ? _workbook : throw new ArgumentOutOfRangeException(nameof(index));
        public IEnumerator<FormatWorkbook> GetEnumerator() { yield return _workbook; }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    public sealed class FormatWorkbook
    {
        public FormatWorkbook(string fullName, string sheetName)
        {
            FullName = fullName;
            Name = Path.GetFileName(fullName);
            Sheet = new FormatWorksheet(this, sheetName);
            Worksheets = new FormatWorksheets(Sheet);
        }

        public FormatExcelApplication? Application { get; set; }
        public string FullName { get; }
        public string Name { get; }
        public FormatWorksheet Sheet { get; }
        public FormatWorksheets Worksheets { get; }
        public FormatWorksheet ActiveSheet => Sheet;
        public int UsedRangeAccessCount { get; internal set; }
        public int MultiCellMergeAreaReads { get; internal set; }
    }

    public sealed class FormatWorksheets : IEnumerable<FormatWorksheet>
    {
        private readonly FormatWorksheet _sheet;
        public FormatWorksheets(FormatWorksheet sheet) => _sheet = sheet;
        public int Count => 1;
        public FormatWorksheet Item(object key) => key switch
        {
            int index when index == 1 => _sheet,
            string name when string.Equals(name, _sheet.Name, StringComparison.OrdinalIgnoreCase) => _sheet,
            _ => throw new ArgumentOutOfRangeException(nameof(key)),
        };
        public IEnumerator<FormatWorksheet> GetEnumerator() { yield return _sheet; }
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    public sealed class FormatWorksheet
    {
        private readonly Dictionary<(int Row, int Column), FormatCell> _cells = new();

        public FormatWorksheet(FormatWorkbook workbook, string name)
        {
            Workbook = workbook;
            Name = name;
        }

        public FormatWorkbook Workbook { get; }
        public string Name { get; }
        public bool ProtectContents => false;
        public Exception? RangeFault { get; set; }
        public string? RangeFaultAddress { get; set; }
        public bool PartialRangeStyleWrites { get; set; }

        public FormatRange UsedRange
        {
            get
            {
                Workbook.UsedRangeAccessCount++;
                return new FormatRange(this, new[] { ((1, 1), (1, 1)) });
            }
        }

        public FormatRange Range(string address)
        {
            if (RangeFault is not null &&
                (RangeFaultAddress is null ||
                 string.Equals(address, RangeFaultAddress, StringComparison.OrdinalIgnoreCase)))
                throw RangeFault;
            return new FormatRange(this, ParseAreas(address));
        }

        public FormatCell Cell(int row, int column)
        {
            if (_cells.TryGetValue((row, column), out var cell)) return cell;
            cell = new FormatCell(this, row, column);
            _cells[(row, column)] = cell;
            return cell;
        }

        public void Merge(int row, int column, int rows, int columns)
        {
            var area = new FormatRange(this, new[] { ((row, column), (row + rows - 1, column + columns - 1)) });
            for (var r = row; r < row + rows; r++)
            for (var c = column; c < column + columns; c++)
            {
                var cell = Cell(r, c);
                cell.MergeCells = true;
                cell.MergeArea = area;
            }
        }

        private static (int Row, int Column) ParseCell(string token)
        {
            var index = 0;
            var column = 0;
            while (index < token.Length && char.IsLetter(token[index]))
            {
                column = column * 26 + char.ToUpperInvariant(token[index]) - 'A' + 1;
                index++;
            }
            return (int.Parse(token[index..], System.Globalization.CultureInfo.InvariantCulture), column);
        }

        private static List<((int Row, int Column) Start, (int Row, int Column) End)> ParseAreas(string address)
        {
            var areas = new List<((int, int), (int, int))>();
            foreach (var part in address.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries))
            {
                var bounds = part.Split(':');
                var start = ParseCell(bounds[0]);
                var end = bounds.Length == 1 ? start : ParseCell(bounds[1]);
                areas.Add((start, end));
            }
            return areas;
        }
    }

    public sealed class FormatRange
    {
        private readonly FormatWorksheet _sheet;
        private readonly IReadOnlyList<((int Row, int Column) Start, (int Row, int Column) End)> _areas;

        public FormatRange(
            FormatWorksheet sheet,
            IReadOnlyList<((int Row, int Column) Start, (int Row, int Column) End)> areas,
            bool leaf = false)
        {
            _sheet = sheet;
            _areas = areas;
            var first = areas[0];
            Row = first.Start.Row;
            Column = first.Start.Column;
            Rows = new CountBox(first.End.Row - first.Start.Row + 1);
            Columns = new CountBox(first.End.Column - first.Start.Column + 1);
            Cells = new FormatCells(sheet, first.Start.Row, first.Start.Column);
            Areas = leaf
                ? new FormatAreas(new[] { this })
                : new FormatAreas(areas.Select(area => new FormatRange(sheet, new[] { area }, leaf: true)).ToArray());
        }

        public int Row { get; }
        public int Column { get; }
        public CountBox Rows { get; }
        public CountBox Columns { get; }
        public FormatCells Cells { get; }
        public FormatAreas Areas { get; }
        public FormatRangeFont Font => new(this);
        public FormatRangeInterior Interior => new(this);
        public object? NumberFormat
        {
            get => Aggregate(cell => cell.NumberFormat);
            set
            {
                WriteStyle(cell =>
                    cell.NumberFormat = Convert.ToString(value, System.Globalization.CultureInfo.InvariantCulture) ?? "General");
            }
        }
        public object? MergeCells
        {
            get
            {
                bool? common = null;
                foreach (var cell in EnumerateCells())
                {
                    var merged = Convert.ToBoolean(cell.MergeCells, System.Globalization.CultureInfo.InvariantCulture);
                    if (common is null) common = merged;
                    else if (common.Value != merged) return null;
                }

                return common ?? false;
            }
        }
        public FormatRange MergeArea
        {
            get
            {
                if (Count > 1)
                {
                    _sheet.Workbook.MultiCellMergeAreaReads++;
                    throw new InvalidOperationException(
                        "Range.MergeArea is only valid for a single-cell range");
                }

                return EnumerateCells().First().MergeArea;
            }
        }
        public int Count => Rows.Count * Columns.Count;
        public long CountLarge => Count;
        public object? Value2 => null;
        public object? Formula => null;

        internal FormatWorksheet Sheet => _sheet;
        private FormatCell FirstCell => _sheet.Cell(Row, Column);

        internal object? AggregateColor(Func<FormatCell, double> read)
        {
            var value = Aggregate(cell => read(cell));
            if (value is null)
                return 0d;
            return value;
        }

        internal void WriteStyle(Action<FormatCell> write)
        {
            if (_sheet.PartialRangeStyleWrites)
            {
                write(FirstCell);
                return;
            }

            foreach (var cell in EnumerateCells())
                write(cell);
        }

        internal IEnumerable<FormatCell> EnumerateCells()
        {
            foreach (var area in _areas)
            {
                for (var row = area.Start.Row; row <= area.End.Row; row++)
                {
                    for (var column = area.Start.Column; column <= area.End.Column; column++)
                        yield return _sheet.Cell(row, column);
                }
            }
        }

        internal object? Aggregate(Func<FormatCell, object?> read)
        {
            object? first = null;
            var started = false;
            foreach (var cell in EnumerateCells())
            {
                var value = read(cell);
                if (!started)
                {
                    first = value;
                    started = true;
                    continue;
                }

                if (!Equals(first, value))
                    return null;
            }

            return first;
        }

        public string Address(bool rowAbs, bool colAbs)
        {
            _ = (rowAbs, colAbs);
            var first = _areas[0];
            return first.Start == first.End
                ? CellName(first.Start.Column, first.Start.Row)
                : $"{CellName(first.Start.Column, first.Start.Row)}:{CellName(first.End.Column, first.End.Row)}";
        }

        private static string CellName(int column, int row)
        {
            var name = "";
            var current = column;
            while (current > 0)
            {
                var m = (current - 1) % 26;
                name = (char)('A' + m) + name;
                current = (current - 1) / 26;
            }
            return name + row.ToString(System.Globalization.CultureInfo.InvariantCulture);
        }
    }

    public sealed class FormatAreas
    {
        private readonly FormatRange[] _areas;
        public FormatAreas(FormatRange[] areas) => _areas = areas;
        public int Count => _areas.Length;
        public FormatRange Item(int index) => _areas[index - 1];
    }

    public sealed class FormatCells
    {
        private readonly FormatWorksheet _sheet;
        private readonly int _startRow;
        private readonly int _startColumn;
        public FormatCells(FormatWorksheet sheet, int startRow, int startColumn)
        {
            _sheet = sheet;
            _startRow = startRow;
            _startColumn = startColumn;
        }
        public FormatCell Item(int row, int column) => _sheet.Cell(_startRow + row - 1, _startColumn + column - 1);
    }

    public sealed class CountBox
    {
        public CountBox(int count) => Count = count;
        public int Count { get; }
    }

    public sealed class FormatCell
    {
        public FormatCell(FormatWorksheet sheet, int row, int column)
        {
            Sheet = sheet;
            Row = row;
            Column = column;
            Font = new FormatFont();
            Interior = new FormatInterior();
            NumberFormat = "General";
            MergeCells = false;
            MergeArea = new FormatRange(sheet, new[] { ((row, column), (row, column)) });
        }

        public FormatWorksheet Sheet { get; }
        public int Row { get; }
        public int Column { get; }
        public FormatFont Font { get; }
        public FormatInterior Interior { get; }
        public string NumberFormat { get; set; }
        public object MergeCells { get; set; }
        public FormatRange MergeArea { get; set; }
    }

    public sealed class FormatRangeFont
    {
        private readonly FormatRange _range;
        public FormatRangeFont(FormatRange range) => _range = range;

        public object? Bold
        {
            get => _range.Aggregate(cell => cell.Font.Bold);
            set => _range.WriteStyle(cell => cell.Font.Bold = value);
        }

        public object? Italic
        {
            get => _range.Aggregate(cell => cell.Font.Italic);
            set => _range.WriteStyle(cell => cell.Font.Italic = value);
        }

        public object? Size
        {
            get => _range.Aggregate(cell => cell.Font.Size);
            set => _range.WriteStyle(cell => cell.Font.Size = value);
        }

        public object? Color
        {
            get => _range.AggregateColor(cell => cell.Font.Color);
            set => _range.WriteStyle(cell =>
                cell.Font.Color = Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture));
        }

        public object? ColorIndex
        {
            get => _range.Aggregate(cell => cell.Font.ColorIndex);
            set => _range.WriteStyle(cell =>
                cell.Font.ColorIndex = Convert.ToInt32(value, System.Globalization.CultureInfo.InvariantCulture));
        }

        public object? TintAndShade
        {
            get => _range.Aggregate(cell => cell.Font.TintAndShade);
            set => _range.WriteStyle(cell =>
                cell.Font.TintAndShade = Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture));
        }

        public int ThemeColor
        {
            get
            {
                int? first = null;
                foreach (var cell in _range.EnumerateCells())
                {
                    var value = cell.Font.ThemeColor;
                    if (first is null) first = value;
                    else if (first != value)
                        throw new COMException("Font.ThemeColor is mixed", unchecked((int)0x800A03EC));
                }

                return first ?? throw new COMException(
                    "Unable to get the ThemeColor property of the Font class",
                    unchecked((int)0x800A03EC));
            }
            set => _range.WriteStyle(cell => cell.Font.ThemeColor = value);
        }
    }

    public sealed class FormatRangeInterior
    {
        private readonly FormatRange _range;
        public FormatRangeInterior(FormatRange range) => _range = range;

        public object? Color
        {
            get => _range.AggregateColor(cell => cell.Interior.Color);
            set => _range.WriteStyle(cell =>
                cell.Interior.Color = Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture));
        }

        public object? ColorIndex
        {
            get => _range.Aggregate(cell => cell.Interior.ColorIndex);
            set => _range.WriteStyle(cell =>
                cell.Interior.ColorIndex = Convert.ToInt32(value, System.Globalization.CultureInfo.InvariantCulture));
        }

        public object? TintAndShade
        {
            get => _range.Aggregate(cell => cell.Interior.TintAndShade);
            set => _range.WriteStyle(cell =>
                cell.Interior.TintAndShade = Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture));
        }

        public object? Pattern
        {
            get => _range.Aggregate(cell => cell.Interior.Pattern);
            set => _range.WriteStyle(cell =>
                cell.Interior.Pattern = Convert.ToInt32(value, System.Globalization.CultureInfo.InvariantCulture));
        }

        public object? PatternColor
        {
            get => _range.AggregateColor(cell => cell.Interior.PatternColor);
            set => _range.WriteStyle(cell =>
                cell.Interior.PatternColor = Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture));
        }

        public object? PatternColorIndex
        {
            get => _range.Aggregate(cell => cell.Interior.PatternColorIndex);
            set { foreach (var cell in _range.EnumerateCells()) cell.Interior.PatternColorIndex = Convert.ToInt32(value, System.Globalization.CultureInfo.InvariantCulture); }
        }

        public object? PatternTintAndShade
        {
            get => _range.Aggregate(cell => cell.Interior.PatternTintAndShade);
            set { foreach (var cell in _range.EnumerateCells()) cell.Interior.PatternTintAndShade = Convert.ToDouble(value, System.Globalization.CultureInfo.InvariantCulture); }
        }

        public int ThemeColor
        {
            get
            {
                int? first = null;
                foreach (var cell in _range.EnumerateCells())
                {
                    var value = cell.Interior.ThemeColor;
                    if (first is null) first = value;
                    else if (first != value)
                        throw new COMException("Interior.ThemeColor is mixed", unchecked((int)0x800A03EC));
                }

                return first ?? throw new COMException(
                    "Unable to get the ThemeColor property of the Interior class",
                    unchecked((int)0x800A03EC));
            }
            set { foreach (var cell in _range.EnumerateCells()) cell.Interior.ThemeColor = value; }
        }

        public int PatternThemeColor
        {
            get
            {
                int? first = null;
                foreach (var cell in _range.EnumerateCells())
                {
                    var value = cell.Interior.PatternThemeColor;
                    if (first is null) first = value;
                    else if (first != value)
                        throw new COMException("Interior.PatternThemeColor is mixed", unchecked((int)0x800A03EC));
                }

                return first ?? throw new COMException(
                    "Unable to get the PatternThemeColor property of the Interior class",
                    unchecked((int)0x800A03EC));
            }
            set { foreach (var cell in _range.EnumerateCells()) cell.Interior.PatternThemeColor = value; }
        }
    }

    public sealed class FormatFont
    {
        private int? _themeColor;
        public object? Bold { get; set; } = false;
        public object? Italic { get; set; } = false;
        public object? Size { get; set; } = 11d;
        private double _color;
        public int ColorReads { get; private set; }
        public int ColorIndexReads { get; private set; }
        public int TintReads { get; private set; }
        public int ThemeReads { get; private set; }
        public double Color
        {
            get { ColorReads++; return _color; }
            set { _color = value; ColorAssignments++; }
        }
        private int _colorIndex = 1;
        public int ColorIndex
        {
            get { ColorIndexReads++; return _colorIndex; }
            set => _colorIndex = value;
        }
        private double _tintAndShade;
        public double TintAndShade
        {
            get { TintReads++; return _tintAndShade; }
            set => _tintAndShade = value;
        }
        public int ColorAssignments { get; set; }
        public Exception? ThemeColorFault { get; set; }
        public int ThemeColor
        {
            get
            {
                ThemeReads++;
                if (ThemeColorFault is not null) throw ThemeColorFault;
                return _themeColor ?? throw new COMException(
                    "Unable to get the ThemeColor property of the Font class",
                    unchecked((int)0x800A03EC));
            }
            set => _themeColor = value;
        }
    }

    public sealed class FormatInterior
    {
        private int? _themeColor;
        private int? _patternThemeColor;
        private double _color = 16_777_215d;
        private int _colorIndex = ExcelStyleContract.XlColorIndexNone;
        private double _tintAndShade;
        public int ColorReads { get; private set; }
        public int ColorIndexReads { get; private set; }
        public int TintReads { get; private set; }
        public int ThemeReads { get; private set; }
        public int PatternReads { get; private set; }
        public double Color
        {
            get { ColorReads++; return _color; }
            set
            {
                _color = value;
                if (_colorIndex is ExcelStyleContract.XlColorIndexAutomatic or ExcelStyleContract.XlColorIndexNone)
                    _colorIndex = 1;
            }
        }
        public int ColorIndex
        {
            get { ColorIndexReads++; return _colorIndex; }
            set => _colorIndex = value;
        }
        public double TintAndShade
        {
            get { TintReads++; return _tintAndShade; }
            set => _tintAndShade = value;
        }
        private int _pattern = ExcelStyleContract.XlPatternNone;
        public int Pattern
        {
            get { PatternReads++; return _pattern; }
            set => _pattern = value;
        }
        public double PatternColor { get; set; }
        public int PatternColorIndex { get; set; } = ExcelStyleContract.XlColorIndexNone;
        public double PatternTintAndShade { get; set; }
        public Exception? ThemeColorFault { get; set; }
        public int ThemeColor
        {
            get
            {
                ThemeReads++;
                if (ThemeColorFault is not null) throw ThemeColorFault;
                return _themeColor ?? throw new COMException(
                    "Unable to get the ThemeColor property of the Interior class",
                    unchecked((int)0x800A03EC));
            }
            set => _themeColor = value;
        }
        public int PatternThemeColor
        {
            get => _patternThemeColor ?? throw new COMException(
                "Unable to get the PatternThemeColor property of the Interior class",
                unchecked((int)0x800A03EC));
            set => _patternThemeColor = value;
        }
    }
}
