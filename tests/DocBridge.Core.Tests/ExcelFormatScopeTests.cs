using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelFormatScopeTests
{
    [Fact]
    public void Bold_only_capture_skips_unrelated_tint_and_theme()
    {
        using var home = new TestHome();
        var workbook = new ExcelFormatOnlySafetyTests.FormatWorkbook(@"C:\docbridge-fixtures\scope-bold-only.xlsx", "Sheet1");
        var cell = workbook.Sheet.Cell(1, 1);
        cell.Font.Bold = false;
        cell.Font.ThemeColorFault = new COMException("theme should not be read", unchecked((int)0x80010108));
        cell.Interior.ThemeColorFault = new COMException("fill theme should not be read", unchecked((int)0x80010108));
        cell.Interior.Color = 42662d;
        cell.Interior.TintAndShade = -0.35d;
        using var adapter = CreateAdapter(workbook);

        adapter.CaptureSnapshot(CreateSnapshotDir(home), new JsonObject(), ExcelFormatOnlySafetyTests.FormatOps("A1"));

        var state = ReadState(home);
        Assert.Equal(ExcelAdapter.FormatOnlyScopedSnapshotVersion, Json.GetInt(state, "snapshotVersion"));
        var formatState = (JsonObject)Json.GetArr(state, "formatStates")![0]!;
        Assert.Equal(ExcelAdapter.FormatStyleModeUniform, Json.GetString(formatState, "styleMode"));
        var style = ExcelFormatOnlySafetyTests.FirstCapturedStyle(formatState);
        Assert.False(Json.GetBool(style, "bold"));
        Assert.False(style.ContainsKey("italic"));
        Assert.False(style.ContainsKey("fontColor"));
        Assert.False(style.ContainsKey("fillColor"));
        Assert.False(style.ContainsKey("fontTintAndShade"));
        Assert.False(style.ContainsKey("fillTintAndShade"));
        Assert.Equal(0, cell.Font.ColorReads);
        Assert.Equal(0, cell.Font.TintReads);
        Assert.Equal(0, cell.Font.ThemeReads);
        Assert.Equal(0, cell.Interior.ColorReads);
        Assert.Equal(0, cell.Interior.TintReads);
        Assert.Equal(0, cell.Interior.ThemeReads);
    }

    [Fact]
    public void Bold_only_rgb_tint_is_not_rejected()
    {
        using var home = new TestHome();
        var workbook = new ExcelFormatOnlySafetyTests.FormatWorkbook(@"C:\docbridge-fixtures\scope-bold-rgb-tint.xlsx", "Sheet1");
        var cell = workbook.Sheet.Cell(3, 3);
        cell.Interior.Color = 42662d;
        cell.Interior.TintAndShade = -0.3499862666707358d;
        using var adapter = CreateAdapter(workbook);

        adapter.CaptureSnapshot(CreateSnapshotDir(home), new JsonObject(), ExcelFormatOnlySafetyTests.FormatOps("C3"));

        var style = ExcelFormatOnlySafetyTests.FirstCapturedStyle((JsonObject)Json.GetArr(ReadState(home), "formatStates")![0]!);
        Assert.True(style.ContainsKey("bold"));
        Assert.False(style.ContainsKey("fillTintAndShade"));
        Assert.Equal(42662d, cell.Interior.Color);
        Assert.Equal(-0.3499862666707358d, cell.Interior.TintAndShade);
    }

    [Fact]
    public void Bold_only_restore_does_not_rewrite_colors()
    {
        using var home = new TestHome();
        var workbook = new ExcelFormatOnlySafetyTests.FormatWorkbook(@"C:\docbridge-fixtures\scope-bold-restore.xlsx", "Sheet1");
        var cell = workbook.Sheet.Cell(1, 1);
        cell.Font.Bold = true;
        cell.Font.Color = 255d;
        cell.Interior.Color = 128d;
        using var adapter = CreateAdapter(workbook);
        var snapshotDir = CreateSnapshotDir(home);
        adapter.CaptureSnapshot(snapshotDir, new JsonObject(), ExcelFormatOnlySafetyTests.FormatOps("A1"));

        cell.Font.Bold = false;
        cell.Font.Color = 1d;
        cell.Interior.Color = 2d;
        cell.Font.ColorAssignments = 0;

        var restored = adapter.RestoreSnapshot(snapshotDir, new JsonObject());
        Assert.True(Json.GetBool(restored, "ok"), restored.ToJsonString());
        Assert.True((bool)cell.Font.Bold!);
        Assert.Equal(1d, cell.Font.Color);
        Assert.Equal(2d, cell.Interior.Color);
        Assert.Equal(0, cell.Font.ColorAssignments);
    }

    [Fact]
    public void Bold_only_fingerprint_ignores_unrelated_fill_changes()
    {
        using var home = new TestHome();
        var workbook = new ExcelFormatOnlySafetyTests.FormatWorkbook(@"C:\docbridge-fixtures\scope-fingerprint-fill.xlsx", "Sheet1");
        var cell = workbook.Sheet.Cell(1, 1);
        cell.Font.Bold = true;
        cell.Interior.Color = 255d;
        using var adapter = CreateAdapter(workbook);
        var snapshotDir = CreateSnapshotDir(home);
        var metadata = new JsonObject();
        var ops = ExcelFormatOnlySafetyTests.FormatOps("A1");
        adapter.CaptureSnapshot(snapshotDir, metadata, ops);

        cell.Interior.Color = 128d;
        var fillChanged = adapter.ValidatePreviewReuse(snapshotDir, metadata, ops);
        Assert.True(Json.GetBool(fillChanged, "ok"));
        Assert.True(Json.GetBool(fillChanged, "reusable"));

        cell.Font.Bold = false;
        var boldChanged = adapter.ValidatePreviewReuse(snapshotDir, metadata, ops);
        Assert.True(Json.GetBool(boldChanged, "ok"));
        Assert.False(Json.GetBool(boldChanged, "reusable"));
    }

    [Fact]
    public void Legacy_v2_full_format_state_still_restores()
    {
        using var home = new TestHome();
        var workbook = new ExcelFormatOnlySafetyTests.FormatWorkbook(@"C:\docbridge-fixtures\scope-legacy-v2.xlsx", "Sheet1");
        var cell = workbook.Sheet.Cell(1, 1);
        cell.Font.Bold = true;
        cell.Font.Color = 255d;
        cell.Interior.Color = 128d;
        using var adapter = CreateAdapter(workbook);
        var snapshotDir = CreateSnapshotDir(home);
        WriteState(snapshotDir, new JsonObject
        {
            ["snapshotVersion"] = 2,
            ["restoreMode"] = "format-only",
            ["documentRef"] = workbook.FullName,
            ["formatStates"] = new JsonArray(ValidFullFormatState()),
            ["coverage"] = new JsonObject { ["complete"] = true, ["cellCount"] = 1 },
        });

        var restored = adapter.RestoreSnapshot(snapshotDir, new JsonObject());
        Assert.True(Json.GetBool(restored, "ok"), restored.ToJsonString());
        Assert.False((bool)cell.Font.Bold!);
        Assert.Equal(0d, cell.Font.Color);
        Assert.Equal(ExcelStyleContract.XlColorIndexNone, cell.Interior.ColorIndex);
        Assert.Equal(ExcelStyleContract.XlPatternNone, cell.Interior.Pattern);
    }

    [Fact]
    public void Uniform_bool_size_and_number_format_use_range_fast_path()
    {
        using var home = new TestHome();
        var workbook = new ExcelFormatOnlySafetyTests.FormatWorkbook(@"C:\docbridge-fixtures\scope-uniform-fast.xlsx", "Sheet1");
        foreach (var (row, column) in new[] { (1, 1), (1, 2), (2, 1), (2, 2) })
        {
            var cell = workbook.Sheet.Cell(row, column);
            cell.Font.Bold = false;
            cell.Font.Size = 14d;
            cell.NumberFormat = "0.00";
        }

        using var adapter = CreateAdapter(workbook);
        adapter.CaptureSnapshot(CreateSnapshotDir(home), new JsonObject(), ExcelFormatOnlySafetyTests.FormatOps("A1:B2", new JsonObject
        {
            ["bold"] = true,
            ["fontSize"] = 12,
            ["numberFormat"] = "General",
        }));

        var formatState = (JsonObject)Json.GetArr(ReadState(home), "formatStates")![0]!;
        Assert.Equal(ExcelAdapter.FormatStyleModeUniform, Json.GetString(formatState, "styleMode"));
        var style = ExcelFormatOnlySafetyTests.FirstCapturedStyle(formatState);
        Assert.False(Json.GetBool(style, "bold"));
        Assert.Equal(14d, style["fontSize"]!.GetValue<double>());
        Assert.Equal("0.00", Json.GetString(style, "numberFormat"));
        Assert.False(style.ContainsKey("fillColor"));
    }

    [Fact]
    public void Mixed_or_null_bool_and_number_format_fall_back_to_groups()
    {
        using var home = new TestHome();
        var workbook = new ExcelFormatOnlySafetyTests.FormatWorkbook(@"C:\docbridge-fixtures\scope-mixed-fast.xlsx", "Sheet1");
        workbook.Sheet.Cell(1, 1).Font.Bold = true;
        workbook.Sheet.Cell(1, 2).Font.Bold = false;
        workbook.Sheet.Cell(1, 1).NumberFormat = "General";
        workbook.Sheet.Cell(1, 2).NumberFormat = "0.00";
        using var adapter = CreateAdapter(workbook);

        adapter.CaptureSnapshot(CreateSnapshotDir(home), new JsonObject(), ExcelFormatOnlySafetyTests.FormatOps("A1:B1", new JsonObject
        {
            ["bold"] = true,
            ["numberFormat"] = "General",
        }));

        var formatState = (JsonObject)Json.GetArr(ReadState(home), "formatStates")![0]!;
        Assert.Equal(ExcelAdapter.FormatStyleModeGroups, Json.GetString(formatState, "styleMode"));
        var groups = Json.GetArr(formatState, "groups")!;
        Assert.Equal(2, groups.Count);
        Assert.True(Json.GetBool((JsonObject)((JsonObject)groups[0]!)["style"]!, "bold"));
        Assert.False(Json.GetBool((JsonObject)((JsonObject)groups[1]!)["style"]!, "bold"));
    }

    [Fact]
    public void Uniform_restore_rewrites_the_whole_rectangle()
    {
        using var home = new TestHome();
        var workbook = new ExcelFormatOnlySafetyTests.FormatWorkbook(@"C:\docbridge-fixtures\scope-uniform-restore.xlsx", "Sheet1");
        foreach (var (row, column) in new[] { (1, 1), (1, 2), (2, 1), (2, 2) })
            workbook.Sheet.Cell(row, column).Font.Bold = false;
        using var adapter = CreateAdapter(workbook);
        var snapshotDir = CreateSnapshotDir(home);
        adapter.CaptureSnapshot(snapshotDir, new JsonObject(), ExcelFormatOnlySafetyTests.FormatOps("A1:B2"));

        foreach (var (row, column) in new[] { (1, 1), (1, 2), (2, 1), (2, 2) })
            workbook.Sheet.Cell(row, column).Font.Bold = true;

        var restored = adapter.RestoreSnapshot(snapshotDir, new JsonObject());
        Assert.True(Json.GetBool(restored, "ok"), restored.ToJsonString());
        foreach (var (row, column) in new[] { (1, 1), (1, 2), (2, 1), (2, 2) })
            Assert.False((bool)workbook.Sheet.Cell(row, column).Font.Bold!);
    }

    [Fact]
    public void Color_zero_and_style_name_are_not_treated_as_uniform()
    {
        using var home = new TestHome();
        var workbook = new ExcelFormatOnlySafetyTests.FormatWorkbook(@"C:\docbridge-fixtures\scope-color-zero.xlsx", "Sheet1");
        var automatic = workbook.Sheet.Cell(1, 1);
        automatic.Interior.Color = 0d;
        automatic.Interior.ColorIndex = ExcelStyleContract.XlColorIndexAutomatic;
        automatic.Interior.Pattern = ExcelStyleContract.XlPatternNone;
        var black = workbook.Sheet.Cell(1, 2);
        black.Interior.Color = 0d;
        black.Interior.ColorIndex = 1;
        black.Interior.Pattern = ExcelStyleContract.XlPatternSolid;
        using var adapter = CreateAdapter(workbook);

        adapter.CaptureSnapshot(CreateSnapshotDir(home), new JsonObject(), ExcelFormatOnlySafetyTests.FormatOps("A1:B1", new JsonObject
        {
            ["fillColor"] = 255,
        }));

        var formatState = (JsonObject)Json.GetArr(ReadState(home), "formatStates")![0]!;
        Assert.NotEqual(ExcelAdapter.FormatStyleModeUniform, Json.GetString(formatState, "styleMode"));
        var groups = Json.GetArr(formatState, "groups")!;
        Assert.Equal(2, groups.Count);
        var a1 = ExcelFormatOnlySafetyTests.StyleAt(formatState, 1, 1);
        var b1 = ExcelFormatOnlySafetyTests.StyleAt(formatState, 1, 2);
        Assert.Equal(0d, a1["fillColor"]!.GetValue<double>());
        Assert.Equal(0d, b1["fillColor"]!.GetValue<double>());
        Assert.Equal(ExcelStyleContract.XlColorIndexAutomatic, a1["fillColorIndex"]!.GetValue<int>());
        Assert.Equal(1, b1["fillColorIndex"]!.GetValue<int>());
        Assert.NotEqual(Json.GetInt(a1, "fillPattern"), Json.GetInt(b1, "fillPattern"));
    }

    [Fact]
    public void Color_ops_capture_linked_index_theme_tint_and_pattern()
    {
        using var home = new TestHome();
        var workbook = new ExcelFormatOnlySafetyTests.FormatWorkbook(@"C:\docbridge-fixtures\scope-color-linked.xlsx", "Sheet1");
        var cell = workbook.Sheet.Cell(2, 2);
        cell.Font.ThemeColor = 5;
        cell.Font.TintAndShade = 0.25d;
        cell.Font.Color = 999d;
        cell.Interior.Color = 128d;
        cell.Interior.ColorIndex = 3;
        cell.Interior.Pattern = 2;
        cell.Interior.PatternColor = 65535d;
        cell.Interior.PatternColorIndex = 6;
        using var adapter = CreateAdapter(workbook);

        adapter.CaptureSnapshot(CreateSnapshotDir(home), new JsonObject(), ExcelFormatOnlySafetyTests.FormatOps("B2", new JsonObject
        {
            ["fontColor"] = 0,
            ["fillColor"] = 255,
        }));

        var style = ExcelFormatOnlySafetyTests.FirstCapturedStyle((JsonObject)Json.GetArr(ReadState(home), "formatStates")![0]!);
        Assert.Equal(5, style["fontThemeColor"]!.GetValue<int>());
        Assert.Equal(0.25d, style["fontTintAndShade"]!.GetValue<double>());
        Assert.Equal(128d, style["fillColor"]!.GetValue<double>());
        Assert.Equal(3, style["fillColorIndex"]!.GetValue<int>());
        Assert.Equal(2, Json.GetInt(style, "fillPattern"));
        Assert.Equal(65535d, style["fillPatternColor"]!.GetValue<double>());
        Assert.Equal(6, Json.GetInt(style, "fillPatternColorIndex"));
        Assert.False(style.ContainsKey("bold"));
    }

    [Fact]
    public void Mixed_color_groups_restore_only_their_rectangles()
    {
        using var home = new TestHome();
        var workbook = new ExcelFormatOnlySafetyTests.FormatWorkbook(@"C:\docbridge-fixtures\scope-color-groups.xlsx", "Sheet1");
        var left = workbook.Sheet.Cell(1, 1);
        left.Interior.Color = 255d;
        left.Interior.Pattern = ExcelStyleContract.XlPatternSolid;
        var right = workbook.Sheet.Cell(1, 2);
        right.Interior.Color = 128d;
        right.Interior.Pattern = ExcelStyleContract.XlPatternSolid;
        using var adapter = CreateAdapter(workbook);
        var snapshotDir = CreateSnapshotDir(home);
        adapter.CaptureSnapshot(snapshotDir, new JsonObject(), ExcelFormatOnlySafetyTests.FormatOps("A1:B1", new JsonObject
        {
            ["fillColor"] = 1,
        }));

        left.Interior.Color = 0d;
        right.Interior.Color = 0d;
        var restored = adapter.RestoreSnapshot(snapshotDir, new JsonObject());
        Assert.True(Json.GetBool(restored, "ok"), restored.ToJsonString());
        Assert.Equal(255d, left.Interior.Color);
        Assert.Equal(128d, right.Interior.Color);
    }

    [Fact]
    public void Apply_readback_rejects_mixed_null_bold_after_partial_write()
    {
        var workbook = new ExcelFormatOnlySafetyTests.FormatWorkbook(@"C:\docbridge-fixtures\apply-mixed-null-bold.xlsx", "Sheet1");
        workbook.Sheet.Cell(1, 1).Font.Bold = true;
        workbook.Sheet.Cell(1, 2).Font.Bold = true;
        workbook.Sheet.PartialRangeStyleWrites = true;
        using var adapter = CreateAdapter(workbook);

        var exec = adapter.Apply(ExcelFormatOnlySafetyTests.FormatOps("A1:B1", new JsonObject { ["bold"] = false }), "snap-apply-mixed-bold");

        Assert.False(exec.Ok);
        Assert.False(Json.GetBool(exec.Readback, "verified"));
        Assert.Contains("bold", string.Join(" ", Json.GetArr(exec.Readback, "mismatches")!.Select(n => n!.GetValue<string>())), StringComparison.OrdinalIgnoreCase);
        Assert.False((bool)workbook.Sheet.Cell(1, 1).Font.Bold!);
        Assert.True((bool)workbook.Sheet.Cell(1, 2).Font.Bold!);
    }

    [Fact]
    public void Apply_readback_rejects_color_zero_after_partial_write()
    {
        var workbook = new ExcelFormatOnlySafetyTests.FormatWorkbook(@"C:\docbridge-fixtures\apply-mixed-color-zero.xlsx", "Sheet1");
        var left = workbook.Sheet.Cell(1, 1);
        var right = workbook.Sheet.Cell(1, 2);
        left.Interior.Color = 255d;
        left.Interior.Pattern = ExcelStyleContract.XlPatternSolid;
        right.Interior.Color = 255d;
        right.Interior.Pattern = ExcelStyleContract.XlPatternSolid;
        workbook.Sheet.PartialRangeStyleWrites = true;
        using var adapter = CreateAdapter(workbook);

        var exec = adapter.Apply(ExcelFormatOnlySafetyTests.FormatOps("A1:B1", new JsonObject { ["fillColor"] = 0 }), "snap-apply-mixed-black");

        Assert.False(exec.Ok);
        Assert.False(Json.GetBool(exec.Readback, "verified"));
        Assert.Contains("fillColor", string.Join(" ", Json.GetArr(exec.Readback, "mismatches")!.Select(n => n!.GetValue<string>())), StringComparison.OrdinalIgnoreCase);
        Assert.Equal(0d, left.Interior.Color);
        Assert.Equal(255d, right.Interior.Color);
        Assert.Equal(ExcelStyleContract.XlPatternSolid, right.Interior.Pattern);
    }

    [Fact]
    public void Apply_readback_accepts_uniform_false_and_black_after_complete_write()
    {
        var workbook = new ExcelFormatOnlySafetyTests.FormatWorkbook(@"C:\docbridge-fixtures\apply-uniform-false-black.xlsx", "Sheet1");
        foreach (var (row, column) in new[] { (1, 1), (1, 2) })
        {
            var cell = workbook.Sheet.Cell(row, column);
            cell.Font.Bold = true;
            cell.Interior.Color = 255d;
            cell.Interior.Pattern = ExcelStyleContract.XlPatternSolid;
        }

        using var adapter = CreateAdapter(workbook);
        var exec = adapter.Apply(ExcelFormatOnlySafetyTests.FormatOps("A1:B1", new JsonObject
        {
            ["bold"] = false,
            ["fillColor"] = 0,
        }), "snap-apply-uniform-false-black");

        Assert.True(exec.Ok, string.Join(" ", exec.Errors));
        var readback = exec.Readback ?? new JsonObject();
        Assert.True(Json.GetBool(readback, "verified"), readback.ToJsonString());
        Assert.False((bool)workbook.Sheet.Cell(1, 1).Font.Bold!);
        Assert.False((bool)workbook.Sheet.Cell(1, 2).Font.Bold!);
        Assert.Equal(0d, workbook.Sheet.Cell(1, 1).Interior.Color);
        Assert.Equal(0d, workbook.Sheet.Cell(1, 2).Interior.Color);
    }

    [Fact]
    public void Fully_merged_target_is_captured_without_reading_area_mergearea()
    {
        using var home = new TestHome();
        var workbook = new ExcelFormatOnlySafetyTests.FormatWorkbook(@"C:\docbridge-fixtures\scope-full-merge.xlsx", "Sheet1");
        workbook.Sheet.Merge(1, 1, 2, 2);
        using var adapter = CreateAdapter(workbook);

        adapter.CaptureSnapshot(CreateSnapshotDir(home), new JsonObject(), ExcelFormatOnlySafetyTests.FormatOps("A1:B2"));

        var state = ReadState(home);
        Assert.Equal(0, workbook.MultiCellMergeAreaReads);
        Assert.Equal(ExcelAdapter.FormatOnlyScopedSnapshotVersion, Json.GetInt(state, "snapshotVersion"));
        Assert.Equal(4, Json.GetInt(Json.GetObj(state, "coverage"), "cellCount"));
        Assert.Equal(ExcelAdapter.FormatStyleModeUniform, Json.GetString((JsonObject)Json.GetArr(state, "formatStates")![0]!, "styleMode"));
    }

    [Fact]
    public void Multiple_merges_fully_inside_the_target_are_captured()
    {
        using var home = new TestHome();
        var workbook = new ExcelFormatOnlySafetyTests.FormatWorkbook(@"C:\docbridge-fixtures\scope-multi-merge.xlsx", "Sheet1");
        workbook.Sheet.Merge(1, 1, 1, 2);
        workbook.Sheet.Merge(2, 1, 1, 2);
        using var adapter = CreateAdapter(workbook);

        adapter.CaptureSnapshot(CreateSnapshotDir(home), new JsonObject(), ExcelFormatOnlySafetyTests.FormatOps("A1:B2"));

        var state = ReadState(home);
        Assert.Equal(0, workbook.MultiCellMergeAreaReads);
        Assert.Equal(4, Json.GetInt(Json.GetObj(state, "coverage"), "cellCount"));
        Assert.True(Json.GetBool(Json.GetObj(state, "coverage"), "complete"));
    }

    [Fact]
    public void Multiple_merges_are_not_assumed_to_be_one_mergearea()
    {
        using var home = new TestHome();
        var workbook = new ExcelFormatOnlySafetyTests.FormatWorkbook(@"C:\docbridge-fixtures\scope-multi-merge-overflow.xlsx", "Sheet1");
        workbook.Sheet.Merge(1, 1, 1, 2);
        workbook.Sheet.Merge(2, 1, 1, 3);
        using var adapter = CreateAdapter(workbook);

        var ex = Assert.ThrowsAny<Exception>(() =>
            adapter.CaptureSnapshot(CreateSnapshotDir(home), new JsonObject(), ExcelFormatOnlySafetyTests.FormatOps("A1:B2")));
        var message = ex is AggregateException aggregate
            ? aggregate.Flatten().InnerExceptions[0].Message
            : ex.Message;
        Assert.Contains("[EXCEL_FORMAT_PARTIAL_MERGE]", message);
        Assert.Equal(0, workbook.MultiCellMergeAreaReads);
    }

    [Fact]
    public void Scoped_restore_rejects_unexpected_style_keys_and_range_mismatch()
    {
        using var home = new TestHome();
        var workbook = new ExcelFormatOnlySafetyTests.FormatWorkbook(@"C:\docbridge-fixtures\scope-shape-reject.xlsx", "Sheet1");
        using var adapter = CreateAdapter(workbook);
        var snapshotDir = CreateSnapshotDir(home);

        WriteState(snapshotDir, ScopedState(new JsonObject
        {
            ["sheet"] = "Sheet1",
            ["range"] = "A1",
            ["row"] = 1,
            ["column"] = 1,
            ["rows"] = 1,
            ["columns"] = 1,
            ["styleMode"] = ExcelAdapter.FormatStyleModeUniform,
            ["styleScope"] = ExcelAdapter.WrittenStyleScope,
            ["scopedProperties"] = new JsonArray("bold"),
            ["style"] = new JsonObject { ["bold"] = false, ["fillColor"] = 255 },
        }));
        var extraKey = adapter.RestoreSnapshot(snapshotDir, new JsonObject());
        Assert.False(Json.GetBool(extraKey, "ok"));
        Assert.Contains("unexpected", string.Join(" ", Json.GetArr(extraKey, "errors")!.Select(n => n!.GetValue<string>())), StringComparison.OrdinalIgnoreCase);

        WriteState(snapshotDir, ScopedState(new JsonObject
        {
            ["sheet"] = "Sheet1",
            ["range"] = "A1:B2",
            ["row"] = 1,
            ["column"] = 1,
            ["rows"] = 1,
            ["columns"] = 1,
            ["styleMode"] = ExcelAdapter.FormatStyleModeUniform,
            ["styleScope"] = ExcelAdapter.WrittenStyleScope,
            ["scopedProperties"] = new JsonArray("bold"),
            ["style"] = new JsonObject { ["bold"] = false },
        }));
        var mismatch = adapter.RestoreSnapshot(snapshotDir, new JsonObject());
        Assert.False(Json.GetBool(mismatch, "ok"));
        Assert.Contains("does not match", string.Join(" ", Json.GetArr(mismatch, "errors")!.Select(n => n!.GetValue<string>())), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Scoped_restore_rejects_coverage_over_the_format_cell_cap()
    {
        using var home = new TestHome();
        var workbook = new ExcelFormatOnlySafetyTests.FormatWorkbook(@"C:\docbridge-fixtures\scope-overflow.xlsx", "Sheet1");
        using var adapter = CreateAdapter(workbook);
        var snapshotDir = CreateSnapshotDir(home);
        WriteState(snapshotDir, new JsonObject
        {
            ["snapshotVersion"] = ExcelAdapter.FormatOnlyScopedSnapshotVersion,
            ["restoreMode"] = "format-only",
            ["styleScope"] = ExcelAdapter.WrittenStyleScope,
            ["documentRef"] = workbook.FullName,
            ["formatStates"] = new JsonArray(new JsonObject
            {
                ["sheet"] = "Sheet1",
                ["range"] = "A1",
                ["row"] = 1,
                ["column"] = 1,
                ["rows"] = 1,
                ["columns"] = 1,
                ["styleMode"] = ExcelAdapter.FormatStyleModeUniform,
                ["styleScope"] = ExcelAdapter.WrittenStyleScope,
                ["scopedProperties"] = new JsonArray("bold"),
                ["style"] = new JsonObject { ["bold"] = false },
            }),
            ["coverage"] = new JsonObject { ["complete"] = true, ["cellCount"] = 100001 },
        });

        var overflow = adapter.RestoreSnapshot(snapshotDir, new JsonObject());
        Assert.False(Json.GetBool(overflow, "ok"));
        Assert.Contains("100000", string.Join(" ", Json.GetArr(overflow, "errors")!.Select(n => n!.GetValue<string>())));
    }

    [Fact]
    public void Malformed_scopedProperties_rejects_arbitrary_keys_and_incomplete_color_couplings()
    {
        using var home = new TestHome();
        var workbook = new ExcelFormatOnlySafetyTests.FormatWorkbook(@"C:\docbridge-fixtures\scope-shape-reject.xlsx", "Sheet1");
        workbook.Sheet.Cell(1, 1).Font.Bold = false;
        workbook.Sheet.Cell(1, 1).Interior.Color = 255;
        using var adapter = CreateAdapter(workbook);
        var snapshotDir = CreateSnapshotDir(home);

        WriteState(snapshotDir, ScopedState(new JsonObject
        {
            ["sheet"] = "Sheet1",
            ["range"] = "A1",
            ["row"] = 1,
            ["column"] = 1,
            ["rows"] = 1,
            ["columns"] = 1,
            ["styleMode"] = ExcelAdapter.FormatStyleModeUniform,
            ["styleScope"] = ExcelAdapter.WrittenStyleScope,
            ["scopedProperties"] = new JsonArray("bold", "notAKey"),
            ["style"] = new JsonObject { ["bold"] = true, ["notAKey"] = 1 },
        }));
        var arbitrary = adapter.RestoreSnapshot(snapshotDir, new JsonObject());
        Assert.False(Json.GetBool(arbitrary, "ok"));
        Assert.Contains("unsupported", string.Join(" ", Json.GetArr(arbitrary, "errors")!.Select(n => n!.GetValue<string>())), StringComparison.OrdinalIgnoreCase);
        Assert.False((bool)workbook.Sheet.Cell(1, 1).Font.Bold!);
        Assert.Equal(255d, workbook.Sheet.Cell(1, 1).Interior.Color);

        WriteState(snapshotDir, ScopedState(new JsonObject
        {
            ["sheet"] = "Sheet1",
            ["range"] = "A1",
            ["row"] = 1,
            ["column"] = 1,
            ["rows"] = 1,
            ["columns"] = 1,
            ["styleMode"] = ExcelAdapter.FormatStyleModeUniform,
            ["styleScope"] = ExcelAdapter.WrittenStyleScope,
            ["scopedProperties"] = new JsonArray("fontColor"),
            ["style"] = new JsonObject { ["fontColor"] = 0 },
        }));
        var missingCoupled = adapter.RestoreSnapshot(snapshotDir, new JsonObject());
        Assert.False(Json.GetBool(missingCoupled, "ok"));
        Assert.Contains("color-coupled", string.Join(" ", Json.GetArr(missingCoupled, "errors")!.Select(n => n!.GetValue<string>())), StringComparison.OrdinalIgnoreCase);
        Assert.Equal(255d, workbook.Sheet.Cell(1, 1).Interior.Color);

        WriteState(snapshotDir, ScopedState(new JsonObject
        {
            ["sheet"] = "Sheet1",
            ["range"] = "A1",
            ["row"] = 1,
            ["column"] = 1,
            ["rows"] = 1,
            ["columns"] = 1,
            ["styleMode"] = ExcelAdapter.FormatStyleModeUniform,
            ["styleScope"] = ExcelAdapter.WrittenStyleScope,
            ["scopedProperties"] = new JsonArray("bold", "fillColorIndex"),
            ["style"] = new JsonObject { ["bold"] = true, ["fillColorIndex"] = 3 },
        }));
        var orphanCoupled = adapter.RestoreSnapshot(snapshotDir, new JsonObject());
        Assert.False(Json.GetBool(orphanCoupled, "ok"));
        Assert.Contains("exactly the written keys", string.Join(" ", Json.GetArr(orphanCoupled, "errors")!.Select(n => n!.GetValue<string>())), StringComparison.OrdinalIgnoreCase);
        Assert.False((bool)workbook.Sheet.Cell(1, 1).Font.Bold!);
        Assert.Equal(255d, workbook.Sheet.Cell(1, 1).Interior.Color);
    }

    [Fact]
    public void Malformed_A1_and_group_range_are_rejected_before_any_write()
    {
        using var home = new TestHome();
        var workbook = new ExcelFormatOnlySafetyTests.FormatWorkbook(@"C:\docbridge-fixtures\scope-range-parse.xlsx", "Sheet1");
        workbook.Sheet.Cell(1, 1).Font.Bold = false;
        workbook.Sheet.Cell(1, 2).Font.Bold = false;
        workbook.Sheet.Cell(1, 26).Font.Bold = false;
        using var adapter = CreateAdapter(workbook);
        var snapshotDir = CreateSnapshotDir(home);

        WriteState(snapshotDir, new JsonObject
        {
            ["snapshotVersion"] = ExcelAdapter.FormatOnlyScopedSnapshotVersion,
            ["restoreMode"] = "format-only",
            ["styleScope"] = ExcelAdapter.WrittenStyleScope,
            ["documentRef"] = workbook.FullName,
            ["formatStates"] = new JsonArray(
                new JsonObject
                {
                    ["sheet"] = "Sheet1",
                    ["range"] = "A1",
                    ["row"] = 1,
                    ["column"] = 1,
                    ["rows"] = 1,
                    ["columns"] = 1,
                    ["styleMode"] = ExcelAdapter.FormatStyleModeUniform,
                    ["styleScope"] = ExcelAdapter.WrittenStyleScope,
                    ["scopedProperties"] = new JsonArray("bold"),
                    ["style"] = new JsonObject { ["bold"] = true },
                },
                new JsonObject
                {
                    ["sheet"] = "Sheet1",
                    ["range"] = "A1:Z10",
                    ["row"] = 1,
                    ["column"] = 1,
                    ["rows"] = 1,
                    ["columns"] = 1,
                    ["styleMode"] = ExcelAdapter.FormatStyleModeUniform,
                    ["styleScope"] = ExcelAdapter.WrittenStyleScope,
                    ["scopedProperties"] = new JsonArray("bold"),
                    ["style"] = new JsonObject { ["bold"] = true },
                }),
            ["coverage"] = new JsonObject { ["complete"] = true, ["cellCount"] = 1 },
        });
        var oversized = adapter.RestoreSnapshot(snapshotDir, new JsonObject());
        Assert.False(Json.GetBool(oversized, "ok"));
        Assert.Contains("does not match row/column/rows/columns", string.Join(" ", Json.GetArr(oversized, "errors")!.Select(n => n!.GetValue<string>())));
        Assert.False((bool)workbook.Sheet.Cell(1, 1).Font.Bold!);
        Assert.False((bool)workbook.Sheet.Cell(1, 26).Font.Bold!);

        WriteState(snapshotDir, new JsonObject
        {
            ["snapshotVersion"] = ExcelAdapter.FormatOnlyScopedSnapshotVersion,
            ["restoreMode"] = "format-only",
            ["styleScope"] = ExcelAdapter.WrittenStyleScope,
            ["documentRef"] = workbook.FullName,
            ["formatStates"] = new JsonArray(new JsonObject
            {
                ["sheet"] = "Sheet1",
                ["range"] = "A1:B1",
                ["row"] = 1,
                ["column"] = 1,
                ["rows"] = 1,
                ["columns"] = 2,
                ["styleMode"] = ExcelAdapter.FormatStyleModeGroups,
                ["styleScope"] = ExcelAdapter.WrittenStyleScope,
                ["scopedProperties"] = new JsonArray("bold"),
                ["groups"] = new JsonArray(new JsonObject
                {
                    ["range"] = "A1:C1",
                    ["row"] = 1,
                    ["column"] = 1,
                    ["rows"] = 1,
                    ["columns"] = 1,
                    ["style"] = new JsonObject { ["bold"] = true },
                }),
            }),
            ["coverage"] = new JsonObject { ["complete"] = true, ["cellCount"] = 2 },
        });
        var groupMismatch = adapter.RestoreSnapshot(snapshotDir, new JsonObject());
        Assert.False(Json.GetBool(groupMismatch, "ok"));
        Assert.Contains("does not match row/column/rows/columns", string.Join(" ", Json.GetArr(groupMismatch, "errors")!.Select(n => n!.GetValue<string>())));
        Assert.False((bool)workbook.Sheet.Cell(1, 1).Font.Bold!);
        Assert.False((bool)workbook.Sheet.Cell(1, 2).Font.Bold!);

        WriteState(snapshotDir, new JsonObject
        {
            ["snapshotVersion"] = ExcelAdapter.FormatOnlyScopedSnapshotVersion,
            ["restoreMode"] = "format-only",
            ["styleScope"] = ExcelAdapter.WrittenStyleScope,
            ["documentRef"] = workbook.FullName,
            ["formatStates"] = new JsonArray(new JsonObject
            {
                ["sheet"] = "Sheet1",
                ["range"] = "A1048576:A1048577",
                ["row"] = 1_048_576,
                ["column"] = 1,
                ["rows"] = 2,
                ["columns"] = 1,
                ["styleMode"] = ExcelAdapter.FormatStyleModeUniform,
                ["styleScope"] = ExcelAdapter.WrittenStyleScope,
                ["scopedProperties"] = new JsonArray("bold"),
                ["style"] = new JsonObject { ["bold"] = true },
            }),
            ["coverage"] = new JsonObject { ["complete"] = true, ["cellCount"] = 2 },
        });
        var overflow = adapter.RestoreSnapshot(snapshotDir, new JsonObject());
        Assert.False(Json.GetBool(overflow, "ok"));
        Assert.Contains("outside Excel bounds", string.Join(" ", Json.GetArr(overflow, "errors")!.Select(n => n!.GetValue<string>())));
        Assert.False((bool)workbook.Sheet.Cell(1, 1).Font.Bold!);
    }

    [Fact]
    public void Color_group_restore_rejects_aggregate_black_when_cells_differ()
    {
        using var home = new TestHome();
        var workbook = new ExcelFormatOnlySafetyTests.FormatWorkbook(@"C:\docbridge-fixtures\scope-color-group-black.xlsx", "Sheet1");
        workbook.Sheet.PartialRangeStyleWrites = true;
        workbook.Sheet.Cell(1, 1).Interior.Color = 0;
        workbook.Sheet.Cell(1, 2).Interior.Color = 255;
        using var adapter = CreateAdapter(workbook);
        var snapshotDir = CreateSnapshotDir(home);

        WriteState(snapshotDir, new JsonObject
        {
            ["snapshotVersion"] = ExcelAdapter.FormatOnlyScopedSnapshotVersion,
            ["restoreMode"] = "format-only",
            ["styleScope"] = ExcelAdapter.WrittenStyleScope,
            ["documentRef"] = workbook.FullName,
            ["formatStates"] = new JsonArray(new JsonObject
            {
                ["sheet"] = "Sheet1",
                ["range"] = "A1:B1",
                ["row"] = 1,
                ["column"] = 1,
                ["rows"] = 1,
                ["columns"] = 2,
                ["styleMode"] = ExcelAdapter.FormatStyleModeGroups,
                ["styleScope"] = ExcelAdapter.WrittenStyleScope,
                ["scopedProperties"] = FillScopedProperties(),
                ["groups"] = new JsonArray(new JsonObject
                {
                    ["range"] = "A1:B1",
                    ["row"] = 1,
                    ["column"] = 1,
                    ["rows"] = 1,
                    ["columns"] = 2,
                    ["style"] = BlackFillStyle(),
                }),
            }),
            ["coverage"] = new JsonObject { ["complete"] = true, ["cellCount"] = 2 },
        });

        var restore = adapter.RestoreSnapshot(snapshotDir, new JsonObject());
        Assert.False(Json.GetBool(restore, "ok"), restore.ToJsonString());
        Assert.False(Json.GetBool(Json.GetObj(restore, "readback"), "verified"));
        Assert.Contains("mismatch", string.Join(" ", Json.GetArr(restore, "errors")!.Select(n => n!.GetValue<string>())), StringComparison.OrdinalIgnoreCase);
        Assert.Equal(255d, workbook.Sheet.Cell(1, 2).Interior.Color);
    }

    private static JsonArray FillScopedProperties() => new(
        "fillColor",
        "fillColorIndex",
        "fillTintAndShade",
        "fillPattern",
        "fillPatternColor",
        "fillPatternColorIndex",
        "fillPatternTintAndShade");

    private static JsonObject BlackFillStyle() => new()
    {
        ["fillColor"] = 0,
        ["fillColorIndex"] = 1,
        ["fillTintAndShade"] = 0,
        ["fillPattern"] = ExcelStyleContract.XlPatternSolid,
        ["fillPatternColor"] = 0,
        ["fillPatternColorIndex"] = ExcelStyleContract.XlColorIndexNone,
        ["fillPatternTintAndShade"] = 0,
    };

    private static JsonObject ScopedState(JsonObject formatState) => new()
    {
        ["snapshotVersion"] = ExcelAdapter.FormatOnlyScopedSnapshotVersion,
        ["restoreMode"] = "format-only",
        ["styleScope"] = ExcelAdapter.WrittenStyleScope,
        ["documentRef"] = @"C:\docbridge-fixtures\scope-shape-reject.xlsx",
        ["formatStates"] = new JsonArray(formatState),
        ["coverage"] = new JsonObject { ["complete"] = true, ["cellCount"] = 1 },
    };

    private static ExcelAdapter CreateAdapter(ExcelFormatOnlySafetyTests.FormatWorkbook workbook) =>
        new(() => new ExcelFormatOnlySafetyTests.FormatExcelApplication(workbook));

    private static string CreateSnapshotDir(TestHome home)
    {
        var dir = Path.Combine(home.Dir, "snapshot");
        Directory.CreateDirectory(dir);
        return dir;
    }

    private static JsonObject ReadState(TestHome home) =>
        JsonNode.Parse(File.ReadAllText(Path.Combine(home.Dir, "snapshot", "state.json")))!.AsObject();

    private static void WriteState(string snapshotDir, JsonObject state)
    {
        Directory.CreateDirectory(snapshotDir);
        File.WriteAllText(Path.Combine(snapshotDir, "state.json"), state.ToJsonString());
    }

    private static JsonObject ValidFullFormatState() => new()
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
}
