using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelMergeSnapshotPerformanceTests
{
    [Fact]
    public void Uniform_grid_expands_every_cell_and_old_reader_can_read_styles()
    {
        var style = SampleStyle("돋움", 255);
        var encoded = ExcelMergeSnapshotPerformance.EncodeUniformStyleRange("A1:C2", 2, 3, style);

        Assert.Equal(ExcelMergeSnapshotPerformance.CaptureUniform, encoded["capture"]!.GetValue<string>());
        Assert.True(ExcelMergeSnapshotPerformance.IsUniformCapture(encoded));
        Assert.True(ExcelMergeSnapshotPerformance.TryReadStyleGrid(encoded, out var address, out var styles));
        Assert.Equal("A1:C2", address);
        Assert.Equal(2, styles!.Count);
        Assert.Equal(3, styles[0]!.AsArray().Count);
        Assert.True(ExcelMergeSnapshotPerformance.CanEncodeUniformFromGrid(styles));
        Assert.False(ExcelMergeSnapshotPerformance.FirstCellOnlyWouldHideMix(styles));
        Assert.True(JsonNode.DeepEquals(style, styles[1]![2]));
    }

    [Fact]
    public void Mixed_second_cell_is_not_uniform_and_first_cell_only_would_hide_it()
    {
        var first = SampleStyle("돋움", 255);
        var mixed = SampleStyle("맑은 고딕", 16711680);
        var grid = new JsonArray
        {
            new JsonArray { first.DeepClone(), mixed.DeepClone() },
        };

        Assert.False(ExcelMergeSnapshotPerformance.CanEncodeUniformFromGrid(grid));
        Assert.True(ExcelMergeSnapshotPerformance.FirstCellOnlyWouldHideMix(grid));
        Assert.False(ExcelMergeSnapshotPerformance.TryGetUniformStyle(grid, out _));

        var encoded = ExcelMergeSnapshotPerformance.EncodeMixedStyleRange("G1:H1", 1, 2, grid);
        Assert.Equal(ExcelMergeSnapshotPerformance.CaptureMixed, encoded["capture"]!.GetValue<string>());
        Assert.False(ExcelMergeSnapshotPerformance.IsUniformCapture(encoded));
        Assert.True(ExcelMergeSnapshotPerformance.TryReadStyleGrid(encoded, out _, out var styles));
        Assert.False(JsonNode.DeepEquals(styles![0]![0], styles[0]![1]));
    }

    [Fact]
    public void Legacy_styles_only_snapshot_still_reads()
    {
        var style = SampleStyle("돋움", 0);
        var legacy = new JsonObject
        {
            ["range"] = "A15:A24",
            ["styles"] = ExcelMergeSnapshotPerformance.ExpandUniformStyles(style, 10, 1),
        };
        Assert.True(ExcelMergeSnapshotPerformance.TryReadStyleGrid(legacy, out var address, out var styles));
        Assert.Equal("A15:A24", address);
        Assert.Equal(10, styles!.Count);
        Assert.False(legacy.ContainsKey("capture"));
    }

    [Fact]
    public void Compact_uniform_without_styles_grid_still_materializes()
    {
        var style = SampleStyle("돋움", 255);
        var compact = new JsonObject
        {
            ["range"] = "A1:BQ1",
            ["rows"] = 1,
            ["columns"] = 69,
            ["capture"] = ExcelMergeSnapshotPerformance.CaptureUniform,
            ["style"] = style.DeepClone(),
        };
        Assert.True(ExcelMergeSnapshotPerformance.TryReadStyleGrid(compact, out _, out var styles));
        var row = Assert.Single(styles!);
        Assert.Equal(69, row!.AsArray().Count);
        Assert.True(JsonNode.DeepEquals(style, row.AsArray()[68]));
    }

    [Fact]
    public void Range_color_zero_plus_last_cell_does_not_prove_middle()
    {
        Assert.True(ExcelMergeSnapshotPerformance.RangeAggregateColorIsAmbiguous(0, cellCount: 3));
        Assert.False(ExcelMergeSnapshotPerformance.RangeAggregateColorIsAmbiguous(255, cellCount: 3));
        Assert.False(ExcelMergeSnapshotPerformance.RangeAggregateColorIsAmbiguous(0, cellCount: 1));
        Assert.False(ExcelMergeSnapshotPerformance.CellColorsProveUniform(new[] { 0d, 255d, 0d }));
        Assert.True(ExcelMergeSnapshotPerformance.CellColorsProveUniform(new[] { 0d, 0d, 0d }));
    }

    [Fact]
    public void Cell_border_readback_requires_every_cell_and_color()
    {
        var requested = new[]
        {
            new ExcelBorderContract.EdgeSpec("left", ExcelBorderContract.ScopeCell, "medium", "continuous", 255),
        };
        var good = new JsonArray
        {
            new JsonArray { BorderCell(255), BorderCell(255) },
            new JsonArray { BorderCell(255), BorderCell(255) },
        };
        var errors = new List<string>();
        Assert.True(ExcelMergeSnapshotPerformance.RequestedCellBordersMatch(requested, good, "예정공정표!G2:H3", errors));
        Assert.Empty(errors);

        var wrongColor = new JsonArray
        {
            new JsonArray { BorderCell(255), BorderCell(16711680) },
        };
        errors.Clear();
        Assert.False(ExcelMergeSnapshotPerformance.RequestedCellBordersMatch(
            requested, wrongColor, "예정공정표!G2:H2", errors));
        Assert.Contains(errors, error => error.Contains("[1,2]", StringComparison.Ordinal));
        Assert.True(
            ExcelMergeSnapshotPerformance.FirstCellOnlyIgnoreColorWouldAccept(requested, wrongColor),
            "old first-cell + ignore-color readback would hide the non-first wrong color");
    }

    private static JsonObject SampleStyle(string fontName, double fontColor) => new()
    {
        ["bold"] = true,
        ["italic"] = false,
        ["fontName"] = fontName,
        ["fontSize"] = 28,
        ["fontColor"] = fontColor,
        ["fillColor"] = 16777215,
        ["numberFormat"] = "General",
    };

    private static JsonObject BorderCell(double color) => new()
    {
        ["left"] = new JsonObject
        {
            ["lineStyle"] = "continuous",
            ["weight"] = "medium",
            ["color"] = color,
        },
    };
}
