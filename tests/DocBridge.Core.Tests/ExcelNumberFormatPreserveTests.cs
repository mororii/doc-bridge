using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelNumberFormatPreserveTests
{
    [Fact]
    public void Mixed_run_null_without_per_cell_refuses_before_at_write()
    {
        Assert.True(ExcelNumberFormatPreserve.NeedsPerCellCapture(null));
        Assert.False(ExcelNumberFormatPreserve.TryNormalizeCapture(
            null, perCell: null, length: 4, out _, out var error));
        Assert.Contains("refusing", error, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("per-cell", error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Mixed_General_fixed_date_and_custom_restore_through_helper()
    {
        var fixture = ExcelNumberFormatPreserve.MixedRunFixtureFormats;
        Assert.Equal("General", fixture[0]);
        Assert.Equal("0.00", fixture[1]);
        Assert.Equal("yyyy-mm-dd", fixture[2]);
        Assert.Equal("\"yyyy\"", fixture[3]);

        Assert.True(ExcelNumberFormatPreserve.TryNormalizeCapture(
            raw: null, perCell: fixture, length: 4, out var captured, out var error));
        Assert.Null(error);
        Assert.True(captured.Mixed);
        Assert.True(captured.RestorePerCell);
        Assert.Equal(fixture, captured.Formats);

        var afterAt = new[] { "@", "@", "@", "@" };
        Assert.False(
            ExcelNumberFormatPreserve.FormatsRetained(captured.Formats, afterAt),
            "leaving @ after a mixed-null capture loses General/0.00/date/custom");

        var restored = captured.Formats.ToArray();
        Assert.True(ExcelNumberFormatPreserve.FormatsRetained(captured.Formats, restored));
        Assert.True(ExcelNumberFormatPreserve.FormatsRetained(fixture, restored));
    }

    [Fact]
    public void Uniform_string_does_not_require_per_cell_and_repeats()
    {
        Assert.False(ExcelNumberFormatPreserve.NeedsPerCellCapture("0.00"));
        Assert.True(ExcelNumberFormatPreserve.TryNormalizeCapture(
            "0.00", perCell: null, length: 3, out var captured, out var error));
        Assert.Null(error);
        Assert.False(captured.Mixed);
        Assert.False(captured.RestorePerCell);
        Assert.Equal(new[] { "0.00", "0.00", "0.00" }, captured.Formats);
        Assert.True(ExcelNumberFormatPreserve.FormatsRetained(
            captured.Formats, new[] { "0.00", "0.00", "0.00" }));
    }

    [Fact]
    public void Mixed_array_and_partial_per_cell_failure_refuse_write()
    {
        object[,] mixed =
        {
            { "General", "0.00", "yyyy-mm-dd", "\"yyyy\"" },
        };
        Assert.True(ExcelNumberFormatPreserve.NeedsPerCellCapture(mixed));
        Assert.True(ExcelNumberFormatPreserve.TryNormalizeCapture(
            mixed, perCell: null, length: 4, out var fromArray, out _));
        Assert.True(fromArray.Mixed);
        Assert.Equal(ExcelNumberFormatPreserve.MixedRunFixtureFormats, fromArray.Formats);

        Assert.False(ExcelNumberFormatPreserve.TryNormalizeCapture(
            null, perCell: new[] { "General", "", "0.00", "yyyy-mm-dd" }, length: 4, out _, out var emptyCell));
        Assert.Contains("index 1", emptyCell, StringComparison.Ordinal);
        Assert.False(ExcelNumberFormatPreserve.TryNormalizeCapture(
            null, perCell: new[] { "General" }, length: 4, out _, out var lengthMismatch));
        Assert.Contains("refusing", lengthMismatch, StringComparison.OrdinalIgnoreCase);
        Assert.False(ExcelNumberFormatPreserve.FormatsRetained(
            new[] { "General", "0.00" }, new[] { "General" }));
        Assert.True(ExcelNumberFormatPreserve.FormatsRetained(
            new[] { "General" }, new[] { "G/표준" }));
        Assert.False(ExcelNumberFormatPreserve.FormatsRetained(
            new[] { "0.00" }, new[] { "General" }));
    }
}
