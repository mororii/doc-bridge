using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelRangeReadContractTests
{
    [Fact]
    public void Plan_single_cell_returns_complete_one_by_one_page()
    {
        var page = ExcelRangeReadContract.Plan(new JsonObject(), totalRows: 1, totalColumns: 1, maxAllowedCells: 10_000);

        Assert.Equal((0, 0, 1, 1, 1), (page.RowOffset, page.ColumnOffset, page.Rows, page.Columns, page.Cells));
        Assert.True(page.Complete);
        Assert.Null(page.Continuation);
    }

    [Fact]
    public void Plan_exact_cap_returns_complete_rectangular_page()
    {
        var page = ExcelRangeReadContract.Plan(new JsonObject { ["maxCells"] = 6 }, totalRows: 2, totalColumns: 3, maxAllowedCells: 10_000);

        Assert.Equal((2, 3, 6), (page.Rows, page.Columns, page.Cells));
        Assert.True(page.Complete);
    }

    [Fact]
    public void Plan_tail_page_continues_without_missing_or_duplicate_cells()
    {
        var first = ExcelRangeReadContract.Plan(new JsonObject { ["maxCells"] = 4 }, totalRows: 3, totalColumns: 2, maxAllowedCells: 10_000);
        var second = ExcelRangeReadContract.Plan(first.Continuation!, totalRows: 3, totalColumns: 2, maxAllowedCells: 10_000);

        Assert.Equal((0, 0, 2, 2), (first.RowOffset, first.ColumnOffset, first.Rows, first.Columns));
        Assert.Equal((2, 0, 1, 2), (second.RowOffset, second.ColumnOffset, second.Rows, second.Columns));
        Assert.Equal(2, second.Cells);
    }

    [Fact]
    public void Plan_rows_wider_than_cap_pages_columns_before_rows()
    {
        var page = ExcelRangeReadContract.Plan(new JsonObject { ["maxCells"] = 5 }, totalRows: 2, totalColumns: 8, maxAllowedCells: 10_000);

        Assert.Equal((0, 0, 1, 5), (page.RowOffset, page.ColumnOffset, page.Rows, page.Columns));
        Assert.Equal(5, page.Cells);
        Assert.Equal(5, Json.GetInt(page.Continuation, "columnOffset"));
        Assert.Equal(0, Json.GetInt(page.Continuation, "rowOffset"));
    }

    [Theory]
    [InlineData(2, 3, 2, 0)]
    [InlineData(4, 5, 5, 2)]
    public void Plan_complete_traversal_covers_every_cell_exactly_once(
        int totalRows, int totalColumns, int maxCells, int maxColumns)
    {
        var args = new JsonObject { ["maxCells"] = maxCells };
        if (maxColumns > 0) args["maxColumns"] = maxColumns;
        var visited = new HashSet<(int Row, int Column)>();
        ExcelRangeReadContract.Page? last = null;

        while (true)
        {
            var page = ExcelRangeReadContract.Plan(args, totalRows, totalColumns, 10_000);
            for (var row = page.RowOffset; row < page.RowOffset + page.Rows; row++)
            for (var column = page.ColumnOffset; column < page.ColumnOffset + page.Columns; column++)
                Assert.True(visited.Add((row, column)), $"duplicate cell ({row}, {column})");
            last = page;
            if (page.Continuation is null) break;
            args = page.Continuation;
        }

        Assert.Equal(totalRows * totalColumns, visited.Count);
        Assert.All(Enumerable.Range(0, totalRows).SelectMany(row =>
            Enumerable.Range(0, totalColumns).Select(column => (row, column))),
            cell => Assert.Contains(cell, visited));
        Assert.False(last!.Complete);
        Assert.False(last.HasMore);
    }

    [Theory]
    [InlineData("rowOffset", -1)]
    [InlineData("columnOffset", -1)]
    [InlineData("maxRows", 0)]
    [InlineData("maxColumns", 0)]
    [InlineData("maxCells", 0)]
    public void Plan_rejects_invalid_bounds(string name, int value)
    {
        var args = new JsonObject { [name] = value };
        Assert.Throws<InvalidOperationException>(() => ExcelRangeReadContract.Plan(args, 2, 2, 10_000));
    }

    [Fact]
    public void Plan_rejects_offset_outside_requested_range()
    {
        var ex = Assert.Throws<InvalidOperationException>(() => ExcelRangeReadContract.Plan(
            new JsonObject { ["rowOffset"] = 2 }, 2, 2, 10_000));
        Assert.Contains("rowOffset", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Plan_rejects_max_cells_above_the_public_cap()
    {
        var ex = Assert.Throws<InvalidOperationException>(() => ExcelRangeReadContract.Plan(
            new JsonObject { ["maxCells"] = 10_001 }, 100, 100, 10_000));
        Assert.Contains("10000", ex.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void Formula_mode_defaults_to_formula_and_requires_explicit_formula2()
    {
        Assert.Equal(ExcelRangeReadContract.FormulaMode.Formula,
            ExcelRangeReadContract.ResolveFormulaMode(new JsonObject()));
        Assert.Equal(ExcelRangeReadContract.FormulaMode.Formula2,
            ExcelRangeReadContract.ResolveFormulaMode(new JsonObject { ["formulaMode"] = "formula2" }));
        Assert.Throws<InvalidOperationException>(() => ExcelRangeReadContract.ResolveFormulaMode(
            new JsonObject { ["formulaMode"] = "dynamic" }));
    }
}
