using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelFormulaTraceBatchTests
{
    [Fact]
    public void Overlapping_reads_reuse_cells_and_batch_new_rows_without_cross_request_cache()
    {
        var calls = new List<ExcelA1Box>();
        var first = new ExcelAdapter.FormulaTraceReadContext((sheet, box) =>
        {
            calls.Add(box);
            return Cells(sheet, box, 1);
        });
        var initial = first.Read("Sheet", new ExcelA1Box(1, 1, 3, 2), 6);
        var expanded = first.Read("Sheet", new ExcelA1Box(1, 1, 5, 2), 10);
        var repeated = first.Read("Sheet", new ExcelA1Box(1, 1, 5, 2), 10);
        Assert.Equal(new[] { "A1:B3", "A4:B5" }, calls.Select(x => x.Address));
        Assert.Equal(10, repeated.Count);
        Assert.Same(initial[0], expanded[0]);
        Assert.Same(expanded[9], repeated[9]);
        var second = new ExcelAdapter.FormulaTraceReadContext((sheet, box) => Cells(sheet, box, 2));
        Assert.Equal(2, second.Read("Sheet", new ExcelA1Box(1, 1, 1, 1), 1)[0].Evidence["value"]!.GetValue<int>());
        Assert.Equal(1, first.Read("Sheet", new ExcelA1Box(1, 1, 1, 1), 1)[0].Evidence["value"]!.GetValue<int>());
    }

    [Fact]
    public void Tall_column_is_one_rectangle_and_cached_holes_are_never_reread()
    {
        var tall = ExcelFormulaTraceReadPlan.MissingRectangles(new ExcelA1Box(1, 1, 500, 1), (_, _) => false);
        Assert.Single(tall);
        Assert.Equal("A1:A500", tall[0].Address);
        var plan = ExcelFormulaTraceReadPlan.MissingRectangles(new ExcelA1Box(1, 1, 3, 3), (r, c) => r == 2 && c == 2);
        var cells = plan.SelectMany(box => Enumerable.Range(box.Row, box.Rows)
            .SelectMany(row => Enumerable.Range(box.Column, box.Columns).Select(col => (row, col)))).ToArray();
        Assert.Equal(8, cells.Length);
        Assert.Equal(8, cells.Distinct().Count());
        Assert.DoesNotContain((2, 2), cells);
    }

    [Fact]
    public void Array_mapping_preserves_types_and_distinguishes_blank_from_missing_shape()
    {
        var values = Array.CreateInstance(typeof(object), new[] { 1, 4 }, new[] { 1, 1 });
        values.SetValue(null, 1, 1); values.SetValue(true, 1, 2);
        values.SetValue("001", 1, 3); values.SetValue(2.5, 1, 4);
        foreach (var column in Enumerable.Range(1, 4))
        {
            Assert.True(ExcelFormulaTraceReadPlan.TryValueAt(values, 1, column, 1, 4, out var value));
            Assert.Equal(values.GetValue(1, column), value);
        }
        Assert.True(ExcelFormulaTraceReadPlan.TryValueAt(null, 1, 1, 1, 1, out _));
        Assert.False(ExcelFormulaTraceReadPlan.TryValueAt(null, 1, 1, 1, 4, out _));
        Assert.False(ExcelFormulaTraceReadPlan.TryValueAt(values, 1, 1, 2, 4, out _));
    }

    private static IReadOnlyList<ExcelFormulaTraceContract.Cell> Cells(string sheet, ExcelA1Box box, int value) =>
        Enumerable.Range(box.Row, box.Rows).SelectMany(row => Enumerable.Range(box.Column, box.Columns)
            .Select(col => new ExcelFormulaTraceContract.Cell(sheet, ExcelA1Box.CellName(col, row), null,
                new JsonObject { ["value"] = value }))).ToArray();
}
