using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelFormatUnionTests
{
    [Fact]
    public void Range_reference_keeps_comma_union_address()
    {
        var bare = ExcelRangeReference.Parse("A1,A3");
        Assert.Null(bare.SheetName);
        Assert.Equal("A1,A3", bare.Address);

        var sheeted = ExcelRangeReference.Parse("'예정공정표'!A1:B2,D1:D3");
        Assert.Equal("예정공정표", sheeted.SheetName);
        Assert.Equal("A1:B2,D1:D3", sheeted.Address);
        Assert.False(ExcelA1Box.TryParse("A1,A3", out _));
    }

    [Fact]
    public void Two_by_two_plus_three_by_one_has_no_off_target_cells()
    {
        const string union = "A1:B2,D1:D3";
        var intended = new[] { "A1", "B1", "A2", "B2", "D1", "D2", "D3" };
        var cells = ExcelFormatUnion.EnumerateCells(union);

        Assert.Equal(intended.Length, cells.Count);
        foreach (var cell in intended)
            Assert.Contains(cell, cells);
        foreach (var off in new[] { "C1", "C2", "C3", "A3", "B3", "E1", "D4", "C4" })
            Assert.DoesNotContain(off, cells);
        Assert.False(ExcelFormatUnion.HasOffTargetCell(union, intended));
        Assert.True(ExcelFormatUnion.HasOffTargetCell("A1:D3", intended));
    }

    [Fact]
    public void Bounded_union_plan_batches_equal_style_areas_without_one_com_per_cell()
    {
        var cells = new List<string>();
        for (var row = 1; row <= 80; row++)
            cells.Add($"A{row}");

        var plans = ExcelFormatUnion.PlanBoundedUnions(cells, maxAreas: 64);
        Assert.Equal(2, plans.Count);
        Assert.Equal(64, ExcelFormatUnion.SplitUnionAddresses(plans[0]).Count);
        Assert.Equal(16, ExcelFormatUnion.SplitUnionAddresses(plans[1]).Count);
        Assert.False(ExcelFormatUnion.HasOffTargetCell(plans[0], cells));
        Assert.False(ExcelFormatUnion.HasOffTargetCell(plans[1], cells));
        Assert.Equal(80, ExcelFormatUnion.EnumerateCells(string.Join(',', plans)).Count);
    }
}
