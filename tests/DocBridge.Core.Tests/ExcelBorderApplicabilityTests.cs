using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelBorderApplicabilityTests
{
    [Fact]
    public void One_row_A4_H4_all_does_not_read_insideHorizontal()
    {
        Assert.False(ExcelBorderApplicability.InsideEdgeApplies("insideHorizontal", rows: 1, columns: 8));
        Assert.True(ExcelBorderApplicability.InsideEdgeApplies("insideVertical", rows: 1, columns: 8));
        Assert.True(ExcelBorderApplicability.InsideEdgeApplies("left", rows: 1, columns: 8));

        var errors = new List<string>();
        Assert.True(ExcelBorderContract.TryNormalize(
            new JsonObject
            {
                ["all"] = new JsonObject
                {
                    ["weight"] = "thin",
                    ["lineStyle"] = "continuous",
                    ["color"] = 0,
                },
            },
            0, errors, out var canonical));
        Assert.Empty(errors);
        Assert.Contains(ExcelBorderContract.ReadEdges(canonical), spec =>
            spec.Name.Equals("insideHorizontal", StringComparison.OrdinalIgnoreCase));

        var filtered = ExcelBorderApplicability.ForRange(canonical, rows: 1, columns: 8);
        var requested = ExcelBorderContract.ReadEdges(filtered);
        Assert.DoesNotContain(requested, spec =>
            spec.Name.Equals("insideHorizontal", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(requested, spec =>
            spec.Name.Equals("insideVertical", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(requested, spec => spec.Name.Equals("left", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(requested, spec => spec.Name.Equals("bottom", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void One_column_keeps_insideHorizontal_and_drops_insideVertical()
    {
        Assert.True(ExcelBorderApplicability.InsideEdgeApplies("insideHorizontal", rows: 3, columns: 1));
        Assert.False(ExcelBorderApplicability.InsideEdgeApplies("insideVertical", rows: 3, columns: 1));
        var explicitInsides = new JsonObject
        {
            ["insideHorizontal"] = true,
            ["insideVertical"] = true,
            ["bottom"] = true,
        };
        var filtered = ExcelBorderApplicability.ForRange(explicitInsides, rows: 3, columns: 1);
        Assert.NotNull(filtered["insideHorizontal"]);
        Assert.Null(filtered["insideVertical"]);
        Assert.NotNull(filtered["bottom"]);
    }

    [Fact]
    public void One_by_one_drops_both_insides()
    {
        Assert.False(ExcelBorderApplicability.InsideEdgeApplies("insideHorizontal", 1, 1));
        Assert.False(ExcelBorderApplicability.InsideEdgeApplies("insideVertical", 1, 1));
    }

    [Fact]
    public void Merge_B9_F9_in_A5_H11_hides_only_interior_merge_edges()
    {
        var geometry = new ExcelBorderApplicability.AreaGeometry(
            7, 8, [new ExcelBorderApplicability.MergeRect(5, 2, 1, 5)]);
        Assert.True(ExcelBorderApplicability.HasVisibleInteriorVertical(geometry));
        Assert.True(ExcelBorderApplicability.HasVisibleInteriorHorizontal(geometry));
        Assert.True(ExcelBorderApplicability.SameMerge(geometry, 5, 2, 5, 6));
        Assert.False(ExcelBorderApplicability.SameMerge(geometry, 5, 1, 5, 2));
        Assert.False(ExcelBorderApplicability.IsVisibleCellEdge(geometry, 5, 3, "left"));
        Assert.False(ExcelBorderApplicability.IsVisibleCellEdge(geometry, 5, 4, "right"));
        Assert.True(ExcelBorderApplicability.IsVisibleCellEdge(geometry, 5, 2, "left"));
        Assert.True(ExcelBorderApplicability.IsVisibleCellEdge(geometry, 5, 6, "right"));
        Assert.True(ExcelBorderApplicability.EdgeApplies("insideVertical", geometry));
        Assert.Contains(
            "no interior vertical",
            ExcelBorderApplicability.SkipReason("insideVertical", ExcelBorderApplicability.AreaGeometry.Ungrouped(7, 1)));
    }

    [Fact]
    public void MergeCells_false_is_zero_scan_true_is_one_merge_area()
    {
        ExcelBorderApplicability.ResetMergeScanCells();
        var unmerged = ExcelBorderReadbackTests.FakeColumn.OneByEight(
            new ExcelBorderReadbackTests.FakeBorder
            {
                LineStyle = ExcelBorderContract.XlContinuous,
                Weight = ExcelBorderContract.XlThin,
                Color = 0,
            },
            insideHorizontal: new ExcelBorderReadbackTests.FakeBorder
            {
                LineStyle = ExcelBorderContract.XlLineStyleNone,
                Weight = ExcelBorderContract.XlThin,
                Color = 0,
            });
        var geometry = ExcelBorderApplicability.ReadGeometry(unmerged);
        Assert.Empty(geometry.Merges);
        Assert.Equal(0, ExcelBorderApplicability.MergeScanCells);

        ExcelBorderApplicability.ResetMergeScanCells();
        var merged = ExcelBorderReadbackTests.FakeColumn.OneByEightMerged(
            new ExcelBorderReadbackTests.FakeBorder
            {
                LineStyle = ExcelBorderContract.XlContinuous,
                Weight = ExcelBorderContract.XlThin,
                Color = 0,
            },
            new ExcelBorderReadbackTests.FakeBorder
            {
                LineStyle = ExcelBorderContract.XlLineStyleNone,
                Weight = ExcelBorderContract.XlThin,
                Color = 0,
            });
        var whole = ExcelBorderApplicability.ReadGeometry(merged);
        Assert.Single(whole.Merges);
        Assert.Equal(1, whole.Merges[0].Rows);
        Assert.Equal(8, whole.Merges[0].Columns);
        Assert.Equal(0, ExcelBorderApplicability.MergeScanCells);

        ExcelBorderApplicability.ResetMergeScanCells();
        var partial = ExcelBorderReadbackTests.FakeColumn.OneByEightContainsPartialMerge(
            new ExcelBorderReadbackTests.FakeBorder
            {
                LineStyle = ExcelBorderContract.XlContinuous,
                Weight = ExcelBorderContract.XlThin,
                Color = 0,
            });
        Assert.True(Convert.ToBoolean(partial.MergeCells));
        var scanned = ExcelBorderApplicability.ReadGeometry(partial);
        Assert.True(ExcelBorderApplicability.MergeScanCells >= 8);
        Assert.Single(scanned.Merges);
        Assert.Equal(2, scanned.Merges[0].Column);
        Assert.Equal(2, scanned.Merges[0].Columns);

        var thrown = Assert.Throws<InvalidOperationException>(() => _ = merged.MergeArea);
        Assert.Contains("single-cell", thrown.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Fully_merged_one_row_has_no_inside_vertical()
    {
        var geometry = new ExcelBorderApplicability.AreaGeometry(
            1, 8, [new ExcelBorderApplicability.MergeRect(1, 1, 1, 8)]);
        Assert.False(ExcelBorderApplicability.HasVisibleInteriorVertical(geometry));
        Assert.False(ExcelBorderApplicability.EdgeApplies("insideHorizontal", geometry));
        Assert.False(ExcelBorderApplicability.EdgeApplies("insideVertical", geometry));
    }
}
