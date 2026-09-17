using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelBorderReadbackTests
{
    [Fact]
    public void A1_A3_bottom_does_not_pass_from_outer_edge_alone()
    {
        Assert.True(ExcelBorderReadback.ScopeCellNeedsInsideCompanion("bottom", rows: 3, columns: 1));
        Assert.Equal("insideHorizontal", ExcelBorderReadback.InsideCompanionName("bottom"));

        var requested = BorderCanonical(255);
        var range = FakeColumn.A1A3(
            outerBottom: MatchBorder(255),
            insideHorizontal: NoneBorder(),
            a1Bottom: NoneBorder(),
            a2Bottom: NoneBorder(),
            a3Bottom: MatchBorder(255));

        Assert.False(
            ExcelBorderReadback.AppliedBordersMatch(range, requested),
            "xlEdgeBottom is only A3; A1/A2 bottoms must be proved via xlInsideHorizontal or per-cell");
    }

    [Fact]
    public void A1_A3_bottom_scans_cells_when_inside_is_mixed_and_rejects_wrong_color()
    {
        var requested = BorderCanonical(255);
        var range = FakeColumn.A1A3(
            outerBottom: MatchBorder(255),
            insideHorizontal: MixedBorder(),
            a1Bottom: MatchBorder(255),
            a2Bottom: MatchBorder(16711680),
            a3Bottom: MatchBorder(255));

        Assert.False(ExcelBorderReadback.AppliedBordersMatch(range, requested));
        Assert.False(ExcelBorderReadback.OneBorderMatches(range.Cells.Item(2, 1), BottomSpec(255)));
    }

    [Fact]
    public void A1_A3_bottom_passes_when_outer_and_inside_match_including_color()
    {
        var requested = BorderCanonical(255);
        var range = FakeColumn.A1A3(
            outerBottom: MatchBorder(255),
            insideHorizontal: MatchBorder(255),
            a1Bottom: MatchBorder(255),
            a2Bottom: MatchBorder(255),
            a3Bottom: MatchBorder(255));

        Assert.True(ExcelBorderReadback.AppliedBordersMatch(range, requested));
    }

    [Fact]
    public void Color_zero_on_multi_cell_cannot_prove_and_middle_red_fails()
    {
        var requested = BorderCanonical(0);
        var range = FakeColumn.A1A3(
            outerBottom: MatchBorder(0),
            insideHorizontal: MatchBorder(0),
            a1Bottom: MatchBorder(0),
            a2Bottom: MatchBorder(16711680),
            a3Bottom: MatchBorder(0));

        Assert.True(ExcelBorderReadback.ColorZeroCannotProve(0d, BottomSpec(0), cellCount: 3));
        Assert.False(
            ExcelBorderReadback.TryUnmixedBorderMatch(range, BottomSpec(0), out _),
            "Color=0 on A1:A3 is mixed-or-black; fast proof must refuse");
        Assert.False(
            ExcelBorderReadback.AppliedBordersMatch(range, requested),
            "requested black + middle red must fail on the adapter AppliedBordersMatch branch");
    }

    [Fact]
    public void Color_zero_on_single_cell_is_real_black()
    {
        var cell = FakeColumn.Cell(MatchBorder(0));
        Assert.False(ExcelBorderReadback.ColorZeroCannotProve(0d, BottomSpec(0), cellCount: 1));
        Assert.True(ExcelBorderReadback.TryUnmixedBorderMatch(cell, BottomSpec(0), out var match));
        Assert.True(match);
        Assert.True(ExcelBorderReadback.AppliedBordersMatch(cell, BorderCanonical(0)));
    }

    [Fact]
    public void Union_A1_A3_second_area_wrong_color_fails_on_adapter_readback()
    {
        var requested = BorderCanonical(255);
        var a1 = FakeColumn.Cell(MatchBorder(255), bold: true);
        var a3 = FakeColumn.Cell(MatchBorder(16711680), bold: true);
        var union = FakeColumn.UnionFirstAreaOnly(a1, a3);

        Assert.Equal(1, union.Rows.Count);
        Assert.Equal(1, union.Columns.Count);
        Assert.True(
            ExcelBorderReadback.OneBorderMatches(union.Cells.Item(1, 1), BottomSpec(255)),
            "first-area Rows/Columns would hide A3");
        Assert.False(
            ExcelBorderReadback.AppliedBordersMatch(union, requested),
            "format_range A1,A3 must verify both Areas");
    }

    [Fact]
    public void Contiguous_A1_A3_applies_edge_plus_inside_not_each_cell()
    {
        var requested = BorderCanonical(255);
        var range = FakeColumn.A1A3(
            MatchBorder(0), MixedBorder(), MatchBorder(0), MatchBorder(0), MatchBorder(0));
        var applied = new List<(object Target, string Name, string Scope)>();

        ExcelBorderApply.ApplyCanonical(range, requested, (target, spec) =>
            applied.Add((target, spec.Name, spec.Scope)));

        Assert.Contains(applied, item => ReferenceEquals(item.Target, range) && item.Name == "bottom");
        Assert.Contains(applied, item =>
            ReferenceEquals(item.Target, range) && item.Name == "insideHorizontal" &&
            item.Scope == ExcelBorderContract.ScopeRange);
        Assert.DoesNotContain(applied, item => ReferenceEquals(item.Target, range.Cells.Item(1, 1)));
        Assert.DoesNotContain(applied, item => ReferenceEquals(item.Target, range.Cells.Item(2, 1)));
        Assert.DoesNotContain(applied, item => ReferenceEquals(item.Target, range.Cells.Item(3, 1)));
        Assert.Equal(2, applied.Count);
    }

    [Fact]
    public void Merged_like_2x2_fallback_reads_every_cell()
    {
        var requested = BorderCanonical(255);
        var range = FakeColumn.TwoByTwo(
            outerBottom: MatchBorder(255),
            insideHorizontal: MixedBorder(),
            a1Bottom: MatchBorder(255),
            b1Bottom: MatchBorder(255),
            a2Bottom: MatchBorder(255),
            b2Bottom: MatchBorder(16711680));

        Assert.False(
            ExcelBorderReadback.TryProveEveryCellEdge(range, BottomSpec(255), 2, 2, out _),
            "mixed inside cannot prove a merged 2x2");
        Assert.False(
            ExcelBorderReadback.AppliedBordersMatch(range, requested),
            "fallback must read all four cells, not only the merge origin");
        Assert.True(ExcelBorderReadback.OneBorderMatches(range.Cells.Item(2, 2), BottomSpec(16711680)));
    }

    [Fact]
    public void Union_A1_A3_apply_visits_both_areas_not_first_area_only()
    {
        var requested = BorderCanonical(255);
        var a1 = FakeColumn.Cell(MatchBorder(0), bold: true);
        var a3 = FakeColumn.Cell(MatchBorder(0), bold: true);
        var union = FakeColumn.UnionFirstAreaOnly(a1, a3);
        var applied = new List<object>();

        ExcelBorderApply.ApplyCanonical(union, requested, (target, spec) =>
        {
            Assert.Equal(ExcelBorderContract.ScopeCell, spec.Scope);
            applied.Add(target);
        });

        Assert.Contains(a1, applied);
        Assert.Contains(a3, applied);
        Assert.DoesNotContain(union, applied);
    }

    [Fact]
    public void Union_outline_applies_per_area_not_bounding_box()
    {
        var errors = new List<string>();
        Assert.True(ExcelBorderContract.TryNormalize(
            new JsonObject
            {
                ["outline"] = new JsonObject
                {
                    ["weight"] = "medium",
                    ["lineStyle"] = "continuous",
                    ["color"] = 255,
                },
            },
            0, errors, out var outline));
        Assert.Empty(errors);

        var a1 = FakeColumn.Cell(MatchBorder(0));
        var a3 = FakeColumn.Cell(MatchBorder(0));
        var union = FakeColumn.UnionFirstAreaOnly(a1, a3);
        var applied = new List<object>();
        ExcelBorderApply.ApplyCanonical(union, outline, (target, spec) =>
        {
            Assert.Equal(ExcelBorderContract.ScopeRange, spec.Scope);
            applied.Add(target);
        });

        Assert.Contains(a1, applied);
        Assert.Contains(a3, applied);
        Assert.DoesNotContain(union, applied);
        Assert.Equal(2, applied.Distinct().Count());
    }

    [Fact]
    public void Unmerged_area_MergeCells_false_does_not_scan_cells()
    {
        ExcelBorderApplicability.ResetMergeScanCells();
        var range = FakeColumn.OneByEight(ThinMatch(0), insideHorizontal: NoneBorder());
        range = new FakeRange
        {
            Rows = range.Rows,
            Columns = range.Columns,
            Cells = range.Cells,
            Borders = range.Borders,
            Row = range.Row,
            Column = range.Column,
            MergeCells = false,
        };
        Assert.True(ExcelBorderReadback.AppliedBordersMatch(range, AllCanonical(0)));
        Assert.Equal(0, ExcelBorderApplicability.MergeScanCells);
    }

    [Fact]
    public void Color_zero_outline_fallback_checks_real_perimeter_segments()
    {
        var errors = new List<string>();
        Assert.True(ExcelBorderContract.TryNormalize(
            new JsonObject
            {
                ["outline"] = new JsonObject
                {
                    ["weight"] = "thin",
                    ["lineStyle"] = "continuous",
                    ["color"] = 0,
                },
            },
            0, errors, out var outline));
        Assert.Empty(errors);

        var a1 = FakeColumn.CellBox(ThinMatch(0), left: ThinMatch(0));
        var a2 = FakeColumn.CellBox(ThinMatch(0), left: ThinMatch(16711680));
        var column = new FakeRange
        {
            Rows = new FakeDim(2),
            Columns = new FakeDim(1),
            Cells = new FakeCells(new[,] { { a1 }, { a2 } }),
            Borders = FakeColumn.AllRangeBorders(ThinMatch(0), NoneBorder(), NoneBorder(), ThinMatch(0)),
            MergeCells = false,
        };

        Assert.True(ExcelBorderReadback.ColorZeroCannotProve(0d, new ExcelBorderContract.EdgeSpec(
            "outlineLeft", ExcelBorderContract.ScopeRange, "thin", "continuous", 0), 2));
        Assert.False(
            ExcelBorderReadback.AppliedBordersMatch(column, outline),
            "range Color=0 must not accept a red A2 left perimeter");
    }

    [Fact]
    public void One_row_A4_H4_all_skips_missing_insideHorizontal_and_matches()
    {
        var requested = AllCanonical(0);
        var range = FakeColumn.OneByEight(ThinMatch(0), insideHorizontal: NoneBorder());
        Assert.False(ExcelBorderApplicability.InsideEdgeApplies("insideHorizontal", 1, 8));
        Assert.True(
            ExcelBorderReadback.AppliedBordersMatch(range, requested),
            "1x8 borders.all must not demand xlInsideHorizontal");
        var explain = ExcelBorderReadback.ExplainAppliedBorders(range, requested);
        Assert.Contains(explain, note => note.Contains("insideHorizontal: skipped", StringComparison.Ordinal));
        Assert.Contains(explain, note => note.Contains("insideVertical: ok", StringComparison.Ordinal));
    }

    [Fact]
    public void One_row_all_rejects_wrong_outer_left()
    {
        var requested = AllCanonical(0);
        var range = FakeColumn.OneByEight(ThinMatch(0), insideHorizontal: NoneBorder(), left: NoneBorder());
        Assert.False(ExcelBorderReadback.AppliedBordersMatch(range, requested));
        var explain = string.Join("; ", ExcelBorderReadback.ExplainAppliedBorders(range, requested));
        Assert.Contains("left: expected", explain, StringComparison.Ordinal);
        Assert.Contains("actual none", explain, StringComparison.Ordinal);
    }

    [Fact]
    public void One_row_all_rejects_wrong_color_on_real_segment()
    {
        var requested = AllCanonical(0);
        var range = FakeColumn.OneByEight(ThinMatch(0), insideHorizontal: NoneBorder(), cellLeftOverride: (2, ThinMatch(16711680)));
        Assert.False(
            ExcelBorderReadback.AppliedBordersMatch(range, requested),
            "B4 left is a real 1-row segment; red must fail");
    }

    [Fact]
    public void Fully_merged_one_row_ignores_interior_lefts()
    {
        var requested = AllCanonical(0);
        var range = FakeColumn.OneByEightMerged(ThinMatch(0), interiorLeft: NoneBorder());
        Assert.True(
            ExcelBorderReadback.AppliedBordersMatch(range, requested),
            "lines inside one MergeArea are not visible boundaries");
        Assert.DoesNotContain(
            ExcelBorderReadback.ExplainAppliedBorders(range, requested),
            note => note.StartsWith("insideVertical: expected", StringComparison.Ordinal));
    }

    [Fact]
    public void Merged_subtotal_row_does_not_treat_mixed_inside_as_blanket()
    {
        var requested = AllCanonical(255);
        var range = FakeColumn.ThreeByThreeWithCenterMerge(
            ThinMatch(255),
            gapRight: ThinMatch(16711680));
        Assert.False(
            ExcelBorderReadback.AppliedBordersMatch(range, requested),
            "visible interior right of A2 must fail; merge interior B2:C2 must not blanket-pass");
    }

    [Fact]
    public void Multi_area_all_second_area_gap_fails()
    {
        var requested = AllCanonical(255);
        var a1 = FakeColumn.CellBox(ThinMatch(255));
        var a3 = FakeColumn.CellBox(ThinMatch(255), bottom: ThinMatch(16711680));
        var union = FakeColumn.UnionFirstAreaOnly(a1, a3);
        Assert.False(ExcelBorderReadback.AppliedBordersMatch(union, requested));
    }

    [Fact]
    public void One_row_all_apply_does_not_write_insideHorizontal()
    {
        var requested = AllCanonical(0);
        var range = FakeColumn.OneByEight(ThinMatch(0), insideHorizontal: NoneBorder());
        var applied = new List<string>();
        ExcelBorderApply.ApplyCanonical(range, requested, (_, spec) => applied.Add(spec.Name));
        Assert.DoesNotContain(applied, name => name.Equals("insideHorizontal", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(applied, name => name.Equals("insideVertical", StringComparison.OrdinalIgnoreCase));
        Assert.Contains(applied, name => name.Equals("bottom", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Union_plain_font_verifies_every_area()
    {
        var a1 = FakeColumn.Cell(MatchBorder(0), bold: true);
        var a3 = FakeColumn.Cell(MatchBorder(0), bold: false);
        var union = FakeColumn.UnionFirstAreaOnly(a1, a3);
        union.Font = new FakeFont { Bold = true };

        Assert.True(ExcelFormatReadback.BoldOnArea(a1, true));
        Assert.False(ExcelFormatReadback.BoldOnArea(a3, true));
        Assert.False(
            ExcelFormatReadback.AppliedBoldMatches(union, true),
            "union Font.Bold can hide a second-area miss");
    }

    private static JsonObject AllCanonical(double color)
    {
        var errors = new List<string>();
        Assert.True(ExcelBorderContract.TryNormalize(
            new JsonObject
            {
                ["all"] = new JsonObject
                {
                    ["weight"] = "thin",
                    ["lineStyle"] = "continuous",
                    ["color"] = color,
                },
            },
            0, errors, out var canonical));
        Assert.Empty(errors);
        return canonical;
    }

    private static FakeBorder ThinMatch(double color) => new()
    {
        LineStyle = ExcelBorderContract.XlContinuous,
        Weight = ExcelBorderContract.XlThin,
        Color = color,
    };

    private static JsonObject BorderCanonical(double color)
    {
        var errors = new List<string>();
        Assert.True(ExcelBorderContract.TryNormalize(
            new JsonObject
            {
                ["bottom"] = new JsonObject
                {
                    ["weight"] = "medium",
                    ["lineStyle"] = "continuous",
                    ["color"] = color,
                },
            },
            0, errors, out var canonical));
        Assert.Empty(errors);
        return canonical;
    }

    private static ExcelBorderContract.EdgeSpec BottomSpec(double color) =>
        new("bottom", ExcelBorderContract.ScopeCell, "medium", "continuous", color);

    private static FakeBorder MatchBorder(double color) => new()
    {
        LineStyle = ExcelBorderContract.XlContinuous,
        Weight = ExcelBorderContract.XlMedium,
        Color = color,
    };

    private static FakeBorder NoneBorder() => new()
    {
        LineStyle = ExcelBorderContract.XlLineStyleNone,
        Weight = ExcelBorderContract.XlThin,
        Color = 0,
    };

    private static FakeBorder MixedBorder() => new()
    {
        LineStyle = null,
        Weight = null,
        Color = null,
    };

    public sealed class FakeBorder
    {
        public object? LineStyle { get; init; }
        public object? Weight { get; init; }
        public object? Color { get; init; }
    }

    public sealed class FakeBorders
    {
        private readonly Dictionary<int, FakeBorder> _items = new();
        public void Set(int index, FakeBorder border) => _items[index] = border;
        public FakeBorder Item(int index) => _items[index];
    }

    public sealed class FakeDim
    {
        public FakeDim(int count) => Count = count;
        public int Count { get; }
    }

    public sealed class FakeCells
    {
        private readonly FakeRange[,] _grid;
        public FakeCells(FakeRange[,] grid) => _grid = grid;
        public FakeRange Item(int row, int column) => _grid[row - 1, column - 1];
    }

    public sealed class FakeFont
    {
        public object? Bold { get; init; }
    }

    public sealed class FakeAreas
    {
        private readonly FakeRange[] _areas;
        public FakeAreas(params FakeRange[] areas) => _areas = areas;
        public int Count => _areas.Length;
        public FakeRange Item(int index) => _areas[index - 1];
    }

    public sealed class FakeRange
    {
        public required FakeDim Rows { get; init; }
        public required FakeDim Columns { get; init; }
        public required FakeCells Cells { get; set; }
        public required FakeBorders Borders { get; init; }
        public FakeAreas? Areas { get; init; }
        public FakeFont? Font { get; set; }
        public object? MergeCells { get; init; } = false;
        private readonly FakeRange? _mergeArea;
        public FakeRange? MergeArea
        {
            get
            {
                if (Rows.Count * Columns.Count != 1)
                    throw new InvalidOperationException(
                        "Range.MergeArea is only valid for a single-cell range");
                return _mergeArea;
            }
            init => _mergeArea = value;
        }
        public int Row { get; init; } = 1;
        public int Column { get; init; } = 1;
    }

    public static class FakeColumn
    {
        public static FakeRange A1A3(
            FakeBorder outerBottom, FakeBorder insideHorizontal,
            FakeBorder a1Bottom, FakeBorder a2Bottom, FakeBorder a3Bottom)
        {
            var a1 = Cell(a1Bottom);
            var a2 = Cell(a2Bottom);
            var a3 = Cell(a3Bottom);
            var grid = new[,] { { a1 }, { a2 }, { a3 } };
            var cells = new FakeCells(grid);
            var range = new FakeRange
            {
                Rows = new FakeDim(3),
                Columns = new FakeDim(1),
                Cells = cells,
                Borders = RangeBorders(outerBottom, insideHorizontal),
            };
            return range;
        }

        public static FakeRange TwoByTwo(
            FakeBorder outerBottom, FakeBorder insideHorizontal,
            FakeBorder a1Bottom, FakeBorder b1Bottom,
            FakeBorder a2Bottom, FakeBorder b2Bottom)
        {
            var grid = new[,]
            {
                { Cell(a1Bottom), Cell(b1Bottom) },
                { Cell(a2Bottom), Cell(b2Bottom) },
            };
            return new FakeRange
            {
                Rows = new FakeDim(2),
                Columns = new FakeDim(2),
                Cells = new FakeCells(grid),
                Borders = RangeBorders(outerBottom, insideHorizontal),
            };
        }

        public static FakeRange UnionFirstAreaOnly(FakeRange first, FakeRange second) => new()
        {
            Rows = new FakeDim(1),
            Columns = new FakeDim(1),
            Cells = new FakeCells(new[,] { { first } }),
            Borders = first.Borders,
            Areas = new FakeAreas(first, second),
            Font = first.Font,
        };

        public static FakeRange OneByEight(
            FakeBorder match,
            FakeBorder? insideHorizontal = null,
            FakeBorder? left = null,
            (int Column, FakeBorder Border)? cellLeftOverride = null)
        {
            var cells = new FakeRange[1, 8];
            for (var column = 1; column <= 8; column++)
            {
                var cellLeft = cellLeftOverride is { } ov && ov.Column == column ? ov.Border : match;
                if (left is not null && column == 1) cellLeft = left;
                cells[0, column - 1] = CellBox(match, left: cellLeft, row: 4, column: column);
            }

            return new FakeRange
            {
                Rows = new FakeDim(1),
                Columns = new FakeDim(8),
                Cells = new FakeCells(cells),
                Borders = AllRangeBorders(match, insideHorizontal ?? NoneBorder(), match, left ?? match),
                Row = 4,
                Column = 1,
            };
        }

        public static FakeRange OneByEightMerged(FakeBorder outer, FakeBorder interiorLeft)
        {
            FakeRange range = null!;
            var cells = new FakeRange[1, 8];
            var mergeBox = new FakeRange
            {
                Rows = new FakeDim(1),
                Columns = new FakeDim(8),
                Cells = null!,
                Borders = AllRangeBorders(outer, NoneBorder(), NoneBorder(), outer),
                Row = 4,
                Column = 1,
                MergeCells = true,
            };
            for (var column = 1; column <= 8; column++)
            {
                cells[0, column - 1] = CellBox(
                    outer,
                    left: column == 1 ? outer : interiorLeft,
                    row: 4,
                    column: column,
                    merge: true,
                    mergeArea: mergeBox);
            }

            range = new FakeRange
            {
                Rows = new FakeDim(1),
                Columns = new FakeDim(8),
                Cells = new FakeCells(cells),
                Borders = AllRangeBorders(outer, NoneBorder(), NoneBorder(), outer),
                Row = 4,
                Column = 1,
                MergeCells = true,
            };
            mergeBox.Cells = range.Cells;
            return range;
        }

        public static FakeRange OneByEightContainsPartialMerge(FakeBorder match)
        {
            var mergeBox = new FakeRange
            {
                Rows = new FakeDim(1),
                Columns = new FakeDim(2),
                Cells = null!,
                Borders = AllRangeBorders(match, NoneBorder(), NoneBorder(), match),
                Row = 4,
                Column = 2,
                MergeCells = true,
            };
            var cells = new FakeRange[1, 8];
            for (var column = 1; column <= 8; column++)
            {
                var inMerge = column is 2 or 3;
                cells[0, column - 1] = CellBox(
                    match,
                    row: 4,
                    column: column,
                    merge: inMerge,
                    mergeArea: inMerge ? mergeBox : null);
            }

            return new FakeRange
            {
                Rows = new FakeDim(1),
                Columns = new FakeDim(8),
                Cells = new FakeCells(cells),
                Borders = AllRangeBorders(match, NoneBorder(), match, match),
                Row = 4,
                Column = 1,
                MergeCells = true,
            };
        }

        public static FakeRange ThreeByThreeWithCenterMerge(FakeBorder match, FakeBorder gapRight)
        {
            var mergeBox = new FakeRange
            {
                Rows = new FakeDim(1),
                Columns = new FakeDim(2),
                Cells = null!,
                Borders = AllRangeBorders(match, match, match, match),
                Row = 2,
                Column = 2,
                MergeCells = true,
            };
            var grid = new FakeRange[3, 3];
            for (var r = 1; r <= 3; r++)
            {
                for (var c = 1; c <= 3; c++)
                {
                    var inMerge = r == 2 && c is 2 or 3;
                    var right = r == 2 && c == 1 ? gapRight : match;
                    grid[r - 1, c - 1] = CellBox(
                        match,
                        right: right,
                        row: r,
                        column: c,
                        merge: inMerge,
                        mergeArea: inMerge ? mergeBox : null);
                }
            }

            return new FakeRange
            {
                Rows = new FakeDim(3),
                Columns = new FakeDim(3),
                Cells = new FakeCells(grid),
                Borders = AllRangeBorders(match, MixedBorder(), MixedBorder(), match),
                Row = 1,
                Column = 1,
                MergeCells = null,
            };
        }

        public static FakeRange CellBox(
            FakeBorder match,
            FakeBorder? left = null,
            FakeBorder? right = null,
            FakeBorder? top = null,
            FakeBorder? bottom = null,
            int row = 1,
            int column = 1,
            bool merge = false,
            FakeRange? mergeArea = null)
        {
            var borders = new FakeBorders();
            borders.Set(ExcelBorderContract.XlEdgeLeft, left ?? match);
            borders.Set(ExcelBorderContract.XlEdgeRight, right ?? match);
            borders.Set(ExcelBorderContract.XlEdgeTop, top ?? match);
            borders.Set(ExcelBorderContract.XlEdgeBottom, bottom ?? match);
            var self = new FakeRange
            {
                Rows = new FakeDim(1),
                Columns = new FakeDim(1),
                Cells = null!,
                Borders = borders,
                Row = row,
                Column = column,
                MergeCells = merge,
                MergeArea = mergeArea,
            };
            self.Cells = new FakeCells(new[,] { { self } });
            return self;
        }

        public static FakeRange Cell(FakeBorder bottom, bool? bold = null)
        {
            var borders = new FakeBorders();
            borders.Set(ExcelBorderContract.XlEdgeBottom, bottom);
            var self = new FakeRange
            {
                Rows = new FakeDim(1),
                Columns = new FakeDim(1),
                Cells = null!,
                Borders = borders,
                Font = bold is null ? null : new FakeFont { Bold = bold.Value },
            };
            self.Cells = new FakeCells(new[,] { { self } });
            return self;
        }

        private static FakeBorders RangeBorders(FakeBorder outerBottom, FakeBorder insideHorizontal)
        {
            var borders = new FakeBorders();
            borders.Set(ExcelBorderContract.XlEdgeBottom, outerBottom);
            borders.Set(ExcelBorderContract.XlInsideHorizontal, insideHorizontal);
            return borders;
        }

        public static FakeBorders AllRangeBorders(
            FakeBorder match, FakeBorder insideHorizontal, FakeBorder insideVertical, FakeBorder left)
        {
            var borders = new FakeBorders();
            borders.Set(ExcelBorderContract.XlEdgeLeft, left);
            borders.Set(ExcelBorderContract.XlEdgeRight, match);
            borders.Set(ExcelBorderContract.XlEdgeTop, match);
            borders.Set(ExcelBorderContract.XlEdgeBottom, match);
            borders.Set(ExcelBorderContract.XlInsideHorizontal, insideHorizontal);
            borders.Set(ExcelBorderContract.XlInsideVertical, insideVertical);
            return borders;
        }
    }
}
