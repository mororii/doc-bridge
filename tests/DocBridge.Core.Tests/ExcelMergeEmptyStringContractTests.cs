using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelMergeEmptyStringContractTests
{
    [Fact]
    public void Preview_blank_empty_string_is_not_merge_refuse_content()
    {
        Assert.False(ExcelMergeEmptyStringContract.IsNonEmptyForMergeRefuse(""));
        Assert.False(ExcelMergeEmptyStringContract.IsNonEmptyForMergeRefuse(null));
        Assert.True(ExcelMergeEmptyStringContract.IsNonEmptyForMergeRefuse("서명"));
        Assert.True(ExcelMergeEmptyStringContract.IsNonEmptyForMergeRefuse("=\"\""));
    }

    [Fact]
    public void Only_proven_constant_empty_strings_are_precleared()
    {
        Assert.True(ExcelMergeEmptyStringContract.IsProvenConstantEmpty("", ""));
        Assert.True(ExcelMergeEmptyStringContract.IsProvenConstantEmpty(null, ""));
        Assert.False(ExcelMergeEmptyStringContract.IsProvenConstantEmpty("", null));
        Assert.False(ExcelMergeEmptyStringContract.IsProvenConstantEmpty("=\"\"", ""));
        Assert.False(ExcelMergeEmptyStringContract.IsProvenConstantEmpty("=A1", ""));
        Assert.False(ExcelMergeEmptyStringContract.IsProvenConstantEmpty("keep", ""));
        Assert.True(ExcelMergeEmptyStringContract.LooksLikeFormula("=\"\""));
        Assert.False(ExcelMergeEmptyStringContract.LooksLikeFormula(""));
        Assert.True(ExcelMergeEmptyStringContract.IsTrueBlank(null));
        Assert.True(ExcelMergeEmptyStringContract.IsTrueBlank(DBNull.Value));
        Assert.False(ExcelMergeEmptyStringContract.IsTrueBlank(""));
        Assert.True(ExcelMergeEmptyStringContract.IsTrueBlank(null, null));
        Assert.False(ExcelMergeEmptyStringContract.IsTrueBlank(null, ""));
        Assert.True(ExcelMergeEmptyStringContract.IsClearedBlank("", null));
        Assert.True(ExcelMergeEmptyStringContract.IsClearedBlank(null, DBNull.Value));
        Assert.False(ExcelMergeEmptyStringContract.IsClearedBlank("", ""));
        Assert.False(ExcelMergeEmptyStringContract.IsClearedBlank(null, ""));
        Assert.False(ExcelMergeEmptyStringContract.IsClearedBlank("unused", ""));
        Assert.True(ExcelValueWriteContract.TypedEqual(
            ExcelValueWriteContract.Classify(JsonValue.Create("")), ""));
    }

    [Fact]
    public void Empty_json_write_grid_is_blank_and_not_an_at_text_run()
    {
        var writes = new ExcelValueWriteContract.CellWrite[3, 2];
        writes[0, 0] = ExcelValueWriteContract.Classify(JsonValue.Create("서명"));
        writes[0, 1] = ExcelValueWriteContract.Classify(JsonValue.Create(""));
        writes[1, 0] = ExcelValueWriteContract.Classify(JsonValue.Create(""));
        writes[1, 1] = ExcelValueWriteContract.Classify(JsonValue.Create(""));
        writes[2, 0] = ExcelValueWriteContract.Classify(JsonValue.Create(""));
        writes[2, 1] = ExcelValueWriteContract.Classify(JsonValue.Create(""));
        Assert.Equal(ExcelValueWriteContract.JsonKind.String, writes[0, 0].Kind);
        Assert.All(new[] { writes[0, 1], writes[1, 0], writes[1, 1], writes[2, 0], writes[2, 1] }, cell =>
        {
            Assert.Equal(ExcelValueWriteContract.JsonKind.EmptyString, cell.Kind);
            Assert.Null(cell.ComValue);
            Assert.Null(cell.Text);
            Assert.False(ExcelValueWriteContract.RequiresTextNumberFormat(cell.Kind));
        });
        Assert.Equal(new[] { (0, 0, 1) }, ExcelValueWriteContract.ContiguousTextRuns(writes, 3, 2));
    }

    [Fact]
    public void Stateful_clear_then_merge_blocks_on_failed_read_or_clear_and_preserves_anchor_formula_nonempty()
    {
        var cleared = FakeMergeRange.SignatureBlock(anchor: "서명", nonUpperLeft: "");
        ExcelMergeEmptyStringContract.PreclearThenMerge(cleared);
        Assert.Equal(new[] { "read:2", "clear:2", "read:2", "read:3", "clear:3", "read:3",
            "read:4", "clear:4", "read:4", "read:5", "clear:5", "read:5", "read:6", "clear:6", "read:6", "merge" }, cleared.Log);
        Assert.True(cleared.Merged);
        Assert.Equal("서명", cleared.Cells[0].Value2);
        Assert.False(cleared.Cells[0].Cleared);

        var failedRead = FakeMergeRange.TwoCells(anchor: "서명", otherFormula: "", otherValue: "");
        failedRead.Cells[1].ThrowOnRead = true;
        var readEx = Assert.Throws<InvalidOperationException>(() => ExcelMergeEmptyStringContract.PreclearThenMerge(failedRead));
        Assert.Contains(ExcelMergeEmptyStringContract.UnprovenCellError, readEx.Message);
        Assert.False(failedRead.Merged);
        Assert.DoesNotContain("merge", failedRead.Log);
        Assert.DoesNotContain("clear:2", failedRead.Log);
        Assert.Equal("서명", failedRead.Cells[0].Value2);
        Assert.Equal("", failedRead.Cells[1].Value2);

        var failedClear = FakeMergeRange.TwoCells(anchor: "서명", otherFormula: "", otherValue: "");
        failedClear.Cells[1].ThrowOnClear = true;
        var clearEx = Assert.Throws<InvalidOperationException>(() => ExcelMergeEmptyStringContract.PreclearThenMerge(failedClear));
        Assert.Contains(ExcelMergeEmptyStringContract.UnprovenCellError, clearEx.Message);
        Assert.False(failedClear.Merged);
        Assert.Equal("", failedClear.Cells[1].Value2);

        var formula = FakeMergeRange.TwoCells(anchor: "서명", otherFormula: "=A1", otherValue: "");
        var formulaEx = Assert.Throws<InvalidOperationException>(() => ExcelMergeEmptyStringContract.PreclearThenMerge(formula));
        Assert.Contains(ExcelMergeEmptyStringContract.WouldDeleteContentError, formulaEx.Message);
        Assert.False(formula.Merged);
        Assert.Equal("=A1", formula.Cells[1].Formula);
        Assert.False(formula.Cells[1].Cleared);

        var nonempty = FakeMergeRange.TwoCells(anchor: "서명", otherFormula: "keep", otherValue: "keep");
        var nonemptyEx = Assert.Throws<InvalidOperationException>(() => ExcelMergeEmptyStringContract.PreclearThenMerge(nonempty));
        Assert.Contains(ExcelMergeEmptyStringContract.WouldDeleteContentError, nonemptyEx.Message);
        Assert.False(nonempty.Merged);
        Assert.Equal("keep", nonempty.Cells[1].Value2);
        Assert.False(nonempty.Cells[1].Cleared);
    }

    [Fact]
    public void Preclear_then_merge_clears_constant_empty_non_ul_and_keeps_anchor()
    {
        var range = FakeMergeRange.SignatureBlock(anchor: "서명", nonUpperLeft: "");
        ExcelMergeEmptyStringContract.PreclearThenMerge(range);
        Assert.True(range.Merged);
        Assert.Equal("서명", range.Cells[0].Formula);
        Assert.Equal("서명", range.Cells[0].Value2);
        Assert.False(range.Cells[0].Cleared);
        Assert.All(range.Cells.Skip(1), cell =>
        {
            Assert.True(cell.Cleared);
            Assert.True(ExcelMergeEmptyStringContract.IsClearedBlank(cell.Formula, cell.Value2));
            Assert.Null(cell.Value2);
        });
        Assert.Equal(new[] { "read:2", "clear:2", "read:2", "read:3", "clear:3", "read:3",
            "read:4", "clear:4", "read:4", "read:5", "clear:5", "read:5", "read:6", "clear:6", "read:6", "merge" }, range.Log);
    }

    [Fact]
    public void Preclear_then_merge_skips_true_blanks_and_still_merges()
    {
        var range = FakeMergeRange.SignatureBlock(anchor: "서명", nonUpperLeft: null);
        ExcelMergeEmptyStringContract.PreclearThenMerge(range);
        Assert.True(range.Merged);
        Assert.All(range.Cells.Skip(1), cell => Assert.False(cell.Cleared));
        Assert.DoesNotContain(range.Log, entry => entry.StartsWith("clear:", StringComparison.Ordinal));
        Assert.Equal("merge", range.Log[^1]);
    }

    [Fact]
    public void Failed_value_read_prevents_merge_and_leaves_cells()
    {
        var range = FakeMergeRange.SignatureBlock(anchor: "서명", nonUpperLeft: "");
        range.Cells[1].ReadFailed = true;
        var ex = Assert.Throws<InvalidOperationException>(() => ExcelMergeEmptyStringContract.PreclearThenMerge(range));
        Assert.Contains(ExcelMergeEmptyStringContract.UnprovenCellError, ex.Message);
        Assert.False(range.Merged);
        Assert.DoesNotContain("merge", range.Log);
        Assert.All(range.Cells, cell => Assert.False(cell.Cleared));
        Assert.Equal("", range.Cells[2].Value2);
    }

    [Fact]
    public void Failed_clear_prevents_merge()
    {
        var range = FakeMergeRange.SignatureBlock(anchor: "서명", nonUpperLeft: "");
        range.Cells[1].ThrowOnClear = true;
        var ex = Assert.Throws<InvalidOperationException>(() => ExcelMergeEmptyStringContract.PreclearThenMerge(range));
        Assert.Contains(ExcelMergeEmptyStringContract.UnprovenCellError, ex.Message);
        Assert.False(range.Merged);
        Assert.False(range.Cells[1].Cleared);
        Assert.Equal("", range.Cells[1].Value2);
    }

    [Fact]
    public void No_op_clear_that_leaves_empty_string_value2_prevents_merge()
    {
        var range = FakeMergeRange.TwoCells(anchor: "서명", otherFormula: "", otherValue: "");
        range.Cells[1].NoOpClear = true;
        Assert.False(ExcelMergeEmptyStringContract.IsTrueBlank(range.Cells[1].Value2));
        Assert.False(ExcelMergeEmptyStringContract.IsClearedBlank(
            range.Cells[1].Formula, range.Cells[1].Value2));
        var ex = Assert.Throws<InvalidOperationException>(() => ExcelMergeEmptyStringContract.PreclearThenMerge(range));
        Assert.Contains(ExcelMergeEmptyStringContract.UnprovenCellError, ex.Message);
        Assert.True(range.Cells[1].Cleared);
        Assert.Equal("", range.Cells[1].Formula);
        Assert.Equal("", range.Cells[1].Value2);
        Assert.False(ExcelMergeEmptyStringContract.IsTrueBlank(range.Cells[1].Value2));
        Assert.False(range.Merged);
        Assert.DoesNotContain("merge", range.Log);
        Assert.Equal(new[] { "read:2", "clear:2", "read:2" }, range.Log);
    }

    [Fact]
    public void Failed_readback_after_clear_prevents_merge()
    {
        var range = FakeMergeRange.SignatureBlock(anchor: "서명", nonUpperLeft: "");
        range.Cells[1].ReadFailedAfterClear = true;
        var ex = Assert.Throws<InvalidOperationException>(() => ExcelMergeEmptyStringContract.PreclearThenMerge(range));
        Assert.Contains(ExcelMergeEmptyStringContract.UnprovenCellError, ex.Message);
        Assert.True(range.Cells[1].Cleared);
        Assert.False(range.Merged);
        Assert.DoesNotContain("merge", range.Log);
    }

    [Fact]
    public void Formula_and_nonempty_non_ul_are_preserved_and_block_merge()
    {
        var formula = FakeMergeRange.TwoCells(anchor: "서명", otherFormula: "=\"\"", otherValue: "");
        var formulaEx = Assert.Throws<InvalidOperationException>(() => ExcelMergeEmptyStringContract.PreclearThenMerge(formula));
        Assert.Contains(ExcelMergeEmptyStringContract.WouldDeleteContentError, formulaEx.Message);
        Assert.False(formula.Merged);
        Assert.Equal("=\"\"", formula.Cells[1].Formula);
        Assert.False(formula.Cells[1].Cleared);

        var nonempty = FakeMergeRange.TwoCells(anchor: "서명", otherFormula: "keep", otherValue: "keep");
        var nonemptyEx = Assert.Throws<InvalidOperationException>(() => ExcelMergeEmptyStringContract.PreclearThenMerge(nonempty));
        Assert.Contains(ExcelMergeEmptyStringContract.WouldDeleteContentError, nonemptyEx.Message);
        Assert.False(nonempty.Merged);
        Assert.Equal("keep", nonempty.Cells[1].Value2);
        Assert.False(nonempty.Cells[1].Cleared);
    }

    [Fact]
    public void Evidence_fixture_is_a_new_file_public_op_repro_not_the_hanging_0820_book()
    {
        var path = Path.Combine(AppContext.BaseDirectory, "evidence", "zero-length-string-merge.json");
        Assert.True(File.Exists(path));
        var root = JsonNode.Parse(File.ReadAllText(path)) as JsonObject;
        Assert.NotNull(root);
        Assert.Equal("zero-length-string-merge", Json.GetString(root, "id"));
        Assert.Contains("0820", Json.GetString(root, "doNotReplayOn") ?? "", StringComparison.Ordinal);
        Assert.Equal("create_workbook", Json.GetString(Json.GetObj(root, "create"), "op"));
        var write = Json.GetObj(root, "write");
        Assert.Equal("set_values", Json.GetString(write, "op"));
        var values = Json.GetArr(write, "values");
        Assert.NotNull(values);
        Assert.Equal(3, values.Count);
        Assert.Equal("", (values[0] as JsonArray)?[1]?.GetValue<string>());
        var merge = Json.GetArr(root, "merge");
        Assert.NotNull(merge);
        Assert.All(merge, node =>
        {
            var op = Assert.IsType<JsonObject>(node);
            Assert.Equal("merge_cells", Json.GetString(op, "op"));
        });
        Assert.Equal("set_formulas", Json.GetString(Json.GetObj(root, "refuseFormula"), "op"));
        Assert.Contains("Cancel", Json.GetString(root, "native") ?? "", StringComparison.Ordinal);
    }

    private sealed class FakeMergeCell
    {
        public object? Formula;
        public object? Value2;
        public bool ReadFailed;
        public bool ReadFailedAfterClear;
        public bool ThrowOnRead;
        public bool ThrowOnClear;
        public bool NoOpClear;
        public bool Cleared;
    }

    private sealed class FakeMergeRange : ExcelMergeEmptyStringContract.IMergeRangeSurface
    {
        public List<FakeMergeCell> Cells { get; }
        public List<string> Log { get; } = [];
        public bool Merged { get; private set; }
        public int CellCount => Cells.Count;

        public FakeMergeRange(IEnumerable<FakeMergeCell> cells) => Cells = cells.ToList();

        public static FakeMergeRange SignatureBlock(object? anchor, object? nonUpperLeft)
        {
            var cells = new List<FakeMergeCell> { new() { Formula = anchor, Value2 = anchor } };
            for (var i = 0; i < 5; i++)
                cells.Add(new FakeMergeCell { Formula = nonUpperLeft, Value2 = nonUpperLeft });
            return new FakeMergeRange(cells);
        }

        public static FakeMergeRange TwoCells(object? anchor, object? otherFormula, object? otherValue) =>
            new([
                new FakeMergeCell { Formula = anchor, Value2 = anchor },
                new FakeMergeCell { Formula = otherFormula, Value2 = otherValue },
            ]);

        public ExcelMergeEmptyStringContract.MergeCellSnapshot Read(int oneBasedIndex)
        {
            Log.Add($"read:{oneBasedIndex}");
            var cell = Cells[oneBasedIndex - 1];
            if (cell.ThrowOnRead)
                throw new InvalidOperationException("Value2 read failed");
            return new ExcelMergeEmptyStringContract.MergeCellSnapshot(cell.Formula, cell.Value2, cell.ReadFailed);
        }

        public void ClearContents(int oneBasedIndex)
        {
            Log.Add($"clear:{oneBasedIndex}");
            var cell = Cells[oneBasedIndex - 1];
            if (cell.ThrowOnClear)
                throw new InvalidOperationException("clear failed");
            cell.Cleared = true;
            if (cell.NoOpClear)
            {
                cell.Formula = "";
                return;
            }
            cell.Formula = "";
            cell.Value2 = null;
            if (cell.ReadFailedAfterClear)
                cell.ReadFailed = true;
        }

        public void MergeAcross()
        {
            Log.Add("merge");
            Merged = true;
        }
    }
}
