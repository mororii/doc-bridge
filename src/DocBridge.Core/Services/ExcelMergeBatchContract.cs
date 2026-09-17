namespace DocBridge.Core.Services;

/// <summary>
/// Homogeneous nonoverlapping merge/unmerge batches. merge_cells and
/// unmerge_cells cannot share a batch or mix with values/format. A 31-month
/// header pair is one merge-only batch of adjacent two-column ranges.
/// </summary>
public static class ExcelMergeBatchContract
{
    public const int MonthHeaderPairCount = 31;
    public const int DefaultMonthHeaderRow = 5;
    public const int DefaultMonthHeaderStartColumn = 7;

    public static IReadOnlyList<string> MonthHeaderPairRanges(
        int startColumn = DefaultMonthHeaderStartColumn,
        int row = DefaultMonthHeaderRow)
    {
        var ranges = new List<string>(MonthHeaderPairCount);
        var column = startColumn;
        for (var index = 0; index < MonthHeaderPairCount; index++)
        {
            ranges.Add($"{ExcelA1Box.ColumnName(column)}{row}:{ExcelA1Box.ColumnName(column + 1)}{row}");
            column += 2;
        }

        return ranges;
    }
}
