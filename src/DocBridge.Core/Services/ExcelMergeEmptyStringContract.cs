namespace DocBridge.Core.Services;

/// <summary>
/// JSON "" may normalize to a true blank on write. Merge still refuses
/// formulas and nonempty non-upper-left cells. Existing constant empty
/// strings (shared-string t=s) may be ClearContents only after a proven
/// read. A failed read, clear, or post-clear readback aborts before Merge.
/// Empty constants as the 0820 warning cause remain a hypothesis.
/// </summary>
public static class ExcelMergeEmptyStringContract
{
    public const string UnprovenCellError = "[EXCEL_MERGE_UNPROVEN_CELL]";
    public const string WouldDeleteContentError = "[EXCEL_MERGE_WOULD_DELETE_CONTENT]";

    public readonly record struct MergeCellSnapshot(object? Formula, object? Value2, bool ReadFailed);

    public interface IMergeRangeSurface
    {
        int CellCount { get; }
        bool Merged { get; }
        MergeCellSnapshot Read(int oneBasedIndex);
        void ClearContents(int oneBasedIndex);
        void MergeAcross();
    }

    public static bool LooksLikeFormula(object? formulaOrValue) =>
        formulaOrValue is string text && text.StartsWith('=');

    public static bool IsNonEmptyForMergeRefuse(object? formulaOrValue) =>
        formulaOrValue is not null && (formulaOrValue is not string text || text.Length != 0);

    /// <summary>
    /// Native empty-string persistence is not a blank. Only Value2 null/DBNull
    /// is a true blank. Do not use set_values TypedEqual here: that treats
    /// leftover "" as equivalent to a cleared cell.
    /// </summary>
    public static bool IsTrueBlank(object? value2) => value2 is null or DBNull;

    public static bool IsTrueBlank(object? formulaOrValue, object? value2) =>
        !LooksLikeFormula(formulaOrValue) &&
        !IsNonEmptyForMergeRefuse(formulaOrValue) &&
        IsTrueBlank(value2);

    public static bool IsProvenConstantEmpty(object? formulaOrValue, object? value2)
    {
        if (LooksLikeFormula(formulaOrValue) || IsNonEmptyForMergeRefuse(formulaOrValue))
            return false;
        return value2 is string empty && empty.Length == 0;
    }

    /// <summary>
    /// Post-clear proof. Formula may read back as "". Value2 "" is unproven.
    /// </summary>
    public static bool IsClearedBlank(object? formulaOrValue, object? value2)
    {
        _ = formulaOrValue;
        return IsTrueBlank(value2);
    }

    public static void PreclearThenMerge(IMergeRangeSurface range)
    {
        if (range.CellCount < 2)
            throw new InvalidOperationException("merge_cells requires a range containing at least two cells");

        for (var index = 2; index <= range.CellCount; index++)
        {
            MergeCellSnapshot snap;
            try
            {
                snap = range.Read(index);
            }
            catch (Exception ex)
            {
                throw Unproven($"read failed at non-upper-left cell {index}: {ex.Message}");
            }

            if (snap.ReadFailed)
                throw Unproven($"read failed at non-upper-left cell {index}");
            if (LooksLikeFormula(snap.Formula) || IsNonEmptyForMergeRefuse(snap.Formula))
                throw new InvalidOperationException(
                    $"{WouldDeleteContentError} merge was blocked because a non-upper-left cell contains a value or formula");
            if (IsTrueBlank(snap.Formula, snap.Value2))
                continue;
            if (!IsProvenConstantEmpty(snap.Formula, snap.Value2))
                throw Unproven($"non-upper-left cell {index} is not a proven blank or constant empty");

            try
            {
                range.ClearContents(index);
            }
            catch (Exception ex)
            {
                throw Unproven($"ClearContents failed at non-upper-left cell {index}: {ex.Message}");
            }

            MergeCellSnapshot after;
            try
            {
                after = range.Read(index);
            }
            catch (Exception ex)
            {
                throw Unproven($"readback failed after ClearContents at cell {index}: {ex.Message}");
            }

            if (after.ReadFailed)
                throw Unproven($"readback failed after ClearContents at cell {index}");
            if (LooksLikeFormula(after.Formula) || !IsTrueBlank(after.Value2))
                throw Unproven($"cell {index} was not a true blank after ClearContents");
        }

        range.MergeAcross();
    }

    private static InvalidOperationException Unproven(string detail) =>
        new($"{UnprovenCellError} {detail}; merge was not called");
}
