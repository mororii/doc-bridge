using System.Globalization;

namespace DocBridge.Core.Services;

/// <summary>
/// Text runs temporarily set NumberFormat=@. A mixed run returns null from
/// Range.NumberFormat; treating that as "nothing to restore" leaves every
/// cell as @. Capture the uniform string, or per-cell originals when mixed.
/// Refuse a failed capture before write. Restore even if Value2 throws.
/// </summary>
public static class ExcelNumberFormatPreserve
{
    public const string General = "General";
    public const string FixedTwoDecimals = "0.00";
    public const string DateYmd = "yyyy-mm-dd";
    public const string CustomQuotedYear = "\"yyyy\"";

    public static readonly string[] MixedRunFixtureFormats =
    [
        General,
        FixedTwoDecimals,
        DateYmd,
        CustomQuotedYear,
    ];

    public sealed record Capture(bool Mixed, IReadOnlyList<string> Formats)
    {
        public bool RestorePerCell =>
            Mixed || Formats.Distinct(StringComparer.Ordinal).Count() > 1;
    }

    public static bool NeedsPerCellCapture(object? raw) =>
        raw is null or DBNull or object[,] ||
        (raw is string text && string.IsNullOrEmpty(text));

    public static bool TryNormalizeCapture(
        object? raw,
        IReadOnlyList<string>? perCell,
        int length,
        out Capture capture,
        out string? error)
    {
        capture = default!;
        error = null;
        if (length < 1)
        {
            error = "text run length must be at least 1";
            return false;
        }

        if (TryFromUniformOrArray(raw, length, out var fromRaw, out var rawError))
        {
            capture = fromRaw;
            return true;
        }

        if (perCell is null || perCell.Count != length)
        {
            error = "mixed NumberFormat must be captured per-cell before write; refusing text write"
                    + (rawError is null ? "" : $" ({rawError})");
            return false;
        }

        for (var i = 0; i < perCell.Count; i++)
        {
            if (string.IsNullOrEmpty(perCell[i]))
            {
                error = $"per-cell NumberFormat capture failed at index {i}; refusing text write";
                return false;
            }
        }

        capture = new Capture(Mixed: true, Formats: perCell.ToArray());
        return true;
    }

    public static bool FormatsRetained(IReadOnlyList<string> before, IReadOnlyList<string> after)
    {
        if (before.Count != after.Count)
            return false;
        for (var i = 0; i < before.Count; i++)
        {
            if (!ExcelNumberFormatContract.ReadbackMatches(before[i], after[i]))
                return false;
        }

        return true;
    }

    public static bool TryFromUniformOrArray(
        object? raw, int length, out Capture capture, out string? error)
    {
        capture = default!;
        error = null;
        if (raw is null or DBNull)
        {
            error = "mixed NumberFormat returned null";
            return false;
        }

        if (raw is string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                error = "NumberFormat was empty";
                return false;
            }

            capture = new Capture(Mixed: false, Formats: Repeat(text, length));
            return true;
        }

        if (raw is object[,] arr)
        {
            var list = FlattenArray(arr);
            if (list.Count != length)
            {
                error = $"NumberFormat array length {list.Count} != run length {length}";
                return false;
            }

            if (list.Any(string.IsNullOrEmpty))
            {
                error = "mixed NumberFormat array contains an empty cell";
                return false;
            }

            capture = new Capture(Mixed: list.Distinct(StringComparer.Ordinal).Count() > 1, Formats: list);
            return true;
        }

        var scalar = Convert.ToString(raw, CultureInfo.InvariantCulture);
        if (string.IsNullOrEmpty(scalar))
        {
            error = "NumberFormat could not be read";
            return false;
        }

        capture = new Capture(Mixed: false, Formats: Repeat(scalar, length));
        return true;
    }

    private static IReadOnlyList<string> FlattenArray(object[,] arr)
    {
        var list = new List<string>();
        var r1 = arr.GetLowerBound(0);
        var r2 = arr.GetUpperBound(0);
        var c1 = arr.GetLowerBound(1);
        var c2 = arr.GetUpperBound(1);
        for (var r = r1; r <= r2; r++)
        {
            for (var c = c1; c <= c2; c++)
                list.Add(Convert.ToString(arr[r, c], CultureInfo.InvariantCulture) ?? "");
        }

        return list;
    }

    private static IReadOnlyList<string> Repeat(string value, int length)
    {
        var formats = new string[length];
        Array.Fill(formats, value);
        return formats;
    }
}
