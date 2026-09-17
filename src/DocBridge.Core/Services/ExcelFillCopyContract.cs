using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace DocBridge.Core.Services;

/// <summary>
/// auto_fill and copy_range readback: filled-band semantics and relative A1 shifts.
/// Comparing only the source block or unshifted formula text is not verification.
/// </summary>
public static class ExcelFillCopyContract
{
    private static readonly Regex A1 = new(
        @"(?<![A-Za-z0-9_])(\$?)([A-Za-z]+)(\$?)(\d+)\b",
        RegexOptions.Compiled);

    public static bool LooksLikeFormula(string? text) =>
        !string.IsNullOrWhiteSpace(text) && text.StartsWith('=');

    public static IReadOnlyList<(int Row, int Column)> FilledBand(int sourceRows, int sourceCols, int destRows, int destCols)
    {
        var cells = new List<(int, int)>();
        for (var r = 0; r < destRows; r++)
        for (var c = 0; c < destCols; c++)
        {
            if (r < sourceRows && c < sourceCols)
                continue;
            cells.Add((r, c));
        }

        return cells;
    }

    public static string ShiftA1Formula(string formula, int rowDelta, int columnDelta)
    {
        if (!LooksLikeFormula(formula) || (rowDelta == 0 && columnDelta == 0))
            return formula;
        return A1.Replace(formula, match =>
        {
            var colAbs = match.Groups[1].Value == "$";
            var rowAbs = match.Groups[3].Value == "$";
            var column = ColIndex(match.Groups[2].Value);
            var row = int.Parse(match.Groups[4].Value, CultureInfo.InvariantCulture);
            if (!colAbs)
                column = Math.Max(1, column + columnDelta);
            if (!rowAbs)
                row = Math.Max(1, row + rowDelta);
            return (colAbs ? "$" : "") + ColName(column) + (rowAbs ? "$" : "") + row.ToString(CultureInfo.InvariantCulture);
        });
    }

    public static bool FormulasMatchWithRelativeShift(
        IReadOnlyList<IReadOnlyList<string>> source,
        IReadOnlyList<IReadOnlyList<string>> dest,
        int rowDelta,
        int columnDelta)
    {
        if (source.Count == 0 || dest.Count < source.Count)
            return false;
        var width = source[0].Count;
        if (dest[0].Count < width)
            return false;
        for (var r = 0; r < source.Count; r++)
        for (var c = 0; c < width; c++)
        {
            var want = ShiftA1Formula(source[r][c], rowDelta, columnDelta);
            var got = dest[r][c];
            if (LooksLikeFormula(source[r][c]))
            {
                if (!LooksLikeFormula(got) || !string.Equals(want, got, StringComparison.OrdinalIgnoreCase))
                    return false;
            }
        }

        return true;
    }

    public static bool FilledFormulasContinue(
        IReadOnlyList<IReadOnlyList<string>> source,
        IReadOnlyList<IReadOnlyList<string>> dest)
    {
        if (source.Count == 0 || dest.Count == 0)
            return false;
        var sourceRows = source.Count;
        var sourceCols = source[0].Count;
        var destRows = dest.Count;
        var destCols = dest[0].Count;
        foreach (var (r, c) in FilledBand(sourceRows, sourceCols, destRows, destCols))
        {
            var templateRow = Math.Min(r, sourceRows - 1);
            var templateCol = Math.Min(c, sourceCols - 1);
            var template = source[templateRow][templateCol];
            if (!LooksLikeFormula(template))
                continue;
            var want = ShiftA1Formula(template, r - templateRow, c - templateCol);
            if (!LooksLikeFormula(dest[r][c]) ||
                !string.Equals(want, dest[r][c], StringComparison.OrdinalIgnoreCase))
                return false;
        }

        return true;
    }

    public static bool FilledSequenceContinues(
        IReadOnlyList<IReadOnlyList<string>> source,
        IReadOnlyList<IReadOnlyList<string>> dest)
    {
        if (source.Count == 0 || dest.Count < 2)
            return source.Count > 0 && dest.Count >= source.Count;
        var sourceRows = source.Count;
        var destRows = dest.Count;
        var cols = Math.Min(source[0].Count, dest[0].Count);
        for (var c = 0; c < cols; c++)
        {
            if (!TryNumber(source[0][c], out var first))
                continue;
            var step = sourceRows >= 2 && TryNumber(source[1][c], out var second)
                ? second - first
                : 0d;
            if (sourceRows == 1)
                step = 0;
            for (var r = 0; r < destRows; r++)
            {
                if (!TryNumber(dest[r][c], out var actual))
                    return false;
                var expected = first + step * r;
                if (Math.Abs(actual - expected) > 1e-9)
                    return false;
            }
        }

        return true;
    }

    public static bool FilledAreaHasContent(
        IReadOnlyList<IReadOnlyList<string>> source,
        IReadOnlyList<IReadOnlyList<string>> dest)
    {
        if (source.Count == 0 || dest.Count == 0)
            return false;
        var sourceRows = source.Count;
        var sourceCols = source[0].Count;
        foreach (var (r, c) in FilledBand(sourceRows, sourceCols, dest.Count, dest[0].Count))
        {
            if (string.IsNullOrEmpty(dest[r][c]))
                return false;
        }

        return FilledBand(sourceRows, sourceCols, dest.Count, dest[0].Count).Count > 0;
    }

    private static bool TryNumber(string text, out double value) =>
        double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out value);

    private static int ColIndex(string name)
    {
        var c = 0;
        foreach (var ch in name.ToUpperInvariant())
            c = c * 26 + (ch - 'A' + 1);
        return c;
    }

    private static string ColName(int column)
    {
        var name = new StringBuilder();
        var c = column;
        while (c > 0)
        {
            var m = (c - 1) % 26;
            name.Insert(0, (char)('A' + m));
            c = (c - 1) / 26;
        }

        return name.ToString();
    }
}
