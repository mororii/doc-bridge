using System.Globalization;
using System.Text;

namespace DocBridge.Core.Services;

/// <summary>
/// A1 rectangle used for merge-batch overlap and layout targeting without COM.
/// </summary>
internal readonly record struct ExcelA1Box(int Row, int Column, int Rows, int Columns)
{
    internal const int MaxRow = 1_048_576;
    internal const int MaxColumn = 16_384;

    internal int EndRow => Row + Rows - 1;
    internal int EndColumn => Column + Columns - 1;
    internal long CellCount => (long)Rows * Columns;

    internal bool Intersects(ExcelA1Box other) =>
        Row <= other.EndRow && other.Row <= EndRow &&
        Column <= other.EndColumn && other.Column <= EndColumn;

    internal string Address =>
        Rows == 1 && Columns == 1
            ? CellName(Column, Row)
            : $"{CellName(Column, Row)}:{CellName(EndColumn, EndRow)}";

    internal static bool TryParse(string? address, out ExcelA1Box box)
    {
        box = default;
        if (string.IsNullOrWhiteSpace(address)) return false;
        var text = address.Trim();
        if (text.Contains('!', StringComparison.Ordinal) ||
            text.Contains(',', StringComparison.Ordinal) ||
            text.Contains(';', StringComparison.Ordinal) ||
            text.Contains(' ', StringComparison.Ordinal))
            return false;

        var parts = text.Split(':');
        if (parts.Length is < 1 or > 2) return false;
        if (!TryParseCell(parts[0], out var startRow, out var startCol)) return false;
        var endRow = startRow;
        var endCol = startCol;
        if (parts.Length == 2 && !TryParseCell(parts[1], out endRow, out endCol)) return false;
        if (endRow < startRow || endCol < startCol) return false;
        try
        {
            checked
            {
                box = new ExcelA1Box(startRow, startCol, endRow - startRow + 1, endCol - startCol + 1);
            }
        }
        catch (OverflowException)
        {
            return false;
        }

        return box.Rows >= 1 && box.Columns >= 1;
    }

    internal static bool TryParseCell(string token, out int row, out int column)
    {
        row = 0;
        column = 0;
        if (string.IsNullOrWhiteSpace(token)) return false;
        var letters = new StringBuilder();
        var digits = new StringBuilder();
        var seenDigit = false;
        foreach (var ch in token.Trim())
        {
            if (ch == '$') continue;
            if (!seenDigit && char.IsAsciiLetter(ch))
            {
                letters.Append(ch);
                continue;
            }

            if (char.IsAsciiDigit(ch))
            {
                seenDigit = true;
                digits.Append(ch);
                continue;
            }

            return false;
        }

        if (letters.Length is < 1 or > 3 || digits.Length == 0) return false;
        if (!int.TryParse(digits.ToString(), NumberStyles.None, CultureInfo.InvariantCulture, out row))
            return false;
        if (row is < 1 or > MaxRow) return false;
        column = ColumnIndex(letters.ToString());
        if (column is < 1 or > MaxColumn) return false;
        return string.Equals(ColumnName(column), letters.ToString(), StringComparison.OrdinalIgnoreCase);
    }

    internal static string ColumnName(int column)
    {
        var name = "";
        var current = column;
        while (current > 0)
        {
            var m = (current - 1) % 26;
            name = (char)('A' + m) + name;
            current = (current - 1) / 26;
        }
        return name;
    }

    internal static int ColumnIndex(string name)
    {
        var column = 0;
        foreach (var ch in name.Trim().ToUpperInvariant())
        {
            if (ch is < 'A' or > 'Z') return 0;
            try { column = checked(column * 26 + ch - 'A' + 1); }
            catch (OverflowException) { return 0; }
        }
        return column;
    }

    internal static string CellName(int column, int row) => $"{ColumnName(column)}{row}";
}
