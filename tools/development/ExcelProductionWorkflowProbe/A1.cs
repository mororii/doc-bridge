namespace DocBridge.Development.ExcelProductionWorkflowProbe;

internal static class A1
{
    public static string Col(int index1)
    {
        if (index1 < 1 || index1 > 16_384) throw new ArgumentOutOfRangeException(nameof(index1));
        var n = index1;
        var chars = new Stack<char>();
        while (n > 0)
        {
            n--;
            chars.Push((char)('A' + (n % 26)));
            n /= 26;
        }
        return new string(chars.ToArray());
    }

    public static int ColIndex(string col)
    {
        var n = 0;
        foreach (var ch in col.ToUpperInvariant())
        {
            if (ch is < 'A' or > 'Z') throw new ArgumentException($"not a column letter: {col}");
            n = n * 26 + (ch - 'A' + 1);
        }
        return n;
    }

    public static string Cell(int row, int col) => $"{Col(col)}{row}";

    public static string Range(int r1, int c1, int r2, int c2) =>
        r1 == r2 && c1 == c2 ? Cell(r1, c1) : $"{Cell(r1, c1)}:{Cell(r2, c2)}";

    public static (int Row, int Col) ParseCell(string address)
    {
        var i = 0;
        while (i < address.Length && char.IsLetter(address[i])) i++;
        if (i == 0 || i == address.Length) throw new ArgumentException($"not an A1 cell: {address}");
        return (int.Parse(address[i..], CultureInfo.InvariantCulture), ColIndex(address[..i]));
    }

    public static (int R1, int C1, int R2, int C2) ParseRange(string range)
    {
        var bang = range.LastIndexOf('!');
        if (bang >= 0) range = range[(bang + 1)..];
        range = range.Replace("$", "", StringComparison.Ordinal);
        var parts = range.Split(':');
        var a = ParseCell(parts[0]);
        var b = parts.Length == 1 ? a : ParseCell(parts[1]);
        return (Math.Min(a.Row, b.Row), Math.Min(a.Col, b.Col), Math.Max(a.Row, b.Row), Math.Max(a.Col, b.Col));
    }

    public static int CellCount(string range)
    {
        var (r1, c1, r2, c2) = ParseRange(range);
        return (r2 - r1 + 1) * (c2 - c1 + 1);
    }

    public static bool Overlaps(string a, string b)
    {
        var x = ParseRange(a);
        var y = ParseRange(b);
        return x.R1 <= y.R2 && y.R1 <= x.R2 && x.C1 <= y.C2 && y.C1 <= x.C2;
    }

    public static bool IsMonthHeaderPair(string merge)
    {
        var (r1, c1, r2, c2) = ParseRange(merge);
        return r1 == 4 && r2 == 4 && c2 == c1 + 1 && c1 >= 7 && c1 <= 67 && (c1 - 7) % 2 == 0;
    }
}
