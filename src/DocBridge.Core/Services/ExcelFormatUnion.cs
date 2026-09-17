namespace DocBridge.Core.Services;

/// <summary>
/// Public comma-union ranges (<c>A1:B2,D1:D3</c>) for equal-style disjoint
/// format_range groups. One union Font/Interior write covers every area.
/// Outline/inside borders stay per-area so the bounding box is not painted.
/// ExcelRangeReference still only splits sheet; it does not reject A1 unions.
/// </summary>
public static class ExcelFormatUnion
{
    public const int DefaultMaxAreas = 64;

    public static IReadOnlyList<string> SplitUnionAddresses(string range)
    {
        if (string.IsNullOrWhiteSpace(range))
            return Array.Empty<string>();
        var address = range.Trim();
        try
        {
            var parsed = ExcelRangeReference.Parse(address);
            address = parsed.Address;
        }
        catch (FormatException)
        {
            // Keep the raw token; callers decide whether it is a valid A1 piece.
        }

        var parts = new List<string>();
        var start = 0;
        for (var i = 0; i < address.Length; i++)
        {
            if (address[i] != ',') continue;
            var piece = address[start..i].Trim();
            if (piece.Length > 0) parts.Add(piece);
            start = i + 1;
        }
        var tail = address[start..].Trim();
        if (tail.Length > 0) parts.Add(tail);
        return parts;
    }

    public static IReadOnlySet<string> EnumerateCells(string range)
    {
        var cells = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var piece in SplitUnionAddresses(range))
        {
            if (!ExcelA1Box.TryParse(piece, out var box)) continue;
            for (var row = box.Row; row <= box.EndRow; row++)
            {
                for (var column = box.Column; column <= box.EndColumn; column++)
                    cells.Add(ExcelA1Box.CellName(column, row));
            }
        }
        return cells;
    }

    public static bool HasOffTargetCell(string unionRange, IReadOnlyCollection<string> intendedCells)
    {
        var intended = new HashSet<string>(intendedCells, StringComparer.OrdinalIgnoreCase);
        foreach (var cell in EnumerateCells(unionRange))
        {
            if (!intended.Contains(cell)) return true;
        }
        return false;
    }

    public static IReadOnlyList<string> PlanBoundedUnions(IReadOnlyList<string> areaAddresses, int maxAreas = DefaultMaxAreas)
    {
        ArgumentNullException.ThrowIfNull(areaAddresses);
        if (maxAreas < 1) maxAreas = DefaultMaxAreas;
        var plans = new List<string>();
        var batch = new List<string>();
        foreach (var raw in areaAddresses)
        {
            foreach (var piece in SplitUnionAddresses(raw))
            {
                batch.Add(piece);
                if (batch.Count >= maxAreas)
                {
                    plans.Add(string.Join(',', batch));
                    batch.Clear();
                }
            }
        }
        if (batch.Count > 0)
            plans.Add(string.Join(',', batch));
        return plans;
    }
}
