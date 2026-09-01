namespace DocBridge.Core.Services;

/// <summary>
/// Parses A1-style Excel range references with an optional worksheet prefix.
/// Supports unquoted names (Sheet1!A1) and Excel-quoted names
/// ('공사 내역'!A1:B3, including doubled apostrophes).
/// </summary>
internal readonly record struct ExcelRangeReference(string? SheetName, string Address)
{
    internal static ExcelRangeReference Parse(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
            throw new FormatException("Excel range must not be empty");

        var text = value.Trim();
        var separator = -1;
        var inQuotes = false;
        for (var index = 0; index < text.Length; index++)
        {
            var ch = text[index];
            if (ch == '\'')
            {
                if (inQuotes && index + 1 < text.Length && text[index + 1] == '\'')
                {
                    index++;
                    continue;
                }
                inQuotes = !inQuotes;
                continue;
            }
            if (ch != '!' || inQuotes) continue;
            if (separator >= 0)
                throw new FormatException($"Excel range contains more than one sheet separator: {value}");
            separator = index;
        }
        if (inQuotes)
            throw new FormatException($"Excel range contains an unterminated sheet quote: {value}");
        if (separator < 0)
            return new ExcelRangeReference(null, text);

        var sheetPart = text[..separator].Trim();
        var address = text[(separator + 1)..].Trim();
        if (sheetPart.Length == 0 || address.Length == 0)
            throw new FormatException($"Excel range must contain both sheet and address: {value}");

        string sheetName;
        if (sheetPart.StartsWith('\'') || sheetPart.EndsWith('\''))
        {
            if (sheetPart.Length < 2 || !sheetPart.StartsWith('\'') || !sheetPart.EndsWith('\''))
                throw new FormatException($"Excel sheet quoting is invalid: {value}");
            sheetName = sheetPart[1..^1].Replace("''", "'", StringComparison.Ordinal);
        }
        else
        {
            sheetName = sheetPart;
        }
        if (string.IsNullOrWhiteSpace(sheetName))
            throw new FormatException($"Excel sheet name must not be empty: {value}");
        return new ExcelRangeReference(sheetName, address);
    }
}
