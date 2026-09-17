using System.Text;
using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// RFC 4180 CSV parse/write for Excel import_csv / export_csv.
/// Quoted commas, escaped quotes, and CRLF are preserved. Explicit empty
/// quoted records stay empty strings. Leading zeros and formula-looking
/// text stay text. Malformed quote sequences are refused before any caller
/// can mutate a sheet. <see cref="NormalizeRectangle"/> checks
/// rows×maxWidth against the import cell limit before allocating padding.
/// Empty fields are written as <c>""</c> so a terminal empty record
/// round-trips. ResolveDelimiter reads the public op JSON directly and
/// does not depend on other service helpers.
/// </summary>
public static class ExcelCsvContract
{
    public const int MaxExportCells = 1_000_000;
    public const int MaxImportCells = 1_000_000;

    public static IReadOnlyList<IReadOnlyList<string>> Parse(string text, char delimiter = ',', int maxCells = MaxImportCells)
    {
        if (text is null) throw new ArgumentNullException(nameof(text));
        var rows = new List<List<string>>();
        if (text.Length == 0)
            return rows;

        var row = new List<string>();
        var field = new StringBuilder();
        var i = 0;
        var quoted = false;
        var fieldOpen = false;
        var lastWasDelimiter = false;
        var cells = 0;

        while (i < text.Length)
        {
            var ch = text[i];
            if (quoted)
            {
                if (ch == '"')
                {
                    if (i + 1 < text.Length && text[i + 1] == '"')
                    {
                        field.Append('"');
                        i += 2;
                        continue;
                    }

                    quoted = false;
                    fieldOpen = true;
                    i++;
                    if (i < text.Length)
                    {
                        var next = text[i];
                        if (next != delimiter && next != '\r' && next != '\n')
                            throw new InvalidOperationException("import_csv has a malformed quote sequence");
                    }

                    continue;
                }

                field.Append(ch);
                i++;
                continue;
            }

            if (ch == '"')
            {
                if (fieldOpen)
                    throw new InvalidOperationException("import_csv has a malformed quote sequence");
                quoted = true;
                fieldOpen = true;
                lastWasDelimiter = false;
                i++;
                continue;
            }

            if (ch == delimiter)
            {
                AddField(row, field, ref cells, maxCells);
                fieldOpen = false;
                lastWasDelimiter = true;
                i++;
                continue;
            }

            if (ch == '\r' || ch == '\n')
            {
                if (ch == '\r' && i + 1 < text.Length && text[i + 1] == '\n')
                    i++;
                AddField(row, field, ref cells, maxCells);
                rows.Add(row);
                row = new List<string>();
                fieldOpen = false;
                lastWasDelimiter = false;
                i++;
                continue;
            }

            field.Append(ch);
            fieldOpen = true;
            lastWasDelimiter = false;
            i++;
        }

        if (quoted)
            throw new InvalidOperationException("import_csv has an unclosed quoted field");
        if (fieldOpen || lastWasDelimiter || row.Count > 0)
        {
            AddField(row, field, ref cells, maxCells);
            rows.Add(row);
        }

        return rows;
    }

    private static void AddField(List<string> row, StringBuilder field, ref int cells, int maxCells)
    {
        cells++;
        if (cells > maxCells)
            throw new InvalidOperationException(
                $"import_csv exceeded {maxCells} cells while parsing; split the file");
        row.Add(field.ToString());
        field.Clear();
    }

    public static IReadOnlyList<IReadOnlyList<string>> NormalizeRectangle(
        IReadOnlyList<IReadOnlyList<string>> rows, int maxCells = MaxImportCells)
    {
        ArgumentNullException.ThrowIfNull(rows);
        var width = 0;
        foreach (var row in rows)
        {
            if (row is not null && row.Count > width)
                width = row.Count;
        }

        if (ImportLimitError(rows.Count, width, maxCells) is string limitError)
            throw new InvalidOperationException(limitError);

        var result = new List<IReadOnlyList<string>>(rows.Count);
        foreach (var row in rows)
        {
            var cells = new string[width];
            var source = row ?? Array.Empty<string>();
            for (var i = 0; i < width; i++)
                cells[i] = i < source.Count ? source[i] ?? "" : "";
            result.Add(cells);
        }

        return result;
    }

    public static string Write(IReadOnlyList<IReadOnlyList<string>> rows, char delimiter = ',')
    {
        ArgumentNullException.ThrowIfNull(rows);
        var builder = new StringBuilder();
        for (var r = 0; r < rows.Count; r++)
        {
            if (r > 0)
                builder.Append("\r\n");
            var row = rows[r] ?? Array.Empty<string>();
            for (var c = 0; c < row.Count; c++)
            {
                if (c > 0)
                    builder.Append(delimiter);
                builder.Append(Escape(row[c], delimiter));
            }
        }

        return builder.ToString();
    }

    public static string Escape(string? value, char delimiter = ',')
    {
        var text = value ?? "";
        if (text.Length == 0)
            return "\"\"";
        var mustQuote = text.Contains(delimiter) || text.Contains('"') ||
                        text.Contains('\r') || text.Contains('\n');
        if (!mustQuote)
            return text;
        return "\"" + text.Replace("\"", "\"\"") + "\"";
    }

    public static bool LooksLikeIdentifierOrFormulaText(string value) =>
        value.Length > 0 &&
        (value[0] == '0' && value.Any(char.IsDigit) ||
         value.StartsWith('=') ||
         value.StartsWith('+') ||
         value.StartsWith('-') && value.Length > 1 && !double.TryParse(value, out _));

    public static string? ExportLimitError(int rows, int columns)
    {
        var cells = (long)rows * columns;
        if (cells <= MaxExportCells)
            return null;
        return $"export_csv range is {rows}x{columns} ({cells} cells) and exceeds {MaxExportCells}; split the range";
    }

    public static string? ImportLimitError(int rows, int columns, int maxCells = MaxImportCells)
    {
        var cells = (long)rows * columns;
        if (cells <= maxCells)
            return null;
        return $"import_csv table is {rows}x{columns} ({cells} cells) and exceeds {maxCells}; split the file";
    }

    public static char ResolveDelimiter(JsonObject? op)
    {
        if (op is null || !op.TryGetPropertyValue("delimiter", out var node) || node is null)
            return ',';
        if (node is not JsonValue value || !value.TryGetValue<string>(out var text) || string.IsNullOrEmpty(text))
            return ',';
        if (text.Length != 1)
            throw new InvalidOperationException("delimiter must be a single character");
        return text[0];
    }

    public static bool TablesEqual(
        IReadOnlyList<IReadOnlyList<string>> left,
        IReadOnlyList<IReadOnlyList<string>> right)
    {
        if (left.Count != right.Count)
            return false;
        for (var r = 0; r < left.Count; r++)
        {
            if (left[r].Count != right[r].Count)
                return false;
            for (var c = 0; c < left[r].Count; c++)
            {
                if (!string.Equals(left[r][c], right[r][c], StringComparison.Ordinal))
                    return false;
            }
        }

        return true;
    }

    public static string CellReadbackMismatch(int row, int column, string expected, string actual) =>
        $"import_csv {ColName(column)}{row}: want '{expected}', got '{actual}'";

    private static string ColName(int column)
    {
        var name = "";
        var c = column;
        while (c > 0)
        {
            var m = (c - 1) % 26;
            name = (char)('A' + m) + name;
            c = (c - 1) / 26;
        }

        return name;
    }
}
