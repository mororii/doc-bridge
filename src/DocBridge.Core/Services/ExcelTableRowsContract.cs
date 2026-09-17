using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

public static class ExcelTableRowsContract
{
    public const int MaxRows = 1000;

    public static void ValidateRows(JsonObject op, int index, ICollection<string> errors)
    {
        var rows = Json.GetArr(op, "rows");
        if (rows is null || rows.Count is < 1 or > MaxRows)
        {
            errors.Add($"ops[{index}] '{Json.GetString(op, "op")}' rows must contain 1..{MaxRows} rows");
            return;
        }
        int? width = null;
        for (var rowIndex = 0; rowIndex < rows.Count; rowIndex++)
        {
            if (rows[rowIndex] is JsonObject sparse)
            {
                if (sparse.Any(p => p.Value is JsonObject or JsonArray))
                    errors.Add($"ops[{index}] rows[{rowIndex}] values must be scalars or null");
                continue;
            }
            if (rows[rowIndex] is not JsonArray row || row.Count == 0)
            {
                errors.Add($"ops[{index}] rows[{rowIndex}] must be a non-empty array or sparse object");
                continue;
            }
            width ??= row.Count;
            if (row.Count != width) errors.Add($"ops[{index}] array rows must be rectangular");
            if (row.Any(value => value is JsonObject or JsonArray))
                errors.Add($"ops[{index}] rows[{rowIndex}] values must be scalars or null");
        }
    }

    public static void ValidateKnownColumns(JsonArray rows, IReadOnlyCollection<string> columnNames)
    {
        var known = new HashSet<string>(columnNames, StringComparer.Ordinal);
        foreach (var row in rows.OfType<JsonObject>())
            foreach (var field in row)
                if (!known.Contains(field.Key))
                    throw new InvalidOperationException($"[EXCEL_TABLE_ROW_COLUMN] unknown table column '{field.Key}'");
    }
}
