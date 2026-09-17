namespace DocBridge.Development.ExcelProductionWorkflowProbe;

/// <summary>
/// Public <c>set_column_widths</c>/<c>set_row_heights</c> items. Oracle
/// <c>unit</c>/<c>note</c> stay in COM/XML evidence, not in apply ops.
/// </summary>
internal static class PublicLayout
{
    private static readonly string[] ColumnKeys = ["col", "count", "widthChars", "autoFit"];
    private static readonly string[] RowKeys = ["row", "count", "heightPoints", "autoFit"];

    public static JsonArray Columns(JsonNode? source) => Project(source, ColumnKeys);

    public static JsonArray Rows(JsonNode? source) => Project(source, RowKeys);

    public static bool HasOracleMetadata(JsonNode? source)
    {
        if (source is not JsonArray arr) return false;
        return arr.OfType<JsonObject>().Any(item =>
            item.ContainsKey("unit") || item.ContainsKey("note"));
    }

    private static JsonArray Project(JsonNode? source, IReadOnlyList<string> keys)
    {
        var dest = new JsonArray();
        if (source is not JsonArray arr) return dest;
        foreach (var item in arr.OfType<JsonObject>())
        {
            var copy = new JsonObject();
            foreach (var key in keys)
            {
                if (item[key] is { } node)
                    copy[key] = node.DeepClone();
            }
            if (copy.Count > 0)
                dest.Add(copy);
        }
        return dest;
    }
}
