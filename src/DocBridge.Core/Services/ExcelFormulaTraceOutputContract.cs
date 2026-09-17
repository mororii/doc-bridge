using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>Removes repeated addressing metadata without discarding cell evidence.</summary>
public static class ExcelFormulaTraceOutputContract
{
    public static JsonObject Compact(JsonObject result)
    {
        var workbook = result["workbook"]?.GetValue<string>();
        if (string.IsNullOrEmpty(workbook))
            throw new ArgumentException("Compact formula trace requires a workbook context.", nameof(result));
        if (result["nodes"] is JsonArray nodes)
        {
            foreach (var node in nodes.OfType<JsonObject>())
            {
                // Preserve an unexpected different workbook rather than erasing provenance.
                if (string.Equals(node["workbook"]?.GetValue<string>(), workbook, StringComparison.OrdinalIgnoreCase))
                    node.Remove("workbook");
            }
        }
        foreach (var collection in new[] { "edges", "cycles", "unresolved" })
        {
            if (result[collection] is not JsonArray items) continue;
            foreach (var item in items.OfType<JsonObject>())
            {
                CompactAddress(item, "from");
                CompactAddress(item, "to");
            }
        }
        result["compact"] = true;
        result["referenceEncoding"] = "quoted-sheet!address; workbook inherited from root";
        return result;
    }

    private static void CompactAddress(JsonObject owner, string property)
    {
        if (owner[property] is not JsonObject address || address.Count != 2 ||
            address["sheet"] is not JsonValue sheetNode || address["address"] is not JsonValue cellNode ||
            !sheetNode.TryGetValue<string>(out var sheet) || !cellNode.TryGetValue<string>(out var cell)) return;
        owner[property] = ExcelFormulaReference.ForceQuoteSheetName(sheet) + "!" + cell;
    }
}
