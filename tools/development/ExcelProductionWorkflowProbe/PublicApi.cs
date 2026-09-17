namespace DocBridge.Development.ExcelProductionWorkflowProbe;

/// <summary>
/// Public Excel read/inspect arguments. The host ignores write-shaped
/// <c>target.workbook</c> on these tools — callers must pass top-level
/// <c>workbook</c> and <c>sheet</c>.
/// </summary>
internal static class PublicApi
{
    public static JsonObject ReadRange(
        PublicClient client,
        string workbook,
        string sheet,
        string range,
        bool formulas = true,
        bool styles = false,
        bool layout = false)
    {
        IdentityGuard.RefuseIfPlaceholder(workbook);
        return AsObject(client.Call("excel_read_range", new JsonObject
        {
            ["workbook"] = workbook,
            ["sheet"] = sheet,
            ["range"] = range,
            ["includeFormulas"] = formulas,
            ["includeStyles"] = styles,
            ["includeLayout"] = layout,
        }));
    }

    public static JsonObject Inspect(
        PublicClient client,
        string workbook,
        string scope,
        string? sheet = null,
        string? objectKind = null)
    {
        IdentityGuard.RefuseIfPlaceholder(workbook);
        var args = new JsonObject
        {
            ["workbook"] = workbook,
            ["scope"] = scope,
        };
        if (!string.IsNullOrWhiteSpace(sheet)) args["sheet"] = sheet;
        if (!string.IsNullOrWhiteSpace(objectKind)) args["objectKind"] = objectKind;
        return AsObject(client.Call("excel_inspect", args));
    }

    public static string? FirstSheetName(JsonNode? scan)
    {
        var sheets = JsonUtil.Get(scan, "sheets") as JsonArray
                     ?? JsonUtil.Get(JsonUtil.Get(scan, "result"), "sheets") as JsonArray;
        if (sheets is null) return null;
        foreach (var sheet in sheets.OfType<JsonObject>())
        {
            var name = JsonUtil.Str(sheet, "name");
            if (!string.IsNullOrWhiteSpace(name)) return name;
        }
        return null;
    }

    public static bool HasReadableValues(JsonNode? read)
    {
        if (JsonUtil.Bool(read, "ok") != true) return false;
        var values = JsonUtil.Get(read, "values") as JsonArray
                     ?? JsonUtil.Get(JsonUtil.Get(read, "result"), "values") as JsonArray;
        var cells = JsonUtil.Get(read, "cells") as JsonArray
                    ?? JsonUtil.Get(JsonUtil.Get(read, "result"), "cells") as JsonArray;
        if (values is { Count: > 0 }) return true;
        if (cells is { Count: > 0 }) return true;
        return false;
    }

    public static JsonNode? ValuesOf(JsonNode? read) =>
        JsonUtil.Get(read, "values")
        ?? JsonUtil.Get(JsonUtil.Get(read, "result"), "values")
        ?? JsonUtil.Get(read, "cells")
        ?? JsonUtil.Get(JsonUtil.Get(read, "result"), "cells");

    public static JsonObject AsObject(JsonNode? node) =>
        node as JsonObject ?? new JsonObject { ["ok"] = false, ["error"] = "public response was not an object" };
}
