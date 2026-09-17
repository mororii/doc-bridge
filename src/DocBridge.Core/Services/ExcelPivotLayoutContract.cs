using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// Pivot layout merge rules. Omitted dimensions stay on update.
/// Values may request more than one aggregate on the same source field.
/// </summary>
public static class ExcelPivotLayoutContract
{
    public readonly record struct ValueSpec(
        string Field,
        string Function,
        string? Caption,
        string? NumberFormat);

    public static bool AppliesDimension(JsonObject op, string key, bool replaceOmitted) =>
        replaceOmitted || op.ContainsKey(key);

    public static IReadOnlyList<ValueSpec> ReadValues(JsonArray? values)
    {
        if (values is null || values.Count == 0) return [];
        var list = new List<ValueSpec>();
        foreach (var node in values.OfType<JsonObject>())
        {
            var field = Json.GetString(node, "field");
            if (string.IsNullOrWhiteSpace(field)) continue;
            list.Add(new ValueSpec(
                field,
                Json.GetString(node, "function") ?? "sum",
                EmptyToNull(Json.GetString(node, "caption")),
                EmptyToNull(Json.GetString(node, "numberFormat"))));
        }
        return list;
    }

    public static JsonArray ToValuesArray(IEnumerable<ValueSpec> values)
    {
        var array = new JsonArray();
        foreach (var value in values)
        {
            var obj = new JsonObject
            {
                ["field"] = value.Field,
                ["function"] = value.Function,
            };
            if (value.Caption is not null) obj["caption"] = value.Caption;
            if (value.NumberFormat is not null) obj["numberFormat"] = value.NumberFormat;
            array.Add(obj);
        }
        return array;
    }

    private static string? EmptyToNull(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;
}
