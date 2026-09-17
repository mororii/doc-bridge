using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

public static class ExcelConditionalFormatContract
{
    public static string Fingerprint(JsonObject rule) => Json.Canonical(new JsonObject
    {
        ["conditionType"] = rule["conditionType"]?.DeepClone(), ["operator"] = rule["operator"]?.DeepClone(), ["formula1"] = rule["formula1"]?.DeepClone(), ["formula2"] = rule["formula2"]?.DeepClone(), ["dupeUnique"] = rule["dupeUnique"]?.DeepClone(), ["style"] = rule["style"]?.DeepClone(),
    });
    public static bool Matches(JsonObject rule, string? expectedFingerprint) => !string.IsNullOrWhiteSpace(expectedFingerprint) && string.Equals(Fingerprint(rule), expectedFingerprint, StringComparison.Ordinal);
    public static bool MatchesRequested(JsonObject actual, JsonObject? rule, JsonObject? style)
    {
        if (Json.GetBool(actual, "unreadable")) return false;
        if (rule is not null)
        {
            var type = Json.GetString(rule, "type");
            var expectedType = type == "expression" ? ExcelDataObjectCatalog.XlExpression : type == "cellValue" ? ExcelDataObjectCatalog.XlCellValue : -1;
            if (expectedType < 0 || Json.GetInt(actual, "conditionType") != expectedType || !string.Equals(Json.GetString(actual, "formula1"), Json.GetString(rule, "formula1"), StringComparison.Ordinal)) return false;
            if (type == "cellValue" && (!ExcelDataObjectCatalog.TryComparisonOperator(Json.GetString(rule, "operator"), out var op) || Json.GetInt(actual, "operator") != op || !string.Equals(Json.GetString(actual, "formula2"), Json.GetString(rule, "formula2"), StringComparison.Ordinal))) return false;
        }
        if (style is not null)
        {
            if (actual["style"] is not JsonObject actualStyle) return false;
            foreach (var pair in style)
                if (!string.Equals(Json.Canonical(pair.Value), Json.Canonical(actualStyle[pair.Key]), StringComparison.Ordinal)) return false;
        }
        return true;
    }
}
