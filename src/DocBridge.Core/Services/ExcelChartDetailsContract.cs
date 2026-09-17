using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

public static class ExcelChartDetailsContract
{
    private static readonly HashSet<string> TrendTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "linear", "exponential", "logarithmic", "polynomial", "power", "movingAverage",
    };

    public static void Validate(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        ValidateLabels(op, index, opName, "dataLabels", errors);
        if (Json.GetArr(op, "series") is not JsonArray series) return;
        for (var i = 0; i < series.Count; i++)
        {
            if (series[i] is not JsonObject item) continue;
            ValidateLabels(item, index, opName, $"series[{i}].dataLabels", errors);
            if (item["trendline"] is JsonObject trend) ValidateTrendline(trend, index, opName, i, errors);
            if (item["points"] is JsonArray points)
            {
                foreach (var point in points.OfType<JsonObject>())
                {
                    if (!ExcelDataOperationsContract.TryGetFiniteNumber(point["index"], out var p) ||
                        p < 1 || p != Math.Truncate(p))
                        errors.Add($"ops[{index}] '{opName}' series[{i}].points[].index must be 1-based");
                    ValidateLabels(point, index, opName, $"series[{i}].points[].dataLabels", errors);
                }
            }
        }
    }

    private static void ValidateTrendline(JsonObject trend, int index, string opName, int seriesIndex,
        ICollection<string> errors)
    {
        var prefix = $"ops[{index}] '{opName}' series[{seriesIndex}].trendline";
        var action = Json.GetString(trend, "action");
        var type = Json.GetString(trend, "type");
        if (action is not ("add" or "update" or "delete"))
        {
            errors.Add(prefix + ".action must be add|update|delete");
            return;
        }
        if (action == "add" && !TrendTypes.Contains(type ?? ""))
            errors.Add(prefix + ".type is unsupported");
        if (action == "update" && trend.ContainsKey("type") && !TrendTypes.Contains(type ?? ""))
            errors.Add(prefix + ".type is unsupported");
        if (trend.ContainsKey("index") && (!ExcelDataOperationsContract.TryGetFiniteNumber(trend["index"], out var trendIndex) ||
                                           trendIndex < 1 || trendIndex != Math.Truncate(trendIndex)))
            errors.Add(prefix + ".index must be 1-based");

        var hasOrder = trend.ContainsKey("order");
        if (trend.ContainsKey("expectedRemainingCount"))
        {
            if (action != "delete") errors.Add(prefix + ".expectedRemainingCount is only valid for delete");
            ValidateIntegerRange(trend["expectedRemainingCount"], 0, 64,
                prefix + ".expectedRemainingCount must be integer 0..64", errors);
        }
        var hasPeriod = trend.ContainsKey("period");
        var isPolynomial = string.Equals(type, "polynomial", StringComparison.OrdinalIgnoreCase);
        var isMovingAverage = string.Equals(type, "movingAverage", StringComparison.OrdinalIgnoreCase);
        if ((action == "add" && isPolynomial) || (action == "update" && isPolynomial))
            ValidateIntegerRange(trend["order"], 2, 6, prefix + ".order must be 2..6", errors);
        else if (hasOrder)
            ValidateIntegerRange(trend["order"], 2, 6, prefix + ".order must be 2..6 when supplied", errors);

        if ((action == "add" && isMovingAverage) || (action == "update" && isMovingAverage))
            ValidateIntegerRange(trend["period"], 2, 255,
                prefix + ".period must be integer 2..255; live point-count check occurs before apply", errors);
        else if (hasPeriod)
            ValidateIntegerRange(trend["period"], 2, 255,
                prefix + ".period must be integer 2..255 when supplied", errors);
    }

    private static void ValidateIntegerRange(JsonNode? node, int minimum, int maximum, string message,
        ICollection<string> errors)
    {
        if (!ExcelDataOperationsContract.TryGetFiniteNumber(node, out var value) ||
            value < minimum || value > maximum || value != Math.Truncate(value))
            errors.Add(message);
    }

    private static void ValidateLabels(JsonObject owner, int index, string op, string field,
        ICollection<string> errors)
    {
        if (!owner.ContainsKey("dataLabels")) return;
        if (Json.GetObj(owner, "dataLabels") is not JsonObject labels)
        {
            errors.Add($"ops[{index}] '{op}' {field} must be an object");
            return;
        }
        foreach (var (key, value) in labels)
        {
            if (key is not ("show" or "value" or "category" or "series" or "percentage") ||
                value is not JsonValue jsonValue || !jsonValue.TryGetValue<bool>(out _))
                errors.Add($"ops[{index}] '{op}' {field}.{key} must be supported boolean");
        }
        if (labels.ContainsKey("show") && !Json.GetBool(labels, "show") &&
            (labels.ContainsKey("value") || labels.ContainsKey("category") || labels.ContainsKey("series") ||
             labels.ContainsKey("percentage")))
            errors.Add($"ops[{index}] '{op}' {field}.show=false cannot be combined with label content options");
    }
}
