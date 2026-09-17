using System.Globalization;
using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// Public format_range border write contract. Weight and line style are distinct.
/// <c>left/right/top/bottom</c> apply to every cell edge; <c>outline</c> is the
/// range perimeter; <c>insideHorizontal/insideVertical</c> are interior grid lines.
/// </summary>
public static class ExcelBorderContract
{
    public const int XlDiagonalDown = 5;
    public const int XlDiagonalUp = 6;
    public const int XlEdgeLeft = 7;
    public const int XlEdgeTop = 8;
    public const int XlEdgeBottom = 9;
    public const int XlEdgeRight = 10;
    public const int XlInsideVertical = 11;
    public const int XlInsideHorizontal = 12;

    public const int XlHairline = 1;
    public const int XlThin = 2;
    public const int XlMedium = -4138;
    public const int XlThick = 4;
    public const int XlLineStyleNone = -4142;
    public const int XlContinuous = 1;
    public const int XlDash = -4115;
    public const int XlDot = -4118;
    public const int XlDashDot = 4;
    public const int XlDashDotDot = 5;
    public const int XlDouble = -4119;
    public const int XlSlantDashDot = 13;

    public const string WeightHairline = "hairline";
    public const string WeightThin = "thin";
    public const string WeightMedium = "medium";
    public const string WeightThick = "thick";

    public const string LineNone = "none";
    public const string LineContinuous = "continuous";
    public const string LineDash = "dash";
    public const string LineDot = "dot";
    public const string LineDashDot = "dashDot";
    public const string LineDashDotDot = "dashDotDot";
    public const string LineDouble = "double";
    public const string LineSlantDashDot = "slantDashDot";

    public const string ScopeCell = "cell";
    public const string ScopeRange = "range";

    public static readonly string[] EdgeNames =
    {
        "left", "right", "top", "bottom", "insideHorizontal", "insideVertical", "outline", "all",
    };

    public static IReadOnlyDictionary<string, int> WeightValues { get; } =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            [WeightHairline] = XlHairline,
            [WeightThin] = XlThin,
            [WeightMedium] = XlMedium,
            [WeightThick] = XlThick,
        };

    public static IReadOnlyDictionary<string, int> LineStyleValues { get; } =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            [LineNone] = XlLineStyleNone,
            [LineContinuous] = XlContinuous,
            [LineDash] = XlDash,
            [LineDot] = XlDot,
            [LineDashDot] = XlDashDot,
            [LineDashDotDot] = XlDashDotDot,
            [LineDouble] = XlDouble,
            [LineSlantDashDot] = XlSlantDashDot,
        };

    public sealed record EdgeSpec(string Name, string Scope, string Weight, string LineStyle, double Color);

    public static bool TryNormalize(JsonNode? node, int opIndex, ICollection<string> errors, out JsonObject canonical)
    {
        canonical = new JsonObject();
        if (node is not JsonObject raw)
        {
            errors.Add($"ops[{opIndex}] 'format_range' style.borders must be an object");
            return false;
        }

        if (!TryReadDefaults(raw, opIndex, errors, out var defaultWeight, out var defaultLine, out var defaultColor))
            return false;

        var selected = new Dictionary<string, EdgeSpec>(StringComparer.OrdinalIgnoreCase);
        var ok = true;
        foreach (var (key, value) in raw)
        {
            if (key is "weight" or "lineStyle" or "color") continue;
            if (!EdgeNames.Contains(key, StringComparer.OrdinalIgnoreCase))
            {
                errors.Add(
                    $"ops[{opIndex}] 'format_range' style.borders.{key} is not supported; " +
                    "allowed: left, right, top, bottom, insideHorizontal, insideVertical, outline, all, weight, lineStyle, color");
                ok = false;
                continue;
            }

            if (!TryReadEdge(key, value, defaultWeight, defaultLine, defaultColor, opIndex, errors, out var specs))
            {
                ok = false;
                continue;
            }

            foreach (var spec in specs)
            {
                if (selected.TryGetValue(spec.Name, out var existing) && !SameEdge(existing, spec))
                {
                    errors.Add(
                        $"ops[{opIndex}] 'format_range' style.borders aliases '{existing.Name}' and '{key}' conflict");
                    ok = false;
                    continue;
                }

                selected[spec.Name] = spec;
            }
        }

        if (!ok) return false;
        if (selected.Count == 0)
        {
            errors.Add(
                $"ops[{opIndex}] 'format_range' style.borders must name at least one edge " +
                "(left/right/top/bottom/insideHorizontal/insideVertical/outline/all)");
            return false;
        }

        var edges = new JsonArray();
        foreach (var spec in selected.Values.OrderBy(item => item.Name, StringComparer.Ordinal))
        {
            edges.Add(new JsonObject
            {
                ["name"] = spec.Name,
                ["scope"] = spec.Scope,
                ["weight"] = spec.Weight,
                ["lineStyle"] = spec.LineStyle,
                ["color"] = spec.Color,
                ["weightValue"] = WeightValues[spec.Weight],
                ["lineStyleValue"] = LineStyleValues[spec.LineStyle],
            });
        }

        canonical["edges"] = edges;
        return true;
    }

    public static IReadOnlyList<EdgeSpec> ReadEdges(JsonObject canonical)
    {
        var result = new List<EdgeSpec>();
        foreach (var node in Json.GetArr(canonical, "edges") ?? new JsonArray())
        {
            if (node is not JsonObject edge) continue;
            var name = Json.GetString(edge, "name");
            var scope = Json.GetString(edge, "scope");
            var weight = Json.GetString(edge, "weight");
            var line = Json.GetString(edge, "lineStyle");
            if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(scope) ||
                string.IsNullOrWhiteSpace(weight) || string.IsNullOrWhiteSpace(line))
                continue;
            result.Add(new EdgeSpec(name, scope, weight, line, edge["color"]?.GetValue<double>() ?? 0));
        }
        return result;
    }

    public static int EdgeIndex(string name) => name.ToLowerInvariant() switch
    {
        "left" or "outlineleft" => XlEdgeLeft,
        "right" or "outlineright" => XlEdgeRight,
        "top" or "outlinetop" => XlEdgeTop,
        "bottom" or "outlinebottom" => XlEdgeBottom,
        "insidevertical" => XlInsideVertical,
        "insidehorizontal" => XlInsideHorizontal,
        _ => 0,
    };

    public static string WeightName(int value) => value switch
    {
        XlHairline => WeightHairline,
        XlThin => WeightThin,
        XlMedium => WeightMedium,
        XlThick => WeightThick,
        _ => $"unknown({value})",
    };

    public static string LineStyleName(int value) => value switch
    {
        XlLineStyleNone => LineNone,
        XlContinuous => LineContinuous,
        XlDash => LineDash,
        XlDot => LineDot,
        XlDashDot => LineDashDot,
        XlDashDotDot => LineDashDotDot,
        XlDouble => LineDouble,
        XlSlantDashDot => LineSlantDashDot,
        _ => $"unknown({value})",
    };

    public static JsonObject DescribeSchema() => new()
    {
        ["type"] = "object",
        ["additionalProperties"] = false,
        ["description"] =
            "weight/lineStyle/color defaults plus edges. left/right/top/bottom are each cell edge; " +
            "outline is the range perimeter; insideHorizontal/insideVertical are the range grid; all expands to cell edges plus inside. " +
            "none clears that edge. medium is a weight, not a dash pattern.",
        ["properties"] = new JsonObject
        {
            ["weight"] = WeightSchema(),
            ["lineStyle"] = LineStyleSchema(),
            ["color"] = ColorSchema(),
            ["left"] = EdgeSchema(),
            ["right"] = EdgeSchema(),
            ["top"] = EdgeSchema(),
            ["bottom"] = EdgeSchema(),
            ["insideHorizontal"] = EdgeSchema(),
            ["insideVertical"] = EdgeSchema(),
            ["outline"] = EdgeSchema(),
            ["all"] = EdgeSchema(),
        },
    };

    public static JsonObject ColorSchema() => new()
    {
        ["description"] = "OLE 0..16777215 or #RRGGBB",
        ["anyOf"] = new JsonArray
        {
            new JsonObject { ["type"] = "integer", ["minimum"] = 0, ["maximum"] = 16777215 },
            new JsonObject { ["type"] = "string", ["pattern"] = "^#[0-9A-Fa-f]{6}$" },
        },
    };

    public static JsonObject WeightSchema() => new()
    {
        ["type"] = "string",
        ["enum"] = new JsonArray(WeightHairline, WeightThin, WeightMedium, WeightThick),
    };

    public static JsonObject LineStyleSchema() => new()
    {
        ["type"] = "string",
        ["enum"] = new JsonArray(LineNone, LineContinuous, LineDash, LineDot, LineDashDot, LineDashDotDot, LineDouble, LineSlantDashDot),
    };

    public static JsonObject EdgeSchema() => new()
    {
        ["description"] = "true uses borders defaults; 'none' clears; or {weight?,lineStyle?,color?}",
        ["anyOf"] = new JsonArray
        {
            new JsonObject { ["type"] = "boolean", ["enum"] = new JsonArray(true) },
            new JsonObject { ["type"] = "string", ["enum"] = new JsonArray("none") },
            new JsonObject
            {
                ["type"] = "object",
                ["additionalProperties"] = false,
                ["properties"] = new JsonObject
                {
                    ["weight"] = WeightSchema(),
                    ["lineStyle"] = LineStyleSchema(),
                    ["color"] = ColorSchema(),
                },
            },
        },
    };

    private static bool TryReadDefaults(JsonObject raw, int opIndex, ICollection<string> errors,
        out string weight, out string line, out double color)
    {
        weight = WeightThin;
        line = LineContinuous;
        color = 0;
        if (raw.ContainsKey("weight") && !TryReadWeight(raw["weight"], "weight", opIndex, errors, out weight))
            return false;
        if (raw.ContainsKey("lineStyle") && !TryReadLineStyle(raw["lineStyle"], "lineStyle", opIndex, errors, out line))
            return false;
        if (raw.ContainsKey("color") && !ExcelStyleContract.TryParseColor(raw["color"], out color, errors, "borders.color", opIndex))
            return false;
        return true;
    }

    private static bool TryReadEdge(string key, JsonNode? value, string defaultWeight, string defaultLine,
        double defaultColor, int opIndex, ICollection<string> errors, out List<EdgeSpec> specs)
    {
        specs = new List<EdgeSpec>();
        var names = ExpandEdgeNames(key);
        if (value is JsonValue flag && flag.TryGetValue<bool>(out var include))
        {
            if (!include)
            {
                errors.Add($"ops[{opIndex}] 'format_range' style.borders.{key} false is ignored; omit the key or use 'none'");
                return false;
            }

            foreach (var name in names)
                specs.Add(new EdgeSpec(name, ScopeOf(name), defaultWeight, defaultLine, defaultColor));
            return true;
        }

        if (value is JsonValue text && text.TryGetValue<string>(out var token) &&
            string.Equals(token, LineNone, StringComparison.OrdinalIgnoreCase))
        {
            foreach (var name in names)
                specs.Add(new EdgeSpec(name, ScopeOf(name), defaultWeight, LineNone, defaultColor));
            return true;
        }

        if (value is not JsonObject obj)
        {
            errors.Add($"ops[{opIndex}] 'format_range' style.borders.{key} must be true, 'none', or an object");
            return false;
        }

        var weight = defaultWeight;
        var line = defaultLine;
        var color = defaultColor;
        if (obj.ContainsKey("weight") && !TryReadWeight(obj["weight"], $"{key}.weight", opIndex, errors, out weight))
            return false;
        if (obj.ContainsKey("lineStyle") && !TryReadLineStyle(obj["lineStyle"], $"{key}.lineStyle", opIndex, errors, out line))
            return false;
        if (obj.ContainsKey("color") && !ExcelStyleContract.TryParseColor(obj["color"], out color, errors, $"borders.{key}.color", opIndex))
            return false;

        foreach (var name in names)
            specs.Add(new EdgeSpec(name, ScopeOf(name), weight, line, color));
        return true;
    }

    private static IReadOnlyList<string> ExpandEdgeNames(string key) => key.ToLowerInvariant() switch
    {
        "all" => new[] { "left", "right", "top", "bottom", "insideHorizontal", "insideVertical" },
        "outline" => new[] { "outlineLeft", "outlineRight", "outlineTop", "outlineBottom" },
        _ => new[] { key },
    };

    private static string ScopeOf(string name) =>
        name.StartsWith("outline", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("insideHorizontal", StringComparison.OrdinalIgnoreCase) ||
        name.Equals("insideVertical", StringComparison.OrdinalIgnoreCase)
            ? ScopeRange
            : ScopeCell;

    private static bool SameEdge(EdgeSpec left, EdgeSpec right) =>
        string.Equals(left.Scope, right.Scope, StringComparison.Ordinal) &&
        string.Equals(left.Weight, right.Weight, StringComparison.OrdinalIgnoreCase) &&
        string.Equals(left.LineStyle, right.LineStyle, StringComparison.OrdinalIgnoreCase) &&
        Math.Abs(left.Color - right.Color) < 1e-9;

    private static bool TryReadWeight(JsonNode? node, string key, int opIndex, ICollection<string> errors, out string weight)
    {
        weight = WeightThin;
        if (node is JsonValue value && value.TryGetValue<string>(out var text) &&
            WeightValues.ContainsKey(text))
        {
            weight = WeightValues.Keys.First(item => item.Equals(text, StringComparison.OrdinalIgnoreCase));
            return true;
        }

        errors.Add($"ops[{opIndex}] 'format_range' style.borders.{key} must be hairline|thin|medium|thick");
        return false;
    }

    private static bool TryReadLineStyle(JsonNode? node, string key, int opIndex, ICollection<string> errors, out string line)
    {
        line = LineContinuous;
        if (node is JsonValue value && value.TryGetValue<string>(out var text) &&
            LineStyleValues.ContainsKey(text))
        {
            line = LineStyleValues.Keys.First(item => item.Equals(text, StringComparison.OrdinalIgnoreCase));
            return true;
        }

        errors.Add(
            $"ops[{opIndex}] 'format_range' style.borders.{key} must be none|continuous|dash|dot|dashDot|dashDotDot|double|slantDashDot");
        return false;
    }
}
