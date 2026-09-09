using System.Globalization;
using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// Excel format_range style contract: write/read aliases, type and color
/// validation, and alias-conflict rejection before snapshot/apply.
/// </summary>
public static class ExcelStyleContract
{
    public const string Bold = "bold";
    public const string Italic = "italic";
    public const string FontSize = "fontSize";
    public const string NumberFormat = "numberFormat";
    public const string FontColor = "fontColor";
    public const string FillColor = "fillColor";

    public const int XlPatternNone = -4142;
    public const int XlPatternSolid = 1;
    public const int XlColorIndexAutomatic = -4105;
    public const int XlColorIndexNone = -4142;
    public const int XlThemeColorMin = 1;
    public const int XlThemeColorMax = 12;
    public const double MaxOleColor = 16_777_215d;

    private static readonly Dictionary<string, string> CanonicalKeys = new(StringComparer.OrdinalIgnoreCase)
    {
        [Bold] = Bold,
        ["fontBold"] = Bold,
        [Italic] = Italic,
        ["fontItalic"] = Italic,
        [FontSize] = FontSize,
        [NumberFormat] = NumberFormat,
        [FontColor] = FontColor,
        [FillColor] = FillColor,
        ["interiorColor"] = FillColor,
        ["fill"] = FillColor,
    };

    public static IReadOnlyList<string> SupportedKeys { get; } = new[]
    {
        Bold, "fontBold", Italic, "fontItalic", FontSize, NumberFormat, FontColor, FillColor, "interiorColor", "fill",
    };

    public static string StyleExpectation =>
        "object using bold|fontBold, italic|fontItalic, fontSize, numberFormat, fontColor, fillColor|interiorColor|fill; " +
        "unknown keys, nulls, conflicting aliases, and invalid colors are rejected before snapshot";

    public static string ReadStyleExpectation =>
        "when true, styles include write keys bold/italic/fontSize/numberFormat/fontColor/fillColor, " +
        "read aliases fontBold/fontItalic/interiorColor/fill, fillPattern, and fillColorIndex; " +
        "no-fill is Pattern=none, not Color=0; mixed cells are null";

    public static bool TryNormalize(JsonObject? style, out JsonObject canonical, ICollection<string> errors, int opIndex = 1)
    {
        canonical = new JsonObject();
        if (style is null)
        {
            errors.Add($"ops[{opIndex}] 'format_range' field 'style' must be {StyleExpectation}");
            return false;
        }

        var pending = new Dictionary<string, (string Source, JsonNode Node)>(StringComparer.Ordinal);
        var ok = true;
        foreach (var (rawKey, node) in style)
        {
            if (!CanonicalKeys.TryGetValue(rawKey, out var canonicalKey))
            {
                errors.Add(
                    $"ops[{opIndex}] 'format_range' style key '{rawKey}' is not supported; " +
                    $"allowed: {string.Join(", ", SupportedKeys)}");
                ok = false;
                continue;
            }

            if (node is null)
            {
                errors.Add($"ops[{opIndex}] 'format_range' style.{rawKey} must not be null; omit the key to leave that property unchanged");
                ok = false;
                continue;
            }

            if (pending.TryGetValue(canonicalKey, out var existing) &&
                !SameCanonicalValue(canonicalKey, existing.Node, node))
            {
                errors.Add(
                    $"ops[{opIndex}] 'format_range' style aliases '{existing.Source}' and '{rawKey}' conflict");
                ok = false;
                continue;
            }

            if (!TryCoerce(canonicalKey, rawKey, node, opIndex, errors, out var coerced))
            {
                ok = false;
                continue;
            }

            pending[canonicalKey] = (rawKey, coerced);
        }

        if (!ok) return false;
        foreach (var (key, pair) in pending.OrderBy(kv => kv.Key, StringComparer.Ordinal))
            canonical[key] = pair.Node.DeepClone();
        return true;
    }

    public static void Validate(JsonObject? style, int opIndex, ICollection<string> errors) =>
        TryNormalize(style, out _, errors, opIndex);

    public static bool TryParseColor(JsonNode? node, out double ole, ICollection<string>? errors = null, string key = FillColor, int opIndex = 1)
    {
        ole = 0;
        if (node is not JsonValue value)
        {
            errors?.Add($"ops[{opIndex}] 'format_range' style.{key} must be an OLE color 0..16777215 or '#RRGGBB'");
            return false;
        }

        if (value.TryGetValue<string>(out var hex))
        {
            if (hex.Length == 7 && hex[0] == '#')
            {
                try
                {
                    var r = Convert.ToInt32(hex[1..3], 16);
                    var g = Convert.ToInt32(hex[3..5], 16);
                    var b = Convert.ToInt32(hex[5..7], 16);
                    ole = r | (g << 8) | (b << 16);
                    return true;
                }
                catch (FormatException)
                {
                    // handled below
                }
            }
            errors?.Add($"ops[{opIndex}] 'format_range' style.{key} color '{hex}' must be '#RRGGBB'");
            return false;
        }

        if (TryGetFiniteNumber(value, out var number) && number is >= 0 and <= MaxOleColor &&
            Math.Abs(number - Math.Round(number)) < 1e-9)
        {
            ole = number;
            return true;
        }

        errors?.Add($"ops[{opIndex}] 'format_range' style.{key} must be an OLE color 0..16777215 or '#RRGGBB'");
        return false;
    }

    public static double ParseColor(JsonNode node) =>
        TryParseColor(node, out var ole, errors: null)
            ? ole
            : throw new ArgumentException("color must be OLE int or '#RRGGBB'");

    public static JsonObject DescribeStyleSchema() => new()
    {
        ["type"] = "object",
        ["description"] = StyleExpectation,
        ["additionalProperties"] = false,
        ["properties"] = new JsonObject
        {
            [Bold] = new JsonObject { ["type"] = "boolean" },
            ["fontBold"] = new JsonObject { ["type"] = "boolean", ["description"] = "read alias for bold" },
            [Italic] = new JsonObject { ["type"] = "boolean" },
            ["fontItalic"] = new JsonObject { ["type"] = "boolean", ["description"] = "read alias for italic" },
            [FontSize] = new JsonObject { ["type"] = "number", ["exclusiveMinimum"] = 0 },
            [NumberFormat] = new JsonObject { ["type"] = "string" },
            [FontColor] = new JsonObject { ["description"] = "OLE 0..16777215 or #RRGGBB" },
            [FillColor] = new JsonObject { ["description"] = "OLE 0..16777215 or #RRGGBB; no-fill is a Pattern=none state, not Color=0" },
            ["interiorColor"] = new JsonObject { ["description"] = "read alias for fillColor" },
            ["fill"] = new JsonObject { ["description"] = "write alias for fillColor" },
        },
    };

    public static JsonObject WithReadAliases(JsonObject style)
    {
        var projected = style.DeepClone().AsObject();
        CopyAlias(projected, Bold, "fontBold");
        CopyAlias(projected, Italic, "fontItalic");
        CopyAlias(projected, FillColor, "interiorColor");
        CopyAlias(projected, FillColor, "fill");
        return projected;
    }

    private static void CopyAlias(JsonObject style, string source, string alias)
    {
        if (!style.ContainsKey(source) || style.ContainsKey(alias)) return;
        style[alias] = style[source]?.DeepClone();
    }

    private static bool TryCoerce(string canonicalKey, string sourceKey, JsonNode node, int opIndex,
        ICollection<string> errors, out JsonNode coerced)
    {
        coerced = node;
        switch (canonicalKey)
        {
            case Bold:
            case Italic:
                if (node is JsonValue flag && flag.TryGetValue<bool>(out var boolean))
                {
                    coerced = JsonValue.Create(boolean)!;
                    return true;
                }
                errors.Add($"ops[{opIndex}] 'format_range' style.{sourceKey} must be boolean");
                return false;
            case FontSize:
                if (node is JsonValue size && TryGetFiniteNumber(size, out var points) && points > 0)
                {
                    coerced = JsonValue.Create(points)!;
                    return true;
                }
                errors.Add($"ops[{opIndex}] 'format_range' style.{sourceKey} must be a finite number greater than 0");
                return false;
            case NumberFormat:
                if (node is JsonValue format && format.TryGetValue<string>(out var text) && text is not null)
                {
                    coerced = JsonValue.Create(text)!;
                    return true;
                }
                errors.Add($"ops[{opIndex}] 'format_range' style.{sourceKey} must be string");
                return false;
            case FontColor:
            case FillColor:
                if (!TryParseColor(node, out var ole, errors, sourceKey, opIndex)) return false;
                coerced = JsonValue.Create(ole)!;
                return true;
            default:
                errors.Add($"ops[{opIndex}] 'format_range' style.{sourceKey} is not supported");
                return false;
        }
    }

    private static bool SameCanonicalValue(string canonicalKey, JsonNode left, JsonNode right)
    {
        if (canonicalKey is FontColor or FillColor &&
            TryParseColor(left, out var leftColor) && TryParseColor(right, out var rightColor))
            return Math.Abs(leftColor - rightColor) < 1e-9;
        return string.Equals(Json.Canonical(left), Json.Canonical(right), StringComparison.Ordinal);
    }

    private static bool TryGetFiniteNumber(JsonValue value, out double number)
    {
        if (value.TryGetValue<int>(out var i)) { number = i; return true; }
        if (value.TryGetValue<long>(out var l)) { number = l; return true; }
        if (value.TryGetValue<uint>(out var ui)) { number = ui; return true; }
        if (value.TryGetValue<ulong>(out var ul) && ul <= (ulong)long.MaxValue) { number = ul; return true; }
        if (value.TryGetValue<decimal>(out var dec))
        {
            number = Convert.ToDouble(dec, CultureInfo.InvariantCulture);
            return double.IsFinite(number);
        }
        if (value.TryGetValue<float>(out var f)) { number = f; return double.IsFinite(number); }
        if (value.TryGetValue<double>(out number)) return double.IsFinite(number);
        number = 0;
        return false;
    }
}
