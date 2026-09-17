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
    public const string FontName = "fontName";
    public const string HorizontalAlign = "horizontalAlign";
    public const string VerticalAlign = "verticalAlign";
    public const string WrapText = "wrapText";
    public const string Borders = "borders";
    public const string ShrinkToFit = "shrinkToFit";
    public const string Underline = "underline";
    public const string Strikethrough = "strikethrough";
    public const string Indent = "indent";
    public const string Orientation = "orientation";
    public const string Locked = "locked";
    public const string FillPattern = "fillPattern";
    public const string NoFill = "noFill";

    public const int XlUnderlineNone = -4142;
    public const int XlUnderlineSingle = 2;
    public const int XlUnderlineDouble = -4119;
    public const int XlUnderlineSingleAccounting = 4;
    public const int XlUnderlineDoubleAccounting = 5;
    public const int XlOrientationStacked = 255;

    public const int XlHAlignGeneral = 1;
    public const int XlHAlignLeft = -4131;
    public const int XlHAlignCenter = -4108;
    public const int XlHAlignRight = -4152;
    public const int XlHAlignFill = 5;
    public const int XlHAlignJustify = -4130;
    public const int XlHAlignCenterAcross = 7;
    public const int XlHAlignDistributed = -4117;
    public const int XlVAlignTop = -4160;
    public const int XlVAlignCenter = -4108;
    public const int XlVAlignBottom = -4107;
    public const int XlVAlignJustify = -4130;
    public const int XlVAlignDistributed = -4117;

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
        [FontName] = FontName,
        ["font"] = FontName,
        [HorizontalAlign] = HorizontalAlign,
        ["hAlign"] = HorizontalAlign,
        ["align"] = HorizontalAlign,
        [VerticalAlign] = VerticalAlign,
        ["vAlign"] = VerticalAlign,
        [WrapText] = WrapText,
        [Borders] = Borders,
        [ShrinkToFit] = ShrinkToFit,
        [Underline] = Underline,
        [Strikethrough] = Strikethrough,
        ["strike"] = Strikethrough,
        [Indent] = Indent,
        [Orientation] = Orientation,
        [Locked] = Locked,
        [FillPattern] = FillPattern,
        [NoFill] = NoFill,
    };

    public static readonly IReadOnlyDictionary<string, int> UnderlineValues =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["none"] = XlUnderlineNone,
            ["single"] = XlUnderlineSingle,
            ["double"] = XlUnderlineDouble,
            ["singleAccounting"] = XlUnderlineSingleAccounting,
            ["doubleAccounting"] = XlUnderlineDoubleAccounting,
        };

    public static readonly IReadOnlyDictionary<string, int> FillPatternValues =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["none"] = XlPatternNone,
            ["solid"] = XlPatternSolid,
            ["gray75"] = -4126,
            ["gray50"] = -4125,
            ["gray25"] = -4124,
            ["gray16"] = 17,
            ["gray8"] = 18,
            ["lightHorizontal"] = 11,
            ["lightVertical"] = 12,
        };

    public static readonly IReadOnlyDictionary<string, int> HorizontalAlignValues =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["general"] = XlHAlignGeneral,
            ["left"] = XlHAlignLeft,
            ["center"] = XlHAlignCenter,
            ["right"] = XlHAlignRight,
            ["fill"] = XlHAlignFill,
            ["justify"] = XlHAlignJustify,
            ["centerAcross"] = XlHAlignCenterAcross,
            ["distributed"] = XlHAlignDistributed,
        };

    public static readonly IReadOnlyDictionary<string, int> VerticalAlignValues =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["top"] = XlVAlignTop,
            ["center"] = XlVAlignCenter,
            ["bottom"] = XlVAlignBottom,
            ["justify"] = XlVAlignJustify,
            ["distributed"] = XlVAlignDistributed,
        };

    public static IReadOnlyList<string> SupportedKeys { get; } = new[]
    {
        Bold, "fontBold", Italic, "fontItalic", FontSize, NumberFormat, FontColor, FillColor, "interiorColor", "fill",
        FontName, "font", HorizontalAlign, "hAlign", "align", VerticalAlign, "vAlign", WrapText, Borders,
        ShrinkToFit, Underline, Strikethrough, "strike", Indent, Orientation, Locked, FillPattern, NoFill,
    };

    public static string StyleExpectation =>
        "object using bold|fontBold, italic|fontItalic, fontSize, numberFormat, fontColor, fillColor|interiorColor|fill, " +
        "fontName|font, horizontalAlign|hAlign|align, verticalAlign|vAlign, wrapText, borders, " +
        "shrinkToFit, underline, strikethrough|strike, indent, orientation, locked, fillPattern, noFill; " +
        "unknown keys, nulls, conflicting aliases, and invalid colors are rejected before snapshot";

    public static string ReadStyleExpectation =>
        "when true, styles include write keys bold/italic/fontSize/numberFormat/fontColor/fillColor/" +
        "fontName/horizontalAlign/verticalAlign/wrapText/borders, " +
        "read aliases fontBold/fontItalic/interiorColor/fill/font/hAlign/vAlign, fillPattern, and fillColorIndex; " +
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
        if (pending.ContainsKey(NoFill) && pending.ContainsKey(FillColor))
        {
            errors.Add($"ops[{opIndex}] 'format_range' style cannot set both noFill and fillColor");
            return false;
        }

        if (pending.ContainsKey(NoFill) && pending.ContainsKey(FillPattern) &&
            pending[FillPattern].Node is JsonValue patternNode &&
            patternNode.TryGetValue<string>(out var patternName) &&
            !patternName.Equals("none", StringComparison.OrdinalIgnoreCase))
        {
            errors.Add($"ops[{opIndex}] 'format_range' style cannot set noFill with fillPattern '{patternName}'");
            return false;
        }

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
            [FontName] = new JsonObject { ["type"] = "string", ["minLength"] = 1 },
            ["font"] = new JsonObject { ["type"] = "string", ["description"] = "write alias for fontName" },
            [HorizontalAlign] = new JsonObject
            {
                ["type"] = "string",
                ["enum"] = new JsonArray("general", "left", "center", "right", "fill", "justify", "centerAcross", "distributed"),
            },
            ["hAlign"] = new JsonObject { ["type"] = "string", ["description"] = "write alias for horizontalAlign" },
            ["align"] = new JsonObject { ["type"] = "string", ["description"] = "write alias for horizontalAlign" },
            [VerticalAlign] = new JsonObject
            {
                ["type"] = "string",
                ["enum"] = new JsonArray("top", "center", "bottom", "justify", "distributed"),
            },
            ["vAlign"] = new JsonObject { ["type"] = "string", ["description"] = "write alias for verticalAlign" },
            [WrapText] = new JsonObject { ["type"] = "boolean" },
            [Borders] = ExcelBorderContract.DescribeSchema(),
            [ShrinkToFit] = new JsonObject { ["type"] = "boolean" },
            [Underline] = new JsonObject
            {
                ["type"] = "string",
                ["enum"] = new JsonArray("none", "single", "double", "singleAccounting", "doubleAccounting"),
            },
            [Strikethrough] = new JsonObject { ["type"] = "boolean" },
            ["strike"] = new JsonObject { ["type"] = "boolean", ["description"] = "write alias for strikethrough" },
            [Indent] = new JsonObject { ["type"] = "integer", ["minimum"] = 0, ["maximum"] = 15 },
            [Orientation] = new JsonObject { ["description"] = "integer -90..90 or stacked" },
            [Locked] = new JsonObject { ["type"] = "boolean" },
            [FillPattern] = new JsonObject
            {
                ["type"] = "string",
                ["enum"] = new JsonArray("none", "solid", "gray75", "gray50", "gray25", "gray16", "gray8", "lightHorizontal", "lightVertical"),
            },
            [NoFill] = new JsonObject { ["type"] = "boolean", ["description"] = "true sets Pattern=none; do not combine with fillColor" },
        },
    };

    public static JsonObject WithReadAliases(JsonObject style)
    {
        var projected = style.DeepClone().AsObject();
        CopyAlias(projected, Bold, "fontBold");
        CopyAlias(projected, Italic, "fontItalic");
        CopyAlias(projected, FillColor, "interiorColor");
        CopyAlias(projected, FillColor, "fill");
        CopyAlias(projected, FontName, "font");
        CopyAlias(projected, HorizontalAlign, "hAlign");
        CopyAlias(projected, VerticalAlign, "vAlign");
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
            case WrapText:
            case ShrinkToFit:
            case Strikethrough:
            case Locked:
            case NoFill:
                if (node is JsonValue flag && flag.TryGetValue<bool>(out var boolean))
                {
                    coerced = JsonValue.Create(boolean)!;
                    return true;
                }
                errors.Add($"ops[{opIndex}] 'format_range' style.{sourceKey} must be boolean");
                return false;
            case FontName:
                if (node is JsonValue font && font.TryGetValue<string>(out var fontName) &&
                    !string.IsNullOrWhiteSpace(fontName))
                {
                    coerced = JsonValue.Create(fontName.Trim())!;
                    return true;
                }
                errors.Add($"ops[{opIndex}] 'format_range' style.{sourceKey} must be a non-empty font name");
                return false;
            case HorizontalAlign:
                if (node is JsonValue h && h.TryGetValue<string>(out var hAlign) &&
                    HorizontalAlignValues.ContainsKey(hAlign))
                {
                    coerced = JsonValue.Create(
                        HorizontalAlignValues.Keys.First(item => item.Equals(hAlign, StringComparison.OrdinalIgnoreCase)))!;
                    return true;
                }
                errors.Add(
                    $"ops[{opIndex}] 'format_range' style.{sourceKey} must be general|left|center|right|fill|justify|centerAcross|distributed");
                return false;
            case VerticalAlign:
                if (node is JsonValue v && v.TryGetValue<string>(out var vAlign) &&
                    VerticalAlignValues.ContainsKey(vAlign))
                {
                    coerced = JsonValue.Create(
                        VerticalAlignValues.Keys.First(item => item.Equals(vAlign, StringComparison.OrdinalIgnoreCase)))!;
                    return true;
                }
                errors.Add($"ops[{opIndex}] 'format_range' style.{sourceKey} must be top|center|bottom|justify|distributed");
                return false;
            case Underline:
                if (node is JsonValue underline && underline.TryGetValue<string>(out var underlineName) &&
                    UnderlineValues.ContainsKey(underlineName))
                {
                    coerced = JsonValue.Create(
                        UnderlineValues.Keys.First(item => item.Equals(underlineName, StringComparison.OrdinalIgnoreCase)))!;
                    return true;
                }
                errors.Add($"ops[{opIndex}] 'format_range' style.{sourceKey} must be none|single|double|singleAccounting|doubleAccounting");
                return false;
            case Indent:
                if (node is JsonValue indent && indent.TryGetValue<int>(out var indentLevel) && indentLevel is >= 0 and <= 15)
                {
                    coerced = JsonValue.Create(indentLevel)!;
                    return true;
                }
                errors.Add($"ops[{opIndex}] 'format_range' style.{sourceKey} must be an integer 0..15");
                return false;
            case Orientation:
                if (node is JsonValue stacked && stacked.TryGetValue<string>(out var orientationName) &&
                    orientationName.Equals("stacked", StringComparison.OrdinalIgnoreCase))
                {
                    coerced = JsonValue.Create("stacked")!;
                    return true;
                }
                if (node is JsonValue degrees && TryGetFiniteNumber(degrees, out var angle) &&
                    angle is >= -90 and <= 90 && Math.Abs(angle - Math.Round(angle)) < 1e-9)
                {
                    coerced = JsonValue.Create((int)Math.Round(angle))!;
                    return true;
                }
                errors.Add($"ops[{opIndex}] 'format_range' style.{sourceKey} must be an integer -90..90 or stacked");
                return false;
            case FillPattern:
                if (node is JsonValue pattern && pattern.TryGetValue<string>(out var patternName) &&
                    FillPatternValues.ContainsKey(patternName))
                {
                    coerced = JsonValue.Create(
                        FillPatternValues.Keys.First(item => item.Equals(patternName, StringComparison.OrdinalIgnoreCase)))!;
                    return true;
                }
                errors.Add($"ops[{opIndex}] 'format_range' style.{sourceKey} must be a supported fillPattern name");
                return false;
            case Borders:
                if (!ExcelBorderContract.TryNormalize(node, opIndex, errors, out var borders))
                    return false;
                coerced = borders;
                return true;
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

    public static string HorizontalAlignName(int value) => value switch
    {
        XlHAlignGeneral => "general",
        XlHAlignLeft => "left",
        XlHAlignCenter => "center",
        XlHAlignRight => "right",
        XlHAlignFill => "fill",
        XlHAlignJustify => "justify",
        XlHAlignCenterAcross => "centerAcross",
        XlHAlignDistributed => "distributed",
        _ => $"unknown({value})",
    };

    public static string VerticalAlignName(int value) => value switch
    {
        XlVAlignTop => "top",
        XlVAlignCenter => "center",
        XlVAlignBottom => "bottom",
        XlVAlignJustify => "justify",
        XlVAlignDistributed => "distributed",
        _ => $"unknown({value})",
    };

    public static int HorizontalAlignValue(string name) => HorizontalAlignValues[name];
    public static int VerticalAlignValue(string name) => VerticalAlignValues[name];
    public static int UnderlineValue(string name) => UnderlineValues[name];
    public static int FillPatternValue(string name) => FillPatternValues[name];

    public static string UnderlineName(int value) => value switch
    {
        XlUnderlineNone => "none",
        XlUnderlineSingle => "single",
        XlUnderlineDouble => "double",
        XlUnderlineSingleAccounting => "singleAccounting",
        XlUnderlineDoubleAccounting => "doubleAccounting",
        _ => $"unknown({value})",
    };

    public static string FillPatternName(int value) =>
        FillPatternValues.FirstOrDefault(pair => pair.Value == value).Key ?? $"pattern({value})";

    public static int OrientationValue(JsonNode node)
    {
        if (node is JsonValue text && text.TryGetValue<string>(out var name) &&
            name.Equals("stacked", StringComparison.OrdinalIgnoreCase))
            return XlOrientationStacked;
        return node!.GetValue<int>();
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
