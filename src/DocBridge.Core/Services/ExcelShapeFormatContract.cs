using System.Globalization;
using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// Public contract and readback comparison for the practical formatting fields
/// shared by insert/update shape and textbox operations.
/// </summary>
public static class ExcelShapeFormatContract
{
    public const double MaxLineWeight = 20;
    public const double MaxFontSize = 409.5;
    private const double RotationEpsilon = 0.0001;

    private static readonly HashSet<string> FontKeys = new(StringComparer.Ordinal)
    {
        "name", "size", "bold", "italic", "color",
    };

    public static void Validate(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        ValidateColor(op, index, opName, "lineColor", errors);
        ValidateFiniteRange(op, index, opName, "lineWeight", 0, MaxLineWeight, exclusiveMinimum: true,
            "a finite point value > 0 and <= 20", errors);
        ValidateOptionalBool(op, index, opName, "lineVisible", errors);
        ValidateRotation(op, index, opName, errors);
        ValidateFont(op, index, opName, errors);
    }

    public static bool HasFormatting(JsonObject op) =>
        op.ContainsKey("lineColor") || op.ContainsKey("lineWeight") || op.ContainsKey("lineVisible") ||
        op.ContainsKey("rotation") || op.ContainsKey("font");

    public static bool TryNormalizeRotation(JsonNode? node, out double rotation)
    {
        rotation = 0;
        if (!ExcelDataOperationsContract.TryGetFiniteNumber(node, out var requested) ||
            requested < 0 || requested > 360 ||
            Math.Abs(requested - Math.Round(requested, MidpointRounding.AwayFromZero)) > RotationEpsilon)
            return false;
        rotation = NormalizeReadbackRotation(requested);
        return true;
    }

    public static double NormalizeReadbackRotation(double rotation)
    {
        if (!double.IsFinite(rotation)) return double.NaN;
        var normalized = Math.Round(rotation, MidpointRounding.AwayFromZero) % 360;
        if (normalized < 0) normalized += 360;
        return Math.Abs(normalized) < RotationEpsilon || Math.Abs(normalized - 360) < RotationEpsilon
            ? 0
            : normalized;
    }

    public static void CompareRequestedReadback(JsonObject actual, JsonObject op, string label, List<string> mismatches)
    {
        CompareOptionalNumber(actual, op, label, "lineWeight", tolerance: 0.01, mismatches);
        CompareOptionalBool(actual, op, label, "lineVisible", mismatches);
        if (op.ContainsKey("lineColor"))
            CompareOptionalColor(actual, op, label, "lineColor", mismatches);

        if (op.ContainsKey("rotation"))
        {
            if (!TryNormalizeRotation(op["rotation"], out var expected) ||
                !ExcelDataOperationsContract.TryGetFiniteNumber(actual["rotation"], out var observed))
            {
                mismatches.Add($"{label}: rotation readback unavailable");
            }
            else if (CircularDifference(expected, NormalizeReadbackRotation(observed)) > RotationEpsilon)
            {
                mismatches.Add($"{label}: rotation readback {observed.ToString("R", CultureInfo.InvariantCulture)} != {expected.ToString("R", CultureInfo.InvariantCulture)}");
            }
        }

        if (Json.GetObj(op, "font") is not JsonObject requestedFont) return;
        var actualFont = Json.GetObj(actual, "font");
        foreach (var (field, expected) in requestedFont)
        {
            if (actualFont is null || !actualFont.TryGetPropertyValue(field, out var observed) || observed is null)
            {
                mismatches.Add($"{label}: font.{field} readback unavailable");
                continue;
            }
            if (!FontValueMatches(field, expected, observed))
                mismatches.Add($"{label}: font.{field} readback mismatch");
        }
    }

    private static void ValidateColor(JsonObject op, int index, string opName, string field, ICollection<string> errors)
    {
        if (!op.ContainsKey(field)) return;
        if (!ExcelStyleContract.TryParseColor(op[field], out _))
            errors.Add($"ops[{index}] '{opName}' {field} must be an OLE color 0..16777215 or '#RRGGBB'");
    }

    private static void ValidateFiniteRange(JsonObject op, int index, string opName, string field,
        double minimum, double maximum, bool exclusiveMinimum, string description, ICollection<string> errors)
    {
        if (!op.ContainsKey(field)) return;
        if (!ExcelDataOperationsContract.TryGetFiniteNumber(op[field], out var number) ||
            (exclusiveMinimum ? number <= minimum : number < minimum) || number > maximum)
            errors.Add($"ops[{index}] '{opName}' {field} must be {description}");
    }

    private static void ValidateOptionalBool(JsonObject op, int index, string opName, string field,
        ICollection<string> errors)
    {
        if (!op.ContainsKey(field)) return;
        if (op[field] is not JsonValue value || !value.TryGetValue<bool>(out _))
            errors.Add($"ops[{index}] '{opName}' {field} must be boolean");
    }

    private static void ValidateRotation(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        if (!op.ContainsKey("rotation")) return;
        if (!TryNormalizeRotation(op["rotation"], out _))
            errors.Add($"ops[{index}] '{opName}' rotation must be an integer degree in 0..360 (360 normalizes to 0)");
    }

    private static void ValidateFont(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        if (!op.TryGetPropertyValue("font", out var node)) return;
        if (node is null)
        {
            errors.Add($"ops[{index}] '{opName}' font must be an object with name|size|bold|italic|color");
            return;
        }
        if (node is not JsonObject font)
        {
            errors.Add($"ops[{index}] '{opName}' font must be an object with name|size|bold|italic|color");
            return;
        }
        if (font.Count == 0)
            errors.Add($"ops[{index}] '{opName}' font must contain at least one supported property");
        foreach (var (field, value) in font)
        {
            if (!FontKeys.Contains(field))
            {
                errors.Add($"ops[{index}] '{opName}' font.{field} is not supported; allowed: name, size, bold, italic, color");
                continue;
            }
            if (value is null)
            {
                errors.Add($"ops[{index}] '{opName}' font.{field} must not be null; omit it to preserve the existing property");
                continue;
            }
            switch (field)
            {
                case "name":
                    if (value is not JsonValue name || !name.TryGetValue<string>(out var text) || string.IsNullOrWhiteSpace(text))
                        errors.Add($"ops[{index}] '{opName}' font.name must be a non-empty string");
                    break;
                case "size":
                    if (!ExcelDataOperationsContract.TryGetFiniteNumber(value, out var size) || size <= 0 || size > MaxFontSize)
                        errors.Add($"ops[{index}] '{opName}' font.size must be a finite point value > 0 and <= {MaxFontSize.ToString(CultureInfo.InvariantCulture)}");
                    break;
                case "bold":
                case "italic":
                    if (value is not JsonValue flag || !flag.TryGetValue<bool>(out _))
                        errors.Add($"ops[{index}] '{opName}' font.{field} must be boolean");
                    break;
                case "color":
                    if (!ExcelStyleContract.TryParseColor(value, out _))
                        errors.Add($"ops[{index}] '{opName}' font.color must be an OLE color 0..16777215 or '#RRGGBB'");
                    break;
            }
        }
    }

    private static void CompareOptionalNumber(JsonObject actual, JsonObject op, string label, string field,
        double tolerance, List<string> mismatches)
    {
        if (!op.ContainsKey(field)) return;
        if (!ExcelDataOperationsContract.TryGetFiniteNumber(op[field], out var expected) ||
            !ExcelDataOperationsContract.TryGetFiniteNumber(actual[field], out var observed) ||
            Math.Abs(expected - observed) > tolerance)
            mismatches.Add($"{label}: {field} readback mismatch");
    }

    private static void CompareOptionalBool(JsonObject actual, JsonObject op, string label, string field,
        List<string> mismatches)
    {
        if (!op.ContainsKey(field)) return;
        if (op[field] is not JsonValue requested || !requested.TryGetValue<bool>(out var expected) ||
            actual[field] is not JsonValue observed || !observed.TryGetValue<bool>(out var actualValue) ||
            expected != actualValue)
            mismatches.Add($"{label}: {field} readback mismatch");
    }

    private static void CompareOptionalColor(JsonObject actual, JsonObject op, string label, string field,
        List<string> mismatches)
    {
        if (!op.ContainsKey(field)) return;
        if (!ExcelStyleContract.TryParseColor(op[field], out var expected) ||
            !ExcelDataOperationsContract.TryGetFiniteNumber(actual[field], out var observed) ||
            Math.Abs(expected - observed) >= 0.5)
            mismatches.Add($"{label}: {field} readback mismatch");
    }

    private static bool FontValueMatches(string field, JsonNode? expected, JsonNode observed) => field switch
    {
        "name" => string.Equals(expected?.ToString()?.Trim(), observed.ToString()?.Trim(), StringComparison.OrdinalIgnoreCase),
        "size" => ExcelDataOperationsContract.TryGetFiniteNumber(expected, out var expectedSize) &&
                  ExcelDataOperationsContract.TryGetFiniteNumber(observed, out var observedSize) &&
                  Math.Abs(expectedSize - observedSize) <= 0.01,
        "bold" or "italic" => expected is JsonValue requested && requested.TryGetValue<bool>(out var expectedBool) &&
                                observed is JsonValue actual && actual.TryGetValue<bool>(out var actualBool) &&
                                expectedBool == actualBool,
        "color" => ExcelStyleContract.TryParseColor(expected, out var expectedColor) &&
                   ExcelDataOperationsContract.TryGetFiniteNumber(observed, out var observedColor) &&
                   Math.Abs(expectedColor - observedColor) < 0.5,
        _ => false,
    };

    private static double CircularDifference(double left, double right)
    {
        var direct = Math.Abs(left - right);
        return Math.Min(direct, 360 - direct);
    }

    public static double NormalizeRequestedRotation(JsonNode node)
    {
        if (!TryNormalizeRotation(node, out var rotation))
            throw new ArgumentOutOfRangeException(nameof(node), "rotation must be an integer degree in 0..360");
        return rotation;
    }

    public static double ReadFiniteNumber(JsonNode node)
    {
        if (!ExcelDataOperationsContract.TryGetFiniteNumber(node, out var number))
            throw new ArgumentOutOfRangeException(nameof(node), "A finite number was required.");
        return number;
    }

    public static bool TryCreateNumericReadbackValue(string field, object? raw, out JsonNode? value)
    {
        value = null;
        if (raw is null or DBNull or bool) return false;
        try
        {
            var number = Convert.ToDouble(raw, CultureInfo.InvariantCulture);
            if (!double.IsFinite(number)) return false;
            switch (field)
            {
                case "rotation" when number >= 0 && number <= 360:
                    value = JsonValue.Create(NormalizeReadbackRotation(number));
                    return true;
                case "lineWeight" when number > 0:
                    value = JsonValue.Create(number);
                    return true;
                case "lineColor" when number >= 0 && number <= 16_777_215 && number == Math.Truncate(number):
                    value = JsonValue.Create((int)number);
                    return true;
                default:
                    return false;
            }
        }
        catch (Exception ex) when (ex is FormatException or InvalidCastException or OverflowException)
        {
            return false;
        }
    }

    public static bool TryCreateFontReadbackValue(string field, object? raw, out JsonNode? value)
    {
        value = null;
        if (raw is null || raw is DBNull) return false;
        try
        {
            switch (field)
            {
                case "name":
                    var name = Convert.ToString(raw, CultureInfo.InvariantCulture);
                    if (string.IsNullOrWhiteSpace(name)) return false;
                    value = JsonValue.Create(name);
                    return true;
                case "size":
                    var size = Convert.ToDouble(raw, CultureInfo.InvariantCulture);
                    if (!double.IsFinite(size) || size <= 0) return false;
                    value = JsonValue.Create(size);
                    return true;
                case "bold":
                case "italic":
                    var flag = Convert.ToInt32(raw, CultureInfo.InvariantCulture);
                    if (flag is not (0 or -1 or 1)) return false;
                    value = JsonValue.Create(flag != 0);
                    return true;
                case "color":
                    var color = Convert.ToInt32(raw, CultureInfo.InvariantCulture);
                    if (color is < 0 or > 16_777_215) return false;
                    value = JsonValue.Create(color);
                    return true;
                default:
                    return false;
            }
        }
        catch (Exception ex) when (ex is FormatException or InvalidCastException or OverflowException)
        {
            return false;
        }
    }

    public static bool TryNormalizeLineVisible(object? raw, out bool visible)
    {
        visible = false;
        if (raw is null || raw is DBNull) return false;
        try
        {
            var flag = Convert.ToInt32(raw, CultureInfo.InvariantCulture);
            if (flag is not (0 or -1 or 1)) return false;
            visible = flag != 0;
            return true;
        }
        catch (Exception ex) when (ex is FormatException or InvalidCastException or OverflowException)
        {
            return false;
        }
    }
}
