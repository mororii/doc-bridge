using System.Globalization;
using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// set_values keeps JSON types. A JSON string stays text (ISO dates are not
/// serial 46275). JSON null is a blank. JSON "" may normalize to a blank;
/// that is not stored by leaving NumberFormat=@. "123" is not 123.
/// Text writes temporarily use @ then restore the original format.
/// Mixed NumberFormat is null on the run; capture per-cell or refuse before @.
/// Formulas belong on set_formulas.
/// </summary>
public static class ExcelValueWriteContract
{
    public const string TextNumberFormat = "@";
    public const string IsoDateFixture = "2026-09-10";
    public const double IsoDateFixtureSerial = 46275d;

    public enum JsonKind
    {
        Null,
        EmptyString,
        String,
        Number,
        Boolean,
    }

    public readonly record struct CellWrite(JsonKind Kind, object? ComValue, string? Text);

    public static JsonKind KindOf(JsonNode? node)
    {
        if (node is null)
            return JsonKind.Null;
        if (node is not JsonValue value)
            return JsonKind.String;
        if (value.GetValueKind() == System.Text.Json.JsonValueKind.Null)
            return JsonKind.Null;
        if (value.TryGetValue<string>(out var s))
            return s.Length == 0 ? JsonKind.EmptyString : JsonKind.String;
        if (value.TryGetValue<bool>(out _))
            return JsonKind.Boolean;
        return JsonKind.Number;
    }

    public static CellWrite Classify(JsonNode? node)
    {
        if (node is null)
            return new CellWrite(JsonKind.Null, null, null);
        if (node is JsonValue value)
        {
            if (value.GetValueKind() == System.Text.Json.JsonValueKind.Null)
                return new CellWrite(JsonKind.Null, null, null);
            if (value.TryGetValue<string>(out var text))
                return text.Length == 0
                    ? new CellWrite(JsonKind.EmptyString, null, null)
                    : new CellWrite(JsonKind.String, text, text);
            if (value.TryGetValue<bool>(out var flag))
                return new CellWrite(JsonKind.Boolean, flag, null);
        }

        return new CellWrite(JsonKind.Number, ExcelAdapterComNumber(node), null);
    }

    public static bool RequiresTextNumberFormat(JsonKind kind) =>
        kind is JsonKind.String;

    public static bool IsIsoDateText(string? text) =>
        DateTime.TryParseExact(text, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out _);

    public static bool IsTextNumberFormat(string? numberFormat) =>
        !string.IsNullOrWhiteSpace(numberFormat) &&
        numberFormat.Trim().Equals(TextNumberFormat, StringComparison.Ordinal);

    public static bool IsExcelBlank(object? value) =>
        value is null or DBNull;

    public static bool IsComNumber(object? value) =>
        value is byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal;

    public static bool TypedEqual(CellWrite written, object? actual)
    {
        switch (written.Kind)
        {
            case JsonKind.Null:
            case JsonKind.EmptyString:
                return IsExcelBlank(actual) || actual is string empty && empty.Length == 0;
            case JsonKind.String:
                return actual is string got &&
                       string.Equals(written.Text, got, StringComparison.Ordinal);
            case JsonKind.Number:
                if (actual is string || !IsComNumber(actual) || written.ComValue is null)
                    return false;
                return Math.Abs(
                    Convert.ToDouble(written.ComValue, CultureInfo.InvariantCulture) -
                    Convert.ToDouble(actual, CultureInfo.InvariantCulture)) < 1e-9;
            case JsonKind.Boolean:
                return written.ComValue is bool want && actual is bool gotFlag && want == gotFlag;
            default:
                return false;
        }
    }

    public static IReadOnlyList<(int Row, int StartColumn, int Length)> ContiguousTextRuns(
        CellWrite[,] writes, int rows, int cols)
    {
        var runs = new List<(int, int, int)>();
        for (var r = 0; r < rows; r++)
        {
            var c = 0;
            while (c < cols)
            {
                if (!RequiresTextNumberFormat(writes[r, c].Kind))
                {
                    c++;
                    continue;
                }

                var start = c;
                while (c < cols && RequiresTextNumberFormat(writes[r, c].Kind))
                    c++;
                runs.Add((r, start, c - start));
            }
        }

        return runs;
    }

    public static bool DisplayEqualIsNotTypedEqual(object? left, object? right)
    {
        var leftText = Convert.ToString(left, CultureInfo.InvariantCulture) ?? "";
        var rightText = Convert.ToString(right, CultureInfo.InvariantCulture) ?? "";
        return string.Equals(leftText, rightText, StringComparison.Ordinal);
    }

    private static object ExcelAdapterComNumber(JsonNode node)
    {
        if (node is not JsonValue value)
            return node.ToJsonString();
        if (value.TryGetValue<int>(out var int32)) return Convert.ToDouble(int32, CultureInfo.InvariantCulture);
        if (value.TryGetValue<long>(out var int64)) return Convert.ToDouble(int64, CultureInfo.InvariantCulture);
        if (value.TryGetValue<uint>(out var uint32)) return Convert.ToDouble(uint32, CultureInfo.InvariantCulture);
        if (value.TryGetValue<ulong>(out var uint64)) return Convert.ToDouble(uint64, CultureInfo.InvariantCulture);
        if (value.TryGetValue<decimal>(out var dec)) return Convert.ToDouble(dec, CultureInfo.InvariantCulture);
        if (value.TryGetValue<float>(out var single)) return Convert.ToDouble(single, CultureInfo.InvariantCulture);
        if (value.TryGetValue<double>(out var number)) return number;
        return Convert.ToDouble(value.ToString(), CultureInfo.InvariantCulture);
    }
}
