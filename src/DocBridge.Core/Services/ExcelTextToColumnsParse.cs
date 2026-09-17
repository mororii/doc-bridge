using System.Globalization;
using System.Text;
using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// Excel <c>Range.TextToColumns</c> delimited parse used by apply readback.
/// TextQualifier (default double quote) keeps delimiters inside quoted fields.
/// Without FieldInfo, omitted columns use xlGeneralFormat — leading zeros
/// become numbers. FieldInfo/fixed-width is a published limitation.
/// https://learn.microsoft.com/en-us/office/vba/api/excel.range.texttocolumns
/// </summary>
public static class ExcelTextToColumnsParse
{
    public const string QualifierDouble = "double";
    public const string QualifierSingle = "single";
    public const string QualifierNone = "none";

    public static string QualifierToken(JsonObject op)
    {
        var raw = Json.GetString(op, "textQualifier");
        if (string.IsNullOrWhiteSpace(raw)) return QualifierDouble;
        if (raw is "\"" or "double" or "DoubleQuote" or "xlTextQualifierDoubleQuote")
            return QualifierDouble;
        if (raw is "'" or "single" or "SingleQuote" or "xlTextQualifierSingleQuote")
            return QualifierSingle;
        if (raw.Equals("none", StringComparison.OrdinalIgnoreCase) ||
            raw.Equals("false", StringComparison.OrdinalIgnoreCase) ||
            raw.Equals("xlTextQualifierNone", StringComparison.OrdinalIgnoreCase))
            return QualifierNone;
        return QualifierDouble;
    }

    public static int ComTextQualifier(JsonObject op) => QualifierToken(op) switch
    {
        QualifierSingle => ExcelDataObjectCatalog.XlTextQualifierSingleQuote,
        QualifierNone => ExcelDataObjectCatalog.XlTextQualifierNone,
        _ => ExcelDataObjectCatalog.XlTextQualifierDoubleQuote,
    };

    public static char? QualifierChar(JsonObject op) => QualifierToken(op) switch
    {
        QualifierSingle => '\'',
        QualifierNone => null,
        _ => '"',
    };

    public static string[] SplitLine(string text, string[] delimiters, char? qualifier, bool consecutive)
    {
        ArgumentNullException.ThrowIfNull(text);
        ArgumentNullException.ThrowIfNull(delimiters);
        var fields = new List<string>();
        var current = new StringBuilder();
        var inQuotes = false;
        for (var i = 0; i < text.Length; i++)
        {
            var c = text[i];
            if (qualifier is char q && c == q)
            {
                if (inQuotes && i + 1 < text.Length && text[i + 1] == q)
                {
                    current.Append(q);
                    i++;
                    continue;
                }
                inQuotes = !inQuotes;
                continue;
            }

            if (!inQuotes && StartsWithDelimiter(text, i, delimiters, out var length))
            {
                fields.Add(current.ToString());
                current.Clear();
                i += length - 1;
                if (consecutive)
                {
                    while (i + 1 < text.Length && StartsWithDelimiter(text, i + 1, delimiters, out var more))
                        i += more;
                }
                continue;
            }

            current.Append(c);
        }

        fields.Add(current.ToString());
        return fields.ToArray();
    }

    /// <summary>
    /// xlGeneralFormat: numeric-looking fields lose leading zeros. Text that is
    /// not a finite number stays as-is. Dates/locales beyond invariant numbers
    /// are a published limitation (no FieldInfo).
    /// </summary>
    public static string GeneralField(string field)
    {
        var trimmed = field.Trim();
        if (trimmed.Length == 0) return field;
        if (ExcelDataOperationsContract.TryGetFiniteNumber(JsonValue.Create(trimmed), out var number) ||
            double.TryParse(trimmed, NumberStyles.Float, CultureInfo.InvariantCulture, out number))
            return number.ToString(CultureInfo.InvariantCulture);
        return field;
    }

    public static string[] ExpectedFields(string text, JsonObject op)
    {
        var delimiters = ExcelDataReadbackContract.TextToColumnsDelimiters(op);
        var consecutive = Json.GetBool(op, "consecutiveDelimiter");
        var raw = SplitLine(text, delimiters, QualifierChar(op), consecutive);
        var general = !UsesTextFieldInfo(op);
        if (!general) return raw;
        return raw.Select(GeneralField).ToArray();
    }

    /// <summary>
    /// FieldInfo is not published. Always General unless a future explicit
    /// text-column list is added.
    /// </summary>
    public static bool UsesTextFieldInfo(JsonObject op) =>
        string.Equals(Json.GetString(op, "columnDataType"), "text", StringComparison.OrdinalIgnoreCase);

    private static bool StartsWithDelimiter(string text, int index, string[] delimiters, out int length)
    {
        length = 0;
        foreach (var delimiter in delimiters)
        {
            if (delimiter.Length == 0) continue;
            if (index + delimiter.Length <= text.Length &&
                string.CompareOrdinal(text, index, delimiter, 0, delimiter.Length) == 0)
            {
                length = delimiter.Length;
                return true;
            }
        }
        return false;
    }
}
