using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace DocBridge.Core.Services;

/// <summary>
/// Compare-only Excel formula / defined-name contract. Apply still writes
/// the caller's RefersTo exactly. Quoted vs unquoted sheet identifiers may
/// match; A1 vs $A$1 must not. String literals keep exact case and quotes.
/// Relative names are selection-dependent
/// (https://learn.microsoft.com/en-us/office/vba/api/excel.names).
/// </summary>
public static class ExcelFormulaReference
{
    private static readonly Regex A1Token = new(
        @"\$?[A-Za-z]{1,3}\$?[0-9]{1,7}",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    public static bool SemanticEquals(string? left, string? right)
    {
        if (string.IsNullOrWhiteSpace(left) && string.IsNullOrWhiteSpace(right))
            return true;
        if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right))
            return false;
        if (string.Equals(left, right, StringComparison.Ordinal))
            return true;
        return string.Equals(Canonicalize(left), Canonicalize(right), StringComparison.Ordinal);
    }

    /// <summary>
    /// Compare key only. Does not strip <c>$</c>. Letters and sheet quoting
    /// outside literals are canonicalized; literal text is copied as-is.
    /// </summary>
    public static string Normalize(string formula) => Canonicalize(formula);

    public static string Canonicalize(string formula)
    {
        var text = formula.Trim();
        var output = new StringBuilder(text.Length);
        var index = 0;
        while (index < text.Length)
        {
            var ch = text[index];
            if (ch == '"')
            {
                CopyStringLiteral(text, ref index, output);
                continue;
            }

            if (ch == '\'' && TryReadQuotedSheet(text, index, out var quoted, out var afterQuoted))
            {
                EmitSheet(output, quoted);
                index = afterQuoted;
                continue;
            }

            if (TryReadUnquotedSheet(text, index, out var unquoted, out var bang))
            {
                EmitSheet(output, unquoted);
                index = bang + 1;
                continue;
            }

            output.Append(char.IsLetter(ch) ? char.ToLowerInvariant(ch) : ch);
            index++;
        }

        return output.ToString();
    }

    public static bool HasRelativeCellRef(string? formula)
    {
        if (string.IsNullOrWhiteSpace(formula)) return false;
        var text = formula.Trim();
        var index = 0;
        while (index < text.Length)
        {
            if (text[index] == '"')
            {
                SkipStringLiteral(text, ref index);
                continue;
            }

            var match = A1Token.Match(text, index);
            if (!match.Success || match.Index != index)
            {
                index++;
                continue;
            }

            if (!IsFullyAbsoluteA1(match.Value))
                return true;
            index = match.Index + match.Length;
        }

        return false;
    }

    public static bool IsFullyAbsoluteA1(string token)
    {
        var match = A1Token.Match(token);
        if (!match.Success || match.Value.Length != token.Length) return false;
        var i = 0;
        if (token[i] != '$') return false;
        i++;
        while (i < token.Length && char.IsLetter(token[i])) i++;
        if (i >= token.Length || token[i] != '$') return false;
        i++;
        return i < token.Length && char.IsDigit(token[i]);
    }

    public static string ForceQuoteSheetName(string sheetName) =>
        "'" + sheetName.Replace("'", "''", StringComparison.Ordinal) + "'";

    public static string AbsoluteQuotedRefersTo(string sheetName, string a1Cell)
    {
        if (!ExcelA1Box.TryParseCell(a1Cell, out var row, out var column))
            throw new FormatException($"A1 cell is required: {a1Cell}");
        return $"={ForceQuoteSheetName(sheetName)}!${ExcelA1Box.ColumnName(column)}${row.ToString(CultureInfo.InvariantCulture)}";
    }

    public static bool TryParseDefinedName(string? actual, out string scope, out string? sheet, out string local)
    {
        scope = "workbook";
        sheet = null;
        local = actual?.Trim() ?? "";
        if (string.IsNullOrWhiteSpace(local)) return false;

        if (local[0] == '\'' && TryReadQuotedSheet(local, 0, out var quoted, out var after) &&
            after <= local.Length)
        {
            scope = "sheet";
            sheet = quoted;
            local = local[after..];
            return local.Length > 0;
        }

        var bang = local.IndexOf('!');
        if (bang <= 0) return true;
        scope = "sheet";
        sheet = local[..bang];
        local = local[(bang + 1)..];
        return local.Length > 0 && !string.IsNullOrWhiteSpace(sheet);
    }

    public static bool NeedsSheetQuotes(string sheetName)
    {
        if (string.IsNullOrEmpty(sheetName)) return true;
        if (char.IsAsciiDigit(sheetName[0])) return true;
        foreach (var ch in sheetName)
        {
            if (char.IsLetter(ch) || char.IsDigit(ch) || ch is '_' or '.' or '\\')
                continue;
            return true;
        }
        return false;
    }

    public static string QuoteSheetName(string sheetName) =>
        NeedsSheetQuotes(sheetName)
            ? ForceQuoteSheetName(sheetName)
            : sheetName;

    /// <summary>
    /// True when a sheet token references <paramref name="sheetName"/>.
    /// String literals and <c>[workbook]Sheet!</c> external links do not count.
    /// </summary>
    public static bool ReferencesSheet(string? formula, string sheetName)
    {
        if (string.IsNullOrWhiteSpace(formula) || string.IsNullOrWhiteSpace(sheetName))
            return false;
        var text = formula;
        var index = 0;
        while (index < text.Length)
        {
            if (text[index] == '"')
            {
                SkipStringLiteral(text, ref index);
                continue;
            }

            if (text[index] == '[')
            {
                var close = text.IndexOf(']', index);
                if (close < 0)
                {
                    index++;
                    continue;
                }

                index = close + 1;
                if (index < text.Length && TryReadQuotedSheet(text, index, out _, out var afterQuotedLink))
                {
                    index = afterQuotedLink;
                    continue;
                }

                if (TryReadUnquotedSheet(text, index, out _, out var bang))
                {
                    index = bang + 1;
                    continue;
                }

                continue;
            }

            if (TryReadQuotedSheet(text, index, out var quoted, out var afterQuoted))
            {
                if (string.Equals(quoted, sheetName, StringComparison.OrdinalIgnoreCase))
                    return true;
                index = afterQuoted;
                continue;
            }

            if (TryReadUnquotedSheet(text, index, out var unquoted, out var bangIndex))
            {
                if (string.Equals(unquoted, sheetName, StringComparison.OrdinalIgnoreCase))
                    return true;
                index = bangIndex + 1;
                continue;
            }

            index++;
        }

        return false;
    }

    private static void CopyStringLiteral(string text, ref int index, StringBuilder output)
    {
        output.Append('"');
        index++;
        while (index < text.Length)
        {
            var ch = text[index++];
            output.Append(ch);
            if (ch != '"') continue;
            if (index < text.Length && text[index] == '"')
            {
                output.Append('"');
                index++;
                continue;
            }
            break;
        }
    }

    private static void SkipStringLiteral(string text, ref int index)
    {
        index++;
        while (index < text.Length)
        {
            var ch = text[index++];
            if (ch != '"') continue;
            if (index < text.Length && text[index] == '"')
            {
                index++;
                continue;
            }
            break;
        }
    }

    private static bool TryReadQuotedSheet(string text, int start, out string sheet, out int afterBang)
    {
        sheet = "";
        afterBang = start;
        if (start >= text.Length || text[start] != '\'') return false;
        var name = new StringBuilder();
        var index = start + 1;
        while (index < text.Length)
        {
            var ch = text[index++];
            if (ch == '\'')
            {
                if (index < text.Length && text[index] == '\'')
                {
                    name.Append('\'');
                    index++;
                    continue;
                }

                if (index < text.Length && text[index] == '!')
                {
                    sheet = name.ToString();
                    afterBang = index + 1;
                    return sheet.Length > 0;
                }
                return false;
            }
            name.Append(ch);
        }
        return false;
    }

    private static bool TryReadUnquotedSheet(string text, int start, out string sheet, out int bang)
    {
        sheet = "";
        bang = -1;
        if (start >= text.Length || !IsSheetNameStart(text[start])) return false;
        var index = start + 1;
        while (index < text.Length && IsSheetNamePart(text[index]))
            index++;
        if (index >= text.Length || text[index] != '!') return false;
        sheet = text[start..index];
        bang = index;
        return sheet.Length > 0;
    }

    private static bool IsSheetNameStart(char ch) =>
        ch is '_' or '\\' || char.IsLetter(ch);

    private static bool IsSheetNamePart(char ch) =>
        IsSheetNameStart(ch) || char.IsDigit(ch) || ch is '.';

    private static void EmitSheet(StringBuilder output, string sheet)
    {
        var folded = new StringBuilder(sheet.Length);
        foreach (var ch in sheet)
            folded.Append(char.IsLetter(ch) ? char.ToLowerInvariant(ch) : ch);
        output.Append(QuoteSheetName(folded.ToString()));
        output.Append('!');
    }
}
