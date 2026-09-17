using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

namespace DocBridge.Core.Services;

/// <summary>
/// Exact workbook/sheet/area or table identity. A1, $A$1, and absolute R1C1
/// compare as the same box; longer ranges and other sheets do not.
/// </summary>
public static class ExcelReferenceIdentity
{
    private static readonly Regex AbsoluteR1C1 = new(
        @"^R(\d+)C(\d+)(?::R(\d+)C(\d+))?$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant | RegexOptions.IgnoreCase);

    internal readonly record struct SourceArea(string? Workbook, ExcelA1Box Box);

    internal readonly record struct Source(
        bool IsTable,
        string? TableName,
        string? Workbook,
        string? Sheet,
        IReadOnlyList<SourceArea> Areas);

    public static bool SameSource(
        string? requested, string? actual, string? defaultSheet = null, string? ownerWorkbook = null)
    {
        if (string.IsNullOrWhiteSpace(requested)) return true;
        if (string.IsNullOrWhiteSpace(actual)) return false;
        if (!TryParseSource(requested, defaultSheet, out var want)) return false;
        if (!TryParseSource(actual, defaultSheet, out var got)) return false;
        return SameSource(want, got, ownerWorkbook);
    }

    internal static bool SameSource(Source requested, Source actual, string? ownerWorkbook = null)
    {
        if (requested.IsTable != actual.IsTable) return false;
        // Resolve missing book from owner first. Do not compare raw null vs [owned.xlsx].
        if (!WorkbookEquals(
                ResolveWorkbook(requested.Workbook, ownerWorkbook),
                ResolveWorkbook(actual.Workbook, ownerWorkbook)))
            return false;
        if (requested.IsTable)
        {
            return string.Equals(requested.TableName, actual.TableName, StringComparison.OrdinalIgnoreCase);
        }

        if (!SheetEquals(requested.Sheet, actual.Sheet)) return false;
        if (requested.Areas.Count == 0 || requested.Areas.Count != actual.Areas.Count)
            return false;
        var remaining = actual.Areas.ToList();
        foreach (var area in requested.Areas)
        {
            var wantBook = ResolveWorkbook(area.Workbook ?? requested.Workbook, ownerWorkbook);
            var index = remaining.FindIndex(item =>
                item.Box == area.Box &&
                WorkbookEquals(wantBook, ResolveWorkbook(item.Workbook ?? actual.Workbook, ownerWorkbook)));
            if (index < 0) return false;
            remaining.RemoveAt(index);
        }
        return remaining.Count == 0;
    }

    public static bool FormulaContainsReference(
        string? formula, string? range, string? defaultSheet, int slot, string? ownerWorkbook = null)
    {
        if (string.IsNullOrWhiteSpace(formula) || string.IsNullOrWhiteSpace(range))
            return false;
        if (!TryReadSeriesArguments(formula, out var args) || slot < 0 || slot >= args.Count)
            return false;
        var arg = args[slot];
        if (IsQuotedString(arg)) return false;
        var token = Unquote(arg);
        if (string.IsNullOrWhiteSpace(token)) return false;
        return SameSource(range, token, defaultSheet, ownerWorkbook);
    }

    public static bool FormulaContainsSourceArea(
        string? formula, string? range, string? defaultSheet = null, string? ownerWorkbook = null) =>
        FormulaContainsReference(formula, range, defaultSheet, slot: 1, ownerWorkbook) ||
        FormulaContainsReference(formula, range, defaultSheet, slot: 2, ownerWorkbook);

    public static bool SeriesNameEquals(string? formula, string? name)
    {
        if (string.IsNullOrWhiteSpace(name)) return true;
        if (!TryReadSeriesArguments(formula, out var args) || args.Count == 0)
            return false;
        var arg = args[0];
        if (IsQuotedString(arg))
            return string.Equals(Unquote(arg), name, StringComparison.Ordinal);
        if (IsExplicitNameReference(arg))
            return false;
        return string.Equals(Unquote(arg), name, StringComparison.Ordinal);
    }

    public static bool SeriesNameReferenceMatches(
        string? formula, string? name, string? defaultSheet = null, string? ownerWorkbook = null)
    {
        if (string.IsNullOrWhiteSpace(name) || !IsExplicitNameReference(name))
            return false;
        if (!TryReadSeriesArguments(formula, out var args) || args.Count == 0)
            return false;
        var arg = args[0];
        if (IsQuotedString(arg)) return false;
        return SameSource(StripEquals(name.Trim()), Unquote(arg), defaultSheet, ownerWorkbook);
    }

    /// <summary>
    /// Microsoft Series.Name is a literal String unless the caller writes an
    /// explicit formula such as <c>=Sheet1!R1C1</c>. A trailing <c>!</c> is not a range.
    /// </summary>
    public static bool IsExplicitNameReference(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;
        return text.TrimStart().StartsWith('=');
    }

    internal static bool TryParseSource(string? text, string? defaultSheet, out Source source)
    {
        source = default;
        if (string.IsNullOrWhiteSpace(text)) return false;
        var raw = StripEquals(text.Trim());
        if (LooksLikeTable(raw))
        {
            source = new Source(true, raw, null, null, []);
            return true;
        }

        if (!TrySplitQualifiedAreas(raw, defaultSheet, out var sheet, out var areas) || areas.Count == 0)
            return false;
        source = new Source(false, null, CommonWorkbook(areas), sheet, areas);
        return true;
    }

    public static IReadOnlyList<string> EnumerateFormulaReferences(string formula)
    {
        var refs = new List<string>();
        if (TryReadSeriesArguments(formula, out var args))
        {
            for (var slot = 0; slot < args.Count && slot <= 2; slot++)
            {
                var arg = args[slot];
                if (IsQuotedString(arg)) continue;
                var token = Unquote(arg);
                if (string.IsNullOrWhiteSpace(token)) continue;
                if (LooksLikeReferenceToken(token))
                    refs.Add(token);
            }
            return refs;
        }

        CollectBareReferences(formula, refs);
        return refs;
    }

    public static bool TryReadSeriesArguments(string? formula, out List<string> args)
    {
        args = [];
        if (string.IsNullOrWhiteSpace(formula)) return false;
        var text = StripEquals(formula.Trim());
        if (!text.StartsWith("SERIES(", StringComparison.OrdinalIgnoreCase) || !text.EndsWith(')'))
            return false;
        return TrySplitTopLevel(text[7..^1], args);
    }

    private static bool TrySplitQualifiedAreas(string text, string? defaultSheet,
        out string? sheet, out List<SourceArea> areas)
    {
        sheet = null;
        areas = [];
        var pieces = SplitTopLevelAreas(text);
        if (pieces.Count == 0) return false;
        string? inheritedBook = null;
        foreach (var piece in pieces)
        {
            if (!TryParseOneArea(piece, defaultSheet, out var pieceBook, out var pieceSheet, out var box))
                return false;
            if (sheet is null)
                sheet = pieceSheet;
            else if (!SheetEquals(sheet, pieceSheet))
                return false;
            if (pieceBook is not null)
                inheritedBook = pieceBook;
            areas.Add(new SourceArea(pieceBook ?? inheritedBook, box));
        }

        return true;
    }

    private static bool TryParseOneArea(string text, string? defaultSheet,
        out string? workbook, out string? sheet, out ExcelA1Box box)
    {
        workbook = null;
        sheet = null;
        box = default;
        var raw = text.Trim();
        if (TryReadWorkbook(raw, out workbook, out var withoutBook))
            raw = withoutBook;
        try
        {
            var parsed = ExcelRangeReference.Parse(raw);
            sheet = string.IsNullOrWhiteSpace(parsed.SheetName) ? EmptyToNull(defaultSheet) : parsed.SheetName;
            raw = parsed.Address;
        }
        catch (FormatException)
        {
            sheet = EmptyToNull(defaultSheet);
        }

        if (raw.Contains(',', StringComparison.Ordinal) || raw.Contains(';', StringComparison.Ordinal))
            return false;
        return TryParseBox(raw, out box);
    }

    private static string? CommonWorkbook(IReadOnlyList<SourceArea> areas)
    {
        string? book = null;
        foreach (var area in areas)
        {
            if (string.IsNullOrWhiteSpace(area.Workbook)) continue;
            if (book is null)
            {
                book = area.Workbook;
                continue;
            }
            if (!WorkbookEquals(book, area.Workbook))
                return null;
        }
        return book;
    }

    private static bool TryParseBox(string address, out ExcelA1Box box)
    {
        var compact = address.Replace("$", "", StringComparison.Ordinal).Trim();
        if (ExcelA1Box.TryParse(compact, out box))
            return true;
        var match = AbsoluteR1C1.Match(compact);
        if (!match.Success) return false;
        var r1 = int.Parse(match.Groups[1].Value, CultureInfo.InvariantCulture);
        var c1 = int.Parse(match.Groups[2].Value, CultureInfo.InvariantCulture);
        var r2 = match.Groups[3].Success ? int.Parse(match.Groups[3].Value, CultureInfo.InvariantCulture) : r1;
        var c2 = match.Groups[4].Success ? int.Parse(match.Groups[4].Value, CultureInfo.InvariantCulture) : c1;
        if (r1 < 1 || c1 < 1 || r2 < r1 || c2 < c1) return false;
        if (r2 > ExcelA1Box.MaxRow || c2 > ExcelA1Box.MaxColumn) return false;
        box = new ExcelA1Box(r1, c1, r2 - r1 + 1, c2 - c1 + 1);
        return true;
    }

    private static bool LooksLikeTable(string text)
    {
        if (text.Contains('!', StringComparison.Ordinal) ||
            text.Contains('[', StringComparison.Ordinal) ||
            text.Contains(']', StringComparison.Ordinal))
            return false;
        if (ExcelDataOperationsContract.TryParseA1(text, out _, allowUnion: false))
            return false;
        if (AbsoluteR1C1.IsMatch(text.Replace("$", "", StringComparison.Ordinal)))
            return false;
        return text.IndexOfAny([' ', '\t', '\r', '\n', ',', ';']) < 0 && text.Length is >= 1 and <= 255;
    }

    private static bool LooksLikeReferenceToken(string token)
    {
        if (token.Contains('!', StringComparison.Ordinal) || token.Contains('$', StringComparison.Ordinal))
            return true;
        return TryParseSource(token, null, out var source) && !source.IsTable;
    }

    /// <summary>
    /// Excel external refs are one quoted unit: <c>'[owned.xlsx]원 장'!R1C1</c>,
    /// including doubled apostrophes. Do not strip only the leading quote.
    /// </summary>
    private static bool TryReadWorkbook(string text, out string? workbook, out string rest)
    {
        workbook = null;
        rest = text;
        var raw = text.Trim();
        if (raw.Length == 0) return false;

        if (raw.StartsWith('\''))
        {
            if (!TrySplitQuotedBookSheet(raw, out workbook, out var sheet, out var address))
                return false;
            rest = ExcelFormulaReference.ForceQuoteSheetName(sheet) + "!" + address;
            return true;
        }

        if (!raw.StartsWith('[')) return false;
        var close = raw.IndexOf(']');
        if (close <= 1) return false;
        workbook = raw[1..close];
        rest = raw[(close + 1)..];
        return workbook.Length > 0;
    }

    private static bool TrySplitQuotedBookSheet(
        string raw, out string? workbook, out string sheet, out string address)
    {
        workbook = null;
        sheet = "";
        address = "";
        if (!TryReadQuotedExternal(raw, out var body, out address))
            return false;
        if (!body.StartsWith('[')) return false;
        var close = body.IndexOf(']');
        if (close <= 1) return false;
        workbook = body[1..close];
        sheet = body[(close + 1)..];
        return workbook.Length > 0 && sheet.Length > 0;
    }

    private static bool TryReadQuotedExternal(string raw, out string body, out string afterBang)
    {
        body = "";
        afterBang = "";
        if (raw.Length < 4 || raw[0] != '\'') return false;
        var current = new StringBuilder();
        for (var i = 1; i < raw.Length; i++)
        {
            if (raw[i] != '\'')
            {
                current.Append(raw[i]);
                continue;
            }

            if (i + 1 < raw.Length && raw[i + 1] == '\'')
            {
                current.Append('\'');
                i++;
                continue;
            }

            if (i + 1 < raw.Length && raw[i + 1] == '!')
            {
                body = current.ToString();
                afterBang = raw[(i + 2)..];
                return body.Length > 0;
            }

            return false;
        }

        return false;
    }

    private static List<string> SplitTopLevelAreas(string text)
    {
        var parts = new List<string>();
        TrySplitTopLevel(text, parts, [',', ';']);
        if (parts.Count == 0 && !string.IsNullOrWhiteSpace(text))
            parts.Add(text.Trim());
        return parts;
    }

    private static bool TrySplitTopLevel(string text, List<string> parts, char[]? separators = null)
    {
        separators ??= [','];
        var current = new StringBuilder();
        var depth = 0;
        var inQuote = false;
        for (var i = 0; i < text.Length; i++)
        {
            var ch = text[i];
            if (ch == '"')
            {
                current.Append(ch);
                if (inQuote && i + 1 < text.Length && text[i + 1] == '"')
                {
                    current.Append('"');
                    i++;
                    continue;
                }
                inQuote = !inQuote;
                continue;
            }

            if (!inQuote && ch == '\'')
            {
                current.Append(ch);
                i++;
                while (i < text.Length)
                {
                    current.Append(text[i]);
                    if (text[i] == '\'' && i + 1 < text.Length && text[i + 1] == '\'')
                    {
                        current.Append('\'');
                        i++;
                    }
                    else if (text[i] == '\'')
                        break;
                    i++;
                }
                continue;
            }

            if (!inQuote && ch == '(')
            {
                depth++;
                current.Append(ch);
                continue;
            }

            if (!inQuote && ch == ')')
            {
                depth--;
                current.Append(ch);
                continue;
            }

            if (!inQuote && depth == 0 && separators.Contains(ch))
            {
                parts.Add(current.ToString().Trim());
                current.Clear();
                continue;
            }

            current.Append(ch);
        }

        if (inQuote || depth != 0) return false;
        parts.Add(current.ToString().Trim());
        return true;
    }

    private static void CollectBareReferences(string formula, List<string> refs)
    {
        var text = formula;
        var index = 0;
        while (index < text.Length)
        {
            if (text[index] == '"')
            {
                index++;
                while (index < text.Length)
                {
                    if (text[index] == '"' && index + 1 < text.Length && text[index + 1] == '"')
                    {
                        index += 2;
                        continue;
                    }
                    if (text[index] == '"')
                    {
                        index++;
                        break;
                    }
                    index++;
                }
                continue;
            }

            if (TryReadQuotedSheetPrefix(text, index, out var prefixed, out var after))
            {
                refs.Add(prefixed);
                index = after;
                continue;
            }

            index++;
        }
    }

    private static bool TryReadQuotedSheetPrefix(string text, int start, out string value, out int after)
    {
        value = "";
        after = start;
        if (start >= text.Length || text[start] != '\'') return false;
        var bang = text.IndexOf("'!", start, StringComparison.Ordinal);
        if (bang < 0) return false;
        var end = bang + 2;
        while (end < text.Length && (char.IsLetterOrDigit(text[end]) || text[end] is '$' or ':' or '_'))
            end++;
        value = text[start..end];
        after = end;
        return value.Contains('!', StringComparison.Ordinal);
    }

    private static string StripEquals(string text) =>
        text.StartsWith('=') ? text[1..] : text;

    private static bool IsQuotedString(string text)
    {
        var value = text.Trim();
        return value.Length >= 2 && value[0] == '"' && value[^1] == '"';
    }

    private static string Unquote(string text)
    {
        var value = text.Trim();
        if (value.Length >= 2 && value[0] == '"' && value[^1] == '"')
            return value[1..^1].Replace("\"\"", "\"", StringComparison.Ordinal);
        return value;
    }

    private static string? ResolveWorkbook(string? qualified, string? ownerWorkbook) =>
        EmptyToNull(qualified) ?? EmptyToNull(ownerWorkbook);

    private static bool WorkbookEquals(string? left, string? right)
    {
        if (string.IsNullOrWhiteSpace(left) && string.IsNullOrWhiteSpace(right)) return true;
        if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right)) return false;
        if (string.Equals(left, right, StringComparison.OrdinalIgnoreCase)) return true;
        var leftFile = WorkbookFileName(left);
        var rightFile = WorkbookFileName(right);
        if (string.Equals(leftFile, rightFile, StringComparison.OrdinalIgnoreCase)) return true;
        return SameWorkbookStem(leftFile, rightFile);
    }

    private static string WorkbookFileName(string path)
    {
        var trimmed = path.Trim();
        var slash = Math.Max(trimmed.LastIndexOf('\\'), trimmed.LastIndexOf('/'));
        return slash >= 0 ? trimmed[(slash + 1)..] : trimmed;
    }

    private static bool SameWorkbookStem(string leftFile, string rightFile)
    {
        SplitWorkbookExt(leftFile, out var leftStem, out var leftExt);
        SplitWorkbookExt(rightFile, out var rightStem, out var rightExt);
        if (leftStem.Length == 0 ||
            !string.Equals(leftStem, rightStem, StringComparison.OrdinalIgnoreCase))
            return false;
        if (leftExt.Length == 0) return IsExcelWorkbookExt(rightExt);
        if (rightExt.Length == 0) return IsExcelWorkbookExt(leftExt);
        return false;
    }

    private static void SplitWorkbookExt(string fileName, out string stem, out string ext)
    {
        var dot = fileName.LastIndexOf('.');
        if (dot <= 0)
        {
            stem = fileName;
            ext = "";
            return;
        }

        stem = fileName[..dot];
        ext = fileName[dot..];
    }

    private static bool IsExcelWorkbookExt(string ext) =>
        ext.Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
        ext.Equals(".xlsm", StringComparison.OrdinalIgnoreCase) ||
        ext.Equals(".xlsb", StringComparison.OrdinalIgnoreCase) ||
        ext.Equals(".xls", StringComparison.OrdinalIgnoreCase) ||
        ext.Equals(".xltx", StringComparison.OrdinalIgnoreCase) ||
        ext.Equals(".xltm", StringComparison.OrdinalIgnoreCase);

    private static bool SheetEquals(string? left, string? right)
    {
        if (string.IsNullOrWhiteSpace(left) && string.IsNullOrWhiteSpace(right)) return true;
        if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right)) return false;
        return string.Equals(left, right, StringComparison.OrdinalIgnoreCase);
    }

    private static string? EmptyToNull(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;
}
