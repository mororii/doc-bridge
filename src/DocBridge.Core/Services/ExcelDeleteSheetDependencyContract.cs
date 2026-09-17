namespace DocBridge.Core.Services;

/// <summary>
/// delete_sheet copy-back does not repair other-sheet formulas, names, or
/// chart series that became #REF!. Capture those in-workbook dependencies
/// before delete and write them back after the sheet copy. Only an actual
/// external workbook or backup file token is excluded. Structured table
/// refs such as Items[Amount], same-workbook [owned.xlsx]Sheet!, and 3D
/// Sheet1:Sheet3! stay capturable.
/// </summary>
public static class ExcelDeleteSheetDependencyContract
{
    public const string CaptureKeyFormulas = "externalFormulas";
    public const string CaptureKeyNames = "names";
    public const string CaptureKeyCharts = "charts";
    public const string CaptureKeyPivots = "pivots";
    public const string UncapturedError = "[EXCEL_DELETE_SHEET_DEPENDENCY_UNCAPTURED]";

    private static readonly string[] WorkbookExtensions =
        [".xlsx", ".xlsm", ".xlsb", ".xls", ".xltx", ".xltm", ".xlw"];

    public static bool IsBrokenRef(string? formula) =>
        !string.IsNullOrWhiteSpace(formula) &&
        formula.Contains("#REF!", StringComparison.OrdinalIgnoreCase);

    public static bool LooksLikeWorkbookFile(string? token)
    {
        if (string.IsNullOrWhiteSpace(token)) return false;
        var file = FileNameOf(token);
        foreach (var ext in WorkbookExtensions)
        {
            if (file.EndsWith(ext, StringComparison.OrdinalIgnoreCase))
                return true;
        }

        return token.IndexOfAny(['\\', '/']) >= 0;
    }

    public static bool IsSameOwnerWorkbook(string? token, string? ownerWorkbook)
    {
        if (string.IsNullOrWhiteSpace(token) || string.IsNullOrWhiteSpace(ownerWorkbook))
            return false;
        return string.Equals(FileNameOf(token), FileNameOf(ownerWorkbook), StringComparison.OrdinalIgnoreCase);
    }

    public static bool IsBackupFileExternalLink(string? formula, string? ownerWorkbook = null)
    {
        if (string.IsNullOrWhiteSpace(formula))
            return false;
        foreach (var token in WorkbookTokens(formula))
        {
            if (IsSameOwnerWorkbook(token, ownerWorkbook))
                continue;
            return true;
        }

        return false;
    }

    public static void RequireFaithfulCapture(string? formula, string? ownerWorkbook)
    {
        if (!CanFaithfullyCapture(formula, ownerWorkbook, out var limitation))
            throw new InvalidOperationException(limitation);
    }

    public static bool IsUncaptured(Exception ex) =>
        ex is InvalidOperationException &&
        ex.Message.Contains(UncapturedError, StringComparison.Ordinal);

    public static bool CanFaithfullyCapture(string? formula, string? ownerWorkbook, out string? limitation)
    {
        limitation = null;
        if (string.IsNullOrWhiteSpace(formula))
            return true;
        var text = formula;
        var index = 0;
        while (index < text.Length)
        {
            if (text[index] == '"')
            {
                SkipQuoted(text, ref index);
                continue;
            }

            if (text[index] != '[')
            {
                index++;
                continue;
            }

            var close = text.IndexOf(']', index);
            if (close < 0)
            {
                limitation = $"{UncapturedError} unclosed '[' cannot be captured faithfully";
                return false;
            }

            if (!IsStructuredTableBracket(text, index) && !IsWorkbookBracket(text, index))
            {
                limitation =
                    $"{UncapturedError} bracketed token '{text[index..(close + 1)]}' is neither a table column nor a workbook link";
                return false;
            }

            index = close + 1;
        }

        _ = ownerWorkbook;
        return true;
    }

    public static bool FormulaReferencesDeletedSheet(
        string? formula, string sheetName, string? ownerWorkbook = null,
        IReadOnlyList<string>? sheetOrder = null)
    {
        if (string.IsNullOrWhiteSpace(formula) || string.IsNullOrWhiteSpace(sheetName))
            return false;
        if (QualifiedOr3DReferencesSheet(formula, sheetName, ownerWorkbook, sheetOrder))
            return true;
        return ExcelFormulaReference.ReferencesSheet(formula, sheetName);
    }

    public static bool NameReferencesDeletedSheet(
        string? refersTo, string sheetName, string? ownerWorkbook = null,
        IReadOnlyList<string>? sheetOrder = null) =>
        FormulaReferencesDeletedSheet(refersTo, sheetName, ownerWorkbook, sheetOrder);

    public static bool ChartFormulaReferencesDeletedSheet(
        string? formula, string sheetName, string? ownerWorkbook = null,
        IReadOnlyList<string>? sheetOrder = null) =>
        FormulaReferencesDeletedSheet(formula, sheetName, ownerWorkbook, sheetOrder);

    public static bool PivotSourceReferencesDeletedSheet(
        string? source, string sheetName, string? ownerWorkbook = null,
        IReadOnlyList<string>? sheetOrder = null) =>
        FormulaReferencesDeletedSheet(source, sheetName, ownerWorkbook, sheetOrder) ||
        FormulaReferencesDeletedSheet("=" + (source ?? ""), sheetName, ownerWorkbook, sheetOrder);

    public static bool IsNoSpecialCells(Exception ex)
    {
        var message = ex.Message ?? "";
        return message.Contains("no cells were found", StringComparison.OrdinalIgnoreCase) ||
               message.Contains("SpecialCells", StringComparison.OrdinalIgnoreCase) ||
               message.Contains("셀을 찾을 수 없습니다", StringComparison.Ordinal) ||
               message.Contains("해당하는 셀이 없습니다", StringComparison.Ordinal);
    }

    public static InvalidOperationException CaptureFailed(string detail) =>
        new($"{UncapturedError} {detail}");

    public static bool RestoredDependencyMatches(string? captured, string? actual) =>
        !IsBrokenRef(actual) &&
        (string.Equals(captured ?? "", actual ?? "", StringComparison.Ordinal) ||
         ExcelFormulaReference.SemanticEquals(captured, actual));

    public static IReadOnlyList<string> WorkbookTokens(string formula)
    {
        var tokens = new List<string>();
        var index = 0;
        while (index < formula.Length)
        {
            if (formula[index] == '"')
            {
                SkipQuoted(formula, ref index);
                continue;
            }

            if (formula[index] == '[')
            {
                var close = formula.IndexOf(']', index);
                if (close < 0)
                    break;
                if (IsWorkbookBracket(formula, index))
                    tokens.Add(formula[(index + 1)..close]);
                index = close + 1;
                continue;
            }

            index++;
        }

        return tokens;
    }

    private static bool QualifiedOr3DReferencesSheet(
        string formula, string sheetName, string? ownerWorkbook, IReadOnlyList<string>? sheetOrder)
    {
        var index = 0;
        while (index < formula.Length)
        {
            if (formula[index] == '"')
            {
                SkipQuoted(formula, ref index);
                continue;
            }

            if (formula[index] == '[')
                {
                    if (IsWorkbookBracket(formula, index) &&
                        TryConsumeWorkbookQualifiedSheet(
                            formula, index, sheetName, ownerWorkbook, sheetOrder, out var afterLink, out var hit))
                    {
                        if (hit)
                            return true;
                        index = afterLink;
                        continue;
                    }
                }

                if (formula[index] == '\'')
                {
                    if (TryReadQuotedSheetToken(formula, index, out var quoted, out var after))
                    {
                        if (SheetTokenReferences(quoted, sheetName, ownerWorkbook, sheetOrder))
                            return true;
                        index = after;
                        continue;
                    }
                }

                if (TryReadUnquoted3DOrSheet(formula, index, out var unquoted, out var bang))
                {
                    if (SheetTokenReferences(unquoted, sheetName, ownerWorkbook, sheetOrder))
                        return true;
                    index = bang + 1;
                    continue;
                }

                index++;
        }

        return false;
    }

    private static bool TryConsumeWorkbookQualifiedSheet(
        string formula, int open, string sheetName, string? ownerWorkbook,
        IReadOnlyList<string>? sheetOrder, out int after, out bool hit)
    {
        after = open;
        hit = false;
        var close = formula.IndexOf(']', open);
        if (close <= open)
            return false;
        var book = formula[(open + 1)..close];
        var next = close + 1;
        if (TryReadQuotedSheetToken(formula, next, out var quoted, out var afterQuoted))
        {
            hit = SheetTokenReferences($"[{book}]{quoted}", sheetName, ownerWorkbook, sheetOrder);
            after = afterQuoted;
            return true;
        }

        if (TryReadUnquoted3DOrSheet(formula, next, out var unquoted, out var bang))
        {
            hit = SheetTokenReferences($"[{book}]{unquoted}", sheetName, ownerWorkbook, sheetOrder);
            after = bang + 1;
            return true;
        }

        after = close + 1;
        return true;
    }

    private static bool SheetTokenReferences(
        string token, string sheetName, string? ownerWorkbook, IReadOnlyList<string>? sheetOrder)
    {
        var sheet = token;
        if (token.StartsWith('[') && token.Contains(']', StringComparison.Ordinal))
        {
            var close = token.IndexOf(']');
            var book = token[1..close];
            sheet = token[(close + 1)..];
            if (LooksLikeWorkbookFile(book) && !IsSameOwnerWorkbook(book, ownerWorkbook))
                return false;
        }

        return SheetOr3DNameEquals(sheet, sheetName, sheetOrder);
    }

    private static bool SheetOr3DNameEquals(string token, string sheetName, IReadOnlyList<string>? sheetOrder)
    {
        if (string.Equals(token, sheetName, StringComparison.OrdinalIgnoreCase))
            return true;
        var split = token.IndexOf(':');
        if (split <= 0 || split >= token.Length - 1)
            return false;
        var left = token[..split];
        var right = token[(split + 1)..];
        if (LooksLikeA1(left) || LooksLikeA1(right))
            return false;
        if (string.Equals(left, sheetName, StringComparison.OrdinalIgnoreCase) ||
            string.Equals(right, sheetName, StringComparison.OrdinalIgnoreCase))
            return true;
        if (sheetOrder is { Count: > 0 })
            return SheetOrderContains(sheetOrder, left, right, sheetName);
        return true;
    }

    private static bool SheetOrderContains(
        IReadOnlyList<string> order, string left, string right, string sheetName)
    {
        var start = IndexOfSheet(order, left);
        var end = IndexOfSheet(order, right);
        var mid = IndexOfSheet(order, sheetName);
        if (start < 0 || end < 0 || mid < 0)
            return true;
        if (start > end)
            (start, end) = (end, start);
        return mid >= start && mid <= end;
    }

    private static int IndexOfSheet(IReadOnlyList<string> order, string name)
    {
        for (var i = 0; i < order.Count; i++)
        {
            if (string.Equals(order[i], name, StringComparison.OrdinalIgnoreCase))
                return i;
        }

        return -1;
    }

    private static bool LooksLikeA1(string token) =>
        ExcelA1Box.TryParse(token.Replace("$", "", StringComparison.Ordinal), out _);

    private static bool TryReadQuotedSheetToken(string text, int start, out string token, out int afterBang)
    {
        token = "";
        afterBang = start;
        if (start >= text.Length || text[start] != '\'') return false;
        var name = new System.Text.StringBuilder();
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
                    token = name.ToString();
                    afterBang = index + 1;
                    return token.Length > 0;
                }

                return false;
            }

            name.Append(ch);
        }

        return false;
    }

    private static bool TryReadUnquoted3DOrSheet(string text, int start, out string token, out int bang)
    {
        token = "";
        bang = -1;
        if (start >= text.Length || !IsSheetNameStart(text[start])) return false;
        var index = start + 1;
        while (index < text.Length && (IsSheetNamePart(text[index]) || text[index] == ':'))
            index++;
        if (index >= text.Length || text[index] != '!') return false;
        token = text[start..index];
        bang = index;
        return token.Length > 0;
    }

    private static bool IsWorkbookBracket(string text, int open)
    {
        var close = text.IndexOf(']', open);
        if (close <= open + 1) return false;
        var inner = text[(open + 1)..close];
        if (LooksLikeWorkbookFile(inner))
            return true;
        var after = close + 1;
        if (after >= text.Length) return false;
        return text[after] == '\'' || IsSheetNameStart(text[after]);
    }

    private static bool IsStructuredTableBracket(string text, int open)
    {
        var close = text.IndexOf(']', open);
        if (close <= open) return false;
        if (IsWorkbookBracket(text, open))
            return false;
        if (open == 0) return false;
        if (IsSheetNamePart(text[open - 1]) || text[open - 1] == '[')
            return true;
        return text[open - 1] == ',' && HasOuterTableName(text, open);
    }

    private static bool HasOuterTableName(string text, int open)
    {
        var depth = 0;
        for (var i = open - 1; i >= 0; i--)
        {
            if (text[i] == ']')
            {
                depth++;
                continue;
            }

            if (text[i] != '[')
                continue;
            if (depth > 0)
            {
                depth--;
                continue;
            }

            return i > 0 && IsSheetNamePart(text[i - 1]);
        }

        return false;
    }

    private static bool IsSheetNameStart(char ch) =>
        ch is '_' or '\\' || char.IsLetter(ch);

    private static bool IsSheetNamePart(char ch) =>
        IsSheetNameStart(ch) || char.IsDigit(ch) || ch is '.';

    private static string FileNameOf(string token)
    {
        var text = token.Trim().Trim('\'');
        var slash = Math.Max(text.LastIndexOf('\\'), text.LastIndexOf('/'));
        return slash >= 0 ? text[(slash + 1)..] : text;
    }

    private static void SkipQuoted(string text, ref int index)
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
}
