using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// target.sheet is the write/destination sheet. An explicitly sheet-qualified
/// sourceRange may resolve to another sheet in the same workbook.
/// Unqualified destination/write ranges still must match target.sheet.
/// </summary>
public static class ExcelDataRangeBinding
{
    public static bool AllowsIndependentSource(string? op) =>
        op is ExcelDataOperationsContract.CreatePivot
            or ExcelDataOperationsContract.UpdatePivot
            or ExcelDataOperationsContract.CreateChart
            or ExcelDataOperationsContract.UpdateChart
            or ExcelDataOperationsContract.PasteSpecial;

    public static string? WriteRangeField(string? op) =>
        op == ExcelDataOperationsContract.CreatePivot ? "destination" :
        op == ExcelDataOperationsContract.PasteSpecial ? "destination" : null;

    public static bool WriteRangeConflictsWithTarget(string? targetSheet, string? writeRange)
    {
        if (string.IsNullOrWhiteSpace(targetSheet) || string.IsNullOrWhiteSpace(writeRange))
            return false;
        try
        {
            var parsed = ExcelRangeReference.Parse(writeRange);
            return !string.IsNullOrWhiteSpace(parsed.SheetName) &&
                   !string.Equals(parsed.SheetName, targetSheet, StringComparison.OrdinalIgnoreCase);
        }
        catch (FormatException)
        {
            return false;
        }
    }

    public static bool TrySourceSheet(
        string? targetSheet,
        string sourceRange,
        out string sheetName,
        out string address,
        out string? error)
    {
        sheetName = "";
        address = "";
        error = null;
        try
        {
            var parsed = ExcelRangeReference.Parse(sourceRange);
            address = parsed.Address;
            if (!string.IsNullOrWhiteSpace(parsed.SheetName))
            {
                sheetName = parsed.SheetName;
                return true;
            }

            if (string.IsNullOrWhiteSpace(targetSheet))
            {
                error = "unqualified sourceRange requires target.sheet";
                return false;
            }

            sheetName = targetSheet;
            return true;
        }
        catch (FormatException ex)
        {
            error = ex.Message;
            return false;
        }
    }

    public static string QualifiedA1(string sheetName, string address) =>
        $"{ExcelFormulaReference.QuoteSheetName(sheetName)}!{address}";

    public static bool IsUnionAddress(string? range)
    {
        if (string.IsNullOrWhiteSpace(range)) return false;
        try
        {
            var parsed = ExcelRangeReference.Parse(range);
            return parsed.Address.Contains(',', StringComparison.Ordinal) ||
                   parsed.Address.Contains(';', StringComparison.Ordinal);
        }
        catch (FormatException)
        {
            return range.Contains(',', StringComparison.Ordinal);
        }
    }

    public static bool LooksLikeTableSource(string? source)
    {
        if (string.IsNullOrWhiteSpace(source)) return false;
        if (source.Contains('!', StringComparison.Ordinal)) return false;
        if (ExcelDataOperationsContract.TryParseA1(source, out _, allowUnion: false))
            return false;
        return source.IndexOfAny([' ', '\t', '\r', '\n']) < 0 && source.Length is >= 1 and <= 255;
    }

    public static bool MustCreatePrivateCache(int pivotsSharingCache) =>
        pivotsSharingCache > 1;

    public static bool SourceReadbackMatches(
        string? requested, string? actual, string? defaultSheet = null, string? ownerWorkbook = null) =>
        ExcelReferenceIdentity.SameSource(requested, actual, defaultSheet, ownerWorkbook);

    public static IReadOnlyList<string> AddressAreas(string? address)
    {
        if (string.IsNullOrWhiteSpace(address)) return [];
        try
        {
            var parsed = ExcelRangeReference.Parse(address);
            return parsed.Address.Split([',', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        }
        catch (FormatException)
        {
            return address.Split([',', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        }
    }

    public static string? TargetSheet(JsonObject op) =>
        Json.GetString(Json.GetObj(op, "target"), "sheet");
}
