using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// calculate is recalculation only. Formula2 writes belong on set_formulas.
/// Workbook calculate must not call Application.Calculate (other user books).
/// </summary>
public static class ExcelCalculateContract
{
    public enum CalculateScope
    {
        Range,
        Workbook,
    }

    public static CalculateScope ResolveScope(JsonObject op) =>
        string.IsNullOrWhiteSpace(Json.GetString(op, "range")) ? CalculateScope.Workbook : CalculateScope.Range;

    public static bool RequestsFormula2Rewrite(JsonObject op) => Json.GetBool(op, "formula2");

    public static string Describe(JsonObject op)
    {
        var scope = ResolveScope(op);
        var rewrite = RequestsFormula2Rewrite(op);
        return scope == CalculateScope.Range
            ? (rewrite
                ? "range.Calculate after Formula2 write is rejected; use set_formulas"
                : "range.Calculate on the target workbook only")
            : "each worksheet.Calculate (not Workbook.Calculate, not Application.Calculate)";
    }

    public const int XlDone = 0;
    public const int XlWorksheet = -4167;

    public static bool RecalculationComplete(int? calculationState) =>
        calculationState is null or XlDone;

    public static bool IsWorksheetType(int? type) =>
        type is null or XlWorksheet;
}
