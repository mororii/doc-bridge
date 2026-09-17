using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// Manual page breaks are written on entire rows or columns.
/// <see href="https://learn.microsoft.com/en-us/office/vba/api/excel.range.pagebreak"/>
/// <c>Rows(25).PageBreak</c> / <c>Columns("J").PageBreak</c>, or
/// <see href="https://learn.microsoft.com/en-us/office/vba/api/excel.hpagebreaks.add"/>
/// <c>HPageBreaks.Add(Before:=Range)</c>. A single cell <c>Range.PageBreak</c>
/// can raise 0x800A03EC. Readback is HPageBreaks/VPageBreaks Location.
/// </summary>
public static class ExcelPageBreakContract
{
    public const int XlPageBreakManual = -4135;
    public const int XlPageBreakNone = -4142;
    public const string AxisRow = "row";
    public const string AxisColumn = "column";

    public enum BreakAxis
    {
        Rows,
        Columns,
    }

    public static BreakAxis ResolveAxis(JsonObject op)
    {
        var text = Json.GetString(op, "orientation") ?? Json.GetString(op, "axis") ?? AxisRow;
        if (text.Equals("column", StringComparison.OrdinalIgnoreCase) ||
            text.Equals("columns", StringComparison.OrdinalIgnoreCase) ||
            text.Equals("col", StringComparison.OrdinalIgnoreCase))
            return BreakAxis.Columns;
        return BreakAxis.Rows;
    }

    public static bool TryAnchor(string? address, out int row, out int column)
    {
        row = 0;
        column = 0;
        if (string.IsNullOrWhiteSpace(address))
            return false;
        var local = ExcelFormula2Contract.CleanA1(address);
        var bang = local.LastIndexOf('!');
        if (bang >= 0)
            local = local[(bang + 1)..];
        var cell = local.Split(':')[0];
        return ExcelFormula2Contract.TryParseA1Cell(cell, out row, out column);
    }

    public static int BreakIndex(string? address, BreakAxis axis) =>
        TryAnchor(address, out var row, out var column)
            ? axis == BreakAxis.Columns ? column : row
            : 0;

    public static bool HasManualBreak(JsonObject locations, BreakAxis axis, int index)
    {
        if (index < 1)
            return false;
        var key = axis == BreakAxis.Columns ? "vertical" : "horizontal";
        foreach (var node in Json.GetArr(locations, key) ?? new JsonArray())
        {
            if (node is JsonValue value && value.TryGetValue<int>(out var found) && found == index)
                return true;
        }

        return false;
    }

    public static string DescribeMissing(string address, BreakAxis axis, int index, JsonObject actual)
    {
        var key = axis == BreakAxis.Columns ? "vertical" : "horizontal";
        var got = string.Join(",",
            (Json.GetArr(actual, key) ?? new JsonArray())
            .Select(node => node is JsonValue value && value.TryGetValue<int>(out var n) ? n.ToString() : "")
            .Where(text => text.Length > 0));
        var axisName = axis == BreakAxis.Columns ? AxisColumn : AxisRow;
        return
            $"set_page_breaks {address} {axisName} expected manual at {index} actual [{got}]";
    }
}
