using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// set_outline semantics. Group/ungroup is not the same as collapse/expand.
/// show:false collapses detail; it does not ClearOutline.
/// </summary>
public static class ExcelOutlineContract
{
    public const int XlSummaryAbove = 0;
    public const int XlSummaryBelow = 1;
    public const int XlSummaryOnLeft = -4131;
    public const int XlSummaryOnRight = -4152;

    public enum OutlineAction
    {
        Group,
        Ungroup,
        Collapse,
        Expand,
    }

    public enum OutlineAxis
    {
        Rows,
        Columns,
    }

    public static OutlineAction ResolveAction(JsonObject op)
    {
        if (Json.GetBool(op, "clear") ||
            string.Equals(Json.GetString(op, "action"), "ungroup", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(Json.GetString(op, "action"), "clear", StringComparison.OrdinalIgnoreCase))
            return OutlineAction.Ungroup;
        if (op.ContainsKey("show"))
            return Json.GetBool(op, "show") ? OutlineAction.Expand : OutlineAction.Collapse;
        return OutlineAction.Group;
    }

    public static int? RequestedLevel(JsonObject op)
    {
        var level = Json.GetInt(op, "level");
        if (level is null) return null;
        if (level is < 1 or > 8)
            throw new InvalidOperationException("set_outline level must be 1..8");
        return level;
    }

    public static int? SummaryRow(JsonObject op) =>
        op.ContainsKey("summaryBelow") ? (Json.GetBool(op, "summaryBelow") ? XlSummaryBelow : XlSummaryAbove) : null;

    public static int? SummaryColumn(JsonObject op) =>
        op.ContainsKey("summaryRight") ? (Json.GetBool(op, "summaryRight") ? XlSummaryOnRight : XlSummaryOnLeft) : null;

    public static OutlineAxis ResolveAxis(JsonObject op)
    {
        var axis = Json.GetString(op, "axis") ?? Json.GetString(op, "orientation");
        if (string.IsNullOrWhiteSpace(axis))
            return OutlineAxis.Rows;
        if (axis.Equals("column", StringComparison.OrdinalIgnoreCase) ||
            axis.Equals("columns", StringComparison.OrdinalIgnoreCase) ||
            axis.Equals("col", StringComparison.OrdinalIgnoreCase))
            return OutlineAxis.Columns;
        if (axis.Equals("row", StringComparison.OrdinalIgnoreCase) ||
            axis.Equals("rows", StringComparison.OrdinalIgnoreCase))
            return OutlineAxis.Rows;
        throw new InvalidOperationException("set_outline axis must be row or column");
    }

    public static string Describe(JsonObject op)
    {
        var action = ResolveAction(op);
        var level = RequestedLevel(op);
        var axis = ResolveAxis(op) == OutlineAxis.Columns ? "columns" : "rows";
        return action switch
        {
            OutlineAction.Ungroup => $"ungroup/clear {axis} outline (not collapse)",
            OutlineAction.Collapse => $"collapse {axis} detail ShowDetail=false",
            OutlineAction.Expand => $"expand {axis} detail ShowDetail=true",
            _ => level is int value ? $"group {axis} OutlineLevel={value}" : $"group {axis}",
        };
    }

    public static bool MatchesReadback(OutlineAction action, int? requestedLevel, int actualLevel, bool showDetail, bool hasOutline)
    {
        return action switch
        {
            OutlineAction.Ungroup => !hasOutline || actualLevel <= 1,
            OutlineAction.Collapse => !showDetail,
            OutlineAction.Expand => showDetail,
            OutlineAction.Group => requestedLevel is null || actualLevel == requestedLevel,
            _ => false,
        };
    }
}
