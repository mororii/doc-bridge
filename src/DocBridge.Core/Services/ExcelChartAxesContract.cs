using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// Chart axes: category / value / valueSecondary with optional group,
/// scale, TickLabels.NumberFormat, and AxisTitle.
/// https://learn.microsoft.com/en-us/office/vba/api/excel.chart.axes
/// https://learn.microsoft.com/en-us/office/vba/api/excel.axis.minimumscale
/// https://learn.microsoft.com/en-us/office/vba/api/excel.ticklabels.numberformat
/// https://learn.microsoft.com/en-us/office/vba/api/excel.axis.axistitle
/// </summary>
public static class ExcelChartAxesContract
{
    public const string Category = "category";
    public const string Value = "value";
    public const string ValueSecondary = "valueSecondary";

    public static readonly string[] Keys = [Category, Value, ValueSecondary];

    public readonly record struct AxisSpec(
        string Key,
        int AxisType,
        int AxisGroup,
        string? GroupToken,
        bool? Visible,
        double? Minimum,
        double? Maximum,
        string? NumberFormat,
        string? Title);

    public static IReadOnlyList<AxisSpec> ReadAxes(JsonObject? axes)
    {
        if (axes is null) return [];
        var list = new List<AxisSpec>(3);
        TryAdd(list, axes, Category, ExcelDataObjectCatalog.XlCategory, ExcelDataObjectCatalog.XlPrimary);
        TryAdd(list, axes, Value, ExcelDataObjectCatalog.XlValue, ExcelDataObjectCatalog.XlPrimary);
        TryAdd(list, axes, ValueSecondary, ExcelDataObjectCatalog.XlValue, ExcelDataObjectCatalog.XlSecondary);
        return list;
    }

    public static string ActualKey(AxisSpec spec) =>
        spec.AxisType == ExcelDataObjectCatalog.XlValue &&
        spec.AxisGroup == ExcelDataObjectCatalog.XlSecondary
            ? ValueSecondary
            : spec.Key;

    public static void CompareRequested(JsonObject actual, JsonObject op, string label,
        ICollection<string> mismatches)
    {
        ArgumentNullException.ThrowIfNull(actual);
        ArgumentNullException.ThrowIfNull(op);
        ArgumentNullException.ThrowIfNull(mismatches);
        if (!op.ContainsKey("axes")) return;
        if (op["axes"] is not JsonObject requested)
        {
            mismatches.Add($"{label}: axes must be an object");
            return;
        }

        var actualAxes = Json.GetObj(actual, "axes");
        if (actualAxes is null)
        {
            mismatches.Add($"{label}: axes readback is missing");
            return;
        }

        foreach (var spec in ReadAxes(requested))
        {
            var key = ActualKey(spec);
            var got = Json.GetObj(actualAxes, key);
            if (got is null || Json.GetBool(got, "unreadable"))
            {
                mismatches.Add($"{label}: axes.{spec.Key} is missing");
                continue;
            }

            if (spec.Visible is bool visible && Json.GetBool(got, "visible") != visible)
                mismatches.Add($"{label}: axes.{spec.Key}.visible readback mismatch");
            if (spec.Minimum is double min &&
                (!ExcelDataOperationsContract.TryGetFiniteNumber(got["minimum"], out var actualMin) ||
                 Math.Abs(actualMin - min) > 0.01))
                mismatches.Add($"{label}: axes.{spec.Key}.minimum readback mismatch");
            if (spec.Maximum is double max &&
                (!ExcelDataOperationsContract.TryGetFiniteNumber(got["maximum"], out var actualMax) ||
                 Math.Abs(actualMax - max) > 0.01))
                mismatches.Add($"{label}: axes.{spec.Key}.maximum readback mismatch");
            if (spec.NumberFormat is not null &&
                !string.Equals(Json.GetString(got, "numberFormat"), spec.NumberFormat, StringComparison.Ordinal))
                mismatches.Add($"{label}: axes.{spec.Key}.numberFormat '{Json.GetString(got, "numberFormat")}' != '{spec.NumberFormat}'");
            if (spec.Title is not null &&
                !string.Equals(Json.GetString(got, "title") ?? "", spec.Title, StringComparison.Ordinal))
                mismatches.Add($"{label}: axes.{spec.Key}.title readback mismatch");
            if (spec.GroupToken is not null &&
                !string.Equals(Json.GetString(got, "group"), spec.GroupToken, StringComparison.OrdinalIgnoreCase))
                mismatches.Add($"{label}: axes.{spec.Key}.group '{Json.GetString(got, "group")}' != '{spec.GroupToken}'");
        }
    }

    private static void TryAdd(List<AxisSpec> list, JsonObject axes, string key, int axisType, int defaultGroup)
    {
        var item = Json.GetObj(axes, key);
        if (item is null) return;
        var groupToken = EmptyToNull(Json.GetString(item, "group"));
        var group = defaultGroup;
        if (groupToken is not null && ExcelDataObjectCatalog.TryAxisGroup(groupToken, out var mapped))
            group = mapped;
        double? minimum = ExcelDataOperationsContract.TryGetFiniteNumber(item["minimum"], out var min) ? min : null;
        double? maximum = ExcelDataOperationsContract.TryGetFiniteNumber(item["maximum"], out var max) ? max : null;
        list.Add(new AxisSpec(
            key,
            axisType,
            group,
            groupToken,
            item.ContainsKey("visible") ? Json.GetBool(item, "visible") : null,
            minimum,
            maximum,
            EmptyToNull(Json.GetString(item, "numberFormat")),
            item.ContainsKey("title") ? Json.GetString(item, "title") ?? "" : null));
    }

    private static string? EmptyToNull(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;
}
