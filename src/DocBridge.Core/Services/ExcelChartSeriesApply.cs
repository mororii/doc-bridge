using System.Globalization;
using System.Linq;
using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// Writes advertised Series.Values / XValues / Name, then ChartType / AxisGroup.
/// https://learn.microsoft.com/en-us/office/vba/api/excel.series.values
/// https://learn.microsoft.com/en-us/office/vba/api/excel.series.charttype
/// https://learn.microsoft.com/en-us/office/vba/api/excel.series.axisgroup
/// </summary>
public static class ExcelChartSeriesApply
{
    public static void ApplyData(object chart, JsonArray? series, Func<string, object?>? resolveRange)
    {
        ArgumentNullException.ThrowIfNull(chart);
        var specs = ExcelChartSeriesContract.ReadSeries(series);
        if (specs.Count == 0) return;

        object? collection = null;
        try
        {
            collection = (object)((dynamic)chart).SeriesCollection();
            var needed = specs.Max(item => item.Index);
            while (Convert.ToInt32(((dynamic)collection).Count, CultureInfo.InvariantCulture) < needed)
            {
                object? created = null;
                try { created = (object)((dynamic)collection).NewSeries(); }
                finally { ExcelOwnedCom.ReleaseOnce(created); }
            }

            foreach (var spec in specs.OrderBy(item => item.Index))
            {
                object? item = null;
                try
                {
                    item = (object)((dynamic)collection).Item(spec.Index);
                    if (spec.Range is not null)
                        AssignRange(item, "Values", spec.Range, resolveRange);
                    if (spec.Values is not null)
                        AssignRange(item, "Values", spec.Values, resolveRange);
                    if (spec.Categories is not null)
                        AssignRange(item, "XValues", spec.Categories, resolveRange);
                    if (spec.Name is not null)
                    {
                        ((dynamic)item).Name = spec.NameIsRange && !spec.Name.StartsWith('=')
                            ? "=" + spec.Name
                            : spec.Name;
                    }
                    ApplyCombo(item, spec);
                }
                finally { ExcelOwnedCom.ReleaseOnce(item); }
            }
        }
        finally { ExcelOwnedCom.ReleaseOnce(collection); }
    }

    private static void AssignRange(object series, string property, string text, Func<string, object?>? resolveRange)
    {
        object? resolved = null;
        try
        {
            resolved = resolveRange?.Invoke(text);
            if (property == "XValues")
                ((dynamic)series).XValues = resolved ?? text;
            else
                ((dynamic)series).Values = resolved ?? text;
        }
        finally
        {
            if (!ReferenceEquals(resolved, series))
                ExcelOwnedCom.ReleaseOnce(resolved);
        }
    }

    private static void ApplyCombo(object series, ExcelChartSeriesContract.SeriesSpec spec)
    {
        if (spec.ChartType is not null)
        {
            if (!ExcelDataObjectCatalog.TryChartType(spec.ChartType, out var chartType))
                throw new InvalidOperationException(
                    $"[EXCEL_CHART_SERIES] series[{spec.Index}].chartType '{spec.ChartType}' is not mapped");
            try
            {
                ((dynamic)series).ChartType = chartType;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"[EXCEL_CHART_SERIES] series[{spec.Index}].ChartType apply failed", ex);
            }
        }

        if (spec.AxisGroup is null) return;
        if (!ExcelDataObjectCatalog.TryAxisGroup(spec.AxisGroup, out var axisGroup))
            throw new InvalidOperationException(
                $"[EXCEL_CHART_SERIES] series[{spec.Index}].axisGroup '{spec.AxisGroup}' is not mapped");
        try
        {
            ((dynamic)series).AxisGroup = axisGroup;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"[EXCEL_CHART_SERIES] series[{spec.Index}].AxisGroup apply failed", ex);
        }
    }
}
