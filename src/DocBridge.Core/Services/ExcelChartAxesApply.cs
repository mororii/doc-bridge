using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// Writes advertised Axis scale / TickLabels.NumberFormat / AxisTitle.
/// Chart.Axes(Type, AxisGroup). Failed sets throw [EXCEL_CHART_AXIS].
/// </summary>
public static class ExcelChartAxesApply
{
    public static void Apply(object chart, JsonObject? axes)
    {
        ArgumentNullException.ThrowIfNull(chart);
        foreach (var spec in ExcelChartAxesContract.ReadAxes(axes))
            ApplyOne(chart, spec);
    }

    private static void ApplyOne(object chart, ExcelChartAxesContract.AxisSpec spec)
    {
        object? axis = null;
        try
        {
            try
            {
                axis = (object)((dynamic)chart).Axes(spec.AxisType, spec.AxisGroup);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"[EXCEL_CHART_AXIS] axes.{spec.Key} resolve failed", ex);
            }

            if (spec.Visible == false)
                HideAxis(axis, spec.Key);
            if (spec.Minimum is double min)
                Set(axis, spec.Key, "MinimumScale", () => ((dynamic)axis).MinimumScale = min);
            if (spec.Maximum is double max)
                Set(axis, spec.Key, "MaximumScale", () => ((dynamic)axis).MaximumScale = max);
            if (spec.NumberFormat is not null)
                ApplyNumberFormat(axis, spec);
            if (spec.Title is not null)
                ApplyTitle(axis, spec);
        }
        finally { ExcelOwnedCom.ReleaseOnce(axis); }
    }

    private static void HideAxis(object axis, string key)
    {
        Set(axis, key, "visible", () =>
        {
            ((dynamic)axis).HasMajorGridlines = false;
            ((dynamic)axis).HasMinorGridlines = false;
            ((dynamic)axis).MajorTickMark = ExcelDataObjectCatalog.XlTickMarkNone;
            ((dynamic)axis).MinorTickMark = ExcelDataObjectCatalog.XlTickMarkNone;
            ((dynamic)axis).TickLabelPosition = ExcelDataObjectCatalog.XlTickLabelPositionNone;
        });
        object? format = null;
        object? line = null;
        try
        {
            format = (object)((dynamic)axis).Format;
            line = (object)((dynamic)format).Line;
            ((dynamic)line).Visible = ExcelDataObjectCatalog.MsoFalse;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"[EXCEL_CHART_AXIS] axes.{key}.visible apply failed", ex);
        }
        finally
        {
            ExcelOwnedCom.ReleaseOnce(line);
            ExcelOwnedCom.ReleaseOnce(format);
        }
    }

    private static void ApplyNumberFormat(object axis, ExcelChartAxesContract.AxisSpec spec)
    {
        object? labels = null;
        try
        {
            labels = (object)((dynamic)axis).TickLabels;
            ((dynamic)labels).NumberFormat = spec.NumberFormat;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"[EXCEL_CHART_AXIS] axes.{spec.Key}.numberFormat apply failed", ex);
        }
        finally { ExcelOwnedCom.ReleaseOnce(labels); }
    }

    private static void ApplyTitle(object axis, ExcelChartAxesContract.AxisSpec spec)
    {
        try
        {
            if (string.IsNullOrEmpty(spec.Title))
            {
                ((dynamic)axis).HasTitle = false;
                return;
            }

            ((dynamic)axis).HasTitle = true;
            object? title = null;
            try
            {
                title = (object)((dynamic)axis).AxisTitle;
                ((dynamic)title).Text = spec.Title;
            }
            finally { ExcelOwnedCom.ReleaseOnce(title); }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"[EXCEL_CHART_AXIS] axes.{spec.Key}.title apply failed", ex);
        }
    }

    private static void Set(object axis, string key, string field, Action assign)
    {
        try { assign(); }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"[EXCEL_CHART_AXIS] axes.{key}.{field} apply failed", ex);
        }
    }
}
