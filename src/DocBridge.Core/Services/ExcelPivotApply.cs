using System.Globalization;
using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// COM-shaped PivotTable layout. Official
/// https://learn.microsoft.com/en-us/office/vba/api/excel.pivottable.adddatafield
/// and PivotField.Orientation / Caption / NumberFormat / Function.
/// </summary>
public static class ExcelPivotApply
{
    public static void ApplyLayout(object pivot, JsonObject op, bool replaceOmitted)
    {
        ArgumentNullException.ThrowIfNull(pivot);
        ArgumentNullException.ThrowIfNull(op);

        if (replaceOmitted)
            HideNonValueFields(pivot);
        else
        {
            if (op.ContainsKey("rows")) HideOrientation(pivot, ExcelDataObjectCatalog.XlRowField);
            if (op.ContainsKey("columns")) HideOrientation(pivot, ExcelDataObjectCatalog.XlColumnField);
            if (op.ContainsKey("filters")) HideOrientation(pivot, ExcelDataObjectCatalog.XlPageField);
        }

        if (ExcelPivotLayoutContract.AppliesDimension(op, "rows", replaceOmitted))
            SetOrientation(pivot, Json.GetArr(op, "rows"), ExcelDataObjectCatalog.XlRowField);
        if (ExcelPivotLayoutContract.AppliesDimension(op, "columns", replaceOmitted))
            SetOrientation(pivot, Json.GetArr(op, "columns"), ExcelDataObjectCatalog.XlColumnField);
        if (ExcelPivotLayoutContract.AppliesDimension(op, "filters", replaceOmitted))
            SetOrientation(pivot, Json.GetArr(op, "filters"), ExcelDataObjectCatalog.XlPageField);
        if (ExcelPivotLayoutContract.AppliesDimension(op, "values", replaceOmitted))
            ApplyValues(pivot, ExcelPivotLayoutContract.ReadValues(Json.GetArr(op, "values")));
    }

    public static void ApplyValues(object pivot, IReadOnlyList<ExcelPivotLayoutContract.ValueSpec> values)
    {
        HideDataFields(pivot);
        foreach (var spec in values)
        {
            if (!ExcelDataObjectCatalog.TryPivotFunction(spec.Function, out var function))
                function = ExcelDataObjectCatalog.XlSum;
            object? source = null;
            object? data = null;
            try
            {
                source = (object)((dynamic)pivot).PivotFields(spec.Field);
                data = spec.Caption is null
                    ? (object)((dynamic)pivot).AddDataField(source, Type.Missing, function)
                    : (object)((dynamic)pivot).AddDataField(source, spec.Caption, function);
                ((dynamic)data).Function = function;
                if (spec.Caption is not null)
                    ((dynamic)data).Caption = spec.Caption;
                if (spec.NumberFormat is not null)
                    ((dynamic)data).NumberFormat = spec.NumberFormat;
            }
            finally
            {
                ExcelOwnedCom.ReleaseOnce(data);
                ExcelOwnedCom.ReleaseOnce(source);
            }
        }
    }

    private static void SetOrientation(object pivot, JsonArray? names, int orientation)
    {
        if (names is null) return;
        foreach (var node in names)
        {
            var fieldName = node?.ToString();
            if (string.IsNullOrWhiteSpace(fieldName)) continue;
            object? field = null;
            try
            {
                field = (object)((dynamic)pivot).PivotFields(fieldName);
                ((dynamic)field).Orientation = orientation;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"[EXCEL_PIVOT_LAYOUT] failed to set field '{fieldName}' orientation: {ex.Message}", ex);
            }
            finally { ExcelOwnedCom.ReleaseOnce(field); }
        }
    }

    private static void HideNonValueFields(object pivot)
    {
        object? fields = null;
        try
        {
            fields = (object)((dynamic)pivot).PivotFields();
            var count = Convert.ToInt32(((dynamic)fields).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                object? field = null;
                try
                {
                    field = (object)((dynamic)fields).Item(index);
                    if (IsCalculatedField(field) ||
                        ReadOrientation(field) == ExcelDataObjectCatalog.XlDataField ||
                        IsDataAxisField(pivot, field))
                        continue;
                    var name = Convert.ToString(((dynamic)field).Name, CultureInfo.InvariantCulture) ?? "";
                    try
                    {
                        ((dynamic)field).Orientation = ExcelDataObjectCatalog.XlHidden;
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException(
                            $"[EXCEL_PIVOT_LAYOUT] failed to hide field '{name}': {ex.Message}", ex);
                    }
                }
                finally { ExcelOwnedCom.ReleaseOnce(field); }
            }
        }
        finally { ExcelOwnedCom.ReleaseOnce(fields); }
    }

    private static void HideOrientation(object pivot, int orientation)
    {
        object? fields = null;
        try
        {
            fields = (object)((dynamic)pivot).PivotFields();
            var count = Convert.ToInt32(((dynamic)fields).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                object? field = null;
                try
                {
                    field = (object)((dynamic)fields).Item(index);
                    if (ReadOrientation(field) != orientation) continue;
                    var name = Convert.ToString(((dynamic)field).Name, CultureInfo.InvariantCulture) ?? "";
                    try
                    {
                        ((dynamic)field).Orientation = ExcelDataObjectCatalog.XlHidden;
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException(
                            $"[EXCEL_PIVOT_LAYOUT] failed to hide field '{name}': {ex.Message}", ex);
                    }
                }
                finally { ExcelOwnedCom.ReleaseOnce(field); }
            }
        }
        finally { ExcelOwnedCom.ReleaseOnce(fields); }
    }

    private static void HideDataFields(object pivot)
    {
        object? fields = null;
        try
        {
            fields = (object)((dynamic)pivot).DataFields;
            var count = Convert.ToInt32(((dynamic)fields).Count, CultureInfo.InvariantCulture);
            for (var index = count; index >= 1; index--)
            {
                object? field = null;
                try
                {
                    field = (object)((dynamic)fields).Item(index);
                    var name = Convert.ToString(((dynamic)field).Name, CultureInfo.InvariantCulture) ?? "";
                    try
                    {
                        ((dynamic)field).Orientation = ExcelDataObjectCatalog.XlHidden;
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException(
                            $"[EXCEL_PIVOT_LAYOUT] failed to hide data field '{name}': {ex.Message}", ex);
                    }
                }
                finally { ExcelOwnedCom.ReleaseOnce(field); }
            }
        }
        catch (InvalidOperationException) { throw; }
        catch { /* DataFields collection is absent */ }
        finally { ExcelOwnedCom.ReleaseOnce(fields); }
    }

    private static int ReadOrientation(object field) =>
        Convert.ToInt32(((dynamic)field).Orientation, CultureInfo.InvariantCulture);

    private static bool IsCalculatedField(object field)
    {
        try { return Convert.ToBoolean(((dynamic)field).IsCalculated, CultureInfo.InvariantCulture); }
        catch { return false; }
    }

    private static bool IsDataAxisField(object pivot, object field)
    {
        object? axis = null;
        try
        {
            axis = (object)((dynamic)pivot).DataPivotField;
            if (axis is null) return false;
            if (ReferenceEquals(axis, field)) return true;
            var axisName = Convert.ToString(((dynamic)axis).Name, CultureInfo.InvariantCulture);
            var fieldName = Convert.ToString(((dynamic)field).Name, CultureInfo.InvariantCulture);
            return !string.IsNullOrWhiteSpace(axisName) &&
                   string.Equals(axisName, fieldName, StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
        finally
        {
            if (!ReferenceEquals(axis, field))
                ExcelOwnedCom.ReleaseOnce(axis);
        }
    }
}
