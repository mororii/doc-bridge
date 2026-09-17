using System.Globalization;

namespace DocBridge.Core.Services;

/// <summary>
/// Walks Excel <c>Range.Areas</c>. <c>Rows</c>/<c>Columns</c>/<c>Count</c> on a
/// union are the first area (or a bounding box) and must not be used to apply
/// or verify a comma union such as <c>A1,A3</c>.
/// </summary>
public static class ExcelFormatAreas
{
    public static void ForEachArea(object rangeObject, Action<object> visit)
    {
        ArgumentNullException.ThrowIfNull(rangeObject);
        ArgumentNullException.ThrowIfNull(visit);
        object? areasObject = null;
        try
        {
            int count;
            try
            {
                areasObject = (object)((dynamic)rangeObject).Areas;
                if (areasObject is null)
                {
                    visit(rangeObject);
                    return;
                }

                count = Convert.ToInt32(((dynamic)areasObject).Count, CultureInfo.InvariantCulture);
            }
            catch
            {
                visit(rangeObject);
                return;
            }

            if (count < 1)
            {
                visit(rangeObject);
                return;
            }

            for (var index = 1; index <= count; index++)
            {
                object? area = null;
                try
                {
                    area = (object)((dynamic)areasObject).Item(index);
                    visit(area);
                }
                finally { RotHelper.ReleaseComReference(area); }
            }
        }
        finally { RotHelper.ReleaseComReference(areasObject); }
    }

    public static bool AllAreas(object rangeObject, Func<object, bool> predicate)
    {
        ArgumentNullException.ThrowIfNull(predicate);
        var ok = true;
        ForEachArea(rangeObject, area =>
        {
            if (!predicate(area)) ok = false;
        });
        return ok;
    }

    public static void ForEachCell(object areaObject, Action<object> visit)
    {
        ArgumentNullException.ThrowIfNull(areaObject);
        ArgumentNullException.ThrowIfNull(visit);
        object? rowsObject = null;
        object? columnsObject = null;
        object? cells = null;
        try
        {
            rowsObject = (object)((dynamic)areaObject).Rows;
            columnsObject = (object)((dynamic)areaObject).Columns;
            var rows = Convert.ToInt32(((dynamic)rowsObject).Count, CultureInfo.InvariantCulture);
            var columns = Convert.ToInt32(((dynamic)columnsObject).Count, CultureInfo.InvariantCulture);
            cells = (object)((dynamic)areaObject).Cells;
            for (var row = 1; row <= rows; row++)
            {
                for (var column = 1; column <= columns; column++)
                {
                    object? cell = null;
                    try
                    {
                        cell = (object)((dynamic)cells).Item(row, column);
                        visit(cell);
                    }
                    finally { RotHelper.ReleaseComReference(cell); }
                }
            }
        }
        finally
        {
            RotHelper.ReleaseComReference(cells);
            RotHelper.ReleaseComReference(columnsObject);
            RotHelper.ReleaseComReference(rowsObject);
        }
    }

    public static int AreaCount(object rangeObject)
    {
        var count = 0;
        ForEachArea(rangeObject, _ => count++);
        return count;
    }

    public static int RectangularCellCount(object areaObject) =>
        TryGetDimensions(areaObject, out var rows, out var columns) ? checked(rows * columns) : 0;

    public static bool TryGetDimensions(object areaObject, out int rows, out int columns)
    {
        rows = 0;
        columns = 0;
        ArgumentNullException.ThrowIfNull(areaObject);
        object? rowsObject = null;
        object? columnsObject = null;
        try
        {
            rowsObject = (object)((dynamic)areaObject).Rows;
            columnsObject = (object)((dynamic)areaObject).Columns;
            rows = Convert.ToInt32(((dynamic)rowsObject).Count, CultureInfo.InvariantCulture);
            columns = Convert.ToInt32(((dynamic)columnsObject).Count, CultureInfo.InvariantCulture);
            return rows >= 1 && columns >= 1;
        }
        catch
        {
            rows = 0;
            columns = 0;
            return false;
        }
        finally
        {
            RotHelper.ReleaseComReference(columnsObject);
            RotHelper.ReleaseComReference(rowsObject);
        }
    }
}
