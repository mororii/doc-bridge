using System.Globalization;
using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// Applicable format_range border edges. <c>borders.all</c> expands
/// insideHorizontal/insideVertical even when the area has no such line:
/// 1xN has no interior horizontal, Nx1 no interior vertical, 1x1 neither.
/// Lines inside one MergeArea are not visible boundaries. Apply and
/// readback skip those; real outer and visible interior segments still fail
/// when wrong or missing.
/// </summary>
public static class ExcelBorderApplicability
{
    public readonly record struct MergeRect(int Row, int Column, int Rows, int Columns)
    {
        public bool Contains(int row, int column) =>
            row >= Row && row < Row + Rows && column >= Column && column < Column + Columns;
    }

    public readonly record struct AreaGeometry(int Rows, int Columns, IReadOnlyList<MergeRect> Merges)
    {
        public static AreaGeometry Ungrouped(int rows, int columns) =>
            new(rows, columns, Array.Empty<MergeRect>());
    }

    public static bool InsideEdgeApplies(string edgeName, int rows, int columns) =>
        EdgeApplies(edgeName, AreaGeometry.Ungrouped(rows, columns));

    public static bool EdgeApplies(string edgeName, AreaGeometry geometry) =>
        edgeName.ToLowerInvariant() switch
        {
            "insidehorizontal" => HasVisibleInteriorHorizontal(geometry),
            "insidevertical" => HasVisibleInteriorVertical(geometry),
            _ => true,
        };

    public static bool ScopeCellNeedsInsideCompanion(string edgeName, AreaGeometry geometry) =>
        edgeName.ToLowerInvariant() switch
        {
            "left" or "right" => HasVisibleInteriorVertical(geometry),
            "top" or "bottom" => HasVisibleInteriorHorizontal(geometry),
            _ => false,
        };

    public static bool HasVisibleInteriorHorizontal(AreaGeometry geometry)
    {
        if (geometry.Rows <= 1) return false;
        for (var row = 1; row < geometry.Rows; row++)
        {
            for (var column = 1; column <= geometry.Columns; column++)
            {
                if (!SameMerge(geometry, row, column, row + 1, column))
                    return true;
            }
        }
        return false;
    }

    public static bool HasVisibleInteriorVertical(AreaGeometry geometry)
    {
        if (geometry.Columns <= 1) return false;
        for (var row = 1; row <= geometry.Rows; row++)
        {
            for (var column = 1; column < geometry.Columns; column++)
            {
                if (!SameMerge(geometry, row, column, row, column + 1))
                    return true;
            }
        }
        return false;
    }

    public static bool SameMerge(AreaGeometry geometry, int r1, int c1, int r2, int c2)
    {
        foreach (var merge in geometry.Merges)
        {
            if (merge.Contains(r1, c1) && merge.Contains(r2, c2))
                return true;
        }
        return false;
    }

    public static bool IsVisibleCellEdge(AreaGeometry geometry, int row, int column, string edgeName)
    {
        if (row < 1 || column < 1 || row > geometry.Rows || column > geometry.Columns)
            return false;
        var merge = FindMerge(geometry, row, column);
        if (merge is null) return true;
        return edgeName.ToLowerInvariant() switch
        {
            "left" => column == merge.Value.Column,
            "right" => column == merge.Value.Column + merge.Value.Columns - 1,
            "top" => row == merge.Value.Row,
            "bottom" => row == merge.Value.Row + merge.Value.Rows - 1,
            _ => true,
        };
    }

    public static MergeRect? FindMerge(AreaGeometry geometry, int row, int column)
    {
        foreach (var merge in geometry.Merges)
        {
            if (merge.Contains(row, column))
                return merge;
        }
        return null;
    }

    public static string SkipReason(string edgeName, AreaGeometry geometry)
    {
        var name = edgeName.ToLowerInvariant();
        if (name == "insidehorizontal" && geometry.Rows <= 1)
            return $"{geometry.Rows}x{geometry.Columns} has no interior horizontal";
        if (name == "insidevertical" && geometry.Columns <= 1)
            return $"{geometry.Rows}x{geometry.Columns} has no interior vertical";
        if (name == "insidehorizontal" && !HasVisibleInteriorHorizontal(geometry))
            return $"{geometry.Rows}x{geometry.Columns} interior horizontals are inside MergeArea";
        if (name == "insidevertical" && !HasVisibleInteriorVertical(geometry))
            return $"{geometry.Rows}x{geometry.Columns} interior verticals are inside MergeArea";
        return $"{edgeName} is not an applicable boundary";
    }

    public static JsonObject ForRange(JsonObject borders, int rows, int columns) =>
        ForRange(borders, AreaGeometry.Ungrouped(rows, columns));

    public static JsonObject ForRange(JsonObject borders, AreaGeometry geometry)
    {
        var edges = Json.GetArr(borders, "edges");
        if (edges is not null)
        {
            var kept = new JsonArray();
            foreach (var node in edges.OfType<JsonObject>())
            {
                var name = Json.GetString(node, "name");
                if (string.IsNullOrWhiteSpace(name) || !EdgeApplies(name, geometry))
                    continue;
                kept.Add(node.DeepClone());
            }

            return new JsonObject { ["edges"] = kept };
        }

        var filtered = new JsonObject();
        foreach (var (key, value) in borders)
        {
            if (key.Equals("all", StringComparison.OrdinalIgnoreCase))
            {
                foreach (var edge in new[]
                         {
                             "left", "right", "top", "bottom", "insideHorizontal", "insideVertical",
                         })
                {
                    if (!EdgeApplies(edge, geometry))
                        continue;
                    filtered[edge] = value?.DeepClone();
                }

                continue;
            }

            if (!EdgeApplies(key, geometry))
                continue;
            filtered[key] = value?.DeepClone();
        }

        return filtered;
    }

    /// <summary>
    /// Cells visited by the mixed/multiple-merge scan only.
    /// <c>MergeCells=false</c> and a first-cell MergeArea that covers the
    /// whole area do not increment this. See
    /// https://learn.microsoft.com/en-us/office/vba/api/excel.range.mergecells
    /// and
    /// https://learn.microsoft.com/en-us/office/vba/api/excel.range.mergearea
    /// (<c>MergeArea</c> is defined for a single-cell range only).
    /// </summary>
    public static int MergeScanCells { get; private set; }

    public static void ResetMergeScanCells() => MergeScanCells = 0;

    public static AreaGeometry ReadGeometry(object areaObject)
    {
        if (!ExcelFormatAreas.TryGetDimensions(areaObject, out var rows, out var columns))
            return AreaGeometry.Ungrouped(0, 0);
        return new AreaGeometry(rows, columns, ReadMergeRects(areaObject, rows, columns));
    }

    public static IReadOnlyList<MergeRect> ReadMergeRects(object areaObject, int rows, int columns)
    {
        var originRow = TryScalar(areaObject, "Row") ?? 1;
        var originCol = TryScalar(areaObject, "Column") ?? 1;
        if (TryBool(areaObject, "MergeCells", out var merged))
        {
            if (!merged)
                return Array.Empty<MergeRect>();
            if (TryReadFirstCellMergeCoveringArea(areaObject, originRow, originCol, rows, columns, out var whole))
                return [whole];
        }

        return ScanMergeRects(areaObject, originRow, originCol, rows, columns);
    }

    public static bool CoversEntireArea(MergeRect rect, int rows, int columns) =>
        rect.Row <= 1 && rect.Column <= 1 &&
        rect.Row + rect.Rows - 1 >= rows &&
        rect.Column + rect.Columns - 1 >= columns;

    private static IReadOnlyList<MergeRect> ScanMergeRects(
        object areaObject, int originRow, int originCol, int rows, int columns)
    {
        var found = new List<MergeRect>();
        ExcelFormatAreas.ForEachCell(areaObject, cell =>
        {
            MergeScanCells++;
            if (!TryBool(cell, "MergeCells", out var cellMerged) || !cellMerged)
                return;
            if (!TryReadMergeArea(cell, originRow, originCol, rows, columns, out var rect))
                return;
            if (found.Exists(existing =>
                    existing.Row == rect.Row && existing.Column == rect.Column &&
                    existing.Rows == rect.Rows && existing.Columns == rect.Columns))
                return;
            found.Add(rect);
        });
        return found;
    }

    /// <summary>
    /// Official <c>Range.MergeArea</c> is valid on a single cell. Peek
    /// <c>Cells(1,1)</c> only; treat as one merge when that box covers every
    /// cell of this area. <c>MergeCells=true</c> alone is not that proof.
    /// </summary>
    private static bool TryReadFirstCellMergeCoveringArea(
        object areaObject, int originRow, int originCol, int rows, int columns, out MergeRect rect)
    {
        rect = default;
        object? cells = null;
        object? first = null;
        try
        {
            cells = (object)((dynamic)areaObject).Cells;
            first = (object)((dynamic)cells).Item(1, 1);
            if (!TryReadMergeArea(first, originRow, originCol, rows, columns, out var raw))
                return false;
            if (!CoversEntireArea(raw, rows, columns))
                return false;
            rect = new MergeRect(1, 1, rows, columns);
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            if (!ReferenceEquals(first, areaObject))
                RotHelper.ReleaseComReference(first);
            RotHelper.ReleaseComReference(cells);
        }
    }

    private static bool TryReadMergeArea(
        object cell, int originRow, int originCol, int rows, int columns, out MergeRect rect)
    {
        rect = default;
        object? area = null;
        object? areaRows = null;
        object? areaCols = null;
        try
        {
            area = (object)((dynamic)cell).MergeArea;
            if (area is null) return false;
            var sheetRow = TryScalar(area, "Row") ?? originRow;
            var sheetCol = TryScalar(area, "Column") ?? originCol;
            areaRows = SafeGet(() => (object)((dynamic)area).Rows);
            areaCols = SafeGet(() => (object)((dynamic)area).Columns);
            var height = areaRows is null ? 1 : Convert.ToInt32(((dynamic)areaRows).Count, CultureInfo.InvariantCulture);
            var width = areaCols is null ? 1 : Convert.ToInt32(((dynamic)areaCols).Count, CultureInfo.InvariantCulture);
            var localRow = sheetRow - originRow + 1;
            var localCol = sheetCol - originCol + 1;
            if (localRow < 1 || localCol < 1) return false;
            if (localRow > rows || localCol > columns) return false;
            rect = new MergeRect(localRow, localCol, height, width);
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            RotHelper.ReleaseComReference(areaCols);
            RotHelper.ReleaseComReference(areaRows);
            if (!ReferenceEquals(area, cell))
                RotHelper.ReleaseComReference(area);
        }
    }

    private static int? TryScalar(object target, string name)
    {
        try
        {
            var value = name switch
            {
                "Row" => (object?)((dynamic)target).Row,
                "Column" => (object?)((dynamic)target).Column,
                _ => null,
            };
            return value is null ? null : Convert.ToInt32(value, CultureInfo.InvariantCulture);
        }
        catch
        {
            return null;
        }
    }

    private static bool TryBool(object target, string name, out bool value)
    {
        value = false;
        try
        {
            var raw = name == "MergeCells" ? (object?)((dynamic)target).MergeCells : null;
            if (raw is null or DBNull) return false;
            value = Convert.ToBoolean(raw, CultureInfo.InvariantCulture);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static object? SafeGet(Func<object?> read)
    {
        try { return read(); }
        catch { return null; }
    }
}
