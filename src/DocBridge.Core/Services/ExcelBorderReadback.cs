using System.Globalization;
using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// format_range border readback. <c>left/right/top/bottom</c> are each cell's
/// edge (XlBordersIndex edge*). Range.Borders(xlEdgeBottom) is only the
/// perimeter — not A1/A2 on A1:A3. Inside companions are xlInsideHorizontal
/// / xlInsideVertical. See
/// https://learn.microsoft.com/en-us/office/vba/api/excel.borders
/// Union ranges must walk <c>Areas</c>; Rows/Columns/Count are first-area only.
/// Color=0 on a multi-cell area cannot prove uniformity (mixed or black).
/// <c>borders.all</c> inside edges that do not exist on 1xN / Nx1 / 1x1 or
/// inside a MergeArea are skipped, not demanded. Wrong real edges still fail.
/// </summary>
public static class ExcelBorderReadback
{
    public static bool AppliedBordersMatch(object rangeObject, JsonObject? borders)
    {
        if (borders is null) return false;
        var specs = ExcelBorderContract.ReadEdges(borders);
        var ok = true;
        ExcelFormatAreas.ForEachArea(rangeObject, area =>
        {
            if (!ok) return;
            var geometry = ExcelBorderApplicability.ReadGeometry(area);
            foreach (var spec in specs)
            {
                if (!AreaEdgeMatches(area, spec, geometry))
                {
                    ok = false;
                    return;
                }
            }
        });
        return ok;
    }

    public static IReadOnlyList<string> ExplainAppliedBorders(object rangeObject, JsonObject? borders)
    {
        var notes = new List<string>();
        if (borders is null)
        {
            notes.Add("borders object is missing");
            return notes;
        }

        ExcelFormatAreas.ForEachArea(rangeObject, area =>
        {
            var geometry = ExcelBorderApplicability.ReadGeometry(area);
            foreach (var spec in ExcelBorderContract.ReadEdges(borders))
            {
                if (!ExcelBorderApplicability.EdgeApplies(spec.Name, geometry))
                {
                    notes.Add(
                        $"{spec.Name}: skipped ({ExcelBorderApplicability.SkipReason(spec.Name, geometry)})");
                    continue;
                }

                if (AreaEdgeMatches(area, spec, geometry))
                {
                    notes.Add($"{spec.Name}: ok {FormatExpected(spec)}");
                    continue;
                }

                notes.Add($"{spec.Name}: expected {FormatExpected(spec)} actual {FormatActual(area, spec, geometry)}");
            }
        });
        return notes;
    }

    public static bool AreaEdgeMatches(object areaObject, ExcelBorderContract.EdgeSpec spec)
    {
        var geometry = ExcelBorderApplicability.ReadGeometry(areaObject);
        return AreaEdgeMatches(areaObject, spec, geometry);
    }

    public static bool AreaEdgeMatches(
        object areaObject, ExcelBorderContract.EdgeSpec spec, ExcelBorderApplicability.AreaGeometry geometry)
    {
        if (geometry.Rows < 1 || geometry.Columns < 1) return false;
        if (!ExcelBorderApplicability.EdgeApplies(spec.Name, geometry))
            return true;

        if (spec.Scope == ExcelBorderContract.ScopeRange)
            return AppliedInsideOrOutlineMatches(areaObject, spec, geometry);
        return AppliedCellBorderMatchesOneArea(areaObject, spec, geometry);
    }

    public static bool AppliedCellBorderMatches(object rangeObject, ExcelBorderContract.EdgeSpec spec) =>
        ExcelFormatAreas.AllAreas(rangeObject, area => AppliedCellBorderMatchesOneArea(area, spec));

    public static bool AppliedCellBorderMatchesOneArea(object areaObject, ExcelBorderContract.EdgeSpec spec)
    {
        var geometry = ExcelBorderApplicability.ReadGeometry(areaObject);
        if (geometry.Rows < 1 || geometry.Columns < 1) return false;
        return AppliedCellBorderMatchesOneArea(areaObject, spec, geometry);
    }

    public static bool AppliedCellBorderMatchesOneArea(
        object areaObject, ExcelBorderContract.EdgeSpec spec, ExcelBorderApplicability.AreaGeometry geometry)
    {
        if (TryProveEveryCellEdge(areaObject, spec, geometry, out bool provedMatch))
            return provedMatch;

        var allMatch = true;
        var row = 0;
        ExcelFormatAreas.ForEachCell(areaObject, cell =>
        {
            row++;
            var localRow = ((row - 1) / geometry.Columns) + 1;
            var localCol = ((row - 1) % geometry.Columns) + 1;
            if (!ExcelBorderApplicability.IsVisibleCellEdge(geometry, localRow, localCol, spec.Name))
                return;
            if (!OneBorderMatches(cell, spec)) allMatch = false;
        });
        return allMatch;
    }

    /// <summary>
    /// True when range-level <c>xlEdge*</c> does not cover every cell of this
    /// ScopeCell edge. A1:A3 bottom needs xlEdgeBottom plus xlInsideHorizontal.
    /// </summary>
    public static bool ScopeCellNeedsInsideCompanion(string edgeName, int rows, int columns) =>
        ExcelBorderApplicability.ScopeCellNeedsInsideCompanion(
            edgeName, ExcelBorderApplicability.AreaGeometry.Ungrouped(rows, columns));

    public static string? InsideCompanionName(string edgeName) =>
        edgeName.ToLowerInvariant() switch
        {
            "left" or "right" => "insideVertical",
            "top" or "bottom" => "insideHorizontal",
            _ => null,
        };

    public static bool TryProveEveryCellEdge(object rangeObject, ExcelBorderContract.EdgeSpec spec,
        int rows, int columns, out bool allMatch) =>
        TryProveEveryCellEdge(
            rangeObject, spec, ExcelBorderApplicability.AreaGeometry.Ungrouped(rows, columns), out allMatch);

    public static bool TryProveEveryCellEdge(
        object rangeObject,
        ExcelBorderContract.EdgeSpec spec,
        ExcelBorderApplicability.AreaGeometry geometry,
        out bool allMatch)
    {
        allMatch = false;
        if (!TryUnmixedBorderMatch(rangeObject, spec, out var outerMatch))
            return false;
        if (!outerMatch)
        {
            allMatch = false;
            return true;
        }

        if (!ExcelBorderApplicability.ScopeCellNeedsInsideCompanion(spec.Name, geometry))
        {
            allMatch = true;
            return true;
        }

        var inside = InsideCompanionName(spec.Name);
        if (inside is null) return false;
        var insideSpec = new ExcelBorderContract.EdgeSpec(
            inside, ExcelBorderContract.ScopeRange, spec.Weight, spec.LineStyle, spec.Color);
        if (!TryUnmixedBorderMatch(rangeObject, insideSpec, out var insideMatch))
            return false;
        allMatch = insideMatch;
        return true;
    }

    public static bool AppliedInsideOrOutlineMatches(
        object areaObject, ExcelBorderContract.EdgeSpec spec, ExcelBorderApplicability.AreaGeometry geometry)
    {
        if (TryUnmixedBorderMatch(areaObject, spec, out var uniform))
            return uniform;

        if (spec.Name.Equals("insideHorizontal", StringComparison.OrdinalIgnoreCase))
            return VisibleInteriorHorizontalsMatch(areaObject, spec, geometry);
        if (spec.Name.Equals("insideVertical", StringComparison.OrdinalIgnoreCase))
            return VisibleInteriorVerticalsMatch(areaObject, spec, geometry);
        return VisiblePerimeterMatches(areaObject, spec, geometry);
    }

    public static bool VisiblePerimeterMatches(
        object areaObject, ExcelBorderContract.EdgeSpec spec, ExcelBorderApplicability.AreaGeometry geometry)
    {
        var cellEdge = PerimeterCellEdge(spec.Name);
        if (cellEdge is null) return false;
        var segment = new ExcelBorderContract.EdgeSpec(
            cellEdge, ExcelBorderContract.ScopeCell, spec.Weight, spec.LineStyle, spec.Color);
        var ok = true;
        var index = 0;
        ExcelFormatAreas.ForEachCell(areaObject, cell =>
        {
            index++;
            var row = ((index - 1) / geometry.Columns) + 1;
            var column = ((index - 1) % geometry.Columns) + 1;
            if (!IsAreaPerimeterCell(geometry, row, column, cellEdge))
                return;
            if (!OneBorderMatches(cell, segment)) ok = false;
        });
        return ok;
    }

    public static string? PerimeterCellEdge(string edgeName) =>
        edgeName.ToLowerInvariant() switch
        {
            "left" or "outlineleft" => "left",
            "right" or "outlineright" => "right",
            "top" or "outlinetop" => "top",
            "bottom" or "outlinebottom" => "bottom",
            _ => null,
        };

    public static bool IsAreaPerimeterCell(
        ExcelBorderApplicability.AreaGeometry geometry, int row, int column, string cellEdge) =>
        cellEdge switch
        {
            "left" => column == 1 && ExcelBorderApplicability.IsVisibleCellEdge(geometry, row, column, "left"),
            "right" => column == geometry.Columns &&
                       ExcelBorderApplicability.IsVisibleCellEdge(geometry, row, column, "right"),
            "top" => row == 1 && ExcelBorderApplicability.IsVisibleCellEdge(geometry, row, column, "top"),
            "bottom" => row == geometry.Rows &&
                        ExcelBorderApplicability.IsVisibleCellEdge(geometry, row, column, "bottom"),
            _ => false,
        };

    public static bool VisibleInteriorHorizontalsMatch(
        object areaObject, ExcelBorderContract.EdgeSpec spec, ExcelBorderApplicability.AreaGeometry geometry)
    {
        var ok = true;
        var index = 0;
        ExcelFormatAreas.ForEachCell(areaObject, cell =>
        {
            index++;
            var row = ((index - 1) / geometry.Columns) + 1;
            var column = ((index - 1) % geometry.Columns) + 1;
            if (row >= geometry.Rows) return;
            if (ExcelBorderApplicability.SameMerge(geometry, row, column, row + 1, column))
                return;
            var segment = new ExcelBorderContract.EdgeSpec(
                "bottom", ExcelBorderContract.ScopeCell, spec.Weight, spec.LineStyle, spec.Color);
            if (!OneBorderMatches(cell, segment)) ok = false;
        });
        return ok;
    }

    public static bool VisibleInteriorVerticalsMatch(
        object areaObject, ExcelBorderContract.EdgeSpec spec, ExcelBorderApplicability.AreaGeometry geometry)
    {
        var ok = true;
        var index = 0;
        ExcelFormatAreas.ForEachCell(areaObject, cell =>
        {
            index++;
            var row = ((index - 1) / geometry.Columns) + 1;
            var column = ((index - 1) % geometry.Columns) + 1;
            if (column >= geometry.Columns) return;
            if (ExcelBorderApplicability.SameMerge(geometry, row, column, row, column + 1))
                return;
            var segment = new ExcelBorderContract.EdgeSpec(
                "right", ExcelBorderContract.ScopeCell, spec.Weight, spec.LineStyle, spec.Color);
            if (!OneBorderMatches(cell, segment)) ok = false;
        });
        return ok;
    }

    public static bool OneBorderMatches(object target, ExcelBorderContract.EdgeSpec spec)
    {
        var index = ExcelBorderContract.EdgeIndex(spec.Name);
        if (index == 0) return false;
        object? borders = null;
        object? border = null;
        try
        {
            borders = (object)((dynamic)target).Borders;
            border = (object)((dynamic)borders).Item(index);
            var line = Convert.ToInt32(((dynamic)border).LineStyle, CultureInfo.InvariantCulture);
            if (string.Equals(spec.LineStyle, ExcelBorderContract.LineNone, StringComparison.OrdinalIgnoreCase))
                return line == ExcelBorderContract.XlLineStyleNone;
            return line == ExcelBorderContract.LineStyleValues[spec.LineStyle]
                && Convert.ToInt32(((dynamic)border).Weight, CultureInfo.InvariantCulture) ==
                   ExcelBorderContract.WeightValues[spec.Weight]
                && Math.Abs(Convert.ToDouble(((dynamic)border).Color, CultureInfo.InvariantCulture) - spec.Color) <= 1e-9;
        }
        catch
        {
            return false;
        }
        finally
        {
            RotHelper.ReleaseComReference(border);
            RotHelper.ReleaseComReference(borders);
        }
    }

    public static bool TryUnmixedBorderMatch(object target, ExcelBorderContract.EdgeSpec spec, out bool match)
    {
        match = false;
        var index = ExcelBorderContract.EdgeIndex(spec.Name);
        if (index == 0) return false;
        object? borders = null;
        object? border = null;
        try
        {
            borders = (object)((dynamic)target).Borders;
            border = (object)((dynamic)borders).Item(index);
            var line = SafeGet(() => (object?)((dynamic)border).LineStyle);
            var weight = SafeGet(() => (object?)((dynamic)border).Weight);
            var color = SafeGet(() => (object?)((dynamic)border).Color);
            if (IsMixed(line) || IsMixed(weight) ||
                (!string.Equals(spec.LineStyle, ExcelBorderContract.LineNone, StringComparison.OrdinalIgnoreCase) &&
                 IsMixed(color)))
                return false;
            if (ColorZeroCannotProve(color, spec, ExcelFormatAreas.RectangularCellCount(target)))
                return false;
            match = OneBorderMatches(target, spec);
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            RotHelper.ReleaseComReference(border);
            RotHelper.ReleaseComReference(borders);
        }
    }

    /// <summary>
    /// Excel may report Color=0 for mixed multi-segment borders or for black.
    /// A multi-cell area therefore cannot use 0 as a uniform fast proof.
    /// Single-cell Color=0 is real black.
    /// </summary>
    public static bool ColorZeroCannotProve(object? color, ExcelBorderContract.EdgeSpec spec, int cellCount)
    {
        if (string.Equals(spec.LineStyle, ExcelBorderContract.LineNone, StringComparison.OrdinalIgnoreCase))
            return false;
        if (cellCount <= 1) return false;
        if (IsMixed(color)) return false;
        try
        {
            return Math.Abs(Convert.ToDouble(color, CultureInfo.InvariantCulture)) <= 1e-9;
        }
        catch
        {
            return true;
        }
    }

    private static string FormatExpected(ExcelBorderContract.EdgeSpec spec) =>
        string.Equals(spec.LineStyle, ExcelBorderContract.LineNone, StringComparison.OrdinalIgnoreCase)
            ? "none"
            : $"{spec.Weight}/{spec.LineStyle}/color={spec.Color.ToString("G17", CultureInfo.InvariantCulture)}";

    private static string FormatActual(
        object areaObject, ExcelBorderContract.EdgeSpec spec, ExcelBorderApplicability.AreaGeometry geometry)
    {
        if (TryReadTriple(areaObject, spec.Name, out var line, out var weight, out var color))
            return DescribeTriple(line, weight, color);
        if (spec.Name.Equals("insideHorizontal", StringComparison.OrdinalIgnoreCase) ||
            spec.Name.Equals("insideVertical", StringComparison.OrdinalIgnoreCase))
            return "mixed-or-unreadable range inside; visible-boundary fallback failed";
        return $"visible {spec.Name} segment mismatch on {geometry.Rows}x{geometry.Columns}";
    }

    private static bool TryReadTriple(object target, string edgeName, out int? line, out int? weight, out double? color)
    {
        line = null;
        weight = null;
        color = null;
        var index = ExcelBorderContract.EdgeIndex(edgeName);
        if (index == 0) return false;
        object? borders = null;
        object? border = null;
        try
        {
            borders = (object)((dynamic)target).Borders;
            border = (object)((dynamic)borders).Item(index);
            var rawLine = SafeGet(() => (object?)((dynamic)border).LineStyle);
            var rawWeight = SafeGet(() => (object?)((dynamic)border).Weight);
            var rawColor = SafeGet(() => (object?)((dynamic)border).Color);
            if (!IsMixed(rawLine))
                line = Convert.ToInt32(rawLine, CultureInfo.InvariantCulture);
            if (!IsMixed(rawWeight))
                weight = Convert.ToInt32(rawWeight, CultureInfo.InvariantCulture);
            if (!IsMixed(rawColor))
                color = Convert.ToDouble(rawColor, CultureInfo.InvariantCulture);
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            RotHelper.ReleaseComReference(border);
            RotHelper.ReleaseComReference(borders);
        }
    }

    private static string DescribeTriple(int? line, int? weight, double? color)
    {
        if (line is null && weight is null && color is null) return "mixed";
        if (line == ExcelBorderContract.XlLineStyleNone) return "none";
        var lineName = line is null ? "mixed-line" : ExcelBorderContract.LineStyleName(line.Value);
        var weightName = weight is null ? "mixed-weight" : ExcelBorderContract.WeightName(weight.Value);
        var colorText = color is null ? "mixed-color" : color.Value.ToString("G17", CultureInfo.InvariantCulture);
        return $"{weightName}/{lineName}/color={colorText}";
    }

    private static object? SafeGet(Func<object?> read)
    {
        try { return read(); }
        catch { return null; }
    }

    private static bool IsMixed(object? value) => value is null or DBNull;
}
