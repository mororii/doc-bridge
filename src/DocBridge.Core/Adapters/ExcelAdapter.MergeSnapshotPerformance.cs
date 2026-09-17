using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// COM-free merge-style snapshot codec and border readback contract.
/// Adapter capture uses range-level uniform reads with mixed per-cell fallback.
/// Old readers still consume a materialized <c>styles</c> grid.
/// </summary>
public static class ExcelMergeSnapshotPerformance
{
    public const string CaptureUniform = "uniform";
    public const string CaptureMixed = "mixed";

    public static JsonArray ExpandUniformStyles(JsonObject style, int rows, int columns)
    {
        ArgumentNullException.ThrowIfNull(style);
        if (rows < 1 || columns < 1)
            throw new ArgumentOutOfRangeException(rows < 1 ? nameof(rows) : nameof(columns));
        var grid = new JsonArray();
        for (var row = 0; row < rows; row++)
        {
            var styleRow = new JsonArray();
            for (var column = 0; column < columns; column++)
                styleRow.Add(style.DeepClone());
            grid.Add(styleRow);
        }
        return grid;
    }

    public static JsonObject EncodeUniformStyleRange(string address, int rows, int columns, JsonObject style)
    {
        ArgumentNullException.ThrowIfNull(style);
        return new JsonObject
        {
            ["range"] = address,
            ["rows"] = rows,
            ["columns"] = columns,
            ["capture"] = CaptureUniform,
            ["style"] = style.DeepClone(),
            ["styles"] = ExpandUniformStyles(style, rows, columns),
        };
    }

    public static JsonObject EncodeMixedStyleRange(string address, int rows, int columns, JsonArray styles)
    {
        ArgumentNullException.ThrowIfNull(styles);
        return new JsonObject
        {
            ["range"] = address,
            ["rows"] = rows,
            ["columns"] = columns,
            ["capture"] = CaptureMixed,
            ["styles"] = styles,
        };
    }

    public static bool IsUniformCapture(JsonObject styleRange) =>
        string.Equals(Json.GetString(styleRange, "capture"), CaptureUniform, StringComparison.OrdinalIgnoreCase) &&
        Json.GetObj(styleRange, "style") is not null;

    /// <summary>
    /// Uniform encoding requires every cell to match. A first-cell-only sample
    /// is never enough when the grid has more than one cell.
    /// </summary>
    public static bool CanEncodeUniformFromGrid(JsonArray? styles)
    {
        return TryGetUniformStyle(styles, out _);
    }

    public static bool TryGetUniformStyle(JsonArray? styles, out JsonObject? style)
    {
        style = null;
        if (styles is null || styles.Count == 0) return false;
        JsonObject? candidate = null;
        var cells = 0;
        foreach (var rowNode in styles)
        {
            if (rowNode is not JsonArray row || row.Count == 0) return false;
            foreach (var cellNode in row)
            {
                if (cellNode is not JsonObject cell) return false;
                cells++;
                if (candidate is null)
                {
                    candidate = cell;
                    continue;
                }
                if (!JsonNode.DeepEquals(candidate, cell))
                    return false;
            }
        }
        if (candidate is null || cells == 0) return false;
        style = candidate;
        return true;
    }

    public static bool FirstCellOnlyWouldHideMix(JsonArray? styles)
    {
        if (styles is null || styles.Count == 0) return false;
        if (styles[0] is not JsonArray firstRow || firstRow.Count == 0 || firstRow[0] is not JsonObject first)
            return false;
        foreach (var rowNode in styles)
        {
            if (rowNode is not JsonArray row) continue;
            foreach (var cellNode in row)
            {
                if (cellNode is JsonObject cell && !JsonNode.DeepEquals(first, cell))
                    return true;
            }
        }
        return false;
    }

    public static bool TryReadStyleGrid(JsonObject styleRange, out string? address, out JsonArray? styles)
    {
        ArgumentNullException.ThrowIfNull(styleRange);
        address = Json.GetString(styleRange, "range");
        styles = Json.GetArr(styleRange, "styles");
        if (styles is { Count: > 0 } && !string.IsNullOrWhiteSpace(address))
            return true;

        if (Json.GetObj(styleRange, "style") is JsonObject style &&
            Json.GetInt(styleRange, "rows") is int rows &&
            Json.GetInt(styleRange, "columns") is int columns &&
            rows > 0 && columns > 0 &&
            !string.IsNullOrWhiteSpace(address))
        {
            styles = ExpandUniformStyles(style, rows, columns);
            return true;
        }

        return false;
    }

    public static bool RequestedCellBordersMatch(
        IReadOnlyList<ExcelBorderContract.EdgeSpec> requested,
        JsonArray capturedCells,
        string label,
        ICollection<string> errors)
    {
        ArgumentNullException.ThrowIfNull(requested);
        ArgumentNullException.ThrowIfNull(capturedCells);
        ArgumentNullException.ThrowIfNull(errors);
        var cellEdges = requested.Where(spec =>
            string.Equals(spec.Scope, ExcelBorderContract.ScopeCell, StringComparison.OrdinalIgnoreCase)).ToList();
        if (cellEdges.Count == 0) return true;

        var ok = true;
        for (var row = 0; row < capturedCells.Count; row++)
        {
            if (capturedCells[row] is not JsonArray cells)
            {
                errors.Add($"{label}: captured border row {row + 1} is missing");
                ok = false;
                continue;
            }
            for (var column = 0; column < cells.Count; column++)
            {
                if (cells[column] is not JsonObject cell)
                {
                    errors.Add($"{label}: captured border cell [{row + 1},{column + 1}] is missing");
                    ok = false;
                    continue;
                }
                foreach (var spec in cellEdges)
                {
                    if (!OneCapturedBorderMatches(cell, spec, ignoreColor: false))
                    {
                        errors.Add(
                            $"{label}[{row + 1},{column + 1}]: {spec.Name} line/weight/color does not match requested");
                        ok = false;
                    }
                }
            }
        }
        return ok;
    }

    public static bool FirstCellOnlyIgnoreColorWouldAccept(
        IReadOnlyList<ExcelBorderContract.EdgeSpec> requested,
        JsonArray capturedCells)
    {
        if (capturedCells.Count == 0 || capturedCells[0] is not JsonArray firstRow ||
            firstRow.Count == 0 || firstRow[0] is not JsonObject first)
            return false;
        var cellEdges = requested.Where(spec =>
            string.Equals(spec.Scope, ExcelBorderContract.ScopeCell, StringComparison.OrdinalIgnoreCase));
        return cellEdges.All(spec => OneCapturedBorderMatches(first, spec, ignoreColor: true));
    }

    public static bool OneCapturedBorderMatches(JsonObject cell, ExcelBorderContract.EdgeSpec spec, bool ignoreColor)
    {
        var edge = Json.GetObj(cell, spec.Name);
        if (edge is null) return false;
        var line = Json.GetString(edge, "lineStyle");
        var weight = Json.GetString(edge, "weight");
        if (!string.Equals(line, spec.LineStyle, StringComparison.OrdinalIgnoreCase))
            return false;
        if (string.Equals(spec.LineStyle, ExcelBorderContract.LineNone, StringComparison.OrdinalIgnoreCase))
            return true;
        if (!string.Equals(weight, spec.Weight, StringComparison.OrdinalIgnoreCase))
            return false;
        if (ignoreColor) return true;
        var color = edge["color"] is JsonValue value && value.TryGetValue<double>(out var number) ? number : double.NaN;
        return !double.IsNaN(color) && Math.Abs(color - spec.Color) <= 1e-9;
    }

    /// <summary>
    /// Multi-cell range Color=0 is not a mixed-null; Excel also uses 0 for black.
    /// A range aggregate plus last-cell sample does not prove interior cells.
    /// </summary>
    public static bool RangeAggregateColorIsAmbiguous(double? color, int cellCount) =>
        cellCount > 1 && (color is null || Math.Abs(color.Value) <= 1e-9);

    public static bool CellColorsProveUniform(IReadOnlyList<double> colors)
    {
        if (colors is null || colors.Count == 0) return false;
        var first = colors[0];
        for (var i = 1; i < colors.Count; i++)
        {
            if (Math.Abs(colors[i] - first) > 1e-9)
                return false;
        }
        return true;
    }
}
