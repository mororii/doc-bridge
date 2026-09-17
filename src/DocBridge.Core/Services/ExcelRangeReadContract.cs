using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// Validates and plans one bounded rectangular page inside an already resolved,
/// single-area Excel range. Offsets are zero-based relative to that range.
/// </summary>
public static class ExcelRangeReadContract
{
    public enum FormulaMode { Formula, Formula2 }

    public sealed record Page(
        int RowOffset,
        int ColumnOffset,
        int Rows,
        int Columns,
        int Cells,
        int TotalRows,
        int TotalColumns,
        JsonObject? Continuation)
    {
        /// <summary>True only when this response itself contains the entire requested range.</summary>
        public bool Complete => RowOffset == 0 && ColumnOffset == 0 && Rows == TotalRows && Columns == TotalColumns;
        /// <summary>True when another page is available after this response.</summary>
        public bool HasMore => Continuation is not null;
    }

    public static Page Plan(JsonObject args, int totalRows, int totalColumns, int maxAllowedCells)
    {
        if (totalRows < 1 || totalColumns < 1)
            throw new InvalidOperationException("Excel range must contain at least one row and column");
        if (maxAllowedCells < 1)
            throw new ArgumentOutOfRangeException(nameof(maxAllowedCells));

        var rowOffset = ReadNonNegative(args, "rowOffset") ?? 0;
        var columnOffset = ReadNonNegative(args, "columnOffset") ?? 0;
        if (rowOffset >= totalRows)
            throw new InvalidOperationException($"rowOffset {rowOffset} is outside the requested range ({totalRows} rows)");
        if (columnOffset >= totalColumns)
            throw new InvalidOperationException($"columnOffset {columnOffset} is outside the requested range ({totalColumns} columns)");

        var maxCells = ReadPositive(args, "maxCells") ?? maxAllowedCells;
        if (maxCells > maxAllowedCells)
            throw new InvalidOperationException($"maxCells must not exceed {maxAllowedCells}");

        var remainingRows = totalRows - rowOffset;
        var remainingColumns = totalColumns - columnOffset;
        var requestedRows = Math.Min(ReadPositive(args, "maxRows") ?? remainingRows, remainingRows);
        var requestedColumns = Math.Min(ReadPositive(args, "maxColumns") ?? remainingColumns, remainingColumns);

        // Horizontal tiles use one row. This keeps every tile in a row band the
        // same height, including the final narrow tile, so continuation never
        // skips cells that were outside a previous wider tile.
        var columns = Math.Min(requestedColumns, maxCells);
        var horizontalPaging = columnOffset > 0 || requestedColumns < remainingColumns || columns < requestedColumns;
        var rows = horizontalPaging ? 1 : Math.Min(requestedRows, maxCells / columns);
        var cells = checked(rows * columns);

        var nextRow = rowOffset;
        var nextColumn = columnOffset + columns;
        if (nextColumn >= totalColumns)
        {
            nextColumn = 0;
            nextRow += rows;
        }

        JsonObject? continuation = null;
        if (nextRow < totalRows)
        {
            continuation = new JsonObject
            {
                ["rowOffset"] = nextRow,
                ["columnOffset"] = nextColumn,
                ["maxCells"] = maxCells,
            };
            CopyOptionalPositive(args, continuation, "maxRows");
            CopyOptionalPositive(args, continuation, "maxColumns");
        }

        return new Page(rowOffset, columnOffset, rows, columns, cells, totalRows, totalColumns, continuation);
    }

    public static FormulaMode ResolveFormulaMode(JsonObject args)
    {
        var mode = Json.GetString(args, "formulaMode");
        if (string.IsNullOrWhiteSpace(mode) || mode.Equals("formula", StringComparison.OrdinalIgnoreCase))
            return FormulaMode.Formula;
        if (mode.Equals("formula2", StringComparison.OrdinalIgnoreCase))
            return FormulaMode.Formula2;
        throw new InvalidOperationException("formulaMode must be formula or formula2");
    }

    private static int? ReadNonNegative(JsonObject args, string property)
    {
        if (!args.ContainsKey(property)) return null;
        var value = Json.GetInt(args, property);
        if (value is null || value < 0)
            throw new InvalidOperationException($"{property} must be a non-negative integer");
        return value.Value;
    }

    private static int? ReadPositive(JsonObject args, string property)
    {
        if (!args.ContainsKey(property)) return null;
        var value = Json.GetInt(args, property);
        if (value is null || value < 1)
            throw new InvalidOperationException($"{property} must be a positive integer");
        return value.Value;
    }

    private static void CopyOptionalPositive(JsonObject source, JsonObject destination, string property)
    {
        if (source.ContainsKey(property)) destination[property] = ReadPositive(source, property);
    }
}
