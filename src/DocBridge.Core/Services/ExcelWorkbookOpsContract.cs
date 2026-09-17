using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// COM-free contract for Excel workbook-level and advanced-object ops:
/// external links, calculation mode, freeze/paste-special, goal seek,
/// workbook protection, split panes, sparklines, slicers, and cell styles.
/// VBA and macros stay forbidden (<c>run_macro</c>) and no op here executes code.
/// </summary>
public static class ExcelWorkbookOpsContract
{
    // XlLinkType / XlLinkInfo / Application.Calculation constants.
    public const int XlExcelLinks = 1;
    public const int XlLinkInfoStatus = 3;
    public const int XlCalculationManual = -4135;
    public const int XlCalculationAutomatic = -4105;
    public const int XlCalculationSemiautomatic = 2;
    public const int XlCalculationDone = 0;

    // XlSparkType. Win/Loss has no distinct XlSparkType value, so it is
    // rejected until a native mapping is proven (same fail-closed policy as
    // text_to_columns fixed-width).
    public const int XlSparkLine = 1;
    public const int XlSparkColumn = 2;

    public const int MaxLinkDependents = 20_000;
    public const int MaxValueSurfaceCells = 20_000;
    public const int MaxCellStyleCells = 5_000;
    public const int MaxLinkSources = 500;
    public const int MaxPasteCells = 20_000;

    public static readonly IReadOnlyDictionary<string, int> CalculationModes =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["manual"] = XlCalculationManual,
            ["automatic"] = XlCalculationAutomatic,
            ["semiautomatic"] = XlCalculationSemiautomatic,
        };

    public static readonly IReadOnlySet<string> RecalculateTokens = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "none", "full", "fullRebuild",
    };

    public static readonly IReadOnlySet<string> PasteTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
    {
        "values", "formulas", "formats", "all",
    };

    public static readonly IReadOnlyDictionary<string, int> PasteOperations =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["none"] = 0,
            ["add"] = 1,
            ["subtract"] = 2,
            ["multiply"] = 3,
            ["divide"] = 4,
        };

    public static readonly IReadOnlyDictionary<string, int> SparklineTypes =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["line"] = XlSparkLine,
            ["column"] = XlSparkColumn,
        };

    public static readonly IReadOnlyDictionary<int, string> LinkStatusNames =
        new Dictionary<int, string>
        {
            [0] = "ok",
            [1] = "missingFile",
            [2] = "missingSheet",
            [3] = "old",
            [4] = "sourceNotCalculated",
            [5] = "indeterminate",
            [6] = "notStarted",
            [7] = "invalidName",
            [8] = "sourceNotOpen",
            [9] = "sourceOpen",
            [10] = "copiedValues",
        };

    public static string LinkStatusName(int status) =>
        LinkStatusNames.TryGetValue(status, out var name) ? name : $"unknown({status})";

    public static bool IsUpdatableLinkStatus(int status) =>
        status is 0 or 3 or 9;

    public static bool TryCalculationMode(string? mode, out int value)
    {
        value = XlCalculationAutomatic;
        return !string.IsNullOrWhiteSpace(mode) && CalculationModes.TryGetValue(mode.Trim(), out value);
    }

    public static bool TryPasteType(string? paste, out string normalized)
    {
        normalized = "all";
        if (string.IsNullOrWhiteSpace(paste)) return true;
        foreach (var token in PasteTypes)
        {
            if (string.Equals(token, paste.Trim(), StringComparison.OrdinalIgnoreCase))
            {
                normalized = token;
                return true;
            }
        }
        return false;
    }

    public static bool TryPasteOperation(string? operation, out int code)
    {
        code = 0;
        if (string.IsNullOrWhiteSpace(operation)) return true;
        return PasteOperations.TryGetValue(operation.Trim(), out code);
    }

    public static bool TrySparklineType(string? type, out int value)
    {
        value = XlSparkLine;
        return !string.IsNullOrWhiteSpace(type) && SparklineTypes.TryGetValue(type.Trim(), out value);
    }

    /// <summary>
    /// Resolve split offsets. A 1-based anchor cell means "rows above / columns
    /// left of the cell stay fixed", matching Window.SplitRow/SplitColumn semantics.
    /// </summary>
    public static bool TryResolveSplit(JsonObject op, out int splitRows, out int splitColumns, out string? error)
    {
        splitRows = 0;
        splitColumns = 0;
        error = null;
        var cell = Json.GetString(op, "cell");
        var hasCell = !string.IsNullOrWhiteSpace(cell);
        var hasRows = op.ContainsKey("splitRows");
        var hasCols = op.ContainsKey("splitColumns");
        if (hasCell && (hasRows || hasCols))
        {
            error = "cell and splitRows/splitColumns are mutually exclusive";
            return false;
        }
        if (hasCell)
        {
            if (!ExcelDataOperationsContract.TryParseCell(cell!.Trim(), out var row, out var col))
            {
                error = "cell must be a single A1 cell";
                return false;
            }
            splitRows = row - 1;
            splitColumns = col - 1;
            return true;
        }
        if (hasRows)
        {
            if (!ExcelDataOperationsContract.TryGetFiniteNumber(op["splitRows"], out var rows) ||
                rows != Math.Truncate(rows) || rows < 0 || rows > 1_048_575)
            {
                error = "splitRows must be an integer 0..1048575";
                return false;
            }
            splitRows = (int)rows;
        }
        if (hasCols)
        {
            if (!ExcelDataOperationsContract.TryGetFiniteNumber(op["splitColumns"], out var cols) ||
                cols != Math.Truncate(cols) || cols < 0 || cols > 16_383)
            {
                error = "splitColumns must be an integer 0..16383";
                return false;
            }
            splitColumns = (int)cols;
        }
        return true;
    }

    public static bool LinkRefersTo(string? formula, string leaf) =>
        !string.IsNullOrWhiteSpace(formula) && !string.IsNullOrWhiteSpace(leaf) &&
        formula!.IndexOf("[" + leaf + "]", StringComparison.OrdinalIgnoreCase) >= 0;

    public static string LeafOf(string source)
    {
        var text = (source ?? "").Replace('/', '\\');
        var index = text.LastIndexOf('\\');
        return index >= 0 ? text[(index + 1)..] : text;
    }

    public static void ValidateLinkSource(JsonObject op, int index, string opName, string field,
        ICollection<string> errors, bool mustExistOnDisk = false)
    {
        var source = Json.GetString(op, field);
        if (string.IsNullOrWhiteSpace(source))
        {
            errors.Add($"ops[{index}] '{opName}' {field} must be a non-empty link source");
            return;
        }
        if (source.Length > 1024)
            errors.Add($"ops[{index}] '{opName}' {field} exceeds 1024 characters");
        if (mustExistOnDisk && !File.Exists(source))
            errors.Add($"ops[{index}] '{opName}' {field} was not found on disk: '{source}'");
    }
}
