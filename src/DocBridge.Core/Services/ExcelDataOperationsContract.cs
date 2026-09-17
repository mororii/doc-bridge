using System.Globalization;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace DocBridge.Core.Services;

/// <summary>
/// Public input contract for Excel data/reporting ops. Validates without COM.
/// Integration Cursor copies RequiredFields/OptionalFields into OperationValidator
/// and calls <see cref="ValidatePublicInput"/> from ValidateExcelTarget.
/// </summary>
public static class ExcelDataOperationsContract
{
    public const string RestoreMode = "data-objects";
    public const int SnapshotVersion = 2;

    public const string CreateTable = "create_table";
    public const string ResizeTable = "resize_table";
    public const string StyleTable = "style_table";
    public const string SetTableTotals = "set_table_totals";
    public const string AddTableColumn = "add_table_column";
    public const string DeleteTable = "delete_table";
    public const string SortRange = "sort_range";
    public const string SortTable = "sort_table";
    public const string SetAutoFilter = "set_auto_filter";
    public const string ClearAutoFilter = "clear_auto_filter";
    public const string RemoveDuplicates = "remove_duplicates";
    public const string TextToColumns = "text_to_columns";
    public const string DefineName = "define_name";
    public const string UpdateName = "update_name";
    public const string DeleteName = "delete_name";
    public const string SetDataValidation = "set_data_validation";
    public const string ClearDataValidation = "clear_data_validation";
    public const string AddConditionalFormat = "add_conditional_format";
    public const string ClearConditionalFormats = "clear_conditional_formats";
    public const string SetRichText = "set_rich_text";
    public const string UpdateConditionalFormat = "update_conditional_format";
    public const string DeleteConditionalFormat = "delete_conditional_format";
    public const string AppendTableRows = "append_table_rows";
    public const string InsertTableRows = "insert_table_rows";
    public const string DeleteTableRows = "delete_table_rows";
    public const string CreateChart = "create_chart";
    public const string UpdateChart = "update_chart";
    public const string DeleteChart = "delete_chart";
    public const string InsertSheetPicture = "insert_sheet_picture";
    public const string UpdatePicture = "update_picture";
    public const string DeletePicture = "delete_picture";
    public const string SetCellNote = "set_cell_note";
    public const string ClearCellNote = "clear_cell_note";
    public const string SetHyperlink = "set_hyperlink";
    public const string ClearHyperlink = "clear_hyperlink";
    public const string CreatePivot = "create_pivot";
    public const string UpdatePivot = "update_pivot";
    public const string RefreshPivot = "refresh_pivot";
    public const string DeletePivot = "delete_pivot";
    public const string InsertShape = "insert_shape";
    public const string UpdateShape = "update_shape";
    public const string DeleteShape = "delete_shape";
    public const string InsertTextbox = "insert_textbox";
    public const string UpdateTextbox = "update_textbox";
    public const string DeleteTextbox = "delete_textbox";
    public const string CreateConnector = "create_connector";
    public const string UpdateConnector = "update_connector";
    public const string UpdateExternalLinks = "update_external_links";
    public const string ChangeLinkSource = "change_link_source";
    public const string BreakExternalLink = "break_external_link";
    public const string SetCalculationMode = "set_calculation_mode";
    public const string FreezeValues = "freeze_values";
    public const string PasteSpecial = "paste_special";
    public const string GoalSeek = "goal_seek";
    public const string ProtectWorkbook = "protect_workbook";
    public const string UnprotectWorkbook = "unprotect_workbook";
    public const string SetSplitPanes = "set_split_panes";
    public const string CreateSparkline = "create_sparkline";
    public const string UpdateSparkline = "update_sparkline";
    public const string DeleteSparkline = "delete_sparkline";
    public const string CreateSlicer = "create_slicer";
    public const string DeleteSlicer = "delete_slicer";
    public const string ApplyCellStyle = "apply_cell_style";

    public static readonly IReadOnlyList<string> WriteOpNames = new[]
    {
        CreateTable, ResizeTable, StyleTable, SetTableTotals, AddTableColumn, DeleteTable,
        SortRange, SortTable, SetAutoFilter, ClearAutoFilter, RemoveDuplicates, TextToColumns,
        DefineName, UpdateName, DeleteName,
        SetDataValidation, ClearDataValidation,
        AddConditionalFormat, ClearConditionalFormats,
        SetRichText, UpdateConditionalFormat, DeleteConditionalFormat,
        AppendTableRows, InsertTableRows, DeleteTableRows,
        CreateChart, UpdateChart, DeleteChart,
        InsertSheetPicture, UpdatePicture, DeletePicture,
        SetCellNote, ClearCellNote, SetHyperlink, ClearHyperlink,
        CreatePivot, UpdatePivot, RefreshPivot, DeletePivot,
        InsertShape, UpdateShape, DeleteShape,
        InsertTextbox, UpdateTextbox, DeleteTextbox, CreateConnector, UpdateConnector,
        UpdateExternalLinks, ChangeLinkSource, BreakExternalLink,
        SetCalculationMode, FreezeValues, PasteSpecial, GoalSeek,
        ProtectWorkbook, UnprotectWorkbook, SetSplitPanes,
        CreateSparkline, UpdateSparkline, DeleteSparkline,
        CreateSlicer, DeleteSlicer, ApplyCellStyle,
    };

    public static readonly IReadOnlyList<string> ReadScopes = new[]
    {
        "tables", "charts", "pictures", "names", "validations",
        "conditionalFormats", "notes", "hyperlinks", "filters",
        "pivots", "shapes", "richText", "connectors",
        "links", "sparklines", "slicers", "cellStyles",
    };

    /// <summary>
    /// Workbook-level ops never assume an active sheet and must not require
    /// target.sheet. They bind the resolved target workbook directly.
    /// </summary>
    public static bool IsWorkbookLevelOp(string? op) =>
        op is UpdateExternalLinks or ChangeLinkSource or BreakExternalLink
            or SetCalculationMode or ProtectWorkbook or UnprotectWorkbook;

    public static readonly IReadOnlySet<string> WriteOpSet =
        new HashSet<string>(WriteOpNames, StringComparer.OrdinalIgnoreCase);

    public static readonly Dictionary<string, (string Field, string Type)[]> RequiredFields =
        new(StringComparer.OrdinalIgnoreCase)
        {
            [CreateTable] = new[] { ("range", "string") },
            [ResizeTable] = new[] { ("name", "string"), ("range", "string") },
            [StyleTable] = new[] { ("name", "string") },
            [SetTableTotals] = new[] { ("name", "string"), ("showTotals", "bool") },
            [AddTableColumn] = new[] { ("name", "string"), ("columnName", "string") },
            [DeleteTable] = new[] { ("name", "string") },
            [SortRange] = new[] { ("range", "string"), ("keys", "array") },
            [SortTable] = new[] { ("name", "string"), ("keys", "array") },
            [RemoveDuplicates] = new[] { ("range", "string") },
            [TextToColumns] = new[] { ("range", "string") },
            [SetAutoFilter] = new[] { ("criteria", "array") },
            [ClearAutoFilter] = Array.Empty<(string, string)>(),
            [DefineName] = new[] { ("name", "string"), ("refersTo", "string"), ("scope", "string") },
            [UpdateName] = new[] { ("name", "string"), ("refersTo", "string"), ("scope", "string") },
            [DeleteName] = new[] { ("name", "string"), ("scope", "string") },
            [SetDataValidation] = new[] { ("range", "string"), ("type", "string") },
            [ClearDataValidation] = new[] { ("range", "string") },
            [AddConditionalFormat] = new[] { ("range", "string"), ("rule", "object") },
            [ClearConditionalFormats] = new[] { ("range", "string") },
            [SetRichText] = new[] { ("cell", "string"), ("baseFont", "object"), ("runs", "array") },
            [UpdateConditionalFormat] = new[] { ("range", "string"), ("index", "int"), ("expectedFingerprint", "string") },
            [DeleteConditionalFormat] = new[] { ("range", "string"), ("index", "int"), ("expectedFingerprint", "string") },
            [AppendTableRows] = new[] { ("name", "string"), ("rows", "array") },
            [InsertTableRows] = new[] { ("name", "string"), ("index", "int"), ("rows", "array") },
            [DeleteTableRows] = new[] { ("name", "string"), ("index", "int"), ("count", "int") },
            [CreateChart] = new[] { ("sourceRange", "string"), ("chartType", "string") },
            [UpdateChart] = new[] { ("name", "string") },
            [DeleteChart] = new[] { ("name", "string") },
            [InsertSheetPicture] = new[] { ("path", "string") },
            [UpdatePicture] = new[] { ("name", "string") },
            [DeletePicture] = new[] { ("name", "string") },
            [SetCellNote] = new[] { ("text", "string") },
            [ClearCellNote] = Array.Empty<(string, string)>(),
            [SetHyperlink] = Array.Empty<(string, string)>(),
            [ClearHyperlink] = Array.Empty<(string, string)>(),
            [CreatePivot] = new[] { ("sourceRange", "string"), ("destination", "string"), ("name", "string") },
            [UpdatePivot] = new[] { ("name", "string") },
            [RefreshPivot] = new[] { ("name", "string") },
            [DeletePivot] = new[] { ("name", "string") },
            [InsertShape] = new[] { ("shapeType", "string") },
            [UpdateShape] = new[] { ("name", "string") },
            [DeleteShape] = new[] { ("name", "string") },
            [InsertTextbox] = Array.Empty<(string, string)>(),
            [UpdateTextbox] = new[] { ("name", "string") },
            [DeleteTextbox] = new[] { ("name", "string") },
            [CreateConnector] = new[] { ("connectorType", "string"), ("begin", "object"), ("end", "object") },
            [UpdateConnector] = new[] { ("name", "string") },
            [UpdateExternalLinks] = Array.Empty<(string, string)>(),
            [ChangeLinkSource] = new[] { ("source", "string"), ("newSource", "string") },
            [BreakExternalLink] = new[] { ("source", "string") },
            [SetCalculationMode] = new[] { ("mode", "string") },
            [FreezeValues] = new[] { ("range", "string") },
            [PasteSpecial] = new[] { ("sourceRange", "string"), ("destination", "string") },
            [GoalSeek] = new[] { ("cell", "string"), ("goalCell", "string"), ("goal", "number") },
            [ProtectWorkbook] = Array.Empty<(string, string)>(),
            [UnprotectWorkbook] = Array.Empty<(string, string)>(),
            [SetSplitPanes] = Array.Empty<(string, string)>(),
            [CreateSparkline] = new[] { ("location", "string"), ("sourceData", "string"), ("type", "string") },
            [UpdateSparkline] = new[] { ("location", "string") },
            [DeleteSparkline] = new[] { ("location", "string") },
            [CreateSlicer] = new[] { ("source", "string"), ("field", "string"), ("name", "string") },
            [DeleteSlicer] = new[] { ("name", "string") },
            [ApplyCellStyle] = new[] { ("range", "string"), ("styleName", "string") },
        };

    public static readonly Dictionary<string, string[]> OptionalFields =
        new(StringComparer.OrdinalIgnoreCase)
        {
            [CreateTable] = new[] { "target", "targetWorkbook", "name", "hasHeaders", "styleName", "showTotals", "totals" },
            [ResizeTable] = new[] { "target", "targetWorkbook" },
            [StyleTable] = new[] { "target", "targetWorkbook", "styleName", "showHeaders", "showTotals", "showAutoFilter", "showRowStripes", "showColumnStripes" },
            [SetTableTotals] = new[] { "target", "targetWorkbook", "columns" },
            [AddTableColumn] = new[] { "target", "targetWorkbook", "formula", "insertAfter" },
            [DeleteTable] = new[] { "target", "targetWorkbook" },
            [SortRange] = new[] { "target", "targetWorkbook", "hasHeaders" },
            [SortTable] = new[] { "target", "targetWorkbook" },
            [RemoveDuplicates] = new[] { "target", "targetWorkbook", "columns", "hasHeaders" },
            [TextToColumns] = new[] { "target", "targetWorkbook", "destination", "dataType", "comma", "tab", "semicolon", "space", "other", "otherChar", "consecutiveDelimiter", "textQualifier" },
            [SetAutoFilter] = new[] { "target", "targetWorkbook", "range", "name" },
            [ClearAutoFilter] = new[] { "target", "targetWorkbook", "range", "name" },
            [DefineName] = new[] { "target", "targetWorkbook", "comment", "replace" },
            [UpdateName] = new[] { "target", "targetWorkbook", "comment" },
            [DeleteName] = new[] { "target", "targetWorkbook" },
            [SetDataValidation] = new[]
            {
                "target", "targetWorkbook", "operator", "formula1", "formula2", "source",
                "inCellDropdown", "ignoreBlank", "showInput", "inputTitle", "inputMessage",
                "showError", "errorTitle", "errorMessage", "errorStyle",
            },
            [ClearDataValidation] = new[] { "target", "targetWorkbook" },
            [AddConditionalFormat] = new[] { "target", "targetWorkbook", "style" },
            [ClearConditionalFormats] = new[] { "target", "targetWorkbook" },
            [SetRichText] = new[] { "target", "targetWorkbook" },
            [UpdateConditionalFormat] = new[] { "target", "targetWorkbook", "rule", "style" },
            [DeleteConditionalFormat] = new[] { "target", "targetWorkbook" },
            [AppendTableRows] = new[] { "target", "targetWorkbook" },
            [InsertTableRows] = new[] { "target", "targetWorkbook" },
            [DeleteTableRows] = new[] { "target", "targetWorkbook" },
            [CreateChart] = new[] { "target", "targetWorkbook", "name", "title", "hasLegend", "legendPosition", "plotBy", "position", "series", "chartFill", "plotFill", "chartBorder", "plotBorder", "axes", "plotArea", "dataLabels" },
            [UpdateChart] = new[] { "target", "targetWorkbook", "sourceRange", "chartType", "title", "hasLegend", "legendPosition", "plotBy", "position", "series", "chartFill", "plotFill", "chartBorder", "plotBorder", "axes", "plotArea", "dataLabels" },
            [DeleteChart] = new[] { "target", "targetWorkbook" },
            [InsertSheetPicture] = new[] { "target", "targetWorkbook", "name", "position", "lockAspectRatio" },
            [UpdatePicture] = new[] { "target", "targetWorkbook", "position", "lockAspectRatio", "path", "crop" },
            [DeletePicture] = new[] { "target", "targetWorkbook", "path" },
            [SetCellNote] = new[] { "target", "targetWorkbook", "range", "cell", "visible" },
            [ClearCellNote] = new[] { "target", "targetWorkbook", "range", "cell" },
            [SetHyperlink] = new[] { "target", "targetWorkbook", "range", "cell", "address", "subAddress", "textToDisplay", "screenTip" },
            [ClearHyperlink] = new[] { "target", "targetWorkbook", "range", "cell" },
            [CreatePivot] = new[] { "target", "targetWorkbook", "rows", "columns", "values", "filters" },
            [UpdatePivot] = new[] { "target", "targetWorkbook", "rows", "columns", "values", "filters", "sourceRange" },
            [RefreshPivot] = new[] { "target", "targetWorkbook" },
            [DeletePivot] = new[] { "target", "targetWorkbook" },
            [InsertShape] = new[] { "target", "targetWorkbook", "name", "position", "text", "fillColor", "lineColor", "lineWeight", "lineVisible", "rotation", "font" },
            [UpdateShape] = new[] { "target", "targetWorkbook", "position", "text", "fillColor", "lineColor", "lineWeight", "lineVisible", "rotation", "font" },
            [DeleteShape] = new[] { "target", "targetWorkbook" },
            [InsertTextbox] = new[] { "target", "targetWorkbook", "name", "text", "position", "fillColor", "lineColor", "lineWeight", "lineVisible", "rotation", "font" },
            [UpdateTextbox] = new[] { "target", "targetWorkbook", "text", "position", "fillColor", "lineColor", "lineWeight", "lineVisible", "rotation", "font" },
            [DeleteTextbox] = new[] { "target", "targetWorkbook" },
            [CreateConnector] = new[] { "target", "targetWorkbook", "name", "lineColor", "lineWeight", "lineVisible" },
            [UpdateConnector] = new[] { "target", "targetWorkbook", "begin", "end", "lineColor", "lineWeight", "lineVisible" },
            [UpdateExternalLinks] = new[] { "targetWorkbook", "source", "sourceContains" },
            [ChangeLinkSource] = new[] { "targetWorkbook" },
            [BreakExternalLink] = new[] { "targetWorkbook" },
            [SetCalculationMode] = new[] { "targetWorkbook", "recalculate" },
            [FreezeValues] = new[] { "target", "targetWorkbook" },
            [PasteSpecial] = new[] { "target", "targetWorkbook", "paste", "transpose", "skipBlanks", "operation" },
            [GoalSeek] = new[] { "target", "targetWorkbook", "tolerance" },
            [ProtectWorkbook] = new[] { "targetWorkbook", "structure", "windows" },
            [UnprotectWorkbook] = new[] { "targetWorkbook" },
            [SetSplitPanes] = new[] { "target", "targetWorkbook", "cell", "splitRows", "splitColumns", "freeze", "remove" },
            [CreateSparkline] = new[] { "target", "targetWorkbook", "markers", "lineColor", "showHigh", "showLow", "showNegative", "showFirst", "showLast" },
            [UpdateSparkline] = new[] { "target", "targetWorkbook", "sourceData", "type", "markers", "lineColor", "showHigh", "showLow", "showNegative", "showFirst", "showLast" },
            [DeleteSparkline] = new[] { "target", "targetWorkbook" },
            [CreateSlicer] = new[] { "target", "targetWorkbook", "caption", "position" },
            [DeleteSlicer] = new[] { "target", "targetWorkbook", "source", "field" },
            [ApplyCellStyle] = new[] { "target", "targetWorkbook" },
        };

    public static readonly Dictionary<(string Op, string Field), string> FieldExpectations = new()
    {
        [(CreateTable, "range")] = "rectangular A1 range, e.g. A1:G20 or '자재대장'!A1:G20",
        [(SortRange, "keys")] = "array of 1..3 objects {column, order:asc|desc}",
        [(SortTable, "keys")] = "array of 1..3 objects {column, order:asc|desc}",
        [(SetAutoFilter, "criteria")] = "array of {column, operator, value?|value2?|values?}",
        [(SetDataValidation, "type")] = "list|whole|decimal|date|time|textLength|custom",
        [(AddConditionalFormat, "rule")] = "object {type:cellValue|expression|uniqueValues|duplicateValues|colorScale|dataBar|iconSet,...}",
        [(SetRichText, "cell")] = "single A1 cell; strings are written literally, never as formulas",
        [(SetRichText, "baseFont")] = "object {name?,size?,bold?,italic?,underline?,color?}",
        [(SetRichText, "runs")] = "ordered 1..256 {text,font?}; combined UTF-16 length <= 32767",
        [(UpdateConditionalFormat, "expectedFingerprint")] = "fingerprint returned by conditionalFormats inspect inventory",
        [(AppendTableRows, "rows")] = "1..1000 rectangular rows with exactly the table column count",
        [(CreateChart, "chartType")] = "columnClustered|columnStacked|barClustered|barStacked|line|lineMarkers|pie|area|scatter|scatterLines",
        [(CreateChart, "series")] = "array of {index?:1..n contiguous, values?, categories?, name?, range?, chartType?, axisGroup:primary|secondary, lineColor?, lineWeight:0.25..10, marker:none|automatic|circle|dash|diamond|dot|square|triangle}",
        [(UpdateChart, "series")] = "array of {index?:1..64, values?, categories?, name?, range?, chartType?, axisGroup:primary|secondary, lineColor?, lineWeight:0.25..10, marker:none|automatic|circle|dash|diamond|dot|square|triangle}",
        [(CreateChart, "sourceRange")] = "rectangular or union A1; may be sheet-qualified independently of target.sheet",
        [(UpdatePivot, "sourceRange")] = "expanded A1, sheet-qualified A1, or ListObject name; shared caches get a private cache",
        [(CreatePivot, "values")] = "array of {field, function, caption?, numberFormat?} — multiple aggregates per source field allowed",
        [(CreateChart, "axes")] = "{category|value|valueSecondary:{group:primary|secondary, visible?, minimum?, maximum?, numberFormat?, title?}}",
        [(UpdateChart, "axes")] = "{category|value|valueSecondary:{group:primary|secondary, visible?, minimum?, maximum?, numberFormat?, title?}}",
        [(CreateChart, "plotArea")] = "{spanChart:true} or {left,top,width,height} points",
        [(CreatePivot, "destination")] = "top-left A1 cell for the pivot table, e.g. A1 or '원가'!A1",
        [(InsertShape, "shapeType")] = "rectangle|roundedRectangle|oval|rightArrow|downArrow",
        [(TextToColumns, "dataType")] = "delimited (fixed is rejected until FieldInfo/column breaks exist)",
        [(AddTableColumn, "columnName")] = "new ListColumn header",
        [(AddTableColumn, "insertAfter")] = "finite 1-based ListColumns index to insert after",
        [(UpdateExternalLinks, "sourceContains")] = "optional substring filter over link sources; omit to update all Excel links",
        [(ChangeLinkSource, "source")] = "existing link source (full path or unique leaf/filename)",
        [(ChangeLinkSource, "newSource")] = "absolute existing workbook path that replaces the link source",
        [(BreakExternalLink, "source")] = "existing link source (full path or unique leaf/filename)",
        [(SetCalculationMode, "mode")] = "manual|automatic|semiautomatic",
        [(SetCalculationMode, "recalculate")] = "none|full|fullRebuild after switching the mode",
        [(FreezeValues, "range")] = "rectangular A1, at most 20000 cells; formulas become cached values in place",
        [(PasteSpecial, "sourceRange")] = "rectangular A1; may be sheet-qualified independently of target.sheet",
        [(PasteSpecial, "destination")] = "top-left A1 anchor or a fitting range; must match target.sheet",
        [(PasteSpecial, "paste")] = "values|formulas|formats|all (default all); all/formats use native Copy so borders transfer, values/formulas keep destination formats",
        [(PasteSpecial, "operation")] = "none|add|subtract|multiply|divide applied to existing numeric destination values",
        [(GoalSeek, "cell")] = "single changing cell",
        [(GoalSeek, "goalCell")] = "single formula cell whose value must reach goal",
        [(GoalSeek, "goal")] = "finite target number",
        [(ProtectWorkbook, "structure")] = "protect workbook structure (default true)",
        [(ProtectWorkbook, "windows")] = "protect workbook windows (default false)",
        [(SetSplitPanes, "cell")] = "A1 anchor; rows above and columns left of it stay fixed (mutually exclusive with splitRows/splitColumns)",
        [(CreateSparkline, "location")] = "single cell, single row, or single column range that hosts the sparkline group",
        [(CreateSparkline, "sourceData")] = "rectangular data range for the sparkline group",
        [(CreateSparkline, "type")] = "line|column (winLoss is rejected until a native mapping is proven)",
        [(CreateSlicer, "source")] = "ListObject or PivotTable name in the same workbook",
        [(CreateSlicer, "field")] = "table column header or pivot field caption",
        [(ApplyCellStyle, "styleName")] = "existing workbook cell style name (see cellStyles inspect scope)",
        [(ApplyCellStyle, "range")] = "rectangular A1, at most 5000 cells",
        [(InsertSheetPicture, "path")] = "absolute existing image path (png/jpg/jpeg/gif/bmp/emf/wmf/tif/tiff)",
        [(DefineName, "refersTo")] = "Excel formula starting with '=', e.g. ='자재대장'!$G$2:$G$20",
        [(SetTableTotals, "columns")] = "array of {column, function:none|sum|count|average|max|min|countNums|stdDev|var|custom, formula?}",
    };

    private static readonly Regex DefinedNamePattern = new(
        @"^[\p{L}_\\][\p{L}\p{Nd}._\\]*$", RegexOptions.Compiled);

    public static bool IsDataOperation(string? op) =>
        !string.IsNullOrWhiteSpace(op) && WriteOpSet.Contains(op);

    public static bool IsWorkbookScopedNameOp(JsonObject op)
    {
        var name = Json.GetString(op, "op");
        if (name is not (DefineName or UpdateName or DeleteName)) return false;
        return string.Equals(Json.GetString(op, "scope"), "workbook", StringComparison.OrdinalIgnoreCase);
    }

    public static bool RequiresExplicitSheet(JsonObject op) =>
        !IsWorkbookScopedNameOp(op) && !IsWorkbookLevelOp(Json.GetString(op, "op"));

    public static bool ValidatePublicInput(JsonObject op, int index, ICollection<string> errors)
    {
        var before = errors.Count;
        var opName = Json.GetString(op, "op") ?? "";
        if (!IsDataOperation(opName))
        {
            errors.Add($"ops[{index}] '{opName}' is not an Excel data/reporting operation");
            return false;
        }

        if (RequiredFields.TryGetValue(opName, out var rules))
        {
            foreach (var (field, type) in rules)
                ValidateRequiredField(op, index, opName, field, type, errors);
        }

        ValidateSheetTarget(op, index, opName, errors);

        switch (opName)
        {
            case CreateTable:
                ValidateA1Field(op, index, opName, "range", errors, allowUnion: false);
                ValidateOptionalObjectName(op, index, opName, "name", errors, definedNameRules: true);
                ValidateOptionalStyleName(op, index, opName, errors);
                ValidateTotalsColumns(op, index, opName, errors, required: false);
                ValidateOptionalBool(op, index, opName, "hasHeaders", errors);
                ValidateOptionalBool(op, index, opName, "showTotals", errors);
                break;
            case ResizeTable:
                ValidateObjectName(op, index, opName, "name", errors, definedNameRules: true);
                ValidateA1Field(op, index, opName, "range", errors, allowUnion: false);
                break;
            case StyleTable:
                ValidateObjectName(op, index, opName, "name", errors, definedNameRules: true);
                ValidateOptionalStyleName(op, index, opName, errors);
                foreach (var flag in new[] { "showHeaders", "showTotals", "showAutoFilter", "showRowStripes", "showColumnStripes" })
                    ValidateOptionalBool(op, index, opName, flag, errors);
                if (!op.ContainsKey("styleName") &&
                    !op.ContainsKey("showHeaders") && !op.ContainsKey("showTotals") &&
                    !op.ContainsKey("showAutoFilter") && !op.ContainsKey("showRowStripes") &&
                    !op.ContainsKey("showColumnStripes"))
                    errors.Add($"ops[{index}] '{opName}' requires styleName or at least one show* flag");
                break;
            case SetTableTotals:
                ValidateObjectName(op, index, opName, "name", errors, definedNameRules: true);
                ValidateTotalsColumns(op, index, opName, errors, required: Json.GetBool(op, "showTotals"));
                break;
            case SortRange:
                ValidateA1Field(op, index, opName, "range", errors, allowUnion: false);
                ValidateSortKeys(op, index, opName, errors, table: false);
                ValidateOptionalBool(op, index, opName, "hasHeaders", errors);
                break;
            case SortTable:
                ValidateObjectName(op, index, opName, "name", errors, definedNameRules: true);
                ValidateSortKeys(op, index, opName, errors, table: true);
                break;
            case SetAutoFilter:
                ValidateRangeXorName(op, index, opName, errors);
                ValidateFilterCriteria(op, index, opName, errors);
                break;
            case ClearAutoFilter:
                ValidateRangeXorName(op, index, opName, errors);
                break;
            case DefineName:
            case UpdateName:
                ValidateDefinedName(op, index, opName, errors);
                ValidateNameScope(op, index, opName, errors);
                ValidateRefersTo(op, index, opName, errors);
                ValidateOptionalBool(op, index, opName, "replace", errors);
                ValidateOptionalBoundedString(op, index, opName, "comment", 255, errors);
                break;
            case DeleteName:
                ValidateDefinedName(op, index, opName, errors);
                ValidateNameScope(op, index, opName, errors);
                break;
            case SetDataValidation:
                ValidateA1Field(op, index, opName, "range", errors, allowUnion: false);
                ValidateDataValidation(op, index, opName, errors);
                break;
            case ClearDataValidation:
            case ClearConditionalFormats:
                ValidateA1Field(op, index, opName, "range", errors, allowUnion: false);
                break;
            case AddConditionalFormat:
                ValidateA1Field(op, index, opName, "range", errors, allowUnion: false);
                ValidateConditionalRule(op, index, opName, errors);
                ValidateConditionalStyle(op, index, opName, errors);
                break;
            case SetRichText:
                ValidateCellAddress(op, index, opName, errors);
                ExcelRichTextContract.Validate(op, index, errors);
                break;
            case UpdateConditionalFormat:
                ValidateA1Field(op, index, opName, "range", errors, allowUnion: false);
                ValidateOptionalPositiveInt(op, index, opName, "index", 64, errors);
                ValidateRequiredBoundedString(op, index, opName, "expectedFingerprint", 4096, errors);
                if (!op.ContainsKey("rule") && !op.ContainsKey("style"))
                    errors.Add($"ops[{index}] '{opName}' requires rule and/or style");
                if (op.ContainsKey("rule")) ValidateConditionalRule(op, index, opName, errors);
                if (op.ContainsKey("style")) ValidateConditionalStyle(op, index, opName, errors);
                break;
            case DeleteConditionalFormat:
                ValidateA1Field(op, index, opName, "range", errors, allowUnion: false);
                ValidateOptionalPositiveInt(op, index, opName, "index", 64, errors);
                ValidateRequiredBoundedString(op, index, opName, "expectedFingerprint", 4096, errors);
                break;
            case AppendTableRows:
            case InsertTableRows:
                ValidateObjectName(op, index, opName, "name", errors, definedNameRules: true);
                if (opName == InsertTableRows) ValidateOptionalPositiveInt(op, index, opName, "index", 1_048_576, errors);
                ExcelTableRowsContract.ValidateRows(op, index, errors);
                break;
            case DeleteTableRows:
                ValidateObjectName(op, index, opName, "name", errors, definedNameRules: true);
                ValidateOptionalPositiveInt(op, index, opName, "index", 1_048_576, errors);
                ValidateOptionalPositiveInt(op, index, opName, "count", 1000, errors);
                break;
            case CreateChart:
                ValidateA1Field(op, index, opName, "sourceRange", errors, allowUnion: true);
                ValidateChartType(op, index, opName, errors);
                ValidateOptionalObjectName(op, index, opName, "name", errors, definedNameRules: false);
                ValidateOptionalBoundedString(op, index, opName, "title", 255, errors);
                ValidateOptionalBool(op, index, opName, "hasLegend", errors);
                ValidateOptionalLegend(op, index, opName, errors);
                ValidateOptionalPlotBy(op, index, opName, errors);
                ValidateOptionalPosition(op, index, opName, errors, requireSize: true);
                ValidateChartPresentation(op, index, opName, errors);
                break;
            case UpdateChart:
                ValidateObjectName(op, index, opName, "name", errors, definedNameRules: false);
                if (op.ContainsKey("sourceRange"))
                    ValidateA1Field(op, index, opName, "sourceRange", errors, allowUnion: true);
                if (op.ContainsKey("chartType")) ValidateChartType(op, index, opName, errors);
                ValidateOptionalBoundedString(op, index, opName, "title", 255, errors);
                ValidateOptionalBool(op, index, opName, "hasLegend", errors);
                ValidateOptionalLegend(op, index, opName, errors);
                ValidateOptionalPlotBy(op, index, opName, errors);
                ValidateOptionalPosition(op, index, opName, errors, requireSize: false);
                ValidateChartPresentation(op, index, opName, errors);
                if (!HasAnyUpdate(op, "sourceRange", "chartType", "title", "hasLegend", "legendPosition", "plotBy",
                        "position", "series", "chartFill", "plotFill", "chartBorder", "plotBorder", "axes", "plotArea", "dataLabels"))
                    errors.Add($"ops[{index}] '{opName}' requires at least one updatable field");
                break;
            case InsertSheetPicture:
                ValidatePicturePath(op, index, opName, errors);
                ValidateOptionalObjectName(op, index, opName, "name", errors, definedNameRules: false);
                ValidateOptionalPosition(op, index, opName, errors, requireSize: false);
                ValidateOptionalBool(op, index, opName, "lockAspectRatio", errors);
                break;
            case UpdatePicture:
                ValidateObjectName(op, index, opName, "name", errors, definedNameRules: false);
                ValidateOptionalPosition(op, index, opName, errors, requireSize: false);
                ValidateOptionalBool(op, index, opName, "lockAspectRatio", errors);
                ExcelDrawingDetailsContract.ValidatePictureDetails(op, index, opName, errors);
                if (!op.ContainsKey("position") && !op.ContainsKey("lockAspectRatio") && !op.ContainsKey("crop") && !op.ContainsKey("path"))
                    errors.Add($"ops[{index}] '{opName}' requires position, lockAspectRatio, crop, and/or path");
                break;
            case CreateConnector:
            case UpdateConnector:
                ExcelDrawingDetailsContract.ValidateConnector(op, index, opName, errors);
                break;
            case UpdateExternalLinks:
                ValidateOptionalBoundedString(op, index, opName, "source", 1024, errors);
                ValidateOptionalBoundedString(op, index, opName, "sourceContains", 1024, errors);
                break;
            case ChangeLinkSource:
                ExcelWorkbookOpsContract.ValidateLinkSource(op, index, opName, "source", errors);
                ExcelWorkbookOpsContract.ValidateLinkSource(op, index, opName, "newSource", errors,
                    mustExistOnDisk: true);
                if (!Path.IsPathFullyQualified(Json.GetString(op, "newSource") ?? ""))
                    errors.Add($"ops[{index}] '{opName}' newSource must be an absolute local path");
                break;
            case BreakExternalLink:
                ExcelWorkbookOpsContract.ValidateLinkSource(op, index, opName, "source", errors);
                break;
            case SetCalculationMode:
                if (!ExcelWorkbookOpsContract.TryCalculationMode(Json.GetString(op, "mode"), out _))
                    errors.Add($"ops[{index}] '{opName}' mode must be manual|automatic|semiautomatic");
                if (op.ContainsKey("recalculate") &&
                    !ExcelWorkbookOpsContract.RecalculateTokens.Contains(Json.GetString(op, "recalculate") ?? ""))
                    errors.Add($"ops[{index}] '{opName}' recalculate must be none|full|fullRebuild");
                break;
            case FreezeValues:
                ValidateA1Field(op, index, opName, "range", errors, allowUnion: false);
                ValidateBoundedCellCount(op, index, opName, "range", ExcelWorkbookOpsContract.MaxValueSurfaceCells, errors);
                break;
            case PasteSpecial:
                ValidateA1Field(op, index, opName, "sourceRange", errors, allowUnion: false);
                ValidateA1Field(op, index, opName, "destination", errors, allowUnion: false);
                ValidateBoundedCellCount(op, index, opName, "sourceRange", ExcelWorkbookOpsContract.MaxPasteCells, errors);
                if (op.ContainsKey("paste") &&
                    !ExcelWorkbookOpsContract.TryPasteType(Json.GetString(op, "paste"), out _))
                    errors.Add($"ops[{index}] '{opName}' paste must be values|formulas|formats|all");
                ValidateOptionalBool(op, index, opName, "transpose", errors);
                ValidateOptionalBool(op, index, opName, "skipBlanks", errors);
                if (op.ContainsKey("operation") &&
                    !ExcelWorkbookOpsContract.TryPasteOperation(Json.GetString(op, "operation"), out _))
                    errors.Add($"ops[{index}] '{opName}' operation must be none|add|subtract|multiply|divide");
                ExcelWorkbookOpsContract.TryPasteType(Json.GetString(op, "paste"), out var pasteToken);
                ExcelWorkbookOpsContract.TryPasteOperation(Json.GetString(op, "operation"), out var pasteOperation);
                if (Json.GetBool(op, "transpose") && pasteToken is "formats" or "all")
                    errors.Add($"ops[{index}] '{opName}' transpose supports paste=values|formulas only");
                if (pasteOperation != 0 && pasteToken is "formulas" or "formats")
                    errors.Add($"ops[{index}] '{opName}' operation requires paste=values|all");
                break;
            case GoalSeek:
                ValidateGoalSeekCell(op, index, opName, "cell", errors);
                ValidateGoalSeekCell(op, index, opName, "goalCell", errors);
                if (!TryGetFiniteNumber(op["goal"], out _))
                    errors.Add($"ops[{index}] '{opName}' goal must be a finite number");
                if (op.ContainsKey("tolerance"))
                {
                    if (!TryGetFiniteNumber(op["tolerance"], out var tolerance) || tolerance <= 0 || tolerance > 1)
                        errors.Add($"ops[{index}] '{opName}' tolerance must be a finite number 0..1 (relative)");
                }
                ValidateGoalSeekSheets(op, index, opName, errors);
                break;
            case ProtectWorkbook:
                ValidateOptionalBool(op, index, opName, "structure", errors);
                ValidateOptionalBool(op, index, opName, "windows", errors);
                break;
            case UnprotectWorkbook:
                break;
            case SetSplitPanes:
                ValidateOptionalBool(op, index, opName, "freeze", errors);
                ValidateOptionalBool(op, index, opName, "remove", errors);
                if (!ExcelWorkbookOpsContract.TryResolveSplit(op, out _, out _,
                        out var splitError))
                    errors.Add($"ops[{index}] '{opName}' {splitError}");
                else if (Json.GetBool(op, "remove") && (op.ContainsKey("cell") || op.ContainsKey("splitRows") || op.ContainsKey("splitColumns")))
                    errors.Add($"ops[{index}] '{opName}' remove cannot be combined with cell/splitRows/splitColumns");
                break;
            case CreateSparkline:
                ValidateA1Field(op, index, opName, "location", errors, allowUnion: false);
                ValidateA1Field(op, index, opName, "sourceData", errors, allowUnion: false);
                ValidateSparklineType(op, index, opName, errors);
                ValidateSparklineOptions(op, index, opName, errors);
                ValidateSparklineSheets(op, index, opName, errors);
                break;
            case UpdateSparkline:
                ValidateA1Field(op, index, opName, "location", errors, allowUnion: false);
                if (op.ContainsKey("sourceData"))
                    ValidateA1Field(op, index, opName, "sourceData", errors, allowUnion: false);
                if (op.ContainsKey("type")) ValidateSparklineType(op, index, opName, errors);
                ValidateSparklineOptions(op, index, opName, errors);
                ValidateSparklineSheets(op, index, opName, errors);
                if (!HasAnyUpdate(op, "sourceData", "type", "markers", "lineColor",
                        "showHigh", "showLow", "showNegative", "showFirst", "showLast"))
                    errors.Add($"ops[{index}] '{opName}' requires at least one updatable field");
                break;
            case DeleteSparkline:
                ValidateA1Field(op, index, opName, "location", errors, allowUnion: false);
                break;
            case CreateSlicer:
                ValidateObjectName(op, index, opName, "source", errors, definedNameRules: false);
                ValidateOptionalBoundedString(op, index, opName, "field", 255, errors);
                if (string.IsNullOrWhiteSpace(Json.GetString(op, "field")))
                    errors.Add($"ops[{index}] '{opName}' field must be a non-empty column/field name");
                ValidateObjectName(op, index, opName, "name", errors, definedNameRules: false);
                ValidateOptionalBoundedString(op, index, opName, "caption", 255, errors);
                ValidateOptionalPosition(op, index, opName, errors, requireSize: false);
                break;
            case DeleteSlicer:
                ValidateObjectName(op, index, opName, "name", errors, definedNameRules: false);
                if (op.ContainsKey("source") || op.ContainsKey("field"))
                {
                    ValidateOptionalObjectName(op, index, opName, "source", errors, definedNameRules: false);
                    ValidateOptionalBoundedString(op, index, opName, "field", 255, errors);
                    if (string.IsNullOrWhiteSpace(Json.GetString(op, "source")) ||
                        string.IsNullOrWhiteSpace(Json.GetString(op, "field")))
                        errors.Add($"ops[{index}] '{opName}' source and field are required together for exact restore");
                }
                break;
            case ApplyCellStyle:
                ValidateA1Field(op, index, opName, "range", errors, allowUnion: false);
                ValidateBoundedCellCount(op, index, opName, "range", ExcelWorkbookOpsContract.MaxCellStyleCells, errors);
                ValidateRequiredBoundedString(op, index, opName, "styleName", 255, errors);
                break;
            case SetCellNote:
                ValidateCellAddress(op, index, opName, errors);
                ValidateRequiredBoundedString(op, index, opName, "text", 32_767, errors);
                ValidateOptionalBool(op, index, opName, "visible", errors);
                break;
            case ClearCellNote:
            case ClearHyperlink:
                ValidateCellAddress(op, index, opName, errors);
                break;
            case SetHyperlink:
                ValidateCellAddress(op, index, opName, errors);
                var address = Json.GetString(op, "address");
                var sub = Json.GetString(op, "subAddress");
                if (string.IsNullOrWhiteSpace(address) && string.IsNullOrWhiteSpace(sub))
                    errors.Add($"ops[{index}] '{opName}' requires address and/or subAddress");
                ValidateOptionalBoundedString(op, index, opName, "address", 2048, errors);
                ValidateOptionalBoundedString(op, index, opName, "subAddress", 255, errors);
                ValidateOptionalBoundedString(op, index, opName, "textToDisplay", 255, errors);
                ValidateOptionalBoundedString(op, index, opName, "screenTip", 255, errors);
                break;
            case AddTableColumn:
                ValidateObjectName(op, index, opName, "name", errors, definedNameRules: true);
                ValidateOptionalBoundedString(op, index, opName, "columnName", 255, errors);
                if (op.ContainsKey("formula"))
                {
                    var formula = Json.GetString(op, "formula");
                    if (string.IsNullOrWhiteSpace(formula) || !formula.StartsWith('='))
                        errors.Add($"ops[{index}] '{opName}' formula must start with '='");
                }
                ValidateOptionalPositiveInt(op, index, opName, "insertAfter", 16_384, errors);
                break;
            case DeleteTable:
            case DeleteChart:
            case DeletePicture:
            case DeleteShape:
            case DeleteTextbox:
            case RefreshPivot:
            case DeletePivot:
                ValidateObjectName(op, index, opName, "name", errors, definedNameRules: opName == DeleteTable);
                break;
            case RemoveDuplicates:
                ValidateA1Field(op, index, opName, "range", errors, allowUnion: false);
                ValidateOptionalBool(op, index, opName, "hasHeaders", errors);
                ValidateOptionalIndexArray(op, index, opName, "columns", errors);
                break;
            case TextToColumns:
                ValidateA1Field(op, index, opName, "range", errors, allowUnion: false);
                if (op.ContainsKey("destination"))
                    ValidateA1Field(op, index, opName, "destination", errors, allowUnion: false);
                if (op.ContainsKey("dataType"))
                {
                    var dataType = Json.GetString(op, "dataType");
                    if (string.Equals(dataType, "fixed", StringComparison.OrdinalIgnoreCase))
                    {
                        errors.Add(
                            $"ops[{index}] '{opName}' dataType=fixed is not implemented (no FieldInfo/column breaks); " +
                            "use delimited or omit dataType");
                    }
                    else if (dataType is not "delimited")
                    {
                        errors.Add($"ops[{index}] '{opName}' dataType must be delimited or omitted");
                    }
                }
                foreach (var flag in new[] { "comma", "tab", "semicolon", "space", "other", "consecutiveDelimiter" })
                    ValidateOptionalBool(op, index, opName, flag, errors);
                ValidateOptionalBoundedString(op, index, opName, "otherChar", 1, errors);
                if (Json.GetBool(op, "other") && string.IsNullOrWhiteSpace(Json.GetString(op, "otherChar")))
                    errors.Add($"ops[{index}] '{opName}' other=true requires otherChar");
                break;
            case CreatePivot:
                ValidatePivotSource(op, index, opName, errors, required: true);
                ValidateA1Field(op, index, opName, "destination", errors, allowUnion: false);
                ValidateObjectName(op, index, opName, "name", errors, definedNameRules: false);
                ValidatePivotLayout(op, index, opName, errors, required: false);
                break;
            case UpdatePivot:
                ValidateObjectName(op, index, opName, "name", errors, definedNameRules: false);
                if (op.ContainsKey("sourceRange"))
                    ValidatePivotSource(op, index, opName, errors, required: false);
                ValidatePivotLayout(op, index, opName, errors, required: false);
                if (!HasAnyUpdate(op, "rows", "columns", "values", "filters", "sourceRange"))
                    errors.Add($"ops[{index}] '{opName}' requires rows, columns, values, filters, or sourceRange");
                break;
            case InsertShape:
                if (!ExcelDataObjectCatalog.TryShapeType(Json.GetString(op, "shapeType"), out _))
                    errors.Add($"ops[{index}] '{opName}' shapeType must be rectangle|roundedRectangle|oval|rightArrow|downArrow");
                ValidateOptionalObjectName(op, index, opName, "name", errors, definedNameRules: false);
                ValidateOptionalPosition(op, index, opName, errors, requireSize: true);
                ValidateOptionalBoundedString(op, index, opName, "text", 32_767, errors);
                ValidateOptionalFillColor(op, index, opName, errors);
                ExcelShapeFormatContract.Validate(op, index, opName, errors);
                break;
            case UpdateShape:
            case UpdateTextbox:
                ValidateObjectName(op, index, opName, "name", errors, definedNameRules: false);
                ValidateOptionalPosition(op, index, opName, errors, requireSize: false);
                ValidateOptionalBoundedString(op, index, opName, "text", 32_767, errors);
                ValidateOptionalFillColor(op, index, opName, errors);
                ExcelShapeFormatContract.Validate(op, index, opName, errors);
                if (!op.ContainsKey("position") && !op.ContainsKey("text") && !op.ContainsKey("fillColor") &&
                    !ExcelShapeFormatContract.HasFormatting(op))
                    errors.Add($"ops[{index}] '{opName}' requires position, text, fillColor, or practical formatting");
                break;
            case InsertTextbox:
                ValidateOptionalObjectName(op, index, opName, "name", errors, definedNameRules: false);
                ValidateOptionalPosition(op, index, opName, errors, requireSize: true);
                ValidateOptionalBoundedString(op, index, opName, "text", 32_767, errors);
                ValidateOptionalFillColor(op, index, opName, errors);
                ExcelShapeFormatContract.Validate(op, index, opName, errors);
                if (!op.ContainsKey("position") && string.IsNullOrWhiteSpace(Json.GetString(op, "text")))
                    errors.Add($"ops[{index}] '{opName}' requires position and/or text");
                break;
        }

        return errors.Count == before;
    }

    public static List<string> ValidatePublicInput(JsonObject op, int index = 1)
    {
        var errors = new List<string>();
        ValidatePublicInput(op, index, errors);
        return errors;
    }

    public static void ValidateDataBatchMixing(IReadOnlyList<JsonObject> ops, ICollection<string> errors)
    {
        var hasData = ops.Any(op => IsDataOperation(Json.GetString(op, "op")));
        var hasOther = ops.Any(op => !IsDataOperation(Json.GetString(op, "op")));
        if (hasData && hasOther)
        {
            errors.Add(
                "Excel data/reporting operations cannot be mixed with value, format, merge, visibility, or copy_sheet operations; " +
                "use a separate dry-run batch so data-objects rollback remains exact");
        }
    }

    internal static bool TryParseA1(string? value, out ExcelA1Box box, bool allowUnion = false)
    {
        box = default;
        if (string.IsNullOrWhiteSpace(value)) return false;
        try
        {
            var parsed = ExcelRangeReference.Parse(value);
            if (parsed.Address.Contains(',', StringComparison.Ordinal) ||
                parsed.Address.Contains(';', StringComparison.Ordinal))
                return allowUnion;
            return ExcelA1Box.TryParse(parsed.Address, out box);
        }
        catch (FormatException)
        {
            return false;
        }
    }

    public static bool TryParseCell(string text, out int row, out int col) =>
        ExcelA1Box.TryParseCell(text, out row, out col);

    public static bool IsValidDefinedName(string? name)
    {
        if (string.IsNullOrWhiteSpace(name) || name.Length > 255) return false;
        if (name.Equals("C", StringComparison.OrdinalIgnoreCase) ||
            name.Equals("R", StringComparison.OrdinalIgnoreCase))
            return false;
        if (!DefinedNamePattern.IsMatch(name)) return false;
        return !TryParseCell(name, out _, out _);
    }

    public static bool IsValidObjectName(string? name)
    {
        if (string.IsNullOrWhiteSpace(name) || name.Length > 255) return false;
        foreach (var ch in name)
        {
            if (char.IsControl(ch) || ch == '!') return false;
        }
        return true;
    }

    public static bool TryGetFiniteNumber(JsonNode? node, out double number)
    {
        number = 0;
        if (node is not JsonValue value) return false;
        if (value.TryGetValue<int>(out var i)) { number = i; return true; }
        if (value.TryGetValue<long>(out var l)) { number = l; return true; }
        if (value.TryGetValue<uint>(out var ui)) { number = ui; return true; }
        if (value.TryGetValue<decimal>(out var dec))
        {
            number = Convert.ToDouble(dec, CultureInfo.InvariantCulture);
            return double.IsFinite(number);
        }
        if (value.TryGetValue<float>(out var f)) { number = f; return double.IsFinite(number); }
        return value.TryGetValue<double>(out number) && double.IsFinite(number);
    }

    public static string? ResolveRangeField(JsonObject op)
    {
        var name = Json.GetString(op, "op");
        if (name == CreatePivot)
            return Json.GetString(op, "destination");
        return Json.GetString(op, "range") ?? Json.GetString(op, "cell") ?? Json.GetString(op, "sourceRange")
               ?? Json.GetString(op, "destination");
    }

    public static bool IsReadScope(string? scope) =>
        !string.IsNullOrWhiteSpace(scope) &&
        ReadScopes.Contains(scope, StringComparer.OrdinalIgnoreCase);

    public static string ResolveReadScope(JsonObject args)
    {
        var objectKind = Json.GetString(args, "objectKind")
                         ?? Json.GetString(args, "objectScope")
                         ?? Json.GetString(args, "kind");
        if (IsReadScope(objectKind)) return objectKind!;
        var scope = Json.GetString(args, "scope") ?? "all";
        if (string.Equals(scope, "objects", StringComparison.OrdinalIgnoreCase) ||
            string.Equals(scope, "all", StringComparison.OrdinalIgnoreCase))
            return "all";
        return scope;
    }

    public static JsonObject DescribeSchema()
    {
        var writeOps = new JsonArray();
        foreach (var name in WriteOpNames)
        {
            var required = new JsonArray();
            if (RequiredFields.TryGetValue(name, out var requiredFields))
            {
                foreach (var (field, type) in requiredFields)
                {
                    required.Add(new JsonObject
                    {
                        ["field"] = field,
                        ["type"] = type,
                        ["description"] = FieldExpectations.TryGetValue((name, field), out var expectation)
                            ? expectation
                            : type,
                    });
                }
            }
            writeOps.Add(new JsonObject
            {
                ["op"] = name,
                ["required"] = required,
                ["optional"] = new JsonArray((OptionalFields.TryGetValue(name, out var optional)
                    ? optional
                    : Array.Empty<string>()).Select(field => (JsonNode)JsonValue.Create(field)!).ToArray()),
            });
        }

        return new JsonObject
        {
            ["snapshotVersion"] = SnapshotVersion,
            ["restoreMode"] = RestoreMode,
            ["tableCreateRollback"] = ExcelDataSnapshotEquality.TableCreateRollbackMethod,
            ["equalityRequired"] = true,
            ["writeOps"] = writeOps,
            ["readScopes"] = new JsonArray(ReadScopes.Select(scope => (JsonNode)JsonValue.Create(scope)!).ToArray()),
            ["inspection"] = new JsonObject
            {
                ["scope"] = "all|objects|<readScope>",
                ["objectKind"] = "preferred when scope=objects; one of readScopes",
                ["sheet"] = "optional worksheet name",
                ["range"] = "optional bounded A1 range for cell-level objects",
                ["limit"] = "1..2000",
            },
            ["fieldTypes"] = new JsonObject
            {
                ["range"] = "rectangular A1, optional sheet qualifier",
                ["name"] = "Excel object or defined name",
                ["scope"] = "workbook|sheet",
                ["chartType"] = "columnClustered|columnStacked|barClustered|barStacked|line|lineMarkers|pie|area|scatter|scatterLines",
                ["marker"] = "none|automatic|circle|dash|diamond|dot|square|triangle",
                ["validationType"] = "list|whole|decimal|date|time|textLength|custom",
                ["conditionalRuleType"] = "cellValue|expression|uniqueValues|duplicateValues|colorScale|dataBar|iconSet",
                ["calculationMode"] = "manual|automatic|semiautomatic",
                ["recalculate"] = "none|full|fullRebuild",
                ["pasteType"] = "values|formulas|formats|all (borders excluded; use format_range)",
                ["pasteOperation"] = "none|add|subtract|multiply|divide",
                ["sparklineType"] = "line|column (winLoss rejected until native mapping is proven)",
                ["pivotFunction"] = "sum|count|average|max|min|countNums|product|stdDev|var",
                ["shapeType"] = "rectangle|roundedRectangle|oval|rightArrow|downArrow",
                ["position"] = "{left,top,width,height} points",
                ["dataType"] = "delimited; fixed is rejected until FieldInfo exists",
                ["recovery"] = "native-copy only when workbookBackupSource=current-memory-savecopyas and fresh+available+filename are proven",
            },
        };
    }

    private static void ValidateRequiredField(JsonObject op, int index, string opName, string field, string type,
        ICollection<string> errors)
    {
        if (!op.TryGetPropertyValue(field, out var node) || node is null)
        {
            errors.Add($"ops[{index}] '{opName}' requires field '{field}' ({Describe(opName, field, type)})");
            return;
        }
        var bad = type switch
        {
            "string" => node is not JsonValue jvs || !jvs.TryGetValue<string>(out var s) || string.IsNullOrWhiteSpace(s),
            "int" => node is not JsonValue jvi || !jvi.TryGetValue<int>(out _),
            "number" => !TryGetFiniteNumber(node, out _),
            "bool" => node is not JsonValue jvb || !jvb.TryGetValue<bool>(out _),
            "array" => node is not JsonArray arr || arr.Count == 0,
            "object" => node is not JsonObject,
            _ => false,
        };
        if (bad)
            errors.Add($"ops[{index}] '{opName}' field '{field}' must be {Describe(opName, field, type)}");
    }

    private static string Describe(string opName, string field, string type) =>
        FieldExpectations.TryGetValue((opName, field), out var expectation)
            ? expectation
            : type;

    private static void ValidateSheetTarget(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        var targetSheet = Json.GetString(Json.GetObj(op, "target"), "sheet");
        var rangeText = ResolveRangeField(op);
        string? rangeSheet = null;
        if (!string.IsNullOrWhiteSpace(rangeText))
        {
            try { rangeSheet = ExcelRangeReference.Parse(rangeText).SheetName; }
            catch (FormatException ex)
            {
                errors.Add($"ops[{index}] '{opName}' has invalid Excel range: {ex.Message}");
                return;
            }
        }

        if (RequiresExplicitSheet(op) &&
            string.IsNullOrWhiteSpace(targetSheet) &&
            string.IsNullOrWhiteSpace(rangeSheet))
        {
            errors.Add(
                $"ops[{index}] '{opName}' requires target.sheet or a sheet-qualified range; active sheet writes are not allowed");
            return;
        }

        if (ExcelDataRangeBinding.AllowsIndependentSource(opName))
        {
            var writeField = ExcelDataRangeBinding.WriteRangeField(opName);
            var writeRange = writeField is null ? null : Json.GetString(op, writeField);
            if (ExcelDataRangeBinding.WriteRangeConflictsWithTarget(targetSheet, writeRange))
            {
                errors.Add(
                    $"ops[{index}] '{opName}' target.sheet '{targetSheet}' does not match write range sheet");
            }
            return;
        }

        if (!string.IsNullOrWhiteSpace(targetSheet) &&
            !string.IsNullOrWhiteSpace(rangeSheet) &&
            !string.Equals(targetSheet, rangeSheet, StringComparison.OrdinalIgnoreCase))
        {
            errors.Add(
                $"ops[{index}] '{opName}' target.sheet '{targetSheet}' does not match range sheet '{rangeSheet}'");
        }
    }

    private static void ValidateA1Field(JsonObject op, int index, string opName, string field,
        ICollection<string> errors, bool allowUnion)
    {
        var text = Json.GetString(op, field);
        if (string.IsNullOrWhiteSpace(text)) return;
        if (!TryParseA1(text, out var box, allowUnion))
        {
            errors.Add($"ops[{index}] '{opName}' {field} must be a rectangular A1 address inside A1:XFD1048576");
            return;
        }
        if (box.CellCount > 1_000_000)
            errors.Add($"ops[{index}] '{opName}' {field} exceeds 1,000,000 cells");
    }

    private static void ValidateBoundedCellCount(JsonObject op, int index, string opName, string field,
        int max, ICollection<string> errors)
    {
        var text = Json.GetString(op, field);
        if (string.IsNullOrWhiteSpace(text)) return;
        if (TryParseA1(text, out var box, allowUnion: false) && box.CellCount > max)
            errors.Add($"ops[{index}] '{opName}' {field} exceeds {max} cells");
    }

    private static void ValidateGoalSeekCell(JsonObject op, int index, string opName, string field,
        ICollection<string> errors)
    {
        var text = Json.GetString(op, field);
        if (string.IsNullOrWhiteSpace(text))
        {
            errors.Add($"ops[{index}] '{opName}' {field} is required");
            return;
        }
        try
        {
            var parsed = ExcelRangeReference.Parse(text);
            if (!TryParseCell(parsed.Address, out _, out _))
                errors.Add($"ops[{index}] '{opName}' {field} must be a single A1 cell");
        }
        catch (FormatException)
        {
            errors.Add($"ops[{index}] '{opName}' {field} must be a single A1 cell");
        }
    }

    private static void ValidateGoalSeekSheets(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        var targetSheet = Json.GetString(Json.GetObj(op, "target"), "sheet");
        string? cellSheet = null;
        string? goalSheet = null;
        foreach (var field in new[] { "cell", "goalCell" })
        {
            var text = Json.GetString(op, field);
            if (string.IsNullOrWhiteSpace(text)) continue;
            try
            {
                var sheet = ExcelRangeReference.Parse(text).SheetName;
                if (field == "cell") cellSheet = sheet; else goalSheet = sheet;
                if (!string.IsNullOrWhiteSpace(sheet) && !string.IsNullOrWhiteSpace(targetSheet) &&
                    !string.Equals(sheet, targetSheet, StringComparison.OrdinalIgnoreCase))
                    errors.Add($"ops[{index}] '{opName}' {field} sheet '{sheet}' does not match target.sheet '{targetSheet}'");
            }
            catch (FormatException)
            {
            }
        }
        if (!string.IsNullOrWhiteSpace(cellSheet) && !string.IsNullOrWhiteSpace(goalSheet) &&
            !string.Equals(cellSheet, goalSheet, StringComparison.OrdinalIgnoreCase))
            errors.Add($"ops[{index}] '{opName}' cell and goalCell must be on the same worksheet");
    }

    private static void ValidateSparklineType(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        var type = Json.GetString(op, "type");
        if (string.Equals(type, "winLoss", StringComparison.OrdinalIgnoreCase))
        {
            errors.Add($"ops[{index}] '{opName}' type=winLoss is not implemented (no distinct XlSparkType); use line or column");
            return;
        }
        if (!ExcelWorkbookOpsContract.TrySparklineType(type, out _))
            errors.Add($"ops[{index}] '{opName}' type must be line|column");
    }

    private static void ValidateSparklineOptions(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        ValidateOptionalBool(op, index, opName, "markers", errors);
        foreach (var flag in new[] { "showHigh", "showLow", "showNegative", "showFirst", "showLast" })
            ValidateOptionalBool(op, index, opName, flag, errors);
        if (op.ContainsKey("lineColor"))
            ExcelStyleContract.TryParseColor(op["lineColor"], out _, errors, "lineColor", index);
    }

    private static void ValidateSparklineSheets(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        var targetSheet = Json.GetString(Json.GetObj(op, "target"), "sheet");
        var location = Json.GetString(op, "location");
        if (!string.IsNullOrWhiteSpace(location))
        {
            try
            {
                var sheet = ExcelRangeReference.Parse(location).SheetName;
                if (!string.IsNullOrWhiteSpace(sheet) && !string.IsNullOrWhiteSpace(targetSheet) &&
                    !string.Equals(sheet, targetSheet, StringComparison.OrdinalIgnoreCase))
                    errors.Add($"ops[{index}] '{opName}' location sheet '{sheet}' does not match target.sheet '{targetSheet}'");
            }
            catch (FormatException)
            {
            }
        }
    }

    private static void ValidateCellAddress(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        var text = Json.GetString(op, "range") ?? Json.GetString(op, "cell");
        if (string.IsNullOrWhiteSpace(text))
        {
            errors.Add($"ops[{index}] '{opName}' requires range or cell");
            return;
        }
        if (op.ContainsKey("range") && op.ContainsKey("cell") &&
            !string.Equals(Json.GetString(op, "range"), Json.GetString(op, "cell"), StringComparison.OrdinalIgnoreCase))
        {
            errors.Add($"ops[{index}] '{opName}' range and cell must not conflict");
            return;
        }
        if (!TryParseA1(text, out var box, allowUnion: false))
        {
            errors.Add($"ops[{index}] '{opName}' range/cell must be a rectangular A1 address");
            return;
        }
        if (opName is SetCellNote or SetHyperlink && box.CellCount != 1)
            errors.Add($"ops[{index}] '{opName}' target must be a single cell");
    }

    private static void ValidateObjectName(JsonObject op, int index, string opName, string field,
        ICollection<string> errors, bool definedNameRules)
    {
        var value = Json.GetString(op, field);
        if (string.IsNullOrWhiteSpace(value)) return;
        if (definedNameRules)
        {
            if (!IsValidDefinedName(value))
                errors.Add($"ops[{index}] '{opName}' {field} is not a valid Excel name");
        }
        else if (!IsValidObjectName(value))
        {
            errors.Add($"ops[{index}] '{opName}' {field} must be 1..255 characters without control chars or '!'");
        }
    }

    private static void ValidateOptionalObjectName(JsonObject op, int index, string opName, string field,
        ICollection<string> errors, bool definedNameRules)
    {
        if (!op.ContainsKey(field)) return;
        ValidateRequiredField(op, index, opName, field, "string", errors);
        ValidateObjectName(op, index, opName, field, errors, definedNameRules);
    }

    private static void ValidateDefinedName(JsonObject op, int index, string opName, ICollection<string> errors) =>
        ValidateObjectName(op, index, opName, "name", errors, definedNameRules: true);

    private static void ValidateNameScope(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        var scope = Json.GetString(op, "scope")?.Trim();
        if (scope is not ("workbook" or "sheet"))
            errors.Add($"ops[{index}] '{opName}' scope must be 'workbook' or 'sheet'");
        else if (scope == "sheet" && string.IsNullOrWhiteSpace(Json.GetString(Json.GetObj(op, "target"), "sheet")))
            errors.Add($"ops[{index}] '{opName}' with scope=sheet requires target.sheet");
    }

    private static void ValidateRefersTo(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        var refersTo = Json.GetString(op, "refersTo");
        if (string.IsNullOrWhiteSpace(refersTo)) return;
        if (!refersTo.StartsWith('=') || refersTo.Length < 2 || refersTo.Length > 255)
            errors.Add($"ops[{index}] '{opName}' refersTo must be an Excel formula starting with '=' (2..255 chars)");
    }

    private static void ValidateOptionalStyleName(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        if (!op.ContainsKey("styleName")) return;
        var value = Json.GetString(op, "styleName");
        if (string.IsNullOrWhiteSpace(value) || value.Length > 255 || value.Any(char.IsControl))
            errors.Add($"ops[{index}] '{opName}' styleName must be a non-empty Excel table style name");
    }

    private static void ValidateOptionalBool(JsonObject op, int index, string opName, string field,
        ICollection<string> errors)
    {
        if (!op.ContainsKey(field)) return;
        if (op[field] is not JsonValue value || !value.TryGetValue<bool>(out _))
            errors.Add($"ops[{index}] '{opName}' {field} must be boolean");
    }

    private static void ValidateOptionalBoundedString(JsonObject op, int index, string opName, string field,
        int max, ICollection<string> errors)
    {
        if (!op.ContainsKey(field)) return;
        if (op[field] is not JsonValue value || !value.TryGetValue<string>(out var text) || text is null)
        {
            errors.Add($"ops[{index}] '{opName}' {field} must be string");
            return;
        }
        if (text.Length > max)
            errors.Add($"ops[{index}] '{opName}' {field} exceeds {max} characters");
    }

    private static void ValidateRequiredBoundedString(JsonObject op, int index, string opName, string field,
        int max, ICollection<string> errors)
    {
        var text = Json.GetString(op, field);
        if (string.IsNullOrEmpty(text))
        {
            errors.Add($"ops[{index}] '{opName}' {field} must be a non-empty string");
            return;
        }
        if (text.Length > max)
            errors.Add($"ops[{index}] '{opName}' {field} exceeds {max} characters");
    }

    private static void ValidateRangeXorName(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        var hasRange = !string.IsNullOrWhiteSpace(Json.GetString(op, "range"));
        var hasName = !string.IsNullOrWhiteSpace(Json.GetString(op, "name"));
        if (hasRange == hasName)
            errors.Add($"ops[{index}] '{opName}' requires exactly one of range or name");
        if (hasRange) ValidateA1Field(op, index, opName, "range", errors, allowUnion: false);
        if (hasName) ValidateObjectName(op, index, opName, "name", errors, definedNameRules: true);
    }

    private static void ValidateSortKeys(JsonObject op, int index, string opName, ICollection<string> errors, bool table)
    {
        var keys = Json.GetArr(op, "keys");
        if (keys is null) return;
        if (keys.Count is < 1 or > 3)
        {
            errors.Add($"ops[{index}] '{opName}' keys must contain 1..3 sort keys");
            return;
        }
        for (var i = 0; i < keys.Count; i++)
        {
            if (keys[i] is not JsonObject key)
            {
                errors.Add($"ops[{index}] '{opName}' keys[{i}] must be an object");
                continue;
            }
            if (!key.ContainsKey("column"))
                errors.Add($"ops[{index}] '{opName}' keys[{i}].column is required");
            else if (!TryValidateColumnKey(key["column"], table))
                errors.Add($"ops[{index}] '{opName}' keys[{i}].column must be a 1-based index, A1 column letter, or header name");
            var order = Json.GetString(key, "order") ?? "asc";
            if (order is not ("asc" or "desc" or "ascending" or "descending"))
                errors.Add($"ops[{index}] '{opName}' keys[{i}].order must be asc or desc");
        }
    }

    private static bool TryValidateColumnKey(JsonNode? node, bool allowName)
    {
        if (node is JsonValue value && value.TryGetValue<int>(out var number))
            return number is >= 1 and <= 16_384;
        if (node is not JsonValue textValue || !textValue.TryGetValue<string>(out var text) ||
            string.IsNullOrWhiteSpace(text))
            return false;
        if (TryParseCell(text + "1", out _, out _)) return true;
        return allowName && IsValidObjectName(text);
    }

    private static void ValidateFilterCriteria(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        var criteria = Json.GetArr(op, "criteria");
        if (criteria is null) return;
        if (criteria.Count is < 1 or > 64)
        {
            errors.Add($"ops[{index}] '{opName}' criteria must contain 1..64 entries");
            return;
        }
        for (var i = 0; i < criteria.Count; i++)
        {
            if (criteria[i] is not JsonObject item)
            {
                errors.Add($"ops[{index}] '{opName}' criteria[{i}] must be an object");
                continue;
            }
            if (!item.ContainsKey("column") || !TryValidateColumnKey(item["column"], allowName: true))
                errors.Add($"ops[{index}] '{opName}' criteria[{i}].column is required");
            var filterOp = Json.GetString(item, "operator") ?? "equals";
            if (!ExcelDataObjectCatalog.FilterOperators.Contains(filterOp))
                errors.Add($"ops[{index}] '{opName}' criteria[{i}].operator is not supported");
            if (string.Equals(filterOp, "values", StringComparison.OrdinalIgnoreCase))
            {
                var values = Json.GetArr(item, "values");
                if (values is null || values.Count == 0 || values.Count > 100)
                    errors.Add($"ops[{index}] '{opName}' criteria[{i}].values must be a non-empty array of at most 100 items");
            }
            else if (string.Equals(filterOp, "between", StringComparison.OrdinalIgnoreCase))
            {
                if (!item.ContainsKey("value") || !item.ContainsKey("value2"))
                    errors.Add($"ops[{index}] '{opName}' criteria[{i}] between requires value and value2");
            }
            else if (!item.ContainsKey("value"))
            {
                errors.Add($"ops[{index}] '{opName}' criteria[{i}].value is required");
            }
        }
    }

    private static void ValidateTotalsColumns(JsonObject op, int index, string opName, ICollection<string> errors,
        bool required)
    {
        var columns = Json.GetArr(op, "columns");
        if (columns is null)
        {
            if (required)
                errors.Add($"ops[{index}] '{opName}' columns are required when showTotals is true");
            return;
        }
        if (columns.Count is < 1 or > 256)
        {
            errors.Add($"ops[{index}] '{opName}' columns must contain 1..256 entries");
            return;
        }
        for (var i = 0; i < columns.Count; i++)
        {
            if (columns[i] is not JsonObject column)
            {
                errors.Add($"ops[{index}] '{opName}' columns[{i}] must be an object");
                continue;
            }
            if (!column.ContainsKey("column") || !TryValidateColumnKey(column["column"], allowName: true))
                errors.Add($"ops[{index}] '{opName}' columns[{i}].column is required");
            var function = Json.GetString(column, "function");
            if (!ExcelDataObjectCatalog.TryTotalsFunction(function, out _))
                errors.Add($"ops[{index}] '{opName}' columns[{i}].function must be a totals token");
            if (string.Equals(function, "custom", StringComparison.OrdinalIgnoreCase))
            {
                var formula = Json.GetString(column, "formula");
                if (string.IsNullOrWhiteSpace(formula) || !formula.StartsWith('='))
                    errors.Add($"ops[{index}] '{opName}' columns[{i}].formula is required for function=custom");
            }
        }
    }

    private static void ValidateDataValidation(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        var type = Json.GetString(op, "type");
        if (!ExcelDataObjectCatalog.TryValidationType(type, out _))
            errors.Add($"ops[{index}] '{opName}' type must be list|whole|decimal|date|time|textLength|custom");

        if (op.ContainsKey("operator") && !ExcelDataObjectCatalog.TryComparisonOperator(Json.GetString(op, "operator"), out _))
            errors.Add($"ops[{index}] '{opName}' operator must be between|notBetween|equal|notEqual|greater|less|greaterEqual|lessEqual");
        if (op.ContainsKey("errorStyle") && !ExcelDataObjectCatalog.TryErrorStyle(Json.GetString(op, "errorStyle"), out _))
            errors.Add($"ops[{index}] '{opName}' errorStyle must be stop|warning|information");

        foreach (var flag in new[] { "inCellDropdown", "ignoreBlank", "showInput", "showError" })
            ValidateOptionalBool(op, index, opName, flag, errors);
        foreach (var field in new[] { "inputTitle", "errorTitle" })
            ValidateOptionalBoundedString(op, index, opName, field, 32, errors);
        foreach (var field in new[] { "inputMessage", "errorMessage" })
            ValidateOptionalBoundedString(op, index, opName, field, 255, errors);

        if (string.Equals(type, "list", StringComparison.OrdinalIgnoreCase))
        {
            var source = Json.GetString(op, "source");
            var formula1 = Json.GetString(op, "formula1");
            if (string.IsNullOrWhiteSpace(source) && string.IsNullOrWhiteSpace(formula1))
                errors.Add($"ops[{index}] '{opName}' list type requires source or formula1");
        }
        else if (type is "whole" or "decimal" or "date" or "time" or "textLength")
        {
            if (string.IsNullOrWhiteSpace(Json.GetString(op, "formula1")) && !op.ContainsKey("formula1"))
                errors.Add($"ops[{index}] '{opName}' {type} requires formula1");
            var comparison = Json.GetString(op, "operator") ?? "between";
            if (comparison is "between" or "notBetween" &&
                string.IsNullOrWhiteSpace(Json.GetString(op, "formula2")) && !op.ContainsKey("formula2"))
                errors.Add($"ops[{index}] '{opName}' {comparison} requires formula2");
        }
        else if (string.Equals(type, "custom", StringComparison.OrdinalIgnoreCase))
        {
            var formula1 = Json.GetString(op, "formula1");
            if (string.IsNullOrWhiteSpace(formula1) || !formula1.StartsWith('='))
                errors.Add($"ops[{index}] '{opName}' custom type requires formula1 starting with '='");
        }
    }

    private static void ValidateConditionalRule(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        var rule = Json.GetObj(op, "rule");
        if (rule is null) return;
        var type = Json.GetString(rule, "type");
        if (string.IsNullOrWhiteSpace(type) || !ExcelDataObjectCatalog.ConditionalRuleTypes.Contains(type))
        {
            errors.Add($"ops[{index}] '{opName}' rule.type must be cellValue|expression|uniqueValues|duplicateValues|colorScale|dataBar|iconSet");
            return;
        }
        if (type is "cellValue")
        {
            if (!ExcelDataObjectCatalog.TryComparisonOperator(Json.GetString(rule, "operator"), out _))
                errors.Add($"ops[{index}] '{opName}' cellValue rule.operator is required");
            if (string.IsNullOrWhiteSpace(Json.GetString(rule, "formula1")))
                errors.Add($"ops[{index}] '{opName}' cellValue rule.formula1 is required");
        }
        else if (type is "expression" && string.IsNullOrWhiteSpace(Json.GetString(rule, "formula1")))
        {
            errors.Add($"ops[{index}] '{opName}' expression rule.formula1 is required");
        }
        else if (type is "colorScale")
        {
            var points = Json.GetInt(rule, "points") ?? 2;
            if (points is not (2 or 3))
                errors.Add($"ops[{index}] '{opName}' colorScale rule.points must be 2 or 3");
        }
        else if (type is "iconSet")
        {
            if (!ExcelDataObjectCatalog.TryIconSet(Json.GetString(rule, "iconSet"), out _))
                errors.Add($"ops[{index}] '{opName}' iconSet rule.iconSet must be arrows3|trafficLights3|symbols3|symbols3Alt|arrows4|arrows5");
        }
    }

    private static void ValidateConditionalStyle(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        var style = Json.GetObj(op, "style");
        if (style is null) return;
        foreach (var (key, node) in style)
        {
            if (key is "bold" or "fontBold" or "italic" or "fontItalic")
            {
                if (node is not JsonValue flag || !flag.TryGetValue<bool>(out _))
                    errors.Add($"ops[{index}] '{opName}' style.{key} must be boolean");
            }
            else if (key is "fontColor" or "fillColor" or "interiorColor" or "fill")
            {
                ExcelStyleContract.TryParseColor(node, out _, errors, key, index);
            }
            else
            {
                errors.Add($"ops[{index}] '{opName}' style key '{key}' is not supported; allowed: bold, italic, fontColor, fillColor");
            }
        }
    }

    private static void ValidateChartType(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        if (!ExcelDataObjectCatalog.TryChartType(Json.GetString(op, "chartType"), out _))
            errors.Add($"ops[{index}] '{opName}' chartType is not a supported native chart token");
    }

    private static void ValidateChartPresentation(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        foreach (var field in new[] { "chartFill", "plotFill", "chartBorder", "plotBorder" })
            ValidateOptionalFillOrNone(op, index, opName, field, errors);
        ValidateChartSeries(op, index, opName, errors);
        ValidateChartAxes(op, index, opName, errors);
        ValidateChartPlotArea(op, index, opName, errors);
        ExcelChartDetailsContract.Validate(op, index, opName, errors);
    }

    private static void ValidateOptionalFillOrNone(JsonObject op, int index, string opName, string field,
        ICollection<string> errors)
    {
        if (!op.ContainsKey(field)) return;
        var text = Json.GetString(op, field);
        if (string.Equals(text, "none", StringComparison.OrdinalIgnoreCase)) return;
        ExcelStyleContract.TryParseColor(op[field], out _, errors, field, index);
    }

    private static void ValidateChartSeries(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        if (!op.ContainsKey("series")) return;
        if (op["series"] is not JsonArray series || series.Count is < 1 or > 64)
        {
            errors.Add($"ops[{index}] '{opName}' series must be an array of 1..64 series objects");
            return;
        }
        for (var i = 0; i < series.Count; i++)
        {
            if (series[i] is not JsonObject item)
            {
                errors.Add($"ops[{index}] '{opName}' series[{i}] must be an object");
                continue;
            }
            if (item.ContainsKey("values"))
                ValidateA1Field(item, index, opName, "values", errors, allowUnion: true);
            if (item.ContainsKey("categories"))
                ValidateA1Field(item, index, opName, "categories", errors, allowUnion: true);
            if (item.ContainsKey("range"))
                ValidateA1Field(item, index, opName, "range", errors, allowUnion: true);
            if (item.ContainsKey("name"))
            {
                var seriesName = Json.GetString(item, "name");
                if (string.IsNullOrWhiteSpace(seriesName))
                    errors.Add($"ops[{index}] '{opName}' series[{i}].name must be a non-empty string or explicit name reference");
                else if (ExcelChartSeriesContract.LooksLikeNameRange(seriesName))
                    ValidateSeriesNameReference(item, index, opName, i, errors);
            }
            if (item.ContainsKey("lineColor"))
                ExcelStyleContract.TryParseColor(item["lineColor"], out _, errors, $"series[{i}].lineColor", index);
            if (item.ContainsKey("lineWeight"))
            {
                if (!TryGetFiniteNumber(item["lineWeight"], out var weight) || weight < 0.25 || weight > 10)
                    errors.Add($"ops[{index}] '{opName}' series[{i}].lineWeight must be a finite number 0.25..10 (pt)");
            }
            if (item.ContainsKey("marker") && !ExcelDataObjectCatalog.TryMarkerStyle(Json.GetString(item, "marker"), out _))
                errors.Add($"ops[{index}] '{opName}' series[{i}].marker must be none|automatic|circle|dash|diamond|dot|square|triangle");
            if (item.ContainsKey("chartType") && !ExcelDataObjectCatalog.TryChartType(Json.GetString(item, "chartType"), out _))
                errors.Add($"ops[{index}] '{opName}' series[{i}].chartType is not a supported native chart token");
            if (item.ContainsKey("axisGroup") && !ExcelDataObjectCatalog.TryAxisGroup(Json.GetString(item, "axisGroup"), out _))
                errors.Add($"ops[{index}] '{opName}' series[{i}].axisGroup must be primary or secondary");
            if (item.ContainsKey("index"))
            {
                if (!TryGetFiniteNumber(item["index"], out var seriesIndex) ||
                    seriesIndex != Math.Truncate(seriesIndex) || seriesIndex < 1 || seriesIndex > 64)
                    errors.Add($"ops[{index}] '{opName}' series[{i}].index must be an integer 1..64");
            }
        }
        var specs = ExcelChartSeriesContract.ReadSeries(series);
        var seen = new HashSet<int>();
        foreach (var spec in specs)
        {
            if (!seen.Add(spec.Index))
            {
                errors.Add($"ops[{index}] '{opName}' series index {spec.Index} is duplicated");
                break;
            }
        }
        if (string.Equals(opName, CreateChart, StringComparison.Ordinal) &&
            !ExcelChartSeriesContract.IndexesAreDenseFromOne(specs))
            errors.Add($"ops[{index}] '{opName}' series index must be contiguous from 1 (sparse index is rejected before apply)");
    }

    private static void ValidateChartAxes(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        if (!op.ContainsKey("axes")) return;
        if (op["axes"] is not JsonObject axes)
        {
            errors.Add($"ops[{index}] '{opName}' axes must be an object {{category,value,valueSecondary}}");
            return;
        }
        foreach (var property in axes)
        {
            if (!ExcelChartAxesContract.Keys.Contains(property.Key, StringComparer.Ordinal))
                errors.Add($"ops[{index}] '{opName}' axes.{property.Key} is not a supported axis key");
        }

        var valueGroup = Json.GetString(Json.GetObj(axes, "value"), "group");
        var secondaryGroup = Json.GetString(Json.GetObj(axes, "valueSecondary"), "group");
        if (axes.ContainsKey("value") && axes.ContainsKey("valueSecondary") &&
            string.Equals(valueGroup, "secondary", StringComparison.OrdinalIgnoreCase))
            errors.Add($"ops[{index}] '{opName}' axes.value.group=secondary conflicts with valueSecondary");
        if (string.Equals(Json.GetString(Json.GetObj(axes, "category"), "group"), "secondary",
                StringComparison.OrdinalIgnoreCase))
            errors.Add($"ops[{index}] '{opName}' axes.category.group must be primary");
        if (string.Equals(secondaryGroup, "primary", StringComparison.OrdinalIgnoreCase))
            errors.Add($"ops[{index}] '{opName}' axes.valueSecondary.group must be secondary");

        foreach (var name in ExcelChartAxesContract.Keys)
        {
            if (!axes.ContainsKey(name)) continue;
            if (axes[name] is not JsonObject axis)
            {
                errors.Add($"ops[{index}] '{opName}' axes.{name} must be an object");
                continue;
            }
            if (axis.ContainsKey("visible") &&
                (axis["visible"] is not JsonValue flag || !flag.TryGetValue<bool>(out _)))
                errors.Add($"ops[{index}] '{opName}' axes.{name}.visible must be boolean");
            if (axis.ContainsKey("group") && !ExcelDataObjectCatalog.TryAxisGroup(Json.GetString(axis, "group"), out _))
                errors.Add($"ops[{index}] '{opName}' axes.{name}.group must be primary or secondary");
            foreach (var scale in new[] { "minimum", "maximum" })
            {
                if (!axis.ContainsKey(scale)) continue;
                if (!TryGetFiniteNumber(axis[scale], out _))
                    errors.Add($"ops[{index}] '{opName}' axes.{name}.{scale} must be a finite number");
            }
            if (axis.ContainsKey("numberFormat"))
            {
                var format = Json.GetString(axis, "numberFormat");
                if (string.IsNullOrWhiteSpace(format) || format.Length > 255)
                    errors.Add($"ops[{index}] '{opName}' axes.{name}.numberFormat must be a non-empty string up to 255 characters");
            }
            if (axis.ContainsKey("title"))
            {
                var title = Json.GetString(axis, "title");
                if (title is null || title.Length > 255)
                    errors.Add($"ops[{index}] '{opName}' axes.{name}.title must be a string up to 255 characters");
            }
        }
    }

    private static void ValidateChartPlotArea(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        if (!op.ContainsKey("plotArea")) return;
        if (op["plotArea"] is not JsonObject plot)
        {
            errors.Add($"ops[{index}] '{opName}' plotArea must be an object");
            return;
        }
        ValidateOptionalBool(plot, index, opName, "spanChart", errors);
        if (plot.ContainsKey("left") || plot.ContainsKey("top") || plot.ContainsKey("width") || plot.ContainsKey("height"))
        {
            var wrapped = new JsonObject { ["position"] = plot.DeepClone() };
            ValidateOptionalPosition(wrapped, index, opName, errors, requireSize: false);
        }
    }

    private static void ValidateOptionalLegend(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        if (!op.ContainsKey("legendPosition")) return;
        if (!ExcelDataObjectCatalog.TryLegendPosition(Json.GetString(op, "legendPosition"), out _))
            errors.Add($"ops[{index}] '{opName}' legendPosition must be bottom|corner|top|right|left");
    }

    private static void ValidateOptionalPlotBy(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        if (!op.ContainsKey("plotBy")) return;
        var plotBy = Json.GetString(op, "plotBy");
        if (plotBy is not ("rows" or "columns"))
            errors.Add($"ops[{index}] '{opName}' plotBy must be rows or columns");
    }

    private static void ValidateOptionalPosition(JsonObject op, int index, string opName, ICollection<string> errors,
        bool requireSize)
    {
        if (!op.ContainsKey("position")) return;
        if (op["position"] is not JsonObject position)
        {
            errors.Add($"ops[{index}] '{opName}' position must be an object {{left,top,width,height}} in points");
            return;
        }
        foreach (var field in new[] { "left", "top", "width", "height" })
        {
            if (!position.ContainsKey(field))
            {
                if (requireSize && field is "width" or "height")
                    errors.Add($"ops[{index}] '{opName}' position.{field} is required");
                else if (field is "left" or "top")
                    errors.Add($"ops[{index}] '{opName}' position.{field} is required when position is present");
                continue;
            }
            if (!TryGetFiniteNumber(position[field], out var number) || number < 0 || number > 20_000)
                errors.Add($"ops[{index}] '{opName}' position.{field} must be a finite number 0..20000 (points)");
        }
    }

    private static void ValidatePicturePath(JsonObject op, int index, string opName, ICollection<string> errors)
    {
        var path = Json.GetString(op, "path");
        if (string.IsNullOrWhiteSpace(path)) return;
        if (!Path.IsPathFullyQualified(path))
        {
            errors.Add($"ops[{index}] '{opName}' path must be an absolute local path");
            return;
        }
        var extension = Path.GetExtension(path);
        if (!ExcelDataObjectCatalog.PictureExtensions.Contains(extension))
            errors.Add($"ops[{index}] '{opName}' path extension '{extension}' is not a supported image type");
    }

    private static bool HasAnyUpdate(JsonObject op, params string[] fields) =>
        fields.Any(op.ContainsKey);

    private static void ValidateOptionalFillColor(JsonObject op, int index, string opName,
        ICollection<string> errors)
    {
        if (!op.ContainsKey("fillColor")) return;
        ExcelStyleContract.TryParseColor(op["fillColor"], out _, errors, "fillColor", index);
    }

    private static void ValidateOptionalPositiveInt(JsonObject op, int index, string opName, string field,
        int max, ICollection<string> errors)
    {
        if (!op.ContainsKey(field)) return;
        if (!TryGetFiniteNumber(op[field], out var number) ||
            number < 1 || number > max || Math.Abs(number - Math.Truncate(number)) > double.Epsilon)
        {
            errors.Add($"ops[{index}] '{opName}' {field} must be a finite integer 1..{max}");
        }
    }

    private static void ValidateOptionalIndexArray(JsonObject op, int index, string opName, string field,
        ICollection<string> errors)
    {
        if (!op.ContainsKey(field)) return;
        if (op[field] is not JsonArray columns || columns.Count == 0)
        {
            errors.Add($"ops[{index}] '{opName}' {field} must be a non-empty array of 1-based indexes");
            return;
        }
        for (var i = 0; i < columns.Count; i++)
        {
            if (columns[i] is not JsonValue value || !value.TryGetValue<int>(out var number) || number < 1)
                errors.Add($"ops[{index}] '{opName}' {field}[{i}] must be a 1-based column index");
        }
    }

    private static void ValidatePivotLayout(JsonObject op, int index, string opName, ICollection<string> errors,
        bool required)
    {
        var hasLayout = op.ContainsKey("rows") || op.ContainsKey("columns") || op.ContainsKey("values") ||
                        op.ContainsKey("filters");
        if (required && !hasLayout)
            errors.Add($"ops[{index}] '{opName}' requires rows, columns, values, or filters");
        ValidateStringArray(op, index, opName, "rows", errors);
        ValidateStringArray(op, index, opName, "columns", errors);
        ValidateStringArray(op, index, opName, "filters", errors);
        var values = Json.GetArr(op, "values");
        if (values is null) return;
        if (values.Count is < 1 or > 64)
        {
            errors.Add($"ops[{index}] '{opName}' values must contain 1..64 entries");
            return;
        }
        for (var i = 0; i < values.Count; i++)
        {
            if (values[i] is not JsonObject item)
            {
                errors.Add($"ops[{index}] '{opName}' values[{i}] must be an object {{field, function}}");
                continue;
            }
            if (string.IsNullOrWhiteSpace(Json.GetString(item, "field")))
                errors.Add($"ops[{index}] '{opName}' values[{i}].field is required");
            if (item.ContainsKey("function"))
            {
                var function = Json.GetString(item, "function");
                if (string.IsNullOrWhiteSpace(function) ||
                    !ExcelDataObjectCatalog.TryPivotFunction(function, out _))
                {
                    errors.Add(
                        $"ops[{index}] '{opName}' values[{i}].function must be a pivot aggregation token; " +
                        "omit the field to default to sum — null is not sum");
                }
            }
            ValidateOptionalBoundedString(item, index, opName, "caption", 255, errors);
            ValidateOptionalBoundedString(item, index, opName, "numberFormat", 255, errors);
        }
    }

    private static void ValidateSeriesNameReference(JsonObject item, int index, string opName, int seriesIndex,
        ICollection<string> errors)
    {
        var text = Json.GetString(item, "name")?.Trim().TrimStart('=');
        if (!ExcelReferenceIdentity.TryParseSource(text, null, out var source) || source.IsTable)
        {
            errors.Add(
                $"ops[{index}] '{opName}' series[{seriesIndex}].name reference must be an explicit A1 or R1C1 range");
        }
    }

    private static void ValidatePivotSource(JsonObject op, int index, string opName, ICollection<string> errors,
        bool required)
    {
        var text = Json.GetString(op, "sourceRange");
        if (string.IsNullOrWhiteSpace(text))
        {
            if (required)
                errors.Add($"ops[{index}] '{opName}' sourceRange is required (A1 or ListObject name)");
            return;
        }
        if (ExcelDataOperationsContract.TryParseA1(text, out var box, allowUnion: false))
        {
            if (box.CellCount > 1_000_000)
                errors.Add($"ops[{index}] '{opName}' sourceRange exceeds 1,000,000 cells");
            return;
        }
        if (ExcelDataRangeBinding.LooksLikeTableSource(text))
            return;
        errors.Add($"ops[{index}] '{opName}' sourceRange must be a rectangular A1 address or a ListObject name");
    }

    private static void ValidateStringArray(JsonObject op, int index, string opName, string field,
        ICollection<string> errors)
    {
        if (!op.ContainsKey(field)) return;
        if (op[field] is not JsonArray array)
        {
            errors.Add($"ops[{index}] '{opName}' {field} must be an array of field names");
            return;
        }
        for (var i = 0; i < array.Count; i++)
        {
            if (array[i] is not JsonValue value || !value.TryGetValue<string>(out var text) ||
                string.IsNullOrWhiteSpace(text))
                errors.Add($"ops[{index}] '{opName}' {field}[{i}] must be a non-empty field name");
        }
    }
}
