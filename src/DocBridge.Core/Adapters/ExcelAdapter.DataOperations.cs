using System.Globalization;
using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

/// <summary>
/// Excel data/reporting writes: ListObject, sort/filter, names, validation,
/// conditional formats, native charts, pictures, notes, and hyperlinks.
/// Integration Cursor must call the Try* entry points from Preview/Apply.
/// </summary>
public sealed partial class ExcelAdapter
{
    internal const string DataObjectsRestoreMode = ExcelDataOperationsContract.RestoreMode;
    internal const int DataObjectsSnapshotVersion = ExcelDataOperationsContract.SnapshotVersion;
    private const int MaxDataSnapshotCells = 20_000;
    private const int MaxValidationSnapshotCells = 2_000;
    private const double DefaultChartLeft = 10;
    private const double DefaultChartTop = 10;
    private const double DefaultChartWidth = 400;
    private const double DefaultChartHeight = 250;
    private const double DefaultPictureLeft = 10;
    private const double DefaultPictureTop = 10;
    private const double DefaultPictureWidth = 160;
    private const double DefaultPictureHeight = 80;

    internal static bool IsDataOperation(string? op) => ExcelDataOperationsContract.IsDataOperation(op);

    internal static bool IsDataOnlySnapshot(IReadOnlyList<JsonObject>? ops) =>
        ops is { Count: > 0 } && ops.All(op => IsDataOperation(Json.GetString(op, "op")));

    internal static bool IsDataObjectsRestoreMode(string? restoreMode) =>
        string.Equals(restoreMode, DataObjectsRestoreMode, StringComparison.Ordinal);

    internal static JsonArray DataWriteOps() => Json.ToArray(ExcelDataOperationsContract.WriteOpNames);

    internal static bool TryPreviewDataOperation(object workbook, JsonObject op, ApplyPreview preview)
    {
        var name = Json.GetString(op, "op");
        if (!IsDataOperation(name)) return false;
        var errors = ExcelDataOperationsContract.ValidatePublicInput(op);
        if (errors.Count > 0)
        {
            foreach (var error in errors) preview.Errors.Add(error);
            return true;
        }
        try { PreviewDataOperation(workbook, op, preview); }
        catch (Exception ex) { preview.Errors.Add(DataComMessage(ex, "preview", name)); }
        return true;
    }

    internal static bool TryApplyDataOperation(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        var name = Json.GetString(op, "op");
        if (!IsDataOperation(name)) return false;
        var errors = ExcelDataOperationsContract.ValidatePublicInput(op);
        if (errors.Count > 0)
            throw new InvalidOperationException(string.Join("; ", errors));
        ApplyDataOperation(workbook, op, execution, mismatches, ref checkedCells);
        return true;
    }

    /// <summary>
    /// Native inventory for inspect/probe. Reads live COM objects only.
    /// </summary>
    internal static JsonObject ReadDataObjects(object workbook, JsonObject args)
    {
        var requestedSheet = Json.GetString(args, "sheet");
        var scope = ExcelDataOperationsContract.ResolveReadScope(args);
        var nameFilter = Json.GetString(args, "name");
        var rangeFilter = Json.GetString(args, "range");
        var limit = Math.Clamp(Json.GetInt(args, "limit") ?? 200, 1, 2000);
        var items = new JsonArray();
        var truncated = false;
        var rangeTruncated = false;

        void Add(JsonObject item)
        {
            if (items.Count >= limit) { truncated = true; return; }
            items.Add(item);
        }

        var includeAll = string.Equals(scope, "all", StringComparison.OrdinalIgnoreCase);
        foreach (var sheet in EnumerateWorksheets(workbook, requestedSheet))
        {
            try
            {
                var sheetName = ReadComString(sheet, "Name") ?? "";
                if (includeAll || scope == "tables")
                    foreach (var table in EnumerateListObjects(sheet))
                        try
                        {
                            var state = ReadListObjectState(table);
                            if (NameMatches(nameFilter, Json.GetString(state, "name")) &&
                                RangeIntersects(Json.GetString(state, "range"), rangeFilter))
                                Add(state);
                        }
                        finally { RotHelper.ReleaseComReference(table); }
                if (includeAll || scope == "charts")
                    foreach (var chart in EnumerateChartObjects(sheet))
                        try
                        {
                            var state = ReadChartState(chart, sheetName);
                            if (NameMatches(nameFilter, Json.GetString(state, "name"))) Add(state);
                        }
                        finally { RotHelper.ReleaseComReference(chart); }
                if (includeAll || scope == "pictures")
                    foreach (var shape in EnumeratePictures(sheet))
                        try
                        {
                            var state = ReadPictureState(shape, sheetName);
                            if (NameMatches(nameFilter, Json.GetString(state, "name"))) Add(state);
                        }
                        finally { RotHelper.ReleaseComReference(shape); }
                if (includeAll || scope is "shapes" or "connectors")
                    foreach (var shape in EnumerateDrawingShapes(sheet))
                        try
                        {
                            var state = ReadShapeState(shape, sheetName);
                            var isConnector = string.Equals(Json.GetString(state, "type"), "connector", StringComparison.Ordinal);
                            if ((scope != "connectors" || isConnector) && (scope != "shapes" || !isConnector) && NameMatches(nameFilter, Json.GetString(state, "name"))) Add(state);
                        }
                        finally { RotHelper.ReleaseComReference(shape); }
                if (includeAll || scope == "sparklines")
                    foreach (var spark in ReadSparklineStates(sheet, sheetName))
                    {
                        if (NameMatches(nameFilter, Json.GetString(spark, "location"))) Add(spark);
                    }
                if (includeAll || scope == "pivots")
                    foreach (var pivot in EnumeratePivotTables(sheet))
                        try
                        {
                            var state = ReadPivotState(pivot, sheetName);
                            if (NameMatches(nameFilter, Json.GetString(state, "name"))) Add(state);
                        }
                        finally { RotHelper.ReleaseComReference(pivot); }
                if (includeAll || scope == "filters")
                    Add(ReadAutoFilterState(sheet, sheetName));
                if (scope == "richText")
                {
                    if (string.IsNullOrWhiteSpace(requestedSheet))
                        throw new InvalidOperationException("richText inspect requires sheet and one requested cell");
                    if (string.IsNullOrWhiteSpace(rangeFilter))
                        throw new InvalidOperationException("richText inspect requires one requested cell in range");
                    if (!ExcelDataOperationsContract.TryParseA1(rangeFilter, out var richCell) || richCell.Rows != 1 || richCell.Columns != 1)
                        throw new InvalidOperationException("richText inspect range must be one cell");
                    Add(ReadRichText(sheet, rangeFilter));
                }
                if (includeAll || scope is "validations" or "conditionalFormats" or "notes" or "hyperlinks")
                    rangeTruncated |= ReadBoundedCellObjects(sheet, sheetName, scope, includeAll, rangeFilter, Add);
            }
            finally { RotHelper.ReleaseComReference(sheet); }
        }

        if (includeAll || scope == "names")
        {
            foreach (var defined in EnumerateNames(workbook))
            {
                try
                {
                    var state = ReadDefinedNameState(defined);
                    if (NameMatches(nameFilter, Json.GetString(state, "name"))) Add(state);
                }
                finally { RotHelper.ReleaseComReference(defined); }
            }
        }

        if (includeAll || scope == "links")
        {
            foreach (var link in ReadExcelLinks(workbook))
            {
                if (NameMatches(nameFilter, Json.GetString(link, "source"))) Add(link);
            }
        }

        if (includeAll || scope == "slicers")
        {
            foreach (var cache in ReadSlicerCacheStates(workbook))
            {
                if (NameMatches(nameFilter, Json.GetString(cache, "name"))) Add(cache);
            }
        }

        if (includeAll || scope == "cellStyles")
        {
            foreach (var style in ReadCellStyleStates(workbook))
            {
                if (NameMatches(nameFilter, Json.GetString(style, "name"))) Add(style);
            }
        }

        return new JsonObject
        {
            ["ok"] = true,
            ["app"] = "excel",
            ["scope"] = scope,
            ["objectKind"] = Json.GetString(args, "objectKind") ?? Json.GetString(args, "objectScope"),
            ["objects"] = items,
            ["coverage"] = new JsonObject
            {
                ["limit"] = limit,
                ["returned"] = items.Count,
                ["truncated"] = truncated || rangeTruncated,
                ["rangeTruncated"] = rangeTruncated,
                ["complete"] = !truncated && !rangeTruncated,
            },
        };
    }

    private static void PreviewDataOperation(object workbook, JsonObject op, ApplyPreview preview)
    {
        var name = Json.GetString(op, "op")!;
        switch (name)
        {
            case ExcelDataOperationsContract.CreateTable: PreviewCreateTable(workbook, op, preview); break;
            case ExcelDataOperationsContract.ResizeTable: PreviewResizeTable(workbook, op, preview); break;
            case ExcelDataOperationsContract.StyleTable: PreviewStyleTable(workbook, op, preview); break;
            case ExcelDataOperationsContract.SetTableTotals: PreviewTableTotals(workbook, op, preview); break;
            case ExcelDataOperationsContract.SortRange: PreviewSortRange(workbook, op, preview); break;
            case ExcelDataOperationsContract.SortTable: PreviewSortTable(workbook, op, preview); break;
            case ExcelDataOperationsContract.SetAutoFilter: PreviewSetAutoFilter(workbook, op, preview); break;
            case ExcelDataOperationsContract.ClearAutoFilter: PreviewClearAutoFilter(workbook, op, preview); break;
            case ExcelDataOperationsContract.DefineName: PreviewDefineName(workbook, op, preview, update: false); break;
            case ExcelDataOperationsContract.UpdateName: PreviewDefineName(workbook, op, preview, update: true); break;
            case ExcelDataOperationsContract.DeleteName: PreviewDeleteName(workbook, op, preview); break;
            case ExcelDataOperationsContract.SetDataValidation: PreviewSetValidation(workbook, op, preview); break;
            case ExcelDataOperationsContract.ClearDataValidation: PreviewClearValidation(workbook, op, preview); break;
            case ExcelDataOperationsContract.AddConditionalFormat: PreviewAddConditional(workbook, op, preview); break;
            case ExcelDataOperationsContract.ClearConditionalFormats: PreviewClearConditional(workbook, op, preview); break;
            case ExcelDataOperationsContract.SetRichText: PreviewSetRichText(workbook, op, preview); break;
            case ExcelDataOperationsContract.UpdateConditionalFormat: PreviewUpdateConditional(workbook, op, preview); break;
            case ExcelDataOperationsContract.DeleteConditionalFormat: PreviewDeleteConditional(workbook, op, preview); break;
            case ExcelDataOperationsContract.AppendTableRows:
            case ExcelDataOperationsContract.InsertTableRows:
            case ExcelDataOperationsContract.DeleteTableRows: PreviewTableRows(workbook, op, preview); break;
            case ExcelDataOperationsContract.CreateChart: PreviewCreateChart(workbook, op, preview); break;
            case ExcelDataOperationsContract.UpdateChart: PreviewUpdateChart(workbook, op, preview); break;
            case ExcelDataOperationsContract.InsertSheetPicture: PreviewInsertPicture(workbook, op, preview); break;
            case ExcelDataOperationsContract.UpdatePicture: PreviewUpdatePicture(workbook, op, preview); break;
            case ExcelDataOperationsContract.SetCellNote: PreviewSetNote(workbook, op, preview); break;
            case ExcelDataOperationsContract.ClearCellNote: PreviewClearNote(workbook, op, preview); break;
            case ExcelDataOperationsContract.SetHyperlink: PreviewSetHyperlink(workbook, op, preview); break;
            case ExcelDataOperationsContract.ClearHyperlink: PreviewClearHyperlink(workbook, op, preview); break;
            case ExcelDataOperationsContract.AddTableColumn: PreviewAddTableColumn(workbook, op, preview); break;
            case ExcelDataOperationsContract.DeleteTable: PreviewDeleteTable(workbook, op, preview); break;
            case ExcelDataOperationsContract.RemoveDuplicates: PreviewRemoveDuplicates(workbook, op, preview); break;
            case ExcelDataOperationsContract.TextToColumns: PreviewTextToColumns(workbook, op, preview); break;
            case ExcelDataOperationsContract.DeleteChart: PreviewDeleteChart(workbook, op, preview); break;
            case ExcelDataOperationsContract.DeletePicture: PreviewDeletePicture(workbook, op, preview); break;
            case ExcelDataOperationsContract.CreatePivot: PreviewCreatePivot(workbook, op, preview); break;
            case ExcelDataOperationsContract.UpdatePivot: PreviewUpdatePivot(workbook, op, preview); break;
            case ExcelDataOperationsContract.RefreshPivot: PreviewRefreshPivot(workbook, op, preview); break;
            case ExcelDataOperationsContract.DeletePivot: PreviewDeletePivot(workbook, op, preview); break;
            case ExcelDataOperationsContract.InsertShape: PreviewInsertShape(workbook, op, preview); break;
            case ExcelDataOperationsContract.UpdateShape: PreviewUpdateShape(workbook, op, preview); break;
            case ExcelDataOperationsContract.DeleteShape: PreviewDeleteShape(workbook, op, preview); break;
            case ExcelDataOperationsContract.InsertTextbox: PreviewInsertTextbox(workbook, op, preview); break;
            case ExcelDataOperationsContract.UpdateTextbox: PreviewUpdateTextbox(workbook, op, preview); break;
            case ExcelDataOperationsContract.DeleteTextbox: PreviewDeleteTextbox(workbook, op, preview); break;
            case ExcelDataOperationsContract.CreateConnector: PreviewCreateConnector(workbook, op, preview); break;
            case ExcelDataOperationsContract.UpdateConnector: PreviewUpdateConnector(workbook, op, preview); break;
            case ExcelDataOperationsContract.UpdateExternalLinks: PreviewUpdateExternalLinks(workbook, op, preview); break;
            case ExcelDataOperationsContract.ChangeLinkSource: PreviewChangeLinkSource(workbook, op, preview); break;
            case ExcelDataOperationsContract.BreakExternalLink: PreviewBreakExternalLink(workbook, op, preview); break;
            case ExcelDataOperationsContract.SetCalculationMode: PreviewSetCalculationMode(workbook, op, preview); break;
            case ExcelDataOperationsContract.FreezeValues: PreviewFreezeValues(workbook, op, preview); break;
            case ExcelDataOperationsContract.PasteSpecial: PreviewPasteSpecial(workbook, op, preview); break;
            case ExcelDataOperationsContract.GoalSeek: PreviewGoalSeek(workbook, op, preview); break;
            case ExcelDataOperationsContract.ProtectWorkbook: PreviewProtectWorkbook(workbook, op, preview); break;
            case ExcelDataOperationsContract.UnprotectWorkbook: PreviewUnprotectWorkbook(workbook, op, preview); break;
            case ExcelDataOperationsContract.SetSplitPanes: PreviewSetSplitPanes(workbook, op, preview); break;
            case ExcelDataOperationsContract.CreateSparkline: PreviewCreateSparkline(workbook, op, preview); break;
            case ExcelDataOperationsContract.UpdateSparkline: PreviewUpdateSparkline(workbook, op, preview); break;
            case ExcelDataOperationsContract.DeleteSparkline: PreviewDeleteSparkline(workbook, op, preview); break;
            case ExcelDataOperationsContract.CreateSlicer: PreviewCreateSlicer(workbook, op, preview); break;
            case ExcelDataOperationsContract.DeleteSlicer: PreviewDeleteSlicer(workbook, op, preview); break;
            case ExcelDataOperationsContract.ApplyCellStyle: PreviewApplyCellStyle(workbook, op, preview); break;
            default:
                preview.Errors.Add($"[EXCEL_DATA_UNSUPPORTED] '{name}' is not implemented as a native data operation");
                break;
        }
    }

    private static void ApplyDataOperation(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        var name = Json.GetString(op, "op")!;
        switch (name)
        {
            case ExcelDataOperationsContract.CreateTable: ApplyCreateTable(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.ResizeTable: ApplyResizeTable(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.StyleTable: ApplyStyleTable(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.SetTableTotals: ApplyTableTotals(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.SortRange: ApplySortRange(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.SortTable: ApplySortTable(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.SetAutoFilter: ApplySetAutoFilter(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.ClearAutoFilter: ApplyClearAutoFilter(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.DefineName: ApplyDefineName(workbook, op, execution, mismatches, ref checkedCells, update: false); break;
            case ExcelDataOperationsContract.UpdateName: ApplyDefineName(workbook, op, execution, mismatches, ref checkedCells, update: true); break;
            case ExcelDataOperationsContract.DeleteName: ApplyDeleteName(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.SetDataValidation: ApplySetValidation(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.ClearDataValidation: ApplyClearValidation(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.AddConditionalFormat: ApplyAddConditional(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.ClearConditionalFormats: ApplyClearConditional(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.SetRichText: ApplySetRichText(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.UpdateConditionalFormat: ApplyUpdateConditional(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.DeleteConditionalFormat: ApplyDeleteConditional(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.AppendTableRows:
            case ExcelDataOperationsContract.InsertTableRows:
            case ExcelDataOperationsContract.DeleteTableRows: ApplyTableRows(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.CreateChart: ApplyCreateChart(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.UpdateChart: ApplyUpdateChart(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.InsertSheetPicture: ApplyInsertPicture(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.UpdatePicture: ApplyUpdatePicture(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.SetCellNote: ApplySetNote(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.ClearCellNote: ApplyClearNote(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.SetHyperlink: ApplySetHyperlink(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.ClearHyperlink: ApplyClearHyperlink(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.AddTableColumn: ApplyAddTableColumn(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.DeleteTable: ApplyDeleteTable(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.RemoveDuplicates: ApplyRemoveDuplicates(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.TextToColumns: ApplyTextToColumns(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.DeleteChart: ApplyDeleteChart(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.DeletePicture: ApplyDeletePicture(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.CreatePivot: ApplyCreatePivot(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.UpdatePivot: ApplyUpdatePivot(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.RefreshPivot: ApplyRefreshPivot(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.DeletePivot: ApplyDeletePivot(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.InsertShape: ApplyInsertShape(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.UpdateShape: ApplyUpdateShape(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.DeleteShape: ApplyDeleteShape(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.InsertTextbox: ApplyInsertTextbox(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.UpdateTextbox: ApplyUpdateTextbox(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.DeleteTextbox: ApplyDeleteTextbox(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.CreateConnector: ApplyCreateConnector(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.UpdateConnector: ApplyUpdateConnector(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.UpdateExternalLinks: ApplyUpdateExternalLinks(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.ChangeLinkSource: ApplyChangeLinkSource(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.BreakExternalLink: ApplyBreakExternalLink(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.SetCalculationMode: ApplySetCalculationMode(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.FreezeValues: ApplyFreezeValues(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.PasteSpecial: ApplyPasteSpecial(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.GoalSeek: ApplyGoalSeek(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.ProtectWorkbook: ApplyProtectWorkbook(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.UnprotectWorkbook: ApplyUnprotectWorkbook(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.SetSplitPanes: ApplySetSplitPanes(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.CreateSparkline: ApplyCreateSparkline(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.UpdateSparkline: ApplyUpdateSparkline(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.DeleteSparkline: ApplyDeleteSparkline(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.CreateSlicer: ApplyCreateSlicer(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.DeleteSlicer: ApplyDeleteSlicer(workbook, op, execution, mismatches, ref checkedCells); break;
            case ExcelDataOperationsContract.ApplyCellStyle: ApplyApplyCellStyle(workbook, op, execution, mismatches, ref checkedCells); break;
            default:
                throw new InvalidOperationException(
                    $"[EXCEL_DATA_UNSUPPORTED] '{name}' is not implemented as a native data operation");
        }
    }

    private static void PreviewCreateTable(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var range = BindRange(workbook, op, "range");
        EnsureSheetWritable(range.Sheet);
        var existing = FindOverlappingTable(range.Sheet, range.Address);
        if (existing is not null)
        {
            preview.Errors.Add(
                $"[EXCEL_TABLE_OVERLAP] {range.SheetName}!{range.Address} already belongs to ListObject '{existing}'");
            return;
        }
        var requestedName = Json.GetString(op, "name");
        if (!string.IsNullOrWhiteSpace(requestedName) && FindListObject(range.Sheet, requestedName) is { } conflict)
        {
            RotHelper.ReleaseComReference(conflict);
            preview.Errors.Add($"[EXCEL_TABLE_EXISTS] ListObject '{requestedName}' already exists on '{range.SheetName}'");
            return;
        }
        preview.Affected.Add(new AffectedRef("table", $"{range.SheetName}!{range.Address}"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{range.SheetName}!{range.Address}:listObject",
            Before = null,
            After = new JsonObject
            {
                ["name"] = requestedName,
                ["hasHeaders"] = Json.GetBool(op, "hasHeaders", true),
                ["styleName"] = Json.GetString(op, "styleName"),
            },
        });
    }

    private static void ApplyCreateTable(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var range = BindRange(workbook, op, "range");
        EnsureSheetWritable(range.Sheet);
        object? listObjects = null;
        object? source = null;
        object? table = null;
        try
        {
            source = (object)((dynamic)range.Sheet).Range(range.Address);
            listObjects = (object)((dynamic)range.Sheet).ListObjects;
            table = (object)((dynamic)listObjects).Add(
                ExcelDataObjectCatalog.XlSrcRange,
                source,
                Type.Missing,
                ExcelDataObjectCatalog.HasHeaders(Json.GetBool(op, "hasHeaders", true)));
            var requestedName = Json.GetString(op, "name");
            if (!string.IsNullOrWhiteSpace(requestedName))
                ((dynamic)table).Name = requestedName;
            var styleName = Json.GetString(op, "styleName");
            if (!string.IsNullOrWhiteSpace(styleName))
                ((dynamic)table).TableStyle = styleName;
            if (op.ContainsKey("showTotals"))
                ((dynamic)table).ShowTotals = Json.GetBool(op, "showTotals");
            ApplyTotalsColumns(table, op);
            var actual = ReadListObjectState(table);
            checkedCells += VerifyTable(actual, op, range, mismatches);
            execution.Affected.Add(new AffectedRef("table", $"{range.SheetName}!{Json.GetString(actual, "name")}"));
        }
        finally
        {
            RotHelper.ReleaseComReference(table);
            RotHelper.ReleaseComReference(source);
            RotHelper.ReleaseComReference(listObjects);
        }
    }

    private static void PreviewResizeTable(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var table = BindTable(workbook, op);
        using var range = BindRange(workbook, op, "range");
        EnsureSheetWritable(table.Sheet);
        preview.Affected.Add(new AffectedRef("table", $"{table.SheetName}!{table.Name}"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{table.SheetName}!{table.Name}:range",
            Before = table.State["range"]?.DeepClone(),
            After = JsonValue.Create(range.Address),
        });
    }

    private static void ApplyResizeTable(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var table = BindTable(workbook, op);
        using var range = BindRange(workbook, op, "range");
        EnsureSheetWritable(table.Sheet);
        object? dest = null;
        try
        {
            dest = (object)((dynamic)range.Sheet).Range(range.Address);
            ((dynamic)table.ListObject).Resize(dest);
            var actual = ReadListObjectState(table.ListObject);
            checkedCells++;
            var actualRange = Json.GetString(actual, "range");
            if (!AddressesEqual(actualRange, range.Address))
                mismatches.Add($"{table.SheetName}!{table.Name}: resize readback {actualRange} != {range.Address}");
            execution.Affected.Add(new AffectedRef("table", $"{table.SheetName}!{table.Name}"));
        }
        finally { RotHelper.ReleaseComReference(dest); }
    }

    private static void PreviewStyleTable(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var table = BindTable(workbook, op);
        EnsureSheetWritable(table.Sheet);
        preview.Affected.Add(new AffectedRef("table", $"{table.SheetName}!{table.Name}"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{table.SheetName}!{table.Name}:style",
            Before = table.State.DeepClone(),
            After = op.DeepClone(),
        });
    }

    private static void ApplyStyleTable(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var table = BindTable(workbook, op);
        EnsureSheetWritable(table.Sheet);
        var styleName = Json.GetString(op, "styleName");
        if (!string.IsNullOrWhiteSpace(styleName))
            ((dynamic)table.ListObject).TableStyle = styleName;
        SetOptionalBool(table.ListObject, op, "showHeaders", "ShowHeaders");
        SetOptionalBool(table.ListObject, op, "showTotals", "ShowTotals");
        SetOptionalBool(table.ListObject, op, "showAutoFilter", "ShowAutoFilter");
        SetOptionalBool(table.ListObject, op, "showTableStyleRowStripes", "ShowRowStripes", "ShowTableStyleRowStripes");
        SetOptionalBool(table.ListObject, op, "showRowStripes", "ShowTableStyleRowStripes");
        SetOptionalBool(table.ListObject, op, "showColumnStripes", "ShowTableStyleColumnStripes");
        var actual = ReadListObjectState(table.ListObject);
        checkedCells += VerifyTableStyle(actual, op, table, mismatches);
        execution.Affected.Add(new AffectedRef("table", $"{table.SheetName}!{table.Name}"));
    }

    private static void PreviewTableTotals(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var table = BindTable(workbook, op);
        EnsureSheetWritable(table.Sheet);
        preview.Affected.Add(new AffectedRef("table", $"{table.SheetName}!{table.Name}:totals"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{table.SheetName}!{table.Name}:totals",
            Before = table.State["totals"]?.DeepClone(),
            After = op["columns"]?.DeepClone() ?? JsonValue.Create(Json.GetBool(op, "showTotals")),
        });
    }

    private static void ApplyTableTotals(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var table = BindTable(workbook, op);
        EnsureSheetWritable(table.Sheet);
        ((dynamic)table.ListObject).ShowTotals = Json.GetBool(op, "showTotals");
        if (Json.GetBool(op, "showTotals")) ApplyTotalsColumns(table.ListObject, op);
        var actual = ReadListObjectState(table.ListObject);
        checkedCells++;
        if (Json.GetBool(actual, "showTotals") != Json.GetBool(op, "showTotals"))
            mismatches.Add($"{table.SheetName}!{table.Name}: ShowTotals readback mismatch");
        ExcelDataReadbackContract.CompareRequestedTotals(
            actual, op, $"{table.SheetName}!{table.Name}", mismatches);
        execution.Affected.Add(new AffectedRef("table", $"{table.SheetName}!{table.Name}:totals"));
    }

    private static void PreviewSortRange(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var range = BindRange(workbook, op, "range");
        EnsureSheetWritable(range.Sheet);
        preview.Affected.Add(new AffectedRef("range", $"{range.SheetName}!{range.Address}"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{range.SheetName}!{range.Address}:sort",
            Before = CaptureRangeFormulas(range.Sheet, range.Address),
            After = op["keys"]?.DeepClone(),
        });
    }

    private static void ApplySortRange(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var range = BindRange(workbook, op, "range");
        EnsureSheetWritable(range.Sheet);
        object? comRange = null;
        try
        {
            comRange = (object)((dynamic)range.Sheet).Range(range.Address);
            var before = CaptureRangeValues(range.Sheet, range.Address);
            var keys = ResolveSortKeys(range.Sheet, range.Address, Json.GetArr(op, "keys")!, table: false);
            InvokeRangeSort(comRange, keys, Json.GetBool(op, "hasHeaders", true));
            checkedCells++;
            execution.Affected.Add(new AffectedRef("range", $"{range.SheetName}!{range.Address}"));
            if (!SortApplied(comRange))
                mismatches.Add($"{range.SheetName}!{range.Address}: Range.Sort completed but the range could not be re-read");
            else
                VerifySortOrder(range.Sheet, range.Address, Json.GetArr(op, "keys")!,
                    Json.GetBool(op, "hasHeaders", true), mismatches, before);
        }
        finally { RotHelper.ReleaseComReference(comRange); }
    }

    private static void PreviewSortTable(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var table = BindTable(workbook, op);
        EnsureSheetWritable(table.Sheet);
        preview.Affected.Add(new AffectedRef("table", $"{table.SheetName}!{table.Name}:sort"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{table.SheetName}!{table.Name}:sort",
            Before = table.State["range"]?.DeepClone(),
            After = op["keys"]?.DeepClone(),
        });
    }

    private static void ApplySortTable(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var table = BindTable(workbook, op);
        EnsureSheetWritable(table.Sheet);
        object? sort = null;
        object? fields = null;
        try
        {
            sort = (object)((dynamic)table.ListObject).Sort;
            fields = (object)((dynamic)sort).SortFields;
            ((dynamic)fields).Clear();
            var hasBody = !string.IsNullOrWhiteSpace(Json.GetString(table.State, "dataRange"));
            var dataRange = Json.GetString(table.State, "dataRange") ?? Json.GetString(table.State, "range")!;
            var before = CaptureRangeValues(table.Sheet, dataRange);
            var keys = ResolveSortKeys(table.Sheet, dataRange, Json.GetArr(op, "keys")!, table: true, tableObject: table.ListObject);
            foreach (var key in keys)
            {
                object? keyRange = null;
                try
                {
                    keyRange = (object)((dynamic)table.Sheet).Range(key.Address);
                    ((dynamic)fields).Add(keyRange, 0, key.Order);
                }
                finally { RotHelper.ReleaseComReference(keyRange); }
            }
            ((dynamic)sort).Header = ExcelDataObjectCatalog.XlYes;
            ((dynamic)sort).Apply();
            checkedCells++;
            var actual = ReadListObjectState(table.ListObject);
            if (string.IsNullOrWhiteSpace(Json.GetString(actual, "name")))
                mismatches.Add($"{table.SheetName}!{table.Name}: table missing after Sort.Apply");
            else
                VerifySortOrder(table.Sheet,
                    hasBody
                        ? Json.GetString(actual, "dataRange") ?? dataRange
                        : Json.GetString(actual, "range") ?? dataRange,
                    Json.GetArr(op, "keys")!,
                    hasHeaders: !hasBody, mismatches, before, tableObject: table.ListObject);
            execution.Affected.Add(new AffectedRef("table", $"{table.SheetName}!{table.Name}:sort"));
        }
        finally
        {
            RotHelper.ReleaseComReference(fields);
            RotHelper.ReleaseComReference(sort);
        }
    }

    private static void PreviewSetAutoFilter(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var target = BindFilterTarget(workbook, op);
        EnsureSheetWritable(target.Sheet);
        preview.Affected.Add(new AffectedRef("filter", $"{target.SheetName}!{target.Address}"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{target.SheetName}!{target.Address}:autoFilter",
            Before = ReadAutoFilterState(target.Sheet, target.SheetName),
            After = op["criteria"]?.DeepClone(),
        });
    }

    private static void ApplySetAutoFilter(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var target = BindFilterTarget(workbook, op);
        EnsureSheetWritable(target.Sheet);
        object? comRange = null;
        try
        {
            comRange = (object)((dynamic)target.Sheet).Range(target.Address);
            if (!Convert.ToBoolean(((dynamic)target.Sheet).AutoFilterMode, CultureInfo.InvariantCulture) &&
                target.ListObject is null)
            {
                ((dynamic)comRange).AutoFilter();
            }
            foreach (var criterion in Json.GetArr(op, "criteria")!.OfType<JsonObject>())
                ApplyOneFilter(comRange, target, criterion);
            checkedCells++;
            var filterMode = Convert.ToBoolean(((dynamic)target.Sheet).FilterMode, CultureInfo.InvariantCulture)
                             || Convert.ToBoolean(((dynamic)target.Sheet).AutoFilterMode, CultureInfo.InvariantCulture)
                             || target.ListObject is not null;
            if (!filterMode)
                mismatches.Add($"{target.SheetName}!{target.Address}: AutoFilter was not present after apply");
            execution.Affected.Add(new AffectedRef("filter", $"{target.SheetName}!{target.Address}"));
        }
        finally { RotHelper.ReleaseComReference(comRange); }
    }

    private static void PreviewClearAutoFilter(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var target = BindFilterTarget(workbook, op);
        preview.Affected.Add(new AffectedRef("filter", $"{target.SheetName}!{target.Address}"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{target.SheetName}!{target.Address}:autoFilter",
            Before = ReadAutoFilterState(target.Sheet, target.SheetName),
            After = JsonValue.Create("cleared"),
        });
    }

    private static void ApplyClearAutoFilter(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var target = BindFilterTarget(workbook, op);
        object? autoFilter = null;
        try
        {
            if (target.ListObject is not null)
            {
                try { autoFilter = (object)((dynamic)target.ListObject).AutoFilter; }
                catch { /* table may not expose AutoFilter until shown */ }
                if (autoFilter is not null) ((dynamic)autoFilter).ShowAllData();
            }
            else if (Convert.ToBoolean(((dynamic)target.Sheet).FilterMode, CultureInfo.InvariantCulture))
            {
                ((dynamic)target.Sheet).ShowAllData();
            }
            checkedCells++;
            if (Convert.ToBoolean(((dynamic)target.Sheet).FilterMode, CultureInfo.InvariantCulture))
                mismatches.Add($"{target.SheetName}!{target.Address}: FilterMode still true after clear");
            execution.Affected.Add(new AffectedRef("filter", $"{target.SheetName}!{target.Address}"));
        }
        finally { RotHelper.ReleaseComReference(autoFilter); }
    }

    private static void PreviewDefineName(object workbook, JsonObject op, ApplyPreview preview, bool update)
    {
        var name = Json.GetString(op, "name")!;
        var scope = Json.GetString(op, "scope")!;
        var existing = FindDefinedName(workbook, name, scope, Json.GetString(Json.GetObj(op, "target"), "sheet"));
        try
        {
            if (update && existing is null)
            {
                preview.Errors.Add($"[EXCEL_NAME_NOT_FOUND] defined name '{name}' ({scope}) does not exist");
                return;
            }
            if (!update && existing is not null && !Json.GetBool(op, "replace"))
            {
                preview.Errors.Add($"[EXCEL_NAME_EXISTS] defined name '{name}' already exists; set replace=true to overwrite");
                return;
            }
            preview.Affected.Add(new AffectedRef("definedName", NameRef(scope, name, op)));
            preview.Diff.Add(new DiffEntry
            {
                Ref = NameRef(scope, name, op),
                Before = existing is null ? null : ReadDefinedNameState(existing),
                After = JsonValue.Create(Json.GetString(op, "refersTo")),
            });
        }
        finally { RotHelper.ReleaseComReference(existing); }
    }

    private static void ApplyDefineName(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells, bool update)
    {
        var name = Json.GetString(op, "name")!;
        var scope = Json.GetString(op, "scope")!;
        var refersTo = Json.GetString(op, "refersTo")!;
        var sheetName = Json.GetString(Json.GetObj(op, "target"), "sheet");
        var existing = FindDefinedName(workbook, name, scope, sheetName);
        object? created = null;
        try
        {
            if (update)
            {
                if (existing is null)
                    throw new InvalidOperationException($"[EXCEL_NAME_NOT_FOUND] defined name '{name}' ({scope}) does not exist");
                ((dynamic)existing).RefersTo = refersTo;
                var comment = Json.GetString(op, "comment");
                if (comment is not null) ((dynamic)existing).Comment = comment;
                created = existing;
                existing = null;
            }
            else
            {
                if (existing is not null)
                {
                    if (!Json.GetBool(op, "replace"))
                        throw new InvalidOperationException($"[EXCEL_NAME_EXISTS] defined name '{name}' already exists");
                    ((dynamic)existing).RefersTo = refersTo;
                    created = existing;
                    existing = null;
                }
                else if (scope == "sheet")
                {
                    object? sheet = null;
                    object? names = null;
                    try
                    {
                        sheet = GetExplicitTargetSheetReference(workbook, op);
                        names = (object)((dynamic)sheet).Names;
                        created = (object)((dynamic)names).Add(name, refersTo);
                    }
                    finally
                    {
                        RotHelper.ReleaseComReference(names);
                        RotHelper.ReleaseComReference(sheet);
                    }
                }
                else
                {
                    object? names = null;
                    try
                    {
                        names = (object)((dynamic)workbook).Names;
                        created = (object)((dynamic)names).Add(name, refersTo);
                    }
                    finally { RotHelper.ReleaseComReference(names); }
                }
            }
            var actual = ReadDefinedNameState(created!);
            checkedCells++;
            if (!string.Equals(Json.GetString(actual, "scope"), scope, StringComparison.OrdinalIgnoreCase))
                mismatches.Add($"{name}: scope readback '{Json.GetString(actual, "scope")}' != '{scope}'");
            if (scope == "sheet" && !string.IsNullOrWhiteSpace(sheetName) &&
                !string.Equals(Json.GetString(actual, "sheet"), sheetName, StringComparison.OrdinalIgnoreCase))
                mismatches.Add($"{name}: sheet identity '{Json.GetString(actual, "sheet")}' != '{sheetName}'");
            if (!RefersToMatches(Json.GetString(actual, "refersTo"), refersTo))
                mismatches.Add($"{name}: RefersTo readback '{Json.GetString(actual, "refersTo")}' != '{refersTo}'");
            execution.Affected.Add(new AffectedRef("definedName", NameRef(scope, name, op)));
        }
        finally
        {
            RotHelper.ReleaseComReference(created);
            RotHelper.ReleaseComReference(existing);
        }
    }

    private static void PreviewDeleteName(object workbook, JsonObject op, ApplyPreview preview)
    {
        var name = Json.GetString(op, "name")!;
        var scope = Json.GetString(op, "scope")!;
        var existing = FindDefinedName(workbook, name, scope, Json.GetString(Json.GetObj(op, "target"), "sheet"));
        try
        {
            if (existing is null)
            {
                preview.Errors.Add($"[EXCEL_NAME_NOT_FOUND] defined name '{name}' ({scope}) does not exist");
                return;
            }
            preview.Affected.Add(new AffectedRef("definedName", NameRef(scope, name, op)));
            preview.Diff.Add(new DiffEntry
            {
                Ref = NameRef(scope, name, op),
                Before = ReadDefinedNameState(existing),
                After = null,
            });
        }
        finally { RotHelper.ReleaseComReference(existing); }
    }

    private static void ApplyDeleteName(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        var name = Json.GetString(op, "name")!;
        var scope = Json.GetString(op, "scope")!;
        var existing = FindDefinedName(workbook, name, scope, Json.GetString(Json.GetObj(op, "target"), "sheet"));
        try
        {
            if (existing is null)
                throw new InvalidOperationException($"[EXCEL_NAME_NOT_FOUND] defined name '{name}' ({scope}) does not exist");
            ((dynamic)existing).Delete();
        }
        finally { RotHelper.ReleaseComReference(existing); }
        checkedCells++;
        if (FindDefinedName(workbook, name, scope, Json.GetString(Json.GetObj(op, "target"), "sheet")) is { } leftover)
        {
            RotHelper.ReleaseComReference(leftover);
            mismatches.Add($"{name}: defined name still present after Delete");
        }
        execution.Affected.Add(new AffectedRef("definedName", NameRef(scope, name, op)));
    }

    private static void PreviewSetValidation(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var range = BindRange(workbook, op, "range");
        EnsureSheetWritable(range.Sheet);
        preview.Affected.Add(new AffectedRef("validation", $"{range.SheetName}!{range.Address}"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{range.SheetName}!{range.Address}:validation",
            Before = ReadValidationState(range.Sheet, range.Address),
            After = new JsonObject { ["type"] = Json.GetString(op, "type"), ["formula1"] = Json.GetString(op, "formula1") ?? Json.GetString(op, "source") },
        });
    }

    private static void ApplySetValidation(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var range = BindRange(workbook, op, "range");
        EnsureSheetWritable(range.Sheet);
        object? comRange = null;
        object? validation = null;
        try
        {
            comRange = (object)((dynamic)range.Sheet).Range(range.Address);
            validation = (object)((dynamic)comRange).Validation;
            try { ((dynamic)validation).Delete(); }
            catch { /* no prior validation */ }
            if (!ExcelDataObjectCatalog.TryValidationType(Json.GetString(op, "type"), out var type))
                throw new InvalidOperationException("[EXCEL_DATA_UNSUPPORTED] validation type is not mapped");
            var alert = ExcelDataObjectCatalog.TryErrorStyle(Json.GetString(op, "errorStyle"), out var alertStyle)
                ? alertStyle
                : ExcelDataObjectCatalog.XlValidAlertStop;
            var comparison = ExcelDataObjectCatalog.TryComparisonOperator(Json.GetString(op, "operator"), out var opValue)
                ? opValue
                : ExcelDataObjectCatalog.XlBetween;
            var formula1 = ResolveValidationFormula(op, "formula1") ?? ResolveListSource(range.SheetName, op);
            var formula2 = ResolveValidationFormula(op, "formula2");
            ((dynamic)validation).Add(type, alert, comparison, formula1 ?? Type.Missing, formula2 ?? Type.Missing);
            if (op.ContainsKey("ignoreBlank")) ((dynamic)validation).IgnoreBlank = Json.GetBool(op, "ignoreBlank", true);
            if (op.ContainsKey("inCellDropdown")) ((dynamic)validation).InCellDropdown = Json.GetBool(op, "inCellDropdown", true);
            if (op.ContainsKey("showInput")) ((dynamic)validation).ShowInput = Json.GetBool(op, "showInput");
            if (op.ContainsKey("showError")) ((dynamic)validation).ShowError = Json.GetBool(op, "showError", true);
            SetOptionalString(validation, op, "inputTitle", "InputTitle");
            SetOptionalString(validation, op, "inputMessage", "InputMessage");
            SetOptionalString(validation, op, "errorTitle", "ErrorTitle");
            SetOptionalString(validation, op, "errorMessage", "ErrorMessage");
            var actual = ReadValidationState(range.Sheet, range.Address);
            checkedCells++;
            if ((Json.GetInt(actual, "validationType") ?? Json.GetInt(actual, "type")) != type)
                mismatches.Add($"{range.SheetName}!{range.Address}: Validation.Type readback {Json.GetInt(actual, "validationType")} != {type}");
            execution.Affected.Add(new AffectedRef("validation", $"{range.SheetName}!{range.Address}"));
        }
        finally
        {
            RotHelper.ReleaseComReference(validation);
            RotHelper.ReleaseComReference(comRange);
        }
    }

    private static void PreviewClearValidation(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var range = BindRange(workbook, op, "range");
        preview.Affected.Add(new AffectedRef("validation", $"{range.SheetName}!{range.Address}"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{range.SheetName}!{range.Address}:validation",
            Before = ReadValidationState(range.Sheet, range.Address),
            After = null,
        });
    }

    private static void ApplyClearValidation(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var range = BindRange(workbook, op, "range");
        object? comRange = null;
        object? validation = null;
        try
        {
            comRange = (object)((dynamic)range.Sheet).Range(range.Address);
            validation = (object)((dynamic)comRange).Validation;
            ((dynamic)validation).Delete();
            checkedCells++;
            var actual = ReadValidationState(range.Sheet, range.Address);
            if (Json.GetBool(actual, "present"))
                mismatches.Add($"{range.SheetName}!{range.Address}: validation still present after Delete");
            execution.Affected.Add(new AffectedRef("validation", $"{range.SheetName}!{range.Address}"));
        }
        finally
        {
            RotHelper.ReleaseComReference(validation);
            RotHelper.ReleaseComReference(comRange);
        }
    }

    private static void PreviewAddConditional(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var range = BindRange(workbook, op, "range");
        EnsureSheetWritable(range.Sheet);
        preview.Affected.Add(new AffectedRef("conditionalFormat", $"{range.SheetName}!{range.Address}"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{range.SheetName}!{range.Address}:formatConditions",
            Before = ReadConditionalFormats(range.Sheet, range.Address),
            After = Json.GetObj(op, "rule")?.DeepClone(),
        });
    }

    private static void ApplyAddConditional(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var range = BindRange(workbook, op, "range");
        EnsureSheetWritable(range.Sheet);
        object? comRange = null;
        object? conditions = null;
        object? added = null;
        try
        {
            comRange = (object)((dynamic)range.Sheet).Range(range.Address);
            conditions = (object)((dynamic)comRange).FormatConditions;
            var before = Convert.ToInt32(((dynamic)conditions).Count, CultureInfo.InvariantCulture);
            added = AddConditionalRule(conditions, Json.GetObj(op, "rule")!);
            ApplyConditionalStyle(added, Json.GetObj(op, "style"));
            var after = Convert.ToInt32(((dynamic)conditions).Count, CultureInfo.InvariantCulture);
            checkedCells++;
            if (after <= before)
                mismatches.Add($"{range.SheetName}!{range.Address}: FormatConditions.Count did not increase");
            execution.Affected.Add(new AffectedRef("conditionalFormat", $"{range.SheetName}!{range.Address}"));
        }
        finally
        {
            RotHelper.ReleaseComReference(added);
            RotHelper.ReleaseComReference(conditions);
            RotHelper.ReleaseComReference(comRange);
        }
    }

    private static void PreviewClearConditional(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var range = BindRange(workbook, op, "range");
        preview.Affected.Add(new AffectedRef("conditionalFormat", $"{range.SheetName}!{range.Address}"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{range.SheetName}!{range.Address}:formatConditions",
            Before = ReadConditionalFormats(range.Sheet, range.Address),
            After = JsonValue.Create("cleared"),
        });
    }

    private static void ApplyClearConditional(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var range = BindRange(workbook, op, "range");
        object? comRange = null;
        object? conditions = null;
        try
        {
            comRange = (object)((dynamic)range.Sheet).Range(range.Address);
            conditions = (object)((dynamic)comRange).FormatConditions;
            ((dynamic)conditions).Delete();
            checkedCells++;
            var remaining = Convert.ToInt32(((dynamic)conditions).Count, CultureInfo.InvariantCulture);
            if (remaining != 0)
                mismatches.Add($"{range.SheetName}!{range.Address}: FormatConditions.Count={remaining} after Delete");
            execution.Affected.Add(new AffectedRef("conditionalFormat", $"{range.SheetName}!{range.Address}"));
        }
        finally
        {
            RotHelper.ReleaseComReference(conditions);
            RotHelper.ReleaseComReference(comRange);
        }
    }

    private static void PreviewCreateChart(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var dest = BindSheet(workbook, op);
        using var source = BindRange(workbook, op, "sourceRange", independentSource: true);
        EnsureSheetWritable(dest.Sheet);
        var requested = Json.GetString(op, "name");
        if (!string.IsNullOrWhiteSpace(requested) && FindChartObject(dest.Sheet, requested) is { } conflict)
        {
            RotHelper.ReleaseComReference(conflict);
            preview.Errors.Add($"[EXCEL_CHART_EXISTS] ChartObject '{requested}' already exists on '{dest.SheetName}'");
            return;
        }
        preview.Affected.Add(new AffectedRef("chart", $"{dest.SheetName}!{requested ?? "new"}"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{dest.SheetName}:chart",
            Before = null,
            After = new JsonObject
            {
                ["name"] = requested,
                ["chartType"] = Json.GetString(op, "chartType"),
                ["sourceRange"] = source.CacheSource,
                ["title"] = Json.GetString(op, "title"),
                ["series"] = op["series"]?.DeepClone(),
            },
        });
    }

    private static void ApplyCreateChart(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var dest = BindSheet(workbook, op);
        using var source = BindRange(workbook, op, "sourceRange", independentSource: true);
        EnsureSheetWritable(dest.Sheet);
        object? charts = null;
        object? chartObject = null;
        object? chart = null;
        try
        {
            var position = ReadPosition(op, DefaultChartLeft, DefaultChartTop, DefaultChartWidth, DefaultChartHeight);
            charts = (object)((dynamic)dest.Sheet).ChartObjects();
            chartObject = (object)((dynamic)charts).Add(position.Left, position.Top, position.Width, position.Height);
            chart = (object)((dynamic)chartObject).Chart;
            TrySetChartSourceData(chart, source, op);
            if (!ExcelDataObjectCatalog.TryChartType(Json.GetString(op, "chartType"), out var chartType))
                throw new InvalidOperationException("[EXCEL_DATA_UNSUPPORTED] chartType is not mapped to XlChartType");
            ((dynamic)chart).ChartType = chartType;
            ApplyAdvertisedChartSeries(workbook, chart, Json.GetArr(op, "series"), source.SheetName);
            ExcelChartDetailsApply.Apply(chart, op);
            ApplyChartTitleLegend(chart, op);
            ApplyChartPresentation(chart, op);
            var requested = Json.GetString(op, "name");
            if (!string.IsNullOrWhiteSpace(requested))
                ((dynamic)chartObject).Name = requested;
            var actual = ReadChartState(chartObject, dest.SheetName);
            checkedCells += VerifyChart(actual, op, source, mismatches);
            execution.Affected.Add(new AffectedRef("chart", $"{dest.SheetName}!{Json.GetString(actual, "name")}"));
        }
        finally
        {
            RotHelper.ReleaseComReference(chart);
            RotHelper.ReleaseComReference(chartObject);
            RotHelper.ReleaseComReference(charts);
        }
    }

    private static void PreviewUpdateChart(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var chart = BindChart(workbook, op);
        EnsureSheetWritable(chart.Sheet);
        preview.Affected.Add(new AffectedRef("chart", $"{chart.SheetName}!{chart.Name}"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{chart.SheetName}!{chart.Name}",
            Before = chart.State.DeepClone(),
            After = op.DeepClone(),
        });
    }

    private static void ApplyUpdateChart(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var chart = BindChart(workbook, op);
        EnsureSheetWritable(chart.Sheet);
        object? inner = null;
        try
        {
            inner = (object)((dynamic)chart.ChartObject).Chart;
            DataRangeLease? source = null;
            try
            {
                if (op.ContainsKey("sourceRange"))
                {
                    source = BindRange(workbook, op, "sourceRange", independentSource: true);
                    TrySetChartSourceData(inner, source, op);
                }
                if (op.ContainsKey("chartType"))
                {
                    if (!ExcelDataObjectCatalog.TryChartType(Json.GetString(op, "chartType"), out var chartType))
                        throw new InvalidOperationException("[EXCEL_DATA_UNSUPPORTED] chartType is not mapped");
                    ((dynamic)inner).ChartType = chartType;
                }
                ApplyAdvertisedChartSeries(
                    workbook,
                    inner,
                    Json.GetArr(op, "series"),
                    source?.SheetName ?? chart.SheetName);
                ExcelChartDetailsApply.Apply(inner, op);
                ApplyChartTitleLegend(inner, op);
                ApplyChartPresentation(inner, op);
                if (op.ContainsKey("position"))
                    ApplyPosition(chart.ChartObject, ReadPosition(op, 0, 0, 0, 0, optionalSize: true));
                var actual = ReadChartState(chart.ChartObject, chart.SheetName);
                checkedCells += VerifyChart(actual, op, source, mismatches);
            }
            finally { source?.Dispose(); }
            execution.Affected.Add(new AffectedRef("chart", $"{chart.SheetName}!{chart.Name}"));
        }
        finally { RotHelper.ReleaseComReference(inner); }
    }

    private static void PreviewInsertPicture(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var sheet = BindSheet(workbook, op);
        EnsureSheetWritable(sheet.Sheet);
        var path = Path.GetFullPath(Json.GetString(op, "path")!);
        if (!File.Exists(path))
        {
            preview.Errors.Add($"[EXCEL_PICTURE_NOT_FOUND] image file does not exist: {path}");
            return;
        }
        var requested = Json.GetString(op, "name");
        if (!string.IsNullOrWhiteSpace(requested) && FindShape(sheet.Sheet, requested) is { } conflict)
        {
            RotHelper.ReleaseComReference(conflict);
            preview.Errors.Add($"[EXCEL_PICTURE_EXISTS] Shape '{requested}' already exists on '{sheet.SheetName}'");
            return;
        }
        preview.Affected.Add(new AffectedRef("picture", $"{sheet.SheetName}!{requested ?? Path.GetFileName(path)}"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{sheet.SheetName}:picture",
            Before = null,
            After = new JsonObject { ["path"] = path, ["name"] = requested },
        });
    }

    private static void ApplyInsertPicture(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var sheet = BindSheet(workbook, op);
        EnsureSheetWritable(sheet.Sheet);
        var path = Path.GetFullPath(Json.GetString(op, "path")!);
        if (!File.Exists(path))
            throw new InvalidOperationException($"[EXCEL_PICTURE_NOT_FOUND] image file does not exist: {path}");
        object? shapes = null;
        object? shape = null;
        try
        {
            var position = ReadPosition(op, DefaultPictureLeft, DefaultPictureTop, DefaultPictureWidth, DefaultPictureHeight);
            shapes = (object)((dynamic)sheet.Sheet).Shapes;
            shape = (object)((dynamic)shapes).AddPicture(
                path, ExcelDataObjectCatalog.MsoFalse, ExcelDataObjectCatalog.MsoTrue,
                position.Left, position.Top, position.Width, position.Height);
            var requested = Json.GetString(op, "name");
            if (!string.IsNullOrWhiteSpace(requested))
                ((dynamic)shape).Name = requested;
            if (op.ContainsKey("lockAspectRatio"))
                ((dynamic)shape).LockAspectRatio = Json.GetBool(op, "lockAspectRatio")
                    ? ExcelDataObjectCatalog.MsoTrue
                    : ExcelDataObjectCatalog.MsoFalse;
            var actual = ReadPictureState(shape, sheet.SheetName);
            checkedCells += VerifyPicture(actual, op, sheet.SheetName, mismatches);
            execution.Affected.Add(new AffectedRef("picture", $"{sheet.SheetName}!{Json.GetString(actual, "name")}"));
        }
        finally
        {
            RotHelper.ReleaseComReference(shape);
            RotHelper.ReleaseComReference(shapes);
        }
    }

    private static void PreviewUpdatePicture(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var picture = BindPicture(workbook, op);
        EnsureSheetWritable(picture.Sheet);
        preview.Affected.Add(new AffectedRef("picture", $"{picture.SheetName}!{picture.Name}"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{picture.SheetName}!{picture.Name}",
            Before = picture.State.DeepClone(),
            After = op.DeepClone(),
        });
    }

    private static void ApplyUpdatePicture(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var picture = BindPicture(workbook, op);
        EnsureSheetWritable(picture.Sheet);
        if (op.ContainsKey("path"))
        {
            var replaced = ApplyPictureReplacement(picture, op);
            checkedCells += VerifyPicture(replaced, op, picture.SheetName, mismatches);
            execution.Affected.Add(new AffectedRef("picture", $"{picture.SheetName}!{picture.Name}"));
            return;
        }
        if (op.ContainsKey("position"))
            ApplyPosition(picture.Shape, ReadPosition(op, 0, 0, 0, 0, optionalSize: true));
        if (op.ContainsKey("lockAspectRatio"))
            ((dynamic)picture.Shape).LockAspectRatio = Json.GetBool(op, "lockAspectRatio")
                ? ExcelDataObjectCatalog.MsoTrue
                : ExcelDataObjectCatalog.MsoFalse;
        if (Json.GetObj(op, "crop") is JsonObject crop)
        {
            object? format = null;
            try
            {
                format = (object)((dynamic)picture.Shape).PictureFormat;
                if (crop.ContainsKey("left")) ((dynamic)format).CropLeft = ExcelShapeFormatContract.ReadFiniteNumber(crop["left"]!);
                if (crop.ContainsKey("top")) ((dynamic)format).CropTop = ExcelShapeFormatContract.ReadFiniteNumber(crop["top"]!);
                if (crop.ContainsKey("right")) ((dynamic)format).CropRight = ExcelShapeFormatContract.ReadFiniteNumber(crop["right"]!);
                if (crop.ContainsKey("bottom")) ((dynamic)format).CropBottom = ExcelShapeFormatContract.ReadFiniteNumber(crop["bottom"]!);
            }
            finally { RotHelper.ReleaseComReference(format); }
        }
        var actual = ReadPictureState(picture.Shape, picture.SheetName);
        checkedCells += VerifyPicture(actual, op, picture.SheetName, mismatches);
        execution.Affected.Add(new AffectedRef("picture", $"{picture.SheetName}!{picture.Name}"));
    }

    private static void PreviewSetNote(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var range = BindCell(workbook, op);
        EnsureSheetWritable(range.Sheet);
        preview.Affected.Add(new AffectedRef("note", $"{range.SheetName}!{range.Address}"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{range.SheetName}!{range.Address}:note",
            Before = ReadNoteState(range.Sheet, range.Address),
            After = JsonValue.Create(Json.GetString(op, "text")),
        });
    }

    private static void ApplySetNote(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var range = BindCell(workbook, op);
        EnsureSheetWritable(range.Sheet);
        object? comRange = null;
        object? comment = null;
        try
        {
            comRange = (object)((dynamic)range.Sheet).Range(range.Address);
            try { comment = (object?)((dynamic)comRange).Comment; }
            catch { comment = null; }
            var text = Json.GetString(op, "text")!;
            if (comment is null)
                comment = (object)((dynamic)comRange).AddComment(text);
            else
                ((dynamic)comment).Text(text);
            if (op.ContainsKey("visible"))
                ((dynamic)comment).Visible = Json.GetBool(op, "visible");
            var actual = ReadNoteState(range.Sheet, range.Address);
            checkedCells++;
            if (!string.Equals(Json.GetString(actual, "text"), text, StringComparison.Ordinal))
                mismatches.Add($"{range.SheetName}!{range.Address}: note text readback mismatch");
            execution.Affected.Add(new AffectedRef("note", $"{range.SheetName}!{range.Address}"));
        }
        finally
        {
            RotHelper.ReleaseComReference(comment);
            RotHelper.ReleaseComReference(comRange);
        }
    }

    private static void PreviewClearNote(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var range = BindRange(workbook, op, Json.GetString(op, "range") is null ? "cell" : "range");
        preview.Affected.Add(new AffectedRef("note", $"{range.SheetName}!{range.Address}"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{range.SheetName}!{range.Address}:note",
            Before = ReadNoteState(range.Sheet, range.Address),
            After = null,
        });
    }

    private static void ApplyClearNote(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var range = BindRange(workbook, op, Json.GetString(op, "range") is null ? "cell" : "range");
        object? comRange = null;
        try
        {
            comRange = (object)((dynamic)range.Sheet).Range(range.Address);
            ((dynamic)comRange).ClearComments();
            checkedCells++;
            var actual = ReadNoteState(range.Sheet, range.Address);
            if (Json.GetBool(actual, "present"))
                mismatches.Add($"{range.SheetName}!{range.Address}: note still present after ClearComments");
            execution.Affected.Add(new AffectedRef("note", $"{range.SheetName}!{range.Address}"));
        }
        finally { RotHelper.ReleaseComReference(comRange); }
    }

    private static void PreviewSetHyperlink(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var range = BindCell(workbook, op);
        EnsureSheetWritable(range.Sheet);
        preview.Affected.Add(new AffectedRef("hyperlink", $"{range.SheetName}!{range.Address}"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{range.SheetName}!{range.Address}:hyperlink",
            Before = ReadHyperlinkState(range.Sheet, range.Address),
            After = new JsonObject
            {
                ["address"] = Json.GetString(op, "address"),
                ["subAddress"] = Json.GetString(op, "subAddress"),
            },
        });
    }

    private static void ApplySetHyperlink(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var range = BindCell(workbook, op);
        EnsureSheetWritable(range.Sheet);
        object? comRange = null;
        object? links = null;
        object? sheetLinks = null;
        try
        {
            comRange = (object)((dynamic)range.Sheet).Range(range.Address);
            links = (object)((dynamic)comRange).Hyperlinks;
            try { ((dynamic)links).Delete(); }
            catch { /* none */ }
            sheetLinks = (object)((dynamic)range.Sheet).Hyperlinks;
            ((dynamic)sheetLinks).Add(
                comRange,
                Json.GetString(op, "address") ?? "",
                Json.GetString(op, "subAddress") ?? Type.Missing,
                Json.GetString(op, "screenTip") ?? Type.Missing,
                Json.GetString(op, "textToDisplay") ?? Type.Missing);
            var actual = ReadHyperlinkState(range.Sheet, range.Address);
            checkedCells++;
            if (!Json.GetBool(actual, "present"))
                mismatches.Add($"{range.SheetName}!{range.Address}: hyperlink missing after Add");
            else if (!HyperlinkMatches(actual, op))
                mismatches.Add($"{range.SheetName}!{range.Address}: hyperlink address/subAddress readback mismatch");
            execution.Affected.Add(new AffectedRef("hyperlink", $"{range.SheetName}!{range.Address}"));
        }
        finally
        {
            RotHelper.ReleaseComReference(sheetLinks);
            RotHelper.ReleaseComReference(links);
            RotHelper.ReleaseComReference(comRange);
        }
    }

    private static void PreviewClearHyperlink(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var range = BindRange(workbook, op, Json.GetString(op, "range") is null ? "cell" : "range");
        preview.Affected.Add(new AffectedRef("hyperlink", $"{range.SheetName}!{range.Address}"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{range.SheetName}!{range.Address}:hyperlink",
            Before = ReadHyperlinkState(range.Sheet, range.Address),
            After = null,
        });
    }

    private static void ApplyClearHyperlink(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var range = BindRange(workbook, op, Json.GetString(op, "range") is null ? "cell" : "range");
        object? comRange = null;
        object? links = null;
        try
        {
            comRange = (object)((dynamic)range.Sheet).Range(range.Address);
            links = (object)((dynamic)comRange).Hyperlinks;
            ((dynamic)links).Delete();
            checkedCells++;
            var actual = ReadHyperlinkState(range.Sheet, range.Address);
            if (Json.GetBool(actual, "present"))
                mismatches.Add($"{range.SheetName}!{range.Address}: hyperlink still present after Delete");
            execution.Affected.Add(new AffectedRef("hyperlink", $"{range.SheetName}!{range.Address}"));
        }
        finally
        {
            RotHelper.ReleaseComReference(links);
            RotHelper.ReleaseComReference(comRange);
        }
    }
}
