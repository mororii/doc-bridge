using System.Globalization;
using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

/// <summary>
/// Operation-scoped snapshot/restore for Excel data objects.
/// Integration Cursor calls these from CaptureSnapshot / RestoreSnapshot
/// when <see cref="IsDataOnlySnapshot"/> / <see cref="IsDataObjectsRestoreMode"/>.
/// verified=true requires <see cref="ExcelDataSnapshotEquality"/> comparisons.
/// </summary>
public sealed partial class ExcelAdapter
{
    [ThreadStatic] private static JsonObject? DataSnapshotBackupMetadata;

    internal static JsonObject CaptureDataOnlyState(object workbook, IReadOnlyList<JsonObject> ops,
        string? documentRef) =>
        CaptureDataOnlyState(workbook, ops, documentRef, backupMetadata: null);

    /// <param name="backupMetadata">
    /// Integration snapshot metadata after <c>CaptureAuxiliaryWorkbookBackup</c>.
    /// Fresh native-copy recovery is recorded only when that copy was actually taken.
    /// </param>
    internal static JsonObject CaptureDataOnlyState(object workbook, IReadOnlyList<JsonObject> ops,
        string? documentRef, JsonObject? backupMetadata)
    {
        DataSnapshotBackupMetadata = backupMetadata;
        try
        {
            var entries = new JsonArray();
            foreach (var op in ops)
            {
                var name = Json.GetString(op, "op") ?? "";
                entries.Add(CaptureDataEntry(workbook, op, name));
            }
            var recovery = ExcelDataRecoveryContract.SurfaceRecoveryFields(backupMetadata);
            return new JsonObject
            {
                ["snapshotVersion"] = DataObjectsSnapshotVersion,
                ["restoreMode"] = DataObjectsRestoreMode,
                ["documentRef"] = documentRef,
                ["originalActiveSheet"] = ReadActiveSheetName(workbook),
                ["equalityRequired"] = true,
                ["tableCreateRollback"] = ExcelDataSnapshotEquality.TableCreateRollbackMethod,
                ["recoveryAvailable"] = recovery["recoveryAvailable"]?.DeepClone(),
                ["recoveryArtifact"] = recovery["recoveryArtifact"]?.DeepClone(),
                ["recoverySource"] = recovery["recoverySource"]?.DeepClone(),
                ["ops"] = CloneOps(ops),
                ["entries"] = entries,
            };
        }
        finally { DataSnapshotBackupMetadata = null; }
    }

    internal static JsonObject RestoreDataOnlyState(object workbook, JsonObject state)
    {
        var mismatches = new RestoreMismatchCollector();
        var version = Json.GetInt(state, "snapshotVersion");
        if (!IsDataObjectsRestoreMode(Json.GetString(state, "restoreMode")) || version is null || version < 1)
        {
            mismatches.Add("data-objects snapshot version or restoreMode is invalid");
            return BuildRestoreResult(false, 0, 0, DataObjectsRestoreMode, mismatches);
        }

        var entries = Json.GetArr(state, "entries") ?? new JsonArray();
        var actuals = new JsonObject?[entries.Count];
        var checkedItems = 0;
        var restoredCells = 0;
        for (var index = entries.Count - 1; index >= 0; index--)
        {
            if (entries[index] is not JsonObject entry) continue;
            try
            {
                restoredCells += RestoreDataEntry(workbook, entry, mismatches);
                actuals[index] = ReadRestoredActual(workbook, entry);
                checkedItems++;
            }
            catch (Exception ex)
            {
                mismatches.Add($"{Json.GetString(entry, "kind")}: restore failed: {ex.Message}");
                actuals[index] = new JsonObject { ["readFailed"] = true, ["error"] = ex.Message };
            }
        }

        var originalActive = Json.GetString(state, "originalActiveSheet");
        if (!string.IsNullOrWhiteSpace(originalActive))
        {
            try { ActivateWorksheet(workbook, originalActive, mismatches); }
            catch (Exception ex) { mismatches.Add($"active sheet restore failed: {ex.Message}"); }
        }

        var live = new JsonArray();
        foreach (var actual in actuals)
            live.Add(actual?.DeepClone() ?? new JsonObject { ["readFailed"] = true, ["error"] = "missing actual" });
        var equality = ExcelDataSnapshotEquality.Evaluate(state, live);
        foreach (var node in Json.GetArr(equality, "mismatches")?.OfType<JsonValue>() ?? Enumerable.Empty<JsonValue>())
        {
            if (node.TryGetValue<string>(out var text) && !string.IsNullOrWhiteSpace(text))
                mismatches.Add(text);
        }

        return BuildRestoreResult(
            mismatches.Count == 0 && Json.GetBool(equality, "verified"),
            restoredCells, checkedItems, DataObjectsRestoreMode, mismatches);
    }

    private static JsonObject CaptureDataEntry(object workbook, JsonObject op, string name) => name switch
    {
        ExcelDataOperationsContract.CreateTable => CaptureCreateTableEntry(workbook, op),
        ExcelDataOperationsContract.ResizeTable or ExcelDataOperationsContract.StyleTable
            or ExcelDataOperationsContract.SetTableTotals or ExcelDataOperationsContract.SortTable
            or ExcelDataOperationsContract.AddTableColumn or ExcelDataOperationsContract.DeleteTable
            or ExcelDataOperationsContract.AppendTableRows or ExcelDataOperationsContract.InsertTableRows
            or ExcelDataOperationsContract.DeleteTableRows =>
            CaptureExistingTableEntry(workbook, op, includeSurface: name is
                ExcelDataOperationsContract.SortTable or ExcelDataOperationsContract.DeleteTable
                or ExcelDataOperationsContract.AddTableColumn or ExcelDataOperationsContract.AppendTableRows
                or ExcelDataOperationsContract.InsertTableRows or ExcelDataOperationsContract.DeleteTableRows),
        ExcelDataOperationsContract.SortRange or ExcelDataOperationsContract.RemoveDuplicates =>
            CaptureSortRangeEntry(workbook, op, name),
        ExcelDataOperationsContract.TextToColumns => CaptureTextToColumnsEntry(workbook, op),
        ExcelDataOperationsContract.SetAutoFilter or ExcelDataOperationsContract.ClearAutoFilter =>
            CaptureFilterEntry(workbook, op),
        ExcelDataOperationsContract.DefineName or ExcelDataOperationsContract.UpdateName
            or ExcelDataOperationsContract.DeleteName => CaptureNameEntry(workbook, op),
        ExcelDataOperationsContract.SetDataValidation or ExcelDataOperationsContract.ClearDataValidation =>
            CaptureValidationEntry(workbook, op),
        ExcelDataOperationsContract.AddConditionalFormat or ExcelDataOperationsContract.ClearConditionalFormats
            or ExcelDataOperationsContract.UpdateConditionalFormat or ExcelDataOperationsContract.DeleteConditionalFormat =>
            CaptureConditionalEntry(workbook, op),
        ExcelDataOperationsContract.SetRichText => CaptureRichTextEntry(workbook, op),
        ExcelDataOperationsContract.CreateChart => CaptureCreateChartEntry(workbook, op),
        ExcelDataOperationsContract.UpdateChart or ExcelDataOperationsContract.DeleteChart =>
            CaptureExistingChartEntry(workbook, op),
        ExcelDataOperationsContract.InsertSheetPicture => CaptureCreatePictureEntry(workbook, op),
        ExcelDataOperationsContract.UpdatePicture or ExcelDataOperationsContract.DeletePicture =>
            CaptureExistingPictureEntry(workbook, op),
        ExcelDataOperationsContract.SetCellNote or ExcelDataOperationsContract.ClearCellNote =>
            CaptureNoteEntry(workbook, op),
        ExcelDataOperationsContract.SetHyperlink or ExcelDataOperationsContract.ClearHyperlink =>
            CaptureHyperlinkEntry(workbook, op),
        ExcelDataOperationsContract.CreatePivot => CaptureCreatePivotEntry(workbook, op),
        ExcelDataOperationsContract.UpdatePivot or ExcelDataOperationsContract.RefreshPivot
            or ExcelDataOperationsContract.DeletePivot => CaptureExistingPivotEntry(workbook, op),
        ExcelDataOperationsContract.InsertShape or ExcelDataOperationsContract.InsertTextbox
            or ExcelDataOperationsContract.CreateConnector =>
            CaptureCreateShapeEntry(workbook, op, name == ExcelDataOperationsContract.InsertTextbox ? "textbox" : name == ExcelDataOperationsContract.CreateConnector ? "connector" : "shape"),
        ExcelDataOperationsContract.UpdateShape or ExcelDataOperationsContract.DeleteShape
            or ExcelDataOperationsContract.UpdateTextbox or ExcelDataOperationsContract.DeleteTextbox
            or ExcelDataOperationsContract.UpdateConnector =>
            CaptureExistingShapeEntry(workbook, op),
        ExcelDataOperationsContract.UpdateExternalLinks or ExcelDataOperationsContract.ChangeLinkSource
            or ExcelDataOperationsContract.BreakExternalLink =>
            CaptureLinkEntry(workbook, op, name),
        ExcelDataOperationsContract.SetCalculationMode => CaptureCalculationEntry(workbook),
        ExcelDataOperationsContract.FreezeValues or ExcelDataOperationsContract.PasteSpecial =>
            CaptureValueSurfaceEntry(workbook, op, name),
        ExcelDataOperationsContract.GoalSeek => CaptureGoalSeekEntry(workbook, op),
        ExcelDataOperationsContract.ProtectWorkbook or ExcelDataOperationsContract.UnprotectWorkbook =>
            CaptureWorkbookProtectionEntry(workbook),
        ExcelDataOperationsContract.SetSplitPanes => CaptureSplitEntry(workbook, op),
        ExcelDataOperationsContract.CreateSparkline or ExcelDataOperationsContract.UpdateSparkline
            or ExcelDataOperationsContract.DeleteSparkline =>
            CaptureSparklineEntry(workbook, op, name),
        ExcelDataOperationsContract.CreateSlicer or ExcelDataOperationsContract.DeleteSlicer =>
            CaptureSlicerEntry(workbook, op, name),
        ExcelDataOperationsContract.ApplyCellStyle => CaptureCellStyleEntry(workbook, op),
        _ => throw new InvalidOperationException(
            $"[EXCEL_DATA_UNSUPPORTED] snapshot cannot capture unknown data op '{name}'"),
    };

    private static int RestoreDataEntry(object workbook, JsonObject entry, RestoreMismatchCollector mismatches)
    {
        var kind = Json.GetString(entry, "kind");
        return kind switch
        {
            "table" => RestoreTableEntry(workbook, entry, mismatches),
            "sortRange" or "duplicates" or "textToColumns" => RestoreSortRangeEntry(workbook, entry, mismatches),
            "autoFilter" => RestoreFilterEntry(workbook, entry, mismatches),
            "definedName" => RestoreNameEntry(workbook, entry, mismatches),
            "validation" => RestoreValidationEntry(workbook, entry, mismatches),
            "conditionalFormat" => RestoreConditionalEntry(workbook, entry, mismatches),
            "richText" => RestoreRichTextEntry(workbook, entry, mismatches),
            "chart" => RestoreChartEntry(workbook, entry, mismatches),
            "picture" => RestorePictureEntry(workbook, entry, mismatches),
            "note" => RestoreNoteEntry(workbook, entry, mismatches),
            "hyperlink" => RestoreHyperlinkEntry(workbook, entry, mismatches),
            "pivot" => RestorePivotEntry(workbook, entry, mismatches),
            "shape" or "textbox" or "connector" => RestoreShapeEntry(workbook, entry, mismatches),
            "externalLinks" => RestoreLinkEntry(workbook, entry, mismatches),
            "calculationMode" => RestoreCalculationEntry(workbook, entry, mismatches),
            "freezeValues" or "pasteSpecial" => RestoreSortRangeEntry(workbook, entry, mismatches),
            "goalSeek" => RestoreGoalSeekEntry(workbook, entry, mismatches),
            "workbookProtection" => RestoreWorkbookProtectionEntry(workbook, entry, mismatches),
            "splitPanes" => RestoreSplitEntry(workbook, entry, mismatches),
            "sparkline" => RestoreSparklineEntry(workbook, entry, mismatches),
            "slicer" => RestoreSlicerEntry(workbook, entry, mismatches),
            "cellStyle" => RestoreCellStyleEntry(workbook, entry, mismatches),
            _ => throw new InvalidOperationException(
                $"[EXCEL_DATA_UNSUPPORTED] snapshot kind '{kind}' cannot be restored"),
        };
    }

    private static JsonObject CaptureCreateTableEntry(object workbook, JsonObject op)
    {
        using var range = BindRange(workbook, op, "range");
        var entry = new JsonObject
        {
            ["kind"] = "table",
            ["existed"] = false,
            ["sheet"] = range.SheetName,
            ["name"] = Json.GetString(op, "name"),
            ["range"] = range.Address,
            ["overlapping"] = FindOverlappingTable(range.Sheet, range.Address),
            ["rollback"] = ExcelDataSnapshotEquality.TableCreateRollbackMethod,
        };
        MergeSurface(entry, CaptureRangeSurface(range.Sheet, range.Address));
        return entry;
    }

    private static JsonObject CaptureExistingTableEntry(object workbook, JsonObject op, bool includeSurface)
    {
        using var table = BindTable(workbook, op);
        var entry = new JsonObject
        {
            ["kind"] = "table",
            ["existed"] = true,
            ["sheet"] = table.SheetName,
            ["name"] = table.Name,
            ["state"] = table.State.DeepClone(),
        };
        var range = Json.GetString(table.State, "range");
        if (includeSurface && !string.IsNullOrWhiteSpace(range))
        {
            entry["formulaRange"] = range;
            entry["range"] = range;
            MergeSurface(entry, CaptureRangeSurface(table.Sheet, range));
        }
        return entry;
    }

    private static JsonObject CaptureSortRangeEntry(object workbook, JsonObject op, string opName)
    {
        using var range = BindRange(workbook, op, "range");
        var kind = opName == ExcelDataOperationsContract.RemoveDuplicates ? "duplicates" : "sortRange";
        var entry = new JsonObject
        {
            ["kind"] = kind,
            ["sheet"] = range.SheetName,
            ["range"] = range.Address,
        };
        MergeSurface(entry, CaptureRangeSurface(range.Sheet, range.Address));
        return entry;
    }

    private static JsonObject CaptureTextToColumnsEntry(object workbook, JsonObject op)
    {
        using var source = BindRange(workbook, op, "range");
        DataRangeLease? destination = null;
        try
        {
            if (!string.IsNullOrWhiteSpace(Json.GetString(op, "destination")))
                destination = BindRange(workbook, op, "destination");
            var destSheet = destination?.Sheet ?? source.Sheet;
            var destSheetName = destination?.SheetName ?? source.SheetName;
            var destOrigin = destination?.Address ?? source.Address;
            var captureAddress = ComputeTextToColumnsCaptureAddress(source, destOrigin, op);
            var entry = new JsonObject
            {
                ["kind"] = "textToColumns",
                ["sheet"] = destSheetName,
                ["range"] = captureAddress,
                ["sourceRange"] = source.Address,
                ["sourceSheet"] = source.SheetName,
                ["destination"] = destOrigin,
            };
            MergeSurface(entry, CaptureRangeSurface(destSheet, captureAddress));
            return entry;
        }
        finally { destination?.Dispose(); }
    }

    private static JsonObject CaptureFilterEntry(object workbook, JsonObject op)
    {
        using var target = BindFilterTarget(workbook, op);
        return new JsonObject
        {
            ["kind"] = "autoFilter",
            ["sheet"] = target.SheetName,
            ["range"] = target.Address,
            ["table"] = target.ListObject is null ? null : ReadComString(target.ListObject, "Name"),
            ["state"] = ReadAutoFilterState(target.Sheet, target.SheetName),
        };
    }

    private static JsonObject CaptureNameEntry(object workbook, JsonObject op)
    {
        var name = Json.GetString(op, "name")!;
        var scope = Json.GetString(op, "scope")!;
        var sheetName = Json.GetString(Json.GetObj(op, "target"), "sheet");
        var existing = FindDefinedName(workbook, name, scope, sheetName);
        try
        {
            return new JsonObject
            {
                ["kind"] = "definedName",
                ["existed"] = existing is not null,
                ["name"] = name,
                ["scope"] = scope,
                ["sheet"] = sheetName,
                ["state"] = existing is null ? null : ReadDefinedNameState(existing),
            };
        }
        finally { RotHelper.ReleaseComReference(existing); }
    }

    private static JsonObject CaptureValidationEntry(object workbook, JsonObject op)
    {
        using var range = BindRange(workbook, op, "range");
        return new JsonObject
        {
            ["kind"] = "validation",
            ["sheet"] = range.SheetName,
            ["range"] = range.Address,
            ["state"] = ReadValidationState(range.Sheet, range.Address),
        };
    }

    private static JsonObject CaptureConditionalEntry(object workbook, JsonObject op)
    {
        using var range = BindRange(workbook, op, "range");
        return new JsonObject
        {
            ["kind"] = "conditionalFormat",
            ["sheet"] = range.SheetName,
            ["range"] = range.Address,
            ["state"] = ReadConditionalFormats(range.Sheet, range.Address),
        };
    }

    private static JsonObject CaptureCreateChartEntry(object workbook, JsonObject op)
    {
        using var sheet = BindSheet(workbook, op);
        return new JsonObject
        {
            ["kind"] = "chart",
            ["existed"] = false,
            ["sheet"] = sheet.SheetName,
            ["name"] = Json.GetString(op, "name"),
        };
    }

    private static JsonObject CaptureExistingChartEntry(object workbook, JsonObject op)
    {
        using var chart = BindChart(workbook, op);
        return new JsonObject
        {
            ["kind"] = "chart",
            ["existed"] = true,
            ["sheet"] = chart.SheetName,
            ["name"] = chart.Name,
            ["state"] = chart.State.DeepClone(),
        };
    }

    private static JsonObject CaptureCreatePictureEntry(object workbook, JsonObject op)
    {
        using var sheet = BindSheet(workbook, op);
        return new JsonObject
        {
            ["kind"] = "picture",
            ["existed"] = false,
            ["sheet"] = sheet.SheetName,
            ["name"] = Json.GetString(op, "name"),
            ["path"] = Json.GetString(op, "path"),
        };
    }

    private static JsonObject CaptureExistingPictureEntry(object workbook, JsonObject op)
    {
        using var picture = BindPicture(workbook, op);
        return new JsonObject
        {
            ["kind"] = "picture",
            ["existed"] = true,
            ["sheet"] = picture.SheetName,
            ["name"] = picture.Name,
            ["path"] = Json.GetString(op, "path"),
            ["state"] = picture.State.DeepClone(),
        };
    }

    private static JsonObject CaptureNoteEntry(object workbook, JsonObject op)
    {
        using var range = BindCell(workbook, op);
        return new JsonObject
        {
            ["kind"] = "note",
            ["sheet"] = range.SheetName,
            ["range"] = range.Address,
            ["state"] = ReadNoteState(range.Sheet, range.Address),
        };
    }

    private static JsonObject CaptureHyperlinkEntry(object workbook, JsonObject op)
    {
        using var range = BindCell(workbook, op);
        return new JsonObject
        {
            ["kind"] = "hyperlink",
            ["sheet"] = range.SheetName,
            ["range"] = range.Address,
            ["state"] = ReadHyperlinkState(range.Sheet, range.Address),
        };
    }

    private static JsonObject CaptureCreatePivotEntry(object workbook, JsonObject op)
    {
        using var sheet = BindSheet(workbook, op);
        return new JsonObject
        {
            ["kind"] = "pivot",
            ["existed"] = false,
            ["sheet"] = sheet.SheetName,
            ["name"] = Json.GetString(op, "name"),
            ["destination"] = Json.GetString(op, "destination"),
        };
    }

    private static JsonObject CaptureExistingPivotEntry(object workbook, JsonObject op)
    {
        using var sheet = BindSheet(workbook, op);
        var name = Json.GetString(op, "name")!;
        var pivot = FindPivotTable(sheet.Sheet, name);
        try
        {
            if (pivot is null)
                throw new InvalidOperationException($"[EXCEL_PIVOT_NOT_FOUND] PivotTable '{name}' was not found");
            return new JsonObject
            {
                ["kind"] = "pivot",
                ["existed"] = true,
                ["sheet"] = sheet.SheetName,
                ["name"] = name,
                ["state"] = ReadPivotState(pivot, sheet.SheetName),
            };
        }
        finally { RotHelper.ReleaseComReference(pivot); }
    }

    private static JsonObject CaptureCreateShapeEntry(object workbook, JsonObject op, string kind)
    {
        using var sheet = BindSheet(workbook, op);
        return new JsonObject
        {
            ["kind"] = kind,
            ["existed"] = false,
            ["sheet"] = sheet.SheetName,
            ["name"] = Json.GetString(op, "name"),
        };
    }

    private static JsonObject CaptureExistingShapeEntry(object workbook, JsonObject op)
    {
        using var sheet = BindSheet(workbook, op);
        var name = Json.GetString(op, "name")!;
        var shape = FindShape(sheet.Sheet, name);
        try
        {
            if (shape is null)
                throw new InvalidOperationException($"[EXCEL_SHAPE_NOT_FOUND] Shape '{name}' was not found");
            var state = ReadShapeState(shape, sheet.SheetName);
            return new JsonObject
            {
                ["kind"] = Json.GetString(state, "type") ?? "shape",
                ["existed"] = true,
                ["sheet"] = sheet.SheetName,
                ["name"] = name,
                ["state"] = state,
            };
        }
        finally { RotHelper.ReleaseComReference(shape); }
    }

    private static int RestoreTableEntry(object workbook, JsonObject entry, RestoreMismatchCollector mismatches)
    {
        var sheetName = Json.GetString(entry, "sheet")!;
        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, SheetOp(sheetName));
            if (!Json.GetBool(entry, "existed"))
            {
                var name = Json.GetString(entry, "name");
                var table = string.IsNullOrWhiteSpace(name)
                    ? FindOverlappingTableObject(sheet, Json.GetString(entry, "range") ?? "")
                    : FindListObject(sheet, name);
                try
                {
                    if (table is null)
                        return 0;
                    ((dynamic)table).Unlist();
                    RestoreRangeSurface(sheet, entry);
                    return 1;
                }
                finally { RotHelper.ReleaseComReference(table); }
            }

            var state = Json.GetObj(entry, "state");
            var tableName = Json.GetString(state, "name") ?? Json.GetString(entry, "name");
            if (string.IsNullOrWhiteSpace(tableName))
            {
                mismatches.Add("table snapshot is missing name");
                return 0;
            }
            var tableObject = FindListObject(sheet, tableName);
            if (tableObject is null && state is not null)
            {
                tableObject = RecreateListObject(sheet, state);
                if (tableObject is null)
                {
                    mismatches.Add($"table '{tableName}' missing during restore and could not be recreated");
                    return 0;
                }
            }
            else if (tableObject is null)
            {
                mismatches.Add($"table '{tableName}' missing during restore");
                return 0;
            }
            try
            {
                var rangeAddress = Json.GetString(state, "range");
                if (!string.IsNullOrWhiteSpace(rangeAddress))
                {
                    object? dest = null;
                    try
                    {
                        dest = (object)((dynamic)sheet).Range(rangeAddress);
                        ((dynamic)tableObject).Resize(dest);
                    }
                    finally { RotHelper.ReleaseComReference(dest); }
                }
                var styleName = Json.GetString(state, "styleName");
                if (!string.IsNullOrWhiteSpace(styleName))
                    ((dynamic)tableObject).TableStyle = styleName;
                if (state!.ContainsKey("showTotals"))
                    ((dynamic)tableObject).ShowTotals = Json.GetBool(state, "showTotals");
                if (state.ContainsKey("showHeaders"))
                    ((dynamic)tableObject).ShowHeaders = Json.GetBool(state, "showHeaders");
                if (state.ContainsKey("showAutoFilter"))
                    ((dynamic)tableObject).ShowAutoFilter = Json.GetBool(state, "showAutoFilter");
                RestoreRangeSurface(sheet, entry);
                return 1;
            }
            finally { RotHelper.ReleaseComReference(tableObject); }
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static object? RecreateListObject(object sheet, JsonObject state)
    {
        var rangeAddress = Json.GetString(state, "range");
        if (string.IsNullOrWhiteSpace(rangeAddress)) return null;
        object? listObjects = null;
        object? source = null;
        try
        {
            source = (object)((dynamic)sheet).Range(rangeAddress);
            listObjects = (object)((dynamic)sheet).ListObjects;
            var table = (object)((dynamic)listObjects).Add(
                ExcelDataObjectCatalog.XlSrcRange, source, Type.Missing, ExcelDataObjectCatalog.XlYes);
            var name = Json.GetString(state, "name");
            if (!string.IsNullOrWhiteSpace(name))
                ((dynamic)table).Name = name;
            return table;
        }
        finally
        {
            RotHelper.ReleaseComReference(source);
            RotHelper.ReleaseComReference(listObjects);
        }
    }

    private static int RestoreSortRangeEntry(object workbook, JsonObject entry, RestoreMismatchCollector mismatches)
    {
        var sheetName = Json.GetString(entry, "sheet");
        var range = Json.GetString(entry, "range");
        if (!entry.ContainsKey("formulas") || !entry.ContainsKey("values") ||
            string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(range))
        {
            mismatches.Add("sort/dedup/text-to-columns snapshot is missing values+formulas");
            return 0;
        }
        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, SheetOp(sheetName));
            RestoreRangeSurface(sheet, entry);
            return Json.GetArr(entry, "formulas")?.Count ?? 1;
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static int RestoreFilterEntry(object workbook, JsonObject entry, RestoreMismatchCollector mismatches)
    {
        var sheetName = Json.GetString(entry, "sheet")!;
        object? sheet = null;
        object? comRange = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, SheetOp(sheetName));
            var state = Json.GetObj(entry, "state");
            var address = Json.GetString(entry, "range") ?? Json.GetString(state, "range");
            if (!string.IsNullOrWhiteSpace(address))
                comRange = (object)((dynamic)sheet).Range(address);

            var enabled = state is not null && Json.GetBool(state, "enabled");
            if (!enabled)
            {
                try
                {
                    if (TryComBool(sheet, "FilterMode") == true)
                        ((dynamic)sheet).ShowAllData();
                }
                catch { /* no filtered rows */ }
                if (TryComBool(sheet, "AutoFilterMode") == true &&
                    string.IsNullOrWhiteSpace(Json.GetString(entry, "table")))
                    ((dynamic)sheet).AutoFilterMode = false;
                return 1;
            }

            if (comRange is not null && TryComBool(sheet, "AutoFilterMode") != true &&
                string.IsNullOrWhiteSpace(Json.GetString(entry, "table")))
                ((dynamic)comRange).AutoFilter();

            var filters = Json.GetArr(state, "filters") ?? new JsonArray();
            foreach (var node in filters.OfType<JsonObject>())
            {
                if (!Json.GetBool(node, "on") || comRange is null) continue;
                var field = Json.GetInt(node, "field") ?? 1;
                var criteria1 = FilterCriteriaComValue(node["criteria1"]);
                var criteria2 = FilterCriteriaComValue(node["criteria2"]);
                var op = Json.GetInt(node, "operator");
                try
                {
                    if (criteria2 is not null && op is not null)
                        ((dynamic)comRange).AutoFilter(field, criteria1 ?? Type.Missing, op.Value, criteria2);
                    else
                        ((dynamic)comRange).AutoFilter(field, criteria1 ?? Type.Missing);
                }
                catch (Exception ex)
                {
                    mismatches.Add($"{sheetName}: filter field {field} restore failed: {ex.Message}");
                }
            }
            return 1;
        }
        finally
        {
            RotHelper.ReleaseComReference(comRange);
            RotHelper.ReleaseComReference(sheet);
        }
    }

    private static object? FilterCriteriaComValue(JsonNode? node)
    {
        if (node is null) return null;
        if (node is JsonArray array)
            return array.Select(item => item?.ToString() ?? "").ToArray();
        return node.ToString();
    }

    private static int RestoreNameEntry(object workbook, JsonObject entry, RestoreMismatchCollector mismatches)
    {
        var name = Json.GetString(entry, "name")!;
        var scope = Json.GetString(entry, "scope")!;
        var sheetName = Json.GetString(entry, "sheet");
        var existed = Json.GetBool(entry, "existed");
        var current = FindDefinedName(workbook, name, scope, sheetName);
        object? created = null;
        try
        {
            if (!existed)
            {
                if (current is not null) ((dynamic)current).Delete();
                return 1;
            }
            var state = Json.GetObj(entry, "state");
            var refersTo = Json.GetString(state, "refersTo");
            if (string.IsNullOrWhiteSpace(refersTo))
            {
                mismatches.Add($"defined name '{name}' snapshot is missing RefersTo");
                return 0;
            }
            if (current is null)
            {
                created = RecreateDefinedName(workbook, name, scope, sheetName, refersTo, Json.GetString(state, "comment"));
                return 1;
            }
            ((dynamic)current).RefersTo = refersTo;
            var comment = Json.GetString(state, "comment");
            if (comment is not null) ((dynamic)current).Comment = comment;
            return 1;
        }
        finally
        {
            RotHelper.ReleaseComReference(created);
            RotHelper.ReleaseComReference(current);
        }
    }

    private static int RestoreValidationEntry(object workbook, JsonObject entry, RestoreMismatchCollector mismatches)
    {
        var sheetName = Json.GetString(entry, "sheet")!;
        var address = Json.GetString(entry, "range")!;
        object? sheet = null;
        object? range = null;
        object? validation = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, SheetOp(sheetName));
            range = (object)((dynamic)sheet).Range(address);
            validation = (object)((dynamic)range).Validation;
            try { ((dynamic)validation).Delete(); }
            catch { /* already empty */ }
            var state = Json.GetObj(entry, "state");
            if (state is not null && Json.GetBool(state, "present"))
            {
                if (Json.GetInt(state, "validationType") is null && Json.GetInt(state, "type") is null)
                {
                    mismatches.Add($"{sheetName}!{address}: validation snapshot missing type");
                    return 0;
                }
                ApplyValidationState(validation, state);
            }
            return 1;
        }
        finally
        {
            RotHelper.ReleaseComReference(validation);
            RotHelper.ReleaseComReference(range);
            RotHelper.ReleaseComReference(sheet);
        }
    }

    private static int RestoreConditionalEntry(object workbook, JsonObject entry, RestoreMismatchCollector mismatches)
    {
        var sheetName = Json.GetString(entry, "sheet")!;
        var address = Json.GetString(entry, "range")!;
        object? sheet = null;
        object? range = null;
        object? conditions = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, SheetOp(sheetName));
            range = (object)((dynamic)sheet).Range(address);
            conditions = (object)((dynamic)range).FormatConditions;
            ((dynamic)conditions).Delete();
            var rules = Json.GetArr(Json.GetObj(entry, "state"), "rules") ?? new JsonArray();
            foreach (var node in rules.OfType<JsonObject>())
            {
                if (Json.GetBool(node, "unreadable"))
                {
                    mismatches.Add($"{sheetName}!{address}: a prior FormatCondition could not be snapshotted and was not restored");
                    continue;
                }
                object? added = null;
                try
                {
                    added = RestoreOneConditionalRule(conditions, node);
                    ApplyConditionalStyle(added, Json.GetObj(node, "style"));
                }
                catch (Exception ex)
                {
                    mismatches.Add($"{sheetName}!{address}: FormatCondition restore failed: {ex.Message}");
                }
                finally { RotHelper.ReleaseComReference(added); }
            }
            return 1;
        }
        finally
        {
            RotHelper.ReleaseComReference(conditions);
            RotHelper.ReleaseComReference(range);
            RotHelper.ReleaseComReference(sheet);
        }
    }

    private static object RestoreOneConditionalRule(object conditions, JsonObject node)
    {
        var token = Json.GetString(node, "ruleType") ?? MapConditionType(Json.GetInt(node, "conditionType"));
        if (!string.IsNullOrWhiteSpace(token) && ExcelDataObjectCatalog.ConditionalRuleTypes.Contains(token))
        {
            var rule = new JsonObject
            {
                ["type"] = token,
                ["operator"] = MapOperatorToken(Json.GetInt(node, "operator")),
                ["formula1"] = Json.GetString(node, "formula1"),
                ["formula2"] = Json.GetString(node, "formula2"),
                ["iconSet"] = Json.GetString(node, "iconSet"),
                ["points"] = Json.GetInt(node, "points"),
            };
            return AddConditionalRule(conditions, rule);
        }
        var type = Json.GetInt(node, "conditionType");
        if (type is null)
            throw new InvalidOperationException("conditional rule type is missing");
        return (object)((dynamic)conditions).Add(
            type.Value,
            Json.GetInt(node, "operator") ?? Type.Missing,
            Json.GetString(node, "formula1") ?? Type.Missing);
    }

    private static string? MapConditionType(int? type) => type switch
    {
        ExcelDataObjectCatalog.XlCellValue => "cellValue",
        ExcelDataObjectCatalog.XlExpression => "expression",
        ExcelDataObjectCatalog.XlColorScale => "colorScale",
        ExcelDataObjectCatalog.XlDatabar => "dataBar",
        ExcelDataObjectCatalog.XlUniqueValues => "uniqueValues",
        ExcelDataObjectCatalog.XlIconSet => "iconSet",
        _ => null,
    };

    private static string? MapOperatorToken(int? value)
    {
        if (value is null) return null;
        return ExcelDataObjectCatalog.TokenFor(ExcelDataObjectCatalog.ComparisonOperators, value.Value, "equal");
    }

    private static int RestoreChartEntry(object workbook, JsonObject entry, RestoreMismatchCollector mismatches)
    {
        var sheetName = Json.GetString(entry, "sheet")!;
        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, SheetOp(sheetName));
            var name = Json.GetString(entry, "name");
            var chart = string.IsNullOrWhiteSpace(name) ? null : FindChartObject(sheet, name);
            try
            {
                if (!Json.GetBool(entry, "existed"))
                {
                    if (chart is not null) ((dynamic)chart).Delete();
                    return 1;
                }
                var state = Json.GetObj(entry, "state") ?? new JsonObject();
                if (chart is null)
                {
                    chart = RecreateChartObject(sheet, name, state);
                    if (chart is null)
                    {
                        mismatches.Add($"chart '{name}' missing during restore and could not be recreated");
                        return 0;
                    }
                }
                ApplyPosition(chart, new DataPosition(
                    JsonDouble(state, "left"), JsonDouble(state, "top"),
                    JsonDouble(state, "width"), JsonDouble(state, "height")));
                object? inner = null;
                try
                {
                    inner = (object)((dynamic)chart).Chart;
                    if (Json.GetInt(state, "chartTypeValue") is int chartType)
                        ((dynamic)inner).ChartType = chartType;
                    if (state.ContainsKey("hasLegend"))
                        ((dynamic)inner).HasLegend = Json.GetBool(state, "hasLegend");
                    if (state.ContainsKey("title"))
                    {
                        var title = Json.GetString(state, "title");
                        ((dynamic)inner).HasTitle = !string.IsNullOrEmpty(title);
                        if (!string.IsNullOrEmpty(title))
                        {
                            object? chartTitle = null;
                            try
                            {
                                chartTitle = (object)((dynamic)inner).ChartTitle;
                                ((dynamic)chartTitle).Text = title;
                            }
                            finally { RotHelper.ReleaseComReference(chartTitle); }
                        }
                    }
                    RestoreChartSeries(workbook, inner, state);
                }
                finally { RotHelper.ReleaseComReference(inner); }
                return 1;
            }
            finally { RotHelper.ReleaseComReference(chart); }
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static object? RecreateChartObject(object sheet, string? name, JsonObject state)
    {
        object? charts = null;
        try
        {
            charts = (object)((dynamic)sheet).ChartObjects();
            var left = JsonDouble(state, "left");
            var top = JsonDouble(state, "top");
            var width = JsonDouble(state, "width");
            var height = JsonDouble(state, "height");
            var chart = (object)((dynamic)charts).Add(
                left > 0 ? left : DefaultChartLeft,
                top > 0 ? top : DefaultChartTop,
                width > 0 ? width : DefaultChartWidth,
                height > 0 ? height : DefaultChartHeight);
            if (!string.IsNullOrWhiteSpace(name))
                ((dynamic)chart).Name = name;
            return chart;
        }
        finally { RotHelper.ReleaseComReference(charts); }
    }

    private static int RestorePictureEntry(object workbook, JsonObject entry, RestoreMismatchCollector mismatches)
    {
        var sheetName = Json.GetString(entry, "sheet")!;
        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, SheetOp(sheetName));
            var name = Json.GetString(entry, "name");
            var shape = string.IsNullOrWhiteSpace(name) ? null : FindShape(sheet, name);
            try
            {
                if (!Json.GetBool(entry, "existed"))
                {
                    if (shape is not null) ((dynamic)shape).Delete();
                    return 1;
                }
                if (shape is null)
                {
                    var path = Json.GetString(entry, "path") ?? Json.GetString(Json.GetObj(entry, "state"), "path");
                    if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                    {
                        mismatches.Add($"picture '{name}' missing during restore; pixels are not in COM state and no fresh in-memory backup was proven");
                        return 0;
                    }
                    var state = Json.GetObj(entry, "state") ?? new JsonObject();
                    object? shapes = null;
                    try
                    {
                        shapes = (object)((dynamic)sheet).Shapes;
                        shape = (object)((dynamic)shapes).AddPicture(
                            path, ExcelDataObjectCatalog.MsoFalse, ExcelDataObjectCatalog.MsoTrue,
                            JsonDouble(state, "left"), JsonDouble(state, "top"),
                            JsonDouble(state, "width") > 0 ? JsonDouble(state, "width") : DefaultPictureWidth,
                            JsonDouble(state, "height") > 0 ? JsonDouble(state, "height") : DefaultPictureHeight);
                        if (!string.IsNullOrWhiteSpace(name))
                            ((dynamic)shape).Name = name;
                    }
                    finally { RotHelper.ReleaseComReference(shapes); }
                }
                var pictureState = Json.GetObj(entry, "state") ?? new JsonObject();
                ApplyPosition(shape!, new DataPosition(
                    JsonDouble(pictureState, "left"), JsonDouble(pictureState, "top"),
                    JsonDouble(pictureState, "width"), JsonDouble(pictureState, "height")));
                return 1;
            }
            finally { RotHelper.ReleaseComReference(shape); }
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static int RestoreNoteEntry(object workbook, JsonObject entry, RestoreMismatchCollector mismatches)
    {
        var sheetName = Json.GetString(entry, "sheet")!;
        var address = Json.GetString(entry, "range")!;
        object? sheet = null;
        object? range = null;
        object? comment = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, SheetOp(sheetName));
            range = (object)((dynamic)sheet).Range(address);
            ((dynamic)range).ClearComments();
            var state = Json.GetObj(entry, "state");
            if (state is not null && Json.GetBool(state, "present"))
            {
                var text = Json.GetString(state, "text") ?? "";
                comment = (object)((dynamic)range).AddComment(text);
                if (state.ContainsKey("visible"))
                    ((dynamic)comment).Visible = Json.GetBool(state, "visible");
            }
            return 1;
        }
        finally
        {
            RotHelper.ReleaseComReference(comment);
            RotHelper.ReleaseComReference(range);
            RotHelper.ReleaseComReference(sheet);
        }
    }

    private static int RestoreHyperlinkEntry(object workbook, JsonObject entry, RestoreMismatchCollector mismatches)
    {
        var sheetName = Json.GetString(entry, "sheet")!;
        var address = Json.GetString(entry, "range")!;
        object? sheet = null;
        object? range = null;
        object? links = null;
        object? sheetLinks = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, SheetOp(sheetName));
            range = (object)((dynamic)sheet).Range(address);
            links = (object)((dynamic)range).Hyperlinks;
            try { ((dynamic)links).Delete(); }
            catch { /* none */ }
            var state = Json.GetObj(entry, "state");
            if (state is not null && Json.GetBool(state, "present"))
            {
                sheetLinks = (object)((dynamic)sheet).Hyperlinks;
                ((dynamic)sheetLinks).Add(
                    range,
                    Json.GetString(state, "address") ?? "",
                    Json.GetString(state, "subAddress") ?? Type.Missing,
                    Json.GetString(state, "screenTip") ?? Type.Missing,
                    Type.Missing);
            }
            RestoreCellContent(range, state);
            return 1;
        }
        finally
        {
            RotHelper.ReleaseComReference(sheetLinks);
            RotHelper.ReleaseComReference(links);
            RotHelper.ReleaseComReference(range);
            RotHelper.ReleaseComReference(sheet);
        }
    }

    private static int RestorePivotEntry(object workbook, JsonObject entry, RestoreMismatchCollector mismatches)
    {
        var sheetName = Json.GetString(entry, "sheet")!;
        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, SheetOp(sheetName));
            var name = Json.GetString(entry, "name");
            var pivot = string.IsNullOrWhiteSpace(name) ? null : FindPivotTable(sheet, name);
            try
            {
                if (!Json.GetBool(entry, "existed"))
                {
                    if (pivot is not null) DeletePivotTable(pivot);
                    return 1;
                }
                var state = Json.GetObj(entry, "state") ?? new JsonObject();
                if (pivot is null)
                {
                    pivot = RecreatePivotFromState(workbook, sheet, name, state, entry);
                    if (pivot is null)
                    {
                        mismatches.Add($"pivot '{name}' missing during restore; source/destination were not captured");
                        return 0;
                    }
                }
                var source = Json.GetString(state, "source");
                if (!string.IsNullOrWhiteSpace(source))
                    ApplyPivotCacheSource(workbook, pivot, source);
                var layout = new JsonObject
                {
                    ["rows"] = Json.GetArr(state, "rowFields")?.DeepClone(),
                    ["columns"] = Json.GetArr(state, "columnFields")?.DeepClone(),
                    ["filters"] = Json.GetArr(state, "pageFields")?.DeepClone(),
                    ["values"] = MapPivotValues(Json.GetArr(state, "dataFields")),
                };
                ApplyPivotLayout(pivot, layout, replaceOmitted: true);
                return 1;
            }
            finally { RotHelper.ReleaseComReference(pivot); }
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static JsonArray MapPivotValues(JsonArray? dataFields)
    {
        var values = new JsonArray();
        if (dataFields is null) return values;
        foreach (var node in dataFields.OfType<JsonObject>())
        {
            var field = Json.GetString(node, "sourceName") ?? Json.GetString(node, "field") ?? Json.GetString(node, "name");
            if (string.IsNullOrWhiteSpace(field)) continue;
            var mapped = new JsonObject
            {
                ["field"] = field,
                ["function"] = Json.GetString(node, "function") ?? "sum",
            };
            var caption = Json.GetString(node, "caption");
            if (!string.IsNullOrWhiteSpace(caption)) mapped["caption"] = caption;
            var numberFormat = Json.GetString(node, "numberFormat");
            if (!string.IsNullOrWhiteSpace(numberFormat)) mapped["numberFormat"] = numberFormat;
            values.Add(mapped);
        }
        return values;
    }

    private static object? RecreatePivotFromState(object workbook, object sheet, string? name, JsonObject state,
        JsonObject entry)
    {
        var source = Json.GetString(state, "source");
        var destText = Json.GetString(state, "range") ?? Json.GetString(entry, "destination");
        if (string.IsNullOrWhiteSpace(source) || string.IsNullOrWhiteSpace(destText))
            return null;
        var destAddress = destText;
        try
        {
            destAddress = ExcelRangeReference.Parse(destText).Address;
        }
        catch (FormatException) { /* keep raw */ }
        var firstArea = destAddress.Split([',', ';'], StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)[0];
        if (ExcelA1Box.TryParse(firstArea, out var box))
            destAddress = ExcelA1Box.CellName(box.Column, box.Row);

        object? caches = null;
        object? cache = null;
        object? destRange = null;
        object? pivot = null;
        try
        {
            destRange = (object)((dynamic)sheet).Range(destAddress);
            caches = (object)((dynamic)workbook).PivotCaches();
            cache = (object)((dynamic)caches).Create(ExcelDataObjectCatalog.XlDatabase, source);
            pivot = (object)((dynamic)cache).CreatePivotTable(destRange, string.IsNullOrWhiteSpace(name) ? Type.Missing : name);
            var keep = pivot;
            pivot = null;
            return keep;
        }
        finally
        {
            RotHelper.ReleaseComReference(pivot);
            RotHelper.ReleaseComReference(destRange);
            RotHelper.ReleaseComReference(cache);
            RotHelper.ReleaseComReference(caches);
        }
    }

    private static void DeletePivotTable(object pivot)
    {
        object? range = null;
        try
        {
            range = (object)((dynamic)pivot).TableRange2;
            ((dynamic)range).Clear();
        }
        finally { RotHelper.ReleaseComReference(range); }
    }

    private static int RestoreShapeEntry(object workbook, JsonObject entry, RestoreMismatchCollector mismatches)
    {
        var sheetName = Json.GetString(entry, "sheet")!;
        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, SheetOp(sheetName));
            var name = Json.GetString(entry, "name");
            var shape = string.IsNullOrWhiteSpace(name) ? null : FindShape(sheet, name);
            try
            {
                if (!Json.GetBool(entry, "existed"))
                {
                    if (shape is not null) ((dynamic)shape).Delete();
                    return 1;
                }
                var state = Json.GetObj(entry, "state") ?? new JsonObject();
                if (Json.GetString(entry, "kind") == "connector" || state.ContainsKey("connectorType"))
                {
                    mismatches.Add("connector topology is not supported by incremental shape restoration; use the existing workbook recovery artifact");
                    return 0;
                }
                if (shape is null)
                {
                    shape = RecreateDrawing(sheet, entry, state);
                    if (shape is null)
                    {
                        mismatches.Add($"{Json.GetString(entry, "kind")} '{name}' missing during restore");
                        return 0;
                    }
                }
                ApplyPosition(shape, new DataPosition(
                    JsonDouble(state, "left"), JsonDouble(state, "top"),
                    JsonDouble(state, "width"), JsonDouble(state, "height")));
                SetShapeText(shape, Json.GetString(state, "text"));
                ApplyShapeFormatting(shape, state);
                return 1;
            }
            finally { RotHelper.ReleaseComReference(shape); }
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static object? RecreateDrawing(object sheet, JsonObject entry, JsonObject state)
    {
        object? shapes = null;
        try
        {
            shapes = (object)((dynamic)sheet).Shapes;
            var left = JsonDouble(state, "left");
            var top = JsonDouble(state, "top");
            var width = JsonDouble(state, "width") > 0 ? JsonDouble(state, "width") : 120;
            var height = JsonDouble(state, "height") > 0 ? JsonDouble(state, "height") : 40;
            object shape;
            if (Json.GetString(entry, "kind") == "textbox")
            {
                shape = (object)((dynamic)shapes).AddTextbox(
                    ExcelDataObjectCatalog.MsoTextOrientationHorizontal, left, top, width, height);
            }
            else
            {
                var type = Json.GetInt(state, "shapeType") ?? ExcelDataObjectCatalog.MsoShapeRectangle;
                shape = (object)((dynamic)shapes).AddShape(type, left, top, width, height);
            }
            var name = Json.GetString(entry, "name");
            if (!string.IsNullOrWhiteSpace(name))
                ((dynamic)shape).Name = name;
            return shape;
        }
        finally { RotHelper.ReleaseComReference(shapes); }
    }

    private static JsonObject ReadRestoredActual(object workbook, JsonObject entry)
    {
        var kind = Json.GetString(entry, "kind");
        var sheetName = Json.GetString(entry, "sheet");
        try
        {
            return kind switch
            {
                "table" => ReadActualTable(workbook, entry),
                "sortRange" or "duplicates" or "textToColumns" => ReadActualSurface(workbook, entry),
                "autoFilter" => ReadActualFilter(workbook, entry),
                "definedName" => ReadActualName(workbook, entry),
                "validation" => ReadActualRangeState(workbook, entry, ReadValidationState),
                "conditionalFormat" => ReadActualRangeState(workbook, entry, ReadConditionalFormats),
                "richText" => ReadActualRangeState(workbook, entry, (sheet, address) => ReadRichText(sheet, address)),
                "chart" => ReadActualChart(workbook, entry),
                "picture" => ReadActualNamedShape(workbook, entry, picture: true),
                "shape" or "textbox" => ReadActualNamedShape(workbook, entry, picture: false),
                "connector" => ReadActualNamedShape(workbook, entry, picture: false),
                "externalLinks" => ReadActualLinkEntry(workbook, entry),
                "calculationMode" => ReadActualCalculation(workbook, entry),
                "freezeValues" or "pasteSpecial" => ReadActualSurface(workbook, entry),
                "goalSeek" => ReadActualGoalSeek(workbook, entry),
                "workbookProtection" => ReadActualWorkbookProtection(workbook, entry),
                "splitPanes" => ReadActualSplit(workbook, entry),
                "sparkline" => ReadActualSparkline(workbook, entry),
                "slicer" => ReadActualSlicer(workbook, entry),
                "cellStyle" => ReadActualCellStyle(workbook, entry),
                "note" => ReadActualRangeState(workbook, entry, ReadNoteState),
                "hyperlink" => ReadActualRangeState(workbook, entry, ReadHyperlinkState),
                "pivot" => ReadActualPivot(workbook, entry),
                _ => new JsonObject { ["readFailed"] = true, ["error"] = $"no actual reader for {kind}" },
            };
        }
        catch (Exception ex)
        {
            return new JsonObject { ["readFailed"] = true, ["error"] = ex.Message, ["sheet"] = sheetName };
        }
    }

    private static JsonObject ReadActualTable(object workbook, JsonObject entry)
    {
        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, SheetOp(Json.GetString(entry, "sheet")!));
            var name = Json.GetString(Json.GetObj(entry, "state"), "name") ?? Json.GetString(entry, "name");
            var table = string.IsNullOrWhiteSpace(name)
                ? FindOverlappingTableObject(sheet, Json.GetString(entry, "range") ?? "")
                : FindListObject(sheet, name ?? "");
            try
            {
                var actual = new JsonObject
                {
                    ["kind"] = "table",
                    ["existed"] = Json.GetBool(entry, "existed"),
                    ["present"] = table is not null,
                    ["name"] = table is null ? null : ReadComString(table, "Name"),
                    ["state"] = table is null ? null : ReadListObjectState(table),
                };
                var range = Json.GetString(entry, "range") ?? Json.GetString(entry, "formulaRange")
                            ?? Json.GetString(Json.GetObj(entry, "state"), "range");
                if (!string.IsNullOrWhiteSpace(range))
                    MergeSurface(actual, CaptureRangeSurface(sheet, range));
                return actual;
            }
            finally { RotHelper.ReleaseComReference(table); }
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static JsonObject ReadActualSurface(object workbook, JsonObject entry)
    {
        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, SheetOp(Json.GetString(entry, "sheet")!));
            var actual = new JsonObject { ["kind"] = Json.GetString(entry, "kind") };
            MergeSurface(actual, CaptureRangeSurface(sheet, Json.GetString(entry, "range")!));
            return actual;
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static JsonObject ReadActualFilter(object workbook, JsonObject entry)
    {
        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, SheetOp(Json.GetString(entry, "sheet")!));
            return new JsonObject
            {
                ["kind"] = "autoFilter",
                ["state"] = ReadAutoFilterState(sheet, Json.GetString(entry, "sheet")!),
            };
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static JsonObject ReadActualName(object workbook, JsonObject entry)
    {
        var current = FindDefinedName(workbook, Json.GetString(entry, "name")!, Json.GetString(entry, "scope")!,
            Json.GetString(entry, "sheet"));
        try
        {
            return new JsonObject
            {
                ["kind"] = "definedName",
                ["present"] = current is not null,
                ["state"] = current is null ? null : ReadDefinedNameState(current),
                ["refersTo"] = current is null ? null : ReadComString(current, "RefersTo"),
            };
        }
        finally { RotHelper.ReleaseComReference(current); }
    }

    private static JsonObject ReadActualRangeState(object workbook, JsonObject entry,
        Func<object, string, JsonObject> reader)
    {
        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, SheetOp(Json.GetString(entry, "sheet")!));
            return new JsonObject
            {
                ["kind"] = Json.GetString(entry, "kind"),
                ["state"] = reader(sheet, Json.GetString(entry, "range")!),
            };
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static JsonObject ReadActualChart(object workbook, JsonObject entry)
    {
        object? sheet = null;
        var name = Json.GetString(entry, "name");
        var chart = (object?)null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, SheetOp(Json.GetString(entry, "sheet")!));
            chart = string.IsNullOrWhiteSpace(name) ? null : FindChartObject(sheet, name);
            return new JsonObject
            {
                ["kind"] = "chart",
                ["present"] = chart is not null,
                ["state"] = chart is null ? null : ReadChartState(chart, Json.GetString(entry, "sheet")!),
            };
        }
        finally
        {
            RotHelper.ReleaseComReference(chart);
            RotHelper.ReleaseComReference(sheet);
        }
    }

    private static JsonObject ReadActualNamedShape(object workbook, JsonObject entry, bool picture)
    {
        object? sheet = null;
        object? shape = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, SheetOp(Json.GetString(entry, "sheet")!));
            var name = Json.GetString(entry, "name");
            shape = string.IsNullOrWhiteSpace(name) ? null : FindShape(sheet, name);
            JsonObject? state = null;
            if (shape is not null)
                state = picture
                    ? ReadPictureState(shape, Json.GetString(entry, "sheet")!)
                    : ReadShapeState(shape, Json.GetString(entry, "sheet")!);
            return new JsonObject
            {
                ["kind"] = Json.GetString(entry, "kind"),
                ["present"] = shape is not null,
                ["state"] = state,
            };
        }
        finally
        {
            RotHelper.ReleaseComReference(shape);
            RotHelper.ReleaseComReference(sheet);
        }
    }

    private static JsonObject ReadActualPivot(object workbook, JsonObject entry)
    {
        object? sheet = null;
        object? pivot = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, SheetOp(Json.GetString(entry, "sheet")!));
            var name = Json.GetString(entry, "name");
            pivot = string.IsNullOrWhiteSpace(name) ? null : FindPivotTable(sheet, name);
            return new JsonObject
            {
                ["kind"] = "pivot",
                ["present"] = pivot is not null,
                ["state"] = pivot is null ? null : ReadPivotState(pivot, Json.GetString(entry, "sheet")!),
            };
        }
        finally
        {
            RotHelper.ReleaseComReference(pivot);
            RotHelper.ReleaseComReference(sheet);
        }
    }

    private static object? FindOverlappingTableObject(object sheet, string address)
    {
        if (string.IsNullOrWhiteSpace(address)) return null;
        object? range = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            return (object?)((dynamic)range).ListObject;
        }
        catch { return null; }
        finally { RotHelper.ReleaseComReference(range); }
    }

    private static JsonObject SheetOp(string sheetName) => new()
    {
        ["op"] = "data_restore",
        ["target"] = new JsonObject { ["sheet"] = sheetName },
    };

    private static double JsonDouble(JsonObject state, string field) =>
        ExcelDataOperationsContract.TryGetFiniteNumber(state[field], out var value) ? value : 0;
}
