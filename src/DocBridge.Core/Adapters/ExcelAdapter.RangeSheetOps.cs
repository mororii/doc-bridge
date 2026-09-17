using System.Globalization;
using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class ExcelAdapter
{
    private const int XlPasteAll = -4104;
    private const int XlPasteValues = -4163;
    private const int XlPasteFormulas = -4123;
    private const int XlPasteFormats = -4122;
    private const int XlShiftDown = -4121;
    private const int XlShiftToRight = -4161;

    private static void PreviewRangeSheetOperation(object workbook, JsonObject op, ApplyPreview preview)
    {
        var name = Json.GetString(op, "op")!;
        switch (name)
        {
            case "clear_range":
            case "copy_range":
                PreviewRangeEdit(workbook, op, preview);
                return;
            case "delete_rows":
            case "delete_cols":
                PreviewDeleteStrip(workbook, op, preview);
                return;
            case "rename_sheet":
                PreviewRenameSheet(workbook, op, preview);
                return;
            case "add_sheet":
            case "move_sheet":
                PreviewStructureOp(workbook, op, preview);
                return;
            case "protect_sheet":
            case "unprotect_sheet":
                PreviewProtectOp(workbook, op, preview);
                return;
        }
    }

    private static void ApplyRangeSheetOperation(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        var name = Json.GetString(op, "op")!;
        switch (name)
        {
            case "clear_range":
                ApplyClearRange(workbook, op, execution, mismatches, ref checkedItems);
                return;
            case "copy_range":
                ApplyCopyRange(workbook, op, execution, mismatches, ref checkedItems);
                return;
            case "delete_rows":
            case "delete_cols":
                ApplyDeleteStrip(workbook, op, execution, mismatches, ref checkedItems);
                return;
            case "rename_sheet":
                ApplyRenameSheet(workbook, op, execution, mismatches, ref checkedItems);
                return;
            case "add_sheet":
                ApplyAddSheet(workbook, op, execution, mismatches, ref checkedItems);
                return;
            case "move_sheet":
                ApplyMoveSheet(workbook, op, execution, mismatches, ref checkedItems);
                return;
            case "protect_sheet":
            case "unprotect_sheet":
                ApplyProtectOp(workbook, op, execution, mismatches, ref checkedItems);
                return;
        }
    }

    private static JsonObject CaptureRenameState(object workbook, IReadOnlyList<JsonObject> ops, string? documentRef)
    {
        var entries = new JsonArray();
        foreach (var op in ops)
        {
            var sheetName = Json.GetString(Json.GetObj(op, "target"), "sheet") ?? "";
            entries.Add(new JsonObject
            {
                ["oldName"] = sheetName,
                ["newName"] = Json.GetString(op, "newName"),
            });
        }

        return new JsonObject
        {
            ["snapshotVersion"] = ExcelLayoutSnapshotVersion,
            ["restoreMode"] = RenameRestoreMode,
            ["documentRef"] = documentRef,
            ["originalActiveSheet"] = ReadActiveSheetName(workbook),
            ["entries"] = entries,
        };
    }

    private static JsonObject RestoreRenameState(object workbook, JsonObject state)
    {
        var mismatches = new RestoreMismatchCollector();
        var checkedItems = 0;
        var entries = Json.GetArr(state, "entries") ?? new JsonArray();
        for (var index = entries.Count - 1; index >= 0; index--)
        {
            if (entries[index] is not JsonObject entry) continue;
            var oldName = Json.GetString(entry, "oldName");
            var newName = Json.GetString(entry, "newName");
            if (string.IsNullOrWhiteSpace(oldName) || string.IsNullOrWhiteSpace(newName)) continue;
            object? sheet = null;
            try
            {
                sheet = FindWorksheet(workbook, newName) ?? FindWorksheet(workbook, oldName);
                if (sheet is null)
                {
                    mismatches.Add($"rename restore: sheet '{oldName}'/'{newName}' missing");
                    continue;
                }

                if (!string.Equals(Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture), oldName,
                        StringComparison.OrdinalIgnoreCase))
                    ((dynamic)sheet).Name = oldName;
                checkedItems++;
            }
            catch (Exception ex) { mismatches.Add($"rename restore failed: {ex.Message}"); }
            finally { RotHelper.ReleaseComReference(sheet); }
        }

        var active = Json.GetString(state, "originalActiveSheet");
        if (!string.IsNullOrWhiteSpace(active)) ActivateWorksheet(workbook, active, mismatches);
        return BuildRestoreResult(mismatches.Count == 0, 0, checkedItems, RenameRestoreMode, mismatches);
    }

    private static JsonObject CaptureRangeEditState(object workbook, IReadOnlyList<JsonObject> ops, string? documentRef)
    {
        var entries = new JsonArray();
        foreach (var op in ops)
        {
            var name = Json.GetString(op, "op")!;
            if (name == "clear_range")
            {
                var resolved = ResolveBasicRange(workbook, op);
                try
                {
                    entries.Add(new JsonObject
                    {
                        ["op"] = name,
                        ["sheet"] = Convert.ToString(((dynamic)resolved.Sheet).Name, CultureInfo.InvariantCulture),
                        ["range"] = resolved.Address,
                        ["what"] = (Json.GetString(op, "what") ?? "all").ToLowerInvariant(),
                        ["payload"] = CaptureRangePayload(resolved.Sheet, resolved.Address),
                    });
                }
                finally
                {
                    RotHelper.ReleaseComReference(resolved.Range);
                    RotHelper.ReleaseComReference(resolved.Sheet);
                }
            }
            else if (name == "copy_range")
            {
                var dest = ResolveCopyDestination(workbook, op);
                try
                {
                    entries.Add(new JsonObject
                    {
                        ["op"] = name,
                        ["sheet"] = dest.SheetName,
                        ["range"] = dest.Address,
                        ["payload"] = CaptureRangePayload(dest.Sheet, dest.Address),
                    });
                }
                finally
                {
                    RotHelper.ReleaseComReference(dest.Range);
                    RotHelper.ReleaseComReference(dest.Sheet);
                }
            }
        }

        return new JsonObject
        {
            ["snapshotVersion"] = ExcelLayoutSnapshotVersion,
            ["restoreMode"] = RangeEditRestoreMode,
            ["documentRef"] = documentRef,
            ["originalActiveSheet"] = ReadActiveSheetName(workbook),
            ["entries"] = entries,
        };
    }

    private static JsonObject RestoreRangeEditState(object workbook, JsonObject state)
    {
        var mismatches = new RestoreMismatchCollector();
        var checkedItems = 0;
        var entries = Json.GetArr(state, "entries") ?? new JsonArray();
        for (var index = entries.Count - 1; index >= 0; index--)
        {
            if (entries[index] is not JsonObject entry) continue;
            object? sheet = null;
            try
            {
                var sheetName = Json.GetString(entry, "sheet");
                sheet = GetExplicitTargetSheetReference(workbook, SheetTargetOp(sheetName));
                checkedItems += RestoreRangePayload(sheet, sheetName, Json.GetString(entry, "range")!,
                    Json.GetObj(entry, "payload"), mismatches);
            }
            catch (Exception ex) { mismatches.Add($"range-edit restore failed: {ex.Message}"); }
            finally { RotHelper.ReleaseComReference(sheet); }
        }

        var active = Json.GetString(state, "originalActiveSheet");
        if (!string.IsNullOrWhiteSpace(active)) ActivateWorksheet(workbook, active, mismatches);
        return BuildRestoreResult(mismatches.Count == 0, 0, checkedItems, RangeEditRestoreMode, mismatches);
    }

    private static JsonObject CaptureDeleteState(object workbook, IReadOnlyList<JsonObject> ops, string? documentRef,
        string? snapshotDir = null)
    {
        var entries = new JsonArray();
        foreach (var op in ops)
        {
            object? sheet = null;
            try
            {
                sheet = GetExplicitTargetSheetReference(workbook, op);
                var sheetName = Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture) ?? "";
                var rows = Json.GetString(op, "op") == "delete_rows";
                var start = rows ? Json.GetInt(op, "row")!.Value : ParseColumnNumber(op["col"]);
                var count = Json.GetInt(op, "count")!.Value;
                var address = rows
                    ? UsedStripAddress(sheet, start, count, rows: true)
                    : UsedStripAddress(sheet, start, count, rows: false);
                entries.Add(new JsonObject
                {
                    ["op"] = Json.GetString(op, "op"),
                    ["sheet"] = sheetName,
                    ["start"] = start,
                    ["count"] = count,
                    ["rows"] = rows,
                    ["applied"] = false,
                    ["address"] = address,
                    ["sizes"] = rows ? CaptureRowHeights(sheet, start, count) : CaptureColumnWidths(sheet, start, count),
                    ["payload"] = string.IsNullOrWhiteSpace(address) ? null : CaptureRangePayload(sheet, address),
                    ["names"] = CaptureIntersectingNames(workbook, sheetName),
                    ["externalFormulas"] = CaptureExternalFormulas(workbook, sheetName),
                });
            }
            finally { RotHelper.ReleaseComReference(sheet); }
        }

        return new JsonObject
        {
            ["snapshotVersion"] = ExcelLayoutSnapshotVersion,
            ["restoreMode"] = DeleteStripRestoreMode,
            ["recoveryCopyRequired"] = true,
            ["documentRef"] = documentRef,
            ["originalActiveSheet"] = ReadActiveSheetName(workbook),
            ["entries"] = entries,
        };
    }

    private static JsonObject RestoreDeleteState(object workbook, JsonObject state)
    {
        var mismatches = new RestoreMismatchCollector();
        var checkedItems = 0;
        var entries = Json.GetArr(state, "entries") ?? new JsonArray();
        for (var index = entries.Count - 1; index >= 0; index--)
        {
            if (entries[index] is not JsonObject entry) continue;
            if (!ExcelAuthoringSchema.ShouldRestoreDeleteEntry(entry))
                continue;
            object? sheet = null;
            object? collection = null;
            object? target = null;
            try
            {
                var sheetName = Json.GetString(entry, "sheet");
                sheet = GetExplicitTargetSheetReference(workbook, SheetTargetOp(sheetName));
                var start = Json.GetInt(entry, "start")!.Value;
                var count = Json.GetInt(entry, "count")!.Value;
                var rows = Json.GetBool(entry, "rows");
                collection = rows ? (object)((dynamic)sheet).Rows : (object)((dynamic)sheet).Columns;
                var key = rows ? $"{start}:{start + count - 1}" : $"{ColName(start)}:{ColName(start + count - 1)}";
                target = (object)((dynamic)collection)[key];
                ((dynamic)target).Insert(rows ? XlShiftDown : XlShiftToRight);
                var address = Json.GetString(entry, "address");
                if (!string.IsNullOrWhiteSpace(address))
                    checkedItems += RestoreRangePayload(sheet, sheetName, address, Json.GetObj(entry, "payload"), mismatches);
                checkedItems += RestoreSizes(sheet, Json.GetArr(entry, "sizes") ?? new JsonArray(), rows, mismatches);
                checkedItems += RestoreIntersectingNames(workbook, Json.GetArr(entry, "names"), mismatches);
                checkedItems += RestoreExternalFormulas(workbook, Json.GetArr(entry, "externalFormulas"), mismatches);
            }
            catch (Exception ex) { mismatches.Add($"delete restore failed: {ex.Message}"); }
            finally
            {
                RotHelper.ReleaseComReference(target);
                RotHelper.ReleaseComReference(collection);
                RotHelper.ReleaseComReference(sheet);
            }
        }

        var active = Json.GetString(state, "originalActiveSheet");
        if (!string.IsNullOrWhiteSpace(active)) ActivateWorksheet(workbook, active, mismatches);
        return BuildRestoreResult(mismatches.Count == 0, 0, checkedItems, DeleteStripRestoreMode, mismatches);
    }

    private static JsonObject CaptureStructureState(object workbook, IReadOnlyList<JsonObject> ops, string? documentRef,
        string snapshotDir, JsonObject backupMetadata)
    {
        var deletedSheets = new JsonArray();
        var names = ReadOrderedWorksheetNames(workbook);
        foreach (var op in ops)
        {
            if (!string.Equals(Json.GetString(op, "op"), "delete_sheet", StringComparison.OrdinalIgnoreCase))
                continue;
            object? sheet = null;
            try
            {
                sheet = GetExplicitTargetSheetReference(workbook, op);
                var sheetName = Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture) ?? "";
                deletedSheets.Add(new JsonObject
                {
                    ["name"] = sheetName,
                    ["index"] = names.FindIndex(item => string.Equals(item, sheetName, StringComparison.OrdinalIgnoreCase)) + 1,
                    [ExcelDeleteSheetDependencyContract.CaptureKeyNames] = CaptureIntersectingNames(workbook, sheetName),
                    [ExcelDeleteSheetDependencyContract.CaptureKeyFormulas] = CaptureExternalFormulas(workbook, sheetName),
                    [ExcelDeleteSheetDependencyContract.CaptureKeyCharts] = CaptureChartDependencies(workbook, sheetName),
                    [ExcelDeleteSheetDependencyContract.CaptureKeyPivots] = CapturePivotDependencies(workbook, sheetName),
                });
            }
            finally { RotHelper.ReleaseComReference(sheet); }
        }

        JsonObject? recovery = null;
        if (deletedSheets.Count > 0)
        {
            var backupFile = Json.GetString(backupMetadata, "workbookBackup");
            var backupPath = string.IsNullOrWhiteSpace(backupFile) ? null : Path.Combine(snapshotDir, backupFile);
            if (!ExcelDeleteSheetRecoveryContract.IsFullRecoveryAvailable(backupMetadata, backupPath))
                throw new InvalidOperationException(
                    ExcelDeleteSheetRecoveryContract.RejectReason(backupMetadata, backupPath));
            recovery = ExcelDeleteSheetRecoveryContract.RecoveryEnvelope(
                ExcelDeleteSheetRecoveryContract.ModeWorkbookCopySheet,
                backupPath,
                Json.GetString(backupMetadata, "workbookBackupSource"),
                documentRef,
                deletedSheets.OfType<JsonObject>().ToList());
        }

        return new JsonObject
        {
            ["snapshotVersion"] = ExcelLayoutSnapshotVersion,
            ["restoreMode"] = StructureRestoreMode,
            ["documentRef"] = documentRef,
            ["originalActiveSheet"] = ReadActiveSheetName(workbook),
            ["originalSheets"] = StringsToJsonArray(names),
            ["deletedSheets"] = deletedSheets,
            ["sheetRecovery"] = recovery,
            ["ops"] = CloneOps(ops),
        };
    }

    private static JsonObject RestoreStructureState(object application, object workbook, JsonObject state,
        string? snapshotDir = null)
    {
        var mismatches = new RestoreMismatchCollector();
        var original = ReadRequiredStringArray(state, "originalSheets") ?? new List<string>();
        var originalSet = original.ToHashSet(StringComparer.OrdinalIgnoreCase);
        var current = ReadOrderedWorksheetNames(workbook);
        var checkedItems = 0;
        checkedItems += RestoreDeletedSheetsFromBackup(application, workbook, state, snapshotDir, mismatches);
        checkedItems += RestoreDeletedSheetDependencies(workbook, state, mismatches);
        current = ReadOrderedWorksheetNames(workbook);
        foreach (var name in current.Where(item => !originalSet.Contains(item)).Reverse())
        {
            if (!TryDeleteWorksheet(application, workbook, name, mismatches))
                mismatches.Add($"structure restore could not delete added sheet '{name}'");
            else
                checkedItems++;
        }

        current = ReadOrderedWorksheetNames(workbook);
        for (var index = 0; index < original.Count; index++)
        {
            var expected = original[index];
            if (index >= current.Count || !string.Equals(current[index], expected, StringComparison.OrdinalIgnoreCase))
            {
                object? sheet = null;
                object? relative = null;
                try
                {
                    sheet = FindWorksheet(workbook, expected);
                    if (sheet is null)
                    {
                        mismatches.Add($"structure restore missing original sheet '{expected}'");
                        continue;
                    }

                    if (index < current.Count)
                    {
                        relative = FindWorksheet(workbook, current[index]);
                        if (relative is not null) ((dynamic)sheet).Move((dynamic)relative);
                    }
                    else
                    {
                        ((dynamic)sheet).Move(Type.Missing, AfterLastWorksheet(workbook));
                    }

                    checkedItems++;
                }
                catch (Exception ex) { mismatches.Add($"structure restore move '{expected}' failed: {ex.Message}"); }
                finally
                {
                    RotHelper.ReleaseComReference(relative);
                    RotHelper.ReleaseComReference(sheet);
                }

                current = ReadOrderedWorksheetNames(workbook);
            }
        }

        var active = Json.GetString(state, "originalActiveSheet");
        if (!string.IsNullOrWhiteSpace(active)) ActivateWorksheet(workbook, active, mismatches);
        return BuildRestoreResult(mismatches.Count == 0, 0, checkedItems, StructureRestoreMode, mismatches);
    }

    private static JsonObject CaptureProtectState(object workbook, IReadOnlyList<JsonObject> ops, string? documentRef)
    {
        var entries = new JsonArray();
        foreach (var op in ops)
        {
            object? sheet = null;
            try
            {
                sheet = GetExplicitTargetSheetReference(workbook, op);
                entries.Add(new JsonObject
                {
                    ["sheet"] = Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture),
                    ["protectContents"] = Convert.ToBoolean(((dynamic)sheet).ProtectContents, CultureInfo.InvariantCulture),
                    ["protectDrawingObjects"] = SafeBool(() => ((dynamic)sheet).ProtectDrawingObjects),
                    ["protectScenarios"] = SafeBool(() => ((dynamic)sheet).ProtectScenarios),
                });
            }
            finally { RotHelper.ReleaseComReference(sheet); }
        }

        return new JsonObject
        {
            ["snapshotVersion"] = ExcelLayoutSnapshotVersion,
            ["restoreMode"] = ProtectRestoreMode,
            ["documentRef"] = documentRef,
            ["originalActiveSheet"] = ReadActiveSheetName(workbook),
            ["entries"] = entries,
        };
    }

    private static JsonObject RestoreProtectState(object workbook, JsonObject state)
    {
        var mismatches = new RestoreMismatchCollector();
        var checkedItems = 0;
        foreach (var node in Json.GetArr(state, "entries") ?? new JsonArray())
        {
            if (node is not JsonObject entry) continue;
            object? sheet = null;
            try
            {
                var sheetName = Json.GetString(entry, "sheet");
                sheet = GetExplicitTargetSheetReference(workbook, SheetTargetOp(sheetName));
                var wanted = Json.GetBool(entry, "protectContents");
                var actual = Convert.ToBoolean(((dynamic)sheet).ProtectContents, CultureInfo.InvariantCulture);
                if (wanted && !actual)
                    ProtectSheetCom(sheet, userInterfaceOnly: true);
                else if (!wanted && actual)
                    ((dynamic)sheet).Unprotect();
                checkedItems++;
                if (Convert.ToBoolean(((dynamic)sheet).ProtectContents, CultureInfo.InvariantCulture) != wanted)
                    mismatches.Add($"{sheetName}: protectContents restore mismatch");
            }
            catch (Exception ex) { mismatches.Add($"protect restore failed: {ex.Message}"); }
            finally { RotHelper.ReleaseComReference(sheet); }
        }

        var active = Json.GetString(state, "originalActiveSheet");
        if (!string.IsNullOrWhiteSpace(active)) ActivateWorksheet(workbook, active, mismatches);
        return BuildRestoreResult(mismatches.Count == 0, 0, checkedItems, ProtectRestoreMode, mismatches);
    }

    private static void PreviewRangeEdit(object workbook, JsonObject op, ApplyPreview preview)
    {
        var name = Json.GetString(op, "op")!;
        if (name == "clear_range")
        {
            var resolved = ResolveBasicRange(workbook, op);
            try
            {
                var sheetName = Convert.ToString(((dynamic)resolved.Sheet).Name, CultureInfo.InvariantCulture);
                preview.Affected.Add(new AffectedRef("range", $"{sheetName}!{resolved.Address}"));
                preview.Diff.Add(new DiffEntry
                {
                    Ref = $"{sheetName}!{resolved.Address}",
                    Before = Json.GetString(op, "what") ?? "all",
                    After = "cleared",
                });
            }
            finally
            {
                RotHelper.ReleaseComReference(resolved.Range);
                RotHelper.ReleaseComReference(resolved.Sheet);
            }
            return;
        }

        var source = ResolveBasicRange(workbook, op);
        var dest = ResolveCopyDestination(workbook, op);
        try
        {
            preview.Affected.Add(new AffectedRef("range", $"{dest.SheetName}!{dest.Address}"));
            preview.Diff.Add(new DiffEntry
            {
                Ref = $"{dest.SheetName}!{dest.Address}",
                Before = CaptureRangePayload(dest.Sheet, dest.Address),
                After = $"{Json.GetString(op, "mode") ?? "all"} from {Convert.ToString(((dynamic)source.Sheet).Name, CultureInfo.InvariantCulture)}!{source.Address}",
            });
        }
        finally
        {
            RotHelper.ReleaseComReference(dest.Range);
            RotHelper.ReleaseComReference(dest.Sheet);
            RotHelper.ReleaseComReference(source.Range);
            RotHelper.ReleaseComReference(source.Sheet);
        }
    }

    private static void PreviewDeleteStrip(object workbook, JsonObject op, ApplyPreview preview)
    {
        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, op);
            var sheetName = Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture);
            var rows = Json.GetString(op, "op") == "delete_rows";
            var start = rows ? Json.GetInt(op, "row")!.Value : ParseColumnNumber(op["col"]);
            var count = Json.GetInt(op, "count")!.Value;
            var label = rows
                ? $"{sheetName}!rows {start}:{start + count - 1}"
                : $"{sheetName}!cols {ColName(start)}:{ColName(start + count - 1)}";
            preview.Affected.Add(new AffectedRef(rows ? "rows" : "cols", label));
            preview.Diff.Add(new DiffEntry { Ref = label, Before = "present", After = "deleted" });
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static void PreviewRenameSheet(object workbook, JsonObject op, ApplyPreview preview)
    {
        var oldName = Json.GetString(Json.GetObj(op, "target"), "sheet")!;
        var newName = Json.GetString(op, "newName")!;
        if (SheetExists(workbook, newName) &&
            !string.Equals(oldName, newName, StringComparison.OrdinalIgnoreCase))
            preview.Errors.Add($"sheet '{newName}' already exists");
        preview.Affected.Add(new AffectedRef("sheet", oldName));
        preview.Diff.Add(new DiffEntry { Ref = "sheetName", Before = oldName, After = newName });
    }

    private static void PreviewStructureOp(object workbook, JsonObject op, ApplyPreview preview)
    {
        var name = Json.GetString(op, "op");
        if (name == "add_sheet")
        {
            var sheetName = Json.GetString(op, "name")!;
            if (SheetExists(workbook, sheetName))
                preview.Errors.Add($"sheet '{sheetName}' already exists");
            preview.Affected.Add(new AffectedRef("sheet", sheetName));
            preview.Diff.Add(new DiffEntry { Ref = "sheet", Before = null, After = sheetName });
            return;
        }

        var sheetNameTarget = Json.GetString(Json.GetObj(op, "target"), "sheet")!;
        preview.Affected.Add(new AffectedRef("sheet", sheetNameTarget));
        preview.Diff.Add(new DiffEntry
        {
            Ref = "sheetOrder",
            Before = StringsToJsonArray(ReadOrderedWorksheetNames(workbook)),
            After = Json.GetString(op, "position"),
        });
    }

    private static void PreviewProtectOp(object workbook, JsonObject op, ApplyPreview preview)
    {
        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, op);
            var sheetName = Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture);
            var protectedNow = Convert.ToBoolean(((dynamic)sheet).ProtectContents, CultureInfo.InvariantCulture);
            preview.Affected.Add(new AffectedRef("sheet", sheetName ?? ""));
            preview.Diff.Add(new DiffEntry
            {
                Ref = $"{sheetName}:protectContents",
                Before = protectedNow,
                After = Json.GetString(op, "op") == "protect_sheet",
            });
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static void ApplyClearRange(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        var resolved = ResolveBasicRange(workbook, op);
        try
        {
            var what = (Json.GetString(op, "what") ?? "all").ToLowerInvariant();
            dynamic range = resolved.Range;
            switch (what)
            {
                case "contents":
                    range.ClearContents();
                    break;
                case "formats":
                    range.ClearFormats();
                    break;
                case "formulas":
                    try { range.SpecialCells(XlCellTypeFormulas).ClearContents(); }
                    catch (Exception) { /* no formula cells */ }
                    break;
                default:
                    range.Clear();
                    break;
            }

            var sheetName = Convert.ToString(((dynamic)resolved.Sheet).Name, CultureInfo.InvariantCulture);
            execution.Affected.Add(new AffectedRef("range", $"{sheetName}!{resolved.Address}"));
            checkedItems++;
        }
        finally
        {
            RotHelper.ReleaseComReference(resolved.Range);
            RotHelper.ReleaseComReference(resolved.Sheet);
        }
    }

    private static void ApplyCopyRange(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        var source = ResolveBasicRange(workbook, op);
        var dest = ResolveCopyDestination(workbook, op);
        object? app = null;
        try
        {
            app = (object)((dynamic)workbook).Application;
            var mode = (Json.GetString(op, "mode") ?? "all").ToLowerInvariant();
            if (mode == "all")
            {
                ((dynamic)source.Range).Copy(dest.Range);
            }
            else
            {
                ((dynamic)source.Range).Copy();
                var paste = mode switch
                {
                    "values" => XlPasteValues,
                    "formulas" => XlPasteFormulas,
                    "formats" => XlPasteFormats,
                    _ => XlPasteAll,
                };
                ((dynamic)dest.Range).PasteSpecial(paste);
            }

            try { ((dynamic)app).CutCopyMode = false; } catch { }

            var destRows = Convert.ToInt32(((dynamic)source.Range).Rows.Count, CultureInfo.InvariantCulture);
            var destCols = Convert.ToInt32(((dynamic)source.Range).Columns.Count, CultureInfo.InvariantCulture);
            object? destCheck = null;
            try
            {
                destCheck = (object)((dynamic)dest.Range).Resize(destRows, destCols);
                if (mode is "all" or "values")
                {
                    var sourceValues = JsonArrayToStringTable(RangeToJson(source.Range, out var sourceCells, maxCells: MaxSnapshotCells));
                    var destValues = JsonArrayToStringTable(RangeToJson(destCheck, out var destCells, maxCells: MaxSnapshotCells));
                    if (!ExcelCsvContract.TablesEqual(sourceValues, destValues))
                        mismatches.Add($"copy_range values readback mismatch {dest.SheetName}!{dest.Address}");
                    else
                        checkedItems += destCells;
                }

                if (mode is "all" or "formulas")
                {
                    var sourceFormulas = JsonArrayToStringTable(RangeToJson(source.Range, out _, formulas: true, maxCells: MaxSnapshotCells));
                    var destFormulas = JsonArrayToStringTable(RangeToJson(destCheck, out var formulaCells, formulas: true, maxCells: MaxSnapshotCells));
                    var sourceRow = Convert.ToInt32(((dynamic)source.Range).Row, CultureInfo.InvariantCulture);
                    var sourceCol = Convert.ToInt32(((dynamic)source.Range).Column, CultureInfo.InvariantCulture);
                    var destRow = Convert.ToInt32(((dynamic)destCheck).Row, CultureInfo.InvariantCulture);
                    var destCol = Convert.ToInt32(((dynamic)destCheck).Column, CultureInfo.InvariantCulture);
                    if (!ExcelFillCopyContract.FormulasMatchWithRelativeShift(
                            sourceFormulas, destFormulas, destRow - sourceRow, destCol - sourceCol))
                        mismatches.Add($"copy_range relative formula readback mismatch {dest.SheetName}!{dest.Address}");
                    else
                        checkedItems += formulaCells;
                }

                if (mode == "formats")
                    checkedItems += CompareCopiedFormats(source.Range, destCheck, dest.SheetName, dest.Address, mismatches);
            }
            finally { RotHelper.ReleaseComReference(destCheck); }

            execution.Affected.Add(new AffectedRef("range", $"{dest.SheetName}!{dest.Address}"));
        }
        finally
        {
            RotHelper.ReleaseComReference(app);
            RotHelper.ReleaseComReference(dest.Range);
            RotHelper.ReleaseComReference(dest.Sheet);
            RotHelper.ReleaseComReference(source.Range);
            RotHelper.ReleaseComReference(source.Sheet);
        }
    }

    private static void ApplyDeleteStrip(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        object? sheet = null;
        object? collection = null;
        object? target = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, op);
            var sheetName = Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture);
            var rows = Json.GetString(op, "op") == "delete_rows";
            var start = rows ? Json.GetInt(op, "row")!.Value : ParseColumnNumber(op["col"]);
            var count = Json.GetInt(op, "count")!.Value;
            collection = rows ? (object)((dynamic)sheet).Rows : (object)((dynamic)sheet).Columns;
            var key = rows ? $"{start}:{start + count - 1}" : $"{ColName(start)}:{ColName(start + count - 1)}";
            target = (object)((dynamic)collection)[key];
            ((dynamic)target).Delete();
            execution.Affected.Add(new AffectedRef(rows ? "rows" : "cols",
                rows ? $"{sheetName}!{start}:{start + count - 1}" : $"{sheetName}!{ColName(start)}:{ColName(start + count - 1)}"));
            checkedItems++;
            PatchSnapshotState(_lastSnapshotDir, state =>
            {
                var entries = Json.GetArr(state, "entries");
                if (entries is null) return;
                for (var index = 0; index < entries.Count; index++)
                {
                    if (entries[index] is not JsonObject entry) continue;
                    if (string.Equals(Json.GetString(entry, "sheet"), sheetName, StringComparison.OrdinalIgnoreCase) &&
                        Json.GetInt(entry, "start") == start &&
                        Json.GetInt(entry, "count") == count)
                        ExcelAuthoringSchema.MarkDeleteEntryApplied(state, index);
                }
            });
        }
        finally
        {
            RotHelper.ReleaseComReference(target);
            RotHelper.ReleaseComReference(collection);
            RotHelper.ReleaseComReference(sheet);
        }
    }

    private static void ApplyRenameSheet(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, op);
            var oldName = Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture);
            var errors = new List<string>();
            if (!ExcelSheetNameContract.TryNormalize(Json.GetString(op, "newName"), 1, "rename_sheet.newName", errors, out var newName))
                throw new InvalidOperationException(string.Join("; ", errors));
            if (SheetExists(workbook, newName) &&
                !string.Equals(oldName, newName, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException($"sheet '{newName}' already exists");
            ((dynamic)sheet).Name = newName;
            checkedItems++;
            var actual = Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture);
            if (!string.Equals(actual, newName, StringComparison.OrdinalIgnoreCase))
                mismatches.Add($"rename_sheet readback mismatch: want '{newName}', got '{actual}'");
            execution.Affected.Add(new AffectedRef("sheet", $"{oldName}->{newName}"));
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static void ApplyAddSheet(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        var errors = new List<string>();
        if (!ExcelSheetNameContract.TryNormalize(Json.GetString(op, "name"), 1, "add_sheet.name", errors, out var name))
            throw new InvalidOperationException(string.Join("; ", errors));
        if (SheetExists(workbook, name))
            throw new InvalidOperationException($"sheet '{name}' already exists");
        object? sheets = null;
        object? created = null;
        object? after = null;
        try
        {
            sheets = (object)((dynamic)workbook).Worksheets;
            var afterName = Json.GetString(op, "afterSheet");
            if (!string.IsNullOrWhiteSpace(afterName))
            {
                after = FindWorksheet(workbook, afterName);
                created = after is null
                    ? (object)((dynamic)sheets).Add()
                    : (object)((dynamic)sheets).Add(Type.Missing, after);
            }
            else
            {
                created = (object)((dynamic)sheets).Add();
            }

            ((dynamic)created).Name = name;
            checkedItems++;
            if (!SheetExists(workbook, name))
                mismatches.Add($"add_sheet '{name}' was not present after apply");
            execution.Affected.Add(new AffectedRef("sheet", name));
        }
        finally
        {
            RotHelper.ReleaseComReference(created);
            RotHelper.ReleaseComReference(after);
            RotHelper.ReleaseComReference(sheets);
        }
    }

    private static void ApplyMoveSheet(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        object? sheet = null;
        object? relative = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, op);
            var position = (Json.GetString(op, "position") ?? "").ToLowerInvariant();
            var names = ReadOrderedWorksheetNames(workbook);
            switch (position)
            {
                case "first":
                    relative = FindWorksheet(workbook, names[0]);
                    if (relative is not null) ((dynamic)sheet).Move(relative);
                    break;
                case "last":
                    ((dynamic)sheet).Move(Type.Missing, AfterLastWorksheet(workbook));
                    break;
                case "before":
                    relative = FindWorksheet(workbook, Json.GetString(op, "relativeSheet"));
                    if (relative is null) throw new InvalidOperationException("move_sheet relativeSheet was not found");
                    ((dynamic)sheet).Move(relative);
                    break;
                case "after":
                    relative = FindWorksheet(workbook, Json.GetString(op, "relativeSheet"));
                    if (relative is null) throw new InvalidOperationException("move_sheet relativeSheet was not found");
                    ((dynamic)sheet).Move(Type.Missing, relative);
                    break;
                default:
                    throw new InvalidOperationException("move_sheet position must be before|after|first|last");
            }

            checkedItems++;
            execution.Affected.Add(new AffectedRef("sheet",
                Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture) ?? ""));
        }
        finally
        {
            RotHelper.ReleaseComReference(relative);
            RotHelper.ReleaseComReference(sheet);
        }
    }

    private static void ApplyProtectOp(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, op);
            var sheetName = Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture);
            var protect = Json.GetString(op, "op") == "protect_sheet";
            var options = ExcelAuthoringSchema.NormalizeProtectionOptions(op);
            if (protect) ProtectSheetCom(sheet, options);
            else ((dynamic)sheet).Unprotect();
            checkedItems++;
            var actual = Convert.ToBoolean(((dynamic)sheet).ProtectContents, CultureInfo.InvariantCulture);
            if (actual != protect)
                mismatches.Add($"{sheetName}: protectContents readback mismatch");
            if (protect && options.ContainsKey("allowFormattingCells"))
            {
                var wanted = Json.GetBool(options, "allowFormattingCells");
                var enabled = SafeBool(() => ((dynamic)sheet).Protection.AllowFormattingCells);
                if (enabled != wanted)
                    mismatches.Add($"{sheetName}: allowFormattingCells readback mismatch");
            }
            execution.Affected.Add(new AffectedRef("sheet", sheetName ?? ""));
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static void ProtectSheetCom(object sheet, JsonObject? options = null, bool userInterfaceOnly = true)
    {
        options ??= new JsonObject();
        ((dynamic)sheet).Protect(
            Type.Missing,
            ExcelAuthoringSchema.OptionalBool(options, "drawingObjects", true),
            ExcelAuthoringSchema.OptionalBool(options, "contents", true),
            ExcelAuthoringSchema.OptionalBool(options, "scenarios", true),
            ExcelAuthoringSchema.OptionalBool(options, "userInterfaceOnly", userInterfaceOnly),
            ExcelAuthoringSchema.OptionalBool(options, "allowFormattingCells", false),
            ExcelAuthoringSchema.OptionalBool(options, "allowFormattingColumns", false),
            ExcelAuthoringSchema.OptionalBool(options, "allowFormattingRows", false),
            ExcelAuthoringSchema.OptionalBool(options, "allowInsertingColumns", false),
            ExcelAuthoringSchema.OptionalBool(options, "allowInsertingRows", false),
            ExcelAuthoringSchema.OptionalBool(options, "allowInsertingHyperlinks", false),
            ExcelAuthoringSchema.OptionalBool(options, "allowDeletingColumns", false),
            ExcelAuthoringSchema.OptionalBool(options, "allowDeletingRows", false),
            ExcelAuthoringSchema.OptionalBool(options, "allowSorting", false),
            ExcelAuthoringSchema.OptionalBool(options, "allowFiltering", false),
            ExcelAuthoringSchema.OptionalBool(options, "allowUsingPivotTables", false));
    }

    private static (object Sheet, object Range, string Address, string SheetName) ResolveCopyDestination(
        object workbook, JsonObject op)
    {
        var destRange = Json.GetString(op, "destRange")!;
        var destSheet = Json.GetString(op, "destSheet") ?? Json.GetString(Json.GetObj(op, "target"), "sheet");
        var resolved = ResolveRangeTarget(workbook, destSheet, destRange, requireExplicitSheet: true);
        object? range = null;
        try
        {
            range = (object)((dynamic)resolved.Sheet).Range(resolved.Address);
            var canonical = Convert.ToString(((dynamic)range).Address(false, false), CultureInfo.InvariantCulture)
                            ?? resolved.Address;
            var sheetName = Convert.ToString(((dynamic)resolved.Sheet).Name, CultureInfo.InvariantCulture) ?? destSheet ?? "";
            return (resolved.Sheet, range, canonical, sheetName);
        }
        catch
        {
            RotHelper.ReleaseComReference(range);
            RotHelper.ReleaseComReference(resolved.Sheet);
            throw;
        }
    }

    private static JsonObject CaptureRangePayload(object sheet, string address)
    {
        object? range = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            var values = RangeToJson(range, out var cells, maxCells: MaxSnapshotCells);
            var formulas = RangeToJson(range, out _, formulas: true, maxCells: MaxSnapshotCells);
            if (cells > MaxSnapshotCells)
                throw new InvalidOperationException(
                    $"range snapshot exceeds {MaxSnapshotCells} cells; write was blocked because restore cannot be guaranteed");
            return new JsonObject
            {
                ["address"] = Convert.ToString(((dynamic)range).Address(false, false), CultureInfo.InvariantCulture),
                ["values"] = values,
                ["formulas"] = formulas,
                ["styles"] = CaptureRangeCellStyles(range),
            };
        }
        finally { RotHelper.ReleaseComReference(range); }
    }

    private static JsonArray CaptureRangeCellStyles(object range)
    {
        var rows = Convert.ToInt32(((dynamic)range).Rows.Count, CultureInfo.InvariantCulture);
        var cols = Convert.ToInt32(((dynamic)range).Columns.Count, CultureInfo.InvariantCulture);
        if ((long)rows * cols > MaxFormatSnapshotCells)
            throw new InvalidOperationException(
                $"format snapshot exceeds {MaxFormatSnapshotCells} cells; write was blocked because cell formatting could not be restored safely");
        var styles = new JsonArray();
        object? cells = null;
        try
        {
            cells = (object)((dynamic)range).Cells;
            for (var row = 1; row <= rows; row++)
            {
                var styleRow = new JsonArray();
                for (var col = 1; col <= cols; col++)
                {
                    object? cell = null;
                    try
                    {
                        cell = (object)((dynamic)cells).Item(row, col);
                        styleRow.Add(CaptureCellStyle(cell));
                    }
                    finally { RotHelper.ReleaseComReference(cell); }
                }
                styles.Add(styleRow);
            }
        }
        finally { RotHelper.ReleaseComReference(cells); }
        return styles;
    }

    private static int RestoreRangePayload(object sheet, string? sheetName, string address, JsonObject? payload,
        RestoreMismatchCollector mismatches)
    {
        if (payload is null || string.IsNullOrWhiteSpace(address)) return 0;
        object? range = null;
        var checkedItems = 0;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            var source = Json.GetArr(payload, "formulas") ?? Json.GetArr(payload, "values");
            if (source is { Count: > 0 } && source[0] is JsonArray first)
            {
                var rows = source.Count;
                var cols = first.Count;
                var data = new object?[rows, cols];
                for (var i = 0; i < rows; i++)
                    for (var j = 0; j < cols; j++)
                        data[i, j] = NodeToComValue(source[i]![j]);
                if (Json.GetArr(payload, "formulas") is not null) ((dynamic)range).Formula = data;
                else ((dynamic)range).Value2 = data;
                checkedItems += rows * cols;
            }

            var styleRows = Json.GetArr(payload, "styles") ?? new JsonArray();
            object? cells = null;
            try
            {
                cells = (object)((dynamic)range).Cells;
                for (var row = 0; row < styleRows.Count; row++)
                {
                    if (styleRows[row] is not JsonArray styleCols) continue;
                    for (var col = 0; col < styleCols.Count; col++)
                    {
                        if (styleCols[col] is not JsonObject style) continue;
                        object? cell = null;
                        try
                        {
                            cell = (object)((dynamic)cells).Item(row + 1, col + 1);
                            RestoreCellStyle(cell, style);
                            checkedItems++;
                            if (!CellStyleMatches(cell, style))
                                mismatches.Add($"{sheetName}!{address}: style restore mismatch");
                        }
                        finally { RotHelper.ReleaseComReference(cell); }
                    }
                }
            }
            finally { RotHelper.ReleaseComReference(cells); }
        }
        finally { RotHelper.ReleaseComReference(range); }
        return checkedItems;
    }

    private static string UsedStripAddress(object sheet, int start, int count, bool rows)
    {
        object? used = null;
        try
        {
            used = (object)((dynamic)sheet).UsedRange;
            var firstRow = Convert.ToInt32(((dynamic)used).Row, CultureInfo.InvariantCulture);
            var firstCol = Convert.ToInt32(((dynamic)used).Column, CultureInfo.InvariantCulture);
            var usedRows = Convert.ToInt32(((dynamic)used).Rows.Count, CultureInfo.InvariantCulture);
            var usedCols = Convert.ToInt32(((dynamic)used).Columns.Count, CultureInfo.InvariantCulture);
            if (rows)
                return $"{ColName(firstCol)}{start}:{ColName(firstCol + usedCols - 1)}{start + count - 1}";
            return $"{ColName(start)}{firstRow}:{ColName(start + count - 1)}{firstRow + usedRows - 1}";
        }
        catch
        {
            return rows ? $"A{start}:A{start + count - 1}" : $"{ColName(start)}1:{ColName(start + count - 1)}1";
        }
        finally { RotHelper.ReleaseComReference(used); }
    }

    private static JsonArray CaptureRowHeights(object sheet, int start, int count)
    {
        var result = new JsonArray();
        object? rows = null;
        try
        {
            rows = (object)((dynamic)sheet).Rows;
            for (var offset = 0; offset < count; offset++)
            {
                object? one = null;
                try
                {
                    one = (object)((dynamic)rows).Item(start + offset);
                    result.Add(new JsonObject
                    {
                        ["row"] = start + offset,
                        ["heightPoints"] = Convert.ToDouble(((dynamic)one).RowHeight, CultureInfo.InvariantCulture),
                    });
                }
                finally { RotHelper.ReleaseComReference(one); }
            }
        }
        finally { RotHelper.ReleaseComReference(rows); }
        return result;
    }

    private static JsonArray CaptureColumnWidths(object sheet, int start, int count)
    {
        var result = new JsonArray();
        object? columns = null;
        try
        {
            columns = (object)((dynamic)sheet).Columns;
            for (var offset = 0; offset < count; offset++)
            {
                object? one = null;
                try
                {
                    one = (object)((dynamic)columns).Item(start + offset);
                    result.Add(new JsonObject
                    {
                        ["column"] = start + offset,
                        ["widthChars"] = Convert.ToDouble(((dynamic)one).ColumnWidth, CultureInfo.InvariantCulture),
                    });
                }
                finally { RotHelper.ReleaseComReference(one); }
            }
        }
        finally { RotHelper.ReleaseComReference(columns); }
        return result;
    }

    private static int RestoreSizes(object sheet, JsonArray sizes, bool rows, RestoreMismatchCollector mismatches)
    {
        if (rows) return RestoreRowHeights(sheet, "", sizes, mismatches);
        return RestoreColumnWidths(sheet, "", sizes, mismatches);
    }

    private static object? FindWorksheet(object workbook, string? name)
    {
        if (string.IsNullOrWhiteSpace(name)) return null;
        object? sheets = null;
        try
        {
            sheets = (object)((dynamic)workbook).Worksheets;
            return (object)((dynamic)sheets).Item(name);
        }
        catch
        {
            return null;
        }
        finally { RotHelper.ReleaseComReference(sheets); }
    }

    private static object AfterLastWorksheet(object workbook)
    {
        object? sheets = null;
        try
        {
            sheets = (object)((dynamic)workbook).Worksheets;
            var count = Convert.ToInt32(((dynamic)sheets).Count, CultureInfo.InvariantCulture);
            return (object)((dynamic)sheets).Item(count);
        }
        finally { RotHelper.ReleaseComReference(sheets); }
    }

    private static JsonObject SheetTargetOp(string? sheetName) => new()
    {
        ["op"] = "restore",
        ["target"] = new JsonObject { ["sheet"] = sheetName },
    };

    private static bool SafeBool(Func<bool> read)
    {
        try { return read(); }
        catch { return false; }
    }

    private static double? SafeDouble(Func<double> read)
    {
        try { return read(); }
        catch { return null; }
    }

    private static int CompareCopiedFormats(object source, object dest, string? sheetName, string address,
        List<string> mismatches)
    {
        var sourceFormats = CaptureNumberFormats(source);
        var destFormats = CaptureNumberFormats(dest);
        var checkedItems = 1;
        if (sourceFormats is JsonArray sourceRows && destFormats is JsonArray destRows)
        {
            if (sourceRows.Count != destRows.Count)
            {
                mismatches.Add($"copy_range formats size mismatch {sheetName}!{address}");
                return checkedItems;
            }

            for (var r = 0; r < sourceRows.Count; r++)
            {
                if (sourceRows[r] is not JsonArray src || destRows[r] is not JsonArray dst) continue;
                for (var c = 0; c < src.Count && c < dst.Count; c++)
                {
                    checkedItems++;
                    if (!string.Equals(src[c]?.ToString(), dst[c]?.ToString(), StringComparison.Ordinal))
                        mismatches.Add($"copy_range format mismatch {sheetName}!{address}");
                }
            }
        }

        return checkedItems;
    }

    private static int RestoreDeletedSheetsFromBackup(object application, object workbook, JsonObject state,
        string? snapshotDir, RestoreMismatchCollector mismatches)
    {
        var recovery = Json.GetObj(state, "sheetRecovery");
        var deleted = Json.GetArr(state, "deletedSheets") ?? new JsonArray();
        if (deleted.Count == 0)
            return 0;
        if (recovery is null ||
            !string.Equals(Json.GetString(recovery, "mode"), ExcelDeleteSheetRecoveryContract.ModeWorkbookCopySheet, StringComparison.Ordinal))
        {
            mismatches.Add(
                "delete_sheet used-range values/numberFormat/tabColor replay is not a full restore; " +
                "refusing to recreate sheets without a current workbook-copy-sheet backup");
            return 0;
        }

        var backupPath = Json.GetString(recovery, "backupPath");
        var targetWorkbook = Json.GetString(recovery, "targetWorkbook") ?? Json.GetString(state, "documentRef");
        string? actualTarget = null;
        try { actualTarget = Convert.ToString(((dynamic)workbook).FullName, CultureInfo.InvariantCulture); }
        catch { actualTarget = null; }
        if (string.IsNullOrWhiteSpace(backupPath) || !File.Exists(backupPath))
        {
            mismatches.Add("delete_sheet recovery backup is missing");
            return 0;
        }

        if (!string.IsNullOrWhiteSpace(targetWorkbook) &&
            !string.Equals(actualTarget, targetWorkbook, StringComparison.OrdinalIgnoreCase))
        {
            mismatches.Add($"delete_sheet recovery target '{actualTarget}' does not match '{targetWorkbook}'");
            return 0;
        }

        object? backup = null;
        var checkedItems = 0;
        try
        {
            object? books = (object)((dynamic)application).Workbooks;
            try
            {
                backup = (object)((dynamic)books).Open(backupPath, 0, true);
            }
            finally { RotHelper.ReleaseComReference(books); }

            var backupName = Convert.ToString(((dynamic)backup).FullName, CultureInfo.InvariantCulture);
            if (!string.Equals(Path.GetFullPath(backupName ?? ""), Path.GetFullPath(backupPath), StringComparison.OrdinalIgnoreCase))
            {
                mismatches.Add($"delete_sheet recovery opened '{backupName}' instead of '{backupPath}'");
                return 0;
            }

            foreach (var node in deleted.OfType<JsonObject>())
            {
                var name = Json.GetString(node, "name");
                if (string.IsNullOrWhiteSpace(name)) continue;
                if (FindWorksheet(workbook, name) is object existing)
                {
                    RotHelper.ReleaseComReference(existing);
                    continue;
                }

                object? sourceSheet = null;
                try
                {
                    sourceSheet = FindWorksheet(backup, name);
                    if (sourceSheet is null)
                    {
                        mismatches.Add($"delete_sheet recovery backup is missing sheet '{name}'");
                        continue;
                    }

                    ((dynamic)sourceSheet).Copy(Type.Missing, AfterLastWorksheet(workbook));
                    if (!SheetExists(workbook, name))
                        mismatches.Add($"delete_sheet recovery did not restore sheet '{name}'");
                    else
                        checkedItems++;
                }
                finally { RotHelper.ReleaseComReference(sourceSheet); }
            }
        }
        catch (Exception ex)
        {
            mismatches.Add($"delete_sheet workbook-copy-sheet restore failed: {ex.Message}");
        }
        finally
        {
            if (backup is not null)
            {
                try { ((dynamic)backup).Close(false); }
                catch { /* already closed */ }
                RotHelper.ReleaseComReference(backup);
            }
        }

        return checkedItems;
    }

    private static int RestoreIntersectingNames(object workbook, JsonArray? names, RestoreMismatchCollector mismatches)
    {
        if (names is null || names.Count == 0) return 0;
        object? collection = null;
        var checkedItems = 0;
        try
        {
            collection = (object)((dynamic)workbook).Names;
            foreach (var node in names)
            {
                if (node is not JsonObject item) continue;
                var name = Json.GetString(item, "name");
                var refersTo = Json.GetString(item, "refersTo");
                if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(refersTo))
                    continue;
                object? existing = null;
                try
                {
                    try { existing = (object)((dynamic)collection).Item(name); }
                    catch { existing = null; }
                    if (existing is null)
                        existing = (object)((dynamic)collection).Add(name, refersTo);
                    else
                        ((dynamic)existing).RefersTo = refersTo;
                    if (item.ContainsKey("visible"))
                        ((dynamic)existing).Visible = Json.GetBool(item, "visible");
                    var comment = Json.GetString(item, "comment");
                    if (!string.IsNullOrWhiteSpace(comment))
                        ((dynamic)existing).Comment = comment;
                    checkedItems++;
                }
                catch (Exception ex)
                {
                    mismatches.Add($"delete restore name '{name}' failed: {ex.Message}");
                }
                finally { RotHelper.ReleaseComReference(existing); }
            }
        }
        finally { RotHelper.ReleaseComReference(collection); }
        return checkedItems;
    }

    private static int RestoreExternalFormulas(object workbook, JsonArray? formulas, RestoreMismatchCollector mismatches)
    {
        if (formulas is null || formulas.Count == 0) return 0;
        var checkedItems = 0;
        foreach (var node in formulas)
        {
            if (node is not JsonObject item) continue;
            var sheetName = Json.GetString(item, "sheet");
            var address = Json.GetString(item, "address");
            var formula = Json.GetString(item, "formula");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(address) ||
                string.IsNullOrWhiteSpace(formula))
                continue;
            object? sheet = null;
            object? range = null;
            try
            {
                sheet = GetExplicitTargetSheetReference(workbook, SheetTargetOp(sheetName));
                range = (object)((dynamic)sheet).Range(address);
                ((dynamic)range).Formula = formula;
                checkedItems++;
            }
            catch (Exception ex)
            {
                mismatches.Add($"delete restore formula {sheetName}!{address} failed: {ex.Message}");
            }
            finally
            {
                RotHelper.ReleaseComReference(range);
                RotHelper.ReleaseComReference(sheet);
            }
        }

        return checkedItems;
    }

    private static JsonArray CaptureIntersectingNames(object workbook, string sheetName)
    {
        var names = new JsonArray();
        object? collection = null;
        try
        {
            collection = (object)((dynamic)workbook).Names;
            var count = Convert.ToInt32(((dynamic)collection).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                object? item = null;
                try
                {
                    item = (object)((dynamic)collection).Item(index);
                    var refersTo = Convert.ToString(((dynamic)item).RefersTo, CultureInfo.InvariantCulture) ?? "";
                    var ownerWorkbook = TryReadWorkbookFileName(workbook);
                    var sheetOrder = ReadOrderedWorksheetNames(workbook);
                    ExcelDeleteSheetDependencyContract.RequireFaithfulCapture(refersTo, ownerWorkbook);
                    if (!ExcelDeleteSheetDependencyContract.NameReferencesDeletedSheet(
                            refersTo, sheetName, ownerWorkbook, sheetOrder))
                        continue;
                    names.Add(new JsonObject
                    {
                        ["name"] = Convert.ToString(((dynamic)item).Name, CultureInfo.InvariantCulture),
                        ["refersTo"] = refersTo,
                        ["visible"] = SafeBool(() => ((dynamic)item).Visible),
                        ["comment"] = Convert.ToString(((dynamic)item).Comment, CultureInfo.InvariantCulture),
                    });
                }
                catch (Exception ex)
                {
                    throw ExcelDeleteSheetDependencyContract.CaptureFailed(
                        $"name capture failed: {ex.Message}");
                }
                finally { RotHelper.ReleaseComReference(item); }
            }
        }
        catch (Exception ex)
        {
            if (ExcelDeleteSheetDependencyContract.IsUncaptured(ex))
                throw;
            throw ExcelDeleteSheetDependencyContract.CaptureFailed(
                $"name collection capture failed: {ex.Message}");
        }
        finally { RotHelper.ReleaseComReference(collection); }
        return names;
    }

    private static JsonArray CaptureExternalFormulas(object workbook, string sheetName)
    {
        var formulas = new JsonArray();
        var ownerWorkbook = TryReadWorkbookFileName(workbook);
        var sheetOrder = ReadOrderedWorksheetNames(workbook);
        object? sheets = null;
        try
        {
            sheets = (object)((dynamic)workbook).Worksheets;
            var count = Convert.ToInt32(((dynamic)sheets).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                object? sheet = null;
                object? used = null;
                try
                {
                    sheet = (object)((dynamic)sheets).Item(index);
                    var name = Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture) ?? "";
                    if (string.Equals(name, sheetName, StringComparison.OrdinalIgnoreCase))
                        continue;
                    used = (object)((dynamic)sheet).UsedRange;
                    object? cells = null;
                    try
                    {
                        try
                        {
                            cells = (object)((dynamic)used).SpecialCells(-4123); // xlCellTypeFormulas
                        }
                        catch (Exception ex) when (ExcelDeleteSheetDependencyContract.IsNoSpecialCells(ex))
                        {
                            continue;
                        }

                        foreach (dynamic cell in ((dynamic)cells).Cells)
                        {
                            object? owned = (object)cell;
                            try
                            {
                                var formula = Convert.ToString(cell.Formula, CultureInfo.InvariantCulture) ?? "";
                                ExcelDeleteSheetDependencyContract.RequireFaithfulCapture(formula, ownerWorkbook);
                                if (!ExcelDeleteSheetDependencyContract.FormulaReferencesDeletedSheet(
                                        formula, sheetName, ownerWorkbook, sheetOrder))
                                    continue;
                                formulas.Add(new JsonObject
                                {
                                    ["sheet"] = name,
                                    ["address"] = Convert.ToString(cell.Address(false, false), CultureInfo.InvariantCulture),
                                    ["formula"] = formula,
                                });
                            }
                            finally { RotHelper.ReleaseComReference(owned); }
                        }
                    }
                    catch (Exception ex)
                    {
                        throw ExcelDeleteSheetDependencyContract.CaptureFailed(
                            $"formula capture failed on '{name}': {ex.Message}");
                    }
                    finally { RotHelper.ReleaseComReference(cells); }
                }
                finally
                {
                    RotHelper.ReleaseComReference(used);
                    RotHelper.ReleaseComReference(sheet);
                }
            }
        }
        finally { RotHelper.ReleaseComReference(sheets); }
        return formulas;
    }

    private static int RestoreDeletedSheetDependencies(object workbook, JsonObject state,
        RestoreMismatchCollector mismatches)
    {
        var checkedItems = 0;
        foreach (var node in Json.GetArr(state, "deletedSheets") ?? new JsonArray())
        {
            if (node is not JsonObject deleted) continue;
            checkedItems += RestoreIntersectingNames(
                workbook, Json.GetArr(deleted, ExcelDeleteSheetDependencyContract.CaptureKeyNames), mismatches);
            checkedItems += RestoreExternalFormulas(
                workbook, Json.GetArr(deleted, ExcelDeleteSheetDependencyContract.CaptureKeyFormulas), mismatches);
            checkedItems += RestoreChartDependencies(
                workbook, Json.GetArr(deleted, ExcelDeleteSheetDependencyContract.CaptureKeyCharts), mismatches);
            checkedItems += RestorePivotDependencies(
                workbook, Json.GetArr(deleted, ExcelDeleteSheetDependencyContract.CaptureKeyPivots), mismatches);
            VerifyDeletedSheetDependencies(workbook, deleted, mismatches);
        }

        return checkedItems;
    }

    private static JsonArray CaptureChartDependencies(object workbook, string sheetName)
    {
        var charts = new JsonArray();
        var ownerWorkbook = TryReadWorkbookFileName(workbook);
        var sheetOrder = ReadOrderedWorksheetNames(workbook);
        object? sheets = null;
        try
        {
            sheets = (object)((dynamic)workbook).Worksheets;
            var count = Convert.ToInt32(((dynamic)sheets).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                object? sheet = null;
                try
                {
                    sheet = (object)((dynamic)sheets).Item(index);
                    var name = Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture) ?? "";
                    if (string.Equals(name, sheetName, StringComparison.OrdinalIgnoreCase))
                        continue;
                    object? chartObjects = null;
                    try
                    {
                        chartObjects = (object)((dynamic)sheet).ChartObjects();
                        var chartCount = Convert.ToInt32(((dynamic)chartObjects).Count, CultureInfo.InvariantCulture);
                        for (var chartIndex = 1; chartIndex <= chartCount; chartIndex++)
                        {
                            object? chartObject = null;
                            object? chart = null;
                            object? series = null;
                            try
                            {
                                chartObject = (object)((dynamic)chartObjects).Item(chartIndex);
                                chart = (object)((dynamic)chartObject).Chart;
                                series = (object)((dynamic)chart).SeriesCollection();
                                var seriesCount = Convert.ToInt32(((dynamic)series).Count, CultureInfo.InvariantCulture);
                                for (var seriesIndex = 1; seriesIndex <= seriesCount; seriesIndex++)
                                {
                                    object? item = null;
                                    try
                                    {
                                        item = (object)((dynamic)series).Item(seriesIndex);
                                        var formula = Convert.ToString(((dynamic)item).Formula, CultureInfo.InvariantCulture);
                                        ExcelDeleteSheetDependencyContract.RequireFaithfulCapture(formula, ownerWorkbook);
                                        if (!ExcelDeleteSheetDependencyContract.ChartFormulaReferencesDeletedSheet(
                                                formula, sheetName, ownerWorkbook, sheetOrder))
                                            continue;
                                        charts.Add(new JsonObject
                                        {
                                            ["sheet"] = name,
                                            ["chart"] = Convert.ToString(((dynamic)chartObject).Name, CultureInfo.InvariantCulture),
                                            ["series"] = seriesIndex,
                                            ["formula"] = formula,
                                        });
                                    }
                                    finally { RotHelper.ReleaseComReference(item); }
                                }
                            }
                            catch (Exception ex)
                            {
                                if (ExcelDeleteSheetDependencyContract.IsUncaptured(ex))
                                    throw;
                            }
                            finally
                            {
                                RotHelper.ReleaseComReference(series);
                                RotHelper.ReleaseComReference(chart);
                                RotHelper.ReleaseComReference(chartObject);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        if (ExcelDeleteSheetDependencyContract.IsUncaptured(ex))
                            throw;
                    }
                    finally { RotHelper.ReleaseComReference(chartObjects); }
                }
                finally { RotHelper.ReleaseComReference(sheet); }
            }
        }
        finally { RotHelper.ReleaseComReference(sheets); }
        return charts;
    }

    private static int RestoreChartDependencies(object workbook, JsonArray? charts, RestoreMismatchCollector mismatches)
    {
        if (charts is null || charts.Count == 0) return 0;
        var checkedItems = 0;
        foreach (var node in charts.OfType<JsonObject>())
        {
            var sheetName = Json.GetString(node, "sheet");
            var chartName = Json.GetString(node, "chart");
            var seriesIndex = Json.GetInt(node, "series") ?? 0;
            var formula = Json.GetString(node, "formula");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(chartName) ||
                seriesIndex < 1 || string.IsNullOrWhiteSpace(formula))
                continue;
            object? sheet = null;
            object? chartObject = null;
            object? chart = null;
            object? series = null;
            object? item = null;
            try
            {
                sheet = GetExplicitTargetSheetReference(workbook, SheetTargetOp(sheetName));
                chartObject = (object)((dynamic)sheet).ChartObjects(chartName);
                chart = (object)((dynamic)chartObject).Chart;
                series = (object)((dynamic)chart).SeriesCollection();
                item = (object)((dynamic)series).Item(seriesIndex);
                ((dynamic)item).Formula = formula;
                checkedItems++;
            }
            catch (Exception ex)
            {
                mismatches.Add($"delete restore chart {sheetName}!{chartName} series {seriesIndex} failed: {ex.Message}");
            }
            finally
            {
                RotHelper.ReleaseComReference(item);
                RotHelper.ReleaseComReference(series);
                RotHelper.ReleaseComReference(chart);
                RotHelper.ReleaseComReference(chartObject);
                RotHelper.ReleaseComReference(sheet);
            }
        }

        return checkedItems;
    }

    private static void VerifyDeletedSheetDependencies(object workbook, JsonObject deleted,
        RestoreMismatchCollector mismatches)
    {
        foreach (var node in Json.GetArr(deleted, ExcelDeleteSheetDependencyContract.CaptureKeyFormulas) ?? new JsonArray())
        {
            if (node is not JsonObject item) continue;
            var sheetName = Json.GetString(item, "sheet");
            var address = Json.GetString(item, "address");
            var expected = Json.GetString(item, "formula");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(address))
                continue;
            object? sheet = null;
            object? range = null;
            try
            {
                sheet = GetExplicitTargetSheetReference(workbook, SheetTargetOp(sheetName));
                range = (object)((dynamic)sheet).Range(address);
                var actual = Convert.ToString(((dynamic)range).Formula, CultureInfo.InvariantCulture);
                if (!ExcelDeleteSheetDependencyContract.RestoredDependencyMatches(expected, actual))
                    mismatches.Add($"delete_sheet formula {sheetName}!{address} restore mismatch: want '{expected}', got '{actual}'");
            }
            catch (Exception ex)
            {
                mismatches.Add($"delete_sheet formula {sheetName}!{address} verify failed: {ex.Message}");
            }
            finally
            {
                RotHelper.ReleaseComReference(range);
                RotHelper.ReleaseComReference(sheet);
            }
        }

        foreach (var node in Json.GetArr(deleted, ExcelDeleteSheetDependencyContract.CaptureKeyNames) ?? new JsonArray())
        {
            if (node is not JsonObject item) continue;
            var name = Json.GetString(item, "name");
            var expected = Json.GetString(item, "refersTo");
            if (string.IsNullOrWhiteSpace(name)) continue;
            object? collection = null;
            object? existing = null;
            try
            {
                collection = (object)((dynamic)workbook).Names;
                existing = (object)((dynamic)collection).Item(name);
                var actual = Convert.ToString(((dynamic)existing).RefersTo, CultureInfo.InvariantCulture);
                if (!ExcelDeleteSheetDependencyContract.RestoredDependencyMatches(expected, actual))
                    mismatches.Add($"delete_sheet name '{name}' restore mismatch: want '{expected}', got '{actual}'");
            }
            catch (Exception ex)
            {
                mismatches.Add($"delete_sheet name '{name}' verify failed: {ex.Message}");
            }
            finally
            {
                RotHelper.ReleaseComReference(existing);
                RotHelper.ReleaseComReference(collection);
            }
        }

        foreach (var node in Json.GetArr(deleted, ExcelDeleteSheetDependencyContract.CaptureKeyCharts) ?? new JsonArray())
        {
            if (node is not JsonObject item) continue;
            var sheetName = Json.GetString(item, "sheet");
            var chartName = Json.GetString(item, "chart");
            var seriesIndex = Json.GetInt(item, "series") ?? 0;
            var expected = Json.GetString(item, "formula");
            object? sheet = null;
            object? chartObject = null;
            object? chart = null;
            object? series = null;
            object? seriesItem = null;
            try
            {
                sheet = GetExplicitTargetSheetReference(workbook, SheetTargetOp(sheetName));
                chartObject = (object)((dynamic)sheet).ChartObjects(chartName);
                chart = (object)((dynamic)chartObject).Chart;
                series = (object)((dynamic)chart).SeriesCollection();
                seriesItem = (object)((dynamic)series).Item(seriesIndex);
                var actual = Convert.ToString(((dynamic)seriesItem).Formula, CultureInfo.InvariantCulture);
                if (!ExcelDeleteSheetDependencyContract.RestoredDependencyMatches(expected, actual))
                    mismatches.Add($"delete_sheet chart {sheetName}!{chartName} series {seriesIndex} restore mismatch: want '{expected}', got '{actual}'");
            }
            catch (Exception ex)
            {
                mismatches.Add($"delete_sheet chart {sheetName}!{chartName} verify failed: {ex.Message}");
            }
            finally
            {
                RotHelper.ReleaseComReference(seriesItem);
                RotHelper.ReleaseComReference(series);
                RotHelper.ReleaseComReference(chart);
                RotHelper.ReleaseComReference(chartObject);
                RotHelper.ReleaseComReference(sheet);
            }
        }

        foreach (var node in Json.GetArr(deleted, ExcelDeleteSheetDependencyContract.CaptureKeyPivots) ?? new JsonArray())
        {
            if (node is not JsonObject item) continue;
            var sheetName = Json.GetString(item, "sheet");
            var pivotName = Json.GetString(item, "name");
            var expected = Json.GetString(item, "source");
            object? sheet = null;
            object? pivot = null;
            object? cache = null;
            try
            {
                sheet = GetExplicitTargetSheetReference(workbook, SheetTargetOp(sheetName));
                pivot = string.IsNullOrWhiteSpace(pivotName) ? null : FindPivotTable(sheet, pivotName);
                if (pivot is null)
                {
                    mismatches.Add($"delete_sheet pivot {sheetName}!{pivotName} missing after restore");
                    continue;
                }

                try { cache = (object)((dynamic)pivot).PivotCache; }
                catch { cache = null; }
                var actual = cache is null ? null : TryComString(cache, "SourceData");
                if (ExcelDeleteSheetDependencyContract.IsBrokenRef(actual) ||
                    !ExcelDeleteSheetDependencyContract.RestoredDependencyMatches(expected, actual))
                    mismatches.Add($"delete_sheet pivot {sheetName}!{pivotName} source restore mismatch: want '{expected}', got '{actual}'");
            }
            catch (Exception ex)
            {
                mismatches.Add($"delete_sheet pivot {sheetName}!{pivotName} verify failed: {ex.Message}");
            }
            finally
            {
                RotHelper.ReleaseComReference(cache);
                RotHelper.ReleaseComReference(pivot);
                RotHelper.ReleaseComReference(sheet);
            }
        }
    }

    private static JsonArray CapturePivotDependencies(object workbook, string sheetName)
    {
        var pivots = new JsonArray();
        object? sheets = null;
        try
        {
            sheets = (object)((dynamic)workbook).Worksheets;
            var count = Convert.ToInt32(((dynamic)sheets).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                object? sheet = null;
                try
                {
                    sheet = (object)((dynamic)sheets).Item(index);
                    var name = Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture) ?? "";
                    if (string.Equals(name, sheetName, StringComparison.OrdinalIgnoreCase))
                        continue;
                    foreach (var pivot in EnumeratePivotTables(sheet))
                    {
                        try
                        {
                            var state = ReadPivotState(pivot, name);
                            var source = Json.GetString(state, "source");
                            var ownerWorkbook = TryReadWorkbookFileName(workbook);
                            var sheetOrder = ReadOrderedWorksheetNames(workbook);
                            ExcelDeleteSheetDependencyContract.RequireFaithfulCapture(source, ownerWorkbook);
                            if (!ExcelDeleteSheetDependencyContract.PivotSourceReferencesDeletedSheet(
                                    source, sheetName, ownerWorkbook, sheetOrder))
                                continue;
                            pivots.Add(new JsonObject
                            {
                                ["sheet"] = name,
                                ["name"] = Json.GetString(state, "name"),
                                ["source"] = source,
                                ["range"] = Json.GetString(state, "range"),
                            });
                        }
                        finally { RotHelper.ReleaseComReference(pivot); }
                    }
                }
                finally { RotHelper.ReleaseComReference(sheet); }
            }
        }
        finally { RotHelper.ReleaseComReference(sheets); }
        return pivots;
    }

    private static int RestorePivotDependencies(object workbook, JsonArray? pivots, RestoreMismatchCollector mismatches)
    {
        if (pivots is null || pivots.Count == 0) return 0;
        var checkedItems = 0;
        foreach (var node in pivots.OfType<JsonObject>())
        {
            var sheetName = Json.GetString(node, "sheet");
            var pivotName = Json.GetString(node, "name");
            var source = Json.GetString(node, "source");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(pivotName) ||
                string.IsNullOrWhiteSpace(source))
                continue;
            object? sheet = null;
            object? pivot = null;
            object? cache = null;
            object? caches = null;
            object? created = null;
            try
            {
                sheet = GetExplicitTargetSheetReference(workbook, SheetTargetOp(sheetName));
                pivot = FindPivotTable(sheet, pivotName);
                if (pivot is null)
                    throw new InvalidOperationException($"PivotTable '{pivotName}' was not found");
                try
                {
                    cache = (object)((dynamic)pivot).PivotCache;
                    ((dynamic)cache).SourceData = source;
                }
                catch
                {
                    caches = (object)((dynamic)workbook).PivotCaches();
                    created = (object)((dynamic)caches).Create(ExcelDataObjectCatalog.XlDatabase, source);
                    ((dynamic)pivot).ChangePivotCache(created);
                }

                try { ((dynamic)pivot).RefreshTable(); }
                catch { /* refresh is best-effort after source restore */ }
                checkedItems++;
            }
            catch (Exception ex)
            {
                mismatches.Add($"delete restore pivot {sheetName}!{pivotName} failed: {ex.Message}");
            }
            finally
            {
                RotHelper.ReleaseComReference(created);
                RotHelper.ReleaseComReference(caches);
                RotHelper.ReleaseComReference(cache);
                RotHelper.ReleaseComReference(pivot);
                RotHelper.ReleaseComReference(sheet);
            }
        }

        return checkedItems;
    }

    private static string? TryReadWorkbookFileName(object workbook)
    {
        try { return Convert.ToString(((dynamic)workbook).Name, CultureInfo.InvariantCulture); }
        catch { return null; }
    }
}
