using System.Globalization;
using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

/// <summary>
/// Operation-scoped snapshot/restore/readback for workbook-level and
/// advanced-object ops. Restore reuses the verified apply paths via
/// synthetic ops wherever the entry carries enough state.
/// </summary>
public sealed partial class ExcelAdapter
{
    // ------------------------------------------------------------------
    // Capture
    // ------------------------------------------------------------------

    private static JsonObject CaptureLinkEntry(object workbook, JsonObject op, string opName)
    {
        var mode = opName == ExcelDataOperationsContract.ChangeLinkSource ? "change"
            : opName == ExcelDataOperationsContract.BreakExternalLink ? "break" : "update";
        var entry = new JsonObject
        {
            ["kind"] = "externalLinks",
            ["mode"] = mode,
            ["sources"] = new JsonArray(
                ReadExcelLinks(workbook)
                    .Select(node => JsonValue.Create(Json.GetString(node, "source") ?? "")!).ToArray()),
        };
        if (mode == "change")
        {
            entry["oldSource"] = ResolveUniqueLink(workbook, Json.GetString(op, "source")!);
            entry["newSource"] = Json.GetString(op, "newSource");
        }
        if (mode == "break")
        {
            var resolved = ResolveUniqueLink(workbook, Json.GetString(op, "source")!);
            entry["oldSource"] = resolved;
            var dependents = ScanLinkDependents(workbook, ExcelWorkbookOpsContract.LeafOf(resolved));
            if (dependents.Count >= ExcelWorkbookOpsContract.MaxLinkDependents)
                throw new InvalidOperationException(
                    $"[EXCEL_LINK_DEPENDENTS_TOO_LARGE] dependents reach the {ExcelWorkbookOpsContract.MaxLinkDependents} snapshot bound");
            var array = new JsonArray();
            foreach (var dep in dependents)
            {
                array.Add(new JsonObject
                {
                    ["sheet"] = dep.Sheet,
                    ["address"] = dep.Address,
                    ["formula"] = dep.Formula,
                });
            }
            entry["dependents"] = array;
        }
        return entry;
    }

    private static JsonObject CaptureCalculationEntry(object workbook)
    {
        object? application = null;
        try
        {
            application = GetWorkbookApplication(workbook);
            return new JsonObject
            {
                ["kind"] = "calculationMode",
                ["mode"] = ReadCalculationMode(application),
            };
        }
        finally { RotHelper.ReleaseComReference(application); }
    }

    private static JsonObject CaptureWorkbookProtectionEntry(object workbook)
    {
        var entry = new JsonObject { ["kind"] = "workbookProtection" };
        foreach (var (key, value) in ReadWorkbookProtection(workbook))
            entry[key] = value?.DeepClone();
        return entry;
    }

    private static JsonObject CaptureSplitEntry(object workbook, JsonObject op)
    {
        using var sheet = BindSheet(workbook, op);
        return new JsonObject
        {
            ["kind"] = "splitPanes",
            ["sheet"] = sheet.SheetName,
            ["state"] = CaptureFreezePanes(sheet.Sheet),
        };
    }

    private static JsonObject CaptureSparklineEntry(object workbook, JsonObject op, string opName)
    {
        using var location = BindRange(workbook, op, "location");
        var group = FindSparklineGroup(location.Sheet, location.Address, out var state, location.SheetName);
        try
        {
            if (opName == ExcelDataOperationsContract.CreateSparkline)
            {
                if (group is not null)
                    throw new InvalidOperationException(
                        $"[EXCEL_SPARKLINE_EXISTS] a sparkline group already covers {location.SheetName}!{location.Address}");
                return new JsonObject
                {
                    ["kind"] = "sparkline",
                    ["sheet"] = location.SheetName,
                    ["location"] = location.Address,
                    ["existed"] = false,
                };
            }
            if (group is null || state is null)
                throw new InvalidOperationException(
                    $"[EXCEL_SPARKLINE_NOT_FOUND] no sparkline group covers {location.SheetName}!{location.Address}");
            return new JsonObject
            {
                ["kind"] = "sparkline",
                ["sheet"] = location.SheetName,
                ["location"] = location.Address,
                ["existed"] = true,
                ["state"] = state.DeepClone(),
            };
        }
        finally { RotHelper.ReleaseComReference(group); }
    }

    private static JsonObject CaptureSlicerEntry(object workbook, JsonObject op, string opName)
    {
        if (opName == ExcelDataOperationsContract.CreateSlicer)
        {
            using var sheet = BindSheet(workbook, op);
            return new JsonObject
            {
                ["kind"] = "slicer",
                ["sheet"] = sheet.SheetName,
                ["name"] = Json.GetString(op, "name"),
                ["existed"] = false,
            };
        }
        var name = Json.GetString(op, "name")!;
        object? owningSheet = null;
        var slicer = FindSlicer(workbook, name, out owningSheet);
        try
        {
            if (slicer is null)
                throw new InvalidOperationException($"[EXCEL_SLICER_NOT_FOUND] slicer '{name}' was not found");
            object? cache = null;
            try
            {
                try { cache = (object)((dynamic)slicer).SlicerCache; }
                catch { cache = null; }
                var cacheState = cache is null ? new JsonObject() : ReadSlicerCacheState(cache);
                object? parent = null;
                JsonObject? position = null;
                string? sheetName = null;
                try
                {
                    parent = (object)((dynamic)slicer).Parent;
                    sheetName = ReadComString(parent, "Name");
                    position = new JsonObject
                    {
                        ["top"] = Js(TryComDouble(slicer, "Top")),
                        ["left"] = Js(TryComDouble(slicer, "Left")),
                        ["width"] = Js(TryComDouble(slicer, "Width")),
                        ["height"] = Js(TryComDouble(slicer, "Height")),
                    };
                }
                finally { RotHelper.ReleaseComReference(parent); }
                var field = Json.GetString(op, "source") is string explicitSource &&
                            Json.GetString(op, "field") is string explicitField &&
                            !string.IsNullOrWhiteSpace(explicitSource) && !string.IsNullOrWhiteSpace(explicitField)
                    ? explicitField
                    : TryComString(slicer, "Caption");
                return new JsonObject
                {
                    ["kind"] = "slicer",
                    ["sheet"] = sheetName,
                    ["name"] = ReadComString(slicer, "Name"),
                    ["existed"] = true,
                    ["state"] = new JsonObject
                    {
                        ["cacheName"] = Json.GetString(cacheState, "name"),
                        ["sourceKind"] = Json.GetString(cacheState, "sourceKind"),
                        ["sourceName"] = Json.GetString(cacheState, "sourceName"),
                        ["field"] = field,
                        ["caption"] = TryComString(slicer, "Caption"),
                        ["position"] = position,
                    },
                };
            }
            finally { RotHelper.ReleaseComReference(cache); }
        }
        finally
        {
            RotHelper.ReleaseComReference(slicer);
            RotHelper.ReleaseComReference(owningSheet);
        }
    }

    private static JsonObject CaptureCellStyleEntry(object workbook, JsonObject op)
    {
        using var range = BindRange(workbook, op, "range");
        return new JsonObject
        {
            ["kind"] = "cellStyle",
            ["sheet"] = range.SheetName,
            ["range"] = range.Address,
            ["styles"] = ReadStyleNameGrid(range.Sheet, range.Address),
        };
    }

    private static JsonObject CaptureValueSurfaceEntry(object workbook, JsonObject op, string opName)
    {
        var kind = opName == ExcelDataOperationsContract.PasteSpecial ? "pasteSpecial" : "freezeValues";
        if (kind == "pasteSpecial")
        {
            using var source = BindRange(workbook, op, "sourceRange", independentSource: true);
            using var dest = BindRange(workbook, op, "destination");
            var plan = PlanPasteSpecial(op, source.Address, dest.Address);
            var entry = new JsonObject
            {
                ["kind"] = kind,
                ["sheet"] = dest.SheetName,
                ["range"] = plan.WriteAddress,
            };
            MergeSurface(entry, CaptureRangeSurface(dest.Sheet, plan.WriteAddress));
            return entry;
        }
        using var range = BindRange(workbook, op, "range");
        var single = new JsonObject
        {
            ["kind"] = kind,
            ["sheet"] = range.SheetName,
            ["range"] = range.Address,
        };
        MergeSurface(single, CaptureRangeSurface(range.Sheet, range.Address));
        return single;
    }

    private static JsonObject CaptureGoalSeekEntry(object workbook, JsonObject op)
    {
        using var changing = BindRange(workbook, op, "cell");
        using var goal = BindRange(workbook, op, "goalCell");
        if (!string.Equals(changing.SheetName, goal.SheetName, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("[EXCEL_GOAL_SEEK_SHEET] cell and goalCell must be on the same worksheet");
        var entry = new JsonObject
        {
            ["kind"] = "goalSeek",
            ["sheet"] = goal.SheetName,
            ["range"] = goal.Address,
            ["changing"] = ReadGoalChangingCell(changing.Sheet, changing.Address),
        };
        MergeSurface(entry, CaptureRangeSurface(goal.Sheet, goal.Address));
        return entry;
    }

    private static JsonObject ReadGoalChangingCell(object sheet, string address)
    {
        object? range = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            object? value = null;
            object? formula = null;
            try
            {
                value = (object)((dynamic)range).Value2;
                formula = (object)((dynamic)range).Formula;
                return new JsonObject
                {
                    ["address"] = address,
                    ["value"] = ToJsonValue(value),
                    ["formula"] = formula as string,
                };
            }
            finally
            {
                RotHelper.ReleaseComReference(formula);
                RotHelper.ReleaseComReference(value);
            }
        }
        finally { RotHelper.ReleaseComReference(range); }
    }

    // ------------------------------------------------------------------
    // Restore (synthetic-op reuse)
    // ------------------------------------------------------------------

    private delegate void DataOpApply(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells);

    private static int RestoreWithSyntheticOp(object workbook, JsonObject synthetic,
        DataOpApply apply,
        RestoreMismatchCollector mismatches, string label)
    {
        var execution = new ApplyExecution();
        var opMismatches = new List<string>();
        var checkedCells = 0;
        try
        {
            apply(workbook, synthetic, execution, opMismatches, ref checkedCells);
        }
        catch (Exception ex)
        {
            mismatches.Add($"{label}: restore failed: {ex.Message}");
            return 0;
        }
        foreach (var mismatch in opMismatches)
            mismatches.Add($"{label}: restore mismatch: {mismatch}");
        return checkedCells;
    }

    private static int RestoreLinkEntry(object workbook, JsonObject entry, RestoreMismatchCollector mismatches)
    {
        var mode = Json.GetString(entry, "mode") ?? "update";
        if (mode == "update") return 0;
        if (mode == "change")
        {
            var current = Json.GetString(entry, "newSource");
            var previous = Json.GetString(entry, "oldSource");
            if (string.IsNullOrWhiteSpace(current) || string.IsNullOrWhiteSpace(previous))
            {
                mismatches.Add("link change snapshot is missing oldSource/newSource");
                return 0;
            }
            var synthetic = new JsonObject
            {
                ["op"] = ExcelDataOperationsContract.ChangeLinkSource,
                ["source"] = current,
                ["newSource"] = previous,
            };
            return RestoreWithSyntheticOp(workbook, synthetic, ApplyChangeLinkSource, mismatches, "link change");
        }
        var dependents = Json.GetArr(entry, "dependents") ?? new JsonArray();
        var restored = 0;
        foreach (var node in dependents.OfType<JsonObject>())
        {
            var sheetName = Json.GetString(node, "sheet");
            var address = Json.GetString(node, "address");
            var formula = Json.GetString(node, "formula");
            if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(address) ||
                string.IsNullOrWhiteSpace(formula))
            {
                mismatches.Add("link break snapshot has an incomplete dependent");
                continue;
            }
            object? sheet = null;
            object? range = null;
            try
            {
                sheet = GetExplicitTargetSheetReference(workbook, SheetOp(sheetName!));
                range = (object)((dynamic)sheet).Range(address);
                ((dynamic)range).Formula = formula;
                restored++;
            }
            catch (Exception ex)
            {
                mismatches.Add($"{sheetName}!{address}: relink failed: {ex.Message}");
            }
            finally
            {
                RotHelper.ReleaseComReference(range);
                RotHelper.ReleaseComReference(sheet);
            }
        }
        object? application = null;
        try
        {
            application = GetWorkbookApplication(workbook);
            try { RecalculateApplication(application, "full"); } catch { /* best effort */ }
        }
        finally { RotHelper.ReleaseComReference(application); }
        return restored;
    }

    private static int RestoreCalculationEntry(object workbook, JsonObject entry, RestoreMismatchCollector mismatches)
    {
        if (Json.GetInt(entry, "mode") is not int mode)
        {
            mismatches.Add("calculation snapshot is missing mode");
            return 0;
        }
        var token = mode == ExcelWorkbookOpsContract.XlCalculationManual ? "manual"
            : mode == ExcelWorkbookOpsContract.XlCalculationAutomatic ? "automatic"
            : mode == ExcelWorkbookOpsContract.XlCalculationSemiautomatic ? "semiautomatic" : null;
        if (token is null)
        {
            mismatches.Add($"calculation snapshot mode {mode} is unknown");
            return 0;
        }
        return RestoreWithSyntheticOp(workbook,
            new JsonObject
            {
                ["op"] = ExcelDataOperationsContract.SetCalculationMode,
                ["mode"] = token,
            }, ApplySetCalculationMode, mismatches, "calculation mode");
    }

    private static int RestoreWorkbookProtectionEntry(object workbook, JsonObject entry, RestoreMismatchCollector mismatches)
    {
        try { ((dynamic)workbook).Unprotect(); }
        catch (Exception ex)
        {
            mismatches.Add($"workbook protection restore failed: {ex.Message}");
            return 0;
        }
        if (Json.GetBool(entry, "structure") || Json.GetBool(entry, "windows"))
        {
            // Protect(Password, Structure, Windows): password stays missing (Type.Missing).
            try { ((dynamic)workbook).Protect(Type.Missing, Json.GetBool(entry, "structure"), Json.GetBool(entry, "windows")); }
            catch (Exception ex)
            {
                mismatches.Add($"workbook protection restore failed: {ex.Message}");
                return 0;
            }
        }
        return 1;
    }

    private static int RestoreSplitEntry(object workbook, JsonObject entry, RestoreMismatchCollector mismatches)
    {
        var sheetName = Json.GetString(entry, "sheet");
        var state = Json.GetObj(entry, "state");
        if (string.IsNullOrWhiteSpace(sheetName) || state is null)
        {
            mismatches.Add("split snapshot is missing sheet/state");
            return 0;
        }
        var synthetic = new JsonObject
        {
            ["op"] = ExcelDataOperationsContract.SetSplitPanes,
            ["target"] = new JsonObject { ["sheet"] = sheetName },
            ["splitRows"] = Json.GetInt(state, "splitRow") ?? 0,
            ["splitColumns"] = Json.GetInt(state, "splitColumn") ?? 0,
        };
        if (!Json.GetBool(state, "frozen") &&
            (Json.GetInt(state, "splitRow") ?? 0) == 0 && (Json.GetInt(state, "splitColumn") ?? 0) == 0)
            synthetic["remove"] = true;
        else if (Json.GetBool(state, "frozen"))
            synthetic["freeze"] = true;
        return RestoreWithSyntheticOp(workbook, synthetic, ApplySetSplitPanes, mismatches, "split panes");
    }

    private static bool SparklineGroupExists(object workbook, string sheetName, string location)
    {
        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, SheetOp(sheetName));
            var group = FindSparklineGroup(sheet, location, out _, sheetName);
            try { return group is not null; }
            finally { RotHelper.ReleaseComReference(group); }
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static int RestoreSparklineEntry(object workbook, JsonObject entry, RestoreMismatchCollector mismatches)
    {
        var sheetName = Json.GetString(entry, "sheet");
        var location = Json.GetString(entry, "location");
        if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(location))
        {
            mismatches.Add("sparkline snapshot is missing sheet/location");
            return 0;
        }
        var target = new JsonObject
        {
            ["op"] = ExcelDataOperationsContract.DeleteSparkline,
            ["target"] = new JsonObject { ["sheet"] = sheetName },
            ["location"] = location,
        };
        if (!Json.GetBool(entry, "existed"))
        {
            if (!SparklineGroupExists(workbook, sheetName!, location!)) return 0;
            return RestoreWithSyntheticOp(workbook, target, ApplyDeleteSparkline, mismatches, "sparkline");
        }
        var removed = RestoreWithSyntheticOp(workbook, target, ApplyDeleteSparkline, mismatches, "sparkline");
        var state = Json.GetObj(entry, "state");
        if (state is null)
        {
            mismatches.Add("sparkline snapshot is missing state for recreation");
            return removed;
        }
        var create = new JsonObject
        {
            ["op"] = ExcelDataOperationsContract.CreateSparkline,
            ["target"] = new JsonObject { ["sheet"] = sheetName },
            ["location"] = location,
            ["sourceData"] = Json.GetString(state, "sourceData"),
            ["type"] = Json.GetString(state, "typeName") == "column" ? "column" : "line",
        };
        foreach (var key in new[] { "markers", "showHigh", "showLow", "showNegative", "showFirst", "showLast" })
        {
            if (state[key] is JsonValue flag && flag.TryGetValue<bool>(out var on)) create[key] = on;
        }
        return removed + RestoreWithSyntheticOp(workbook, create, ApplyCreateSparkline, mismatches, "sparkline");
    }

    private static int RestoreSlicerEntry(object workbook, JsonObject entry, RestoreMismatchCollector mismatches)
    {
        var name = Json.GetString(entry, "name");
        if (!string.IsNullOrWhiteSpace(name))
        {
            var sheetName = Json.GetString(entry, "sheet") ?? "";
            var existing = FindSlicer(workbook, name, out var owningSheet);
            try
            {
                if (existing is null)
                {
                    if (!Json.GetBool(entry, "existed")) return 0;
                    mismatches.Add($"slicer '{name}' is missing during restore");
                    return 0;
                }
            }
            finally
            {
                RotHelper.ReleaseComReference(existing);
                RotHelper.ReleaseComReference(owningSheet);
            }
            RestoreWithSyntheticOp(workbook,
                new JsonObject
                {
                    ["op"] = ExcelDataOperationsContract.DeleteSlicer,
                    ["target"] = new JsonObject { ["sheet"] = sheetName },
                    ["name"] = name,
                }, ApplyDeleteSlicer, mismatches, "slicer");
        }
        if (!Json.GetBool(entry, "existed")) return 1;
        var state = Json.GetObj(entry, "state");
        var sourceName = Json.GetString(state, "sourceName");
        var field = Json.GetString(state, "field");
        if (state is null || string.IsNullOrWhiteSpace(sourceName) || string.IsNullOrWhiteSpace(field))
        {
            mismatches.Add("slicer snapshot is missing source/field for recreation");
            return 0;
        }
        var create = new JsonObject
        {
            ["op"] = ExcelDataOperationsContract.CreateSlicer,
            ["target"] = new JsonObject { ["sheet"] = Json.GetString(entry, "sheet") ?? "" },
            ["source"] = sourceName,
            ["field"] = field,
            ["name"] = name ?? "",
        };
        if (Json.GetString(state, "caption") is string caption && !string.IsNullOrWhiteSpace(caption))
            create["caption"] = caption;
        if (Json.GetObj(state, "position") is JsonObject position)
            create["position"] = position.DeepClone();
        return RestoreWithSyntheticOp(workbook, create, ApplyCreateSlicer, mismatches, "slicer");
    }

    private static int RestoreCellStyleEntry(object workbook, JsonObject entry, RestoreMismatchCollector mismatches)
    {
        var sheetName = Json.GetString(entry, "sheet");
        var range = Json.GetString(entry, "range");
        var styles = Json.GetArr(entry, "styles");
        if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(range) || styles is null)
        {
            mismatches.Add("cellStyle snapshot is missing sheet/range/styles");
            return 0;
        }
        if (!ExcelDataOperationsContract.TryParseA1(range, out var box, allowUnion: false))
        {
            mismatches.Add("cellStyle snapshot range is invalid");
            return 0;
        }
        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, SheetOp(sheetName!));
            var restored = 0;
            var index = 0;
            for (var r = 0; r < box.Rows; r++)
            {
                var runStart = -1;
                string? runStyle = null;
                for (var c = 0; c <= box.Columns; c++)
                {
                    string? current = null;
                    if (c < box.Columns && index < styles.Count)
                    {
                        var node = styles[index];
                        current = node is JsonValue value && value.TryGetValue<string>(out var text) ? text : null;
                    }
                    if (!string.Equals(current, runStyle, StringComparison.Ordinal))
                    {
                        if (runStart >= 0 && !string.IsNullOrWhiteSpace(runStyle))
                        {
                            WriteStyleRun(sheet, box, r, runStart, c - 1, runStyle!);
                            restored += c - runStart;
                        }
                        runStart = c;
                        runStyle = current;
                    }
                    index++;
                }
            }
            return restored;
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static void WriteStyleRun(object sheet, ExcelA1Box box, int rowOffset, int colStart, int colEnd, string styleName)
    {
        object? range = null;
        try
        {
            var address = colStart == colEnd
                ? $"{ColName(box.Column + colStart)}{box.Row + rowOffset}"
                : $"{ColName(box.Column + colStart)}{box.Row + rowOffset}:{ColName(box.Column + colEnd)}{box.Row + rowOffset}";
            range = (object)((dynamic)sheet).Range(address);
            ((dynamic)range).Style = styleName;
        }
        finally { RotHelper.ReleaseComReference(range); }
    }

    private static int RestoreGoalSeekEntry(object workbook, JsonObject entry, RestoreMismatchCollector mismatches)
    {
        var sheetName = Json.GetString(entry, "sheet");
        var range = Json.GetString(entry, "range");
        if (string.IsNullOrWhiteSpace(sheetName) || string.IsNullOrWhiteSpace(range))
        {
            mismatches.Add("goalSeek snapshot is missing sheet/range");
            return 0;
        }
        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, SheetOp(sheetName!));
            RestoreRangeSurface(sheet, entry);
            var changing = Json.GetObj(entry, "changing");
            var address = changing is null ? null : Json.GetString(changing, "address");
            if (!string.IsNullOrWhiteSpace(address))
            {
                var formula = changing is null ? null : Json.GetString(changing, "formula");
                object? cell = null;
                try
                {
                    cell = (object)((dynamic)sheet).Range(address);
                    if (!string.IsNullOrWhiteSpace(formula))
                        ((dynamic)cell).Formula = formula;
                    else if (changing is not null && changing["value"] is JsonNode value)
                        ((dynamic)cell).Value2 = NodeToComValue(value) ?? "";
                }
                finally { RotHelper.ReleaseComReference(cell); }
            }
            return 2;
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    // ------------------------------------------------------------------
    // ReadActual
    // ------------------------------------------------------------------

    private static JsonObject ReadActualLinkEntry(object workbook, JsonObject entry)
    {
        var actual = new JsonObject
        {
            ["kind"] = "externalLinks",
            ["mode"] = Json.GetString(entry, "mode"),
            ["sources"] = new JsonArray(
                ReadExcelLinks(workbook)
                    .Select(node => JsonValue.Create(Json.GetString(node, "source") ?? "")!).ToArray()),
        };
        if ((Json.GetString(entry, "mode") ?? "") == "break" &&
            Json.GetString(entry, "oldSource") is string old && !string.IsNullOrWhiteSpace(old))
        {
            var dependents = ScanLinkDependents(workbook, ExcelWorkbookOpsContract.LeafOf(old));
            var array = new JsonArray();
            foreach (var dep in dependents)
            {
                array.Add(new JsonObject
                {
                    ["sheet"] = dep.Sheet,
                    ["address"] = dep.Address,
                    ["formula"] = dep.Formula,
                });
            }
            actual["dependents"] = array;
        }
        return actual;
    }

    private static JsonObject ReadActualCalculation(object workbook, JsonObject entry)
    {
        object? application = null;
        try
        {
            application = GetWorkbookApplication(workbook);
            return new JsonObject
            {
                ["kind"] = "calculationMode",
                ["mode"] = ReadCalculationMode(application),
            };
        }
        finally { RotHelper.ReleaseComReference(application); }
    }

    private static JsonObject ReadActualWorkbookProtection(object workbook, JsonObject entry)
    {
        var actual = new JsonObject { ["kind"] = "workbookProtection" };
        foreach (var (key, value) in ReadWorkbookProtection(workbook))
            actual[key] = value?.DeepClone();
        return actual;
    }

    private static JsonObject ReadActualSplit(object workbook, JsonObject entry)
    {
        var sheetName = Json.GetString(entry, "sheet") ?? "";
        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, SheetOp(sheetName));
            return new JsonObject
            {
                ["kind"] = "splitPanes",
                ["sheet"] = sheetName,
                ["state"] = CaptureFreezePanes(sheet),
            };
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static JsonObject ReadActualSparkline(object workbook, JsonObject entry)
    {
        var sheetName = Json.GetString(entry, "sheet") ?? "";
        var location = Json.GetString(entry, "location") ?? "";
        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, SheetOp(sheetName));
            var group = FindSparklineGroup(sheet, location, out var state, sheetName);
            try
            {
                return new JsonObject
                {
                    ["kind"] = "sparkline",
                    ["existed"] = Json.GetBool(entry, "existed"),
                    ["present"] = group is not null,
                    ["state"] = state?.DeepClone(),
                };
            }
            finally { RotHelper.ReleaseComReference(group); }
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static JsonObject ReadActualSlicer(object workbook, JsonObject entry)
    {
        var name = Json.GetString(entry, "name") ?? "";
        object? owningSheet = null;
        var slicer = FindSlicer(workbook, name, out owningSheet);
        try
        {
            JsonObject? state = null;
            if (slicer is not null)
            {
                object? cache = null;
                try
                {
                    try { cache = (object)((dynamic)slicer).SlicerCache; }
                    catch { cache = null; }
                    object? parent = null;
                    try
                    {
                        parent = (object)((dynamic)slicer).Parent;
                        state = new JsonObject
                        {
                            ["cacheName"] = cache is null ? null : ReadComString(cache, "Name"),
                            ["caption"] = TryComString(slicer, "Caption"),
                            ["sheet"] = TryComString(parent, "Name"),
                            ["position"] = new JsonObject
                            {
                                ["top"] = Js(TryComDouble(slicer, "Top")),
                                ["left"] = Js(TryComDouble(slicer, "Left")),
                                ["width"] = Js(TryComDouble(slicer, "Width")),
                                ["height"] = Js(TryComDouble(slicer, "Height")),
                            },
                        };
                    }
                    finally { RotHelper.ReleaseComReference(parent); }
                }
                finally { RotHelper.ReleaseComReference(cache); }
            }
            return new JsonObject
            {
                ["kind"] = "slicer",
                ["existed"] = Json.GetBool(entry, "existed"),
                ["present"] = slicer is not null,
                ["state"] = state?.DeepClone(),
            };
        }
        finally
        {
            RotHelper.ReleaseComReference(slicer);
            RotHelper.ReleaseComReference(owningSheet);
        }
    }

    private static JsonObject ReadActualCellStyle(object workbook, JsonObject entry)
    {
        var sheetName = Json.GetString(entry, "sheet") ?? "";
        var range = Json.GetString(entry, "range") ?? "";
        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, SheetOp(sheetName));
            return new JsonObject
            {
                ["kind"] = "cellStyle",
                ["sheet"] = sheetName,
                ["range"] = range,
                ["styles"] = ReadStyleNameGrid(sheet, range),
            };
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static JsonObject ReadActualGoalSeek(object workbook, JsonObject entry)
    {
        var sheetName = Json.GetString(entry, "sheet") ?? "";
        var range = Json.GetString(entry, "range") ?? "";
        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, SheetOp(sheetName));
            var actual = new JsonObject
            {
                ["kind"] = "goalSeek",
                ["sheet"] = sheetName,
                ["range"] = range,
            };
            MergeSurface(actual, CaptureRangeSurface(sheet, range));
            var changing = Json.GetObj(entry, "changing");
            var address = Json.GetString(changing, "address");
            if (changing is not null && !string.IsNullOrWhiteSpace(address))
                actual["changing"] = ReadGoalChangingCell(sheet, address!);
            return actual;
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }
}
