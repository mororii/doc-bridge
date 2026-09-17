using System.Globalization;
using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class ExcelAdapter
{
    private static void PreviewSheetLayoutOperation(object workbook, JsonObject op, ApplyPreview preview)
    {
        var name = Json.GetString(op, "op")!;
        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, op);
            var sheetName = Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture) ?? "";
            switch (name)
            {
                case "set_row_heights":
                    preview.Affected.Add(new AffectedRef("rows", $"{sheetName}!{DescribeSizeSpan(op, rows: true)}"));
                    preview.Diff.Add(new DiffEntry
                    {
                        Ref = $"{sheetName}:rowHeights",
                        Before = CaptureRowHeightSummary(sheet, op),
                        After = op["rows"]?.DeepClone(),
                    });
                    return;
                case "set_column_widths":
                    preview.Affected.Add(new AffectedRef("cols", $"{sheetName}!{DescribeSizeSpan(op, rows: false)}"));
                    preview.Diff.Add(new DiffEntry
                    {
                        Ref = $"{sheetName}:columnWidths",
                        Before = CaptureColumnWidthSummary(sheet, op),
                        After = op["columns"]?.DeepClone(),
                    });
                    return;
                case "freeze_panes":
                    var before = CaptureFreezePanes(sheet);
                    var after = PlannedFreezePanes(op);
                    preview.Affected.Add(new AffectedRef("window", $"{sheetName}:freezePanes"));
                    preview.Diff.Add(new DiffEntry
                    {
                        Ref = $"{sheetName}:freezePanes",
                        Before = before,
                        After = after,
                    });
                    return;
                case "set_page_setup":
                    var pageErrors = new List<string>();
                    ExcelPageSetupContract.TryNormalize(Json.GetObj(op, "page"), 1, pageErrors, out var canonical);
                    if (pageErrors.Count > 0) throw new InvalidOperationException(string.Join("; ", pageErrors));
                    preview.Affected.Add(new AffectedRef("page", sheetName));
                    preview.Diff.Add(new DiffEntry
                    {
                        Ref = $"{sheetName}:pageSetup",
                        Before = CapturePageSetup(sheet),
                        After = canonical,
                    });
                    if (Json.GetBool(canonical, "fitIgnored"))
                        preview.Warnings.Add("set_page_setup: scale is applied and fitToWidth/fitToHeight are ignored");
                    return;
                case "set_view":
                    preview.Affected.Add(new AffectedRef("window", $"{sheetName}:view"));
                    preview.Diff.Add(new DiffEntry
                    {
                        Ref = $"{sheetName}:view",
                        Before = CaptureSheetView(workbook, sheet),
                        After = PlannedSheetView(op),
                    });
                    return;
            }
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static void ApplySheetLayoutOperation(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        var name = Json.GetString(op, "op")!;
        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, op);
            var sheetName = Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture) ?? "";
            switch (name)
            {
                case "set_row_heights":
                    ApplyRowHeights(sheet, sheetName, op, execution, mismatches, ref checkedItems);
                    return;
                case "set_column_widths":
                    ApplyColumnWidths(sheet, sheetName, op, execution, mismatches, ref checkedItems);
                    return;
                case "freeze_panes":
                    ApplyFreezePanes(workbook, sheet, sheetName, op, execution, mismatches, ref checkedItems);
                    return;
                case "set_page_setup":
                    ApplyPageSetup(sheet, sheetName, op, execution, mismatches, ref checkedItems);
                    return;
                case "set_view":
                    ApplySheetView(workbook, sheet, sheetName, op, execution, mismatches, ref checkedItems);
                    return;
            }
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static JsonObject CaptureSheetLayoutState(object workbook, IReadOnlyList<JsonObject> ops, string? documentRef)
    {
        var entries = new JsonArray();
        foreach (var op in ops)
        {
            object? sheet = null;
            try
            {
                sheet = GetExplicitTargetSheetReference(workbook, op);
                var sheetName = Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture) ?? "";
                var name = Json.GetString(op, "op")!;
                var entry = new JsonObject
                {
                    ["op"] = name,
                    ["sheet"] = sheetName,
                };
                switch (name)
                {
                    case "set_row_heights":
                        entry["rows"] = CaptureRowHeightEntries(sheet, op);
                        break;
                    case "set_column_widths":
                        entry["columns"] = CaptureColumnWidthEntries(sheet, op);
                        break;
                    case "freeze_panes":
                        entry["freezePanes"] = CaptureFreezePanes(sheet);
                        entry["window"] = CaptureWindowState(workbook);
                        break;
                    case "set_page_setup":
                        entry["pageSetup"] = CapturePageSetup(sheet);
                        break;
                    case "set_view":
                        entry["view"] = CaptureSheetView(workbook, sheet);
                        entry["window"] = CaptureWindowState(workbook);
                        break;
                }
                entries.Add(entry);
            }
            finally { RotHelper.ReleaseComReference(sheet); }
        }

        return new JsonObject
        {
            ["snapshotVersion"] = ExcelLayoutSnapshotVersion,
            ["restoreMode"] = SheetLayoutRestoreMode,
            ["documentRef"] = documentRef,
            ["originalActiveSheet"] = ReadActiveSheetName(workbook),
            ["entries"] = entries,
        };
    }

    private static JsonObject RestoreSheetLayoutState(object workbook, JsonObject state)
    {
        var mismatches = new RestoreMismatchCollector();
        if (Json.GetInt(state, "snapshotVersion") != ExcelLayoutSnapshotVersion)
        {
            mismatches.Add("unsupported Excel sheet-layout snapshot version");
            return BuildRestoreResult(false, 0, 0, SheetLayoutRestoreMode, mismatches);
        }

        var checkedItems = 0;
        foreach (var node in Json.GetArr(state, "entries") ?? new JsonArray())
        {
            if (node is not JsonObject entry) continue;
            object? sheet = null;
            try
            {
                var sheetName = Json.GetString(entry, "sheet") ?? "";
                sheet = GetExplicitTargetSheetReference(workbook, new JsonObject
                {
                    ["op"] = "restore_sheet_layout",
                    ["target"] = new JsonObject { ["sheet"] = sheetName },
                });
                var kind = Json.GetString(entry, "op");
                if (kind == "set_row_heights")
                    checkedItems += RestoreRowHeights(sheet, sheetName, Json.GetArr(entry, "rows") ?? new JsonArray(), mismatches);
                else if (kind == "set_column_widths")
                    checkedItems += RestoreColumnWidths(sheet, sheetName, Json.GetArr(entry, "columns") ?? new JsonArray(), mismatches);
                else if (kind == "freeze_panes")
                    checkedItems += RestoreFreezePanes(workbook, sheet, sheetName, Json.GetObj(entry, "freezePanes"), mismatches);
                else if (kind == "set_page_setup")
                    checkedItems += RestorePageSetup(sheet, sheetName, Json.GetObj(entry, "pageSetup"), mismatches);
                else if (kind == "set_view")
                    checkedItems += RestoreSheetView(workbook, sheet, sheetName, Json.GetObj(entry, "view"), mismatches);
            }
            catch (Exception ex) { mismatches.Add($"sheet-layout restore failed: {ex.Message}"); }
            finally { RotHelper.ReleaseComReference(sheet); }
        }

        var active = Json.GetString(state, "originalActiveSheet");
        if (!string.IsNullOrWhiteSpace(active)) ActivateWorksheet(workbook, active, mismatches);
        return BuildRestoreResult(mismatches.Count == 0, 0, checkedItems, SheetLayoutRestoreMode, mismatches);
    }

    private static void ApplyRowHeights(object sheet, string sheetName, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        foreach (var node in Json.GetArr(op, "rows") ?? new JsonArray())
        {
            if (node is not JsonObject item) continue;
            var row = Json.GetInt(item, "row")!.Value;
            var count = Json.GetInt(item, "count") ?? 1;
            var autoFit = Json.GetBool(item, "autoFit");
            object? rows = null;
            try
            {
                rows = (object)((dynamic)sheet).Rows;
                object? target = null;
                try
                {
                    target = (object)((dynamic)rows)[$"{row}:{row + count - 1}"];
                    if (autoFit) ((dynamic)target).AutoFit();
                    else ((dynamic)target).RowHeight = item["heightPoints"]!.GetValue<double>();
                }
                finally { RotHelper.ReleaseComReference(target); }

                for (var offset = 0; offset < count; offset++)
                {
                    object? one = null;
                    try
                    {
                        one = (object)((dynamic)rows).Item(row + offset);
                        checkedItems++;
                        var actual = Convert.ToDouble((object)((dynamic)one).RowHeight, CultureInfo.InvariantCulture);
                        if (autoFit)
                        {
                            execution.Diff.Add(new DiffEntry
                            {
                                Ref = $"{sheetName}!row {row + offset}",
                                After = JsonValue.Create(actual),
                            });
                            continue;
                        }

                        var requested = item["heightPoints"]!.GetValue<double>();
                        var mapping = ReadOwnerRowHeightPixelMapping(sheet);
                        if (!ExcelRowHeightContract.ObservedMatchesNearestPixelStep(
                                requested, actual, mapping, out double normalized, out _))
                            mismatches.Add(ExcelRowHeightContract.MismatchMessage(
                                sheetName, row + offset, requested, actual, mapping));
                        execution.Diff.Add(new DiffEntry
                        {
                            Ref = $"{sheetName}!row {row + offset}",
                            Before = JsonValue.Create(requested),
                            After = JsonValue.Create(normalized),
                        });
                    }
                    finally { RotHelper.ReleaseComReference(one); }
                }
            }
            finally { RotHelper.ReleaseComReference(rows); }
        }

        execution.Affected.Add(new AffectedRef("rows", $"{sheetName}!{DescribeSizeSpan(op, rows: true)}"));
    }

    private static void ApplyColumnWidths(object sheet, string sheetName, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        foreach (var node in Json.GetArr(op, "columns") ?? new JsonArray())
        {
            if (node is not JsonObject item) continue;
            var col = ParseColumnNumber(item["col"]);
            var count = Json.GetInt(item, "count") ?? 1;
            var autoFit = Json.GetBool(item, "autoFit");
            object? columns = null;
            try
            {
                columns = (object)((dynamic)sheet).Columns;
                object? target = null;
                try
                {
                    target = (object)((dynamic)columns)[$"{ColName(col)}:{ColName(col + count - 1)}"];
                    if (autoFit) ((dynamic)target).AutoFit();
                    else ((dynamic)target).ColumnWidth = item["widthChars"]!.GetValue<double>();
                }
                finally { RotHelper.ReleaseComReference(target); }

                for (var offset = 0; offset < count; offset++)
                {
                    object? one = null;
                    try
                    {
                        one = (object)((dynamic)columns).Item(col + offset);
                        checkedItems++;
                        if (!autoFit &&
                            Math.Abs(Convert.ToDouble(((dynamic)one).ColumnWidth, CultureInfo.InvariantCulture) -
                                     item["widthChars"]!.GetValue<double>()) > 0.05)
                            mismatches.Add($"{sheetName}!col {ColName(col + offset)}: widthChars readback mismatch");
                    }
                    finally { RotHelper.ReleaseComReference(one); }
                }
            }
            finally { RotHelper.ReleaseComReference(columns); }
        }

        execution.Affected.Add(new AffectedRef("cols", $"{sheetName}!{DescribeSizeSpan(op, rows: false)}"));
    }

    private static void ApplyFreezePanes(object workbook, object sheet, string sheetName, JsonObject op,
        ApplyExecution execution, List<string> mismatches, ref int checkedItems)
    {
        var window = CaptureWindowState(workbook);
        try
        {
            object? sheetWindow = null;
            try
            {
                if (!TryGetWorkbookWindow(workbook, out sheetWindow, out var windowError) || sheetWindow is null)
                    throw new InvalidOperationException(windowError ?? "workbook window is unavailable");
                var visible = Convert.ToInt32(((dynamic)sheet).Visible, CultureInfo.InvariantCulture);
                if (visible != XlSheetVisible)
                    throw new InvalidOperationException($"freeze_panes cannot activate hidden sheet '{sheetName}'");
                ActivateWorksheet(workbook, sheetName, new RestoreMismatchCollector());
                if (Json.GetBool(op, "unfreeze"))
                {
                    ((dynamic)sheetWindow).FreezePanes = false;
                    ((dynamic)sheetWindow).SplitRow = 0;
                    ((dynamic)sheetWindow).SplitColumn = 0;
                }
                else
                {
                    PlannedFreezeSplit(op, out var splitRow, out var splitCol);
                    ((dynamic)sheetWindow).FreezePanes = false;
                    ((dynamic)sheetWindow).SplitRow = splitRow;
                    ((dynamic)sheetWindow).SplitColumn = splitCol;
                    ((dynamic)sheetWindow).FreezePanes = true;
                }
            }
            finally
            {
                RotHelper.ReleaseComReference(sheetWindow);
            }

            var actual = CaptureFreezePanes(sheet);
            var expected = PlannedFreezePanes(op);
            checkedItems++;
            if (Json.GetBool(actual, "readable") == false)
                mismatches.Add($"{sheetName}: freezePanes read failed: {Json.GetString(actual, "error")}");
            else if (!FreezeEquals(actual, expected))
                mismatches.Add($"{sheetName}: freezePanes readback mismatch");
            execution.Affected.Add(new AffectedRef("window", $"{sheetName}:freezePanes"));
        }
        finally
        {
            RestoreWindowState(workbook, window, new RestoreMismatchCollector());
        }
    }

    private static void ApplyPageSetup(object sheet, string sheetName, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        var errors = new List<string>();
        if (!ExcelPageSetupContract.TryNormalize(Json.GetObj(op, "page"), 1, errors, out var page))
            throw new InvalidOperationException(string.Join("; ", errors));
        WritePageSetup(sheet, page);
        var actual = CapturePageSetup(sheet);
        checkedItems++;
        if (!PageSetupMatches(actual, page))
            mismatches.Add($"{sheetName}: pageSetup readback mismatch: {string.Join("; ", ExcelPageSetupContract.DescribeMismatches(actual, page))}");
        execution.Affected.Add(new AffectedRef("page", sheetName));
    }

    private static JsonArray CaptureRowHeightEntries(object sheet, JsonObject op)
    {
        var result = new JsonArray();
        object? rows = null;
        try
        {
            rows = (object)((dynamic)sheet).Rows;
            foreach (var node in Json.GetArr(op, "rows") ?? new JsonArray())
            {
                if (node is not JsonObject item) continue;
                var row = Json.GetInt(item, "row")!.Value;
                var count = Json.GetInt(item, "count") ?? 1;
                for (var offset = 0; offset < count; offset++)
                {
                    object? one = null;
                    try
                    {
                        one = (object)((dynamic)rows).Item(row + offset);
                        result.Add(new JsonObject
                        {
                            ["row"] = row + offset,
                            ["heightPoints"] = Convert.ToDouble(((dynamic)one).RowHeight, CultureInfo.InvariantCulture),
                        });
                    }
                    finally { RotHelper.ReleaseComReference(one); }
                }
            }
        }
        finally { RotHelper.ReleaseComReference(rows); }
        return result;
    }

    private static JsonArray CaptureColumnWidthEntries(object sheet, JsonObject op)
    {
        var result = new JsonArray();
        object? columns = null;
        try
        {
            columns = (object)((dynamic)sheet).Columns;
            foreach (var node in Json.GetArr(op, "columns") ?? new JsonArray())
            {
                if (node is not JsonObject item) continue;
                var col = ParseColumnNumber(item["col"]);
                var count = Json.GetInt(item, "count") ?? 1;
                for (var offset = 0; offset < count; offset++)
                {
                    object? one = null;
                    try
                    {
                        one = (object)((dynamic)columns).Item(col + offset);
                        result.Add(new JsonObject
                        {
                            ["col"] = ColName(col + offset),
                            ["column"] = col + offset,
                            ["widthChars"] = Convert.ToDouble(((dynamic)one).ColumnWidth, CultureInfo.InvariantCulture),
                        });
                    }
                    finally { RotHelper.ReleaseComReference(one); }
                }
            }
        }
        finally { RotHelper.ReleaseComReference(columns); }
        return result;
    }

    private static int RestoreRowHeights(object sheet, string sheetName, JsonArray rows, RestoreMismatchCollector mismatches)
    {
        object? collection = null;
        var checkedItems = 0;
        try
        {
            collection = (object)((dynamic)sheet).Rows;
            foreach (var node in rows)
            {
                if (node is not JsonObject item || Json.GetInt(item, "row") is not int row) continue;
                object? one = null;
                try
                {
                    one = (object)((dynamic)collection).Item(row);
                    var expected = item["heightPoints"]!.GetValue<double>();
                    ((dynamic)one).RowHeight = expected;
                    checkedItems++;
                    var actual = Convert.ToDouble((object)((dynamic)one).RowHeight, CultureInfo.InvariantCulture);
                    var mapping = ReadOwnerRowHeightPixelMapping(sheet);
                    if (!ExcelRowHeightContract.ObservedMatchesNearestPixelStep(
                            expected, actual, mapping, out _, out _))
                        mismatches.Add(ExcelRowHeightContract.MismatchMessage(sheetName, row, expected, actual, mapping));
                }
                finally { RotHelper.ReleaseComReference(one); }
            }
        }
        finally { RotHelper.ReleaseComReference(collection); }
        return checkedItems;
    }

    private static int RestoreColumnWidths(object sheet, string sheetName, JsonArray columns, RestoreMismatchCollector mismatches)
    {
        object? collection = null;
        var checkedItems = 0;
        try
        {
            collection = (object)((dynamic)sheet).Columns;
            foreach (var node in columns)
            {
                if (node is not JsonObject item || Json.GetInt(item, "column") is not int col) continue;
                object? one = null;
                try
                {
                    one = (object)((dynamic)collection).Item(col);
                    var expected = item["widthChars"]!.GetValue<double>();
                    ((dynamic)one).ColumnWidth = expected;
                    checkedItems++;
                    if (Math.Abs(Convert.ToDouble(((dynamic)one).ColumnWidth, CultureInfo.InvariantCulture) - expected) > 0.05)
                        mismatches.Add($"{sheetName}!col {ColName(col)}: width restore mismatch");
                }
                finally { RotHelper.ReleaseComReference(one); }
            }
        }
        finally { RotHelper.ReleaseComReference(collection); }
        return checkedItems;
    }

    private static JsonObject CaptureFreezePanes(object sheet)
    {
        object? workbook = null;
        object? window = null;
        var original = "";
        try
        {
            workbook = (object)((dynamic)sheet).Parent;
            original = ReadActiveSheetName(workbook);
            var sheetName = Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture) ?? "";
            var visible = Convert.ToInt32(((dynamic)sheet).Visible, CultureInfo.InvariantCulture);
            if (visible != XlSheetVisible)
            {
                return new JsonObject
                {
                    ["readable"] = false,
                    ["error"] = $"freeze panes cannot be read from hidden sheet '{sheetName}'",
                    ["sheet"] = sheetName,
                    ["hidden"] = true,
                };
            }

            if (!TryGetWorkbookWindow(workbook, out window, out var windowError) || window is null)
            {
                return new JsonObject
                {
                    ["readable"] = false,
                    ["error"] = windowError ?? "workbook window is unavailable",
                    ["sheet"] = sheetName,
                };
            }

            ActivateWorksheet(workbook, sheetName, new RestoreMismatchCollector());
            var splitRow = Convert.ToInt32(((dynamic)window).SplitRow, CultureInfo.InvariantCulture);
            var splitCol = Convert.ToInt32(((dynamic)window).SplitColumn, CultureInfo.InvariantCulture);
            var frozen = Convert.ToBoolean(((dynamic)window).FreezePanes, CultureInfo.InvariantCulture);
            return new JsonObject
            {
                ["readable"] = true,
                ["frozen"] = frozen,
                ["splitRow"] = splitRow,
                ["splitColumn"] = splitCol,
                ["topLeftCell"] = splitRow > 0 || splitCol > 0
                    ? $"{ColName(Math.Max(1, splitCol + 1))}{Math.Max(1, splitRow + 1)}"
                    : "",
                ["sheet"] = sheetName,
            };
        }
        catch (Exception ex)
        {
            return new JsonObject
            {
                ["readable"] = false,
                ["error"] = ex.Message,
                ["frozen"] = null,
                ["splitRow"] = null,
                ["splitColumn"] = null,
                ["topLeftCell"] = "",
            };
        }
        finally
        {
            RotHelper.ReleaseComReference(window);
            if (workbook is not null && !string.IsNullOrWhiteSpace(original))
                ActivateWorksheet(workbook, original, new RestoreMismatchCollector());
            RotHelper.ReleaseComReference(workbook);
        }
    }

    private static JsonObject PlannedFreezePanes(JsonObject op)
    {
        if (Json.GetBool(op, "unfreeze"))
            return new JsonObject { ["frozen"] = false, ["splitRow"] = 0, ["splitColumn"] = 0, ["topLeftCell"] = "" };
        PlannedFreezeSplit(op, out var splitRow, out var splitCol);
        return new JsonObject
        {
            ["frozen"] = splitRow > 0 || splitCol > 0,
            ["splitRow"] = splitRow,
            ["splitColumn"] = splitCol,
            ["topLeftCell"] = $"{ColName(Math.Max(1, splitCol + 1))}{Math.Max(1, splitRow + 1)}",
        };
    }

    private static void PlannedFreezeSplit(JsonObject op, out int splitRow, out int splitCol)
    {
        var cell = Json.GetString(op, "cell");
        if (!string.IsNullOrWhiteSpace(cell) && ExcelA1Box.TryParseCell(cell, out var row, out var column))
        {
            splitRow = Math.Max(0, row - 1);
            splitCol = Math.Max(0, column - 1);
            return;
        }

        splitRow = Json.GetInt(op, "rows") ?? 0;
        splitCol = Json.GetInt(op, "columns") ?? 0;
    }

    private static bool FreezeEquals(JsonObject actual, JsonObject expected) =>
        (!actual.ContainsKey("readable") || Json.GetBool(actual, "readable")) &&
        !actual.ContainsKey("error") &&
        Json.GetBool(actual, "frozen") == Json.GetBool(expected, "frozen") &&
        Json.GetInt(actual, "splitRow") == Json.GetInt(expected, "splitRow") &&
        Json.GetInt(actual, "splitColumn") == Json.GetInt(expected, "splitColumn");

    private static int RestoreFreezePanes(object workbook, object sheet, string sheetName, JsonObject? freeze,
        RestoreMismatchCollector mismatches)
    {
        if (freeze is not null && Json.GetBool(freeze, "readable") == false)
        {
            mismatches.Add($"{sheetName}: freezePanes snapshot was unreadable ({Json.GetString(freeze, "error")})");
            return 0;
        }

        var window = CaptureWindowState(workbook);
        try
        {
            object? sheetWindow = null;
            try
            {
                if (!TryGetWorkbookWindow(workbook, out sheetWindow, out var windowError) || sheetWindow is null)
                {
                    mismatches.Add($"{sheetName}: freezePanes restore failed: {windowError}");
                    return 0;
                }

                ActivateWorksheet(workbook, sheetName, mismatches);
                var splitRow = Json.GetInt(freeze, "splitRow") ?? 0;
                var splitCol = Json.GetInt(freeze, "splitColumn") ?? 0;
                ((dynamic)sheetWindow).FreezePanes = false;
                ((dynamic)sheetWindow).SplitRow = splitRow;
                ((dynamic)sheetWindow).SplitColumn = splitCol;
                ((dynamic)sheetWindow).FreezePanes = Json.GetBool(freeze, "frozen");
            }
            finally
            {
                RotHelper.ReleaseComReference(sheetWindow);
            }

            var actual = CaptureFreezePanes(sheet);
            if (freeze is not null && !FreezeEquals(actual, freeze))
                mismatches.Add($"{sheetName}: freezePanes restore mismatch");
            return 1;
        }
        finally { RestoreWindowState(workbook, window, mismatches); }
    }

    private static bool TryGetWorkbookWindow(object workbook, out object? window, out string? error)
    {
        window = null;
        error = null;
        object? windows = null;
        try
        {
            windows = (object)((dynamic)workbook).Windows;
            var count = Convert.ToInt32(((dynamic)windows).Count, CultureInfo.InvariantCulture);
            if (count < 1)
            {
                error = "workbook has no windows";
                return false;
            }

            window = (object)((dynamic)windows).Item(1);
            return true;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return false;
        }
        finally { RotHelper.ReleaseComReference(windows); }
    }

    private static JsonObject CaptureWindowState(object workbook)
    {
        object? window = null;
        object? selection = null;
        try
        {
            if (!TryGetWorkbookWindow(workbook, out window, out var error) || window is null)
                return new JsonObject { ["activeSheet"] = ReadActiveSheetName(workbook), ["error"] = error };
            try { selection = (object)((dynamic)window).Selection; } catch { }
            return new JsonObject
            {
                ["activeSheet"] = ReadActiveSheetName(workbook),
                ["scrollRow"] = SafeInt(() => ((dynamic)window).ScrollRow),
                ["scrollColumn"] = SafeInt(() => ((dynamic)window).ScrollColumn),
                ["selection"] = selection is null
                    ? null
                    : Convert.ToString(((dynamic)selection).Address(false, false), CultureInfo.InvariantCulture),
            };
        }
        catch (Exception ex)
        {
            return new JsonObject { ["activeSheet"] = ReadActiveSheetName(workbook), ["error"] = ex.Message };
        }
        finally
        {
            RotHelper.ReleaseComReference(selection);
            RotHelper.ReleaseComReference(window);
        }
    }

    private static void RestoreWindowState(object workbook, JsonObject state, RestoreMismatchCollector mismatches)
    {
        var sheetName = Json.GetString(state, "activeSheet");
        if (!string.IsNullOrWhiteSpace(sheetName)) ActivateWorksheet(workbook, sheetName, mismatches);
        object? window = null;
        object? sheet = null;
        try
        {
            if (!TryGetWorkbookWindow(workbook, out window, out _) || window is null)
                return;
            dynamic liveWindow = window;
            if (Json.GetInt(state, "scrollRow") is int scrollRow && scrollRow > 0)
                liveWindow.ScrollRow = scrollRow;
            if (Json.GetInt(state, "scrollColumn") is int scrollCol && scrollCol > 0)
                liveWindow.ScrollColumn = scrollCol;
            var selection = Json.GetString(state, "selection");
            if (!string.IsNullOrWhiteSpace(selection) && !string.IsNullOrWhiteSpace(sheetName))
            {
                sheet = GetExplicitTargetSheetReference(workbook, new JsonObject
                {
                    ["op"] = "restore_window",
                    ["target"] = new JsonObject { ["sheet"] = sheetName },
                });
                ((dynamic)sheet).Range(selection).Select();
            }
        }
        catch
        {
            // Window restore is best-effort; freeze/page state is the verified payload.
        }
        finally
        {
            RotHelper.ReleaseComReference(sheet);
            RotHelper.ReleaseComReference(window);
        }
    }

    private static JsonObject CapturePageSetup(object sheet)
    {
        object? page = null;
        try
        {
            page = (object)((dynamic)sheet).PageSetup;
            dynamic setup = page;
            var zoom = SafeGet(() => (object?)setup.Zoom);
            var zoomNumber = zoom is bool ? (int?)null : SafeInt(() => Convert.ToInt32(zoom, CultureInfo.InvariantCulture));
            return new JsonObject
            {
                ["paperSize"] = SafeInt(() => (int)setup.PaperSize),
                ["paperSizeName"] = ExcelPageSetupContract.PaperName(SafeInt(() => (int)setup.PaperSize)),
                ["orientation"] = SafeInt(() => (int)setup.Orientation) == ExcelPageSetupContract.XlLandscape
                    ? "landscape"
                    : "portrait",
                ["orientationValue"] = SafeInt(() => (int)setup.Orientation),
                ["zoomMode"] = zoom is bool && (bool)zoom == false
                    ? ExcelPageSetupContract.ModeFit
                    : ExcelPageSetupContract.ModeScale,
                ["scale"] = zoomNumber,
                ["fitToWidth"] = SafeInt(() => (int)setup.FitToPagesWide),
                ["fitToHeight"] = SafeInt(() => (int)setup.FitToPagesTall),
                ["printArea"] = Convert.ToString(setup.PrintArea, CultureInfo.InvariantCulture) ?? "",
                ["printTitleRows"] = Convert.ToString(setup.PrintTitleRows, CultureInfo.InvariantCulture) ?? "",
                ["printTitleColumns"] = Convert.ToString(setup.PrintTitleColumns, CultureInfo.InvariantCulture) ?? "",
                ["leftMarginMm"] = ExcelPageSetupContract.PointsToMillimetres(Convert.ToDouble(setup.LeftMargin, CultureInfo.InvariantCulture)),
                ["rightMarginMm"] = ExcelPageSetupContract.PointsToMillimetres(Convert.ToDouble(setup.RightMargin, CultureInfo.InvariantCulture)),
                ["topMarginMm"] = ExcelPageSetupContract.PointsToMillimetres(Convert.ToDouble(setup.TopMargin, CultureInfo.InvariantCulture)),
                ["bottomMarginMm"] = ExcelPageSetupContract.PointsToMillimetres(Convert.ToDouble(setup.BottomMargin, CultureInfo.InvariantCulture)),
                ["headerMarginMm"] = ExcelPageSetupContract.PointsToMillimetres(Convert.ToDouble(setup.HeaderMargin, CultureInfo.InvariantCulture)),
                ["footerMarginMm"] = ExcelPageSetupContract.PointsToMillimetres(Convert.ToDouble(setup.FooterMargin, CultureInfo.InvariantCulture)),
                ["leftMarginPoints"] = Convert.ToDouble(setup.LeftMargin, CultureInfo.InvariantCulture),
                ["rightMarginPoints"] = Convert.ToDouble(setup.RightMargin, CultureInfo.InvariantCulture),
                ["topMarginPoints"] = Convert.ToDouble(setup.TopMargin, CultureInfo.InvariantCulture),
                ["bottomMarginPoints"] = Convert.ToDouble(setup.BottomMargin, CultureInfo.InvariantCulture),
                ["headerMarginPoints"] = Convert.ToDouble(setup.HeaderMargin, CultureInfo.InvariantCulture),
                ["footerMarginPoints"] = Convert.ToDouble(setup.FooterMargin, CultureInfo.InvariantCulture),
                ["leftHeader"] = Convert.ToString(setup.LeftHeader, CultureInfo.InvariantCulture) ?? "",
                ["centerHeader"] = Convert.ToString(setup.CenterHeader, CultureInfo.InvariantCulture) ?? "",
                ["rightHeader"] = Convert.ToString(setup.RightHeader, CultureInfo.InvariantCulture) ?? "",
                ["leftFooter"] = Convert.ToString(setup.LeftFooter, CultureInfo.InvariantCulture) ?? "",
                ["centerFooter"] = Convert.ToString(setup.CenterFooter, CultureInfo.InvariantCulture) ?? "",
                ["rightFooter"] = Convert.ToString(setup.RightFooter, CultureInfo.InvariantCulture) ?? "",
            };
        }
        catch (Exception ex)
        {
            return new JsonObject { ["error"] = ex.Message };
        }
        finally { RotHelper.ReleaseComReference(page); }
    }

    private static void WritePageSetup(object sheet, JsonObject page)
    {
        object? setupObject = null;
        try
        {
            setupObject = (object)((dynamic)sheet).PageSetup;
            dynamic setup = setupObject;
            ExcelPageSetupContract.EnsurePointKeys(page);
            if (page.ContainsKey("paperSize")) setup.PaperSize = Json.GetInt(page, "paperSize")!.Value;
            if (page.ContainsKey("orientationValue")) setup.Orientation = Json.GetInt(page, "orientationValue")!.Value;
            if (string.Equals(Json.GetString(page, "zoomMode"), ExcelPageSetupContract.ModeScale, StringComparison.Ordinal))
            {
                setup.Zoom = Json.GetInt(page, "scale")!.Value;
            }
            else if (string.Equals(Json.GetString(page, "zoomMode"), ExcelPageSetupContract.ModeFit, StringComparison.Ordinal))
            {
                setup.Zoom = false;
                if (page.ContainsKey("fitToWidth")) setup.FitToPagesWide = Json.GetInt(page, "fitToWidth")!.Value;
                if (page.ContainsKey("fitToHeight")) setup.FitToPagesTall = Json.GetInt(page, "fitToHeight")!.Value;
            }

            AssignIfPresent(page, "printArea", value => setup.PrintArea = value);
            AssignIfPresent(page, "printTitleRows", value => setup.PrintTitleRows = value);
            AssignIfPresent(page, "printTitleColumns", value => setup.PrintTitleColumns = value);
            AssignIfPresent(page, "leftHeader", value => setup.LeftHeader = value);
            AssignIfPresent(page, "centerHeader", value => setup.CenterHeader = value);
            AssignIfPresent(page, "rightHeader", value => setup.RightHeader = value);
            AssignIfPresent(page, "leftFooter", value => setup.LeftFooter = value);
            AssignIfPresent(page, "centerFooter", value => setup.CenterFooter = value);
            AssignIfPresent(page, "rightFooter", value => setup.RightFooter = value);
            if (page.ContainsKey("leftMarginPoints")) setup.LeftMargin = page["leftMarginPoints"]!.GetValue<double>();
            if (page.ContainsKey("rightMarginPoints")) setup.RightMargin = page["rightMarginPoints"]!.GetValue<double>();
            if (page.ContainsKey("topMarginPoints")) setup.TopMargin = page["topMarginPoints"]!.GetValue<double>();
            if (page.ContainsKey("bottomMarginPoints")) setup.BottomMargin = page["bottomMarginPoints"]!.GetValue<double>();
            if (page.ContainsKey("headerMarginPoints")) setup.HeaderMargin = page["headerMarginPoints"]!.GetValue<double>();
            if (page.ContainsKey("footerMarginPoints")) setup.FooterMargin = page["footerMarginPoints"]!.GetValue<double>();
        }
        finally { RotHelper.ReleaseComReference(setupObject); }
    }

    private static bool PageSetupMatches(JsonObject actual, JsonObject expected) =>
        ExcelPageSetupContract.MatchesRequested(actual, expected);

    private static int RestorePageSetup(object sheet, string sheetName, JsonObject? page, RestoreMismatchCollector mismatches)
    {
        if (page is null)
        {
            mismatches.Add($"{sheetName}: pageSetup snapshot missing");
            return 0;
        }

        WritePageSetup(sheet, page);
        var actual = CapturePageSetup(sheet);
        if (!PageSetupMatches(actual, page))
            mismatches.Add($"{sheetName}: pageSetup restore mismatch: {string.Join("; ", ExcelPageSetupContract.DescribeMismatches(actual, page))}");
        return 1;
    }

    private static string NormalizePrintArea(string? value) =>
        (value ?? "").Replace("$", "", StringComparison.Ordinal).Trim();

    private static void AssignIfPresent(JsonObject page, string key, Action<string> assign)
    {
        if (page.ContainsKey(key)) assign(Json.GetString(page, key) ?? "");
    }

    private static JsonNode CaptureRowHeightSummary(object sheet, JsonObject op) =>
        CaptureRowHeightEntries(sheet, op);

    private static JsonNode CaptureColumnWidthSummary(object sheet, JsonObject op) =>
        CaptureColumnWidthEntries(sheet, op);

    private static string DescribeSizeSpan(JsonObject op, bool rows)
    {
        var items = Json.GetArr(op, rows ? "rows" : "columns") ?? new JsonArray();
        return $"{items.Count} span(s)";
    }

    private static int SafeInt(Func<int> read)
    {
        try { return read(); }
        catch { return 0; }
    }

    private const int XlNormalView = 1;
    private const int XlPageBreakPreview = 2;
    private const int XlPageLayoutView = 3;

    private static void ApplySheetView(object workbook, object sheet, string sheetName, JsonObject op,
        ApplyExecution execution, List<string> mismatches, ref int checkedItems)
    {
        WriteSheetView(workbook, sheet, op);
        var actual = CaptureSheetView(workbook, sheet);
        checkedItems++;
        if (Json.GetBool(actual, "readable") == false)
            mismatches.Add($"{sheetName}: view read failed: {Json.GetString(actual, "error")}");
        else if (!SheetViewMatches(actual, PlannedSheetView(op)))
            mismatches.Add($"{sheetName}: view readback mismatch");
        execution.Affected.Add(new AffectedRef("window", $"{sheetName}:view"));
    }

    private static int RestoreSheetView(object workbook, object sheet, string sheetName, JsonObject? view,
        RestoreMismatchCollector mismatches)
    {
        if (view is null)
        {
            mismatches.Add($"{sheetName}: view snapshot missing");
            return 0;
        }

        WriteSheetView(workbook, sheet, view);
        var actual = CaptureSheetView(workbook, sheet);
        if (Json.GetBool(actual, "readable") == false)
            mismatches.Add($"{sheetName}: view restore read failed: {Json.GetString(actual, "error")}");
        else if (!SheetViewMatches(actual, view))
            mismatches.Add($"{sheetName}: view restore mismatch");
        return 1;
    }

    private static void WriteSheetView(object workbook, object sheet, JsonObject requested)
    {
        object? window = null;
        try
        {
            var visible = Convert.ToInt32(((dynamic)sheet).Visible, CultureInfo.InvariantCulture);
            if (visible != XlSheetVisible)
                throw new InvalidOperationException("set_view cannot change a hidden sheet window");
            if (!TryGetWorkbookWindow(workbook, out window, out var error) || window is null)
                throw new InvalidOperationException(error ?? "workbook window is unavailable");
            ActivateWorksheet(workbook, Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture),
                new RestoreMismatchCollector());
            if (Json.GetInt(requested, "zoom") is int zoom)
                ((dynamic)window).Zoom = zoom;
            if (requested.ContainsKey("displayGridlines"))
                ((dynamic)window).DisplayGridlines = Json.GetBool(requested, "displayGridlines");
            if (requested.ContainsKey("displayHeadings"))
                ((dynamic)window).DisplayHeadings = Json.GetBool(requested, "displayHeadings");
            if (requested.ContainsKey("displayZeros"))
                ((dynamic)window).DisplayZeros = Json.GetBool(requested, "displayZeros");
            var view = Json.GetString(requested, "view");
            if (!string.IsNullOrWhiteSpace(view))
                ((dynamic)window).View = ViewNameToXl(view);
        }
        finally { RotHelper.ReleaseComReference(window); }
    }

    private static JsonObject CaptureSheetView(object workbook, object sheet)
    {
        var sheetName = Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture) ?? "";
        var visible = Convert.ToInt32(((dynamic)sheet).Visible, CultureInfo.InvariantCulture);
        if (visible != XlSheetVisible)
            return ExcelSheetViewCaptureContract.Unreadable(sheetName, $"view cannot be read from hidden sheet '{sheetName}'");
        return ExcelSheetViewCaptureContract.CaptureTargetView(
            new WorkbookSheetViewSurface(workbook, sheet), sheetName);
    }

    private static JsonObject PlannedSheetView(JsonObject op)
    {
        var planned = new JsonObject();
        if (Json.GetInt(op, "zoom") is int zoom) planned["zoom"] = zoom;
        if (op.ContainsKey("displayGridlines")) planned["displayGridlines"] = Json.GetBool(op, "displayGridlines");
        if (op.ContainsKey("displayHeadings")) planned["displayHeadings"] = Json.GetBool(op, "displayHeadings");
        if (op.ContainsKey("displayZeros")) planned["displayZeros"] = Json.GetBool(op, "displayZeros");
        if (!string.IsNullOrWhiteSpace(Json.GetString(op, "view")))
            planned["view"] = Json.GetString(op, "view");
        return planned;
    }

    private static bool SheetViewMatches(JsonObject actual, JsonObject expected)
    {
        if (expected.ContainsKey("zoom") && Json.GetInt(actual, "zoom") != Json.GetInt(expected, "zoom"))
            return false;
        if (expected.ContainsKey("displayGridlines") &&
            Json.GetBool(actual, "displayGridlines") != Json.GetBool(expected, "displayGridlines"))
            return false;
        if (expected.ContainsKey("displayHeadings") &&
            Json.GetBool(actual, "displayHeadings") != Json.GetBool(expected, "displayHeadings"))
            return false;
        if (expected.ContainsKey("displayZeros") &&
            Json.GetBool(actual, "displayZeros") != Json.GetBool(expected, "displayZeros"))
            return false;
        if (expected.ContainsKey("view") &&
            !string.Equals(Json.GetString(actual, "view"), Json.GetString(expected, "view"), StringComparison.OrdinalIgnoreCase))
            return false;
        return true;
    }

    private static int ViewNameToXl(string name) => name.ToLowerInvariant() switch
    {
        "pagelayout" => XlPageLayoutView,
        "pagebreakpreview" => XlPageBreakPreview,
        _ => XlNormalView,
    };

    private static string ViewXlToName(int value) => value switch
    {
        XlPageLayoutView => "pageLayout",
        XlPageBreakPreview => "pageBreakPreview",
        _ => "normal",
    };

    private static ExcelRowHeightContract.PixelMapping ReadOwnerRowHeightPixelMapping(object sheet)
    {
        object? window = null;
        object? parent = null;
        try
        {
            parent = (object)((dynamic)sheet).Parent;
            if (!TryGetWorkbookWindow(parent, out window, out _) || window is null)
                return ExcelRowHeightContract.PixelMapping.Unmeasured;
            return TryReadPointsToScreenPixelsY(window, out var windowMapping)
                ? windowMapping
                : ExcelRowHeightContract.PixelMapping.Unmeasured;
        }
        catch
        {
            return ExcelRowHeightContract.PixelMapping.Unmeasured;
        }
        finally
        {
            RotHelper.ReleaseComReference(window);
            RotHelper.ReleaseComReference(parent);
        }
    }

    private sealed class WorkbookSheetViewSurface : ExcelSheetViewCaptureContract.ISheetViewSurface
    {
        private readonly object _workbook;
        private readonly object _sheet;

        public WorkbookSheetViewSurface(object workbook, object sheet)
        {
            _workbook = workbook;
            _sheet = sheet;
        }

        public string? ActiveSheet => ReadActiveSheetName(_workbook);

        public ExcelSheetViewCaptureContract.WindowBookmark CaptureOriginal()
        {
            var window = CaptureWindowState(_workbook);
            return new ExcelSheetViewCaptureContract.WindowBookmark(
                Json.GetString(window, "activeSheet"),
                Json.GetString(window, "selection"),
                Json.GetInt(window, "scrollRow"),
                Json.GetInt(window, "scrollColumn"),
                ReadActiveWorkbookName(_sheet));
        }

        public void Activate(string sheet)
        {
            var mismatches = new RestoreMismatchCollector();
            ActivateWorksheet(_workbook, sheet, mismatches);
            if (mismatches.Count > 0)
                throw new InvalidOperationException(string.Join("; ", mismatches.Samples));
        }

        public JsonObject ReadActiveView()
        {
            object? window = null;
            try
            {
                if (!TryGetWorkbookWindow(_workbook, out window, out var error) || window is null)
                    throw new InvalidOperationException(error ?? "workbook window is unavailable");
                return new JsonObject
                {
                    ["zoom"] = SafeInt(() => Convert.ToInt32(((dynamic)window).Zoom, CultureInfo.InvariantCulture)),
                    ["displayGridlines"] = SafeBool(() => ((dynamic)window).DisplayGridlines),
                    ["displayHeadings"] = SafeBool(() => ((dynamic)window).DisplayHeadings),
                    ["displayZeros"] = SafeBool(() => ((dynamic)window).DisplayZeros),
                    ["view"] = ViewXlToName(SafeInt(() => Convert.ToInt32(((dynamic)window).View, CultureInfo.InvariantCulture))),
                };
            }
            finally { RotHelper.ReleaseComReference(window); }
        }

        public void Restore(ExcelSheetViewCaptureContract.WindowBookmark original)
        {
            RestoreActiveWorkbook(_sheet, original.ActiveWorkbook);
            var mismatches = new RestoreMismatchCollector();
            RestoreWindowState(_workbook, new JsonObject
            {
                ["activeSheet"] = original.ActiveSheet,
                ["selection"] = original.Selection,
                ["scrollRow"] = original.ScrollRow,
                ["scrollColumn"] = original.ScrollColumn,
            }, mismatches);
            if (mismatches.Count > 0)
                throw new InvalidOperationException(string.Join("; ", mismatches.Samples));
        }
    }

    private static string? ReadActiveWorkbookName(object excelObject)
    {
        object? application = null;
        object? active = null;
        try
        {
            application = (object)((dynamic)excelObject).Application;
            active = (object)((dynamic)application).ActiveWorkbook;
            return Convert.ToString(((dynamic)active).Name, CultureInfo.InvariantCulture);
        }
        catch
        {
            return null;
        }
        finally
        {
            RotHelper.ReleaseComReference(active);
            RotHelper.ReleaseComReference(application);
        }
    }

    private static void RestoreActiveWorkbook(object excelObject, string? name)
    {
        if (string.IsNullOrWhiteSpace(name))
            return;
        var current = ReadActiveWorkbookName(excelObject);
        if (string.Equals(current, name, StringComparison.OrdinalIgnoreCase))
            return;
        object? application = null;
        object? books = null;
        object? book = null;
        try
        {
            application = (object)((dynamic)excelObject).Application;
            books = (object)((dynamic)application).Workbooks;
            book = (object)((dynamic)books).Item(name);
            ((dynamic)book).Activate();
        }
        finally
        {
            RotHelper.ReleaseComReference(book);
            RotHelper.ReleaseComReference(books);
            RotHelper.ReleaseComReference(application);
        }
    }

    private static bool TryReadPointsToScreenPixelsY(object source, out ExcelRowHeightContract.PixelMapping mapping)
    {
        mapping = ExcelRowHeightContract.PixelMapping.Unmeasured;
        try
        {
            var y0 = Convert.ToInt32((object)((dynamic)source).PointsToScreenPixelsY(0), CultureInfo.InvariantCulture);
            var y72 = Convert.ToInt32(
                (object)((dynamic)source).PointsToScreenPixelsY(ExcelRowHeightContract.MeasureSpanPoints),
                CultureInfo.InvariantCulture);
            mapping = ExcelRowHeightContract.FromScreenPixelDelta(y0, y72);
            return mapping.Measured;
        }
        catch
        {
            return false;
        }
    }
}
