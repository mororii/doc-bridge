using System.Globalization;
using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

/// <summary>
/// Pivot, dedup, text-to-columns, lifecycle deletes, shapes, and textboxes.
/// Native COM only. Unsupported host states throw; they are not stubbed.
/// </summary>
public sealed partial class ExcelAdapter
{
    private static void PreviewAddTableColumn(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var table = BindTable(workbook, op);
        EnsureSheetWritable(table.Sheet);
        preview.Affected.Add(new AffectedRef("table", $"{table.SheetName}!{table.Name}:column"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{table.SheetName}!{table.Name}:column",
            Before = table.State["columns"]?.DeepClone(),
            After = JsonValue.Create(Json.GetString(op, "columnName")),
        });
    }

    private static void ApplyAddTableColumn(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var table = BindTable(workbook, op);
        EnsureSheetWritable(table.Sheet);
        object? columns = null;
        object? added = null;
        try
        {
            columns = (object)((dynamic)table.ListObject).ListColumns;
            var insertAfter = Json.GetInt(op, "insertAfter");
            added = insertAfter is int index
                ? (object)((dynamic)columns).Add(index)
                : (object)((dynamic)columns).Add();
            ((dynamic)added).Name = Json.GetString(op, "columnName");
            var formula = Json.GetString(op, "formula");
            if (!string.IsNullOrWhiteSpace(formula))
            {
                object? body = null;
                try
                {
                    body = (object)((dynamic)added).DataBodyRange;
                    ((dynamic)body).Formula = formula;
                }
                finally { RotHelper.ReleaseComReference(body); }
            }
            var actual = ReadListObjectState(table.ListObject);
            checkedCells++;
            var names = Json.GetArr(actual, "columns") ?? new JsonArray();
            var columnName = Json.GetString(op, "columnName");
            if (!names.Any(node => string.Equals(node?.ToString(), columnName, StringComparison.OrdinalIgnoreCase)))
                mismatches.Add($"{table.SheetName}!{table.Name}: column '{columnName}' missing after Add");
            if (!string.IsNullOrWhiteSpace(formula) && added is not null)
                VerifyCalculatedColumn(added, formula, $"{table.SheetName}!{table.Name}", mismatches);
            execution.Affected.Add(new AffectedRef("table", $"{table.SheetName}!{table.Name}:column"));
        }
        finally
        {
            RotHelper.ReleaseComReference(added);
            RotHelper.ReleaseComReference(columns);
        }
    }

    private static void PreviewDeleteTable(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var table = BindTable(workbook, op);
        EnsureSheetWritable(table.Sheet);
        preview.Affected.Add(new AffectedRef("table", $"{table.SheetName}!{table.Name}"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{table.SheetName}!{table.Name}",
            Before = table.State.DeepClone(),
            After = null,
        });
    }

    private static void ApplyDeleteTable(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var table = BindTable(workbook, op);
        EnsureSheetWritable(table.Sheet);
        var name = table.Name;
        var sheetName = table.SheetName;
        ((dynamic)table.ListObject).Delete();
        checkedCells++;
        if (FindListObject(table.Sheet, name) is { } leftover)
        {
            RotHelper.ReleaseComReference(leftover);
            mismatches.Add($"{sheetName}!{name}: ListObject still present after Delete");
        }
        execution.Affected.Add(new AffectedRef("table", $"{sheetName}!{name}"));
    }

    private static void PreviewRemoveDuplicates(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var range = BindRange(workbook, op, "range");
        EnsureSheetWritable(range.Sheet);
        preview.Affected.Add(new AffectedRef("range", $"{range.SheetName}!{range.Address}:duplicates"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{range.SheetName}!{range.Address}:duplicates",
            Before = CaptureRangeValues(range.Sheet, range.Address),
            After = op["columns"]?.DeepClone() ?? JsonValue.Create("all"),
        });
    }

    private static void ApplyRemoveDuplicates(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var range = BindRange(workbook, op, "range");
        EnsureSheetWritable(range.Sheet);
        object? comRange = null;
        try
        {
            comRange = (object)((dynamic)range.Sheet).Range(range.Address);
            if (!ExcelDataOperationsContract.TryParseA1(range.Address, out var box))
                throw new InvalidOperationException("remove_duplicates range is not rectangular");
            var columns = Json.GetArr(op, "columns");
            object keys = columns is { Count: > 0 }
                ? columns.Select(node => node is JsonValue value && value.TryGetValue<int>(out var n) ? n : 0)
                    .Where(n => n > 0).Cast<object>().ToArray()
                : Enumerable.Range(1, box.Columns).Cast<object>().ToArray();
            ((dynamic)comRange).RemoveDuplicates(keys, ExcelDataObjectCatalog.HasHeaders(Json.GetBool(op, "hasHeaders", true)));
            checkedCells++;
            VerifyDuplicatesRemoved(range.Sheet, range.Address, columns, Json.GetBool(op, "hasHeaders", true), mismatches);
            execution.Affected.Add(new AffectedRef("range", $"{range.SheetName}!{range.Address}:duplicates"));
        }
        finally { RotHelper.ReleaseComReference(comRange); }
    }

    private static void PreviewTextToColumns(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var range = BindRange(workbook, op, "range");
        DataRangeLease? destination = null;
        try
        {
            if (!string.IsNullOrWhiteSpace(Json.GetString(op, "destination")))
                destination = BindRange(workbook, op, "destination");
            var destSheet = destination?.Sheet ?? range.Sheet;
            var destSheetName = destination?.SheetName ?? range.SheetName;
            var destOrigin = destination?.Address ?? range.Address;
            EnsureSheetWritable(destSheet);
            var captureAddress = ComputeTextToColumnsCaptureAddress(range, destOrigin, op);
            CollectTextToColumnsOverwriteConflicts(
                destSheet, destSheetName, range.SheetName, range.Address, captureAddress, preview);
            preview.Affected.Add(new AffectedRef("range", $"{destSheetName}!{captureAddress}:textToColumns"));
            preview.Diff.Add(new DiffEntry
            {
                Ref = $"{destSheetName}!{captureAddress}:textToColumns",
                Before = CaptureRangeValues(destSheet, captureAddress),
                After = new JsonObject
                {
                    ["dataType"] = Json.GetString(op, "dataType") ?? "delimited",
                    ["source"] = $"{range.SheetName}!{range.Address}",
                    ["destination"] = destOrigin,
                    ["footprint"] = captureAddress,
                },
            });
        }
        finally { destination?.Dispose(); }
    }

    private static void ApplyTextToColumns(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var range = BindRange(workbook, op, "range");
        DataRangeLease? destLease = null;
        object? comRange = null;
        object? destination = null;
        try
        {
            if (!string.IsNullOrWhiteSpace(Json.GetString(op, "destination")))
                destLease = BindRange(workbook, op, "destination");
            var destSheet = destLease?.Sheet ?? range.Sheet;
            var destSheetName = destLease?.SheetName ?? range.SheetName;
            var destOrigin = destLease?.Address ?? range.Address;
            EnsureSheetWritable(destSheet);
            comRange = (object)((dynamic)range.Sheet).Range(range.Address);
            var before = CaptureRangeValues(range.Sheet, range.Address);
            if (destLease is not null)
                destination = (object)((dynamic)destLease.Sheet).Range(destLease.Address);
            ((dynamic)comRange).TextToColumns(
                Destination: destination ?? comRange,
                DataType: ExcelDataObjectCatalog.XlDelimited,
                TextQualifier: ExcelTextToColumnsParse.ComTextQualifier(op),
                ConsecutiveDelimiter: Json.GetBool(op, "consecutiveDelimiter"),
                Tab: Json.GetBool(op, "tab", !op.ContainsKey("comma") && !op.ContainsKey("semicolon") && !op.ContainsKey("space") && !op.ContainsKey("other")),
                Semicolon: Json.GetBool(op, "semicolon"),
                Comma: Json.GetBool(op, "comma", true),
                Space: Json.GetBool(op, "space"),
                Other: Json.GetBool(op, "other"),
                OtherChar: Json.GetString(op, "otherChar") ?? Type.Missing);
            checkedCells++;
            var afterAddress = ComputeTextToColumnsCaptureAddress(range, destOrigin, op);
            VerifyTextToColumns(destSheet, range.Address, afterAddress, before, op, mismatches);
            execution.Affected.Add(new AffectedRef("range", $"{destSheetName}!{afterAddress}:textToColumns"));
        }
        finally
        {
            RotHelper.ReleaseComReference(destination);
            RotHelper.ReleaseComReference(comRange);
            destLease?.Dispose();
        }
    }

    private static void PreviewDeleteChart(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var chart = BindChart(workbook, op);
        EnsureSheetWritable(chart.Sheet);
        preview.Affected.Add(new AffectedRef("chart", $"{chart.SheetName}!{chart.Name}"));
        preview.Diff.Add(new DiffEntry { Ref = $"{chart.SheetName}!{chart.Name}", Before = chart.State.DeepClone(), After = null });
    }

    private static void ApplyDeleteChart(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var chart = BindChart(workbook, op);
        EnsureSheetWritable(chart.Sheet);
        var name = chart.Name;
        var sheetName = chart.SheetName;
        ((dynamic)chart.ChartObject).Delete();
        checkedCells++;
        if (FindChartObject(chart.Sheet, name) is { } leftover)
        {
            RotHelper.ReleaseComReference(leftover);
            mismatches.Add($"{sheetName}!{name}: ChartObject still present after Delete");
        }
        execution.Affected.Add(new AffectedRef("chart", $"{sheetName}!{name}"));
    }

    private static void PreviewDeletePicture(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var picture = BindPicture(workbook, op);
        EnsureSheetWritable(picture.Sheet);
        preview.Affected.Add(new AffectedRef("picture", $"{picture.SheetName}!{picture.Name}"));
        preview.Diff.Add(new DiffEntry { Ref = $"{picture.SheetName}!{picture.Name}", Before = picture.State.DeepClone(), After = null });
    }

    private static void ApplyDeletePicture(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var picture = BindPicture(workbook, op);
        EnsureSheetWritable(picture.Sheet);
        var name = picture.Name;
        var sheetName = picture.SheetName;
        ((dynamic)picture.Shape).Delete();
        checkedCells++;
        if (FindShape(picture.Sheet, name) is { } leftover)
        {
            RotHelper.ReleaseComReference(leftover);
            mismatches.Add($"{sheetName}!{name}: picture still present after Delete");
        }
        execution.Affected.Add(new AffectedRef("picture", $"{sheetName}!{name}"));
    }

    private static void PreviewCreatePivot(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var source = BindPivotSource(workbook, op);
        using var dest = BindRange(workbook, op, "destination");
        EnsureSheetWritable(dest.Sheet);
        var name = Json.GetString(op, "name")!;
        if (FindPivotTable(dest.Sheet, name) is { } conflict)
        {
            RotHelper.ReleaseComReference(conflict);
            preview.Errors.Add($"[EXCEL_PIVOT_EXISTS] PivotTable '{name}' already exists on '{dest.SheetName}'");
            return;
        }
        preview.Affected.Add(new AffectedRef("pivot", $"{dest.SheetName}!{name}"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{dest.SheetName}!{name}",
            Before = null,
            After = new JsonObject
            {
                ["source"] = source.CacheSource,
                ["destination"] = dest.Address,
                ["rows"] = op["rows"]?.DeepClone(),
                ["values"] = op["values"]?.DeepClone(),
            },
        });
    }

    private static void ApplyCreatePivot(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var source = BindPivotSource(workbook, op);
        using var dest = BindRange(workbook, op, "destination");
        EnsureSheetWritable(dest.Sheet);
        object? caches = null;
        object? cache = null;
        object? destRange = null;
        object? pivot = null;
        try
        {
            destRange = (object)((dynamic)dest.Sheet).Range(dest.Address);
            caches = (object)((dynamic)workbook).PivotCaches();
            cache = (object)((dynamic)caches).Create(ExcelDataObjectCatalog.XlDatabase, source.CacheSource);
            var name = Json.GetString(op, "name")!;
            pivot = (object)((dynamic)cache).CreatePivotTable(destRange, name);
            ApplyPivotLayout(pivot, op, replaceOmitted: true);
            var actual = ReadPivotState(pivot, dest.SheetName);
            checkedCells++;
            VerifyPivotRequested(actual, op, dest.SheetName, mismatches);
            execution.Affected.Add(new AffectedRef("pivot", $"{dest.SheetName}!{name}"));
        }
        finally
        {
            RotHelper.ReleaseComReference(pivot);
            RotHelper.ReleaseComReference(destRange);
            RotHelper.ReleaseComReference(cache);
            RotHelper.ReleaseComReference(caches);
        }
    }

    private static void PreviewUpdatePivot(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var sheet = BindSheet(workbook, op);
        var name = Json.GetString(op, "name")!;
        var pivot = FindPivotTable(sheet.Sheet, name);
        try
        {
            if (pivot is null)
            {
                preview.Errors.Add($"[EXCEL_PIVOT_NOT_FOUND] PivotTable '{name}' was not found");
                return;
            }
            EnsureSheetWritable(sheet.Sheet);
            if (op.ContainsKey("sourceRange"))
            {
                using var source = BindPivotSource(workbook, op);
                preview.Affected.Add(new AffectedRef("pivot", $"{sheet.SheetName}!{name}:{source.CacheSource}"));
            }
            preview.Affected.Add(new AffectedRef("pivot", $"{sheet.SheetName}!{name}"));
            preview.Diff.Add(new DiffEntry
            {
                Ref = $"{sheet.SheetName}!{name}",
                Before = ReadPivotState(pivot, sheet.SheetName),
                After = op.DeepClone(),
            });
        }
        finally { RotHelper.ReleaseComReference(pivot); }
    }

    private static void ApplyUpdatePivot(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var sheet = BindSheet(workbook, op);
        EnsureSheetWritable(sheet.Sheet);
        var name = Json.GetString(op, "name")!;
        var pivot = FindPivotTable(sheet.Sheet, name);
        try
        {
            if (pivot is null)
                throw new InvalidOperationException($"[EXCEL_PIVOT_NOT_FOUND] PivotTable '{name}' was not found");
            if (op.ContainsKey("sourceRange"))
            {
                using var source = BindPivotSource(workbook, op);
                ApplyPivotCacheSource(workbook, pivot, source.CacheSource);
            }
            ApplyPivotLayout(pivot, op, replaceOmitted: false);
            var actual = ReadPivotState(pivot, sheet.SheetName);
            checkedCells++;
            VerifyPivotRequested(actual, op, sheet.SheetName, mismatches);
            execution.Affected.Add(new AffectedRef("pivot", $"{sheet.SheetName}!{name}"));
        }
        finally { RotHelper.ReleaseComReference(pivot); }
    }

    private static void PreviewRefreshPivot(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var sheet = BindSheet(workbook, op);
        var name = Json.GetString(op, "name")!;
        var pivot = FindPivotTable(sheet.Sheet, name);
        try
        {
            if (pivot is null)
            {
                preview.Errors.Add($"[EXCEL_PIVOT_NOT_FOUND] PivotTable '{name}' was not found");
                return;
            }
            preview.Affected.Add(new AffectedRef("pivot", $"{sheet.SheetName}!{name}:refresh"));
            preview.Diff.Add(new DiffEntry
            {
                Ref = $"{sheet.SheetName}!{name}:refresh",
                Before = ReadPivotState(pivot, sheet.SheetName),
                After = JsonValue.Create("refresh"),
            });
        }
        finally { RotHelper.ReleaseComReference(pivot); }
    }

    private static void ApplyRefreshPivot(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var sheet = BindSheet(workbook, op);
        var name = Json.GetString(op, "name")!;
        var pivot = FindPivotTable(sheet.Sheet, name);
        try
        {
            if (pivot is null)
                throw new InvalidOperationException($"[EXCEL_PIVOT_NOT_FOUND] PivotTable '{name}' was not found");
            var refreshed = Convert.ToBoolean(((dynamic)pivot).RefreshTable(), CultureInfo.InvariantCulture);
            checkedCells++;
            if (!refreshed)
                mismatches.Add($"{sheet.SheetName}!{name}: RefreshTable returned false");
            var actual = ReadPivotState(pivot, sheet.SheetName);
            VerifyPivotRequested(actual, op, sheet.SheetName, mismatches, requireAggregates: true);
            execution.Affected.Add(new AffectedRef("pivot", $"{sheet.SheetName}!{name}:refresh"));
        }
        finally { RotHelper.ReleaseComReference(pivot); }
    }

    private static void PreviewDeletePivot(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var sheet = BindSheet(workbook, op);
        var name = Json.GetString(op, "name")!;
        var pivot = FindPivotTable(sheet.Sheet, name);
        try
        {
            if (pivot is null)
            {
                preview.Errors.Add($"[EXCEL_PIVOT_NOT_FOUND] PivotTable '{name}' was not found");
                return;
            }
            EnsureSheetWritable(sheet.Sheet);
            preview.Affected.Add(new AffectedRef("pivot", $"{sheet.SheetName}!{name}"));
            preview.Diff.Add(new DiffEntry
            {
                Ref = $"{sheet.SheetName}!{name}",
                Before = ReadPivotState(pivot, sheet.SheetName),
                After = null,
            });
        }
        finally { RotHelper.ReleaseComReference(pivot); }
    }

    private static void ApplyDeletePivot(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var sheet = BindSheet(workbook, op);
        EnsureSheetWritable(sheet.Sheet);
        var name = Json.GetString(op, "name")!;
        var pivot = FindPivotTable(sheet.Sheet, name);
        try
        {
            if (pivot is null)
                throw new InvalidOperationException($"[EXCEL_PIVOT_NOT_FOUND] PivotTable '{name}' was not found");
            DeletePivotTable(pivot);
        }
        finally { RotHelper.ReleaseComReference(pivot); }
        checkedCells++;
        if (FindPivotTable(sheet.Sheet, name) is { } leftover)
        {
            RotHelper.ReleaseComReference(leftover);
            mismatches.Add($"{sheet.SheetName}!{name}: PivotTable still present after delete");
        }
        execution.Affected.Add(new AffectedRef("pivot", $"{sheet.SheetName}!{name}"));
    }

    private static void PreviewInsertShape(object workbook, JsonObject op, ApplyPreview preview) =>
        PreviewInsertDrawing(workbook, op, preview, "shape");

    private static void PreviewInsertTextbox(object workbook, JsonObject op, ApplyPreview preview) =>
        PreviewInsertDrawing(workbook, op, preview, "textbox");

    private static void PreviewInsertDrawing(object workbook, JsonObject op, ApplyPreview preview, string kind)
    {
        using var sheet = BindSheet(workbook, op);
        EnsureSheetWritable(sheet.Sheet);
        var requested = Json.GetString(op, "name");
        if (!string.IsNullOrWhiteSpace(requested) && FindShape(sheet.Sheet, requested) is { } conflict)
        {
            RotHelper.ReleaseComReference(conflict);
            preview.Errors.Add($"[EXCEL_SHAPE_EXISTS] Shape '{requested}' already exists on '{sheet.SheetName}'");
            return;
        }
        preview.Affected.Add(new AffectedRef(kind, $"{sheet.SheetName}!{requested ?? "new"}"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{sheet.SheetName}:{kind}",
            Before = null,
            After = new JsonObject { ["name"] = requested, ["shapeType"] = Json.GetString(op, "shapeType"), ["text"] = Json.GetString(op, "text") },
        });
    }

    private static void ApplyInsertShape(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells) =>
        ApplyInsertDrawing(workbook, op, execution, mismatches, ref checkedCells, textbox: false);

    private static void ApplyInsertTextbox(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells) =>
        ApplyInsertDrawing(workbook, op, execution, mismatches, ref checkedCells, textbox: true);

    private static void ApplyInsertDrawing(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells, bool textbox)
    {
        using var sheet = BindSheet(workbook, op);
        EnsureSheetWritable(sheet.Sheet);
        object? shapes = null;
        object? shape = null;
        try
        {
            var position = ReadPosition(op, DefaultPictureLeft, DefaultPictureTop, 160, 60);
            shapes = (object)((dynamic)sheet.Sheet).Shapes;
            shape = textbox
                ? (object)((dynamic)shapes).AddTextbox(
                    ExcelDataObjectCatalog.MsoTextOrientationHorizontal,
                    position.Left, position.Top, position.Width, position.Height)
                : (object)((dynamic)shapes).AddShape(
                    ExcelDataObjectCatalog.TryShapeType(Json.GetString(op, "shapeType"), out var type)
                        ? type
                        : ExcelDataObjectCatalog.MsoShapeRectangle,
                    position.Left, position.Top, position.Width, position.Height);
            var requested = Json.GetString(op, "name");
            if (!string.IsNullOrWhiteSpace(requested))
                ((dynamic)shape).Name = requested;
            SetShapeText(shape, Json.GetString(op, "text"));
            ApplyShapeFill(shape, op);
            ApplyShapeFormatting(shape, op);
            var actual = ReadShapeState(shape, sheet.SheetName);
            checkedCells++;
            VerifyRequestedShape(actual, op, sheet.SheetName, textbox, mismatches);
            execution.Affected.Add(new AffectedRef(textbox ? "textbox" : "shape",
                $"{sheet.SheetName}!{Json.GetString(actual, "name")}"));
        }
        finally
        {
            RotHelper.ReleaseComReference(shape);
            RotHelper.ReleaseComReference(shapes);
        }
    }

    private static void PreviewUpdateShape(object workbook, JsonObject op, ApplyPreview preview) =>
        PreviewUpdateDrawing(workbook, op, preview);

    private static void PreviewUpdateTextbox(object workbook, JsonObject op, ApplyPreview preview) =>
        PreviewUpdateDrawing(workbook, op, preview);

    private static void PreviewUpdateDrawing(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var sheet = BindSheet(workbook, op);
        var name = Json.GetString(op, "name")!;
        var shape = FindShape(sheet.Sheet, name);
        try
        {
            if (shape is null)
            {
                preview.Errors.Add($"[EXCEL_SHAPE_NOT_FOUND] Shape '{name}' was not found");
                return;
            }
            EnsureSheetWritable(sheet.Sheet);
            preview.Affected.Add(new AffectedRef("shape", $"{sheet.SheetName}!{name}"));
            preview.Diff.Add(new DiffEntry
            {
                Ref = $"{sheet.SheetName}!{name}",
                Before = ReadShapeState(shape, sheet.SheetName),
                After = op.DeepClone(),
            });
        }
        finally { RotHelper.ReleaseComReference(shape); }
    }

    private static void ApplyUpdateShape(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells) =>
        ApplyUpdateDrawing(workbook, op, execution, mismatches, ref checkedCells);

    private static void ApplyUpdateTextbox(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells) =>
        ApplyUpdateDrawing(workbook, op, execution, mismatches, ref checkedCells);

    private static void ApplyUpdateDrawing(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var sheet = BindSheet(workbook, op);
        EnsureSheetWritable(sheet.Sheet);
        var name = Json.GetString(op, "name")!;
        var shape = FindShape(sheet.Sheet, name);
        try
        {
            if (shape is null)
                throw new InvalidOperationException($"[EXCEL_SHAPE_NOT_FOUND] Shape '{name}' was not found");
            if (op.ContainsKey("position"))
                ApplyPosition(shape, ReadPosition(op, 0, 0, 0, 0, optionalSize: true));
            if (op.ContainsKey("text"))
                SetShapeText(shape, Json.GetString(op, "text"));
            ApplyShapeFill(shape, op);
            ApplyShapeFormatting(shape, op);
            var actual = ReadShapeState(shape, sheet.SheetName);
            checkedCells++;
            VerifyRequestedShape(actual, op, sheet.SheetName,
                textbox: Json.GetInt(actual, "shapeType") == ExcelDataObjectCatalog.MsoTextBox, mismatches);
            execution.Affected.Add(new AffectedRef("shape", $"{sheet.SheetName}!{name}"));
        }
        finally { RotHelper.ReleaseComReference(shape); }
    }

    private static void PreviewDeleteShape(object workbook, JsonObject op, ApplyPreview preview) =>
        PreviewDeleteDrawing(workbook, op, preview);

    private static void PreviewDeleteTextbox(object workbook, JsonObject op, ApplyPreview preview) =>
        PreviewDeleteDrawing(workbook, op, preview);

    private static void PreviewDeleteDrawing(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var sheet = BindSheet(workbook, op);
        var name = Json.GetString(op, "name")!;
        var shape = FindShape(sheet.Sheet, name);
        try
        {
            if (shape is null)
            {
                preview.Errors.Add($"[EXCEL_SHAPE_NOT_FOUND] Shape '{name}' was not found");
                return;
            }
            EnsureSheetWritable(sheet.Sheet);
            preview.Affected.Add(new AffectedRef("shape", $"{sheet.SheetName}!{name}"));
            preview.Diff.Add(new DiffEntry
            {
                Ref = $"{sheet.SheetName}!{name}",
                Before = ReadShapeState(shape, sheet.SheetName),
                After = null,
            });
        }
        finally { RotHelper.ReleaseComReference(shape); }
    }

    private static void ApplyDeleteShape(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells) =>
        ApplyDeleteDrawing(workbook, op, execution, mismatches, ref checkedCells);

    private static void ApplyDeleteTextbox(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells) =>
        ApplyDeleteDrawing(workbook, op, execution, mismatches, ref checkedCells);

    private static void ApplyDeleteDrawing(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var sheet = BindSheet(workbook, op);
        EnsureSheetWritable(sheet.Sheet);
        var name = Json.GetString(op, "name")!;
        var shape = FindShape(sheet.Sheet, name);
        try
        {
            if (shape is null)
                throw new InvalidOperationException($"[EXCEL_SHAPE_NOT_FOUND] Shape '{name}' was not found");
            ((dynamic)shape).Delete();
        }
        finally { RotHelper.ReleaseComReference(shape); }
        checkedCells++;
        if (FindShape(sheet.Sheet, name) is { } leftover)
        {
            RotHelper.ReleaseComReference(leftover);
            mismatches.Add($"{sheet.SheetName}!{name}: shape still present after Delete");
        }
        execution.Affected.Add(new AffectedRef("shape", $"{sheet.SheetName}!{name}"));
    }
}
