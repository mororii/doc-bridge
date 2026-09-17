using System.Globalization;
using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class ExcelAdapter
{
    private const int XlFillDefault = 0;
    private const int XlPageBreakManual = -4135;
    private const int XlPageBreakNone = -4142;
    internal const string ExtendedOpsRestoreMode = "extended-ops";

    internal static bool IsExtendedOnlySnapshot(IReadOnlyList<JsonObject>? ops) =>
        ops is { Count: > 0 } &&
        ops.All(op => OperationValidator.ExcelExtendedOpNames.Contains(Json.GetString(op, "op") ?? ""));

    internal static bool IsFormula2WriteSnapshot(IReadOnlyList<JsonObject>? ops) =>
        ops is { Count: > 0 } &&
        ops.All(op => string.Equals(Json.GetString(op, "op"), "set_formulas", StringComparison.OrdinalIgnoreCase)) &&
        ops.Any(ExcelFormula2Contract.WritesFormula2);

    private static bool TryPreviewExtendedOperation(object workbook, JsonObject op, ApplyPreview preview)
    {
        var name = Json.GetString(op, "op")!;
        if (!OperationValidator.ExcelExtendedOpNames.Contains(name) &&
            name is not ("delete_sheet" or "set_tab_color"))
            return false;
        switch (name)
        {
            case "fill_range":
            case "auto_fill":
            case "set_page_breaks":
            case "import_csv":
            {
                var range = Json.GetString(op, "range") ?? Json.GetString(op, "destination") ?? "A1";
                var sheet = Json.GetString(Json.GetObj(op, "target"), "sheet") ?? "";
                preview.Affected.Add(new AffectedRef("range", $"{sheet}!{range}"));
                preview.Diff.Add(new DiffEntry { Ref = name, Before = null, After = range });
                return true;
            }
            case "set_outline":
            {
                var range = Json.GetString(op, "range") ?? "";
                var sheet = Json.GetString(Json.GetObj(op, "target"), "sheet") ?? "";
                var description = ExcelOutlineContract.Describe(op);
                preview.Affected.Add(new AffectedRef("outline", $"{sheet}!{range}"));
                preview.Diff.Add(new DiffEntry { Ref = "set_outline", Before = null, After = description });
                return true;
            }
            case "export_csv":
            {
                var output = Json.GetString(op, "output") ?? "";
                preview.Affected.Add(new AffectedRef("file", output));
                preview.Diff.Add(new DiffEntry { Ref = "export_csv", Before = null, After = output });
                return true;
            }
            case "calculate":
            {
                if (ExcelCalculateContract.RequestsFormula2Rewrite(op))
                    preview.Errors.Add("calculate.formula2 is not a Formula2 writer; use set_formulas");
                var description = ExcelCalculateContract.Describe(op);
                preview.Affected.Add(new AffectedRef(
                    ExcelCalculateContract.ResolveScope(op) == ExcelCalculateContract.CalculateScope.Range
                        ? "range"
                        : "workbook",
                    Json.GetString(op, "range") ?? "workbook"));
                preview.Diff.Add(new DiffEntry { Ref = "calculate", Before = null, After = description });
                return true;
            }
            case "delete_sheet":
                preview.Affected.Add(new AffectedRef("sheet", Json.GetString(Json.GetObj(op, "target"), "sheet") ?? ""));
                preview.Warnings.Add(
                    "delete_sheet restore copies the whole sheet from a current SaveCopyAs/last-saved workbook backup, " +
                    "then restores captured other-sheet formulas/names/charts; used-range replay is refused; " +
                    "backup-file external links are not a restore source");
                return true;
            case "set_tab_color":
                preview.Affected.Add(new AffectedRef("sheet", Json.GetString(Json.GetObj(op, "target"), "sheet") ?? ""));
                return true;
            default:
                return false;
        }
    }

    private static bool TryApplyExtendedOperation(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        var name = Json.GetString(op, "op")!;
        switch (name)
        {
            case "fill_range":
                ApplyFillRange(workbook, op, execution, mismatches, ref checkedItems);
                return true;
            case "auto_fill":
                ApplyAutoFill(workbook, op, execution, mismatches, ref checkedItems);
                return true;
            case "calculate":
                ApplyCalculate(workbook, op, execution, mismatches, ref checkedItems);
                return true;
            case "set_tab_color":
                ApplyTabColor(workbook, op, execution, mismatches, ref checkedItems);
                return true;
            case "set_outline":
                ApplyOutline(workbook, op, execution, mismatches, ref checkedItems);
                return true;
            case "import_csv":
                ApplyImportCsv(workbook, op, execution, mismatches, ref checkedItems);
                return true;
            case "export_csv":
                ApplyExportCsv(workbook, op, execution, mismatches, ref checkedItems);
                return true;
            case "set_page_breaks":
                ApplyPageBreaks(workbook, op, execution, mismatches, ref checkedItems);
                return true;
            case "delete_sheet":
                ApplyDeleteSheet(workbook, op, execution, mismatches, ref checkedItems);
                return true;
            default:
                return false;
        }
    }

    private static void ApplyFillRange(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        var resolved = ResolveBasicRange(workbook, op);
        try
        {
            object written;
            if (Json.GetArr(op, "values") is { Count: > 0 } values)
            {
                written = ToComArray(values);
                ((dynamic)resolved.Range).Value2 = written;
            }
            else
            {
                written = JsonNodeToComValue(op["value"]) ?? "";
                ((dynamic)resolved.Range).Value2 = written;
            }

            object? raw = ((dynamic)resolved.Range).Value2;
            var rows = Convert.ToInt32(((dynamic)resolved.Range).Rows.Count, CultureInfo.InvariantCulture);
            var cols = Convert.ToInt32(((dynamic)resolved.Range).Columns.Count, CultureInfo.InvariantCulture);
            for (var r = 0; r < rows; r++)
            for (var c = 0; c < cols; c++)
            {
                checkedItems++;
                object? want = written is object[,] matrix ? matrix[r, c] : written;
                object? got = raw is object[,] arr ? arr[r + 1, c + 1] : raw;
                if (!ComValuesEqual(want, got))
                    mismatches.Add($"fill_range {resolved.Address}: want '{ToDisp(want)}', got '{ToDisp(got)}'");
            }

            execution.Affected.Add(new AffectedRef("range",
                $"{Convert.ToString(((dynamic)resolved.Sheet).Name, CultureInfo.InvariantCulture)}!{resolved.Address}"));
        }
        finally
        {
            RotHelper.ReleaseComReference(resolved.Range);
            RotHelper.ReleaseComReference(resolved.Sheet);
        }
    }

    private static void ApplyAutoFill(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        var source = ResolveBasicRange(workbook, op);
        object? dest = null;
        try
        {
            dest = (object)((dynamic)source.Sheet).Range(Json.GetString(op, "destRange"));
            object? sourceValues = ((dynamic)source.Range).Value2;
            var sourceRows = Convert.ToInt32(((dynamic)source.Range).Rows.Count, CultureInfo.InvariantCulture);
            var sourceCols = Convert.ToInt32(((dynamic)source.Range).Columns.Count, CultureInfo.InvariantCulture);
            ((dynamic)source.Range).AutoFill(dest, XlFillDefault);

            var destRows = Convert.ToInt32(((dynamic)dest).Rows.Count, CultureInfo.InvariantCulture);
            var destCols = Convert.ToInt32(((dynamic)dest).Columns.Count, CultureInfo.InvariantCulture);
            if (destRows < sourceRows || destCols < sourceCols)
                mismatches.Add($"auto_fill dest {Json.GetString(op, "destRange")} is smaller than source {source.Address}");

            for (var r = 0; r < sourceRows; r++)
            for (var c = 0; c < sourceCols; c++)
            {
                checkedItems++;
                object? want = sourceValues is object[,] src ? src[r + 1, c + 1] : sourceValues;
                object? destValues = ((dynamic)dest).Value2;
                object? got = destValues is object[,] dst ? dst[r + 1, c + 1] : destValues;
                if (!ComValuesEqual(want, got))
                    mismatches.Add($"auto_fill dest does not preserve source {source.Address}: want '{ToDisp(want)}', got '{ToDisp(got)}'");
            }

            var sourceTable = JsonArrayToStringTable(RangeToJson(source.Range, out _, maxCells: MaxSnapshotCells));
            var destTable = JsonArrayToStringTable(RangeToJson(dest, out _, maxCells: MaxSnapshotCells));
            var sourceFormulas = JsonArrayToStringTable(RangeToJson(source.Range, out _, formulas: true, maxCells: MaxSnapshotCells));
            var destFormulas = JsonArrayToStringTable(RangeToJson(dest, out _, formulas: true, maxCells: MaxSnapshotCells));
            if (!ExcelFillCopyContract.FilledAreaHasContent(sourceTable, destTable))
                mismatches.Add($"auto_fill dest {Json.GetString(op, "destRange")} filled area is empty or not larger than source");
            else
                checkedItems += ExcelFillCopyContract.FilledBand(sourceRows, sourceCols, destRows, destCols).Count;

            var sourceHasFormula = sourceFormulas.Any(row => row.Any(ExcelFillCopyContract.LooksLikeFormula));
            if (sourceHasFormula && !ExcelFillCopyContract.FilledFormulasContinue(sourceFormulas, destFormulas))
                mismatches.Add($"auto_fill dest {Json.GetString(op, "destRange")} filled formulas did not continue with relative references");
            else if (!sourceHasFormula && !ExcelFillCopyContract.FilledSequenceContinues(sourceTable, destTable))
                mismatches.Add($"auto_fill dest {Json.GetString(op, "destRange")} filled sequence/values did not continue the source series");

            execution.Affected.Add(new AffectedRef("range", Json.GetString(op, "destRange") ?? source.Address));
        }
        finally
        {
            RotHelper.ReleaseComReference(dest);
            RotHelper.ReleaseComReference(source.Range);
            RotHelper.ReleaseComReference(source.Sheet);
        }
    }

    private static void ApplyCalculate(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        var range = Json.GetString(op, "range");
        if (ExcelCalculateContract.RequestsFormula2Rewrite(op))
            throw new InvalidOperationException("calculate.formula2 is not a Formula2 writer; use set_formulas");

        if (!string.IsNullOrWhiteSpace(range))
        {
            var resolved = ResolveBasicRange(workbook, op);
            try
            {
                ((dynamic)resolved.Range).Calculate();
                object? application = null;
                try
                {
                    application = (object)((dynamic)workbook).Application;
                    var state = Convert.ToInt32(((dynamic)application).CalculationState, CultureInfo.InvariantCulture);
                    if (!ExcelCalculateContract.RecalculationComplete(state))
                        mismatches.Add($"calculate range left CalculationState {state}, not xlDone");
                }
                catch { /* CalculationState is optional verification */ }
                finally { RotHelper.ReleaseComReference(application); }
                checkedItems++;
                execution.Affected.Add(new AffectedRef("range",
                    $"{Convert.ToString(((dynamic)resolved.Sheet).Name, CultureInfo.InvariantCulture)}!{resolved.Address}"));
            }
            finally
            {
                RotHelper.ReleaseComReference(resolved.Range);
                RotHelper.ReleaseComReference(resolved.Sheet);
            }
            return;
        }

        object? sheets = null;
        object? workbookApplication = null;
        var calculated = 0;
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
                    int? type = null;
                    try { type = Convert.ToInt32(((dynamic)sheet).Type, CultureInfo.InvariantCulture); }
                    catch { type = ExcelCalculateContract.XlWorksheet; }
                    if (!ExcelCalculateContract.IsWorksheetType(type))
                        continue;
                    ((dynamic)sheet).Calculate();
                    calculated++;
                }
                finally { RotHelper.ReleaseComReference(sheet); }
            }

            try
            {
                workbookApplication = (object)((dynamic)workbook).Application;
                var state = Convert.ToInt32(((dynamic)workbookApplication).CalculationState, CultureInfo.InvariantCulture);
                if (!ExcelCalculateContract.RecalculationComplete(state))
                    mismatches.Add($"calculate worksheets left CalculationState {state}, not xlDone");
            }
            catch
            {
                if (calculated == 0)
                    mismatches.Add("calculate did not reach any worksheet.Calculate");
            }

            if (calculated == 0)
                mismatches.Add("calculate found no worksheet to Calculate");
            else
                checkedItems += calculated;
            execution.Affected.Add(new AffectedRef("workbook", ExcelCalculateContract.Describe(op)));
        }
        finally
        {
            RotHelper.ReleaseComReference(workbookApplication);
            RotHelper.ReleaseComReference(sheets);
        }
    }

    private static void ApplyTabColor(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        object? sheet = null;
        object? tab = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, op);
            tab = (object)((dynamic)sheet).Tab;
            if (!ExcelStyleContract.TryParseColor(op["color"], out var ole, errors: null, key: "color"))
                throw new InvalidOperationException("set_tab_color color must be OLE or #RRGGBB");
            ((dynamic)tab).Color = ole;
            checkedItems++;
            execution.Affected.Add(new AffectedRef("sheet", Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture)));
        }
        finally
        {
            RotHelper.ReleaseComReference(tab);
            RotHelper.ReleaseComReference(sheet);
        }
    }

    private static void ApplyOutline(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        var resolved = ResolveBasicRange(workbook, op);
        object? grouped = null;
        object? entire = null;
        object? outline = null;
        try
        {
            var action = ExcelOutlineContract.ResolveAction(op);
            var level = ExcelOutlineContract.RequestedLevel(op);
            var axis = ExcelOutlineContract.ResolveAxis(op);
            grouped = axis == ExcelOutlineContract.OutlineAxis.Columns
                ? (object)((dynamic)resolved.Range).Columns
                : (object)((dynamic)resolved.Range).Rows;
            entire = axis == ExcelOutlineContract.OutlineAxis.Columns
                ? (object)((dynamic)resolved.Range).EntireColumn
                : (object)((dynamic)resolved.Range).EntireRow;
            switch (action)
            {
                case ExcelOutlineContract.OutlineAction.Ungroup:
                    ((dynamic)resolved.Range).ClearOutline();
                    break;
                case ExcelOutlineContract.OutlineAction.Collapse:
                    ((dynamic)grouped).ShowDetail = false;
                    break;
                case ExcelOutlineContract.OutlineAction.Expand:
                    ((dynamic)grouped).ShowDetail = true;
                    break;
                default:
                    ((dynamic)grouped).Group();
                    if (level is int outlineLevel)
                        ((dynamic)entire).OutlineLevel = outlineLevel;
                    break;
            }

            if (ExcelOutlineContract.SummaryRow(op) is int summaryRow)
            {
                outline = (object)((dynamic)resolved.Sheet).Outline;
                ((dynamic)outline).SummaryRow = summaryRow;
            }

            if (ExcelOutlineContract.SummaryColumn(op) is int summaryColumn)
            {
                outline ??= (object)((dynamic)resolved.Sheet).Outline;
                ((dynamic)outline).SummaryColumn = summaryColumn;
            }

            var actualLevel = Convert.ToInt32(((dynamic)entire).OutlineLevel, CultureInfo.InvariantCulture);
            var showDetail = true;
            try { showDetail = Convert.ToBoolean(((dynamic)grouped).ShowDetail, CultureInfo.InvariantCulture); }
            catch { /* grouped items without a single ShowDetail */ }
            var hasOutline = actualLevel > 1;
            if (!ExcelOutlineContract.MatchesReadback(action, level, actualLevel, showDetail, hasOutline))
                mismatches.Add($"set_outline {axis} readback {ExcelOutlineContract.Describe(op)}: level={actualLevel} showDetail={showDetail}");
            if (ExcelOutlineContract.SummaryRow(op) is int wantedRow)
            {
                outline ??= (object)((dynamic)resolved.Sheet).Outline;
                var actualSummaryRow = Convert.ToInt32(((dynamic)outline).SummaryRow, CultureInfo.InvariantCulture);
                if (actualSummaryRow != wantedRow)
                    mismatches.Add($"set_outline SummaryRow want {wantedRow}, got {actualSummaryRow}");
            }

            if (ExcelOutlineContract.SummaryColumn(op) is int wantedColumn)
            {
                outline ??= (object)((dynamic)resolved.Sheet).Outline;
                var actualSummaryColumn = Convert.ToInt32(((dynamic)outline).SummaryColumn, CultureInfo.InvariantCulture);
                if (actualSummaryColumn != wantedColumn)
                    mismatches.Add($"set_outline SummaryColumn want {wantedColumn}, got {actualSummaryColumn}");
            }

            checkedItems++;
            execution.Affected.Add(new AffectedRef("outline", $"{axis}:{resolved.Address}"));
        }
        finally
        {
            RotHelper.ReleaseComReference(outline);
            RotHelper.ReleaseComReference(entire);
            RotHelper.ReleaseComReference(grouped);
            RotHelper.ReleaseComReference(resolved.Range);
            RotHelper.ReleaseComReference(resolved.Sheet);
        }
    }

    private static void ApplyImportCsv(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        var path = RequireAbsoluteOutput(Json.GetString(op, "path"), "import_csv");
        if (!File.Exists(path))
            throw new InvalidOperationException($"import_csv path does not exist: {path}");
        var destination = Json.GetString(op, "destination") ?? "A1";
        var delimiter = ExcelCsvContract.ResolveDelimiter(op);
        var parsed = ExcelCsvContract.NormalizeRectangle(
            ExcelCsvContract.Parse(File.ReadAllText(path), delimiter));
        var rows = parsed.Count;
        var columns = rows == 0 ? 0 : parsed[0].Count;
        if (ExcelCsvContract.ImportLimitError(rows, columns) is string limitError)
            throw new InvalidOperationException(limitError);

        var sheet = GetExplicitTargetSheetReference(workbook, op);
        object? range = null;
        try
        {
            if (rows > 0 && columns > 0)
            {
                var matrix = new object[rows, columns];
                for (var r = 0; r < rows; r++)
                    for (var c = 0; c < columns; c++)
                        matrix[r, c] = parsed[r][c];
                range = (object)((dynamic)sheet).Range(destination).Resize(rows, columns);
                ((dynamic)range).NumberFormat = "@";
                ((dynamic)range).Value2 = matrix;

                object? raw = ((dynamic)range).Value2;
                for (var r = 0; r < rows; r++)
                for (var c = 0; c < columns; c++)
                {
                    checkedItems++;
                    var want = parsed[r][c];
                    var got = ToDisp(raw is object[,] arr ? arr[r + 1, c + 1] : raw);
                    if (!string.Equals(want, got, StringComparison.Ordinal))
                        mismatches.Add(ExcelCsvContract.CellReadbackMismatch(
                            Convert.ToInt32(((dynamic)range).Row, CultureInfo.InvariantCulture) + r,
                            Convert.ToInt32(((dynamic)range).Column, CultureInfo.InvariantCulture) + c,
                            want, got));
                }
            }
            else
            {
                mismatches.Add("import_csv produced no rows or columns");
            }

            execution.Affected.Add(new AffectedRef("range", destination));
        }
        finally
        {
            RotHelper.ReleaseComReference(range);
            RotHelper.ReleaseComReference(sheet);
        }
    }

    private static void ApplyExportCsv(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        var output = RequireAbsoluteOutput(Json.GetString(op, "output"), "export_csv");
        if (ExcelAuthoringSchema.SaveConflictsWithExistingFile(output, null, Json.GetBool(op, "overwrite"), File.Exists(output)))
            throw new InvalidOperationException($"export_csv '{output}' exists; set overwrite:true");
        var delimiter = ExcelCsvContract.ResolveDelimiter(op);
        object? sheet = null;
        object? used = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, op);
            var rangeRef = Json.GetString(op, "range");
            used = string.IsNullOrWhiteSpace(rangeRef)
                ? (object)((dynamic)sheet).UsedRange
                : (object)((dynamic)sheet).Range(rangeRef);
            var rows = Convert.ToInt32(((dynamic)used).Rows.Count, CultureInfo.InvariantCulture);
            var columns = Convert.ToInt32(((dynamic)used).Columns.Count, CultureInfo.InvariantCulture);
            if (ExcelCsvContract.ExportLimitError(rows, columns) is string limitError)
                throw new InvalidOperationException(limitError);
            var values = RangeToJson(used, out var cells, maxCells: ExcelCsvContract.MaxExportCells);
            if (cells < rows * columns)
                throw new InvalidOperationException(
                    $"export_csv read {cells} of {rows * columns} cells; refusing a truncated file");
            var table = JsonArrayToStringTable(values);
            Directory.CreateDirectory(Path.GetDirectoryName(output) ?? ".");
            File.WriteAllText(output, ExcelCsvContract.Write(table, delimiter));
            if (!File.Exists(output))
            {
                mismatches.Add($"export_csv did not create '{output}'");
            }
            else
            {
                var roundTrip = ExcelCsvContract.NormalizeRectangle(
                    ExcelCsvContract.Parse(File.ReadAllText(output), delimiter));
                if (!ExcelCsvContract.TablesEqual(table, roundTrip))
                    mismatches.Add("export_csv value equality failed (quoted CRLF, empty, delimiter, or text cells)");
                else
                    checkedItems += Math.Max(1, table.Count * (table.Count == 0 ? 0 : table[0].Count));
            }

            execution.Affected.Add(new AffectedRef("file", output));
        }
        finally
        {
            RotHelper.ReleaseComReference(used);
            RotHelper.ReleaseComReference(sheet);
        }
    }

    private static void ApplyPageBreaks(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        var resolved = ResolveBasicRange(workbook, op);
        try
        {
            if (Json.GetBool(op, "clear"))
            {
                ((dynamic)resolved.Sheet).ResetAllPageBreaks();
                var remaining = CountManualPageBreaks(resolved.Sheet);
                checkedItems++;
                if (remaining > 0)
                    mismatches.Add($"set_page_breaks clear left {remaining} manual break(s)");
            }
            else
            {
                var axis = ExcelPageBreakContract.ResolveAxis(op);
                var index = ExcelPageBreakContract.BreakIndex(resolved.Address, axis);
                if (index < 1)
                    throw new InvalidOperationException($"set_page_breaks '{resolved.Address}' is not a row/column anchor");
                WriteManualPageBreak(resolved.Sheet, resolved.Address, axis, index);
                var actual = CapturePageBreakLocations(resolved.Sheet);
                checkedItems++;
                if (!ExcelPageBreakContract.HasManualBreak(actual, axis, index))
                    mismatches.Add(ExcelPageBreakContract.DescribeMissing(resolved.Address, axis, index, actual));
            }

            execution.Affected.Add(new AffectedRef("range", resolved.Address));
        }
        finally
        {
            RotHelper.ReleaseComReference(resolved.Range);
            RotHelper.ReleaseComReference(resolved.Sheet);
        }
    }

    private static void ApplyDeleteSheet(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        object? sheet = null;
        object? app = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, op);
            var name = Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture);
            app = (object)((dynamic)workbook).Application;
            var alerts = ((dynamic)app).DisplayAlerts;
            try
            {
                ((dynamic)app).DisplayAlerts = false;
                ((dynamic)sheet).Delete();
            }
            finally { try { ((dynamic)app).DisplayAlerts = alerts; } catch { } }

            checkedItems++;
            if (!string.IsNullOrWhiteSpace(name) && SheetExists(workbook, name))
                mismatches.Add($"delete_sheet '{name}' is still present");
            execution.Affected.Add(new AffectedRef("sheet", name ?? ""));
        }
        finally
        {
            RotHelper.ReleaseComReference(app);
            RotHelper.ReleaseComReference(sheet);
        }
    }

    private static int CountManualPageBreaks(object sheet)
    {
        var count = 0;
        object? horizontal = null;
        object? vertical = null;
        try
        {
            horizontal = (object)((dynamic)sheet).HPageBreaks;
            vertical = (object)((dynamic)sheet).VPageBreaks;
            count += Convert.ToInt32(((dynamic)horizontal).Count, CultureInfo.InvariantCulture);
            count += Convert.ToInt32(((dynamic)vertical).Count, CultureInfo.InvariantCulture);
        }
        catch
        {
            return count;
        }
        finally
        {
            RotHelper.ReleaseComReference(vertical);
            RotHelper.ReleaseComReference(horizontal);
        }

        return count;
    }

    private static object? JsonNodeToComValue(JsonNode? node)
    {
        if (node is null) return "";
        if (node is not JsonValue value)
            return node.ToJsonString();
        if (value.TryGetValue<bool>(out var flag)) return flag;
        if (value.TryGetValue<int>(out var integer)) return integer;
        if (value.TryGetValue<long>(out var longInteger) && longInteger is >= int.MinValue and <= int.MaxValue)
            return (int)longInteger;
        if (value.TryGetValue<double>(out var number) && double.IsFinite(number)) return number;
        if (value.TryGetValue<string>(out var text)) return text ?? "";
        return value.ToString();
    }

    private static object ToComArray(JsonArray values)
    {
        var rows = values.Count;
        var cols = values[0] is JsonArray first ? first.Count : 1;
        var matrix = new object[rows, cols];
        for (var r = 0; r < rows; r++)
        {
            var row = values[r] as JsonArray ?? new JsonArray();
            for (var c = 0; c < cols; c++)
                matrix[r, c] = JsonNodeToComValue(row[c]) ?? "";
        }
        return matrix;
    }

    private static object ToObjectMatrix(object[][] rows)
    {
        var matrix = new object[rows.Length, rows[0].Length];
        for (var r = 0; r < rows.Length; r++)
            for (var c = 0; c < rows[0].Length; c++)
                matrix[r, c] = rows[r][c];
        return matrix;
    }

    private static IReadOnlyList<IReadOnlyList<string>> JsonArrayToStringTable(JsonNode? values)
    {
        var table = new List<IReadOnlyList<string>>();
        if (values is not JsonArray rows) return table;
        foreach (var row in rows)
        {
            if (row is not JsonArray cells) continue;
            table.Add(cells.Select(cell =>
            {
                if (cell is null) return "";
                if (cell is JsonValue value)
                {
                    if (value.TryGetValue<string>(out var text)) return text ?? "";
                    if (value.TryGetValue<bool>(out var flag)) return flag ? "TRUE" : "FALSE";
                    return value.ToString();
                }

                return cell.ToString() ?? "";
            }).ToArray());
        }

        return ExcelCsvContract.NormalizeRectangle(table);
    }

    private static JsonObject CaptureExtendedOpsState(object workbook, IReadOnlyList<JsonObject> ops, string? documentRef)
    {
        var entries = new JsonArray();
        foreach (var op in ops)
        {
            var name = Json.GetString(op, "op") ?? "";
            entries.Add(name switch
            {
                "import_csv" => CaptureImportCsvEntry(workbook, op),
                "export_csv" => CaptureExportCsvEntry(op),
                "fill_range" => CaptureFillEntry(workbook, op),
                "auto_fill" => CaptureAutoFillEntry(workbook, op),
                "set_outline" => CaptureOutlineEntry(workbook, op),
                "set_page_breaks" => CapturePageBreakEntry(workbook, op),
                "calculate" => CaptureCalculateEntry(workbook, op),
                _ => new JsonObject { ["op"] = name },
            });
        }

        return new JsonObject
        {
            ["snapshotVersion"] = ExcelLayoutSnapshotVersion,
            ["restoreMode"] = ExtendedOpsRestoreMode,
            ["documentRef"] = documentRef,
            ["originalActiveSheet"] = ReadActiveSheetName(workbook),
            ["entries"] = entries,
        };
    }

    private static JsonObject CaptureImportCsvEntry(object workbook, JsonObject op)
    {
        var path = RequireAbsoluteOutput(Json.GetString(op, "path"), "import_csv");
        var destination = Json.GetString(op, "destination") ?? "A1";
        var delimiter = ExcelCsvContract.ResolveDelimiter(op);
        var parsed = File.Exists(path)
            ? ExcelCsvContract.NormalizeRectangle(ExcelCsvContract.Parse(File.ReadAllText(path), delimiter))
            : Array.Empty<IReadOnlyList<string>>();
        var rows = parsed.Count;
        var columns = rows == 0 ? 0 : parsed[0].Count;
        if (ExcelCsvContract.ImportLimitError(rows, columns) is string limitError)
            throw new InvalidOperationException(limitError);
        object? sheet = null;
        object? range = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, op);
            var sheetName = Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture);
            if (rows == 0 || columns == 0)
            {
                return new JsonObject
                {
                    ["op"] = "import_csv",
                    ["sheet"] = sheetName,
                    ["range"] = destination,
                };
            }

            range = (object)((dynamic)sheet).Range(destination).Resize(rows, columns);
            var address = Convert.ToString(((dynamic)range).Address(false, false), CultureInfo.InvariantCulture)
                          ?? destination;
            return new JsonObject
            {
                ["op"] = "import_csv",
                ["sheet"] = sheetName,
                ["range"] = address,
                ["payload"] = CaptureRangePayload(sheet, address),
                ["numberFormats"] = CaptureNumberFormats(range),
            };
        }
        finally
        {
            RotHelper.ReleaseComReference(range);
            RotHelper.ReleaseComReference(sheet);
        }
    }

    private static JsonObject CaptureExportCsvEntry(JsonObject op)
    {
        var output = RequireAbsoluteOutput(Json.GetString(op, "output"), "export_csv");
        return new JsonObject
        {
            ["op"] = "export_csv",
            ["output"] = output,
            ["existed"] = File.Exists(output),
            ["previous"] = File.Exists(output) ? File.ReadAllText(output) : null,
        };
    }

    private static JsonObject CaptureFillEntry(object workbook, JsonObject op)
    {
        var resolved = ResolveBasicRange(workbook, op);
        try
        {
            var sheetName = Convert.ToString(((dynamic)resolved.Sheet).Name, CultureInfo.InvariantCulture);
            return new JsonObject
            {
                ["op"] = "fill_range",
                ["sheet"] = sheetName,
                ["range"] = resolved.Address,
                ["payload"] = CaptureRangePayload(resolved.Sheet, resolved.Address),
                ["numberFormats"] = CaptureNumberFormats(resolved.Range),
            };
        }
        finally
        {
            RotHelper.ReleaseComReference(resolved.Range);
            RotHelper.ReleaseComReference(resolved.Sheet);
        }
    }

    private static JsonObject CaptureAutoFillEntry(object workbook, JsonObject op)
    {
        var source = ResolveBasicRange(workbook, op);
        object? dest = null;
        try
        {
            dest = (object)((dynamic)source.Sheet).Range(Json.GetString(op, "destRange"));
            var destAddress = Convert.ToString(((dynamic)dest).Address(false, false), CultureInfo.InvariantCulture)
                              ?? Json.GetString(op, "destRange");
            var sheetName = Convert.ToString(((dynamic)source.Sheet).Name, CultureInfo.InvariantCulture);
            return new JsonObject
            {
                ["op"] = "auto_fill",
                ["sheet"] = sheetName,
                ["range"] = destAddress,
                ["payload"] = CaptureRangePayload(source.Sheet, destAddress!),
                ["numberFormats"] = CaptureNumberFormats(dest),
            };
        }
        finally
        {
            RotHelper.ReleaseComReference(dest);
            RotHelper.ReleaseComReference(source.Range);
            RotHelper.ReleaseComReference(source.Sheet);
        }
    }

    private static JsonObject CaptureOutlineEntry(object workbook, JsonObject op)
    {
        var resolved = ResolveBasicRange(workbook, op);
        object? outline = null;
        try
        {
            outline = (object)((dynamic)resolved.Sheet).Outline;
            return new JsonObject
            {
                ["op"] = "set_outline",
                ["sheet"] = Convert.ToString(((dynamic)resolved.Sheet).Name, CultureInfo.InvariantCulture),
                ["range"] = resolved.Address,
                ["axis"] = ExcelOutlineContract.ResolveAxis(op) == ExcelOutlineContract.OutlineAxis.Columns
                    ? "column"
                    : "row",
                ["rowLevels"] = CaptureOutlineLevels(resolved.Range, rows: true),
                ["columnLevels"] = CaptureOutlineLevels(resolved.Range, rows: false),
                ["summaryRow"] = Convert.ToInt32(((dynamic)outline).SummaryRow, CultureInfo.InvariantCulture),
                ["summaryColumn"] = Convert.ToInt32(((dynamic)outline).SummaryColumn, CultureInfo.InvariantCulture),
            };
        }
        finally
        {
            RotHelper.ReleaseComReference(outline);
            RotHelper.ReleaseComReference(resolved.Range);
            RotHelper.ReleaseComReference(resolved.Sheet);
        }
    }

    private static JsonArray CaptureOutlineLevels(object range, bool rows)
    {
        var levels = new JsonArray();
        object? collection = null;
        try
        {
            collection = rows ? (object)((dynamic)range).Rows : (object)((dynamic)range).Columns;
            var count = Convert.ToInt32(((dynamic)collection).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                object? item = null;
                object? entire = null;
                try
                {
                    item = (object)((dynamic)collection).Item(index);
                    entire = rows ? (object)((dynamic)item).EntireRow : (object)((dynamic)item).EntireColumn;
                    var showDetail = true;
                    try { showDetail = Convert.ToBoolean(((dynamic)item).ShowDetail, CultureInfo.InvariantCulture); }
                    catch { /* no grouped detail */ }
                    levels.Add(new JsonObject
                    {
                        ["index"] = rows
                            ? Convert.ToInt32(((dynamic)item).Row, CultureInfo.InvariantCulture)
                            : Convert.ToInt32(((dynamic)item).Column, CultureInfo.InvariantCulture),
                        ["level"] = Convert.ToInt32(((dynamic)entire).OutlineLevel, CultureInfo.InvariantCulture),
                        ["showDetail"] = showDetail,
                    });
                }
                finally
                {
                    RotHelper.ReleaseComReference(entire);
                    RotHelper.ReleaseComReference(item);
                }
            }
        }
        finally { RotHelper.ReleaseComReference(collection); }
        return levels;
    }

    private static JsonObject CapturePageBreakEntry(object workbook, JsonObject op)
    {
        var resolved = ResolveBasicRange(workbook, op);
        try
        {
            return new JsonObject
            {
                ["op"] = "set_page_breaks",
                ["sheet"] = Convert.ToString(((dynamic)resolved.Sheet).Name, CultureInfo.InvariantCulture),
                ["range"] = resolved.Address,
                ["breaks"] = CapturePageBreakLocations(resolved.Sheet),
            };
        }
        finally
        {
            RotHelper.ReleaseComReference(resolved.Range);
            RotHelper.ReleaseComReference(resolved.Sheet);
        }
    }

    private static JsonObject CapturePageBreakLocations(object sheet)
    {
        var horizontal = new JsonArray();
        var vertical = new JsonArray();
        object? hBreaks = null;
        object? vBreaks = null;
        try
        {
            hBreaks = (object)((dynamic)sheet).HPageBreaks;
            vBreaks = (object)((dynamic)sheet).VPageBreaks;
            var hCount = Convert.ToInt32(((dynamic)hBreaks).Count, CultureInfo.InvariantCulture);
            for (var i = 1; i <= hCount; i++)
            {
                object? item = null;
                object? location = null;
                try
                {
                    item = (object)((dynamic)hBreaks).Item(i);
                    location = (object)((dynamic)item).Location;
                    horizontal.Add(Convert.ToInt32(((dynamic)location).Row, CultureInfo.InvariantCulture));
                }
                catch { /* skip one break */ }
                finally
                {
                    RotHelper.ReleaseComReference(location);
                    RotHelper.ReleaseComReference(item);
                }
            }

            var vCount = Convert.ToInt32(((dynamic)vBreaks).Count, CultureInfo.InvariantCulture);
            for (var i = 1; i <= vCount; i++)
            {
                object? item = null;
                object? location = null;
                try
                {
                    item = (object)((dynamic)vBreaks).Item(i);
                    location = (object)((dynamic)item).Location;
                    vertical.Add(Convert.ToInt32(((dynamic)location).Column, CultureInfo.InvariantCulture));
                }
                catch { /* skip one break */ }
                finally
                {
                    RotHelper.ReleaseComReference(location);
                    RotHelper.ReleaseComReference(item);
                }
            }
        }
        finally
        {
            RotHelper.ReleaseComReference(vBreaks);
            RotHelper.ReleaseComReference(hBreaks);
        }

        return new JsonObject { ["horizontal"] = horizontal, ["vertical"] = vertical };
    }

    private static JsonObject CaptureCalculateEntry(object workbook, JsonObject op)
    {
        if (string.IsNullOrWhiteSpace(Json.GetString(op, "range")))
            return new JsonObject { ["op"] = "calculate", ["scope"] = "workbook" };

        var resolved = ResolveBasicRange(workbook, op);
        try
        {
            var sheetName = Convert.ToString(((dynamic)resolved.Sheet).Name, CultureInfo.InvariantCulture);
            return new JsonObject
            {
                ["op"] = "calculate",
                ["sheet"] = sheetName,
                ["range"] = resolved.Address,
                ["payload"] = CaptureRangePayload(resolved.Sheet, resolved.Address),
            };
        }
        finally
        {
            RotHelper.ReleaseComReference(resolved.Range);
            RotHelper.ReleaseComReference(resolved.Sheet);
        }
    }

    private static JsonObject RestoreExtendedOpsState(object workbook, JsonObject state)
    {
        var mismatches = new RestoreMismatchCollector();
        var checkedItems = 0;
        var entries = Json.GetArr(state, "entries") ?? new JsonArray();
        for (var index = entries.Count - 1; index >= 0; index--)
        {
            if (entries[index] is not JsonObject entry) continue;
            try
            {
                checkedItems += RestoreExtendedEntry(workbook, entry, mismatches);
            }
            catch (Exception ex)
            {
                mismatches.Add($"extended restore {Json.GetString(entry, "op")} failed: {ex.Message}");
            }
        }

        var active = Json.GetString(state, "originalActiveSheet");
        if (!string.IsNullOrWhiteSpace(active)) ActivateWorksheet(workbook, active, mismatches);
        return BuildRestoreResult(mismatches.Count == 0, 0, checkedItems, ExtendedOpsRestoreMode, mismatches);
    }

    private static int RestoreExtendedEntry(object workbook, JsonObject entry, RestoreMismatchCollector mismatches)
    {
        var name = Json.GetString(entry, "op") ?? "";
        if (name == "export_csv")
        {
            var output = Json.GetString(entry, "output");
            if (string.IsNullOrWhiteSpace(output)) return 0;
            if (Json.GetBool(entry, "existed"))
                File.WriteAllText(output, Json.GetString(entry, "previous") ?? "");
            else if (File.Exists(output))
                File.Delete(output);
            return 1;
        }

        if (name == "calculate" && string.Equals(Json.GetString(entry, "scope"), "workbook", StringComparison.Ordinal))
            return 0;

        object? sheet = null;
        try
        {
            var sheetName = Json.GetString(entry, "sheet");
            if (string.IsNullOrWhiteSpace(sheetName)) return 0;
            sheet = GetExplicitTargetSheetReference(workbook, SheetTargetOp(sheetName));
            var address = Json.GetString(entry, "range");
            var checkedItems = 0;
            if (!string.IsNullOrWhiteSpace(address) && entry["payload"] is JsonObject)
                checkedItems += RestoreRangePayload(sheet, sheetName, address, Json.GetObj(entry, "payload"), mismatches);
            if (!string.IsNullOrWhiteSpace(address) && entry["numberFormats"] is JsonArray formats)
            {
                RestoreNumberFormats(sheet, address, formats);
                checkedItems += formats.Count;
            }

            if (name == "set_outline")
                checkedItems += RestoreOutlineEntry(sheet, entry, mismatches);
            if (name == "set_page_breaks" && entry["breaks"] is JsonObject breaks)
                checkedItems += RestorePageBreakLocations(sheet, breaks, mismatches);
            return checkedItems;
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static int RestoreOutlineEntry(object sheet, JsonObject entry, RestoreMismatchCollector mismatches)
    {
        var checkedItems = 0;
        object? outline = null;
        try
        {
            outline = (object)((dynamic)sheet).Outline;
            if (entry.ContainsKey("summaryRow"))
                ((dynamic)outline).SummaryRow = Json.GetInt(entry, "summaryRow") ?? ExcelOutlineContract.XlSummaryBelow;
            if (entry.ContainsKey("summaryColumn"))
                ((dynamic)outline).SummaryColumn = Json.GetInt(entry, "summaryColumn") ?? ExcelOutlineContract.XlSummaryOnRight;
            checkedItems += RestoreOutlineLevels(sheet, Json.GetArr(entry, "rowLevels"), rows: true, mismatches);
            checkedItems += RestoreOutlineLevels(sheet, Json.GetArr(entry, "columnLevels"), rows: false, mismatches);
        }
        finally { RotHelper.ReleaseComReference(outline); }
        return checkedItems;
    }

    private static int RestoreOutlineLevels(object sheet, JsonArray? levels, bool rows, RestoreMismatchCollector mismatches)
    {
        if (levels is null) return 0;
        var checkedItems = 0;
        object? collection = null;
        try
        {
            collection = rows ? (object)((dynamic)sheet).Rows : (object)((dynamic)sheet).Columns;
            foreach (var node in levels.OfType<JsonObject>())
            {
                var index = Json.GetInt(node, "index");
                if (index is null) continue;
                object? item = null;
                object? entire = null;
                try
                {
                    item = (object)((dynamic)collection).Item(index.Value);
                    entire = rows ? (object)((dynamic)item).EntireRow : (object)((dynamic)item).EntireColumn;
                    ((dynamic)entire).OutlineLevel = Json.GetInt(node, "level") ?? 1;
                    if (node.ContainsKey("showDetail"))
                    {
                        try { ((dynamic)item).ShowDetail = Json.GetBool(node, "showDetail"); }
                        catch { /* not grouped */ }
                    }

                    checkedItems++;
                    var actual = Convert.ToInt32(((dynamic)entire).OutlineLevel, CultureInfo.InvariantCulture);
                    if (actual != (Json.GetInt(node, "level") ?? 1))
                        mismatches.Add($"{(rows ? "row" : "col")} {index}: outline level restore mismatch");
                }
                finally
                {
                    RotHelper.ReleaseComReference(entire);
                    RotHelper.ReleaseComReference(item);
                }
            }
        }
        finally { RotHelper.ReleaseComReference(collection); }
        return checkedItems;
    }

    private static void WriteManualPageBreak(object sheet, string address, ExcelPageBreakContract.BreakAxis axis, int index)
    {
        object? entire = null;
        object? breaks = null;
        try
        {
            entire = axis == ExcelPageBreakContract.BreakAxis.Columns
                ? (object)((dynamic)sheet).Columns[index]
                : (object)((dynamic)sheet).Rows[index];
            try
            {
                ((dynamic)entire).PageBreak = ExcelPageBreakContract.XlPageBreakManual;
                return;
            }
            catch
            {
                breaks = axis == ExcelPageBreakContract.BreakAxis.Columns
                    ? (object)((dynamic)sheet).VPageBreaks
                    : (object)((dynamic)sheet).HPageBreaks;
                ((dynamic)breaks).Add(entire);
            }
        }
        finally
        {
            RotHelper.ReleaseComReference(breaks);
            RotHelper.ReleaseComReference(entire);
        }
    }

    private static int RestorePageBreakLocations(object sheet, JsonObject breaks, RestoreMismatchCollector mismatches)
    {
        ((dynamic)sheet).ResetAllPageBreaks();
        var checkedItems = 0;
        foreach (var node in Json.GetArr(breaks, "horizontal") ?? new JsonArray())
        {
            if (node is not JsonValue value || !value.TryGetValue<int>(out var row)) continue;
            WriteManualPageBreak(sheet, "A" + row.ToString(CultureInfo.InvariantCulture), ExcelPageBreakContract.BreakAxis.Rows, row);
            checkedItems++;
        }

        foreach (var node in Json.GetArr(breaks, "vertical") ?? new JsonArray())
        {
            if (node is not JsonValue value || !value.TryGetValue<int>(out var column)) continue;
            WriteManualPageBreak(sheet, ColName(column) + "1", ExcelPageBreakContract.BreakAxis.Columns, column);
            checkedItems++;
        }

        var actual = CapturePageBreakLocations(sheet);
        if (!PageBreakLocationsEqual(breaks, actual))
            mismatches.Add("set_page_breaks restore mismatch");
        return Math.Max(1, checkedItems);
    }

    private static bool PageBreakLocationsEqual(JsonObject left, JsonObject right)
    {
        static List<int> Read(JsonObject obj, string key) =>
            (Json.GetArr(obj, key) ?? new JsonArray())
            .Select(node => node is JsonValue value && value.TryGetValue<int>(out var n) ? n : 0)
            .Where(n => n > 0)
            .OrderBy(n => n)
            .ToList();

        return Read(left, "horizontal").SequenceEqual(Read(right, "horizontal")) &&
               Read(left, "vertical").SequenceEqual(Read(right, "vertical"));
    }

    private static JsonObject CaptureFormula2WriteState(object workbook, IReadOnlyList<JsonObject> ops, string? documentRef)
    {
        var entries = new JsonArray();
        foreach (var op in ops)
        {
            if (!string.Equals(Json.GetString(op, "op"), "set_formulas", StringComparison.OrdinalIgnoreCase))
                continue;
            var resolved = ResolveBasicRange(workbook, op);
            object? captureRange = null;
            try
            {
                // Keep named tuples statically typed across the COM dynamic boundary.
                int destRows = Convert.ToInt32(((dynamic)resolved.Range).Rows.Count, CultureInfo.InvariantCulture);
                int destCols = Convert.ToInt32(((dynamic)resolved.Range).Columns.Count, CultureInfo.InvariantCulture);
                (int Rows, int Columns) footprint = ExcelFormula2Contract.WritesFormula2(op)
                    ? ExcelFormula2Contract.ResolveCaptureFootprint(op, destRows, destCols)
                    : (destRows, destCols);
                var oldSpill = TryReadSpillingToRange(resolved.Range);
                footprint = ExcelFormula2Contract.UnionFootprint(
                    resolved.Address, footprint.Rows, footprint.Columns, oldSpill);
                captureRange = (object)((dynamic)resolved.Range).Resize(footprint.Rows, footprint.Columns);
                var address = Convert.ToString(((dynamic)captureRange).Address(false, false), CultureInfo.InvariantCulture)
                              ?? resolved.Address;
                var sheetName = Convert.ToString(((dynamic)resolved.Sheet).Name, CultureInfo.InvariantCulture);
                entries.Add(new JsonObject
                {
                    ["sheet"] = sheetName,
                    ["range"] = address,
                    ["dest"] = resolved.Address,
                    ["engine"] = ExcelFormula2Contract.ResolveEngine(op),
                    ["capturedRows"] = footprint.Rows,
                    ["capturedColumns"] = footprint.Columns,
                    ["payload"] = CaptureRangePayload(resolved.Sheet, address),
                    ["numberFormats"] = CaptureNumberFormats(captureRange),
                    ["formula2"] = RangeToJson(captureRange, out _, formulas: true, formula2: true, maxCells: MaxSnapshotCells),
                    ["oldSpill"] = oldSpill,
                    ["neighborBlockers"] = NeighborBlockerAddresses(resolved.Address, address),
                });
            }
            finally
            {
                RotHelper.ReleaseComReference(captureRange);
                RotHelper.ReleaseComReference(resolved.Range);
                RotHelper.ReleaseComReference(resolved.Sheet);
            }
        }

        return new JsonObject
        {
            ["snapshotVersion"] = ExcelLayoutSnapshotVersion,
            ["restoreMode"] = ExcelFormula2Contract.RestoreMode,
            ["documentRef"] = documentRef,
            ["originalActiveSheet"] = ReadActiveSheetName(workbook),
            ["entries"] = entries,
        };
    }

    private static JsonObject RestoreFormula2WriteState(object workbook, JsonObject state)
    {
        var mismatches = new RestoreMismatchCollector();
        var checkedItems = 0;
        checkedItems += ClearNewlyOwnedFormula2Spills(workbook, state);
        foreach (var node in Json.GetArr(state, "entries") ?? new JsonArray())
        {
            if (node is not JsonObject entry) continue;
            object? sheet = null;
            try
            {
                var sheetName = Json.GetString(entry, "sheet");
                sheet = GetExplicitTargetSheetReference(workbook, SheetTargetOp(sheetName));
                var address = Json.GetString(entry, "range")!;
                checkedItems += RestoreRangePayload(sheet, sheetName, address, Json.GetObj(entry, "payload"), mismatches);
                if (entry["numberFormats"] is JsonArray formats)
                    RestoreNumberFormats(sheet, address, formats);
                if (entry["formula2"] is JsonNode formula2)
                    checkedItems += WriteFormulaGrid(sheet, address, formula2, formula2Engine: true);
                VerifyFormula2Restored(sheet, entry, mismatches);
            }
            catch (Exception ex)
            {
                mismatches.Add($"formula2 restore failed: {ex.Message}");
            }
            finally { RotHelper.ReleaseComReference(sheet); }
        }

        var active = Json.GetString(state, "originalActiveSheet");
        if (!string.IsNullOrWhiteSpace(active)) ActivateWorksheet(workbook, active, mismatches);
        return BuildRestoreResult(mismatches.Count == 0, 0, checkedItems, ExcelFormula2Contract.RestoreMode, mismatches);
    }

    private static int WriteFormulaGrid(object sheet, string address, JsonNode? grid, bool formula2Engine)
    {
        if (grid is not JsonArray rows || rows.Count == 0 || rows[0] is not JsonArray first)
            return 0;
        object? range = null;
        try
        {
            var height = rows.Count;
            var width = first.Count;
            var data = new object?[height, width];
            for (var r = 0; r < height; r++)
            {
                if (rows[r] is not JsonArray row) continue;
                for (var c = 0; c < width && c < row.Count; c++)
                    data[r, c] = NodeToComValue(row[c]);
            }

            range = (object)((dynamic)sheet).Range(address);
            if (formula2Engine)
                ((dynamic)range).Formula2 = data;
            else
                ((dynamic)range).Formula = data;
            return height * width;
        }
        finally { RotHelper.ReleaseComReference(range); }
    }

    private static int ClearNewlyOwnedFormula2Spills(object workbook, JsonObject state)
    {
        var checkedItems = 0;
        foreach (var node in Json.GetArr(state, "formula2Spills") ?? new JsonArray())
        {
            if (node is not JsonObject entry) continue;
            var dest = ExcelFormula2Contract.CleanA1(Json.GetString(entry, "dest"));
            var captured = ExcelFormula2Contract.CleanA1(Json.GetString(entry, "capturedRange"));
            if (string.IsNullOrWhiteSpace(captured))
            {
                var capturedRows = Json.GetInt(entry, "capturedRows") ?? 1;
                var capturedCols = Json.GetInt(entry, "capturedColumns") ?? 1;
                captured = ExcelFormula2Contract.ResizeA1(dest, capturedRows, capturedCols);
            }
            var spill = ExcelFormula2Contract.CleanA1(Json.GetString(entry, "spillRange"));
            bool? hasSpill = entry.ContainsKey("hasSpill") ? Json.GetBool(entry, "hasSpill") : true;
            if (!string.IsNullOrWhiteSpace(dest))
            {
                object? destSheet = null;
                object? destRange = null;
                try
                {
                    destSheet = GetExplicitTargetSheetReference(workbook, SheetTargetOp(Json.GetString(entry, "sheet")));
                    destRange = (object)((dynamic)destSheet).Range(dest);
                    ((dynamic)destRange).ClearContents();
                    checkedItems++;
                }
                catch { /* dest may already be gone */ }
                finally
                {
                    RotHelper.ReleaseComReference(destRange);
                    RotHelper.ReleaseComReference(destSheet);
                }
            }

            if (ExcelFormula2Contract.MustNotClearNeighbors(hasSpill) ||
                !ExcelFormula2Contract.ShouldRecordSpill(hasSpill, spill))
                continue;
            foreach (var extra in ExcelFormula2Contract.NewlyOwnedSpillRanges(captured, spill))
            {
                object? sheet = null;
                object? range = null;
                try
                {
                    sheet = GetExplicitTargetSheetReference(workbook, SheetTargetOp(Json.GetString(entry, "sheet")));
                    range = (object)((dynamic)sheet).Range(extra);
                    ((dynamic)range).ClearContents();
                    checkedItems++;
                }
                catch { /* newly owned spill cell may already be gone */ }
                finally
                {
                    RotHelper.ReleaseComReference(range);
                    RotHelper.ReleaseComReference(sheet);
                }
            }
        }

        return checkedItems;
    }

    private static void RecordFormula2Spill(string? sheetName, string dest, string? spillRange, int capturedRows, int capturedColumns,
        bool hasSpill = true)
    {
        PatchSnapshotState(_lastSnapshotDir, state =>
        {
            var spills = Json.GetArr(state, "formula2Spills") ?? new JsonArray();
            spills.Add(new JsonObject
            {
                ["sheet"] = sheetName,
                ["dest"] = dest,
                ["capturedRange"] = ExcelFormula2Contract.ResizeA1(dest, capturedRows, capturedColumns),
                ["spillRange"] = spillRange,
                ["hasSpill"] = hasSpill,
                ["capturedRows"] = capturedRows,
                ["capturedColumns"] = capturedColumns,
            });
            state["formula2Spills"] = spills;
        });
    }

    private static void ClearFormula2SpillFootprints(object workbook, JsonObject state)
    {
        foreach (var node in Json.GetArr(state, "formula2Spills") ?? new JsonArray())
        {
            if (node is not JsonObject entry) continue;
            var spill = Json.GetString(entry, "spillRange");
            var dest = Json.GetString(entry, "dest");
            if (!ExcelFormula2Contract.ShouldRecordSpill(
                    entry.ContainsKey("hasSpill") ? Json.GetBool(entry, "hasSpill") : true,
                    spill) ||
                string.Equals(spill, dest, StringComparison.OrdinalIgnoreCase))
                continue;
            object? sheet = null;
            object? range = null;
            try
            {
                sheet = GetExplicitTargetSheetReference(workbook, SheetTargetOp(Json.GetString(entry, "sheet")));
                range = (object)((dynamic)sheet).Range(spill);
                ((dynamic)range).ClearContents();
            }
            catch { /* spill cell may already be gone */ }
            finally
            {
                RotHelper.ReleaseComReference(range);
                RotHelper.ReleaseComReference(sheet);
            }
        }
    }

    private static JsonObject CaptureFormula2PublicReadback(object range, string address, string engine)
    {
        bool? hasSpill = null;
        string? spillRange = null;
        string? formula2 = null;
        JsonNode? spillValues = null;
        object? spill = null;
        try
        {
            try { formula2 = Convert.ToString(((dynamic)range).Formula2, CultureInfo.InvariantCulture); }
            catch { formula2 = Convert.ToString(((dynamic)range).Formula, CultureInfo.InvariantCulture); }
            try { hasSpill = ExcelFormula2Contract.HasSpillFromCom((object?)((dynamic)range).HasSpill); }
            catch { hasSpill = null; }
            if (hasSpill == true)
            {
                try
                {
                    try { spill = (object)((dynamic)range).SpillingToRange; }
                    catch { spill = null; }
                    spillRange = ExcelFormula2Contract.AcceptSpillAddress(
                        address,
                        spill is null
                            ? null
                            : Convert.ToString(((dynamic)spill).Address(false, false), CultureInfo.InvariantCulture),
                        spillingToRangeOk: spill is not null);
                    if (spill is not null)
                        spillValues = RangeToJson(spill, out _, maxCells: MaxSnapshotCells);
                }
                catch { spillRange = null; }
            }
        }
        finally { RotHelper.ReleaseComReference(spill); }

        return ExcelFormula2Contract.SpillReadback(
            ExcelFormula2Contract.CleanA1(address), engine, hasSpill, spillRange, formula2,
            formulas: null, spillValues: spillValues);
    }

    private static string? TryReadSpillingToRange(object range)
    {
        object? spill = null;
        try
        {
            bool? hasSpill;
            try { hasSpill = ExcelFormula2Contract.HasSpillFromCom((object?)((dynamic)range).HasSpill); }
            catch (System.Runtime.InteropServices.COMException) { hasSpill = null; }
            if (hasSpill != true)
                return null;
            try { spill = (object)((dynamic)range).SpillingToRange; }
            catch (System.Runtime.InteropServices.COMException) { return null; }

            return ExcelFormula2Contract.AcceptSpillAddress(
                Convert.ToString(((dynamic)range).Address(false, false), CultureInfo.InvariantCulture) ?? "",
                Convert.ToString(((dynamic)spill).Address(false, false), CultureInfo.InvariantCulture),
                spillingToRangeOk: spill is not null);
        }
        finally { RotHelper.ReleaseComReference(spill); }
    }

    private static JsonArray NeighborBlockerAddresses(string dest, string capturedRange)
    {
        var neighbors = new JsonArray();
        foreach (var extra in ExcelFormula2Contract.NeighborAndOldSpillRanges(dest, capturedRange))
            neighbors.Add(extra);
        return neighbors;
    }

    private static void VerifyFormula2Restored(object sheet, JsonObject entry, RestoreMismatchCollector mismatches)
    {
        var dest = ExcelFormula2Contract.CleanA1(Json.GetString(entry, "dest") ?? Json.GetString(entry, "range"));
        if (string.IsNullOrWhiteSpace(dest))
            return;
        object? range = null;
        try
        {
            range = (object)((dynamic)sheet).Range(dest);
            string? actual = null;
            try { actual = Convert.ToString(((dynamic)range).Formula2, CultureInfo.InvariantCulture); }
            catch { actual = Convert.ToString(((dynamic)range).Formula, CultureInfo.InvariantCulture); }
            var expected = FirstFormula2Cell(entry["formula2"]);
            if (!string.IsNullOrWhiteSpace(expected) &&
                !ExcelFormula2Contract.RestoredAnchorMatches(expected, actual))
                mismatches.Add($"formula2 dest {dest} restore mismatch: want '{expected}', got '{actual}'");

            var payload = Json.GetObj(entry, "payload");
            var expectedValues = payload?["values"];
            var capturedAddress = ExcelFormula2Contract.CleanA1(Json.GetString(entry, "range") ?? dest);
            if (expectedValues is not null && !string.IsNullOrWhiteSpace(capturedAddress))
            {
                object? captured = null;
                try
                {
                    captured = (object)((dynamic)sheet).Range(capturedAddress);
                    var actualValues = RangeToJson(captured, out _, maxCells: MaxSnapshotCells);
                    if (!ExcelFormula2Contract.RestoredValuesMatch(expectedValues, actualValues))
                        mismatches.Add($"formula2 values {capturedAddress} restore mismatch");
                }
                finally { RotHelper.ReleaseComReference(captured); }
            }

            foreach (var neighbor in Json.GetArr(entry, "neighborBlockers") ?? new JsonArray())
            {
                var address = neighbor is JsonValue value
                    ? value.ToString()
                    : neighbor is JsonObject obj ? Json.GetString(obj, "range") : null;
                if (string.IsNullOrWhiteSpace(address))
                    continue;
                object? neighborRange = null;
                try
                {
                    neighborRange = (object)((dynamic)sheet).Range(address);
                    string? got = null;
                    try { got = Convert.ToString(((dynamic)neighborRange).Formula2, CultureInfo.InvariantCulture); }
                    catch { got = Convert.ToString(((dynamic)neighborRange).Formula, CultureInfo.InvariantCulture); }
                    if (ExcelDeleteSheetDependencyContract.IsBrokenRef(got))
                        mismatches.Add($"formula2 neighbor {address} is #REF! after restore");
                }
                finally { RotHelper.ReleaseComReference(neighborRange); }
            }
        }
        finally { RotHelper.ReleaseComReference(range); }
    }

    private static string? FirstFormula2Cell(JsonNode? grid)
    {
        if (grid is JsonArray rows && rows.Count > 0 && rows[0] is JsonArray first && first.Count > 0)
            return first[0]?.ToString();
        return grid?.ToString();
    }
}
