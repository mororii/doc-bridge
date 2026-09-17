using System.Collections;
using System.Globalization;
using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters
{
public sealed partial class ExcelAdapter
{
    internal sealed class FormulaTraceReadContext
    {
        private readonly dynamic _workbook;
        private readonly string _workbookPath;
        private readonly Dictionary<string, ExcelFormulaTraceContract.Cell> _cells = new(StringComparer.OrdinalIgnoreCase);
        private readonly Func<string, ExcelA1Box, IReadOnlyList<ExcelFormulaTraceContract.Cell>>? _reader;

        internal FormulaTraceReadContext(object workbook, string workbookPath) { _workbook = workbook; _workbookPath = workbookPath; }
        internal FormulaTraceReadContext(Func<string, ExcelA1Box, IReadOnlyList<ExcelFormulaTraceContract.Cell>> reader)
        { _reader = reader; _workbook = null!; _workbookPath = ""; }

        internal IReadOnlyList<ExcelFormulaTraceContract.Cell> Read(string sheetName, ExcelA1Box requested, long limit)
        {
            var result = new List<ExcelFormulaTraceContract.Cell>();
            foreach (var planned in ExcelFormulaTraceReadPlan.Plan(requested, limit))
            {
                foreach (var missing in MissingRuns(sheetName, planned)) ReadRectangle(sheetName, missing);
                for (var row = planned.Row; row <= planned.EndRow; row++)
                    for (var col = planned.Column; col <= planned.EndColumn; col++)
                        result.Add(_cells[Key(sheetName, ExcelA1Box.CellName(col, row))]);
            }
            return result;
        }

        private IEnumerable<ExcelA1Box> MissingRuns(string sheetName, ExcelA1Box box)
        {
            return ExcelFormulaTraceReadPlan.MissingRectangles(box,
                (row, col) => _cells.ContainsKey(Key(sheetName, ExcelA1Box.CellName(col, row))));
        }

        private void ReadRectangle(string sheetName, ExcelA1Box box)
        {
            if (_reader is not null)
            {
                foreach (var item in _reader(sheetName, box)) _cells[Key(item.Sheet, item.Address)] = item;
                return;
            }
            object? sheet = null; object? range = null; object? application = null; object? worksheetFunction = null;
            try
            {
                sheet = GetSheet(_workbook, sheetName);
                range = (object)((dynamic)sheet).Range(box.Address);
                dynamic nativeRange = range;
                object? formulas = null; object? values = null;
                var formulaRead = "formula2";
                var formulaSucceeded = false;
                try { formulas = nativeRange.Formula2; formulaSucceeded = true; }
                catch (Exception formula2Error)
                {
                    try { formulas = nativeRange.Formula; formulaSucceeded = true; formulaRead = "formula (Formula2 fallback: " + formula2Error.GetType().Name + ")"; }
                    catch (Exception formulaError) { formulaRead = "unavailable (Formula2: " + formula2Error.GetType().Name + "; Formula: " + formulaError.GetType().Name + ")"; }
                }
                var valueSucceeded = false;
                try { values = nativeRange.Value2; valueSucceeded = true; } catch { }
                try { application = (object)nativeRange.Application; worksheetFunction = (object)((dynamic)application).WorksheetFunction; } catch { }
                for (var row = box.Row; row <= box.EndRow; row++)
                for (var col = box.Column; col <= box.EndColumn; col++)
                {
                    var address = ExcelA1Box.CellName(col, row);
                    object? cell = null;
                    try
                    {
                        var formulaMapped = ExcelFormulaTraceReadPlan.TryValueAt(formulas, row - box.Row + 1, col - box.Column + 1, box.Rows, box.Columns, out var formulaValue);
                        var valueMapped = ExcelFormulaTraceReadPlan.TryValueAt(values, row - box.Row + 1, col - box.Column + 1, box.Rows, box.Columns, out var value);
                        var formula = formulaValue is string text && text.StartsWith("=", StringComparison.Ordinal) ? text : null;
                        bool? isError = null; string? display = null;
                        try
                        {
                            cell = (object)((dynamic)sheet).Range(address);
                            if (worksheetFunction is not null) isError = Convert.ToBoolean(((dynamic)worksheetFunction).IsError(cell), CultureInfo.InvariantCulture);
                            if (formula is not null || isError is true) display = TryString(() => ((dynamic)cell).Text);
                        }
                        catch { }
                        var evidence = new JsonObject
                        {
                            ["workbook"] = _workbookPath, ["sheet"] = sheetName, ["address"] = address,
                            ["value"] = FormulaTraceValue(value), ["formula"] = formula, ["formulaRead"] = formulaRead,
                            ["errorClassification"] = isError switch { true => "excel_error", false => "not_excel_error", _ => "unknown" },
                            ["readEvidenceComplete"] = formulaSucceeded && formulaMapped && valueSucceeded && valueMapped && isError is not null,
                        };
                        if (!string.IsNullOrWhiteSpace(display) && (formula is not null || isError is true)) evidence["displayText"] = display;
                        if (isError is true && !string.IsNullOrWhiteSpace(display)) evidence["errorDisplay"] = display;
                        _cells.Add(Key(sheetName, address), new ExcelFormulaTraceContract.Cell(sheetName, address, formula, evidence));
                    }
                    finally { RotHelper.ReleaseComReference(cell); }
                }
            }
            finally { RotHelper.ReleaseComReference(worksheetFunction); RotHelper.ReleaseComReference(application); RotHelper.ReleaseComReference(range); RotHelper.ReleaseComReference(sheet); }
        }
        private static string Key(string sheet, string address) => sheet + "!" + address;
    }
}
}

namespace DocBridge.Core.Services
{
internal static class ExcelFormulaTraceReadPlan
{
    internal static IReadOnlyList<ExcelA1Box> MissingRectangles(ExcelA1Box box, Func<int, int, bool> isCached)
    {
        var completed = new List<ExcelA1Box>();
        var active = new Dictionary<(int Column, int Width), ExcelA1Box>();
        for (var row = box.Row; row <= box.EndRow; row++)
        {
            var next = new Dictionary<(int Column, int Width), ExcelA1Box>();
            var col = box.Column;
            while (col <= box.EndColumn)
            {
                if (isCached(row, col)) { col++; continue; }
                var first = col++;
                while (col <= box.EndColumn && !isCached(row, col)) col++;
                var key = (first, col - first);
                next[key] = active.Remove(key, out var previous)
                    ? new ExcelA1Box(previous.Row, first, previous.Rows + 1, col - first)
                    : new ExcelA1Box(row, first, 1, col - first);
            }
            completed.AddRange(active.Values);
            active = next;
        }
        completed.AddRange(active.Values);
        return completed;
    }

    internal static bool TryValueAt(object? source, int row, int column, int rows, int columns, out object? value)
    {
        value = null;
        if (row < 1 || column < 1 || row > rows || column > columns) return false;
        if (source is not Array values)
        {
            if (rows != 1 || columns != 1) return false;
            value = source;
            return true;
        }
        if (values.Rank != 2 || values.GetLength(0) != rows || values.GetLength(1) != columns) return false;
        value = values.GetValue(values.GetLowerBound(0) + row - 1, values.GetLowerBound(1) + column - 1);
        return true;
    }

    internal static IReadOnlyList<ExcelA1Box> Plan(ExcelA1Box requested, long maximumCells)
    {
        var count = Math.Min(requested.CellCount, Math.Max(0, maximumCells));
        if (count == 0) return Array.Empty<ExcelA1Box>();
        var fullRows = (int)(count / requested.Columns);
        var tail = (int)(count % requested.Columns);
        var result = new List<ExcelA1Box>(tail == 0 ? 1 : 2);
        if (fullRows > 0) result.Add(new ExcelA1Box(requested.Row, requested.Column, fullRows, requested.Columns));
        if (tail > 0) result.Add(new ExcelA1Box(requested.Row + fullRows, requested.Column, 1, tail));
        return result;
    }

    internal static object? ValueAt(object? source, int row, int column)
    {
        if (source is not Array values) return row == 1 && column == 1 ? source : null;
        try { return values.GetValue(values.GetLowerBound(0) + row - 1, values.GetLowerBound(1) + column - 1); }
        catch (IndexOutOfRangeException) { return null; }
        catch (RankException) { return null; }
    }
}
}
