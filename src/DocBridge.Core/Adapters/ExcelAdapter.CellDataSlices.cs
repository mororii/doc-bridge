using System.Globalization;
using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class ExcelAdapter
{
    private static JsonObject ReadRichText(object sheet, string address, int maxChars = ExcelRichTextContract.DefaultInspectMaxChars)
    {
        object? cell = null;
        try
        {
            cell = (object)((dynamic)sheet).Range(address);
            var text = TryComString(cell, "Value2") ?? "";
            var result = new JsonObject { ["type"] = "richText", ["range"] = address, ["text"] = text, ["runs"] = new JsonArray(), ["complete"] = true, ["truncated"] = false, ["unreadable"] = false };
            if (text.Any(char.IsSurrogate)) { result["complete"] = false; result["unreadable"] = true; result["reason"] = "surrogate character runs are unsupported"; return result; }
            var inspectedLength = Math.Min(text.Length, Math.Clamp(maxChars, 1, ExcelRichTextContract.DefaultInspectMaxChars));
            if (inspectedLength < text.Length) { result["complete"] = false; result["truncated"] = true; }
            var runs = (JsonArray)result["runs"]!;
            var start = 1;
            while (start <= inspectedLength && runs.Count < ExcelRichTextContract.MaxRuns)
            {
                object? chars = null; object? font = null;
                try
                {
                    chars = (object)((dynamic)cell).Characters(start, 1); font = (object)((dynamic)chars).Font;
                    var signature = RichFontSignature(font);
                    if (JsonNode.Parse(signature) is JsonObject currentFont && currentFont.Any(p => p.Value is null))
                    { result["complete"] = false; result["unreadable"] = true; }
                    var end = start + 1;
                    while (end <= inspectedLength && end - start < 1024)
                    {
                        object? next = null; object? nextFont = null;
                        try { next = (object)((dynamic)cell).Characters(end, 1); nextFont = (object)((dynamic)next).Font; if (!string.Equals(signature, RichFontSignature(nextFont), StringComparison.Ordinal)) break; }
                        finally { RotHelper.ReleaseComReference(nextFont); RotHelper.ReleaseComReference(next); }
                        end++;
                    }
                    runs.Add(new JsonObject { ["text"] = text.Substring(start - 1, end - start), ["font"] = JsonNode.Parse(signature) }); start = end;
                }
                finally { RotHelper.ReleaseComReference(font); RotHelper.ReleaseComReference(chars); }
            }
            if (start <= inspectedLength || inspectedLength < text.Length) { result["complete"] = false; result["truncated"] = true; }
            return result;
        }
        catch { return new JsonObject { ["type"] = "richText", ["range"] = address, ["complete"] = false, ["truncated"] = false, ["unreadable"] = true }; }
        finally { RotHelper.ReleaseComReference(cell); }
    }

    private static string RichFontSignature(object font) => Json.Canonical(ReadRichFont(font));

    private static void PreviewTableRows(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var table = BindTable(workbook, op); EnsureSheetWritable(table.Sheet);
        preview.Affected.Add(new AffectedRef("tableRows", $"{table.SheetName}!{table.Name}"));
        preview.Diff.Add(new DiffEntry { Ref = $"{table.SheetName}!{table.Name}:rows", Before = table.State["dataRange"]?.DeepClone(), After = op["rows"]?.DeepClone() ?? JsonValue.Create("delete") });
    }

    private static void ApplyTableRows(object workbook, JsonObject op, ApplyExecution execution, List<string> mismatches, ref int checkedCells)
    {
        using var table = BindTable(workbook, op); EnsureSheetWritable(table.Sheet);
        var columns = Json.GetArr(table.State, "columns")?.Count ?? 0;
        var existing = CountTableRows(table.ListObject);
        var index = Json.GetInt(op, "index") ?? (Json.GetString(op, "op") == ExcelDataOperationsContract.AppendTableRows ? existing + 1 : 0);
        if (index < 1 || index > existing + 1) throw new InvalidOperationException("[EXCEL_TABLE_ROW_INDEX] index must address a data row or the append position");
        if (Json.GetString(op, "op") == ExcelDataOperationsContract.DeleteTableRows)
        {
            var count = Json.GetInt(op, "count")!.Value;
            if (index + count - 1 > existing) throw new InvalidOperationException("[EXCEL_TABLE_ROW_INDEX] delete exceeds table data rows");
            for (var i = 0; i < count; i++) ((dynamic)table.ListObject).ListRows.Item(index).Delete();
            checkedCells += count;
        }
        else
        {
            var rows = Json.GetArr(op, "rows")!;
            var columnNames = Json.GetArr(table.State, "columns")!.Select(x => x!.GetValue<string>()).ToArray();
            ExcelTableRowsContract.ValidateKnownColumns(rows, columnNames);
            if (rows.Any(row => row is JsonArray values && values.Count != columns || row is not JsonArray and not JsonObject)) throw new InvalidOperationException("[EXCEL_TABLE_ROW_WIDTH] array rows must exactly match the table column count");
            foreach (var node in rows)
            {
                var values = node as JsonArray;
                var sparse = node as JsonObject;
                if (values is null && sparse is null) continue;
                object? added = null; object? rowRange = null;
                try
                {
                    added = (object)((dynamic)table.ListObject).ListRows.Add(index);
                    rowRange = (object)((dynamic)added).Range;
                    for (var col = 0; col < columns; col++)
                    {
                        var supplied = values is not null || sparse!.ContainsKey(columnNames[col]);
                        object? cell = null;
                        try
                        {
                            cell = (object)((dynamic)rowRange).Cells(1, col + 1);
                            if (!supplied)
                            {
                                bool hasFormula;
                                try { hasFormula = Convert.ToBoolean(((dynamic)cell).HasFormula, CultureInfo.InvariantCulture); }
                                catch { hasFormula = false; }
                                if (!hasFormula) mismatches.Add($"{table.SheetName}!{table.Name}: omitted column '{columnNames[col]}' did not retain a calculated formula");
                                continue;
                            }
                            var value = values is not null ? values[col] : sparse![columnNames[col]];
                            WriteLiteralValue(cell, value);
                            if (!ExcelValueWriteContract.TypedEqual(ExcelValueWriteContract.Classify(value), ((dynamic)cell).Value2))
                                mismatches.Add($"{table.SheetName}!{table.Name}: written value mismatch at row {index}, column '{columnNames[col]}'");
                        }
                        finally { RotHelper.ReleaseComReference(cell); }
                    }
                }
                finally { RotHelper.ReleaseComReference(rowRange); RotHelper.ReleaseComReference(added); }
                index++; checkedCells++;
            }
        }
        var actual = CountTableRows(table.ListObject);
        var expected = Json.GetString(op, "op") == ExcelDataOperationsContract.DeleteTableRows ? existing - Json.GetInt(op, "count")!.Value : existing + Json.GetArr(op, "rows")!.Count;
        if (actual != expected) mismatches.Add($"{table.SheetName}!{table.Name}: expected {expected} data rows, got {actual}");
        execution.Affected.Add(new AffectedRef("tableRows", $"{table.SheetName}!{table.Name}"));
    }

    private static int CountTableRows(object table)
    {
        try { return Convert.ToInt32(((dynamic)table).ListRows.Count, CultureInfo.InvariantCulture); }
        catch (Exception ex) { throw new InvalidOperationException("[EXCEL_TABLE_ROWS_UNREADABLE] ListRows.Count could not be read", ex); }
    }

    private static void WriteLiteralValue(object cell, JsonNode? node)
    {
        var write = ExcelValueWriteContract.Classify(node);
        object? originalFormat = null;
        try
        {
            if (ExcelValueWriteContract.RequiresTextNumberFormat(write.Kind))
            { originalFormat = ((dynamic)cell).NumberFormat; ((dynamic)cell).NumberFormat = ExcelValueWriteContract.TextNumberFormat; }
            ((dynamic)cell).Value2 = write.ComValue;
        }
        finally
        {
            if (originalFormat is not null) ((dynamic)cell).NumberFormat = originalFormat;
        }
    }
}
