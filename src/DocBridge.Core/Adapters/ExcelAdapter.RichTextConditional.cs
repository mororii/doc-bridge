using System.Globalization;
using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class ExcelAdapter
{
    private static void PreviewSetRichText(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var range = BindRange(workbook, op, "cell"); EnsureSheetWritable(range.Sheet);
        preview.Affected.Add(new AffectedRef("richText", $"{range.SheetName}!{range.Address}"));
        preview.Diff.Add(new DiffEntry { Ref = $"{range.SheetName}!{range.Address}", Before = ReadRichText(range.Sheet, range.Address), After = op["runs"]?.DeepClone() });
    }

    private static void ApplySetRichText(object workbook, JsonObject op, ApplyExecution execution, List<string> mismatches, ref int checkedCells)
    {
        using var range = BindRange(workbook, op, "cell"); EnsureSheetWritable(range.Sheet);
        object? cell = null; object? baseComFont = null;
        try
        {
            cell = (object)((dynamic)range.Sheet).Range(range.Address);
            bool hasFormula;
            try { hasFormula = Convert.ToBoolean(((dynamic)cell).HasFormula, CultureInfo.InvariantCulture); }
            catch { throw new InvalidOperationException("[EXCEL_RICH_TEXT_UNREADABLE] HasFormula could not be read"); }
            if (hasFormula) throw new InvalidOperationException("[EXCEL_RICH_TEXT_UNSUPPORTED] formula cells cannot be character-formatted");
            var runs = Json.GetArr(op, "runs")!;
            var text = string.Concat(runs.OfType<JsonObject>().Select(run => Json.GetString(run, "text")!));
            object? oldNumberFormat = ((dynamic)cell).NumberFormat;
            try
            {
                ((dynamic)cell).NumberFormat = ExcelValueWriteContract.TextNumberFormat;
                ((dynamic)cell).Value2 = text;
            }
            finally { ((dynamic)cell).NumberFormat = oldNumberFormat; }
            baseComFont = (object)((dynamic)cell).Font;
            ApplyRichFont(baseComFont, Json.GetObj(op, "baseFont")!);
            var start = 1; // Excel documents Start as 1-based; surrogate text is rejected by the public contract.
            foreach (var run in runs.OfType<JsonObject>())
            {
                var runText = Json.GetString(run, "text")!; object? characters = null; object? font = null;
                try { characters = (object)((dynamic)cell).Characters(start, runText.Length); font = (object)((dynamic)characters).Font; ApplyRichFont(font, ExcelRichTextContract.EffectiveFont(Json.GetObj(op, "baseFont")!, Json.GetObj(run, "font"))); }
                finally { RotHelper.ReleaseComReference(font); RotHelper.ReleaseComReference(characters); }
                start += runText.Length;
            }
            VerifyRequestedRichText(cell, op, text, mismatches, $"{range.SheetName}!{range.Address}"); checkedCells++;
            execution.Affected.Add(new AffectedRef("richText", $"{range.SheetName}!{range.Address}"));
        }
        finally { RotHelper.ReleaseComReference(baseComFont); RotHelper.ReleaseComReference(cell); }
    }

    private static void VerifyRequestedRichText(object cell, JsonObject op, string expectedText, List<string> mismatches, string reference)
    {
        if (!string.Equals(TryComString(cell, "Value2"), expectedText, StringComparison.Ordinal)) { mismatches.Add(reference + ": rich text differs after write"); return; }
        var start = 1;
        foreach (var run in Json.GetArr(op, "runs")!.OfType<JsonObject>())
        {
            var text = Json.GetString(run, "text")!; object? chars = null; object? font = null;
            try
            {
                chars = (object)((dynamic)cell).Characters(start, text.Length); font = (object)((dynamic)chars).Font;
                var actual = ReadRichFont(font);
                if (!ExcelRichTextContract.FontEquals(ExcelRichTextContract.EffectiveFont(Json.GetObj(op, "baseFont")!, Json.GetObj(run, "font")), actual)) mismatches.Add($"{reference}: effective font mismatch at requested span {start}:{text.Length}");
            }
            catch { mismatches.Add($"{reference}: requested rich-text span {start}:{text.Length} is unreadable"); }
            finally { RotHelper.ReleaseComReference(font); RotHelper.ReleaseComReference(chars); }
            start += text.Length;
        }
    }

    private static JsonObject ReadRichFont(object font) => new()
    {
        ["name"] = TryComString(font, "Name"), ["size"] = Js(TryComDouble(font, "Size")), ["bold"] = Js(TryComBool(font, "Bold")),
        ["italic"] = Js(TryComBool(font, "Italic")), ["underline"] = TryComInt(font, "Underline") is int underline ? underline != -4142 : null,
        ["color"] = Js(TryComDouble(font, "Color")),
    };

    private static void ApplyRichFont(object font, JsonObject value)
    {
        ((dynamic)font).Name = Json.GetString(value, "name"); ((dynamic)font).Size = Convert.ToDouble(value["size"]!.ToString(), CultureInfo.InvariantCulture);
        ((dynamic)font).Bold = Json.GetBool(value, "bold"); ((dynamic)font).Italic = Json.GetBool(value, "italic");
        ((dynamic)font).Underline = Json.GetBool(value, "underline") ? -4119 : -4142;
        if (ExcelStyleContract.TryParseColor(value["color"], out var color)) ((dynamic)font).Color = color;
    }

    private static void PreviewUpdateConditional(object workbook, JsonObject op, ApplyPreview preview) => PreviewConditionalTarget(workbook, op, preview, "update");
    private static void PreviewDeleteConditional(object workbook, JsonObject op, ApplyPreview preview) => PreviewConditionalTarget(workbook, op, preview, "delete");
    private static void PreviewConditionalTarget(object workbook, JsonObject op, ApplyPreview preview, string action)
    { using var range = BindRange(workbook, op, "range"); var target = FindConditionalTarget(range.Sheet, range.Address, op); preview.Affected.Add(new AffectedRef("conditionalFormat", $"{range.SheetName}!{range.Address}:{Json.GetInt(op, "index")}")); preview.Diff.Add(new DiffEntry { Ref = $"{range.SheetName}!{range.Address}:formatCondition", Before = target.DeepClone(), After = JsonValue.Create(action) }); }

    private static void ApplyUpdateConditional(object workbook, JsonObject op, ApplyExecution execution, List<string> mismatches, ref int checkedCells)
    {
        using var range = BindRange(workbook, op, "range"); EnsureSheetWritable(range.Sheet); object? conditions = null; object? item = null;
        try
        {
            _ = FindConditionalTarget(range.Sheet, range.Address, op); conditions = (object)((dynamic)((dynamic)range.Sheet).Range(range.Address)).FormatConditions; item = (object)((dynamic)conditions).Item(Json.GetInt(op, "index")!.Value);
            if (Json.GetObj(op, "rule") is { } rule)
            { var type = Json.GetString(rule, "type"); if (type is not ("expression" or "cellValue")) throw new InvalidOperationException("[EXCEL_DATA_UNSUPPORTED] only expression and cellValue conditional rules can be updated"); var expectedType = type == "expression" ? ExcelDataObjectCatalog.XlExpression : ExcelDataObjectCatalog.XlCellValue; if (TryComInt(item, "Type") != expectedType) throw new InvalidOperationException("[EXCEL_CONDITIONAL_CHANGED] rule type changed; inspect and retry"); var comparison = type == "cellValue" && ExcelDataObjectCatalog.TryComparisonOperator(Json.GetString(rule, "operator"), out var mapped) ? mapped : Type.Missing; ((dynamic)item).Modify(expectedType, comparison, Json.GetString(rule, "formula1"), Json.GetString(rule, "formula2") ?? Type.Missing); }
            ApplyConditionalStyle(item, Json.GetObj(op, "style")); checkedCells++;
            var actual = ReadOneConditionalRule(item, Json.GetInt(op, "index")!.Value);
            if (!ExcelConditionalFormatContract.MatchesRequested(actual, Json.GetObj(op, "rule"), Json.GetObj(op, "style"))) mismatches.Add($"{range.SheetName}!{range.Address}: requested conditional rule/style was not read back");
            execution.Affected.Add(new AffectedRef("conditionalFormat", $"{range.SheetName}!{range.Address}:{Json.GetInt(op, "index")}"));
        }
        finally { RotHelper.ReleaseComReference(item); RotHelper.ReleaseComReference(conditions); }
    }

    private static void ApplyDeleteConditional(object workbook, JsonObject op, ApplyExecution execution, List<string> mismatches, ref int checkedCells)
    { using var range = BindRange(workbook, op, "range"); EnsureSheetWritable(range.Sheet); object? conditions = null; object? item = null; try { _ = FindConditionalTarget(range.Sheet, range.Address, op); conditions = (object)((dynamic)((dynamic)range.Sheet).Range(range.Address)).FormatConditions; item = (object)((dynamic)conditions).Item(Json.GetInt(op, "index")!.Value); ((dynamic)item).Delete(); checkedCells++; execution.Affected.Add(new AffectedRef("conditionalFormat", $"{range.SheetName}!{range.Address}:{Json.GetInt(op, "index")}")); } finally { RotHelper.ReleaseComReference(item); RotHelper.ReleaseComReference(conditions); } }
    private static JsonObject FindConditionalTarget(object sheet, string address, JsonObject op)
    { var inventory = ReadConditionalFormats(sheet, address); var index = Json.GetInt(op, "index")!.Value; var rule = Json.GetArr(inventory, "rules")?.ElementAtOrDefault(index - 1) as JsonObject; if (rule is null || Json.GetBool(rule, "unreadable")) throw new InvalidOperationException("[EXCEL_CONDITIONAL_UNREADABLE] selected rule is unavailable"); if (!ExcelConditionalFormatContract.Matches(rule, Json.GetString(op, "expectedFingerprint"))) throw new InvalidOperationException("[EXCEL_CONDITIONAL_CHANGED] selected rule fingerprint changed; inspect and retry"); return rule; }
}
