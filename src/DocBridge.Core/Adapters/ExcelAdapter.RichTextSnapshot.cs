using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class ExcelAdapter
{
    private static JsonObject CaptureRichTextEntry(object workbook, JsonObject op)
    {
        using var range = BindRange(workbook, op, "cell");
        var state = ReadRichText(range.Sheet, range.Address);
        // Run-only capture cannot prove restoration of empty-cell font or nontext types.
        object? cell = null;
        try
        {
            cell = (object)((dynamic)range.Sheet).Range(range.Address);
            object? value = ((dynamic)cell).Value2;
            object? hasFormula = ((dynamic)cell).HasFormula;
            if (value is not string text || text.Length == 0 || hasFormula is not bool formula || formula)
            { state["complete"] = false; state["restoreLimitation"] = "exact run restore requires a nonempty literal text cell"; }
        }
        finally { RotHelper.ReleaseComReference(cell); }
        return new JsonObject { ["kind"] = "richText", ["sheet"] = range.SheetName, ["range"] = range.Address, ["state"] = state, ["coverageComplete"] = Json.GetBool(state, "complete") };
    }

    private static int RestoreRichTextEntry(object workbook, JsonObject entry, RestoreMismatchCollector mismatches)
    {
        var state = Json.GetObj(entry, "state");
        if (state is null || !Json.GetBool(state, "complete")) { mismatches.Add("richText snapshot coverage is incomplete; exact restore is unavailable"); return 0; }
        object? sheet = null; object? cell = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, SheetOp(Json.GetString(entry, "sheet")!)); cell = (object)((dynamic)sheet).Range(Json.GetString(entry, "range")!);
            var text = Json.GetString(state, "text") ?? ""; WriteLiteralValue(cell, JsonValue.Create(text));
            var start = 1;
            foreach (var run in Json.GetArr(state, "runs")!.OfType<JsonObject>())
            {
                var part = Json.GetString(run, "text") ?? ""; object? chars = null; object? font = null;
                try { chars = (object)((dynamic)cell).Characters(start, part.Length); font = (object)((dynamic)chars).Font; if (Json.GetObj(run, "font") is { } richFont) ApplyRichFont(font, richFont); }
                finally { RotHelper.ReleaseComReference(font); RotHelper.ReleaseComReference(chars); }
                start += part.Length;
            }
            return 1;
        }
        finally { RotHelper.ReleaseComReference(cell); RotHelper.ReleaseComReference(sheet); }
    }
}
