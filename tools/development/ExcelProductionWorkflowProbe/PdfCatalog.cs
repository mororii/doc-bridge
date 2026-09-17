namespace DocBridge.Development.ExcelProductionWorkflowProbe;

internal static class PdfCatalog
{
    public static JsonObject Checklist() => new()
    {
        ["visualQaOwner"] = "root",
        ["probeRole"] = "produce PDFs and list pages; do not mark visual quality pass",
        ["pages"] = new JsonArray
        {
            Page("e1-schedule.pdf", 1, 1, "A3 landscape requested scale 55 / freeze G6. Title, month headers, row-5 dates, footer, 941 bars, Chart1, no clip on BQ."),
            Page("e2-cost-estimate.pdf", 1, 1, "A4. Title/merge, four items, subtotal/VAT/grand, approval row, no clipped 합계."),
            Page("e3-daily-report.pdf", 1, 2, "One-page form plus overflow variant. Wrapped work notes, two sample photos, approval grid, row heights."),
            Page("e4-materials-ledger.pdf", 1, 1, "ListObject header, filter dropdowns, CF coloring after data change — visual, not XML-only."),
            Page("e5-progress-payment.pdf", 1, 2, "Input vs summary sheets. Named totals. Hyperlink target readable."),
            Page("e6-monthly-report.pdf", 2, 3, "Repeating title rows 1:2, page numbers, line+column charts not overlapping the table."),
        },
    };

    public static int? CountPages(string pdfPath)
    {
        if (!File.Exists(pdfPath)) return null;
        var text = File.ReadAllText(pdfPath);
        var pages = 0;
        var idx = 0;
        while (true)
        {
            var spaced = text.IndexOf("/Type /Page", idx, StringComparison.Ordinal);
            var compact = text.IndexOf("/Type/Page", idx, StringComparison.Ordinal);
            if (spaced < 0 && compact < 0) break;
            idx = spaced < 0 ? compact : compact < 0 ? spaced : Math.Min(spaced, compact);
            var tokenEnd = idx + (text.AsSpan(idx).StartsWith("/Type /Page") ? "/Type /Page".Length : "/Type/Page".Length);
            if (tokenEnd < text.Length && text[tokenEnd] == 's')
            {
                idx = tokenEnd;
                continue;
            }
            pages++;
            idx = tokenEnd;
        }
        return pages;
    }

    private static JsonObject Page(string file, int min, int max, string inspect) => new()
    {
        ["file"] = file,
        ["expectedPagesMin"] = min,
        ["expectedPagesMax"] = max,
        ["inspect"] = inspect,
    };
}
