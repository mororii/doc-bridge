namespace DocBridge.Development.ExcelProductionWorkflowProbe;

/// <summary>
/// Typed <c>set_page_setup.page</c> for the public Excel schema.
/// XML-only oddHeader/oddFooter and null fit keys are not sent.
/// </summary>
internal static class PagePayload
{
    public static JsonObject E1Acceptance(JsonObject reconstructed)
    {
        var mm = JsonUtil.Get(reconstructed, "marginsMm") as JsonObject;
        var footer = PrintFooter.StripCenterMarker(
            JsonUtil.Str(reconstructed, "centerFooter") ?? JsonUtil.Str(reconstructed, "oddFooterExact"));
        var paper = JsonUtil.Str(reconstructed, "paperSize");
        var page = new JsonObject
        {
            ["paperSize"] = paper is "8" or "A3" or "a3" ? "A3" : (paper ?? "A3"),
            ["orientation"] = JsonUtil.Str(reconstructed, "orientation") ?? "landscape",
            ["scale"] = 55,
            ["printArea"] = "A1:BQ96",
        };
        if (!string.IsNullOrWhiteSpace(footer))
            page["centerFooter"] = footer;
        CopyMm(mm, page, "left");
        CopyMm(mm, page, "right");
        CopyMm(mm, page, "top");
        CopyMm(mm, page, "bottom");
        CopyMm(mm, page, "header");
        CopyMm(mm, page, "footer");
        return page;
    }

    public static JsonArray HygieneErrors(JsonObject page)
    {
        var errors = new JsonArray();
        foreach (var banned in new[] { "oddHeader", "oddFooter", "oddHeaderExact", "oddFooterExact" })
        {
            if (page.ContainsKey(banned))
                errors.Add($"{banned} is XML-only; use centerFooter");
        }

        if (!page.ContainsKey("scale"))
            errors.Add("E1 page.scale 55 is required");
        else if (page["scale"] is not JsonValue scale || !scale.TryGetValue<int>(out var s) || s != 55)
            errors.Add("page.scale must be the integer 55 (doubles fail the product TryReadInt contract)");

        if (page.ContainsKey("fitToWidth") || page.ContainsKey("fitToHeight"))
            errors.Add("omit fitToWidth/fitToHeight when authoring scale 55");

        if (page.ContainsKey("centerFooter"))
        {
            if (page["centerFooter"] is not JsonValue fv || !fv.TryGetValue<string>(out var footer))
                errors.Add("centerFooter must be a string");
            else if (footer.StartsWith("&C", StringComparison.Ordinal))
                errors.Add("centerFooter must not start with the XML &C routing marker");
        }

        foreach (var key in page.Select(p => p.Key).ToList())
        {
            if (page[key] is null)
                errors.Add($"{key} is null; omit the property");
        }

        return errors;
    }

    public static JsonArray ChartHygieneErrors(JsonObject create)
    {
        var errors = new JsonArray();
        if (JsonUtil.Str(create, "op") != "create_chart")
            errors.Add("chart op must be create_chart");
        if (JsonUtil.Str(create, "chartType") != "line")
            errors.Add("chartType must be line");
        var source = JsonUtil.Str(create, "sourceRange") ?? "";
        if (!source.Contains("G95:BP95", StringComparison.OrdinalIgnoreCase)
            && !source.Contains("$G$95:$BP$95", StringComparison.OrdinalIgnoreCase))
            errors.Add("sourceRange must be G95:BP95");
        if (JsonUtil.Bool(create, "hasLegend") != false)
            errors.Add("hasLegend must be false");
        var series = (JsonUtil.Get(create, "series") as JsonArray)?.OfType<JsonObject>().FirstOrDefault();
        if (series is null)
            errors.Add("series overlay required by data Line_chart_presentation_contract");
        else
        {
            if (JsonUtil.Str(series, "lineColor") is not { Length: > 0 })
                errors.Add("series.lineColor required");
            if (JsonUtil.Num(series, "lineWeight") is null)
                errors.Add("series.lineWeight required");
            if (JsonUtil.Str(series, "marker") != "none")
                errors.Add("series.marker must be none");
        }
        if (JsonUtil.Str(create, "chartFill") != "none" || JsonUtil.Str(create, "plotFill") != "none")
            errors.Add("chart/plot fill must be none");
        var plot = JsonUtil.Get(create, "plotArea") as JsonObject;
        if (JsonUtil.Bool(plot, "spanChart") != true)
            errors.Add("plotArea.spanChart must be true");
        return errors;
    }

    private static void CopyMm(JsonObject? mm, JsonObject page, string edge)
    {
        var value = JsonUtil.Num(mm, edge);
        if (value is null) return;
        var key = edge switch
        {
            "left" => "leftMarginMm",
            "right" => "rightMarginMm",
            "top" => "topMarginMm",
            "bottom" => "bottomMarginMm",
            "header" => "headerMarginMm",
            "footer" => "footerMarginMm",
            _ => null,
        };
        if (key is not null)
            page[key] = value.Value;
    }
}
