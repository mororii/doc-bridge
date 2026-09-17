namespace DocBridge.Development.ExcelProductionWorkflowProbe;

internal static class ScheduleOracle
{
    public static JsonObject Export(ProbeOptions options)
    {
        var hash = SourceIntegrity.Capture(options.SourceXlsx, options.ExpectedSha256);
        var xml = GoldenXml.Inspect(options.SourceXlsx, options.ScheduleSheet);
        var reconstruction = StyleReconstructor.Reconstruct(options.SourceXlsx, options.ScheduleSheet);
        var meta = JsonUtil.Load(Path.Combine(options.RebuildDir, "meta.json")).AsObject();
        var bars = JsonUtil.Load(Path.Combine(options.RebuildDir, "bars.json")).AsArray();
        var formulas = JsonUtil.Load(Path.Combine(options.RebuildDir, "formulas.json")).AsArray();
        var values = JsonUtil.Load(Path.Combine(options.RebuildDir, "values.json")).AsArray();

        var barCells = bars.Select(ParseRc).ToList();
        var styleCounts = GoldenXml.CountUsedStylesAndMediumBottoms(options.SourceXlsx, options.ScheduleSheet, barCells);

        var rebuildMerges = meta["merges"]!.AsArray().Select(n => n!.GetValue<string>()).ToList();
        var xmlMerges = xml["merges"]!.AsArray().Select(n => n!.GetValue<string>()).ToList();
        var month = xml["monthHeaderMerges"]!.AsArray().Select(n => n!.GetValue<string>()).ToList();

        var row5 = ExtractRow(values, 5);
        var expectedDates = new JsonArray();
        foreach (var n in row5.Skip(6).Take(62))
            expectedDates.Add(JsonUtil.Clone(n));

        var bp95 = FindFormula(formulas, 95, 68);
        var bp96 = FindFormula(formulas, 96, 68);
        var ab95 = FindFormula(formulas, 95, 28);

        var rebuildMonth = rebuildMerges.Where(A1.IsMonthHeaderPair).OrderBy(m => A1.ParseRange(m).C1).ToList();
        var rebuildOutsideGrid = barCells
            .Where(b => b.R < 1 || b.R > 96 || b.C < 7 || b.C > 68)
            .Select(b => A1.Cell(b.R, b.C))
            .ToList();

        return new JsonObject
        {
            ["kind"] = "schedule-oracle",
            ["authoring"] = "none — read-only XML + rebuild JSON",
            ["source"] = hash,
            ["xml"] = xml,
            ["stylesAndBars"] = styleCounts,
            ["reconstruction"] = new JsonObject
            {
                ["usedCellStyleCount"] = reconstruction["usedCellStyleCount"]?.DeepClone(),
                ["uniqueStylePayloads"] = reconstruction["uniqueStylePayloads"]?.DeepClone(),
                ["formatOpCount"] = reconstruction["formatOpCount"]?.DeepClone(),
                ["cellsWithStyle"] = reconstruction["cellsWithStyle"]?.DeepClone(),
                ["columnGroups"] = (reconstruction["columns"] as JsonArray)?.Count,
                ["rowGroups"] = (reconstruction["rows"] as JsonArray)?.Count,
                ["oddFooterExact"] = JsonUtil.Get(JsonUtil.Get(reconstruction, "page"), "oddFooterExact")?.DeepClone(),
                ["notInferredFromTitleAndBarsAlone"] = true,
            },
            ["rebuild"] = new JsonObject
            {
                ["dir"] = options.RebuildDir,
                ["mergeCount"] = rebuildMerges.Count,
                ["formulaCount"] = formulas.Count,
                ["barCount"] = bars.Count,
                ["valueRows"] = values.Count,
                ["metaMaxRow"] = JsonUtil.Num(meta, "max_row"),
                ["metaMaxCol"] = JsonUtil.Num(meta, "max_col"),
                ["monthHeaderCount"] = rebuildMonth.Count,
            },
            ["crossCheck"] = new JsonObject
            {
                ["mergeCountEqual"] = rebuildMerges.Count == xmlMerges.Count,
                ["mergeSetEqual"] = SetsEqual(rebuildMerges, xmlMerges),
                ["monthHeaderCountXml"] = month.Count,
                ["monthHeaderCountRebuild"] = rebuildMonth.Count,
                ["monthHeadersEqual"] = SetsEqual(month, rebuildMonth),
                ["expectedMonthHeaders"] = 31,
                ["expectedTotalMerges"] = 281,
                ["expectedBarsPublished"] = 941,
                ["rebuildBarCount"] = bars.Count,
                ["scheduleGridBarCountXml"] = JsonUtil.Num(styleCounts, "xmlScheduleMediumBottomCellCount"),
                ["scheduleBarSetEqualG1BP96"] = styleCounts["barCompare"] is JsonObject bc
                    && (bc["onlyInRebuild"] as JsonArray)?.Count == 0
                    && (bc["onlyInXmlScheduleGrid"] as JsonArray)?.Count == 0,
                ["rebuildBarsOutsideG1BP96"] = new JsonArray(rebuildOutsideGrid.Select(s => JsonValue.Create(s)).ToArray()),
                ["barNote"] = "Published 941 includes every bars.json entry. Exact G1:BP96 medium-bottom set matches XML. Leftover rebuild cells and left-of-G / BQ mediums are table/boundary, not schedule bars.",
                ["row5DateCellsPreservedInRebuild"] = row5.Count >= 68,
            },
            ["numericFormulas"] = new JsonObject
            {
                ["BP95"] = new JsonObject { ["formula"] = bp95, ["expectedApprox"] = 100d },
                ["BP96"] = new JsonObject { ["formula"] = bp96, ["expected"] = 22656d },
                ["AB95"] = new JsonObject { ["formula"] = ab95, ["expected"] = 17.659d },
            },
            ["row5DatesFromRebuild"] = expectedDates,
            ["e1MergePlan"] = new JsonObject
            {
                ["batch1_monthHeaders"] = new JsonArray(month.Select(m => JsonValue.Create(m)).ToArray()),
                ["batch2_remaining"] = xml["remainingMerges"]!.DeepClone(),
            },
            ["layoutGolden"] = new JsonObject
            {
                ["freezeCell"] = "G6",
                ["paper"] = "A3",
                ["orientation"] = "landscape",
                ["scale"] = 55,
                ["printArea"] = "A1:BQ96",
                ["defaultRowHeightPoints"] = xml["defaultRowHeightPoints"]?.DeepClone(),
                ["columnWidthXmlGtoBP"] = 4.09765625,
                ["columnWidthNote"] = "Compare XML width and COM ColumnWidth with documented tolerance; do not require equality.",
                ["titleFont"] = new JsonObject { ["name"] = "돋움", ["size"] = 28, ["bold"] = true },
                ["barBorder"] = new JsonObject { ["edge"] = "bottom", ["weight"] = "medium", ["lineStyle"] = "continuous", ["scope"] = "each-cell" },
            },
        };
    }

    public static JsonArray LoadValues(string rebuildDir) =>
        JsonUtil.Load(Path.Combine(rebuildDir, "values.json")).AsArray();

    public static JsonArray LoadFormulas(string rebuildDir) =>
        JsonUtil.Load(Path.Combine(rebuildDir, "formulas.json")).AsArray();

    public static JsonArray LoadBars(string rebuildDir) =>
        JsonUtil.Load(Path.Combine(rebuildDir, "bars.json")).AsArray();

    public static JsonArray LoadNumberFormats(string rebuildDir) =>
        JsonUtil.Load(Path.Combine(rebuildDir, "nf.json")).AsArray();

    private static List<JsonNode?> ExtractRow(JsonArray values, int excelRow)
    {
        var idx = excelRow - 1;
        if (idx < 0 || idx >= values.Count || values[idx] is not JsonArray row) return [];
        return row.Select(n => n).ToList();
    }

    private static string? FindFormula(JsonArray formulas, int r, int c)
    {
        foreach (var item in formulas.OfType<JsonObject>())
        {
            if ((int)(JsonUtil.Num(item, "r") ?? -1) == r && (int)(JsonUtil.Num(item, "c") ?? -1) == c)
                return JsonUtil.Str(item, "f");
        }
        return null;
    }

    private static (int R, int C) ParseRc(JsonNode? node)
    {
        var o = node as JsonObject ?? throw new InvalidOperationException("bar entry must be object");
        return ((int)(JsonUtil.Num(o, "r") ?? 0), (int)(JsonUtil.Num(o, "c") ?? 0));
    }

    private static bool SetsEqual(IEnumerable<string> a, IEnumerable<string> b) =>
        a.ToHashSet(StringComparer.OrdinalIgnoreCase)
            .SetEquals(b.ToHashSet(StringComparer.OrdinalIgnoreCase));
}
