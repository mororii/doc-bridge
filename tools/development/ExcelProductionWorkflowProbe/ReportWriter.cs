namespace DocBridge.Development.ExcelProductionWorkflowProbe;

internal static class ReportWriter
{
    public static void WritePlan(
        ProbeOptions options,
        JsonObject oracle,
        JsonObject hash,
        IReadOnlyList<string> scenarios,
        IReadOnlyDictionary<string, List<PlannedBatch>> plans,
        JsonObject selfCheck)
    {
        var planDir = Path.Combine(options.OutputDir, "planned-ops");
        Directory.CreateDirectory(planDir);
        var expected = new JsonObject();
        var planIndex = new JsonArray();
        var warnings = new JsonArray();

        foreach (var id in scenarios)
        {
            expected[id] = ScenarioRunner.Expected(id);
            var batches = plans[id];
            var file = Path.Combine(planDir, $"{id}.json");
            JsonUtil.Write(file, new JsonObject
            {
                ["scenario"] = id,
                ["batchCount"] = batches.Count,
                ["batches"] = new JsonArray(batches.Select(b => b.ToJson()).ToArray()),
            });
            planIndex.Add(new JsonObject
            {
                ["id"] = id,
                ["file"] = file,
                ["batches"] = batches.Count,
                ["ops"] = batches.Sum(b => b.Ops.Count),
            });
            if (id == "e8")
                warnings.Add("e8 unlocked B1=11 must recalc B3 to 11000; fixture b3-still-10000 is not acceptance");
            if (id == "e2")
                warnings.Add("e2 signature blocks are A12:B14/C12:D14/E12:F14/G12:H14; e2-sig-gap-relabel is G12:H12 only; do not replay historical H12:H14 or wipe B12/D12/F12 empty constants");
        }

        JsonUtil.Write(Path.Combine(options.OutputDir, "schedule-oracle.json"), oracle);
        JsonUtil.Write(Path.Combine(options.OutputDir, "source-hash.json"), hash);
        JsonUtil.Write(Path.Combine(options.OutputDir, "expected-checks.json"), expected);
        JsonUtil.Write(Path.Combine(options.OutputDir, "contract-catalog.json"), ContractCatalog.Describe());
        JsonUtil.Write(Path.Combine(options.OutputDir, "pdf-page-checklist.json"), PdfCatalog.Checklist());
        JsonUtil.Write(Path.Combine(options.OutputDir, "live-checks.json"), LiveVerifier.Catalog());
        var coverage = OpCoverage.FromPlans(plans);
        JsonUtil.Write(Path.Combine(options.OutputDir, "op-coverage.json"), coverage);
        File.WriteAllText(Path.Combine(options.OutputDir, "pdf-page-checklist.md"), PdfMarkdown(), new UTF8Encoding(false));

        var remainingWithoutE1 = !scenarios.Any(s =>
            s.Equals("e1", StringComparison.OrdinalIgnoreCase)
            || s.Equals("all", StringComparison.OrdinalIgnoreCase));
        var oracleOk = JsonUtil.Bool(selfCheck, "ok") == true
                       && (remainingWithoutE1
                           || (JsonUtil.Bool(hash, "match") == true
                               && JsonUtil.Bool(JsonUtil.Get(oracle, "crossCheck"), "mergeSetEqual") == true));

        var report = new JsonObject
        {
            ["ok"] = oracleOk,
            ["mode"] = options.Mode.ToString().ToLowerInvariant(),
            ["liveComRun"] = false,
            ["productEdited"] = false,
            ["coreCompiled"] = false,
            ["selfCheck"] = selfCheck,
            ["sourceHash"] = hash,
            ["scenarios"] = planIndex,
            ["warnings"] = warnings,
            ["opCoverage"] = new JsonObject
            {
                ["advertisedOps"] = coverage["advertisedOps"]!.DeepClone(),
                ["plannedDistinctOps"] = coverage["plannedDistinctOps"]!.DeepClone(),
                ["advertisedWithoutPlanCount"] = coverage["advertisedWithoutPlanCount"]!.DeepClone(),
                ["file"] = Path.Combine(options.OutputDir, "op-coverage.json"),
            },
            ["oracleNotes"] = new JsonArray(
                "OOXML column width is not COM ColumnWidth",
                "Full style equality is not inferred from title/bar counts",
                "PDF existence is not visual QA",
                "941 rebuild bar entries = 940 G1:BP96 medium bottoms matching XML + BQ93 outside the schedule grid",
                "usedCellStyleCount from XML cellXfs-in-use is 60"),
        };
        JsonUtil.Write(Path.Combine(options.OutputDir, "harness-report.json"), report);
        File.WriteAllText(Path.Combine(options.OutputDir, "harness-report.md"), ReportMd(report, oracle), new UTF8Encoding(false));
    }

    private static string PdfMarkdown()
    {
        var sb = new StringBuilder();
        sb.AppendLine("# PDF page checklist (root visual QA)");
        sb.AppendLine();
        sb.AppendLine("The probe only produces files and expected page counts. Root inspects every page.");
        sb.AppendLine();
        foreach (var page in PdfCatalog.Checklist()["pages"]!.AsArray().OfType<JsonObject>())
        {
            sb.AppendLine($"- `{page["file"]}` pages {page["expectedPagesMin"]}–{page["expectedPagesMax"]}: {page["inspect"]}");
        }
        return sb.ToString();
    }

    private static string ReportMd(JsonObject report, JsonObject oracle)
    {
        var sb = new StringBuilder();
        sb.AppendLine("# Excel production harness — plan report");
        sb.AppendLine();
        sb.AppendLine($"ok: **{report["ok"]}**. Live COM was not run. Product code was not edited.");
        sb.AppendLine();
        sb.AppendLine("## Source");
        sb.AppendLine();
        sb.AppendLine($"- SHA-256 match: {JsonUtil.Get(report, "sourceHash")?["match"]}");
        var cc = JsonUtil.Get(oracle, "crossCheck");
        sb.AppendLine($"- Merge set equal (XML vs rebuild): {JsonUtil.Get(cc, "mergeSetEqual")}");
        sb.AppendLine($"- Month headers: {JsonUtil.Get(cc, "monthHeaderCountXml")} (expect 31)");
        sb.AppendLine($"- Schedule-grid bars equal: {JsonUtil.Get(cc, "scheduleBarSetEqualG1BP96")}; rebuild total {JsonUtil.Get(cc, "rebuildBarCount")}; outside G1:BP96 {JsonUtil.Get(cc, "rebuildBarsOutsideG1BP96")}");
        sb.AppendLine($"- {JsonUtil.Get(cc, "barNote")}");
        sb.AppendLine();
        sb.AppendLine("## Scenarios planned");
        sb.AppendLine();
        foreach (var s in report["scenarios"]!.AsArray().OfType<JsonObject>())
            sb.AppendLine($"- `{s["id"]}`: {s["batches"]} batches / {s["ops"]} ops");
        sb.AppendLine();
        var cov = JsonUtil.Get(report, "opCoverage");
        if (cov is not null)
            sb.AppendLine($"Advertised ops planned: {JsonUtil.Get(cov, "plannedDistinctOps")} / {JsonUtil.Get(cov, "advertisedOps")} (without plan: {JsonUtil.Get(cov, "advertisedWithoutPlanCount")}). See `op-coverage.json`.");
        sb.AppendLine();
        sb.AppendLine("See `schedule-oracle.json`, `planned-ops/`, `expected-checks.json`, `CONTRACT-PROPOSAL.md`.");
        return sb.ToString();
    }
}
