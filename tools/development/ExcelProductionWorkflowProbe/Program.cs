namespace DocBridge.Development.ExcelProductionWorkflowProbe;

internal static class Program
{
    [STAThread]
    public static int Main(string[] args)
    {
        ProbeOptions options;
        try
        {
            if (!ProbeOptions.TryParse(args, out options, out var error))
                return error == "help" ? ProbeOptions.PrintHelp() : Fail(error);
        }
        catch (Exception ex)
        {
            return Fail(ex.Message);
        }

        try
        {
            Directory.CreateDirectory(options.OutputDir);
            foreach (var name in new[] { "harness-report.json", "schedule-oracle.json", "source-hash.json", "com-oracle.json", "golden-source.pdf" })
            {
                var path = Path.Combine(options.OutputDir, name);
                if (File.Exists(path))
                    return Fail($"refusing to overwrite existing path: {path}");
            }

            var hash = SourceIntegrity.Capture(options.SourceXlsx, options.ExpectedSha256);
            if (JsonUtil.Bool(hash, "match") != true)
            {
                var liveResume = options.Mode == ProbeMode.Live
                    && !string.IsNullOrWhiteSpace(options.ResumeOwnedPath);
                var includesE1 = options.Scenarios.Any(s =>
                    s.Equals("e1", StringComparison.OrdinalIgnoreCase)
                    || s.Equals("all", StringComparison.OrdinalIgnoreCase));
                var remainingPlan = options.Mode == ProbeMode.Plan && !includesE1;
                if (!liveResume && !remainingPlan)
                    return Fail($"source SHA-256 mismatch: {hash["sha256"]} != {options.ExpectedSha256}");
                hash["continuesDespiteMismatch"] = true;
                hash["frozenExpectedSha256"] = options.ExpectedSha256;
                hash["note"] = "E1 frozen baseline stays F1D0B0DB; current source bytes are the user 09:41 save, not the 0918 pin";
            }

            if (options.Mode == ProbeMode.Protocol)
            {
                if (options.Transport != "mcp")
                    return Fail("protocol mode requires --transport mcp");
                using var client = new PublicClient(options);
                var proto = client.ProtocolSelfTest();
                JsonUtil.Write(Path.Combine(options.OutputDir, "protocol-self-test.json"), proto);
                Console.WriteLine(JsonUtil.Bool(proto, "ok") == true
                    ? $"protocol ok → {Path.Combine(options.OutputDir, "protocol-self-test.json")}"
                    : proto.ToJsonString(JsonUtil.WritePretty));
                return JsonUtil.Bool(proto, "ok") == true ? 0 : 2;
            }

            if (options.Mode == ProbeMode.Live)
            {
                var live = ScenarioRunner.RunLive(options, ScenarioRunner.Expand(options.Scenarios));
                JsonUtil.Write(Path.Combine(options.OutputDir, "harness-report.json"), live);
                return JsonUtil.Bool(live, "ok") == true ? 0 : 1;
            }

            var oracle = ScheduleOracle.Export(options);
            if (options.Mode == ProbeMode.Oracle)
            {
                JsonUtil.Write(Path.Combine(options.OutputDir, "schedule-oracle.json"), oracle);
                JsonUtil.Write(Path.Combine(options.OutputDir, "source-hash.json"), hash);
                if (options.OracleCom)
                {
                    var com = ComOracle.ReadWorkbook(options, options.SourceXlsx);
                    JsonUtil.Write(Path.Combine(options.OutputDir, "com-oracle.json"), com);
                    File.WriteAllText(Path.Combine(options.OutputDir, "oracle-native.md"),
                        $"ok={com["ok"]} sourceUnchanged={com["sourceUnchanged"]} pdf={com["goldenPdf"]}\n",
                        new UTF8Encoding(false));
                    return JsonUtil.Bool(com, "ok") == true ? 0 : 2;
                }
                return 0;
            }

            var self = AcceptanceNumbers.SelfCheck();
            var ids = ScenarioRunner.Expand(options.Scenarios);
            var artifactDir = options.ArtifactDir;
            var pictureDir = Path.Combine(options.OutputDir, "fixtures");
            Directory.CreateDirectory(pictureDir);
            WriteFixtureManifest(options, pictureDir);
            var plans = new Dictionary<string, List<PlannedBatch>>(StringComparer.OrdinalIgnoreCase);
            foreach (var id in ids)
            {
                var scenarioDir = Path.Combine(artifactDir, id);
                Directory.CreateDirectory(scenarioDir);
                plans[id] = ScenarioRunner.Plan(id, options, scenarioDir, pictureDir);
            }

            ReportWriter.WritePlan(options, oracle, hash, ids, plans, self);

            if (JsonUtil.Bool(self, "ok") != true)
                return Fail("acceptance self-check failed; see harness-report.json");
            var remainingWithoutE1 = !ids.Any(s =>
                s.Equals("e1", StringComparison.OrdinalIgnoreCase)
                || s.Equals("all", StringComparison.OrdinalIgnoreCase));
            if (!remainingWithoutE1)
            {
                if (JsonUtil.Bool(JsonUtil.Get(oracle, "crossCheck"), "mergeSetEqual") != true)
                    return Fail("XML vs rebuild merge sets differ; see schedule-oracle.json");
                var month = (int)(JsonUtil.Num(JsonUtil.Get(oracle, "crossCheck"), "monthHeaderCountXml") ?? 0);
                if (month != 31)
                    return Fail($"month-header merges were {month}, expected 31");
            }

            Console.WriteLine($"plan ok → {Path.Combine(options.OutputDir, "harness-report.json")}");
            return 0;
        }
        catch (Exception ex)
        {
            return Fail(ex.ToString());
        }
    }

    internal static void WriteFixtureManifest(ProbeOptions options, string pictureDir)
    {
        TinyPng.WriteSample(pictureDir, "e3-photo-1.png", 40, 90, 160);
        TinyPng.WriteSample(pictureDir, "e3-photo-2.png", 160, 90, 40);
        TinyPng.WriteSample(pictureDir, "e6-logo.png", 20, 20, 20);
        File.WriteAllText(Path.Combine(pictureDir, "README.txt"),
            "Sample images generated by the probe. Not site photography.\n", new UTF8Encoding(false));
        JsonUtil.Write(Path.Combine(options.OutputDir, "fixture-manifest.json"), new JsonObject
        {
            ["disclaimer"] = "E2–E9 are functional virtual data. E1 uses _excel_rebuild + original XML golden.",
            ["rebuild"] = options.RebuildDir,
            ["source"] = options.SourceXlsx,
            ["pictures"] = pictureDir,
        });
    }

    private static int Fail(string message)
    {
        Console.Error.WriteLine(message);
        return 2;
    }
}
