namespace DocBridge.Development.ExcelProductionWorkflowProbe;

internal enum ProbeMode
{
    Plan,
    Oracle,
    Protocol,
    Live,
}

internal sealed class ProbeOptions
{
    public ProbeMode Mode { get; init; } = ProbeMode.Plan;
    public string OutputDir { get; init; } = "";
    public string ArtifactDir { get; init; } = "";
    public string SourceXlsx { get; init; } = "";
    public string RebuildDir { get; init; } = "";
    public string ExpectedSha256 { get; init; } = "";
    public string ScheduleSheet { get; init; } = "예정공정표";
    public string? ConfigPath { get; init; }
    public string? CliPath { get; init; }
    public string? McpPath { get; init; }
    public string Transport { get; init; } = "cli";
    public IReadOnlyList<string> Scenarios { get; init; } = ["all"];
    public bool LeaseOk { get; init; }
    public bool OracleCom { get; init; }
    public string? ComOraclePath { get; init; }
    public string? ResumeOwnedPath { get; init; }
    public string? ResumeFromBatchId { get; init; }
    public string? StopBeforeBatchId { get; init; }

    public static int PrintHelp()
    {
        Console.WriteLine("""
            DocBridge.ExcelProductionWorkflowProbe
            Public CLI/MCP authoring only. Isolated artifacts/bin. No product ProjectReference.

            --mode plan|oracle|protocol|live
                                    plan (default) never talks to Excel
                                    protocol tests persistent MCP only (no Excel writes)
            --output <dir>              required. JSON/MD written here (not user Documents)
            --config <json>             task fixture: sourceXlsx, rebuildDir, expectedSha256, scheduleSheet
            --source <xlsx>             original schedule (read-only); overrides config
            --rebuild <dir>             _excel_rebuild inputs; overrides config
            --expected-sha <hex>        expected source SHA-256; overrides config
            --schedule-sheet <name>     default from config
            --artifact-dir <dir>        live XLSX/PDF destination (default <output>/artifacts)
            --scenario <csv>            e1..e9 or all
            --cli <exe>                 already-built doc-bridge-cli
            --mcp <exe>                 already-built doc-bridge-mcp
            --transport cli|mcp         live requires mcp (persistent). CLI is refused for live.
            --lease-ok                  coordinator granted the Excel app lease
            --oracle-com                isolated COM oracle (not product authoring)
            --com-oracle <json>         prior com-oracle.json to overlay native ColumnWidth/Height
            --resume-owned <xlsx>       already-open owned workbook; bind it, do not create/SaveAs
            --resume-from <batch-id>    skip planned batches before this id
            --stop-before <batch-id>    run through the previous batch, then stop (PDF before mixedNF)

            Source path and hash are NOT baked into the probe. Pass --config or the explicit flags.
            Live also requires DOCBRIDGE_E2E=1. Never writes the user original or 통합 문서1.
            """);
        return 0;
    }

    public static bool TryParse(string[] args, out ProbeOptions options, out string error)
    {
        options = new ProbeOptions();
        error = "";
        if (args.Any(a => a is "-h" or "--help" or "/?"))
        {
            error = "help";
            return false;
        }

        var mode = ProbeMode.Plan;
        string? output = null;
        string? artifacts = null;
        string? config = null;
        string? source = null;
        string? rebuild = null;
        string? sha = null;
        string? sheet = null;
        string? cli = null;
        string? mcp = null;
        var transport = "cli";
        var scenarios = new List<string> { "all" };
        var lease = false;
        var oracleCom = false;
        string? comOracle = null;
        string? resumeOwned = null;
        string? resumeFrom = null;
        string? stopBefore = null;

        for (var i = 0; i < args.Length; i++)
        {
            string Need() => i + 1 < args.Length ? args[++i] : throw new ArgumentException($"{args[i]} requires a value");
            switch (args[i])
            {
                case "--mode":
                    mode = Need() switch
                    {
                        "plan" => ProbeMode.Plan,
                        "oracle" => ProbeMode.Oracle,
                        "protocol" => ProbeMode.Protocol,
                        "live" => ProbeMode.Live,
                        var other => throw new ArgumentException($"--mode must be plan|oracle|protocol|live, got {other}"),
                    };
                    break;
                case "--output": output = Need(); break;
                case "--artifact-dir": artifacts = Need(); break;
                case "--config": config = Need(); break;
                case "--source": source = Need(); break;
                case "--rebuild": rebuild = Need(); break;
                case "--expected-sha": sha = Need(); break;
                case "--schedule-sheet": sheet = Need(); break;
                case "--cli": cli = Need(); break;
                case "--mcp": mcp = Need(); break;
                case "--transport":
                    transport = Need();
                    if (transport is not ("cli" or "mcp"))
                        throw new ArgumentException("--transport must be cli or mcp");
                    break;
                case "--scenario":
                    scenarios = Need().Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                        .Select(s => s.ToLowerInvariant()).ToList();
                    break;
                case "--lease-ok": lease = true; break;
                case "--oracle-com": oracleCom = true; break;
                case "--com-oracle": comOracle = Need(); break;
                case "--resume-owned": resumeOwned = Need(); break;
                case "--resume-from": resumeFrom = Need(); break;
                case "--stop-before": stopBefore = Need(); break;
                default:
                    error = $"unknown argument: {args[i]}";
                    return false;
            }
        }

        if (string.IsNullOrWhiteSpace(output))
        {
            error = "--output <dir> is required";
            return false;
        }

        output = Path.GetFullPath(output);
        if (IsUserDocumentsPath(output))
        {
            error = "refusing: --output must not be a user Documents path";
            return false;
        }

        string? cfgSource = null, cfgRebuild = null, cfgSha = null, cfgSheet = null, cfgCom = null;
        if (!string.IsNullOrWhiteSpace(config))
        {
            config = Path.GetFullPath(config);
            if (!File.Exists(config))
            {
                error = $"--config not found: {config}";
                return false;
            }
            var node = JsonUtil.Load(config).AsObject();
            cfgSource = JsonUtil.Str(node, "sourceXlsx");
            cfgRebuild = JsonUtil.Str(node, "rebuildDir");
            cfgSha = JsonUtil.Str(node, "expectedSha256");
            cfgSheet = JsonUtil.Str(node, "scheduleSheet");
            cfgCom = JsonUtil.Str(node, "comOracle");
        }

        source = source ?? cfgSource;
        rebuild = rebuild ?? cfgRebuild;
        sha = sha ?? cfgSha;
        sheet = sheet ?? cfgSheet ?? "예정공정표";
        comOracle = comOracle ?? cfgCom;

        if (string.IsNullOrWhiteSpace(source) || string.IsNullOrWhiteSpace(rebuild) || string.IsNullOrWhiteSpace(sha))
        {
            error = "source, rebuild, and expected SHA-256 are required via --config or --source/--rebuild/--expected-sha";
            return false;
        }

        options = new ProbeOptions
        {
            Mode = mode,
            OutputDir = output,
            ArtifactDir = string.IsNullOrWhiteSpace(artifacts) ? Path.Combine(output, "artifacts") : Path.GetFullPath(artifacts),
            SourceXlsx = Path.GetFullPath(source),
            RebuildDir = Path.GetFullPath(rebuild),
            ExpectedSha256 = sha.Trim(),
            ScheduleSheet = sheet,
            ConfigPath = config,
            CliPath = cli is null ? null : Path.GetFullPath(cli),
            McpPath = mcp is null ? null : Path.GetFullPath(mcp),
            Transport = transport,
            Scenarios = scenarios,
            LeaseOk = lease,
            OracleCom = oracleCom,
            ComOraclePath = string.IsNullOrWhiteSpace(comOracle) ? null : Path.GetFullPath(comOracle),
            ResumeOwnedPath = string.IsNullOrWhiteSpace(resumeOwned) ? null : Path.GetFullPath(resumeOwned),
            ResumeFromBatchId = string.IsNullOrWhiteSpace(resumeFrom) ? null : resumeFrom,
            StopBeforeBatchId = string.IsNullOrWhiteSpace(stopBefore) ? null : stopBefore,
        };
        return true;
    }

    public static bool IsUserDocumentsPath(string path)
    {
        var docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        return !string.IsNullOrWhiteSpace(docs) &&
               path.StartsWith(docs.TrimEnd(Path.DirectorySeparatorChar), StringComparison.OrdinalIgnoreCase);
    }
}
