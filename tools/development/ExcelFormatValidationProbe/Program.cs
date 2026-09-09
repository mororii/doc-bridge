using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Development.ExcelFormatValidationProbe;

/// <summary>
/// Guarded Excel format-validation probe. Requires DOCBRIDGE_E2E=1.
/// Owns a new Excel PID only. Never Kill/taskkill Office. Never Quit unless ownership is proven.
/// </summary>
internal static partial class Program
{
    private const int XlPatternNone = -4142;
    private const int XlColorIndexNone = -4142;
    private const int XlThemeColorAccent1 = 5;
    private const double CanonicalFillOle = 255d;
    private const double AliasFillOle = 16_711_680d; // #0000FF OLE BGR
    private const double FingerprintSeedOle = 12_632_256d; // known solid untinted RGB
    private const double FingerprintMutateOle = 65_280d;
    private const double UntintedRgbOle = 65_535d;
    private const double OversizedFont = 4097d;
    private const string ProbeSheet = "Target";
    private const string BloatSheet = "Bloat";
    private const string LockedSheet = "Locked";
    private const string BystanderSheet = "Keep";
    private const string ProtectPassword = "probe";
    private const string Marker = "DO-NOT-TOUCH";

    private static string? ProbePath;
    private static string? BystanderPath;
    private static ExcelAdapter? SessionAdapter;
    private static object? SessionApp;
    private static SnapshotService? Snapshots;
    private static bool OwnershipProven;
    private static int OwnedPid;
    private static HashSet<string> KnownWorkbookPaths = new(StringComparer.OrdinalIgnoreCase);

    public static int Main(string[] args)
    {
        if (args.Any(a => a is "-h" or "--help" or "/?"))
        {
            Console.WriteLine("""
                DocBridge.ExcelFormatValidationProbe
                Requires DOCBRIDGE_E2E=1. Creates a NEW owned Excel PID only.

                --output <dir>     required. XLSX + JSON written here (not user Documents)
                --scale <n>        optional 1 (default), 100, 1000, or 7000 extra format cells
                                   100=A20:J29  1000=A20:J119  7000=A20:J719
                --mode <name>      legacy (default, baseline DLL) | execute | both
                --case <csv>       optional filters: uniform-bold,mixed-bold-colors,restore,
                                   tint-preservation,fault-after-success,execute-replay,
                                   execute-conflict, plus existing regression names
                --repeat <n>       optional 1-3 repetitions of timing cases (default 1)
                --label <name>     optional run label (baseline|candidate)
                --bloat            seed unrelated XFD131 UsedRange bloat
                --bloat-only       bloat fixture plus UsedRange blocker case only (no speed claim)

                Isolated baseline: copy this exe directory, replace DocBridge.Core.dll, rerun.
                Never logs confirmToken. Does not set DisplayAlerts. Does not kill Office.
                """);
            return 0;
        }

        if (!string.Equals(Environment.GetEnvironmentVariable("DOCBRIDGE_E2E"), "1", StringComparison.Ordinal))
        {
            Console.Error.WriteLine("refusing: set DOCBRIDGE_E2E=1 to run this probe");
            return 2;
        }

        string? outputDir = null;
        var scale = 1;
        var label = "unspecified";
        var modeArg = "legacy";
        var caseArg = "";
        var repeat = 1;
        var bloat = false;
        var bloatOnly = false;
        for (var i = 0; i < args.Length; i++)
        {
            if (args[i] == "--output" && i + 1 < args.Length) outputDir = args[++i];
            else if (args[i] == "--scale" && i + 1 < args.Length) scale = int.Parse(args[++i], CultureInfo.InvariantCulture);
            else if (args[i] == "--label" && i + 1 < args.Length) label = args[++i];
            else if (args[i] == "--mode" && i + 1 < args.Length) modeArg = args[++i];
            else if (args[i] == "--case" && i + 1 < args.Length) caseArg = args[++i];
            else if (args[i] == "--repeat" && i + 1 < args.Length) repeat = int.Parse(args[++i], CultureInfo.InvariantCulture);
            else if (args[i] == "--bloat") bloat = true;
            else if (args[i] == "--bloat-only") { bloat = true; bloatOnly = true; }
            else
            {
                Console.Error.WriteLine($"unknown argument: {args[i]}");
                return 2;
            }
        }

        if (string.IsNullOrWhiteSpace(outputDir))
        {
            Console.Error.WriteLine("--output <dir> is required");
            return 2;
        }

        if (scale is not (1 or 100 or 1000 or 7000))
        {
            Console.Error.WriteLine("--scale must be 1, 100, 1000, or 7000");
            return 2;
        }

        if (modeArg is not ("legacy" or "execute" or "both"))
        {
            Console.Error.WriteLine("--mode must be legacy, execute, or both");
            return 2;
        }

        if (repeat is < 1 or > 3)
        {
            Console.Error.WriteLine("--repeat must be 1, 2, or 3");
            return 2;
        }

        ModeArg = modeArg;
        Repeat = repeat;
        CaseFilter = ParseCaseFilter(caseArg);

        outputDir = Path.GetFullPath(outputDir);
        Directory.CreateDirectory(outputDir);
        if (IsUserDocumentsPath(outputDir))
        {
            Console.Error.WriteLine("refusing: --output must not be a user Documents path");
            return 2;
        }

        var homeDir = Path.Combine(outputDir, "probe-home");
        var resultPath = Path.Combine(outputDir, "excel-format-probe-result.json");
        var probePath = Path.Combine(outputDir, "excel-format-probe-target.xlsx");
        var bystanderPath = Path.Combine(outputDir, "excel-format-probe-bystander.xlsx");
        foreach (var path in new[] { resultPath, probePath, bystanderPath })
        {
            if (File.Exists(path))
            {
                Console.Error.WriteLine($"refusing to overwrite existing path: {path}");
                return 2;
            }
        }

        Directory.CreateDirectory(homeDir);
        var report = NewReport(label, scale, modeArg, repeat, bloat, bloatOnly, outputDir, homeDir);
        ExcelAdapter? adapter = null;
        DocBridgeHost? host = null;
        object? createdApp = null;
        var preexisting = InventoryExcel();
        report["preexistingExcel"] = preexisting.DeepClone();

        try
        {
            var options = new DocBridgeOptions(homeDir);
            Snapshots = new SnapshotService(options);
            adapter = new ExcelAdapter(() =>
            {
                var type = Type.GetTypeFromProgID("Excel.Application")
                    ?? throw new InvalidOperationException("Excel.Application ProgID not registered");
                object? appObj = Activator.CreateInstance(type);
                if (appObj is null) throw new InvalidOperationException("Excel.Application create returned null");
                var pid = ProcessIdOf(appObj);
                if (pid <= 0 || PidSet(preexisting).Contains(pid))
                {
                    RotHelper.ReleaseComObject(appObj);
                    throw new InvalidOperationException(
                        pid <= 0
                            ? "Excel PID was not readable immediately after COM creation"
                            : $"factory reused preexisting Excel PID {pid}");
                }

                OwnershipProven = true;
                OwnedPid = pid;
                createdApp = appObj;
                SessionApp = appObj;

                object? workbooks = null;
                object? probe = null;
                object? bystander = null;
                try
                {
                    dynamic app = appObj;
                    app.Visible = false;
                    workbooks = (object)app.Workbooks;
                    ProbePath = probePath;
                    BystanderPath = bystanderPath;
                    KnownWorkbookPaths.Add(probePath);
                    KnownWorkbookPaths.Add(bystanderPath);
                    probe = CreateProbeWorkbook(workbooks, probePath, scale, bloat);
                    bystander = CreateBystanderWorkbook(workbooks, bystanderPath);
                    ((dynamic)probe).Activate();
                    return appObj;
                }
                catch
                {
                    TryCloseWorkbook(probe);
                    TryCloseWorkbook(bystander);
                    throw;
                }
                finally
                {
                    RotHelper.ReleaseComObject(bystander);
                    RotHelper.ReleaseComObject(probe);
                    RotHelper.ReleaseComObject(workbooks);
                }
            }, appFactoryOwnsInstance: true);

            host = new DocBridgeHost(options);
            host.Router.Register("excel", adapter);
            SessionAdapter = adapter;
            var ctx = host.GetActiveContext("excel");
            SessionApp = createdApp;
            report["context"] = CompactResult(ctx);
            var contextErrors = JoinErrors(ctx);
            if (!OwnershipProven || OwnedPid <= 0)
            {
                throw new InvalidOperationException(
                    string.IsNullOrWhiteSpace(contextErrors)
                        ? "owned Excel PID was not proven"
                        : $"owned Excel PID was not proven; GetActiveContext errors: {contextErrors}");
            }
            report["ownership"] = new JsonObject
            {
                ["ownedPid"] = OwnedPid,
                ["ownedPidIsNew"] = true,
                ["provenBeforePreferencesOrWorkbooks"] = true,
                ["probeWorkbook"] = probePath,
                ["bystanderWorkbook"] = bystanderPath,
                ["appFactoryOwnsInstance"] = true,
            };

            var cases = new JsonArray();
            if (bloatOnly)
            {
                cases.Add(RunUsedRangeBloatBlocker(host, adapter, createdApp!));
            }
            else
            {
                AddRegressionCases(cases, host, adapter, createdApp!, bystanderPath, bloat);
                foreach (var mode in SelectedModes())
                {
                    CurrentMode = mode;
                    for (var rep = 1; rep <= Repeat; rep++)
                        AddSpeedCases(cases, host, adapter, createdApp!, scale, rep);
                }
            }

            report["cases"] = cases;
            SaveOwnedCopies(adapter, createdApp!, probePath, bystanderPath);
            report["artifacts"] = new JsonObject
            {
                ["resultJson"] = resultPath,
                ["targetXlsx"] = probePath,
                ["bystanderXlsx"] = bystanderPath,
            };
            report["ok"] = cases.All(n => n is JsonObject o && (Json.GetBool(o, "ok") || Json.GetBool(o, "skipped")));
            report["candidatePass"] = cases.All(n => n is JsonObject o && Json.GetBool(o, "candidatePass"));
            report["executeUnsupportedCases"] = cases.Count(n => n is JsonObject o && Json.GetBool(o, "executeUnsupported"));
            return Json.GetBool(report, "ok") ? 0 : 4;
        }
        catch (Exception ex)
        {
            report["ok"] = false;
            report["harnessError"] = ex.Message;
            report["harnessType"] = ex.GetType().FullName;
            Console.Error.WriteLine(ex);
            return 3;
        }
        finally
        {
            try
            {
                report["cleanup"] = CleanupOwned(adapter, createdApp);
                report["excelAfterCleanup"] = InventoryExcel();
                report["preexistingPidsStillPresent"] = PreexistingStillPresent(preexisting, report["excelAfterCleanup"] as JsonArray);
            }
            catch (Exception cleanupEx)
            {
                report["cleanupError"] = cleanupEx.Message;
            }
            host?.Dispose();
            WriteResult(resultPath, report);
            Console.WriteLine(resultPath);
        }
    }

    private static JsonObject RunCanonicalAndAliases(DocBridgeHost host)
    {
        var sw = Stopwatch.StartNew();
        ActivateProbeWorkbook();
        var canonical = FormatFlow(host, ProbeSheet, "A1", new JsonObject { ["fillColor"] = CanonicalFillOle, ["bold"] = true }, restore: false);
        var readA1 = host.Read("excel", ReadArgs(ProbeSheet, "A1"));
        var alias = FormatFlow(host, ProbeSheet, "A2", new JsonObject { ["fontBold"] = true, ["fill"] = "#0000FF" }, restore: false);
        var readA2 = host.Read("excel", ReadArgs(ProbeSheet, "A2"));
        var a1 = Json.GetObj(readA1, "styles");
        var a2 = Json.GetObj(readA2, "styles");
        var a1Bold = Json.GetBool(a1, "fontBold");
        var a2Bold = Json.GetBool(a2, "fontBold");
        var a1Fill = StyleNumber(a1, "interiorColor");
        var a2Fill = StyleNumber(a2, "interiorColor");
        var match = Json.GetBool(canonical, "ok") && Json.GetBool(alias, "ok") &&
                    a1Bold && a2Bold &&
                    NumbersEqual(a1Fill, CanonicalFillOle) &&
                    NumbersEqual(a2Fill, AliasFillOle);
        sw.Stop();
        return Case("canonical-fillColor-and-aliases", match, sw.ElapsedMilliseconds,
            new JsonArray(canonical, Phase("readback-A1", readA1), Phase("readback-A2", readA2)),
            new JsonArray(
                $"A1 expected fontBold=true interiorColor={CanonicalFillOle}; actual fontBold={a1Bold} interiorColor={a1Fill}",
                $"A2 expected fontBold=true interiorColor={AliasFillOle}; actual fontBold={a2Bold} interiorColor={a2Fill}"));
    }

    private static JsonObject RunJsonReorder(DocBridgeHost host)
    {
        var sw = Stopwatch.StartNew();
        ActivateProbeWorkbook();
        var first = ReorderOp(orderA: true);
        var second = ReorderOp(orderA: false);
        var hash1 = ConfirmTokenService.HashOps(new[] { first });
        var hash2 = ConfirmTokenService.HashOps(new[] { second });
        var (dry1, dryMs1) = Timed(() => host.ApplyOps("excel", MakeBatch(new JsonArray(first.DeepClone()), dryRun: true)));
        var token = Json.GetString(dry1, "confirmToken");
        var (applyReordered, applyMs) = Timed(() => host.ApplyOps("excel", MakeBatch(new JsonArray(second.DeepClone()), dryRun: false, token)));
        var (reuseToken, reuseMs) = Timed(() => host.ApplyOps("excel", MakeBatch(new JsonArray(first.DeepClone()), dryRun: false, token)));
        var reuseDenied = !Json.GetBool(reuseToken, "ok") &&
                          JoinErrors(reuseToken).Contains("already used", StringComparison.OrdinalIgnoreCase);
        var ok = Json.GetBool(dry1, "ok") && Json.GetBool(applyReordered, "ok") &&
                 string.Equals(hash1, hash2, StringComparison.Ordinal) && reuseDenied;
        sw.Stop();
        return Case("nested-json-reorder", ok, sw.ElapsedMilliseconds,
            new JsonArray(Phase("dryrun-order-a", dry1, dryMs1), Phase("apply-order-b-with-token-a", applyReordered, applyMs), Phase("token-singleuse-denied", reuseToken, reuseMs)),
            new JsonArray($"hashOpsA={hash1}", $"hashOpsB={hash2}", $"hashesEqual={string.Equals(hash1, hash2, StringComparison.Ordinal)}", $"secondApplyDenied={reuseDenied}"));
    }

    private static JsonObject RunProtectedSheet(DocBridgeHost host, ExcelAdapter adapter, object app)
    {
        var sw = Stopwatch.StartNew();
        ActivateProbeWorkbook();
        SetSheetProtected(adapter, app, LockedSheet, protect: true);
        try
        {
            var style = new JsonObject { ["fillColor"] = CanonicalFillOle, ["bold"] = true };
            var classified = ClassifyWrite(host, LockedSheet, "A1", style);
            var comTested = Json.GetBool(classified, "comWriteTested");
            var rollbackTested = Json.GetBool(classified, "rollbackTested");
            var rollback = Json.GetObj(classified, "rollback");
            var after = ReadComStyle(adapter, app, LockedSheet, "A1");
            var kind = Json.GetString(classified, "failureKind");
            var ok = Json.GetBool(classified, "classified") && rollback is not null &&
                     (kind == "dry-run-preflight"
                         ? !comTested && !rollbackTested
                         : comTested);
            sw.Stop();
            return Case("protected-sheet-com-failure", ok, sw.ElapsedMilliseconds, new JsonArray(classified),
                new JsonArray(
                    $"failureKind={kind}",
                    $"comWriteTested={comTested}",
                    $"rollbackTested={rollbackTested}",
                    $"rollbackAttempted={Json.GetBool(rollback, "attempted")} verified={Json.GetBool(rollback, "verified")} error={Json.GetString(rollback, "error")}",
                    $"lockedA1After fontBold={Json.GetBool(after, "fontBold")} fillColor={StyleNumber(after, "fillColor")}",
                    "dry-run preflight is not a COM write/rollback result; rollback fields are recorded exactly"));
        }
        finally { SetSheetProtected(adapter, app, LockedSheet, protect: false); }
    }

    private static JsonObject RunOversizedFontComFailure(DocBridgeHost host)
    {
        var sw = Stopwatch.StartNew();
        ActivateProbeWorkbook();
        var beforeE1 = ReadComStyle(SessionAdapter!, SessionApp!, ProbeSheet, "E1");
        var beforeD1 = ReadComStyle(SessionAdapter!, SessionApp!, ProbeSheet, "D1");
        var beforeE2 = ReadComStyle(SessionAdapter!, SessionApp!, ProbeSheet, "E2");
        var ops = new JsonArray(
            FormatOp(ProbeSheet, "E1", new JsonObject { ["bold"] = true }),
            FormatOp(ProbeSheet, "D1", new JsonObject { ["fontSize"] = OversizedFont }),
            FormatOp(ProbeSheet, "E2", new JsonObject { ["italic"] = true }));
        var batch = ApplyBatch(host, ops);
        var apply = batch.Apply;
        var applyMs = batch.ApplyMs;
        var dry = batch.Dry;
        var dryMs = batch.DryMs;
        var kind = ClassifyFailure(dry ?? new JsonObject { ["ok"] = true }, apply);
        var rollback = ExactRollback(apply);
        var results = Json.GetArr(apply, "operationResults");
        var firstOk = OpResultOk(results, 0);
        var secondOk = OpResultOk(results, 1);
        var thirdOk = OpResultOk(results, 2);
        var e1After = ReadComStyle(SessionAdapter!, SessionApp!, ProbeSheet, "E1");
        var d1After = ReadComStyle(SessionAdapter!, SessionApp!, ProbeSheet, "D1");
        var e2After = ReadComStyle(SessionAdapter!, SessionApp!, ProbeSheet, "E2");
        var nativeRestored = StylesEqual(beforeE1, e1After) &&
                             StylesEqual(beforeD1, d1After) &&
                             StylesEqual(beforeE2, e2After);
        var firstUndone = !Json.GetBool(e1After, "fontBold");
        var laterSkipped = !Json.GetBool(e2After, "fontItalic");
        var comTested = kind == "com-apply-exception";
        // Native COM state is the pass criterion. rollback.verified alone is not accepted.
        var ok = comTested && !Json.GetBool(apply, "ok") && firstOk && !secondOk && !thirdOk &&
                 nativeRestored && firstUndone && laterSkipped;
        sw.Stop();
        var phases = new JsonArray();
        if (dry is not null) phases.Add(Phase("dryrun", dry, dryMs));
        phases.Add(Phase("apply", apply, applyMs));
        phases.Add(rollback);
        phases.Add(new JsonObject { ["name"] = "native-before", ["E1"] = beforeE1, ["D1"] = beforeD1, ["E2"] = beforeE2 });
        phases.Add(new JsonObject { ["name"] = "native-after", ["E1"] = e1After, ["D1"] = d1After, ["E2"] = e2After });
        return Case("oversized-fontsize-com-failure", ok, sw.ElapsedMilliseconds, phases,
            new JsonArray(
                $"failureKind={kind}",
                $"comWriteTested={comTested}",
                $"op0ok={firstOk} op1ok={secondOk} op2ok={thirdOk} (later op must not be success)",
                $"nativeRestored={nativeRestored} (required; rollback.verified={Json.GetBool(rollback, "verified")} is not sufficient)",
                $"successfulOpUndone={firstUndone} laterOpSkippedNative={laterSkipped}",
                $"rollbackAttempted={Json.GetBool(rollback, "attempted")} error={Json.GetString(rollback, "error")}"));
    }

    private static JsonObject RunStyleChangedAfterDryRun(DocBridgeHost host, ExcelAdapter adapter, object app)
    {
        var sw = Stopwatch.StartNew();
        ActivateProbeWorkbook();
        // Seed a known solid untinted RGB before dry-run so restore is Color-only, not no-fill rewrite.
        SetCellFill(adapter, app, ProbeSheet, "B1", FingerprintSeedOle);
        var before = ReadComStyle(adapter, app, ProbeSheet, "B1");
        var style = new JsonObject { ["fillColor"] = CanonicalFillOle, ["bold"] = true };
        var ops = new JsonArray(FormatOp(ProbeSheet, "B1", style));
        var (dry, dryMs) = Timed(() => host.ApplyOps("excel", MakeBatch(ops, dryRun: true)));
        var token = Json.GetString(dry, "confirmToken");
        SetCellFill(adapter, app, ProbeSheet, "B1", FingerprintMutateOle);
        var (denied, denyMs) = Timed(() => host.ApplyOps("excel", MakeBatch(CloneOps(ops), dryRun: false, token)));
        SetCellFill(adapter, app, ProbeSheet, "B1", FingerprintSeedOle);
        var afterRestore = ReadComStyle(adapter, app, ProbeSheet, "B1");
        var nativeEqual = StylesEqual(before, afterRestore);
        var applyMeasured = nativeEqual
            ? Timed(() => host.ApplyOps("excel", MakeBatch(CloneOps(ops), dryRun: false, token)))
            : new TimedCall(new JsonObject { ["ok"] = false, ["skipped"] = true, ["errors"] = new JsonArray("same-token apply skipped; native B1 fill was not restored") }, 0);
        var applySame = applyMeasured.Result;
        var applyMs = applyMeasured.Ms;
        var denyErrors = JoinErrors(denied);
        var retained = !Json.GetBool(denied, "ok") &&
                       denyErrors.Contains("changed after dry-run", StringComparison.OrdinalIgnoreCase) &&
                       !denyErrors.Contains("already used", StringComparison.OrdinalIgnoreCase) &&
                       Json.GetBool(applySame, "ok");
        sw.Stop();
        return Case("style-changed-after-dryrun", Json.GetBool(dry, "ok") && retained && nativeEqual,
            sw.ElapsedMilliseconds,
            new JsonArray(Phase("dryrun", dry, dryMs), Phase("apply-after-com-style-change", denied, denyMs), Phase("same-token-after-restore", applySame, applyMs)),
            new JsonArray(
                $"seedFill={FingerprintSeedOle} mutateFill={FingerprintMutateOle}",
                $"deniedErrors={denyErrors}",
                $"sameTokenAppliedAfterRestore={Json.GetBool(applySame, "ok")}",
                $"targetRestoredBeforeReuse={nativeEqual}"));
    }

    private static JsonObject RunMixedThemeNoFillRollback(DocBridgeHost host, ExcelAdapter adapter, object app)
    {
        var sw = Stopwatch.StartNew();
        ActivateProbeWorkbook();
        SeedMixedStyles(adapter, app);
        var before = new JsonObject
        {
            ["C1"] = ReadComStyle(adapter, app, ProbeSheet, "C1"),
            ["C2"] = ReadComStyle(adapter, app, ProbeSheet, "C2"),
            ["C3"] = ReadComStyle(adapter, app, ProbeSheet, "C3"),
        };
        var flow = FormatFlow(host, ProbeSheet, "C1:C3", new JsonObject { ["fillColor"] = "#FFE699", ["bold"] = true }, restore: true);
        var after = new JsonObject
        {
            ["C1"] = ReadComStyle(adapter, app, ProbeSheet, "C1"),
            ["C2"] = ReadComStyle(adapter, app, ProbeSheet, "C2"),
            ["C3"] = ReadComStyle(adapter, app, ProbeSheet, "C3"),
        };
        var restored = StylesEqual(Json.GetObj(before, "C1"), Json.GetObj(after, "C1")) &&
                       StylesEqual(Json.GetObj(before, "C2"), Json.GetObj(after, "C2")) &&
                       StylesEqual(Json.GetObj(before, "C3"), Json.GetObj(after, "C3"));
        sw.Stop();
        return Case("mixed-theme-nofill-rollback", Json.GetBool(flow, "ok") && restored, sw.ElapsedMilliseconds,
            new JsonArray(flow, new JsonObject { ["name"] = "com-styles-before", ["cells"] = before.DeepClone() }, new JsonObject { ["name"] = "com-styles-after-restore", ["cells"] = after.DeepClone() }),
            new JsonArray(
                "C1 theme font, C2 no-fill, C3 untinted non-theme RGB (no TintAndShade)",
                "compare Color, ColorIndex, ThemeColor presence, TintAndShade, Pattern/PatternColor",
                $"nativeStatesRestored={restored}"));
    }

    private static JsonObject RunUnsupportedRgbTintColorChangeDenied(DocBridgeHost host, ExcelAdapter adapter, object app)
    {
        var sw = Stopwatch.StartNew();
        ActivateProbeWorkbook();
        SeedUnsupportedRgbTint(adapter, app);
        var before = ReadComStyle(adapter, app, ProbeSheet, "C4");
        var batch = ApplyBatch(host, new JsonArray(FormatOp(ProbeSheet, "C4", new JsonObject { ["fillColor"] = CanonicalFillOle })));
        var apply = batch.Apply;
        var applyMs = batch.ApplyMs;
        var dry = batch.Dry;
        var dryMs = batch.DryMs;
        var after = ReadComStyle(adapter, app, ProbeSheet, "C4");
        var errors = JoinErrors(dry ?? apply);
        var denied = !Json.GetBool(dry ?? apply, "ok") &&
                     errors.Contains("EXCEL_FORMAT_UNSUPPORTED_RGB_TINT", StringComparison.Ordinal);
        var unchanged = StylesEqual(before, after);
        sw.Stop();
        var phases = new JsonArray();
        if (dry is not null) phases.Add(Phase("dryrun", dry, dryMs));
        phases.Add(Phase("apply-or-deny", apply, applyMs));
        phases.Add(new JsonObject { ["name"] = "com-style-before", ["cell"] = before.DeepClone() });
        phases.Add(new JsonObject { ["name"] = "com-style-after", ["cell"] = after.DeepClone() });
        return Case("unsupported-rgb-tint-color-change-denied", denied && unchanged, sw.ElapsedMilliseconds, phases,
            new JsonArray(
                "C4 is non-theme RGB + nonzero tint; color-changing write must fail before mutation",
                $"denied={denied} nativeUnchanged={unchanged}",
                $"errors={errors}"));
    }

    private static JsonObject RunOtherWorkbookPreservation(DocBridgeHost host, ExcelAdapter adapter, object app, string bystanderPath)
    {
        var sw = Stopwatch.StartNew();
        ActivateWorkbook(adapter, app, bystanderPath);
        var beforeCtx = host.GetActiveContext("excel");
        var beforeActive = Json.GetString(beforeCtx, "documentRef");
        var beforeMarker = ReadBystanderMarker(adapter, app);
        var ops = new JsonArray(FormatOp(ProbeSheet, "A4", new JsonObject { ["bold"] = true }));
        var (dry, dryMs) = Timed(() => host.ApplyOps("excel", MakeBatch(ops, dryRun: true)));
        var token = Json.GetString(dry, "confirmToken");
        var applyMeasured = Json.GetBool(dry, "ok")
            ? Timed(() => host.ApplyOps("excel", MakeBatch(CloneOps(ops), dryRun: false, token)))
            : new TimedCall(new JsonObject { ["ok"] = false, ["skipped"] = true, ["errors"] = Json.GetArr(dry, "errors")?.DeepClone() ?? new JsonArray() }, 0);
        var apply = applyMeasured.Result;
        var applyMs = applyMeasured.Ms;
        var afterCtx = host.GetActiveContext("excel");
        var afterActive = Json.GetString(afterCtx, "documentRef");
        var afterMarker = ReadBystanderMarker(adapter, app);
        var ok = Json.GetBool(dry, "ok") && Json.GetBool(apply, "ok") &&
                 string.Equals(beforeActive, bystanderPath, StringComparison.OrdinalIgnoreCase) &&
                 string.Equals(afterActive, bystanderPath, StringComparison.OrdinalIgnoreCase) &&
                 string.Equals(beforeMarker, Marker, StringComparison.Ordinal) &&
                 string.Equals(afterMarker, Marker, StringComparison.Ordinal);
        sw.Stop();
        return Case("other-active-workbook-preservation", ok, sw.ElapsedMilliseconds,
            new JsonArray(Phase("context-before", beforeCtx), Phase("dryrun-explicit-target", dry, dryMs), Phase("apply-explicit-target", apply, applyMs), Phase("context-after", afterCtx)),
            new JsonArray(
                $"activeBefore={beforeActive}",
                $"activeAfter={afterActive}",
                $"markerBefore={beforeMarker} markerAfter={afterMarker}",
                "bystander stayed active; probe ops used target.workbook only"));
    }

    private static JsonObject RunUsedRangeBloatBlocker(DocBridgeHost host, ExcelAdapter adapter, object app)
    {
        var sw = Stopwatch.StartNew();
        var used = ReadOwnedBloatUsedRange(adapter, app);
        var cells = used["cells"] is JsonValue cellValue && TryNumber(cellValue, out var n) ? n : 0;
        var fixtureOk = cells > 1_000_000;
        var flow = FormatFlow(host, ProbeSheet, "A1", new JsonObject { ["bold"] = true }, restore: true);
        var dry = Json.GetObj(flow, "dryrun");
        var apply = Json.GetObj(flow, "apply");
        var restore = Json.GetObj(flow, "restore");
        var dryErrors = JoinErrors(dry ?? new JsonObject());
        var snapshotLimit = dry is not null && !Json.GetBool(dry, "ok") && IsSnapshotLimitReject(dryErrors);
        var targetCompleted = Json.GetBool(dry, "ok") && Json.GetBool(apply, "ok") &&
                              restore is not null && Json.GetBool(restore, "ok");
        var outcome = snapshotLimit ? "snapshot-limit-reject"
            : targetCompleted ? "target-a1-completed"
            : "unclassified";
        // Honest observed outcomes only. Candidate scoped A1 success is not a regression
        // just because the old UsedRange blocker is gone. Do not key off --label.
        var ok = fixtureOk && (snapshotLimit || targetCompleted);
        sw.Stop();
        return Case("usedrange-bloat-blocker", ok, sw.ElapsedMilliseconds,
            new JsonArray(used, flow),
            new JsonArray(
                $"ownedBloatUsedRange={used["address"]} rows={used["rows"]} cols={used["columns"]} cells={used["cells"]}",
                $"fixtureCellsOver1M={fixtureOk}",
                $"outcome={outcome}",
                $"snapshotLimitReject={snapshotLimit} targetA1Completed={targetCompleted}",
                $"dryErrors={dryErrors}"));
    }

    private static bool IsSnapshotLimitReject(string errors) =>
        errors.Contains("snapshot-limit", StringComparison.OrdinalIgnoreCase) ||
        errors.Contains("snapshot exceeds", StringComparison.OrdinalIgnoreCase) ||
        errors.Contains("format snapshot exceeds", StringComparison.OrdinalIgnoreCase);

    private static JsonObject ReadOwnedBloatUsedRange(ExcelAdapter adapter, object app)
    {
        return adapter.RunOnAdapterThread(() =>
        {
            var sheet = Sheet(app, BloatSheet);
            object? used = null;
            object? rows = null;
            object? columns = null;
            try
            {
                used = (object)((dynamic)sheet).UsedRange;
                rows = (object)((dynamic)used).Rows;
                columns = (object)((dynamic)used).Columns;
                var rowCount = Convert.ToInt32(((dynamic)rows).Count, CultureInfo.InvariantCulture);
                var colCount = Convert.ToInt32(((dynamic)columns).Count, CultureInfo.InvariantCulture);
                var address = Convert.ToString((object?)((dynamic)used).Address(false, false), CultureInfo.InvariantCulture);
                return new JsonObject
                {
                    ["name"] = "owned-bloat-usedrange",
                    ["ok"] = rowCount > 0 && colCount > 0,
                    ["sheet"] = BloatSheet,
                    ["address"] = address,
                    ["rows"] = rowCount,
                    ["columns"] = colCount,
                    ["cells"] = (long)rowCount * colCount,
                };
            }
            finally
            {
                RotHelper.ReleaseComObject(columns);
                RotHelper.ReleaseComObject(rows);
                RotHelper.ReleaseComObject(used);
                RotHelper.ReleaseComObject(sheet);
            }
        });
    }

    private static JsonObject RunScaleFormat(DocBridgeHost host, int scale)
    {
        var sw = Stopwatch.StartNew();
        ActivateProbeWorkbook();
        var range = ScaleRange(scale);
        var flow = FormatFlow(host, ProbeSheet, range, new JsonObject { ["bold"] = true }, restore: true);
        sw.Stop();
        return Case($"scale-{scale}-cells", Json.GetBool(flow, "ok"), sw.ElapsedMilliseconds, new JsonArray(flow),
            new JsonArray($"optional {scale}-cell format; not a bloat speed claim", $"range={range}"));
    }

    private static JsonObject FormatFlow(DocBridgeHost host, string sheet, string range, JsonObject style, bool restore)
    {
        var ops = new JsonArray(FormatOp(sheet, range, style));
        var batch = ApplyBatch(host, ops);
        var apply = batch.Apply;
        var applyMs = batch.ApplyMs;
        var dry = batch.Dry;
        var dryMs = batch.DryMs;
        var executeUnsupported = CurrentMode == ProbeExecutionMode.Execute && IsExecuteUnsupported(apply, dry);
        JsonObject? restoreResult = null;
        var restoreMs = 0L;
        var snapshotId = Json.GetString(apply, "snapshotId") ?? Json.GetString(dry, "snapshotId");
        var writeOk = executeUnsupported
            ? false
            : (dry is null || Json.GetBool(dry, "ok")) && Json.GetBool(apply, "ok");
        if (restore && writeOk)
        {
            var restored = RestoreRaw(host, snapshotId);
            restoreResult = restored.Result;
            restoreMs = restored.Ms;
        }
        var ok = executeUnsupported
            ? false
            : writeOk && (!restore || (restoreResult is not null && Json.GetBool(restoreResult, "ok")));
        var node = new JsonObject
        {
            ["ok"] = ok,
            ["mode"] = ModeName(CurrentMode),
            ["executeUnsupported"] = executeUnsupported,
            ["snapshotId"] = snapshotId,
            ["candidatePass"] = ok && !executeUnsupported,
        };
        if (dry is not null) node["dryrun"] = Phase("dryrun", dry, dryMs);
        node["apply"] = Phase(CurrentMode == ProbeExecutionMode.Execute ? "execute" : "apply", apply, applyMs);
        if (restoreResult is not null)
            node["restore"] = Phase("restore", restoreResult, restoreMs);
        return node;
    }

    private static JsonObject ClassifyWrite(DocBridgeHost host, string sheet, string range, JsonObject style)
    {
        var flow = FormatFlow(host, sheet, range, style, restore: false);
        var dry = Json.GetObj(flow, "dryrun");
        var apply = Json.GetObj(flow, "apply");
        var kind = ClassifyFailure(Hostish(dry), Hostish(apply));
        var rollback = ExactRollback(Hostish(apply));
        var comTested = kind == "com-apply-exception";
        var rollbackTested = comTested && Json.GetBool(rollback, "present");
        return new JsonObject
        {
            ["classified"] = true,
            ["failureKind"] = kind,
            ["comWriteTested"] = comTested,
            ["rollbackTested"] = rollbackTested,
            ["dryrun"] = dry?.DeepClone(),
            ["apply"] = apply?.DeepClone(),
            ["rollback"] = rollback,
        };
    }

    private static JsonObject Hostish(JsonObject? phase)
    {
        if (phase is null) return new JsonObject { ["ok"] = false };
        var clone = phase.DeepClone().AsObject();
        if (clone["errors"] is null) clone["errors"] = new JsonArray();
        return clone;
    }

    private static string ClassifyFailure(JsonObject dry, JsonObject apply)
    {
        if (!Json.GetBool(dry, "ok")) return "dry-run-preflight";
        if (Json.GetBool(apply, "skipped")) return "dry-run-preflight";
        if (Json.GetBool(apply, "ok")) return "none";
        var errors = JoinErrors(apply);
        if (errors.Contains("already used", StringComparison.OrdinalIgnoreCase) ||
            errors.Contains("does not match", StringComparison.OrdinalIgnoreCase) ||
            errors.Contains("changed after dry-run", StringComparison.OrdinalIgnoreCase) ||
            errors.Contains("confirmToken", StringComparison.OrdinalIgnoreCase))
            return "token-or-fingerprint";
        if (errors.Contains("apply threw", StringComparison.OrdinalIgnoreCase) ||
            errors.Contains("HRESULT", StringComparison.OrdinalIgnoreCase) ||
            errors.Contains("0x", StringComparison.OrdinalIgnoreCase) ||
            errors.Contains("COM", StringComparison.OrdinalIgnoreCase) ||
            HasComOperationError(apply))
            return "com-apply-exception";
        return "apply-denied";
    }

    private static bool HasComOperationError(JsonObject apply)
    {
        var results = Json.GetArr(apply, "operationResults");
        if (results is null) return false;
        foreach (var node in results)
        {
            if (node is not JsonObject op) continue;
            var text = string.Join(" ", Json.GetArr(op, "errors")?.Select(n => n?.ToString()) ?? Array.Empty<string?>());
            text += Json.GetString(op, "error") ?? "";
            text += Json.GetString(op, "exceptionType") ?? "";
            if (text.Contains("COM", StringComparison.OrdinalIgnoreCase) ||
                text.Contains("HRESULT", StringComparison.OrdinalIgnoreCase) ||
                (Json.GetInt(op, "hresult") is not null))
                return true;
        }
        return false;
    }

    private static JsonObject ExactRollback(JsonObject apply)
    {
        var rollback = Json.GetObj(apply, "rollback");
        return new JsonObject
        {
            ["name"] = "rollback",
            ["present"] = rollback is not null,
            ["attempted"] = Json.GetBool(rollback, "attempted"),
            ["verified"] = Json.GetBool(rollback, "verified"),
            ["error"] = Json.GetString(rollback, "error") ?? Json.GetString(Json.GetObj(rollback, "result"), "error"),
            ["elapsedMs"] = rollback?["elapsedMs"]?.DeepClone(),
            ["raw"] = rollback?.DeepClone(),
        };
    }

    private static TimedCall RestoreRaw(DocBridgeHost host, string? snapshotId)
    {
        if (string.IsNullOrWhiteSpace(snapshotId))
            return new TimedCall(new JsonObject { ["ok"] = false, ["errors"] = new JsonArray("snapshotId missing") }, 0);
        ActivateProbeWorkbook();
        var restoreDry = Timed(() => host.CoreRestoreSnapshot(RoundtripJson(new JsonObject { ["snapshotId"] = snapshotId })));
        var restored = Timed(() => host.CoreRestoreSnapshot(RoundtripJson(new JsonObject
        {
            ["snapshotId"] = snapshotId,
            ["confirmToken"] = Json.GetString(restoreDry.Result, "confirmToken"),
        })));
        restored.Result["restoreDryMs"] = restoreDry.Ms;
        return restored;
    }

    private static JsonObject FormatOp(string sheet, string range, JsonObject style) => new()
    {
        ["op"] = "format_range",
        ["target"] = new JsonObject { ["sheet"] = sheet, ["workbook"] = ProbePath },
        ["range"] = range,
        ["style"] = style.DeepClone(),
    };

    private static JsonObject ReorderOp(bool orderA)
    {
        var workbook = ProbePath;
        if (orderA)
        {
            return new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = new JsonObject { ["sheet"] = ProbeSheet, ["workbook"] = workbook },
                ["range"] = "A3",
                ["style"] = new JsonObject { ["bold"] = true, ["fillColor"] = 255 },
            };
        }

        return new JsonObject
        {
            ["style"] = new JsonObject { ["fillColor"] = 255, ["bold"] = true },
            ["range"] = "A3",
            ["target"] = new JsonObject { ["workbook"] = workbook, ["sheet"] = ProbeSheet },
            ["op"] = "format_range",
        };
    }

    private static JsonObject MakeBatch(JsonArray ops, bool dryRun, string? token = null)
    {
        var batch = new JsonObject { ["ops"] = ops, ["dryRun"] = dryRun };
        if (!string.IsNullOrWhiteSpace(token)) batch["confirmToken"] = token;
        return RoundtripJson(batch);
    }

    private static JsonObject MakeExecuteBatch(JsonArray ops, string requestId) =>
        RoundtripJson(new JsonObject
        {
            ["executionMode"] = "execute",
            ["requestId"] = requestId,
            ["expectedDocumentRef"] = ProbePath,
            ["ops"] = ops,
        });

    private readonly record struct ApplyCall(JsonObject Apply, JsonObject? Dry, long ApplyMs, long DryMs);

    private static ApplyCall ApplyBatch(DocBridgeHost host, JsonArray ops)
    {
        if (CurrentMode == ProbeExecutionMode.Execute)
        {
            var exec = Timed(() => host.ApplyOps("excel", MakeExecuteBatch(ops, NewRequestId())));
            return new ApplyCall(exec.Result, Dry: null, exec.Ms, DryMs: 0);
        }

        var dryCall = Timed(() => host.ApplyOps("excel", MakeBatch(ops, dryRun: true)));
        var dryResult = dryCall.Result;
        if (!Json.GetBool(dryResult, "ok"))
        {
            return new ApplyCall(
                new JsonObject
                {
                    ["ok"] = false,
                    ["skipped"] = true,
                    ["errors"] = new JsonArray("apply not attempted; dry-run preflight rejected"),
                },
                dryResult,
                ApplyMs: 0,
                dryCall.Ms);
        }

        var token = Json.GetString(dryResult, "confirmToken");
        var applyCall = Timed(
            () => host.ApplyOps("excel", MakeBatch(CloneOps(ops), dryRun: false, token)));
        return new ApplyCall(applyCall.Result, dryResult, applyCall.Ms, dryCall.Ms);
    }

    private static JsonObject RoundtripJson(JsonObject payload)
    {
        var parsed = JsonNode.Parse(payload.ToJsonString());
        if (parsed is not JsonObject obj)
            throw new InvalidOperationException("ops payload JSON roundtrip did not produce an object");
        return obj;
    }

    private static JsonArray CloneOps(JsonArray ops) => (JsonArray)ops.DeepClone();

    private static JsonObject ReadArgs(string sheet, string range) => RoundtripJson(new JsonObject
    {
        ["workbook"] = ProbePath,
        ["sheet"] = sheet,
        ["range"] = range,
        ["includeStyles"] = true,
    });

    private static void ActivateProbeWorkbook() => ActivateWorkbook(SessionAdapter, SessionApp, ProbePath);

    private static void ActivateWorkbook(ExcelAdapter? adapter, object? app, string? path)
    {
        if (adapter is null || app is null || string.IsNullOrWhiteSpace(path)) return;
        adapter.RunOnAdapterThread<object?>(() =>
        {
            object? workbooks = null;
            object? workbook = null;
            try
            {
                workbooks = (object)((dynamic)app).Workbooks;
                var count = Convert.ToInt32(((dynamic)workbooks).Count, CultureInfo.InvariantCulture);
                for (var i = 1; i <= count; i++)
                {
                    RotHelper.ReleaseComObject(workbook);
                    workbook = (object)((dynamic)workbooks).Item(i);
                    var full = Convert.ToString((object?)((dynamic)workbook).FullName, CultureInfo.InvariantCulture);
                    if (!string.Equals(full, path, StringComparison.OrdinalIgnoreCase)) continue;
                    ((dynamic)workbook).Activate();
                    break;
                }
            }
            finally
            {
                RotHelper.ReleaseComObject(workbook);
                RotHelper.ReleaseComObject(workbooks);
            }
            return null;
        });
    }

    private static JsonObject Phase(string name, JsonObject result, long? wallMs = null)
    {
        var snapshotId = Json.GetString(result, "snapshotId");
        var phase = new JsonObject
        {
            ["name"] = name,
            ["ok"] = Json.GetBool(result, "ok"),
            ["errors"] = Json.GetArr(result, "errors")?.DeepClone() ?? new JsonArray(),
            ["snapshotId"] = snapshotId,
            ["snapshotBytes"] = SnapshotBytes(snapshotId),
            ["snapshotDir"] = SnapshotDir(snapshotId),
            ["hostTimings"] = Json.GetObj(result, "timings")?.DeepClone() ?? new JsonObject(),
            ["readbackVerified"] = Json.GetBool(Json.GetObj(result, "readback"), "verified"),
            ["rollback"] = ExactRollback(result),
            ["operationResults"] = Json.GetArr(result, "operationResults")?.DeepClone(),
            ["replayed"] = result.ContainsKey("replayed") ? result["replayed"]?.DeepClone() : null,
            ["idempotent"] = result.ContainsKey("idempotent") ? result["idempotent"]?.DeepClone() : null,
            ["idempotentReplay"] = result.ContainsKey("idempotentReplay") ? result["idempotentReplay"]?.DeepClone() : null,
            ["requestId"] = Json.GetString(result, "requestId"),
        };
        if (wallMs is not null) phase["wallMs"] = wallMs.Value;
        if (result.ContainsKey("confirmToken") && !string.IsNullOrWhiteSpace(Json.GetString(result, "confirmToken")))
            phase["confirmTokenIssued"] = true;
        StripSecrets(phase);
        return phase;
    }

    private static JsonObject Case(string name, bool ok, long wallMs, JsonArray phases, JsonArray notes, bool executeUnsupported = false) => new()
    {
        ["name"] = name,
        ["ok"] = ok && !executeUnsupported,
        ["skipped"] = executeUnsupported,
        ["skipReason"] = executeUnsupported ? "executeUnsupported" : null,
        ["candidatePass"] = ok && !executeUnsupported,
        ["executeUnsupported"] = executeUnsupported,
        ["wallMs"] = wallMs,
        ["phases"] = phases,
        ["notes"] = notes,
    };

    private static JsonObject CompactResult(JsonObject result)
    {
        var compact = new JsonObject
        {
            ["ok"] = Json.GetBool(result, "ok"),
            ["documentRef"] = Json.GetString(result, "documentRef"),
            ["errors"] = Json.GetArr(result, "errors")?.DeepClone() ?? new JsonArray(),
        };
        var summary = Json.GetObj(result, "summary");
        if (summary is not null)
        {
            compact["activeSheet"] = Json.GetString(summary, "activeSheet");
            compact["usedRange"] = Json.GetString(summary, "usedRange");
            compact["sheets"] = Json.GetArr(summary, "sheets")?.DeepClone();
        }
        return compact;
    }

    private static object CreateProbeWorkbook(object workbooks, string path, int scale, bool bloat)
    {
        object? workbook = null;
        object? sheets = null;
        object? target = null;
        object? bloatSheet = null;
        object? locked = null;
        try
        {
            workbook = (object)((dynamic)workbooks).Add();
            sheets = (object)((dynamic)workbook).Worksheets;
            target = (object)((dynamic)workbook).ActiveSheet;
            ((dynamic)target).Name = ProbeSheet;
            SetCellValue(target, "A1", "canonical");
            SetCellValue(target, "A2", "alias");
            SetCellValue(target, "A3", "reorder");
            SetCellValue(target, "A4", "bystander-op");
            SetCellValue(target, "A5", "execute-replay");
            SetCellValue(target, "A6", "execute-conflict");
            SetCellValue(target, "B1", "fingerprint");
            SetCellValue(target, "C1", "theme");
            SetCellValue(target, "C2", "nofill");
            SetCellValue(target, "C3", "rgb-untinted");
            SetCellValue(target, "C4", "rgb-tint");
            SetCellValue(target, "D1", "fontsize");
            SetCellValue(target, "E1", "first-ok");
            SetCellValue(target, "E2", "later-skip");
            SeedScaleBlock(target, scale);
            if (bloat)
            {
                bloatSheet = (object)((dynamic)sheets).Add(After: target);
                ((dynamic)bloatSheet).Name = BloatSheet;
                SetCellValue(bloatSheet, "A1", "unrelated");
                SetCellFillRaw(bloatSheet, "XFD131", 12632256d);
            }

            locked = (object)((dynamic)sheets).Add(After: bloatSheet ?? target);
            ((dynamic)locked).Name = LockedSheet;
            SetCellValue(locked, "A1", "locked");
            ((dynamic)target).Activate();
            ((dynamic)workbook).SaveAs(path);
            return workbook;
        }
        finally
        {
            RotHelper.ReleaseComObject(locked);
            RotHelper.ReleaseComObject(bloatSheet);
            RotHelper.ReleaseComObject(target);
            RotHelper.ReleaseComObject(sheets);
        }
    }

    private static object CreateBystanderWorkbook(object workbooks, string path)
    {
        object? workbook = null;
        object? sheet = null;
        try
        {
            workbook = (object)((dynamic)workbooks).Add();
            sheet = (object)((dynamic)workbook).ActiveSheet;
            ((dynamic)sheet).Name = BystanderSheet;
            SetCellValue(sheet, "A1", Marker);
            SetCellFillRaw(sheet, "A1", 5287936d);
            ((dynamic)workbook).SaveAs(path);
            return workbook;
        }
        finally { RotHelper.ReleaseComObject(sheet); }
    }

    private static void SeedScaleBlock(object sheet, int scale)
    {
        var rows = ScaleRows(scale);
        if (rows <= 0) return;
        var values = new object[rows, 10];
        for (var r = 0; r < rows; r++)
            for (var c = 0; c < 10; c++)
                values[r, c] = $"s{r}c{c + 1}";
        object? range = null;
        try
        {
            range = (object)((dynamic)sheet).Range(ScaleRange(scale));
            ((dynamic)range).Value2 = values;
        }
        finally { RotHelper.ReleaseComObject(range); }
    }

    private static void SeedMixedStyles(ExcelAdapter adapter, object app)
    {
        adapter.RunOnAdapterThread<object?>(() =>
        {
            var sheet = Sheet(app, ProbeSheet);
            object? c1 = null; object? c2 = null; object? c3 = null;
            object? f1 = null; object? i2 = null; object? i3 = null;
            try
            {
                c1 = (object)((dynamic)sheet).Range("C1");
                c2 = (object)((dynamic)sheet).Range("C2");
                c3 = (object)((dynamic)sheet).Range("C3");
                f1 = (object)((dynamic)c1).Font;
                i2 = (object)((dynamic)c2).Interior;
                i3 = (object)((dynamic)c3).Interior;

                // C1: themed font (not RGB).
                ((dynamic)f1).ThemeColor = XlThemeColorAccent1;
                ((dynamic)f1).TintAndShade = 0.25d;

                // C2: no-fill.
                ((dynamic)i2).Pattern = XlPatternNone;
                ((dynamic)i2).ColorIndex = XlColorIndexNone;

                // C3: non-themed untinted RGB. Color assignment clears theme; do not set TintAndShade.
                ((dynamic)i3).Pattern = 1;
                ((dynamic)i3).Color = UntintedRgbOle;
            }
            finally
            {
                RotHelper.ReleaseComReference(i3);
                RotHelper.ReleaseComReference(i2);
                RotHelper.ReleaseComReference(f1);
                RotHelper.ReleaseComReference(c3);
                RotHelper.ReleaseComReference(c2);
                RotHelper.ReleaseComReference(c1);
                RotHelper.ReleaseComObject(sheet);
            }
            return null;
        });
    }

    private static void SeedUnsupportedRgbTint(ExcelAdapter adapter, object app)
    {
        adapter.RunOnAdapterThread<object?>(() =>
        {
            var sheet = Sheet(app, ProbeSheet);
            object? cell = null;
            object? interior = null;
            try
            {
                cell = (object)((dynamic)sheet).Range("C4");
                interior = (object)((dynamic)cell).Interior;
                ((dynamic)interior).Pattern = 1;
                ((dynamic)interior).Color = UntintedRgbOle;
                ((dynamic)interior).TintAndShade = -0.35d;
            }
            finally
            {
                RotHelper.ReleaseComReference(interior);
                RotHelper.ReleaseComReference(cell);
                RotHelper.ReleaseComObject(sheet);
            }
            return null;
        });
    }

    private static JsonObject ReadComStyle(ExcelAdapter adapter, object app, string sheetName, string address)
    {
        return adapter.RunOnAdapterThread(() =>
        {
            var sheet = Sheet(app, sheetName);
            object? range = null; object? font = null; object? interior = null;
            try
            {
                range = (object)((dynamic)sheet).Range(address);
                font = (object)((dynamic)range).Font;
                interior = (object)((dynamic)range).Interior;
                var fontTheme = TryTheme(() => ((dynamic)font).ThemeColor);
                var fillTheme = TryTheme(() => ((dynamic)interior).ThemeColor);
                return new JsonObject
                {
                    ["fontColor"] = TryDouble(() => ((dynamic)font).Color),
                    ["fontColorIndex"] = TryInt(() => ((dynamic)font).ColorIndex),
                    ["fontThemeColorPresent"] = fontTheme.Present,
                    ["fontThemeColor"] = fontTheme.Value,
                    ["fontTint"] = TryDouble(() => ((dynamic)font).TintAndShade),
                    ["fillColor"] = TryDouble(() => ((dynamic)interior).Color),
                    ["fillColorIndex"] = TryInt(() => ((dynamic)interior).ColorIndex),
                    ["fillThemeColorPresent"] = fillTheme.Present,
                    ["fillThemeColor"] = fillTheme.Value,
                    ["fillTint"] = TryDouble(() => ((dynamic)interior).TintAndShade),
                    ["fillPattern"] = TryInt(() => ((dynamic)interior).Pattern),
                    ["fillPatternColor"] = TryDouble(() => ((dynamic)interior).PatternColor),
                    ["fontBold"] = TryBool(() => ((dynamic)font).Bold),
                    ["fontItalic"] = TryBool(() => ((dynamic)font).Italic),
                };
            }
            finally
            {
                RotHelper.ReleaseComReference(interior);
                RotHelper.ReleaseComReference(font);
                RotHelper.ReleaseComReference(range);
                RotHelper.ReleaseComObject(sheet);
            }
        });
    }

    private static void SetSheetProtected(ExcelAdapter adapter, object app, string sheetName, bool protect)
    {
        adapter.RunOnAdapterThread<object?>(() =>
        {
            var sheet = Sheet(app, sheetName);
            try
            {
                if (protect) ((dynamic)sheet).Protect(ProtectPassword);
                else ((dynamic)sheet).Unprotect(ProtectPassword);
            }
            finally { RotHelper.ReleaseComObject(sheet); }
            return null;
        });
    }

    private static void SetCellFill(ExcelAdapter adapter, object app, string sheetName, string address, double ole)
    {
        adapter.RunOnAdapterThread<object?>(() =>
        {
            var sheet = Sheet(app, sheetName);
            try { SetCellFillRaw(sheet, address, ole); }
            finally { RotHelper.ReleaseComObject(sheet); }
            return null;
        });
    }

    private static string? ReadBystanderMarker(ExcelAdapter adapter, object app)
    {
        return adapter.RunOnAdapterThread(() =>
        {
            object? workbooks = null; object? workbook = null; object? sheet = null; object? range = null;
            try
            {
                workbooks = (object)((dynamic)app).Workbooks;
                var count = Convert.ToInt32(((dynamic)workbooks).Count, CultureInfo.InvariantCulture);
                for (var i = 1; i <= count; i++)
                {
                    RotHelper.ReleaseComObject(workbook);
                    workbook = (object)((dynamic)workbooks).Item(i);
                    var full = Convert.ToString((object?)((dynamic)workbook).FullName, CultureInfo.InvariantCulture);
                    if (!string.Equals(full, BystanderPath, StringComparison.OrdinalIgnoreCase)) continue;
                    sheet = (object)((dynamic)workbook).Worksheets.Item(BystanderSheet);
                    range = (object)((dynamic)sheet).Range("A1");
                    return Convert.ToString((object?)((dynamic)range).Value2, CultureInfo.InvariantCulture);
                }
                return null;
            }
            finally
            {
                RotHelper.ReleaseComObject(range);
                RotHelper.ReleaseComObject(sheet);
                RotHelper.ReleaseComObject(workbook);
                RotHelper.ReleaseComObject(workbooks);
            }
        });
    }

    private static object Sheet(object app, string name)
    {
        object? workbooks = null; object? workbook = null;
        try
        {
            workbooks = (object)((dynamic)app).Workbooks;
            var count = Convert.ToInt32(((dynamic)workbooks).Count, CultureInfo.InvariantCulture);
            for (var i = 1; i <= count; i++)
            {
                RotHelper.ReleaseComObject(workbook);
                workbook = (object)((dynamic)workbooks).Item(i);
                var full = Convert.ToString((object?)((dynamic)workbook).FullName, CultureInfo.InvariantCulture) ?? "";
                if (!string.Equals(full, ProbePath, StringComparison.OrdinalIgnoreCase)) continue;
                return (object)((dynamic)workbook).Worksheets.Item(name);
            }
            throw new InvalidOperationException("owned probe workbook not found");
        }
        finally
        {
            RotHelper.ReleaseComObject(workbook);
            RotHelper.ReleaseComObject(workbooks);
        }
    }

    private static void SaveOwnedCopies(ExcelAdapter adapter, object app, string probePath, string bystanderPath)
    {
        if (!OwnershipProven || !PidMatchesOnSta(adapter, app, OwnedPid)) return;
        adapter.RunOnAdapterThread<object?>(() =>
        {
            SaveIfKnown(app, probePath);
            SaveIfKnown(app, bystanderPath);
            return null;
        });
    }

    private static void SaveIfKnown(object app, string path)
    {
        if (!KnownWorkbookPaths.Contains(path)) return;
        object? workbooks = null; object? workbook = null;
        try
        {
            workbooks = (object)((dynamic)app).Workbooks;
            var count = Convert.ToInt32(((dynamic)workbooks).Count, CultureInfo.InvariantCulture);
            for (var i = 1; i <= count; i++)
            {
                RotHelper.ReleaseComObject(workbook);
                workbook = (object)((dynamic)workbooks).Item(i);
                var full = Convert.ToString((object?)((dynamic)workbook).FullName, CultureInfo.InvariantCulture);
                if (string.Equals(full, path, StringComparison.OrdinalIgnoreCase))
                    ((dynamic)workbook).Save();
            }
        }
        finally
        {
            RotHelper.ReleaseComObject(workbook);
            RotHelper.ReleaseComObject(workbooks);
        }
    }

    private static JsonObject CleanupOwned(ExcelAdapter? adapter, object? app)
    {
        var closed = new JsonArray();
        if (!OwnershipProven || OwnedPid <= 0)
        {
            return new JsonObject
            {
                ["releasedOnly"] = true,
                ["quit"] = false,
                ["killedOffice"] = false,
                ["reason"] = "ownership unproven; no host-thread RCW access",
            };
        }

        var ownedAdapter = adapter;
        var ownedApp = app;
        var pidMatches = ownedAdapter is not null && ownedApp is not null &&
                         PidMatchesOnSta(ownedAdapter, ownedApp, OwnedPid);
        if (pidMatches && ownedAdapter is not null && ownedApp is not null)
        {
            object staApp = ownedApp;
            ownedAdapter.RunOnAdapterThread<object?>(() =>
            {
                object? workbooks = null;
                try
                {
                    workbooks = (object)((dynamic)staApp).Workbooks;
                    for (var i = Convert.ToInt32(((dynamic)workbooks).Count, CultureInfo.InvariantCulture); i >= 1; i--)
                    {
                        object? workbook = null;
                        try
                        {
                            workbook = (object)((dynamic)workbooks).Item(i);
                            var full = Convert.ToString((object?)((dynamic)workbook).FullName, CultureInfo.InvariantCulture);
                            if (full is null || !KnownWorkbookPaths.Contains(full)) continue;
                            ((dynamic)workbook).Close(false);
                            closed.Add(full);
                        }
                        finally { RotHelper.ReleaseComObject(workbook); }
                    }
                }
                finally { RotHelper.ReleaseComObject(workbooks); }
                return null;
            });
        }

        JsonObject? disconnect = null;
        if (pidMatches && ownedAdapter is not null)
            disconnect = ownedAdapter.Disconnect();

        return new JsonObject
        {
            ["closedWorkbooks"] = closed,
            ["disconnect"] = disconnect?.DeepClone(),
            ["provenPid"] = OwnedPid,
            ["killedOffice"] = false,
            ["quitWithoutProof"] = false,
        };
    }

    private static JsonArray InventoryExcel()
    {
        var list = new JsonArray();
        foreach (var process in Process.GetProcessesByName("EXCEL"))
        {
            using (process)
            {
                list.Add(new JsonObject
                {
                    ["pid"] = process.Id,
                    ["processName"] = process.ProcessName,
                    ["mainWindowTitle"] = process.MainWindowTitle,
                });
            }
        }
        return list;
    }

    private static HashSet<int> PidSet(JsonArray inventory)
    {
        var set = new HashSet<int>();
        foreach (var node in inventory)
        {
            if (node is JsonObject o)
            {
                var pid = Json.GetInt(o, "pid");
                if (pid is int value && value > 0) set.Add(value);
            }
        }
        return set;
    }

    private static bool PreexistingStillPresent(JsonArray before, JsonArray? after)
    {
        if (after is null) return false;
        var afterPids = PidSet(after);
        foreach (var pid in PidSet(before))
            if (!afterPids.Contains(pid)) return false;
        return true;
    }

    private static bool PidMatchesOnSta(ExcelAdapter adapter, object app, int expected)
    {
        try
        {
            return adapter.RunOnAdapterThread(() =>
            {
                try { return ProcessIdOf(app) == expected; }
                catch { return false; }
            });
        }
        catch { return false; }
    }

    private static int ProcessIdOf(object app)
    {
        var hwnd = Convert.ToInt64(((dynamic)app).Hwnd, CultureInfo.InvariantCulture);
        return RotHelper.ProcessIdFromWindowHandle(hwnd);
    }

    private static void SetCellValue(object sheet, string address, object value)
    {
        object? range = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            ((dynamic)range).Value2 = value;
        }
        finally { RotHelper.ReleaseComObject(range); }
    }

    private static void SetCellFillRaw(object sheet, string address, double ole)
    {
        object? range = null; object? interior = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            interior = (object)((dynamic)range).Interior;
            ((dynamic)interior).Color = ole;
        }
        finally
        {
            RotHelper.ReleaseComReference(interior);
            RotHelper.ReleaseComReference(range);
        }
    }

    private static void TryCloseWorkbook(object? workbook)
    {
        if (workbook is null) return;
        try { ((dynamic)workbook).Close(false); } catch { /* owned workbook only */ }
    }

    private readonly record struct TimedCall(JsonObject Result, long Ms);

    private static TimedCall Timed(Func<JsonObject> work)
    {
        var sw = Stopwatch.StartNew();
        var result = work();
        sw.Stop();
        return new TimedCall(result, sw.ElapsedMilliseconds);
    }

    private static string? SnapshotDir(string? snapshotId)
    {
        if (Snapshots is null || string.IsNullOrWhiteSpace(snapshotId)) return null;
        return Snapshots.Get(snapshotId)?.Info.Dir;
    }

    private static long SnapshotBytes(string? snapshotId)
    {
        var dir = SnapshotDir(snapshotId);
        if (string.IsNullOrWhiteSpace(dir) || !Directory.Exists(dir)) return 0;
        return Directory.EnumerateFiles(dir, "*", SearchOption.AllDirectories).Sum(f => new FileInfo(f).Length);
    }

    private static string JoinErrors(JsonObject result)
    {
        var parts = new List<string>();
        CollectErrors(result, parts);
        return string.Join(" | ", parts);
    }

    private static void CollectErrors(JsonNode? node, List<string> parts)
    {
        if (node is JsonObject obj)
        {
            if (Json.GetArr(obj, "errors") is { } arr)
                foreach (var item in arr)
                    if (item is JsonValue v && v.TryGetValue<string>(out var s) && !string.IsNullOrWhiteSpace(s))
                        parts.Add(s);
            foreach (var kv in obj)
                if (kv.Key != "errors") CollectErrors(kv.Value, parts);
        }
        else if (node is JsonArray array)
        {
            foreach (var item in array) CollectErrors(item, parts);
        }
    }

    private static bool OpResultOk(JsonArray? results, int index)
    {
        if (results is null || index < 0 || index >= results.Count || results[index] is not JsonObject op)
            return false;
        return Json.GetBool(op, "ok");
    }

    private static bool StylesEqual(JsonObject? left, JsonObject? right)
    {
        if (left is null || right is null) return false;
        foreach (var key in new[]
                 {
                     "fontColor", "fontColorIndex", "fontThemeColorPresent", "fontThemeColor", "fontTint",
                     "fillColor", "fillColorIndex", "fillThemeColorPresent", "fillThemeColor", "fillTint",
                     "fillPattern", "fillPatternColor", "fontBold", "fontItalic",
                 })
        {
            var a = left[key];
            var b = right[key];
            if (a is null && b is null) continue;
            if (a is JsonValue av && b is JsonValue bv)
            {
                if (av.TryGetValue<bool>(out var ab) && bv.TryGetValue<bool>(out var bb))
                {
                    if (ab != bb) return false;
                    continue;
                }
                if (TryNumber(av, out var an) && TryNumber(bv, out var bn))
                {
                    if (!NumbersEqual(an, bn)) return false;
                    continue;
                }
            }
            if (!string.Equals(Json.ToCompact(a), Json.ToCompact(b), StringComparison.Ordinal)) return false;
        }
        return true;
    }

    private static double? StyleNumber(JsonObject? style, string key)
    {
        if (style?[key] is not JsonValue value || !TryNumber(value, out var n)) return null;
        return n;
    }

    private static bool TryNumber(JsonValue value, out double n)
    {
        if (value.TryGetValue<double>(out n)) return true;
        if (value.TryGetValue<int>(out var i)) { n = i; return true; }
        if (value.TryGetValue<long>(out var l)) { n = l; return true; }
        n = 0;
        return false;
    }

    private static bool NumbersEqual(double? a, double? b) =>
        a is not null && b is not null && Math.Abs(a.Value - b.Value) < 0.51;

    private static (bool Present, JsonNode? Value) TryTheme(Func<object?> read)
    {
        try
        {
            var v = read();
            if (v is null or DBNull) return (false, null);
            var n = Convert.ToInt32(v, CultureInfo.InvariantCulture);
            return n is >= 1 and <= 12 ? (true, JsonValue.Create(n)) : (false, JsonValue.Create(n));
        }
        catch
        {
            return (false, null);
        }
    }

    private static JsonNode? TryInt(Func<object?> read)
    {
        try
        {
            var v = read();
            if (v is null or DBNull) return null;
            return JsonValue.Create(Convert.ToInt32(v, CultureInfo.InvariantCulture));
        }
        catch { return null; }
    }

    private static JsonNode? TryDouble(Func<object?> read)
    {
        try
        {
            var v = read();
            if (v is null or DBNull) return null;
            return JsonValue.Create(Convert.ToDouble(v, CultureInfo.InvariantCulture));
        }
        catch { return null; }
    }

    private static JsonNode? TryBool(Func<object?> read)
    {
        try
        {
            var v = read();
            if (v is null or DBNull) return null;
            return JsonValue.Create(Convert.ToBoolean(v, CultureInfo.InvariantCulture));
        }
        catch { return null; }
    }

    private static bool IsUserDocumentsPath(string path)
    {
        var docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        var personal = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
        return StartsWithPath(path, docs) || StartsWithPath(path, personal);
    }

    private static bool StartsWithPath(string path, string root) =>
        !string.IsNullOrWhiteSpace(root) &&
        path.StartsWith(root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
            StringComparison.OrdinalIgnoreCase);

    private static string ColName(int c)
    {
        var name = "";
        while (c > 0) { var m = (c - 1) % 26; name = (char)('A' + m) + name; c = (c - 1) / 26; }
        return name;
    }

    private static JsonObject NewReport(string label, int scale, string mode, int repeat, bool bloat, bool bloatOnly, string outputDir, string homeDir)
    {
        var core = typeof(DocBridgeHost).Assembly;
        return new JsonObject
        {
            ["ok"] = false,
            ["label"] = label,
            ["scale"] = scale,
            ["scaleRange"] = ScaleRange(scale),
            ["mode"] = mode,
            ["repeat"] = repeat,
            ["casesRequested"] = CaseFilter.Count == 0 ? "all" : string.Join(",", CaseFilter.OrderBy(s => s)),
            ["bloat"] = bloat,
            ["bloatOnly"] = bloatOnly,
            ["outputDir"] = outputDir,
            ["homeDir"] = homeDir,
            ["core"] = new JsonObject
            {
                ["location"] = core.Location,
                ["fullName"] = core.FullName,
                ["informationalVersion"] = core.GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion,
                ["hostVersion"] = DocBridgeHost.Version,
            },
            ["probe"] = new JsonObject
            {
                ["exe"] = Environment.ProcessPath,
                ["baseDirectory"] = AppContext.BaseDirectory,
            },
        };
    }

    private static void WriteResult(string path, JsonObject report)
    {
        StripSecrets(report);
        File.WriteAllText(path, Json.ToPretty(report));
    }
}
