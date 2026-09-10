using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Development.ExcelSnapshotStrategyProbe;

/// <summary>
/// Guarded capture-strategy experiment. Requires DOCBRIDGE_E2E=1.
/// New owned Excel PID only. Never Kill. Never Quit unless ownership is proven.
/// Does not set DisplayAlerts. Does not launch until an operator runs this exe.
/// </summary>
internal static partial class Program
{
    private const int XlPatternNone = -4142;
    private const int XlColorIndexNone = -4142;
    private const int XlPatternSolid = 1;
    private const int XlPatternGray16 = 17;
    private const int XlThemeColorAccent1 = 5;
    private const int ExcelThemeReadUnavailable = unchecked((int)0x800A03EC);
    private const double RedOle = 255d;
    private const double BlueOle = 16_711_680d;
    private const double ThemeTint = 0.25d;
    private const int NativeScanChunkCells = 200;
    private const string TargetSheet = "Target";
    private const string KeepSheet = "Keep";
    private const string Marker = "DO-NOT-TOUCH";
    private const string SeedValue = "seed-value";
    private const string DirtyValue = "dirty-value";
    private const string SeedFormula = "=1+1";
    private const string DirtyFormula = "=3+4";

    private static string? ProbePath;
    private static ExcelAdapter? SessionAdapter;
    private static object? SessionApp;
    private static SnapshotService? Snapshots;
    private static bool OwnershipProven;
    private static int OwnedPid;
    private static readonly HashSet<string> KnownWorkbookPaths = new(StringComparer.OrdinalIgnoreCase);
    private static string ProbeCase = "mixed-fill";
    private static bool MixedFill => ProbeCase == "mixed-fill";
    private static bool DeferredRestoreRequested;

    public static int Main(string[] args)
    {
        if (args.Any(a => a is "-h" or "--help" or "/?"))
        {
            Console.WriteLine("""
                DocBridge.ExcelSnapshotStrategyProbe
                Requires DOCBRIDGE_E2E=1. Creates a NEW owned Excel PID only.

                --output <dir>   required. XLSX + JSON written here (not user Documents)
                --scale <n>      required 100, 1000, or 7000 target cells
                                 100=A20:J29  1000=A20:J119  7000=A20:J719
                --repeat <n>     required 1-3 capture-timing repetitions
                --case <name>    optional uniform-bold | mixed-fill (default mixed-fill)
                --deferred-restore  optional. SaveCopyAs checkpoint + deferred native
                                 style extract/restore. Fixture-only; does not change
                                 the default capture-benchmark path. Scale 100|1000 only.

                mixed-fill requests fillColor so capture hits the coupled color snapshot path.
                uniform-bold is an explicit Bold-only control. SaveCopyAs is not scoped rollback.
                One-cell production apply/restore is smoke only. Does not set DisplayAlerts.
                """);
            return 0;
        }

        if (!string.Equals(Environment.GetEnvironmentVariable("DOCBRIDGE_E2E"), "1", StringComparison.Ordinal))
        {
            Console.Error.WriteLine("refusing: set DOCBRIDGE_E2E=1 to run this probe");
            return 2;
        }

        string? outputDir = null;
        int? scale = null;
        int? repeat = null;
        var caseArg = "mixed-fill";
        var deferredRestore = false;
        for (var i = 0; i < args.Length; i++)
        {
            if (args[i] == "--output" && i + 1 < args.Length) outputDir = args[++i];
            else if (args[i] == "--scale" && i + 1 < args.Length)
                scale = int.Parse(args[++i], CultureInfo.InvariantCulture);
            else if (args[i] == "--repeat" && i + 1 < args.Length)
                repeat = int.Parse(args[++i], CultureInfo.InvariantCulture);
            else if (args[i] == "--case" && i + 1 < args.Length)
                caseArg = args[++i];
            else if (args[i] == "--deferred-restore")
                deferredRestore = true;
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

        if (scale is not (100 or 1000 or 7000))
        {
            Console.Error.WriteLine("--scale must be 100, 1000, or 7000");
            return 2;
        }

        if (repeat is not (>= 1 and <= 3))
        {
            Console.Error.WriteLine("--repeat must be 1, 2, or 3");
            return 2;
        }

        if (caseArg is not ("uniform-bold" or "mixed-fill"))
        {
            Console.Error.WriteLine("--case must be uniform-bold or mixed-fill");
            return 2;
        }

        if (deferredRestore && scale is not (100 or 1000))
        {
            Console.Error.WriteLine("--deferred-restore supports --scale 100 or 1000 only");
            return 2;
        }

        ProbeCase = caseArg;
        DeferredRestoreRequested = deferredRestore;
        var scaleValue = scale.Value;
        var repeatValue = repeat.Value;
        outputDir = Path.GetFullPath(outputDir);
        Directory.CreateDirectory(outputDir);
        if (IsUserDocumentsPath(outputDir))
        {
            Console.Error.WriteLine("refusing: --output must not be a user Documents path");
            return 2;
        }

        var homeDir = Path.Combine(outputDir, "probe-home");
        var resultPath = Path.Combine(outputDir, "excel-snapshot-strategy-probe-result.json");
        var probePath = Path.Combine(outputDir, "excel-snapshot-strategy-target.xlsx");
        var baselineCopyPath = Path.Combine(outputDir, "excel-snapshot-strategy-baseline-copy.xlsx");
        var expectedPath = Path.Combine(outputDir, "excel-snapshot-strategy-expected.json");
        foreach (var path in new[] { resultPath, probePath, baselineCopyPath, expectedPath })
        {
            if (File.Exists(path))
            {
                Console.Error.WriteLine($"refusing to overwrite existing path: {path}");
                return 2;
            }
        }

        Directory.CreateDirectory(homeDir);
        var nativeBeforePath = Path.Combine(outputDir, "excel-snapshot-strategy-native-before.json");
        var auditExpectationPath = Path.Combine(outputDir, "excel-snapshot-strategy-audit-expectation.json");
        Console.WriteLine($"INTERFACE case={ProbeCase}");
        Console.WriteLine($"INTERFACE writtenProperties={string.Join(",", WrittenPropertyNames())}");
        Console.WriteLine($"INTERFACE resultJson={resultPath}");
        Console.WriteLine($"INTERFACE targetXlsx={probePath}");
        Console.WriteLine($"INTERFACE baselineCopyXlsx={baselineCopyPath}");
        Console.WriteLine($"INTERFACE expectedJson={expectedPath}");
        Console.WriteLine($"INTERFACE nativeBeforeJson={nativeBeforePath}");
        Console.WriteLine($"INTERFACE auditExpectationJson={auditExpectationPath}");
        Console.WriteLine($"INTERFACE saveCopyAsPattern={Path.Combine(outputDir, "savecopyas-*.xlsx")}");
        Console.WriteLine($"INTERFACE scopedSnapshotDirPattern={Path.Combine(homeDir, "captures", "scoped-rep*-*")}");
        Console.WriteLine($"INTERFACE hostSnapshotDir={Path.Combine(homeDir, "snapshots", "excel")}");
        Console.WriteLine($"INTERFACE deferredRestore={DeferredRestoreRequested}");
        if (DeferredRestoreRequested)
            Console.WriteLine($"INTERFACE deferredResultKey=deferredRestore");
        Progress("CASE", $"{ProbeCase} scale={scaleValue} repeat={repeatValue} deferredRestore={DeferredRestoreRequested} written={string.Join(",", WrittenPropertyNames())}");

        var report = NewReport(
            scaleValue, repeatValue, outputDir, homeDir, resultPath, probePath,
            baselineCopyPath, expectedPath, nativeBeforePath, auditExpectationPath);
        ExcelAdapter? adapter = null;
        DocBridgeHost? host = null;
        object? createdApp = null;
        var preexisting = InventoryExcel();
        report["preexistingExcel"] = preexisting.DeepClone();
        var exitCode = 3;

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
                try
                {
                    dynamic app = appObj;
                    app.Visible = false;
                    workbooks = (object)app.Workbooks;
                    ProbePath = probePath;
                    KnownWorkbookPaths.Add(probePath);
                    var setupSw = Stopwatch.StartNew();
                    probe = CreateOwnedWorkbook(workbooks, probePath, scaleValue);
                    setupSw.Stop();
                    report["setup"] = new JsonObject
                    {
                        ["createAndSaveBaselineMs"] = setupSw.ElapsedMilliseconds,
                        ["note"] = "setup is isolated from capture/apply/restore/native validation",
                    };
                    ((dynamic)probe).Activate();
                    return appObj;
                }
                catch
                {
                    TryCloseWorkbook(probe);
                    throw;
                }
                finally
                {
                    RotHelper.ReleaseComReference(probe);
                    RotHelper.ReleaseComReference(workbooks);
                }
            }, appFactoryOwnsInstance: true);

            host = new DocBridgeHost(options);
            host.Router.Register("excel", adapter);
            SessionAdapter = adapter;
            var ctx = host.GetActiveContext("excel");
            SessionApp = createdApp;
            if (!OwnershipProven || OwnedPid <= 0)
                throw new InvalidOperationException("owned Excel PID was not proven");

            File.Copy(probePath, baselineCopyPath, overwrite: false);
            WriteExpectedCatalog(expectedPath, scaleValue, probePath, baselineCopyPath);

            var identityAfterSave = ReadIdentity(adapter, createdApp!);
            ApplyUnsavedMutations(adapter, createdApp!, scaleValue);
            var identityAfterDirty = ReadIdentity(adapter, createdApp!);
            report["identity"] = new JsonObject
            {
                ["afterSavedBaseline"] = identityAfterSave,
                ["afterUnsavedMutations"] = identityAfterDirty,
                ["unsavedSavedFlag"] = Json.GetBool(identityAfterDirty, "saved") == false,
            };

            Progress("native-scan", "native-before start (pre-SaveCopyAs)");
            var nativeBefore = ScanScaleCells(adapter, createdApp!, scaleValue, compareTo: null);
            var keepBefore = ReadCellAudit(adapter, createdApp!, KeepSheet, "A1");
            WriteNativeBeforeArtifact(nativeBeforePath, nativeBefore, keepBefore, identityAfterDirty, probePath);
            WriteAuditExpectation(
                auditExpectationPath, workbookPath: null, probePath, identityAfterDirty, nativeBefore, keepBefore);
            report["nativeBefore"] = NativePhase("native-before-pre-savecopyas", nativeBefore);
            Progress("native-scan", $"native-before done ok={nativeBefore.Ok} cells={nativeBefore.Checked} unreadable={nativeBefore.Unreadable} dir-pending");

            report["ownership"] = new JsonObject
            {
                ["ownedPid"] = OwnedPid,
                ["ownedPidIsNew"] = true,
                ["provenBeforeWorkbooks"] = true,
                ["probeWorkbook"] = probePath,
                ["appFactoryOwnsInstance"] = true,
                ["killedOffice"] = false,
                ["displayAlertsOverride"] = false,
            };

            report["rejectLimitation"] = RunUnknownStyleKeyReject(host);
            if (DeferredRestoreRequested)
            {
                report["repeats"] = new JsonArray();
                report["captureBenchmarkSkipped"] = true;
                report["applyRestore"] = new JsonObject
                {
                    ["skipped"] = true,
                    ["reason"] = "host apply/restore smoke skipped; deferred path uses ExcelAdapter.Apply after checkpoint",
                };
                report["deferredRestore"] = RunDeferredRestoreExperiment(
                    adapter, createdApp!, outputDir, scaleValue, nativeBefore, keepBefore);
                report["ok"] = nativeBefore.Ok &&
                               Json.GetBool(Json.GetObj(report, "rejectLimitation"), "ok") &&
                               Json.GetBool(Json.GetObj(report, "deferredRestore"), "ok");
            }
            else
            {
                var repeats = new JsonArray();
                for (var rep = 1; rep <= repeatValue; rep++)
                {
                    repeats.Add(RunCaptureRepeat(
                        adapter, createdApp!, host, outputDir, homeDir, scaleValue, rep,
                        nativeBefore, auditExpectationPath, keepBefore, identityAfterDirty));
                }

                report["repeats"] = repeats;
                report["applyRestore"] = RunScopedApplyRestore(host, adapter, createdApp!, scaleValue);
                report["ok"] = nativeBefore.Ok &&
                               RepeatsOk(repeats) &&
                               Json.GetBool(Json.GetObj(report, "rejectLimitation"), "ok") &&
                               Json.GetBool(Json.GetObj(report, "applyRestore"), "ok");
            }
            exitCode = Json.GetBool(report, "ok") ? 0 : 4;
        }
        catch (Exception ex)
        {
            report["ok"] = false;
            report["harnessError"] = ex.Message;
            report["harnessType"] = ex.GetType().FullName;
            if (ex is COMException com)
                report["harnessHresult"] = com.ErrorCode;
            Console.Error.WriteLine(ex);
            exitCode = 3;
        }
        finally
        {
            Progress("cleanup", "start");
            var cleanupOk = false;
            var pidsPreserved = false;
            try
            {
                report["cleanup"] = CleanupOwned(adapter, createdApp);
                report["excelAfterCleanup"] = InventoryExcel();
                pidsPreserved = PreexistingStillPresent(preexisting, report["excelAfterCleanup"] as JsonArray);
                report["preexistingPidsStillPresent"] = pidsPreserved;
                cleanupOk = CleanupSucceeded(report);
            }
            catch (Exception cleanupEx)
            {
                report["cleanupError"] = cleanupEx.Message;
                cleanupOk = false;
            }

            report["cleanupOk"] = cleanupOk;
            if (!cleanupOk || !pidsPreserved)
            {
                report["ok"] = false;
                if (exitCode == 0) exitCode = 5;
            }

            host?.Dispose();
            WriteResult(resultPath, report);
            Progress("cleanup", $"ok={cleanupOk} preexistingPidsStillPresent={pidsPreserved} result.ok={Json.GetBool(report, "ok")} exit={exitCode}");
            Console.WriteLine(resultPath);
        }

        return exitCode;
    }

    private static JsonObject NewReport(
        int scale, int repeat, string outputDir, string homeDir,
        string resultPath, string probePath, string baselineCopyPath, string expectedPath,
        string nativeBeforePath, string auditExpectationPath)
    {
        var deferred = DeferredRestoreRequested;
        return new JsonObject
        {
            ["probe"] = "ExcelSnapshotStrategyProbe",
            ["sourceVersion"] = deferred
                ? "local experimental probe implementation with --deferred-restore; not v0.4.20 public baseline"
                : "local experimental probe; default capture-benchmark path; not a published v0.4.20 public-baseline package",
            ["publicBaselinePackage"] = false,
            ["ok"] = false,
            ["deferredRestoreRequested"] = deferred,
            ["case"] = ProbeCase,
            ["writtenProperties"] = ToJsonArray(WrittenPropertyNames()),
            ["coupledFillFields"] = MixedFill ? ToJsonArray(CoupledFillFieldNames()) : new JsonArray(),
            ["scale"] = scale,
            ["repeat"] = repeat,
            ["targetRange"] = ScaleRange(scale),
            ["targetGeometry"] = new JsonObject
            {
                ["sheet"] = TargetSheet,
                ["range"] = ScaleRange(scale),
                ["rows"] = ScaleRows(scale),
                ["columns"] = 10,
                ["cells"] = scale,
                ["startRow"] = 20,
                ["startColumn"] = 1,
            },
            ["timingBoundary"] = new JsonObject
            {
                ["scopedCapture"] = "ExcelAdapter.CaptureSnapshot on STA via ComInvoke. Format-only ops write written-properties state.json. Auxiliary workbook-backup source is taken from CaptureSnapshot metadata.workbookBackupSource (dirty ordinary .xlsx uses current-memory-savecopyas; other cases last-saved-file or none). STA 120s is unchanged.",
                ["saveCopyAs"] = "Workbook.SaveCopyAs writes current memory to a unique file. It is not Save, not operation-scoped rollback, and must not change FullName/Saved/document count.",
                ["excludedFromCaptureMs"] = "setup, unsaved mutation, native scans, apply, restore",
                ["nativeRangeCopy"] = "not implemented; secondary candidate",
                ["candidateRestore"] = deferred
                    ? "local experimental probe: SaveCopyAs checkpoint + SHA/OOXML preflight + deferred native extract/restore after eligibility; fixture-only generated .xlsx scale 100|1000; not a product API or generic rollout"
                    : "unimplemented on the default capture-benchmark path; applyRestore is production smoke only",
            },
            ["claims"] = new JsonObject
            {
                ["saveCopyAsIsOperationScopedRollback"] = false,
                ["missingNativeReadsCoercedToFalseOrZero"] = false,
                ["timeoutsRaised"] = false,
                ["seededTintComparedExactToNativeWithoutRecordingBoth"] = false,
                ["candidateRestoreImplemented"] = deferred,
                ["publicBaselinePackage"] = false,
            },
            ["artifacts"] = new JsonObject
            {
                ["resultJson"] = resultPath,
                ["targetXlsx"] = probePath,
                ["baselineCopyXlsx"] = baselineCopyPath,
                ["expectedJson"] = expectedPath,
                ["nativeBeforeJson"] = nativeBeforePath,
                ["auditExpectationJson"] = auditExpectationPath,
                ["outputDir"] = outputDir,
                ["probeHome"] = homeDir,
                ["scopedSnapshotDirPattern"] = Path.Combine(homeDir, "captures", "scoped-rep*-*"),
            },
            ["pendingWork"] = deferred
                ? new JsonArray(
                    "Deferred restore is probe-local experimental implementation, not a product API or generic rollout",
                    "Native range-copy strategy is secondary and not implemented",
                    "Root compares this exe logic before accepting native evidence")
                : new JsonArray(
                    "SaveCopyAs-based restore is not implemented and must not be claimed as scoped rollback",
                    "Candidate native-checkpoint restore remains unimplemented on the default capture-benchmark path",
                    "Native range-copy strategy is secondary and not implemented",
                    "Codex owns live execution of this compiled exe"),
        };
    }

    private static JsonObject RunCaptureRepeat(
        ExcelAdapter adapter, object app, DocBridgeHost host,
        string outputDir, string homeDir, int scale, int rep,
        ScaleCellScan nativeBefore, string auditExpectationPath,
        JsonObject keepBefore, JsonObject identityAfterDirty)
    {
        ActivateOwnedWorkbook(adapter, app);
        var before = ReadIdentity(adapter, app);
        var ops = FormatOpList(ScaleRange(scale));
        var snapshotDir = Path.Combine(homeDir, "captures", $"scoped-rep{rep}-{Guid.NewGuid():N}");
        Directory.CreateDirectory(snapshotDir);
        Console.WriteLine($"INTERFACE scopedCaptureDir={snapshotDir}");
        var saveCopyAsFirst = rep == 2;
        var order = saveCopyAsFirst ? "savecopyas-then-scoped" : "scoped-then-savecopyas";
        Progress("CASE", $"{ProbeCase} repeat={rep} order={order}");

        JsonObject scoped;
        JsonObject copy;
        var copyPath = Path.Combine(outputDir, $"savecopyas-rep{rep}-{Guid.NewGuid():N}.xlsx");
        if (saveCopyAsFirst)
        {
            copy = TimeSaveCopyAs(adapter, app, copyPath);
            scoped = TimeScopedCapture(adapter, snapshotDir, ops);
        }
        else
        {
            scoped = TimeScopedCapture(adapter, snapshotDir, ops);
            copy = TimeSaveCopyAs(adapter, app, copyPath);
        }

        if (rep == 1)
        {
            WriteAuditExpectation(
                auditExpectationPath, copyPath, ProbePath ?? "", identityAfterDirty, nativeBefore, keepBefore);
        }

        var after = ReadIdentity(adapter, app);
        Progress("native-scan", $"native-after-captures start rep={rep}");
        var native = ScanScaleCells(adapter, app, scale, compareTo: nativeBefore);
        Progress("native-scan", $"native-after-captures done rep={rep} ok={native.Ok} mismatches={native.Mismatches}");
        var identityOk = IdentityUnchanged(before, after) &&
                         Json.GetBool(before, "saved") == false &&
                         Json.GetBool(after, "saved") == false &&
                         Json.GetBool(copy, "ok");
        return new JsonObject
        {
            ["repeat"] = rep,
            ["case"] = ProbeCase,
            ["writtenProperties"] = ToJsonArray(WrittenPropertyNames()),
            ["coupledFillFields"] = MixedFill ? ToJsonArray(CoupledFillFieldNames()) : new JsonArray(),
            ["captureOrder"] = order,
            ["snapshotDir"] = snapshotDir,
            ["ok"] = identityOk && native.Ok && Json.GetBool(scoped, "ok"),
            ["identityBefore"] = before,
            ["identityAfter"] = after,
            ["identityPreserved"] = identityOk,
            ["scopedCapture"] = scoped,
            ["saveCopyAs"] = copy,
            ["nativeScan"] = NativePhase("native-after-captures", native),
            ["note"] = "capture ms exclude setup/scan/apply/restore; after-scan compares to native-before observations, not seeded tint",
            ["hostContextTouched"] = host.GetType().Name,
            ["candidateRestoreImplemented"] = false,
        };
    }

    private static JsonObject TimeScopedCapture(ExcelAdapter adapter, string snapshotDir, List<JsonObject> ops)
    {
        Progress("capture-start", $"dir={snapshotDir}");
        var result = TimeCapture(snapshotDir, () =>
        {
            var metadata = new JsonObject();
            adapter.CaptureSnapshot(snapshotDir, metadata, ops);
            return metadata;
        });
        Progress("capture-done", $"dir={snapshotDir} ok={Json.GetBool(result, "ok")} wallMs={Json.GetLong(result, "wallMs")}");
        return result;
    }

    private static JsonObject TimeCapture(string snapshotDir, Func<JsonObject> work)
    {
        var sw = Stopwatch.StartNew();
        try
        {
            var metadata = work();
            sw.Stop();
            return new JsonObject
            {
                ["ok"] = true,
                ["wallMs"] = sw.ElapsedMilliseconds,
                ["method"] = "ExcelAdapter.CaptureSnapshot",
                ["snapshotDir"] = snapshotDir,
                ["payload"] = Json.GetString(metadata, "payload"),
                ["documentRef"] = Json.GetString(metadata, "documentRef"),
                ["workbookBackup"] = Json.GetString(metadata, "workbookBackup"),
                ["workbookBackupError"] = Json.GetString(metadata, "workbookBackupError"),
                ["styleScope"] = "written-properties when format-only",
                ["writtenProperties"] = ToJsonArray(WrittenPropertyNames()),
                ["metadata"] = metadata.DeepClone(),
                ["workbookBackupSource"] = Json.GetString(metadata, "workbookBackupSource"),
                ["workbookBackupFresh"] = metadata["workbookBackupFresh"]?.DeepClone(),
                ["workbookBackupReason"] = Json.GetString(metadata, "workbookBackupReason"),
                ["workbookBackupIsLastSavedDiskFile"] = string.Equals(
                    Json.GetString(metadata, "workbookBackupSource"),
                    "last-saved-file",
                    StringComparison.Ordinal),
                ["workbookBackupIsCurrentMemory"] = string.Equals(
                    Json.GetString(metadata, "workbookBackupSource"),
                    "current-memory-savecopyas",
                    StringComparison.Ordinal),
            };
        }
        catch (Exception ex)
        {
            sw.Stop();
            var fail = Fail("ExcelAdapter.CaptureSnapshot", sw.ElapsedMilliseconds, ex);
            fail["snapshotDir"] = snapshotDir;
            return fail;
        }
    }

    private static JsonObject TimeSaveCopyAs(ExcelAdapter adapter, object app, string copyPath)
    {
        return adapter.RunOnAdapterThread(() =>
        {
            object? workbooks = null;
            object? workbook = null;
            var sw = Stopwatch.StartNew();
            try
            {
                workbooks = (object)((dynamic)app).Workbooks;
                workbook = FindOwnedWorkbook(workbooks);
                var beforeName = Convert.ToString((object?)((dynamic)workbook).FullName, CultureInfo.InvariantCulture);
                var beforeSaved = (bool)((dynamic)workbook).Saved;
                var beforeCount = Convert.ToInt32(((dynamic)workbooks).Count, CultureInfo.InvariantCulture);
                ((dynamic)workbook).SaveCopyAs(copyPath);
                sw.Stop();
                var afterName = Convert.ToString((object?)((dynamic)workbook).FullName, CultureInfo.InvariantCulture);
                var afterSaved = (bool)((dynamic)workbook).Saved;
                var afterCount = Convert.ToInt32(((dynamic)workbooks).Count, CultureInfo.InvariantCulture);
                var ok = string.Equals(beforeName, afterName, StringComparison.OrdinalIgnoreCase) &&
                         beforeSaved == afterSaved &&
                         beforeCount == afterCount &&
                         File.Exists(copyPath);
                Progress("SaveCopyAs-done", $"path={copyPath} ok={ok} wallMs={sw.ElapsedMilliseconds}");
                return new JsonObject
                {
                    ["ok"] = ok,
                    ["wallMs"] = sw.ElapsedMilliseconds,
                    ["method"] = "Workbook.SaveCopyAs",
                    ["path"] = copyPath,
                    ["bytes"] = File.Exists(copyPath) ? new FileInfo(copyPath).Length : 0,
                    ["fullNameBefore"] = beforeName,
                    ["fullNameAfter"] = afterName,
                    ["savedBefore"] = beforeSaved,
                    ["savedAfter"] = afterSaved,
                    ["workbookCountBefore"] = beforeCount,
                    ["workbookCountAfter"] = afterCount,
                    ["originalBaselinePath"] = ProbePath,
                    ["operationScopedRollback"] = false,
                };
            }
            catch (Exception ex)
            {
                sw.Stop();
                Progress("SaveCopyAs-done", $"path={copyPath} ok=false wallMs={sw.ElapsedMilliseconds}");
                return Fail("Workbook.SaveCopyAs", sw.ElapsedMilliseconds, ex);
            }
            finally
            {
                RotHelper.ReleaseComReference(workbook);
                RotHelper.ReleaseComReference(workbooks);
            }
        });
    }

    private static JsonObject RunUnknownStyleKeyReject(DocBridgeHost host)
    {
        var sw = Stopwatch.StartNew();
        ActivateProbe();
        var ops = new JsonArray(new JsonObject
        {
            ["op"] = "format_range",
            ["target"] = new JsonObject { ["sheet"] = TargetSheet },
            ["range"] = "C4",
            ["style"] = new JsonObject
            {
                ["fillColor"] = RedOle,
                ["fillTint"] = -0.35d,
            },
        });
        var dry = host.ApplyOps("excel", new JsonObject { ["ops"] = ops, ["dryRun"] = true });
        sw.Stop();
        var errors = JoinErrors(dry);
        var rejected = !Json.GetBool(dry, "ok") &&
                       (errors.Contains("fillTint", StringComparison.Ordinal) ||
                        errors.Contains("is not supported", StringComparison.OrdinalIgnoreCase));
        return new JsonObject
        {
            ["ok"] = rejected,
            ["name"] = "unknown-style-key-reject",
            ["wallMs"] = sw.ElapsedMilliseconds,
            ["successCase"] = false,
            ["errors"] = errors,
            ["provesNativeRgbTintSnapshotRejection"] = false,
            ["note"] = "unknown style-key validation only: dry-run sends unsupported fillTint. Does not prove native RGB+nonzero-tint snapshot rejection.",
        };
    }

    private static JsonObject RunScopedApplyRestore(DocBridgeHost host, ExcelAdapter adapter, object app, int scale)
    {
        ActivateOwnedWorkbook(adapter, app);
        var before = ReadCellAudit(adapter, app, TargetSheet, "A3");
        var swApply = Stopwatch.StartNew();
        var apply = host.ApplyOps("excel", new JsonObject
        {
            ["ops"] = new JsonArray(new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = new JsonObject { ["sheet"] = TargetSheet },
                ["range"] = "A3",
                ["style"] = new JsonObject { ["bold"] = true },
            }),
            ["executionMode"] = "execute",
            ["requestId"] = Guid.NewGuid().ToString("D"),
            ["expectedDocumentRef"] = ProbePath,
        });
        swApply.Stop();
        var after = ReadCellAudit(adapter, app, TargetSheet, "A3");
        var snapshotId = Json.GetString(apply, "snapshotId");
        JsonObject restore;
        long restoreMs;
        if (string.IsNullOrWhiteSpace(snapshotId))
        {
            restore = new JsonObject { ["ok"] = false, ["errors"] = new JsonArray("snapshotId missing") };
            restoreMs = 0;
        }
        else
        {
            var dry = Timed(() => host.CoreRestoreSnapshot(new JsonObject { ["snapshotId"] = snapshotId }));
            var confirmed = Timed(() => host.CoreRestoreSnapshot(new JsonObject
            {
                ["snapshotId"] = snapshotId,
                ["confirmToken"] = Json.GetString(dry.Result, "confirmToken"),
            }));
            restore = confirmed.Result;
            restoreMs = confirmed.Ms;
            restore["restoreDryMs"] = dry.Ms;
        }

        var restored = ReadCellAudit(adapter, app, TargetSheet, "A3");
        var keep = ReadCellAudit(adapter, app, KeepSheet, "A1");
        var nativeOk = AuditReadable(before) && AuditReadable(after) && AuditReadable(restored) &&
                       Json.GetBool(before, "fontBold") == Json.GetBool(restored, "fontBold") &&
                       Json.GetString(keep, "value") == Marker;
        return new JsonObject
        {
            ["ok"] = Json.GetBool(apply, "ok") && Json.GetBool(restore, "ok") && nativeOk,
            ["applyMs"] = swApply.ElapsedMilliseconds,
            ["restoreMs"] = restoreMs,
            ["snapshotId"] = snapshotId,
            ["apply"] = CompactResult(apply),
            ["restore"] = CompactResult(restore),
            ["nativeBefore"] = before,
            ["nativeAfterApply"] = after,
            ["nativeAfterRestore"] = restored,
            ["bystanderKeep"] = keep,
            ["scaleContextCells"] = scale,
            ["role"] = "production-baseline-smoke-only",
            ["candidateRestoreImplemented"] = false,
            ["candidateNativeCheckpointRestore"] = "unimplemented",
            ["note"] = "baseline smoke only: one-cell A3 production format_range execute + core_restore_snapshot. Candidate SaveCopyAs/native-checkpoint restore remains unimplemented. Do not treat this as the measured candidate.",
        };
    }

    private static object CreateOwnedWorkbook(object workbooks, string path, int scale)
    {
        object? workbook = null;
        object? sheets = null;
        object? target = null;
        object? keep = null;
        try
        {
            workbook = (object)((dynamic)workbooks).Add();
            sheets = (object)((dynamic)workbook).Worksheets;
            target = (object)((dynamic)workbook).ActiveSheet;
            ((dynamic)target).Name = TargetSheet;
            SetCellValue(target, "A1", SeedValue);
            SetCellFormula(target, "A2", SeedFormula);
            SetCellValue(target, "A3", "uniform-bold-restore");
            SetCellValue(target, "C4", "rgb-tint-reject");
            SeedScaleBlock(target, scale);
            SeedScaleStyles(target, scale, dirty: false);
            keep = (object)((dynamic)sheets).Add(After: target);
            ((dynamic)keep).Name = KeepSheet;
            SetCellValue(keep, "A1", Marker);
            SetCellFillRaw(keep, "A1", 5_287_936d);
            ((dynamic)target).Activate();
            ((dynamic)workbook).SaveAs(path);
            return workbook;
        }
        finally
        {
            RotHelper.ReleaseComReference(keep);
            RotHelper.ReleaseComReference(target);
            RotHelper.ReleaseComReference(sheets);
        }
    }

    private static void ApplyUnsavedMutations(ExcelAdapter adapter, object app, int scale)
    {
        adapter.RunOnAdapterThread<object?>(() =>
        {
            var sheet = Sheet(app, TargetSheet);
            try
            {
                SetCellValue(sheet, "A1", DirtyValue);
                SetCellFormula(sheet, "A2", DirtyFormula);
                SetCellBold(sheet, "A3", false);
                if (MixedFill)
                {
                    var first = ScaleAddress(0);
                    SetCellFillRaw(sheet, first, BlueOle);
                    SetCellBold(sheet, first, true);
                }
            }
            finally { RotHelper.ReleaseComReference(sheet); }
            return null;
        });
        _ = scale;
    }

    private static void SeedScaleBlock(object sheet, int scale)
    {
        var rows = ScaleRows(scale);
        var values = new object[rows, 10];
        for (var r = 0; r < rows; r++)
        {
            for (var c = 0; c < 10; c++)
                values[r, c] = c == 9 ? (object)$"={100 + r}" : $"s{r}c{c + 1}";
        }

        object? range = null;
        try
        {
            range = (object)((dynamic)sheet).Range(ScaleRange(scale));
            ((dynamic)range).Value2 = values;
        }
        finally { RotHelper.ReleaseComReference(range); }
    }

    private static void SeedScaleStyles(object sheet, int scale, bool dirty)
    {
        if (!MixedFill)
        {
            SeedUniformBold(sheet, scale);
            return;
        }

        var expected = scale;
        for (var start = 0; start < expected; start += NativeScanChunkCells)
        {
            var count = Math.Min(NativeScanChunkCells, expected - start);
            for (var i = 0; i < count; i++)
            {
                var index = start + i;
                var kind = StyleKind(index, dirty && index == 0);
                object? range = null;
                object? font = null;
                object? interior = null;
                try
                {
                    range = (object)((dynamic)sheet).Range(ScaleAddress(index));
                    font = (object)((dynamic)range).Font;
                    interior = (object)((dynamic)range).Interior;
                    ApplyKind((dynamic)font, (dynamic)interior, kind);
                }
                finally
                {
                    RotHelper.ReleaseComReference(interior);
                    RotHelper.ReleaseComReference(font);
                    RotHelper.ReleaseComReference(range);
                }
            }
        }
    }

    private static void SeedUniformBold(object sheet, int scale)
    {
        object? range = null;
        object? font = null;
        object? interior = null;
        try
        {
            range = (object)((dynamic)sheet).Range(ScaleRange(scale));
            font = (object)((dynamic)range).Font;
            interior = (object)((dynamic)range).Interior;
            ((dynamic)font).Bold = true;
            ((dynamic)interior).Pattern = XlPatternNone;
            ((dynamic)interior).ColorIndex = XlColorIndexNone;
        }
        finally
        {
            RotHelper.ReleaseComReference(interior);
            RotHelper.ReleaseComReference(font);
            RotHelper.ReleaseComReference(range);
        }
    }

    private static void ApplyKind(dynamic font, dynamic interior, int kind)
    {
        switch (kind)
        {
            case 0:
                font.Bold = true;
                interior.Pattern = XlPatternNone;
                interior.ColorIndex = XlColorIndexNone;
                break;
            case 1:
                font.Bold = true;
                interior.Pattern = XlPatternSolid;
                interior.Color = RedOle;
                break;
            case 2:
                font.Bold = false;
                interior.Pattern = XlPatternSolid;
                interior.ThemeColor = XlThemeColorAccent1;
                interior.TintAndShade = ThemeTint;
                break;
            case 3:
                font.Bold = true;
                interior.Pattern = XlPatternGray16;
                interior.Color = RedOle;
                interior.PatternColor = BlueOle;
                break;
            default:
                font.Bold = true;
                interior.Pattern = XlPatternSolid;
                interior.Color = BlueOle;
                break;
        }
    }

    private static int StyleKind(int index, bool dirtyFirst)
    {
        if (!MixedFill) return 0;
        return dirtyFirst ? 4 : index % 4;
    }

    private sealed class ScaleCellScan
    {
        public int Checked { get; set; }
        public int Expected { get; set; }
        public int Chunks { get; set; }
        public int Unreadable { get; set; }
        public int Mismatches { get; set; }
        public int SemanticDrift { get; set; }
        public long Ms { get; set; }
        public JsonArray Cells { get; } = new();
        public JsonArray Samples { get; } = new();
        public bool Ok => Checked == Expected && Mismatches == 0 && Unreadable == 0;
    }

    private static ScaleCellScan ScanScaleCells(ExcelAdapter adapter, object app, int scale, ScaleCellScan? compareTo)
    {
        var scan = new ScaleCellScan { Expected = scale };
        var sw = Stopwatch.StartNew();
        for (var start = 0; start < scale; start += NativeScanChunkCells)
        {
            var count = Math.Min(NativeScanChunkCells, scale - start);
            scan.Chunks++;
            var prior = SliceCells(compareTo?.Cells, start, count);
            var chunk = adapter.RunOnAdapterThread(() => ScanChunk(app, start, count, prior));
            scan.Checked += chunk.Checked;
            scan.Unreadable += chunk.Unreadable;
            scan.Mismatches += chunk.Mismatches;
            scan.SemanticDrift += chunk.SemanticDrift;
            foreach (var cell in chunk.Cells)
            {
                if (cell is null) continue;
                scan.Cells.Add(cell.DeepClone());
            }

            foreach (var sample in chunk.Samples)
            {
                if (sample is null || scan.Samples.Count >= 12) continue;
                scan.Samples.Add(sample.DeepClone());
            }
        }

        scan.Ms = sw.ElapsedMilliseconds;
        return scan;
    }

    private static JsonArray? SliceCells(JsonArray? cells, int start, int count)
    {
        if (cells is null || cells.Count == 0) return null;
        var slice = new JsonArray();
        for (var i = 0; i < count; i++)
        {
            var idx = start + i;
            slice.Add(idx < cells.Count && cells[idx] is JsonNode n ? n.DeepClone() : null);
        }

        return slice;
    }

    private sealed class ChunkResult
    {
        public int Checked;
        public int Unreadable;
        public int Mismatches;
        public int SemanticDrift;
        public JsonArray Cells { get; } = new();
        public JsonArray Samples { get; } = new();
    }

    private static ChunkResult ScanChunk(object app, int start, int count, JsonArray? prior)
    {
        var result = new ChunkResult();
        var sheet = Sheet(app, TargetSheet);
        try
        {
            for (var i = 0; i < count; i++)
            {
                var index = start + i;
                object? range = null;
                object? font = null;
                object? interior = null;
                try
                {
                    range = (object)((dynamic)sheet).Range(ScaleAddress(index));
                    font = (object)((dynamic)range).Font;
                    interior = (object)((dynamic)range).Interior;
                    var cell = ReadNativeCell((dynamic)range, (dynamic)font, (dynamic)interior, index);
                    result.Cells.Add(cell);
                    if (Json.GetBool(cell, "unreadable"))
                    {
                        result.Unreadable++;
                        AddSample(result, cell);
                        continue;
                    }

                    result.Checked++;
                    if (prior is not null && !ObservationsMatch(prior[i] as JsonObject, cell))
                    {
                        result.Mismatches++;
                        AddSample(result, cell);
                    }

                    if (Json.GetBool(cell, "semanticDrift"))
                        result.SemanticDrift++;
                }
                finally
                {
                    RotHelper.ReleaseComReference(interior);
                    RotHelper.ReleaseComReference(font);
                    RotHelper.ReleaseComReference(range);
                }
            }
        }
        finally { RotHelper.ReleaseComReference(sheet); }

        return result;
    }

    private static void AddSample(ChunkResult result, JsonObject cell)
    {
        if (result.Samples.Count >= 12) return;
        result.Samples.Add(cell.DeepClone());
    }

    private static JsonObject ReadNativeCell(dynamic range, dynamic font, dynamic interior, int index)
    {
        var address = ScaleAddress(index);
        var kind = StyleKind(index, MixedFill && index == 0);
        var seeded = SeededSemantics(kind);
        var required = RequiredNativeFields();
        var unreadable = new JsonArray();
        var style = new JsonObject();

        void NeedBool(string name, bool? value)
        {
            if (value is null)
            {
                if (required.Contains(name)) unreadable.Add(name);
                return;
            }

            style[name] = value.Value;
        }

        void NeedInt(string name, int? value)
        {
            if (value is null)
            {
                if (required.Contains(name)) unreadable.Add(name);
                return;
            }

            style[name] = value.Value;
        }

        void NeedDouble(string name, double? value)
        {
            if (value is null)
            {
                if (required.Contains(name)) unreadable.Add(name);
                return;
            }

            style[name] = value.Value;
        }

        NeedBool("bold", TryBool(() => (object?)font.Bold));
        NeedDouble("fillColor", TryDouble(() => (object?)interior.Color));
        NeedInt("fillColorIndex", TryInt(() => (object?)interior.ColorIndex));
        NeedDouble("fillTintAndShade", TryDouble(() => (object?)interior.TintAndShade));
        NeedInt("fillPattern", TryInt(() => (object?)interior.Pattern));
        NeedDouble("fillPatternColor", TryDouble(() => (object?)interior.PatternColor));
        NeedInt("fillPatternColorIndex", TryInt(() => (object?)interior.PatternColorIndex));
        NeedDouble("fillPatternTintAndShade", TryDouble(() => (object?)interior.PatternTintAndShade));

        var fillTheme = ReadThemeColor(() => (object?)interior.ThemeColor, $"{address} Interior.ThemeColor");
        var patternTheme = ReadThemeColor(() => (object?)interior.PatternThemeColor, $"{address} Interior.PatternThemeColor");
        if (fillTheme.Status == "read-failed") unreadable.Add("fillThemeColor");
        if (patternTheme.Status == "read-failed") unreadable.Add("fillPatternThemeColor");
        if (fillTheme.Status == "present") style["fillThemeColor"] = fillTheme.Value;
        if (patternTheme.Status == "present") style["fillPatternThemeColor"] = patternTheme.Value;

        var hasFormula = TryBool(() => (object?)range.HasFormula);
        object? value2 = null;
        var value2Read = false;
        try
        {
            value2 = (object?)range.Value2;
            if (value2 is DBNull) value2 = null;
            value2Read = true;
        }
        catch
        {
            value2Read = false;
        }

        string? formulaText = null;
        if (hasFormula == true)
            formulaText = TryString(() => Convert.ToString((object?)range.Formula, CultureInfo.InvariantCulture));

        var observedTint = style["fillTintAndShade"] as JsonValue;
        double? nativeTint = observedTint is not null && observedTint.TryGetValue<double>(out var tintVal) ? tintVal : null;
        var semanticDrift = SeededSemanticDrift(seeded, style, fillTheme);

        return new JsonObject
        {
            ["ref"] = address,
            ["index"] = index,
            ["unreadable"] = unreadable.Count > 0,
            ["unreadableFields"] = unreadable,
            ["requiredFields"] = ToJsonArray(required),
            ["hasFormula"] = hasFormula,
            ["value2Read"] = value2Read,
            ["value2ClrType"] = value2Read ? value2?.GetType().FullName : null,
            ["value"] = value2Read ? ClassifyTypedValue(value2) : null,
            ["formula"] = hasFormula == true ? formulaText : null,
            ["style"] = style,
            ["theme"] = new JsonObject
            {
                ["fillThemeColor"] = ThemeNode(fillTheme),
                ["fillPatternThemeColor"] = ThemeNode(patternTheme),
            },
            ["seeded"] = seeded,
            ["observedNativeTint"] = nativeTint,
            ["seededSemanticTint"] = seeded["tintAndShade"]?.DeepClone(),
            ["tintExactMatchToSeed"] = seeded["tintAndShade"] is JsonValue seedTint &&
                                       seedTint.TryGetValue<double>(out var seed) &&
                                       nativeTint is double native &&
                                       Math.Abs(native - seed) < 1e-9,
            ["semanticDrift"] = semanticDrift,
            ["comparedSeededTintAsExactIdentity"] = false,
        };
    }

    private static JsonObject SeededSemantics(int kind) => kind switch
    {
        0 => new JsonObject
        {
            ["kind"] = "no-fill",
            ["bold"] = true,
            ["fillPattern"] = XlPatternNone,
        },
        1 => new JsonObject
        {
            ["kind"] = "solid-rgb",
            ["bold"] = true,
            ["fillColor"] = RedOle,
            ["fillPattern"] = XlPatternSolid,
        },
        2 => new JsonObject
        {
            ["kind"] = "theme-tint",
            ["bold"] = false,
            ["fillPattern"] = XlPatternSolid,
            ["themeColor"] = XlThemeColorAccent1,
            ["tintAndShade"] = ThemeTint,
            ["tintNote"] = "seeded semantic 0.25; record native tint separately; do not treat quantized native as exact seed",
        },
        3 => new JsonObject
        {
            ["kind"] = "patterned",
            ["bold"] = true,
            ["fillColor"] = RedOle,
            ["fillPattern"] = XlPatternGray16,
            ["fillPatternColor"] = BlueOle,
        },
        _ => new JsonObject
        {
            ["kind"] = "dirty-solid-rgb",
            ["bold"] = true,
            ["fillColor"] = BlueOle,
            ["fillPattern"] = XlPatternSolid,
        },
    };

    private static bool SeededSemanticDrift(JsonObject seeded, JsonObject style, ThemeObservation theme)
    {
        if (seeded["bold"] is JsonValue b && b.TryGetValue<bool>(out var expectBold) &&
            style["bold"] is JsonValue ab && ab.TryGetValue<bool>(out var actualBold) &&
            expectBold != actualBold)
            return true;
        if (seeded["fillPattern"] is JsonValue p && p.TryGetValue<int>(out var expectPattern) &&
            style["fillPattern"] is JsonValue ap && ap.TryGetValue<int>(out var actualPattern) &&
            expectPattern != actualPattern)
            return true;
        if (seeded["fillColor"] is JsonValue c && c.TryGetValue<double>(out var expectColor) &&
            style["fillColor"] is JsonValue ac && ac.TryGetValue<double>(out var actualColor) &&
            Math.Abs(actualColor - expectColor) > 0.5d)
            return true;
        if (seeded["themeColor"] is JsonValue t && t.TryGetValue<int>(out var expectTheme) &&
            (theme.Status != "present" || theme.Value != expectTheme))
            return true;
        return false;
    }

    private static bool ObservationsMatch(JsonObject? before, JsonObject after)
    {
        if (before is null) return false;
        if (Json.GetBool(before, "unreadable") || Json.GetBool(after, "unreadable")) return false;
        var beforeStyle = Json.GetObj(before, "style");
        var afterStyle = Json.GetObj(after, "style");
        if (beforeStyle is null || afterStyle is null) return false;
        foreach (var field in RequiredNativeFields())
        {
            if (!StyleFieldEqual(beforeStyle[field], afterStyle[field], field))
                return false;
        }

        var beforeTheme = Json.GetObj(Json.GetObj(before, "theme"), "fillThemeColor");
        var afterTheme = Json.GetObj(Json.GetObj(after, "theme"), "fillThemeColor");
        var beforePatternTheme = Json.GetObj(Json.GetObj(before, "theme"), "fillPatternThemeColor");
        var afterPatternTheme = Json.GetObj(Json.GetObj(after, "theme"), "fillPatternThemeColor");
        return ThemeEqual(beforeTheme, afterTheme) && ThemeEqual(beforePatternTheme, afterPatternTheme);
    }

    private static bool StyleFieldEqual(JsonNode? left, JsonNode? right, string field)
    {
        if (left is null && right is null) return true;
        if (left is not JsonValue lv || right is not JsonValue rv) return false;
        if (field is "fillColor" or "fillPatternColor" or "fillTintAndShade" or "fillPatternTintAndShade")
        {
            if (!lv.TryGetValue<double>(out var ld) || !rv.TryGetValue<double>(out var rd)) return false;
            var tol = field.Contains("Tint", StringComparison.Ordinal) ? 1e-6 : 0.5d;
            return Math.Abs(ld - rd) <= tol;
        }

        return string.Equals(Json.Canonical(left), Json.Canonical(right), StringComparison.Ordinal);
    }

    private static bool ThemeEqual(JsonObject? left, JsonObject? right)
    {
        if (left is null || right is null) return false;
        return Json.GetString(left, "status") == Json.GetString(right, "status") &&
               Json.GetInt(left, "value") == Json.GetInt(right, "value");
    }

    private readonly record struct ThemeObservation(string Status, int? Value, int? HResult, string? Error);

    private static ThemeObservation ReadThemeColor(Func<object?> read, string property)
    {
        try
        {
            var raw = read();
            if (raw is null or DBNull)
                return new ThemeObservation("absent-valid", null, null, null);
            var value = Convert.ToInt32(raw, CultureInfo.InvariantCulture);
            if (value is >= ExcelStyleContract.XlThemeColorMin and <= ExcelStyleContract.XlThemeColorMax)
                return new ThemeObservation("present", value, null, null);
            return new ThemeObservation("absent-valid", value, null, "theme value outside 1..12");
        }
        catch (COMException ex) when (ex.HResult == ExcelThemeReadUnavailable)
        {
            return new ThemeObservation("absent-valid", null, ex.ErrorCode, "Excel 1004 ThemeColor unset (direct RGB / no-fill)");
        }
        catch (Exception ex)
        {
            var hr = ex is COMException com ? com.ErrorCode : ex.HResult;
            return new ThemeObservation("read-failed", null, hr, $"{property}: {ex.GetType().Name} {ex.Message}");
        }
    }

    private static JsonObject ThemeNode(ThemeObservation theme)
    {
        var node = new JsonObject
        {
            ["status"] = theme.Status,
            ["value"] = theme.Value,
        };
        if (theme.HResult is int hr) node["hresult"] = hr;
        if (theme.Error is not null) node["error"] = theme.Error;
        return node;
    }

    private static JsonObject NativePhase(string name, ScaleCellScan scan) => new()
    {
        ["name"] = name,
        ["ok"] = scan.Ok,
        ["checked"] = scan.Checked,
        ["expected"] = scan.Expected,
        ["chunks"] = scan.Chunks,
        ["chunkCells"] = NativeScanChunkCells,
        ["unreadable"] = scan.Unreadable,
        ["mismatches"] = scan.Mismatches,
        ["semanticDriftVsSeed"] = scan.SemanticDrift,
        ["wallMs"] = scan.Ms,
        ["cellCount"] = scan.Cells.Count,
        ["missingTreatedAsFalseOrZero"] = false,
        ["themeAbsenceCatchAllComErrors"] = false,
        ["comparedSeededTintAsExactIdentity"] = false,
        ["samples"] = scan.Samples.DeepClone(),
    };

    private static JsonObject ReadForegroundSnapshot()
    {
        try
        {
            var type = typeof(RotHelper).Assembly.GetType("DocBridge.Core.Services.Win32ForegroundWindowNative");
            var instance = type?.GetField("Instance", BindingFlags.Public | BindingFlags.Static)?.GetValue(null);
            var getter = type?.GetMethod("GetForegroundWindow", BindingFlags.Public | BindingFlags.Instance);
            if (instance is null || getter is null)
            {
                return new JsonObject
                {
                    ["measured"] = false,
                    ["reason"] = "Win32ForegroundWindowNative.GetForegroundWindow was not found",
                };
            }

            var hwnd = Convert.ToInt64(getter.Invoke(instance, null), CultureInfo.InvariantCulture);
            return new JsonObject
            {
                ["measured"] = true,
                ["hwnd"] = hwnd,
                ["source"] = "Win32ForegroundWindowNative.GetForegroundWindow",
            };
        }
        catch (Exception ex)
        {
            return new JsonObject
            {
                ["measured"] = false,
                ["reason"] = ex.Message,
            };
        }
    }

    private static JsonObject ReadIdentity(ExcelAdapter adapter, object app)
    {
        return adapter.RunOnAdapterThread(() =>
        {
            object? workbooks = null;
            object? workbook = null;
            try
            {
                workbooks = (object)((dynamic)app).Workbooks;
                workbook = FindOwnedWorkbook(workbooks);
                return new JsonObject
                {
                    ["pid"] = ProcessIdOf(app),
                    ["fullName"] = Convert.ToString((object?)((dynamic)workbook).FullName, CultureInfo.InvariantCulture),
                    ["saved"] = (bool)((dynamic)workbook).Saved,
                    ["workbookCount"] = Convert.ToInt32(((dynamic)workbooks).Count, CultureInfo.InvariantCulture),
                    ["hwnd"] = Convert.ToInt64(((dynamic)app).Hwnd, CultureInfo.InvariantCulture),
                };
            }
            finally
            {
                RotHelper.ReleaseComReference(workbook);
                RotHelper.ReleaseComReference(workbooks);
            }
        });
    }

    private static JsonObject ReadCellAudit(ExcelAdapter adapter, object app, string sheetName, string address)
    {
        return adapter.RunOnAdapterThread(() =>
        {
            var sheet = Sheet(app, sheetName);
            object? range = null;
            object? font = null;
            object? interior = null;
            try
            {
                range = (object)((dynamic)sheet).Range(address);
                font = (object)((dynamic)range).Font;
                interior = (object)((dynamic)range).Interior;
                var bold = TryBool(() => ((dynamic)font).Bold);
                var fill = TryDouble(() => ((dynamic)interior).Color);
                var pattern = TryInt(() => ((dynamic)interior).Pattern);
                object? value2 = null;
                var value2Read = false;
                try
                {
                    value2 = (object?)((dynamic)range).Value2;
                    if (value2 is DBNull) value2 = null;
                    value2Read = true;
                }
                catch
                {
                    value2Read = false;
                }

                var hasFormula = TryBool(() => (object?)((dynamic)range).HasFormula);
                string? formulaText = null;
                if (hasFormula == true)
                    formulaText = TryString(() => Convert.ToString((object?)((dynamic)range).Formula, CultureInfo.InvariantCulture));
                var display = value2Read
                    ? Convert.ToString(value2, CultureInfo.InvariantCulture)
                    : null;
                return new JsonObject
                {
                    ["sheet"] = sheetName,
                    ["address"] = address,
                    ["readable"] = bold is not null && pattern is not null,
                    ["unreadable"] = bold is null || pattern is null,
                    ["fontBold"] = bold,
                    ["fillColor"] = fill,
                    ["fillPattern"] = pattern,
                    ["value"] = display,
                    ["hasFormula"] = hasFormula,
                    ["value2Read"] = value2Read,
                    ["typedValue"] = value2Read ? ClassifyTypedValue(value2) : null,
                    ["formula"] = hasFormula == true ? formulaText : null,
                    ["coercedMissingToFalseOrZero"] = false,
                };
            }
            finally
            {
                RotHelper.ReleaseComReference(interior);
                RotHelper.ReleaseComReference(font);
                RotHelper.ReleaseComReference(range);
                RotHelper.ReleaseComReference(sheet);
            }
        });
    }

    private static bool AuditReadable(JsonObject audit) =>
        Json.GetBool(audit, "readable") && !Json.GetBool(audit, "unreadable");

    private static object FindOwnedWorkbook(object workbooks)
    {
        var count = Convert.ToInt32(((dynamic)workbooks).Count, CultureInfo.InvariantCulture);
        for (var i = 1; i <= count; i++)
        {
            object? workbook = (object)((dynamic)workbooks).Item(i);
            var full = Convert.ToString((object?)((dynamic)workbook).FullName, CultureInfo.InvariantCulture) ?? "";
            if (string.Equals(full, ProbePath, StringComparison.OrdinalIgnoreCase))
                return workbook;
            RotHelper.ReleaseComReference(workbook);
        }

        throw new InvalidOperationException("owned probe workbook not found");
    }

    private static object Sheet(object app, string name)
    {
        object? workbooks = null;
        object? workbook = null;
        try
        {
            workbooks = (object)((dynamic)app).Workbooks;
            workbook = FindOwnedWorkbook(workbooks);
            return (object)((dynamic)workbook).Worksheets.Item(name);
        }
        finally
        {
            RotHelper.ReleaseComReference(workbook);
            RotHelper.ReleaseComReference(workbooks);
        }
    }

    private static void ActivateOwnedWorkbook(ExcelAdapter adapter, object app)
    {
        adapter.RunOnAdapterThread<object?>(() =>
        {
            object? workbooks = null;
            object? workbook = null;
            try
            {
                workbooks = (object)((dynamic)app).Workbooks;
                workbook = FindOwnedWorkbook(workbooks);
                ((dynamic)workbook).Activate();
                return null;
            }
            finally
            {
                RotHelper.ReleaseComReference(workbook);
                RotHelper.ReleaseComReference(workbooks);
            }
        });
    }

    private static void ActivateProbe()
    {
        if (SessionAdapter is null || SessionApp is null) return;
        ActivateOwnedWorkbook(SessionAdapter, SessionApp);
    }

    private static List<JsonObject> FormatOpList(string range)
    {
        var style = new JsonObject { ["bold"] = true };
        if (MixedFill) style["fillColor"] = RedOle;
        return
        [
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = new JsonObject { ["sheet"] = TargetSheet },
                ["range"] = range,
                ["style"] = style,
            },
        ];
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

        var pidMatches = adapter is not null && app is not null && PidMatchesOnSta(adapter, app, OwnedPid);
        if (pidMatches && adapter is not null && app is not null)
        {
            object staApp = app;
            adapter.RunOnAdapterThread<object?>(() =>
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
                finally { RotHelper.ReleaseComReference(workbooks); }

                return null;
            });
        }

        JsonObject? disconnect = null;
        if (pidMatches && adapter is not null)
            disconnect = adapter.Disconnect();

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
            if (node is JsonObject o && Json.GetInt(o, "pid") is int value && value > 0)
                set.Add(value);
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
        finally { RotHelper.ReleaseComReference(range); }
    }

    private static void SetCellFormula(object sheet, string address, string formula)
    {
        object? range = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            ((dynamic)range).Formula = formula;
        }
        finally { RotHelper.ReleaseComReference(range); }
    }

    private static void SetCellBold(object sheet, string address, bool bold)
    {
        object? range = null;
        object? font = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            font = (object)((dynamic)range).Font;
            ((dynamic)font).Bold = bold;
        }
        finally
        {
            RotHelper.ReleaseComReference(font);
            RotHelper.ReleaseComReference(range);
        }
    }

    private static void SetCellFillRaw(object sheet, string address, double ole)
    {
        object? range = null;
        object? interior = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            interior = (object)((dynamic)range).Interior;
            ((dynamic)interior).Pattern = XlPatternSolid;
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

    private static int ScaleRows(int scale) => scale / 10;

    private static string ScaleRange(int scale) => $"A20:J{19 + ScaleRows(scale)}";

    private static string ScaleAddress(int index)
    {
        var row = 20 + (index / 10);
        var col = index % 10;
        return $"{(char)('A' + col)}{row}";
    }

    private static bool IdentityReadable(JsonObject identity) =>
        !string.IsNullOrWhiteSpace(Json.GetString(identity, "fullName")) &&
        Json.GetInt(identity, "pid") > 0 &&
        Json.GetInt(identity, "workbookCount") >= 1;

    private static bool IdentityUnchanged(JsonObject before, JsonObject after) =>
        string.Equals(Json.GetString(before, "fullName"), Json.GetString(after, "fullName"), StringComparison.OrdinalIgnoreCase) &&
        Json.GetBool(before, "saved") == Json.GetBool(after, "saved") &&
        Json.GetInt(before, "workbookCount") == Json.GetInt(after, "workbookCount") &&
        Json.GetInt(before, "pid") == Json.GetInt(after, "pid");

    private static bool RepeatsOk(JsonArray repeats) =>
        repeats.All(n => n is JsonObject o && Json.GetBool(o, "ok"));

    private static JsonObject Fail(string method, long ms, Exception ex)
    {
        var node = new JsonObject
        {
            ["ok"] = false,
            ["wallMs"] = ms,
            ["method"] = method,
            ["error"] = ex.Message,
            ["errorType"] = ex.GetType().FullName,
        };
        if (ex is COMException com) node["hresult"] = com.ErrorCode;
        return node;
    }

    private static bool? TryBool(Func<object?> read)
    {
        try
        {
            var value = read();
            return value is null or DBNull ? null : Convert.ToBoolean(value, CultureInfo.InvariantCulture);
        }
        catch { return null; }
    }

    private static double? TryDouble(Func<object?> read)
    {
        try
        {
            var value = read();
            return value is null or DBNull ? null : Convert.ToDouble(value, CultureInfo.InvariantCulture);
        }
        catch { return null; }
    }

    private static int? TryInt(Func<object?> read)
    {
        try
        {
            var value = read();
            return value is null or DBNull ? null : Convert.ToInt32(value, CultureInfo.InvariantCulture);
        }
        catch { return null; }
    }

    private static string? TryString(Func<string?> read)
    {
        try { return read(); }
        catch { return null; }
    }

    private static string[] WrittenPropertyNames() =>
        MixedFill ? ["bold", "fillColor"] : ["bold"];

    private static string[] CoupledFillFieldNames() =>
    [
        "fillColor",
        "fillColorIndex",
        "fillTintAndShade",
        "fillPattern",
        "fillPatternColor",
        "fillPatternColorIndex",
        "fillPatternTintAndShade",
    ];

    private static string[] RequiredNativeFields() =>
        MixedFill ? ["bold", .. CoupledFillFieldNames()] : ["bold"];

    private static JsonArray ToJsonArray(IEnumerable<string> items)
    {
        var array = new JsonArray();
        foreach (var item in items) array.Add(item);
        return array;
    }

    private static void Progress(string phase, string detail) =>
        Console.WriteLine($"PROGRESS {DateTimeOffset.Now:O} {phase} {detail}");

    private static bool CleanupSucceeded(JsonObject report)
    {
        if (report["cleanupError"] is not null) return false;
        var cleanup = Json.GetObj(report, "cleanup");
        if (cleanup is null) return false;
        if (Json.GetBool(cleanup, "killedOffice")) return false;
        return true;
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

    private static JsonObject CompactResult(JsonObject result)
    {
        var compact = new JsonObject
        {
            ["ok"] = Json.GetBool(result, "ok"),
            ["errors"] = JoinErrors(result),
            ["snapshotId"] = Json.GetString(result, "snapshotId"),
        };
        if (result["rollback"] is JsonNode rollback) compact["rollback"] = rollback.DeepClone();
        return compact;
    }

    private readonly record struct TimedCall(JsonObject Result, long Ms);

    private static TimedCall Timed(Func<JsonObject> work)
    {
        var sw = Stopwatch.StartNew();
        var result = work();
        sw.Stop();
        return new TimedCall(result, sw.ElapsedMilliseconds);
    }

    private static void WriteExpectedCatalog(string path, int scale, string baselinePath, string baselineCopyPath)
    {
        var node = new JsonObject
        {
            ["role"] = "seeded-semantic-catalog-separate-from-native-before-observations",
            ["case"] = ProbeCase,
            ["writtenProperties"] = ToJsonArray(WrittenPropertyNames()),
            ["sheet"] = TargetSheet,
            ["keepSheet"] = KeepSheet,
            ["range"] = ScaleRange(scale),
            ["cells"] = scale,
            ["savedScalars"] = new JsonObject
            {
                ["A1.value"] = SeedValue,
                ["A2.formula"] = SeedFormula,
                ["A3.value"] = "uniform-bold-restore",
                ["Keep!A1.value"] = Marker,
            },
            ["unsavedScalars"] = new JsonObject
            {
                ["A1.value"] = DirtyValue,
                ["A2.formula"] = DirtyFormula,
                ["A3.fontBold"] = false,
                ["firstScaleCell.fill"] = MixedFill ? JsonValue.Create(BlueOle) : JsonValue.Create("unchanged-uniform-bold"),
            },
            ["styleCycle"] = MixedFill
                ? new JsonArray(
                    "0 bold + no-fill",
                    "1 bold + solid RGB red OLE 255",
                    "2 not-bold + theme Accent1 + seeded tint 0.25 (record native tint separately)",
                    "3 bold + patterned Gray16 + PatternColor blue",
                    "dirty first cell: bold + solid RGB blue")
                : new JsonArray("uniform bold + no-fill on the whole target range"),
            ["unsupportedSuccessForbidden"] = "unknown style key fillTint (not a native RGB+tint snapshot proof)",
            ["baselinePath"] = baselinePath,
            ["baselineCopyPath"] = baselineCopyPath,
            ["seededCellCount"] = scale,
            ["dirtyFirstCellOverridesCycle"] = MixedFill,
            ["doNotCompareSeededTintExactToNative"] = true,
        };
        File.WriteAllText(path, node.ToJsonString(Json.Pretty));
    }

    private static void WriteNativeBeforeArtifact(
        string path, ScaleCellScan scan, JsonObject keep, JsonObject identity, string sourceWorkbook)
    {
        var node = new JsonObject
        {
            ["role"] = "per-cell native-before observations for independent audit",
            ["recordedBeforeSaveCopyAs"] = true,
            ["case"] = ProbeCase,
            ["writtenProperties"] = ToJsonArray(WrittenPropertyNames()),
            ["requiredFields"] = ToJsonArray(RequiredNativeFields()),
            ["sourceWorkbookFullName"] = sourceWorkbook,
            ["liveIdentity"] = identity.DeepClone(),
            ["scan"] = NativePhase("native-before-pre-savecopyas", scan),
            ["cells"] = scan.Cells.DeepClone(),
            ["untouched"] = new JsonArray(keep.DeepClone()),
            ["seededSemanticsAreSeparate"] = true,
            ["comparedSeededTintAsExactIdentity"] = false,
        };
        File.WriteAllText(path, node.ToJsonString(Json.Pretty));
    }

    private static void WriteAuditExpectation(
        string path, string? workbookPath, string sourceWorkbook, JsonObject identity,
        ScaleCellScan nativeBefore, JsonObject keep)
    {
        var targetCells = new JsonArray();
        foreach (var node in nativeBefore.Cells)
        {
            if (node is not JsonObject cell) continue;
            var style = Json.GetObj(cell, "style")?.DeepClone().AsObject() ?? new JsonObject();
            var row = new JsonObject
            {
                ["ref"] = Json.GetString(cell, "ref"),
                ["style"] = style,
            };
            if (ToAuditValue(cell) is { } auditValue)
                row["value"] = auditValue;
            if (cell["hasFormula"] is JsonValue hasNode && hasNode.TryGetValue<bool>(out var hasFormula))
                row["formula"] = hasFormula ? cell["formula"]?.DeepClone() : JsonValue.Create((string?)null);
            targetCells.Add(row);
        }

        var keepStyle = new JsonObject();
        if (keep["fontBold"] is JsonNode bold) keepStyle["bold"] = bold.DeepClone();
        if (keep["fillColor"] is JsonNode fill) keepStyle["fillColor"] = fill.DeepClone();
        if (keep["fillPattern"] is JsonNode pattern) keepStyle["fillPattern"] = pattern.DeepClone();

        var expectation = new JsonObject
        {
            ["schemaVersion"] = "docbridge-excel-audit-expectation/1",
            ["note"] = "Audit value.kind is the COM Value2 CLR type (string|number|bool|blank); formula is null when HasFormula is false. Seeded tint 0.25 is not asserted. Unsupported COM-to-OOXML mappings may be reported honestly.",
            ["workbook"] = new JsonObject
            {
                ["path"] = workbookPath,
                ["sourceWorkbookFullName"] = sourceWorkbook,
                ["liveSavedFlagBefore"] = Json.GetBool(identity, "saved"),
                ["liveSavedFlagAfter"] = Json.GetBool(identity, "saved"),
            },
            ["targets"] = new JsonArray(new JsonObject
            {
                ["sheet"] = TargetSheet,
                ["range"] = ScaleRange(nativeBefore.Expected),
                ["scopedProperties"] = ToJsonArray(WrittenPropertyNames()),
                ["cells"] = targetCells,
            }),
            ["untouched"] = new JsonArray(new JsonObject
            {
                ["sheet"] = KeepSheet,
                ["range"] = "A1",
                ["cells"] = new JsonArray(BuildUntouchedKeepCell(keep, keepStyle)),
            }),
        };
        File.WriteAllText(path, expectation.ToJsonString(Json.Pretty));
    }

    private static JsonObject ClassifyTypedValue(object? raw) => raw switch
    {
        null => new JsonObject { ["kind"] = "blank", ["raw"] = null },
        bool flag => new JsonObject { ["kind"] = "bool", ["raw"] = flag },
        byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal
            => new JsonObject
            {
                ["kind"] = "number",
                ["raw"] = Convert.ToDouble(raw, CultureInfo.InvariantCulture),
            },
        DateTime date => new JsonObject
        {
            ["kind"] = "iso-date",
            ["raw"] = date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
        },
        string text => new JsonObject { ["kind"] = "string", ["raw"] = text },
        _ => new JsonObject
        {
            ["kind"] = "string",
            ["raw"] = Convert.ToString(raw, CultureInfo.InvariantCulture),
        },
    };

    private static JsonObject? ToAuditValue(JsonObject cell)
    {
        if (cell["value2Read"] is JsonValue read && read.TryGetValue<bool>(out var ok) && !ok)
            return null;
        if (Json.GetObj(cell, "value") is { } typed && IsAuditValueKind(Json.GetString(typed, "kind")))
            return typed.DeepClone().AsObject();
        if (Json.GetObj(cell, "typedValue") is { } named && IsAuditValueKind(Json.GetString(named, "kind")))
            return named.DeepClone().AsObject();
        return null;
    }

    private static bool IsAuditValueKind(string? kind) =>
        kind is "string" or "number" or "bool" or "blank" or "iso-date" or "error";

    private static JsonObject BuildUntouchedKeepCell(JsonObject keep, JsonObject keepStyle)
    {
        var row = new JsonObject
        {
            ["ref"] = "A1",
            ["style"] = keepStyle,
        };
        if (ToAuditValue(keep) is { } auditValue)
            row["value"] = auditValue;
        if (keep["hasFormula"] is JsonValue hasNode && hasNode.TryGetValue<bool>(out var hasFormula))
            row["formula"] = hasFormula ? keep["formula"]?.DeepClone() : JsonValue.Create((string?)null);
        return row;
    }

    private static void WriteResult(string path, JsonObject report) =>
        File.WriteAllText(path, report.ToJsonString(Json.Pretty));

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
}
