using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

namespace DocBridge.Development.ExcelFormatValidationProbe;

internal static partial class Program
{
    private enum ProbeExecutionMode { Legacy, Execute }

    private static string ModeArg = "legacy";
    private static int Repeat = 1;
    private static HashSet<string> CaseFilter = new(StringComparer.OrdinalIgnoreCase);
    private static ProbeExecutionMode CurrentMode = ProbeExecutionMode.Legacy;

    private static HashSet<string> ParseCaseFilter(string raw)
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(raw)) return set;
        foreach (var part in raw.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries))
            set.Add(part);
        return set;
    }

    private static bool CaseWanted(params string[] names)
    {
        if (CaseFilter.Count == 0) return true;
        return names.Any(name => CaseFilter.Contains(name));
    }

    private static IEnumerable<ProbeExecutionMode> SelectedModes()
    {
        if (ModeArg == "execute") yield return ProbeExecutionMode.Execute;
        else if (ModeArg == "both")
        {
            yield return ProbeExecutionMode.Legacy;
            yield return ProbeExecutionMode.Execute;
        }
        else yield return ProbeExecutionMode.Legacy;
    }

    private static string ModeName(ProbeExecutionMode mode) =>
        mode == ProbeExecutionMode.Execute ? "execute" : "legacy";

    private static string NewRequestId() => Guid.NewGuid().ToString("D");

    private static int ScaleRows(int scale) => scale switch
    {
        7000 => 700,
        1000 => 100,
        100 => 10,
        _ => 0,
    };

    private static int ScaleLastRow(int scale) => 19 + ScaleRows(scale);

    private static string ScaleRange(int scale) => scale switch
    {
        7000 => "A20:J719",
        1000 => "A20:J119",
        100 => "A20:J29",
        _ => "A20:A20",
    };

    private static void AddRegressionCases(
        JsonArray cases, DocBridgeHost host, ExcelAdapter adapter, object app, string bystanderPath, bool bloat)
    {
        var previous = CurrentMode;
        CurrentMode = ProbeExecutionMode.Legacy;
        try
        {
            if (CaseWanted("canonical-fillColor-and-aliases", "canonical"))
                cases.Add(RunCanonicalAndAliases(host));
            if (CaseWanted("nested-json-reorder", "json-reorder"))
                cases.Add(RunJsonReorder(host));
            if (CaseWanted("protected-sheet-com-failure", "protected"))
                cases.Add(RunProtectedSheet(host, adapter, app));
            if (CaseWanted("oversized-fontsize-com-failure", "fault-after-success", "fault"))
                cases.Add(RunOversizedFontComFailure(host));
            if (CaseWanted("style-changed-after-dryrun", "fingerprint"))
                cases.Add(RunStyleChangedAfterDryRun(host, adapter, app));
            if (CaseWanted("mixed-theme-nofill-rollback", "mixed-theme"))
                cases.Add(RunMixedThemeNoFillRollback(host, adapter, app));
            if (CaseWanted("unsupported-rgb-tint-color-change-denied", "unsupported-rgb-tint-denied", "color-change"))
                cases.Add(RunUnsupportedRgbTintColorChangeDenied(host, adapter, app));
            if (CaseWanted("other-active-workbook-preservation", "bystander"))
                cases.Add(RunOtherWorkbookPreservation(host, adapter, app, bystanderPath));
            if (bloat && CaseWanted("usedrange-bloat-blocker", "bloat"))
                cases.Add(RunUsedRangeBloatBlocker(host, adapter, app));
        }
        finally { CurrentMode = previous; }
    }

    private static void AddSpeedCases(
        JsonArray cases, DocBridgeHost host, ExcelAdapter adapter, object app, int scale, int repetition)
    {
        if (CaseWanted("uniform-bold", "scale") && scale > 1)
            cases.Add(WithRepetition(RunUniformBold(host, adapter, app, scale), repetition));
        if (CaseWanted("mixed-bold-colors", "mixed") && scale > 1)
            cases.Add(WithRepetition(RunMixedBoldAndColors(host, adapter, app, scale), repetition));
        if (CaseWanted("restore") && scale > 1)
            cases.Add(WithRepetition(RunRestoreTiming(host, adapter, app, scale), repetition));
        if (CaseWanted("tint-preservation", "tint"))
            cases.Add(WithRepetition(RunTintedRgbBoldPreservation(host, adapter, app), repetition));
        if (CurrentMode == ProbeExecutionMode.Execute && CaseWanted("execute-replay", "replay"))
            cases.Add(WithRepetition(RunExecuteReplay(host, adapter, app), repetition));
        if (CurrentMode == ProbeExecutionMode.Execute && CaseWanted("execute-conflict", "conflict"))
            cases.Add(WithRepetition(RunExecuteConflict(host, adapter, app), repetition));
        if (CurrentMode == ProbeExecutionMode.Execute &&
            CaseWanted("fault-after-success", "fault", "execute-fault"))
            cases.Add(WithRepetition(RunExecuteFaultAfterSuccess(host, adapter, app), repetition));
    }

    private static JsonObject WithRepetition(JsonObject result, int repetition)
    {
        result["repetition"] = repetition;
        result["mode"] = ModeName(CurrentMode);
        if (repetition > 1)
            result["name"] = $"{Json.GetString(result, "name")}#{repetition}";
        return result;
    }

    private static JsonObject RunUniformBold(DocBridgeHost host, ExcelAdapter adapter, object app, int scale)
    {
        var sw = Stopwatch.StartNew();
        ActivateProbeWorkbook();
        var range = ScaleRange(scale);
        var untouched = ReadComStyle(adapter, app, ProbeSheet, "C4");
        var before = ScanScaleCells(adapter, app, scale);
        var flow = FormatFlow(host, ProbeSheet, range, new JsonObject { ["bold"] = true }, restore: false);
        var afterApply = ScanScaleCells(adapter, app, scale, requireBold: true, expectFill: before.Fill, expectPattern: before.Pattern);
        var executeUnsupported = Json.GetBool(flow, "executeUnsupported");
        var applied = afterApply.Ok;
        var restore = AttachRestore(host, flow, executeUnsupported);
        var afterRestore = ScanScaleCells(adapter, app, scale, expectBold: before.Bold, expectFill: before.Fill, expectPattern: before.Pattern);
        var afterUntouched = ReadComStyle(adapter, app, ProbeSheet, "C4");
        var restored = afterRestore.Ok && afterRestore.Digest == before.Digest;
        var bystander = StylesEqual(untouched, afterUntouched);
        var restoreOk = restore is not null && Json.GetBool(restore, "ok");
        var nativeVerifyMs = before.Ms + afterApply.Ms + afterRestore.Ms;
        var hostMs = HostPhaseMs(flow);
        var ok = !executeUnsupported && Json.GetBool(flow, "ok") && applied && restoreOk && restored && bystander;
        sw.Stop();
        var result = Case($"uniform-bold-{scale}", ok, hostMs,
            new JsonArray(flow,
                NativeScanPhase("native-before", before),
                NativeScanPhase("native-after-apply-before-restore", afterApply),
                NativeScanPhase("native-after-restore", afterRestore),
                new JsonObject { ["name"] = "untouched-C4-before", ["cell"] = untouched },
                new JsonObject { ["name"] = "untouched-C4-after", ["cell"] = afterUntouched }),
            new JsonArray(
                $"mode={ModeName(CurrentMode)} range={range} cells={scale} scanned={afterApply.Checked}",
                $"executeUnsupported={executeUnsupported} (unsupported is skip/failure, not a candidate pass)",
                $"requestedAppliedBeforeRestore={applied} nativeRestored={restored} untouchedC4Preserved={bystander}",
                "every scale cell is scanned for requested Bold/fill after apply and before restore; nativeVerifyMs is not host apply/restore time"),
            executeUnsupported);
        result["nativeVerifyMs"] = nativeVerifyMs;
        result["totalMs"] = sw.ElapsedMilliseconds;
        return result;
    }

    private static JsonObject RunMixedBoldAndColors(DocBridgeHost host, ExcelAdapter adapter, object app, int scale)
    {
        var sw = Stopwatch.StartNew();
        ActivateProbeWorkbook();
        var range = ScaleRange(scale);
        SeedMixedScaleStyles(adapter, app, scale);
        var before = ScanScaleCells(adapter, app, scale);
        var flow = FormatFlow(host, ProbeSheet, range,
            new JsonObject { ["bold"] = true, ["fillColor"] = CanonicalFillOle }, restore: false);
        var afterApply = ScanScaleCells(adapter, app, scale, requireBold: true, requireFill: CanonicalFillOle);
        var executeUnsupported = Json.GetBool(flow, "executeUnsupported");
        var applied = afterApply.Ok && afterApply.Checked == scale;
        var restore = AttachRestore(host, flow, executeUnsupported);
        var afterRestore = ScanScaleCells(adapter, app, scale, expectBold: before.Bold, expectFill: before.Fill, expectPattern: before.Pattern);
        var restored = afterRestore.Ok && afterRestore.Digest == before.Digest;
        var restoreOk = restore is not null && Json.GetBool(restore, "ok");
        var nativeVerifyMs = before.Ms + afterApply.Ms + afterRestore.Ms;
        var hostMs = HostPhaseMs(flow);
        var ok = !executeUnsupported && Json.GetBool(flow, "ok") && applied && restoreOk && restored;
        sw.Stop();
        var result = Case($"mixed-bold-colors-{scale}", ok, hostMs,
            new JsonArray(flow,
                NativeScanPhase("native-before", before),
                NativeScanPhase("native-after-apply-before-restore", afterApply),
                NativeScanPhase("native-after-restore", afterRestore)),
            new JsonArray(
                $"mode={ModeName(CurrentMode)} mixed groups over {range} ({scale} cells, not A20:D20)",
                $"executeUnsupported={executeUnsupported} (unsupported is skip/failure, not a candidate pass)",
                $"requestedAppliedBeforeRestore={applied} scanned={afterApply.Checked} nativeRestored={restored}"),
            executeUnsupported);
        result["nativeVerifyMs"] = nativeVerifyMs;
        result["totalMs"] = sw.ElapsedMilliseconds;
        result["editedRange"] = range;
        result["editedCells"] = scale;
        return result;
    }

    private static JsonObject RunRestoreTiming(DocBridgeHost host, ExcelAdapter adapter, object app, int scale)
    {
        var sw = Stopwatch.StartNew();
        ActivateProbeWorkbook();
        var range = ScaleRange(scale);
        var before = ScanScaleCells(adapter, app, scale);
        var flow = FormatFlow(host, ProbeSheet, range, new JsonObject { ["bold"] = true }, restore: false);
        var afterApply = ScanScaleCells(adapter, app, scale, requireBold: true, expectFill: before.Fill, expectPattern: before.Pattern);
        var executeUnsupported = Json.GetBool(flow, "executeUnsupported");
        var applied = afterApply.Ok;
        var restore = AttachRestore(host, flow, executeUnsupported);
        var afterRestore = ScanScaleCells(adapter, app, scale, expectBold: before.Bold, expectFill: before.Fill, expectPattern: before.Pattern);
        var restored = afterRestore.Ok && afterRestore.Digest == before.Digest;
        var restoreOk = restore is not null && Json.GetBool(restore, "ok");
        var nativeVerifyMs = before.Ms + afterApply.Ms + afterRestore.Ms;
        var hostMs = HostPhaseMs(flow);
        var ok = !executeUnsupported && Json.GetBool(flow, "ok") && applied && restoreOk && restored;
        sw.Stop();
        var result = Case($"restore-{scale}", ok, hostMs,
            new JsonArray(flow,
                NativeScanPhase("native-before", before),
                NativeScanPhase("native-after-apply-before-restore", afterApply),
                NativeScanPhase("native-after-restore", afterRestore)),
            new JsonArray(
                $"mode={ModeName(CurrentMode)} restore wallMs={restore?["wallMs"]?.GetValue<long>()} snapshotBytes={restore?["snapshotBytes"]?.GetValue<long>()}",
                $"executeUnsupported={executeUnsupported} (unsupported is skip/failure, not a candidate pass)",
                $"requestedAppliedBeforeRestore={applied} scanned={afterApply.Checked} nativeRestored={restored}"),
            executeUnsupported);
        result["nativeVerifyMs"] = nativeVerifyMs;
        result["totalMs"] = sw.ElapsedMilliseconds;
        return result;
    }

    private static JsonObject RunTintedRgbBoldPreservation(DocBridgeHost host, ExcelAdapter adapter, object app)
    {
        var sw = Stopwatch.StartNew();
        ActivateProbeWorkbook();
        SeedUnsupportedRgbTint(adapter, app);
        SetCellBold(adapter, app, ProbeSheet, "C4", bold: false);
        var before = ReadComStyle(adapter, app, ProbeSheet, "C4");
        var flow = FormatFlow(host, ProbeSheet, "C4", new JsonObject { ["bold"] = true }, restore: false);
        var after = ReadComStyle(adapter, app, ProbeSheet, "C4");
        var errors = JoinErrors(Json.GetObj(flow, "apply") ?? flow);
        if (Json.GetObj(flow, "dryrun") is { } dry) errors = JoinErrors(dry) + " | " + errors;
        var executeUnsupported = Json.GetBool(flow, "executeUnsupported");
        var colorPreserved = ColorFieldsEqual(before, after);
        var boldBefore = Json.GetBool(before, "fontBold");
        var boldApplied = Json.GetBool(after, "fontBold");
        var boldTransition = !boldBefore && boldApplied;
        var denied = !Json.GetBool(flow, "ok") &&
                     errors.Contains("EXCEL_FORMAT_UNSUPPORTED_RGB_TINT", StringComparison.Ordinal);
        var unchanged = StylesEqual(before, after);
        var candidateSuccess = Json.GetBool(flow, "ok") && boldTransition && colorPreserved;
        var baselineDenied = denied && unchanged;
        var ok = !executeUnsupported && (candidateSuccess || baselineDenied);
        sw.Stop();
        return Case("tint-preservation", ok, sw.ElapsedMilliseconds,
            new JsonArray(flow,
                new JsonObject { ["name"] = "native-before", ["cell"] = before.DeepClone() },
                new JsonObject { ["name"] = "native-after", ["cell"] = after.DeepClone() }),
            new JsonArray(
                $"mode={ModeName(CurrentMode)}",
                $"outcome={(candidateSuccess ? "bold-applied-tint-preserved" : baselineDenied ? "legacy-denied-unchanged" : executeUnsupported ? "execute-unsupported-not-a-pass" : "unexpected")}",
                $"boldBefore={boldBefore} boldApplied={boldApplied} boldTransition={boldTransition} colorPreserved={colorPreserved} denied={denied} nativeUnchanged={unchanged}",
                $"errors={errors}"),
            executeUnsupported);
    }

    private static JsonObject RunExecuteReplay(DocBridgeHost host, ExcelAdapter adapter, object app)
    {
        var sw = Stopwatch.StartNew();
        ActivateProbeWorkbook();
        var requestId = NewRequestId();
        var ops = new JsonArray(FormatOp(ProbeSheet, "A5", new JsonObject { ["bold"] = true }));
        var before = ReadComStyle(adapter, app, ProbeSheet, "A5");
        var first = Timed(() => host.ApplyOps("excel", MakeExecuteBatch(ops, requestId)));
        var afterFirst = ReadComStyle(adapter, app, ProbeSheet, "A5");
        var firstMutated = Json.GetBool(afterFirst, "fontBold") && !Json.GetBool(before, "fontBold");
        SetCellBold(adapter, app, ProbeSheet, "A5", bold: false);
        var challenged = ReadComStyle(adapter, app, ProbeSheet, "A5");
        var challengeCleared = !Json.GetBool(challenged, "fontBold");
        var replay = Timed(() => host.ApplyOps("excel", MakeExecuteBatch(CloneOps(ops), requestId)));
        var afterReplay = ReadComStyle(adapter, app, ProbeSheet, "A5");
        var unsupported = IsExecuteUnsupported(first.Result, null);
        var idempotentReplay = Json.GetBool(replay.Result, "idempotentReplay");
        var challengeHeld = !Json.GetBool(afterReplay, "fontBold") && StylesEqual(challenged, afterReplay);
        var ok = !unsupported &&
                 Json.GetBool(first.Result, "ok") &&
                 Json.GetBool(replay.Result, "ok") &&
                 firstMutated &&
                 challengeCleared &&
                 idempotentReplay &&
                 challengeHeld;
        sw.Stop();
        return Case("execute-replay", ok, sw.ElapsedMilliseconds,
            new JsonArray(
                Phase("execute-first", first.Result, first.Ms),
                Phase("execute-replay", replay.Result, replay.Ms),
                new JsonObject { ["name"] = "native-before", ["cell"] = before },
                new JsonObject { ["name"] = "native-after-first", ["cell"] = afterFirst },
                new JsonObject { ["name"] = "native-after-challenge-unbold", ["cell"] = challenged },
                new JsonObject { ["name"] = "native-after-replay", ["cell"] = afterReplay }),
            new JsonArray(
                $"requestId={requestId}",
                $"executeUnsupported={unsupported} (unsupported is not a candidate pass)",
                $"firstMutated={firstMutated} challengeCleared={challengeCleared} idempotentReplay={idempotentReplay} challengeHeld={challengeHeld}",
                "owned A5 is cleared after the first edit; replay must set idempotentReplay and must not re-apply bold"),
            unsupported);
    }

    private static JsonObject RunExecuteConflict(DocBridgeHost host, ExcelAdapter adapter, object app)
    {
        var sw = Stopwatch.StartNew();
        ActivateProbeWorkbook();
        var requestId = NewRequestId();
        var firstOps = new JsonArray(FormatOp(ProbeSheet, "A6", new JsonObject { ["bold"] = true }));
        var conflictOps = new JsonArray(FormatOp(ProbeSheet, "A6", new JsonObject { ["italic"] = true }));
        var before = ReadComStyle(adapter, app, ProbeSheet, "A6");
        var first = Timed(() => host.ApplyOps("excel", MakeExecuteBatch(firstOps, requestId)));
        var afterFirst = ReadComStyle(adapter, app, ProbeSheet, "A6");
        var conflict = Timed(() => host.ApplyOps("excel", MakeExecuteBatch(conflictOps, requestId)));
        var afterConflict = ReadComStyle(adapter, app, ProbeSheet, "A6");
        var unsupported = IsExecuteUnsupported(first.Result, null);
        var rejected = !Json.GetBool(conflict.Result, "ok");
        var unchangedAfterReject = StylesEqual(afterFirst, afterConflict);
        var firstMutated = Json.GetBool(afterFirst, "fontBold") && !Json.GetBool(before, "fontBold");
        var ok = !unsupported &&
                 Json.GetBool(first.Result, "ok") &&
                 firstMutated &&
                 rejected &&
                 unchangedAfterReject;
        sw.Stop();
        return Case("execute-conflict", ok, sw.ElapsedMilliseconds,
            new JsonArray(
                Phase("execute-first", first.Result, first.Ms),
                Phase("execute-conflict", conflict.Result, conflict.Ms),
                new JsonObject { ["name"] = "native-before", ["cell"] = before },
                new JsonObject { ["name"] = "native-after-first", ["cell"] = afterFirst },
                new JsonObject { ["name"] = "native-after-conflict", ["cell"] = afterConflict }),
            new JsonArray(
                $"requestId={requestId}",
                $"executeUnsupported={unsupported} (unsupported is not a candidate pass)",
                $"firstMutated={firstMutated} conflictRejected={rejected} nativeUnchangedAfterReject={unchangedAfterReject}",
                "same UUID + different payload must be rejected without a second mutation"),
            unsupported);
    }

    private static JsonObject RunExecuteFaultAfterSuccess(DocBridgeHost host, ExcelAdapter adapter, object app)
    {
        var previous = CurrentMode;
        CurrentMode = ProbeExecutionMode.Execute;
        try
        {
            var sw = Stopwatch.StartNew();
            ActivateProbeWorkbook();
            var beforeE1 = ReadComStyle(adapter, app, ProbeSheet, "E1");
            var beforeD1 = ReadComStyle(adapter, app, ProbeSheet, "D1");
            var beforeE2 = ReadComStyle(adapter, app, ProbeSheet, "E2");
            var ops = new JsonArray(
                FormatOp(ProbeSheet, "E1", new JsonObject { ["bold"] = true }),
                FormatOp(ProbeSheet, "D1", new JsonObject { ["fontSize"] = OversizedFont }),
                FormatOp(ProbeSheet, "E2", new JsonObject { ["italic"] = true }));
            var batch = ApplyBatch(host, ops);
            var apply = batch.Apply;
            var applyMs = batch.ApplyMs;
            var afterE1 = ReadComStyle(adapter, app, ProbeSheet, "E1");
            var afterD1 = ReadComStyle(adapter, app, ProbeSheet, "D1");
            var afterE2 = ReadComStyle(adapter, app, ProbeSheet, "E2");
            var unsupported = IsExecuteUnsupported(apply, null);
            var results = Json.GetArr(apply, "operationResults");
            var op0Applied = OpResultOk(results, 0);
            var op1ComFailure = OpHasComFailure(results, 1);
            var op2Skipped = OpWasSkipped(results, 2);
            var nativeRestored = StylesEqual(beforeE1, afterE1) &&
                                 StylesEqual(beforeD1, afterD1) &&
                                 StylesEqual(beforeE2, afterE2);
            var rollback = ExactRollback(apply);
            var ok = !unsupported && !Json.GetBool(apply, "ok") &&
                     op0Applied && op1ComFailure && op2Skipped && nativeRestored;
            sw.Stop();
            return Case("execute-fault-after-success", ok, sw.ElapsedMilliseconds,
                new JsonArray(
                    Phase("execute", apply, applyMs),
                    rollback,
                    new JsonObject { ["name"] = "native-before", ["E1"] = beforeE1, ["D1"] = beforeD1, ["E2"] = beforeE2 },
                    new JsonObject { ["name"] = "native-after", ["E1"] = afterE1, ["D1"] = afterD1, ["E2"] = afterE2 }),
                new JsonArray(
                    $"executeUnsupported={unsupported} (unsupported is skip/failure, not a candidate pass)",
                    $"op0Applied={op0Applied} op1ComFailure={op1ComFailure} op2Skipped={op2Skipped}",
                    $"nativeRestored={nativeRestored} (required; rollback.verified={Json.GetBool(rollback, "verified")} is not sufficient)",
                    "pass requires op0 success, op1 COM failure, op2 skipped, and native COM restore"),
                unsupported);
        }
        finally { CurrentMode = previous; }
    }

    private static JsonObject? AttachRestore(DocBridgeHost host, JsonObject flow, bool executeUnsupported)
    {
        if (executeUnsupported || !Json.GetBool(flow, "ok")) return null;
        var restored = RestoreRaw(host, Json.GetString(flow, "snapshotId"));
        var phase = Phase("restore", restored.Result, restored.Ms);
        flow["restore"] = phase;
        return phase;
    }

    private static bool OpHasComFailure(JsonArray? results, int index)
    {
        if (results is null || index < 0 || index >= results.Count || results[index] is not JsonObject op)
            return false;
        if (Json.GetBool(op, "ok")) return false;
        var text = string.Join(" ", Json.GetArr(op, "errors")?.Select(n => n?.ToString()) ?? Array.Empty<string?>());
        text += Json.GetString(op, "error") ?? "";
        text += Json.GetString(op, "exceptionType") ?? "";
        return text.Contains("COM", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("HRESULT", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("0x", StringComparison.OrdinalIgnoreCase) ||
               Json.GetInt(op, "hresult") is not null;
    }

    private static bool OpWasSkipped(JsonArray? results, int index)
    {
        if (results is null || index < 0 || index >= results.Count || results[index] is null)
            return true;
        if (results[index] is not JsonObject op) return true;
        if (Json.GetBool(op, "ok")) return false;
        if (Json.GetBool(op, "skipped")) return true;
        var text = string.Join(" ", Json.GetArr(op, "errors")?.Select(n => n?.ToString()) ?? Array.Empty<string?>());
        text += Json.GetString(op, "error") ?? "";
        if (text.Contains("skipped", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("not attempted", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("not executed", StringComparison.OrdinalIgnoreCase))
            return true;
        return !OpHasComFailure(results, index);
    }

    private static bool IsExecuteUnsupported(JsonObject apply, JsonObject? dry)
    {
        if (Json.GetBool(apply, "dryRun")) return true;
        var errors = JoinErrors(apply);
        if (dry is not null) errors = JoinErrors(dry) + " | " + errors;
        return errors.Contains("requires confirmToken", StringComparison.OrdinalIgnoreCase) ||
               errors.Contains("unknown executionMode", StringComparison.OrdinalIgnoreCase) ||
               errors.Contains("executionMode is not supported", StringComparison.OrdinalIgnoreCase) ||
               errors.Contains("unsupported executionMode", StringComparison.OrdinalIgnoreCase);
    }

    private static long HostPhaseMs(JsonObject flow)
    {
        long ms = 0;
        foreach (var key in new[] { "dryrun", "apply", "restore" })
        {
            var phase = Json.GetObj(flow, key);
            if (phase?["wallMs"] is JsonValue value && value.TryGetValue<long>(out var wall))
                ms += wall;
        }
        return ms;
    }

    // Live100 measured 2–13s per 100 independent cell reads. Keep each STA hop bounded
    // so a 7000-cell verifier does not hit ExcelAdapter's 120s production COM timeout.
    private const int NativeScanChunkCells = 200;

    private sealed class ScaleCellScan
    {
        public required bool?[] Bold { get; init; }
        public required double?[] Fill { get; init; }
        public required int?[] Pattern { get; init; }
        public required string Digest { get; init; }
        public int Checked { get; init; }
        public int Expected { get; init; }
        public int Chunks { get; init; }
        public int ChunkCells { get; init; }
        public int BoldMismatches { get; init; }
        public int FillMismatches { get; init; }
        public int PatternMismatches { get; init; }
        public int Unreadable { get; init; }
        public int Mismatches { get; init; }
        public long Ms { get; set; }
        public JsonArray Samples { get; init; } = new();
        public bool Ok => Checked == Expected && Mismatches == 0 && Unreadable == 0;
    }

    private static JsonObject NativeScanPhase(string name, ScaleCellScan scan) => new()
    {
        ["name"] = name,
        ["ok"] = scan.Ok,
        ["kind"] = "per-cell-digest",
        ["checked"] = scan.Checked,
        ["expected"] = scan.Expected,
        ["digest"] = scan.Digest,
        ["chunks"] = scan.Chunks,
        ["chunkCells"] = scan.ChunkCells,
        ["boldMismatches"] = scan.BoldMismatches,
        ["fillMismatches"] = scan.FillMismatches,
        ["patternMismatches"] = scan.PatternMismatches,
        ["unreadable"] = scan.Unreadable,
        ["mismatches"] = scan.Mismatches,
        ["wallMs"] = scan.Ms,
        ["samples"] = scan.Samples.DeepClone(),
    };

    private static ScaleCellScan ScanScaleCells(
        ExcelAdapter adapter, object app, int scale,
        bool? requireBold = null, double? requireFill = null, int? requirePattern = null,
        bool?[]? expectBold = null, double?[]? expectFill = null, int?[]? expectPattern = null)
    {
        var rows = ScaleRows(scale);
        var cols = 10;
        var expected = rows * cols;
        var bold = new bool?[expected];
        var fill = new double?[expected];
        var pattern = new int?[expected];
        var boldMismatches = 0;
        var fillMismatches = 0;
        var patternMismatches = 0;
        var unreadable = 0;
        var checkedCount = 0;
        var samples = new JsonArray();
        var chunks = 0;
        var sw = Stopwatch.StartNew();
        for (var start = 0; start < expected; start += NativeScanChunkCells)
        {
            var count = Math.Min(NativeScanChunkCells, expected - start);
            chunks++;
            var chunk = adapter.RunOnAdapterThread(() =>
                ScanScaleCellChunk(app, start, count, cols, requireBold, requireFill, requirePattern,
                    expectBold, expectFill, expectPattern));
            for (var offset = 0; offset < count; offset++)
            {
                var i = start + offset;
                bold[i] = chunk.Bold[offset];
                fill[i] = chunk.Fill[offset];
                pattern[i] = chunk.Pattern[offset];
            }
            boldMismatches += chunk.BoldMismatches;
            fillMismatches += chunk.FillMismatches;
            patternMismatches += chunk.PatternMismatches;
            unreadable += chunk.Unreadable;
            checkedCount += chunk.Checked;
            foreach (var sample in chunk.Samples)
            {
                if (samples.Count >= 8) break;
                if (sample is null) continue;
                samples.Add(sample.DeepClone());
            }
        }
        sw.Stop();
        return new ScaleCellScan
        {
            Bold = bold,
            Fill = fill,
            Pattern = pattern,
            Digest = DigestScaleCells(bold, fill, pattern),
            Checked = checkedCount,
            Expected = expected,
            Chunks = chunks,
            ChunkCells = NativeScanChunkCells,
            BoldMismatches = boldMismatches,
            FillMismatches = fillMismatches,
            PatternMismatches = patternMismatches,
            Unreadable = unreadable,
            Mismatches = boldMismatches + fillMismatches + patternMismatches,
            Ms = sw.ElapsedMilliseconds,
            Samples = samples,
        };
    }

    private sealed class ScaleCellChunk
    {
        public required bool?[] Bold { get; init; }
        public required double?[] Fill { get; init; }
        public required int?[] Pattern { get; init; }
        public int Checked { get; init; }
        public int BoldMismatches { get; init; }
        public int FillMismatches { get; init; }
        public int PatternMismatches { get; init; }
        public int Unreadable { get; init; }
        public JsonArray Samples { get; init; } = new();
    }

    private static ScaleCellChunk ScanScaleCellChunk(
        object app, int start, int count, int cols,
        bool? requireBold, double? requireFill, int? requirePattern,
        bool?[]? expectBold, double?[]? expectFill, int?[]? expectPattern)
    {
        var bold = new bool?[count];
        var fill = new double?[count];
        var pattern = new int?[count];
        var boldMismatches = 0;
        var fillMismatches = 0;
        var patternMismatches = 0;
        var unreadable = 0;
        var checkedCount = 0;
        var samples = new JsonArray();
        var sheet = Sheet(app, ProbeSheet);
        try
        {
            for (var offset = 0; offset < count; offset++)
            {
                var i = start + offset;
                var r = i / cols + 1;
                var c = i % cols + 1;
                object? cell = null; object? font = null; object? interior = null;
                try
                {
                    // Independent verifier: one COM cell at a time. Do not read Font.Bold /
                    // Interior.Color on the multi-cell scale range (that is the shortcut under test).
                    var address = $"{ColName(c)}{19 + r}";
                    cell = (object)((dynamic)sheet).Range(address);
                    font = (object)((dynamic)cell).Font;
                    interior = (object)((dynamic)cell).Interior;
                    var cellBold = ReadCellBool(() => ((dynamic)font).Bold);
                    var cellFill = ReadCellDouble(() => ((dynamic)interior).Color);
                    var cellPattern = ReadCellInt(() => ((dynamic)interior).Pattern);
                    bold[offset] = cellBold;
                    fill[offset] = cellFill;
                    pattern[offset] = cellPattern;

                    var boldUnread = cellBold is null;
                    var fillUnread = cellFill is null;
                    var patternUnread = cellPattern is null;
                    if (boldUnread || fillUnread || patternUnread) unreadable++;
                    else checkedCount++;

                    // Missing/unreadable is a failure. Do not treat null as false, and do not
                    // treat two unreadable values as equal.
                    var boldFail = boldUnread
                        || (requireBold is { } rb && cellBold!.Value != rb)
                        || (expectBold is not null && (expectBold[i] is not { } eb || cellBold!.Value != eb));
                    var fillFail = fillUnread
                        || (requireFill is { } rf && !NumbersEqual(cellFill, rf))
                        || (expectFill is not null && (expectFill[i] is not { } ef || !NumbersEqual(cellFill, ef)));
                    var patternFail = patternUnread
                        || (requirePattern is { } rp && cellPattern!.Value != rp)
                        || (expectPattern is not null && (expectPattern[i] is not { } ep || cellPattern!.Value != ep));
                    if (boldFail) boldMismatches++;
                    if (fillFail) fillMismatches++;
                    if (patternFail) patternMismatches++;
                    if ((boldFail || fillFail || patternFail) && samples.Count < 8)
                        samples.Add($"{address}: bold={Fmt(cellBold)} fill={Fmt(cellFill)} pattern={Fmt(cellPattern)}");
                }
                finally
                {
                    RotHelper.ReleaseComReference(interior);
                    RotHelper.ReleaseComReference(font);
                    RotHelper.ReleaseComReference(cell);
                }
            }
        }
        finally { RotHelper.ReleaseComObject(sheet); }

        return new ScaleCellChunk
        {
            Bold = bold,
            Fill = fill,
            Pattern = pattern,
            Checked = checkedCount,
            BoldMismatches = boldMismatches,
            FillMismatches = fillMismatches,
            PatternMismatches = patternMismatches,
            Unreadable = unreadable,
            Samples = samples,
        };
    }

    private static string DigestScaleCells(bool?[] bold, double?[] fill, int?[] pattern)
    {
        var sb = new StringBuilder(bold.Length * 24);
        for (var i = 0; i < bold.Length; i++)
        {
            sb.Append(bold[i] is null ? 'N' : bold[i]!.Value ? 'T' : 'F').Append('|');
            sb.Append(fill[i] is { } f ? f.ToString("R", CultureInfo.InvariantCulture) : "N").Append('|');
            sb.Append(pattern[i] is { } p ? p.ToString(CultureInfo.InvariantCulture) : "N").Append('\n');
        }
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(sb.ToString()))).ToLowerInvariant();
    }

    private static bool? ReadCellBool(Func<object?> read)
    {
        try
        {
            var v = read();
            if (IsMissingComValue(v)) return null;
            if (v is bool flag) return flag;
            if (v is sbyte or byte or short or ushort or int or uint or long or ulong)
            {
                var n = Convert.ToInt64(v, CultureInfo.InvariantCulture);
                if (n == 0) return false;
                if (n is -1 or 1) return true;
            }
            return null;
        }
        catch
        {
            return null;
        }
    }

    private static double? ReadCellDouble(Func<object?> read)
    {
        try
        {
            var v = read();
            if (IsMissingComValue(v)) return null;
            if (v is double d) return d;
            if (v is float f) return f;
            if (v is decimal m) return (double)m;
            if (v is sbyte or byte or short or ushort or int or uint or long or ulong)
                return Convert.ToDouble(v, CultureInfo.InvariantCulture);
            return null;
        }
        catch
        {
            return null;
        }
    }

    private static int? ReadCellInt(Func<object?> read)
    {
        try
        {
            var v = read();
            if (IsMissingComValue(v)) return null;
            if (v is int i) return i;
            if (v is sbyte or byte or short or ushort or uint or long or ulong)
                return Convert.ToInt32(v, CultureInfo.InvariantCulture);
            return null;
        }
        catch
        {
            return null;
        }
    }

    private static bool IsMissingComValue(object? value) =>
        value is null or DBNull ||
        ReferenceEquals(value, Type.Missing) ||
        value is Missing;

    private static string Fmt(bool? value) => value is null ? "null" : value.Value ? "true" : "false";
    private static string Fmt(double? value) => value is null ? "null" : value.Value.ToString("R", CultureInfo.InvariantCulture);
    private static string Fmt(int? value) => value is null ? "null" : value.Value.ToString(CultureInfo.InvariantCulture);

    private static bool ColorFieldsEqual(JsonObject? before, JsonObject? after)
    {
        if (before is null || after is null) return false;
        foreach (var key in new[]
                 {
                     "fontColor", "fontColorIndex", "fontThemeColorPresent", "fontThemeColor", "fontTint",
                     "fillColor", "fillColorIndex", "fillThemeColorPresent", "fillThemeColor", "fillTint",
                     "fillPattern", "fillPatternColor",
                 })
        {
            var a = before[key];
            var b = after[key];
            if (a is null && b is null) continue;
            if (a is JsonValue av && b is JsonValue bv && TryNumber(av, out var an) && TryNumber(bv, out var bn))
            {
                if (!NumbersEqual(an, bn)) return false;
                continue;
            }
            if (!string.Equals(Json.ToCompact(a), Json.ToCompact(b), StringComparison.Ordinal)) return false;
        }
        return true;
    }

    private static void SeedMixedScaleStyles(ExcelAdapter adapter, object app, int scale)
    {
        var lastRow = ScaleLastRow(scale);
        var fills = new[]
        {
            CanonicalFillOle, AliasFillOle, UntintedRgbOle, FingerprintSeedOle, CanonicalFillOle,
            AliasFillOle, UntintedRgbOle, FingerprintSeedOle, CanonicalFillOle, AliasFillOle,
        };
        adapter.RunOnAdapterThread<object?>(() =>
        {
            var sheet = Sheet(app, ProbeSheet);
            try
            {
                for (var c = 1; c <= 10; c++)
                    SetCellFillRaw(sheet, $"{ColName(c)}20:{ColName(c)}{lastRow}", fills[c - 1]);
                SetRangeBoldRaw(sheet, $"D20:D{lastRow}", true);
            }
            finally { RotHelper.ReleaseComObject(sheet); }
            return null;
        });
    }

    private static void SetRangeBoldRaw(object sheet, string address, bool bold)
    {
        object? range = null; object? font = null;
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

    private static void SetCellBold(ExcelAdapter adapter, object app, string sheetName, string address, bool bold)
    {
        adapter.RunOnAdapterThread<object?>(() =>
        {
            var sheet = Sheet(app, sheetName);
            try { SetRangeBoldRaw(sheet, address, bold); }
            finally { RotHelper.ReleaseComObject(sheet); }
            return null;
        });
    }

    private static void StripSecrets(JsonNode? node)
    {
        if (node is JsonObject obj)
        {
            obj.Remove("confirmToken");
            foreach (var kv in obj.ToList())
                StripSecrets(kv.Value);
        }
        else if (node is JsonArray array)
        {
            foreach (var item in array) StripSecrets(item);
        }
    }
}
