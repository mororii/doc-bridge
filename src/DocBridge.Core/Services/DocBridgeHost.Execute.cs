using System.Diagnostics;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Models;

namespace DocBridge.Core.Services;

public sealed partial class DocBridgeHost
{
    private void AttachExecutionCapabilities(string app, JsonObject capability)
    {
        var auto = _policy.AutoExecuteOps(app);
        var coverage = AdapterRollbackCoverageHint(app);
        var execution = new JsonObject
        {
            ["modes"] = new JsonArray("legacy", "execute"),
            ["executeRequires"] = new JsonArray("requestId", "expectedDocumentRef"),
            ["executeConflictsWith"] = new JsonArray("dryRun", "confirmToken", "highRiskConfirm"),
            ["highRiskRequiresLegacyReview"] = true,
            ["rollbackCoverage"] = coverage.DeepClone(),
        };
        if (IsCadApp(app))
            execution["cadRollbackCoverage"] = coverage.DeepClone();
        capability["autoExecuteOps"] = Json.ToArray(auto);
        capability["execution"] = execution;
    }

    private JsonObject ApplyExecute(
        string app,
        string tool,
        OperationValidator.ParsedBatch parsed,
        string opsHash,
        JsonObject timings,
        Func<JsonObject> auditTimings)
    {
        var requestId = parsed.RequestId!;
        var expected = parsed.ExpectedDocumentRef!;

        // Idempotency lookup precedes every adapter call so a completed result
        // can be replayed after process restart without touching the app.
        var early = TryReplayOrDeny(app, tool, parsed, opsHash);
        if (early is not null) return early;

        return WithAutomationLock(() =>
        {
            var raced = TryReplayOrDeny(app, tool, parsed, opsHash);
            if (raced is not null) return raced;

            if (!_executeJournal.TryBegin(requestId, app, expected, opsHash))
            {
                var after = _executeJournal.Lookup(requestId);
                if (after.State == ExecuteJournalState.Completed &&
                    after.Result is not null &&
                    _executeJournal.Matches(after, app, expected, opsHash))
                    return Replay(after, timings);
                return DenyExecute(
                    tool, app, parsed, opsHash,
                    after.State == ExecuteJournalState.Started
                        ? "requestId is pending with outcome unknown; it will not be retried automatically"
                        : "execute journal could not record start; fail closed",
                    journalStarted: false,
                    outcomeUnknown: after.State is ExecuteJournalState.Started or ExecuteJournalState.Unreadable,
                    journalStatus: after.State == ExecuteJournalState.Started ? "started" : "unreadable");
            }

            IAppAdapter adapter;
            try
            {
                adapter = _router.Get(app);
            }
            catch (Exception ex)
            {
                return FinishExecuteDeny(tool, app, parsed, opsHash,
                    $"adapter unavailable: {ex.Message}", journalStarted: true);
            }

            AdapterStatus status;
            var statusStarted = Stopwatch.StartNew();
            try
            {
                status = adapter.GetStatus();
            }
            catch (Exception ex)
            {
                return UncertainExecuteDeny(tool, app, parsed, opsHash,
                    $"execute status failed with outcome unknown: {ex.Message}");
            }
            statusStarted.Stop();
            timings["statusMs"] = statusStarted.ElapsedMilliseconds;
            if (!status.Available)
            {
                return FinishExecuteDeny(tool, app, parsed, opsHash,
                    $"{app} program not available: {status.Detail ?? "not detected"}",
                    journalStarted: true);
            }

            if (!TryBindExecuteDocument(app, expected, parsed.ExplicitDocumentRef, status.Document, out var targetDocument, out var bindError))
            {
                return FinishExecuteDeny(tool, app, parsed, opsHash, bindError, journalStarted: true);
            }

            ApplyPreview preview;
            var previewStarted = Stopwatch.StartNew();
            try
            {
                preview = adapter.Preview(parsed.Ops);
            }
            catch (Exception ex)
            {
                return UncertainExecuteDeny(tool, app, parsed, opsHash,
                    $"execute preview failed with outcome unknown: {ex.Message}");
            }
            previewStarted.Stop();
            timings["previewMs"] = previewStarted.ElapsedMilliseconds;
            timings["previewReused"] = false;
            timings["previewCacheHit"] = false;
            AddDistinctWarnings(preview.Warnings, parsed.OptimizationWarnings);

            if (preview.Errors.Count > 0)
            {
                var pe = new JsonObject
                {
                    ["ok"] = false,
                    ["dryRun"] = false,
                    ["executionMode"] = OperationValidator.ExecutionModeExecute,
                    ["requestId"] = requestId,
                };
                pe["errors"] = Json.ToArray(preview.Errors);
                pe["warnings"] = Json.ToArray(preview.Warnings);
                pe["interaction"] = preview.Interaction?.DeepClone();
                _audit.Write(tool, app, "deny", null, false, preview.Errors);
                _executeJournal.TryComplete(requestId, app, expected, opsHash, pe);
                return pe;
            }

            if (preview.RequiresHighRiskApproval || parsed.HasHighRiskOps)
            {
                return FinishExecuteDeny(tool, app, parsed, opsHash,
                    "executionMode=execute is blocked for high-risk work; use dry-run + confirmToken. highRiskConfirm does not authenticate human approval",
                    journalStarted: true);
            }

            var snapshotStarted = Stopwatch.StartNew();
            SnapshotInfo info;
            try
            {
                info = _snapshots.Create(app, $"{tool} execute", targetDocument,
                    (dir, meta) => adapter.CaptureSnapshot(dir, meta, parsed.Ops));
            }
            catch (Exception ex)
            {
                return UncertainExecuteDeny(tool, app, parsed, opsHash,
                    $"execute snapshot failed with outcome unknown: {ex.Message}");
            }
            snapshotStarted.Stop();
            timings["snapshotMs"] = snapshotStarted.ElapsedMilliseconds;
            timings["snapshotReused"] = false;

            var found = _snapshots.Get(info.SnapshotId);
            if (found is null)
            {
                return FinishExecuteDeny(tool, app, parsed, opsHash,
                    "execute snapshot could not be read back; fail closed",
                    journalStarted: true);
            }

            if (!CapturedSnapshotMatchesExpected(app, expected, found.Value.Info, found.Value.Metadata, out var snapshotError))
            {
                return FinishExecuteDeny(tool, app, parsed, opsHash, snapshotError, journalStarted: true);
            }

            var written = CompleteWrite(
                tool, app, adapter, parsed, info.SnapshotId,
                found.Value.Info, found.Value.Metadata,
                timings, auditTimings);
            written["executionMode"] = OperationValidator.ExecutionModeExecute;
            written["requestId"] = requestId;

            if (!_executeJournal.TryComplete(requestId, app, expected, opsHash, written))
            {
                var warnings = Json.GetArr(written, "warnings") ?? new JsonArray();
                warnings.Add("execute journal could not persist the completed result; the same requestId will not be retried automatically");
                written["warnings"] = warnings;
            }
            return written;
        }, elapsed => timings["lockWaitMs"] = elapsed);
    }

    private JsonObject? TryReplayOrDeny(
        string app, string tool, OperationValidator.ParsedBatch parsed, string opsHash)
    {
        var requestId = parsed.RequestId!;
        var expected = parsed.ExpectedDocumentRef!;
        var record = _executeJournal.Lookup(requestId);
        switch (record.State)
        {
            case ExecuteJournalState.Missing:
                return null;
            case ExecuteJournalState.Unreadable:
                return DenyExecute(tool, app, parsed, opsHash,
                    "execute journal is corrupt or missing a completed result; fail closed",
                    journalStarted: false, outcomeUnknown: true, journalStatus: "unreadable");
            case ExecuteJournalState.Started:
                return DenyExecute(tool, app, parsed, opsHash,
                    "requestId is pending with outcome unknown; it will not be retried automatically",
                    journalStarted: false, outcomeUnknown: true, journalStatus: "started");
            case ExecuteJournalState.Completed when record.Result is null:
                return DenyExecute(tool, app, parsed, opsHash,
                    "execute journal completed record is missing its result; fail closed",
                    journalStarted: false, outcomeUnknown: true, journalStatus: "completed-missing-result");
            case ExecuteJournalState.Completed:
                if (!_executeJournal.Matches(record, app, expected, opsHash))
                    return DenyExecute(tool, app, parsed, opsHash,
                        "requestId was already used with a different app, document, or ops",
                        journalStarted: false);
                _audit.Write(tool, app, "replay", new JsonObject
                {
                    ["requestId"] = requestId,
                    ["ops"] = OpsSummary(parsed.Ops),
                }, Json.GetBool(record.Result, "ok"));
                return Replay(record, timings: null);
            default:
                return DenyExecute(tool, app, parsed, opsHash,
                    "execute journal state is unknown; fail closed",
                    journalStarted: false);
        }
    }

    private static JsonObject Replay(ExecuteJournalRecord record, JsonObject? timings)
    {
        var replayed = (JsonObject)record.Result!.DeepClone();
        replayed["idempotentReplay"] = true;
        replayed["executionMode"] = OperationValidator.ExecutionModeExecute;
        replayed["requestId"] = record.RequestId;
        if (timings is not null) replayed["timings"] = timings.DeepClone();
        return replayed;
    }

    private JsonObject FinishExecuteDeny(
        string tool, string app, OperationValidator.ParsedBatch parsed, string opsHash,
        string message, bool journalStarted) =>
        DenyExecute(tool, app, parsed, opsHash, message, journalStarted);

    private JsonObject UncertainExecuteDeny(
        string tool, string app, OperationValidator.ParsedBatch parsed, string opsHash,
        string message) =>
        DenyExecute(tool, app, parsed, opsHash, message,
            journalStarted: false, outcomeUnknown: true, journalStatus: "started");

    private JsonObject DenyExecute(
        string tool, string app, OperationValidator.ParsedBatch parsed, string opsHash,
        string message, bool journalStarted, bool outcomeUnknown = false, string? journalStatus = null)
    {
        var result = Json.ErrorResult(message, app);
        result["dryRun"] = false;
        result["executionMode"] = OperationValidator.ExecutionModeExecute;
        result["requestId"] = parsed.RequestId;
        result["safeToRetry"] = false;
        if (outcomeUnknown)
        {
            result["outcomeUnknown"] = true;
            result["journalStatus"] = journalStatus ?? "unknown";
        }
        _audit.Write(tool, app, "deny", new JsonObject { ["requestId"] = parsed.RequestId }, false, new[] { message });
        if (journalStarted)
            _executeJournal.TryComplete(parsed.RequestId!, app, parsed.ExpectedDocumentRef!, opsHash, result);
        return result;
    }

    private JsonObject CompleteWrite(
        string tool,
        string app,
        IAppAdapter adapter,
        OperationValidator.ParsedBatch parsed,
        string snapshotId,
        SnapshotInfo snapshotInfo,
        JsonObject snapshotMetadata,
        JsonObject timings,
        Func<JsonObject> auditTimings)
    {
        var applyStarted = Stopwatch.StartNew();
        ApplyExecution exec;
        HwpAutomationException? hwpApplyError = null;
        try
        {
            exec = adapter.Apply(parsed.Ops, snapshotId);
        }
        catch (Exception applyError)
        {
            exec = new ApplyExecution { Ok = false };
            hwpApplyError = app.Equals("hwp", StringComparison.OrdinalIgnoreCase)
                ? FindHwpAutomationException(applyError)
                : null;
            exec.Errors.Add(hwpApplyError is null
                ? $"apply threw: {applyError.Message}"
                : $"[{hwpApplyError.Code}] {hwpApplyError.Message}");
        }
        applyStarted.Stop();
        timings["applyMs"] = applyStarted.ElapsedMilliseconds;

        if (exec.Errors.Count == 0)
        {
            foreach (var node in Json.GetArr(exec.Readback, "mismatches") ?? new JsonArray())
            {
                if (node is JsonValue value && value.TryGetValue<string>(out var text) &&
                    !string.IsNullOrWhiteSpace(text))
                    exec.Errors.Add(text);
            }
        }

        if (exec.Ok && !HasExplicitReadbackProof(exec.Readback))
        {
            exec.Ok = false;
            exec.Errors.Add("apply did not return explicit readback.verified=true; the write is not treated as proven");
        }

        if (exec.OperationResults.Count == 0)
        {
            for (var index = 0; index < parsed.Ops.Count; index++)
            {
                exec.OperationResults.Add(new JsonObject
                {
                    ["index"] = index,
                    ["op"] = Json.GetString(parsed.Ops[index], "op") ?? "?",
                    ["ok"] = exec.Ok,
                    ["elapsedMs"] = index == 0 ? applyStarted.ElapsedMilliseconds : 0,
                    ["timingScope"] = "batch-fallback",
                });
            }
        }

        var rollback = new JsonObject
        {
            ["attempted"] = false,
            ["verified"] = false,
        };
        if (!exec.Ok)
        {
            rollback["attempted"] = true;
            var rollbackStarted = Stopwatch.StartNew();
            try
            {
                var restored = adapter.RestoreSnapshot(snapshotInfo.Dir, snapshotMetadata);
                rollback["result"] = restored.DeepClone();
                rollback["verified"] = IsVerifiedRestore(restored, snapshotMetadata);
                if (Json.GetBool(rollback, "verified"))
                    exec.Warnings.Add("apply failed; the pre-apply snapshot was restored automatically");
                else
                    exec.Errors.Add("apply failed and automatic rollback could not be verified");
            }
            catch (Exception rollbackError)
            {
                rollback["error"] = rollbackError.Message;
                exec.Errors.Add($"automatic rollback failed: {rollbackError.Message}");
            }
            finally
            {
                rollbackStarted.Stop();
                rollback["elapsedMs"] = rollbackStarted.ElapsedMilliseconds;
                timings["rollbackMs"] = rollbackStarted.ElapsedMilliseconds;
            }
        }

        AttachAdapterRollbackCoverage(app, rollback, snapshotMetadata, exec);

        _audit.Write(tool, app, "apply", new JsonObject
        {
            ["snapshotId"] = snapshotId,
            ["ops"] = OpsSummary(parsed.Ops),
            ["readbackOk"] = HasExplicitReadbackProof(exec.Readback),
            ["elapsedMs"] = applyStarted.ElapsedMilliseconds,
            ["rollback"] = rollback.DeepClone(),
            ["operationResults"] = OperationTimingSummary(exec.OperationResults),
            ["timings"] = auditTimings(),
        }, exec.Ok, exec.Errors);

        var applyResult = new JsonObject
        {
            ["ok"] = exec.Ok,
            ["dryRun"] = false,
            ["snapshotId"] = snapshotId,
            ["affected"] = Json.ToArray(exec.Affected),
            ["diff"] = Json.ToArray(exec.Diff),
            ["operationResults"] = exec.OperationResults.DeepClone(),
            ["elapsedMs"] = applyStarted.ElapsedMilliseconds,
            ["readback"] = exec.Readback?.DeepClone(),
            ["interaction"] = exec.Interaction?.DeepClone(),
            ["rollback"] = rollback,
            ["warnings"] = Json.ToArray(exec.Warnings),
            ["errors"] = Json.ToArray(exec.Errors),
        };
        if (hwpApplyError is not null) AddHwpAutomationError(applyResult, app, hwpApplyError);
        applyResult["rollbackCoverage"] = DescribeRollbackCoverage(app, snapshotMetadata, rollback);
        return applyResult;
    }

    private static bool TryBindExecuteDocument(
        string app,
        string expected,
        string? explicitDocumentRef,
        string? statusDocument,
        out string targetDocument,
        out string error)
    {
        targetDocument = expected;
        error = "";

        // CAD/Gstar Preview+Apply use ActiveDocWait and ignore op.document except
        // activate_document. An explicit binder match is not adapter routing.
        if (IsCadApp(app))
        {
            if (!DocumentIdentity.TryNormalizeExecuteRef(app, statusDocument, out _, out _))
            {
                error = $"current document '{statusDocument}' is not a stable identity. " +
                        DocumentIdentity.LegacyPreviewGuidance;
                return false;
            }

            if (!SameDocumentRef(app, expected, statusDocument))
            {
                error =
                    $"expectedDocumentRef '{expected}' does not match current document '{statusDocument}'. " +
                    "CAD/Gstar execute always binds the active drawing; op.document is not used to select a target.";
                return false;
            }

            return true;
        }

        if (explicitDocumentRef is not null)
            return true;

        if (!DocumentIdentity.TryNormalizeExecuteRef(app, statusDocument, out _, out _))
        {
            error = $"current document '{statusDocument}' is not a stable identity. " +
                    DocumentIdentity.LegacyPreviewGuidance;
            return false;
        }

        if (!SameDocumentRef(app, expected, statusDocument))
        {
            error = $"expectedDocumentRef '{expected}' does not match current document '{statusDocument}'";
            return false;
        }

        return true;
    }

    internal static bool CapturedSnapshotMatchesExpected(
        string app, string expected, SnapshotInfo info, JsonObject metadata, out string error)
    {
        error = "";
        var infoRef = info.DocumentRef;
        var metaRef = Json.GetString(metadata, "documentRef");
        if (!HasConcreteDocumentRef(infoRef) || !HasConcreteDocumentRef(metaRef))
        {
            error = "captured snapshot is missing documentRef; execute will not apply";
            return false;
        }

        if (!SameDocumentRef(app, expected, infoRef) || !SameDocumentRef(app, expected, metaRef))
        {
            error =
                $"captured snapshot document info='{infoRef}' metadata='{metaRef}' " +
                $"does not match expectedDocumentRef '{expected}'";
            return false;
        }

        return true;
    }

    private static bool HasConcreteDocumentRef(string? documentRef) =>
        !string.IsNullOrWhiteSpace(documentRef);

    private static bool IsCadApp(string app) =>
        app.Equals("cad", StringComparison.OrdinalIgnoreCase) ||
        app.Equals("gstarcad", StringComparison.OrdinalIgnoreCase);

    private static JsonObject AdapterRollbackCoverageHint(string app) => new()
    {
        ["fullDocumentRestored"] = false,
        ["coverage"] = "adapter-snapshot-restore",
        ["note"] = IsCadApp(app)
            ? "Adapter-scoped snapshot/restore for the requested CAD execute ops."
            : "Adapter-scoped snapshot/restore for the requested execute ops.",
    };

    private static JsonObject DescribeRollbackCoverage(string app, JsonObject snapshotMetadata, JsonObject rollback)
    {
        var restored = Json.GetObj(rollback, "result");
        var coverage = Json.GetString(restored, "coverage")
                       ?? Json.GetString(restored, "snapshotCoverage")
                       ?? Json.GetString(snapshotMetadata, "snapshotCoverage")
                       ?? Json.GetString(snapshotMetadata, "coverage")
                       ?? "adapter-snapshot-restore";
        var described = new JsonObject
        {
            ["coverage"] = coverage,
            ["source"] = restored is null ? "snapshot-metadata" : "adapter-restore",
            ["fullDocumentRestored"] = restored is not null && IsExplicitTrue(restored, "fullDocumentRestored"),
        };
        var kind = Json.GetString(restored, "kind") ?? Json.GetString(snapshotMetadata, "snapshotKind");
        if (!string.IsNullOrWhiteSpace(kind))
            described["snapshotKind"] = kind;
        return described;
    }

    private static void AttachAdapterRollbackCoverage(
        string app, JsonObject rollback, JsonObject snapshotMetadata, ApplyExecution exec)
    {
        var described = DescribeRollbackCoverage(app, snapshotMetadata, rollback);
        rollback["coverage"] = described["coverage"]?.DeepClone();
        rollback["fullDocumentRestored"] = described["fullDocumentRestored"]?.DeepClone();
        if (described["snapshotKind"] is not null)
            rollback["snapshotKind"] = described["snapshotKind"]!.DeepClone();

        if (!Json.GetBool(rollback, "attempted"))
            return;

        var restored = Json.GetObj(rollback, "result");
        if (HasNonFullCoverage(restored) || HasNonFullCoverage(snapshotMetadata) ||
            HasNonFullCoverage(Json.GetObj(restored, "readback")) ||
            HasNonFullCoverage(described) ||
            (restored is not null && !HasExplicitReadbackProof(Json.GetObj(restored, "readback"))))
        {
            rollback["verified"] = false;
            const string unverified =
                "automatic rollback is not verified: adapter coverage is incomplete/partial or readback.verified is missing";
            if (!exec.Warnings.Contains(unverified, StringComparer.Ordinal))
                exec.Warnings.Add(unverified);
        }
    }

    internal static bool HasExplicitReadbackProof(JsonObject? readback) =>
        readback is not null && IsExplicitTrue(readback, "verified");

    internal static bool IsVerifiedRestore(JsonObject restored, JsonObject? snapshotMetadata = null)
    {
        if (!Json.GetBool(restored, "ok")) return false;
        if (HasNonFullCoverage(restored) || HasNonFullCoverage(snapshotMetadata))
            return false;
        if (IsExplicitFalse(restored, "fullDocumentRestored") || IsExplicitFalse(restored, "complete"))
            return false;
        if (!HasExplicitReadbackProof(Json.GetObj(restored, "readback")))
            return false;
        var readback = Json.GetObj(restored, "readback")!;
        if (HasNonFullCoverage(readback)) return false;
        if (IsExplicitFalse(readback, "fullDocumentRestored")) return false;
        return true;
    }

    private static bool HasNonFullCoverage(JsonObject? obj)
    {
        if (obj is null) return false;
        var coverage = Json.GetString(obj, "coverage") ?? Json.GetString(obj, "snapshotCoverage");
        return coverage is not null &&
               (coverage.Equals("partial", StringComparison.OrdinalIgnoreCase) ||
                coverage.Equals("incomplete", StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsExplicitTrue(JsonObject obj, string key) =>
        obj.TryGetPropertyValue(key, out var node) &&
        node is JsonValue value &&
        value.TryGetValue<bool>(out var flag) &&
        flag;

    private static bool IsExplicitFalse(JsonObject obj, string key) =>
        obj.TryGetPropertyValue(key, out var node) &&
        node is JsonValue value &&
        value.TryGetValue<bool>(out var flag) &&
        !flag;
}
