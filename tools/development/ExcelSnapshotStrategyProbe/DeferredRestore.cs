using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

namespace DocBridge.Development.ExcelSnapshotStrategyProbe;

internal static partial class Program
{
    private const string DeferredValue = "post-checkpoint-value";
    private const string DeferredNumberFormat = "\"deferred-nfmt\"";

    private static JsonObject RunDeferredRestoreExperiment(
        ExcelAdapter adapter, object app, string outputDir, int scale,
        ScaleCellScan nativeBefore, JsonObject keepBefore)
    {
        Progress("CASE", $"deferred-restore fixture-only scale={scale}");
        var ops = FormatOpList(ScaleRange(scale));
        var first = ScaleAddress(0);
        var checkpointPath = Path.Combine(outputDir, $"deferred-checkpoint-{Guid.NewGuid():N}.xlsx");
        var corruptPath = Path.Combine(outputDir, $"deferred-corrupt-{Guid.NewGuid():N}.xlsx");
        Console.WriteLine($"INTERFACE deferredCheckpoint={checkpointPath}");
        var foregroundBefore = ReadForegroundSnapshot();

        var eligibility = FixtureEligibility(ProbePath ?? "", scale);
        if (!Json.GetBool(eligibility, "eligible"))
        {
            return new JsonObject
            {
                ["ok"] = false,
                ["experimental"] = true,
                ["fixtureOnly"] = true,
                ["genericRollout"] = false,
                ["eligibility"] = eligibility,
            };
        }

        var checkpointTotalSw = Stopwatch.StartNew();
        var identitySw = Stopwatch.StartNew();
        var identityBefore = ReadIdentity(adapter, app);
        identitySw.Stop();
        var copy = TimeSaveCopyAs(adapter, app, checkpointPath);
        KnownWorkbookPaths.Add(checkpointPath);
        var shaSw = Stopwatch.StartNew();
        var sha = File.Exists(checkpointPath) ? Sha256File(checkpointPath) : "";
        shaSw.Stop();
        var preflightSw = Stopwatch.StartNew();
        var preflight = PrefightGeneratedXlsx(checkpointPath);
        preflightSw.Stop();
        identitySw.Start();
        var identityAfterCopy = ReadIdentity(adapter, app);
        identitySw.Stop();
        checkpointTotalSw.Stop();
        var checkpoint = copy;
        checkpoint["sha256"] = sha;
        checkpoint["originalIdentity"] = identityBefore.DeepClone();
        checkpoint["documentRef"] = ProbePath;
        checkpoint["preflight"] = preflight;
        checkpoint["copyMs"] = Json.GetLong(copy, "wallMs");
        checkpoint["sha256Ms"] = shaSw.ElapsedMilliseconds;
        checkpoint["ooxmlPreflightMs"] = preflightSw.ElapsedMilliseconds;
        checkpoint["identityMs"] = identitySw.ElapsedMilliseconds;
        checkpoint["checkpointTotalMs"] = checkpointTotalSw.ElapsedMilliseconds;
        checkpoint["copyMsIsNotFullCheckpoint"] = true;

        var result = new JsonObject
        {
            ["ok"] = false,
            ["experimental"] = true,
            ["fixtureOnly"] = true,
            ["genericRollout"] = false,
            ["eligibility"] = eligibility,
            ["originalIdentity"] = identityBefore,
            ["identityAfterCopy"] = identityAfterCopy,
            ["checkpoint"] = checkpoint,
            ["phaseIdentities"] = new JsonArray(),
        };
        AppendPhase(result, RecordPhase(adapter, app, "after-checkpoint"));

        var checkpointOk = Json.GetBool(checkpoint, "ok");
        var preflightOk = Json.GetBool(preflight, "ok");
        var shaNonempty = !string.IsNullOrWhiteSpace(sha);
        var identityReadable = IdentityReadable(identityBefore) && IdentityReadable(identityAfterCopy);
        var identityPreservedAfterCopy = identityReadable && IdentityUnchanged(identityBefore, identityAfterCopy);
        var checkpointEligible = checkpointOk && preflightOk && shaNonempty && identityPreservedAfterCopy;
        result["checkpointEligible"] = new JsonObject
        {
            ["ok"] = checkpointEligible,
            ["checkpointOk"] = checkpointOk,
            ["preflightOk"] = preflightOk,
            ["shaNonempty"] = shaNonempty,
            ["identityReadable"] = identityReadable,
            ["identityPreservedAfterCopy"] = identityPreservedAfterCopy,
        };
        if (!checkpointEligible)
        {
            result["refusedBeforeApply"] = true;
            result["apply"] = new JsonObject
            {
                ["ok"] = false,
                ["skipped"] = true,
                ["reason"] = "checkpoint-ineligible",
            };
            result["error"] = "no checkpoint eligibility; refusing mutation";
            return result;
        }

        try
        {
            var applySw = Stopwatch.StartNew();
            var applyExec = adapter.Apply(ops, snapshotId: "");
            applySw.Stop();
            var apply = new JsonObject
            {
                ["ok"] = applyExec.Ok && applyExec.Errors.Count == 0,
                ["wallMs"] = applySw.ElapsedMilliseconds,
                ["method"] = "ExcelAdapter.Apply",
                ["hostEagerSnapshot"] = false,
                ["errors"] = string.Join(" | ", applyExec.Errors),
            };
            result["apply"] = apply;
            result["timings"] = new JsonObject
            {
                ["copyMs"] = Json.GetLong(checkpoint, "copyMs"),
                ["sha256Ms"] = Json.GetLong(checkpoint, "sha256Ms"),
                ["ooxmlPreflightMs"] = Json.GetLong(checkpoint, "ooxmlPreflightMs"),
                ["identityMs"] = Json.GetLong(checkpoint, "identityMs"),
                ["checkpointTotalMs"] = Json.GetLong(checkpoint, "checkpointTotalMs"),
                ["copyMsIsNotFullCheckpoint"] = true,
                ["applyMs"] = applySw.ElapsedMilliseconds,
            };
            Progress("deferred-apply", $"ok={Json.GetBool(apply, "ok")} wallMs={applySw.ElapsedMilliseconds}");
            AppendPhase(result, RecordPhase(adapter, app, "after-apply"));

            Progress("native-scan", "native-after-deferred-apply");
            var afterApply = ScanScaleCells(adapter, app, scale, compareTo: null);
            var requestedVisible = VerifyRequestedApplyStyles(afterApply);
            var afterApplyVsOld = CompareScaleScans(afterApply, nativeBefore);
            result["nativeAfterApply"] = new JsonObject
            {
                ["readable"] = NativePhase("native-after-deferred-apply", afterApply),
                ["requestedVisibleResult"] = requestedVisible,
                ["oldStateComparison"] = afterApplyVsOld,
                ["note"] = "requestedVisibleResult is the apply assertion; oldStateComparison mismatch is expected after format_range and is not that check",
            };
            SetCellValueAndNumberFormat(adapter, app, first, DeferredValue, DeferredNumberFormat);
            var newer = ReadValueFormat(adapter, app, TargetSheet, first);
            result["newerNonScoped"] = new JsonObject
            {
                ["address"] = first,
                ["beforeRestore"] = newer,
            };

            var success = DeferredExtractAndRestore(adapter, app, checkpointPath, sha, ops, "success");
            result["successRestore"] = success;
            Json.GetObj(result, "timings")!["extractMs"] = Json.GetLong(success, "extractMs");
            Json.GetObj(result, "timings")!["restoreMs"] = Json.GetLong(success, "restoreMs");
            AppendPhase(result, Json.GetObj(success, "identityAfterExtract")?.DeepClone().AsObject()
                ?? RecordPhase(adapter, app, "after-success-extract"));
            AppendPhase(result, RecordPhase(adapter, app, "after-success-restore"));

            Progress("native-scan", "native-after-deferred-restore");
            var afterRestore = ScanScaleCells(adapter, app, scale, compareTo: nativeBefore);
            var newerAfter = ReadValueFormat(adapter, app, TargetSheet, first);
            var keepAfter = ReadCellAudit(adapter, app, KeepSheet, "A1");
            var identityAfterRestore = ReadIdentity(adapter, app);
            var newerPreserved = NewerValueFormatPreserved(newerAfter);
            var keepAfterOk = KeepAuditsEqual(keepBefore, keepAfter);
            var identityOk = IdentityUnchanged(identityBefore, identityAfterCopy) &&
                             IdentityUnchanged(identityBefore, identityAfterRestore) &&
                             Json.GetBool(identityAfterRestore, "saved") == false;
            result["nativeAfterRestore"] = NativePhase("native-after-deferred-restore", afterRestore);
            result["identityAfterRestore"] = identityAfterRestore;
            result["identityPreserved"] = identityOk;
            result["keep"] = new JsonObject
            {
                ["before"] = keepBefore.DeepClone(),
                ["afterRestore"] = keepAfter,
                ["comparedFields"] = new JsonArray("value", "typedValue", "hasFormula", "formula", "fontBold", "fillColor", "fillPattern"),
            };
            if (result["newerNonScoped"] is JsonObject newerNode)
            {
                newerNode["afterRestore"] = newerAfter;
                newerNode["preservedAfterRestore"] = newerPreserved;
            }

            var partial = RunPartialWriteInjectedFailure(adapter, app, scale);
            result["partialWriteInjectedFailure"] = partial;
            AppendPhase(result, RecordPhase(adapter, app, "after-partial-inject"));
            var recovered = DeferredExtractAndRestore(adapter, app, checkpointPath, sha, ops, "partial-recovery");
            result["partialRecovery"] = recovered;
            Json.GetObj(result, "timings")!["partialExtractMs"] = Json.GetLong(recovered, "extractMs");
            Json.GetObj(result, "timings")!["partialRestoreMs"] = Json.GetLong(recovered, "restoreMs");
            var afterPartial = ScanScaleCells(adapter, app, scale, compareTo: nativeBefore);
            var newerAfterPartial = ReadValueFormat(adapter, app, TargetSheet, first);
            var newerPartialPreserved = NewerValueFormatPreserved(newerAfterPartial);
            result["nativeAfterPartialRecovery"] = NativePhase("native-after-partial-recovery", afterPartial);
            if (result["newerNonScoped"] is JsonObject newerNode2)
            {
                newerNode2["afterPartialRecovery"] = newerAfterPartial;
                newerNode2["preservedAfterPartialRecovery"] = newerPartialPreserved;
            }

            var corruption = RunChecksumCorruption(adapter, app, checkpointPath, sha, ops, corruptPath, afterPartial);
            result["checksumCorruption"] = corruption;
            var keepFinal = ReadCellAudit(adapter, app, KeepSheet, "A1");
            var keepFinalOk = KeepAuditsEqual(keepBefore, keepFinal);
            if (result["keep"] is JsonObject keepNode)
            {
                keepNode["final"] = keepFinal;
                keepNode["ok"] = keepAfterOk && keepFinalOk;
            }

            var foregroundAfter = ReadForegroundSnapshot();
            result["foreground"] = ForegroundCompare(foregroundBefore, foregroundAfter);
            result["note"] = "fixture-only generated .xlsx proof; not a generic rollout or crash-restart claim";
            result["ok"] = Json.GetBool(apply, "ok") &&
                           Json.GetBool(requestedVisible, "ok") &&
                           Json.GetBool(success, "ok") &&
                           Json.GetBool(success, "checkpointClosed") &&
                           Json.GetBool(success, "workbookCountPreserved") &&
                           afterRestore.Ok &&
                           newerPreserved &&
                           newerPartialPreserved &&
                           keepAfterOk &&
                           keepFinalOk &&
                           identityOk &&
                           Json.GetBool(partial, "injected") &&
                           Json.GetBool(recovered, "ok") &&
                           Json.GetBool(recovered, "checkpointClosed") &&
                           Json.GetBool(recovered, "workbookCountPreserved") &&
                           afterPartial.Ok &&
                           Json.GetBool(corruption, "rejectedBeforeMutation") &&
                           Json.GetBool(corruption, "stylesUnchangedFromPriorScan");
            return result;
        }
        catch (Exception ex)
        {
            result["ok"] = false;
            result["retainedAfterFailure"] = true;
            result["phaseError"] = ex.Message;
            result["phaseErrorType"] = ex.GetType().FullName;
            if (ex is COMException com) result["phaseHresult"] = com.ErrorCode;
            AppendPhase(result, RecordPhase(adapter, app, "after-failure"));
            result["foreground"] = ForegroundCompare(foregroundBefore, ReadForegroundSnapshot());
            Console.Error.WriteLine(ex);
            return result;
        }
    }

    private static void AppendPhase(JsonObject result, JsonObject? phase)
    {
        if (phase is null) return;
        if (result["phaseIdentities"] is not JsonArray list)
        {
            list = new JsonArray();
            result["phaseIdentities"] = list;
        }

        list.Add(phase.DeepClone());
    }

    private static JsonObject RecordPhase(ExcelAdapter adapter, object app, string phase)
    {
        try
        {
            var identity = ReadIdentity(adapter, app);
            identity["phase"] = phase;
            identity["ok"] = true;
            Progress("phase-identity", $"{phase} pid={Json.GetInt(identity, "pid")} count={Json.GetInt(identity, "workbookCount")} saved={Json.GetBool(identity, "saved")}");
            return identity;
        }
        catch (Exception ex)
        {
            var node = new JsonObject
            {
                ["phase"] = phase,
                ["ok"] = false,
                ["error"] = ex.Message,
                ["errorType"] = ex.GetType().FullName,
            };
            if (ex is COMException com) node["hresult"] = com.ErrorCode;
            Progress("phase-identity", $"{phase} FAILED {ex.GetType().Name}");
            return node;
        }
    }

    private static JsonObject FixtureEligibility(string path, int scale) => new()
    {
        ["eligible"] = scale is 100 or 1000 &&
                       path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) &&
                       string.Equals(Path.GetFileName(path), "excel-snapshot-strategy-target.xlsx", StringComparison.OrdinalIgnoreCase),
        ["generator"] = "ExcelSnapshotStrategyProbe",
        ["format"] = "xlsx",
        ["scale"] = scale,
        ["supportedOldStyles"] = MixedFill
            ? new JsonArray("no-fill", "solid-rgb", "theme-tint", "patterned")
            : new JsonArray("uniform-bold-no-fill"),
        ["genericRollout"] = false,
        ["unsupported"] = new JsonArray("xls", "xlsm", "xlsb", "old-rgb-plus-tint-as-success", "untrusted-offline-color-inverse"),
    };

    private static JsonObject PrefightGeneratedXlsx(string path)
    {
        if (!File.Exists(path) || !path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
            return new JsonObject { ["ok"] = false, ["reason"] = "not a present .xlsx fixture" };
        try
        {
            using var zip = ZipFile.OpenRead(path);
            var names = zip.Entries.Select(e => e.FullName).ToHashSet(StringComparer.OrdinalIgnoreCase);
            var workbook = names.Contains("xl/workbook.xml");
            var types = names.Contains("[Content_Types].xml");
            var vba = names.Any(n => n.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase));
            return new JsonObject
            {
                ["ok"] = workbook && types && !vba,
                ["hasWorkbookXml"] = workbook,
                ["hasContentTypes"] = types,
                ["hasVba"] = vba,
                ["genericRollout"] = false,
            };
        }
        catch (Exception ex)
        {
            return new JsonObject { ["ok"] = false, ["reason"] = ex.Message };
        }
    }

    private static JsonObject DeferredExtractAndRestore(
        ExcelAdapter adapter, object app, string checkpointPath, string expectedSha,
        List<JsonObject> ops, string phase)
    {
        var actualSha = Sha256File(checkpointPath);
        if (!string.Equals(actualSha, expectedSha, StringComparison.OrdinalIgnoreCase))
        {
            return new JsonObject
            {
                ["ok"] = false,
                ["phase"] = phase,
                ["rejectedBeforeMutation"] = true,
                ["error"] = "checkpoint checksum mismatch; restore refused before open/mutation",
                ["expectedSha256"] = expectedSha,
                ["actualSha256"] = actualSha,
            };
        }

        ActivateOwnedWorkbook(adapter, app);
        var extractSw = Stopwatch.StartNew();
        JsonObject extracted;
        try
        {
            extracted = ExtractDeferredState(adapter, app, checkpointPath, ops);
        }
        catch (Exception ex)
        {
            extractSw.Stop();
            return new JsonObject
            {
                ["ok"] = false,
                ["phase"] = phase,
                ["extractMs"] = extractSw.ElapsedMilliseconds,
                ["error"] = ex.Message,
                ["errorType"] = ex.GetType().FullName,
            };
        }

        extractSw.Stop();
        Progress("deferred-extract", $"phase={phase} wallMs={extractSw.ElapsedMilliseconds}");
        extracted["identityAfterExtract"] = RecordPhase(adapter, app, $"{phase}-after-extract");
        if (!Json.GetBool(extracted, "ok"))
        {
            extracted["phase"] = phase;
            extracted["extractMs"] = extractSw.ElapsedMilliseconds;
            return extracted;
        }

        var restoreSw = Stopwatch.StartNew();
        JsonObject restored;
        try
        {
            restored = RestoreDeferredState(adapter, app, Json.GetObj(extracted, "state")!);
        }
        catch (Exception ex)
        {
            restoreSw.Stop();
            return new JsonObject
            {
                ["ok"] = false,
                ["phase"] = phase,
                ["extractMs"] = extractSw.ElapsedMilliseconds,
                ["restoreMs"] = restoreSw.ElapsedMilliseconds,
                ["error"] = ex.Message,
            };
        }

        restoreSw.Stop();
        Progress("deferred-restore", $"phase={phase} ok={Json.GetBool(restored, "ok")} wallMs={restoreSw.ElapsedMilliseconds}");
        return new JsonObject
        {
            ["ok"] = Json.GetBool(restored, "ok"),
            ["phase"] = phase,
            ["extractMs"] = extractSw.ElapsedMilliseconds,
            ["restoreMs"] = restoreSw.ElapsedMilliseconds,
            ["checkpointClosed"] = Json.GetBool(extracted, "checkpointClosed"),
            ["originalActivated"] = Json.GetBool(extracted, "originalActivated"),
            ["documentRefPassed"] = Json.GetString(extracted, "documentRefPassed"),
            ["workbookCountPreserved"] = Json.GetBool(extracted, "workbookCountPreserved"),
            ["fingerprint"] = Json.GetString(Json.GetObj(extracted, "state"), "fingerprint"),
            ["restore"] = restored,
            ["rejectedBeforeMutation"] = false,
        };
    }

    private static JsonObject ExtractDeferredState(
        ExcelAdapter adapter, object app, string checkpointPath, List<JsonObject> ops)
    {
        return adapter.RunOnAdapterThread(() =>
        {
            object? workbooks = null;
            object? original = null;
            object? checkpoint = null;
            try
            {
                workbooks = (object)((dynamic)app).Workbooks;
                original = FindOwnedWorkbook(workbooks);
                var beforeCount = Convert.ToInt32(((dynamic)workbooks).Count, CultureInfo.InvariantCulture);
                var originalName = Convert.ToString((object?)((dynamic)original).FullName, CultureInfo.InvariantCulture);
                checkpoint = (object)((dynamic)workbooks).Open(checkpointPath, 0, true);
                var openedName = Convert.ToString((object?)((dynamic)checkpoint).FullName, CultureInfo.InvariantCulture);
                var state = InvokeCaptureFormatOnlyState(adapter, checkpoint, ops, ProbePath ?? originalName ?? "");
                ((dynamic)checkpoint).Close(false);
                RotHelper.ReleaseComReference(checkpoint);
                checkpoint = null;
                ((dynamic)original).Activate();
                var afterCount = Convert.ToInt32(((dynamic)workbooks).Count, CultureInfo.InvariantCulture);
                return new JsonObject
                {
                    ["ok"] = true,
                    ["state"] = state,
                    ["checkpointClosed"] = true,
                    ["originalActivated"] = string.Equals(originalName, ProbePath, StringComparison.OrdinalIgnoreCase),
                    ["documentRefPassed"] = ProbePath,
                    ["openedCheckpointName"] = openedName,
                    ["workbookCountBefore"] = beforeCount,
                    ["workbookCountAfter"] = afterCount,
                    ["workbookCountPreserved"] = afterCount == beforeCount,
                    ["releaseMode"] = "ReleaseComReference-borrowed",
                };
            }
            finally
            {
                if (checkpoint is not null)
                {
                    try { ((dynamic)checkpoint).Close(false); } catch { /* owned checkpoint only */ }
                    RotHelper.ReleaseComReference(checkpoint);
                }

                RotHelper.ReleaseComReference(original);
                RotHelper.ReleaseComReference(workbooks);
            }
        });
    }

    private static JsonObject RestoreDeferredState(ExcelAdapter adapter, object app, JsonObject state)
    {
        ActivateOwnedWorkbook(adapter, app);
        return adapter.RunOnAdapterThread(() =>
        {
            object? workbooks = null;
            object? workbook = null;
            try
            {
                workbooks = (object)((dynamic)app).Workbooks;
                workbook = FindOwnedWorkbook(workbooks);
                return InvokeRestoreFormatOnlyState(adapter, workbook, state);
            }
            finally
            {
                RotHelper.ReleaseComReference(workbook);
                RotHelper.ReleaseComReference(workbooks);
            }
        });
    }

    private static JsonObject RunPartialWriteInjectedFailure(ExcelAdapter adapter, object app, int scale)
    {
        var written = new JsonArray();
        Exception? thrown = null;
        try
        {
            adapter.RunOnAdapterThread<object?>(() =>
            {
                var sheet = Sheet(app, TargetSheet);
                try
                {
                    for (var i = 0; i < 3 && i < scale; i++)
                    {
                        var address = ScaleAddress(i);
                        SetCellFillRaw(sheet, address, BlueOle);
                        SetCellBold(sheet, address, true);
                        written.Add(address);
                    }
                }
                finally { RotHelper.ReleaseComReference(sheet); }

                throw new InjectedPartialWriteFailureException();
            });
        }
        catch (Exception ex)
        {
            thrown = Unwrap(ex);
        }

        return new JsonObject
        {
            ["ok"] = false,
            ["injected"] = thrown is InjectedPartialWriteFailureException,
            ["name"] = "partial-write-injected-failure",
            ["writtenCells"] = written,
            ["error"] = thrown?.Message,
            ["errorType"] = thrown?.GetType().FullName,
            ["successCase"] = false,
            ["adapterAutomaticErrorPath"] = false,
            ["fixtureInjectedOnly"] = true,
            ["note"] = "sentinel thrown after fixture partial writes; not adapter/automatic error-path proof",
        };
    }

    private static JsonObject RunChecksumCorruption(
        ExcelAdapter adapter, object app, string checkpointPath, string expectedSha,
        List<JsonObject> ops, string corruptPath, ScaleCellScan beforeReject)
    {
        File.Copy(checkpointPath, corruptPath, overwrite: false);
        var bytes = File.ReadAllBytes(corruptPath);
        if (bytes.Length > 64) bytes[64] ^= 0x5A;
        File.WriteAllBytes(corruptPath, bytes);
        var attempt = DeferredExtractAndRestore(adapter, app, corruptPath, expectedSha, ops, "checksum-corruption");
        var after = ReadIdentity(adapter, app);
        var first = ReadValueFormat(adapter, app, TargetSheet, ScaleAddress(0));
        Progress("native-scan", "native-after-checksum-reject");
        var afterReject = ScanScaleCells(adapter, app, scale: beforeReject.Expected, compareTo: beforeReject);
        return new JsonObject
        {
            ["ok"] = Json.GetBool(attempt, "rejectedBeforeMutation") && !Json.GetBool(attempt, "ok") && afterReject.Ok,
            ["rejectedBeforeMutation"] = Json.GetBool(attempt, "rejectedBeforeMutation"),
            ["attempt"] = attempt,
            ["identityAfterReject"] = after,
            ["sampleAfterReject"] = first,
            ["nativeAfterReject"] = NativePhase("native-after-checksum-reject", afterReject),
            ["stylesUnchangedFromPriorScan"] = afterReject.Ok,
            ["corruptPath"] = corruptPath,
        };
    }

    private static JsonObject InvokeCaptureFormatOnlyState(
        ExcelAdapter adapter, object workbook, List<JsonObject> ops, string documentRef)
    {
        var method = typeof(ExcelAdapter).GetMethod(
            "CaptureFormatOnlyState", BindingFlags.Instance | BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("CaptureFormatOnlyState not found for reflection");
        try
        {
            return (JsonObject)(method.Invoke(adapter, [workbook, ops, documentRef])
                ?? throw new InvalidOperationException("CaptureFormatOnlyState returned null"));
        }
        catch (TargetInvocationException ex) when (ex.InnerException is not null)
        {
            throw ex.InnerException;
        }
    }

    private static JsonObject InvokeRestoreFormatOnlyState(
        ExcelAdapter adapter, object workbook, JsonObject state)
    {
        var method = typeof(ExcelAdapter).GetMethod(
            "RestoreFormatOnlyState", BindingFlags.Instance | BindingFlags.NonPublic)
            ?? throw new InvalidOperationException("RestoreFormatOnlyState not found for reflection");
        try
        {
            return (JsonObject)(method.Invoke(adapter, [workbook, state])
                ?? throw new InvalidOperationException("RestoreFormatOnlyState returned null"));
        }
        catch (TargetInvocationException ex) when (ex.InnerException is not null)
        {
            throw ex.InnerException;
        }
    }

    private static void SetCellValueAndNumberFormat(
        ExcelAdapter adapter, object app, string address, string value, string numberFormat)
    {
        adapter.RunOnAdapterThread<object?>(() =>
        {
            var sheet = Sheet(app, TargetSheet);
            object? range = null;
            try
            {
                range = (object)((dynamic)sheet).Range(address);
                ((dynamic)range).Value2 = value;
                ((dynamic)range).NumberFormat = numberFormat;
            }
            finally
            {
                RotHelper.ReleaseComReference(range);
                RotHelper.ReleaseComReference(sheet);
            }

            return null;
        });
    }

    private static JsonObject ReadValueFormat(ExcelAdapter adapter, object app, string sheetName, string address)
    {
        return adapter.RunOnAdapterThread(() =>
        {
            var sheet = Sheet(app, sheetName);
            object? range = null;
            try
            {
                range = (object)((dynamic)sheet).Range(address);
                return new JsonObject
                {
                    ["sheet"] = sheetName,
                    ["address"] = address,
                    ["value2"] = Convert.ToString((object?)((dynamic)range).Value2, CultureInfo.InvariantCulture),
                    ["numberFormat"] = Convert.ToString((object?)((dynamic)range).NumberFormat, CultureInfo.InvariantCulture),
                    ["hasFormula"] = TryBool(() => (object?)((dynamic)range).HasFormula),
                };
            }
            finally
            {
                RotHelper.ReleaseComReference(range);
                RotHelper.ReleaseComReference(sheet);
            }
        });
    }

    private static bool NumberFormatsEquivalent(string? actual, string expected)
    {
        if (string.Equals(actual, expected, StringComparison.Ordinal)) return true;
        var trim = actual?.Trim('"', '\'');
        var want = expected.Trim('"', '\'');
        return string.Equals(trim, want, StringComparison.Ordinal);
    }

    private static string Sha256File(string path) =>
        Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();

    private static JsonObject VerifyRequestedApplyStyles(ScaleCellScan scan)
    {
        var mismatches = new JsonArray();
        var checkedCells = 0;
        foreach (var node in scan.Cells)
        {
            if (node is not JsonObject cell) continue;
            checkedCells++;
            var style = Json.GetObj(cell, "style") ?? new JsonObject();
            var address = Json.GetString(cell, "ref");
            if (style["bold"] is not JsonValue boldNode || !boldNode.TryGetValue<bool>(out var bold) || !bold)
                mismatches.Add($"{address}: requested bold=true not observed");
            if (!MixedFill) continue;
            if (style["fillColor"] is not JsonValue fillNode || !fillNode.TryGetValue<double>(out var fill) ||
                Math.Abs(fill - RedOle) > 0.5d)
                mismatches.Add($"{address}: requested fillColor={RedOle} not observed");
            if (style["fillPattern"] is not JsonValue patternNode || !patternNode.TryGetValue<int>(out var pattern) ||
                pattern != XlPatternSolid)
                mismatches.Add($"{address}: requested solid fillPattern not observed");
        }

        return new JsonObject
        {
            ["ok"] = scan.Unreadable == 0 && checkedCells == scan.Expected && mismatches.Count == 0,
            ["checked"] = checkedCells,
            ["expected"] = scan.Expected,
            ["unreadable"] = scan.Unreadable,
            ["requestedFields"] = MixedFill ? ToJsonArray(["bold", "fillColor", "fillPattern"]) : ToJsonArray(["bold"]),
            ["mismatchCount"] = mismatches.Count,
            ["samples"] = SampleMismatches(mismatches, 8),
        };
    }

    private static JsonArray SampleMismatches(JsonArray mismatches, int max)
    {
        if (mismatches.Count <= max) return mismatches;
        var samples = new JsonArray();
        for (var i = 0; i < max; i++)
        {
            if (mismatches[i] is JsonNode node)
                samples.Add(node.DeepClone());
        }

        return samples;
    }

    private static JsonObject CompareScaleScans(ScaleCellScan actual, ScaleCellScan prior)
    {
        var mismatches = 0;
        var n = Math.Min(actual.Cells.Count, prior.Cells.Count);
        for (var i = 0; i < n; i++)
        {
            if (!ObservationsMatch(prior.Cells[i] as JsonObject, actual.Cells[i] as JsonObject ?? new JsonObject()))
                mismatches++;
        }

        return new JsonObject
        {
            ["comparedToNativeBefore"] = true,
            ["expectedToDifferAfterApply"] = true,
            ["mismatches"] = mismatches,
            ["note"] = "old-state differences after apply are expected; they are not the requested-visible-result assertion",
        };
    }

    private static bool NewerValueFormatPreserved(JsonObject observed) =>
        Json.GetString(observed, "value2") == DeferredValue &&
        NumberFormatsEquivalent(Json.GetString(observed, "numberFormat"), DeferredNumberFormat);

    private static bool KeepAuditsEqual(JsonObject before, JsonObject after)
    {
        if (!AuditReadable(before) || !AuditReadable(after)) return false;
        if (Json.GetString(before, "value") != Json.GetString(after, "value")) return false;
        if (Json.GetString(before, "value") != Marker) return false;
        if (before["hasFormula"] is JsonValue hb && after["hasFormula"] is JsonValue ha &&
            hb.TryGetValue<bool>(out var bHas) && ha.TryGetValue<bool>(out var aHas) && bHas != aHas)
            return false;
        if (!string.Equals(Json.Canonical(before["typedValue"]), Json.Canonical(after["typedValue"]), StringComparison.Ordinal))
            return false;
        if (!string.Equals(Json.Canonical(before["formula"]), Json.Canonical(after["formula"]), StringComparison.Ordinal))
            return false;
        if (before["fontBold"] is JsonValue bb && after["fontBold"] is JsonValue ab &&
            bb.TryGetValue<bool>(out var bBold) && ab.TryGetValue<bool>(out var aBold) && bBold != aBold)
            return false;
        if (!StyleFieldEqual(before["fillColor"], after["fillColor"], "fillColor")) return false;
        if (!StyleFieldEqual(before["fillPattern"], after["fillPattern"], "fillPattern")) return false;
        return true;
    }

    private static JsonObject ForegroundCompare(JsonObject before, JsonObject after)
    {
        var measured = Json.GetBool(before, "measured") && Json.GetBool(after, "measured");
        var result = new JsonObject
        {
            ["before"] = before.DeepClone(),
            ["after"] = after.DeepClone(),
            ["measured"] = measured,
            ["claimedWithoutEvidence"] = false,
            ["note"] = measured
                ? "read-only GetForegroundWindow before/after; not ForegroundInteractionGuard.Complete"
                : "foreground unmeasured; not claimed preserved",
        };
        if (measured)
            result["preserved"] = Json.GetLong(before, "hwnd") == Json.GetLong(after, "hwnd");
        return result;
    }

    private static Exception Unwrap(Exception ex)
    {
        while (ex is TargetInvocationException { InnerException: { } inner })
            ex = inner;
        return ex;
    }

}

internal sealed class InjectedPartialWriteFailureException : Exception
{
    public InjectedPartialWriteFailureException()
        : base("injected failure after partial native format write")
    {
    }
}
