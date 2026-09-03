using System.Diagnostics;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Models;

namespace DocBridge.Core.Services;

/// <summary>
/// tool 호출 orchestration. MCP 서버와 CLI가 공유하는 단일 진입점.
///
/// 쓰기 안전 순서 (명령서 §0, §3):
///   dry-run(preview + diff) → snapshot → confirmToken 발급
///   → (사용자 승인) → confirmToken 검증/소비 → apply → readback verify → audit
///
/// COM 자동화 호출은 named mutex(Global\DocBridge.Automation)로
/// 크로스 프로세스 직렬화한다 (Claude/Codex 동시 stdio 서버 대비).
/// </summary>
public sealed class DocBridgeHost : IDisposable
{
    public const string Version = "0.4.19";
    private const string AutomationMutex = @"Global\DocBridge.Automation";

    private readonly DocBridgeOptions _options;
    private readonly PolicyEngine _policy;
    private readonly OperationValidator _validator;
    private readonly ConfirmTokenService _tokens;
    private readonly SnapshotService _snapshots;
    private readonly AuditLog _audit;
    private readonly SessionRouter _router;

    public DocBridgeHost(DocBridgeOptions? options = null)
    {
        _options = options ?? new DocBridgeOptions();
        _options.EnsureDirectories();
        _policy = new PolicyEngine(_options.PolicyPath);
        _validator = new OperationValidator(_policy);
        ConfirmTokenService.WarmUpCrypto();
        _tokens = new ConfirmTokenService(_options, _policy.TokenTtlSeconds);
        _snapshots = new SnapshotService(_options);
        _audit = new AuditLog(_options);
        _router = new SessionRouter();
    }

    public PolicyEngine Policy => _policy;
    public SessionRouter Router => _router;
    public AuditLog Audit => _audit;

    /// <summary>크로스 프로세스 자동화 직렬화</summary>
    private T WithAutomationLock<T>(Func<T> work, Action<long>? lockTiming = null)
    {
        using var mutex = new Mutex(false, AutomationMutex);
        var acquired = false;
        var wait = Stopwatch.StartNew();
        try
        {
            acquired = mutex.WaitOne(TimeSpan.FromSeconds(60));
            wait.Stop();
            lockTiming?.Invoke(wait.ElapsedMilliseconds);
            if (!acquired)
                throw new TimeoutException("another doc-bridge process is busy (automation lock timeout)");
            return work();
        }
        finally
        {
            if (acquired) mutex.ReleaseMutex();
        }
    }

    // ---------- core tools ----------

    public JsonObject CorePing()
    {
        var adapters = new JsonArray();
        foreach (var app in _router.Apps) adapters.Add(app);
        return new JsonObject
        {
            ["ok"] = true,
            ["version"] = Version,
            ["adapters"] = adapters,
        };
    }

    public JsonObject CoreGetStatus()
    {
        var apps = new JsonObject();
        foreach (var app in _router.Apps)
        {
            try
            {
                var st = _router.Get(app).GetStatus();
                apps[app] = new JsonObject
                {
                    ["available"] = st.Available,
                    ["connected"] = st.Connected,
                    ["program"] = st.Program,
                    ["version"] = st.Version,
                    ["document"] = st.Document,
                    ["detail"] = st.Detail,
                };
            }
            catch (Exception ex)
            {
                var error = new JsonObject { ["available"] = false, ["error"] = ex.Message };
                AddHwpAutomationError(error, app, ex);
                apps[app] = error;
            }
        }
        return new JsonObject { ["ok"] = true, ["version"] = Version, ["apps"] = apps };
    }

    public JsonObject CoreDisconnect(JsonObject args)
    {
        var app = Json.GetString(args, "app");
        if (string.IsNullOrWhiteSpace(app))
            return Json.ErrorResult("core_disconnect requires 'app'");

        try
        {
            return WithAutomationLock(() =>
            {
                var adapter = _router.Get(app);
                if (adapter is not IConnectionLifecycleAdapter lifecycle)
                    return Json.ErrorResult($"{app} adapter does not support explicit disconnect", app);
                var result = lifecycle.Disconnect();
                _audit.Write("core_disconnect", app, "disconnect", result.DeepClone() as JsonObject,
                    Json.GetBool(result, "ok"), Json.GetBool(result, "ok") ? null : new[] { Json.GetString(result, "error") ?? "disconnect failed" });
                return result;
            });
        }
        catch (Exception ex)
        {
            _audit.Write("core_disconnect", app, "disconnect", null, false, new[] { ex.Message });
            return Json.ErrorResult($"disconnect failed: {ex.Message}", app);
        }
    }

    public JsonObject CoreGetCapabilities(JsonObject? args)
    {
        var requested = Json.GetString(args, "app");
        var names = string.IsNullOrWhiteSpace(requested)
            ? _router.Apps.Where(name => !name.Equals("fake", StringComparison.OrdinalIgnoreCase)).ToArray()
            : _router.Apps.Where(name => name.Equals(requested, StringComparison.OrdinalIgnoreCase)).ToArray();
        if (names.Length == 0)
            return Json.ErrorResult($"unknown app '{requested}'. supported: excel, hwp, cad");

        var apps = new JsonObject();
        foreach (var app in names)
        {
            try
            {
                var adapter = _router.Get(app);
                var capability = adapter.GetCapabilities();
                var status = adapter.GetStatus();
                capability["registeredApp"] = app;
                capability["available"] = status.Available;
                capability["connected"] = status.Connected;
                capability["programVersion"] = status.Version;
                capability["documentRef"] = status.Document;
                apps[app] = capability;
            }
            catch (Exception ex)
            {
                var error = new JsonObject
                {
                    ["app"] = app,
                    ["available"] = false,
                    ["connected"] = false,
                    ["error"] = ex.Message,
                };
                AddHwpAutomationError(error, app, ex);
                apps[app] = error;
            }
        }
        return new JsonObject { ["ok"] = true, ["version"] = Version, ["apps"] = apps };
    }

    public JsonObject CoreCreateSnapshot(JsonObject args)
    {
        var app = Json.GetString(args, "app");
        var reason = Json.GetString(args, "reason") ?? "manual";
        if (string.IsNullOrWhiteSpace(app))
            return Json.ErrorResult("core_create_snapshot requires 'app'");

        try
        {
            return WithAutomationLock(() =>
            {
                var adapter = _router.Get(app);
                var info = _snapshots.Create(app, reason, CurrentDocRef(adapter),
                    (dir, meta) => adapter.CaptureSnapshot(dir, meta));
                _audit.Write("core_create_snapshot", app, "snapshot",
                    new JsonObject { ["snapshotId"] = info.SnapshotId, ["reason"] = reason });
                return new JsonObject
                {
                    ["ok"] = true,
                    ["snapshotId"] = info.SnapshotId,
                    ["createdAt"] = info.CreatedAt,
                    ["app"] = info.App,
                    ["documentRef"] = info.DocumentRef,
                };
            });
        }
        catch (Exception ex)
        {
            _audit.Write("core_create_snapshot", app, "snapshot", null, false, new[] { ex.Message });
            return Json.ErrorResult($"snapshot failed: {ex.Message}", app);
        }
    }

    public JsonObject CoreListSnapshots(JsonObject? args)
    {
        var app = Json.GetString(args, "app");
        var limit = Json.GetInt(args, "limit") ?? 20;
        var list = _snapshots.List(app, limit);
        var arr = new JsonArray();
        foreach (var s in list)
            arr.Add(new JsonObject
            {
                ["snapshotId"] = s.SnapshotId,
                ["createdAt"] = s.CreatedAt,
                ["app"] = s.App,
                ["documentRef"] = s.DocumentRef,
                ["reason"] = s.Reason,
            });
        return new JsonObject { ["ok"] = true, ["snapshots"] = arr, ["count"] = arr.Count };
    }

    /// <summary>고위험 tool: 스냅샷 복원. dry-run → confirmToken 흐름을 그대로 따른다.</summary>
    public JsonObject CoreRestoreSnapshot(JsonObject args)
    {
        var snapshotId = Json.GetString(args, "snapshotId");
        if (string.IsNullOrWhiteSpace(snapshotId))
            return Json.ErrorResult("core_restore_snapshot requires 'snapshotId'");

        var found = _snapshots.Get(snapshotId);
        if (found is null)
            return Json.ErrorResult($"snapshot '{snapshotId}' not found");
        var (info, metadata) = found.Value;

        // restore 의도를 ops 해시 대신 snapshotId 해시로 바인딩한다
        var scope = $"restore:{info.App}";
        var restoreHash = ConfirmTokenService.HashScope($"restore:{snapshotId}");

        var confirmToken = Json.GetString(args, "confirmToken");
        if (confirmToken is null)
        {
            // 1단계: dry-run (토큰 발급)
            var (token, expires) = _tokens.Create(scope, restoreHash, snapshotId);
            _audit.Write("core_restore_snapshot", info.App, "dry_run",
                new JsonObject { ["snapshotId"] = snapshotId });
            return new JsonObject
            {
                ["ok"] = true,
                ["dryRun"] = true,
                ["highRisk"] = true,
                ["snapshotId"] = snapshotId,
                ["app"] = info.App,
                ["documentRef"] = info.DocumentRef,
                ["confirmToken"] = token,
                ["expiresInSec"] = expires,
                ["warnings"] = Json.ToArray(new[]
                {
                    "복원은 현재 문서 상태를 스냅샷 시점으로 되돌립니다.",
                    "적용하려면 같은 snapshotId와 이 confirmToken으로 다시 호출하세요.",
                }),
            };
        }

        // 2단계: 실제 복원
        var check = _tokens.ValidateAndConsume(confirmToken, scope, restoreHash);
        if (!check.Ok)
        {
            _audit.Write("core_restore_snapshot", info.App, "deny",
                new JsonObject { ["snapshotId"] = snapshotId }, false, new[] { check.Error! });
            return Json.ErrorResult(check.Error!, info.App);
        }

        try
        {
            return WithAutomationLock(() =>
            {
                var adapter = _router.Get(info.App);
                var result = adapter.RestoreSnapshot(info.Dir, metadata);
                var ok = Json.GetBool(result, "ok");
                _audit.Write("core_restore_snapshot", info.App, "restore",
                    new JsonObject { ["snapshotId"] = snapshotId, ["result"] = result.DeepClone() },
                    ok, ok ? null : new[] { "restore reported failure" });
                result["snapshotId"] = snapshotId;
                result["dryRun"] = false;
                return result;
            });
        }
        catch (Exception ex)
        {
            _audit.Write("core_restore_snapshot", info.App, "restore",
                new JsonObject { ["snapshotId"] = snapshotId }, false, new[] { ex.Message });
            return Json.ErrorResult($"restore failed: {ex.Message}", info.App);
        }
    }

    // ---------- context/read ----------

    public JsonObject HwpPlanCreation(JsonObject? args)
    {
        const string tool = "hwp_plan_creation";
        try
        {
            var result = HwpCreationPolicy.Evaluate(args);
            _audit.Write(tool, "hwp", "plan", result.DeepClone() as JsonObject,
                Json.GetBool(result, "ok"),
                Json.GetBool(result, "ok") ? null : new[] { Json.GetString(result, "error") ?? "planning failed" });
            return result;
        }
        catch (Exception ex)
        {
            _audit.Write(tool, "hwp", "plan", null, false, new[] { ex.Message });
            return Json.ErrorResult($"hwp creation planning failed: {ex.Message}", "hwp");
        }
    }

    public JsonObject HwpLaunch(JsonObject? args)
    {
        const string tool = "hwp_launch";
        try
        {
            return WithAutomationLock(() =>
            {
                if (_router.Get("hwp") is not IHwpAutomationAdapter adapter)
                    return Json.ErrorResult("hwp adapter does not support launching", "hwp");
                var result = adapter.Launch(args ?? new JsonObject());
                _audit.Write(tool, "hwp", "launch", new JsonObject
                {
                    ["documentRef"] = Json.GetString(result, "documentRef"),
                    ["createdDocument"] = Json.GetBool(Json.GetObj(result, "summary"), "createdDocument"),
                }, Json.GetBool(result, "ok"),
                    Json.GetArr(result, "errors")?.Select(n => n?.GetValue<string>() ?? "") ?? Array.Empty<string>());
                return result;
            });
        }
        catch (Exception ex)
        {
            _audit.Write(tool, "hwp", "launch", null, false, new[] { ex.Message });
            return AutomationErrorResult("hwp launch failed", "hwp", ex);
        }
    }

    public JsonObject HwpDoctor(JsonObject? args)
    {
        const string tool = "hwp_doctor";
        try
        {
            if (_router.Get("hwp") is not IHwpAutomationAdapter adapter)
                return Json.ErrorResult("hwp adapter does not support environment diagnostics", "hwp");
            var result = adapter.Doctor(args ?? new JsonObject());
            _audit.Write(tool, "hwp", "diagnose", result.DeepClone() as JsonObject, Json.GetBool(result, "ok"));
            return result;
        }
        catch (Exception ex)
        {
            _audit.Write(tool, "hwp", "diagnose", null, false, new[] { ex.Message });
            return AutomationErrorResult("hwp doctor failed", "hwp", ex);
        }
    }

    public JsonObject HwpRepairTypeLib(JsonObject args)
    {
        const string tool = "hwp_repair_typelib";
        if (!Json.GetBool(args, "confirm"))
            return Json.ErrorResult("hwp_repair_typelib requires confirm=true after explicit user approval", "hwp");
        try
        {
            return WithAutomationLock(() =>
            {
                if (_router.Get("hwp") is not IHwpAutomationAdapter adapter)
                    return Json.ErrorResult("hwp adapter does not support TypeLib repair", "hwp");
                var result = adapter.RepairTypeLib(args);
                _audit.Write(tool, "hwp", "repair", result.DeepClone() as JsonObject, Json.GetBool(result, "ok"),
                    Json.GetArr(result, "errors")?.Select(node => node?.GetValue<string>() ?? "") ?? Array.Empty<string>());
                return result;
            });
        }
        catch (Exception ex)
        {
            _audit.Write(tool, "hwp", "repair", null, false, new[] { ex.Message });
            return Json.ErrorResult($"hwp TypeLib repair failed: {ex.Message}", "hwp");
        }
    }

    public JsonObject CadLaunch(JsonObject? args)
    {
        const string tool = "cad_launch";
        try
        {
            return WithAutomationLock(() =>
            {
                if (_router.Get("cad") is not CadAdapter adapter)
                    return Json.ErrorResult("cad adapter does not support launching", "cad");
                var result = adapter.Launch(args ?? new JsonObject());
                _audit.Write(tool, "cad", "launch", new JsonObject
                {
                    ["documentRef"] = Json.GetString(result, "documentRef"),
                }, Json.GetBool(result, "ok"),
                    Json.GetArr(result, "errors")?.Select(n => n?.GetValue<string>() ?? "") ?? Array.Empty<string>());
                return result;
            });
        }
        catch (Exception ex)
        {
            _audit.Write(tool, "cad", "launch", null, false, new[] { ex.Message });
            return Json.ErrorResult($"cad launch failed: {ex.Message}", "cad");
        }
    }

    public JsonObject GetActiveContext(string app, JsonObject? args = null)
    {
        try
        {
            return WithAutomationLock(() =>
            {
                var adapter = _router.Get(app);
                var ctx = app.Equals("cad", StringComparison.OrdinalIgnoreCase) && adapter is CadAdapter cad
                    ? cad.GetActiveContext(args)
                    : adapter.GetActiveContext();
                _audit.Write($"{app}_get_active_context", app, "read",
                    new JsonObject
                    {
                        ["documentRef"] = ctx.DocumentRef,
                        ["detailLevel"] = Json.GetString(args, "detailLevel"),
                    }, ctx.Ok, ctx.Errors);
                return ctx.ToJson();
            });
        }
        catch (Exception ex)
        {
            _audit.Write($"{app}_get_active_context", app, "read", null, false, new[] { ex.Message });
            return AutomationErrorResult("get_active_context failed", app, ex);
        }
    }

    public JsonObject Read(string app, JsonObject? args)
    {
        try
        {
            return WithAutomationLock(() =>
            {
                var adapter = _router.Get(app);
                var result = adapter.Read(args ?? new JsonObject());
                _audit.Write($"{app}_read", app, "read", null, Json.GetBool(result, "ok"));
                return result;
            });
        }
        catch (Exception ex)
        {
            _audit.Write($"{app}_read", app, "read", null, false, new[] { ex.Message });
            return AutomationErrorResult("read failed", app, ex);
        }
    }

    // ---------- apply (핵심 안전 흐름) ----------

    public JsonObject ApplyOps(string app, JsonObject? batch)
    {
        var tool = $"{app}_apply_ops";
        var totalStarted = Stopwatch.StartNew();
        var timings = new JsonObject();
        JsonObject Finish(JsonObject result)
        {
            totalStarted.Stop();
            timings["totalMs"] = totalStarted.ElapsedMilliseconds;
            result["timings"] = timings.DeepClone();
            return result;
        }

        JsonObject AuditTimings()
        {
            // Audit entries are written before Finish() returns the response.  Capture the
            // elapsed wall-clock time at the exact audit boundary so production telemetry
            // contains a useful total without logging document contents.
            timings["totalMs"] = totalStarted.ElapsedMilliseconds;
            return (JsonObject)timings.DeepClone();
        }

        // 1) 구조/정책 검증
        var validationStarted = Stopwatch.StartNew();
        var errors = new List<string>();
        var parsed = _validator.Validate(batch, app, errors);
        validationStarted.Stop();
        timings["validationMs"] = validationStarted.ElapsedMilliseconds;
        if (parsed is null)
        {
            _audit.Write(tool, app, "deny", null, false, errors);
            var deny = new JsonObject { ["ok"] = false, ["dryRun"] = Json.GetBool(batch, "dryRun", true) };
            deny["errors"] = Json.ToArray(errors);
            var expectedSchemas = _validator.DescribeExpectedSchemas(batch, app);
            if (expectedSchemas.Count > 0) deny["expectedSchema"] = expectedSchemas;
            return Finish(deny);
        }

        var opsHash = ConfirmTokenService.HashOps(parsed.Ops);
        var scope = $"apply:{app}";

        try
        {
            var result = WithAutomationLock(() =>
            {
                var adapter = _router.Get(app);
                var statusStarted = Stopwatch.StartNew();
                var status = adapter.GetStatus();
                statusStarted.Stop();
                timings["statusMs"] = statusStarted.ElapsedMilliseconds;
                if (!status.Available)
                {
                    var msg = $"{app} program not available: {status.Detail ?? "not detected"}";
                    _audit.Write(tool, app, "deny", null, false, new[] { msg });
                    return Json.ErrorResult(msg, app);
                }
                // 한 호출 안에서는 방금 읽은 상태의 문서 식별자를 재사용한다. TargetDocumentRef가
                // GetStatus()를 다시 부르지 않게 해 COM 왕복을 줄이되, dry-run과 apply 사이에는
                // 여전히 각각 새 상태를 읽어 문서 전환을 놓치지 않는다.
                var targetDocument = TargetDocumentRef(app, parsed.Ops, status.Document);

                if (parsed.DryRun)
                {
                    // 2) 동일 문서·동일 ops의 반복 dry-run이면 최신 스냅샷 후보를 찾는다.
                    // 후보 메타데이터만으로 신뢰하지 않고, 전체 문서 fingerprint가 현재 상태와
                    // 일치할 때만 preview와 snapshot을 함께 재사용한다.
                    (SnapshotInfo Info, JsonObject Metadata)? reuseCandidate = null;
                    ApplyPreview? dryPreview = null;
                    var cacheLookupStarted = Stopwatch.StartNew();
                    if (adapter is IPreviewReuseAdapter dryRunReusableAdapter)
                    {
                        reuseCandidate = _snapshots.FindLatestReusableCandidate(
                            app, targetDocument, opsHash,
                            (expected, current) => SameDocumentRef(app, expected, current));
                        if (reuseCandidate is not null)
                        {
                            var reuseValidationStarted = Stopwatch.StartNew();
                            try
                            {
                                var validation = dryRunReusableAdapter.ValidatePreviewReuse(
                                    reuseCandidate.Value.Info.Dir,
                                    reuseCandidate.Value.Metadata,
                                    parsed.Ops);
                                timings["dryRunFingerprintMethod"] = Json.GetString(validation, "fingerprintMethod");
                                if (Json.GetBool(validation, "reusable"))
                                    dryPreview = ApplyPreviewArtifact.FromMetadata(
                                        reuseCandidate.Value.Metadata, opsHash);
                                else
                                    timings["dryRunCacheMissReason"] =
                                        Json.GetString(validation, "reason") ?? "document fingerprint changed";
                            }
                            catch (Exception ex)
                            {
                                // 캐시 검증 실패는 안전하게 일반 preview+새 snapshot 경로로 내린다.
                                timings["dryRunCacheMissReason"] = $"fingerprint validation failed: {ex.Message}";
                            }
                            finally
                            {
                                reuseValidationStarted.Stop();
                                timings["dryRunFingerprintValidationMs"] = reuseValidationStarted.ElapsedMilliseconds;
                            }
                        }
                    }
                    cacheLookupStarted.Stop();
                    timings["dryRunCacheLookupMs"] = cacheLookupStarted.ElapsedMilliseconds;

                    if (dryPreview is null)
                    {
                        var previewStarted = Stopwatch.StartNew();
                        dryPreview = adapter.Preview(parsed.Ops);
                        previewStarted.Stop();
                        timings["previewMs"] = previewStarted.ElapsedMilliseconds;
                        timings["previewReused"] = false;
                        timings["previewCacheHit"] = false;
                        reuseCandidate = null;
                    }
                    else
                    {
                        timings["previewMs"] = 0L;
                        timings["previewReused"] = true;
                        timings["previewCacheHit"] = true;
                    }
                    AddDistinctWarnings(dryPreview.Warnings, parsed.OptimizationWarnings);
                    if (dryPreview.Errors.Count > 0)
                    {
                        _audit.Write(tool, app, "deny", null, false, dryPreview.Errors);
                        var pe = new JsonObject { ["ok"] = false, ["dryRun"] = true };
                        pe["errors"] = Json.ToArray(dryPreview.Errors);
                        pe["warnings"] = Json.ToArray(dryPreview.Warnings);
                        pe["interaction"] = dryPreview.Interaction?.DeepClone();
                        return pe;
                    }

                    // 3) snapshot 생성/재사용 + 4) confirmToken 발급
                    var snapshotStarted = Stopwatch.StartNew();
                    SnapshotInfo info;
                    if (reuseCandidate is not null)
                    {
                        info = reuseCandidate.Value.Info;
                        timings["snapshotReused"] = true;
                    }
                    else
                    {
                        info = _snapshots.Create(app, $"{tool} dry-run", targetDocument,
                            (dir, meta) =>
                            {
                                adapter.CaptureSnapshot(dir, meta, parsed.Ops);
                                meta["snapshotReuseVersion"] = 1;
                                ApplyPreviewArtifact.StoreInMetadata(meta, opsHash, dryPreview);
                            });
                        timings["snapshotReused"] = false;
                    }
                    snapshotStarted.Stop();
                    timings["snapshotMs"] = reuseCandidate is null
                        ? snapshotStarted.ElapsedMilliseconds
                        : 0;
                    var tokenStarted = Stopwatch.StartNew();
                    var (token, expires) = _tokens.Create(scope, opsHash, info.SnapshotId);
                    tokenStarted.Stop();
                    timings["tokenMs"] = tokenStarted.ElapsedMilliseconds;
                    _audit.Write(tool, app, "dry_run", new JsonObject
                    {
                        ["snapshotId"] = info.SnapshotId,
                        ["ops"] = OpsSummary(parsed.Ops),
                        ["diffCount"] = dryPreview.Diff.Count,
                        ["timings"] = AuditTimings(),
                    });

                    return new JsonObject
                    {
                        ["ok"] = true,
                        ["dryRun"] = true,
                        ["snapshotId"] = info.SnapshotId,
                        ["confirmToken"] = token,
                        ["expiresInSec"] = expires,
                        ["requiresHighRiskApproval"] = dryPreview.RequiresHighRiskApproval || parsed.HasHighRiskOps,
                        ["affected"] = Json.ToArray(dryPreview.Affected),
                        ["diff"] = Json.ToArray(dryPreview.Diff),
                        ["diffTruncated"] = dryPreview.DiffTruncated,
                        ["warnings"] = Json.ToArray(dryPreview.Warnings),
                        ["errors"] = Json.ToArray(Array.Empty<string>()),
                        ["interaction"] = dryPreview.Interaction?.DeepClone(),
                    };
                }

                // 5) 실제 적용: confirmToken 필수 (보안 원칙 3)
                if (string.IsNullOrWhiteSpace(parsed.ConfirmToken))
                {
                    var msg = "dryRun=false requires confirmToken from a prior dry-run call (security rule 3)";
                    _audit.Write(tool, app, "deny", null, false, new[] { msg });
                    return Json.ErrorResult(msg, app);
                }

                // 토큰을 먼저 소비하지 않고 검증해 dry-run 스냅샷과 preview artifact를 찾는다.
                // 고위험 승인/문서 fingerprint 검증이 실패하면 토큰은 여전히 재사용 가능하다.
                var tokenValidationStarted = Stopwatch.StartNew();
                var peek = _tokens.Validate(parsed.ConfirmToken, scope, opsHash);
                tokenValidationStarted.Stop();
                timings["tokenValidationMs"] = tokenValidationStarted.ElapsedMilliseconds;
                if (!peek.Ok)
                {
                    _audit.Write(tool, app, "deny", null, false, new[] { peek.Error! });
                    return Json.ErrorResult(peek.Error!, app);
                }

                var snapshotLookupStarted = Stopwatch.StartNew();
                var snapshot = peek.SnapshotId is null ? null : _snapshots.Get(peek.SnapshotId);
                snapshotLookupStarted.Stop();
                timings["snapshotLookupMs"] = snapshotLookupStarted.ElapsedMilliseconds;
                if (snapshot is null)
                {
                    const string msg = "confirmToken snapshot not found; run dry-run again";
                    _audit.Write(tool, app, "deny", null, false, new[] { msg });
                    return Json.ErrorResult(msg, app);
                }

                // 토큰을 발급한 dry-run의 문서와 현재 적용 대상이 같아야 한다.
                // ops만 바인딩하면 사용자가 중간에 활성 workbook/도면을 바꿨을 때
                // 같은 토큰이 다른 문서에 적용될 수 있다.
                var identityStarted = Stopwatch.StartNew();
                var expectedDocument = snapshot.Value.Info.DocumentRef;
                var currentDocument = targetDocument;
                identityStarted.Stop();
                timings["documentIdentityMs"] = identityStarted.ElapsedMilliseconds;
                if (!SameDocumentRef(app, expectedDocument, currentDocument))
                {
                    var msg = $"document changed after dry-run: expected '{expectedDocument}', current '{currentDocument}'. run dry-run again";
                    _audit.Write(tool, app, "deny", null, false, new[] { msg });
                    return Json.ErrorResult(msg, app);
                }

                ApplyPreview? preview = null;
                var cachedPreview = ApplyPreviewArtifact.FromMetadata(snapshot.Value.Metadata, opsHash);
                if (cachedPreview is not null && adapter is IPreviewReuseAdapter reusableAdapter)
                {
                    var fingerprintStarted = Stopwatch.StartNew();
                    var validation = reusableAdapter.ValidatePreviewReuse(
                        snapshot.Value.Info.Dir, snapshot.Value.Metadata, parsed.Ops);
                    fingerprintStarted.Stop();
                    timings["fingerprintValidationMs"] = fingerprintStarted.ElapsedMilliseconds;
                    timings["fingerprintMethod"] = Json.GetString(validation, "fingerprintMethod");
                    if (!Json.GetBool(validation, "reusable"))
                    {
                        var reason = Json.GetString(validation, "reason") ?? "document fingerprint changed";
                        var msg = $"document changed after dry-run ({reason}); run dry-run again";
                        _audit.Write(tool, app, "deny", validation, false, new[] { msg });
                        return Json.ErrorResult(msg, app);
                    }
                    preview = cachedPreview;
                    timings["previewReused"] = true;
                    timings["previewMs"] = 0L;
                }
                else
                {
                    var previewStarted = Stopwatch.StartNew();
                    preview = adapter.Preview(parsed.Ops);
                    previewStarted.Stop();
                    timings["previewMs"] = previewStarted.ElapsedMilliseconds;
                    timings["previewReused"] = false;
                    timings["previewReuseReason"] = cachedPreview is null
                        ? "dry-run artifact unavailable"
                        : "adapter does not support full fingerprint validation";
                }

                AddDistinctWarnings(preview.Warnings, parsed.OptimizationWarnings);

                if (preview.Errors.Count > 0)
                {
                    _audit.Write(tool, app, "deny", null, false, preview.Errors);
                    var pe = new JsonObject { ["ok"] = false, ["dryRun"] = false };
                    pe["errors"] = Json.ToArray(preview.Errors);
                    pe["warnings"] = Json.ToArray(preview.Warnings);
                    pe["interaction"] = preview.Interaction?.DeepClone();
                    return pe;
                }

                // 고위험 op 추가 승인 (보안 원칙 9)
                if ((preview.RequiresHighRiskApproval || parsed.HasHighRiskOps) && !parsed.HighRiskConfirm)
                {
                    var msg = "batch contains high-risk ops; set highRiskConfirm=true together with confirmToken (security rule 9)";
                    _audit.Write(tool, app, "deny", null, false, new[] { msg });
                    return Json.ErrorResult(msg, app);
                }

                var tokenConsumeStarted = Stopwatch.StartNew();
                var check = _tokens.ValidateAndConsume(parsed.ConfirmToken, scope, opsHash);
                tokenConsumeStarted.Stop();
                timings["tokenConsumeMs"] = tokenConsumeStarted.ElapsedMilliseconds;
                if (!check.Ok)
                {
                    _audit.Write(tool, app, "deny", null, false, new[] { check.Error! });
                    return Json.ErrorResult(check.Error!, app);
                }

                // 6) apply + 7) readback verify. A failed or thrown batch is restored
                // from the exact pre-apply snapshot without requiring another approval.
                var applyStarted = Stopwatch.StartNew();
                ApplyExecution exec;
                HwpAutomationException? hwpApplyError = null;
                try
                {
                    exec = adapter.Apply(parsed.Ops, check.SnapshotId ?? "");
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
                        var restored = adapter.RestoreSnapshot(snapshot.Value.Info.Dir, snapshot.Value.Metadata);
                        rollback["result"] = restored.DeepClone();
                        rollback["verified"] = Json.GetBool(restored, "ok") &&
                            (Json.GetObj(restored, "readback") is not { } rb || Json.GetBool(rb, "verified", true));
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
                _audit.Write(tool, app, "apply", new JsonObject
                {
                    ["snapshotId"] = check.SnapshotId,
                    ["ops"] = OpsSummary(parsed.Ops),
                    ["readbackOk"] = exec.Readback is not null && Json.GetBool(exec.Readback, "verified"),
                    ["elapsedMs"] = applyStarted.ElapsedMilliseconds,
                    ["rollback"] = rollback.DeepClone(),
                    ["operationResults"] = OperationTimingSummary(exec.OperationResults),
                    ["timings"] = AuditTimings(),
                }, exec.Ok, exec.Errors);

                var applyResult = new JsonObject
                {
                    ["ok"] = exec.Ok,
                    ["dryRun"] = false,
                    ["snapshotId"] = check.SnapshotId,
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
                return applyResult;
            }, elapsed => timings["lockWaitMs"] = elapsed);
            return Finish(result);
        }
        catch (Exception ex)
        {
            _audit.Write(tool, app, "apply", null, false, new[] { ex.Message });
            return Finish(AutomationErrorResult("apply failed", app, ex));
        }
    }

    private static JsonObject AutomationErrorResult(string prefix, string app, Exception error)
    {
        var hwp = app.Equals("hwp", StringComparison.OrdinalIgnoreCase)
            ? FindHwpAutomationException(error)
            : null;
        if (hwp is null) return Json.ErrorResult($"{prefix}: {error.Message}", app);
        var result = hwp.ToResult(app);
        result["operation"] = prefix;
        return result;
    }

    private static void AddHwpAutomationError(JsonObject result, string app, Exception error)
    {
        if (!app.Equals("hwp", StringComparison.OrdinalIgnoreCase)) return;
        var hwp = FindHwpAutomationException(error);
        if (hwp is null) return;
        result["errorCode"] = hwp.Code;
        var structured = hwp.ToResult(app);
        result["retryable"] = structured["retryable"]?.DeepClone();
        result["automaticRetry"] = structured["automaticRetry"]?.DeepClone();
        result["retryPolicy"] = structured["retryPolicy"]?.DeepClone();
        if (!string.IsNullOrWhiteSpace(hwp.UserAction)) result["userAction"] = hwp.UserAction;
        if (hwp.RetryAfterMs is not null) result["retryAfterMs"] = hwp.RetryAfterMs.Value;
    }

    private static HwpAutomationException? FindHwpAutomationException(Exception error)
    {
        if (error is HwpAutomationException hwp) return hwp;
        if (error is AggregateException aggregate)
            foreach (var inner in aggregate.Flatten().InnerExceptions)
                if (FindHwpAutomationException(inner) is { } nested) return nested;
        return error.InnerException is null ? null : FindHwpAutomationException(error.InnerException);
    }

    private static JsonArray OpsSummary(IEnumerable<JsonObject> ops)
    {
        var a = new JsonArray();
        foreach (var op in ops) a.Add(Json.GetString(op, "op") ?? "?");
        return a;
    }

    private static void AddDistinctWarnings(List<string> target, IEnumerable<string> warnings)
    {
        foreach (var warning in warnings)
            if (!target.Contains(warning, StringComparer.Ordinal))
                target.Add(warning);
    }

    private static JsonArray OperationTimingSummary(JsonArray operationResults)
    {
        var summary = new JsonArray();
        foreach (var node in operationResults)
        {
            if (node is not JsonObject operation) continue;
            summary.Add(new JsonObject
            {
                ["index"] = Json.GetInt(operation, "index"),
                ["op"] = Json.GetString(operation, "op") ?? "?",
                ["ok"] = Json.GetBool(operation, "ok"),
                ["elapsedMs"] = Json.GetLong(operation, "elapsedMs"),
                ["timingScope"] = Json.GetString(operation, "timingScope"),
            });
        }
        return summary;
    }

    private static string? CurrentDocRef(IAppAdapter adapter)
    {
        try { return adapter.GetStatus().Document; }
        catch { return null; }
    }

    private static string? TargetDocumentRef(
        string app,
        IReadOnlyList<JsonObject> ops,
        string? statusDocument = null)
    {
        if (string.Equals(app, "hwp", StringComparison.OrdinalIgnoreCase))
        {
            var documentRef = ops.Select(op => Json.GetString(op, "documentRef"))
                .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value));
            if (documentRef is not null) return documentRef;
            var file = ops.Select(op => Json.GetString(op, "file"))
                .FirstOrDefault(path => !string.IsNullOrWhiteSpace(path));
            if (file is not null) return Path.GetFullPath(file);
        }
        if (string.Equals(app, "excel", StringComparison.OrdinalIgnoreCase))
        {
            var workbook = ops.Select(op => Json.GetString(op, "targetWorkbook") ?? Json.GetString(Json.GetObj(op, "target"), "workbook"))
                .FirstOrDefault(path => !string.IsNullOrWhiteSpace(path));
            if (workbook is not null && Path.IsPathFullyQualified(workbook)) return Path.GetFullPath(workbook);
        }
        return statusDocument;
    }

    internal static bool SameDocumentRef(string app, string? expected, string? current)
    {
        if (string.IsNullOrWhiteSpace(expected) || string.IsNullOrWhiteSpace(current))
            return string.Equals(expected ?? "", current ?? "", StringComparison.OrdinalIgnoreCase);
        if (string.Equals(app, "hwp", StringComparison.OrdinalIgnoreCase) &&
            TryParseHwpTransientRef(expected, out var expectedProcess, out var expectedDocument) &&
            TryParseHwpTransientRef(current, out var currentProcess, out var currentDocument))
            return expectedProcess == currentProcess && expectedDocument == currentDocument;
        try
        {
            if (Path.IsPathFullyQualified(expected) || Path.IsPathFullyQualified(current))
                return string.Equals(Path.GetFullPath(expected), Path.GetFullPath(current), StringComparison.OrdinalIgnoreCase);
        }
        catch { return false; }
        return string.Equals(expected, current, StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryParseHwpTransientRef(string value, out string processId, out string documentId)
    {
        processId = "";
        documentId = "";
        string[] parts;
        if (value.StartsWith("hwp:", StringComparison.OrdinalIgnoreCase))
            parts = value.Split(':', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        else if (value.StartsWith("untitled-", StringComparison.OrdinalIgnoreCase))
            parts = value.Split('-', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        else
            return false;

        if (parts.Length < 3) return false;
        processId = parts[1];
        documentId = parts[^1];
        return int.TryParse(processId, out _) && int.TryParse(documentId, out _);
    }

    public void Dispose() => _router.Dispose();
}
