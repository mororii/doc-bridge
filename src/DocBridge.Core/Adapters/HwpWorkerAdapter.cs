using System.Collections.Concurrent;
using System.Diagnostics;
using System.Text;
using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

/// <summary>
/// 한글 COM을 MCP 프로세스 밖에서 실행한다. 모달/COM hang 시 worker만 종료·교체되어
/// Excel/CAD 도구와 MCP stdio 세션은 계속 사용할 수 있다.
/// </summary>
public sealed class HwpWorkerAdapter : IAppAdapter, IHwpAutomationAdapter, IPreviewReuseAdapter
{
    private static readonly TimeSpan CircuitCooldown = TimeSpan.FromSeconds(45);
    private static readonly TimeSpan TransportCircuitCooldown = TimeSpan.FromSeconds(15);
    private readonly object _gate = new();
    private readonly string? _workerPath;
    private readonly HwpAdapter? _fallback;
    private readonly ConcurrentQueue<string> _stderr = new();
    private readonly HwpWorkerCircuitBreaker _circuit = new();
    private Process? _process;
    private bool _disposed;
    private bool _poisoned;

    public HwpWorkerAdapter(string? workerPath = null, bool allowDirectFallback = false)
    {
        _workerPath = ResolveWorkerPath(workerPath);
        if (_workerPath is null && allowDirectFallback) _fallback = new HwpAdapter();
    }

    public string App => "hwp";

    public JsonObject GetCapabilities()
    {
        if (_fallback is not null)
        {
            var direct = _fallback.GetCapabilities();
            direct["processIsolation"] = false;
            direct["workerFallback"] = true;
            return direct;
        }
        var result = Call("getCapabilities", new JsonObject(), readOnly: true);
        result["processIsolation"] = true;
        result["workerExecutable"] = _workerPath;
        result["workerRestartLimit"] = 1;
        result["requestTimeoutsSec"] = TimeoutCapabilities();
        result["circuitBreaker"] = new JsonObject
        {
            ["cooldownSec"] = (int)CircuitCooldown.TotalSeconds,
            ["writeRetry"] = false,
            ["readTransportRetry"] = 1,
        };
        return result;
    }

    public AdapterStatus GetStatus() => _fallback is not null
        ? _fallback.GetStatus()
        : HwpWorkerProtocol.StatusFromJson(Call("getStatus", new JsonObject(), readOnly: true));

    public ContextResult GetActiveContext() => _fallback is not null
        ? _fallback.GetActiveContext()
        : HwpWorkerProtocol.ContextFromJson(Call("getActiveContext", new JsonObject(), readOnly: true));

    public JsonObject Read(JsonObject args) => _fallback is not null
        ? _fallback.Read(args)
        : Call("read", args, readOnly: true);

    public ApplyPreview Preview(IReadOnlyList<JsonObject> ops)
    {
        if (_fallback is not null) return _fallback.Preview(ops);
        return HwpWorkerProtocol.PreviewFromJson(Call("preview", OpsPayload(ops), readOnly: true));
    }

    public ApplyExecution Apply(IReadOnlyList<JsonObject> ops, string snapshotId)
    {
        if (_fallback is not null) return _fallback.Apply(ops, snapshotId);
        var payload = OpsPayload(ops);
        payload["snapshotId"] = snapshotId;
        return HwpWorkerProtocol.ExecutionFromJson(Call("apply", payload, readOnly: false));
    }

    public void CaptureSnapshot(string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject>? ops = null)
    {
        if (_fallback is not null)
        {
            _fallback.CaptureSnapshot(snapshotDir, metadata, ops);
            return;
        }
        var payload = OpsPayload(ops ?? Array.Empty<JsonObject>());
        payload["snapshotDir"] = snapshotDir;
        payload["metadata"] = metadata.DeepClone();
        var result = Call("captureSnapshot", payload, readOnly: false);
        ReplaceSnapshotMetadata(result, metadata);
    }

    internal static void ReplaceSnapshotMetadata(JsonObject result, JsonObject metadata)
    {
        if (!string.IsNullOrWhiteSpace(Json.GetString(result, "errorCode")))
            throw StructuredFailure(result);
        var updated = Json.GetObj(result, "metadata") ?? new JsonObject();
        metadata.Clear();
        foreach (var pair in updated) metadata[pair.Key] = pair.Value?.DeepClone();
    }

    public JsonObject RestoreSnapshot(string snapshotDir, JsonObject metadata)
    {
        if (_fallback is not null) return _fallback.RestoreSnapshot(snapshotDir, metadata);
        return Call("restoreSnapshot", new JsonObject
        {
            ["snapshotDir"] = snapshotDir,
            ["metadata"] = metadata.DeepClone(),
        }, readOnly: false);
    }

    public JsonObject ValidatePreviewReuse(
        string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject> ops)
    {
        if (_fallback is IPreviewReuseAdapter direct)
            return direct.ValidatePreviewReuse(snapshotDir, metadata, ops);
        var payload = OpsPayload(ops);
        payload["snapshotDir"] = snapshotDir;
        payload["metadata"] = metadata.DeepClone();
        return Call("validatePreviewReuse", payload, readOnly: true);
    }

    public JsonObject Launch(JsonObject args) => _fallback is not null
        ? _fallback.Launch(args)
        : Call("launch", args, readOnly: false);

    public JsonObject Doctor(JsonObject args) => _fallback is not null
        ? _fallback.Doctor(args)
        : Call("doctor", args, readOnly: true);

    public JsonObject RepairTypeLib(JsonObject args) => _fallback is not null
        ? _fallback.RepairTypeLib(args)
        : Call("repairTypeLib", args, readOnly: false);

    private JsonObject Call(string method, JsonObject payload, bool readOnly)
    {
        if (_workerPath is null)
            throw new HwpAutomationException("HWP_WORKER_NOT_FOUND",
                "doc-bridge-hwp-worker.exe를 찾지 못했습니다.", "DocBridge를 다시 설치하거나 배포 무결성 검사를 실행하세요.");

        var timeout = ResolveRequestTimeout(method, payload);
        var maxAttempts = CanRetryReadTransport(method, payload, readOnly) ? 2 : 1;
        Exception? last = null;
        for (var attempt = 1; attempt <= maxAttempts; attempt++)
        {
            try
            {
                lock (_gate)
                {
                    ThrowIfDisposed();
                    ThrowIfCircuitOpen(method);
                    var diagnosticOrRepair = method is "doctor" or "repairTypeLib";
                    if (!diagnosticOrRepair && HwpUiFailureDetector.Detect() is { } existingFailure)
                    {
                        _circuit.Open("HWP_UI_INITIALIZATION_FAILED", CircuitCooldown);
                        throw UiInitializationException(existingFailure);
                    }
                    if (_poisoned) RestartWorker();
                    EnsureWorker();
                    var requestId = Guid.NewGuid().ToString("n");
                    var request = new JsonObject
                    {
                        ["id"] = requestId,
                        ["method"] = method,
                        ["payload"] = payload.DeepClone(),
                    };
                    _process!.StandardInput.WriteLine(Json.ToCompact(request));
                    _process.StandardInput.Flush();

                    var lineTask = _process.StandardOutput.ReadLineAsync();
                    var line = WaitForWorkerResponse(lineTask, timeout);
                    if (line is null)
                        throw new IOException($"HWP worker exited before replying (exit={SafeExitCode(_process)}; stderr={LastStderr()})");
                    var response = Json.ParseObject(line)
                        ?? throw new InvalidDataException("HWP worker returned invalid JSON");
                    if (!string.Equals(Json.GetString(response, "id"), requestId, StringComparison.Ordinal))
                        throw new InvalidDataException("HWP worker response id mismatch");
                    if (!Json.GetBool(response, "transportOk"))
                        throw new IOException(Json.GetString(response, "error") ?? "HWP worker transport failure");
                    var result = Json.GetObj(response, "result")?.DeepClone() as JsonObject ?? new JsonObject();
                    var resultCode = Json.GetString(result, "errorCode");
                    _poisoned = HwpWorkerProtocol.ContainsComTimeout(result) ||
                                resultCode == "HWP_UI_INITIALIZATION_FAILED" ||
                                Json.GetBool(response, "restartRequired");
                    if (_poisoned)
                        _circuit.Open(resultCode ?? "HWP_COM_TIMEOUT", CircuitCooldown);
                    if (!string.IsNullOrWhiteSpace(resultCode) && method is not ("doctor" or "repairTypeLib"))
                        throw StructuredFailure(result);
                    else if (ShouldResetCircuitAfterResponse(method, result))
                        // doctor는 등록/환경 진단 성공일 뿐 COM 호출 정상화를 증명하지 않는다.
                        // 실패 결과도 transportOk=true일 수 있으므로 실제 ok 응답만 회로를 닫는다.
                        _circuit.Reset();
                    return result;
                }
            }
            catch (HwpAutomationException ex) when (ex.Code == "HWP_CIRCUIT_OPEN")
            {
                throw;
            }
            catch (HwpAutomationException ex)
            {
                if (ex.Code is "HWP_UI_INITIALIZATION_FAILED" or "HWP_COM_TIMEOUT")
                {
                    lock (_gate)
                    {
                        _poisoned = true;
                        _circuit.Open(ex.Code, CircuitCooldown);
                        RestartWorker();
                    }
                }
                // Application-level failures (document not found, invalid target,
                // repair declined, and so on) are not transport failures. Preserve
                // their stable code instead of restarting and wrapping them.
                throw;
            }
            catch (Exception ex)
            {
                last = ex;
                lock (_gate) RestartWorker();
                if (HwpUiFailureDetector.Detect() is { } detectedFailure)
                {
                    lock (_gate) _circuit.Open("HWP_UI_INITIALIZATION_FAILED", CircuitCooldown);
                    throw UiInitializationException(detectedFailure);
                }
                if (HwpUiFailureDetector.IsDeterministicFailureMessage(ex.Message) ||
                    HwpUiFailureDetector.IsDeterministicFailureMessage(LastStderr()))
                {
                    lock (_gate) _circuit.Open("HWP_UI_INITIALIZATION_FAILED", CircuitCooldown);
                    throw new HwpAutomationException(
                        "HWP_UI_INITIALIZATION_FAILED",
                        "한글 UI 형식 초기화 오류가 감지되어 빈 창 반복 실행을 중단했습니다.",
                        HwpUiFailureDetector.UpdateAction,
                        ex);
                }
                if (ex is TimeoutException)
                {
                    lock (_gate) _circuit.Open("HWP_COM_TIMEOUT", CircuitCooldown);
                    break;
                }
                if (attempt < maxAttempts && IsRetryableTransportFailure(ex))
                {
                    Thread.Sleep(250 * attempt);
                    continue;
                }
                break;
            }
        }

        if (last is TimeoutException)
            throw new HwpAutomationException("HWP_COM_TIMEOUT",
                $"한글 COM worker의 '{method}' 작업이 {timeout.TotalSeconds:0}초 안에 응답하지 않았습니다.",
                $"한글의 팝업/대화상자를 닫고 {CircuitCooldown.TotalSeconds:0}초 뒤 다시 시도하세요. 그 전에는 빈 창 반복 실행을 막기 위해 요청을 즉시 중단합니다.",
                last, (int)CircuitCooldown.TotalMilliseconds);
        lock (_gate) _circuit.Open("HWP_WORKER_FAILED", TransportCircuitCooldown);
        throw new HwpAutomationException("HWP_WORKER_FAILED",
            $"한글 worker 요청에 실패했습니다: {last?.Message}",
            $"hwp_doctor와 설치 무결성 검사를 실행하고 {TransportCircuitCooldown.TotalSeconds:0}초 뒤 다시 시도하세요.",
            last, (int)TransportCircuitCooldown.TotalMilliseconds);
    }

    internal static TimeSpan ResolveRequestTimeout(string method, JsonObject payload)
    {
        return method switch
        {
            "getCapabilities" or "getStatus" or "doctor" => TimeSpan.FromSeconds(15),
            "getActiveContext" => TimeSpan.FromSeconds(20),
            "read" => ReadTimeout(payload),
            "preview" => OpsTimeout(payload, preview: true),
            "apply" => OpsTimeout(payload, preview: false),
            "validatePreviewReuse" => TimeSpan.FromSeconds(30),
            "captureSnapshot" => TimeSpan.FromSeconds(60),
            "restoreSnapshot" => TimeSpan.FromSeconds(90),
            "launch" => LaunchTimeout(payload),
            "repairTypeLib" => TimeSpan.FromSeconds(60),
            _ => TimeSpan.FromSeconds(45),
        };
    }

    private static TimeSpan ReadTimeout(JsonObject payload)
    {
        // The public hwp_read_text contract uses `scope`. Keep `mode` only as a
        // compatibility fallback for older internal callers.
        var scope = (Json.GetString(payload, "scope") ??
                     Json.GetString(payload, "mode") ??
                     "document").ToLowerInvariant();
        return scope is "bundle" or "document_map" or "structure" or "fields" or "tables"
            ? TimeSpan.FromSeconds(45)
            : TimeSpan.FromSeconds(30);
    }

    private static TimeSpan LaunchTimeout(JsonObject payload)
    {
        // The public hwp_launch contract is creationMode=docx-first plus
        // sourceFile. Legacy mode/docx keys remain supported for compatibility.
        var creationMode = (Json.GetString(payload, "creationMode") ?? "").ToLowerInvariant();
        var legacyMode = (Json.GetString(payload, "mode") ?? "").ToLowerInvariant();
        return creationMode == "docx-first" ||
               !string.IsNullOrWhiteSpace(Json.GetString(payload, "sourceFile")) ||
               legacyMode.Contains("docx", StringComparison.Ordinal) ||
               !string.IsNullOrWhiteSpace(Json.GetString(payload, "docx"))
            ? TimeSpan.FromSeconds(90)
            : TimeSpan.FromSeconds(45);
    }

    private static TimeSpan OpsTimeout(JsonObject payload, bool preview)
    {
        var ops = Json.GetArr(payload, "ops") ?? new JsonArray();
        if (preview)
            return ops.Count > 20 ? TimeSpan.FromSeconds(60) : TimeSpan.FromSeconds(45);

        var expensive = ops.Any(node => node is JsonObject op &&
            (Json.GetString(op, "op") ?? "") is
                "insert_table" or "insert_picture" or "export_pdf" or
                "set_header_footer_text" or "table_set_cells" or "table_set_row_heights" or "format_paragraphs");
        if (expensive || ops.Count > 20) return TimeSpan.FromSeconds(90);
        if (ops.Count > 10) return TimeSpan.FromSeconds(60);
        return TimeSpan.FromSeconds(45);
    }

    internal static bool CanRetryReadTransport(string method, JsonObject payload, bool readOnly)
    {
        if (!readOnly) return false;
        if (method is "read" or "preview")
            return string.IsNullOrWhiteSpace(Json.GetString(payload, "file")) &&
                   !PayloadOpsHaveFile(payload);
        return method is "getCapabilities" or "getStatus" or "getActiveContext" or
            "doctor" or "validatePreviewReuse";
    }

    private static bool PayloadOpsHaveFile(JsonObject payload) =>
        (Json.GetArr(payload, "ops") ?? new JsonArray()).Any(node =>
            node is JsonObject op && !string.IsNullOrWhiteSpace(Json.GetString(op, "file")));

    internal static bool IsRetryableTransportFailure(Exception error) =>
        error is IOException or InvalidDataException;

    internal static bool ShouldResetCircuitAfterResponse(string method, JsonObject result) =>
        method != "doctor" && Json.GetBool(result, "ok");

    internal static HwpAutomationException StructuredFailure(JsonObject result)
    {
        var code = Json.GetString(result, "errorCode") ?? "HWP_AUTOMATION_FAILED";
        var message = (Json.GetArr(result, "errors") ?? new JsonArray())
            .Select(node => node?.GetValue<string>())
            .FirstOrDefault(value => !string.IsNullOrWhiteSpace(value))
            ?? Json.GetString(result, "message")
            ?? "한글 자동화 작업에 실패했습니다.";
        return new HwpAutomationException(
            code,
            message,
            Json.GetString(result, "userAction"),
            retryAfterMs: Json.GetInt(result, "retryAfterMs"));
    }

    private void ThrowIfCircuitOpen(string method)
    {
        // doctor/repair와 자동 rollback은 회로가 열린 원인을 확인·복구하기 위한
        // 명시적 경로이므로 허용한다. 일반 읽기/쓰기는 새 빈 창 반복을 막기 위해 차단한다.
        if (method is "doctor" or "repairTypeLib" or "restoreSnapshot") return;
        if (!_circuit.TryGetOpen(out var state)) return;
        throw new HwpAutomationException(
            "HWP_CIRCUIT_OPEN",
            $"한글 자동화가 직전 오류({state.Code}) 뒤 보호 대기 중입니다. 새 worker나 빈 한글 창을 다시 만들지 않았습니다.",
            $"열린 팝업을 닫고 약 {Math.Max(1, (int)Math.Ceiling(state.Remaining.TotalSeconds))}초 뒤 다시 시도하거나 hwp_doctor를 실행하세요.",
            retryAfterMs: Math.Max(1, (int)Math.Ceiling(state.Remaining.TotalMilliseconds)));
    }

    private static JsonObject TimeoutCapabilities() => new()
    {
        ["status"] = 15,
        ["context"] = 20,
        ["readText"] = 30,
        ["readComplex"] = 45,
        ["preview"] = "45-60",
        ["apply"] = "45-90",
        ["snapshot"] = 60,
        ["restore"] = 90,
        ["docxOrPdf"] = 90,
    };

    private static string? WaitForWorkerResponse(Task<string?> lineTask, TimeSpan timeout)
    {
        var stopwatch = Stopwatch.StartNew();
        while (stopwatch.Elapsed < timeout)
        {
            var remaining = timeout - stopwatch.Elapsed;
            var delay = remaining < TimeSpan.FromMilliseconds(250)
                ? remaining
                : TimeSpan.FromMilliseconds(250);
            var completed = Task.WhenAny(lineTask, Task.Delay(delay)).GetAwaiter().GetResult();
            if (completed == lineTask) return lineTask.GetAwaiter().GetResult();
            if (HwpUiFailureDetector.Detect() is { } failure)
                throw UiInitializationException(failure);
        }
        throw new TimeoutException($"HWP worker response timed out after {timeout.TotalSeconds:0} seconds");
    }

    private static HwpAutomationException UiInitializationException(HwpUiFailure failure) =>
        new(
            "HWP_UI_INITIALIZATION_FAILED",
            $"한글 UI 초기화 오류가 감지되었습니다 ({failure.Signature}). 빈 창 반복 실행을 중단했습니다.",
            HwpUiFailureDetector.UpdateAction);

    private void EnsureWorker()
    {
        if (_process is { HasExited: false }) return;
        RestartWorker();
        var start = new ProcessStartInfo
        {
            FileName = _workerPath!,
            UseShellExecute = false,
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            StandardInputEncoding = new UTF8Encoding(false),
            StandardOutputEncoding = new UTF8Encoding(false),
            StandardErrorEncoding = new UTF8Encoding(false),
        };
        HwpEnvironmentDoctor.ApplyAutomationEnvironment(start);
        _process = Process.Start(start) ?? throw new IOException("HWP worker process start failed");
        _process.ErrorDataReceived += (sender, eventArgs) =>
        {
            if (string.IsNullOrWhiteSpace(eventArgs.Data)) return;
            _stderr.Enqueue(eventArgs.Data);
            while (_stderr.Count > 20) _stderr.TryDequeue(out _);
        };
        _process.BeginErrorReadLine();
        _poisoned = false;
    }

    private void RestartWorker(bool graceful = false)
    {
        var process = _process;
        _process = null;
        _poisoned = false;
        if (process is null) return;
        try
        {
            if (!process.HasExited && graceful)
            {
                try
                {
                    process.StandardInput.WriteLine(Json.ToCompact(new JsonObject
                    {
                        ["id"] = Guid.NewGuid().ToString("n"),
                        ["method"] = "shutdown",
                        ["payload"] = new JsonObject(),
                    }));
                    process.StandardInput.Flush();
                    // The worker intentionally sends no response.  Reaching EOF lets its
                    // `using HwpAdapter` dispose the private COM instance and Hwp.exe.
                    _ = process.WaitForExit(7000);
                }
                catch { }
            }
            if (!process.HasExited) process.Kill(entireProcessTree: true);
            process.WaitForExit(5000);
        }
        catch { }
        finally { process.Dispose(); }
    }

    private static JsonObject OpsPayload(IEnumerable<JsonObject> ops)
    {
        var array = new JsonArray();
        foreach (var op in ops) array.Add(op.DeepClone());
        return new JsonObject { ["ops"] = array };
    }

    private static string? ResolveWorkerPath(string? explicitPath)
    {
        var candidates = new[]
        {
            explicitPath,
            Environment.GetEnvironmentVariable("DOCBRIDGE_HWP_WORKER"),
            Path.Combine(AppContext.BaseDirectory, "doc-bridge-hwp-worker.exe"),
        };
        foreach (var candidate in candidates)
            if (!string.IsNullOrWhiteSpace(candidate) && File.Exists(candidate)) return Path.GetFullPath(candidate);
        return null;
    }

    private string LastStderr() => string.Join(" | ", _stderr.ToArray().TakeLast(5));
    private static string SafeExitCode(Process process) { try { return process.HasExited ? process.ExitCode.ToString() : "running"; } catch { return "unknown"; } }
    private void ThrowIfDisposed() { if (_disposed) throw new ObjectDisposedException(nameof(HwpWorkerAdapter)); }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        _fallback?.Dispose();
        lock (_gate) RestartWorker(graceful: true);
    }
}

internal sealed class HwpWorkerCircuitBreaker
{
    private readonly TimeProvider _timeProvider;
    private DateTimeOffset _openUntil;
    private string? _code;

    internal HwpWorkerCircuitBreaker(TimeProvider? timeProvider = null) =>
        _timeProvider = timeProvider ?? TimeProvider.System;

    internal void Open(string code, TimeSpan cooldown)
    {
        var until = _timeProvider.GetUtcNow() + cooldown;
        if (until > _openUntil) _openUntil = until;
        _code = code;
    }

    internal void Reset()
    {
        _openUntil = default;
        _code = null;
    }

    internal bool TryGetOpen(out HwpCircuitState state)
    {
        var now = _timeProvider.GetUtcNow();
        if (_code is not null && _openUntil > now)
        {
            state = new HwpCircuitState(_code, _openUntil - now);
            return true;
        }
        Reset();
        state = default;
        return false;
    }
}

internal readonly record struct HwpCircuitState(string Code, TimeSpan Remaining);
