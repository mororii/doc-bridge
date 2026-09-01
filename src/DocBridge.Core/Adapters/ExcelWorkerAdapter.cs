using System.Diagnostics;
using System.Globalization;
using System.Text;
using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

/// <summary>
/// Production Excel boundary. The worker owns every Excel RCW; the MCP/CLI host only exchanges
/// JSON over redirected standard streams. If the host is terminated, the pipe closes and the
/// worker disposes its ExcelAdapter through the normal save-state-aware path.
/// </summary>
public sealed class ExcelWorkerAdapter : IAppAdapter, IConnectionLifecycleAdapter
{
    private static readonly TimeSpan CallTimeout = TimeSpan.FromSeconds(150);
    private static readonly TimeSpan DiscoveryCallTimeout = TimeSpan.FromSeconds(45);
    private readonly object _gate = new();
    private Process? _process;
    private StreamWriter? _input;
    private StreamReader? _output;
    private long _nextId;
    private bool _disposed;

    public string App => "excel";

    public static bool CanUseCurrentHost
    {
        get
        {
            var executable = Environment.ProcessPath;
            if (string.IsNullOrWhiteSpace(executable)) return false;
            var name = Path.GetFileNameWithoutExtension(executable);
            return name.Equals("doc-bridge-mcp", StringComparison.OrdinalIgnoreCase) ||
                   name.Equals("doc-bridge-cli", StringComparison.OrdinalIgnoreCase);
        }
    }

    public AdapterStatus GetStatus()
    {
        var value = Call("status");
        return new AdapterStatus(
            Json.GetBool(value, "available"), Json.GetBool(value, "connected"),
            Json.GetString(value, "program") ?? App, Json.GetString(value, "version"),
            Json.GetString(value, "document"), Json.GetString(value, "detail"));
    }

    public JsonObject GetCapabilities() => Call("capabilities");

    public ContextResult GetActiveContext()
    {
        var value = Call("context");
        var result = new ContextResult
        {
            Ok = Json.GetBool(value, "ok"),
            App = Json.GetString(value, "app") ?? App,
            DocumentRef = Json.GetString(value, "documentRef"),
            Summary = Json.GetObj(value, "summary")?.DeepClone().AsObject() ?? new JsonObject(),
            Selection = value["selection"]?.DeepClone() as JsonObject,
            Interaction = value["interaction"]?.DeepClone() as JsonObject,
        };
        CopyStrings(value["warnings"] as JsonArray, result.Warnings);
        CopyStrings(value["errors"] as JsonArray, result.Errors);
        return result;
    }

    public JsonObject Read(JsonObject args) => Call("read", new JsonObject { ["args"] = args.DeepClone() });

    public ApplyPreview Preview(IReadOnlyList<JsonObject> ops)
    {
        var value = Call("preview", new JsonObject { ["ops"] = OpsToJson(ops) });
        var result = new ApplyPreview
        {
            DiffTruncated = Json.GetBool(value, "diffTruncated"),
            RequiresHighRiskApproval = Json.GetBool(value, "requiresHighRiskApproval"),
            Interaction = value["interaction"]?.DeepClone() as JsonObject,
        };
        ParseAffected(value["affected"] as JsonArray, result.Affected);
        ParseDiff(value["diff"] as JsonArray, result.Diff);
        CopyStrings(value["warnings"] as JsonArray, result.Warnings);
        CopyStrings(value["errors"] as JsonArray, result.Errors);
        return result;
    }

    public ApplyExecution Apply(IReadOnlyList<JsonObject> ops, string snapshotId)
    {
        var value = Call("apply", new JsonObject
        {
            ["ops"] = OpsToJson(ops),
            ["snapshotId"] = snapshotId,
        });
        var result = new ApplyExecution
        {
            Ok = Json.GetBool(value, "ok"),
            Readback = value["readback"]?.DeepClone() as JsonObject,
            Interaction = value["interaction"]?.DeepClone() as JsonObject,
        };
        ParseAffected(value["affected"] as JsonArray, result.Affected);
        ParseDiff(value["diff"] as JsonArray, result.Diff);
        if (value["operationResults"] is JsonArray operationResults)
            foreach (var item in operationResults) result.OperationResults.Add(item?.DeepClone());
        CopyStrings(value["warnings"] as JsonArray, result.Warnings);
        CopyStrings(value["errors"] as JsonArray, result.Errors);
        return result;
    }

    public void CaptureSnapshot(string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject>? ops = null)
    {
        _ = Call("captureSnapshot", new JsonObject
        {
            ["snapshotDir"] = snapshotDir,
            ["metadata"] = metadata.DeepClone(),
            ["ops"] = ops is null ? null : OpsToJson(ops),
        });
    }

    public JsonObject RestoreSnapshot(string snapshotDir, JsonObject metadata) =>
        Call("restoreSnapshot", new JsonObject
        {
            ["snapshotDir"] = snapshotDir,
            ["metadata"] = metadata.DeepClone(),
        });

    public JsonObject Disconnect() => Call("disconnect");

    private JsonObject Call(string method, JsonObject? payload = null)
    {
        lock (_gate)
        {
            ObjectDisposedException.ThrowIf(_disposed, this);
            EnsureWorker();
            var id = ++_nextId;
            var request = payload ?? new JsonObject();
            request["id"] = id;
            request["method"] = method;
            try
            {
                _input!.WriteLine(request.ToJsonString());
                _input.Flush();
                var read = _output!.ReadLineAsync();
                if (!read.Wait(TimeoutForMethod(method)))
                    throw new TimeoutException($"Excel worker timed out in {method}");
                var line = read.Result ?? throw new EndOfStreamException("Excel worker closed its output pipe");
                var response = JsonNode.Parse(line) as JsonObject
                    ?? throw new InvalidDataException("Excel worker returned invalid JSON");
                if (Json.GetInt(response, "id") != id)
                    throw new InvalidDataException("Excel worker response id mismatch");
                if (!Json.GetBool(response, "ok"))
                    throw new InvalidOperationException(Json.GetString(response, "error") ?? "Excel worker failed");
                return (response["result"] as JsonObject)?.DeepClone().AsObject() ?? new JsonObject();
            }
            catch
            {
                // This exact child worker is poisoned (timeout, broken pipe, or invalid
                // protocol).  Closing stdio is not enough when its STA is stuck in a COM call:
                // it cannot observe EOF and would remain as an orphan holding Excel RCWs.
                // Terminate only the DocBridge child worker, never EXCEL.EXE and never a process
                // tree.  Releasing the worker process also lets the owner watchdog perform its
                // save-state-aware cleanup for an instance created by DocBridge.
                ResetWorker(terminatePoisonedWorker: true);
                throw;
            }
        }
    }

    internal static TimeSpan TimeoutForMethod(string method) => method switch
    {
        // Keep discovery/read below the common 60-second MCP client deadline so this adapter can
        // reclaim a poisoned worker before the client abandons and restarts the MCP server.
        "status" or "context" or "read" => DiscoveryCallTimeout,
        _ => CallTimeout,
    };

    private void EnsureWorker()
    {
        if (_process is { HasExited: false } && _input is not null && _output is not null) return;
        ResetWorker();
        var executable = Environment.ProcessPath
            ?? throw new InvalidOperationException("Excel worker host executable path is unavailable");
        var start = new ProcessStartInfo
        {
            FileName = executable,
            UseShellExecute = false,
            CreateNoWindow = true,
            WindowStyle = ProcessWindowStyle.Hidden,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
        };
        start.ArgumentList.Add(ExcelWorkerProcess.ModeArgument);
        var process = Process.Start(start) ?? throw new InvalidOperationException("Excel worker did not start");
        var utf8 = new UTF8Encoding(false);
        _process = process;
        _input = new StreamWriter(process.StandardInput.BaseStream, utf8) { AutoFlush = true, NewLine = "\n" };
        _output = new StreamReader(process.StandardOutput.BaseStream, utf8, false);
    }

    private void ResetWorker(bool terminatePoisonedWorker = false)
    {
        try { _input?.Dispose(); } catch { }
        try { _output?.Dispose(); } catch { }
        if (_process is not null)
        {
            var exited = false;
            try
            {
                exited = _process.WaitForExit(terminatePoisonedWorker ? 2000 : 20000);
            }
            catch { }
            if (terminatePoisonedWorker && !exited)
            {
                try
                {
                    // The process object is the exact child started in EnsureWorker.  Do not use
                    // taskkill and do not terminate its watchdog child or Excel process.
                    _process.Kill(entireProcessTree: false);
                    _process.WaitForExit(5000);
                }
                catch { }
            }
            _process.Dispose();
        }
        _input = null;
        _output = null;
        _process = null;
    }

    public void Dispose()
    {
        lock (_gate)
        {
            if (_disposed) return;
            _disposed = true;
            ResetWorker();
        }
    }

    private static JsonArray OpsToJson(IReadOnlyList<JsonObject> ops)
    {
        var result = new JsonArray();
        foreach (var op in ops) result.Add(op.DeepClone());
        return result;
    }

    private static void CopyStrings(JsonArray? source, ICollection<string> target)
    {
        if (source is null) return;
        foreach (var item in source)
            if (item is not null) target.Add(item.GetValue<string>());
    }

    private static void ParseAffected(JsonArray? source, ICollection<AffectedRef> target)
    {
        if (source is null) return;
        foreach (var item in source.OfType<JsonObject>())
            target.Add(new AffectedRef(Json.GetString(item, "type") ?? "", Json.GetString(item, "ref") ?? ""));
    }

    private static void ParseDiff(JsonArray? source, ICollection<DiffEntry> target)
    {
        if (source is null) return;
        foreach (var item in source.OfType<JsonObject>())
            target.Add(new DiffEntry
            {
                Ref = Json.GetString(item, "ref") ?? "",
                Before = item["before"]?.DeepClone(),
                After = item["after"]?.DeepClone(),
            });
    }
}

public static class ExcelWorkerProcess
{
    public const string ModeArgument = "--excel-worker";

    public static bool TryRun(string[] args, out int exitCode)
    {
        exitCode = 0;
        if (args.Length != 1 || !string.Equals(args[0], ModeArgument, StringComparison.Ordinal))
            return false;
        exitCode = Run();
        return true;
    }

    private static int Run()
    {
        var utf8 = new UTF8Encoding(false);
        using var input = new StreamReader(Console.OpenStandardInput(), utf8, false);
        using var output = new StreamWriter(Console.OpenStandardOutput(), utf8) { AutoFlush = true, NewLine = "\n" };
        using var adapter = new ExcelAdapter();
        string? line;
        while ((line = input.ReadLine()) is not null)
        {
            JsonObject response;
            long id = 0;
            try
            {
                var request = JsonNode.Parse(line) as JsonObject
                    ?? throw new InvalidDataException("request must be a JSON object");
                id = request["id"]?.GetValue<long>() ?? 0;
                var method = Json.GetString(request, "method") ?? throw new InvalidDataException("method is required");
                response = new JsonObject
                {
                    ["id"] = id,
                    ["ok"] = true,
                    ["result"] = Dispatch(adapter, method, request),
                };
            }
            catch (Exception ex)
            {
                response = new JsonObject { ["id"] = id, ["ok"] = false, ["error"] = ex.Message };
            }
            output.WriteLine(response.ToJsonString());
        }
        return 0;
    }

    private static JsonObject Dispatch(ExcelAdapter adapter, string method, JsonObject request) => method switch
    {
        "status" => StatusToJson(adapter.GetStatus()),
        "capabilities" => adapter.GetCapabilities(),
        "context" => adapter.GetActiveContext().ToJson(),
        "read" => adapter.Read(Json.GetObj(request, "args") ?? new JsonObject()),
        "preview" => PreviewToJson(adapter.Preview(ParseOps(request))),
        "apply" => ExecutionToJson(adapter.Apply(ParseOps(request), Json.GetString(request, "snapshotId") ?? "")),
        "captureSnapshot" => CaptureSnapshot(adapter, request),
        "restoreSnapshot" => adapter.RestoreSnapshot(
            Json.GetString(request, "snapshotDir") ?? throw new InvalidDataException("snapshotDir is required"),
            Json.GetObj(request, "metadata") ?? new JsonObject()),
        "disconnect" => adapter.Disconnect(),
        _ => throw new InvalidDataException($"unknown Excel worker method '{method}'"),
    };

    private static JsonObject CaptureSnapshot(ExcelAdapter adapter, JsonObject request)
    {
        var ops = request["ops"] is JsonArray ? ParseOps(request) : null;
        adapter.CaptureSnapshot(
            Json.GetString(request, "snapshotDir") ?? throw new InvalidDataException("snapshotDir is required"),
            Json.GetObj(request, "metadata") ?? new JsonObject(), ops);
        return new JsonObject { ["captured"] = true };
    }

    private static IReadOnlyList<JsonObject> ParseOps(JsonObject request)
    {
        var result = new List<JsonObject>();
        if (request["ops"] is not JsonArray array) return result;
        foreach (var item in array.OfType<JsonObject>()) result.Add(item.DeepClone().AsObject());
        return result;
    }

    private static JsonObject StatusToJson(AdapterStatus status) => new()
    {
        ["available"] = status.Available,
        ["connected"] = status.Connected,
        ["program"] = status.Program,
        ["version"] = status.Version,
        ["document"] = status.Document,
        ["detail"] = status.Detail,
    };

    private static JsonObject PreviewToJson(ApplyPreview preview) => new()
    {
        ["affected"] = AffectedToJson(preview.Affected),
        ["diff"] = DiffToJson(preview.Diff),
        ["diffTruncated"] = preview.DiffTruncated,
        ["requiresHighRiskApproval"] = preview.RequiresHighRiskApproval,
        ["warnings"] = StringsToJson(preview.Warnings),
        ["errors"] = StringsToJson(preview.Errors),
        ["interaction"] = preview.Interaction?.DeepClone(),
    };

    private static JsonObject ExecutionToJson(ApplyExecution execution) => new()
    {
        ["ok"] = execution.Ok,
        ["affected"] = AffectedToJson(execution.Affected),
        ["diff"] = DiffToJson(execution.Diff),
        ["operationResults"] = execution.OperationResults.DeepClone(),
        ["readback"] = execution.Readback?.DeepClone(),
        ["interaction"] = execution.Interaction?.DeepClone(),
        ["warnings"] = StringsToJson(execution.Warnings),
        ["errors"] = StringsToJson(execution.Errors),
    };

    private static JsonArray AffectedToJson(IEnumerable<AffectedRef> affected)
    {
        var result = new JsonArray();
        foreach (var item in affected)
            result.Add(new JsonObject { ["type"] = item.Type, ["ref"] = item.Ref });
        return result;
    }

    private static JsonArray DiffToJson(IEnumerable<DiffEntry> diff)
    {
        var result = new JsonArray();
        foreach (var item in diff) result.Add(item.ToJson());
        return result;
    }

    private static JsonArray StringsToJson(IEnumerable<string> values)
    {
        var result = new JsonArray();
        foreach (var value in values) result.Add(value);
        return result;
    }
}
