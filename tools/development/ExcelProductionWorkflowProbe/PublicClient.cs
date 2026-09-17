using System.Diagnostics;

namespace DocBridge.Development.ExcelProductionWorkflowProbe;

internal sealed class ToolCall
{
    public required string Id { get; init; }
    public required string Tool { get; init; }
    public required JsonObject Arguments { get; init; }
    public JsonNode? Response { get; set; }
    public int ExitCode { get; set; }
    public long ElapsedMs { get; set; }
    public string? Error { get; set; }
}

/// <summary>
/// Persistent MCP is the only live transport. CLI is process-per-call and cannot
/// prove workbook ownership; it also cannot carry confirmToken without writing it to disk.
/// </summary>
internal sealed class PublicClient : IDisposable
{
    private readonly ProbeOptions _options;
    private readonly string _rawDir;
    private readonly string _homeDir;
    private int _seq;
    private readonly List<ToolCall> _calls = [];
    private Process? _mcp;
    private Task? _mcpStderr;
    private readonly StringBuilder _mcpErr = new();
    private readonly object _mcpGate = new();

    public PublicClient(ProbeOptions options)
    {
        _options = options;
        _rawDir = Path.Combine(options.OutputDir, "raw");
        _homeDir = Path.Combine(options.OutputDir, "probe-home");
        Directory.CreateDirectory(_rawDir);
        Directory.CreateDirectory(_homeDir);
    }

    public IReadOnlyList<ToolCall> Calls => _calls;

    public JsonNode Call(string tool, JsonObject args)
    {
        var id = $"call-{++_seq:0000}-{tool}";
        var started = Stopwatch.StartNew();
        var record = new ToolCall { Id = id, Tool = tool, Arguments = args };
        try
        {
            record.Response = Detached(_options.Transport == "mcp"
                ? CallMcp(tool, args, out var exit)
                : CallCli(tool, args, out exit));
            record.ExitCode = exit;
        }
        catch (Exception ex)
        {
            record.Error = ex.Message;
            record.Response = new JsonObject { ["ok"] = false, ["error"] = ex.Message };
            record.ExitCode = 2;
        }
        record.ElapsedMs = started.ElapsedMilliseconds;
        _calls.Add(record);
        JsonUtil.Write(Path.Combine(_rawDir, $"{id}.json"), new JsonObject
        {
            ["tool"] = tool,
            ["elapsedMs"] = record.ElapsedMs,
            ["exitCode"] = record.ExitCode,
            ["arguments"] = Redact.Tokens(args),
            ["response"] = Redact.Tokens(record.Response),
            ["error"] = record.Error,
        });
        return Detached(record.Response);
    }

    private static JsonNode Detached(JsonNode? node) =>
        node?.DeepClone() ?? new JsonObject { ["ok"] = false };

    public JsonObject ProtocolSelfTest()
    {
        if (_options.Transport != "mcp")
            return new JsonObject { ["ok"] = false, ["error"] = "protocol self-test requires --transport mcp" };

        var ping = Call("core_ping", new JsonObject());
        var caps = Call("core_get_capabilities", new JsonObject());
        var status = Call("core_get_status", new JsonObject { ["app"] = "excel" });
        var ok = JsonUtil.Bool(ping, "ok") != false && _mcp is { HasExited: false };
        return new JsonObject
        {
            ["ok"] = ok,
            ["persistentProcess"] = _mcp is { HasExited: false },
            ["pid"] = _mcp?.Id,
            ["ping"] = Redact.Tokens(ping),
            ["capabilitiesKeys"] = CapsHint(caps),
            ["statusExcelConnected"] = JsonUtil.Bool(JsonUtil.Get(JsonUtil.Get(status, "apps"), "excel"), "connected"),
            ["note"] = "No excel_apply_ops during protocol self-test.",
        };
    }

    private static JsonNode CapsHint(JsonNode caps)
    {
        var write = ContractCatalog.AdvertisedWriteOps(caps);
        return new JsonArray(write.OrderBy(s => s, StringComparer.OrdinalIgnoreCase).Select(s => JsonValue.Create(s)).ToArray());
    }

    private JsonNode CallCli(string tool, JsonObject args, out int exit)
    {
        if (ContainsToken(args))
            throw new InvalidOperationException("CLI transport refuses to write confirmToken to disk; use --transport mcp");
        if (string.IsNullOrWhiteSpace(_options.CliPath) || !File.Exists(_options.CliPath))
            throw new InvalidOperationException("--cli must point at an already-built doc-bridge-cli");

        var jsonPath = Path.Combine(_rawDir, $"arg-{_seq:0000}.json");
        try
        {
            JsonUtil.Write(jsonPath, Redact.Tokens(args) as JsonNode ?? new JsonObject(), pretty: false);
            var start = new ProcessStartInfo
            {
                FileName = _options.CliPath,
                Arguments = $"{tool} --json-file \"{jsonPath}\"",
                UseShellExecute = false,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                CreateNoWindow = true,
                StandardOutputEncoding = new UTF8Encoding(false),
                StandardErrorEncoding = new UTF8Encoding(false),
            };
            start.Environment["DOCBRIDGE_HOME"] = _homeDir;

            using var proc = Process.Start(start) ?? throw new InvalidOperationException("failed to start CLI");
            var stdoutTask = proc.StandardOutput.ReadToEndAsync();
            var stderrTask = proc.StandardError.ReadToEndAsync();
            if (!proc.WaitForExit(180_000))
            {
                try { proc.Kill(entireProcessTree: false); } catch { /* leave Office alone */ }
                throw new TimeoutException($"{tool} exceeded 180s");
            }
            Task.WaitAll(new Task[] { stdoutTask, stderrTask }, 5_000);
            exit = proc.ExitCode;
            var stdout = stdoutTask.IsCompletedSuccessfully ? stdoutTask.Result : "";
            return JsonNode.Parse(string.IsNullOrWhiteSpace(stdout) ? "{}" : stdout.Trim())
                   ?? new JsonObject { ["ok"] = false, ["error"] = "empty CLI stdout" };
        }
        finally
        {
            try { if (File.Exists(jsonPath)) File.Delete(jsonPath); } catch { /* best-effort */ }
        }
    }

    private JsonNode CallMcp(string tool, JsonObject args, out int exit)
    {
        EnsureMcp();
        var rpcId = _seq;
        var payload = new JsonObject
        {
            ["jsonrpc"] = "2.0",
            ["id"] = rpcId,
            ["method"] = "tools/call",
            ["params"] = new JsonObject
            {
                ["name"] = tool,
                ["arguments"] = args.DeepClone(),
            },
        };

        JsonNode? envelope;
        lock (_mcpGate)
        {
            _mcp!.StandardInput.WriteLine(payload.ToJsonString(JsonUtil.WriteCompact));
            _mcp.StandardInput.Flush();
            var timeout = string.Equals(tool, "excel_apply_ops", StringComparison.OrdinalIgnoreCase)
                ? TimeSpan.FromMinutes(12)
                : TimeSpan.FromSeconds(180);
            envelope = ReadMatchingResponse(rpcId, timeout);
        }

        exit = 0;
        if (envelope is JsonObject err && err["error"] is not null)
            throw new InvalidOperationException($"MCP error: {err["error"]}");
        return UnwrapToolResult(envelope);
    }

    private JsonNode ReadMatchingResponse(int rpcId, TimeSpan timeout)
    {
        var deadline = DateTime.UtcNow + timeout;
        while (DateTime.UtcNow < deadline)
        {
            var remain = deadline - DateTime.UtcNow;
            var line = ReadLineWithTimeout(remain);
            if (line is null)
                throw new TimeoutException($"MCP produced no response for id {rpcId}");
            var parsed = JsonNode.Parse(line);
            if (parsed is JsonObject obj)
            {
                if (!obj.ContainsKey("id"))
                    continue; // notification
                var idNode = obj["id"];
                var got = idNode is JsonValue v && v.TryGetValue<int>(out var n) ? n
                    : int.TryParse(idNode?.ToString(), out var p) ? p : int.MinValue;
                if (got != rpcId)
                    throw new InvalidOperationException($"MCP response id {got} != request {rpcId}");
                return obj;
            }
        }
        throw new TimeoutException($"MCP timed out waiting for id {rpcId}");
    }

    private string? ReadLineWithTimeout(TimeSpan timeout)
    {
        var read = _mcp!.StandardOutput.ReadLineAsync();
        if (read.Wait(timeout))
            return read.Result;
        throw new TimeoutException("MCP stdout ReadLine timed out");
    }

    private static JsonNode UnwrapToolResult(JsonNode envelope)
    {
        var result = JsonUtil.Get(envelope, "result") ?? envelope;
        if (JsonUtil.Bool(result, "isError") == true)
        {
            var text = (JsonUtil.Get(result, "content") as JsonArray)?
                .OfType<JsonObject>()
                .Select(o => JsonUtil.Str(o, "text"))
                .FirstOrDefault(s => !string.IsNullOrWhiteSpace(s));
            throw new InvalidOperationException(text ?? "MCP tools/call isError");
        }

        var structured = JsonUtil.Get(result, "structuredContent");
        if (structured is not null)
            return structured.DeepClone();

        if (JsonUtil.Get(result, "content") is JsonArray content)
        {
            foreach (var item in content.OfType<JsonObject>())
            {
                var text = JsonUtil.Str(item, "text");
                if (string.IsNullOrWhiteSpace(text)) continue;
                try { return JsonNode.Parse(text) ?? result.DeepClone(); }
                catch { return new JsonObject { ["ok"] = false, ["text"] = text }; }
            }
        }
        return result.DeepClone();
    }

    private void EnsureMcp()
    {
        if (_mcp is { HasExited: false }) return;
        if (string.IsNullOrWhiteSpace(_options.McpPath) || !File.Exists(_options.McpPath))
            throw new InvalidOperationException("--mcp must point at an already-built doc-bridge-mcp");
        var start = new ProcessStartInfo
        {
            FileName = _options.McpPath,
            Arguments = "--stdio",
            UseShellExecute = false,
            RedirectStandardInput = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true,
            StandardInputEncoding = new UTF8Encoding(false),
            StandardOutputEncoding = new UTF8Encoding(false),
            StandardErrorEncoding = new UTF8Encoding(false),
        };
        start.Environment["DOCBRIDGE_HOME"] = _homeDir;
        _mcp = Process.Start(start) ?? throw new InvalidOperationException("failed to start MCP");
        _mcpStderr = Task.Run(() =>
        {
            try
            {
                while (_mcp is { HasExited: false } && _mcp.StandardError.ReadLine() is { } line)
                    lock (_mcpErr) _mcpErr.AppendLine(line);
            }
            catch { /* process ended */ }
        });

        WriteRpc(new JsonObject
        {
            ["jsonrpc"] = "2.0",
            ["id"] = 0,
            ["method"] = "initialize",
            ["params"] = new JsonObject
            {
                ["protocolVersion"] = "2025-06-18",
                ["capabilities"] = new JsonObject(),
                ["clientInfo"] = new JsonObject { ["name"] = "excel-production-workflow-probe", ["version"] = "1" },
            },
        });
        var init = ReadMatchingResponse(0, TimeSpan.FromSeconds(30));
        if (init is null) throw new InvalidOperationException("MCP initialize failed");
        _mcp.StandardInput.WriteLine("""{"jsonrpc":"2.0","method":"notifications/initialized"}""");
        _mcp.StandardInput.Flush();
    }

    private void WriteRpc(JsonObject payload)
    {
        _mcp!.StandardInput.WriteLine(payload.ToJsonString(JsonUtil.WriteCompact));
        _mcp.StandardInput.Flush();
    }

    private static bool ContainsToken(JsonNode? node)
    {
        if (node is JsonObject o)
        {
            foreach (var (k, v) in o)
            {
                if (k.Contains("confirmToken", StringComparison.OrdinalIgnoreCase)) return true;
                if (ContainsToken(v)) return true;
            }
        }
        else if (node is JsonArray a)
        {
            foreach (var item in a)
                if (ContainsToken(item)) return true;
        }
        return false;
    }

    public JsonArray TimingLedger()
    {
        var a = new JsonArray();
        foreach (var c in _calls)
        {
            a.Add(new JsonObject
            {
                ["id"] = c.Id,
                ["tool"] = c.Tool,
                ["elapsedMs"] = c.ElapsedMs,
                ["exitCode"] = c.ExitCode,
                ["ok"] = JsonUtil.Bool(c.Response, "ok") ?? c.ExitCode == 0,
            });
        }
        return a;
    }

    public void Dispose()
    {
        if (_mcp is { HasExited: false })
        {
            try { _mcp.StandardInput.Close(); } catch { /* ignore */ }
            _mcp.WaitForExit(5_000);
        }
        _mcp?.Dispose();
    }
}
