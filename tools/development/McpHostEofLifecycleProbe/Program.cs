using System.Diagnostics;
using System.Text;
using System.Text.Json.Nodes;

namespace DocBridge.Development.McpHostEofLifecycleProbe;

/// <summary>
/// Production MCP-host EOF lifecycle probe. Distinct from ExcelOwnerCrashProbe
/// (no watchdog hang, no process kill, no Office taskkill).
/// Requires DOCBRIDGE_E2E=1. Starts an already-built doc-bridge-mcp, talks NDJSON,
/// closes stdin, and records host exit plus Excel PID inventory. Never kills Office.
/// </summary>
internal static class Program
{
    public static int Main(string[] args)
    {
        if (args.Any(a => a is "-h" or "--help" or "/?"))
        {
            Console.WriteLine("""
                DocBridge.McpHostEofLifecycleProbe
                Requires DOCBRIDGE_E2E=1. Does not launch Excel. Does not kill Office.

                --mcp <exe>        required. Already-built doc-bridge-mcp path
                --output <dir>     required. JSON written here (not user Documents)
                --timeout-ms <n>   optional wait after stdin EOF (default 20000)

                This is a production MCP stdin-EOF fixture, not a COM owner-crash probe.
                """);
            return 0;
        }

        if (!string.Equals(Environment.GetEnvironmentVariable("DOCBRIDGE_E2E"), "1", StringComparison.Ordinal))
        {
            Console.Error.WriteLine("refusing: set DOCBRIDGE_E2E=1 to run this probe");
            return 2;
        }

        string? mcp = null;
        string? outputDir = null;
        var timeoutMs = 20_000;
        for (var i = 0; i < args.Length; i++)
        {
            if (args[i] == "--mcp" && i + 1 < args.Length) mcp = args[++i];
            else if (args[i] == "--output" && i + 1 < args.Length) outputDir = args[++i];
            else if (args[i] == "--timeout-ms" && i + 1 < args.Length) timeoutMs = int.Parse(args[++i]);
            else
            {
                Console.Error.WriteLine($"unknown argument: {args[i]}");
                return 2;
            }
        }

        if (string.IsNullOrWhiteSpace(mcp) || !File.Exists(mcp))
        {
            Console.Error.WriteLine("--mcp must point to an existing doc-bridge-mcp executable");
            return 2;
        }
        if (string.IsNullOrWhiteSpace(outputDir))
        {
            Console.Error.WriteLine("--output <dir> is required");
            return 2;
        }

        outputDir = Path.GetFullPath(outputDir);
        var docs = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
        if (!string.IsNullOrWhiteSpace(docs) &&
            outputDir.StartsWith(docs.TrimEnd(Path.DirectorySeparatorChar), StringComparison.OrdinalIgnoreCase))
        {
            Console.Error.WriteLine("refusing: --output must not be a user Documents path");
            return 2;
        }

        Directory.CreateDirectory(outputDir);
        var resultPath = Path.Combine(outputDir, "mcp-host-eof-result.json");
        if (File.Exists(resultPath))
        {
            Console.Error.WriteLine($"refusing to overwrite existing path: {resultPath}");
            return 2;
        }

        var homeDir = Path.Combine(outputDir, "probe-home");
        Directory.CreateDirectory(homeDir);
        var before = InventoryExcel();
        var report = new JsonObject
        {
            ["ok"] = false,
            ["kind"] = "production-mcp-host-eof",
            ["distinctFrom"] = "ExcelOwnerCrashProbe",
            ["mcp"] = Path.GetFullPath(mcp),
            ["outputDir"] = outputDir,
            ["homeDir"] = homeDir,
            ["excelBefore"] = before,
            ["killedOffice"] = false,
        };

        try
        {
            var start = new ProcessStartInfo
            {
                FileName = Path.GetFullPath(mcp),
                Arguments = "--stdio",
                WorkingDirectory = Path.GetDirectoryName(Path.GetFullPath(mcp)),
                UseShellExecute = false,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true,
                StandardInputEncoding = new UTF8Encoding(false),
                StandardOutputEncoding = new UTF8Encoding(false),
                StandardErrorEncoding = new UTF8Encoding(false),
                CreateNoWindow = true,
            };
            start.Environment["DOCBRIDGE_HOME"] = homeDir;

            using var process = Process.Start(start)
                ?? throw new InvalidOperationException("failed to start MCP host");
            report["hostPid"] = process.Id;

            var stderr = new StringBuilder();
            process.ErrorDataReceived += (_, e) => { if (e.Data is not null) stderr.AppendLine(e.Data); };
            process.BeginErrorReadLine();

            WriteLine(process, Rpc(1, "initialize", new JsonObject
            {
                ["protocolVersion"] = "2025-06-18",
                ["capabilities"] = new JsonObject(),
                ["clientInfo"] = new JsonObject { ["name"] = "mcp-host-eof-probe", ["version"] = "0" },
            }));
            var initialize = ReadLine(process, 10_000);
            WriteLine(process, """{"jsonrpc":"2.0","method":"notifications/initialized"}""");
            WriteLine(process, Rpc(2, "tools/call", new JsonObject
            {
                ["name"] = "core_get_status",
                ["arguments"] = new JsonObject(),
            }));
            var status = ReadLine(process, 20_000);
            WriteLine(process, Rpc(3, "tools/call", new JsonObject
            {
                ["name"] = "excel_get_active_context",
                ["arguments"] = new JsonObject(),
            }));
            var excelCtx = ReadLine(process, 20_000);

            report["initialize"] = Compact(initialize);
            report["core_get_status"] = Compact(status);
            report["excel_get_active_context"] = Compact(excelCtx);
            report["excelAfterStatus"] = InventoryExcel();
            report["createdExcelDuringStatus"] = !SamePids(before, report["excelAfterStatus"] as JsonArray);

            process.StandardInput.Close();
            var exited = process.WaitForExit(timeoutMs);
            report["stdinClosed"] = true;
            report["exited"] = exited;
            report["exitCode"] = exited ? process.ExitCode : null;
            report["stderr"] = stderr.ToString();
            if (!exited)
            {
                try { process.Kill(entireProcessTree: false); } catch { /* host only; never Excel */ }
                report["hostKillAfterTimeout"] = true;
            }

            report["excelAfterEof"] = InventoryExcel();
            report["preexistingPidsStillPresent"] = PreexistingStillPresent(before, report["excelAfterEof"] as JsonArray);
            report["ok"] = exited && GetBoolish(report["preexistingPidsStillPresent"]) &&
                           !GetBoolish(report["createdExcelDuringStatus"]);
            return GetBoolish(report["ok"]) ? 0 : 4;
        }
        catch (Exception ex)
        {
            report["ok"] = false;
            report["harnessError"] = ex.Message;
            Console.Error.WriteLine(ex);
            return 3;
        }
        finally
        {
            File.WriteAllText(resultPath, report.ToJsonString(new System.Text.Json.JsonSerializerOptions { WriteIndented = true }));
            Console.WriteLine(resultPath);
        }
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

    private static bool SamePids(JsonArray before, JsonArray? after)
    {
        if (after is null) return false;
        return PidSet(before).SetEquals(PidSet(after));
    }

    private static bool PreexistingStillPresent(JsonArray before, JsonArray? after)
    {
        if (after is null) return false;
        var afterPids = PidSet(after);
        return PidSet(before).All(afterPids.Contains);
    }

    private static HashSet<int> PidSet(JsonArray inventory)
    {
        var set = new HashSet<int>();
        foreach (var node in inventory)
            if (node is JsonObject o && o["pid"] is JsonValue v && v.TryGetValue<int>(out var pid))
                set.Add(pid);
        return set;
    }

    private static void WriteLine(Process process, string line)
    {
        process.StandardInput.WriteLine(line);
        process.StandardInput.Flush();
    }

    private static string ReadLine(Process process, int timeoutMs)
    {
        var task = process.StandardOutput.ReadLineAsync();
        if (!task.Wait(timeoutMs))
            throw new TimeoutException("timed out waiting for MCP stdout");
        return task.Result ?? throw new InvalidOperationException("MCP stdout EOF");
    }

    private static string Rpc(int id, string method, JsonObject? @params)
    {
        var req = new JsonObject
        {
            ["jsonrpc"] = "2.0",
            ["id"] = id,
            ["method"] = method,
        };
        if (@params is not null) req["params"] = @params;
        return req.ToJsonString();
    }

    private static JsonNode Compact(string raw)
    {
        try { return JsonNode.Parse(raw) ?? JsonValue.Create(raw)!; }
        catch { return JsonValue.Create(raw)!; }
    }

    private static bool GetBoolish(JsonNode? node) =>
        node is JsonValue v && v.TryGetValue<bool>(out var b) && b;
}
