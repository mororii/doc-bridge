using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// 감사 로그 (보안 원칙 7: stdio MCP 서버는 stdout에 로그 금지 → 파일 JSONL).
/// {RootDir}/logs/audit-yyyyMMdd.jsonl, named mutex로 크로스 프로세스 직렬화.
/// </summary>
public sealed class AuditLog
{
    private const string MutexName = @"Global\DocBridge.Audit";
    private readonly string _logsDir;

    public AuditLog(Models.DocBridgeOptions options)
    {
        _logsDir = options.LogsDir;
        Directory.CreateDirectory(_logsDir);
    }

    public void Write(string tool, string? app, string action, JsonObject? detail = null, bool ok = true, IEnumerable<string>? errors = null)
    {
        var entry = new JsonObject
        {
            ["ts"] = DateTimeOffset.Now.ToString("o"),
            ["tool"] = tool,
            ["app"] = app,
            ["action"] = action,      // "read" | "dry_run" | "apply" | "restore" | "snapshot" | "deny"
            ["ok"] = ok,
            ["detail"] = detail?.DeepClone(),
            ["errors"] = Json.ToArray(errors ?? Array.Empty<string>()),
            ["pid"] = Environment.ProcessId,
        };
        var line = entry.ToJsonString(Json.Compact);
        var path = Path.Combine(_logsDir, $"audit-{DateTimeOffset.Now:yyyyMMdd}.jsonl");

        try
        {
            using var mutex = new Mutex(false, MutexName);
            mutex.WaitOne(TimeSpan.FromSeconds(10));
            try { File.AppendAllText(path, line + Environment.NewLine); }
            finally { mutex.ReleaseMutex(); }
        }
        catch
        {
            // 감사 로그 실패가 본 작업을 중단시키면 안 되지만, 최후 수단으로 stderr에 남긴다 (stdout 금지)
            Console.Error.WriteLine($"[doc-bridge] audit log write failed: {path}");
        }
    }
}
