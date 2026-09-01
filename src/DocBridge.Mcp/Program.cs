using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;
using DocBridge.Mcp;
using Microsoft.AspNetCore.Builder;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using Microsoft.Extensions.Logging;

// doc-bridge MCP 서버 (명령서 §4, §6)
//   --stdio (기본)     : NDJSON stdio. stdout은 프로토콜 전용, 로그는 stderr.
//   --http [--port N]  : Streamable HTTP (POST /mcp, GET /health)
//
// 클라이언트: Claude Desktop / Codex / Kimi CLI 모두 stdio 방식을 우선 권장한다.

if (ExcelWorkerProcess.TryRun(args, out var excelWorkerExitCode))
    return excelWorkerExitCode;
if (ExcelOwnerWatchdog.TryRun(args, out var watchdogExitCode))
    return watchdogExitCode;

if (args.Contains("--help") || args.Contains("-h"))
{
    Console.Error.WriteLine("doc-bridge-mcp [--stdio | --http [--port 5177]] [--version]");
    return 0;
}
if (args.Contains("--version"))
{
    Console.Error.WriteLine(DocBridgeHost.Version);
    return 0;
}

var http = args.Contains("--http");
var portArg = args.SkipWhile(a => a != "--port").Skip(1).FirstOrDefault();
var port = int.TryParse(portArg, out var p) ? p : 5177;

// 어댑터는 SessionRouter가 필요할 때 지연 생성한다 (excel/hwp/cad/fake 팩터리 내장).
// 시작 시 미리 만들면 COM/STA 스레드가 즉시 생기고, 앱 하나가 없어도 서버 전체가 죽는다.
using var host = new DocBridgeHost();
using var shutdownCleanup = new ShutdownCleanupRegistration(host.Dispose);

var tools = new ToolRegistry(host);
var dispatcher = new JsonRpcDispatcher(tools, DocBridgeHost.Version);

// MCP stdio는 UTF-8이 표준. Console.InputEncoding/OutputEncoding 프로퍼티는 콘솔이 없는
// 상태(클라이언트가 파이프로 띄운 경우)에서 IOException을 던질 수 있으므로 쓰지 않고,
// 표준 스트림을 UTF-8(BOM 없음)로 직접 감싼다.
var utf8 = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false);
var stderr = new StreamWriter(Console.OpenStandardError(), utf8) { AutoFlush = true };

if (!http)
{
    var stdin = new StreamReader(Console.OpenStandardInput(), utf8, detectEncodingFromByteOrderMarks: false);
    var stdout = new StreamWriter(Console.OpenStandardOutput(), utf8) { AutoFlush = false, NewLine = "\n" };

    stderr.WriteLine($"[doc-bridge] MCP stdio server v{DocBridgeHost.Version} ({tools.All.Count} tools)");
    var server = new McpServer(dispatcher, stdin, stdout, stderr);
    await server.RunAsync();
    await stdout.FlushAsync();
    return 0;
}

// ---------- HTTP 모드 (Streamable HTTP, 명령서 M4) ----------
// 127.0.0.1에만 바인딩한다. DOCBRIDGE_HTTP_TOKEN 환경변수를 설정하면 Bearer 인증을 요구한다.
var authToken = Environment.GetEnvironmentVariable("DOCBRIDGE_HTTP_TOKEN");

// args를 넘기지 않는다 — "--http" 처럼 값 없는 플래그가 명령줄 구성 공급자와 충돌할 수 있다.
var builder = WebApplication.CreateBuilder();
builder.Logging.ClearProviders();
builder.Logging.SetMinimumLevel(LogLevel.Warning);
builder.WebHost.UseUrls($"http://127.0.0.1:{port}");
var app = builder.Build();

static JsonObject RpcError(int code, string message) => new()
{
    ["jsonrpc"] = "2.0",
    ["id"] = null,
    ["error"] = new JsonObject { ["code"] = code, ["message"] = message },
};

bool Authorized(HttpContext ctx)
{
    if (string.IsNullOrEmpty(authToken)) return true;
    var header = ctx.Request.Headers.Authorization.ToString();
    return header.StartsWith("Bearer ", StringComparison.OrdinalIgnoreCase)
           && string.Equals(header[7..].Trim(), authToken, StringComparison.Ordinal);
}

app.MapPost("/mcp", async (HttpContext ctx) =>
{
    if (!Authorized(ctx))
    {
        ctx.Response.StatusCode = 401;
        await ctx.Response.WriteAsJsonAsync(RpcError(-32001, "unauthorized"));
        return;
    }

    JsonNode? req;
    try
    {
        // leaveOpen: true — reader를 dispose해도 요청 본문 스트림을 닫지 않는다.
        using var reader = new StreamReader(ctx.Request.Body, utf8,
            detectEncodingFromByteOrderMarks: false, bufferSize: -1, leaveOpen: true);
        var body = await reader.ReadToEndAsync(ctx.RequestAborted);
        req = string.IsNullOrWhiteSpace(body) ? null : JsonNode.Parse(body);
    }
    catch (JsonException)
    {
        ctx.Response.StatusCode = 400;
        await ctx.Response.WriteAsJsonAsync(RpcError(-32700, "parse error"));
        return;
    }

    if (req is null)
    {
        ctx.Response.StatusCode = 400;
        await ctx.Response.WriteAsJsonAsync(RpcError(-32600, "invalid request: empty body"));
        return;
    }

    // Streamable HTTP: initialize 응답에 세션 ID를 실어 준다 (클라이언트가 이후 헤더로 되돌려 보냄).
    // 프로토콜 버전 헤더는 이 요청 하나만 보고 정한다 — 서로 다른 버전의 클라이언트가
    // 동시에 붙어도 남의 버전을 받지 않도록 인스턴스 상태를 쓰지 않는다.
    if (req is JsonObject reqObj && Json.GetString(reqObj, "method") == "initialize")
        ctx.Response.Headers["Mcp-Session-Id"] = Guid.NewGuid().ToString("N");

    // Streamable HTTP에서 클라이언트는 협상한 버전을 요청 헤더로 되돌려 보낸다. 그게 있으면 그걸 쓰고,
    // 없으면(첫 initialize) 요청 본문에서 뽑는다.
    var clientProtocol = ctx.Request.Headers["MCP-Protocol-Version"].ToString();
    ctx.Response.Headers["MCP-Protocol-Version"] = string.IsNullOrEmpty(clientProtocol)
        ? JsonRpcDispatcher.NegotiateProtocolVersion(req as JsonObject)
        : JsonRpcDispatcher.NegotiateProtocolVersion(clientProtocol);

    var response = dispatcher.DispatchNode(req);

    if (response is null)
    {
        ctx.Response.StatusCode = 202; // notification만 있었음 — 본문 없음
        return;
    }

    var payload = Json.ToCompact(response);

    // 클라이언트가 SSE만 받겠다고 하면 SSE로, 아니면 일반 JSON으로 (둘 다 스펙 허용).
    var accept = ctx.Request.Headers.Accept.ToString();
    var wantsSseOnly = accept.Contains("text/event-stream", StringComparison.OrdinalIgnoreCase)
                       && !accept.Contains("application/json", StringComparison.OrdinalIgnoreCase);

    if (wantsSseOnly)
    {
        ctx.Response.ContentType = "text/event-stream";
        ctx.Response.Headers.CacheControl = "no-cache";
        await ctx.Response.WriteAsync($"event: message\ndata: {payload}\n\n", ctx.RequestAborted);
        await ctx.Response.Body.FlushAsync(ctx.RequestAborted);
        return;
    }

    ctx.Response.ContentType = "application/json; charset=utf-8";
    await ctx.Response.WriteAsync(payload, ctx.RequestAborted);
});

// 서버 → 클라이언트 단방향 스트림은 쓰지 않는다. 스펙상 405가 정상 응답이다.
app.MapGet("/mcp", (HttpContext ctx) =>
{
    ctx.Response.Headers.Allow = "POST, DELETE";
    return Results.Json(RpcError(-32000, "SSE stream not supported; use POST /mcp"), statusCode: 405);
});

// 세션 종료 — 서버가 세션 상태를 보관하지 않으므로 항상 성공.
app.MapDelete("/mcp", () => Results.StatusCode(204));

app.MapGet("/health", () => Results.Json(new JsonObject
{
    ["ok"] = true,
    ["version"] = DocBridgeHost.Version,
    ["protocolVersion"] = JsonRpcDispatcher.ProtocolVersion,
    ["tools"] = tools.All.Count,
}));

stderr.WriteLine($"[doc-bridge] MCP http server on http://127.0.0.1:{port}/mcp"
                 + (string.IsNullOrEmpty(authToken) ? "" : " (Bearer 인증 필요)"));
await app.RunAsync();
return 0;
