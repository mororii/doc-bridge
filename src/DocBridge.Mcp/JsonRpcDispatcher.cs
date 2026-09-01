using System.Text.Json;
using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Mcp;

/// <summary>
/// JSON-RPC 2.0 디스패처 (MCP). stdio(NDJSON)/HTTP 양쪽에서 공유한다.
/// notification(id 멤버 없음)은 응답하지 않는다. 로그는 절대 이 클래스의 출력에 섞지 않는다.
///
/// 클라이언트 호환:
///   - protocolVersion 협상: 클라이언트가 보낸 버전이 지원 목록에 있으면 그대로 echo,
///     없으면 서버 기본(ProtocolVersion)을 돌려준다. Claude Desktop(2024-11-05),
///     Codex/Kimi CLI(2025-03-26 / 2025-06-18) 모두 수용된다.
///   - resources/prompts 를 쓰지 않지만, 일부 클라이언트가 기동 시 무조건 조회하므로
///     빈 목록으로 정상 응답한다 (-32601로 실패시키면 연결을 끊는 클라이언트가 있다).
///   - JSON-RPC 배치(배열) 요청 지원.
/// </summary>
public sealed class JsonRpcDispatcher
{
    /// <summary>서버 기본(선호) 프로토콜 버전.</summary>
    public const string ProtocolVersion = "2025-06-18";

    /// <summary>협상 시 그대로 수용하는 버전 목록 (최신 우선).</summary>
    public static readonly string[] SupportedProtocolVersions =
    {
        "2025-06-18",
        "2025-03-26",
        "2024-11-05",
    };

    private const string Instructions =
        "doc-bridge는 실행 중인 Excel/한글(HWP)/AutoCAD 문서를 읽고 쓴다. " +
        "앱별 작업 전 core_get_status를 먼저 호출한다. Excel은 apps.excel.connected=true이고 document가 비어 있지 않을 때만 excel_get_active_context를 한 번 호출한다. " +
        "Excel이 닫혔거나 workbook이 없으면 context를 실행 probe로 호출하거나 같은 실패를 반복하지 않는다. " +
        "DocBridge 오류나 제약을 openpyxl 파일 덮어쓰기, pywin32/직접 Excel COM, PowerShell Excel COM, Start-Process/쉘/UI 자동화로 우회하지 말고 오류와 필요한 사용자 조치를 그대로 보고한다. " +
        "allowOpenFile은 기본 false이며, 사용자가 닫힌 기존 파일을 열어 읽으라고 명시하고 절대 workbook 경로를 제공한 경우에만 true로 쓸 수 있다. 쓰기에는 사용할 수 없다. " +
        "쓰기(*_apply_ops)는 반드시 dryRun=true로 먼저 호출해 diff와 confirmToken을 받고, " +
        "사용자에게 diff를 보여 승인받은 뒤 같은 ops를 dryRun=false + confirmToken으로 재호출한다. " +
        "confirmToken은 5분 TTL·1회용이며 ops 내용에 바인딩되어 있어 ops를 바꾸면 무효가 된다. " +
        "delete_entities/run_script_template 같은 고위험 op는 highRiskConfirm=true가 추가로 필요하다. " +
        "Excel 쓰기는 활성 시트를 추정하지 말고 각 시트 범위 op에 target.sheet 또는 '시트 이름'!A1 형식의 range를 사용한다. " +
        "한글(HWP)은 hwp_get_active_context.summary.openDocuments로 모든 표시 창과 탭을 확인한다. " +
        "여러 문서 중 하나를 읽거나 편집할 때는 반환된 documentRef를 사용하고, 디스크 파일 작업만 file 절대 경로를 사용한다.";

    private const string InteractionInstructions =
        " DocBridge preserves the user's foreground application and restores internal document/view state after COM work. " +
        "The user may work in another application, but must not edit the same target Excel/HWP/AutoCAD window concurrently. " +
        "If interaction.interrupted or userActivityDetected is true, reread the document and create a new dry-run for only the remaining work.";

    private readonly ToolRegistry _tools;
    private readonly string _serverVersion;

    public JsonRpcDispatcher(ToolRegistry tools, string serverVersion)
    {
        _tools = tools;
        _serverVersion = serverVersion;
    }

    /// <summary>
    /// 클라이언트가 요청한 버전을 협상한다. 지원 목록에 있으면 그대로, 없으면 서버 기본값.
    /// 상태를 두지 않는다 — HTTP 모드에서 한 인스턴스가 서로 다른 버전의 클라이언트를
    /// 동시에 상대할 수 있기 때문이다 (예: Claude 2024-11-05 + Codex 2025-06-18).
    /// </summary>
    public static string NegotiateProtocolVersion(string? requested) =>
        requested is not null && SupportedProtocolVersions.Contains(requested)
            ? requested
            : ProtocolVersion;

    /// <summary>요청 객체(initialize)에서 협상 결과 버전을 뽑는다. initialize가 아니면 서버 기본값.</summary>
    public static string NegotiateProtocolVersion(JsonObject? request) =>
        Json.GetString(request, "method") == "initialize"
            ? NegotiateProtocolVersion(Json.GetString(request?["params"] as JsonObject, "protocolVersion"))
            : ProtocolVersion;

    /// <summary>요청 한 건 처리. notification이면 null 반환 (응답 전송 금지).</summary>
    public JsonObject? Dispatch(JsonObject req)
    {
        var hasId = req.ContainsKey("id");
        var id = req["id"];
        var method = Json.GetString(req, "method");

        if (method is null)
            return hasId ? Error(id, -32600, "invalid request: missing 'method'") : null;

        // notification (id 멤버 없음): 조용히 무시 (notifications/initialized, cancelled 등)
        if (!hasId)
            return null;

        try
        {
            return method switch
            {
                "initialize" => Result(id, Initialize(req["params"] as JsonObject)),
                "ping" => Result(id, new JsonObject()),
                "tools/list" => Result(id, new JsonObject { ["tools"] = _tools.ListSpec() }),
                "tools/call" => Result(id, CallTool(req["params"] as JsonObject)),

                // 미사용 기능 — 빈 목록으로 정상 응답해 클라이언트 기동 실패를 막는다.
                "resources/list" => Result(id, new JsonObject { ["resources"] = new JsonArray() }),
                "resources/templates/list" => Result(id, new JsonObject { ["resourceTemplates"] = new JsonArray() }),
                "prompts/list" => Result(id, new JsonObject { ["prompts"] = new JsonArray() }),
                "logging/setLevel" => Result(id, new JsonObject()),

                _ => Error(id, -32601, $"method not found: {method}"),
            };
        }
        catch (Exception ex)
        {
            return Error(id, -32603, $"internal error: {ex.Message}");
        }
    }

    private JsonObject Initialize(JsonObject? p)
    {
        return new JsonObject
        {
            ["protocolVersion"] = NegotiateProtocolVersion(Json.GetString(p, "protocolVersion")),
            ["capabilities"] = new JsonObject
            {
                ["tools"] = new JsonObject { ["listChanged"] = false },
            },
            ["serverInfo"] = new JsonObject
            {
                ["name"] = "doc-bridge",
                ["title"] = "doc-bridge (Excel / 한글 / AutoCAD)",
                ["version"] = _serverVersion,
            },
            ["instructions"] = Instructions + InteractionInstructions,
        };
    }

    private JsonObject CallTool(JsonObject? p)
    {
        var name = Json.GetString(p, "name") ?? throw new InvalidOperationException("tools/call requires params.name");
        var tool = _tools.Find(name)
            ?? throw new InvalidOperationException($"unknown tool: {name}");
        var args = p?["arguments"] as JsonObject ?? new JsonObject();

        JsonObject output;
        try
        {
            output = tool.Handler(args);
        }
        catch (Exception ex)
        {
            output = Json.ErrorResult($"{name} failed: {ex.Message}");
        }

        return new JsonObject
        {
            ["content"] = new JsonArray(new JsonObject
            {
                ["type"] = "text",
                ["text"] = Json.ToCompact(output),
            }),
            // 2025-06-18 이상 클라이언트가 JSON을 그대로 쓸 수 있도록 함께 제공한다.
            ["structuredContent"] = output.DeepClone(),
            ["isError"] = !Json.GetBool(output, "ok"),
        };
    }

    private static JsonObject Result(JsonNode? id, JsonObject result) => new()
    {
        ["jsonrpc"] = "2.0",
        ["id"] = id?.DeepClone(),
        ["result"] = result,
    };

    private static JsonObject Error(JsonNode? id, int code, string message) => new()
    {
        ["jsonrpc"] = "2.0",
        ["id"] = id?.DeepClone(),
        ["error"] = new JsonObject { ["code"] = code, ["message"] = message },
    };

    /// <summary>
    /// 요청 노드(단건 객체 또는 배치 배열) 처리.
    /// 배치는 응답이 있는 항목만 모아 배열로 반환하고, 전부 notification이면 null.
    /// </summary>
    public JsonNode? DispatchNode(JsonNode? node)
    {
        switch (node)
        {
            case JsonObject obj:
                return Dispatch(obj);

            case JsonArray batch:
            {
                if (batch.Count == 0)
                    return Error(null, -32600, "invalid request: empty batch");

                var responses = new JsonArray();
                foreach (var item in batch.ToList())
                {
                    var res = item is JsonObject o
                        ? Dispatch(o)
                        : Error(null, -32600, "invalid request: batch item is not an object");
                    if (res is not null) responses.Add(res);
                }
                return responses.Count == 0 ? null : responses;
            }

            default:
                return Error(null, -32600, "invalid request: expected JSON object or array");
        }
    }

    /// <summary>한 줄(NDJSON) 처리. 파싱 불가 시 parse error 응답 반환.</summary>
    public JsonNode? DispatchLine(string line)
    {
        JsonNode? req;
        try { req = JsonNode.Parse(line); }
        catch (JsonException)
        {
            return Error(null, -32700, "parse error: line is not valid JSON");
        }
        return DispatchNode(req);
    }
}
