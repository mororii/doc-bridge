using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

namespace DocBridge.Mcp.Tests;

/// <summary>
/// MCP 핸드셰이크/stdio 순수성 테스트 (M4 인수 조건).
/// Office 없이 돌아가도록 FakeAdapter를 excel/hwp/cad 이름으로 등록한 host를 쓴다.
/// </summary>
public class McpHandshakeTests : IDisposable
{
    private readonly TestHome _home = new();
    private readonly DocBridgeHost _host;
    private readonly JsonRpcDispatcher _dispatcher;

    public McpHandshakeTests()
    {
        _host = new DocBridgeHost(_home.Options);
        _host.Router.Register("excel", new FakeAdapter());
        _host.Router.Register("hwp", new FakeAdapter());
        _host.Router.Register("cad", new FakeAdapter());
        _dispatcher = new JsonRpcDispatcher(new ToolRegistry(_host), DocBridgeHost.Version);
    }

    public void Dispose() { _host.Dispose(); _home.Dispose(); }

    private static JsonObject Req(string method, object? id = null, JsonObject? p = null)
    {
        var o = new JsonObject { ["jsonrpc"] = "2.0", ["method"] = method };
        if (id is not null) o["id"] = JsonValue.Create(id);
        if (p is not null) o["params"] = p;
        return o;
    }

    private static JsonObject? ResultOf(JsonObject? res) => res?["result"] as JsonObject;

    [Fact]
    public void Initialize_returns_protocol_and_serverInfo()
    {
        var res = _dispatcher.Dispatch(Req("initialize", 1));
        var result = ResultOf(res)!;
        Assert.Equal(JsonRpcDispatcher.ProtocolVersion, Json.GetString(result, "protocolVersion"));
        Assert.Equal("doc-bridge", Json.GetString(Json.GetObj(result, "serverInfo"), "name"));
        Assert.NotNull(Json.GetObj(Json.GetObj(result, "capabilities"), "tools"));
        var instructions = Json.GetString(result, "instructions")!;
        Assert.Contains("openDocuments", instructions);
        Assert.Contains("documentRef", instructions);
        Assert.DoesNotContain("실행 중 창에 붙을 수 없", instructions);
    }

    [Fact]
    public void Tools_list_exposes_all_25_underscore_tools()
    {
        var res = _dispatcher.Dispatch(Req("tools/list", 2));
        var tools = Json.GetArr(ResultOf(res), "tools")!;
        Assert.Equal(25, tools.Count);

        var names = tools.Select(t => Json.GetString(t as JsonObject, "name")!).ToHashSet();
        foreach (var want in new[]
        {
            "core_ping", "core_get_status", "core_get_capabilities", "core_disconnect", "core_create_snapshot", "core_list_snapshots", "core_restore_snapshot",
            "excel_get_active_context", "excel_read_range", "excel_inspect", "excel_apply_ops", "excel_disconnect",
            "hwp_plan_creation", "hwp_launch", "hwp_get_active_context", "hwp_doctor", "hwp_repair_typelib", "hwp_read_text", "hwp_apply_ops", "hwp_submit_ops", "hwp_get_job",
            "cad_launch", "cad_get_active_context", "cad_query_entities", "cad_apply_ops",
        })
            Assert.Contains(want, names);

        Assert.All(names, n => Assert.DoesNotContain('.', n)); // 밑줄 명명 규칙
        Assert.All(tools, t => Assert.NotNull(Json.GetObj(t as JsonObject, "inputSchema")));
        Assert.All(tools, t => Assert.NotNull(Json.GetObj(t as JsonObject, "annotations")));

        var readTool = tools.Select(t => t as JsonObject).First(t => Json.GetString(t, "name") == "excel_read_range")!;
        Assert.True(Json.GetBool(Json.GetObj(readTool, "annotations"), "readOnlyHint"));
        Assert.False(Json.GetBool(Json.GetObj(readTool, "annotations"), "destructiveHint"));

        var writeTool = tools.Select(t => t as JsonObject).First(t => Json.GetString(t, "name") == "excel_apply_ops")!;
        Assert.False(Json.GetBool(Json.GetObj(writeTool, "annotations"), "readOnlyHint"));
        Assert.True(Json.GetBool(Json.GetObj(writeTool, "annotations"), "destructiveHint"));

        var hwpSubmitTool = tools.Select(t => t as JsonObject).First(t => Json.GetString(t, "name") == "hwp_submit_ops")!;
        Assert.False(Json.GetBool(Json.GetObj(hwpSubmitTool, "annotations"), "readOnlyHint"));
        Assert.True(Json.GetBool(Json.GetObj(hwpSubmitTool, "annotations"), "destructiveHint"));
        Assert.False(Json.GetBool(Json.GetObj(hwpSubmitTool, "annotations"), "idempotentHint"));

        var hwpJobTool = tools.Select(t => t as JsonObject).First(t => Json.GetString(t, "name") == "hwp_get_job")!;
        Assert.True(Json.GetBool(Json.GetObj(hwpJobTool, "annotations"), "readOnlyHint"));
        Assert.False(Json.GetBool(Json.GetObj(hwpJobTool, "annotations"), "destructiveHint"));
        Assert.True(Json.GetBool(Json.GetObj(hwpJobTool, "annotations"), "idempotentHint"));

        var disconnectTool = tools.Select(t => t as JsonObject).First(t => Json.GetString(t, "name") == "excel_disconnect")!;
        Assert.False(Json.GetBool(Json.GetObj(disconnectTool, "annotations"), "readOnlyHint"));
        Assert.False(Json.GetBool(Json.GetObj(disconnectTool, "annotations"), "destructiveHint"));
        Assert.True(Json.GetBool(Json.GetObj(disconnectTool, "annotations"), "idempotentHint"));
        var excelWriteSchema = Json.GetObj(writeTool, "inputSchema")!;
        Assert.Contains("활성 시트", Json.GetString(excelWriteSchema, "description"));
        var excelWriteItems = Json.GetObj(Json.GetObj(Json.GetObj(excelWriteSchema, "properties"), "ops"), "items")!;
        var excelWriteProperties = Json.GetObj(excelWriteItems, "properties")!;
        Assert.NotNull(Json.GetObj(excelWriteProperties, "target"));
        Assert.NotNull(Json.GetObj(excelWriteProperties, "range"));
        Assert.NotNull(Json.GetObj(excelWriteProperties, "values"));
        var excelOpEnum = Json.GetArr(Json.GetObj(excelWriteProperties, "op"), "enum")!;
        var excelOps = excelOpEnum.Select(item => item!.GetValue<string>()).ToHashSet(StringComparer.OrdinalIgnoreCase);
        Assert.True(new[]
        {
            "set_values", "merge_cells", "unmerge_cells", "set_rows_hidden", "set_cols_hidden",
            "set_sheet_visibility",
        }.All(excelOps.Contains));
        Assert.NotNull(Json.GetObj(excelWriteProperties, "hidden"));
        Assert.NotNull(Json.GetObj(excelWriteProperties, "visibility"));

        var excelReadProperties = Json.GetObj(Json.GetObj(readTool, "inputSchema"), "properties")!;
        Assert.NotNull(Json.GetObj(excelReadProperties, "includeLayout"));
        Assert.NotNull(Json.GetObj(excelReadProperties, "allowOpenFile"));

        var cadQueryTool = tools.Select(t => t as JsonObject).First(t => Json.GetString(t, "name") == "cad_query_entities")!;
        var cadScopeEnum = Json.GetArr(Json.GetObj(Json.GetObj(Json.GetObj(cadQueryTool, "inputSchema"), "properties"), "scope"), "enum")!;
        var cadScopes = cadScopeEnum.Select(item => item!.GetValue<string>()).ToHashSet(StringComparer.OrdinalIgnoreCase);
        Assert.True(new[] { "entities", "layouts", "layers", "xrefs", "window", "regions" }.All(cadScopes.Contains));

        var cadContextTool = tools.Select(t => t as JsonObject).First(t => Json.GetString(t, "name") == "cad_get_active_context")!;
        var cadContextProperties = Json.GetObj(Json.GetObj(cadContextTool, "inputSchema"), "properties")!;
        var cadDetailLevel = Json.GetObj(cadContextProperties, "detailLevel")!;
        Assert.Equal("basic", Json.GetString(cadDetailLevel, "default"));
        var cadDetailLevels = Json.GetArr(cadDetailLevel, "enum")!
            .Select(node => node!.GetValue<string>()).ToHashSet(StringComparer.OrdinalIgnoreCase);
        Assert.True(new[] { "basic", "summary" }.All(cadDetailLevels.Contains));
        Assert.Contains("nextActions", Json.GetString(cadDetailLevel, "description"));

        var hwpWriteTool = tools.Select(t => t as JsonObject).First(t => Json.GetString(t, "name") == "hwp_apply_ops")!;
        var hwpDescription = Json.GetString(hwpWriteTool, "description")! + " " +
                             Json.GetString(Json.GetObj(hwpWriteTool, "inputSchema"), "description");
        Assert.Contains("append_text", hwpDescription);
        Assert.Contains("insert_before_text", hwpDescription);
        Assert.Contains("insert_after_text", hwpDescription);
        Assert.Contains("서식", hwpDescription);
        Assert.Contains("다중 문단", hwpDescription);
        Assert.Contains("PowerShell", hwpDescription);
        var hwpItems = Json.GetObj(Json.GetObj(Json.GetObj(hwpWriteTool, "inputSchema"), "properties"), "ops")!["items"] as JsonObject;
        var hwpItemProperties = Json.GetObj(hwpItems, "properties")!;
        Assert.NotNull(Json.GetObj(hwpItemProperties, "text"));
        var hwpOpEnum = Json.GetArr(Json.GetObj(hwpItemProperties, "op"), "enum")!;
        Assert.Contains(hwpOpEnum, item => item!.GetValue<string>() == "append_text");
        Assert.Contains(hwpOpEnum, item => item!.GetValue<string>() == "insert_before_text");
        Assert.Contains(hwpOpEnum, item => item!.GetValue<string>() == "insert_after_text");
        Assert.Contains(hwpOpEnum, item => item!.GetValue<string>() == "table_set_row_height");
        Assert.Contains(hwpOpEnum, item => item!.GetValue<string>() == "table_set_row_heights");
        Assert.Contains(hwpOpEnum, item => item!.GetValue<string>() == "format_paragraphs");
        Assert.Contains(hwpOpEnum, item => item!.GetValue<string>() == "table_set_cells");
        var hwpRows = Json.GetObj(hwpItemProperties, "rows");
        Assert.NotNull(hwpRows);
        Assert.Equal(500, Json.GetInt(hwpRows, "maxItems"));
        Assert.NotNull(Json.GetObj(hwpItemProperties, "items"));
        Assert.NotNull(Json.GetObj(hwpItemProperties, "anchor"));
        Assert.NotNull(Json.GetObj(hwpItemProperties, "occurrence"));
        Assert.NotNull(Json.GetObj(hwpItemProperties, "mode"));
        Assert.NotNull(Json.GetObj(hwpItemProperties, "preserveStyle"));
        Assert.NotNull(Json.GetObj(hwpItemProperties, "styleSource"));
        Assert.NotNull(Json.GetObj(hwpItemProperties, "cellIndex"));
        Assert.NotNull(Json.GetObj(hwpItemProperties, "heightMm"));
        Assert.NotNull(Json.GetObj(hwpItemProperties, "documentRef"));
        Assert.NotNull(Json.GetObj(hwpItemProperties, "count"));
        Assert.NotNull(Json.GetObj(hwpItemProperties, "position"));
        Assert.NotNull(Json.GetObj(hwpItemProperties, "path"));
        Assert.NotNull(Json.GetObj(hwpItemProperties, "sizeOption"));
        Assert.NotNull(Json.GetObj(hwpItemProperties, "clearCell"));
        Assert.NotNull(Json.GetObj(hwpItemProperties, "find"));
        Assert.NotNull(Json.GetObj(hwpItemProperties, "replace"));
        Assert.NotNull(Json.GetObj(hwpItemProperties, "scope"));
        Assert.NotNull(Json.GetObj(hwpItemProperties, "cells"));
        Assert.Contains("순차 시뮬레이션", hwpDescription);

        var hwpLaunchTool = tools.Select(t => t as JsonObject)
            .First(t => Json.GetString(t, "name") == "hwp_launch")!;
        var hwpLaunchDescription = Json.GetString(hwpLaunchTool, "description")! + " " +
                                   Json.GetString(Json.GetObj(hwpLaunchTool, "inputSchema"), "description");
        Assert.Contains("DOCX", hwpLaunchDescription);
        Assert.Contains("OOXML", hwpLaunchDescription);
        Assert.Contains("hwp_plan_creation", hwpLaunchDescription);
        var hwpLaunchProperties = Json.GetObj(Json.GetObj(hwpLaunchTool, "inputSchema"), "properties")!;
        Assert.NotNull(Json.GetObj(hwpLaunchProperties, "creationMode"));
        Assert.NotNull(Json.GetObj(hwpLaunchProperties, "sourceFile"));
        Assert.NotNull(Json.GetObj(hwpLaunchProperties, "outputFile"));
        Assert.NotNull(Json.GetObj(hwpLaunchProperties, "closeAfterImport"));
        Assert.NotNull(Json.GetObj(hwpLaunchProperties, "expectedPageCount"));
        Assert.NotNull(Json.GetObj(hwpLaunchProperties, "expectedTableCount"));
        Assert.NotNull(Json.GetObj(hwpLaunchProperties, "requiredText"));

        var hwpPlanTool = tools.Select(t => t as JsonObject)
            .First(t => Json.GetString(t, "name") == "hwp_plan_creation")!;
        Assert.True(Json.GetBool(Json.GetObj(hwpPlanTool, "annotations"), "readOnlyHint"));
        Assert.False(Json.GetBool(Json.GetObj(hwpPlanTool, "annotations"), "destructiveHint"));
        var hwpPlanProperties = Json.GetObj(Json.GetObj(hwpPlanTool, "inputSchema"), "properties")!;
        Assert.NotNull(Json.GetObj(hwpPlanProperties, "documentState"));
        Assert.NotNull(Json.GetObj(hwpPlanProperties, "requiresNativeFields"));
        Assert.NotNull(Json.GetObj(hwpPlanProperties, "requiresComplexMergedTables"));

        var hwpReadTool = tools.Select(t => t as JsonObject).First(t => Json.GetString(t, "name") == "hwp_read_text")!;
        var hwpReadScope = Json.GetObj(Json.GetObj(Json.GetObj(hwpReadTool, "inputSchema"), "properties"), "scope");
        var hwpReadScopes = Json.GetArr(hwpReadScope, "enum")!.Select(item => item!.GetValue<string>()).ToHashSet();
        Assert.Contains("tables", hwpReadScopes);
        Assert.Contains("document_map", hwpReadScopes);
        Assert.Contains("bundle", hwpReadScopes);
        var hwpReadProperties = Json.GetObj(Json.GetObj(hwpReadTool, "inputSchema"), "properties")!;
        Assert.NotNull(Json.GetObj(hwpReadProperties, "documentRef"));
        var sectionItems = Json.GetObj(Json.GetObj(hwpReadProperties, "sections"), "items");
        var sectionEnum = Json.GetArr(sectionItems, "enum")!;
        Assert.Contains(sectionEnum, item => item!.GetValue<string>() == "tables");

        var excelInspect = tools.Select(t => t as JsonObject).First(t => Json.GetString(t, "name") == "excel_inspect")!;
        Assert.True(Json.GetBool(Json.GetObj(excelInspect, "annotations"), "readOnlyHint"));
        var inspectProperties = Json.GetObj(Json.GetObj(excelInspect, "inputSchema"), "properties")!;
        var inspectScopes = Json.GetArr(Json.GetObj(inspectProperties, "scope"), "enum")!;
        Assert.Contains(inspectScopes, item => item!.GetValue<string>() == "errors");
        Assert.Contains(inspectScopes, item => item!.GetValue<string>() == "diagnostics");
        Assert.NotNull(Json.GetObj(inspectProperties, "allowOpenFile"));

        var hwpRepair = tools.Select(t => t as JsonObject).First(t => Json.GetString(t, "name") == "hwp_repair_typelib")!;
        Assert.False(Json.GetBool(Json.GetObj(hwpRepair, "annotations"), "readOnlyHint"));
        Assert.True(Json.GetBool(Json.GetObj(hwpRepair, "annotations"), "destructiveHint"));
    }

    [Fact]
    public void Tools_call_core_ping_returns_ok_content()
    {
        var res = _dispatcher.Dispatch(Req("tools/call", 3, new JsonObject
        {
            ["name"] = "core_ping",
            ["arguments"] = new JsonObject(),
        }));
        var result = ResultOf(res)!;
        Assert.False(Json.GetBool(result, "isError"));
        var content = Json.GetArr(result, "content")!;
        var payload = JsonNode.Parse(Json.GetString(content[0] as JsonObject, "text")!) as JsonObject;
        Assert.True(Json.GetBool(payload, "ok"));
        var adapters = Json.GetArr(payload, "adapters")!;
        var names = adapters.Select(a => a!.GetValue<string>()).ToHashSet();
        Assert.True(new HashSet<string> { "excel", "hwp", "cad" }.IsSubsetOf(names)); // fake 기본 등록 포함 4개 중 3개 확인
    }

    [Fact]
    public void Tools_call_excel_context_via_fake_adapter()
    {
        var res = _dispatcher.Dispatch(Req("tools/call", 4, new JsonObject
        {
            ["name"] = "excel_get_active_context",
            ["arguments"] = new JsonObject(),
        }));
        var result = ResultOf(res)!;
        Assert.False(Json.GetBool(result, "isError"));
    }

    [Fact]
    public void Tools_call_hwp_plan_creation_returns_the_hybrid_route_without_starting_hwp()
    {
        var res = _dispatcher.Dispatch(Req("tools/call", 41, new JsonObject
        {
            ["name"] = "hwp_plan_creation",
            ["arguments"] = new JsonObject
            {
                ["documentState"] = "new",
                ["requiresNativeFields"] = false,
            },
        }));

        var result = ResultOf(res)!;
        Assert.False(Json.GetBool(result, "isError"));
        var content = Json.GetArr(result, "content")!;
        var payload = JsonNode.Parse(Json.GetString(content[0] as JsonObject, "text")!) as JsonObject;
        Assert.True(Json.GetBool(payload, "ok"));
        Assert.Equal("docx-first", Json.GetString(payload, "mode"));
        Assert.False(Json.GetBool(payload, "wordComRequired"));
    }

    [Fact]
    public void Hwp_async_job_returns_id_then_terminal_result_without_resubmitting()
    {
        var submit = _dispatcher.Dispatch(Req("tools/call", 42, new JsonObject
        {
            ["name"] = "hwp_submit_ops",
            ["arguments"] = new JsonObject
            {
                ["ops"] = new JsonArray(new JsonObject { ["op"] = "insert_text", ["text"] = "job" }),
                ["dryRun"] = true,
            },
        }));
        var submitContent = Json.GetArr(ResultOf(submit), "content")!;
        var submitted = JsonNode.Parse(Json.GetString(submitContent[0] as JsonObject, "text")!) as JsonObject;
        var jobId = Json.GetString(submitted, "jobId");
        Assert.False(string.IsNullOrWhiteSpace(jobId));
        Assert.False(Json.GetBool(submitted, "safeToRetrySubmit", true));

        JsonObject? job = null;
        for (var attempt = 0; attempt < 50; attempt++)
        {
            var poll = _dispatcher.Dispatch(Req("tools/call", 43, new JsonObject
            {
                ["name"] = "hwp_get_job",
                ["arguments"] = new JsonObject { ["jobId"] = jobId },
            }));
            var pollContent = Json.GetArr(ResultOf(poll), "content")!;
            job = JsonNode.Parse(Json.GetString(pollContent[0] as JsonObject, "text")!) as JsonObject;
            if (Json.GetBool(job, "terminal")) break;
            Thread.Sleep(10);
        }
        Assert.NotNull(job);
        Assert.Equal("succeeded", Json.GetString(job, "status"));
        Assert.True(Json.GetBool(Json.GetObj(job, "result"), "ok"));
    }

    [Fact]
    public void Unknown_tool_and_unknown_method_are_jsonrpc_errors()
    {
        var badTool = _dispatcher.Dispatch(Req("tools/call", 5, new JsonObject { ["name"] = "nope" }))!;
        Assert.Equal(-32603, Json.GetInt(Json.GetObj(badTool, "error"), "code"));

        var badMethod = _dispatcher.Dispatch(Req("bogus/method", 6))!;
        Assert.Equal(-32601, Json.GetInt(Json.GetObj(badMethod, "error"), "code"));
    }

    [Fact]
    public void Notifications_get_no_response()
    {
        Assert.Null(_dispatcher.Dispatch(Req("notifications/initialized")));
        Assert.Null(_dispatcher.Dispatch(Req("initialize"))); // id 없으면 notification 취급
    }

    // ---------- 클라이언트 호환 (Kimi / Codex / Claude) ----------

    [Theory]
    [InlineData("2024-11-05")] // Claude Desktop
    [InlineData("2025-03-26")]
    [InlineData("2025-06-18")] // Codex / Kimi CLI
    public void Initialize_echoes_supported_client_protocol_version(string clientVersion)
    {
        var res = _dispatcher.Dispatch(Req("initialize", 1, new JsonObject
        {
            ["protocolVersion"] = clientVersion,
            ["capabilities"] = new JsonObject(),
            ["clientInfo"] = new JsonObject { ["name"] = "test", ["version"] = "1.0" },
        }));
        Assert.Equal(clientVersion, Json.GetString(ResultOf(res), "protocolVersion"));
        Assert.Equal(clientVersion, JsonRpcDispatcher.NegotiateProtocolVersion(clientVersion));
    }

    [Fact]
    public void Protocol_negotiation_is_stateless_across_clients()
    {
        // HTTP 모드에서 서로 다른 버전의 클라이언트가 동시에 붙어도 서로 영향을 주면 안 된다.
        var claude = _dispatcher.Dispatch(Req("initialize", 1,
            new JsonObject { ["protocolVersion"] = "2024-11-05" }));
        var codex = _dispatcher.Dispatch(Req("initialize", 2,
            new JsonObject { ["protocolVersion"] = "2025-06-18" }));
        var claudeAgain = _dispatcher.Dispatch(Req("initialize", 3,
            new JsonObject { ["protocolVersion"] = "2024-11-05" }));

        Assert.Equal("2024-11-05", Json.GetString(ResultOf(claude), "protocolVersion"));
        Assert.Equal("2025-06-18", Json.GetString(ResultOf(codex), "protocolVersion"));
        Assert.Equal("2024-11-05", Json.GetString(ResultOf(claudeAgain), "protocolVersion"));

        // HTTP 헤더 경로가 쓰는 요청객체 오버로드도 같은 결과여야 한다.
        Assert.Equal("2024-11-05", JsonRpcDispatcher.NegotiateProtocolVersion(
            Req("initialize", 1, new JsonObject { ["protocolVersion"] = "2024-11-05" })));
        Assert.Equal(JsonRpcDispatcher.ProtocolVersion,
            JsonRpcDispatcher.NegotiateProtocolVersion(Req("tools/list", 2)));
    }

    [Fact]
    public void Initialize_falls_back_to_server_version_for_unknown_protocol()
    {
        var res = _dispatcher.Dispatch(Req("initialize", 1, new JsonObject
        {
            ["protocolVersion"] = "1999-01-01",
        }));
        Assert.Equal(JsonRpcDispatcher.ProtocolVersion, Json.GetString(ResultOf(res), "protocolVersion"));
    }

    [Fact]
    public void Initialize_includes_agent_instructions()
    {
        var result = ResultOf(_dispatcher.Dispatch(Req("initialize", 1)))!;
        var instructions = Json.GetString(result, "instructions");
        Assert.False(string.IsNullOrWhiteSpace(instructions));
        Assert.Contains("dryRun", instructions);      // 안전 흐름을 에이전트가 알 수 있어야 한다
        Assert.Contains("confirmToken", instructions);
        Assert.Contains("target.sheet", instructions);
        Assert.Contains("core_get_status", instructions);
        Assert.Contains("allowOpenFile", instructions);
        Assert.Contains("반복하지", instructions);
        Assert.Contains("openpyxl", instructions);
        Assert.Contains("pywin32/직접 Excel COM", instructions);
        Assert.Contains("Start-Process/쉘/UI 자동화", instructions);
    }

    [Fact]
    public void Unused_capabilities_answer_with_empty_lists_instead_of_failing()
    {
        // 일부 클라이언트는 기동 시 무조건 조회한다. -32601로 실패시키면 연결이 끊긴다.
        Assert.Empty(Json.GetArr(ResultOf(_dispatcher.Dispatch(Req("resources/list", 1))), "resources")!);
        Assert.Empty(Json.GetArr(ResultOf(_dispatcher.Dispatch(Req("prompts/list", 2))), "prompts")!);
        Assert.Empty(Json.GetArr(ResultOf(_dispatcher.Dispatch(Req("resources/templates/list", 3))), "resourceTemplates")!);
        Assert.NotNull(ResultOf(_dispatcher.Dispatch(Req("logging/setLevel", 4))));
        Assert.NotNull(ResultOf(_dispatcher.Dispatch(Req("ping", 5))));
    }

    [Fact]
    public void Tools_call_returns_structured_content_matching_text()
    {
        var result = ResultOf(_dispatcher.Dispatch(Req("tools/call", 1, new JsonObject
        {
            ["name"] = "core_ping",
            ["arguments"] = new JsonObject(),
        })))!;

        var structured = Json.GetObj(result, "structuredContent");
        Assert.NotNull(structured);
        Assert.True(Json.GetBool(structured, "ok"));

        var text = Json.GetString(Json.GetArr(result, "content")![0] as JsonObject, "text")!;
        Assert.Equal(Json.ToCompact(structured), text); // 두 표현이 어긋나면 안 된다
    }

    [Fact]
    public void Batch_request_returns_array_without_notification_entries()
    {
        var line = """[{"jsonrpc":"2.0","id":1,"method":"initialize"},{"jsonrpc":"2.0","method":"notifications/initialized"},{"jsonrpc":"2.0","id":2,"method":"tools/list"}]""";

        var res = _dispatcher.DispatchLine(line) as JsonArray;
        Assert.NotNull(res);
        Assert.Equal(2, res.Count); // notification은 응답에 포함되지 않는다
        Assert.Equal(1, Json.GetInt(res[0] as JsonObject, "id"));
        Assert.Equal(2, Json.GetInt(res[1] as JsonObject, "id"));
    }

    [Fact]
    public void Missing_method_with_id_is_invalid_request()
    {
        var res = _dispatcher.Dispatch(new JsonObject { ["jsonrpc"] = "2.0", ["id"] = 9 })!;
        Assert.Equal(-32600, Json.GetInt(Json.GetObj(res, "error"), "code"));

        // id가 없으면 notification이므로 조용히 무시한다
        Assert.Null(_dispatcher.Dispatch(new JsonObject { ["jsonrpc"] = "2.0" }));
    }

    [Fact]
    public async Task Stdio_loop_writes_only_protocol_lines()
    {
        var input = new StringReader(string.Join('\n',
            "\uFEFF" + """{"jsonrpc":"2.0","id":1,"method":"initialize"}""",
            """{"jsonrpc":"2.0","method":"notifications/initialized"}""",
            """{"jsonrpc":"2.0","id":2,"method":"tools/list"}""",
            """{"jsonrpc":"2.0","id":3,"method":"tools/call","params":{"name":"core_ping","arguments":{}}}""",
            "not json at all",
            ""));
        var output = new StringWriter();
        var server = new McpServer(_dispatcher, input, output);

        await server.RunAsync();

        var lines = output.ToString().Split('\n', StringSplitOptions.RemoveEmptyEntries);
        Assert.Equal(4, lines.Length); // initialize + tools/list + tools/call + parse error; notification은 응답 없음
        foreach (var line in lines)
        {
            // stdout의 모든 줄은 JSON-RPC여야 한다 — 로그/배너가 섞이면 여기서 실패
            var msg = JsonNode.Parse(line) as JsonObject;
            Assert.NotNull(msg);
            Assert.Equal("2.0", Json.GetString(msg, "jsonrpc"));
        }
        var parseErr = JsonNode.Parse(lines[3]) as JsonObject;
        Assert.Equal(-32700, Json.GetInt(Json.GetObj(parseErr, "error"), "code"));
        var initialize = JsonNode.Parse(lines[0]) as JsonObject;
        Assert.Equal(1, Json.GetInt(initialize, "id"));
    }
}
