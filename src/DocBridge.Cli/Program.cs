using System.Text.Json;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;
using DocBridge.Mcp;

// doc-bridge-cli — Kimi fallback (명령서 §8.3)
//   doc-bridge-cli <tool> [--json '<json>' | --json-file args.json]
//   doc-bridge-cli <apply_tool> --ops ops.json [--dry-run | --confirm-token conf_...] [--high-risk-confirm]
// 출력: 어떤 경로로 끝나든 stdout에 결과 JSON 한 줄.
// exit code: 0 = ok, 1 = tool이 ok=false 반환, 2 = 인자/기동/미지원 tool 오류.

if (ExcelWorkerProcess.TryRun(args, out var excelWorkerExitCode))
    return excelWorkerExitCode;
if (ExcelOwnerWatchdog.TryRun(args, out var watchdogExitCode))
    return watchdogExitCode;

// 콘솔이 없는 상태(파이프 실행)에서 Console.*Encoding 세터는 IOException을 던지므로
// 실패해도 무시하고, 출력은 UTF-8 StreamWriter로 직접 쓴다.
var utf8 = new System.Text.UTF8Encoding(false);
try { Console.OutputEncoding = utf8; } catch { /* 콘솔 없음 — 무시 */ }
var stdout = new StreamWriter(Console.OpenStandardOutput(), utf8) { AutoFlush = true, NewLine = "\n" };

if (args.Length == 0 || args[0] is "-h" or "--help")
{
    Console.Error.WriteLine("usage: doc-bridge-cli <tool> [--json '<json>'|--json-file file] [--ops file [--dry-run|--confirm-token T] [--high-risk-confirm]]");
    Console.Error.WriteLine("tools: core_ping core_get_status core_get_capabilities core_disconnect core_create_snapshot core_list_snapshots core_restore_snapshot");
    Console.Error.WriteLine("       excel_get_active_context excel_read_range excel_inspect excel_apply_ops excel_disconnect");
    Console.Error.WriteLine("       hwp_plan_creation hwp_launch hwp_get_active_context hwp_doctor hwp_repair_typelib hwp_read_text hwp_apply_ops hwp_submit_ops hwp_get_job");
    Console.Error.WriteLine("       cad_launch cad_get_active_context cad_query_entities cad_apply_ops");
    if (args.Length == 0)
    {
        // stdout만 파싱하는 호출자를 위해 결과 JSON도 남긴다.
        stdout.WriteLine(Json.ToCompact(Json.ErrorResult("no tool specified; see --help")));
        stdout.Flush();
        return 2;
    }
    return 0;
}

var toolName = args[0];

// ---------- 인자 파싱 ----------
var toolArgs = new JsonObject();
string? opsFile = null;
string? confirmToken = null;
var dryRunFlag = false;
var highRiskConfirm = false;

// 인자 오류도 stdout에 결과 JSON으로 내보낸다 — 호출자(Kimi 등)가 stdout만 파싱하기 때문.
try
{
    for (var i = 1; i < args.Length; i++)
    {
        var a = args[i];
        string? Next() => i + 1 < args.Length ? args[++i] : null;
        switch (a)
        {
            case "--json":
            {
                var raw = Next() ?? throw new ArgumentException("--json requires a value");
                var fromJson = JsonNode.Parse(raw) as JsonObject
                    ?? throw new ArgumentException("--json must be a JSON object");
                // 누적된 인자에 병합한다 (같은 키는 나중 값 우선). clone 필수: 부모 재사용 금지
                foreach (var (k, v) in fromJson) toolArgs[k] = v?.DeepClone();
                break;
            }
            case "--json-file":
            {
                var path = Next() ?? throw new ArgumentException("--json-file requires a file");
                if (!File.Exists(path)) throw new FileNotFoundException($"json file not found: {path}");
                var fromFile = JsonNode.Parse(await File.ReadAllTextAsync(path)) as JsonObject
                    ?? throw new ArgumentException("--json-file must contain a JSON object");
                foreach (var (k, v) in fromFile) toolArgs[k] = v?.DeepClone();
                break;
            }
            case "--ops": opsFile = Next() ?? throw new ArgumentException("--ops requires a file"); break;
            case "--confirm-token": confirmToken = Next() ?? throw new ArgumentException("--confirm-token requires a value"); break;
            case "--dry-run": dryRunFlag = true; break;
            case "--high-risk-confirm": highRiskConfirm = true; break;
            default:
            {
                if (a.StartsWith("--"))
                {
                    var key = a[2..].Replace('-', '_');
                    var val = Next() ?? throw new ArgumentException($"{a} requires a value");
                    // JSON 값으로 파싱 시도, 실패 시 문자열
                    try { toolArgs[key] = JsonNode.Parse(val); }
                    catch (JsonException) { toolArgs[key] = val; }
                }
                else throw new ArgumentException($"unexpected argument: {a}");
                break;
            }
        }
    }

    if (opsFile is not null)
    {
        if (!File.Exists(opsFile)) throw new FileNotFoundException($"ops file not found: {opsFile}");
        var parsed = JsonNode.Parse(await File.ReadAllTextAsync(opsFile))
            ?? throw new ArgumentException($"ops file is not valid JSON: {opsFile}");
        if (parsed is JsonArray arr)
            toolArgs["ops"] = arr.DeepClone();
        else if (parsed is JsonObject obj)
            foreach (var (k, v) in obj) toolArgs[k] = v?.DeepClone();
        else
            throw new ArgumentException("ops file must be a JSON array or object");
        toolArgs["dryRun"] = confirmToken is null; // 토큰 있으면 apply, 없으면 dry-run
    }
    if (dryRunFlag) toolArgs["dryRun"] = true;
    if (confirmToken is not null)
    {
        toolArgs["dryRun"] = false;
        toolArgs["confirmToken"] = confirmToken;
    }
    if (highRiskConfirm) toolArgs["highRiskConfirm"] = true;
}
catch (Exception ex)
{
    stdout.WriteLine(Json.ToCompact(Json.ErrorResult($"argument error: {ex.Message}")));
    stdout.Flush();
    return 2;
}

// ---------- 실행 ----------
// 기동 실패(정책 파일 손상, %LOCALAPPDATA% 쓰기 권한 없음 등)도 stdout 결과 JSON으로 낸다.
DocBridgeHost host;
ToolRegistry tools;
try
{
    // 어댑터는 SessionRouter가 지연 생성한다 (excel/hwp/cad/fake 팩터리 내장).
    host = new DocBridgeHost();
    tools = new ToolRegistry(host);
}
catch (Exception ex)
{
    stdout.WriteLine(Json.ToCompact(Json.ErrorResult($"startup failed: {ex.Message}")));
    stdout.Flush();
    return 2;
}
using var hostScope = host;

var tool = tools.Find(toolName);
if (tool is null)
{
    Console.Error.WriteLine($"unknown tool: {toolName} (--help 참조)");
    stdout.WriteLine(Json.ToCompact(Json.ErrorResult($"unknown tool: {toolName}")));
    stdout.Flush();
    return 2;
}

JsonObject result;
try { result = tool.Handler(toolArgs); }
catch (Exception ex) { result = Json.ErrorResult($"{toolName} failed: {ex.Message}"); }

stdout.WriteLine(Json.ToCompact(result));
stdout.Flush();
return Json.GetBool(result, "ok") ? 0 : 1;
