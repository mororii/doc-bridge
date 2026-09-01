using System.Text;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

Console.InputEncoding = new UTF8Encoding(false);
Console.OutputEncoding = new UTF8Encoding(false);

using var adapter = new HwpAdapter();
string? line;
while ((line = Console.ReadLine()) is not null)
{
    var request = Json.ParseObject(line);
    var id = Json.GetString(request, "id") ?? "";
    var method = Json.GetString(request, "method") ?? "";
    var payload = Json.GetObj(request, "payload") ?? new JsonObject();
    JsonObject response;
    var restartRequired = false;
    try
    {
        if (method == "shutdown") break;
        JsonObject result = method switch
        {
            "getCapabilities" => adapter.GetCapabilities(),
            "getStatus" => HwpWorkerProtocol.StatusToJson(adapter.GetStatus()),
            "getActiveContext" => adapter.GetActiveContext().ToJson(),
            "read" => adapter.Read(payload),
            "preview" => HwpWorkerProtocol.PreviewToJson(adapter.Preview(ParseOps(payload))),
            "apply" => HwpWorkerProtocol.ExecutionToJson(adapter.Apply(ParseOps(payload), Json.GetString(payload, "snapshotId") ?? "")),
            "captureSnapshot" => CaptureSnapshot(adapter, payload),
            "validatePreviewReuse" => adapter.ValidatePreviewReuse(
                Json.GetString(payload, "snapshotDir") ?? throw new ArgumentException("snapshotDir is required"),
                Json.GetObj(payload, "metadata") ?? new JsonObject(),
                ParseOps(payload)),
            "restoreSnapshot" => adapter.RestoreSnapshot(
                Json.GetString(payload, "snapshotDir") ?? throw new ArgumentException("snapshotDir is required"),
                Json.GetObj(payload, "metadata") ?? new JsonObject()),
            "launch" => adapter.Launch(payload),
            "doctor" => adapter.Doctor(payload),
            "repairTypeLib" => adapter.RepairTypeLib(payload),
            _ => throw new ArgumentException($"unknown worker method '{method}'"),
        };
        restartRequired = HwpWorkerProtocol.ContainsComTimeout(result);
        response = new JsonObject
        {
            ["id"] = id,
            ["transportOk"] = true,
            ["result"] = result,
            ["restartRequired"] = restartRequired,
        };
    }
    catch (HwpAutomationException ex)
    {
        response = new JsonObject
        {
            ["id"] = id,
            ["transportOk"] = true,
            ["result"] = ex.ToResult(),
        };
    }
    catch (Exception ex)
    {
        response = new JsonObject
        {
            ["id"] = id,
            ["transportOk"] = false,
            ["error"] = ex.Message,
            ["exceptionType"] = ex.GetType().FullName,
        };
    }

    Console.WriteLine(Json.ToCompact(response));
    Console.Out.Flush();
    if (restartRequired) break;
}

static IReadOnlyList<JsonObject> ParseOps(JsonObject payload)
{
    var result = new List<JsonObject>();
    foreach (var node in Json.GetArr(payload, "ops") ?? new JsonArray())
        if (node is JsonObject op) result.Add(op.DeepClone() as JsonObject ?? new JsonObject());
    return result;
}

static JsonObject CaptureSnapshot(HwpAdapter adapter, JsonObject payload)
{
    var dir = Json.GetString(payload, "snapshotDir") ?? throw new ArgumentException("snapshotDir is required");
    var metadata = Json.GetObj(payload, "metadata")?.DeepClone() as JsonObject ?? new JsonObject();
    adapter.CaptureSnapshot(dir, metadata, ParseOps(payload));
    return new JsonObject { ["ok"] = true, ["metadata"] = metadata };
}
