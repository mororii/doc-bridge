using System.Text.Json.Nodes;
using DocBridge.Core.Models;

namespace DocBridge.Core.Services;

/// <summary>한글 외부 worker와 Core 프록시가 공유하는 JSON 직렬화 계약.</summary>
public static class HwpWorkerProtocol
{
    public static JsonObject StatusToJson(AdapterStatus value) => new()
    {
        ["available"] = value.Available,
        ["connected"] = value.Connected,
        ["program"] = value.Program,
        ["version"] = value.Version,
        ["document"] = value.Document,
        ["detail"] = value.Detail,
    };

    public static AdapterStatus StatusFromJson(JsonObject value) => new(
        Json.GetBool(value, "available"), Json.GetBool(value, "connected"),
        Json.GetString(value, "program") ?? "hwp", Json.GetString(value, "version"),
        Json.GetString(value, "document"), Json.GetString(value, "detail"));

    public static ContextResult ContextFromJson(JsonObject value)
    {
        var result = new ContextResult
        {
            Ok = Json.GetBool(value, "ok"),
            App = Json.GetString(value, "app") ?? "hwp",
            DocumentRef = Json.GetString(value, "documentRef"),
            Summary = Json.GetObj(value, "summary")?.DeepClone() as JsonObject ?? new JsonObject(),
            Selection = Json.GetObj(value, "selection")?.DeepClone() as JsonObject,
            Interaction = Json.GetObj(value, "interaction")?.DeepClone() as JsonObject,
        };
        AddStrings(result.Warnings, Json.GetArr(value, "warnings"));
        AddStrings(result.Errors, Json.GetArr(value, "errors"));
        return result;
    }

    public static JsonObject PreviewToJson(ApplyPreview value) => new()
    {
        ["affected"] = Json.ToArray(value.Affected),
        ["diff"] = Json.ToArray(value.Diff),
        ["diffTruncated"] = value.DiffTruncated,
        ["requiresHighRiskApproval"] = value.RequiresHighRiskApproval,
        ["warnings"] = Json.ToArray(value.Warnings),
        ["errors"] = Json.ToArray(value.Errors),
        ["interaction"] = value.Interaction?.DeepClone(),
    };

    public static ApplyPreview PreviewFromJson(JsonObject value)
    {
        var result = new ApplyPreview
        {
            DiffTruncated = Json.GetBool(value, "diffTruncated"),
            RequiresHighRiskApproval = Json.GetBool(value, "requiresHighRiskApproval"),
            Interaction = Json.GetObj(value, "interaction")?.DeepClone() as JsonObject,
        };
        AddAffected(result.Affected, Json.GetArr(value, "affected"));
        AddDiff(result.Diff, Json.GetArr(value, "diff"));
        AddStrings(result.Warnings, Json.GetArr(value, "warnings"));
        AddStrings(result.Errors, Json.GetArr(value, "errors"));
        return result;
    }

    public static JsonObject ExecutionToJson(ApplyExecution value) => new()
    {
        ["ok"] = value.Ok,
        ["affected"] = Json.ToArray(value.Affected),
        ["diff"] = Json.ToArray(value.Diff),
        ["operationResults"] = value.OperationResults.DeepClone(),
        ["readback"] = value.Readback?.DeepClone(),
        ["warnings"] = Json.ToArray(value.Warnings),
        ["errors"] = Json.ToArray(value.Errors),
        ["interaction"] = value.Interaction?.DeepClone(),
    };

    public static ApplyExecution ExecutionFromJson(JsonObject value)
    {
        var result = new ApplyExecution
        {
            Ok = Json.GetBool(value, "ok"),
            Readback = Json.GetObj(value, "readback")?.DeepClone() as JsonObject,
            Interaction = Json.GetObj(value, "interaction")?.DeepClone() as JsonObject,
        };
        AddAffected(result.Affected, Json.GetArr(value, "affected"));
        AddDiff(result.Diff, Json.GetArr(value, "diff"));
        foreach (var node in Json.GetArr(value, "operationResults") ?? new JsonArray())
            if (node is JsonObject item) result.OperationResults.Add(item.DeepClone());
        AddStrings(result.Warnings, Json.GetArr(value, "warnings"));
        AddStrings(result.Errors, Json.GetArr(value, "errors"));
        return result;
    }

    public static bool ContainsComTimeout(JsonNode? value)
    {
        if (value is null) return false;
        var text = value.ToJsonString();
        return text.Contains("STA work item did not complete", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("HWP_COM_TIMEOUT", StringComparison.OrdinalIgnoreCase) ||
               text.Contains("possible COM modal dialog", StringComparison.OrdinalIgnoreCase);
    }

    private static void AddStrings(ICollection<string> target, JsonArray? source)
    {
        foreach (var node in source ?? new JsonArray())
            if (node is JsonValue value && value.TryGetValue<string>(out var text)) target.Add(text);
    }

    private static void AddAffected(ICollection<AffectedRef> target, JsonArray? source)
    {
        foreach (var node in source ?? new JsonArray())
            if (node is JsonObject item)
                target.Add(new AffectedRef(Json.GetString(item, "type") ?? "unknown", Json.GetString(item, "ref") ?? ""));
    }

    private static void AddDiff(ICollection<DiffEntry> target, JsonArray? source)
    {
        foreach (var node in source ?? new JsonArray())
            if (node is JsonObject item)
                target.Add(new DiffEntry
                {
                    Ref = Json.GetString(item, "ref") ?? "",
                    Before = item["before"]?.DeepClone(),
                    After = item["after"]?.DeepClone(),
                });
    }
}
