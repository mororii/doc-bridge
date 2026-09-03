using System.Text.Json.Nodes;
using DocBridge.Core.Models;

namespace DocBridge.Core.Services;

/// <summary>
/// dry-run preview를 스냅샷 metadata에 보존해 별도 MCP/CLI 프로세스에서도 재사용한다.
/// 실제 재사용 여부는 IPreviewReuseAdapter의 전체 문서 fingerprint 검증 뒤에만 결정한다.
/// </summary>
public static class ApplyPreviewArtifact
{
    public const int Version = 2;

    public static JsonObject ToJson(ApplyPreview preview) => new()
    {
        ["version"] = Version,
        ["affected"] = Json.ToArray(preview.Affected),
        ["diff"] = Json.ToArray(preview.Diff),
        ["diffTruncated"] = preview.DiffTruncated,
        ["requiresHighRiskApproval"] = preview.RequiresHighRiskApproval,
        ["warnings"] = Json.ToArray(preview.Warnings),
        ["errors"] = Json.ToArray(preview.Errors),
        ["interaction"] = preview.Interaction?.DeepClone(),
    };

    public static void StoreInMetadata(JsonObject metadata, string opsHash, ApplyPreview preview)
    {
        metadata["opsHash"] = opsHash;
        metadata["previewArtifact"] = ToJson(preview);
    }

    public static ApplyPreview? FromMetadata(JsonObject metadata, string opsHash)
    {
        if (!string.Equals(Json.GetString(metadata, "opsHash"), opsHash, StringComparison.Ordinal))
            return null;
        return FromJson(Json.GetObj(metadata, "previewArtifact"));
    }

    public static ApplyPreview? FromJson(JsonObject? artifact)
    {
        var version = Json.GetInt(artifact, "version");
        if (artifact is null || version is not (1 or Version)) return null;
        var preview = new ApplyPreview
        {
            DiffTruncated = Json.GetBool(artifact, "diffTruncated"),
            RequiresHighRiskApproval = Json.GetBool(artifact, "requiresHighRiskApproval"),
            Interaction = Json.GetObj(artifact, "interaction")?.DeepClone() as JsonObject,
        };
        foreach (var node in Json.GetArr(artifact, "affected") ?? new JsonArray())
        {
            if (node is not JsonObject item) continue;
            preview.Affected.Add(new AffectedRef(
                Json.GetString(item, "type") ?? "unknown",
                Json.GetString(item, "ref") ?? ""));
        }
        foreach (var node in Json.GetArr(artifact, "diff") ?? new JsonArray())
        {
            if (node is not JsonObject item) continue;
            preview.Diff.Add(new DiffEntry
            {
                Ref = Json.GetString(item, "ref") ?? "",
                Before = item["before"]?.DeepClone(),
                After = item["after"]?.DeepClone(),
            });
        }
        foreach (var node in Json.GetArr(artifact, "warnings") ?? new JsonArray())
            if (node is not null) preview.Warnings.Add(node.GetValue<string>());
        foreach (var node in Json.GetArr(artifact, "errors") ?? new JsonArray())
            if (node is not null) preview.Errors.Add(node.GetValue<string>());
        return preview;
    }
}
