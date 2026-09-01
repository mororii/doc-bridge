using System.Text.Json.Nodes;

namespace DocBridge.Core.Models;

/// <summary>실행 중인 문서 프로그램 연결 상태</summary>
public sealed record AdapterStatus(
    bool Available,      // 프로그램 설치/실행 감지됨
    bool Connected,      // 어댑터가 프로그램에 연결됨
    string Program,      // "excel" | "hwp" | "cad" | "fake"
    string? Version,
    string? Document,    // 현재 문서 경로/이름
    string? Detail);     // 추가 설명/오류

/// <summary>get_active_context 공통 응답 모델</summary>
public sealed class ContextResult
{
    public bool Ok { get; set; }
    public string App { get; set; } = "";
    public string? DocumentRef { get; set; }
    public JsonObject Summary { get; set; } = new();
    public JsonObject? Selection { get; set; }
    public JsonObject? Interaction { get; set; }
    public List<string> Warnings { get; } = new();
    public List<string> Errors { get; } = new();

    public JsonObject ToJson()
    {
        var o = new JsonObject
        {
            ["ok"] = Ok,
            ["app"] = App,
            ["documentRef"] = DocumentRef,
            ["summary"] = Summary.DeepClone(),
            ["selection"] = Selection?.DeepClone(),
            ["interaction"] = Interaction?.DeepClone(),
        };
        var w = new JsonArray(); foreach (var s in Warnings) w.Add(s);
        var e = new JsonArray(); foreach (var s in Errors) e.Add(s);
        o["warnings"] = w; o["errors"] = e;
        return o;
    }
}

public sealed record AffectedRef(string Type, string Ref);

public sealed class DiffEntry
{
    public string Ref { get; set; } = "";
    public JsonNode? Before { get; set; }
    public JsonNode? After { get; set; }

    public JsonObject ToJson() => new()
    {
        ["ref"] = Ref,
        ["before"] = Before?.DeepClone(),
        ["after"] = After?.DeepClone(),
    };
}

/// <summary>dry-run 단계에서 어댑터가 계산한 적용 예상 결과</summary>
public sealed class ApplyPreview
{
    public List<AffectedRef> Affected { get; } = new();
    public List<DiffEntry> Diff { get; } = new();
    public bool DiffTruncated { get; set; }
    public bool RequiresHighRiskApproval { get; set; }
    public List<string> Warnings { get; } = new();
    public List<string> Errors { get; } = new();
    public JsonObject? Interaction { get; set; }
}

/// <summary>실제 적용 + readback 검증 결과</summary>
public sealed class ApplyExecution
{
    public bool Ok { get; set; }
    public List<AffectedRef> Affected { get; } = new();
    public List<DiffEntry> Diff { get; } = new();
    public JsonArray OperationResults { get; } = new();
    public JsonObject? Readback { get; set; }
    public JsonObject? Interaction { get; set; }
    public List<string> Warnings { get; } = new();
    public List<string> Errors { get; } = new();
}

public sealed record SnapshotInfo(
    string SnapshotId,
    string CreatedAt,
    string App,
    string? DocumentRef,
    string Reason,
    string Dir);
