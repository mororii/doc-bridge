using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

public enum OpClass { Allowed, HighRisk, Forbidden, Unknown }

/// <summary>
/// allowlist 기반 op 분류기. 정책 파일 우선, 없으면 EmbeddedResource 기본 정책.
/// 보안 원칙 6: tool은 allowlist 기반 operation만 받는다.
/// </summary>
public sealed class PolicyEngine
{
    private readonly JsonObject _policy;

    public PolicyEngine(string? policyPath = null)
    {
        _policy = Load(policyPath);
    }

    public int TokenTtlSeconds => Json.GetInt(_policy, "tokenTtlSeconds") ?? 300;
    public int MaxDiffEntries => Json.GetInt(_policy, "maxDiffEntries") ?? 100;
    public int MaxReadCells => Json.GetInt(_policy, "maxReadCells") ?? 10000;
    public int MaxReadChars => Json.GetInt(_policy, "maxReadChars") ?? 20000;

    private static JsonObject Load(string? path)
    {
        if (path is not null && File.Exists(path))
            return JsonNode.Parse(File.ReadAllText(path)) as JsonObject
                   ?? throw new InvalidOperationException($"invalid policy json: {path}");

        // repo 루트 탐색 (ops/policies/default.policy.json)
        var dir = AppContext.BaseDirectory;
        for (var i = 0; i < 8 && dir is not null; i++)
        {
            var candidate = Path.Combine(dir, "ops", "policies", "default.policy.json");
            if (File.Exists(candidate))
                return JsonNode.Parse(File.ReadAllText(candidate)) as JsonObject
                       ?? throw new InvalidOperationException($"invalid policy json: {candidate}");
            dir = Directory.GetParent(dir)?.FullName;
        }

        var asm = typeof(PolicyEngine).Assembly;
        using var stream = asm.GetManifestResourceStream("DocBridge.ops.policies.default.policy.json")
            ?? throw new InvalidOperationException("embedded default.policy.json not found");
        return JsonNode.Parse(stream) as JsonObject
               ?? throw new InvalidOperationException("embedded default.policy.json invalid");
    }

    private JsonObject? AppPolicy(string app) => Json.GetObj(_policy, "apps")?[app] as JsonObject;

    private static HashSet<string> Set(JsonObject? appPolicy, string key)
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (appPolicy is null) return set;
        if (Json.GetArr(appPolicy, key) is { } arr)
            foreach (var v in arr)
                if (v is JsonValue jv && jv.TryGetValue<string>(out var s))
                    set.Add(s);
        return set;
    }

    /// <summary>op 분류: Forbidden > HighRisk > Allowed > Unknown</summary>
    public OpClass ClassifyOp(string app, string op)
    {
        var p = AppPolicy(app);
        if (p is null) return OpClass.Unknown;
        if (Set(p, "forbiddenOps").Contains(op)) return OpClass.Forbidden;
        if (Set(p, "highRiskOps").Contains(op)) return OpClass.HighRisk;
        if (Set(p, "writeOps").Contains(op)) return OpClass.Allowed;
        return OpClass.Unknown;
    }

    public bool IsToolHighRisk(string tool)
    {
        if (Json.GetArr(_policy, "highRiskTools") is { } arr)
            foreach (var v in arr)
                if (v is JsonValue jv && jv.TryGetValue<string>(out var s) &&
                    string.Equals(s, tool, StringComparison.OrdinalIgnoreCase))
                    return true;
        return false;
    }
}
