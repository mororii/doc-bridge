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
        {
            var loaded = JsonNode.Parse(File.ReadAllText(path)) as JsonObject
                   ?? throw new InvalidOperationException($"invalid policy json: {path}");
            MergeDynamicExcelOps(loaded);
            return loaded;
        }

        // repo 루트 탐색 (ops/policies/default.policy.json)
        var dir = AppContext.BaseDirectory;
        for (var i = 0; i < 8 && dir is not null; i++)
        {
            var candidate = Path.Combine(dir, "ops", "policies", "default.policy.json");
            if (File.Exists(candidate))
            {
                var loaded = JsonNode.Parse(File.ReadAllText(candidate)) as JsonObject
                       ?? throw new InvalidOperationException($"invalid policy json: {candidate}");
                MergeDynamicExcelOps(loaded);
                return loaded;
            }
            dir = Directory.GetParent(dir)?.FullName;
        }

        var asm = typeof(PolicyEngine).Assembly;
        using var stream = asm.GetManifestResourceStream("DocBridge.ops.policies.default.policy.json")
            ?? throw new InvalidOperationException("embedded default.policy.json not found");
        var embedded = JsonNode.Parse(stream) as JsonObject
               ?? throw new InvalidOperationException("embedded default.policy.json invalid");
        MergeDynamicExcelOps(embedded);
        return embedded;
    }

    private static void MergeDynamicExcelOps(JsonObject policy)
    {
        if (Json.GetObj(policy, "apps")?["excel"] is not JsonObject excel)
            return;
        excel["writeOps"] ??= new JsonArray();
        excel["highRiskOps"] ??= new JsonArray();
        excel["forbiddenOps"] ??= new JsonArray();
        var write = (JsonArray)excel["writeOps"]!;
        var high = (JsonArray)excel["highRiskOps"]!;
        var forbidden = (JsonArray)excel["forbiddenOps"]!;

        void Ensure(JsonArray target, string name)
        {
            if (Contains(write, name) || Contains(high, name) || Contains(forbidden, name) || Contains(target, name))
                return;
            target.Add(name);
        }

        foreach (var name in ExcelDataOperationsContract.WriteOpNames)
        {
            if (name.StartsWith("delete_", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(name, ExcelDataOperationsContract.BreakExternalLink, StringComparison.OrdinalIgnoreCase))
                Ensure(high, name);
            else
                Ensure(write, name);
        }

        for (var i = forbidden.Count - 1; i >= 0; i--)
        {
            if (forbidden[i] is JsonValue value && value.TryGetValue<string>(out var name) &&
                name.Equals("delete_sheet", StringComparison.OrdinalIgnoreCase))
                forbidden.RemoveAt(i);
        }

        foreach (var name in new[]
                 {
                     "close_workbook", "fill_range", "auto_fill", "calculate",
                     "set_tab_color", "set_outline", "set_view", "import_csv", "export_csv", "set_page_breaks",
                 })
            Ensure(write, name);
        foreach (var name in new[] { "delete_sheet", "delete_rows", "delete_cols", "save_workbook", "export_pdf" })
            Ensure(high, name);
    }

    private static bool Contains(JsonArray array, string name) =>
        array.Any(node => node is JsonValue value && value.TryGetValue<string>(out var item) &&
                          item.Equals(name, StringComparison.OrdinalIgnoreCase));

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

    /// <summary>
    /// Direct execute is only for policy-allowed ordinary ops on the explicit autoExecute list.
    /// High-risk, forbidden, and unknown ops can never be auto-executed, even if listed by mistake.
    /// </summary>
    public bool IsAutoExecutable(string app, string op) =>
        ClassifyOp(app, op) == OpClass.Allowed && Set(AppPolicy(app), "autoExecuteOps").Contains(op);

    public IReadOnlyList<string> AutoExecuteOps(string app) =>
        Set(AppPolicy(app), "autoExecuteOps")
            .Where(op => ClassifyOp(app, op) == OpClass.Allowed)
            .OrderBy(op => op, StringComparer.OrdinalIgnoreCase)
            .ToArray();

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
