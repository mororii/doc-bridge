using System.Text.Json;
using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>System.Text.Json 공용 헬퍼</summary>
public static class Json
{
    public static readonly JsonSerializerOptions Pretty = new(JsonSerializerDefaults.General)
    {
        WriteIndented = true,
        TypeInfoResolver = new System.Text.Json.Serialization.Metadata.DefaultJsonTypeInfoResolver(),
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public static readonly JsonSerializerOptions Compact = new(JsonSerializerDefaults.General)
    {
        WriteIndented = false,
        TypeInfoResolver = new System.Text.Json.Serialization.Metadata.DefaultJsonTypeInfoResolver(),
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public static JsonObject? ParseObject(string? json)
    {
        if (string.IsNullOrWhiteSpace(json)) return null;
        return JsonNode.Parse(json) as JsonObject;
    }

    public static string ToPretty(JsonNode? node) =>
        node is null ? "null" : node.ToJsonString(Pretty);

    public static string ToCompact(JsonNode? node) =>
        node is null ? "null" : node.ToJsonString(Compact);

    public static JsonArray ToArray(IEnumerable<string> items)
    {
        var a = new JsonArray();
        foreach (var s in items) a.Add(s);
        return a;
    }

    public static JsonArray ToArray(IEnumerable<Models.AffectedRef> items)
    {
        var a = new JsonArray();
        foreach (var r in items) a.Add(new JsonObject { ["type"] = r.Type, ["ref"] = r.Ref });
        return a;
    }

    public static JsonArray ToArray(IEnumerable<Models.DiffEntry> items)
    {
        var a = new JsonArray();
        foreach (var d in items) a.Add(d.ToJson());
        return a;
    }

    /// <summary>토큰 바인딩용 canonical 문자열 (키 순서 정규화 없이 compact 직렬화 사용)</summary>
    public static string Canonical(JsonNode? node) => ToCompact(node);

    public static string? GetString(JsonObject? o, string key) =>
        o is not null && o.TryGetPropertyValue(key, out var v) && v is JsonValue jv && jv.TryGetValue<string>(out var s) ? s : null;

    public static bool GetBool(JsonObject? o, string key, bool fallback = false) =>
        o is not null && o.TryGetPropertyValue(key, out var v) && v is JsonValue jv && jv.TryGetValue<bool>(out var b) ? b : fallback;

    public static int? GetInt(JsonObject? o, string key) =>
        o is not null && o.TryGetPropertyValue(key, out var v) && v is JsonValue jv && jv.TryGetValue<int>(out var i) ? i : null;

    public static long? GetLong(JsonObject? o, string key) =>
        o is not null && o.TryGetPropertyValue(key, out var v) && v is JsonValue jv && jv.TryGetValue<long>(out var i) ? i : null;

    public static JsonObject? GetObj(JsonObject? o, string key) =>
        o is not null && o.TryGetPropertyValue(key, out var v) ? v as JsonObject : null;

    public static JsonArray? GetArr(JsonObject? o, string key) =>
        o is not null && o.TryGetPropertyValue(key, out var v) ? v as JsonArray : null;

    public static JsonObject ErrorResult(string message, string? app = null)
    {
        var o = new JsonObject { ["ok"] = false };
        if (app is not null) o["app"] = app;
        o["errors"] = ToArray(new[] { message });
        return o;
    }
}
