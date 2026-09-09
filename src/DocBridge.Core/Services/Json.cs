using System.Text;
using System.Text.Encodings.Web;
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

    /// <summary>
    /// Deterministic token-binding serialization. Object keys are sorted recursively;
    /// array order, missing vs null, strings, and Unicode are preserved. Numbers use
    /// exact lexical coefficient+exponent identity (1, 1.0, and 1e0 match; 1e-100
    /// stays distinct from 0). Values are never converted through decimal or IEEE-754
    /// double, and huge exponents are not expanded into zeros. This is not RFC 8785.
    /// </summary>
    public static string Canonical(JsonNode? node)
    {
        var sb = new StringBuilder();
        WriteCanonical(node, sb);
        return sb.ToString();
    }

    private static void WriteCanonical(JsonNode? node, StringBuilder sb)
    {
        switch (node)
        {
            case null:
                sb.Append("null");
                return;
            case JsonObject obj:
                sb.Append('{');
                var first = true;
                foreach (var key in obj.Select(kv => kv.Key).OrderBy(k => k, StringComparer.Ordinal))
                {
                    if (!first) sb.Append(',');
                    first = false;
                    WriteCanonicalString(key, sb);
                    sb.Append(':');
                    WriteCanonical(obj[key], sb);
                }
                sb.Append('}');
                return;
            case JsonArray arr:
                sb.Append('[');
                for (var i = 0; i < arr.Count; i++)
                {
                    if (i > 0) sb.Append(',');
                    WriteCanonical(arr[i], sb);
                }
                sb.Append(']');
                return;
            case JsonValue value:
                WriteCanonicalValue(value, sb);
                return;
            default:
                sb.Append(ToCompact(node));
                return;
        }
    }

    private static void WriteCanonicalValue(JsonValue value, StringBuilder sb)
    {
        RejectNonfiniteNumber(value);
        switch (value.GetValueKind())
        {
            case JsonValueKind.Null:
                sb.Append("null");
                return;
            case JsonValueKind.True:
                sb.Append("true");
                return;
            case JsonValueKind.False:
                sb.Append("false");
                return;
            case JsonValueKind.String:
                WriteCanonicalStringValue(value, sb);
                return;
            case JsonValueKind.Number:
                WriteCanonicalNumber(value, sb);
                return;
        }

        if (value.TryGetValue<string>(out var s))
        {
            WriteCanonicalString(s, sb);
            return;
        }
        if (value.TryGetValue<bool>(out var b))
        {
            sb.Append(b ? "true" : "false");
            return;
        }
        if (TryWriteJsonStringRepresentation(value, sb)) return;
        WriteCanonicalNumber(value, sb);
    }

    private static void WriteCanonicalStringValue(JsonValue value, StringBuilder sb)
    {
        if (value.TryGetValue<string>(out var text))
        {
            WriteCanonicalString(text, sb);
            return;
        }

        // DateTime/DateTimeOffset/Guid and other string-valued JsonValue types
        // serialize as JSON strings even when TryGetValue<string> fails.
        sb.Append(value.ToJsonString(Compact));
    }

    private static bool TryWriteJsonStringRepresentation(JsonValue value, StringBuilder sb)
    {
        var compact = value.ToJsonString(Compact);
        if (compact.Length >= 2 && compact[0] == '"')
        {
            sb.Append(compact);
            return true;
        }
        return false;
    }

    private static void RejectNonfiniteNumber(JsonValue value)
    {
        // Reject only a CLR float/double that is actually NaN or Infinity.
        // TryGetValue<float>/double on a parsed JsonElement converts 1e100 / 1e9999
        // to Infinity even though the JSON token is a finite lexical number.
        if (value.TryGetValue<JsonElement>(out _))
            return;

        object? stored;
        try { stored = value.GetValue<object>(); }
        catch { return; }

        switch (stored)
        {
            case float f32 when !float.IsFinite(f32):
                throw new ArgumentOutOfRangeException(nameof(value), f32, "JSON canonical numbers must be finite");
            case double f64 when !double.IsFinite(f64):
                throw new ArgumentOutOfRangeException(nameof(value), f64, "JSON canonical numbers must be finite");
        }
    }

    private static void WriteCanonicalNumber(JsonValue value, StringBuilder sb)
    {
        RejectNonfiniteNumber(value);
        if (value.TryGetValue<JsonElement>(out var element) && element.ValueKind == JsonValueKind.Number)
        {
            sb.Append(JsonCanonicalNumber.Normalize(element.GetRawText()));
            return;
        }

        var compact = value.ToJsonString(Compact);
        sb.Append(JsonCanonicalNumber.Normalize(compact));
    }

    private static void WriteCanonicalString(string value, StringBuilder sb)
    {
        using var stream = new MemoryStream();
        using (var writer = new Utf8JsonWriter(stream, new JsonWriterOptions
        {
            Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        }))
        {
            writer.WriteStringValue(value);
            writer.Flush();
        }
        sb.Append(Encoding.UTF8.GetString(stream.ToArray()));
    }

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
