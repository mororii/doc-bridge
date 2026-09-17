using System.Text.Json.Serialization.Metadata;

namespace DocBridge.Development.ExcelProductionWorkflowProbe;

internal static class JsonUtil
{
    public static readonly JsonSerializerOptions WritePretty = Create(true);
    public static readonly JsonSerializerOptions WriteCompact = Create(false);

    private static JsonSerializerOptions Create(bool pretty) => new()
    {
        WriteIndented = pretty,
        Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        TypeInfoResolver = new DefaultJsonTypeInfoResolver(),
    };

    public static JsonNode? Clone(JsonNode? node) => node?.DeepClone();

    public static JsonArray Arr(params JsonNode?[] items)
    {
        var a = new JsonArray();
        foreach (var item in items) a.Add(item);
        return a;
    }

    public static void Write(string path, JsonNode node, bool pretty = true)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, node.ToJsonString(pretty ? WritePretty : WriteCompact), new UTF8Encoding(false));
    }

    public static JsonNode Load(string path) =>
        JsonNode.Parse(File.ReadAllText(path)) ?? throw new InvalidOperationException($"empty JSON: {path}");

    public static JsonNode? Get(JsonNode? node, string name) =>
        node is JsonObject o && o.TryGetPropertyValue(name, out var v) ? v : null;

    public static string? Str(JsonNode? node, string name) => Get(node, name)?.GetValue<string>();

    public static bool? Bool(JsonNode? node, string name) =>
        Get(node, name) is JsonValue v && v.TryGetValue<bool>(out var b) ? b : null;

    public static double? Num(JsonNode? node, string name)
    {
        var n = Get(node, name);
        if (n is JsonValue v)
        {
            if (v.TryGetValue<double>(out var d)) return d;
            if (v.TryGetValue<int>(out var i)) return i;
            if (v.TryGetValue<long>(out var l)) return l;
        }
        return null;
    }

    public static JsonObject Target(string sheet, string? workbook = null)
    {
        var t = new JsonObject { ["sheet"] = sheet };
        if (!string.IsNullOrWhiteSpace(workbook)) t["workbook"] = workbook;
        return t;
    }

    public static JsonObject Op(string name, string sheet, string? range = null)
    {
        var o = new JsonObject
        {
            ["op"] = name,
            ["target"] = Target(sheet),
        };
        if (range is not null) o["range"] = range;
        return o;
    }

    public static JsonNode? FromClr(object? value) => value switch
    {
        null => null,
        JsonNode n => n.DeepClone(),
        string s => JsonValue.Create(s),
        bool b => JsonValue.Create(b),
        int i => JsonValue.Create(i),
        long l => JsonValue.Create(l),
        double d => JsonValue.Create(d),
        decimal m => JsonValue.Create((double)m),
        JsonElement e => JsonNode.Parse(e.GetRawText()),
        _ => JsonValue.Create(Convert.ToString(value, CultureInfo.InvariantCulture)),
    };
}
