namespace DocBridge.Development.ExcelProductionWorkflowProbe;

internal static class Redact
{
    public static JsonNode? Tokens(JsonNode? node)
    {
        if (node is JsonObject o)
        {
            var copy = new JsonObject();
            foreach (var (k, v) in o)
            {
                if (k.Contains("confirmToken", StringComparison.OrdinalIgnoreCase) ||
                    k.Contains("token", StringComparison.OrdinalIgnoreCase) &&
                    !k.Contains("timeout", StringComparison.OrdinalIgnoreCase))
                {
                    copy[k] = v is null ? null : "<redacted>";
                }
                else copy[k] = Tokens(v);
            }
            return copy;
        }

        if (node is JsonArray a)
        {
            var copy = new JsonArray();
            foreach (var item in a) copy.Add(Tokens(item));
            return copy;
        }

        return node?.DeepClone();
    }
}
