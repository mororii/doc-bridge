using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// System.Text.Json nodes may have only one parent. Inspect/read responses
/// must clone before attaching a node that already lives in another tree.
/// </summary>
public static class ExcelJsonOwnership
{
    public static JsonObject DetachObject(JsonObject node)
    {
        ArgumentNullException.ThrowIfNull(node);
        return node.Parent is null ? node : (JsonObject)node.DeepClone();
    }

    public static JsonNode? Clone(JsonNode? node) => node?.DeepClone();

    public static void AddDetached(JsonArray destination, JsonObject item)
    {
        ArgumentNullException.ThrowIfNull(destination);
        ArgumentNullException.ThrowIfNull(item);
        destination.Add(DetachObject(item));
    }
}
