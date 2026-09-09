using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class CadAdapter
{
    internal static readonly string[] GstarWriteOps =
    {
        "activate_document", "regen_document", "set_layer_visibility", "set_layer_color",
        "move_entities", "rotate_entities", "set_text_value", "zoom_window", "draw_entities",
        "copy_entities", "scale_entities", "mirror_entities", "offset_entities", "set_entity_properties",
        "set_block_attributes", "save_document", "delete_entities", "delete_entities_in_bounds", "delete_entities_from_index",
    };
    internal static readonly string[] GstarDrawTypes =
    {
        "line", "text", "mtext", "lwpolyline", "circle", "arc", "ellipse", "point", "dim_aligned", "dim_rotated",
    };

    // Check the entire batch before touching COM. Prevent an allowed prefix being applied
    // before a later op reaches an AutoCAD-specific interop or command path.
    private List<string> ValidateProductOperations(IReadOnlyList<JsonObject> ops)
    {
        var errors = new List<string>();
        if (_product != CadProduct.GstarCad) return errors;
        for (var i = 0; i < ops.Count; i++)
        {
            var op = ops[i];
            var name = Json.GetString(op, "op") ?? "";
            if (!GstarWriteOps.Contains(name, StringComparer.Ordinal))
                errors.Add($"ops[{i}] '{name}' is not supported by GstarCAD basic ActiveX support; no AutoCAD fallback is permitted.");
            if (name == "draw_entities")
                foreach (var entity in Json.GetArr(op, "entities")?.OfType<JsonObject>() ?? Enumerable.Empty<JsonObject>())
                {
                    if (!GstarDrawTypes.Contains((Json.GetString(entity, "type") ?? "").ToLowerInvariant()))
                        errors.Add($"ops[{i}] GstarCAD draw type must be one of: {string.Join(", ", GstarDrawTypes)}");
                    CheckColor(entity["color"], i, errors);
                }
            CheckColor(op["color"], i, errors);
            CheckColor((op["properties"] as JsonObject)?["color"], i, errors);
        }
        return errors;
    }

    private static void CheckColor(JsonNode? color, int index, List<string> errors)
    {
        if (color is JsonObject c && c["rgb"] is not null)
            errors.Add($"ops[{index}] GstarCAD RGB color is not yet supported; use color:{{aci:1}} instead.");
    }
}
