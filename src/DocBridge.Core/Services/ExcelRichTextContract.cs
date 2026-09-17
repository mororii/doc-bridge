using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>COM-free validation/planning for full-cell rich text.</summary>
public static class ExcelRichTextContract
{
    public const int MaxUtf16Length = 32_767;
    public const int MaxRuns = 256;
    public const int DefaultInspectMaxChars = 4_096;
    public static readonly string[] FontFields = ["name", "size", "bold", "italic", "underline", "color"];
    private static readonly HashSet<string> FontKeys = new(FontFields, StringComparer.OrdinalIgnoreCase);

    public static void Validate(JsonObject op, int index, ICollection<string> errors)
    {
        if (Json.GetObj(op, "baseFont") is not { } baseFont) return;
        ValidateFont(baseFont, index, "baseFont", complete: true, errors);
        var runs = Json.GetArr(op, "runs");
        if (runs is null || runs.Count is < 1 or > MaxRuns)
        { errors.Add($"ops[{index}] 'set_rich_text' runs must contain 1..{MaxRuns} items"); return; }
        var length = 0;
        for (var runIndex = 0; runIndex < runs.Count; runIndex++)
        {
            if (runs[runIndex] is not JsonObject run || run["text"] is not JsonValue textValue || !textValue.TryGetValue<string>(out var text))
            { errors.Add($"ops[{index}] 'set_rich_text' runs[{runIndex}].text must be a string"); continue; }
            if (text.Length == 0) errors.Add($"ops[{index}] 'set_rich_text' runs[{runIndex}].text cannot be empty");
            if (text.Any(char.IsSurrogate)) errors.Add($"ops[{index}] 'set_rich_text' does not support surrogate characters because native Characters indexing is unverified");
            length += text.Length;
            if (run["font"] is JsonObject font) ValidateFont(font, index, $"runs[{runIndex}].font", complete: false, errors);
            else if (run.ContainsKey("font")) errors.Add($"ops[{index}] 'set_rich_text' runs[{runIndex}].font must be an object, not null");
        }
        if (length > MaxUtf16Length) errors.Add($"ops[{index}] 'set_rich_text' combined text exceeds {MaxUtf16Length} characters");
    }

    public static JsonObject EffectiveFont(JsonObject baseFont, JsonObject? overrideFont)
    { var effective = (JsonObject)baseFont.DeepClone(); if (overrideFont is not null) foreach (var pair in overrideFont) effective[pair.Key] = pair.Value!.DeepClone(); return effective; }

    public static bool FontEquals(JsonObject expected, JsonObject? actual)
    {
        if (actual is null) return false;
        foreach (var field in FontFields)
            if (!actual.ContainsKey(field) || !string.Equals(Json.Canonical(expected[field]), Json.Canonical(actual[field]), StringComparison.Ordinal)) return false;
        return true;
    }

    private static void ValidateFont(JsonObject font, int index, string field, bool complete, ICollection<string> errors)
    {
        foreach (var pair in font)
            if (!FontKeys.Contains(pair.Key)) errors.Add($"ops[{index}] 'set_rich_text' {field}.{pair.Key} is unsupported");
        foreach (var key in FontFields)
            if (font.ContainsKey(key) && font[key] is null) errors.Add($"ops[{index}] 'set_rich_text' {field}.{key} cannot be null");
            else if (complete && !font.ContainsKey(key)) errors.Add($"ops[{index}] 'set_rich_text' {field}.{key} is required");
        if (font["name"] is not null && (Json.GetString(font, "name") is not { Length: > 0 and <= 255 })) errors.Add($"ops[{index}] 'set_rich_text' {field}.name must be 1..255 characters");
        if (font["size"] is not null && (!ExcelDataOperationsContract.TryGetFiniteNumber(font["size"], out var size) || size is < 1 or > 409)) errors.Add($"ops[{index}] 'set_rich_text' {field}.size must be a finite number from 1 to 409");
        foreach (var key in new[] { "bold", "italic", "underline" })
            if (font[key] is JsonValue value && !value.TryGetValue<bool>(out _)) errors.Add($"ops[{index}] 'set_rich_text' {field}.{key} must be boolean");
        if (font["color"] is not null && !ExcelStyleContract.TryParseColor(font["color"], out _)) errors.Add($"ops[{index}] 'set_rich_text' {field}.color must use the Excel color convention");
    }
}
