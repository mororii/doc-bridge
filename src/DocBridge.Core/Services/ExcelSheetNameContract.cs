using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

public static class ExcelSheetNameContract
{
    public const int MaxLength = 31;
    private static readonly char[] Forbidden = { ':', '\\', '/', '?', '*', '[', ']' };

    public static bool TryNormalize(string? name, int opIndex, string field, ICollection<string> errors, out string normalized)
    {
        normalized = "";
        if (string.IsNullOrWhiteSpace(name))
        {
            errors.Add($"ops[{opIndex}] '{field}' must be a non-empty worksheet name");
            return false;
        }

        var text = name.Trim();
        if (text.Length > MaxLength)
        {
            errors.Add($"ops[{opIndex}] '{field}' exceeds {MaxLength} characters");
            return false;
        }

        if (text.IndexOfAny(Forbidden) >= 0)
        {
            errors.Add($"ops[{opIndex}] '{field}' contains a forbidden character (: \\ / ? * [ ])");
            return false;
        }

        if (text.StartsWith('\'') || text.EndsWith('\''))
        {
            errors.Add($"ops[{opIndex}] '{field}' must not start or end with an apostrophe");
            return false;
        }

        normalized = text;
        return true;
    }
}
