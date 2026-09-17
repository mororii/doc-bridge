using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// COM-free E5-style name/formula dependency checks. Summary formulas that
/// mention ContractAmount/JanAmt must see workbook-scoped names whose
/// RefersTo matches the source cells. Sheet-local names and zero
/// divisors explain 0 / #DIV/0! / #NAME? without opening Excel.
/// </summary>
public static class ExcelDefinedNameDependencies
{
    public sealed record NameBinding(string Name, string Scope, string? Sheet, string? RefersTo, double? Value);

    public static IReadOnlyList<string> EnumerateNameTokens(string formula)
    {
        var tokens = new List<string>();
        var text = formula.Trim();
        var index = 0;
        while (index < text.Length)
        {
            var ch = text[index];
            if (ch == '"')
            {
                index++;
                while (index < text.Length)
                {
                    if (text[index++] == '"' &&
                        (index >= text.Length || text[index] != '"'))
                        break;
                    if (index < text.Length && text[index - 1] == '"' && text[index] == '"')
                        index++;
                }
                continue;
            }

            if (IsNameStart(ch))
            {
                var start = index;
                index++;
                while (index < text.Length && IsNamePart(text[index]))
                    index++;
                var token = text[start..index];
                if (index < text.Length && text[index] == '(')
                    continue;
                if (index < text.Length && text[index] == '!')
                    continue;
                if (ExcelA1Box.TryParseCell(token, out _, out _))
                    continue;
                if (ExcelDataOperationsContract.IsValidDefinedName(token))
                    tokens.Add(token);
                continue;
            }

            index++;
        }

        return tokens;
    }

    public static IReadOnlyList<string> ExplainMissingWorkbookBindings(
        string formula, IReadOnlyCollection<NameBinding> names)
    {
        var errors = new List<string>();
        var map = names.ToDictionary(item => item.Name, StringComparer.OrdinalIgnoreCase);
        foreach (var token in EnumerateNameTokens(formula))
        {
            if (!map.TryGetValue(token, out var binding))
            {
                errors.Add($"{formula}: #{token} is #NAME? (no defined name)");
                continue;
            }

            if (!string.Equals(binding.Scope, "workbook", StringComparison.OrdinalIgnoreCase))
                errors.Add($"{formula}: #{token} is sheet-local ({binding.Sheet}); summary formulas need workbook scope");
        }

        return errors;
    }

    public static IReadOnlyList<string> ExplainZeroDivisor(
        string formula, IReadOnlyCollection<NameBinding> names)
    {
        var errors = new List<string>();
        var slash = formula.IndexOf('/');
        if (slash <= 0) return errors;
        var right = formula[(slash + 1)..].Trim();
        if (right.StartsWith('=')) right = right[1..];
        foreach (var token in EnumerateNameTokens("=" + right))
        {
            var binding = names.FirstOrDefault(item =>
                string.Equals(item.Name, token, StringComparison.OrdinalIgnoreCase));
            if (binding is null) continue;
            if (binding.Value is 0)
                errors.Add($"{formula}: #{token} is 0 so this evaluates to #DIV/0!");
        }

        return errors;
    }

    public static IReadOnlyList<string> ExplainStaleCachedZero(
        string formula, double? cached, IReadOnlyCollection<NameBinding> names)
    {
        if (cached is not 0) return Array.Empty<string>();
        var used = new HashSet<string>(EnumerateNameTokens(formula), StringComparer.OrdinalIgnoreCase);
        var bindings = names.Where(item => used.Contains(item.Name)).ToList();
        if (bindings.Count == 0 || bindings.Any(item => item.Value is null))
            return Array.Empty<string>();
        if (bindings.Any(item => ExcelFormulaReference.HasRelativeCellRef(item.RefersTo)))
            return Array.Empty<string>();
        if (bindings.All(item => item.Value > 0))
            return new[]
            {
                $"{formula}: cached 0 while workbook names have nonzero sources; public calculate did not run",
            };
        return Array.Empty<string>();
    }

    /// <summary>
    /// Excel Names with relative A1 are selection-dependent
    /// (https://learn.microsoft.com/en-us/office/vba/api/excel.names).
    /// Saved 0/#DIV/0! cannot be blamed on stale cache alone.
    /// </summary>
    public static IReadOnlyList<string> ExplainRelativeRefersTo(IReadOnlyCollection<NameBinding> names)
    {
        var errors = new List<string>();
        foreach (var binding in names)
        {
            if (!ExcelFormulaReference.HasRelativeCellRef(binding.RefersTo)) continue;
            errors.Add(
                $"{binding.Name}: RefersTo '{binding.RefersTo}' is relative and selection-dependent; " +
                "intended fixed name is an absolute quoted cell such as ='월별기성'!$B$2");
        }
        return errors;
    }

    public static bool SourceRefersToExpected(NameBinding binding, string expectedRefersTo) =>
        string.Equals(binding.Scope, "workbook", StringComparison.OrdinalIgnoreCase) &&
        ExcelFormulaReference.SemanticEquals(binding.RefersTo, expectedRefersTo);

    public static NameBinding FromState(JsonObject state) =>
        new(
            Json.GetString(state, "localName") ?? Json.GetString(state, "name") ?? "",
            Json.GetString(state, "scope") ?? "workbook",
            Json.GetString(state, "sheet"),
            Json.GetString(state, "refersTo"),
            ExcelDataOperationsContract.TryGetFiniteNumber(state["value"], out var value) ? value : null);

    private static bool IsNameStart(char ch) => ch is '_' or '\\' || char.IsLetter(ch);

    private static bool IsNamePart(char ch) => IsNameStart(ch) || char.IsDigit(ch) || ch is '.' or '\\';
}
