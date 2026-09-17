namespace DocBridge.Core.Services;

/// <summary>Pure parsing and selection planning for Excel structured formula references.</summary>
public static class ExcelStructuredReferenceContract
{
    public sealed record Selector(string? Column = null, string? LastColumn = null, string? Part = null, bool CurrentRow = false);
    public sealed record StructuredReference(string? TableName, IReadOnlyList<Selector> Selectors);
    public sealed record TableBounds(int FirstColumn, int ColumnCount, int? HeaderRow, int? DataFirstRow, int? DataLastRow, int? TotalsRow);
    public sealed record SelectionPlan(IReadOnlyList<string> Addresses, string? Reason = null)
    {
        public bool Resolved => Reason is null;
    }

    public static bool TryParse(string token, out StructuredReference reference, out string? reason)
    {
        reference = new StructuredReference(null, Array.Empty<Selector>());
        reason = null;
        if (string.IsNullOrWhiteSpace(token)) { reason = "structured_reference_malformed"; return false; }
        var open = token.IndexOf('[');
        if (open < 0 || !token.EndsWith(']')) { reason = "structured_reference_malformed"; return false; }
        var table = open == 0 ? null : token[..open];
        if (table is not null && !IsTableName(table)) { reason = "structured_table_name_invalid"; return false; }
        if (!TryReadBracket(token, open, out var content, out var end) || end != token.Length) { reason = "structured_reference_malformed"; return false; }
        var items = new List<Selector>();
        if (content.StartsWith("[", StringComparison.Ordinal))
        {
            var cursor = 0;
            while (cursor < content.Length)
            {
                if (!TryReadBracket(content, cursor, out var item, out cursor)) { reason = "structured_selector_malformed"; return false; }
                if (!TryParseItem(item, out var selector, out reason)) return false;
                if (cursor < content.Length && content[cursor] == ':')
                {
                    if (!TryReadBracket(content, cursor + 1, out var last, out cursor) || !TryColumn(last, out var lastColumn)) { reason = "structured_column_range_invalid"; return false; }
                    selector = selector with { LastColumn = lastColumn };
                }
                items.Add(selector);
                if (cursor == content.Length) break;
                if (content[cursor++] != ',') { reason = "structured_selector_malformed"; return false; }
            }
        }
        else if (!TryParseItem(content, out var selector, out reason)) return false;
        else items.Add(selector);
        if (items.Count == 0) { reason = "structured_selector_empty"; return false; }
        reference = new StructuredReference(table, items);
        return true;
    }

    public static SelectionPlan Plan(StructuredReference reference, TableBounds bounds, IReadOnlyDictionary<string, int> columns, int sourceRow)
    {
        if (bounds.ColumnCount < 1) return new SelectionPlan(Array.Empty<string>(), "structured_table_columns_unavailable");
        var selectedColumns = new SortedSet<int>();
        var requestedParts = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var currentRow = false;
        foreach (var selector in reference.Selectors)
        {
            if (selector.CurrentRow) currentRow = true;
            if (selector.Part is not null) requestedParts.Add(selector.Part);
            if (selector.Column is not null && !TryColumns(selector, columns, selectedColumns, out var columnReason))
                return new SelectionPlan(Array.Empty<string>(), columnReason);
        }
        if (selectedColumns.Count == 0)
            for (var column = 1; column <= bounds.ColumnCount; column++) selectedColumns.Add(column);
        if (requestedParts.Count == 0) requestedParts.Add("#Data");
        if (requestedParts.Contains("#This Row")) { requestedParts.Remove("#This Row"); requestedParts.Add("#Data"); currentRow = true; }
        if (currentRow && (bounds.DataFirstRow is null || bounds.DataLastRow is null || sourceRow < bounds.DataFirstRow || sourceRow > bounds.DataLastRow))
            return new SelectionPlan(Array.Empty<string>(), "structured_current_row_outside_data_body");
        if (requestedParts.Remove("#All")) { requestedParts.Add("#Headers"); requestedParts.Add("#Data"); requestedParts.Add("#Totals"); }
        var rows = new List<(int? First, int? Last)>();
        if (currentRow) rows.Add((sourceRow, sourceRow));
        else
        {
            if (requestedParts.Contains("#Headers")) rows.Add((bounds.HeaderRow, bounds.HeaderRow));
            if (requestedParts.Contains("#Data")) rows.Add((bounds.DataFirstRow, bounds.DataLastRow));
            if (requestedParts.Contains("#Totals")) rows.Add((bounds.TotalsRow, bounds.TotalsRow));
        }
        var addresses = new List<string>();
        foreach (var (first, last) in rows)
        {
            if (first is null || last is null) continue;
            foreach (var run in Runs(selectedColumns))
                addresses.Add(new ExcelA1Box(first.Value, bounds.FirstColumn + run.Start - 1, last.Value - first.Value + 1, run.End - run.Start + 1).Address);
        }
        return new SelectionPlan(addresses);
    }

    private static bool TryColumns(Selector selector, IReadOnlyDictionary<string, int> columns, SortedSet<int> output, out string reason)
    {
        reason = "structured_column_not_found";
        if (!columns.TryGetValue(selector.Column!, out var first)) return false;
        var last = first;
        if (selector.LastColumn is not null && !columns.TryGetValue(selector.LastColumn, out last)) return false;
        if (last < first) { reason = "structured_column_range_reversed"; return false; }
        for (var value = first; value <= last; value++) output.Add(value);
        return true;
    }

    private static IEnumerable<(int Start, int End)> Runs(SortedSet<int> values)
    {
        var start = 0; var previous = 0;
        foreach (var value in values)
        {
            if (start == 0) { start = previous = value; continue; }
            if (value == previous + 1) { previous = value; continue; }
            yield return (start, previous); start = previous = value;
        }
        if (start != 0) yield return (start, previous);
    }

    private static bool TryParseItem(string text, out Selector selector, out string? reason)
    {
        selector = new Selector(); reason = null;
        if (string.IsNullOrEmpty(text)) { reason = "structured_selector_empty"; return false; }
        if (TrySpecial(text, out var special)) { selector = new Selector(Part: special, CurrentRow: special == "#This Row"); return true; }
        var current = IsUnescapedPrefix(text, '@');
        var value = Unescape(current ? text[1..] : text);
        if (string.IsNullOrWhiteSpace(value)) { reason = "structured_column_empty"; return false; }
        selector = new Selector(value, CurrentRow: current);
        return true;
    }

    private static bool TryColumn(string text, out string column)
    {
        column = "";
        if (TrySpecial(text, out _) || IsUnescapedPrefix(text, '@')) return false;
        column = Unescape(text);
        return !string.IsNullOrWhiteSpace(column);
    }

    private static bool TrySpecial(string text, out string special)
    {
        special = "";
        if (!IsUnescapedPrefix(text, '#')) return false;
        var candidate = Unescape(text);
        if (candidate.Equals("#Headers", StringComparison.OrdinalIgnoreCase) || candidate.Equals("#Data", StringComparison.OrdinalIgnoreCase) || candidate.Equals("#Totals", StringComparison.OrdinalIgnoreCase) || candidate.Equals("#All", StringComparison.OrdinalIgnoreCase) || candidate.Equals("#This Row", StringComparison.OrdinalIgnoreCase)) { special = candidate; return true; }
        return false;
    }

    private static bool TryReadBracket(string text, int open, out string content, out int end)
    {
        content = ""; end = open;
        if (open >= text.Length || text[open] != '[') return false;
        var depth = 1; var builder = new System.Text.StringBuilder();
        for (var cursor = open + 1; cursor < text.Length; cursor++)
        {
            var ch = text[cursor];
            if (ch == '\'' && cursor + 1 < text.Length) { builder.Append(ch).Append(text[++cursor]); continue; }
            if (ch == '[') { depth++; builder.Append(ch); continue; }
            if (ch == ']' && --depth == 0) { content = builder.ToString(); end = cursor + 1; return true; }
            builder.Append(ch);
        }
        return false;
    }

    private static bool IsUnescapedPrefix(string text, char prefix) => text.Length > 0 && text[0] == prefix;
    private static string Unescape(string text)
    {
        var builder = new System.Text.StringBuilder();
        for (var index = 0; index < text.Length; index++)
        {
            if (text[index] == '\'' && index + 1 < text.Length) index++;
            builder.Append(text[index]);
        }
        return builder.ToString();
    }
    private static bool IsTableName(string value) => value.All(ch => char.IsLetterOrDigit(ch) || ch is '_' or '.');
}
