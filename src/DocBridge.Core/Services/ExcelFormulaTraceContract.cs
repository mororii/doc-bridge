using System.Text;
using System.Text.Json.Nodes;


namespace DocBridge.Core.Services;

/// <summary>
/// Conservative parser and bounded graph walker for public Excel formula traces.
/// It recognizes only A1 references and defined-name tokens; callers resolve every
/// emitted reference through Excel before it becomes an edge.
/// </summary>
public static class ExcelFormulaTraceContract
{
    public sealed record Reference(string Kind, string Token, string? Sheet = null, string? Address = null, string? Reason = null);
    public sealed record Cell(string Sheet, string Address, string? Formula, JsonObject Evidence);
    public sealed record Issue(string SourceSheet, string SourceAddress, Reference Reference, string Reason);
    public sealed record Edge(string SourceSheet, string SourceAddress, string TargetSheet, string TargetAddress, string Via);
    public sealed record Cycle(string SourceSheet, string SourceAddress, string TargetSheet, string TargetAddress);
    public sealed record Resolution(IReadOnlyList<Cell> Targets, IReadOnlyList<Issue> Unresolved, bool Truncated = false, string? TruncationReason = null);
    public sealed record Traversal(
        IReadOnlyList<Cell> Nodes,
        IReadOnlyList<Edge> Edges,
        IReadOnlyList<Issue> Unresolved,
        IReadOnlyList<Cycle> Cycles,
        IReadOnlyList<string> TruncationReasons,
        IReadOnlyList<string> Frontier)
    {
        public bool Complete => Unresolved.Count == 0 && TruncationReasons.Count == 0;
    }


    public static IReadOnlyList<Reference> ParseReferences(string? formula)
    {
        var references = new List<Reference>();
        if (string.IsNullOrWhiteSpace(formula) || !formula.TrimStart().StartsWith("=", StringComparison.Ordinal))
            return references;

        var text = formula;
        var index = 0;
        while (index < text.Length)
        {
            if (text[index] == '"')
            {
                SkipStringLiteral(text, ref index);
                continue;
            }

            if (StartsAt(text, index, "#REF!"))
            {
                references.Add(new Reference("unresolved", "#REF!", Reason: "missing_reference"));
                index += 5;
                continue;
            }

            if (TryReadWholeAxisReference(text, index, out var leadingAxisAddress, out var leadingAxisEnd))
            {
                references.Add(new Reference("axis", text[index..leadingAxisEnd], Address: leadingAxisAddress));
                index = leadingAxisEnd;
                continue;
            }

            if (TrySkipNumericLiteral(text, index, out var numericEnd))
            {
                index = numericEnd;
                continue;
            }

            if (text[index] == '[')
            {
                var after = text.IndexOf(']', index + 1);
                if (after >= 0 && TryReadQualifierAndA1(text, after + 1, out _, out _, out var externalEnd))
                {
                    references.Add(new Reference("unresolved", text[index..externalEnd], Reason: "external_workbook_reference"));
                    index = externalEnd;
                    continue;
                }
            }

            if (text[index] == '\'' && TryReadQuotedQualifierAndA1(text, index, out var quotedSheet, out var quotedAddress, out var quotedEnd))
            {
                var reason = quotedSheet.Contains('[', StringComparison.Ordinal)
                    ? "external_workbook_reference"
                    : quotedSheet.Contains(':', StringComparison.Ordinal)
                        ? "three_dimensional_reference"
                    : null;
                references.Add(reason is null
                    ? new Reference(quotedEnd < text.Length && text[quotedEnd] == '#' ? "spill" : "cell", text[index..(quotedEnd < text.Length && text[quotedEnd] == '#' ? quotedEnd + 1 : quotedEnd)], quotedSheet, quotedAddress)
                    : new Reference("unresolved", text[index..quotedEnd], Reason: reason));
                index = quotedEnd < text.Length && text[quotedEnd] == '#' ? quotedEnd + 1 : quotedEnd;
                continue;
            }

            if (TryReadQualifierAndA1(text, index, out var sheet, out var address, out var qualifiedEnd))
            {
                if (sheet.Contains(':', StringComparison.Ordinal))
                    references.Add(new Reference("unresolved", text[index..qualifiedEnd], Reason: "three_dimensional_reference"));
                else
                    references.Add(new Reference(qualifiedEnd < text.Length && text[qualifiedEnd] == '#' ? "spill" : "cell", text[index..(qualifiedEnd < text.Length && text[qualifiedEnd] == '#' ? qualifiedEnd + 1 : qualifiedEnd)], sheet, address));
                index = qualifiedEnd < text.Length && text[qualifiedEnd] == '#' ? qualifiedEnd + 1 : qualifiedEnd;
                continue;
            }

            if (TryReadWholeAxisReference(text, index, out var axisAddress, out var axisEnd))
            {
                references.Add(new Reference("axis", text[index..axisEnd], Address: axisAddress));
                index = axisEnd;
                continue;
            }

            if (TryReadIdentifier(text, index, out var identifier, out var identifierEnd))
            {
                if (identifierEnd < text.Length && text[identifierEnd] == '[')
                {
                    var end = ReadStructuredTokenEnd(text, identifierEnd);
                    if (end < 0)
                    {
                        references.Add(new Reference("structured", text[index..], Address: identifier));
                        break;
                    }
                    references.Add(new Reference("structured", text[index..end], Address: identifier));
                    index = end;
                    continue;
                }
                if (identifierEnd < text.Length && text[identifierEnd] == '(')
                {
                    if (identifier.Equals("INDIRECT", StringComparison.OrdinalIgnoreCase) ||
                        identifier.Equals("OFFSET", StringComparison.OrdinalIgnoreCase))
                    {
                        var end = SkipBalancedParentheses(text, identifierEnd);
                        references.Add(new Reference("unresolved", text[index..end], Reason: "dynamic_reference"));
                        index = end;
                        continue;
                    }
                    index = identifierEnd;
                    continue;
                }

                if (TryReadA1Range(text, index, out var directAddress, out var directEnd))
                {
                    references.Add(new Reference(directEnd < text.Length && text[directEnd] == '#' ? "spill" : "cell", text[index..(directEnd < text.Length && text[directEnd] == '#' ? directEnd + 1 : directEnd)], Address: directAddress));
                    index = directEnd < text.Length && text[directEnd] == '#' ? directEnd + 1 : directEnd;
                    continue;
                }

                if (!IsBooleanLiteral(identifier))
                    references.Add(new Reference("name", identifier));
                index = identifierEnd;
                continue;
            }

            if (TryReadA1Range(text, index, out var unqualifiedAddress, out var unqualifiedEnd))
            {
                references.Add(new Reference(unqualifiedEnd < text.Length && text[unqualifiedEnd] == '#' ? "spill" : "cell", text[index..(unqualifiedEnd < text.Length && text[unqualifiedEnd] == '#' ? unqualifiedEnd + 1 : unqualifiedEnd)], Address: unqualifiedAddress));
                index = unqualifiedEnd < text.Length && text[unqualifiedEnd] == '#' ? unqualifiedEnd + 1 : unqualifiedEnd;
                continue;
            }

            if (text[index] == '[' && ReadStructuredTokenEnd(text, index) is var structuredEnd && structuredEnd >= 0)
            {
                references.Add(new Reference("structured", text[index..structuredEnd]));
                index = structuredEnd;
                continue;
            }

            index++;
        }

        return references;
    }

    public static Traversal Trace(
        IEnumerable<Cell> roots,
        int maxDepth,
        int maxCells,
        Func<Cell, Reference, int, Resolution> resolve)
    {
        if (maxDepth < 0) throw new ArgumentOutOfRangeException(nameof(maxDepth));
        if (maxCells < 1) throw new ArgumentOutOfRangeException(nameof(maxCells));

        var nodes = new List<Cell>();
        var edges = new List<Edge>();
        var unresolved = new List<Issue>();
        var cycles = new List<Cycle>();
        var truncationReasons = new List<string>();
        var frontier = new List<string>();
        var known = new Dictionary<string, Cell>(StringComparer.OrdinalIgnoreCase);
        var pending = new Queue<(Cell Cell, int Depth, HashSet<string> Path)>();

        foreach (var root in roots)
        {
            if (!TryAdd(root, 0, new HashSet<string>(StringComparer.OrdinalIgnoreCase)))
            {
                AddTruncation("max_cells", Citation(root));
                break;
            }
        }

        while (pending.Count > 0)
        {
            var current = pending.Dequeue();
            foreach (var reference in ParseReferences(current.Cell.Formula))
            {
                if (reference.Kind == "unresolved")
                {
                    unresolved.Add(new Issue(current.Cell.Sheet, current.Cell.Address, reference,
                        reference.Reason ?? "unsupported_reference"));
                    continue;
                }

                if (current.Depth >= maxDepth)
                {
                    AddTruncation("max_depth", $"{Citation(current.Cell)} -> {reference.Token}");
                    continue;
                }

                var remaining = maxCells - nodes.Count;
                if (remaining < 1)
                {
                    AddTruncation("max_cells", $"{Citation(current.Cell)} -> {reference.Token}");
                    continue;
                }

                var resolution = resolve(current.Cell, reference, remaining);
                unresolved.AddRange(resolution.Unresolved);
                if (resolution.Truncated)
                    AddTruncation(resolution.TruncationReason ?? "max_cells", $"{Citation(current.Cell)} -> {reference.Token}");

                foreach (var target in resolution.Targets)
                {
                    var targetKey = Key(target);
                    edges.Add(new Edge(current.Cell.Sheet, current.Cell.Address, target.Sheet, target.Address, reference.Token));
                    if (known.ContainsKey(targetKey))
                    {
                        continue;
                    }
                    if (!TryAdd(target, current.Depth + 1, current.Path))
                    {
                        AddTruncation("max_cells", $"{Citation(current.Cell)} -> {Citation(target)}");
                        break;
                    }
                }
            }
        }

        // Traverse the verified graph independently of root discovery. Iterative DFS
        // reports closing edges once and avoids recursion or a reachability walk per edge.
        var adjacency = edges.GroupBy(edge => edge.SourceSheet + "!" + edge.SourceAddress, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(group => group.Key, group => group.ToArray(), StringComparer.OrdinalIgnoreCase);
        var colors = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        foreach (var start in known.Keys)
        {
            if (colors.ContainsKey(start)) continue;
            var stack = new Stack<(string Node, int Next)>();
            colors[start] = 1;
            stack.Push((start, 0));
            while (stack.Count > 0)
            {
                var frame = stack.Pop();
                if (!adjacency.TryGetValue(frame.Node, out var outgoing) || frame.Next >= outgoing.Length)
                {
                    colors[frame.Node] = 2;
                    continue;
                }
                stack.Push((frame.Node, frame.Next + 1));
                var edge = outgoing[frame.Next];
                var targetKey = edge.TargetSheet + "!" + edge.TargetAddress;
                if (colors.TryGetValue(targetKey, out var color))
                {
                    if (color == 1)
                    {
                        var cycle = new Cycle(edge.SourceSheet, edge.SourceAddress, edge.TargetSheet, edge.TargetAddress);
                        if (!cycles.Contains(cycle)) cycles.Add(cycle);
                    }
                }
                else
                {
                    colors[targetKey] = 1;
                    stack.Push((targetKey, 0));
                }
            }
        }
        return new Traversal(nodes, edges, unresolved, cycles, truncationReasons, frontier);

        bool TryAdd(Cell cell, int depth, HashSet<string> parentPath)
        {
            var key = Key(cell);
            if (known.ContainsKey(key)) return true;
            if (nodes.Count >= maxCells) return false;
            known.Add(key, cell);
            nodes.Add(cell);
            var path = new HashSet<string>(parentPath, StringComparer.OrdinalIgnoreCase) { key };
            pending.Enqueue((cell, depth, path));
            return true;
        }

        void AddTruncation(string reason, string item)
        {
            if (!truncationReasons.Contains(reason, StringComparer.OrdinalIgnoreCase))
                truncationReasons.Add(reason);
            if (!frontier.Contains(item, StringComparer.OrdinalIgnoreCase))
                frontier.Add(item);
        }
    }

    private static bool TryReadQuotedQualifierAndA1(string text, int index, out string sheet, out string address, out int end)
    {
        sheet = "";
        address = "";
        end = index;
        if (text[index] != '\'') return false;
        var builder = new StringBuilder();
        var cursor = index + 1;
        while (cursor < text.Length)
        {
            if (text[cursor] != '\'')
            {
                builder.Append(text[cursor++]);
                continue;
            }
            if (cursor + 1 < text.Length && text[cursor + 1] == '\'')
            {
                builder.Append('\'');
                cursor += 2;
                continue;
            }
            cursor++;
            if (cursor >= text.Length || text[cursor] != '!') return false;
            cursor++;
            if (!TryReadA1Range(text, cursor, out address, out end) && !TryReadWholeAxisReference(text, cursor, out address, out end)) return false;
            sheet = builder.ToString();
            return sheet.Length > 0;
        }
        return false;
    }

    private static bool TryReadQualifierAndA1(string text, int index, out string sheet, out string address, out int end)
    {
        sheet = "";
        address = "";
        end = index;
        if (index >= text.Length || !IsQualifierStart(text[index])) return false;
        var cursor = index + 1;
        while (cursor < text.Length && IsQualifierPart(text[cursor])) cursor++;
        if (cursor >= text.Length || text[cursor] != '!') return false;
        sheet = text[index..cursor];
        cursor++;
        if (!TryReadA1Range(text, cursor, out address, out end) && !TryReadWholeAxisReference(text, cursor, out address, out end)) return false;
        return true;
    }

    private static bool TryReadA1Range(string text, int index, out string address, out int end)
    {
        address = "";
        end = index;
        if (index > 0 && IsIdentifierPart(text[index - 1])) return false;
        if (!TryReadA1Cell(text, index, out var first, out var cursor)) return false;
        if (cursor < text.Length && text[cursor] == '(') return false; // LOG10( is a function, not cell LOG10.
        var last = first;
        if (cursor < text.Length && text[cursor] == ':')
        {
            if (!TryReadA1Cell(text, cursor + 1, out last, out cursor)) return false;
        }
        if (cursor < text.Length && IsIdentifierPart(text[cursor])) return false;
        if (!ExcelA1Box.TryParse(first + (last == first ? "" : ":" + last), out var box)) return false;
        address = box.Address;
        end = cursor;
        return true;
    }

    private static bool TryReadA1Cell(string text, int index, out string cell, out int end)
    {
        cell = "";
        end = index;
        var cursor = index;
        if (cursor < text.Length && text[cursor] == '$') cursor++;
        var lettersStart = cursor;
        while (cursor < text.Length && char.IsAsciiLetter(text[cursor]) && cursor - lettersStart < 4) cursor++;
        var letters = cursor - lettersStart;
        if (letters is < 1 or > 3) return false;
        if (cursor < text.Length && text[cursor] == '$') cursor++;
        var digitsStart = cursor;
        while (cursor < text.Length && char.IsAsciiDigit(text[cursor])) cursor++;
        if (cursor == digitsStart) return false;
        cell = text[index..cursor];
        if (!ExcelA1Box.TryParseCell(cell, out _, out _)) return false;
        end = cursor;
        return true;
    }

    private static bool TryReadWholeAxisReference(string text, int index, out string address, out int end)
    {
        address = "";
        end = index;
        if (index > 0 && IsIdentifierPart(text[index - 1])) return false;
        var cursor = index;
        if (cursor < text.Length && text[cursor] == '$') cursor++;
        var firstStart = cursor;
        while (cursor < text.Length && char.IsAsciiLetter(text[cursor]) && cursor - firstStart < 4) cursor++;
        var firstColumns = cursor - firstStart;
        if (firstColumns is > 0 and <= 3 && cursor < text.Length && text[cursor] == ':')
        {
            cursor++;
            if (cursor < text.Length && text[cursor] == '$') cursor++;
            var secondStart = cursor;
            while (cursor < text.Length && char.IsAsciiLetter(text[cursor]) && cursor - secondStart < 4) cursor++;
            if (cursor - secondStart is > 0 and <= 3 && (cursor == text.Length || !IsIdentifierPart(text[cursor])))
            {
                var first = text[firstStart..(firstStart + firstColumns)];
                var second = text[secondStart..cursor];
                var firstColumn = ExcelA1Box.ColumnIndex(first);
                var secondColumn = ExcelA1Box.ColumnIndex(second);
                if (firstColumn < 1 || secondColumn < firstColumn || secondColumn > ExcelA1Box.MaxColumn) return false;
                address = ExcelA1Box.CellName(firstColumn, 1) + ":" + ExcelA1Box.CellName(secondColumn, ExcelA1Box.MaxRow);
                end = cursor;
                return true;
            }
        }

        cursor = index;
        if (cursor < text.Length && text[cursor] == '$') cursor++;
        var rowStart = cursor;
        while (cursor < text.Length && char.IsAsciiDigit(text[cursor])) cursor++;
        if (cursor == rowStart || cursor >= text.Length || text[cursor] != ':') return false;
        cursor++;
        if (cursor < text.Length && text[cursor] == '$') cursor++;
        var secondRowStart = cursor;
        while (cursor < text.Length && char.IsAsciiDigit(text[cursor])) cursor++;
        if (cursor == secondRowStart || (cursor < text.Length && IsIdentifierPart(text[cursor]))) return false;
        var colon = text.IndexOf(':', rowStart);
        if (!int.TryParse(text[rowStart..colon], out var firstRow) || !int.TryParse(text[secondRowStart..cursor], out var secondRow) ||
            firstRow < 1 || secondRow < firstRow || secondRow > ExcelA1Box.MaxRow) return false;
        address = ExcelA1Box.CellName(1, firstRow) + ":" + ExcelA1Box.CellName(ExcelA1Box.MaxColumn, secondRow);
        end = cursor;
        return true;
    }

    private static bool TrySkipNumericLiteral(string text, int index, out int end)
    {
        end = index;
        var cursor = index;
        var hasDigits = false;
        if (cursor < text.Length && text[cursor] == '.') cursor++;
        while (cursor < text.Length && char.IsAsciiDigit(text[cursor]))
        {
            hasDigits = true;
            cursor++;
        }
        if (cursor < text.Length && text[cursor] == '.')
        {
            cursor++;
            while (cursor < text.Length && char.IsAsciiDigit(text[cursor]))
            {
                hasDigits = true;
                cursor++;
            }
        }
        if (!hasDigits) return false;
        if (cursor < text.Length && text[cursor] is 'E' or 'e')
        {
            var exponent = cursor + 1;
            if (exponent < text.Length && text[exponent] is '+' or '-') exponent++;
            var exponentStart = exponent;
            while (exponent < text.Length && char.IsAsciiDigit(text[exponent])) exponent++;
            if (exponent > exponentStart) cursor = exponent;
        }
        end = cursor;
        return true;
    }

    private static bool TryReadIdentifier(string text, int index, out string identifier, out int end)
    {
        identifier = "";
        end = index;
        if (index >= text.Length || !IsIdentifierStart(text[index])) return false;
        var cursor = index + 1;
        while (cursor < text.Length && IsIdentifierPart(text[cursor])) cursor++;
        identifier = text[index..cursor];
        end = cursor;
        return true;
    }

    private static int SkipBalancedParentheses(string text, int open)
    {
        var depth = 0;
        var index = open;
        while (index < text.Length)
        {
            if (text[index] == '"')
            {
                SkipStringLiteral(text, ref index);
                continue;
            }
            if (text[index] == '(') depth++;
            if (text[index] == ')' && --depth == 0) return index + 1;
            index++;
        }
        return text.Length;
    }

    private static int FindClosingBracket(string text, int open)
    {
        var depth = 0;
        for (var index = open; index < text.Length; index++)
        {
            if (text[index] == '[') depth++;
            else if (text[index] == ']' && --depth == 0) return index;
        }
        return -1;
    }

    // Apostrophe escapes the following bracket inside structured column names.
    private static int ReadStructuredTokenEnd(string text, int open)
    {
        var depth = 0;
        for (var index = open; index < text.Length; index++)
        {
            if (text[index] == '\'' && index + 1 < text.Length) { index++; continue; }
            if (text[index] == '[') depth++;
            else if (text[index] == ']' && --depth == 0) return index + 1;
        }
        return -1;
    }

    private static void SkipStringLiteral(string text, ref int index)
    {
        index++;
        while (index < text.Length)
        {
            if (text[index++] != '"') continue;
            if (index < text.Length && text[index] == '"') { index++; continue; }
            break;
        }
    }

    private static bool StartsAt(string text, int index, string token) =>
        index + token.Length <= text.Length && text.AsSpan(index, token.Length).Equals(token, StringComparison.OrdinalIgnoreCase);
    private static bool IsQualifierStart(char ch) => char.IsLetter(ch) || ch is '_' or '\\';
    private static bool IsQualifierPart(char ch) => IsQualifierStart(ch) || char.IsDigit(ch) || ch is '.' or ':';
    private static bool IsIdentifierStart(char ch) => char.IsLetter(ch) || ch is '_' or '\\';
    private static bool IsIdentifierPart(char ch) => IsIdentifierStart(ch) || char.IsDigit(ch) || ch == '.';
    private static bool IsBooleanLiteral(string token) => token.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || token.Equals("FALSE", StringComparison.OrdinalIgnoreCase);
    private static string Key(Cell cell) => cell.Sheet + "!" + cell.Address;
    private static string Citation(Cell cell) => cell.Sheet + "!" + cell.Address;
}
