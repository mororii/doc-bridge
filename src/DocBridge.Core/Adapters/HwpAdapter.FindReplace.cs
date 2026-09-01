using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class HwpAdapter
{
    private sealed record HwpFindReplaceSimulation(
        string Before,
        string After,
        IReadOnlyList<int> SelectedOrdinals,
        int AvailableMatches,
        int ReplacedMatches);

    private static string FindReplaceValue(JsonObject op, string field)
    {
        var value = Json.GetString(op, field) ?? string.Empty;
        return Json.GetBool(Json.GetObj(op, "options"), "literalEntities")
            ? value
            : DecodeHwpSerializedText(value);
    }

    private static bool FindReplaceMatchCase(JsonObject op) =>
        Json.GetBool(op, "matchCase",
            Json.GetBool(Json.GetObj(op, "options"), "matchCase"));

    private static JsonObject? FindReplaceScope(JsonObject op) =>
        Json.GetObj(op, "scope") ?? Json.GetObj(op, "target");

    private static IReadOnlyList<int> FindMatchPositions(string text, string find, bool matchCase)
    {
        var positions = new List<int>();
        if (string.IsNullOrEmpty(find)) return positions;
        var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var offset = 0;
        while ((offset = text.IndexOf(find, offset, comparison)) >= 0)
        {
            positions.Add(offset);
            offset += find.Length;
        }
        return positions;
    }

    private static (int Start, int End) ParagraphCharacterRange(
        string text, int startParagraph, int endParagraph)
    {
        if (startParagraph < 0 || endParagraph < startParagraph)
            throw new ArgumentException("find_replace.scope의 startParagraph/endParagraph 범위가 올바르지 않습니다");
        var paragraph = 0;
        var start = startParagraph == 0 ? 0 : -1;
        var end = -1;
        for (var index = 0; index < text.Length; index++)
        {
            if (text[index] != '\n') continue;
            paragraph++;
            if (paragraph == startParagraph) start = index + 1;
            if (paragraph == endParagraph + 1)
            {
                end = index;
                break;
            }
        }
        if (start < 0) throw new ArgumentOutOfRangeException(
            nameof(startParagraph), $"startParagraph={startParagraph}가 문서 범위를 벗어났습니다");
        if (end < 0) end = text.Length;
        return (start, end);
    }

    private static HwpFindReplaceSimulation SimulateFindReplace(
        string input, string find, string replace, int? occurrence,
        bool matchCase, JsonObject? scope)
    {
        var before = DecodeHwpSerializedText(NormalizeNewlines(input));
        var positions = FindMatchPositions(before, find, matchCase);
        IEnumerable<int> candidateIndexes = Enumerable.Range(0, positions.Count);

        if (scope is not null && scope.ContainsKey("startParagraph"))
        {
            var startParagraph = Json.GetInt(scope, "startParagraph") ?? 0;
            var endParagraph = Json.GetInt(scope, "endParagraph") ?? startParagraph;
            var range = ParagraphCharacterRange(before, startParagraph, endParagraph);
            candidateIndexes = candidateIndexes.Where(index =>
                positions[index] >= range.Start && positions[index] + find.Length <= range.End);
        }

        var candidates = candidateIndexes.ToList();
        if (occurrence is not null)
        {
            if (occurrence.Value < 1 || occurrence.Value > candidates.Count)
                throw new ArgumentOutOfRangeException(nameof(occurrence),
                    $"find_replace occurrence={occurrence}의 유효 범위는 1..{candidates.Count}입니다");
            candidates = new List<int> { candidates[occurrence.Value - 1] };
        }

        var after = before;
        foreach (var candidate in candidates.OrderByDescending(value => positions[value]))
            after = after.Remove(positions[candidate], find.Length).Insert(positions[candidate], replace);
        return new HwpFindReplaceSimulation(
            before, after, candidates.Select(index => index + 1).ToArray(),
            candidateIndexes.Count(), candidates.Count);
    }

    private static HwpFindReplaceSimulation PreviewFindReplace(
        dynamic hwp, JsonObject op, string document, ApplyPreview preview)
    {
        var find = FindReplaceValue(op, "find");
        var replace = FindReplaceValue(op, "replace");
        if (find.Length == 0) throw new ArgumentException("find_replace.find는 빈 문자열일 수 없습니다");
        var occurrence = Json.GetInt(op, "occurrence");
        var matchCase = FindReplaceMatchCase(op);
        var scope = FindReplaceScope(op);

        if (scope is not null && scope.ContainsKey("tableIndex"))
        {
            var tableIndex = Json.GetInt(scope, "tableIndex") ?? 0;
            var row = Json.GetInt(scope, "row") ?? 0;
            var col = Json.GetInt(scope, "col") ?? 0;
            var cellIndex = Json.GetInt(scope, "cellIndex");
            if (!TryReadTableCell((object)hwp, tableIndex, row, col, cellIndex,
                    out var cell, out var error))
                throw new ArgumentException($"find_replace 표 셀 범위를 읽지 못했습니다: {error}");
            var simulation = SimulateFindReplace(
                cell!.Text, find, replace, occurrence, matchCase, scope: null);
            var locator = cellIndex is null ? $"cell:{row},{col}" : $"cellIndex:{cellIndex}";
            preview.Affected.Add(new AffectedRef(
                $"table:{tableIndex}/{locator}", $"{simulation.ReplacedMatches} occurrence(s)"));
            preview.Diff.Add(new DiffEntry
            {
                Ref = $"table:{tableIndex}/{locator}",
                Before = simulation.Before,
                After = simulation.After,
            });
            if (simulation.ReplacedMatches == 0)
                preview.Warnings.Add($"표 셀에서 '{find}' 일치 항목이 없습니다");
            return simulation;
        }

        var documentSimulation = SimulateFindReplace(
            document, find, replace, occurrence, matchCase, scope);
        preview.Affected.Add(new AffectedRef(
            scope is not null && scope.ContainsKey("startParagraph") ? "paragraph-range" : "matches",
            $"{documentSimulation.ReplacedMatches} occurrence(s)"));
        if (documentSimulation.ReplacedMatches == 0)
            preview.Warnings.Add($"'{find}' 일치 항목이 없습니다");
        else
        {
            preview.Diff.Add(new DiffEntry
            {
                Ref = scope is not null && scope.ContainsKey("startParagraph")
                    ? "paragraph-range"
                    : occurrence is null ? "document" : $"document/occurrence:{occurrence}",
                Before = documentSimulation.Before[..Math.Min(200, documentSimulation.Before.Length)],
                After = documentSimulation.After[..Math.Min(200, documentSimulation.After.Length)],
            });
        }
        return documentSimulation;
    }

    private static HwpWriteResult ApplyFindReplaceInTableCell(dynamic hwp, JsonObject op)
    {
        var scope = FindReplaceScope(op) ?? throw new ArgumentException("find_replace.scope가 필요합니다");
        var tableIndex = Json.GetInt(scope, "tableIndex") ?? 0;
        var row = Json.GetInt(scope, "row") ?? 0;
        var col = Json.GetInt(scope, "col") ?? 0;
        var cellIndex = Json.GetInt(scope, "cellIndex");
        if (!TryReadTableCell((object)hwp, tableIndex, row, col, cellIndex,
                out var cell, out var error))
            return new HwpWriteResult(false, $"table:{tableIndex}", error);

        var simulation = SimulateFindReplace(
            cell!.Text,
            FindReplaceValue(op, "find"),
            FindReplaceValue(op, "replace"),
            Json.GetInt(op, "occurrence"),
            FindReplaceMatchCase(op),
            scope: null);
        if (simulation.ReplacedMatches == 0)
            return new HwpWriteResult(true, $"table:{tableIndex}",
                "0 occurrence(s) replaced", simulation.Before, simulation.After);

        var cellWrite = new JsonObject
        {
            ["tableIndex"] = tableIndex,
            ["row"] = row,
            ["col"] = col,
            ["text"] = simulation.After,
            ["preserveStyle"] = Json.GetBool(op, "preserveStyle", true),
        };
        if (cellIndex is not null) cellWrite["cellIndex"] = cellIndex.Value;
        if (Json.GetObj(op, "style") is { } style) cellWrite["style"] = style.DeepClone();
        HwpWriteResult result = ExecTableCellSetText(hwp, cellWrite);
        return result with
        {
            Detail = result.Ok
                ? $"{simulation.ReplacedMatches} occurrence(s) replaced in scoped cell; {result.Detail}"
                : result.Detail,
        };
    }

    private static HwpWriteResult ApplyFindReplaceInDocument(dynamic hwp, JsonObject op)
    {
        var find = FindReplaceValue(op, "find");
        var replace = FindReplaceValue(op, "replace");
        if (find.Length == 0) throw new ArgumentException("find_replace.find는 빈 문자열일 수 없습니다");
        var scope = FindReplaceScope(op);
        var occurrence = Json.GetInt(op, "occurrence");
        var matchCase = FindReplaceMatchCase(op);
        string before = GetDocText(hwp);
        var simulation = SimulateFindReplace(before, find, replace, occurrence, matchCase, scope);
        if (simulation.ReplacedMatches == 0)
            return new HwpWriteResult(true, "matches", "0 occurrence(s) replaced", before, before);

        bool actionReturned = true;
        if (occurrence is null && (scope is null || !scope.ContainsKey("startParagraph")))
        {
            dynamic act = hwp.HAction;
            dynamic ps = hwp.HParameterSet.HFindReplace;
            act.GetDefault("AllReplace", ps.HSet);
            ps.FindString = find;
            ps.ReplaceString = replace;
            try { ps.IgnoreMessage = 1; } catch { }
            try { ps.MatchCase = matchCase ? 1 : 0; } catch { }
            actionReturned = (bool)act.Execute("AllReplace", ps.HSet);
        }
        else
        {
            foreach (var ordinal in simulation.SelectedOrdinals.OrderByDescending(value => value))
            {
                if (!SelectTextOccurrence(hwp, find, ordinal, matchCase))
                    return new HwpWriteResult(false, $"matches/occurrence:{ordinal}",
                        "지정한 항목을 한글 문서에서 다시 선택하지 못했습니다", before, GetDocText(hwp));
                var context = CaptureCurrentNativeStyle(hwp, $"find_replace:{find}#{ordinal}");
                if (!PrepareContextualWriteStyle(hwp, op, context) || !ExecInsertText(hwp, replace))
                    return new HwpWriteResult(false, $"matches/occurrence:{ordinal}",
                        "선택 항목 교체에 실패했습니다", before, GetDocText(hwp));
            }
        }

        var after = GetDocText(hwp);
        var verified = string.Equals(
            DecodeHwpSerializedText(NormalizeNewlines(after)),
            simulation.After,
            StringComparison.Ordinal);
        return new HwpWriteResult(
            verified,
            scope is not null && scope.ContainsKey("startParagraph") ? "paragraph-range" : "matches",
            verified
                ? $"{simulation.ReplacedMatches} occurrence(s) replaced; actionReturned={actionReturned}; exact readback verified"
                : $"find_replace exact readback mismatch; actionReturned={actionReturned}",
            before,
            after);
    }
}
