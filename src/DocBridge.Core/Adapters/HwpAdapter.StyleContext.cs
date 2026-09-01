using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

/// <summary>
/// 한글 쓰기 시 글자/문단 서식을 문맥으로 해석하고 보존한다.
/// "대상 위치" 하나만 믿지 않고 기존 값, 같은 라벨의 반복 양식, 위/아래의 같은 역할 셀 순으로 찾는다.
/// </summary>
public sealed partial class HwpAdapter
{
    private sealed record HwpNativeStyle(object? Character, object? Paragraph, string Source)
    {
        public bool Available => Character is not null || Paragraph is not null;
    }

    private sealed record HwpCellRead(string Text, HwpNativeStyle Style, bool HasFormula = false);

    private sealed record HwpTableRead(IReadOnlyList<HwpCellRead> Cells, bool Truncated);

    private static HwpNativeStyle CaptureCurrentNativeStyle(dynamic hwp, string source)
    {
        object? character = null;
        object? paragraph = null;
        try
        {
            dynamic value = hwp.CharShape;
            try { character = (object)value.Clone(); }
            catch { character = (object)value; }
        }
        catch { }
        try
        {
            dynamic value = hwp.ParaShape;
            try { paragraph = (object)value.Clone(); }
            catch { paragraph = (object)value; }
        }
        catch { }
        return new HwpNativeStyle(character, paragraph, source);
    }

    private static bool ApplyNativeStyle(dynamic hwp, HwpNativeStyle? style)
    {
        if (style is null || !style.Available) return false;
        var applied = false;
        if (style.Character is not null)
        {
            try { hwp.CharShape = (dynamic)style.Character; applied = true; }
            catch
            {
                try
                {
                    dynamic shape = style.Character;
                    applied |= (bool)hwp.HAction.Execute("CharShape", shape);
                }
                catch { }
            }
        }
        if (style.Paragraph is not null)
        {
            try { hwp.ParaShape = (dynamic)style.Paragraph; applied = true; }
            catch
            {
                try
                {
                    dynamic shape = style.Paragraph;
                    applied |= (bool)hwp.HAction.Execute("ParagraphShape", shape);
                }
                catch { }
            }
        }
        return applied;
    }

    private static bool PreserveStyle(JsonObject op) => Json.GetBool(op, "preserveStyle", true);

    private static bool ApplyExplicitWriteStyle(dynamic hwp, JsonObject op)
    {
        var style = Json.GetObj(op, "style");
        if (style is null) return true;
        ValidateCharacterStyle(style);
        ValidateParagraphStyle(style);
        return ApplyCharShape(hwp, style) && ApplyParagraphShape(hwp, style);
    }

    private static bool PrepareContextualWriteStyle(dynamic hwp, JsonObject op, HwpNativeStyle? context)
    {
        if (PreserveStyle(op) && context is not null) _ = ApplyNativeStyle(hwp, context);
        return ApplyExplicitWriteStyle(hwp, op);
    }

    private static bool NativeStylesEquivalent(HwpNativeStyle left, HwpNativeStyle right)
    {
        return string.Equals(
            NativeStyleFingerprint(left),
            NativeStyleFingerprint(right),
            StringComparison.Ordinal);
    }

    private static string NativeStyleFingerprint(HwpNativeStyle style)
    {
        var summary = NativeStyleSummary(style);
        return (summary["character"]?.ToJsonString() ?? "null") + "|" +
               (summary["paragraph"]?.ToJsonString() ?? "null");
    }

    /// <summary>
    /// 새 문단 삽입은 기준 문구 하나만 보지 않고 바로 위·아래 문단까지 비교한다.
    /// 위/아래가 같은 서식이면 그 합의를 우선하고, 서로 다르면 기준 문단을 사용한다.
    /// </summary>
    private static HwpNativeStyle ResolveParagraphContextStyle(
        object hwpObject, object startPosition, object endPosition, HwpNativeStyle anchor)
    {
        dynamic hwp = hwpObject;
        HwpNativeStyle? previous = null;
        HwpNativeStyle? next = null;
        try
        {
            hwp.SetPosBySet((dynamic)startPosition);
            if ((bool)hwp.HAction.Run("MovePrevParaBegin"))
                previous = CaptureCurrentNativeStyle(hwp, "previous-paragraph");
        }
        catch { }
        try
        {
            hwp.SetPosBySet((dynamic)endPosition);
            if ((bool)hwp.HAction.Run("MoveNextParaBegin"))
                next = CaptureCurrentNativeStyle(hwp, "next-paragraph");
        }
        catch { }

        if (previous is not null && next is not null && NativeStylesEquivalent(previous, next))
            return previous with { Source = "surrounding-paragraph-consensus" };
        if (previous is not null && NativeStylesEquivalent(previous, anchor))
            return anchor with { Source = "anchor+previous-paragraph" };
        if (next is not null && NativeStylesEquivalent(next, anchor))
            return anchor with { Source = "anchor+next-paragraph" };
        return anchor with
        {
            Source = previous is not null || next is not null
                ? "anchor-paragraph(neighbor-style-conflict)"
                : "anchor-paragraph",
        };
    }

    private static HwpNativeStyle CaptureCaretContextStyle(object hwpObject, string source)
    {
        dynamic hwp = hwpObject;
        var anchor = CaptureCurrentNativeStyle(hwp, source);
        try
        {
            dynamic position = hwp.CreateSet("ListParaPos");
            if (!(bool)hwp.GetPosBySet(position)) return anchor;
            var resolved = ResolveParagraphContextStyle((object)hwp, (object)position, (object)position, anchor);
            try { hwp.SetPosBySet(position); } catch { }
            return resolved;
        }
        catch { return anchor; }
    }

    private static HwpNativeStyle? CaptureExplicitStyleSource(object hwpObject, JsonObject op)
    {
        dynamic hwp = hwpObject;
        var source = Json.GetObj(op, "styleSource");
        if (source is null) return null;

        if (Json.GetString(source, "text") is { Length: > 0 } sourceText)
        {
            var occurrence = Json.GetInt(source, "occurrence") ?? 1;
            if (!SelectTextOccurrence(hwp, sourceText, occurrence, Json.GetBool(source, "matchCase", true)))
                throw new ArgumentException($"styleSource.text를 찾지 못했습니다: {sourceText}");
            return CaptureCurrentNativeStyle(hwp, $"explicit-text:{sourceText}#{occurrence}");
        }

        var tableIndex = Json.GetInt(source, "tableIndex");
        var cellIndex = Json.GetInt(source, "cellIndex");
        var error = "";
        if (tableIndex is not null && cellIndex is not null &&
            TryReadTableCell((object)hwp, tableIndex.Value, 0, 0, cellIndex, out var cell, out error))
            return cell!.Style with { Source = $"explicit-table:{tableIndex}/cellIndex:{cellIndex}" };
        if (tableIndex is not null && cellIndex is not null)
            throw new ArgumentException($"styleSource 셀을 읽지 못했습니다: {error}");

        throw new ArgumentException("styleSource에는 text 또는 tableIndex+cellIndex가 필요합니다");
    }

    private static bool TryReadTableCell(
        object hwpObject, int tableIndex, int row, int col, int? cellIndex,
        out HwpCellRead? result, out string error)
    {
        dynamic hwp = hwpObject;
        result = null;
        dynamic? table = FindControl(hwp, "tbl", tableIndex);
        if (table is null)
        {
            error = $"표 {tableIndex}을 찾을 수 없습니다";
            return false;
        }
        return TryReadTableCellOnControl(hwpObject, (object)table, tableIndex, row, col, cellIndex,
            out result, out error);
    }

    private static bool TryReadTableCellOnControl(
        object hwpObject, object tableObject, int tableIndex, int row, int col, int? cellIndex,
        out HwpCellRead? result, out string error,
        IReadOnlySet<(int List, int Para)>? formulaPositions = null)
    {
        dynamic hwp = hwpObject;
        result = null;
        if (!SelectTableCellBlockOnControl(hwpObject, tableObject, tableIndex, row, col, cellIndex, out error)) return false;
        try { hwp.HAction.Run("Cancel"); } catch { }
        try
        {
            if ((Convert.ToInt32(hwp.CurFieldState) & 0x0F) != 1)
            {
                error = "선택한 위치가 표 셀이 아닙니다";
                return false;
            }
        }
        catch { }

        var hasFormula = formulaPositions is null
            ? IsCurrentTableCellFormula(hwpObject)
            : CurrentListParaPosition(hwpObject) is { } key && formulaPositions.Contains(key);
        _ = hwp.HAction.Run("SelectAll");
        var text = NormalizeCellText(GetSelectionText(hwp));
        var locator = cellIndex is null ? $"{row},{col}" : $"cellIndex:{cellIndex}";
        var style = CaptureCurrentNativeStyle(hwp, $"table:{tableIndex}/{locator}");
        try { hwp.HAction.Run("Cancel"); } catch { }
        result = new HwpCellRead(text, style, hasFormula);
        return true;
    }

    /// <summary>표를 처음부터 한 번만 순회한다. 셀마다 표 시작점으로 되돌아가는 O(n²) COM 호출을 피한다.</summary>
    private static HwpTableRead ReadTableCellsSequential(object hwpObject, int tableIndex, int maxCells)
    {
        dynamic hwp = hwpObject;
        dynamic? table = FindControl(hwp, "tbl", tableIndex);
        if (table is null) return new HwpTableRead(Array.Empty<HwpCellRead>(), false);
        return ReadTableCellsSequentialOnControl(hwpObject, (object)table, tableIndex, maxCells, true, null);
    }

    private static HwpTableRead ReadTableCellsSequentialOnControl(
        object hwpObject, object tableObject, int tableIndex, int maxCells,
        bool includeStyles, IReadOnlySet<(int List, int Para)>? formulaPositions)
    {
        dynamic hwp = hwpObject;
        var cells = new List<HwpCellRead>();
        if (!SelectTableCellBlockOnControl(hwpObject, tableObject, tableIndex, 0, 0, 0, out _))
            return new HwpTableRead(cells, false);

        var truncated = true;
        for (var cellIndex = 0; cellIndex < maxCells; cellIndex++)
        {
            try { hwp.HAction.Run("Cancel"); } catch { }
            var hasFormula = formulaPositions is null
                ? IsCurrentTableCellFormula(hwpObject)
                : CurrentListParaPosition(hwpObject) is { } key && formulaPositions.Contains(key);
            _ = hwp.HAction.Run("SelectAll");
            var text = NormalizeCellText(GetSelectionText(hwp));
            var style = includeStyles
                ? CaptureCurrentNativeStyle(hwp, $"table:{tableIndex}/cellIndex:{cellIndex}")
                : new HwpNativeStyle(null, null, "not-requested");
            cells.Add(new HwpCellRead(text, style, hasFormula));
            try { hwp.HAction.Run("Cancel"); } catch { }

            if (cellIndex == maxCells - 1) break;
            if (!(bool)hwp.HAction.Run("TableCellBlock") || !(bool)hwp.HAction.Run("TableRightCell"))
            {
                truncated = false;
                try { hwp.HAction.Run("Cancel"); } catch { }
                break;
            }
        }
        return new HwpTableRead(cells, truncated);
    }

    private static string ControlReference(object controlObject, int fallbackIndex)
    {
        dynamic control = controlObject;
        try
        {
            dynamic anchor = control.GetAnchorPos(0);
            int Read(string name)
            {
                try { return Convert.ToInt32(anchor.Item(name)); } catch { return -1; }
            }
            return $"tbl:list={Read("List")};para={Read("Para")};pos={Read("Pos")}";
        }
        catch { return $"tbl:index={fallbackIndex}"; }
    }

    private static bool SameLabel(string left, string right) =>
        !string.IsNullOrWhiteSpace(left) &&
        string.Equals(NormalizeCellText(left), NormalizeCellText(right), StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// 빈 셀 주변의 같은 역할 후보를 위·아래 모두 본다. 가장 가까운 한쪽을 즉시 택하지 않고,
    /// 양쪽 합의 → 대상 셀 기본 서식과 일치 → 주변 다수 서식 순으로 결정한다.
    /// </summary>
    private static HwpNativeStyle? ChooseSurroundingCellStyle(
        IReadOnlyList<(HwpCellRead Cell, int Distance, string Direction)> candidates,
        HwpCellRead target, string scope)
    {
        if (candidates.Count == 0) return null;
        var rankedCandidates = candidates.Select(candidate => new
        {
            Candidate = candidate,
            Fingerprint = NativeStyleFingerprint(candidate.Cell.Style),
        }).ToArray();
        var targetFingerprint = NativeStyleFingerprint(target.Style);

        foreach (var distance in rankedCandidates.Select(item => item.Candidate.Distance).Distinct().OrderBy(value => value))
        {
            var sameDistance = rankedCandidates.Where(item => item.Candidate.Distance == distance).ToArray();
            var before = sameDistance.FirstOrDefault(item => item.Candidate.Direction == "before");
            var after = sameDistance.FirstOrDefault(item => item.Candidate.Direction == "after");
            if (before is not null && after is not null &&
                string.Equals(before.Fingerprint, after.Fingerprint, StringComparison.Ordinal))
            {
                return before.Candidate.Cell.Style with
                {
                    Source = $"surrounding-{scope}-consensus:distance:{distance}",
                };
            }
        }

        var targetMatch = rankedCandidates
            .Where(item => string.Equals(item.Fingerprint, targetFingerprint, StringComparison.Ordinal))
            .OrderBy(item => item.Candidate.Distance)
            .FirstOrDefault();
        if (targetMatch is not null)
        {
            return targetMatch.Candidate.Cell.Style with
            {
                Source = $"surrounding-{scope}+target-default:{targetMatch.Candidate.Direction}/distance:{targetMatch.Candidate.Distance}",
            };
        }

        var winner = rankedCandidates
            .Select(item => new
            {
                item.Candidate,
                Agreement = rankedCandidates.Count(peer =>
                    string.Equals(item.Fingerprint, peer.Fingerprint, StringComparison.Ordinal)),
            })
            .OrderByDescending(item => item.Agreement)
            .ThenBy(item => item.Candidate.Distance)
            .First();
        return winner.Candidate.Cell.Style with
        {
            Source = $"surrounding-{scope}-majority:{winner.Candidate.Direction}/distance:{winner.Candidate.Distance}/agreement:{winner.Agreement}",
        };
    }

    /// <summary>
    /// 빈 셀의 서식은 반복 양식의 같은 라벨 값 셀을 우선하고, 없으면 같은 열/역할의 위아래 값을 사용한다.
    /// 마지막에만 대상 빈 셀 자체의 기본 서식을 사용한다.
    /// </summary>
    private static HwpNativeStyle ResolveTableContextStyle(
        object hwpObject, JsonObject op, int tableIndex, int row, int col, int? cellIndex,
        HwpCellRead target)
    {
        dynamic hwp = hwpObject;
        var explicitSource = CaptureExplicitStyleSource((object)hwp, op);
        if (explicitSource is not null) return explicitSource;
        if (!string.IsNullOrWhiteSpace(target.Text)) return target.Style with { Source = "existing-target-text" };

        if (cellIndex is not null)
        {
            var currentTable = ReadTableCellsSequential((object)hwp, tableIndex,
                Math.Clamp(Math.Max(cellIndex.Value + 41, 100), 1, 500));
            var label = cellIndex.Value > 0 && cellIndex.Value - 1 < currentTable.Cells.Count
                ? currentTable.Cells[cellIndex.Value - 1].Text
                : "";

            // 동일한 라벨을 가진 이전 반복 표의 값 셀을 찾는다.
            if (!string.IsNullOrWhiteSpace(label))
            {
                for (var previousTable = tableIndex - 1; previousTable >= 0; previousTable--)
                {
                    var previousCells = ReadTableCellsSequential((object)hwp, previousTable, 500).Cells;

                    // 반복 양식은 먼저 같은 실제 셀 순서를 비교한다(가장 빠르고 역할도 가장 정확함).
                    if (cellIndex.Value < previousCells.Count && cellIndex.Value > 0 &&
                        SameLabel(label, previousCells[cellIndex.Value - 1].Text) &&
                        !string.IsNullOrWhiteSpace(previousCells[cellIndex.Value].Text))
                    {
                        return previousCells[cellIndex.Value].Style with
                        {
                            Source = $"repeated-label:{label}@table:{previousTable}/cellIndex:{cellIndex.Value}",
                        };
                    }

                    for (var i = 0; i + 1 < previousCells.Count; i++)
                    {
                        if (!SameLabel(label, previousCells[i].Text)) continue;
                        if (string.IsNullOrWhiteSpace(previousCells[i + 1].Text)) continue;
                        return previousCells[i + 1].Style with
                        {
                            Source = $"repeated-label:{label}@table:{previousTable}/cellIndex:{i + 1}",
                        };
                    }
                }
            }

            // 같은 표에서 좌우 짝 구조(라벨/값)를 유지하는 위·아래 후보를 함께 비교한다.
            var surrounding = new List<(HwpCellRead Cell, int Distance, string Direction)>();
            for (var distance = 2; distance <= 40; distance += 2)
            {
                foreach (var candidate in new[]
                {
                    (Index: cellIndex.Value - distance, Direction: "before"),
                    (Index: cellIndex.Value + distance, Direction: "after"),
                })
                {
                    var peerIndex = candidate.Index;
                    if (peerIndex < 0 || peerIndex >= currentTable.Cells.Count) continue;
                    var peer = currentTable.Cells[peerIndex];
                    if (string.IsNullOrWhiteSpace(peer.Text)) continue;
                    surrounding.Add((peer, distance, candidate.Direction));
                }
            }
            var resolved = ChooseSurroundingCellStyle(surrounding, target, $"same-role:table:{tableIndex}");
            if (resolved is not null) return resolved;
        }
        else
        {
            string label = "";
            if (col > 0 && TryReadTableCell((object)hwp, tableIndex, row, col - 1, null, out var left, out _))
                label = left!.Text;

            for (var previousTable = tableIndex - 1; previousTable >= 0; previousTable--)
            {
                if (!TryReadTableCell((object)hwp, previousTable, row, col, null, out var samePosition, out _)) continue;
                if (string.IsNullOrWhiteSpace(samePosition!.Text)) continue;
                if (!string.IsNullOrWhiteSpace(label) && col > 0 &&
                    TryReadTableCell((object)hwp, previousTable, row, col - 1, null, out var candidateLabel, out _) &&
                    !SameLabel(label, candidateLabel!.Text)) continue;
                return samePosition.Style with { Source = $"repeated-form:table:{previousTable}/cell:{row},{col}" };
            }

            var surrounding = new List<(HwpCellRead Cell, int Distance, string Direction)>();
            for (var distance = 1; distance <= 20; distance++)
            {
                foreach (var candidate in new[]
                {
                    (Row: row - distance, Direction: "before"),
                    (Row: row + distance, Direction: "after"),
                })
                {
                    var peerRow = candidate.Row;
                    if (peerRow < 0) continue;
                    if (!TryReadTableCell((object)hwp, tableIndex, peerRow, col, null, out var peer, out _)) continue;
                    if (string.IsNullOrWhiteSpace(peer!.Text)) continue;
                    surrounding.Add((peer, distance, candidate.Direction));
                }
            }
            var resolved = ChooseSurroundingCellStyle(surrounding, target, $"vertical:table:{tableIndex}/col:{col}");
            if (resolved is not null) return resolved;
        }

        return target.Style with { Source = "target-cell-default" };
    }

    private static string? DynamicString(object? value, string property)
    {
        if (value is null) return null;
        try
        {
            dynamic v = value;
            return Convert.ToString(v.Item(property));
        }
        catch { return null; }
    }

    private static int? DynamicInt(object? value, string property)
    {
        if (value is null) return null;
        try
        {
            dynamic v = value;
            return Convert.ToInt32(v.Item(property));
        }
        catch { return null; }
    }

    private static JsonObject NativeStyleSummary(HwpNativeStyle style)
    {
        var charShape = new JsonObject
        {
            ["fontName"] = DynamicString(style.Character, "FaceNameHangul"),
            ["fontSizePt"] = DynamicInt(style.Character, "Height") is { } height ? height / 100.0 : null,
            ["bold"] = DynamicInt(style.Character, "Bold") is { } bold ? bold != 0 : null,
            ["italic"] = DynamicInt(style.Character, "Italic") is { } italic ? italic != 0 : null,
            ["textColor"] = DynamicInt(style.Character, "TextColor"),
        };
        var paraShape = new JsonObject
        {
            ["alignType"] = DynamicInt(style.Paragraph, "AlignType"),
            ["leftMargin"] = DynamicInt(style.Paragraph, "LeftMargin"),
            ["rightMargin"] = DynamicInt(style.Paragraph, "RightMargin"),
            ["indentation"] = DynamicInt(style.Paragraph, "Indentation"),
            ["spaceBefore"] = DynamicInt(style.Paragraph, "PrevSpacing"),
            ["spaceAfter"] = DynamicInt(style.Paragraph, "NextSpacing"),
            ["lineSpacingType"] = DynamicInt(style.Paragraph, "LineSpacingType"),
            ["lineSpacing"] = DynamicInt(style.Paragraph, "LineSpacing"),
        };
        return new JsonObject
        {
            ["source"] = style.Source,
            ["character"] = charShape,
            ["paragraph"] = paraShape,
        };
    }

    private static JsonObject InspectTables(object hwpObject, int? requestedTable, int maxCells, bool includeStyles)
    {
        dynamic hwp = hwpObject;
        // 컨트롤을 먼저 한 번 스냅샷으로 고정한다. 표를 읽는 동안 선택/커서가 이동해도
        // 뒤 표의 인덱스가 앞 표로 드리프트하거나 중복되지 않는다.
        var tableControls = FindControls(hwpObject, "tbl");
        var tableCount = tableControls.Count;
        var tables = new JsonArray();
        var formulaPositions = FormulaCellPositions(hwpObject);
        var indexes = requestedTable is null
            ? Enumerable.Range(0, tableCount)
            : new[] { requestedTable.Value };

        object? originalPosition = null;
        try
        {
            dynamic position = hwp.CreateSet("ListParaPos");
            if ((bool)hwp.GetPosBySet(position)) originalPosition = (object)position;
        }
        catch { }

        foreach (var tableIndex in indexes)
        {
            if (tableIndex < 0 || tableIndex >= tableCount)
                throw new ArgumentOutOfRangeException(nameof(requestedTable), $"표 {tableIndex}이 없습니다 (표 개수={tableCount})");
            var cells = new JsonArray();
            var tableControl = tableControls[tableIndex];
            var dimensions = TryReadTableDimensionsOnControl(hwpObject, tableControl, tableIndex);
            var tableRead = ReadTableCellsSequentialOnControl(
                hwpObject, tableControl, tableIndex, maxCells, includeStyles, formulaPositions);
            for (var cellIndex = 0; cellIndex < tableRead.Cells.Count; cellIndex++)
            {
                var cell = tableRead.Cells[cellIndex];
                var item = new JsonObject
                {
                    ["cellIndex"] = cellIndex,
                    ["text"] = cell.Text,
                    ["hasFormula"] = cell.HasFormula,
                };
                if (includeStyles) item["style"] = NativeStyleSummary(cell.Style);
                cells.Add(item);
            }
            tables.Add(new JsonObject
            {
                ["tableIndex"] = tableIndex,
                ["controlRef"] = ControlReference(tableControl, tableIndex),
                ["rowCount"] = dimensions?.Rows,
                ["columnCount"] = dimensions?.Columns,
                ["cells"] = cells,
                ["cellCountRead"] = cells.Count,
                ["truncated"] = tableRead.Truncated,
            });
        }

        if (originalPosition is not null)
            try { hwp.SetPosBySet((dynamic)originalPosition); } catch { }

        return new JsonObject
        {
            ["tableCount"] = tableCount,
            ["tables"] = tables,
            ["stylesIncluded"] = includeStyles,
        };
    }
}
