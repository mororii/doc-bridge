using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

/// <summary>
/// 한컴 공식 HWP Automation 2025 문서의 Action/ParameterSet을 사용하는 실무 편집 기능.
/// UI 자동화나 스크립트 매크로를 사용하지 않고 HAction과 공개 COM 메서드만 호출한다.
/// </summary>
public sealed partial class HwpAdapter
{
    private sealed record HwpWriteResult(bool Ok, string Ref, string Detail, string? Before = null, string? After = null);
    private sealed record HwpFormatResult(string Ref, int AppliedCount);

    private static int ToHwpColorRef(string value)
    {
        var text = value.Trim();
        if (text.StartsWith('#')) text = text[1..];
        if (text.Length != 6 || !int.TryParse(text, System.Globalization.NumberStyles.HexNumber,
                System.Globalization.CultureInfo.InvariantCulture, out var rgb))
            throw new ArgumentException($"색상은 #RRGGBB 형식이어야 합니다: {value}");
        var red = (rgb >> 16) & 0xFF;
        var green = (rgb >> 8) & 0xFF;
        var blue = rgb & 0xFF;
        return red | (green << 8) | (blue << 16); // COLORREF = 0x00BBGGRR
    }

    private static int UnderlineType(string? value) => value?.ToLowerInvariant() switch
    {
        null or "none" => 0,
        "bottom" or "single" => 1,
        "center" => 2,
        "top" => 3,
        _ => throw new ArgumentException("underline은 none|bottom|center|top 중 하나여야 합니다"),
    };

    private static int StrikeOutType(string? value) => value?.ToLowerInvariant() switch
    {
        null or "none" => 0,
        "red-single" => 1,
        "red-double" => 2,
        "single" or "text-single" => 3,
        "double" or "text-double" => 4,
        _ => throw new ArgumentException("strikeout은 none|single|double|red-single|red-double 중 하나여야 합니다"),
    };

    private static void SetAllLanguageCharShapeFields(dynamic ps, string prefix, int value)
    {
        switch (prefix)
        {
            case "Spacing":
                if (value is < -50 or > 50) throw new ArgumentOutOfRangeException(nameof(value), "letterSpacing은 -50~50입니다");
                ps.SpacingHangul = value; ps.SpacingLatin = value; ps.SpacingHanja = value;
                ps.SpacingJapanese = value; ps.SpacingOther = value; ps.SpacingSymbol = value; ps.SpacingUser = value;
                break;
            case "Ratio":
                if (value is < 50 or > 200) throw new ArgumentOutOfRangeException(nameof(value), "widthRatio는 50~200입니다");
                ps.RatioHangul = value; ps.RatioLatin = value; ps.RatioHanja = value;
                ps.RatioJapanese = value; ps.RatioOther = value; ps.RatioSymbol = value; ps.RatioUser = value;
                break;
            case "Offset":
                if (value is < -100 or > 100) throw new ArgumentOutOfRangeException(nameof(value), "offset은 -100~100입니다");
                ps.OffsetHangul = value; ps.OffsetLatin = value; ps.OffsetHanja = value;
                ps.OffsetJapanese = value; ps.OffsetOther = value; ps.OffsetSymbol = value; ps.OffsetUser = value;
                break;
            default:
                throw new ArgumentException($"지원하지 않는 글자 모양 필드: {prefix}");
        }
    }

    private static int ParagraphAlignType(string value) => value.ToLowerInvariant() switch
    {
        "justify" => 0,
        "left" => 1,
        "right" => 2,
        "center" => 3,
        "distribute" => 4,
        "division" => 5,
        _ => throw new ArgumentException("align은 justify|left|right|center|distribute|division 중 하나여야 합니다"),
    };

    private static bool ApplyParagraphShape(dynamic hwp, JsonObject style)
    {
        dynamic action = hwp.HAction;
        dynamic shape = hwp.HParameterSet.HParaShape;
        action.GetDefault("ParagraphShape", shape.HSet);

        if (Json.GetString(style, "align") is { Length: > 0 } align) shape.AlignType = ParagraphAlignType(align);
        if (TryJsonNumber(style, "leftMarginMm", out var left)) shape.LeftMargin = hwp.MiliToHwpUnit(left);
        if (TryJsonNumber(style, "rightMarginMm", out var right)) shape.RightMargin = hwp.MiliToHwpUnit(right);
        if (TryJsonNumber(style, "firstLineIndentMm", out var indent)) shape.Indentation = hwp.MiliToHwpUnit(indent);
        if (TryJsonNumber(style, "spaceBeforePt", out var before)) shape.PrevSpacing = hwp.PointToHwpUnit(before);
        if (TryJsonNumber(style, "spaceAfterPt", out var after)) shape.NextSpacing = hwp.PointToHwpUnit(after);
        if (TryJsonNumber(style, "lineSpacingPercent", out var spacing))
        {
            if (spacing is < 50 or > 500) throw new ArgumentOutOfRangeException(nameof(spacing), "lineSpacingPercent는 50~500입니다");
            shape.LineSpacingType = 0;
            shape.LineSpacing = Convert.ToInt32(Math.Round(spacing));
        }
        if (style.TryGetPropertyValue("widowOrphan", out var widow) && widow is not null) shape.WidowOrphan = widow.GetValue<bool>();
        if (style.TryGetPropertyValue("keepWithNext", out var keepNext) && keepNext is not null) shape.KeepWithNext = keepNext.GetValue<bool>();
        if (style.TryGetPropertyValue("keepLinesTogether", out var keepLines) && keepLines is not null) shape.KeepLinesTogether = keepLines.GetValue<bool>();
        if (style.TryGetPropertyValue("pageBreakBefore", out var pageBreak) && pageBreak is not null) shape.PagebreakBefore = pageBreak.GetValue<bool>();
        return (bool)action.Execute("ParagraphShape", shape.HSet);
    }

    private static bool TryJsonNumber(JsonObject obj, string key, out double value)
    {
        value = 0;
        if (!obj.TryGetPropertyValue(key, out var node) || node is not JsonValue json) return false;
        return json.TryGetValue<double>(out value) || (json.TryGetValue<int>(out var integer) && (value = integer) == integer);
    }

    private static int ApplyParagraphShapeToTextMatches(dynamic hwp, string targetText, JsonObject style)
    {
        if (string.IsNullOrEmpty(targetText)) return 0;
        var count = 0;
        int? previousMessageMode = null;
        try
        {
            try { previousMessageMode = Convert.ToInt32(hwp.GetMessageBoxMode()); } catch { }
            try { hwp.SetMessageBoxMode(0x2FFF1); } catch { }
            hwp.HAction.Run("MoveDocBegin");
            dynamic action = hwp.HAction;
            dynamic find = hwp.HParameterSet.HFindReplace;
            action.GetDefault("FindDlg", find.HSet);
            find.FindString = targetText;
            try { find.Direction = hwp.FindDir("Forward"); } catch { }
            try { find.MatchCase = 1; find.SeveralWords = 0; find.UseWildCards = 0; find.WholeWordOnly = 0; } catch { }
            try { find.IgnoreMessage = 1; find.FindRegExp = 0; find.FindType = 1; } catch { }
            while (count < 1000 && (bool)action.Execute("RepeatFind", find.HSet))
            {
                if (!ApplyParagraphShape(hwp, style)) break;
                count++;
            }
        }
        finally
        {
            try { hwp.HAction.Run("Cancel"); } catch { }
            try { hwp.HAction.Run("MoveDocBegin"); } catch { }
            try { hwp.SetMessageBoxMode(previousMessageMode ?? 0xFFFFF); } catch { }
        }
        return count;
    }

    private static int ApplyParagraphFormatTarget(dynamic hwp, JsonObject op)
    {
        var style = Json.GetObj(op, "style") ?? throw new ArgumentException("set_paragraph_format.style이 필요합니다");
        var target = Json.GetObj(op, "target");
        var targetText = Json.GetString(target, "text");
        var scope = (Json.GetString(target, "scope") ?? "selection").ToLowerInvariant();
        if (!string.IsNullOrEmpty(targetText)) return ApplyParagraphShapeToTextMatches(hwp, targetText, style);
        if (scope == "document")
        {
            hwp.HAction.Run("MoveDocBegin");
            if (!(bool)hwp.HAction.Run("SelectAll")) return 0;
            var ok = ApplyParagraphShape(hwp, style);
            try { hwp.HAction.Run("Cancel"); hwp.HAction.Run("MoveDocBegin"); } catch { }
            return ok ? 1 : 0;
        }
        if (scope is not ("selection" or "paragraph"))
            throw new ArgumentException("target.scope은 selection|paragraph|document 중 하나여야 합니다");
        return ApplyParagraphShape(hwp, style) ? 1 : 0;
    }

    internal static void ValidateFormatParagraphItems(JsonObject op)
    {
        var items = Json.GetArr(op, "items") ??
            throw new ArgumentException("format_paragraphs.items 배열이 필요합니다");
        if (items.Count is < 1 or > 100)
            throw new ArgumentOutOfRangeException("items", "format_paragraphs.items는 1~100개입니다");
        for (var index = 0; index < items.Count; index++)
        {
            if (items[index] is not JsonObject item)
                throw new ArgumentException($"format_paragraphs.items[{index}]는 객체여야 합니다");
            var characterStyle = Json.GetObj(item, "characterStyle");
            var paragraphStyle = Json.GetObj(item, "paragraphStyle");
            if (characterStyle is null && paragraphStyle is null)
                throw new ArgumentException($"items[{index}]에는 characterStyle 또는 paragraphStyle이 필요합니다");
            if (characterStyle is not null) ValidateCharacterStyle(characterStyle);
            if (paragraphStyle is not null) ValidateParagraphStyle(paragraphStyle);

            var target = Json.GetObj(item, "target");
            var targetText = Json.GetString(target, "text");
            var scope = (Json.GetString(target, "scope") ?? "selection").ToLowerInvariant();
            if (string.IsNullOrEmpty(targetText) && scope is not ("selection" or "paragraph" or "document"))
                throw new ArgumentException($"items[{index}].target.scope은 selection|paragraph|document 중 하나여야 합니다");
        }
    }

    private static IReadOnlyList<HwpFormatResult> ExecFormatParagraphs(dynamic hwp, JsonObject op)
    {
        ValidateFormatParagraphItems(op);
        var results = new List<HwpFormatResult>();
        foreach (var node in Json.GetArr(op, "items")!)
        {
            var item = (JsonObject)node!;
            var characterStyle = Json.GetObj(item, "characterStyle");
            var paragraphStyle = Json.GetObj(item, "paragraphStyle");
            var target = Json.GetObj(item, "target");
            var targetText = Json.GetString(target, "text");
            var scope = (Json.GetString(target, "scope") ?? "selection").ToLowerInvariant();
            int count;
            if (!string.IsNullOrEmpty(targetText))
            {
                count = ApplyCombinedStylesToTextMatches(hwp, targetText, characterStyle, paragraphStyle);
            }
            else if (scope == "document")
            {
                hwp.HAction.Run("MoveDocBegin");
                count = (bool)hwp.HAction.Run("SelectAll") &&
                        ApplyCombinedStyles(hwp, characterStyle, paragraphStyle) ? 1 : 0;
                try { hwp.HAction.Run("Cancel"); hwp.HAction.Run("MoveDocBegin"); } catch { }
            }
            else
            {
                count = ApplyCombinedStyles(hwp, characterStyle, paragraphStyle) ? 1 : 0;
            }
            results.Add(new HwpFormatResult(
                !string.IsNullOrEmpty(targetText) ? $"text:{targetText}" : $"scope:{scope}", count));
            if (count == 0) break;
        }
        return results;
    }

    private static bool ApplyCombinedStyles(dynamic hwp, JsonObject? characterStyle, JsonObject? paragraphStyle)
    {
        if (characterStyle is not null &&
            (!ApplyCharShape(hwp, characterStyle) || !ApplyParagraphAlignment(hwp, characterStyle)))
            return false;
        return paragraphStyle is null || ApplyParagraphShape(hwp, paragraphStyle);
    }

    private static int ApplyCombinedStylesToTextMatches(
        dynamic hwp, string targetText, JsonObject? characterStyle, JsonObject? paragraphStyle)
    {
        if (string.IsNullOrEmpty(targetText)) return 0;
        var count = 0;
        int? previousMessageMode = null;
        try
        {
            try { previousMessageMode = Convert.ToInt32(hwp.GetMessageBoxMode()); } catch { }
            try { hwp.SetMessageBoxMode(0x2FFF1); } catch { }
            hwp.HAction.Run("MoveDocBegin");
            dynamic action = hwp.HAction;
            dynamic find = hwp.HParameterSet.HFindReplace;
            action.GetDefault("FindDlg", find.HSet);
            find.FindString = targetText;
            try { find.Direction = hwp.FindDir("Forward"); } catch { }
            try { find.MatchCase = 1; find.SeveralWords = 0; find.UseWildCards = 0; find.WholeWordOnly = 0; } catch { }
            try { find.IgnoreMessage = 1; find.FindRegExp = 0; find.FindType = 1; } catch { }
            while (count < 1000 && (bool)action.Execute("RepeatFind", find.HSet))
            {
                if (!ApplyCombinedStyles(hwp, characterStyle, paragraphStyle)) break;
                count++;
            }
        }
        finally
        {
            try { hwp.HAction.Run("Cancel"); } catch { }
            try { hwp.HAction.Run("MoveDocBegin"); } catch { }
            try { hwp.SetMessageBoxMode(previousMessageMode ?? 0xFFFFF); } catch { }
        }
        return count;
    }

    private static HwpWriteResult ExecSetPageSetup(dynamic hwp, JsonObject op)
    {
        var page = Json.GetObj(op, "page") ?? throw new ArgumentException("set_page_setup.page가 필요합니다");
        dynamic section = hwp.HParameterSet.HSecDef;
        hwp.HAction.GetDefault("PageSetup", section.HSet);
        dynamic def = section.PageDef;
        var before = $"{Convert.ToDouble(def.PaperWidth) / 283.465:0.##}x{Convert.ToDouble(def.PaperHeight) / 283.465:0.##}mm";

        if (TryJsonNumber(page, "widthMm", out var width)) def.PaperWidth = hwp.MiliToHwpUnit(width);
        if (TryJsonNumber(page, "heightMm", out var height)) def.PaperHeight = hwp.MiliToHwpUnit(height);
        if (Json.GetString(page, "orientation") is { Length: > 0 } orientation)
            def.Landscape = orientation.ToLowerInvariant() switch
            {
                "portrait" => 0,
                "landscape" => 1,
                _ => throw new ArgumentException("orientation은 portrait|landscape 중 하나여야 합니다"),
            };
        if (TryPageLength((object)hwp, page, "leftMarginMm", out var leftMargin)) def.LeftMargin = leftMargin;
        if (TryPageLength((object)hwp, page, "rightMarginMm", out var rightMargin)) def.RightMargin = rightMargin;
        if (TryPageLength((object)hwp, page, "topMarginMm", out var topMargin)) def.TopMargin = topMargin;
        if (TryPageLength((object)hwp, page, "bottomMarginMm", out var bottomMargin)) def.BottomMargin = bottomMargin;
        if (TryPageLength((object)hwp, page, "headerMm", out var headerLength)) def.HeaderLen = headerLength;
        if (TryPageLength((object)hwp, page, "footerMm", out var footerLength)) def.FooterLen = footerLength;
        if (TryPageLength((object)hwp, page, "gutterMm", out var gutterLength)) def.GutterLen = gutterLength;
        var applyTo = (Json.GetString(op, "applyTo") ?? "current-section").ToLowerInvariant() switch
        {
            "selection" => 1,
            "current-section" => 2,
            "document" => 3,
            "new-section" => 4,
            _ => throw new ArgumentException("applyTo는 selection|current-section|document|new-section 중 하나여야 합니다"),
        };
        section.HSet.SetItem("ApplyTo", applyTo);
        var ok = (bool)hwp.HAction.Execute("PageSetup", section.HSet);
        var after = $"{Convert.ToDouble(def.PaperWidth) / 283.465:0.##}x{Convert.ToDouble(def.PaperHeight) / 283.465:0.##}mm";
        return new HwpWriteResult(ok, "page-setup", $"page setup {before} -> {after}", before, after);
    }

    private static bool TryPageLength(object hwpObject, JsonObject page, string key, out int value)
    {
        value = 0;
        if (!TryJsonNumber(page, key, out var millimeters)) return false;
        if (millimeters < 0) throw new ArgumentOutOfRangeException(key, "쪽 여백/길이는 0 이상이어야 합니다");
        dynamic hwp = hwpObject;
        value = Convert.ToInt32(hwp.MiliToHwpUnit(millimeters));
        return true;
    }

    private static HwpWriteResult ExecInsertBreak(dynamic hwp, JsonObject op)
    {
        var type = (Json.GetString(op, "type") ?? "page").ToLowerInvariant();
        var action = type switch
        {
            "line" => "BreakLine",
            "paragraph" => "BreakPara",
            "page" => "BreakPage",
            "section" => "BreakSection",
            "column" => "BreakColumn",
            _ => throw new ArgumentException("type은 line|paragraph|page|section|column 중 하나여야 합니다"),
        };
        var ok = (bool)hwp.HAction.Run(action);
        return new HwpWriteResult(ok, $"break:{type}", $"inserted {type} break");
    }

    private static dynamic? FindControl(dynamic hwp, string controlId, int zeroBasedIndex)
    {
        if (zeroBasedIndex < 0) throw new ArgumentOutOfRangeException(nameof(zeroBasedIndex));
        dynamic? ctrl = null;
        try { ctrl = hwp.HeadCtrl; } catch { }
        var found = 0;
        while (ctrl is not null)
        {
            string id = "";
            try { id = (string)(ctrl.CtrlID ?? ""); } catch { }
            if (string.Equals(id, controlId, StringComparison.OrdinalIgnoreCase))
            {
                if (found == zeroBasedIndex) return ctrl;
                found++;
            }
            try { ctrl = ctrl.Next; } catch { ctrl = null; }
        }
        return null;
    }

    private static IReadOnlyList<object> FindControls(object hwpObject, string controlId)
    {
        dynamic hwp = hwpObject;
        var controls = new List<object>();
        dynamic? ctrl = null;
        try { ctrl = hwp.HeadCtrl; } catch { }
        while (ctrl is not null)
        {
            string id = "";
            try { id = Convert.ToString(ctrl.CtrlID) ?? ""; } catch { }
            if (string.Equals(id, controlId, StringComparison.OrdinalIgnoreCase))
                controls.Add((object)ctrl);
            try { ctrl = ctrl.Next; } catch { ctrl = null; }
        }
        return controls;
    }

    private static HwpWriteResult ExecTableCellSetText(dynamic hwp, JsonObject op) =>
        ExecTableCellSetTextCore((object)hwp, op, null, null);

    private static HwpWriteResult ExecTableCellSetTextCore(
        object hwpObject, JsonObject op, object? tableControl,
        IReadOnlySet<(int List, int Para)>? formulaPositions)
    {
        dynamic hwp = hwpObject;
        var tableIndex = Json.GetInt(op, "tableIndex") ?? 0;
        var cellIndex = Json.GetInt(op, "cellIndex");
        var row = Json.GetInt(op, "row") ?? 0;
        var col = Json.GetInt(op, "col") ?? 0;
        var text = Json.GetString(op, "text") ?? "";
        if (row < 0 || col < 0 || cellIndex < 0)
            throw new ArgumentException("row, col, cellIndex는 0부터 시작하는 값이어야 합니다");
        if (cellIndex is null && !op.ContainsKey("row") && !op.ContainsKey("col"))
            throw new ArgumentException("table_cell_set_text에는 cellIndex 또는 row+col이 필요합니다");

        var locator = cellIndex is null ? $"cell:{row},{col}" : $"cellIndex:{cellIndex}";
        var reference = $"table:{tableIndex}/{locator}";
        var read = tableControl is null
            ? TryReadTableCell(hwpObject, tableIndex, row, col, cellIndex, out var target, out var readError)
            : TryReadTableCellOnControl(hwpObject, tableControl, tableIndex, row, col, cellIndex,
                out target, out readError, formulaPositions);
        if (!read)
            return new HwpWriteResult(false, reference, readError);
        if (target!.HasFormula)
            return new HwpWriteResult(false, reference,
                "대상 셀에는 한글 수식 컨트롤(%fmu)이 있습니다(hasFormula=true). 수식 셀을 일반 텍스트로 덮어쓰지 않았습니다.");
        var before = target.Text;

        HwpNativeStyle? chosenStyle = null;
        if (PreserveStyle(op))
            chosenStyle = ResolveTableContextStyle((object)hwp, op, tableIndex, row, col, cellIndex, target);

        // 문맥 탐색 과정에서 캐럿이 이동하므로 대상 셀을 다시 정확히 선택한다.
        var selected = tableControl is null
            ? SelectTableCellBlock(hwpObject, tableIndex, row, col, cellIndex, out var selectError)
            : SelectTableCellBlockOnControl(hwpObject, tableControl, tableIndex, row, col, cellIndex, out selectError);
        if (!selected)
            return new HwpWriteResult(false, reference, selectError, before, text);
        try { hwp.HAction.Run("Cancel"); } catch { }
        hwp.HAction.Run("SelectAll");
        if (!string.IsNullOrEmpty(before))
        {
            // Delete도 선택 텍스트를 정상 삭제하고 false를 반환하는 버전이 있어 적용 후 정확한 셀 내용으로 검증한다.
            hwp.HAction.Run("Delete");
        }
        else
        {
            try { hwp.HAction.Run("Cancel"); } catch { }
        }

        if (chosenStyle is not null) _ = ApplyNativeStyle(hwp, chosenStyle);
        if (!ApplyExplicitWriteStyle(hwp, op))
            return new HwpWriteResult(false, reference, "명시한 글자/문단 서식 적용 실패", before, text);
        if (!ExecInsertText(hwp, text))
            return new HwpWriteResult(false, reference, "셀 텍스트 입력 실패", before, text);
        hwp.HAction.Run("SelectAll");
        var after = GetSelectionText(hwp);
        try { hwp.HAction.Run("Cancel"); } catch { }
        var normalizedAfter = NormalizeCellText(after);
        var normalizedExpected = NormalizeCellText(text);
        var verified = string.Equals(normalizedAfter, normalizedExpected, StringComparison.Ordinal);
        var detail = verified
            ? $"cell text replaced; style={(chosenStyle?.Source ?? (Json.GetObj(op, "style") is null ? "not-preserved" : "explicit"))}"
            : $"cell readback mismatch (expected='{normalizedExpected}', actual='{normalizedAfter}', before='{NormalizeCellText(before)}')";
        return new HwpWriteResult(verified, reference, detail, before, after);
    }

    private static IReadOnlyList<HwpWriteResult> ExecTableSetCells(dynamic hwp, JsonObject op)
    {
        var tableIndex = Json.GetInt(op, "tableIndex") ?? 0;
        var cells = Json.GetArr(op, "cells") ??
            throw new ArgumentException("table_set_cells.cells 배열이 필요합니다");
        if (cells.Count is < 1 or > 500)
            throw new ArgumentOutOfRangeException("cells", "table_set_cells.cells는 1~500개입니다");
        dynamic? table = FindControl(hwp, "tbl", tableIndex);
        if (table is null)
            return new[] { new HwpWriteResult(false, $"table:{tableIndex}", $"표 {tableIndex}을 찾을 수 없습니다") };

        var defaultPreserveStyle = Json.GetBool(op, "preserveStyle", true);
        var formulaPositions = FormulaCellPositions((object)hwp);
        var results = new List<HwpWriteResult>(cells.Count);
        foreach (var node in cells)
        {
            if (node is not JsonObject cell)
                throw new ArgumentException("table_set_cells.cells의 각 항목은 객체여야 합니다");
            var single = (JsonObject)cell.DeepClone();
            single["tableIndex"] = tableIndex;
            if (!single.ContainsKey("preserveStyle")) single["preserveStyle"] = defaultPreserveStyle;
            if (!single.ContainsKey("style") && Json.GetObj(op, "style") is { } style)
                single["style"] = style.DeepClone();
            var result = ExecTableCellSetTextCore((object)hwp, single, (object)table, formulaPositions);
            results.Add(result);
            if (!result.Ok) break;
        }
        return results;
    }

    private static string NormalizeCellText(string value) =>
        NormalizeNewlines(value).Trim('\n', '\0', '\u0002', '\u0003', ' ');

    private static HwpWriteResult ExecInsertPicture(dynamic hwp, JsonObject op)
    {
        var path = Path.GetFullPath(Json.GetString(op, "path") ?? throw new ArgumentException("insert_picture.path가 필요합니다"));
        if (!File.Exists(path)) throw new FileNotFoundException("삽입할 그림을 찾을 수 없습니다", path);
        var tableIndex = Json.GetInt(op, "tableIndex");
        var row = Json.GetInt(op, "row") ?? 0;
        var col = Json.GetInt(op, "col") ?? 0;
        var cellIndex = Json.GetInt(op, "cellIndex");
        var reference = "picture:new";
        if (tableIndex is not null)
        {
            if (!SelectTableCellBlock((object)hwp, tableIndex.Value, row, col, cellIndex, out var error))
                return new HwpWriteResult(false, $"table:{tableIndex}", error);
            reference = cellIndex is null
                ? $"table:{tableIndex}/cell:{row},{col}/picture"
                : $"table:{tableIndex}/cellIndex:{cellIndex}/picture";
            try { hwp.HAction.Run("Cancel"); } catch { }
            if (Json.GetBool(op, "clearCell"))
            {
                _ = hwp.HAction.Run("SelectAll");
                _ = hwp.HAction.Run("Delete");
                try { hwp.HAction.Run("Cancel"); } catch { }
            }
        }
        var sizeOptionName = Json.GetString(op, "sizeOption") ??
            (tableIndex is null ? "real" : "cell-ratio");
        var sizeOption = sizeOptionName.ToLowerInvariant() switch
        {
            "real" => 0,
            "specific" => 1,
            "cell" => 2,
            "cell-ratio" => 3,
            _ => throw new ArgumentException("sizeOption은 real|specific|cell|cell-ratio 중 하나여야 합니다"),
        };
        var width = TryJsonNumber(op, "widthMm", out var widthMm) ? widthMm : 0;
        var height = TryJsonNumber(op, "heightMm", out var heightMm) ? heightMm : 0;
        if (sizeOption == 1 && (width <= 0 || height <= 0))
            throw new ArgumentException("specific 크기에는 양수 widthMm과 heightMm이 필요합니다");
        var effect = (Json.GetString(op, "effect") ?? "original").ToLowerInvariant() switch
        {
            "original" => 0,
            "grayscale" => 1,
            "black-white" => 2,
            _ => throw new ArgumentException("effect는 original|grayscale|black-white 중 하나여야 합니다"),
        };
        var before = CountPictureControls(hwp);
        dynamic? inserted = hwp.InsertPicture(path, Json.GetBool(op, "embedded", true), sizeOption,
            Json.GetBool(op, "reverse"), Json.GetBool(op, "watermark"), effect, width, height);
        var after = CountPictureControls(hwp);
        string insertedId = "";
        try { insertedId = (string)(inserted?.CtrlID ?? ""); } catch { }
        // 최신 한글은 그림을 일반 shape-object 컨트롤 ID "gso"로 반환한다.
        // 따라서 $pic 하나만 센 0->0을 성공으로 오인하지 않고 두 실제 ID의 합계를 검증한다.
        var returnedPictureControl = insertedId is "$pic" or "gso";
        var ok = inserted is not null && returnedPictureControl && after == before + 1;
        return new HwpWriteResult(ok, reference,
            $"picture controls($pic+gso, document scope) {before}->{after}, returned={insertedId}, exact delta={(after - before)}",
            before.ToString(), after.ToString());
    }

    private static int CountPictureControls(dynamic hwp) =>
        CountControlId(hwp, "$pic") + CountControlId(hwp, "gso");

    private static int CountControlId(dynamic hwp, string controlId)
    {
        var count = 0;
        dynamic? ctrl = null;
        try { ctrl = hwp.HeadCtrl; } catch { }
        while (ctrl is not null)
        {
            try { if (string.Equals((string)(ctrl.CtrlID ?? ""), controlId, StringComparison.OrdinalIgnoreCase)) count++; } catch { }
            try { ctrl = ctrl.Next; } catch { ctrl = null; }
        }
        return count;
    }

    private static HwpWriteResult ExecInsertPageNumber(dynamic hwp, JsonObject op)
    {
        dynamic pageNum = hwp.HParameterSet.HPageNumPos;
        hwp.HAction.GetDefault("PageNumPos", pageNum.HSet);
        pageNum.DrawPos = (Json.GetString(op, "position") ?? "bottom-center").ToLowerInvariant() switch
        {
            "none" => 0,
            "top-left" => 1,
            "top-center" => 2,
            "top-right" => 3,
            "bottom-left" => 4,
            "bottom-center" => 5,
            "bottom-right" => 6,
            "top-outside" => 7,
            "bottom-outside" => 8,
            "top-inside" => 9,
            "bottom-inside" => 10,
            _ => throw new ArgumentException("지원하지 않는 쪽 번호 위치입니다"),
        };
        pageNum.NumberFormat = (Json.GetString(op, "format") ?? "arabic").ToLowerInvariant() switch
        {
            "arabic" => 0,
            "circled" => 1,
            "roman-upper" => 2,
            "roman-lower" => 3,
            "alpha-upper" => 4,
            "hangul" => 8,
            "chinese" => 13,
            _ => throw new ArgumentException("지원하지 않는 쪽 번호 형식입니다"),
        };
        pageNum.NewNumber = Json.GetInt(op, "startNumber") ?? 1;
        var before = CountControlId(hwp, "pgnp");
        var ok = (bool)hwp.HAction.Execute("PageNumPos", pageNum.HSet);
        var after = CountControlId(hwp, "pgnp");
        return new HwpWriteResult(ok && after >= before, "page-number", $"page-number controls {before}->{after}", before.ToString(), after.ToString());
    }

    private static HwpWriteResult ExecSetHeaderFooterText(dynamic hwp, JsonObject op)
    {
        var kind = (Json.GetString(op, "kind") ?? "header").ToLowerInvariant();
        var text = Json.GetString(op, "text") ?? "";
        dynamic headerFooter = hwp.HParameterSet.HHeaderFooter;
        hwp.HAction.GetDefault("HeaderFooter", headerFooter.HSet);
        var kindValue = kind switch
        {
            "header" => 0,
            "footer" => 1,
            _ => throw new ArgumentException("kind는 header|footer 중 하나여야 합니다"),
        };
        var pageType = (Json.GetString(op, "pages") ?? "both").ToLowerInvariant() switch
        {
            "both" => 0,
            "even" => 1,
            "odd" => 2,
            _ => throw new ArgumentException("pages는 both|even|odd 중 하나여야 합니다"),
        };
        headerFooter.HSet.SetItem("HeaderFooterCtrlType", kindValue);
        headerFooter.HSet.SetItem("Type", pageType);
        var controlId = kind == "header" ? "head" : "foot";
        var before = CountControlId(hwp, controlId);
        if (!(bool)hwp.HAction.Execute("HeaderFooter", headerFooter.HSet))
            return new HwpWriteResult(false, kind, $"{kind} control create failed");
        hwp.HAction.Run("MoveListBegin");
        hwp.HAction.Run("MoveSelListEnd");
        var oldText = GetSelectionText(hwp);
        if (!ExecInsertText(hwp, text)) return new HwpWriteResult(false, kind, $"{kind} text insert failed", oldText, text);
        try { hwp.HAction.Run("CloseEx"); } catch { try { hwp.HAction.Run("Cancel"); } catch { } }
        var after = CountControlId(hwp, controlId);
        return new HwpWriteResult(after >= before, kind, $"{kind} controls {before}->{after}", oldText, text);
    }

    private static HwpWriteResult ExecExportPdf(dynamic hwp, JsonObject op)
    {
        var output = Path.GetFullPath(Json.GetString(op, "output") ?? throw new ArgumentException("export_pdf.output이 필요합니다"));
        if (!output.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("export_pdf.output은 .pdf 파일이어야 합니다");
        Directory.CreateDirectory(Path.GetDirectoryName(output)!);
        var ok = (bool)hwp.SaveAs(output, "PDF", "");
        var exists = File.Exists(output) && new FileInfo(output).Length > 0;
        return new HwpWriteResult(ok && exists, "pdf:export", exists ? $"exported {new FileInfo(output).Length} bytes" : "PDF output missing", null, output);
    }

    private static void ValidateCharacterStyle(JsonObject style)
    {
        if (TryJsonNumber(style, "fontSize", out var fontSize) && fontSize is < 1 or > 4096)
            throw new ArgumentOutOfRangeException(nameof(fontSize), "fontSize는 1~4096pt입니다");
        if (Json.GetString(style, "textColor") is { Length: > 0 } textColor) _ = ToHwpColorRef(textColor);
        if (Json.GetString(style, "shadeColor") is { Length: > 0 } shadeColor) _ = ToHwpColorRef(shadeColor);
        if (Json.GetString(style, "underlineColor") is { Length: > 0 } underlineColor) _ = ToHwpColorRef(underlineColor);
        if (Json.GetString(style, "strikeoutColor") is { Length: > 0 } strikeoutColor) _ = ToHwpColorRef(strikeoutColor);
        if (style["underline"] is JsonValue underline && underline.TryGetValue<string>(out var underlineName)) _ = UnderlineType(underlineName);
        if (style["strikeout"] is JsonValue strikeout && strikeout.TryGetValue<string>(out var strikeoutName)) _ = StrikeOutType(strikeoutName);
        if (style["letterSpacing"] is JsonValue spacing && spacing.TryGetValue<int>(out var spacingValue) && spacingValue is < -50 or > 50)
            throw new ArgumentOutOfRangeException(nameof(spacingValue), "letterSpacing은 -50~50입니다");
        if (style["widthRatio"] is JsonValue ratio && ratio.TryGetValue<int>(out var ratioValue) && ratioValue is < 50 or > 200)
            throw new ArgumentOutOfRangeException(nameof(ratioValue), "widthRatio는 50~200입니다");
        if (style["offset"] is JsonValue offset && offset.TryGetValue<int>(out var offsetValue) && offsetValue is < -100 or > 100)
            throw new ArgumentOutOfRangeException(nameof(offsetValue), "offset은 -100~100입니다");
        if (Json.GetString(style, "align") is { Length: > 0 } align) _ = ParagraphAlignType(align);
    }

    private static void ValidateParagraphStyle(JsonObject style)
    {
        if (Json.GetString(style, "align") is { Length: > 0 } align) _ = ParagraphAlignType(align);
        if (TryJsonNumber(style, "lineSpacingPercent", out var spacing) && spacing is < 50 or > 500)
            throw new ArgumentOutOfRangeException(nameof(spacing), "lineSpacingPercent는 50~500입니다");
        foreach (var key in new[] { "spaceBeforePt", "spaceAfterPt" })
            if (TryJsonNumber(style, key, out var value) && value < 0)
                throw new ArgumentOutOfRangeException(key, "문단 간격은 0 이상이어야 합니다");
    }

    private static void ValidatePageSetup(JsonObject op)
    {
        var page = Json.GetObj(op, "page") ?? throw new ArgumentException("set_page_setup.page가 필요합니다");
        foreach (var key in new[] { "widthMm", "heightMm" })
            if (TryJsonNumber(page, key, out var value) && value <= 0)
                throw new ArgumentOutOfRangeException(key, "용지 크기는 0보다 커야 합니다");
        foreach (var key in new[] { "leftMarginMm", "rightMarginMm", "topMarginMm", "bottomMarginMm", "headerMm", "footerMm", "gutterMm" })
            if (TryJsonNumber(page, key, out var value) && value < 0)
                throw new ArgumentOutOfRangeException(key, "쪽 여백/길이는 0 이상이어야 합니다");
        if (Json.GetString(page, "orientation") is { Length: > 0 } orientation && orientation.ToLowerInvariant() is not ("portrait" or "landscape"))
            throw new ArgumentException("orientation은 portrait|landscape 중 하나여야 합니다");
        if ((Json.GetString(op, "applyTo") ?? "current-section").ToLowerInvariant() is not ("selection" or "current-section" or "document" or "new-section"))
            throw new ArgumentException("applyTo는 selection|current-section|document|new-section 중 하나여야 합니다");
    }

    private static void ValidateBreak(JsonObject op)
    {
        if ((Json.GetString(op, "type") ?? "page").ToLowerInvariant() is not ("line" or "paragraph" or "page" or "section" or "column"))
            throw new ArgumentException("type은 line|paragraph|page|section|column 중 하나여야 합니다");
    }

    private static void ValidatePicture(JsonObject op)
    {
        var path = Path.GetFullPath(Json.GetString(op, "path") ?? throw new ArgumentException("insert_picture.path가 필요합니다"));
        if (!File.Exists(path)) throw new FileNotFoundException("삽입할 그림을 찾을 수 없습니다", path);
        var tableIndex = Json.GetInt(op, "tableIndex");
        var row = Json.GetInt(op, "row") ?? 0;
        var col = Json.GetInt(op, "col") ?? 0;
        var cellIndex = Json.GetInt(op, "cellIndex");
        if (tableIndex < 0 || row < 0 || col < 0 || cellIndex < 0)
            throw new ArgumentException("tableIndex, row, col, cellIndex는 0 이상이어야 합니다");
        if (tableIndex is null && (op.ContainsKey("row") || op.ContainsKey("col") || cellIndex is not null || Json.GetBool(op, "clearCell")))
            throw new ArgumentException("row, col, cellIndex, clearCell은 tableIndex와 함께 사용해야 합니다");
        if (cellIndex is not null && (op.ContainsKey("row") || op.ContainsKey("col")))
            throw new ArgumentException("insert_picture 표 셀은 cellIndex 또는 row+col 중 한 방식으로 지정하세요");
        var option = (Json.GetString(op, "sizeOption") ??
            (tableIndex is null ? "real" : "cell-ratio")).ToLowerInvariant();
        if (option is not ("real" or "specific" or "cell" or "cell-ratio")) throw new ArgumentException("sizeOption은 real|specific|cell|cell-ratio 중 하나여야 합니다");
        if (option == "specific" && (!TryJsonNumber(op, "widthMm", out var width) || width <= 0 || !TryJsonNumber(op, "heightMm", out var height) || height <= 0))
            throw new ArgumentException("specific 크기에는 양수 widthMm과 heightMm이 필요합니다");
        if ((Json.GetString(op, "effect") ?? "original").ToLowerInvariant() is not ("original" or "grayscale" or "black-white"))
            throw new ArgumentException("effect는 original|grayscale|black-white 중 하나여야 합니다");
    }

    private static void ValidatePageNumber(JsonObject op)
    {
        if ((Json.GetString(op, "position") ?? "bottom-center").ToLowerInvariant() is not
            ("none" or "top-left" or "top-center" or "top-right" or "bottom-left" or "bottom-center" or "bottom-right" or "top-outside" or "bottom-outside" or "top-inside" or "bottom-inside"))
            throw new ArgumentException("지원하지 않는 쪽 번호 위치입니다");
        if ((Json.GetString(op, "format") ?? "arabic").ToLowerInvariant() is not
            ("arabic" or "circled" or "roman-upper" or "roman-lower" or "alpha-upper" or "hangul" or "chinese"))
            throw new ArgumentException("지원하지 않는 쪽 번호 형식입니다");
        if ((Json.GetInt(op, "startNumber") ?? 1) < 1) throw new ArgumentException("startNumber는 1 이상이어야 합니다");
    }

    private static void ValidateHeaderFooter(JsonObject op)
    {
        if ((Json.GetString(op, "kind") ?? "header").ToLowerInvariant() is not ("header" or "footer")) throw new ArgumentException("kind는 header|footer 중 하나여야 합니다");
        if ((Json.GetString(op, "pages") ?? "both").ToLowerInvariant() is not ("both" or "even" or "odd")) throw new ArgumentException("pages는 both|even|odd 중 하나여야 합니다");
    }

    private static void ValidateExportPdf(JsonObject op)
    {
        var output = Path.GetFullPath(Json.GetString(op, "output") ?? throw new ArgumentException("export_pdf.output이 필요합니다"));
        if (!output.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("export_pdf.output은 .pdf 파일이어야 합니다");
    }
}
