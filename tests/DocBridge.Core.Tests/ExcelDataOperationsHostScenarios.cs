using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

/// <summary>
/// Host-based Excel scenario inputs for the materials ledger and monthly
/// chart report. Values/formulas/layout must be written first by integration
/// or the probe via existing set_values ops. These batches are data-only.
/// </summary>
public static class ExcelDataOperationsHostScenarios
{
    public const string MaterialsSheet = "자재대장";
    public const string ReportSheet = "월간보고";
    public const string TableName = "Materials";
    public const string ChartName = "MonthlyOutput";
    public const string PictureName = "ReportLogo";

    public static string SamplePicturePath { get; } =
        @"C:\DocBridgeTest\Desktop\plug-in\output\docbridge-production-20260910\cursor-data\fixtures\logo.png";

    public static JsonObject MaterialsSeedValues() => Json.ParseObject("""
        {
          "op": "set_values",
          "target": { "sheet": "자재대장" },
          "range": "A1:G8",
          "values": [
            ["일자", "전표", "구분", "품목", "단위", "수량", "재고"],
            ["2026-09-01", "IN-001", "입고", "관부속", "EA", 120, 120],
            ["2026-09-02", "OUT-014", "출고", "관부속", "EA", 18, 102],
            ["2026-09-03", "IN-002", "입고", "맨홀뚜껑", "SET", 6, 6],
            ["2026-09-04", "MOVE-3", "이동", "관부속", "EA", 10, 92],
            ["2026-09-05", "OUT-015", "출고", "맨홀뚜껑", "SET", 1, 5],
            ["2026-09-06", "IN-003", "입고", "관부속", "EA", 40, 132],
            ["2026-09-07", "OUT-016", "출고", "관부속", "EA", 25, 107]
          ]
        }
        """)!;

    public static JsonObject ReportSeedValues() => Json.ParseObject("""
        {
          "op": "set_values",
          "target": { "sheet": "월간보고" },
          "range": "A1:B6",
          "values": [
            ["월", "실적"],
            ["4월", 210],
            ["5월", 245],
            ["6월", 198],
            ["7월", 276],
            ["8월", 254]
          ]
        }
        """)!;

    public static IReadOnlyList<JsonObject> MaterialsLedgerOps() => new[]
    {
        Op("""
        { "op": "create_table", "target": { "sheet": "자재대장" },
          "range": "A1:G8", "name": "Materials", "hasHeaders": true,
          "styleName": "TableStyleMedium2", "showTotals": false }
        """),
        Op("""
        { "op": "style_table", "target": { "sheet": "자재대장" },
          "name": "Materials", "showRowStripes": true, "showAutoFilter": true }
        """),
        Op("""
        { "op": "set_table_totals", "target": { "sheet": "자재대장" },
          "name": "Materials", "showTotals": true,
          "columns": [
            { "column": "수량", "function": "sum" },
            { "column": "재고", "function": "sum" }
          ] }
        """),
        Op("""
        { "op": "set_data_validation", "target": { "sheet": "자재대장" },
          "range": "C2:C8", "type": "list", "source": "입고,출고,이동",
          "inCellDropdown": true, "ignoreBlank": true, "errorStyle": "stop",
          "errorTitle": "구분", "errorMessage": "입고/출고/이동만 입력하세요." }
        """),
        Op("""
        { "op": "set_data_validation", "target": { "sheet": "자재대장" },
          "range": "F2:F8", "type": "whole", "operator": "greater",
          "formula1": "0", "showError": true, "errorStyle": "warning",
          "errorMessage": "수량은 0보다 커야 합니다." }
        """),
        Op("""
        { "op": "add_conditional_format", "target": { "sheet": "자재대장" },
          "range": "G2:G8",
          "rule": { "type": "cellValue", "operator": "less", "formula1": "10" },
          "style": { "bold": true, "fontColor": "#9C0006", "fillColor": "#FFC7CE" } }
        """),
        Op("""
        { "op": "define_name", "name": "StockOnHand", "scope": "sheet",
          "target": { "sheet": "자재대장" },
          "refersTo": "='자재대장'!$G$2:$G$8", "comment": "재고 열" }
        """),
        Op("""
        { "op": "sort_table", "target": { "sheet": "자재대장" },
          "name": "Materials",
          "keys": [{ "column": "일자", "order": "asc" }, { "column": "전표", "order": "asc" }] }
        """),
        Op("""
        { "op": "set_auto_filter", "target": { "sheet": "자재대장" },
          "name": "Materials",
          "criteria": [{ "column": "구분", "operator": "equals", "value": "입고" }] }
        """),
        Op("""
        { "op": "set_cell_note", "target": { "sheet": "자재대장" },
          "cell": "B4", "text": "현장 검수 후 입고 확정", "visible": false }
        """),
        Op("""
        { "op": "set_hyperlink", "target": { "sheet": "자재대장" },
          "cell": "D2", "address": "https://example.invalid/spec/pipe-fitting",
          "textToDisplay": "관부속", "screenTip": "자재 시방" }
        """),
        Op("""
        { "op": "clear_auto_filter", "target": { "sheet": "자재대장" },
          "name": "Materials" }
        """),
    };

    public static IReadOnlyList<JsonObject> ChartReportOps()
    {
        var picture = new JsonObject
        {
            ["op"] = "insert_sheet_picture",
            ["target"] = new JsonObject { ["sheet"] = ReportSheet },
            ["path"] = SamplePicturePath,
            ["name"] = PictureName,
            ["lockAspectRatio"] = true,
            ["position"] = new JsonObject
            {
                ["left"] = 420, ["top"] = 12, ["width"] = 120, ["height"] = 48,
            },
        };
        return new[]
        {
            Op("""
            { "op": "create_chart", "target": { "sheet": "월간보고" },
              "sourceRange": "A1:B6", "chartType": "columnClustered",
              "name": "MonthlyOutput", "title": "월별 실적",
              "hasLegend": true, "legendPosition": "bottom", "plotBy": "columns",
              "position": { "left": 20, "top": 140, "width": 380, "height": 220 } }
            """),
            Op("""
            { "op": "update_chart", "target": { "sheet": "월간보고" },
              "name": "MonthlyOutput", "title": "2026년 월별 실적",
              "legendPosition": "right",
              "position": { "left": 24, "top": 148, "width": 400, "height": 240 } }
            """),
            picture,
            Op("""
            { "op": "update_picture", "target": { "sheet": "월간보고" },
              "name": "ReportLogo",
              "position": { "left": 430, "top": 16, "width": 110, "height": 44 } }
            """),
            Op("""
            { "op": "define_name", "name": "MonthlyActuals", "scope": "workbook",
              "refersTo": "='월간보고'!$B$2:$B$6" }
            """),
            Op("""
            { "op": "set_hyperlink", "target": { "sheet": "월간보고" },
              "cell": "A1", "subAddress": "'자재대장'!A1",
              "textToDisplay": "월", "screenTip": "자재대장으로 이동" }
            """),
        };
    }

    public static IReadOnlyList<JsonObject> ExpansionOps() => new[]
    {
        Op("""
        { "op": "remove_duplicates", "target": { "sheet": "자재대장" },
          "range": "A1:G8", "hasHeaders": true, "columns": [4] }
        """),
        Op("""
        { "op": "text_to_columns", "target": { "sheet": "자재대장" },
          "range": "B2:B8", "dataType": "delimited", "comma": false, "other": true, "otherChar": "-" }
        """),
        Op("""
        { "op": "create_pivot", "target": { "sheet": "월간보고" },
          "sourceRange": "'자재대장'!A1:G8", "destination": "E12", "name": "MaterialPivot",
          "rows": ["품목"], "columns": ["구분"],
          "values": [{ "field": "수량", "function": "sum" }] }
        """),
        Op("""
        { "op": "refresh_pivot", "target": { "sheet": "월간보고" }, "name": "MaterialPivot" }
        """),
        Op("""
        { "op": "insert_textbox", "target": { "sheet": "월간보고" },
          "name": "ReportNote", "text": "피벗 새로고침 후 확인",
          "position": { "left": 420, "top": 80, "width": 160, "height": 36 } }
        """),
        Op("""
        { "op": "insert_shape", "target": { "sheet": "월간보고" },
          "shapeType": "rectangle", "name": "HighlightBox",
          "position": { "left": 20, "top": 12, "width": 80, "height": 24 } }
        """),
        Op("""
        { "op": "add_table_column", "target": { "sheet": "자재대장" },
          "name": "Materials", "columnName": "검수", "formula": "=[@수량]" }
        """),
        Op("""
        { "op": "delete_chart", "target": { "sheet": "월간보고" }, "name": "MonthlyOutput" }
        """),
        Op("""
        { "op": "delete_pivot", "target": { "sheet": "월간보고" }, "name": "MaterialPivot" }
        """),
    };

    public static IEnumerable<JsonObject> AllOps()
    {
        foreach (var op in MaterialsLedgerOps()) yield return op;
        foreach (var op in ChartReportOps()) yield return op;
        foreach (var op in ExpansionOps()) yield return op;
    }

    public static JsonObject DryRunBatch(IEnumerable<JsonObject> ops) => new()
    {
        ["ops"] = new JsonArray(ops.Select(op => op.DeepClone()).ToArray()),
        ["dryRun"] = true,
    };

    public static JsonObject ApplyBatch(IEnumerable<JsonObject> ops, string confirmToken) => new()
    {
        ["ops"] = new JsonArray(ops.Select(op => op.DeepClone()).ToArray()),
        ["dryRun"] = false,
        ["confirmToken"] = confirmToken,
    };

    public static JsonObject InspectArgs(string sheet, string scope) => new()
    {
        ["sheet"] = sheet,
        ["scope"] = scope,
        ["limit"] = 200,
    };

    private static JsonObject Op(string json) =>
        Json.ParseObject(json) ?? throw new InvalidOperationException(json);
}
