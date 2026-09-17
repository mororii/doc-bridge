using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelDataOperationsSchemaTests
{
    [Fact]
    public void DescribeSchema_is_metadata_not_inputSchema()
    {
        var metadata = ExcelDataOperationsContract.DescribeSchema();
        Assert.Null(metadata["anyOf"]);
        Assert.Null(metadata["type"]);
        Assert.NotNull(metadata["writeOps"]);
        Assert.NotNull(metadata["fieldTypes"]);
        Assert.False(metadata.ContainsKey("properties"));
    }

    [Fact]
    public void AttachToApplyItems_publishes_leftover_plus_typed_branches()
    {
        var items = FlattenedItemsLikeToolRegistry();
        ExcelDataOperationSchema.AttachToApplyItems(items);

        Assert.False(items.ContainsKey("properties"));
        Assert.False(items.ContainsKey("required"));
        var anyOf = items["anyOf"]!.AsArray();
        Assert.Equal(1 + ExcelDataOperationsContract.WriteOpNames.Count, anyOf.Count);

        var leftover = anyOf[0]!.AsObject();
        var leftoverEnum = leftover["properties"]!["op"]!["enum"]!.AsArray()
            .Select(node => node!.GetValue<string>())
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        Assert.Contains("move_sheet", leftoverEnum);
        Assert.Contains("set_values", leftoverEnum);
        foreach (var name in ExcelDataOperationsContract.WriteOpNames)
            Assert.DoesNotContain(name, leftoverEnum);

        Assert.Equal("string", leftover["properties"]!["position"]!["type"]!.GetValue<string>());
        Assert.Equal("array", leftover["properties"]!["values"]!["items"]!["type"]!.GetValue<string>());
        Assert.Equal("object", leftover["properties"]!["rows"]!["items"]!["type"]!.GetValue<string>());

        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        for (var i = 1; i < anyOf.Count; i++)
        {
            var branch = anyOf[i]!.AsObject();
            Assert.False(branch["additionalProperties"]!.GetValue<bool>());
            var op = branch["properties"]!["op"]!["const"]!.GetValue<string>();
            Assert.True(names.Add(op), op);
        }
        Assert.Equal(ExcelDataOperationsContract.WriteOpNames.Count, names.Count);
        foreach (var name in ExcelDataOperationsContract.WriteOpNames)
            Assert.Contains(name, names);
    }

    [Fact]
    public void Published_anyOf_accepts_e1_chart_redline_and_rejects_leftover_position_string()
    {
        var items = PublishedItems();
        var e1 = Op(
            """
            { "op": "create_chart", "target": { "sheet": "예정공정표" },
              "sourceRange": "'예정공정표'!$G$95:$BP$95", "chartType": "line",
              "name": "Chart1", "title": "", "hasLegend": false, "plotBy": "rows",
              "position": { "left": 12, "top": 18, "width": 420, "height": 180 },
              "series": [{ "lineColor": "#FF0000", "lineWeight": 2.25, "marker": "none" }],
              "chartFill": "none", "plotFill": "none", "chartBorder": "none", "plotBorder": "none",
              "axes": { "category": { "visible": false }, "value": { "visible": false, "maximum": 100 } },
              "plotArea": { "spanChart": true } }
            """);

        var errors = new List<string>();
        Assert.True(ExcelDataOperationSchema.TryValidatePublishedItems(items, e1, errors),
            string.Join("; ", errors));
        Assert.Empty(errors);
        Assert.Equal(1, ExcelDataOperationSchema.CountPublishedMatches(items, e1));
        Assert.Empty(ExcelDataOperationsContract.ValidatePublicInput(e1));

        var chartBranch = FindBranch(items, "create_chart");
        Assert.Equal("object", chartBranch["properties"]!["position"]!["type"]!.GetValue<string>());
        Assert.Contains("left", chartBranch["properties"]!["position"]!["required"]!.AsArray()
            .Select(node => node!.GetValue<string>()));
        Assert.Equal("none", chartBranch["properties"]!["series"]!["items"]!["properties"]!["marker"]!["enum"]!
            .AsArray().First(node => node!.GetValue<string>() == "none")!.GetValue<string>());
        var seriesProps = chartBranch["properties"]!["series"]!["items"]!["properties"]!;
        Assert.Contains("lineMarkers", seriesProps["chartType"]!["enum"]!.AsArray()
            .Select(node => node!.GetValue<string>()));
        Assert.Contains("primary", seriesProps["axisGroup"]!["enum"]!.AsArray()
            .Select(node => node!.GetValue<string>()));
        Assert.Contains("secondary", seriesProps["axisGroup"]!["enum"]!.AsArray()
            .Select(node => node!.GetValue<string>()));
    }

    [Fact]
    public void Published_anyOf_accepts_combo_series_and_keeps_e1_style_only()
    {
        var items = PublishedItems();
        var combo = Op(
            """
            { "op": "create_chart", "target": { "sheet": "월간보고" },
              "sourceRange": "A1:C7", "chartType": "columnClustered", "name": "계획대비실적",
              "series": [
                { "values": "B2:B7", "categories": "A2:A7", "name": "금액",
                  "chartType": "columnClustered", "axisGroup": "primary" },
                { "values": "C2:C7", "categories": "A2:A7", "name": "진척률",
                  "chartType": "lineMarkers", "axisGroup": "secondary" }
              ],
              "axes": {
                "value": { "group": "primary", "numberFormat": "#,##0", "title": "금액" },
                "valueSecondary": { "group": "secondary", "minimum": 0, "maximum": 100, "numberFormat": "0\\%", "title": "진척률" }
              } }
            """);
        AssertPublishedOk(items, combo);
        Assert.Empty(ExcelDataOperationsContract.ValidatePublicInput(combo));
        var e1 = Op(
            """
            { "op": "create_chart", "target": { "sheet": "예정공정표" },
              "sourceRange": "'예정공정표'!$G$95:$BP$95", "chartType": "line",
              "name": "Chart1", "title": "", "hasLegend": false, "plotBy": "rows",
              "position": { "left": 12, "top": 18, "width": 420, "height": 180 },
              "series": [{ "lineColor": "#FF0000", "lineWeight": 2.25, "marker": "none" }],
              "chartFill": "none", "plotFill": "none", "chartBorder": "none", "plotBorder": "none",
              "axes": { "value": { "visible": false, "maximum": 100 } },
              "plotArea": { "spanChart": true } }
            """);
        AssertPublishedOk(items, e1);
        Assert.Empty(ExcelDataOperationsContract.ValidatePublicInput(e1));
        Assert.False(ExcelChartSeriesContract.HasComboOverride(e1["series"]!.AsArray()));
        var chartBranch = FindBranch(items, "create_chart");
        Assert.Contains("valueSecondary", chartBranch["properties"]!["axes"]!["properties"]!.AsObject()
            .Select(pair => pair.Key));
        AssertPublishedFails(items, Op(
            """
            { "op": "create_chart", "target": { "sheet": "월간보고" },
              "sourceRange": "A1:C7", "chartType": "columnClustered",
              "axes": { "valueSecondary": { "group": "tertiary" } } }
            """), "group");
        AssertPublishedFails(items, Op(
            """
            { "op": "create_chart", "target": { "sheet": "월간보고" },
              "sourceRange": "A1:C7", "chartType": "columnClustered",
              "axes": { "value": { "numberFormat": "" } } }
            """), "numberFormat");
    }

    [Fact]
    public void Published_anyOf_accepts_e4_table_and_e7_published_pivot()
    {
        var items = PublishedItems();
        var e4 = Op(
            """
            { "op": "create_table", "target": { "sheet": "자재대장" },
              "range": "A2:I5", "name": "자재표", "hasHeaders": true,
              "styleName": "TableStyleMedium2", "showTotals": true }
            """);
        var e7 = Op(
            """
            { "op": "create_pivot", "target": { "sheet": "피벗" },
              "sourceRange": "'원장'!A1:C5", "destination": "A1", "name": "LedgerPivot",
              "rows": ["품목"], "columns": ["월"],
              "values": [{ "field": "금액", "function": "sum" }] }
            """);

        AssertPublishedOk(items, e4);
        AssertPublishedOk(items, e7);
        Assert.Empty(ExcelDataOperationsContract.ValidatePublicInput(e4));
        Assert.Empty(ExcelDataOperationsContract.ValidatePublicInput(e7));

        var staleDestSheet = Op(
            """
            { "op": "create_pivot", "target": { "sheet": "피벗" },
              "sourceRange": "'원장'!A1:C5", "destSheet": "피벗", "name": "LedgerPivot" }
            """);
        AssertPublishedFails(items, staleDestSheet, "destSheet");
    }

    [Fact]
    public void Published_anyOf_rejects_nested_field_conflicts()
    {
        var items = PublishedItems();
        AssertPublishedFails(items, Op(
            """
            { "op": "create_pivot", "target": { "sheet": "원가" },
              "sourceRange": "A1:G20", "destination": "A24", "name": "CostPivot",
              "values": [{ "field": "수량", "function": null }] }
            """), "function");
        AssertPublishedFails(items, Op(
            """
            { "op": "create_pivot", "target": { "sheet": "원가" },
              "sourceRange": "A1:G20", "destination": "A24", "name": "CostPivot",
              "rows": ["품목", ""] }
            """), "rows");
        AssertPublishedFails(items, Op(
            """
            { "op": "create_pivot", "target": { "sheet": "원가" },
              "sourceRange": "A1:G20", "destination": "A24", "name": "CostPivot",
              "values": [[1, 2]] }
            """), "values");
        AssertPublishedFails(items, Op(
            """
            { "op": "create_chart", "target": { "sheet": "월간보고" },
              "sourceRange": "A1:B6", "chartType": "combo",
              "series": [{ "values": "B2:B7" }] }
            """), "chartType");
        AssertPublishedFails(items, Op(
            """
            { "op": "create_chart", "target": { "sheet": "월간보고" },
              "sourceRange": "A1:C7", "chartType": "columnClustered",
              "series": [{ "values": "B2:B7", "chartType": "banana" }] }
            """), "chartType");
        AssertPublishedFails(items, Op(
            """
            { "op": "create_chart", "target": { "sheet": "월간보고" },
              "sourceRange": "A1:C7", "chartType": "columnClustered",
              "series": [{ "values": "C2:C7", "axisGroup": "tertiary" }] }
            """), "axisGroup");
        AssertPublishedFails(items, Op(
            """
            { "op": "create_chart", "target": { "sheet": "월간보고" },
              "sourceRange": "A1:C7", "chartType": "columnClustered",
              "series": [{ "index": 0, "values": "B2:B7" }] }
            """), "index");
        AssertPublishedFails(items, Op(
            """
            { "op": "create_chart", "target": { "sheet": "월간보고" },
              "sourceRange": "A1:B6", "chartType": "line",
              "series": [{ "marker": "star" }] }
            """), "marker");
        AssertPublishedFails(items, Op(
            """
            { "op": "create_chart", "target": { "sheet": "월간보고" },
              "sourceRange": "A1:B6", "chartType": "line",
              "series": [{ "lineWeight": 0 }] }
            """), "lineWeight");
        AssertPublishedFails(items, Op(
            """
            { "op": "create_chart", "target": { "sheet": "월간보고" },
              "sourceRange": "A1:B6", "chartType": "line",
              "position": "A1" }
            """), "position");
        AssertPublishedFails(items, Op(
            """
            { "op": "remove_duplicates", "target": { "sheet": "자재대장" },
              "range": "A1:G8", "columns": ["구분"] }
            """), "columns");
        AssertPublishedFails(items, Op(
            """
            { "op": "set_table_totals", "target": { "sheet": "자재대장" },
              "name": "자재표", "showTotals": true,
              "columns": [{ "col": "A", "width": 12 }] }
            """), "columns");
        AssertPublishedFails(items, Op(
            """
            { "op": "add_conditional_format", "target": { "sheet": "자재대장" },
              "range": "G2:G20", "rule": { "type": "expression", "formula": "=G2<10" } }
            """), "formula");
    }

    [Fact]
    public void Update_chart_allows_width_height_only_position()
    {
        var items = PublishedItems();
        var update = Op(
            """
            { "op": "update_chart", "target": { "sheet": "예정공정표" },
              "name": "Chart1", "position": { "width": 440, "height": 200 } }
            """);
        AssertPublishedOk(items, update);
    }

    [Fact]
    public void Outer_position_string_conjunct_with_chart_object_rejects()
    {
        var items = PublishedItems();
        var chart = Op(
            """
            { "op": "create_chart", "target": { "sheet": "예정공정표" },
              "sourceRange": "'예정공정표'!$G$95:$BP$95", "chartType": "line",
              "name": "Chart1",
              "position": { "left": 12, "top": 18, "width": 420, "height": 180 } }
            """);
        AssertPublishedOk(items, chart);

        items["properties"] = new JsonObject
        {
            ["position"] = new JsonObject
            {
                ["type"] = "string",
                ["enum"] = new JsonArray("before", "after", "first", "last"),
            },
        };
        var errors = new List<string>();
        Assert.False(ExcelDataOperationSchema.TryValidatePublishedItems(items, chart, errors));
        Assert.Contains(errors, error => error.Contains("position", StringComparison.OrdinalIgnoreCase));
        Assert.True(ExcelDataOperationSchema.CountPublishedMatches(items, chart) >= 1,
            "anyOf chart branch still matches; outer leftover position:string must be what rejects");
    }

    [Fact]
    public void Leftover_non_data_branch_still_accepts_move_sheet_position_string()
    {
        var items = PublishedItems();
        var move = Op("""{ "op": "move_sheet", "position": "before" }""");
        Assert.True(ExcelDataOperationSchema.CountPublishedMatches(items, move) >= 1);
        var chartAsMove = Op(
            """
            { "op": "create_chart", "target": { "sheet": "월간보고" },
              "sourceRange": "A1:B6", "chartType": "line", "position": "before" }
            """);
        AssertPublishedFails(items, chartAsMove, "position");
    }

    [Fact]
    public void Every_write_op_has_a_real_payload_that_matches_published_anyOf()
    {
        var items = PublishedItems();
        var examples = RealPayloads();
        Assert.Equal(ExcelDataOperationsContract.WriteOpNames.Count, examples.Count);
        foreach (var name in ExcelDataOperationsContract.WriteOpNames)
        {
            Assert.True(examples.ContainsKey(name), name);
            var errors = new List<string>();
            Assert.True(
                ExcelDataOperationSchema.TryValidatePublishedItems(items, examples[name], errors),
                name + ": " + string.Join("; ", errors));
            Assert.Equal(1, ExcelDataOperationSchema.CountPublishedMatches(items, examples[name]));
        }
    }

    [Fact]
    public void Required_and_optional_fields_are_declared_on_each_branch()
    {
        foreach (var name in ExcelDataOperationsContract.WriteOpNames)
        {
            var properties = ExcelDataOperationSchema.OpSchema(name)["properties"]!.AsObject();
            Assert.Equal(name, properties["op"]!["const"]!.GetValue<string>());
            foreach (var (field, _) in ExcelDataOperationsContract.RequiredFields[name])
                Assert.True(properties.ContainsKey(field), name + "." + field);
            foreach (var field in ExcelDataOperationsContract.OptionalFields[name])
                Assert.True(properties.ContainsKey(field), name + "." + field);
        }
    }

    private static void AssertPublishedOk(JsonObject items, JsonObject instance)
    {
        var errors = new List<string>();
        Assert.True(ExcelDataOperationSchema.TryValidatePublishedItems(items, instance, errors),
            string.Join("; ", errors));
        Assert.Equal(1, ExcelDataOperationSchema.CountPublishedMatches(items, instance));
    }

    private static void AssertPublishedFails(JsonObject items, JsonObject instance, string token)
    {
        var errors = new List<string>();
        Assert.False(ExcelDataOperationSchema.TryValidatePublishedItems(items, instance, errors));
        Assert.Contains(errors, error => error.Contains(token, StringComparison.OrdinalIgnoreCase));
        Assert.Equal(0, ExcelDataOperationSchema.CountPublishedMatches(items, instance));
    }

    [Fact]
    public void Published_shape_branches_expose_practical_formatting_and_reject_invalid_schema_values()
    {
        var items = PublishedItems();
        var valid = Op("""
            { "op": "insert_textbox", "target": { "sheet": "월간보고" }, "name": "ReportNote", "text": "",
              "position": { "left": 1, "top": 2, "width": 120, "height": 30 },
              "lineColor": "#112233", "lineWeight": 0.25, "lineVisible": true, "rotation": 360,
              "font": { "name": "Arial", "size": 11, "bold": true, "italic": false, "color": "#445566" } }
            """);
        var errors = new List<string>();
        Assert.True(ExcelDataOperationSchema.TryValidatePublishedItems(items, valid, errors), string.Join("; ", errors));
        Assert.Empty(errors);

        var invalid = Op("""
            { "op": "update_shape", "target": { "sheet": "월간보고" }, "name": "HighlightBox",
              "lineWeight": 0, "font": { "bold": "yes", "unsupported": true } }
            """);
        AssertPublishedFails(items, invalid, "lineWeight");
    }

    private static JsonObject FindBranch(JsonObject items, string opName) =>
        items["anyOf"]!.AsArray().OfType<JsonObject>().First(branch =>
            branch["properties"]?["op"]?["const"]?.GetValue<string>() == opName);

    private static JsonObject PublishedItems()
    {
        var items = FlattenedItemsLikeToolRegistry();
        ExcelDataOperationSchema.AttachToApplyItems(items);
        return items;
    }

    private static JsonObject FlattenedItemsLikeToolRegistry() => new()
    {
        ["type"] = "object",
        ["required"] = new JsonArray("op"),
        ["properties"] = new JsonObject
        {
            ["op"] = new JsonObject
            {
                ["type"] = "string",
                ["enum"] = new JsonArray(
                    "set_values", "move_sheet", "set_row_heights", "set_column_widths",
                    "create_table", "create_chart", "create_pivot", "remove_duplicates",
                    "set_table_totals"),
            },
            ["position"] = new JsonObject
            {
                ["type"] = "string",
                ["enum"] = new JsonArray("before", "after", "first", "last"),
                ["description"] = "move_sheet position",
            },
            ["rows"] = new JsonObject
            {
                ["type"] = "array",
                ["items"] = new JsonObject { ["type"] = "object" },
            },
            ["columns"] = new JsonObject
            {
                ["type"] = "array",
                ["items"] = new JsonObject { ["type"] = "object" },
            },
            ["values"] = new JsonObject
            {
                ["type"] = "array",
                ["items"] = new JsonObject { ["type"] = "array" },
            },
        },
    };

    private static Dictionary<string, JsonObject> RealPayloads() => new(StringComparer.OrdinalIgnoreCase)
    {
        ["create_table"] = Op(
            """{ "op": "create_table", "target": { "sheet": "자재대장" }, "range": "A2:I5", "name": "자재표", "hasHeaders": true, "styleName": "TableStyleMedium2", "showTotals": true }"""),
        ["resize_table"] = Op(
            """{ "op": "resize_table", "target": { "sheet": "자재대장" }, "name": "자재표", "range": "A2:I9" }"""),
        ["style_table"] = Op(
            """{ "op": "style_table", "target": { "sheet": "자재대장" }, "name": "자재표", "showRowStripes": true, "showAutoFilter": true }"""),
        ["set_table_totals"] = Op(
            """{ "op": "set_table_totals", "target": { "sheet": "자재대장" }, "name": "자재표", "showTotals": true, "columns": [{ "column": "수량", "function": "sum" }] }"""),
        ["add_table_column"] = Op(
            """{ "op": "add_table_column", "target": { "sheet": "자재대장" }, "name": "자재표", "columnName": "검수", "formula": "=[@수량]" }"""),
        ["delete_table"] = Op(
            """{ "op": "delete_table", "target": { "sheet": "자재대장" }, "name": "자재표" }"""),
        ["sort_range"] = Op(
            """{ "op": "sort_range", "target": { "sheet": "자재대장" }, "range": "A2:I5", "hasHeaders": true, "keys": [{ "column": "일자", "order": "asc" }] }"""),
        ["sort_table"] = Op(
            """{ "op": "sort_table", "target": { "sheet": "자재대장" }, "name": "자재표", "keys": [{ "column": "일자", "order": "asc" }] }"""),
        ["set_auto_filter"] = Op(
            """{ "op": "set_auto_filter", "target": { "sheet": "자재대장" }, "name": "자재표", "criteria": [{ "column": "구분", "operator": "equals", "value": "입고" }] }"""),
        ["clear_auto_filter"] = Op(
            """{ "op": "clear_auto_filter", "target": { "sheet": "자재대장" }, "name": "자재표" }"""),
        ["remove_duplicates"] = Op(
            """{ "op": "remove_duplicates", "target": { "sheet": "자재대장" }, "range": "A1:G8", "hasHeaders": true, "columns": [4] }"""),
        ["text_to_columns"] = Op(
            """{ "op": "text_to_columns", "target": { "sheet": "자재대장" }, "range": "B2:B8", "dataType": "delimited", "other": true, "otherChar": "-" }"""),
        ["define_name"] = Op(
            """{ "op": "define_name", "name": "StockOnHand", "scope": "sheet", "target": { "sheet": "자재대장" }, "refersTo": "='자재대장'!$G$2:$G$8" }"""),
        ["update_name"] = Op(
            """{ "op": "update_name", "name": "StockOnHand", "scope": "sheet", "target": { "sheet": "자재대장" }, "refersTo": "='자재대장'!$G$2:$G$9" }"""),
        ["delete_name"] = Op(
            """{ "op": "delete_name", "name": "StockOnHand", "scope": "sheet", "target": { "sheet": "자재대장" } }"""),
        ["set_data_validation"] = Op(
            """{ "op": "set_data_validation", "target": { "sheet": "자재대장" }, "range": "C2:C8", "type": "list", "source": "입고,출고,이동", "inCellDropdown": true, "errorStyle": "stop" }"""),
        ["clear_data_validation"] = Op(
            """{ "op": "clear_data_validation", "target": { "sheet": "자재대장" }, "range": "C2:C8" }"""),
        ["add_conditional_format"] = Op(
            """{ "op": "add_conditional_format", "target": { "sheet": "자재대장" }, "range": "G2:G8", "rule": { "type": "cellValue", "operator": "less", "formula1": "10" }, "style": { "bold": true, "fontColor": "#9C0006", "fillColor": "#FFC7CE" } }"""),
        ["clear_conditional_formats"] = Op(
            """{ "op": "clear_conditional_formats", "target": { "sheet": "자재대장" }, "range": "G2:G8" }"""),
        ["create_chart"] = Op(
            """{ "op": "create_chart", "target": { "sheet": "예정공정표" }, "sourceRange": "'예정공정표'!$G$95:$BP$95", "chartType": "line", "name": "Chart1", "title": "", "hasLegend": false, "plotBy": "rows", "position": { "left": 12, "top": 18, "width": 420, "height": 180 }, "series": [{ "lineColor": "#FF0000", "lineWeight": 2.25, "marker": "none" }], "chartFill": "none", "plotFill": "none", "axes": { "value": { "visible": false, "maximum": 100 } }, "plotArea": { "spanChart": true } }"""),
        ["update_chart"] = Op(
            """{ "op": "update_chart", "target": { "sheet": "월간보고" }, "name": "MonthlyOutput", "title": "2026년 월별 실적", "legendPosition": "right" }"""),
        ["delete_chart"] = Op(
            """{ "op": "delete_chart", "target": { "sheet": "월간보고" }, "name": "MonthlyOutput" }"""),
        ["insert_sheet_picture"] = Op(
            """{ "op": "insert_sheet_picture", "target": { "sheet": "월간보고" }, "path": "C:\\DocBridgeTest\\Desktop\\plug-in\\logo.png", "name": "ReportLogo", "lockAspectRatio": true, "position": { "left": 420, "top": 12, "width": 120, "height": 48 } }"""),
        ["update_picture"] = Op(
            """{ "op": "update_picture", "target": { "sheet": "월간보고" }, "name": "ReportLogo", "position": { "left": 430, "top": 16, "width": 110, "height": 44 } }"""),
        ["delete_picture"] = Op(
            """{ "op": "delete_picture", "target": { "sheet": "월간보고" }, "name": "ReportLogo" }"""),
        ["set_cell_note"] = Op(
            """{ "op": "set_cell_note", "target": { "sheet": "자재대장" }, "cell": "B4", "text": "현장 검수 후 입고 확정", "visible": false }"""),
        ["clear_cell_note"] = Op(
            """{ "op": "clear_cell_note", "target": { "sheet": "자재대장" }, "cell": "B4" }"""),
        ["set_hyperlink"] = Op(
            """{ "op": "set_hyperlink", "target": { "sheet": "자재대장" }, "cell": "D2", "address": "https://example.invalid/spec", "textToDisplay": "관부속" }"""),
        ["clear_hyperlink"] = Op(
            """{ "op": "clear_hyperlink", "target": { "sheet": "자재대장" }, "cell": "D2" }"""),
        ["create_pivot"] = Op(
            """{ "op": "create_pivot", "target": { "sheet": "피벗" }, "sourceRange": "'원장'!A1:C5", "destination": "A1", "name": "LedgerPivot", "rows": ["품목"], "columns": ["월"], "values": [{ "field": "금액", "function": "sum" }] }"""),
        ["update_pivot"] = Op(
            """{ "op": "update_pivot", "target": { "sheet": "피벗" }, "name": "LedgerPivot", "rows": ["품목"], "values": [{ "field": "금액" }] }"""),
        ["refresh_pivot"] = Op(
            """{ "op": "refresh_pivot", "target": { "sheet": "피벗" }, "name": "LedgerPivot" }"""),
        ["delete_pivot"] = Op(
            """{ "op": "delete_pivot", "target": { "sheet": "피벗" }, "name": "LedgerPivot" }"""),
        ["insert_shape"] = Op(
            """{ "op": "insert_shape", "target": { "sheet": "월간보고" }, "shapeType": "rectangle", "name": "HighlightBox", "position": { "left": 20, "top": 12, "width": 80, "height": 24 } }"""),
        ["update_shape"] = Op(
            """{ "op": "update_shape", "target": { "sheet": "월간보고" }, "name": "HighlightBox", "text": "확인" }"""),
        ["delete_shape"] = Op(
            """{ "op": "delete_shape", "target": { "sheet": "월간보고" }, "name": "HighlightBox" }"""),
        ["insert_textbox"] = Op(
            """{ "op": "insert_textbox", "target": { "sheet": "월간보고" }, "name": "ReportNote", "text": "피벗 새로고침 후 확인", "position": { "left": 420, "top": 80, "width": 160, "height": 36 } }"""),
        ["update_textbox"] = Op(
            """{ "op": "update_textbox", "target": { "sheet": "월간보고" }, "name": "ReportNote", "text": "확인 완료" }"""),
        ["delete_textbox"] = Op(
            """{ "op": "delete_textbox", "target": { "sheet": "월간보고" }, "name": "ReportNote" }"""),
        ["set_rich_text"] = Op("""{ "op":"set_rich_text", "target":{"sheet":"자재대장"}, "cell":"A2", "baseFont":{"name":"Aptos","size":11,"bold":false,"italic":false,"underline":false,"color":"#000000"}, "runs":[{"text":"text"}] }"""),
        ["update_conditional_format"] = Op("""{ "op":"update_conditional_format", "target":{"sheet":"자재대장"}, "range":"G2:G8", "index":1, "expectedFingerprint":"f", "rule":{"type":"expression","formula1":"=G2<10"} }"""),
        ["delete_conditional_format"] = Op("""{ "op":"delete_conditional_format", "target":{"sheet":"자재대장"}, "range":"G2:G8", "index":1, "expectedFingerprint":"f" }"""),
        ["append_table_rows"] = Op("""{ "op":"append_table_rows", "target":{"sheet":"자재대장"}, "name":"자재표", "rows":[[1]] }"""),
        ["insert_table_rows"] = Op("""{ "op":"insert_table_rows", "target":{"sheet":"자재대장"}, "name":"자재표", "index":1, "rows":[[1]] }"""),
        ["delete_table_rows"] = Op("""{ "op":"delete_table_rows", "target":{"sheet":"자재대장"}, "name":"자재표", "index":1, "count":1 }"""),
        ["create_connector"] = Op(
            """{ "op": "create_connector", "target": { "sheet": "월간보고" }, "connectorType": "straight", "begin": { "x": 1, "y": 2 }, "end": { "x": 20, "y": 30 } }"""),
        ["update_connector"] = Op(
            """{ "op": "update_connector", "target": { "sheet": "월간보고" }, "name": "Connector 1", "begin": { "x": 1, "y": 2 } }"""),
        ["update_external_links"] = Op(
            """{ "op": "update_external_links", "sourceContains": "적용수량" }"""),
        ["change_link_source"] = Op(
            """{ "op": "change_link_source", "source": "적용수량집계_r2.xlsx", "newSource": "C:\\작업\\적용수량집계_r3.xlsx" }"""),
        ["break_external_link"] = Op(
            """{ "op": "break_external_link", "source": "적용수량집계_r2.xlsx" }"""),
        ["set_calculation_mode"] = Op(
            """{ "op": "set_calculation_mode", "mode": "manual", "recalculate": "fullRebuild" }"""),
        ["freeze_values"] = Op(
            """{ "op": "freeze_values", "target": { "sheet": "자재대장" }, "range": "A1:G8" }"""),
        ["paste_special"] = Op(
            """{ "op": "paste_special", "target": { "sheet": "자재대장" }, "sourceRange": "A1:G8", "destination": "A10", "paste": "values", "skipBlanks": true, "operation": "none" }"""),
        ["goal_seek"] = Op(
            """{ "op": "goal_seek", "target": { "sheet": "원가" }, "cell": "B1", "goalCell": "B3", "goal": 11000 }"""),
        ["protect_workbook"] = Op(
            """{ "op": "protect_workbook", "structure": true, "windows": false }"""),
        ["unprotect_workbook"] = Op(
            """{ "op": "unprotect_workbook" }"""),
        ["set_split_panes"] = Op(
            """{ "op": "set_split_panes", "target": { "sheet": "자재대장" }, "cell": "C5" }"""),
        ["create_sparkline"] = Op(
            """{ "op": "create_sparkline", "target": { "sheet": "월간보고" }, "location": "C2", "sourceData": "B2:B6", "type": "line", "markers": true }"""),
        ["update_sparkline"] = Op(
            """{ "op": "update_sparkline", "target": { "sheet": "월간보고" }, "location": "C2", "showHigh": true }"""),
        ["delete_sparkline"] = Op(
            """{ "op": "delete_sparkline", "target": { "sheet": "월간보고" }, "location": "C2" }"""),
        ["create_slicer"] = Op(
            """{ "op": "create_slicer", "target": { "sheet": "월간보고" }, "source": "Materials", "field": "구분", "name": "Slicer_구분" }"""),
        ["delete_slicer"] = Op(
            """{ "op": "delete_slicer", "target": { "sheet": "월간보고" }, "name": "Slicer_구분" }"""),
        ["apply_cell_style"] = Op(
            """{ "op": "apply_cell_style", "target": { "sheet": "자재대장" }, "range": "A1:G1", "styleName": "Heading 1" }"""),
    };

    private static JsonObject Op(string json) =>
        Json.ParseObject(json) ?? throw new InvalidOperationException(json);
}
