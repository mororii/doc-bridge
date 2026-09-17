using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelDataOperationsContractTests
{
    [Fact]
    public void Write_op_catalog_covers_data_and_reporting_ops()
    {
        Assert.True(ExcelDataOperationsContract.WriteOpNames.Count >= 39);
        Assert.Equal(ExcelDataOperationsContract.WriteOpNames.Count, ExcelDataOperationsContract.WriteOpSet.Count);
        foreach (var name in ExcelDataOperationsContract.WriteOpNames)
            Assert.True(ExcelDataOperationsContract.IsDataOperation(name), name);
        Assert.Contains("create_pivot", ExcelDataOperationsContract.WriteOpNames);
        Assert.Contains("remove_duplicates", ExcelDataOperationsContract.WriteOpNames);
        Assert.Contains("insert_textbox", ExcelDataOperationsContract.WriteOpNames);
        Assert.False(ExcelDataOperationsContract.IsDataOperation("set_values"));
        Assert.False(ExcelDataOperationsContract.IsDataOperation("insert_picture"));
        Assert.Equal("data-objects", ExcelDataOperationsContract.RestoreMode);
        Assert.Equal(2, ExcelDataOperationsContract.SnapshotVersion);
    }

    [Fact]
    public void Sheet_scoped_writes_reject_active_sheet_assumption()
    {
        var errors = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "create_table", """{ "op": "create_table", "range": "A1:D8" }"""));
        Assert.Contains(errors, error => error.Contains("target.sheet") && error.Contains("active sheet"));
    }

    [Fact]
    public void Sheet_qualified_range_is_accepted_without_target_object()
    {
        var errors = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "create_table", """{ "op": "create_table", "range": "'자재대장'!A1:G12" }"""));
        Assert.Empty(errors);
    }

    [Fact]
    public void Conflicting_sheet_and_range_are_rejected()
    {
        var errors = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "sort_range",
            """
            { "op": "sort_range", "target": { "sheet": "자재대장" }, "range": "원가!A1:C4",
              "keys": [{ "column": "A", "order": "asc" }] }
            """));
        Assert.Contains(errors, error => error.Contains("does not match"));
    }

    [Theory]
    [InlineData("A1:XFD1048576", true)]
    [InlineData("A1:B2", true)]
    [InlineData("$G$2:$G$20", true)]
    [InlineData("A1,C3", false)]
    [InlineData("R1C1", false)]
    [InlineData("AAAA1", false)]
    public void Rectangular_a1_addresses_are_validated(string address, bool ok)
    {
        Assert.Equal(ok, ExcelDataOperationsContract.TryParseA1(address, out _));
    }

    [Theory]
    [InlineData("Materials", true)]
    [InlineData("_창고", true)]
    [InlineData("A1", false)]
    [InlineData("C", false)]
    [InlineData("자재 대장", false)]
    [InlineData("", false)]
    public void Defined_names_follow_excel_rules(string name, bool ok) =>
        Assert.Equal(ok, ExcelDataOperationsContract.IsValidDefinedName(name));

    [Fact]
    public void List_validation_requires_source_or_formula()
    {
        var errors = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "set_data_validation",
            """
            { "op": "set_data_validation", "target": { "sheet": "자재대장" },
              "range": "C2:C20", "type": "list" }
            """));
        Assert.Contains(errors, error => error.Contains("source or formula1"));
    }

    [Fact]
    public void Chart_type_and_picture_path_are_strict()
    {
        var chart = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "create_chart",
            """
            { "op": "create_chart", "target": { "sheet": "월간보고" },
              "sourceRange": "A1:B6", "chartType": "pivot" }
            """));
        Assert.Contains(chart, error => error.Contains("chartType"));

        var relative = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "insert_sheet_picture",
            """
            { "op": "insert_sheet_picture", "target": { "sheet": "월간보고" },
              "path": "logo.png" }
            """));
        Assert.Contains(relative, error => error.Contains("absolute"));

        var ext = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "insert_sheet_picture",
            """
            { "op": "insert_sheet_picture", "target": { "sheet": "월간보고" },
              "path": "C:\\\\temp\\\\logo.exe" }
            """));
        Assert.Contains(ext, error => error.Contains("extension"));
    }

    [Fact]
    public void Filter_requires_exactly_one_of_range_or_name()
    {
        var both = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "set_auto_filter",
            """
            { "op": "set_auto_filter", "target": { "sheet": "자재대장" },
              "range": "A1:G12", "name": "Materials",
              "criteria": [{ "column": "구분", "operator": "equals", "value": "입고" }] }
            """));
        Assert.Contains(both, error => error.Contains("exactly one"));

        var neither = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "clear_auto_filter",
            """{ "op": "clear_auto_filter", "target": { "sheet": "자재대장" } }"""));
        Assert.Contains(neither, error => error.Contains("exactly one"));
    }

    [Fact]
    public void Workbook_scoped_names_do_not_require_a_sheet()
    {
        var errors = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "define_name",
            """{ "op": "define_name", "name": "VatRate", "refersTo": "=0.1", "scope": "workbook" }"""));
        Assert.Empty(errors);
    }

    [Fact]
    public void Sheet_scoped_names_require_target_sheet()
    {
        var errors = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "define_name",
            """{ "op": "define_name", "name": "Warehouse", "refersTo": "=A1", "scope": "sheet" }"""));
        Assert.Contains(errors, error => error.Contains("target.sheet"));
    }

    [Fact]
    public void Data_ops_cannot_mix_with_value_or_merge_batches()
    {
        var errors = new List<string>();
        ExcelDataOperationsContract.ValidateDataBatchMixing(
            new[]
            {
                Op("create_table", """{ "op": "create_table", "target": { "sheet": "자재대장" }, "range": "A1:B2" }"""),
                Op("set_values", """{ "op": "set_values", "target": { "sheet": "자재대장" }, "range": "A1", "values": [[1]] }"""),
            },
            errors);
        Assert.Contains(errors, error => error.Contains("cannot be mixed"));
    }

    [Fact]
    public void Note_and_hyperlink_writes_are_single_cell()
    {
        var note = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "set_cell_note",
            """
            { "op": "set_cell_note", "target": { "sheet": "자재대장" },
              "range": "A1:A2", "text": "검수" }
            """));
        Assert.Contains(note, error => error.Contains("single cell"));

        var link = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "set_hyperlink",
            """{ "op": "set_hyperlink", "target": { "sheet": "자재대장" }, "cell": "A1" }"""));
        Assert.Contains(link, error => error.Contains("address and/or subAddress"));
    }

    [Fact]
    public void Catalog_maps_official_excel_constants()
    {
        Assert.True(ExcelDataObjectCatalog.TryChartType("columnClustered", out var chart));
        Assert.Equal(51, chart);
        Assert.True(ExcelDataObjectCatalog.TryValidationType("list", out var list));
        Assert.Equal(3, list);
        Assert.True(ExcelDataObjectCatalog.TryTotalsFunction("sum", out var sum));
        Assert.Equal(6, sum);
        Assert.Equal(ExcelDataObjectCatalog.XlDescending, ExcelDataObjectCatalog.SortOrder("desc"));
    }

    [Fact]
    public void Required_and_optional_field_tables_cover_every_write_op()
    {
        foreach (var op in ExcelDataOperationsContract.WriteOpNames)
        {
            Assert.True(ExcelDataOperationsContract.RequiredFields.ContainsKey(op), op);
            Assert.True(ExcelDataOperationsContract.OptionalFields.ContainsKey(op), op);
        }
    }

    [Fact]
    public void Valid_materials_and_chart_ops_pass_public_validation()
    {
        foreach (var op in ExcelDataOperationsHostScenarios.AllOps())
        {
            var errors = ExcelDataOperationsContract.ValidatePublicInput(op);
            Assert.True(errors.Count == 0, $"{Json.GetString(op, "op")}: {string.Join("; ", errors)}");
        }
    }

    [Fact]
    public void Unsupported_conditional_rule_is_rejected_not_stubbed()
    {
        var errors = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "add_conditional_format",
            """
            { "op": "add_conditional_format", "target": { "sheet": "자재대장" },
              "range": "G2:G20", "rule": { "type": "top10" } }
            """));
        Assert.Contains(errors, error => error.Contains("rule.type"));
    }

    [Fact]
    public void Icon_set_and_pivot_ops_are_accepted()
    {
        var icon = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "add_conditional_format",
            """
            { "op": "add_conditional_format", "target": { "sheet": "자재대장" },
              "range": "G2:G20", "rule": { "type": "iconSet", "iconSet": "arrows3" } }
            """));
        Assert.Empty(icon);

        var pivot = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "create_pivot",
            """
            { "op": "create_pivot", "target": { "sheet": "원가" },
              "sourceRange": "A1:G20", "destination": "A24", "name": "CostPivot",
              "rows": ["품목"], "values": [{ "field": "수량", "function": "sum" }] }
            """));
        Assert.Empty(pivot);
    }

    [Fact]
    public void Inspection_objectKind_normalizes_to_a_read_scope()
    {
        var args = Json.ParseObject("""{ "scope": "objects", "objectKind": "tables" }""")!;
        Assert.Equal("tables", ExcelDataOperationsContract.ResolveReadScope(args));
        Assert.Equal("all", ExcelDataOperationsContract.ResolveReadScope(
            Json.ParseObject("""{ "scope": "objects" }""")!));
    }

    [Fact]
    public void DescribeSchema_exposes_field_types_not_just_op_names()
    {
        var schema = ExcelDataOperationsContract.DescribeSchema();
        Assert.Equal(2, Json.GetInt(schema, "snapshotVersion"));
        Assert.Equal("Unlist", Json.GetString(schema, "tableCreateRollback"));
        Assert.True(Json.GetBool(schema, "equalityRequired"));
        Assert.NotNull(Json.GetObj(schema, "fieldTypes"));
        Assert.NotNull(Json.GetArr(schema, "writeOps"));
        Assert.True((Json.GetArr(schema, "writeOps")?.Count ?? 0) >= 39);
    }

    [Fact]
    public void Pivot_function_omitted_defaults_but_explicit_null_is_rejected()
    {
        var omitted = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "create_pivot",
            """
            { "op": "create_pivot", "target": { "sheet": "원가" },
              "sourceRange": "A1:G20", "destination": "A24", "name": "CostPivot",
              "values": [{ "field": "수량" }] }
            """));
        Assert.Empty(omitted);

        var explicitNull = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "create_pivot",
            """
            { "op": "create_pivot", "target": { "sheet": "원가" },
              "sourceRange": "A1:G20", "destination": "A24", "name": "CostPivot",
              "values": [{ "field": "수량", "function": null }] }
            """));
        Assert.Contains(explicitNull, error => error.Contains("function") && error.Contains("null"));

        var unknown = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "create_pivot",
            """
            { "op": "create_pivot", "target": { "sheet": "원가" },
              "sourceRange": "A1:G20", "destination": "A24", "name": "CostPivot",
              "values": [{ "field": "수량", "function": "median" }] }
            """));
        Assert.Contains(unknown, error => error.Contains("function"));
    }

    [Fact]
    public void Add_table_column_insertAfter_must_be_finite_positive_int()
    {
        var ok = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "add_table_column",
            """
            { "op": "add_table_column", "target": { "sheet": "자재대장" },
              "name": "Materials", "columnName": "검수", "insertAfter": 2 }
            """));
        Assert.Empty(ok);

        foreach (var json in new[]
                 {
                     """{ "op": "add_table_column", "target": { "sheet": "자재대장" }, "name": "Materials", "columnName": "검수", "insertAfter": 0 }""",
                     """{ "op": "add_table_column", "target": { "sheet": "자재대장" }, "name": "Materials", "columnName": "검수", "insertAfter": -1 }""",
                     """{ "op": "add_table_column", "target": { "sheet": "자재대장" }, "name": "Materials", "columnName": "검수", "insertAfter": 1.5 }""",
                     """{ "op": "add_table_column", "target": { "sheet": "자재대장" }, "name": "Materials", "columnName": "검수", "insertAfter": null }""",
                 })
        {
            var errors = ExcelDataOperationsContract.ValidatePublicInput(Op("add_table_column", json));
            Assert.Contains(errors, error => error.Contains("insertAfter"));
        }
    }

    [Fact]
    public void Text_to_columns_other_requires_otherChar_and_rejects_unknown_dataType()
    {
        var missing = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "text_to_columns",
            """
            { "op": "text_to_columns", "target": { "sheet": "자재대장" },
              "range": "B2:B8", "other": true }
            """));
        Assert.Contains(missing, error => error.Contains("otherChar"));

        var dataTypeNull = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "text_to_columns",
            """
            { "op": "text_to_columns", "target": { "sheet": "자재대장" },
              "range": "B2:B8", "dataType": null }
            """));
        Assert.Contains(dataTypeNull, error => error.Contains("dataType"));

        var unknown = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "text_to_columns",
            """
            { "op": "text_to_columns", "target": { "sheet": "자재대장" },
              "range": "B2:B8", "dataType": "csv" }
            """));
        Assert.Contains(unknown, error => error.Contains("dataType"));

        var fixedWidth = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "text_to_columns",
            """
            { "op": "text_to_columns", "target": { "sheet": "자재대장" },
              "range": "B2:B8", "dataType": "fixed" }
            """));
        Assert.Contains(fixedWidth, error => error.Contains("FieldInfo") || error.Contains("fixed"));
    }

    [Fact]
    public void Shape_and_icon_tokens_reject_null_and_unknown()
    {
        var unknownShape = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "insert_shape",
            """
            { "op": "insert_shape", "target": { "sheet": "월간보고" },
              "shapeType": "triangle",
              "position": { "left": 1, "top": 1, "width": 20, "height": 20 } }
            """));
        Assert.Contains(unknownShape, error => error.Contains("shapeType"));

        var nullShape = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "insert_shape",
            """
            { "op": "insert_shape", "target": { "sheet": "월간보고" },
              "shapeType": null,
              "position": { "left": 1, "top": 1, "width": 20, "height": 20 } }
            """));
        Assert.Contains(nullShape, error => error.Contains("shapeType"));

        var badFill = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "insert_shape",
            """
            { "op": "insert_shape", "target": { "sheet": "월간보고" },
              "shapeType": "rectangle", "fillColor": "not-a-color",
              "position": { "left": 1, "top": 1, "width": 20, "height": 20 } }
            """));
        Assert.Contains(badFill, error => error.Contains("fillColor"));

        var icon = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "add_conditional_format",
            """
            { "op": "add_conditional_format", "target": { "sheet": "자재대장" },
              "range": "G2:G20", "rule": { "type": "iconSet", "iconSet": null } }
            """));
        Assert.Contains(icon, error => error.Contains("iconSet"));
    }

    [Fact]
    public void Line_chart_presentation_contract_accepts_e1_overlay()
    {
        var errors = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "create_chart",
            """
            { "op": "create_chart", "target": { "sheet": "예정공정표" },
              "sourceRange": "'예정공정표'!$G$95:$BP$95", "chartType": "line",
              "name": "Chart1", "title": "", "hasLegend": false, "plotBy": "rows",
              "series": [{ "lineColor": "#FF0000", "lineWeight": 2.25, "marker": "none" }],
              "chartFill": "none", "plotFill": "none", "chartBorder": "none", "plotBorder": "none",
              "axes": { "category": { "visible": false }, "value": { "visible": false, "maximum": 100 } },
              "plotArea": { "spanChart": true } }
            """));
        Assert.Empty(errors);
    }

    [Fact]
    public void Chart_presentation_rejects_unknown_marker_and_weight()
    {
        var marker = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "create_chart",
            """
            { "op": "create_chart", "target": { "sheet": "월간보고" },
              "sourceRange": "A1:B6", "chartType": "line",
              "series": [{ "marker": "star" }] }
            """));
        Assert.Contains(marker, error => error.Contains("marker"));

        var weight = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "create_chart",
            """
            { "op": "create_chart", "target": { "sheet": "월간보고" },
              "sourceRange": "A1:B6", "chartType": "line",
              "series": [{ "lineWeight": 0 }] }
            """));
        Assert.Contains(weight, error => error.Contains("lineWeight"));
    }

    [Fact]
    public void Pivot_string_arrays_reject_empty_items()
    {
        var errors = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "create_pivot",
            """
            { "op": "create_pivot", "target": { "sheet": "원가" },
              "sourceRange": "A1:G20", "destination": "A24", "name": "CostPivot",
              "rows": ["품목", ""] }
            """));
        Assert.Contains(errors, error => error.Contains("rows[1]"));
    }

    [Fact]
    public void Shape_formatting_contract_accepts_partial_updates_and_rejects_invalid_values()
    {
        var valid = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "update_shape",
            """
            { "op": "update_shape", "target": { "sheet": "월간보고" }, "name": "HighlightBox",
              "lineColor": "#112233", "lineWeight": 1.5, "lineVisible": false, "rotation": 360,
              "font": { "name": "Arial", "size": 10.5, "bold": true, "italic": false, "color": "#445566" } }
            """));
        Assert.Empty(valid);

        var invalid = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "update_textbox",
            """
            { "op": "update_textbox", "target": { "sheet": "월간보고" }, "name": "ReportNote",
              "lineColor": "not-a-color", "lineWeight": 0, "lineVisible": "no", "rotation": 12.5,
              "font": { "size": 0, "unknown": true, "color": "#XYZXYZ" } }
            """));
        Assert.Contains(invalid, error => error.Contains("lineColor"));
        Assert.Contains(invalid, error => error.Contains("lineWeight"));
        Assert.Contains(invalid, error => error.Contains("lineVisible"));
        Assert.Contains(invalid, error => error.Contains("rotation"));
        Assert.Contains(invalid, error => error.Contains("font.size"));
        Assert.Contains(invalid, error => error.Contains("font.unknown"));
        Assert.Contains(invalid, error => error.Contains("font.color"));

        var explicitNulls = ExcelDataOperationsContract.ValidatePublicInput(Op(
            "update_shape",
            """
            { "op": "update_shape", "target": { "sheet": "월간보고" }, "name": "HighlightBox",
              "lineVisible": null, "font": null }
            """));
        Assert.Contains(explicitNulls, error => error.Contains("lineVisible"));
        Assert.Contains(explicitNulls, error => error.Contains("font must be"));
    }

    [Fact]
    public void Shape_formatting_readback_normalizes_rotation_and_checks_only_requested_fields()
    {
        Assert.Equal(0, ExcelShapeFormatContract.NormalizeRequestedRotation(JsonValue.Create(0)!));
        Assert.Equal(0, ExcelShapeFormatContract.NormalizeRequestedRotation(JsonValue.Create(360)!));

        var requested = Op("update_shape", """
            { "op": "update_shape", "name": "HighlightBox", "target": { "sheet": "월간보고" },
              "rotation": 360, "font": { "bold": true } }
            """);
        var actual = new JsonObject
        {
            ["rotation"] = 0,
            ["lineColor"] = 3351057,
            ["font"] = new JsonObject { ["bold"] = true, ["name"] = "unchanged" },
        };
        requested["lineColor"] = "#112233";
        var mismatches = new List<string>();
        ExcelShapeFormatContract.CompareRequestedReadback(actual, requested, "월간보고!HighlightBox", mismatches);
        Assert.Empty(mismatches);

        actual["font"] = new JsonObject { ["bold"] = false };
        ExcelShapeFormatContract.CompareRequestedReadback(actual, requested, "월간보고!HighlightBox", mismatches);
        Assert.Contains(mismatches, mismatch => mismatch.Contains("font.bold"));
    }

    [Fact]
    public void Shape_font_readback_does_not_coerce_null_or_mixed_com_values()
    {
        Assert.False(ExcelShapeFormatContract.TryCreateFontReadbackValue("bold", null, out _));
        Assert.False(ExcelShapeFormatContract.TryCreateFontReadbackValue("color", DBNull.Value, out _));
        Assert.False(ExcelShapeFormatContract.TryCreateFontReadbackValue("italic", -2, out _));
        Assert.False(ExcelShapeFormatContract.TryNormalizeLineVisible(-2, out _));

        Assert.True(ExcelShapeFormatContract.TryCreateFontReadbackValue("bold", -1, out var bold));
        Assert.True(bold!.GetValue<bool>());
        Assert.True(ExcelShapeFormatContract.TryNormalizeLineVisible(0, out var visible));
        Assert.False(visible);
    }

    [Fact]
    public void Shape_numeric_readback_distinguishes_unavailable_properties_from_zero()
    {
        foreach (var field in new[] { "rotation", "lineWeight", "lineColor" })
        {
            Assert.False(ExcelShapeFormatContract.TryCreateNumericReadbackValue(field, null, out _));
            Assert.False(ExcelShapeFormatContract.TryCreateNumericReadbackValue(field, DBNull.Value, out _));
            Assert.False(ExcelShapeFormatContract.TryCreateNumericReadbackValue(field, double.NaN, out _));
            Assert.False(ExcelShapeFormatContract.TryCreateNumericReadbackValue(field, -2, out _));
        }
        Assert.True(ExcelShapeFormatContract.TryCreateNumericReadbackValue("rotation", 360f, out var rotation));
        Assert.Equal(0, rotation!.GetValue<double>());
        Assert.True(ExcelShapeFormatContract.TryCreateNumericReadbackValue("lineColor", 0, out var black));
        Assert.Equal(0, black!.GetValue<int>());
        Assert.True(ExcelShapeFormatContract.TryCreateNumericReadbackValue("lineWeight", 1.5f, out var weight));
        Assert.Equal(1.5, weight!.GetValue<double>());
        Assert.False(ExcelShapeFormatContract.TryCreateNumericReadbackValue("lineColor", 1.5, out _));
        Assert.False(ExcelShapeFormatContract.TryCreateNumericReadbackValue("lineWeight", 0, out _));
    }

    private static JsonObject Op(string _, string json) =>
        Json.ParseObject(json) ?? throw new InvalidOperationException(json);
}
