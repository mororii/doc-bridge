using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelDataOperationsReadbackTests
{
    [Fact]
    public void E4_totals_require_requested_function_and_finite_value()
    {
        var actual = Json.ParseObject(
            """
            {
              "name": "자재표",
              "showTotals": true,
              "totals": [
                { "column": "수량", "function": "sum", "value": 42 },
                { "column": "재고", "function": "sum", "value": 17 }
              ]
            }
            """)!;
        var op = Json.ParseObject(
            """
            {
              "op": "set_table_totals",
              "name": "자재표",
              "showTotals": true,
              "columns": [
                { "column": "수량", "function": "sum" },
                { "column": "재고", "function": "sum" }
              ]
            }
            """)!;
        var errors = new List<string>();
        ExcelDataReadbackContract.CompareRequestedTotals(actual, op, "자재대장!자재표", errors);
        Assert.Empty(errors);
    }

    [Fact]
    public void E4_totals_reject_function_only_without_aggregate()
    {
        var actual = Json.ParseObject(
            """
            {
              "name": "자재표",
              "showTotals": true,
              "totals": [{ "column": "수량", "function": "sum" }]
            }
            """)!;
        var op = Json.ParseObject(
            """
            { "op": "set_table_totals", "columns": [{ "column": "수량", "function": "sum" }] }
            """)!;
        var errors = new List<string>();
        ExcelDataReadbackContract.CompareRequestedTotals(actual, op, "자재대장!자재표", errors);
        Assert.Contains(errors, error => error.Contains("finite aggregate", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void E7_pivot_requires_layout_and_numeric_aggregates()
    {
        var actual = Json.ParseObject(
            """
            {
              "name": "LedgerPivot",
              "rowFields": ["품목"],
              "columnFields": ["월"],
              "dataFields": [{ "name": "합계 : 금액", "function": "sum" }],
              "aggregates": [
                { "name": "합계 : 금액", "field": "금액", "function": "sum", "values": [2500000, 630000], "value": 2500000, "hasAggregate": true }
              ]
            }
            """)!;
        var op = Json.ParseObject(
            """
            {
              "op": "create_pivot",
              "name": "LedgerPivot",
              "rows": ["품목"],
              "columns": ["월"],
              "values": [{ "field": "금액", "function": "sum" }]
            }
            """)!;
        var errors = new List<string>();
        ExcelDataReadbackContract.CompareRequestedPivot(actual, op, "피벗!LedgerPivot", errors, requireAggregates: true);
        Assert.Empty(errors);
    }

    [Fact]
    public void Refresh_pivot_name_only_is_not_enough()
    {
        var actual = Json.ParseObject("""{ "name": "LedgerPivot" }""")!;
        var op = Json.ParseObject("""{ "op": "refresh_pivot", "name": "LedgerPivot" }""")!;
        var errors = new List<string>();
        ExcelDataReadbackContract.CompareRequestedPivot(actual, op, "피벗!LedgerPivot", errors, requireAggregates: true);
        Assert.Contains(errors, error => error.Contains("name-only", StringComparison.OrdinalIgnoreCase));
        Assert.True(ExcelDataReadbackContract.IsNameOnlyPivot(actual));
    }

    [Fact]
    public void Multi_key_sort_requires_order_and_row_associations()
    {
        var before = Json.ParseObject(
            """
            { "rows": [
              ["일자", "전표", "품목"],
              ["2026-09-02", "B", "관"],
              ["2026-09-01", "A", "볼트"],
              ["2026-09-01", "C", "관"]
            ] }
            """)!["rows"]!.AsArray();
        var after = Json.ParseObject(
            """
            { "rows": [
              ["일자", "전표", "품목"],
              ["2026-09-01", "A", "볼트"],
              ["2026-09-01", "C", "관"],
              ["2026-09-02", "B", "관"]
            ] }
            """)!["rows"]!.AsArray();
        var keys = new List<(int GridColumn, bool Descending)> { (0, false), (1, false) };
        var errors = new List<string>();
        ExcelDataReadbackContract.CompareSortReadback(before, after, keys, hasHeaders: true, "자재대장!A1:C4", errors);
        Assert.Empty(errors);

        var brokenAssoc = Json.ParseObject(
            """
            { "rows": [
              ["일자", "전표", "품목"],
              ["2026-09-01", "A", "관"],
              ["2026-09-01", "C", "볼트"],
              ["2026-09-02", "B", "관"]
            ] }
            """)!["rows"]!.AsArray();
        errors.Clear();
        ExcelDataReadbackContract.CompareSortReadback(before, brokenAssoc, keys, hasHeaders: true, "자재대장!A1:C4", errors);
        Assert.Contains(errors, error => error.Contains("association", StringComparison.OrdinalIgnoreCase));

        var brokenOrder = Json.ParseObject(
            """
            { "rows": [
              ["일자", "전표", "품목"],
              ["2026-09-02", "B", "관"],
              ["2026-09-01", "A", "볼트"],
              ["2026-09-01", "C", "관"]
            ] }
            """)!["rows"]!.AsArray();
        errors.Clear();
        ExcelDataReadbackContract.CompareSortReadback(before, brokenOrder, keys, hasHeaders: true, "자재대장!A1:C4", errors);
        Assert.Contains(errors, error => error.Contains("multi-key", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Text_to_columns_requires_every_destination_field()
    {
        var source = Json.ParseObject(
            """{ "rows": [["A-1"], ["B-2"], ["C-3"]] }""")!["rows"]!.AsArray();
        var dest = Json.ParseObject(
            """{ "rows": [["A", "1"], ["B", "2"], ["C", "3"]] }""")!["rows"]!.AsArray();
        var op = Json.ParseObject("""{ "op": "text_to_columns", "other": true, "otherChar": "-", "comma": false }""")!;
        var errors = new List<string>();
        ExcelDataReadbackContract.CompareTextToColumnsReadback(source, dest, op, "B2:B4->B2:C4", errors);
        Assert.Empty(errors);

        var firstOnly = Json.ParseObject(
            """{ "rows": [["A", "1"], ["B-2"], ["C-3"]] }""")!["rows"]!.AsArray();
        errors.Clear();
        ExcelDataReadbackContract.CompareTextToColumnsReadback(source, firstOnly, op, "B2:B4->B2:C4", errors);
        Assert.Contains(errors, error => error.Contains("row 2", StringComparison.OrdinalIgnoreCase));

        var leftover = Json.ParseObject(
            """{ "rows": [["A", "1", "x"], ["B", "2"], ["C", "3"]] }""")!["rows"]!.AsArray();
        errors.Clear();
        ExcelDataReadbackContract.CompareTextToColumnsReadback(source, leftover, op, "B2:B4->B2:C4", errors);
        Assert.Contains(errors, error => error.Contains("leftover", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Text_to_columns_quoted_comma_and_leading_zeros_use_excel_general()
    {
        var op = Json.ParseObject(
            """{ "op": "text_to_columns", "comma": true, "textQualifier": "double" }""")!;
        Assert.Equal(ExcelDataObjectCatalog.XlTextQualifierDoubleQuote, ExcelTextToColumnsParse.ComTextQualifier(op));

        var quoted = ExcelTextToColumnsParse.ExpectedFields("\"a,b\",c", op);
        Assert.Equal(new[] { "a,b", "c" }, quoted);

        var zeros = ExcelTextToColumnsParse.ExpectedFields("\"001\",002", op);
        Assert.Equal(new[] { "1", "2" }, zeros);

        var source = Json.ParseObject(
            """{ "rows": [["\"001\",002"], ["\"a,b\",c"]] }""")!["rows"]!.AsArray();
        var dest = Json.ParseObject(
            """{ "rows": [["1", "2"], ["a,b", "c"]] }""")!["rows"]!.AsArray();
        var errors = new List<string>();
        ExcelDataReadbackContract.CompareTextToColumnsReadback(source, dest, op, "B2:B3->B2:C3", errors);
        Assert.Empty(errors);

        var naiveDest = Json.ParseObject(
            """{ "rows": [["\"001\"", "002"], ["\"a", "b\"", "c"]] }""")!["rows"]!.AsArray();
        errors.Clear();
        ExcelDataReadbackContract.CompareTextToColumnsReadback(source, naiveDest, op, "B2:B3->B2:C3", errors);
        Assert.NotEmpty(errors);
    }
}
