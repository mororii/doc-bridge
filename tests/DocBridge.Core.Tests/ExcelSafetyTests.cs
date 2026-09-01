using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelSafetyTests
{
    [Theory]
    [InlineData("Sheet1!A1", "Sheet1", "A1")]
    [InlineData("'공사 내역'!B2:D5", "공사 내역", "B2:D5")]
    [InlineData("'홍길동''s'!$C$7", "홍길동's", "$C$7")]
    [InlineData("A1:B3", null, "A1:B3")]
    public void Sheet_qualified_ranges_are_parsed(string input, string? expectedSheet, string expectedAddress)
    {
        var parsed = ExcelRangeReference.Parse(input);
        Assert.Equal(expectedSheet, parsed.SheetName);
        Assert.Equal(expectedAddress, parsed.Address);
    }

    [Theory]
    [InlineData("'미완성!A1")]
    [InlineData("!A1")]
    [InlineData("Sheet1!")]
    [InlineData("Sheet1!A1!B2")]
    public void Invalid_sheet_qualified_ranges_are_rejected(string input)
        => Assert.Throws<FormatException>(() => ExcelRangeReference.Parse(input));

    [Fact]
    public void Excel_numeric_json_values_are_normalized_to_double()
    {
        Assert.IsType<double>(ExcelAdapter.NodeToComValue(JsonValue.Create(1500)));
        Assert.Equal(1500d, ExcelAdapter.NodeToComValue(JsonValue.Create(1500)));
        Assert.Equal(4_294_967_296d, ExcelAdapter.NodeToComValue(JsonValue.Create(4_294_967_296L)));
        Assert.Equal(12.5d, ExcelAdapter.NodeToComValue(JsonValue.Create(12.5m)));
        Assert.Equal(true, ExcelAdapter.NodeToComValue(JsonValue.Create(true)));
        Assert.Equal("001500", ExcelAdapter.NodeToComValue(JsonValue.Create("001500")));
        Assert.Null(ExcelAdapter.NodeToComValue(null));
    }

    [Fact]
    public void Excel_refuses_integer_values_that_cannot_round_trip_exactly()
        => Assert.Throws<ArgumentOutOfRangeException>(() =>
            ExcelAdapter.NodeToComValue(JsonValue.Create(9_007_199_254_740_992L)));

    [Fact]
    public void Excel_write_requires_an_explicit_sheet_but_accepts_two_explicit_forms()
    {
        var validator = new OperationValidator(new PolicyEngine());

        var ambiguous = Json.ParseObject("""
        { "ops": [ { "op": "set_values", "range": "A1", "values": [[1]] } ], "dryRun": true }
        """);
        var ambiguousErrors = new List<string>();
        Assert.Null(validator.Validate(ambiguous, "excel", ambiguousErrors));
        Assert.Contains(ambiguousErrors, error => error.Contains("active sheet writes are not allowed"));

        var targetObject = Json.ParseObject("""
        { "ops": [ { "op": "set_values", "target": { "sheet": "매출" }, "range": "A1", "values": [[1]] } ], "dryRun": true }
        """);
        var targetErrors = new List<string>();
        Assert.NotNull(validator.Validate(targetObject, "excel", targetErrors));
        Assert.Empty(targetErrors);

        var qualifiedRange = Json.ParseObject("""
        { "ops": [ { "op": "set_values", "range": "'매출 내역'!A1", "values": [[1]] } ], "dryRun": true }
        """);
        var qualifiedErrors = new List<string>();
        Assert.NotNull(validator.Validate(qualifiedRange, "excel", qualifiedErrors));
        Assert.Empty(qualifiedErrors);
    }

    [Fact]
    public void Excel_rejects_conflicting_sheet_targets_and_ambiguous_sheet_scopes()
    {
        var validator = new OperationValidator(new PolicyEngine());
        var conflicting = Json.ParseObject("""
        { "ops": [ { "op": "format_range", "target": { "sheet": "매출" }, "range": "원가!A1", "style": { "bold": true } } ], "dryRun": true }
        """);
        var conflictingErrors = new List<string>();
        Assert.Null(validator.Validate(conflicting, "excel", conflictingErrors));
        Assert.Contains(conflictingErrors, error => error.Contains("does not match"));

        var insert = Json.ParseObject("""
        { "ops": [ { "op": "insert_rows", "row": 2, "count": 1 } ], "dryRun": true }
        """);
        var insertErrors = new List<string>();
        Assert.Null(validator.Validate(insert, "excel", insertErrors));
        Assert.Contains(insertErrors, error => error.Contains("requires target.sheet"));

        var sheetReplace = Json.ParseObject("""
        { "ops": [ { "op": "find_replace", "find": "A", "replace": "B", "target": { "scope": "sheet" } } ], "dryRun": true }
        """);
        var sheetErrors = new List<string>();
        Assert.Null(validator.Validate(sheetReplace, "excel", sheetErrors));

        var workbookReplace = Json.ParseObject("""
        { "ops": [ { "op": "find_replace", "find": "A", "replace": "B", "target": { "scope": "workbook" } } ], "dryRun": true }
        """);
        var workbookErrors = new List<string>();
        Assert.NotNull(validator.Validate(workbookReplace, "excel", workbookErrors));
        Assert.Empty(workbookErrors);
    }

    [Fact]
    public void Excel_capabilities_publish_real_read_and_write_operations()
    {
        using var adapter = new ExcelAdapter(() => null);
        var capabilities = adapter.GetCapabilities();
        var readOps = Json.GetArr(capabilities, "readOps")!
            .Select(node => node!.GetValue<string>()).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var writeOps = Json.GetArr(capabilities, "writeOps")!
            .Select(node => node!.GetValue<string>()).ToHashSet(StringComparer.OrdinalIgnoreCase);

        Assert.Contains("layout", readOps);
        Assert.Contains("merge_cells", writeOps);
        Assert.Contains("unmerge_cells", writeOps);
        Assert.Contains("set_rows_hidden", writeOps);
        Assert.Contains("set_cols_hidden", writeOps);
        Assert.Contains("set_sheet_visibility", writeOps);
        Assert.Equal(2_000, Json.GetInt(Json.GetObj(capabilities, "limits"), "maxMergeCells"));
        Assert.Contains(Json.GetArr(capabilities, "safety")!, node =>
            string.Equals(node?.GetValue<string>(), "merge-content-loss-block", StringComparison.Ordinal));
    }
}
