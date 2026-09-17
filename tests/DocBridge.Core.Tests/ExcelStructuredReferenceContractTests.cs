using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelStructuredReferenceContractTests
{
    [Theory]
    [InlineData("Table1[Amount]", "Table1")]
    [InlineData("Table1[[#Data],[금액]]", "Table1")]
    [InlineData("[@[Column Name]]", null)]
    [InlineData("[[#This Row],[Amount]]", null)]
    public void Parser_accepts_named_and_local_structured_forms(string token, string? table)
    {
        Assert.True(ExcelStructuredReferenceContract.TryParse(token, out var parsed, out var reason), reason);
        Assert.Equal(table, parsed.TableName);
    }

    [Fact]
    public void Parser_preserves_escaped_special_column_names_and_rejects_bad_selector()
    {
        Assert.True(ExcelStructuredReferenceContract.TryParse("Table1['#상태]", out var parsed, out _));
        Assert.Equal("#상태", parsed.Selectors.Single().Column);
        Assert.False(ExcelStructuredReferenceContract.TryParse("Table1[[#Data][Amount]]", out _, out var reason));
        Assert.Equal("structured_selector_malformed", reason);
    }

    [Fact]
    public void Formula_parser_keeps_complex_structured_tokens_and_external_links_separate()
    {
        var references = ExcelFormulaTraceContract.ParseReferences("=Table1[[#Headers],[#Data],[금 액]]+[@[A'[B']]]+[book.xlsx]Data!A1");
        Assert.Contains(references, reference => reference.Kind == "structured" && reference.Token == "Table1[[#Headers],[#Data],[금 액]]");
        Assert.Contains(references, reference => reference.Kind == "structured" && reference.Token == "[@[A'[B']]]");
        Assert.Contains(references, reference => reference.Reason == "external_workbook_reference");
    }

    [Fact]
    public void Planner_selects_disjoint_parts_and_column_ranges()
    {
        Assert.True(ExcelStructuredReferenceContract.TryParse("Table1[[#Headers],[#Data],[B]:[C]]", out var parsed, out _));
        var plan = ExcelStructuredReferenceContract.Plan(parsed, Bounds(), Columns(), 11);
        Assert.True(plan.Resolved);
        Assert.Equal(new[] { "C10:D10", "C11:D12" }, plan.Addresses);
    }

    [Fact]
    public void Planner_current_row_requires_data_body_membership_and_handles_empty_optional_parts()
    {
        Assert.True(ExcelStructuredReferenceContract.TryParse("[[#This Row],[금액]]", out var current, out _));
        var selected = ExcelStructuredReferenceContract.Plan(current, Bounds(), Columns(), 12);
        Assert.Equal(new[] { "C12" }, selected.Addresses);
        Assert.Equal("structured_current_row_outside_data_body", ExcelStructuredReferenceContract.Plan(current, Bounds(), Columns(), 9).Reason);

        Assert.True(ExcelStructuredReferenceContract.TryParse("Table1[[#All],[Amount]]", out var all, out _));
        var empty = ExcelStructuredReferenceContract.Plan(all, new ExcelStructuredReferenceContract.TableBounds(2, 3, 10, null, null, null), Columns(), 0);
        Assert.Equal(new[] { "B10" }, empty.Addresses);
    }

    private static ExcelStructuredReferenceContract.TableBounds Bounds() => new(2, 3, 10, 11, 12, 13);
    private static IReadOnlyDictionary<string, int> Columns() => new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
    {
        ["A"] = 1, ["B"] = 2, ["C"] = 3, ["Amount"] = 1, ["금액"] = 2,
    };
}