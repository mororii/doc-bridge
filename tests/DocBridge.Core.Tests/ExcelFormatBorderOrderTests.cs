using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelFormatBorderOrderTests
{
    private readonly OperationValidator _v = new(new PolicyEngine());

    [Fact]
    public void Draw_then_clear_on_shared_edge_is_rejected()
    {
        var errors = ValidateBatch(
            """{ "op": "format_range", "target": { "sheet": "검수요청서" }, "range": "F16:H18", "style": { "borders": { "outline": { "weight": "medium" } } } }""",
            """{ "op": "format_range", "target": { "sheet": "검수요청서" }, "range": "A16:E18", "style": { "borders": { "all": "none" } } }""");
        Assert.Contains(errors, error => error.Contains("EXCEL_BORDER_ORDER"));
    }

    [Fact]
    public void Clear_then_draw_on_shared_edge_is_accepted()
    {
        var errors = ValidateBatch(
            """{ "op": "format_range", "target": { "sheet": "검수요청서" }, "range": "A16:E18", "style": { "borders": { "all": "none" } } }""",
            """{ "op": "format_range", "target": { "sheet": "검수요청서" }, "range": "F16:H18", "style": { "borders": { "outline": { "weight": "medium" } } } }""");
        Assert.DoesNotContain(errors, error => error.Contains("EXCEL_BORDER_ORDER"));
    }

    [Fact]
    public void Same_range_and_other_sheets_are_exempt()
    {
        var same = ValidateBatch(
            """{ "op": "format_range", "target": { "sheet": "검수요청서" }, "range": "F16:H18", "style": { "borders": { "outline": { "weight": "medium" } } } }""",
            """{ "op": "format_range", "target": { "sheet": "검수요청서" }, "range": "F16:H18", "style": { "borders": { "all": "none" } } }""");
        Assert.DoesNotContain(same, error => error.Contains("EXCEL_BORDER_ORDER"));

        var other = ValidateBatch(
            """{ "op": "format_range", "target": { "sheet": "계획" }, "range": "F16:H18", "style": { "borders": { "outline": { "weight": "medium" } } } }""",
            """{ "op": "format_range", "target": { "sheet": "집계" }, "range": "A16:E18", "style": { "borders": { "all": "none" } } }""");
        Assert.DoesNotContain(other, error => error.Contains("EXCEL_BORDER_ORDER"));
    }

    [Fact]
    public void Non_adjacent_ranges_are_exempt()
    {
        var errors = ValidateBatch(
            """{ "op": "format_range", "target": { "sheet": "검수요청서" }, "range": "F16:H18", "style": { "borders": { "outline": { "weight": "medium" } } } }""",
            """{ "op": "format_range", "target": { "sheet": "검수요청서" }, "range": "A20:H21", "style": { "borders": { "all": "none" } } }""");
        Assert.DoesNotContain(errors, error => error.Contains("EXCEL_BORDER_ORDER"));
    }

    private List<string> ValidateBatch(params string[] ops)
    {
        var array = new JsonArray();
        foreach (var json in ops)
            array.Add(Json.ParseObject(json));
        var batch = new JsonObject { ["ops"] = array, ["dryRun"] = true };
        var errors = new List<string>();
        _v.Validate(batch, "excel", errors);
        return errors;
    }
}
