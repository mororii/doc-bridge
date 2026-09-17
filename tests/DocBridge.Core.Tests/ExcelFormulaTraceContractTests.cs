using System.Text.Json.Nodes;

using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelFormulaTraceContractTests
{
    [Fact]
    public void Batch_plan_is_row_major_and_maps_com_scalar_and_one_based_arrays()
    {
        Assert.Collection(ExcelFormulaTraceReadPlan.Plan(new ExcelA1Box(2, 3, 3, 4), 10),
            full => Assert.Equal("C2:F3", full.Address),
            tail => Assert.Equal("C4:D4", tail.Address));
        Assert.Single(ExcelFormulaTraceReadPlan.Plan(new ExcelA1Box(1, 1, ExcelA1Box.MaxRow, ExcelA1Box.MaxColumn), 2));
        Assert.Equal(12d, ExcelFormulaTraceReadPlan.ValueAt(12d, 1, 1));
        Assert.Null(ExcelFormulaTraceReadPlan.ValueAt(null, 1, 1));
        var values = Array.CreateInstance(typeof(object), new[] { 2, 2 }, new[] { 1, 1 });
        values.SetValue("A", 1, 1); values.SetValue("D", 2, 2);
        Assert.Equal("A", ExcelFormulaTraceReadPlan.ValueAt(values, 1, 1));
        Assert.Equal("D", ExcelFormulaTraceReadPlan.ValueAt(values, 2, 2));
    }

    [Fact]
    public void Traversal_reports_one_closing_edge_when_both_cycle_cells_are_roots()
    {
        var a = Cell("Calc", "A1", "=A2");
        var b = Cell("Calc", "A2", "=A1");
        var trace = ExcelFormulaTraceContract.Trace(new[] { a, b }, 4, 10,
            (_, reference, _) => new ExcelFormulaTraceContract.Resolution(
                new[] { reference.Address == "A1" ? a : b }, Array.Empty<ExcelFormulaTraceContract.Issue>()));
        Assert.Equal(2, trace.Nodes.Count);
        Assert.Equal(2, trace.Edges.Count);
        Assert.Single(trace.Cycles);
        Assert.True(trace.Complete);
    }


    [Fact]
    public void Parser_handles_quoted_sheets_ranges_names_and_skips_literals_and_function_tokens()
    {
        var references = ExcelFormulaTraceContract.ParseReferences(
            "=SUM('원가 표'!$A$1:B2,Data!C3)+Named_Rate+\"A1 and Sheet!B2\"+LOG10(10)");

        Assert.Contains(references, item => item.Kind == "cell" && item.Sheet == "원가 표" && item.Address == "A1:B2");
        Assert.Contains(references, item => item.Kind == "cell" && item.Sheet == "Data" && item.Address == "C3");
        Assert.Contains(references, item => item.Kind == "name" && item.Token == "Named_Rate");
        Assert.DoesNotContain(references, item => item.Token.Contains("A1 and", StringComparison.Ordinal));
        Assert.DoesNotContain(references, item => item.Kind == "cell" && item.Token.StartsWith("LOG10", StringComparison.Ordinal));
    }

    [Fact]
    public void Parser_reports_external_dynamic_3d_structured_and_missing_references()
    {
        var references = ExcelFormulaTraceContract.ParseReferences(
            "=[other.xlsx]Data!A1+INDIRECT(\"B2\")+OFFSET(A1,1,0)+Jan:Dec!A1+Table1[Amount]+#REF!");

        Assert.Contains(references, item => item.Reason == "external_workbook_reference");
        Assert.Contains(references, item => item.Reason == "dynamic_reference" && item.Token.StartsWith("INDIRECT", StringComparison.Ordinal));
        Assert.Contains(references, item => item.Reason == "dynamic_reference" && item.Token.StartsWith("OFFSET", StringComparison.Ordinal));
        Assert.Contains(references, item => item.Reason == "three_dimensional_reference");
        Assert.Contains(references, item => item.Kind == "structured" && item.Token == "Table1[Amount]");
        Assert.Contains(references, item => item.Reason == "missing_reference");
    }

    [Fact]
    public void Parser_preserves_spill_axis_and_nested_structured_references()
    {
        var wholeRows = ExcelFormulaTraceContract.ParseReferences("=SUM(1:3)");
        var wholeColumns = ExcelFormulaTraceContract.ParseReferences("=SUM(A:A)");
        var spill = ExcelFormulaTraceContract.ParseReferences("=SUM(A1#)");
        var scientific = ExcelFormulaTraceContract.ParseReferences("=1E3+A1");
        var quotedThreeDimensional = ExcelFormulaTraceContract.ParseReferences("='Jan:Dec'!A1");

        Assert.Contains(wholeRows, item => item.Kind == "axis" && item.Address == "A1:XFD3");
        Assert.Contains(wholeColumns, item => item.Kind == "axis" && item.Address == "A1:A1048576");
        Assert.Contains(spill, item => item.Kind == "spill" && item.Address == "A1");
        Assert.DoesNotContain(scientific, item => item.Kind == "name" && item.Token == "E3");
        Assert.Contains(scientific, item => item.Kind == "cell" && item.Address == "A1");
        Assert.Single(quotedThreeDimensional);
        Assert.Equal("three_dimensional_reference", quotedThreeDimensional[0].Reason);
    }

    [Fact]
    public void Parser_preserves_qualified_axis_spills_and_nested_table_tokens()
    {
        var references = ExcelFormulaTraceContract.ParseReferences("=Data!A1#+'원가 표'!$1:$3+Table1[[#Data],[Amount]]+Table1[Amount]");

        Assert.Contains(references, item => item.Kind == "spill" && item.Sheet == "Data" && item.Address == "A1");
        Assert.Contains(references, item => item.Kind == "cell" && item.Sheet == "원가 표" && item.Address == "A1:XFD3");
        Assert.Contains(references, item => item.Kind == "structured" && item.Token == "Table1[[#Data],[Amount]]" && item.Address == "Table1");
        Assert.Contains(references, item => item.Kind == "structured" && item.Token == "Table1[Amount]" && item.Address == "Table1");
    }

    [Fact]
    public void Traversal_detects_cycle_after_resolved_cell_edges()
    {
        var cells = new Dictionary<string, ExcelFormulaTraceContract.Cell>(StringComparer.OrdinalIgnoreCase)
        {
            ["Calc!A1"] = Cell("Calc", "A1", "=B1"),
            ["Calc!B1"] = Cell("Calc", "B1", "=A1"),
        };

        var trace = ExcelFormulaTraceContract.Trace(
            new[] { cells["Calc!A1"] }, 4, 10,
            (source, reference, _) => Resolve(cells, source, reference));

        Assert.Equal(2, trace.Nodes.Count);
        Assert.Equal(2, trace.Edges.Count);
        Assert.Single(trace.Cycles);
        Assert.True(trace.Complete);
    }

    [Fact]
    public void Traversal_respects_depth_and_cell_limits_without_allocating_unbounded_range_nodes()
    {
        var cells = new Dictionary<string, ExcelFormulaTraceContract.Cell>(StringComparer.OrdinalIgnoreCase)
        {
            ["Calc!A1"] = Cell("Calc", "A1", "=B1:C1"),
            ["Calc!B1"] = Cell("Calc", "B1", "=D1"),
            ["Calc!C1"] = Cell("Calc", "C1", null),
            ["Calc!D1"] = Cell("Calc", "D1", null),
        };

        var cellsLimited = ExcelFormulaTraceContract.Trace(
            new[] { cells["Calc!A1"] }, 4, 2,
            (source, reference, remaining) => Resolve(cells, source, reference, remaining));
        Assert.Equal(2, cellsLimited.Nodes.Count);
        Assert.Contains("max_cells", cellsLimited.TruncationReasons);
        Assert.False(cellsLimited.Complete);

        var depthLimited = ExcelFormulaTraceContract.Trace(
            new[] { cells["Calc!A1"] }, 0, 10,
            (source, reference, remaining) => Resolve(cells, source, reference, remaining));
        Assert.Single(depthLimited.Nodes);
        Assert.Empty(depthLimited.Edges);
        Assert.Contains("max_depth", depthLimited.TruncationReasons);
    }

    private static ExcelFormulaTraceContract.Resolution Resolve(
        IReadOnlyDictionary<string, ExcelFormulaTraceContract.Cell> cells,
        ExcelFormulaTraceContract.Cell source,
        ExcelFormulaTraceContract.Reference reference,
        int remaining = int.MaxValue)
    {
        if (reference.Kind != "cell" || string.IsNullOrWhiteSpace(reference.Address))
            return new ExcelFormulaTraceContract.Resolution(Array.Empty<ExcelFormulaTraceContract.Cell>(),
                new[] { new ExcelFormulaTraceContract.Issue(source.Sheet, source.Address, reference, "unresolved") });
        var sheet = reference.Sheet ?? source.Sheet;
        var requested = reference.Address.Split(':');
        var targets = requested
            .Select(address => cells.TryGetValue(sheet + "!" + address, out var target) ? target : null)
            .Where(target => target is not null)
            .Cast<ExcelFormulaTraceContract.Cell>()
            .Take(remaining)
            .ToArray();
        return new ExcelFormulaTraceContract.Resolution(targets, Array.Empty<ExcelFormulaTraceContract.Issue>(),
            requested.Length > remaining, requested.Length > remaining ? "max_cells" : null);
    }

    private static ExcelFormulaTraceContract.Cell Cell(string sheet, string address, string? formula) =>
        new(sheet, address, formula, new JsonObject { ["sheet"] = sheet, ["address"] = address, ["formula"] = formula });
}
