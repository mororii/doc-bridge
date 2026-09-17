using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelFormulaTraceOutputContractTests
{
    [Fact]
    public void Compact_trace_preserves_evidence_and_unambiguous_cross_sheet_citations()
    {
        var full = JsonNode.Parse("""
        {"workbook":"C:/work/budget.xlsx","nodes":[
          {"workbook":"C:/work/budget.xlsx","sheet":"O'Brien","address":"A1","value":12,"formula":"=B1*2","errorClassification":"not_excel_error"},
          {"workbook":"C:/other.xlsx","sheet":"Data","address":"B1","value":6}],
         "edges":[{"from":{"sheet":"O'Brien","address":"A1"},"to":{"sheet":"원가 표","address":"B1"},"via":"'원가 표'!B1"}],
         "unresolved":[{"from":{"sheet":"O'Brien","address":"A1"},"reference":"INDIRECT(\"C1\")","reason":"dynamic_reference"}],
         "coverage":{"complete":false,"truncated":true,"frontier":["원가 표!C1"]}}
        """)!.AsObject();
        var compact = ExcelFormulaTraceOutputContract.Compact(full.DeepClone().AsObject());
        Assert.Equal(full["workbook"]!.GetValue<string>(), compact["workbook"]!.GetValue<string>());
        Assert.Null(compact["nodes"]![0]!["workbook"]);
        Assert.Equal("C:/other.xlsx", compact["nodes"]![1]!["workbook"]!.GetValue<string>());
        foreach (var field in new[] { "sheet", "address", "value", "formula", "errorClassification" })
            Assert.True(JsonNode.DeepEquals(full["nodes"]![0]![field], compact["nodes"]![0]![field]));
        Assert.Equal("'O''Brien'!A1", compact["edges"]![0]!["from"]!.GetValue<string>());
        Assert.Equal("'원가 표'!B1", compact["edges"]![0]!["to"]!.GetValue<string>());
        Assert.Equal("dynamic_reference", compact["unresolved"]![0]!["reason"]!.GetValue<string>());
        Assert.True(JsonNode.DeepEquals(full["coverage"], compact["coverage"]));
        Assert.True(JsonNode.DeepEquals(compact, ExcelFormulaTraceOutputContract.Compact(compact.DeepClone().AsObject())));
    }
}
