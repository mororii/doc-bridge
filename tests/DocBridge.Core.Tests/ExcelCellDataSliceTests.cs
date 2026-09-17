using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelCellDataSliceTests
{
    [Fact]
    public void Rich_text_validates_utf16_length_and_resets_each_run_to_base_font()
    {
        var op = JsonNode.Parse("""{"op":"set_rich_text","target":{"sheet":"S"},"cell":"A1","baseFont":{"name":"Aptos","size":11,"bold":false,"italic":false,"underline":false,"color":"#112233"},"runs":[{"text":"a","font":{"bold":true}},{"text":"b"}]}""")!.AsObject();
        Assert.Empty(ExcelDataOperationsContract.ValidatePublicInput(op));
        var second = ExcelRichTextContract.EffectiveFont(Json.GetObj(op, "baseFont")!, null);
        Assert.False(Json.GetBool(second, "bold"));
        Assert.Equal("#112233", Json.GetString(second, "color"));

        var oversized = (JsonObject)op.DeepClone();
        oversized["runs"] = new JsonArray(new JsonObject { ["text"] = new string('x', ExcelRichTextContract.MaxUtf16Length + 1) });
        Assert.Contains(ExcelDataOperationsContract.ValidatePublicInput(oversized), x => x.Contains("exceeds"));
    }

    [Fact]
    public void Rich_text_rejects_unpaired_run_shape_and_table_rows_are_rectangular()
    {
        var rich = JsonNode.Parse("""{"op":"set_rich_text","target":{"sheet":"S"},"cell":"A1","baseFont":{},"runs":[{"text":""}]}""")!.AsObject();
        Assert.NotEmpty(ExcelDataOperationsContract.ValidatePublicInput(rich));
        var rows = JsonNode.Parse("""{"op":"append_table_rows","target":{"sheet":"S"},"name":"Items","rows":[[1,2],[3]]}""")!.AsObject();
        Assert.Contains(ExcelDataOperationsContract.ValidatePublicInput(rows), x => x.Contains("rectangular"));
    }

    [Fact]
    public void Rich_text_rejects_null_surrogate_and_missing_base_fields()
    {
        var rich = JsonNode.Parse("""{"op":"set_rich_text","target":{"sheet":"S"},"cell":"A1","baseFont":{"name":null},"runs":[{"text":"😀"}]}""")!.AsObject();
        var errors = ExcelDataOperationsContract.ValidatePublicInput(rich);
        Assert.Contains(errors, x => x.Contains("cannot be null"));
        Assert.Contains(errors, x => x.Contains("surrogate"));
    }

    [Fact]
    public void Table_scalar_contract_preserves_scalar_kinds_and_sparse_rows()
    {
        Assert.Equal(ExcelValueWriteContract.JsonKind.Number, ExcelValueWriteContract.Classify(JsonValue.Create(1)).Kind);
        Assert.Equal(ExcelValueWriteContract.JsonKind.Boolean, ExcelValueWriteContract.Classify(JsonValue.Create(true)).Kind);
        Assert.Equal(ExcelValueWriteContract.JsonKind.Null, ExcelValueWriteContract.Classify(null).Kind);
        var sparse = JsonNode.Parse("""{"op":"append_table_rows","target":{"sheet":"S"},"name":"Items","rows":[{"Amount":1,"Enabled":true}]}""")!.AsObject();
        Assert.Empty(ExcelDataOperationsContract.ValidatePublicInput(sparse));
    }

    [Fact]
    public void Conditional_rule_fingerprint_is_stable_and_detects_identity_change()
    {
        var rule = JsonNode.Parse("""{"index":3,"conditionType":2,"operator":5,"formula1":"=A1>0","style":{"bold":true}}""")!.AsObject();
        var fingerprint = ExcelConditionalFormatContract.Fingerprint(rule);
        Assert.True(ExcelConditionalFormatContract.Matches(rule, fingerprint));
        rule["formula1"] = "=A1>1";
        Assert.False(ExcelConditionalFormatContract.Matches(rule, fingerprint));
        var requested = JsonNode.Parse("""{"type":"expression","formula1":"=A1>1"}""")!.AsObject();
        Assert.True(ExcelConditionalFormatContract.MatchesRequested(rule, requested, null));
        requested["formula1"] = "=A1>2";
        Assert.False(ExcelConditionalFormatContract.MatchesRequested(rule, requested, null));
    }

    [Fact]
    public void New_public_operations_have_schema_branches()
    {
        foreach (var op in new[] { "set_rich_text", "update_conditional_format", "delete_conditional_format", "append_table_rows", "insert_table_rows", "delete_table_rows" })
            Assert.NotNull(ExcelDataOperationSchema.OpSchema(op));
    }

    [Fact]
    public void Table_rows_reject_unknown_columns_and_nested_values_before_writing()
    {
        var rows = JsonNode.Parse("""[{"Amount":1,"Enabled":true},{"Amuont":2}]""")!.AsArray();
        Assert.Throws<InvalidOperationException>(() => ExcelTableRowsContract.ValidateKnownColumns(rows, new[] { "Amount", "Enabled" }));
        rows.RemoveAt(1);
        ExcelTableRowsContract.ValidateKnownColumns(rows, new[] { "Amount", "Enabled", "Total" });
        var op = JsonNode.Parse("""{"op":"append_table_rows","target":{"sheet":"S"},"name":"Items","rows":[{"Amount":{"nested":1}}]}""")!.AsObject();
        Assert.NotEmpty(ExcelDataOperationsContract.ValidatePublicInput(op));
    }
}
