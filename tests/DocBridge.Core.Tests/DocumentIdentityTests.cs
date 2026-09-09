using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public sealed class DocumentIdentityTests
{
    private static readonly string AbsExcel = Path.GetFullPath(@"C:\docbridge-identity\book.xlsx");
    private static readonly string AbsCad = Path.GetFullPath(@"C:\docbridge-identity\drawing.dwg");

    [Theory]
    [InlineData("untitled-18636-2", "hwp:18636:2")]
    [InlineData("untitled-18636-333780-2", "hwp:18636:333780:2")]
    public void Hwp_transient_refs_for_the_same_process_and_document_are_equivalent(
        string expected, string current)
    {
        Assert.True(DocBridgeHost.SameDocumentRef("hwp", expected, current));
        Assert.True(DocumentIdentity.TryNormalizeExecuteRef("hwp", expected, out var left, out _));
        Assert.True(DocumentIdentity.TryNormalizeExecuteRef("hwp", current, out var right, out _));
        Assert.Equal(left, right);
        Assert.Equal("hwp:18636:2", left);
    }

    [Fact]
    public void Hwp_transient_refs_for_different_documents_are_not_equivalent()
    {
        Assert.False(DocBridgeHost.SameDocumentRef("hwp", "untitled-18636-2", "hwp:18636:3"));
    }

    [Fact]
    public void Excel_and_cad_paths_normalize_full_path_and_case()
    {
        Assert.True(DocumentIdentity.TryNormalizeExecuteRef("excel", @"C:\docbridge-identity\book.xlsx", out var excel, out _));
        Assert.True(DocumentIdentity.TryNormalizeExecuteRef("excel", @"C:\docbridge-identity\.\BOOK.XLSX", out var excelDot, out _));
        Assert.Equal(AbsExcel, excel);
        Assert.True(DocBridgeHost.SameDocumentRef("excel", excel, excelDot));

        Assert.True(DocumentIdentity.TryNormalizeExecuteRef("cad", @"C:\docbridge-identity\drawing.dwg", out var cad, out _));
        Assert.Equal(AbsCad, cad);
    }

    [Theory]
    [InlineData("excel", "Book1")]
    [InlineData("excel", "Book1.xlsx")]
    [InlineData("cad", "Drawing1")]
    [InlineData("cad", "unsaved-Drawing1")]
    [InlineData("gstarcad", "Drawing1")]
    [InlineData("hwp", "새 문서1")]
    public void Bare_unsaved_names_are_not_stable_execute_refs(string app, string raw)
    {
        Assert.False(DocumentIdentity.TryNormalizeExecuteRef(app, raw, out _, out var error));
        Assert.Contains("dryRun=true", error, StringComparison.Ordinal);
        Assert.Contains("confirmToken", error, StringComparison.Ordinal);
        Assert.Contains("will not save", error, StringComparison.Ordinal);
    }

    [Fact]
    public void Adapter_instance_bound_refs_are_stable_without_saving()
    {
        Assert.True(DocumentIdentity.TryNormalizeExecuteRef("excel", "excel-instance:7:1", out var excel, out _));
        Assert.Equal("excel-instance:7:1", excel);
        Assert.True(DocumentIdentity.TryNormalizeExecuteRef("cad", "cad-instance:hwnd:3", out var cad, out _));
        Assert.Equal("cad-instance:hwnd:3", cad);
    }

    [Fact]
    public void Conflicting_explicit_per_op_targets_are_denied()
    {
        var one = Path.GetFullPath(@"C:\data\one.xlsx");
        var two = Path.GetFullPath(@"C:\data\two.xlsx");
        var bind = DocumentTargetBinder.BindForExecute("excel", new[]
        {
            new JsonObject
            {
                ["op"] = "set_values",
                ["range"] = "Sheet1!A1",
                ["values"] = new JsonArray(new JsonArray("1")),
                ["targetWorkbook"] = one,
            },
            new JsonObject
            {
                ["op"] = "set_values",
                ["range"] = "Sheet1!A2",
                ["values"] = new JsonArray(new JsonArray("2")),
                ["targetWorkbook"] = two,
            },
        }, one);
        Assert.False(bind.Ok);
        Assert.Contains(bind.Errors, e => e.Contains("mixed explicit", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Clone_bound_ops_stamp_expected_and_do_not_mutate_source()
    {
        var source = new JsonObject
        {
            ["op"] = "set_values",
            ["range"] = "Sheet1!A1",
            ["values"] = new JsonArray(new JsonArray("1")),
        };
        var bound = DocumentTargetBinder.CloneBoundOps("excel", new[] { source }, AbsExcel);
        Assert.Equal(AbsExcel, Json.GetString(bound[0], "targetWorkbook"));
        Assert.False(source.ContainsKey("targetWorkbook"));
        Assert.NotEqual(ConfirmTokenService.HashOps(new[] { source }), ConfirmTokenService.HashOps(bound));

        var hwpFile = new JsonObject { ["op"] = "append_text", ["text"] = "x", ["file"] = AbsExcel };
        var hwpBound = DocumentTargetBinder.CloneBoundOps("hwp", new[] { hwpFile }, AbsExcel);
        Assert.Equal(AbsExcel, Json.GetString(hwpBound[0], "file"));
        Assert.False(hwpBound[0].ContainsKey("documentRef"));

        var cad = new JsonObject { ["op"] = "set_layer_visibility", ["layer"] = "PLAN", ["visible"] = false };
        var cadBound = DocumentTargetBinder.CloneBoundOps("cad", new[] { cad }, AbsCad);
        Assert.Equal(AbsCad, Json.GetString(cadBound[0], "document"));
        Assert.False(cad.ContainsKey("document"));
    }
}
