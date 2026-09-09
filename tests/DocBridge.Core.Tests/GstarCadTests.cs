using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class GstarCadTests
{
    [Fact]
    public void Router_keeps_separate_product_instances_and_capabilities()
    {
        using var router = new SessionRouter();
        var auto = router.Get("cad");
        var gstar = router.Get("gstarcad");
        Assert.NotSame(auto, gstar);
        Assert.Equal("cad", auto.App);
        Assert.Equal("gstarcad", gstar.App);
        Assert.Equal("Gcad.Application", (string?)gstar.GetCapabilities()["progId"]);
        Assert.False((bool)gstar.GetCapabilities()["crossProductFallback"]!);
        Assert.DoesNotContain("plot_pdf", ((JsonArray)gstar.GetCapabilities()["writeOps"]!).Select(n => (string?)n));
        Assert.Contains("plot_pdf", ((JsonArray)auto.GetCapabilities()["writeOps"]!).Select(n => (string?)n));
    }

    [Theory]
    [InlineData("cad", CadProduct.AutoCad)]
    [InlineData("gstarcad", CadProduct.GstarCad)]
    public void Context_and_pagination_actions_keep_product_identity(string key, CadProduct product)
    {
        var app = CadContextOptimizationTests.FakeCadApp.Create(8, 3);
        using var adapter = new CadAdapter(() => app, product);
        Assert.Equal(key, adapter.GetActiveContext().App);
        var context = adapter.GetActiveContext();
        Assert.All((JsonArray)context.Summary["nextActions"]!, n => Assert.Equal(key + "_query_entities", (string?)n?["tool"]));
        foreach (var scope in new[] { "layers", "entities" })
        {
            var result = adapter.Read(new JsonObject { ["scope"] = scope, ["limit"] = 1 });
            Assert.True(Json.GetBool(result, "ok"), result.ToJsonString());
            Assert.Equal(key, (string?)result["app"]);
            Assert.All((JsonArray)result["nextActions"]!, n => Assert.Equal(key + "_query_entities", (string?)n?["tool"]));
        }
    }

    [Theory]
    [InlineData("{\"op\":\"copy_entities_between_documents\"}")]
    [InlineData("{\"op\":\"plot_pdf\",\"output\":\"test.pdf\"}")]
    [InlineData("{\"op\":\"draw_entities\",\"entities\":[{\"type\":\"hatch\"}]}")]
    [InlineData("{\"op\":\"draw_entities\",\"entities\":[{\"type\":\"text\",\"color\":{\"rgb\":[1,2,3]}}]}")]
    public void Unsupported_tail_is_rejected_before_any_com_or_allowed_prefix(string tail)
    {
        var calls = 0;
        using var adapter = new CadAdapter(() => { calls++; return null; }, CadProduct.GstarCad);
        var ops = new[] { new JsonObject { ["op"] = "regen_document" }, JsonNode.Parse(tail)!.AsObject() };
        Assert.NotEmpty(adapter.Preview(ops).Errors);
        Assert.False(adapter.Apply(ops, "unused").Ok);
        Assert.Equal(0, calls);
    }

    [Fact]
    public void Disconnected_status_does_not_report_connected_or_reuse_dead_reference()
    {
        var calls = 0;
        using var adapter = new CadAdapter(() => { calls++; return new DisconnectedApp(); }, CadProduct.GstarCad);
        Assert.False(adapter.GetStatus().Connected);
        Assert.False(adapter.GetStatus().Connected);
        Assert.Equal(2, calls);
    }

    [Theory]
    [InlineData(CadProduct.AutoCad, "cad")]
    [InlineData(CadProduct.GstarCad, "gstarcad")]
    public void Dxf_fallback_preserves_requested_product_without_creating_an_instance(CadProduct product, string key)
    {
        using var home = new TestHome();
        var file = Path.Combine(home.Dir, "probe.dxf");
        File.WriteAllText(file, "0\nSECTION\n2\nENTITIES\n0\nTEXT\n5\nA1\n8\n0\n1\nPROBE\n0\nENDSEC\n0\nEOF\n");
        using var adapter = new CadAdapter(() => null, product);
        var result = adapter.Read(new JsonObject { ["file"] = file });
        Assert.True(Json.GetBool(result, "ok"), result.ToJsonString());
        Assert.Equal(key, (string?)result["app"]);
    }

    [Fact]
    public void Gstar_policy_preserves_safety_and_does_not_widen_autocad_policy()
    {
        var policy = new PolicyEngine();
        Assert.Equal(OpClass.Allowed, policy.ClassifyOp("gstarcad", "set_text_value"));
        Assert.Equal(OpClass.HighRisk, policy.ClassifyOp("gstarcad", "save_document"));
        Assert.Equal(OpClass.Forbidden, policy.ClassifyOp("gstarcad", "run_script_template"));
        Assert.Equal(OpClass.HighRisk, policy.ClassifyOp("cad", "run_script_template"));
    }

    [Fact]
    public void Identical_document_and_ops_still_require_separate_product_tokens_and_snapshots()
    {
        using var home = new TestHome();
        var snapshots = new SnapshotService(home.Options);
        var auto = snapshots.Create("cad", "probe", "same.dwg", (_, _) => { });
        var gstar = snapshots.Create("gstarcad", "probe", "same.dwg", (_, _) => { });
        Assert.NotEqual(auto.Dir, gstar.Dir);
        Assert.Equal("cad", auto.App);
        Assert.Equal("gstarcad", gstar.App);
        var tokens = new ConfirmTokenService(home.Options);
        var hash = ConfirmTokenService.HashOps(new[] { new JsonObject { ["op"] = "regen_document" } });
        var token = tokens.Create("apply:gstarcad", hash, gstar.SnapshotId).Token;
        var wrong = tokens.Validate(token, "apply:cad", hash);
        Assert.False(wrong.Ok);
        Assert.True(tokens.Validate(token, "apply:gstarcad", hash).Ok);
    }

    public sealed class DisconnectedApp
    {
        public string Version => throw new COMException("disconnected", unchecked((int)0x80010108));
    }
}
