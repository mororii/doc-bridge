using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public sealed class HwpCreationPolicyTests
{
    [Fact]
    public void New_general_document_defaults_to_docx_first()
    {
        var result = HwpCreationPolicy.Evaluate(new JsonObject
        {
            ["documentState"] = "new",
        });

        Assert.True(Json.GetBool(result, "ok"));
        Assert.Equal("docx-first", Json.GetString(result, "mode"));
        Assert.Equal(HwpCreationPolicy.PolicyVersion, Json.GetString(result, "policyVersion"));
        Assert.False(Json.GetBool(result, "wordComRequired"));
        Assert.Contains(Json.GetArr(result, "qualityGates")!, node =>
            node!.GetValue<string>().Contains("SHA-256", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("existing-hwp")]
    [InlineData("existing-hwpx")]
    public void Existing_native_document_always_uses_direct_hwp(string documentState)
    {
        var result = HwpCreationPolicy.Evaluate(new JsonObject
        {
            ["documentState"] = documentState,
        });

        Assert.True(Json.GetBool(result, "ok"));
        Assert.Equal("native-hwp", Json.GetString(result, "mode"));
    }

    [Theory]
    [InlineData("hasExistingHwpTemplate")]
    [InlineData("requiresNativeFields")]
    [InlineData("requiresHwpOnlyObjects")]
    [InlineData("requiresComplexMergedTables")]
    [InlineData("mustPreserveOriginalLayout")]
    public void Native_requirement_routes_a_new_document_to_direct_hwp(string flag)
    {
        var args = new JsonObject { ["documentState"] = "new" };
        args[flag] = true;

        var result = HwpCreationPolicy.Evaluate(args);

        Assert.True(Json.GetBool(result, "ok"));
        Assert.Equal("native-hwp", Json.GetString(result, "mode"));
        Assert.NotEmpty(Json.GetArr(result, "reasons")!);
    }

    [Fact]
    public void Missing_docx_generator_has_a_safe_native_fallback()
    {
        var result = HwpCreationPolicy.Evaluate(new JsonObject
        {
            ["documentState"] = "new",
            ["docxGeneratorAvailable"] = false,
        });

        Assert.True(Json.GetBool(result, "ok"));
        Assert.Equal("native-hwp", Json.GetString(result, "mode"));
    }

    [Fact]
    public void Unknown_document_state_is_rejected()
    {
        var result = HwpCreationPolicy.Evaluate(new JsonObject
        {
            ["documentState"] = "word",
        });

        Assert.False(Json.GetBool(result, "ok"));
    }
}
