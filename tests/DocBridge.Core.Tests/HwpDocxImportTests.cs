using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;

namespace DocBridge.Core.Tests;

public sealed class HwpDocxImportTests : IDisposable
{
    private readonly string _root = Path.Combine(
        Path.GetTempPath(), "docbridge-docx-import-" + Guid.NewGuid().ToString("n"));

    public HwpDocxImportTests() => Directory.CreateDirectory(_root);

    public void Dispose()
    {
        try { Directory.Delete(_root, recursive: true); } catch { }
    }

    [Fact]
    public void No_source_keeps_the_existing_blank_document_launch_mode()
    {
        Assert.Null(HwpAdapter.ParseDocxImportRequest(new JsonObject()));
        Assert.Null(HwpAdapter.ParseDocxImportRequest(new JsonObject { ["newDocument"] = true }));
        Assert.Null(HwpAdapter.ParseDocxImportRequest(new JsonObject
        {
            ["creationMode"] = "native-hwp",
            ["newDocument"] = true,
        }));
    }

    [Fact]
    public void Docx_source_defaults_to_a_non_overwriting_hwpx_target()
    {
        var source = Path.Combine(_root, "현장 일지.docx");
        File.WriteAllBytes(source, new byte[] { 1, 2, 3 });

        var request = HwpAdapter.ParseDocxImportRequest(new JsonObject
        {
            ["creationMode"] = "docx-first",
            ["sourceFile"] = source,
            ["closeAfterImport"] = true,
            ["expectedPageCount"] = 1,
            ["expectedTableCount"] = 6,
            ["requiredText"] = new JsonArray("현장기술인 변경계", "홍길동", "홍길동"),
        });

        Assert.NotNull(request);
        Assert.Equal(Path.GetFullPath(source), request!.SourceFile);
        Assert.Equal(Path.ChangeExtension(Path.GetFullPath(source), ".hwpx"), request.OutputFile);
        Assert.True(request.CloseAfterImport);
        Assert.Equal(1, request.ExpectedPageCount);
        Assert.Equal(6, request.ExpectedTableCount);
        Assert.Equal(new[] { "현장기술인 변경계", "홍길동" }, request.RequiredText);
        Assert.Equal("OOXML", HwpAdapter.HwpAutomationFormatForPath(source));
        Assert.Equal("HWPX", HwpAdapter.HwpAutomationFormatForPath(request.OutputFile));
    }

    [Theory]
    [InlineData("result.hwpx", "HWPX")]
    [InlineData("result.hwp", "HWP")]
    public void Native_output_formats_are_explicit(string fileName, string expectedFormat)
    {
        var source = Path.Combine(_root, Guid.NewGuid().ToString("n") + ".docx");
        var output = Path.Combine(_root, fileName);
        File.WriteAllText(source, "fixture");

        var request = HwpAdapter.ParseDocxImportRequest(new JsonObject
        {
            ["sourceFile"] = source,
            ["outputFile"] = output,
        });

        Assert.NotNull(request);
        Assert.Equal(expectedFormat, HwpAdapter.HwpAutomationFormatForPath(request!.OutputFile));
    }

    [Fact]
    public void Existing_output_is_refused_before_Hwp_is_started()
    {
        var source = Path.Combine(_root, "source.docx");
        var output = Path.Combine(_root, "existing.hwpx");
        File.WriteAllText(source, "source");
        File.WriteAllText(output, "existing");

        var error = Assert.Throws<HwpAutomationException>(() =>
            HwpAdapter.ParseDocxImportRequest(new JsonObject
            {
                ["sourceFile"] = source,
                ["outputFile"] = output,
            }));

        Assert.Equal("HWP_OUTPUT_EXISTS", error.Code);
        Assert.Equal("existing", File.ReadAllText(output));
    }

    [Fact]
    public void Invalid_or_conflicting_import_requests_fail_deterministically()
    {
        var source = Path.Combine(_root, "source.docx");
        File.WriteAllText(source, "source");

        var outputOnly = Assert.Throws<HwpAutomationException>(() =>
            HwpAdapter.ParseDocxImportRequest(new JsonObject
            {
                ["outputFile"] = Path.Combine(_root, "out.hwpx"),
            }));
        Assert.Equal("HWP_DOCX_SOURCE_REQUIRED", outputOnly.Code);

        var conflict = Assert.Throws<HwpAutomationException>(() =>
            HwpAdapter.ParseDocxImportRequest(new JsonObject
            {
                ["sourceFile"] = source,
                ["newDocument"] = true,
            }));
        Assert.Equal("HWP_LAUNCH_MODE_CONFLICT", conflict.Code);

        var badOutput = Assert.Throws<HwpAutomationException>(() =>
            HwpAdapter.ParseDocxImportRequest(new JsonObject
            {
                ["sourceFile"] = source,
                ["outputFile"] = Path.Combine(_root, "out.pdf"),
            }));
        Assert.Equal("HWP_OUTPUT_FORMAT_INVALID", badOutput.Code);

        var badPageCount = Assert.Throws<HwpAutomationException>(() =>
            HwpAdapter.ParseDocxImportRequest(new JsonObject
            {
                ["sourceFile"] = source,
                ["expectedPageCount"] = 0,
            }));
        Assert.Equal("HWP_EXPECTED_PAGE_COUNT_INVALID", badPageCount.Code);

        var missingDocx = Assert.Throws<HwpAutomationException>(() =>
            HwpAdapter.ParseDocxImportRequest(new JsonObject
            {
                ["creationMode"] = "docx-first",
            }));
        Assert.Equal("HWP_DOCX_SOURCE_REQUIRED", missingDocx.Code);

        var nativeWithDocx = Assert.Throws<HwpAutomationException>(() =>
            HwpAdapter.ParseDocxImportRequest(new JsonObject
            {
                ["creationMode"] = "native-hwp",
                ["sourceFile"] = source,
            }));
        Assert.Equal("HWP_CREATION_MODE_CONFLICT", nativeWithDocx.Code);

        var unknownMode = Assert.Throws<HwpAutomationException>(() =>
            HwpAdapter.ParseDocxImportRequest(new JsonObject
            {
                ["creationMode"] = "automatic",
            }));
        Assert.Equal("HWP_CREATION_MODE_INVALID", unknownMode.Code);
    }

    [Theory]
    [InlineData(2, 1, "", true)]
    [InlineData(2, 1, "\r\n\0\u0002\u0003", true)]
    [InlineData(3, 1, "", false)]
    [InlineData(1, 1, "", false)]
    [InlineData(2, null, "", false)]
    [InlineData(2, 1, "content", false)]
    public void Trailing_blank_page_compaction_is_narrowly_gated(
        int actualPageCount,
        int? expectedPageCount,
        string trailingParagraphText,
        bool expected)
    {
        Assert.Equal(
            expected,
            HwpAdapter.ShouldCompactTrailingBlankPage(
                actualPageCount,
                expectedPageCount,
                trailingParagraphText));
    }
}
