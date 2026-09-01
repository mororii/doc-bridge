using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public sealed class DocumentIdentityTests
{
    [Theory]
    [InlineData("untitled-18636-2", "hwp:18636:2")]
    [InlineData("untitled-18636-333780-2", "hwp:18636:333780:2")]
    public void Hwp_transient_refs_for_the_same_process_and_document_are_equivalent(
        string expected, string current)
    {
        Assert.True(DocBridgeHost.SameDocumentRef("hwp", expected, current));
    }

    [Fact]
    public void Hwp_transient_refs_for_different_documents_are_not_equivalent()
    {
        Assert.False(DocBridgeHost.SameDocumentRef("hwp", "untitled-18636-2", "hwp:18636:3"));
    }
}
