using DocBridge.Core.Adapters;

namespace DocBridge.Core.Tests;

public sealed class HwpFindReplaceReadbackTests
{
    [Fact]
    public void False_action_return_is_accepted_when_readback_proves_replacement()
    {
        Assert.True(HwpAdapter.FindReplaceReadbackVerified(
            "첨부 문서", "[[MOVED]] 문서", "첨부", "[[MOVED]]", 1, 0));
    }

    [Fact]
    public void Readback_rejects_when_original_text_remains()
    {
        Assert.False(HwpAdapter.FindReplaceReadbackVerified(
            "첨부 문서", "첨부 문서", "첨부", "[[MOVED]]", 1, 1));
    }

    [Fact]
    public void Replacement_containing_find_requires_an_increased_replacement_count()
    {
        Assert.True(HwpAdapter.FindReplaceReadbackVerified(
            "제목", "제목(수정됨)", "제목", "제목(수정됨)", 1, 1));
    }
}
