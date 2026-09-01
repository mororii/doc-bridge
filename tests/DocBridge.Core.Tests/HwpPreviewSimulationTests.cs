using DocBridge.Core.Adapters;

namespace DocBridge.Core.Tests;

public sealed class HwpPreviewSimulationTests
{
    [Fact]
    public void Relative_insert_readback_accepts_equivalent_minus_and_compatibility_unit()
    {
        const string inserted = "관경 − 300㎜ · 굴착 깊이 1,200㎜";
        const string hwpReadback = "관경 &#8722; 300mm · 굴착 깊이 1,200mm";

        Assert.True(HwpAdapter.HwpReadbackContainsEquivalent(hwpReadback, inserted));
        Assert.False(HwpAdapter.HwpReadbackContainsEquivalent(
            "관경 - 250mm · 굴착 깊이 1,200mm", inserted));
    }

    [Fact]
    public void Preview_text_simulation_applies_operations_in_batch_order()
    {
        var state = HwpAdapter.SimulatePreviewAppend("제목", "기준 문단", true);
        state = HwpAdapter.SimulatePreviewRelativeInsert(
            state, "기준 문단", "앞 작업 결과를 참조한 후속 문단", 1,
            matchCase: true, before: false, mode: "paragraph");

        Assert.Equal("제목\n기준 문단\n앞 작업 결과를 참조한 후속 문단", state);
        Assert.Contains("앞 작업 결과", state);
    }

    [Fact]
    public void Preview_inline_simulation_inserts_at_requested_occurrence()
    {
        var state = HwpAdapter.SimulatePreviewRelativeInsert(
            "기준/기준", "기준", "[삽입]", 2,
            matchCase: true, before: true, mode: "inline");

        Assert.Equal("기준/[삽입]기준", state);
    }
}
