namespace DocBridge.Core.Tests;

public sealed class HwpE2EOwnershipTests
{
    [Fact]
    public void Window_process_id_wins_over_newcomer_scan()
    {
        var pid = HwpE2EOwnershipPolicy.ResolveProcessId(4242, [100], [100, 4242, 9000]);
        Assert.Equal(4242, pid);
    }

    [Fact]
    public void Exactly_one_new_pid_proves_process_when_window_id_is_missing()
    {
        var pid = HwpE2EOwnershipPolicy.ResolveProcessId(0, [100], [100, 777]);
        Assert.Equal(777, pid);
    }

    [Theory]
    [InlineData(new int[0], new int[0])]
    [InlineData(new[] { 100 }, new[] { 100 })]
    [InlineData(new[] { 100 }, new[] { 100, 200, 300 })]
    public void Unknown_or_ambiguous_process_ids_fail_closed(int[] existing, int[] current)
    {
        Assert.Equal(0, HwpE2EOwnershipPolicy.ResolveProcessId(0, existing, current));
    }

    [Fact]
    public void New_process_with_successful_file_new_may_seed_the_empty_tab()
    {
        var after = HwpE2EDocumentSnapshot.Create(["default", "fresh"], "fresh", true);
        Assert.True(HwpE2EOwnershipPolicy.TryProve(
            [], 501, fileNewSucceeded: true,
            HwpE2EDocumentSnapshot.Create(["default"], "default", true),
            after, out var claim, out var error), error);
        Assert.Equal(HwpE2EOwnershipMode.ExclusiveNewProcess, claim.Mode);
        Assert.True(HwpE2EOwnershipPolicy.AllowsDocumentMutation(claim));
        Assert.False(HwpE2EOwnershipPolicy.AllowsClearOrSelectAllDelete(claim));
        Assert.Contains("fresh", claim.OwnedDocumentIds);
        Assert.Contains("default", claim.OwnedDocumentIds);
        Assert.Equal("fresh", claim.SeededDocumentId);
    }

    [Fact]
    public void New_process_may_use_the_single_empty_default_tab_when_file_new_fails()
    {
        var after = HwpE2EDocumentSnapshot.Create(["only"], "only", true);
        Assert.True(HwpE2EOwnershipPolicy.TryProve(
            [], 502, fileNewSucceeded: false, before: null, after, out var claim, out var error), error);
        Assert.Equal(HwpE2EOwnershipMode.ExclusiveNewProcess, claim.Mode);
        Assert.Equal(new[] { "only" }, claim.OwnedDocumentIds);
        Assert.Equal("only", claim.SeededDocumentId);
    }

    [Theory]
    [InlineData(false, "only", false)]
    [InlineData(false, "only", null)]
    [InlineData(false, "", true)]
    [InlineData(true, "", true)]
    [InlineData(true, "only", null)]
    public void New_process_fails_closed_when_file_new_fails_or_tab_identity_is_unusable(
        bool fileNewSucceeded, string activeId, bool? empty)
    {
        var ids = string.IsNullOrEmpty(activeId) ? Array.Empty<string>() : new[] { activeId };
        var after = HwpE2EDocumentSnapshot.Create(ids, string.IsNullOrEmpty(activeId) ? null : activeId, empty);
        Assert.False(HwpE2EOwnershipPolicy.TryProve(
            [], 503, fileNewSucceeded, before: null, after, out var claim, out var error));
        Assert.Equal(HwpE2EOwnershipMode.None, claim.Mode);
        Assert.False(HwpE2EOwnershipPolicy.AllowsDocumentMutation(claim));
        Assert.Contains("refused to mutate", error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void New_process_with_file_new_success_refuses_a_non_empty_tab()
    {
        var after = HwpE2EDocumentSnapshot.Create(["fresh"], "fresh", false);
        Assert.False(HwpE2EOwnershipPolicy.TryProve(
            [], 504, fileNewSucceeded: true,
            HwpE2EDocumentSnapshot.Empty, after, out _, out var error));
        Assert.Contains("not proven empty", error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Unknown_process_id_fails_closed_before_any_mutation_decision()
    {
        Assert.False(HwpE2EOwnershipPolicy.TryProve(
            [10], 0, fileNewSucceeded: true,
            HwpE2EDocumentSnapshot.Empty,
            HwpE2EDocumentSnapshot.Create(["a"], "a", true),
            out var claim, out var error));
        Assert.False(claim.CanMutate);
        Assert.Contains("unknown", error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Existing_process_requires_exactly_one_new_empty_active_tab()
    {
        var before = HwpE2EDocumentSnapshot.Create(["user-doc"], "user-doc", false);
        var after = HwpE2EDocumentSnapshot.Create(["user-doc", "e2e-tab"], "e2e-tab", true);
        Assert.True(HwpE2EOwnershipPolicy.TryProve(
            [900], 900, fileNewSucceeded: true, before, after, out var claim, out var error), error);
        Assert.Equal(HwpE2EOwnershipMode.IsolatedTab, claim.Mode);
        Assert.Equal(new[] { "e2e-tab" }, claim.OwnedDocumentIds);
        Assert.Equal("e2e-tab", claim.SeededDocumentId);
        Assert.Equal(new[] { "user-doc" }, claim.PreexistingDocumentIds);
        Assert.True(HwpE2EOwnershipPolicy.AllowsDocumentMutation(claim));
        Assert.False(HwpE2EOwnershipPolicy.AllowsClearOrSelectAllDelete(claim));
    }

    [Theory]
    [InlineData(false, "user-doc", "user-doc", "e2e-tab", 2, true)]
    [InlineData(true, "user-doc", "user-doc", "user-doc", 1, true)]
    [InlineData(true, "user-doc", "user-doc", "e2e-tab", 3, true)]
    [InlineData(true, "user-doc", "other", "e2e-tab", 2, true)]
    [InlineData(true, "user-doc", "e2e-tab", "e2e-tab", 2, false)]
    public void Existing_process_fails_closed_without_isolated_tab_proof(
        bool fileNewSucceeded,
        string beforeId,
        string activeAfter,
        string extraAfter,
        int afterCount,
        bool empty)
    {
        var before = HwpE2EDocumentSnapshot.Create([beforeId], beforeId, false);
        var afterIds = afterCount switch
        {
            1 => new[] { activeAfter },
            2 => new[] { beforeId, extraAfter },
            _ => new[] { beforeId, extraAfter, "third" },
        };
        var after = new HwpE2EDocumentSnapshot(afterIds, activeAfter, empty, afterCount);
        Assert.False(HwpE2EOwnershipPolicy.TryProve(
            [900], 900, fileNewSucceeded, before, after, out var claim, out _));
        Assert.False(claim.CanMutate);
    }

    [Fact]
    public void Existing_process_without_a_before_inventory_fails_closed()
    {
        var after = HwpE2EDocumentSnapshot.Create(["e2e-tab"], "e2e-tab", true);
        Assert.False(HwpE2EOwnershipPolicy.TryProve(
            [900], 900, fileNewSucceeded: true, before: null, after, out var claim, out var error));
        Assert.False(claim.CanMutate);
        Assert.Contains("preexisting", error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Created_tab_identity_requires_exactly_one_new_id()
    {
        var before = HwpE2EDocumentSnapshot.Create(["a"], "a", false);
        Assert.True(HwpE2EOwnershipPolicy.TryIdentifyCreatedTab(
            before, HwpE2EDocumentSnapshot.Create(["a", "b"], "b", true), out var created));
        Assert.Equal("b", created);
        Assert.False(HwpE2EOwnershipPolicy.TryIdentifyCreatedTab(
            before, HwpE2EDocumentSnapshot.Create(["a"], "a", true), out _));
        Assert.False(HwpE2EOwnershipPolicy.TryIdentifyCreatedTab(
            before, HwpE2EDocumentSnapshot.Create(["a", "b", "c"], "b", true), out _));
    }

    [Fact]
    public void Cleanup_of_unproven_or_unreadable_state_retains_and_never_kills()
    {
        var unproven = HwpE2EOwnershipPolicy.PlanCleanup(null, ["any"]);
        Assert.True(unproven.RetainUnknownState);
        Assert.Empty(unproven.DocumentIdsToClose);
        Assert.False(unproven.AllowCloseAllWindows);
        Assert.False(unproven.AllowProcessKill);

        var claim = new HwpE2EOwnershipClaim(
            HwpE2EOwnershipMode.ExclusiveNewProcess, 1, ["owned"], [], "test", "owned");
        var unread = HwpE2EOwnershipPolicy.PlanCleanup(claim, currentDocumentIds: null);
        Assert.True(unread.RetainUnknownState);
        Assert.Empty(unread.DocumentIdsToClose);
        Assert.False(unread.AllowProcessKill);
        Assert.False(unread.AllowCloseAllWindows);
    }

    [Fact]
    public void Cleanup_closes_only_explicit_owned_ids_and_preserves_a_user_opened_later_tab()
    {
        var exclusive = new HwpE2EOwnershipClaim(
            HwpE2EOwnershipMode.ExclusiveNewProcess,
            12,
            ["owned", "default"],
            [],
            "exclusive",
            "owned");
        var exclusivePlan = HwpE2EOwnershipPolicy.PlanCleanup(
            exclusive, ["owned", "default", "later-launch", "user-opened-later"]);
        Assert.False(exclusivePlan.RetainUnknownState);
        Assert.Equal(new[] { "owned", "default" }, exclusivePlan.DocumentIdsToClose);
        Assert.Equal(new[] { "later-launch", "user-opened-later" }, exclusivePlan.RetainedUnknownDocumentIds);
        Assert.DoesNotContain("user-opened-later", exclusivePlan.DocumentIdsToClose);
        Assert.False(exclusivePlan.AllowCloseAllWindows);
        Assert.False(exclusivePlan.AllowProcessKill);
        Assert.False(HwpE2EOwnershipPolicy.AllowsClearOrSelectAllDelete(exclusive));

        var isolated = new HwpE2EOwnershipClaim(
            HwpE2EOwnershipMode.IsolatedTab,
            12,
            ["e2e-tab"],
            ["user-doc"],
            "isolated",
            "e2e-tab");
        var isolatedPlan = HwpE2EOwnershipPolicy.PlanCleanup(
            isolated, ["user-doc", "e2e-tab", "launch-tab", "user-opened-later"]);
        Assert.Equal(new[] { "e2e-tab" }, isolatedPlan.DocumentIdsToClose);
        Assert.Contains("user-doc", isolatedPlan.RetainedUnknownDocumentIds);
        Assert.Contains("launch-tab", isolatedPlan.RetainedUnknownDocumentIds);
        Assert.Contains("user-opened-later", isolatedPlan.RetainedUnknownDocumentIds);
        Assert.DoesNotContain("user-opened-later", isolatedPlan.DocumentIdsToClose);
        Assert.False(isolatedPlan.AllowCloseAllWindows);
        Assert.False(isolatedPlan.AllowProcessKill);
    }

    [Fact]
    public void Inventory_validation_fails_on_empty_duplicate_or_count_mismatch()
    {
        Assert.True(HwpE2EOwnershipPolicy.TryValidateInventory(
            2, ["a", "b"], "b", true, out var ok, out _));
        Assert.Equal(2, ok.DocumentCount);
        Assert.Equal("b", ok.ActiveDocumentId);

        Assert.False(HwpE2EOwnershipPolicy.TryValidateInventory(
            2, ["a"], "a", true, out _, out var countError));
        Assert.Contains("differ", countError, StringComparison.OrdinalIgnoreCase);

        Assert.False(HwpE2EOwnershipPolicy.TryValidateInventory(
            2, ["a", ""], "a", true, out _, out var emptyError));
        Assert.Contains("empty", emptyError, StringComparison.OrdinalIgnoreCase);

        Assert.False(HwpE2EOwnershipPolicy.TryValidateInventory(
            2, ["a", "a"], "a", true, out _, out var dupError));
        Assert.Contains("duplicate", dupError, StringComparison.OrdinalIgnoreCase);

        Assert.False(HwpE2EOwnershipPolicy.TryValidateInventory(
            1, [null], "a", true, out _, out var failedRead));
        Assert.Contains("could not be read", failedRead, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Launch_capture_records_only_the_exact_before_after_delta()
    {
        var before = HwpE2EDocumentSnapshot.Create(["owned", "default"], "owned", false);
        var after = HwpE2EDocumentSnapshot.Create(["owned", "default", "launch-tab"], "launch-tab", true);
        Assert.True(HwpE2EOwnershipPolicy.TryCaptureCreatedDocumentIds(before, after, out var created));
        Assert.Equal(new[] { "launch-tab" }, created);

        var claim = new HwpE2EOwnershipClaim(
            HwpE2EOwnershipMode.ExclusiveNewProcess, 1, ["owned", "default"], [], "exclusive", "owned");
        var updated = HwpE2EOwnershipPolicy.WithAdditionalOwnedDocuments(claim, created);
        var plan = HwpE2EOwnershipPolicy.PlanCleanup(
            updated, ["owned", "default", "launch-tab", "user-opened-later"]);
        Assert.Equal(new[] { "owned", "default", "launch-tab" }, plan.DocumentIdsToClose);
        Assert.Equal(new[] { "user-opened-later" }, plan.RetainedUnknownDocumentIds);

        var staleAfter = new HwpE2EDocumentSnapshot(["owned", "default", "launch-tab"], "launch-tab", true, 4);
        Assert.False(HwpE2EOwnershipPolicy.TryCaptureCreatedDocumentIds(before, staleAfter, out _));
    }

    [Fact]
    public void Artifacts_are_retained_only_when_the_directory_env_is_explicit()
    {
        Assert.False(HwpE2EOwnershipPolicy.ShouldRetainArtifacts(null));
        Assert.False(HwpE2EOwnershipPolicy.ShouldRetainArtifacts(""));
        Assert.True(HwpE2EOwnershipPolicy.ShouldRetainArtifacts(@"C:\tmp\hwp-artifacts"));
    }

    [Fact]
    public void Owned_hwp_artifact_paths_are_unique_and_include_the_seeded_document()
    {
        var first = HwpE2EOwnershipPolicy.UniqueOwnedHwpArtifactPath(@"C:\tmp\hwp-artifacts", "fresh-tab");
        var second = HwpE2EOwnershipPolicy.UniqueOwnedHwpArtifactPath(@"C:\tmp\hwp-artifacts", "fresh-tab");
        Assert.NotEqual(first, second);
        Assert.Contains("fresh-tab", first, StringComparison.Ordinal);
        Assert.DoesNotContain("hwp-e2e-owned.hwp", new[] { Path.GetFileName(first), Path.GetFileName(second) });
    }
}
