using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelLifecycleOwnershipTests
{
    [Fact]
    public void Auto_quit_only_empty_owned_application()
    {
        Assert.True(ExcelApplicationQuitContract.MayAutoQuit(ownedInstance: true, workbookCount: 0));
        Assert.False(ExcelApplicationQuitContract.MayAutoQuit(ownedInstance: false, workbookCount: 0));
        Assert.False(ExcelApplicationQuitContract.MayAutoQuit(ownedInstance: true, workbookCount: null));
    }

    [Fact]
    public void Saved_preexisting_book_does_not_authorize_quit()
    {
        Assert.False(ExcelApplicationQuitContract.MayAutoQuit(ownedInstance: true, workbookCount: 1));
        Assert.Contains("Saved does not authorize Quit",
            ExcelApplicationQuitContract.DetachWithoutQuitMessage(1), StringComparison.Ordinal);
    }

    [Fact]
    public void Saved_new_output_does_not_authorize_quit()
    {
        Assert.False(ExcelApplicationQuitContract.MayAutoQuit(ownedInstance: true, workbookCount: 2));
        Assert.Contains("2 open workbook",
            ExcelApplicationQuitContract.DetachWithoutQuitMessage(2), StringComparison.Ordinal);
    }

    [Fact]
    public void Unsaved_books_do_not_authorize_quit()
    {
        Assert.False(ExcelApplicationQuitContract.MayAutoQuit(ownedInstance: true, workbookCount: 1));
    }

    [Fact]
    public void Zero_restore_checks_are_unproven_not_complete_or_verified()
    {
        Assert.False(ExcelLifecycleRestoreContract.ZeroChecksIsComplete);
        Assert.False(ExcelLifecycleRestoreContract.IsVerified(0, 0));
        Assert.Equal(ExcelLifecycleRestoreContract.CompletenessUnproven,
            ExcelLifecycleRestoreContract.Completeness(0, 0));
        Assert.Contains("unproven", ExcelLifecycleRestoreContract.UnprovenMessage, StringComparison.Ordinal);
    }

    [Fact]
    public void Lost_preexisting_bystander_is_incomplete_and_not_recreated()
    {
        var before = new[]
        {
            Book("e2-cost-estimate.xlsx", @"C:\out\e2-cost-estimate.xlsx"),
            Book("e6-weekly.xlsx", @"C:\out\e6-weekly.xlsx"),
        };
        var after = new[]
        {
            Book("e6-weekly.xlsx", @"C:\out\e6-weekly.xlsx"),
        };
        var owned = new JsonArray();
        var lost = ExcelLifecycleRestoreContract.LostPreexisting(before, after, owned);
        Assert.Contains(lost, name => name.Contains("e2-cost-estimate", StringComparison.OrdinalIgnoreCase));
        Assert.Equal(1, ExcelLifecycleRestoreContract.VerifiedPreexistingCount(before, after, owned));
        Assert.False(ExcelLifecycleRestoreContract.IsVerified(1, lost.Count));
        Assert.Equal(ExcelLifecycleRestoreContract.CompletenessIncomplete,
            ExcelLifecycleRestoreContract.Completeness(1, lost.Count));
        Assert.False(ExcelAuthoringSchema.ShouldCloseOnLifecycleRestore(
            "e2-cost-estimate.xlsx", @"C:\out\e2-cost-estimate.xlsx", owned));
    }

    [Fact]
    public void Owned_close_still_allowed_and_does_not_count_as_lost_bystander()
    {
        var owned = new JsonArray
        {
            Book("e6-weekly.xlsx", @"C:\out\e6-weekly.xlsx"),
        };
        var before = new[]
        {
            Book("e2-cost-estimate.xlsx", @"C:\out\e2-cost-estimate.xlsx"),
            Book("e6-weekly.xlsx", @"C:\out\e6-weekly.xlsx"),
        };
        var afterClose = new[]
        {
            Book("e2-cost-estimate.xlsx", @"C:\out\e2-cost-estimate.xlsx"),
        };
        Assert.True(ExcelAuthoringSchema.ShouldCloseOnLifecycleRestore(
            "e6-weekly.xlsx", @"C:\out\e6-weekly.xlsx", owned));
        Assert.Empty(ExcelLifecycleRestoreContract.LostPreexisting(before, afterClose, owned));
        Assert.Equal(1, ExcelLifecycleRestoreContract.VerifiedPreexistingCount(before, afterClose, owned));
        Assert.True(ExcelLifecycleRestoreContract.IsVerified(1, 0));
        Assert.Equal(ExcelLifecycleRestoreContract.CompletenessComplete,
            ExcelLifecycleRestoreContract.Completeness(1, 0));
    }

    [Fact]
    public void Discovery_must_preserve_borrowed_application_alias()
    {
        var live = new object();
        Assert.True(ExcelComAliasContract.MustPreserveLiveAlias(live, live));
        Assert.False(ExcelComAliasContract.MustPreserveLiveAlias(new object(), live));
        Assert.False(ExcelComAliasContract.MustPreserveLiveAlias(null, live));
        Assert.False(ExcelComAliasContract.DiscoveryFinalReleasesBorrowedChild);
        Assert.False(ExcelComAliasContract.DiscoveryFinalReleasesApplicationAlias);
    }

    [Fact]
    public void Busy_or_rejected_visible_instance_is_unknown_not_empty()
    {
        var rejected = new System.Runtime.InteropServices.COMException(
            "rejected", unchecked((int)0x80010001));
        Assert.True(ExcelDiscoveryCoverageContract.IsBusyOrRejected(rejected));
        Assert.Equal(ExcelDiscoveryCoverageContract.RejectedOrBusy,
            ExcelDiscoveryCoverageContract.DescribeFailure(rejected));
        Assert.Equal(ExcelDiscoveryCoverageContract.CoverageUnknown,
            ExcelDiscoveryCoverageContract.Coverage(workbookCount: null, readFailed: true));
        Assert.False(ExcelDiscoveryCoverageContract.IsProvenEmpty(null, readFailed: true));
        Assert.False(ExcelDiscoveryCoverageContract.SkipWhenRequireWorkbook(null, readFailed: true));
        Assert.True(ExcelDiscoveryCoverageContract.DiscoveryScore(null, readFailed: true) >
                    ExcelDiscoveryCoverageContract.DiscoveryScore(0, readFailed: false));
        Assert.True(ExcelDiscoveryCoverageContract.IsProvenEmpty(0, readFailed: false));
        Assert.True(ExcelDiscoveryCoverageContract.ShouldRetainBusyCandidate(1));
        Assert.False(ExcelDiscoveryCoverageContract.ShouldRetainBusyCandidate(1001));
    }

    [Fact]
    public void Unknown_same_instance_coverage_does_not_prove_bystander_lost()
    {
        var before = new[]
        {
            new JsonObject
            {
                ["name"] = "e2-cost-estimate.xlsx",
                ["fullName"] = @"C:\out\e2-cost-estimate.xlsx",
                ["processId"] = 46688,
                ["excelHwnd"] = 28120064,
            },
        };
        var after = new[]
        {
            ExcelDiscoveryCoverageContract.UnknownInstance(28120064, 46688,
                ExcelDiscoveryCoverageContract.RejectedOrBusy),
        };
        Assert.True(ExcelDiscoveryCoverageContract.SameInstanceUnknownCoverage(before[0], after));
        Assert.Empty(ExcelLifecycleRestoreContract.LostPreexisting(before, after, new JsonArray()));
        Assert.Equal(0, ExcelLifecycleRestoreContract.VerifiedPreexistingCount(before, after, new JsonArray()));
        Assert.Equal(ExcelLifecycleRestoreContract.CompletenessUnproven,
            ExcelLifecycleRestoreContract.Completeness(0, 0));
    }

    [Fact]
    public void Page_break_A20_defaults_to_row_20_and_column_J_is_vertical()
    {
        Assert.Equal(ExcelPageBreakContract.BreakAxis.Rows,
            ExcelPageBreakContract.ResolveAxis(new JsonObject { ["range"] = "A20" }));
        Assert.Equal(20, ExcelPageBreakContract.BreakIndex("A20", ExcelPageBreakContract.BreakAxis.Rows));
        Assert.Equal(ExcelPageBreakContract.BreakAxis.Columns,
            ExcelPageBreakContract.ResolveAxis(new JsonObject { ["orientation"] = "column" }));
        Assert.Equal(10, ExcelPageBreakContract.BreakIndex("J1", ExcelPageBreakContract.BreakAxis.Columns));

        var missing = new JsonObject { ["horizontal"] = new JsonArray(), ["vertical"] = new JsonArray() };
        var why = ExcelPageBreakContract.DescribeMissing("A20", ExcelPageBreakContract.BreakAxis.Rows, 20, missing);
        Assert.Contains("A20", why, StringComparison.Ordinal);
        Assert.Contains("expected manual at 20", why, StringComparison.Ordinal);
        Assert.False(ExcelPageBreakContract.HasManualBreak(missing, ExcelPageBreakContract.BreakAxis.Rows, 20));
        Assert.True(ExcelPageBreakContract.HasManualBreak(
            new JsonObject { ["horizontal"] = new JsonArray(20) },
            ExcelPageBreakContract.BreakAxis.Rows, 20));
    }

    private static JsonObject Book(string name, string fullName) =>
        new()
        {
            ["name"] = name,
            ["fullName"] = fullName,
        };
}
