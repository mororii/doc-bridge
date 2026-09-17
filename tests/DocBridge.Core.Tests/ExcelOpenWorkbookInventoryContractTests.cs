using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelOpenWorkbookInventoryContractTests
{
    [Fact]
    public void Lost_blank_saved_book1_is_aggregate_failure_not_waived()
    {
        Assert.True(ExcelOpenWorkbookInventoryContract.IsBlankUntitledName("Book1"));
        Assert.True(ExcelOpenWorkbookInventoryContract.IsBlankUntitledName("통합 문서1"));
        var before = new[]
        {
            Book(pid: 1001, hwnd: 4722654, name: "통합 문서1", fullName: "통합 문서1", saved: true),
        };
        var after = new[]
        {
            Book(pid: 1001, hwnd: 5247156, name: "e2-cost-estimate.xlsx",
                fullName: @"C:\tmp\e2-cost-estimate.xlsx", saved: true),
        };
        var opened = @"C:\tmp\e2-cost-estimate.xlsx";
        var lost = ExcelOpenWorkbookInventoryContract.LostBystanders(before, after, opened);
        Assert.Equal(new[] { "통합 문서1" }, lost);
        Assert.True(ExcelOpenWorkbookInventoryContract.PreserveAggregateFailure(lost));
        var message = ExcelOpenWorkbookInventoryContract.FailureMessage(lost, 1001, [1001, 2204]);
        Assert.Contains("통합 문서1", message, StringComparison.Ordinal);
        Assert.Contains("target-owner PID 1001", message, StringComparison.Ordinal);
        Assert.Contains("all-instance PIDs", message, StringComparison.Ordinal);
        Assert.Contains("not waived and not recreated", message, StringComparison.Ordinal);
    }

    [Fact]
    public void Opened_target_is_not_a_lost_bystander()
    {
        var path = @"C:\tmp\e2-cost-estimate.xlsx";
        var before = new[]
        {
            Book(pid: 8, hwnd: 1, name: "e2-cost-estimate.xlsx", fullName: path, saved: true),
            Book(pid: 9, hwnd: 2, name: "Book1", fullName: "Book1", saved: true),
        };
        var after = new[]
        {
            Book(pid: 8, hwnd: 1, name: "e2-cost-estimate.xlsx", fullName: path, saved: true),
            Book(pid: 9, hwnd: 2, name: "Book1", fullName: "Book1", saved: true),
        };
        Assert.Empty(ExcelOpenWorkbookInventoryContract.LostBystanders(before, after, path));
        Assert.False(ExcelOpenWorkbookInventoryContract.PreserveAggregateFailure([]));
    }

    private static JsonObject Book(int pid, long hwnd, string name, string fullName, bool saved) =>
        new()
        {
            ["processId"] = pid,
            ["excelHwnd"] = hwnd,
            ["name"] = name,
            ["fullName"] = fullName,
            ["saved"] = saved,
        };
}
