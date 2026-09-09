using System.Reflection;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

/// <summary>
/// Invokes product <see cref="ExcelOwnerWatchdog"/> behavior. Does not launch Excel or copy policy into locals.
/// </summary>
public sealed class ExcelLifecycleProbePolicyTests
{
    [Fact]
    public void Watchdog_try_run_rejects_malformed_arguments_without_starting_excel()
    {
        Assert.False(ExcelOwnerWatchdog.TryRun(["status"], out var ignored));
        Assert.Equal(0, ignored);

        Assert.True(ExcelOwnerWatchdog.TryRun(["--excel-owner-watchdog"], out var missingArgs));
        Assert.Equal(2, missingArgs);

        Assert.True(ExcelOwnerWatchdog.TryRun(
            ["--excel-owner-watchdog", "parent", "excel", "ready", "release"], out var badNumbers));
        Assert.Equal(2, badNumbers);
    }

    [Fact]
    public void TryQuitSafely_quits_empty_application()
    {
        var app = new ExcelReconnectTests.EmptyOwnedExcelApplication(openThrows: false);
        Assert.True(InvokeTryQuitSafely(app));
        Assert.Equal(1, app.QuitCalls);
    }

    [Fact]
    public void TryQuitSafely_quits_when_single_workbook_is_saved()
    {
        var app = new ExcelReconnectTests.LifecycleExcelApplication(saved: true);
        Assert.True(InvokeTryQuitSafely(app));
        Assert.Equal(1, app.QuitCalls);
    }

    [Fact]
    public void TryQuitSafely_refuses_unsaved_workbook()
    {
        var app = new ExcelReconnectTests.LifecycleExcelApplication(saved: false);
        Assert.False(InvokeTryQuitSafely(app));
        Assert.Equal(0, app.QuitCalls);
    }

    [Fact]
    public void TryQuitSafely_refuses_when_saved_state_cannot_be_read()
    {
        var app = new UnreadableSavedExcelApplication();
        // Fakes must be public: Core's dynamic binder cannot see private Workbooks.
        // TryQuitSafely's outer try/finally does not catch a Workbooks bind failure;
        // Watchdog.Run does. Either return-false or throw is acceptable; Quit is not.
        var quitAllowed = TryInvokeTryQuitSafely(app, out var helperThrew);
        if (!helperThrew)
            Assert.False(quitAllowed);
        Assert.Equal(0, app.QuitCalls);
    }

    private static bool InvokeTryQuitSafely(object application)
    {
        var quitAllowed = TryInvokeTryQuitSafely(application, out var helperThrew);
        Assert.False(helperThrew);
        return quitAllowed;
    }

    private static bool TryInvokeTryQuitSafely(object application, out bool helperThrew)
    {
        var method = typeof(ExcelOwnerWatchdog).GetMethod(
            "TryQuitSafely",
            BindingFlags.NonPublic | BindingFlags.Static);
        Assert.NotNull(method);
        try
        {
            helperThrew = false;
            return Assert.IsType<bool>(method!.Invoke(null, [application]));
        }
        catch (TargetInvocationException)
        {
            helperThrew = true;
            return false;
        }
    }

    public sealed class UnreadableSavedExcelApplication
    {
        public int QuitCalls { get; private set; }
        public UnreadableSavedWorkbooks Workbooks { get; } = new();
        public void Quit() => QuitCalls++;
    }

    public sealed class UnreadableSavedWorkbooks
    {
        public int Count => 1;
        public UnreadableSavedWorkbook Item(int index) =>
            index == 1 ? new UnreadableSavedWorkbook() : throw new ArgumentOutOfRangeException(nameof(index));
    }

    public sealed class UnreadableSavedWorkbook
    {
        public string Name => "unreadable.xlsx";
        public bool Saved => throw new InvalidOperationException("synthetic unread Saved");
    }
}
