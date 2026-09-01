using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public sealed class ForegroundInteractionGuardTests
{
    [Fact]
    public void Win32_native_guard_entry_points_are_resolvable()
    {
        if (!OperatingSystem.IsWindows()) return;

        var guard = new ForegroundInteractionGuard("excel");
        var result = guard.Complete();

        Assert.Equal("preserve-foreground", result["policy"]!.GetValue<string>());
    }

    [Fact]
    public void Restores_original_window_when_target_steals_foreground()
    {
        var native = new FakeForegroundNative(10, 100);
        native.Processes[10] = 1;
        native.Processes[20] = 2;
        var guard = new ForegroundInteractionGuard("hwp", native);
        guard.TrackTargetWindow(20);

        native.Foreground = 20;
        Assert.True(guard.Checkpoint());
        var result = guard.Complete();

        Assert.Equal(10, native.Foreground);
        Assert.True(result["foregroundPreserved"]!.GetValue<bool>());
        Assert.True(result["foregroundRestored"]!.GetValue<bool>());
        Assert.False(result["userActivityDetected"]!.GetValue<bool>());
    }

    [Fact]
    public void Newly_activated_target_does_not_misattribute_other_app_input()
    {
        var native = new FakeForegroundNative(10, 100);
        native.Processes[10] = 1;
        native.Processes[20] = 2;
        var guard = new ForegroundInteractionGuard("hwp", native);
        guard.TrackTargetWindow(20);

        native.LastInputTick = 101;
        native.Foreground = 20;
        Assert.True(guard.Checkpoint(stopOnConcurrentInput: true));
        var result = guard.Complete();

        Assert.Equal(10, native.Foreground);
        Assert.False(result["userActivityDetected"]!.GetValue<bool>());
        Assert.False(result["interrupted"]!.GetValue<bool>());
    }

    [Fact]
    public void Follows_latest_non_target_window_before_restoring()
    {
        var native = new FakeForegroundNative(10, 100);
        native.Processes[10] = 1;
        native.Processes[11] = 3;
        native.Processes[20] = 2;
        var guard = new ForegroundInteractionGuard("cad", native);
        guard.TrackTargetWindow(20);

        native.Foreground = 11;
        Assert.True(guard.Checkpoint());
        native.Foreground = 20;
        Assert.True(guard.Checkpoint());

        Assert.Equal(11, native.Foreground);
    }

    [Fact]
    public void Sibling_window_in_same_process_is_treated_as_user_foreground()
    {
        var native = new FakeForegroundNative(10, 100);
        native.Processes[10] = 1;
        native.Processes[20] = 2;
        native.Processes[21] = 2;
        var guard = new ForegroundInteractionGuard("hwp", native);
        guard.TrackTargetWindow(20);

        native.Foreground = 21;
        Assert.True(guard.Checkpoint());
        var result = guard.Complete();

        Assert.Equal(21, native.Foreground);
        Assert.False(result["foregroundRestoreAttempted"]!.GetValue<bool>());
        Assert.True(result["foregroundPreserved"]!.GetValue<bool>());
    }

    [Fact]
    public void Process_tracking_remains_available_for_windowless_fallback()
    {
        var native = new FakeForegroundNative(20, 100);
        native.Processes[20] = 2;
        var guard = new ForegroundInteractionGuard("hwp", native);
        guard.TrackTargetProcess(2);

        Assert.True(guard.Checkpoint());
        var result = guard.Complete();

        Assert.True(result["targetWasForegroundAtStart"]!.GetValue<bool>());
        Assert.False(result["foregroundRestoreAttempted"]!.GetValue<bool>());
    }

    [Fact]
    public void Does_not_move_a_target_that_was_already_foreground()
    {
        var native = new FakeForegroundNative(20, 100);
        native.Processes[20] = 2;
        var guard = new ForegroundInteractionGuard("excel", native);
        guard.TrackTargetWindow(20);

        Assert.True(guard.Checkpoint());
        var result = guard.Complete();

        Assert.Equal(20, native.Foreground);
        Assert.True(result["targetWasForegroundAtStart"]!.GetValue<bool>());
        Assert.False(result["foregroundRestoreAttempted"]!.GetValue<bool>());
    }

    [Fact]
    public void Concurrent_input_in_an_already_foreground_target_interrupts_without_moving_it()
    {
        var native = new FakeForegroundNative(20, 100);
        native.Processes[20] = 2;
        var guard = new ForegroundInteractionGuard("cad", native);
        guard.TrackTargetWindow(20);

        native.LastInputTick = 101;
        Assert.False(guard.Checkpoint(stopOnConcurrentInput: true));
        var result = guard.Complete();

        Assert.Equal(20, native.Foreground);
        Assert.True(result["userActivityDetected"]!.GetValue<bool>());
        Assert.True(result["interrupted"]!.GetValue<bool>());
        Assert.False(result["foregroundRestoreAttempted"]!.GetValue<bool>());
    }

    private sealed class FakeForegroundNative : IForegroundWindowNative
    {
        public FakeForegroundNative(long foreground, uint lastInputTick)
        {
            Foreground = foreground;
            LastInputTick = lastInputTick;
        }

        public long Foreground { get; set; }
        public uint LastInputTick { get; set; }
        public Dictionary<long, uint> Processes { get; } = new();
        public long GetForegroundWindow() => Foreground;
        public long GetRootWindow(long windowHandle) => windowHandle;
        public uint GetWindowProcessId(long windowHandle) => Processes.GetValueOrDefault(windowHandle);
        public uint GetLastInputTick() => LastInputTick;
        public bool IsWindow(long windowHandle) => windowHandle != 0;
        public bool SetForegroundWindow(long windowHandle)
        {
            Foreground = windowHandle;
            return true;
        }
    }
}
