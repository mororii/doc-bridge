using System.Runtime.InteropServices;
using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// COM 서버가 내부 문서/탭을 활성화하면서 Windows 전경 창까지 가져가는 경우를 감시한다.
/// 사용자가 작업 중인 최신 비대상 창을 기억하고, 대상 앱이 전경을 가져간 경우에만 복구한다.
/// </summary>
internal sealed class ForegroundInteractionGuard
{
    private const int WatchdogIntervalMilliseconds = 30;
    private readonly string _app;
    private readonly IForegroundWindowNative _native;
    private readonly object _sync = new();
    private readonly HashSet<long> _targetRoots = new();
    private readonly HashSet<uint> _targetProcesses = new();
    private long _preferredForeground;
    private long _lastObservedForeground;
    private uint _lastInputTick;
    private bool _completed;
    private bool _foregroundChangeDetected;
    private bool _restoreAttempted;
    private bool _restored;
    private bool _restoreFailed;
    private bool _userActivityDetected;
    private bool _interrupted;
    private bool _targetWasForegroundAtStart;
    private int _checkpoints;
    private Timer? _watchdog;

    public ForegroundInteractionGuard(string app, IForegroundWindowNative? native = null)
    {
        _app = app;
        _native = native ?? Win32ForegroundWindowNative.Instance;
        _preferredForeground = Root(_native.GetForegroundWindow());
        _lastObservedForeground = _preferredForeground;
        _lastInputTick = _native.GetLastInputTick();
    }

    public bool Interrupted => _interrupted;
    public bool UserActivityDetected => _userActivityDetected;
    public bool TargetWasForegroundAtStart => _targetWasForegroundAtStart;

    public void TrackTargetWindow(long windowHandle)
    {
        lock (_sync)
        {
            var root = Root(windowHandle);
            if (root == 0) return;
            _targetRoots.Add(root);
            if (_preferredForeground == root) _targetWasForegroundAtStart = true;
            // 실제 Win32 환경에서만 짧은 주기 감시를 사용한다. COM 호출 한 번이 오래
            // 걸리더라도 작업 경계까지 기다리지 않고 사용자의 원래 창을 복구한다.
            if (ReferenceEquals(_native, Win32ForegroundWindowNative.Instance) && _watchdog is null)
                _watchdog = new Timer(_ => WatchdogCheckpoint(), null,
                    WatchdogIntervalMilliseconds, WatchdogIntervalMilliseconds);
        }
    }

    public void TrackTargetProcess(int processId)
    {
        lock (_sync)
        {
            if (processId <= 0) return;
            var targetProcess = unchecked((uint)processId);
            _targetProcesses.Add(targetProcess);
            if (_preferredForeground != 0 &&
                _native.GetWindowProcessId(_preferredForeground) == targetProcess)
                _targetWasForegroundAtStart = true;
            if (ReferenceEquals(_native, Win32ForegroundWindowNative.Instance) && _watchdog is null)
                _watchdog = new Timer(_ => WatchdogCheckpoint(), null,
                    WatchdogIntervalMilliseconds, WatchdogIntervalMilliseconds);
        }
    }

    /// <summary>
    /// 안전한 작업 경계에서 호출한다. 대상 앱이 전경을 가져갔으면 최신 비대상 전경 창을 복구한다.
    /// 같은 구간에 사용자 입력도 있었다면 다음 작업을 중단하도록 false를 반환한다.
    /// </summary>
    public bool Checkpoint(bool stopOnConcurrentInput = true)
    {
        lock (_sync) return CheckpointCore(stopOnConcurrentInput);
    }

    private void WatchdogCheckpoint()
    {
        try
        {
            lock (_sync)
                if (!_completed) _ = CheckpointCore(stopOnConcurrentInput: false);
        }
        catch
        {
            // 포커스 보존은 문서 편집 결과를 실패시키지 않는 보조 안전장치다.
        }
    }

    private bool CheckpointCore(bool stopOnConcurrentInput)
    {
        if (_completed) return !_interrupted;
        _checkpoints++;
        var current = Root(_native.GetForegroundWindow());
        var inputTick = _native.GetLastInputTick();
        var inputChanged = inputTick != _lastInputTick;

        var currentIsTarget = current != 0 && IsTarget(current);
        if (current != 0 && !currentIsTarget)
        {
            // 사용자가 다른 프로그램으로 옮겼다면 그 창을 이후 복구 대상으로 삼는다.
            _preferredForeground = current;
        }
        else if (currentIsTarget)
        {
            // A COM action can activate its own application after the user typed in a
            // different app.  Global GetLastInputInfo cannot attribute that earlier input
            // to a window, so only classify it as target-app activity when the target was
            // already foreground at the preceding boundary.  A newly activated target is
            // restored without producing a false concurrent-edit interruption.
            if (inputChanged && (_targetWasForegroundAtStart || IsTarget(_lastObservedForeground)))
            {
                _userActivityDetected = true;
                if (stopOnConcurrentInput) _interrupted = true;
            }

            if (_preferredForeground != 0 && current != _preferredForeground)
            {
                _foregroundChangeDetected = true;
                if (_native.IsWindow(_preferredForeground))
                {
                    _restoreAttempted = true;
                    _restored = _native.SetForegroundWindow(_preferredForeground) ||
                                Root(_native.GetForegroundWindow()) == _preferredForeground;
                    _restoreFailed |= !_restored;
                }
            }
        }

        _lastInputTick = inputTick;
        _lastObservedForeground = Root(_native.GetForegroundWindow());
        return !_interrupted;
    }

    public JsonObject Complete()
    {
        lock (_sync)
        {
            _watchdog?.Dispose();
            _watchdog = null;
            if (!_completed)
            {
                CheckpointCore(stopOnConcurrentInput: false);
                _completed = true;
            }

            var final = Root(_native.GetForegroundWindow());
            var finalIsTarget = final != 0 && IsTarget(final);
            var preserved = _targetWasForegroundAtStart || !finalIsTarget ||
                            (_preferredForeground != 0 && final == _preferredForeground);
            return new JsonObject
            {
                ["policy"] = "preserve-foreground",
                ["app"] = _app,
                ["foregroundPreserved"] = preserved && !_restoreFailed,
                ["foregroundChangeDetected"] = _foregroundChangeDetected,
                ["foregroundRestoreAttempted"] = _restoreAttempted,
                ["foregroundRestored"] = _restoreAttempted ? _restored && !_restoreFailed : null,
                ["targetWasForegroundAtStart"] = _targetWasForegroundAtStart,
                ["userActivityDetected"] = _userActivityDetected,
                ["interrupted"] = _interrupted,
                ["checkpoints"] = _checkpoints,
                ["watchdogIntervalMs"] = WatchdogIntervalMilliseconds,
            };
        }
    }

    private bool IsTarget(long root)
    {
        if (_targetRoots.Contains(root)) return true;
        var processId = _native.GetWindowProcessId(root);
        return processId != 0 && _targetProcesses.Contains(processId);
    }

    private long Root(long windowHandle) => windowHandle == 0 ? 0 : _native.GetRootWindow(windowHandle);
}

internal interface IForegroundWindowNative
{
    long GetForegroundWindow();
    long GetRootWindow(long windowHandle);
    uint GetWindowProcessId(long windowHandle);
    uint GetLastInputTick();
    bool IsWindow(long windowHandle);
    bool SetForegroundWindow(long windowHandle);
}

internal sealed class Win32ForegroundWindowNative : IForegroundWindowNative
{
    public static readonly Win32ForegroundWindowNative Instance = new();
    private const uint GaRoot = 2;

    [StructLayout(LayoutKind.Sequential)]
    private struct LastInputInfo
    {
        public uint CbSize;
        public uint DwTime;
    }

    [DllImport("user32.dll", EntryPoint = "GetForegroundWindow")]
    private static extern IntPtr GetForegroundWindowNative();

    [DllImport("user32.dll")]
    private static extern IntPtr GetAncestor(IntPtr hWnd, uint gaFlags);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetLastInputInfo(ref LastInputInfo info);

    [DllImport("user32.dll", EntryPoint = "IsWindow")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowNative(IntPtr hWnd);

    [DllImport("user32.dll", EntryPoint = "SetForegroundWindow")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetForegroundWindowNative(IntPtr hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool BringWindowToTop(IntPtr hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool AttachThreadInput(uint idAttach, uint idAttachTo, bool attach);

    [DllImport("kernel32.dll")]
    private static extern uint GetCurrentThreadId();

    public long GetForegroundWindow() => GetForegroundWindowNative().ToInt64();

    public long GetRootWindow(long windowHandle)
    {
        var handle = new IntPtr(windowHandle);
        var root = GetAncestor(handle, GaRoot);
        return (root == IntPtr.Zero ? handle : root).ToInt64();
    }

    public uint GetWindowProcessId(long windowHandle)
    {
        GetWindowThreadProcessId(new IntPtr(windowHandle), out var processId);
        return processId;
    }

    public uint GetLastInputTick()
    {
        var info = new LastInputInfo { CbSize = (uint)Marshal.SizeOf<LastInputInfo>() };
        return GetLastInputInfo(ref info) ? info.DwTime : 0;
    }

    public bool IsWindow(long windowHandle) => IsWindowNative(new IntPtr(windowHandle));
    public bool SetForegroundWindow(long windowHandle)
    {
        var target = new IntPtr(windowHandle);
        if (SetForegroundWindowNative(target) || GetForegroundWindowNative() == target) return true;

        // Windows normally prevents a background COM server from stealing focus.  Here the
        // destination is the user's previously active window, so temporarily join the input
        // queues and restore that exact window.  This is window-state repair, not UI driving.
        var currentThread = GetCurrentThreadId();
        var foreground = GetForegroundWindowNative();
        var foregroundThread = foreground == IntPtr.Zero ? 0 : GetWindowThreadProcessId(foreground, out _);
        var targetThread = GetWindowThreadProcessId(target, out _);
        var attachedForeground = foregroundThread != 0 && foregroundThread != currentThread &&
                                 AttachThreadInput(currentThread, foregroundThread, true);
        var attachedTarget = targetThread != 0 && targetThread != currentThread && targetThread != foregroundThread &&
                             AttachThreadInput(currentThread, targetThread, true);
        try
        {
            _ = BringWindowToTop(target);
            return SetForegroundWindowNative(target) || GetForegroundWindowNative() == target;
        }
        finally
        {
            if (attachedTarget) _ = AttachThreadInput(currentThread, targetThread, false);
            if (attachedForeground) _ = AttachThreadInput(currentThread, foregroundThread, false);
        }
    }
}
