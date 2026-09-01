using System.Runtime.InteropServices;
using System.Runtime.Loader;

namespace DocBridge.Mcp;

/// <summary>
/// Runs adapter cleanup on every managed server shutdown path. A hard OS kill cannot execute
/// managed code, but terminating the client process still destroys its COM apartment; these hooks
/// cover EOF, Ctrl+C, SIGINT/SIGTERM, runtime unloading, normal exit, and unhandled unwind paths.
/// </summary>
internal sealed class ShutdownCleanupRegistration : IDisposable
{
    private readonly Action _cleanup;
    private readonly EventHandler _processExit;
    private readonly ConsoleCancelEventHandler _cancelKeyPress;
    private readonly Action<AssemblyLoadContext> _unloading;
    private readonly PosixSignalRegistration? _sigInt;
    private readonly PosixSignalRegistration? _sigTerm;
    private int _cleaned;
    private int _disposed;

    public ShutdownCleanupRegistration(Action cleanup)
    {
        _cleanup = cleanup;
        _processExit = (_, _) => CleanupOnce();
        _cancelKeyPress = (_, _) => CleanupOnce();
        _unloading = _ => CleanupOnce();

        AppDomain.CurrentDomain.ProcessExit += _processExit;
        Console.CancelKeyPress += _cancelKeyPress;
        AssemblyLoadContext.Default.Unloading += _unloading;

        try { _sigInt = PosixSignalRegistration.Create(PosixSignal.SIGINT, _ => CleanupOnce()); }
        catch (PlatformNotSupportedException) { }
        try { _sigTerm = PosixSignalRegistration.Create(PosixSignal.SIGTERM, _ => CleanupOnce()); }
        catch (PlatformNotSupportedException) { }
    }

    private void CleanupOnce()
    {
        if (Interlocked.Exchange(ref _cleaned, 1) != 0) return;
        try { _cleanup(); }
        catch { }
    }

    public void Dispose()
    {
        if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
        AppDomain.CurrentDomain.ProcessExit -= _processExit;
        Console.CancelKeyPress -= _cancelKeyPress;
        AssemblyLoadContext.Default.Unloading -= _unloading;
        _sigInt?.Dispose();
        _sigTerm?.Dispose();
        CleanupOnce();
    }
}
