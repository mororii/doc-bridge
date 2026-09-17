using System.Runtime.InteropServices;

namespace DocBridge.Development.ExcelProductionWorkflowProbe;

[ComImport]
[Guid("00000016-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
internal interface IOleMessageFilter
{
    [PreserveSig] int HandleInComingCall(int callType, IntPtr taskCaller, int tickCount, IntPtr reject);
    [PreserveSig] int RetryRejectedCall(IntPtr taskCallee, int tickCount, int rejectType);
    [PreserveSig] int MessagePending(IntPtr taskCallee, int tickCount, int pendingType);
}

/// <summary>Retries Excel RPC_E_SERVERCALL_RETRYLATER instead of failing the first busy tick.</summary>
internal sealed class ExcelMessageFilter : IOleMessageFilter
{
    [DllImport("ole32.dll")]
    private static extern int CoRegisterMessageFilter(IOleMessageFilter? newFilter, out IOleMessageFilter oldFilter);

    private static IOleMessageFilter? _previous;

    public static void Register()
    {
        var filter = new ExcelMessageFilter();
        CoRegisterMessageFilter(filter, out _previous);
    }

    public static void Revoke()
    {
        CoRegisterMessageFilter(_previous, out _);
        _previous = null;
    }

    public int HandleInComingCall(int callType, IntPtr taskCaller, int tickCount, IntPtr reject) => 0;

    public int RetryRejectedCall(IntPtr taskCallee, int tickCount, int rejectType) =>
        rejectType == 2 ? 200 : -1;

    public int MessagePending(IntPtr taskCallee, int tickCount, int pendingType) => 2;
}
