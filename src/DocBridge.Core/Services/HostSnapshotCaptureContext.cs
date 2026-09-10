namespace DocBridge.Core.Services;

/// <summary>
/// Host-owned snapshot capture context. Only <see cref="DocBridgeHost"/> stamps
/// this onto snapshot metadata. Adapters must not infer execute vs dry-run.
/// </summary>
internal static class HostSnapshotCaptureContext
{
    internal const string MetadataKey = "hostSnapshotContext";
    internal const string Execute = "execute";
}
