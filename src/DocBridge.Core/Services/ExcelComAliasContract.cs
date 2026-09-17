namespace DocBridge.Core.Services;

/// <summary>
/// ROT/window discovery can return the same Application RCW the adapter already
/// holds. Final-releasing that alias separates live attached locals from the
/// native object. Borrowed child collections (Workbooks/Workbook from a live
/// Application) are released once, not FinalReleased.
/// </summary>
public static class ExcelComAliasContract
{
    public static bool MustPreserveLiveAlias(object? discovered, object? liveAlias) =>
        discovered is not null && liveAlias is not null && ReferenceEquals(discovered, liveAlias);

    public static bool DiscoveryFinalReleasesBorrowedChild => false;

    public static bool DiscoveryFinalReleasesApplicationAlias => false;
}
