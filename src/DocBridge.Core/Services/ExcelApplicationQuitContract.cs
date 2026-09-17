namespace DocBridge.Core.Services;

/// <summary>
/// An owned Application is not ownership of every workbook later opened in it.
/// Saved does not authorize Quit. Implicit dispose/detach must preserve any
/// nonempty Workbooks collection. Auto-Quit is allowed only for a proven empty
/// owned Application. Explicit <c>close_workbook</c> remains the exact-target route.
/// </summary>
public static class ExcelApplicationQuitContract
{
    public static bool MayAutoQuit(bool ownedInstance, int? workbookCount) =>
        ownedInstance && workbookCount == 0;

    public static string DetachWithoutQuitMessage(int? workbookCount)
    {
        if (workbookCount is null)
            return "owned Application workbook count is unproven; detaching without Quit so nonempty collections stay preserved";
        if (workbookCount > 0)
            return
                $"owned Application still has {workbookCount.Value} open workbook(s); Saved does not authorize Quit. Detaching with the collection preserved. Explicit close_workbook remains the exact-target route.";
        return "not quitting this Application";
    }
}
