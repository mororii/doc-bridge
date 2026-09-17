using System.Runtime.InteropServices;

namespace DocBridge.Core.Services;

/// <summary>
/// Release one owned COM RCW. Managed fakes are ignored. Does not
/// FinalRelease or touch borrowed aliases passed in by the caller.
/// </summary>
public static class ExcelOwnedCom
{
    public static void ReleaseOnce(object? value)
    {
        if (value is null) return;
        try
        {
            if (Marshal.IsComObject(value))
                Marshal.ReleaseComObject(value);
        }
        catch (InvalidComObjectException)
        {
            /* already disconnected */
        }
        catch (COMException)
        {
            /* native object already gone */
        }
    }
}
