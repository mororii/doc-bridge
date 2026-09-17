using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// Workbook.Windows.Item(1) Zoom/DisplayGridlines belong to the active sheet.
/// Capture the named target by activating it, verify the active sheet, then
/// restore the original workbook/sheet/selection/scroll. Activation, read,
/// or restore failure is readable:false.
/// </summary>
public static class ExcelSheetViewCaptureContract
{
    public readonly record struct WindowBookmark(
        string? ActiveSheet,
        string? Selection,
        int? ScrollRow,
        int? ScrollColumn,
        string? ActiveWorkbook = null);

    public interface ISheetViewSurface
    {
        WindowBookmark CaptureOriginal();
        void Activate(string sheet);
        JsonObject ReadActiveView();
        void Restore(WindowBookmark original);
        string? ActiveSheet { get; }
    }

    public static bool MustActivateTarget(string targetSheet, string? activeSheet) =>
        !string.Equals(targetSheet, activeSheet, StringComparison.OrdinalIgnoreCase);

    public static bool CapturedViewBelongsTo(JsonObject view, string sheet) =>
        string.Equals(Json.GetString(view, "sheet"), sheet, StringComparison.OrdinalIgnoreCase);

    public static JsonObject Unreadable(string sheet, string error) =>
        new()
        {
            ["readable"] = false,
            ["sheet"] = sheet,
            ["error"] = error,
        };

    public static JsonObject CaptureTargetView(ISheetViewSurface surface, string targetSheet)
    {
        var original = surface.CaptureOriginal();
        JsonObject? view = null;
        var restoreFailed = false;
        string? restoreError = null;
        try
        {
            if (MustActivateTarget(targetSheet, original.ActiveSheet))
                surface.Activate(targetSheet);
            if (MustActivateTarget(targetSheet, surface.ActiveSheet))
            {
                view = Unreadable(targetSheet, "active sheet is not the capture target");
                return view;
            }

            view = surface.ReadActiveView();
            view["sheet"] = targetSheet;
            view["readable"] = true;
            return view;
        }
        catch (Exception ex)
        {
            view = Unreadable(targetSheet, ex.Message);
            return view;
        }
        finally
        {
            try
            {
                surface.Restore(original);
            }
            catch (Exception ex)
            {
                restoreFailed = true;
                restoreError = ex.Message;
            }

            if (restoreFailed && view is not null)
            {
                view["readable"] = false;
                view["error"] = restoreError ?? "window bookmark restore failed";
            }
        }
    }
}
