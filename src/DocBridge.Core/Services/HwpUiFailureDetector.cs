using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// Detects fatal or blocking HWP UI dialogs before the worker retries COM activation.
/// A HWP TourPopup/FontCache initialization failure otherwise looks like a generic COM
/// timeout and used to create and close several blank HWP processes in succession.
/// </summary>
public static class HwpUiFailureDetector
{
    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    private static extern IntPtr GetWindow(IntPtr hWnd, uint uCmd);

    private const uint GwOwner = 4;

    public const string UpdateAction =
        "오류 창에서는 아니요(N)를 눌러 문서를 유지하고, 한글과 AI 프로그램을 모두 종료한 뒤 " +
        "DocBridge 0.4.10 이상을 설치해 windir/SystemRoot 복구가 적용됐는지 hwp_doctor로 확인하세요. " +
        "같은 오류가 남으면 [한컴 자동 업데이트 2024]로 최신 패치를 설치한 뒤 다시 시작하세요.";

    public static HwpUiFailure? Detect()
    {
        if (!OperatingSystem.IsWindows()) return null;

        HwpUiFailure? detected = null;
        try
        {
            EnumWindows((window, _) =>
            {
                if (detected is not null || !IsWindowVisible(window)) return detected is null;
                GetWindowThreadProcessId(window, out var processId);
                if (processId == 0 || !IsHwpProcess(unchecked((int)processId))) return true;

                var title = WindowText(window);
                var windowClass = WindowClass(window);
                var text = new StringBuilder(title);
                EnumChildWindows(window, (child, _) =>
                {
                    var childText = WindowText(child);
                    if (!string.IsNullOrWhiteSpace(childText)) text.Append('\n').Append(childText);
                    return true;
                }, IntPtr.Zero);

                var signature = ClassifyText(text.ToString());
                if (signature is null && IsBlockingHwpDialog(window, title, windowClass))
                    signature = "hwp-modal-dialog";
                if (signature is null) return true;

                detected = new HwpUiFailure(
                    unchecked((int)processId), window.ToInt64(), title, windowClass,
                    signature, text.ToString());
                return false;
            }, IntPtr.Zero);
        }
        catch
        {
            return null;
        }
        return detected;
    }

    public static HwpUiFailure? WaitForFailure(TimeSpan duration, TimeSpan? pollInterval = null)
    {
        var interval = pollInterval ?? TimeSpan.FromMilliseconds(125);
        var stopwatch = Stopwatch.StartNew();
        do
        {
            var failure = Detect();
            if (failure is not null) return failure;
            if (stopwatch.Elapsed >= duration) break;
            Thread.Sleep(interval);
        } while (true);
        return null;
    }

    internal static string? ClassifyText(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;
        if (text.Contains("PopupBorderImpl", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("TourPopup", StringComparison.OrdinalIgnoreCase))
            return "tour-popup-type-initializer";
        if (text.Contains("MS.Internal.FontCache.Util", StringComparison.OrdinalIgnoreCase) ||
            text.Contains("CultureFontManager", StringComparison.OrdinalIgnoreCase))
            return "font-cache-type-initializer";
        if (text.Contains("TypeInitializationException", StringComparison.OrdinalIgnoreCase) &&
            text.Contains("Hnc.Controls", StringComparison.OrdinalIgnoreCase))
            return "hnc-controls-type-initializer";
        return null;
    }

    internal static bool IsDeterministicFailureMessage(string? message) =>
        ClassifyText(message) is not null ||
        message?.Contains("HWP_UI_INITIALIZATION_FAILED", StringComparison.OrdinalIgnoreCase) == true;

    public static JsonObject ToJson(HwpUiFailure failure) => new()
    {
        ["processId"] = failure.ProcessId,
        ["windowHandle"] = failure.WindowHandle.ToString(),
        ["title"] = failure.Title,
        ["windowClass"] = failure.WindowClass,
        ["signature"] = failure.Signature,
    };

    private static bool IsBlockingHwpDialog(IntPtr window, string title, string windowClass)
    {
        if (!string.Equals(title.Trim(), "Hwp", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(title.Trim(), "한글", StringComparison.OrdinalIgnoreCase))
            return false;
        return GetWindow(window, GwOwner) != IntPtr.Zero ||
               string.Equals(windowClass, "#32770", StringComparison.Ordinal);
    }

    private static bool IsHwpProcess(int processId)
    {
        try
        {
            using var process = Process.GetProcessById(processId);
            return string.Equals(process.ProcessName, "Hwp", StringComparison.OrdinalIgnoreCase);
        }
        catch { return false; }
    }

    private static string WindowText(IntPtr window)
    {
        var buffer = new StringBuilder(8192);
        return GetWindowText(window, buffer, buffer.Capacity) > 0 ? buffer.ToString() : "";
    }

    private static string WindowClass(IntPtr window)
    {
        var buffer = new StringBuilder(256);
        return GetClassName(window, buffer, buffer.Capacity) > 0 ? buffer.ToString() : "";
    }
}

public sealed record HwpUiFailure(
    int ProcessId,
    long WindowHandle,
    string Title,
    string WindowClass,
    string Signature,
    string Text);
