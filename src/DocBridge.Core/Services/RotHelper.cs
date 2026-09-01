using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using System.Text;

namespace DocBridge.Core.Services;

/// <summary>
/// 실행 중인 COM 서버에 연결하기 위한 Running Object Table 헬퍼.
/// .NET Core에는 Marshal.GetActiveObject가 없으므로 oleaut32!GetActiveObject를 P/Invoke한다.
/// </summary>
public static class RotHelper
{
    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("ole32.dll", CharSet = CharSet.Unicode)]
    private static extern int CLSIDFromProgID(string lpszProgID, out Guid pclsid);

    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(uint reserved, out IRunningObjectTable? pprot);

    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(uint reserved, out IBindCtx? ppbc);

    [DllImport("oleaut32.dll", PreserveSig = true)]
    private static extern int GetActiveObject(ref Guid rclsid, IntPtr pvReserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object? ppunk);

    [DllImport("user32.dll")]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern bool EnumChildWindows(IntPtr hWndParent, EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool IsWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern IntPtr GetAncestor(IntPtr hWnd, uint gaFlags);

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("oleacc.dll")]
    private static extern int AccessibleObjectFromWindow(
        IntPtr hwnd,
        uint dwObjectId,
        ref Guid riid,
        [MarshalAs(UnmanagedType.Interface)] out object? ppvObject);

    private const uint ObjIdNativeOm = 0xFFFFFFF0;
    private const uint GaRoot = 2;
    private static readonly Guid IidIDispatch = new("00020400-0000-0000-C000-000000000046");

    /// <summary>Checks whether a cached Office application HWND still identifies a live window.</summary>
    public static bool IsWindowAlive(long hWnd)
    {
        if (hWnd == 0) return false;
        try { return IsWindow(new IntPtr(hWnd)); }
        catch { return false; }
    }

    /// <summary>실행 중인 인스턴스 연결. 없으면 null.</summary>
    public static object? GetActiveObject(string progId)
    {
        try
        {
            if (CLSIDFromProgID(progId, out var clsid) != 0) return null;
            var hr = GetActiveObject(ref clsid, IntPtr.Zero, out var obj);
            return hr == 0 ? obj : null;
        }
        catch { return null; }
    }

    /// <summary>
    /// 실행 중인 모든 Excel.Application 인스턴스를 반환한다.
    /// GetActiveObject는 여러 Excel 프로세스 중 하나만 반환하므로, 각 XLMAIN 창의
    /// EXCEL7 네이티브 오브젝트 모델을 통해 인스턴스를 찾아 HWND 기준으로 중복을 제거한다.
    /// 반드시 호출 측의 STA 스레드에서 사용해야 한다.
    /// </summary>
    public static IReadOnlyList<object> GetExcelApplications()
    {
        var result = new List<object>();
        var seenHwnds = new Dictionary<long, object>();

        void Add(object? app)
        {
            if (app is null) return;
            try
            {
                dynamic d = app;
                long hwnd = Convert.ToInt64((object)d.Hwnd);
                if (!seenHwnds.TryGetValue(hwnd, out object? existing))
                {
                    seenHwnds[hwnd] = app;
                    result.Add(app);
                }
                else if (ReferenceEquals(existing, app))
                {
                    // The same STA can receive the same RCW for another COM pointer acquisition.
                    // Balance one acquisition without invalidating the RCW retained in result.
                    ReleaseComReference(app);
                }
                else ReleaseComObject(app);
            }
            catch
            {
                // Hwnd를 읽지 못한 RCW는 안전하게 제외한다.
                ReleaseComObject(app);
            }
        }

        try
        {
            EnumWindows((top, _) =>
            {
                if (!string.Equals(WindowClass(top), "XLMAIN", StringComparison.Ordinal)) return true;

                object? application = null;
                EnumChildWindows(top, (child, _) =>
                {
                    if (!string.Equals(WindowClass(child), "EXCEL7", StringComparison.Ordinal)) return true;
                    var iid = IidIDispatch;
                    object? native = null;
                    if (AccessibleObjectFromWindow(child, ObjIdNativeOm, ref iid, out native) != 0 || native is null)
                        return true;
                    try
                    {
                        dynamic window = native;
                        application = (object)window.Application;
                        return false;
                    }
                    catch
                    {
                        return true;
                    }
                    finally
                    {
                        // AccessibleObjectFromWindow returns a Window RCW separate from the
                        // Application RCW. Leaving it alive is enough to keep EXCEL.EXE resident.
                        ReleaseComObject(native);
                    }
                }, IntPtr.Zero);

                Add(application);
                return true;
            }, IntPtr.Zero);
        }
        catch
        {
            // 아래 GetActiveObject fallback을 사용한다.
        }

        Add(GetActiveObject("Excel.Application"));
        return result;
    }

    /// <summary>
    /// 실행 중인 한글 자동화 객체를 한글 전용 ROT 모니커(!HwpObject.*)에서 찾는다.
    /// 공개 pyhwpx와 같은 연결 경로이며, 사용자가 열어 둔 창을 새 인스턴스보다 우선한다.
    /// 반환 순서는 전경 창, 표시 중인 창의 Z 순서, ROT 최신 항목 순이다.
    /// 반드시 호출 측의 STA 스레드에서 사용해야 한다.
    /// </summary>
    public static IReadOnlyList<object> GetHwpApplications()
    {
        var candidates = new List<(object App, long Hwnd, bool Foreground, bool Visible, int ZOrder, int RotOrder)>();
        var seenWindows = new HashSet<long>();
        var seenMonikers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        IRunningObjectTable? rot = null;
        IBindCtx? bindCtx = null;
        IEnumMoniker? enumerator = null;

        var foreground = RootWindow(GetForegroundWindow()).ToInt64();
        var zOrder = new Dictionary<long, int>();
        var z = 0;
        try
        {
            EnumWindows((hwnd, _) =>
            {
                var root = RootWindow(hwnd).ToInt64();
                if (root != 0 && !zOrder.ContainsKey(root)) zOrder[root] = z++;
                return true;
            }, IntPtr.Zero);
        }
        catch { }

        try
        {
            if (GetRunningObjectTable(0, out rot) != 0 || rot is null) return Array.Empty<object>();
            if (CreateBindCtx(0, out bindCtx) != 0 || bindCtx is null) return Array.Empty<object>();

            rot.EnumRunning(out enumerator);
            if (enumerator is null) return Array.Empty<object>();

            var monikers = new IMoniker[1];
            var order = 0;
            while (enumerator.Next(1, monikers, IntPtr.Zero) == 0)
            {
                var moniker = monikers[0];
                try
                {
                    moniker.GetDisplayName(bindCtx, null, out var displayName);
                    if (string.IsNullOrWhiteSpace(displayName) ||
                        !displayName.StartsWith("!HwpObject.", StringComparison.OrdinalIgnoreCase) ||
                        !seenMonikers.Add(displayName))
                        continue;

                    rot.GetObject(moniker, out var app);
                    if (app is null) continue;

                    var hwnd = HwpWindowHandle(app);
                    var root = RootWindow(new IntPtr(hwnd)).ToInt64();
                    if (root != 0 && !seenWindows.Add(root))
                    {
                        ReleaseComObject(app);
                        continue;
                    }

                    candidates.Add((
                        app,
                        root,
                        root != 0 && root == foreground,
                        root != 0 && IsWindowVisible(new IntPtr(root)),
                        root != 0 && zOrder.TryGetValue(root, out var rank) ? rank : int.MaxValue,
                        order++));
                }
                catch
                {
                    // 닫히는 중인 문서나 손상된 ROT 항목은 건너뛴다.
                }
                finally
                {
                    ReleaseComObject(moniker);
                    monikers[0] = null!;
                }
            }
        }
        catch
        {
            // 아래 ProgID fallback까지 진행한다.
        }
        finally
        {
            ReleaseComObject(enumerator);
            ReleaseComObject(bindCtx);
            ReleaseComObject(rot);
        }

        var fallback = GetActiveObject("HWPFrame.HwpObject");
        if (fallback is not null)
        {
            var hwnd = RootWindow(new IntPtr(HwpWindowHandle(fallback))).ToInt64();
            if (hwnd == 0 || seenWindows.Add(hwnd))
            {
                candidates.Add((
                    fallback,
                    hwnd,
                    hwnd != 0 && hwnd == foreground,
                    hwnd != 0 && IsWindowVisible(new IntPtr(hwnd)),
                    hwnd != 0 && zOrder.TryGetValue(hwnd, out var rank) ? rank : int.MaxValue,
                    int.MaxValue));
            }
            else
            {
                ReleaseComObject(fallback);
            }
        }

        return candidates
            .OrderByDescending(candidate => candidate.Foreground)
            .ThenByDescending(candidate => candidate.Visible)
            .ThenBy(candidate => candidate.ZOrder)
            .ThenByDescending(candidate => candidate.RotOrder)
            .Select(candidate => candidate.App)
            .ToList();
    }

    public static long HwpWindowHandle(object app)
    {
        try
        {
            dynamic hwp = app;
            return Convert.ToInt64(hwp.XHwpWindows.Active_XHwpWindow.WindowHandle);
        }
        catch { return 0; }
    }

    /// <summary>
    /// 사용자가 볼 수 있는 한글 최상위 창인지 확인한다. -Automation -Embedding으로
    /// 남은 숨은 인스턴스는 라이브 문서로 오인하지 않아야 한다.
    /// </summary>
    public static bool HwpWindowVisible(object app)
    {
        try
        {
            var hwnd = RootWindow(new IntPtr(HwpWindowHandle(app)));
            return hwnd != IntPtr.Zero && IsWindow(hwnd) && IsWindowVisible(hwnd);
        }
        catch { return false; }
    }

    /// <summary>창 핸들로 소유 프로세스 ID를 구한다. 잘못된 핸들이면 0.</summary>
    public static int ProcessIdFromWindowHandle(long windowHandle)
    {
        if (windowHandle == 0) return 0;
        try
        {
            GetWindowThreadProcessId(new IntPtr(windowHandle), out var pid);
            return unchecked((int)pid);
        }
        catch { return 0; }
    }

    public static void ReleaseComObject(object? value)
    {
        if (value is null) return;
        try
        {
            if (Marshal.IsComObject(value)) Marshal.FinalReleaseComObject(value);
        }
        catch { }
    }

    public static void ReleaseComReference(object? value)
    {
        if (value is null) return;
        try
        {
            if (Marshal.IsComObject(value)) Marshal.ReleaseComObject(value);
        }
        catch { }
    }

    private static IntPtr RootWindow(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero) return IntPtr.Zero;
        try
        {
            var root = GetAncestor(hwnd, GaRoot);
            return root == IntPtr.Zero ? hwnd : root;
        }
        catch { return hwnd; }
    }

    private static string WindowClass(IntPtr hwnd)
    {
        var buffer = new StringBuilder(128);
        return GetClassName(hwnd, buffer, buffer.Capacity) > 0 ? buffer.ToString() : "";
    }

    /// <summary>새 인스턴스 생성. ProgID 미등록이면 null.</summary>
    public static object? CreateInstance(string progId)
    {
        try
        {
            var type = Type.GetTypeFromProgID(progId);
            return type is null ? null : Activator.CreateInstance(type);
        }
        catch { return null; }
    }

    public static string? VersionOf(string progId)
    {
        object? app = null;
        try
        {
            // A version probe must never create a hidden Office process. Installation checks use
            // registry/file metadata; this helper only reports a running COM server.
            app = GetActiveObject(progId);
            if (app is null) return null;
            return (string?)((dynamic)app).Version?.ToString();
        }
        catch { return null; }
        finally { ReleaseComObject(app); }
    }
}
