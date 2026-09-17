using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

/// <summary>테스트 전용 ROT 원시 덤프. 바인딩/핸들/가시성 단계별 실패를 구분한다.</summary>
internal static class RotHelperProbe
{
    [DllImport("ole32.dll")]
    private static extern int GetRunningObjectTable(uint reserved, out IRunningObjectTable pprot);

    [DllImport("ole32.dll")]
    private static extern int CreateBindCtx(uint reserved, out IBindCtx ppbc);

    internal static void DumpHwpRotEntries(Action<string> log)
    {
        IRunningObjectTable? rot = null;
        IBindCtx? bindCtx = null;
        IEnumMoniker? enumerator = null;
        try
        {
            if (GetRunningObjectTable(0, out rot) != 0 || rot is null) { log("rot-unavailable"); return; }
            if (CreateBindCtx(0, out bindCtx) != 0 || bindCtx is null) { log("bindctx-unavailable"); return; }
            rot.EnumRunning(out enumerator);
            if (enumerator is null) { log("enum-unavailable"); return; }
            var monikers = new IMoniker[1];
            var total = 0;
            var hwp = 0;
            while (enumerator.Next(1, monikers, IntPtr.Zero) == 0)
            {
                total++;
                var moniker = monikers[0];
                try
                {
                    moniker.GetDisplayName(bindCtx, null, out var displayName);
                    if (displayName is not null &&
                        displayName.StartsWith("!HwpObject.", StringComparison.OrdinalIgnoreCase))
                    {
                        hwp++;
                        object? app = null;
                        try
                        {
                            rot.GetObject(moniker, out app);
                            if (app is null) { log($"entry {displayName}: bind-null"); continue; }
                            long hwnd;
                            try { hwnd = RotHelper.HwpWindowHandle(app); }
                            catch (Exception ex) { log($"entry {displayName}: handle-throw {ex.Message.Split('\n')[0]}"); continue; }
                            var pid = RotHelper.ProcessIdFromWindowHandle(hwnd);
                            bool visible;
                            try { visible = RotHelper.HwpWindowVisible(app); }
                            catch (Exception ex) { log($"entry {displayName}: visible-throw {ex.Message.Split('\n')[0]}"); continue; }
                            log($"entry {displayName}: bind-ok hwnd={hwnd} pid={pid} visible={visible}");
                            try
                            {
                                using var proc = System.Diagnostics.Process.GetProcessById(pid);
                                var main = proc.MainWindowHandle.ToInt64();
                                log($"entry {displayName}: main={main} title={proc.MainWindowTitle}");
                            }
                            catch (Exception ex) { log($"entry {displayName}: main-throw {ex.Message.Split('\n')[0]}"); }
                        }
                        catch (Exception ex) { log($"entry {displayName}: bind-throw {ex.Message.Split('\n')[0]}"); }
                        finally
                        {
                            try { if (app is not null && Marshal.IsComObject(app)) Marshal.ReleaseComObject(app); }
                            catch { }
                        }
                    }
                }
                catch (Exception ex) { log($"moniker-throw {ex.Message.Split('\n')[0]}"); }
                finally
                {
                    try { if (Marshal.IsComObject(moniker)) Marshal.ReleaseComObject(moniker); }
                    catch { }
                    monikers[0] = null!;
                }
            }
            log($"total-monikers={total} hwp-entries={hwp}");
        }
        finally
        {
            try { if (enumerator is not null && Marshal.IsComObject(enumerator)) Marshal.ReleaseComObject(enumerator); }
            catch { }
            try { if (bindCtx is not null && Marshal.IsComObject(bindCtx)) Marshal.ReleaseComObject(bindCtx); }
            catch { }
            try { if (rot is not null && Marshal.IsComObject(rot)) Marshal.ReleaseComObject(rot); }
            catch { }
        }
    }
}

/// <summary>
/// ROT 탐색 진단 전용. 바인딩/핸들/가시성 단계별 실패를 구분한다. E2E 전용.
/// </summary>
[Trait("Category", "E2E")]
public sealed class HwpRotLifetimeTests
{
    private static bool Enabled =>
        string.Equals(Environment.GetEnvironmentVariable("DOCBRIDGE_E2E"), "1", StringComparison.Ordinal);

    [Fact]
    public void ProbeRotEntriesWithoutBinding()
    {
        if (!Enabled) return;
        Exception? err = null;
        var t = new Thread(() =>
        {
            try
            {
                RotHelperProbe.DumpHwpRotEntries(
                    line => Console.Error.WriteLine("HwpRotProbe: " + line));
            }
            catch (Exception ex) { err = ex; }
        });
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        if (!t.Join(TimeSpan.FromMinutes(2))) throw new TimeoutException("rot probe timeout");
        if (err is not null) throw err;
    }
}
