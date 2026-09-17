using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

/// <summary>
/// 핀 폴백 증명: 홀더가 끝난 숨은 소유 인스턴스를 새 어댑터가 핀으로 찾아
/// 다시 표시하고 읽는다. DOCBRIDGE_E2E=1일 때만 실행. Category=E2E.
/// </summary>
[Trait("Category", "E2E")]
public sealed class HwpInstancePinE2ETests : IDisposable
{
    private static bool Enabled =>
        string.Equals(Environment.GetEnvironmentVariable("DOCBRIDGE_E2E"), "1", StringComparison.Ordinal);

    private readonly TestHome _home = new();
    private int _ownedPid;
    private bool _disposed;

    private static void RunOnSta(Action work, TimeSpan timeout)
    {
        Exception? err = null;
        var t = new Thread(() =>
        {
            try { work(); }
            catch (Exception ex) { err = ex; }
        });
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        if (!t.Join(timeout)) throw new TimeoutException("pin e2e sta timeout");
        if (err is not null) throw err;
    }

    [Fact]
    public void Pinned_hidden_instance_is_found_unhidden_and_read()
    {
        if (!Enabled) return;
        var existing = HwpE2EOwnership.CurrentHwpProcessIds();
        string createdId = "";
        RunOnSta(() =>
        {
            var type = Type.GetTypeFromProgID("HWPFrame.HwpObject")
                ?? throw new InvalidOperationException("HWP not installed");
            dynamic hwp = HwpEnvironmentDoctor.RunWithAutomationWorkingDirectory(
                () => Activator.CreateInstance(type)!)!;
            try
            {
                if (!HwpE2EOwnership.TryReadInventory(hwp, out HwpE2EDocumentSnapshot before))
                    throw new InvalidOperationException("pin: before-inventory failed");
                var windowPid = RotHelper.ProcessIdFromWindowHandle(
                    RotHelper.HwpWindowHandle((object)hwp));
                var pid = HwpE2EOwnershipPolicy.ResolveProcessId(
                    windowPid, existing, HwpE2EOwnership.CurrentHwpProcessIds());
                if (!HwpE2EOwnership.TryFileNew(hwp))
                    throw new InvalidOperationException("pin: FileNew failed");
                if (!HwpE2EOwnership.TryReadInventory(hwp, out HwpE2EDocumentSnapshot after))
                    throw new InvalidOperationException("pin: after-inventory failed");
                if (!HwpE2EOwnershipPolicy.TryProve(existing, pid, true, before, after,
                        out _, out var error))
                    throw new InvalidOperationException("pin: ownership unproven: " + error);
                if (!HwpE2EOwnershipPolicy.TryIdentifyCreatedTab(before, after, out var tabId))
                    throw new InvalidOperationException("pin: created tab unknown");
                createdId = tabId;
                _ownedPid = pid;
                HwpE2EOwnership.InsertSeedText(hwp, "핀복구검증문장");
                var hwnd = RotHelper.HwpWindowHandle((object)hwp);
                if (!HwpInstancePin.TrySavePin(pid, hwnd, "e2e", out var saveError))
                    throw new InvalidOperationException("pin: save failed: " + saveError);
            }
            finally
            {
                // 홀더 종료를 재현한다: 모든 참조 해제 후 GC.
                try
                {
                    if (Marshal.IsComObject((object)hwp))
                        Marshal.ReleaseComObject((object)hwp);
                }
                catch { }
            }
        }, TimeSpan.FromMinutes(3));
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();

        try
        {
            Assert.True(HwpInstancePin.TryValidatePin(HwpInstancePin.TryLoadPin()!));
            using var host = new DocBridgeHost(_home.Options);
            host.Router.Register("hwp", new HwpAdapter());
            var docRef = $"untitled-{_ownedPid}-{createdId}";
            var read = host.Read("hwp", new JsonObject { ["scope"] = "document", ["documentRef"] = docRef });
            Assert.True(Json.GetBool(read, "ok"), $"pinned read failed: {read}");
            Assert.Contains("핀복구검증문장", Json.GetString(read, "text"));
            var shown = false;
            var deadline = DateTime.UtcNow.AddSeconds(10);
            while (!shown && DateTime.UtcNow < deadline)
            {
                try
                {
                    using var proc = System.Diagnostics.Process.GetProcessById(_ownedPid);
                    shown = proc.MainWindowHandle.ToInt64() != 0;
                }
                catch { break; }
                if (!shown) Thread.Sleep(500);
            }
            Assert.True(shown);
        }
        finally
        {
            HwpInstancePin.ClearPin();
        }
    }

    [Fact]
    public void Pin_file_roundtrip_and_validation_rules()
    {
        if (!Enabled) return;
        var path = HwpInstancePin.PinPath();
        var backup = File.Exists(path) ? File.ReadAllBytes(path) : null;
        try
        {
            HwpInstancePin.ClearPin();
            Assert.Null(HwpInstancePin.TryLoadPin());
            Assert.True(HwpInstancePin.TrySavePin(424242, 0, "unit", out _));
            var loaded = HwpInstancePin.TryLoadPin();
            Assert.NotNull(loaded);
            // 존재하지 않는 PID는 검증 탈락한다.
            Assert.False(HwpInstancePin.TryValidatePin(loaded!));
        }
        finally
        {
            try
            {
                if (backup is null) HwpInstancePin.ClearPin();
                else File.WriteAllBytes(path, backup);
            }
            catch { }
        }
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        try
        {
            HwpInstancePin.ClearPin();
            if (_ownedPid > 0)
            {
                using var proc = System.Diagnostics.Process.GetProcessById(_ownedPid);
                proc.Kill();
                proc.WaitForExit(10000);
            }
        }
        catch { }
        _home.Dispose();
    }
}
