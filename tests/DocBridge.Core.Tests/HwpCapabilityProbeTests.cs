using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;
using Xunit.Abstractions;

namespace DocBridge.Core.Tests;

/// <summary>
/// HWP 신규 실무 기능 후보의 Action/속성 존재를 소유 인스턴스에서만 실측한다.
/// GetDefault + IDispatch 열거만 사용하며 Execute/Run(문서 변경)은 FileNew(소유 탭) 외에 호출하지 않는다.
/// DOCBRIDGE_E2E=1일 때만 실행. Category=E2E.
/// </summary>
[Trait("Category", "E2E")]
public sealed class HwpCapabilityProbeTests : IDisposable
{
    private static bool Enabled =>
        string.Equals(Environment.GetEnvironmentVariable("DOCBRIDGE_E2E"), "1", StringComparison.Ordinal);

    private readonly ITestOutputHelper _output;
    private dynamic? _hwp;
    private string _createdTabId = "";
    private HwpE2EOwnershipClaim? _claim;
    private bool _disposed;

    public HwpCapabilityProbeTests(ITestOutputHelper output) => _output = output;

    [ComImport, Guid("00020400-0000-0000-C000-000000000046"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    private interface IDispatchRaw
    {
        void GetTypeInfoCount(out int pctinfo);
        void GetTypeInfo(int iTInfo, int lcid, out System.Runtime.InteropServices.ComTypes.ITypeInfo ppTInfo);
    }

    private static List<string> EnumerateDispatchMembers(object comObject)
    {
        var names = new List<string>();
        System.Runtime.InteropServices.ComTypes.ITypeInfo? typeInfo = null;
        try
        {
            ((IDispatchRaw)comObject).GetTypeInfo(0, 0, out typeInfo);
            if (typeInfo is null) return names;
            typeInfo.GetTypeAttr(out var typeAttrPtr);
            try
            {
                var typeAttr = Marshal.PtrToStructure<TYPEATTR>(typeAttrPtr);
                for (var i = 0; i < typeAttr.cFuncs; i++)
                {
                    typeInfo.GetFuncDesc(i, out var funcDescPtr);
                    try
                    {
                        var funcDesc = Marshal.PtrToStructure<FUNCDESC>(funcDescPtr);
                        var rg = new string[1];
                        typeInfo.GetNames(funcDesc.memid, rg, 1, out var got);
                        if (got > 0 && !string.IsNullOrEmpty(rg[0])) names.Add(rg[0]);
                    }
                    catch { }
                    finally { typeInfo.ReleaseFuncDesc(funcDescPtr); }
                }
                for (var i = 0; i < typeAttr.cVars; i++)
                {
                    typeInfo.GetVarDesc(i, out var varDescPtr);
                    try
                    {
                        var varDesc = Marshal.PtrToStructure<VARDESC>(varDescPtr);
                        var rg = new string[1];
                        typeInfo.GetNames(varDesc.memid, rg, 1, out var got);
                        if (got > 0 && !string.IsNullOrEmpty(rg[0])) names.Add(rg[0]);
                    }
                    catch { }
                    finally { typeInfo.ReleaseVarDesc(varDescPtr); }
                }
            }
            finally { typeInfo.ReleaseTypeAttr(typeAttrPtr); }
        }
        catch { }
        finally { if (typeInfo is not null) Marshal.ReleaseComObject(typeInfo); }
        return names.Distinct().OrderBy(n => n).ToList();
    }

    private void EnsureOwnedApp()
    {
        if (_hwp is not null) return;
        var existing = HwpE2EOwnership.CurrentHwpProcessIds();
        var type = Type.GetTypeFromProgID("HWPFrame.HwpObject")
            ?? throw new InvalidOperationException("HWP not installed");
        dynamic hwp = DocBridge.Core.Services.HwpEnvironmentDoctor.RunWithAutomationWorkingDirectory(
            () => Activator.CreateInstance(type)!)!;
        if (!HwpE2EOwnership.TryReadInventory(hwp, out HwpE2EDocumentSnapshot before))
            throw new InvalidOperationException("probe: before-inventory failed");
        var windowPid = DocBridge.Core.Services.RotHelper.ProcessIdFromWindowHandle(
            DocBridge.Core.Services.RotHelper.HwpWindowHandle((object)hwp));
        var pid = HwpE2EOwnershipPolicy.ResolveProcessId(
            windowPid, existing, HwpE2EOwnership.CurrentHwpProcessIds());
        if (!HwpE2EOwnership.TryFileNew(hwp))
            throw new InvalidOperationException("probe: FileNew failed");
        if (!HwpE2EOwnership.TryReadInventory(hwp, out HwpE2EDocumentSnapshot after))
            throw new InvalidOperationException("probe: after-inventory failed");
        if (!HwpE2EOwnershipPolicy.TryProve(existing, pid, true, before, after,
                out var claim, out var error))
            throw new InvalidOperationException("probe: ownership unproven: " + error);
        _claim = claim;
        if (HwpE2EOwnershipPolicy.TryIdentifyCreatedTab(before, after, out var createdId))
            _createdTabId = createdId;
        _hwp = hwp;
    }

    private static dynamic GetParameterSet(dynamic hwp, string setName) => setName switch
    {
        "HParaShape" => hwp.HParameterSet.HParaShape,
        "HCharShape" => hwp.HParameterSet.HCharShape,
        "HSecDef" => hwp.HParameterSet.HSecDef,
        "HTableCreation" => hwp.HParameterSet.HTableCreation,
        "HNumberingShape" => hwp.HParameterSet.HNumberingShape,
        "HBulletShape" => hwp.HParameterSet.HBulletShape,
        "HFootnoteShape" => hwp.HParameterSet.HFootnoteShape,
        "HMarkpenShape" => hwp.HParameterSet.HMarkpenShape,
        "HHeaderFooter" => hwp.HParameterSet.HHeaderFooter,
        "HPageNumPos" => hwp.HParameterSet.HPageNumPos,
        "HCellBorderFill" => hwp.HParameterSet.HCellBorderFill,
        "HShapeObject" => hwp.HParameterSet.HShapeObject,
        _ => throw new ArgumentOutOfRangeException(nameof(setName)),
    };

    [Fact]
    public void ProbeCandidateActionsAndParameterSets()
    {
        if (!Enabled) return;
        Exception? workerError = null;
        var worker = new Thread(() =>
        {
            try { RunProbe(); }
            catch (Exception ex) { workerError = ex; }
            finally { CloseOwnedSession(); }
        });
        worker.SetApartmentState(ApartmentState.STA);
        worker.Start();
        if (!worker.Join(TimeSpan.FromMinutes(5)))
            throw new TimeoutException("probe: STA worker timed out");
        if (workerError is not null) throw workerError;
    }

    private void RunProbe()
    {
        EnsureOwnedApp();
        dynamic hwp = _hwp!;

        var setNames = new[]
        {
            "HParaShape", "HCharShape", "HSecDef", "HTableCreation",
            "HNumberingShape", "HBulletShape", "HFootnoteShape", "HMarkpenShape",
            "HHeaderFooter", "HPageNumPos", "HCellBorderFill", "HShapeObject",
        };

        foreach (var setName in setNames)
        {
            try
            {
                dynamic ps = GetParameterSet(hwp, setName);
                var members = EnumerateDispatchMembers((object)ps);
                _output.WriteLine($"[SET-OK] {setName} props({members.Count}): {string.Join(" ", members)}");
                HwpE2EOwnership.ReleaseComReference((object)ps);
            }
            catch (Exception ex)
            {
                _output.WriteLine($"[SET-FAIL] {setName}: {ex.Message.Split('\n')[0]}");
            }
        }

        // NOTE: HAction.GetDefault never rejects unknown ActionIDs, so action
        // existence is verified only by Execute + readback per feature (see E2E tests).
        _output.WriteLine("[PROBE-DONE]");
    }

    /// <summary>
    /// 소유 탭 닫기 + 소유 프로세스 Quit까지 STA 워커 스레드에서 수행한다.
    /// 스레드 종료 뒤 MTA에서 호출하면 COM 아파트먼트가 사라져 조용히 실패한다.
    /// </summary>
    private void CloseOwnedSession()
    {
        try
        {
            if (_hwp is not null && !string.IsNullOrEmpty(_createdTabId))
                HwpE2EOwnership.TryCloseDocumentById(_hwp, _createdTabId);
        }
        catch { }
        try
        {
            // 소유 프로세스만 완전히 종료한다. 저장하지 않은 테스트 문서는 버린다.
            // HWP 자동화에 Quit이 없으므로 증명된 소유 PID만 Kill한다(제품 Dispose와 동일 방식).
            var pid = _claim?.ProcessId ?? 0;
            if (_claim?.Mode == HwpE2EOwnershipMode.ExclusiveNewProcess && pid > 0)
            {
                try
                {
                    using var proc = System.Diagnostics.Process.GetProcessById(pid);
                    proc.Kill();
                    proc.WaitForExit(10000);
                }
                catch { }
            }
        }
        catch { }
        _createdTabId = "";
    }

    public void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        CloseOwnedSession();
        try
        {
            if (_hwp is not null && Marshal.IsComObject((object)_hwp))
                Marshal.ReleaseComObject((object)_hwp);
        }
        catch { }
        _hwp = null;
    }
}
