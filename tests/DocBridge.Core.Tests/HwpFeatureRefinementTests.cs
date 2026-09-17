using System.Runtime.InteropServices;
using Xunit.Abstractions;

namespace DocBridge.Core.Tests;

/// <summary>형광펜/다단/번호ID 후보를 올바른 세트+선택 영역으로 재검증한다. E2E 전용.</summary>
[Trait("Category", "E2E")]
public sealed class HwpFeatureRefinementTests : IDisposable
{
    private static bool Enabled =>
        string.Equals(Environment.GetEnvironmentVariable("DOCBRIDGE_E2E"), "1", StringComparison.Ordinal);

    private readonly ITestOutputHelper _output;
    private dynamic? _hwp;
    private string _createdTabId = "";
    private HwpE2EOwnershipClaim? _claim;

    public HwpFeatureRefinementTests(ITestOutputHelper output) => _output = output;

    private void EnsureOwnedApp()
    {
        if (_hwp is not null) return;
        var existing = HwpE2EOwnership.CurrentHwpProcessIds();
        var type = Type.GetTypeFromProgID("HWPFrame.HwpObject")
            ?? throw new InvalidOperationException("HWP not installed");
        dynamic hwp = DocBridge.Core.Services.HwpEnvironmentDoctor.RunWithAutomationWorkingDirectory(
            () => Activator.CreateInstance(type)!)!;
        if (!HwpE2EOwnership.TryReadInventory(hwp, out HwpE2EDocumentSnapshot before))
            throw new InvalidOperationException("refine: before-inventory failed");
        var windowPid = DocBridge.Core.Services.RotHelper.ProcessIdFromWindowHandle(
            DocBridge.Core.Services.RotHelper.HwpWindowHandle((object)hwp));
        var pid = HwpE2EOwnershipPolicy.ResolveProcessId(
            windowPid, existing, HwpE2EOwnership.CurrentHwpProcessIds());
        if (!HwpE2EOwnership.TryFileNew(hwp))
            throw new InvalidOperationException("refine: FileNew failed");
        if (!HwpE2EOwnership.TryReadInventory(hwp, out HwpE2EDocumentSnapshot after))
            throw new InvalidOperationException("refine: after-inventory failed");
        if (!HwpE2EOwnershipPolicy.TryProve(existing, pid, true, before, after,
                out var refineClaim, out var error))
            throw new InvalidOperationException("refine: ownership unproven: " + error);
        _claim = refineClaim;
        if (HwpE2EOwnershipPolicy.TryIdentifyCreatedTab(before, after, out var createdId))
            _createdTabId = createdId;
        _hwp = hwp;
    }

    [Fact]
    public void RefineMarkpenColumnsNumbering()
    {
        if (!Enabled) return;
        Exception? err = null;
        var t = new Thread(() =>
        {
            try
            {
                EnsureOwnedApp();
                dynamic hwp = _hwp!;
                dynamic act = hwp.HAction;
                dynamic ins = hwp.HParameterSet.HInsertText;
                act.GetDefault("InsertText", ins.HSet);
                ins.Text = "형광펜다단번호검증문장";
                if (!(bool)act.Execute("InsertText", ins.HSet))
                    throw new InvalidOperationException("refine: seed failed");
                hwp.HAction.Run("MoveDocBegin");
                hwp.HAction.Run("SelectAll");

                foreach (var action in new[] { "Markpen", "MarkpenShape", "Highlight" })
                {
                    try
                    {
                        dynamic ps = hwp.HParameterSet.HMarkpenShape;
                        act.GetDefault(action, ps.HSet);
                        try { ps.Color = 65535; } catch (Exception ex) { _output.WriteLine($"[MARKPEN] {action} Color-set: {ex.Message.Split('\n')[0]}"); }
                        bool ok;
                        try { ok = (bool)act.Execute(action, ps.HSet); }
                        catch (Exception ex) { _output.WriteLine($"[MARKPEN] {action}: EXEC-THROW {ex.Message.Split('\n')[0]}"); continue; }
                        _output.WriteLine($"[MARKPEN] {action}: ok={ok}");
                    }
                    catch (Exception ex) { _output.WriteLine($"[MARKPEN] {action}: SET-THROW {ex.Message.Split('\n')[0]}"); }
                    try { hwp.HAction.Run("Cancel"); } catch { }
                }

                foreach (var action in new[] { "Column", "Columns", "SectionColumn", "DanDefine", "ColumnDefine2" })
                {
                    try
                    {
                        dynamic ps = hwp.HParameterSet.HSecDef;
                        act.GetDefault(action, ps.HSet);
                        bool ok;
                        try { ok = (bool)act.Execute(action, ps.HSet); }
                        catch (Exception ex) { _output.WriteLine($"[COLUMN] {action}: EXEC-THROW {ex.Message.Split('\n')[0]}"); continue; }
                        _output.WriteLine($"[COLUMN] {action}: ok={ok}");
                    }
                    catch (Exception ex) { _output.WriteLine($"[COLUMN] {action}: SET-THROW {ex.Message.Split('\n')[0]}"); }
                    try { hwp.HAction.Run("Cancel"); } catch { }
                }

                foreach (var numberingId in new[] { 1, 2 })
                {
                    try
                    {
                        dynamic ps = hwp.HParameterSet.HParaShape;
                        act.GetDefault("ParagraphShape", ps.HSet);
                        ps.NumberingID = numberingId;
                        bool ok;
                        try { ok = (bool)act.Execute("ParagraphShape", ps.HSet); }
                        catch (Exception ex) { _output.WriteLine($"[NUMBERING] id={numberingId}: EXEC-THROW {ex.Message.Split('\n')[0]}"); continue; }
                        dynamic ps2 = hwp.HParameterSet.HParaShape;
                        act.GetDefault("ParagraphShape", ps2.HSet);
                        object? rb = null;
                        try { rb = ps2.NumberingID; } catch { }
                        _output.WriteLine($"[NUMBERING] id={numberingId}: ok={ok} readback={rb}");
                    }
                    catch (Exception ex) { _output.WriteLine($"[NUMBERING] id={numberingId}: SET-THROW {ex.Message.Split('\n')[0]}"); }
                    try { hwp.HAction.Run("Cancel"); } catch { }
                }
            }
            catch (Exception ex) { err = ex; }
            finally { CloseOwnedSession(); }
        });
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        if (!t.Join(TimeSpan.FromMinutes(5))) throw new TimeoutException("refine timeout");
        if (err is not null) throw err;
    }

    private static int CountControl(dynamic hwp, string controlId)
    {
        try
        {
            dynamic? ctrl = null;
            try { ctrl = hwp.HeadCtrl; } catch { }
            var count = 0;
            while (ctrl is not null)
            {
                try
                {
                    if (string.Equals(Convert.ToString(ctrl.CtrlID) ?? "",
                            controlId, StringComparison.OrdinalIgnoreCase))
                        count++;
                }
                catch { }
                try { ctrl = ctrl.Next; } catch { ctrl = null; }
            }
            return count;
        }
        catch { return -999; }
    }

    private static string ListControlIds(dynamic hwp)
    {
        try
        {
            var ids = new List<string>();
            dynamic? ctrl = null;
            try { ctrl = hwp.HeadCtrl; } catch { }
            var guard = 0;
            while (ctrl is not null && guard++ < 200)
            {
                try { ids.Add(Convert.ToString(ctrl.CtrlID) ?? "?"); } catch { ids.Add("?"); }
                try { ctrl = ctrl.Next; } catch { ctrl = null; }
            }
            return string.Join(",", ids.Distinct());
        }
        catch (Exception ex) { return "ERR:" + ex.Message.Split('\n')[0]; }
    }

    [Fact]
    public void ExploreFootnoteFlow()
    {
        if (!Enabled) return;
        Exception? err = null;
        var t = new Thread(() =>
        {
            try
            {
                EnsureOwnedApp();
                dynamic hwp = _hwp!;
                dynamic act = hwp.HAction;
                dynamic ins = hwp.HParameterSet.HInsertText;
                act.GetDefault("InsertText", ins.HSet);
                ins.Text = "각주흐름검증본문";
                if (!(bool)act.Execute("InsertText", ins.HSet))
                    throw new InvalidOperationException("refine: seed failed");
                foreach (var id in new[] { "fn", "en", "footnote", "endnote", "note" })
                    _output.WriteLine($"[FN-BEFORE] {id}={CountControl(hwp, id)}");
                dynamic fns = hwp.HParameterSet.HFootnoteShape;
                act.GetDefault("InsertFootnote", fns.HSet);
                var ok = (bool)act.Execute("InsertFootnote", fns.HSet);
                _output.WriteLine($"[FN-EXEC] ok={ok}");
                foreach (var id in new[] { "fn", "en", "footnote", "endnote", "note" })
                    _output.WriteLine($"[FN-AFTER] {id}={CountControl(hwp, id)}");
                dynamic ins2 = hwp.HParameterSet.HInsertText;
                act.GetDefault("InsertText", ins2.HSet);
                ins2.Text = "각주내용123";
                _output.WriteLine($"[FN-TEXT] ok={act.Execute("InsertText", ins2.HSet)}");
                try { hwp.HAction.Run("MoveDocBegin"); } catch { }
                _output.WriteLine("[FN-CTRLS] " + ListControlIds(hwp));
                foreach (var id in new[] { "fn", "en", "footnote", "endnote", "note" })
                    _output.WriteLine($"[FN-FINAL] {id}={CountControl(hwp, id)}");
                try
                {
                    hwp.HAction.Run("MoveDocBegin");
                    hwp.HAction.Run("SelectAll");
                    _output.WriteLine("[FN-DOC-TEXT-LEN] " + ((string)hwp.GetSelectedText()).Length);
                }
                catch (Exception ex) { _output.WriteLine("[FN-DOC-TEXT] ERR " + ex.Message.Split('\n')[0]); }
            }
            catch (Exception ex) { err = ex; }
            finally { CloseOwnedSession(); }
        });
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        if (!t.Join(TimeSpan.FromMinutes(5))) throw new TimeoutException("refine footnote timeout");
        if (err is not null) throw err;
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

    private static string TryReadProp(dynamic obj, string name)
    {
        try
        {
            var value = obj.GetType().InvokeMember(
                name,
                System.Reflection.BindingFlags.GetProperty,
                null, obj, null);
            return value?.ToString() ?? "<null>";
        }
        catch (Exception ex) { return "<ERR:" + ex.Message.Split('\n')[0] + ">"; }
    }

    [Fact]
    public void ExploreTableBorderWidths()
    {
        if (!Enabled) return;
        Exception? err = null;
        var t = new Thread(() =>
        {
            try
            {
                EnsureOwnedApp();
                dynamic hwp = _hwp!;
                dynamic act = hwp.HAction;
                dynamic create = hwp.HParameterSet.HTableCreation;
                act.GetDefault("TableCreate", create.HSet);
                create.Rows = 3;
                create.Cols = 3;
                if (!(bool)act.Execute("TableCreate", create.HSet))
                    throw new InvalidOperationException("border: table create failed");
                dynamic border = hwp.HParameterSet.HCellBorderFill;
                act.GetDefault("CellBorderFill", border.HSet);
                _output.WriteLine("[BORDER-DEFAULTS] TypeTop={0} WidthTop={1} ColorTop={2} ApplyToEdge={3}",
                    TryReadProp(border, "BorderTypeTop"), TryReadProp(border, "BorderWidthTop"),
                    TryReadProp(border, "BorderColorTop"), TryReadProp(border, "ApplyBorderToEdge"));
                try
                {
                    _output.WriteLine("[BORDER-LINE] Solid={0} W012={1} W05={2} W10={3}",
                        hwp.HwpLineType("Solid"), hwp.HwpLineWidth("0.12mm"),
                        hwp.HwpLineWidth("0.5mm"), hwp.HwpLineWidth("1.0mm"));
                }
                catch (Exception ex) { _output.WriteLine("[BORDER-LINE] ERR " + ex.Message.Split('\n')[0]); }
                hwp.HAction.Run("TableCellBlock");
                hwp.HAction.Run("TableCellBlockExtend");
                hwp.HAction.Run("TableLowerCell");
                hwp.HAction.Run("TableLowerCell");
                hwp.HAction.Run("TableRightCell");
                hwp.HAction.Run("TableRightCell");
                dynamic b2 = hwp.HParameterSet.HCellBorderFill;
                act.GetDefault("CellBorderFill", b2.HSet);
                try
                {
                    b2.BorderTypeTop = hwp.HwpLineType("Solid");
                    b2.BorderTypeBottom = hwp.HwpLineType("Solid");
                    b2.BorderTypeLeft = hwp.HwpLineType("Solid");
                    b2.BorderTypeRight = hwp.HwpLineType("Solid");
                    b2.BorderWidthTop = hwp.HwpLineWidth("0.5mm");
                    b2.BorderWidthBottom = hwp.HwpLineWidth("0.5mm");
                    b2.BorderWidthLeft = hwp.HwpLineWidth("0.5mm");
                    b2.BorderWidthRight = hwp.HwpLineWidth("0.5mm");
                    _output.WriteLine("[BORDER-EXEC] ok=" + act.Execute("CellBorderFill", b2.HSet));
                }
                catch (Exception ex) { _output.WriteLine("[BORDER-EXEC] ERR " + ex.Message.Split('\n')[0]); }
                try { hwp.HAction.Run("Cancel"); } catch { }
                dynamic b3 = hwp.HParameterSet.HCellBorderFill;
                act.GetDefault("CellBorderFill", b3.HSet);
                _output.WriteLine("[BORDER-READBACK] TypeTop={0} WidthTop={1}",
                    TryReadProp(b3, "BorderTypeTop"), TryReadProp(b3, "BorderWidthTop"));
            }
            catch (Exception ex) { err = ex; }
            finally { CloseOwnedSession(); }
        });
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        if (!t.Join(TimeSpan.FromMinutes(5))) throw new TimeoutException("border explore timeout");
        if (err is not null) throw err;
    }

    private static string PosKey(dynamic hwp)
    {
        try
        {
            dynamic set = hwp.CreateSet("ListParaPos");
            if (!(bool)hwp.GetPosBySet(set)) return "<nopos>";
            return $"{set.Item("List")}:{set.Item("Para")}:{set.Item("Pos")}";
        }
        catch (Exception ex) { return "<ERR:" + ex.Message.Split('\n')[0] + ">"; }
    }

    [Fact]
    public void ExploreTableNavigationBounds()
    {
        if (!Enabled) return;
        Exception? err = null;
        var t = new Thread(() =>
        {
            try
            {
                EnsureOwnedApp();
                dynamic hwp = _hwp!;
                dynamic act = hwp.HAction;
                dynamic create = hwp.HParameterSet.HTableCreation;
                act.GetDefault("TableCreate", create.HSet);
                create.Rows = 3;
                create.Cols = 4;
                if (!(bool)act.Execute("TableCreate", create.HSet))
                    throw new InvalidOperationException("nav: table create failed");
                hwp.HAction.Run("MoveDocBegin");
                // go to (0,0): move into table then up/left to edges
                hwp.HAction.Run("TableCellBlock");
                for (var i = 0; i < 5; i++) { try { hwp.HAction.Run("TableUpperCell"); } catch { } }
                for (var i = 0; i < 6; i++) { try { hwp.HAction.Run("TableLeftCell"); } catch { } }
                _output.WriteLine("[NAV-START] key=" + PosKey(hwp));
                for (var i = 0; i < 5; i++)
                {
                    bool ok;
                    try { ok = (bool)hwp.HAction.Run("TableLowerCell"); }
                    catch (Exception ex) { _output.WriteLine($"[NAV-DOWN] {i}: THROW {ex.Message.Split('\n')[0]}"); break; }
                    _output.WriteLine($"[NAV-DOWN] {i}: ok={ok} key=" + PosKey(hwp));
                    if (!ok) break;
                }
                for (var i = 0; i < 5; i++) { try { hwp.HAction.Run("TableUpperCell"); } catch { } }
                for (var i = 0; i < 6; i++) { try { hwp.HAction.Run("TableLeftCell"); } catch { } }
                for (var i = 0; i < 6; i++)
                {
                    bool ok;
                    try { ok = (bool)hwp.HAction.Run("TableRightCell"); }
                    catch (Exception ex) { _output.WriteLine($"[NAV-RIGHT] {i}: THROW {ex.Message.Split('\n')[0]}"); break; }
                    _output.WriteLine($"[NAV-RIGHT] {i}: ok={ok} key=" + PosKey(hwp));
                    if (!ok) break;
                }
                // wrap-up test from (1,0)
                for (var i = 0; i < 6; i++) { try { hwp.HAction.Run("TableLeftCell"); } catch { } }
                try { hwp.HAction.Run("TableLowerCell"); } catch { }
                _output.WriteLine("[NAV-ROW1] key=" + PosKey(hwp));
                try
                {
                    var ok = (bool)hwp.HAction.Run("TableLeftCell");
                    _output.WriteLine("[NAV-WRAPUP] ok=" + ok + " key=" + PosKey(hwp));
                }
                catch (Exception ex) { _output.WriteLine("[NAV-WRAPUP] THROW " + ex.Message.Split('\n')[0]); }
            }
            catch (Exception ex) { err = ex; }
            finally { CloseOwnedSession(); }
        });
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        if (!t.Join(TimeSpan.FromMinutes(5))) throw new TimeoutException("nav explore timeout");
        if (err is not null) throw err;
    }

    public void Dispose()
    {
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
