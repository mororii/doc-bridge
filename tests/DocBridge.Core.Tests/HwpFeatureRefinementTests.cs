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
                out _, out var error))
            throw new InvalidOperationException("refine: ownership unproven: " + error);
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
        });
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        if (!t.Join(TimeSpan.FromMinutes(5))) throw new TimeoutException("refine footnote timeout");
        if (err is not null) throw err;
    }

    public void Dispose()
    {
        try
        {
            if (_hwp is not null && !string.IsNullOrEmpty(_createdTabId))
                HwpE2EOwnership.TryCloseDocumentById(_hwp, _createdTabId);
        }
        catch { }
        try
        {
            if (_hwp is not null && Marshal.IsComObject((object)_hwp))
                Marshal.ReleaseComObject((object)_hwp);
        }
        catch { }
        _hwp = null;
    }
}
