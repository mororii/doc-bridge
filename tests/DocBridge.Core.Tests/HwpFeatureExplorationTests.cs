using System.Runtime.InteropServices;
using Xunit.Abstractions;

namespace DocBridge.Core.Tests;

/// <summary>
/// Wave 1/2 후보 기능의 값 범위와 액션 존재를 소유 탭에서 Execute로 확정한다.
/// 읽기 전용이 아니므로 DOCBRIDGE_E2E=1 + 소유권 증명 필수. Category=E2E.
/// </summary>
[Trait("Category", "E2E")]
public sealed class HwpFeatureExplorationTests : IDisposable
{
    private static bool Enabled =>
        string.Equals(Environment.GetEnvironmentVariable("DOCBRIDGE_E2E"), "1", StringComparison.Ordinal);

    private readonly ITestOutputHelper _output;
    private dynamic? _hwp;
    private string _createdTabId = "";
    private HwpE2EOwnershipClaim? _claim;

    public HwpFeatureExplorationTests(ITestOutputHelper output) => _output = output;

    private void EnsureOwnedApp()
    {
        if (_hwp is not null) return;
        var existing = HwpE2EOwnership.CurrentHwpProcessIds();
        var type = Type.GetTypeFromProgID("HWPFrame.HwpObject")
            ?? throw new InvalidOperationException("HWP not installed");
        dynamic hwp = DocBridge.Core.Services.HwpEnvironmentDoctor.RunWithAutomationWorkingDirectory(
            () => Activator.CreateInstance(type)!)!;
        if (!HwpE2EOwnership.TryReadInventory(hwp, out HwpE2EDocumentSnapshot before))
            throw new InvalidOperationException("explore: before-inventory failed");
        var windowPid = DocBridge.Core.Services.RotHelper.ProcessIdFromWindowHandle(
            DocBridge.Core.Services.RotHelper.HwpWindowHandle((object)hwp));
        var pid = HwpE2EOwnershipPolicy.ResolveProcessId(
            windowPid, existing, HwpE2EOwnership.CurrentHwpProcessIds());
        if (!HwpE2EOwnership.TryFileNew(hwp))
            throw new InvalidOperationException("explore: FileNew failed");
        if (!HwpE2EOwnership.TryReadInventory(hwp, out HwpE2EDocumentSnapshot after))
            throw new InvalidOperationException("explore: after-inventory failed");
        if (!HwpE2EOwnershipPolicy.TryProve(existing, pid, true, before, after,
                out var exploreClaim, out var error))
            throw new InvalidOperationException("explore: ownership unproven: " + error);
        _claim = exploreClaim;
        if (HwpE2EOwnershipPolicy.TryIdentifyCreatedTab(before, after, out var createdId))
            _createdTabId = createdId;
        _hwp = hwp;
    }

    private static string TryRead(dynamic obj, string name)
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

    private static string TryExecute(dynamic hwp, string action, dynamic hSet)
    {
        try { return "ok=" + ((object)hwp.HAction.Execute(action, hSet)).ToString(); }
        catch (Exception ex) { return "THROW:" + ex.Message.Split('\n')[0]; }
    }

    private void SeedText(string text)
    {
        dynamic hwp = _hwp!;
        dynamic act = hwp.HAction;
        dynamic ps = hwp.HParameterSet.HInsertText;
        act.GetDefault("InsertText", ps.HSet);
        ps.Text = text;
        if (!(bool)act.Execute("InsertText", ps.HSet))
            throw new InvalidOperationException("explore: seed InsertText failed");
    }

    private void SelectAll()
    {
        dynamic hwp = _hwp!;
        hwp.HAction.Run("MoveDocBegin");
        hwp.HAction.Run("SelectAll");
    }

    [Fact]
    public void ExploreCharShapeEffects()
    {
        if (!Enabled) return;
        Exception? err = null;
        var t = new Thread(() =>
        {
            try
            {
                EnsureOwnedApp();
                dynamic hwp = _hwp!;
                SeedText("효과검증문단");
                SelectAll();
                dynamic act = hwp.HAction;
                dynamic ps = hwp.HParameterSet.HCharShape;
                act.GetDefault("CharShape", ps.HSet);
                _output.WriteLine("[CHAR-DEFAULT] OutLineType={0} ShadowType={1} ShadowColor={2} Emboss={3} Engrave={4} SmallCaps={5} UseKerning={6}",
                    TryRead(ps, "OutLineType"), TryRead(ps, "ShadowType"), TryRead(ps, "ShadowColor"),
                    TryRead(ps, "Emboss"), TryRead(ps, "Engrave"), TryRead(ps, "SmallCaps"), TryRead(ps, "UseKerning"));
                ps.OutLineType = 1;
                ps.ShadowType = 1;
                ps.ShadowColor = 255;
                ps.Emboss = true;
                ps.SmallCaps = true;
                ps.UseKerning = true;
                _output.WriteLine("[CHAR-EXEC] " + TryExecute(hwp, "CharShape", ps.HSet));
                dynamic ps2 = hwp.HParameterSet.HCharShape;
                act.GetDefault("CharShape", ps2.HSet);
                _output.WriteLine("[CHAR-READBACK] OutLineType={0} ShadowType={1} ShadowColor={2} Emboss={3} SmallCaps={4} UseKerning={5}",
                    TryRead(ps2, "OutLineType"), TryRead(ps2, "ShadowType"), TryRead(ps2, "ShadowColor"),
                    TryRead(ps2, "Emboss"), TryRead(ps2, "SmallCaps"), TryRead(ps2, "UseKerning"));
            }
            catch (Exception ex) { err = ex; }
            finally { CloseOwnedSession(); }
        });
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        if (!t.Join(TimeSpan.FromMinutes(3))) throw new TimeoutException("explore char timeout");
        if (err is not null) throw err;
    }

    [Fact]
    public void ExploreParaLevelAndNumbering()
    {
        if (!Enabled) return;
        Exception? err = null;
        var t = new Thread(() =>
        {
            try
            {
                EnsureOwnedApp();
                dynamic hwp = _hwp!;
                SeedText("개요검증문단");
                SelectAll();
                dynamic act = hwp.HAction;
                dynamic ps = hwp.HParameterSet.HParaShape;
                act.GetDefault("ParagraphShape", ps.HSet);
                _output.WriteLine("[PARA-DEFAULT] Level={0} HeadingType={1} Numbering={2} NumberingID={3} Bullet={4} BulletID={5}",
                    TryRead(ps, "Level"), TryRead(ps, "HeadingType"), TryRead(ps, "Numbering"),
                    TryRead(ps, "NumberingID"), TryRead(ps, "Bullet"), TryRead(ps, "BulletID"));
                ps.Level = 1;
                _output.WriteLine("[PARA-LEVEL-EXEC] " + TryExecute(hwp, "ParagraphShape", ps.HSet));
                dynamic ps2 = hwp.HParameterSet.HParaShape;
                act.GetDefault("ParagraphShape", ps2.HSet);
                _output.WriteLine("[PARA-LEVEL-READBACK] Level={0}", TryRead(ps2, "Level"));
            }
            catch (Exception ex) { err = ex; }
            finally { CloseOwnedSession(); }
        });
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        if (!t.Join(TimeSpan.FromMinutes(3))) throw new TimeoutException("explore para timeout");
        if (err is not null) throw err;
    }

    [Fact]
    public void ExploreCandidateInsertActions()
    {
        if (!Enabled) return;
        var candidates = new[]
        {
            "Footnote", "InsertFootnote", "Endnote", "InsertEndnote",
            "Markpen", "Numbering", "ColumnDefine", "MultiColumn",
            "Caption", "InsertCaption", "TextBox", "InsertTextBox",
            "EquationCreate", "InsertEquation", "Style", "ApplyStyle",
        };
        Exception? err = null;
        var t = new Thread(() =>
        {
            try
            {
                EnsureOwnedApp();
                dynamic hwp = _hwp!;
                SeedText("액션후보검증");
                hwp.HAction.Run("MoveDocEnd");
                dynamic act = hwp.HAction;
                foreach (var action in candidates)
                {
                    dynamic ps = hwp.HParameterSet.HShapeObject;
                    string result;
                    try
                    {
                        act.GetDefault(action, ps.HSet);
                        result = TryExecute(hwp, action, ps.HSet);
                    }
                    catch (Exception ex) { result = "GETDEFAULT-THROW:" + ex.Message.Split('\n')[0]; }
                    _output.WriteLine($"[CANDIDATE] {action}: {result}");
                    try { hwp.HAction.Run("Cancel"); } catch { }
                }
            }
            catch (Exception ex) { err = ex; }
            finally { CloseOwnedSession(); }
        });
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        if (!t.Join(TimeSpan.FromMinutes(5))) throw new TimeoutException("explore candidates timeout");
        if (err is not null) throw err;
    }

    [Fact]
    public void ExploreSectionAndTableProperties()
    {
        if (!Enabled) return;
        Exception? err = null;
        var t = new Thread(() =>
        {
            try
            {
                EnsureOwnedApp();
                dynamic hwp = _hwp!;
                SeedText("구역표검증");
                dynamic act = hwp.HAction;
                dynamic sec = hwp.HParameterSet.HSecDef;
                act.GetDefault("PageSetup", sec.HSet);
                _output.WriteLine("[SEC-DEFAULT] ShowLineNumbers={0} LineNumberStart={1} LineNumberCountBy={2} LineNumberDistance={3} LineNumberRestart={4} SpaceBetweenColumns={5} PageBorderFillBoth={6} WongojiFormat={7}",
                    TryRead(sec, "ShowLineNumbers"), TryRead(sec, "LineNumberStart"), TryRead(sec, "LineNumberCountBy"),
                    TryRead(sec, "LineNumberDistance"), TryRead(sec, "LineNumberRestart"),
                    TryRead(sec, "SpaceBetweenColumns"), TryRead(sec, "PageBorderFillBoth"), TryRead(sec, "WongojiFormat"));
                sec.ShowLineNumbers = true;
                _output.WriteLine("[SEC-LINENUM-EXEC] " + TryExecute(hwp, "PageSetup", sec.HSet));
                dynamic sec2 = hwp.HParameterSet.HSecDef;
                act.GetDefault("PageSetup", sec2.HSet);
                _output.WriteLine("[SEC-LINENUM-READBACK] ShowLineNumbers={0}", TryRead(sec2, "ShowLineNumbers"));

                dynamic create = hwp.HParameterSet.HTableCreation;
                act.GetDefault("TableCreate", create.HSet);
                create.Rows = 3;
                create.Cols = 2;
                _output.WriteLine("[TABLE-CREATE] " + TryExecute(hwp, "TableCreate", create.HSet));
                hwp.HAction.Run("TableCellBlockExtend");
                dynamic shape = hwp.HParameterSet.HShapeObject;
                act.GetDefault("TablePropertyDialog", shape.HSet);
                _output.WriteLine("[TABLEPROP-DEFAULT] RepeatHeader={0}", TryRead(shape, "RepeatHeader"));
                shape.RepeatHeader = true;
                _output.WriteLine("[TABLEPROP-REPEAT-EXEC] " + TryExecute(hwp, "TablePropertyDialog", shape.HSet));
            }
            catch (Exception ex) { err = ex; }
            finally { CloseOwnedSession(); }
        });
        t.SetApartmentState(ApartmentState.STA);
        t.Start();
        if (!t.Join(TimeSpan.FromMinutes(5))) throw new TimeoutException("explore section/table timeout");
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
