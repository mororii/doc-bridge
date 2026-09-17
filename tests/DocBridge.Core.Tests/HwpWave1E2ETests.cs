using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

/// <summary>
/// Wave 1 신규/확장 HWP op의 제품 경로(dry-run → confirmToken → apply → readback) 검증.
/// DOCBRIDGE_E2E=1일 때만 실행. 소유 탭만 건드리며 종료 시 소유 문서만 닫는다.
/// </summary>
[Trait("Category", "E2E")]
public sealed class HwpWave1E2ETests : IDisposable
{
    private static bool Enabled =>
        string.Equals(Environment.GetEnvironmentVariable("DOCBRIDGE_E2E"), "1", StringComparison.Ordinal);

    private readonly TestHome _home = new();
    private HwpAdapter? _adapter;
    private object? _createdApp;
    private HwpE2EOwnershipClaim? _ownership;
    private bool _sessionReleased;
    private bool _adapterDisposed;

    private DocBridgeHost CreateHostWithHwp()
    {
        var existingProcessIds = HwpE2EOwnership.CurrentHwpProcessIds();
        _adapter = new HwpAdapter(() =>
        {
            var type = Type.GetTypeFromProgID("HWPFrame.HwpObject")
                ?? throw new InvalidOperationException("HWP not installed");
            dynamic hwp = HwpEnvironmentDoctor.RunWithAutomationWorkingDirectory(
                () => Activator.CreateInstance(type)!)!;
            if (!HwpE2EOwnership.TryReadInventory(hwp, out HwpE2EDocumentSnapshot before))
                throw new InvalidOperationException("wave1: before-inventory failed");
            var windowProcessId = RotHelper.ProcessIdFromWindowHandle(RotHelper.HwpWindowHandle(hwp));
            var processId = HwpE2EOwnershipPolicy.ResolveProcessId(
                windowProcessId, existingProcessIds, HwpE2EOwnership.CurrentHwpProcessIds());
            if (!HwpE2EOwnership.TryFileNew(hwp))
                throw new InvalidOperationException("wave1: FileNew failed");
            if (!HwpE2EOwnership.TryReadInventory(hwp, out HwpE2EDocumentSnapshot after))
                throw new InvalidOperationException("wave1: after-inventory failed");
            if (!HwpE2EOwnershipPolicy.TryProve(existingProcessIds, processId, true, before, after,
                    out HwpE2EOwnershipClaim claim, out string error))
            {
                if (HwpE2EOwnershipPolicy.TryIdentifyCreatedTab(before, after, out var createdId))
                    HwpE2EOwnership.TryCloseDocumentById(hwp, createdId);
                throw new InvalidOperationException(error);
            }
            _ownership = claim;
            _createdApp = hwp;
            HwpE2EOwnership.InsertSeedText(hwp, HwpE2EOwnershipPolicy.SeedText);
            return (object)hwp;
        });
        var host = new DocBridgeHost(_home.Options);
        host.Router.Register("hwp", new HwpE2ESafeAdapter(_adapter, ReleaseOwnedSession));
        return host;
    }

    private static JsonObject ApplyBatch(DocBridgeHost host, JsonArray ops)
    {
        var dry = host.ApplyOps("hwp", new JsonObject { ["ops"] = ops.DeepClone(), ["dryRun"] = true });
        Assert.True(Json.GetBool(dry, "ok"), $"dry-run failed: {dry}");
        var applied = host.ApplyOps("hwp", new JsonObject
        {
            ["ops"] = ops.DeepClone(),
            ["dryRun"] = false,
            ["confirmToken"] = Json.GetString(dry, "confirmToken"),
        });
        Assert.True(Json.GetBool(applied, "ok"), $"apply failed: {applied}");
        Assert.True(Json.GetBool(Json.GetObj(applied, "readback"), "verified"),
            $"readback unverified: {applied}");
        return applied;
    }

    private static void SelectFirstMatch(dynamic hwp, string text)
    {
        hwp.HAction.Run("MoveDocBegin");
        dynamic act = hwp.HAction;
        dynamic find = hwp.HParameterSet.HFindReplace;
        act.GetDefault("FindDlg", find.HSet);
        _ = act.Execute("FindDlg", find.HSet);
        find = hwp.HParameterSet.HFindReplace;
        try { find.MatchCase = 1; } catch { }
        try { find.Direction = hwp.FindDir("Forward"); } catch { }
        find.FindString = text;
        try { find.IgnoreMessage = 1; } catch { }
        try { find.FindRegExp = 0; } catch { }
        try { find.FindType = 1; } catch { }
        if (!(bool)act.Execute("RepeatFind", find.HSet))
            throw new InvalidOperationException("wave1: seed text not found: " + text);
    }

    [Fact]
    public void Char_effects_and_para_level_apply_with_readback()
    {
        if (!Enabled) return;
        using var host = CreateHostWithHwp();
        ApplyBatch(host, new JsonArray
        {
            new JsonObject { ["op"] = "append_text", ["text"] = "효과검증문장\n개요수준문장" },
        });
        ApplyBatch(host, new JsonArray
        {
            new JsonObject
            {
                ["op"] = "set_paragraph_style_basic",
                ["target"] = new JsonObject { ["text"] = "효과검증문장" },
                ["style"] = new JsonObject
                {
                    ["outline"] = true, ["shadow"] = true, ["shadowColor"] = "#0000FF",
                    ["emboss"] = true, ["smallCaps"] = true, ["kerning"] = true,
                },
            },
        });
        ApplyBatch(host, new JsonArray
        {
            new JsonObject
            {
                ["op"] = "set_paragraph_style_basic",
                ["target"] = new JsonObject { ["text"] = "개요수준문장" },
                ["style"] = new JsonObject { ["engrave"] = true },
            },
        });
        ApplyBatch(host, new JsonArray
        {
            new JsonObject
            {
                ["op"] = "set_paragraph_format",
                ["target"] = new JsonObject { ["text"] = "개요수준문장" },
                ["style"] = new JsonObject { ["level"] = 2 },
            },
        });
        _adapter!.RunOnAdapterThread<object?>(() =>
        {
            dynamic hwp = _createdApp!;
            SelectFirstMatch(hwp, "효과검증문장");
            dynamic act = hwp.HAction;
            dynamic ps = hwp.HParameterSet.HCharShape;
            act.GetDefault("CharShape", ps.HSet);
            Assert.Equal(1, Convert.ToInt32(ps.OutLineType));
            Assert.Equal(1, Convert.ToInt32(ps.ShadowType));
            Assert.Equal(1, Convert.ToInt32(ps.Emboss));
            Assert.Equal(1, Convert.ToInt32(ps.SmallCaps));
            Assert.Equal(1, Convert.ToInt32(ps.UseKerning));
            SelectFirstMatch(hwp, "개요수준문장");
            dynamic ps2 = hwp.HParameterSet.HCharShape;
            act.GetDefault("CharShape", ps2.HSet);
            Assert.Equal(1, Convert.ToInt32(ps2.Engrave));
            dynamic para = hwp.HParameterSet.HParaShape;
            act.GetDefault("ParagraphShape", para.HSet);
            Assert.Equal(2, Convert.ToInt32(para.Level));
            return null;
        });
    }

    [Fact]
    public void Footnote_and_endnote_insert_with_control_readback()
    {
        if (!Enabled) return;
        using var host = CreateHostWithHwp();
        ApplyBatch(host, new JsonArray
        {
            new JsonObject { ["op"] = "append_text", ["text"] = "각주대상문장" },
        });
        var footnote = ApplyBatch(host, new JsonArray
        {
            new JsonObject
            {
                ["op"] = "insert_footnote",
                ["target"] = new JsonObject { ["text"] = "각주대상문장" },
                ["text"] = "각주내용검증",
            },
        });
        Assert.True(Json.GetBool(footnote, "ok"));
        var endnote = ApplyBatch(host, new JsonArray
        {
            new JsonObject { ["op"] = "insert_endnote", ["text"] = "미주내용검증" },
        });
        Assert.True(Json.GetBool(endnote, "ok"));
        _adapter!.RunOnAdapterThread<object?>(() =>
        {
            dynamic hwp = _createdApp!;
            Assert.Equal(1, CountControls(hwp, "fn"));
            Assert.Equal(1, CountControls(hwp, "en"));
            return null;
        });
    }

    [Fact]
    public void Repeat_header_and_line_numbers_apply_with_readback()
    {
        if (!Enabled) return;
        using var host = CreateHostWithHwp();
        ApplyBatch(host, new JsonArray
        {
            new JsonObject
            {
                ["op"] = "insert_table",
                ["rows"] = new JsonArray
                {
                    new JsonArray("항목", "내용"),
                    new JsonArray("검측", "실시"),
                    new JsonArray("결과", "적합"),
                },
                ["header"] = true,
            },
        });
        var repeatOff = ApplyBatch(host, new JsonArray
        {
            new JsonObject { ["op"] = "table_set_repeat_header", ["tableIndex"] = 0, ["repeat"] = false },
        });
        Assert.True(Json.GetBool(repeatOff, "ok"), $"repeat-off apply failed: {repeatOff}");
        var repeat = ApplyBatch(host, new JsonArray
        {
            new JsonObject { ["op"] = "table_set_repeat_header", ["tableIndex"] = 0, ["repeat"] = true },
        });
        Assert.True(Json.GetBool(repeat, "ok"), $"repeat apply failed: {repeat}");
        ApplyBatch(host, new JsonArray
        {
            new JsonObject
            {
                ["op"] = "set_page_setup",
                ["page"] = new JsonObject { ["lineNumbers"] = true, ["lineNumberStart"] = 5 },
            },
        });
        _adapter!.RunOnAdapterThread<object?>(() =>
        {
            dynamic hwp = _createdApp!;
            // 표 개체 재조회는 제품 readback(off→on 토글 검증됨)에 맡기고,
            // 여기서는 선택 없이 읽을 수 있는 구역(줄번호) 값만 원시 확인한다.
            dynamic sec = hwp.HParameterSet.HSecDef;
            hwp.HAction.GetDefault("PageSetup", sec.HSet);
            Assert.True(Convert.ToBoolean(sec.ShowLineNumbers));
            Assert.Equal(5, Convert.ToInt32(sec.LineNumberStart));
            return null;
        });
    }

    private static int CountControls(dynamic hwp, string controlId)
    {
        dynamic? ctrl = null;
        try { ctrl = hwp.HeadCtrl; } catch { }
        var count = 0;
        var guard = 0;
        while (ctrl is not null && guard++ < 500)
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

    public void Dispose()
    {
        ReleaseOwnedSession();
        if (!_adapterDisposed)
        {
            HwpE2EOwnership.DisposeAdapterWithoutKilling(_adapter);
            _adapterDisposed = true;
        }
        _adapter = null;
        _createdApp = null;
        GC.Collect();
        GC.WaitForPendingFinalizers();
        _home.Dispose();
    }

    private void ReleaseOwnedSession()
    {
        if (_sessionReleased) return;
        _sessionReleased = true;
        if (_adapter is null || _createdApp is null || _ownership is null) return;
        try
        {
            _adapter.RunOnAdapterThread<object?>(() =>
            {
                foreach (var id in _ownership.OwnedDocumentIds)
                    HwpE2EOwnership.TryCloseDocumentById(_createdApp, id);
                return null;
            });
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("HwpWave1E2E: owned-session release reported: " + ex.Message);
        }
    }
}
