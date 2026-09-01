using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class HwpAdapter
{
    /// <summary>
    /// GetTextFile("HWP") 전체 직렬화는 내용이 같아도 내부 비결정 값이 달라질 수 있다.
    /// Preview가 실제로 의존하는 본문·선택·커서·control/field 구조·현재 서식을
    /// 결정적인 JSON 순서로 묶어 apply 전 상태 동일성을 검증한다.
    /// </summary>
    private static string CapturePreviewFingerprint(dynamic hwp, string? knownText = null)
    {
        var text = NormalizeNewlines(knownText ?? GetDocText(hwp));
        var selection = NormalizeNewlines(GetSelectionText(hwp));
        var controls = new JsonArray();
        var control = hwp.HeadCtrl;
        var scanned = 0;
        while (control is not null && scanned < 5000)
        {
            try { controls.Add(Convert.ToString(control.CtrlID) ?? ""); }
            catch { controls.Add("?"); }
            try { control = control.Next; }
            catch { break; }
            scanned++;
        }

        var position = new JsonObject();
        try
        {
            dynamic pos = hwp.CreateSet("ListParaPos");
            if ((bool)hwp.GetPosBySet(pos))
            {
                position["list"] = DynamicInt((object)pos, "List");
                position["para"] = DynamicInt((object)pos, "Para");
                position["pos"] = DynamicInt((object)pos, "Pos");
            }
        }
        catch { }

        JsonObject? caretStyle = null;
        try { caretStyle = NativeStyleSummary(CaptureCurrentNativeStyle(hwp, "preview-fingerprint")); }
        catch { }
        string fields = "";
        try { fields = Convert.ToString(hwp.GetFieldList(0, 0)) ?? ""; }
        catch { }

        var state = new JsonObject
        {
            ["documentIdentity"] = CaptureDocumentIdentity(hwp),
            ["textSha256"] = Sha(text),
            ["textLength"] = text.Length,
            ["selectionSha256"] = Sha(selection),
            ["selectionLength"] = selection.Length,
            ["position"] = position,
            ["controls"] = controls,
            ["fieldsSha256"] = Sha(fields),
            ["caretStyle"] = caretStyle,
        };
        return Sha(Json.Canonical(state));
    }

    private static string Sha(string value) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(value))).ToLowerInvariant();
}
