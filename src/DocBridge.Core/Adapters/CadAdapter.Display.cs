using System.Diagnostics;
using System.Text.Json.Nodes;
using DocBridge.Core.Models;

namespace DocBridge.Core.Adapters;

public sealed partial class CadAdapter
{
    private static string? CurrentLayerName(dynamic doc)
    {
        try { return (string)doc.ActiveLayer.Name; }
        catch { return null; }
    }

    // Unknown COM properties are null, never silently reported as off/unlocked.
    private static JsonObject LayerState(object layerObject, string? currentLayer)
    {
        dynamic layer = layerObject;
        var item = new JsonObject();
        var unavailable = new JsonArray();
        void Read(string key, Func<JsonNode?> read)
        {
            try { item[key] = read(); }
            catch { item[key] = null; unavailable.Add(key); }
        }
        Read("name", () => JsonValue.Create((string)layer.Name));
        Read("on", () => JsonValue.Create((bool)layer.LayerOn));
        Read("freeze", () => JsonValue.Create((bool)layer.Freeze));
        Read("locked", () => JsonValue.Create((bool)layer.Lock));
        Read("plottable", () => JsonValue.Create((bool)layer.Plottable));
        Read("color", () => JsonValue.Create((int)layer.Color));
        Read("linetype", () => JsonValue.Create((string)layer.Linetype));
        item["current"] = currentLayer is null || item["name"] is null ? null :
            string.Equals(item["name"]!.GetValue<string>(), currentLayer, StringComparison.OrdinalIgnoreCase);
        item["modelVisible"] = item["on"] is null || item["freeze"] is null ? null :
            item["on"]!.GetValue<bool>() && !item["freeze"]!.GetValue<bool>();
        item["unavailableProperties"] = unavailable;
        return item;
    }

    private static JsonObject LayerStateSemantics() => new()
    {
        ["current"] = "현재 작업 레이어. 켜짐/꺼짐과 별개입니다.",
        ["on"] = "LayerOn: 켜짐/꺼짐", ["freeze"] = "Freeze: 동결/해동",
        ["locked"] = "Lock: 편집 잠금 (잠금만으로 숨겨지지 않음)",
        ["modelVisible"] = "on && !freeze. 뷰포트별 동결·객체 표시/투명도·가림·화면 갱신은 별도입니다.",
        ["null"] = "조회 불가/미지원. false로 해석하지 마세요.",
    };

    private static void AddEntityDisplayState(object entity, JsonObject item)
    {
        dynamic ent = entity;
        var unavailable = new JsonArray();
        void Read(string key, Func<JsonNode?> read)
        {
            try { item[key] = read(); }
            catch { item[key] = null; unavailable.Add(key); }
        }
        Read("visible", () => JsonValue.Create((bool)ent.Visible));
        Read("color", () => JsonValue.Create((int)ent.Color));
        Read("transparency", () => JsonValue.Create((string)ent.EntityTransparency));
        item["displayPropertiesUnavailable"] = unavailable;
    }

    // Register before execution: a later COM error can leave a partially edited document.
    // Keep object identity rather than a path key (SaveAs may change the path mid-batch).
    private sealed class CadDisplayRefresh
    {
        private readonly List<object> _documents = new();

        public void Track(object doc)
        {
            if (!_documents.Any(existing => ReferenceEquals(existing, doc))) _documents.Add(doc);
        }

        public void Complete(ApplyExecution execution)
        {
            var results = new JsonArray();
            var failed = false;
            foreach (var document in _documents)
            {
                var timer = Stopwatch.StartNew();
                var result = new JsonObject { ["document"] = CadDocumentIdentity(document) };
                try
                {
                    ((dynamic)document).Regen(1); // acAllViewports, direct ActiveX; no command injection.
                    result["ok"] = true;
                }
                catch (Exception ex)
                {
                    failed = true;
                    result["ok"] = false;
                    result["error"] = ex.Message;
                    execution.Warnings.Add("[CAD_DISPLAY_REFRESH_FAILED] 도면 편집과 화면 갱신은 별개입니다. " +
                        "화면 재생성에 실패했습니다. move/scale 등을 다시 적용하지 말고 상태 조회 후 regen_document만 실행하세요.");
                }
                result["elapsedMs"] = timer.ElapsedMilliseconds;
                results.Add(result);
            }
            execution.Readback ??= new JsonObject();
            execution.Readback["displayRefresh"] = new JsonObject
            {
                ["status"] = results.Count == 0 ? "not-required" : failed ? "failed" : "completed",
                ["method"] = "ActiveX.Regen(acAllViewports)",
                ["documents"] = results,
                ["visualVerification"] = "not-performed",
            };
            // A graphics failure must not roll back or repeat already-applied geometry.
        }
    }
}
