using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

/// <summary>
/// DXF 파일 최소 파서 (CAD fallback): AutoCAD가 실행 중이지 않을 때
/// DXF 파일을 읽어 레이어/엔티티 정보를 제공한다 (읽기 전용, 쓰기 op 미지원).
/// 지원 엔티티: TEXT, MTEXT, LINE, CIRCLE, ARC, LWPOLYLINE (요약 수준).
/// </summary>
public static class DxfReader
{
    private sealed class DxfEntity
    {
        public string Type = "";
        public string Layer = "0";
        public string? Text;
        public string Handle = "";
    }

    /// <summary>group code 쌍 읽기</summary>
    private static IEnumerable<(int Code, string Value)> ReadPairs(string path)
    {
        var lines = File.ReadLines(path);
        using var e = lines.GetEnumerator();
        while (e.MoveNext())
        {
            var codeLine = e.Current.Trim();
            if (!e.MoveNext()) yield break;
            if (int.TryParse(codeLine, out var code))
                yield return (code, e.Current);
        }
    }

    public static JsonObject Analyze(string path, int maxEntities = 500)
    {
        if (!File.Exists(path))
            return Json.ErrorResult($"DXF file not found: {path}", "cad");

        var layers = new Dictionary<string, JsonObject>(StringComparer.OrdinalIgnoreCase);
        var entities = new List<DxfEntity>();
        var countsByType = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        // 상태기계: section(TABLES/ENTITIES 등) / table(LAYER 등) 을 code-0 마커로만 전이한다.
        string? section = null;
        string? table = null;
        var expectSectionName = false;
        var expectTableName = false;
        DxfEntity? current = null;
        string? curLayerName = null;
        string? curLayerColor = null;
        var curLayerOn = true;

        foreach (var (code, value) in ReadPairs(path))
        {
            var v = value.Trim();

            if (code == 0)
            {
                // 열린 LAYER 레코드 커밋 ("" = 이름 미확정 레코드는 버린다)
                if (curLayerName is not null and not "")
                {
                    if (!layers.ContainsKey(curLayerName))
                        layers[curLayerName] = new JsonObject
                        {
                            ["name"] = curLayerName,
                            ["color"] = curLayerColor,
                            ["on"] = curLayerOn,
                        };
                    curLayerName = null;
                }
                // 열린 엔티티 커밋
                if (current is not null)
                {
                    entities.Add(current);
                    current = null;
                }

                switch (v)
                {
                    case "SECTION": expectSectionName = true; break;
                    case "ENDSEC": section = null; table = null; break;
                    case "TABLE" when section == "TABLES": expectTableName = true; break;
                    case "ENDTAB": table = null; break;
                    case "LAYER" when table == "LAYER":
                        curLayerName = null; curLayerColor = null; curLayerOn = true;
                        // 다음 code 2 가 레이어 이름 — 아래 code==2 분기에서 table=="LAYER" && curLayerName==null 로 잡는다
                        curLayerName = ""; // 레코드 오픈 마커 (빈 문자열 = 이름 대기)
                        break;
                    default:
                        if (section == "ENTITIES" &&
                            v is "TEXT" or "MTEXT" or "LINE" or "CIRCLE" or "ARC" or "LWPOLYLINE" or "INSERT")
                        {
                            current = new DxfEntity { Type = v };
                            countsByType[v] = countsByType.GetValueOrDefault(v) + 1;
                        }
                        break;
                }
                continue;
            }

            if (code == 2)
            {
                if (expectSectionName) { section = v; expectSectionName = false; continue; }
                if (expectTableName) { table = v; expectTableName = false; continue; }
                if (table == "LAYER" && curLayerName == "") { curLayerName = v; continue; }
            }

            // LAYER 레코드 필드 (이름이 잡힌 뒤에만)
            if (table == "LAYER" && curLayerName is not null and not "")
            {
                if (code == 62) curLayerColor = v;
                else if (code == 70) curLayerOn = (int.TryParse(v, out var f) ? f : 0) >= 0;
                continue;
            }

            // 엔티티 필드
            if (current is not null)
            {
                if (code == 5) current.Handle = v;
                else if (code == 8) current.Layer = v;
                else if (code == 1 && current.Type is "TEXT" or "MTEXT") current.Text = v;
                else if (code == 3 && current.Type == "MTEXT") current.Text = (current.Text ?? "") + v;
            }
        }
        // 파일 끝 커밋
        if (curLayerName is not null and not "" && !layers.ContainsKey(curLayerName))
            layers[curLayerName] = new JsonObject
            {
                ["name"] = curLayerName,
                ["color"] = curLayerColor,
                ["on"] = curLayerOn,
            };
        if (current is not null) entities.Add(current);

        var layerArr = new JsonArray();
        foreach (var l in layers.Values.Take(100)) layerArr.Add(l.DeepClone());
        var entArr = new JsonArray();
        foreach (var ent in entities.Take(maxEntities))
            entArr.Add(new JsonObject
            {
                ["handle"] = ent.Handle,
                ["type"] = ent.Type,
                ["layer"] = ent.Layer,
                ["text"] = ent.Text,
            });
        var counts = new JsonObject();
        foreach (var (t, c) in countsByType) counts[t] = c;

        return new JsonObject
        {
            ["ok"] = true,
            ["app"] = "cad",
            ["mode"] = "dxf-fallback (read-only)",
            ["documentRef"] = path,
            ["summary"] = new JsonObject
            {
                ["layers"] = layerArr,
                ["entityCount"] = entities.Count,
                ["countsByType"] = counts,
                ["truncated"] = entities.Count > maxEntities,
            },
            ["entities"] = entArr,
        };
    }
}
