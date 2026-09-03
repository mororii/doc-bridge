using System.Diagnostics;
using System.Globalization;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

/// <summary>
/// AutoCAD 어댑터: 실행 중인 AutoCAD에 COM ActiveX(late binding)로 연결.
/// - 도형/수정/블록속성/배치·뷰포트는 ActiveX 객체를 직접 생성·편집한다.
/// - 고위험 op(별도 승인): 삭제, 저장/SaveAs, PDF 출력, 등록 스크립트 템플릿
/// - 문서 간 복사: AutoCAD interop의 AcadEntity[] 배열로 ActiveX CopyObjects를 우선 사용하고,
///   실패할 때만 지원 엔티티를 ModelSpace Add* 메서드로 직접 재생성한다.
/// - run_script_template은 임의 스크립트가 아니라 repo 등록 template(ops/script-templates/*.scr)만 허용
/// - fallback: AutoCAD 미실행 + args.file 지정 시 DXF 분석(읽기 전용)
/// - snapshot: 도면 파일 복사 + 레이어/텍스트 state.json
///
/// ※ 실측 기반: 이 PC AutoCAD 26.0에서 Layers.Add/LayerOn/Color/AddText/
///   HandleToObject/TextString/SaveAs 동작을 확인하고 구현했다.
/// </summary>
public sealed partial class CadAdapter : ComAdapterBase, IPreviewReuseAdapter
{
    private const int MaxContextSummaryEntities = 500;
    private const int MaxQueryEntities = 5000;
    private const int MaxDiff = 100;
    private const int MaxDrawEntities = 1000;
    private const int MaxContextLayers = 50;
    private readonly Func<object?>? _appFactory;
    private object? _attached;

    public CadAdapter(Func<object?>? appFactory = null) : base("cad", "AutoCAD.Application")
    {
        _appFactory = appFactory;
    }

    private object? AttachCad() => _attached ??= _appFactory is not null
        ? _appFactory()
        : RotHelper.GetActiveObject("AutoCAD.Application");

    /// <summary>
    /// Launch AutoCAD through COM when needed, make its main window visible, and
    /// ensure there is an active drawing that later CAD tools can edit directly.
    /// This is intentionally separate from GetActiveContext because reads must not
    /// create application or document state as a side effect.
    /// </summary>
    public JsonObject Launch(JsonObject args)
    {
        return ComInvokeWithRetry(() =>
        {
            var foreground = new ForegroundInteractionGuard(App);
            try
            {
            var warnings = new List<string>();
            var app = AttachCad();
            var launched = false;
            if (app is null)
            {
                var type = Type.GetTypeFromProgID("AutoCAD.Application");
                if (type is null)
                    return Json.ErrorResult("AutoCAD.Application COM registration was not found", App);

                app = Activator.CreateInstance(type);
                if (app is null)
                    return Json.ErrorResult("AutoCAD.Application could not be created", App);
                _attached = app;
                launched = true;
            }
            TrackCadInteraction(app, foreground, state: null);

            dynamic d = app;
            try { d.Visible = true; }
            catch (Exception ex) when (IsCallRejected(ex)) { throw; }
            catch (Exception ex) { warnings.Add($"Visible=true failed: {ex.Message}"); }

            dynamic? doc = ActiveDocWait(d, attempts: 2);
            var createdDrawing = false;
            if (doc is null)
            {
                var template = Json.GetString(args, "template") ?? "acad.dwt";
                if (template is not ("acad.dwt" or "acadiso.dwt"))
                    return Json.ErrorResult("template must be 'acad.dwt' or 'acadiso.dwt'", App);

                doc = d.Documents.Add(template);
                createdDrawing = true;
            }

            try { if (createdDrawing) doc.Activate(); }
            catch (Exception ex) when (IsCallRejected(ex)) { throw; }
            catch (Exception ex) { warnings.Add($"drawing activation failed: {ex.Message}"); }
            try { d.Visible = true; } catch { }
            try { if (launched) d.WindowState = 3; } catch { } // acMax

            string name = "";
            string fullName = "";
            try { name = (string)(doc.Name ?? ""); } catch { }
            try { fullName = (string)(doc.FullName ?? ""); } catch { }

            return new JsonObject
            {
                ["ok"] = true,
                ["app"] = App,
                ["documentRef"] = string.IsNullOrEmpty(fullName) ? $"unsaved-{name}" : fullName,
                ["summary"] = new JsonObject
                {
                    ["launched"] = launched,
                    ["visible"] = true,
                    ["createdDrawing"] = createdDrawing,
                    ["drawing"] = name,
                    ["fullName"] = fullName,
                },
                ["warnings"] = Json.ToArray(warnings),
                ["errors"] = new JsonArray(),
            };
            }
            finally { _ = foreground.Complete(); }
        }, timeoutSec: 90, maxAttempts: 30, delayMs: 1000);
    }

    private static dynamic? ActiveDoc(dynamic app)
    {
        try { return app.ActiveDocument; }
        catch (Exception ex) when (IsCallRejected(ex)) { throw; } // busy 상태는 재시도 대상 — "도면 없음"으로 둔갑시키지 않는다
        catch { return null; }
    }

    /// <summary>
    /// ActiveDocument가 부팅/로딩 레이스로 일시적으로 null을 주는 경우 대비:
    /// 짧게 폴링한다 (기본 5회×1초). 진짜로 도면이 없으면 null 반환.
    /// </summary>
    private static dynamic? ActiveDocWait(dynamic app, int attempts = 5)
    {
        for (var i = 0; i < attempts; i++)
        {
            var doc = ActiveDoc(app);
            if (doc is not null) return doc;
            Thread.Sleep(1000);
        }
        return null;
    }

    private static bool DocumentMatches(string value, string selector)
    {
        if (string.IsNullOrWhiteSpace(value) || string.IsNullOrWhiteSpace(selector)) return false;
        if (!selector.Contains('*') && !selector.Contains('?'))
            return value.Equals(selector, StringComparison.OrdinalIgnoreCase) ||
                   value.Contains(selector, StringComparison.OrdinalIgnoreCase);
        var pattern = "^" + Regex.Escape(selector).Replace("\\*", ".*").Replace("\\?", ".") + "$";
        return Regex.IsMatch(value, pattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    }

    private static dynamic? FindOpenDocument(dynamic app, string? selector)
    {
        if (string.IsNullOrWhiteSpace(selector)) return ActiveDocWait(app);
        foreach (dynamic candidate in app.Documents)
        {
            string name = "";
            string fullName = "";
            try { name = (string)(candidate.Name ?? ""); } catch { }
            try { fullName = (string)(candidate.FullName ?? ""); } catch { }
            if (DocumentMatches(name, selector) || DocumentMatches(fullName, selector)) return candidate;
        }
        return null;
    }

    private static bool TryBoundingBox(object entityObject, out double minX, out double minY, out double maxX, out double maxY)
    {
        minX = minY = maxX = maxY = 0;
        try
        {
            dynamic entity = entityObject;
            object minValue = null!;
            object maxValue = null!;
            entity.GetBoundingBox(out minValue, out maxValue);
            if (minValue is not Array min || maxValue is not Array max) return false;
            minX = Convert.ToDouble(min.GetValue(0), CultureInfo.InvariantCulture);
            minY = Convert.ToDouble(min.GetValue(1), CultureInfo.InvariantCulture);
            maxX = Convert.ToDouble(max.GetValue(0), CultureInfo.InvariantCulture);
            maxY = Convert.ToDouble(max.GetValue(1), CultureInfo.InvariantCulture);
            return true;
        }
        catch { return false; }
    }

    private static bool BoundsMatch(
        object entity, double minX, double minY, double maxX, double maxY, string mode)
    {
        if (!TryBoundingBox(entity, out var ex0, out var ey0, out var ex1, out var ey1)) return false;
        return mode.ToLowerInvariant() switch
        {
            "inside" => ex0 >= minX && ey0 >= minY && ex1 <= maxX && ey1 <= maxY,
            "intersect" => ex1 >= minX && ex0 <= maxX && ey1 >= minY && ey0 <= maxY,
            _ => (ex0 + ex1) / 2 >= minX && (ex0 + ex1) / 2 <= maxX &&
                 (ey0 + ey1) / 2 >= minY && (ey0 + ey1) / 2 <= maxY,
        };
    }

    private static JsonArray PointJson(object? value)
    {
        var result = new JsonArray();
        if (value is not Array array) return result;
        foreach (var item in array) result.Add(Convert.ToDouble(item, CultureInfo.InvariantCulture));
        return result;
    }

    private static List<object> ObjectList(object? value)
    {
        var result = new List<object>();
        if (value is Array array)
        {
            foreach (var item in array) if (item is not null) result.Add(item);
        }
        else if (value is not null) result.Add(value);
        return result;
    }

    private static JsonObject EntityJson(dynamic ent, int index, bool includeGeometry)
    {
        string type = "";
        string layer = "";
        string handle = "";
        try { type = (string)ent.EntityName; } catch { }
        try { layer = (string)ent.Layer; } catch { }
        try { handle = (string)ent.Handle; } catch { }
        var item = new JsonObject
        {
            ["index"] = index,
            ["handle"] = handle,
            ["type"] = type,
            ["layer"] = layer,
        };
        if (IsTextLike(type)) item["text"] = TextOf(ent);
        if (TryBoundingBox((object)ent, out var minX, out var minY, out var maxX, out var maxY))
            item["bounds"] = new JsonObject { ["minX"] = minX, ["minY"] = minY, ["maxX"] = maxX, ["maxY"] = maxY };
        if (!includeGeometry) return item;

        AddEntityDisplayState((object)ent, item);

        try { item["insertionPoint"] = PointJson((object?)ent.InsertionPoint); } catch { }
        try { item["position"] = PointJson((object?)ent.Position); } catch { }
        try { item["startPoint"] = PointJson((object?)ent.StartPoint); } catch { }
        try { item["endPoint"] = PointJson((object?)ent.EndPoint); } catch { }
        try { item["center"] = PointJson((object?)ent.Center); } catch { }
        try { item["coordinates"] = PointJson((object?)ent.Coordinates); } catch { }
        try { item["rotation"] = (double)ent.Rotation; } catch { }
        try { item["height"] = (double)ent.Height; } catch { }
        try { item["name"] = (string)ent.Name; } catch { }
        try { item["effectiveName"] = (string)ent.EffectiveName; } catch { }
        try { item["path"] = (string)ent.Path; } catch { }
        try { item["xScale"] = (double)ent.XScaleFactor; } catch { }
        try { item["yScale"] = (double)ent.YScaleFactor; } catch { }
        try
        {
            if ((bool)ent.HasAttributes)
            {
                var attributes = new JsonObject();
                foreach (dynamic attribute in ComObjects(ent.GetAttributes()))
                    attributes[(string)attribute.TagString] = (string)attribute.TextString;
                item["attributes"] = attributes;
            }
        }
        catch { }
        return item;
    }

    private static double[] Point(double x, double y, double z = 0) => new[] { x, y, z };

    private static double Dbl(JsonNode? node) => node switch
    {
        null => 0,
        JsonValue jv when jv.TryGetValue<double>(out var d) => d,
        JsonValue jv when jv.TryGetValue<int>(out var i) => i,
        _ => double.Parse(node.ToJsonString(), CultureInfo.InvariantCulture),
    };

    private static string TextOf(dynamic e)
    {
        try { return (string)(e.TextString ?? ""); } catch { return ""; }
    }

    private static bool IsTextLike(string entityName) =>
        entityName.Contains("Text", StringComparison.OrdinalIgnoreCase) ||
        entityName.Contains("Attribute", StringComparison.OrdinalIgnoreCase);

    // ---------- draw_entities helpers ----------

    /// <summary>[[x,y],...] + 선택적 bulge 배열로 LWPolyline 생성. 반환값은 COM RCW(object).</summary>
    private static object AddLwPolyline(dynamic doc, JsonArray points, JsonArray? bulges, bool closed)
    {
        var n = points.Count;
        if (n < 2) throw new InvalidOperationException($"lwpolyline needs >= 2 points, got {n}");
        var coords = new double[n * 2];
        for (var i = 0; i < n; i++)
        {
            if (points[i] is not JsonArray pt || pt.Count < 2)
                throw new InvalidOperationException($"point[{i}] must be [x,y]");
            coords[i * 2] = Dbl(pt[0]);
            coords[i * 2 + 1] = Dbl(pt[1]);
        }
        dynamic pl = doc.ModelSpace.AddLightWeightPolyline(coords);
        if (bulges is not null)
            for (var i = 0; i < Math.Min(bulges.Count, n); i++)
            {
                var b = Dbl(bulges[i]);
                if (b != 0) pl.SetBulge(i, b);
            }
        if (closed) pl.Closed = true;
        return (object)pl;
    }

    /// <summary>엔티티에 색 적용. {"aci":n} → ACI, {"rgb":[r,g,b]} → TrueColor(실패 시 근사 ACI 폴백)</summary>
    private static void SetEntityColor(dynamic app, dynamic ent, JsonObject? color, List<string> warnings)
    {
        if (color is null) return;
        if (color.TryGetPropertyValue("aci", out var aciNode) && aciNode is JsonValue jv && jv.TryGetValue<int>(out var aci))
        {
            ent.Color = aci;
            return;
        }
        if (Json.GetArr(color, "rgb") is { Count: >= 3 } rgb)
        {
            var r = (int)Dbl(rgb[0]);
            var g = (int)Dbl(rgb[1]);
            var b = (int)Dbl(rgb[2]);
            try
            {
                dynamic cm = TrueColorObject(app);
                cm.SetRGB(r, g, b);
                ent.TrueColor = cm;
            }
            catch (Exception ex) when (IsCallRejected(ex)) { throw; }
            catch (Exception ex)
            {
                var fallback = NearestAci(r, g, b);
                ent.Color = fallback;
                warnings.Add($"TrueColor 실패({ex.Message}) → ACI {fallback} 근사 적용");
            }
        }
    }

    /// <summary>AcCmColor ProgID 버전 접미사를 app.Version("26.0s (LMS Tech)")에서 파생, 실패 시 최근 버전들 시도</summary>
    private static dynamic TrueColorObject(dynamic app)
    {
        string version = "";
        try { version = (string)app.Version; } catch { }
        var major = "";
        foreach (var ch in version)
        {
            if (char.IsDigit(ch)) major += ch;
            else if (major.Length > 0) break;
        }
        var suffixes = new List<string>();
        if (major.Length > 0) suffixes.Add(major);
        suffixes.AddRange(new[] { "26", "25", "24", "23" });
        Exception? last = null;
        foreach (var s in suffixes.Distinct())
        {
            try { return app.GetInterfaceObject($"AutoCAD.AcCmColor.{s}"); }
            catch (Exception ex) when (IsCallRejected(ex)) { throw; }
            catch (Exception ex) { last = ex; }
        }
        throw last ?? new InvalidOperationException("AcCmColor 생성 실패");
    }

    /// <summary>TrueColor 폴백용 대표 색 근사 (태극기 팔레트 기준)</summary>
    private static int NearestAci(int r, int g, int b)
    {
        if (r > 200 && g > 200 && b > 200) return 7;  // 흰색 → 7 (배경 반전)
        if (r < 60 && g < 60 && b < 60) return 7;     // 검정 → 7
        if (r > 150 && g < 110 && b < 110) return 1;  // 빨강
        if (b > 120 && r < 110) return 5;             // 파랑
        return 7;
    }

    /// <summary>AutoCAD 명령 인수용 문화권 독립 실수 리터럴.</summary>
    private static string FNum(double v) => v.ToString("0.0#########", CultureInfo.InvariantCulture);

    private static int EntityCount(dynamic doc)
    {
        try { return (int)doc.ModelSpace.Count; } catch { return -1; }
    }

    // ---------- direct COM Taegeukgi drawing (no AutoLISP / SendCommand) ----------

    private static (double X, double Y) RotatePoint(
        (double X, double Y) point, (double X, double Y) center, double angleRad)
    {
        var dx = point.X - center.X;
        var dy = point.Y - center.Y;
        var c = Math.Cos(angleRad);
        var s = Math.Sin(angleRad);
        return (center.X + dx * c - dy * s, center.Y + dx * s + dy * c);
    }

    private static int AddSolidTriangle(
        dynamic app, dynamic doc,
        (double X, double Y) a, (double X, double Y) b, (double X, double Y) c,
        JsonObject color, List<string> warnings, string? layerName = null)
    {
        dynamic solid = doc.ModelSpace.AddSolid(
            Point(a.X, a.Y), Point(b.X, b.Y), Point(c.X, c.Y), Point(c.X, c.Y));
        if (!string.IsNullOrWhiteSpace(layerName)) solid.Layer = layerName;
        SetEntityColor(app, solid, color, warnings);
        return 1;
    }

    private static int AddSolidQuad(
        dynamic app, dynamic doc,
        (double X, double Y) a, (double X, double Y) b,
        (double X, double Y) c, (double X, double Y) d,
        JsonObject color, List<string> warnings, string? layerName = null)
    {
        return AddSolidTriangle(app, doc, a, b, c, color, warnings, layerName)
             + AddSolidTriangle(app, doc, a, c, d, color, warnings, layerName);
    }

    private static int AddSolidPolygon(
        dynamic app, dynamic doc, IReadOnlyList<(double X, double Y)> points,
        JsonObject color, List<string> warnings, string? layerName = null)
    {
        if (points.Count < 3) throw new InvalidOperationException("solid polygon needs at least three points");
        var count = 0;
        for (var i = 1; i < points.Count - 1; i++)
            count += AddSolidTriangle(app, doc, points[0], points[i], points[i + 1], color, warnings, layerName);
        return count;
    }

    private static int AddDirectPolyline(
        dynamic app, dynamic doc, IReadOnlyList<(double X, double Y)> points,
        bool closed, JsonObject color, List<string> warnings, string? layerName = null)
    {
        var arr = new JsonArray();
        foreach (var point in points)
            arr.Add(new JsonArray(point.X, point.Y));
        var polyline = AddLwPolyline(doc, arr, null, closed);
        if (!string.IsNullOrWhiteSpace(layerName)) ((dynamic)polyline).Layer = layerName;
        SetEntityColor(app, polyline, color, warnings);
        return 1;
    }

    private static int AddDirectText(
        dynamic app, dynamic doc, string text,
        double x, double y, double height, double rotation,
        JsonObject color, List<string> warnings, string? layerName = null)
    {
        dynamic entity = doc.ModelSpace.AddText(text, Point(x, y), height);
        if (!string.IsNullOrWhiteSpace(layerName)) entity.Layer = layerName;
        if (rotation != 0) entity.Rotation = rotation;
        SetEntityColor(app, entity, color, warnings);
        return 1;
    }

    private static void EnsureLayer(dynamic doc, string layerName)
    {
        try { _ = doc.Layers.Item(layerName); }
        catch { _ = doc.Layers.Add(layerName); }
    }

    private static List<string> EntityHandlesInBounds(
        dynamic doc, double minX, double minY, double maxX, double maxY)
    {
        var handles = new List<string>();
        foreach (dynamic entity in doc.ModelSpace)
        {
            try
            {
                var args = new object?[] { null, null };
                ((object)entity).GetType().InvokeMember(
                    "GetBoundingBox",
                    BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                    null, (object)entity, args, CultureInfo.InvariantCulture);
                if (args[0] is not Array min || args[1] is not Array max) continue;
                var centerX = (Convert.ToDouble(min.GetValue(0)) + Convert.ToDouble(max.GetValue(0))) / 2.0;
                var centerY = (Convert.ToDouble(min.GetValue(1)) + Convert.ToDouble(max.GetValue(1))) / 2.0;
                if (centerX >= minX && centerX <= maxX && centerY >= minY && centerY <= maxY)
                    handles.Add((string)entity.Handle);
            }
            catch { }
        }
        return handles;
    }

    private static List<string> EntityHandlesFromIndex(dynamic doc, int startIndex)
    {
        var handles = new List<string>();
        var count = (int)doc.ModelSpace.Count;
        if (startIndex < 0 || startIndex > count)
            throw new InvalidOperationException($"startIndex {startIndex} is outside ModelSpace count {count}");
        for (var i = startIndex; i < count; i++)
            handles.Add((string)doc.ModelSpace.Item(i).Handle);
        return handles;
    }

    private static (object Document, bool CloseAfter) ResolveSourceDocument(object appObject, object targetObject, JsonObject op)
    {
        dynamic app = appObject;
        dynamic targetDoc = targetObject;
        var sourceFile = Json.GetString(op, "sourceFile");
        var sourceDocument = Json.GetString(op, "sourceDocument");
        if (!string.IsNullOrWhiteSpace(sourceFile))
        {
            var fullPath = Path.GetFullPath(sourceFile);
            var open = FindOpenDocument(app, fullPath);
            if (open is not null) return ((object)open, false);
            if (!File.Exists(fullPath)) throw new FileNotFoundException("source DWG not found", fullPath);
            dynamic opened = app.Documents.Open(fullPath, true);
            return ((object)opened, true);
        }
        if (!string.IsNullOrWhiteSpace(sourceDocument))
        {
            var open = FindOpenDocument(app, sourceDocument)
                ?? throw new InvalidOperationException($"source CAD document not open: {sourceDocument}");
            return ((object)open, false);
        }
        return ((object)targetDoc, false);
    }

    private static List<object> CollectCopyObjects(dynamic sourceDoc, JsonObject op)
    {
        var result = new List<object>();
        var handles = Json.GetArr(op, "handles");
        if (handles is not null && handles.Count > 0)
        {
            foreach (var handle in handles)
                result.Add((object)sourceDoc.HandleToObject(handle!.GetValue<string>()));
            return result;
        }

        var bounds = Json.GetObj(op, "sourceBounds");
        var mode = Json.GetString(op, "selectionMode") ?? "center";
        var layerFilters = Json.GetArr(op, "layers")?.Select(n => n?.GetValue<string>() ?? "")
            .Where(s => !string.IsNullOrWhiteSpace(s)).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var typeFilters = Json.GetArr(op, "entityTypes")?.Select(n => n?.GetValue<string>() ?? "")
            .Where(s => !string.IsNullOrWhiteSpace(s)).ToHashSet(StringComparer.OrdinalIgnoreCase);
        var modelSpaceCount = Convert.ToInt32(sourceDoc.ModelSpace.Count, CultureInfo.InvariantCulture);
        for (var entityIndex = 0; entityIndex < modelSpaceCount; entityIndex++)
        {
            dynamic entity = sourceDoc.ModelSpace.Item(entityIndex);
            if (layerFilters is { Count: > 0 })
            {
                string layer = "";
                try { layer = (string)entity.Layer; } catch { }
                if (!layerFilters.Contains(layer)) continue;
            }
            if (typeFilters is { Count: > 0 })
            {
                string type = "";
                try { type = (string)entity.EntityName; } catch { }
                if (!typeFilters.Contains(type) && !typeFilters.Contains(type.Replace("AcDb", "", StringComparison.OrdinalIgnoreCase))) continue;
            }
            if (bounds is not null && !BoundsMatch((object)entity,
                Dbl(bounds["minX"]), Dbl(bounds["minY"]), Dbl(bounds["maxX"]), Dbl(bounds["maxY"]), mode)) continue;
            result.Add((object)entity);
            if (result.Count > 10000) throw new InvalidOperationException("copy selection exceeds 10,000 entities; narrow sourceBounds");
        }
        return result;
    }

    private static (double X, double Y) Origin(JsonObject? origin)
        => origin is null ? (0, 0) : (Dbl(origin["x"]), Dbl(origin["y"]));

    private static double[] ComDoubles(object? value)
    {
        if (value is not Array array) return Array.Empty<double>();
        var result = new double[array.Length];
        var i = 0;
        foreach (var item in array)
            result[i++] = Convert.ToDouble(item, CultureInfo.InvariantCulture);
        return result;
    }

    private static List<object> ComObjects(object? value)
    {
        var result = new List<object>();
        if (value is null) return result;
        if (value is Array array)
        {
            foreach (var item in array)
                if (item is not null) result.Add(item);
            return result;
        }
        result.Add(value);
        return result;
    }

    private static Array AutoCadEntityArray(dynamic document, IReadOnlyList<object> entities)
    {
        string installPath;
        try { installPath = (string)document.Application.Path; }
        catch (Exception ex)
        {
            throw new InvalidOperationException("AutoCAD install path was not available for COM entity-array marshalling", ex);
        }
        var interopPath = Path.Combine(installPath, "Autodesk.AutoCAD.Interop.Common.dll");
        if (!File.Exists(interopPath))
            throw new FileNotFoundException("AutoCAD interop assembly was not found", interopPath);
        var assembly = Assembly.LoadFrom(interopPath);
        var entityType = assembly.GetType("Autodesk.AutoCAD.Interop.Common.AcadEntity")
            ?? assembly.GetType("Autodesk.AutoCAD.Interop.Common.IAcadEntity")
            ?? throw new InvalidOperationException("AutoCAD AcadEntity interface was not found in the interop assembly");
        var result = Array.CreateInstance(entityType, entities.Count);
        for (var i = 0; i < entities.Count; i++) result.SetValue(entities[i], i);
        return result;
    }

    /// <summary>
    /// 닫힌 임시 폴리선을 경계로 비연관 솔리드 해치를 만든다. AutoCAD interop의
    /// AcadEntity[] SAFEARRAY를 사용하므로 SendCommand/AutoLISP가 전혀 필요 없다.
    /// </summary>
    private static object AddHatchDirect(
        dynamic app, dynamic doc, JsonObject loop, JsonObject? color,
        List<string> warnings, string? layerName)
    {
        var points = Json.GetArr(loop, "points")
            ?? throw new InvalidOperationException("hatch loop.points is required");
        dynamic boundary = AddLwPolyline(doc, points, Json.GetArr(loop, "bulges"), closed: true);
        dynamic? hatch = null;
        try
        {
            hatch = doc.ModelSpace.AddHatch(0, "SOLID", false, 0);
            object boundaryArray = AutoCadEntityArray(doc, new object[] { (object)boundary });
            hatch.AppendOuterLoop(boundaryArray);
            hatch.Evaluate();
            if (!string.IsNullOrWhiteSpace(layerName))
            {
                EnsureLayer(doc, layerName);
                hatch.Layer = layerName;
            }
            SetEntityColor(app, hatch, color, warnings);
            return (object)hatch;
        }
        catch
        {
            if (hatch is not null) try { hatch.Delete(); } catch { }
            throw;
        }
        finally
        {
            try { boundary.Delete(); } catch { }
        }
    }

    private static void CopyCommonEntityProperties(
        dynamic targetDoc, dynamic source, dynamic target, List<string> warnings)
    {
        try
        {
            var layer = (string)source.Layer;
            EnsureLayer(targetDoc, layer);
            target.Layer = layer;
        }
        catch (Exception ex) { warnings.Add($"copy layer property failed: {ex.Message}"); }
        try { target.Color = source.Color; } catch { }
        try { target.Linetype = source.Linetype; } catch { }
        try { target.LinetypeScale = source.LinetypeScale; } catch { }
        try { target.Lineweight = source.Lineweight; } catch { }
        try { target.Visible = source.Visible; } catch { }
    }

    /// <summary>
    /// AutoCAD의 late-bound CopyObjects는 object[]를 VT_ARRAY|VT_VARIANT로 마샬링해
    /// 일부 버전에서 "잘못된 객체 배열"을 반환한다. 문서 간 복사는 화면 명령이나
    /// AutoLISP로 우회하지 않고 ModelSpace Add* 메서드로 엔티티를 직접 재생성한다.
    /// </summary>
    private static object CloneEntityAcrossDocuments(
        dynamic targetDoc, dynamic source, List<string> warnings)
    {
        string type = "";
        try { type = (string)source.EntityName; } catch { }
        dynamic target;
        switch (type)
        {
            case "AcDbLine":
                target = targetDoc.ModelSpace.AddLine(
                    ComDoubles((object?)source.StartPoint),
                    ComDoubles((object?)source.EndPoint));
                break;

            case "AcDbText":
            {
                var insertion = ComDoubles((object?)source.InsertionPoint);
                target = targetDoc.ModelSpace.AddText(
                    (string)source.TextString, insertion, Convert.ToDouble(source.Height, CultureInfo.InvariantCulture));
                try { target.StyleName = source.StyleName; } catch { }
                try { target.Rotation = source.Rotation; } catch { }
                try { target.ObliqueAngle = source.ObliqueAngle; } catch { }
                try { target.ScaleFactor = source.ScaleFactor; } catch { }
                try { target.TextGenerationFlag = source.TextGenerationFlag; } catch { }
                try
                {
                    var alignment = Convert.ToInt32(source.Alignment, CultureInfo.InvariantCulture);
                    target.Alignment = alignment;
                    if (alignment != 0)
                        target.TextAlignmentPoint = ComDoubles((object?)source.TextAlignmentPoint);
                }
                catch { }
                break;
            }

            case "AcDbPolyline":
            {
                var coordinates = ComDoubles((object?)source.Coordinates);
                target = targetDoc.ModelSpace.AddLightWeightPolyline(coordinates);
                try { target.Closed = source.Closed; } catch { }
                try { target.ConstantWidth = source.ConstantWidth; } catch { }
                try { target.Elevation = source.Elevation; } catch { }
                var vertexCount = coordinates.Length / 2;
                for (var i = 0; i < vertexCount; i++)
                {
                    try { target.SetBulge(i, source.GetBulge(i)); } catch { }
                    try
                    {
                        object startWidth = 0.0, endWidth = 0.0;
                        source.GetWidth(i, out startWidth, out endWidth);
                        target.SetWidth(i,
                            Convert.ToDouble(startWidth, CultureInfo.InvariantCulture),
                            Convert.ToDouble(endWidth, CultureInfo.InvariantCulture));
                    }
                    catch { }
                }
                break;
            }

            case "AcDbEllipse":
                target = targetDoc.ModelSpace.AddEllipse(
                    ComDoubles((object?)source.Center),
                    ComDoubles((object?)source.MajorAxis),
                    Convert.ToDouble(source.RadiusRatio, CultureInfo.InvariantCulture));
                try { target.StartParameter = source.StartParameter; } catch { }
                try { target.EndParameter = source.EndParameter; } catch { }
                break;

            case "AcDbBlockReference":
            {
                string name;
                try { name = (string)source.Name; }
                catch { name = (string)source.EffectiveName; }
                target = targetDoc.ModelSpace.InsertBlock(
                    ComDoubles((object?)source.InsertionPoint), name,
                    Convert.ToDouble(source.XScaleFactor, CultureInfo.InvariantCulture),
                    Convert.ToDouble(source.YScaleFactor, CultureInfo.InvariantCulture),
                    Convert.ToDouble(source.ZScaleFactor, CultureInfo.InvariantCulture),
                    Convert.ToDouble(source.Rotation, CultureInfo.InvariantCulture));
                break;
            }

            case "AcDbCircle":
                target = targetDoc.ModelSpace.AddCircle(
                    ComDoubles((object?)source.Center),
                    Convert.ToDouble(source.Radius, CultureInfo.InvariantCulture));
                break;

            case "AcDbArc":
                target = targetDoc.ModelSpace.AddArc(
                    ComDoubles((object?)source.Center),
                    Convert.ToDouble(source.Radius, CultureInfo.InvariantCulture),
                    Convert.ToDouble(source.StartAngle, CultureInfo.InvariantCulture),
                    Convert.ToDouble(source.EndAngle, CultureInfo.InvariantCulture));
                break;

            case "AcDbPoint":
                target = targetDoc.ModelSpace.AddPoint(ComDoubles((object?)source.Coordinates));
                break;

            case "AcDbHatch":
            {
                Exception? nativeCopyError = null;
                try
                {
                    dynamic sourceDocument = source.Document;
                    object primaryObjects = AutoCadEntityArray(
                        sourceDocument, new object[] { (object)source });
                    object copyResult = sourceDocument.CopyObjects(
                        primaryObjects, targetDoc.ModelSpace);
                    var copiedPrimaryObjects = ComObjects(copyResult);
                    if (copiedPrimaryObjects.Count != 1)
                        throw new InvalidOperationException(
                            $"native CopyObjects returned {copiedPrimaryObjects.Count} primary objects for one AcDbHatch");
                    target = copiedPrimaryObjects[0];
                    warnings.Add("AcDbHatch was deep-cloned directly between open drawings with ActiveX CopyObjects");
                    break;
                }
                catch (Exception ex)
                {
                    nativeCopyError = ex;
                    warnings.Add($"native AcDbHatch CopyObjects fallback: {ex.Message}");
                }

                var loopCount = Convert.ToInt32(source.NumberOfLoops, CultureInfo.InvariantCulture);
                if (loopCount < 1)
                    throw new InvalidOperationException("source AcDbHatch has no boundary loops");

                // GetLoopAt returns source-database boundary objects. Recreate temporary
                // equivalents in the target database, use them to define a non-associative
                // hatch, then delete only the temporary boundaries. The real boundary
                // entities in the selection are copied separately in their original order.
                var targetLoops = new List<object[]>(loopCount);
                var temporaryBoundaries = new List<object>();
                dynamic? hatch = null;
                try
                {
                    for (var loopIndex = 0; loopIndex < loopCount; loopIndex++)
                    {
                        object? sourceLoop = null;
                        source.GetLoopAt(loopIndex, out sourceLoop);
                        var sourceBoundaries = ComObjects(sourceLoop);
                        if (sourceBoundaries.Count == 0)
                            throw new InvalidOperationException(
                                $"source AcDbHatch loop {loopIndex} is empty; native CopyObjects failed first: {nativeCopyError?.Message}");
                        var targetBoundaries = new object[sourceBoundaries.Count];
                        for (var boundaryIndex = 0; boundaryIndex < sourceBoundaries.Count; boundaryIndex++)
                        {
                            var clonedBoundary = CloneEntityAcrossDocuments(
                                targetDoc, (dynamic)sourceBoundaries[boundaryIndex], warnings);
                            temporaryBoundaries.Add(clonedBoundary);
                            targetBoundaries[boundaryIndex] = clonedBoundary;
                        }
                        targetLoops.Add(targetBoundaries);
                    }

                    var patternType = Convert.ToInt32(source.PatternType, CultureInfo.InvariantCulture);
                    var patternName = (string)source.PatternName;
                    var hatchObjectType = 0;
                    try { hatchObjectType = Convert.ToInt32(source.HatchObjectType, CultureInfo.InvariantCulture); }
                    catch { }

                    // AutoCAD requires AppendOuterLoop to be the very first operation after
                    // AddHatch. Do not move property copies above this call.
                    hatch = targetDoc.ModelSpace.AddHatch(patternType, patternName, false, hatchObjectType);
                    hatch.AppendOuterLoop(targetLoops[0]);
                    for (var loopIndex = 1; loopIndex < targetLoops.Count; loopIndex++)
                        hatch.AppendInnerLoop(targetLoops[loopIndex]);

                    try { hatch.HatchStyle = source.HatchStyle; } catch { }
                    try { hatch.PatternAngle = source.PatternAngle; } catch { }
                    try { hatch.PatternScale = source.PatternScale; } catch { }
                    try { hatch.PatternSpace = source.PatternSpace; } catch { }
                    try { hatch.PatternDouble = source.PatternDouble; } catch { }
                    try { hatch.Elevation = source.Elevation; } catch { }
                    try { hatch.Origin = ComDoubles((object?)source.Origin); } catch { }
                    try { hatch.Normal = ComDoubles((object?)source.Normal); } catch { }
                    hatch.Evaluate();
                    target = hatch;
                    warnings.Add("AcDbHatch was recreated from its source boundary loops as a non-associative target hatch");
                    break;
                }
                catch
                {
                    if (hatch is not null)
                    {
                        try { hatch.Delete(); } catch { }
                    }
                    throw;
                }
                finally
                {
                    for (var i = temporaryBoundaries.Count - 1; i >= 0; i--)
                    {
                        try { ((dynamic)temporaryBoundaries[i]).Delete(); } catch { }
                    }
                }
            }

            default:
                throw new InvalidOperationException(
                    $"direct cross-document copy does not yet support {type}; no objects were copied for this entity");
        }
        CopyCommonEntityProperties(targetDoc, source, target, warnings);
        return (object)target;
    }

    private static List<object> CopyEntitiesDirect(
        object appObject, object targetObject, JsonObject op, List<string> warnings)
    {
        dynamic app = appObject;
        dynamic targetDoc = targetObject;
        var (sourceObject, closeAfter) = ResolveSourceDocument(appObject, targetObject, op);
        dynamic sourceDoc = sourceObject;
        try
        {
            var sourceEntities = CollectCopyObjects(sourceDoc, op);
            if (sourceEntities.Count == 0) throw new InvalidOperationException("copy selection is empty");
            var sourceOrigin = Origin(Json.GetObj(op, "sourceOrigin"));
            var targetOrigin = Origin(Json.GetObj(op, "targetOrigin"));
            var scale = op["scale"] is null ? 1.0 : Dbl(op["scale"]);
            var angle = (op["rotationDeg"] is null ? 0.0 : Dbl(op["rotationDeg"])) * Math.PI / 180.0;
            if (scale <= 0) throw new InvalidOperationException("copy scale must be positive");

            string sourceName = "", targetName = "", sourceFullName = "", targetFullName = "";
            try { sourceName = (string)(sourceDoc.Name ?? ""); } catch { }
            try { targetName = (string)(targetDoc.Name ?? ""); } catch { }
            try { sourceFullName = (string)(sourceDoc.FullName ?? ""); } catch { }
            try { targetFullName = (string)(targetDoc.FullName ?? ""); } catch { }
            var sameDocument = (!string.IsNullOrWhiteSpace(sourceFullName) &&
                                sourceFullName.Equals(targetFullName, StringComparison.OrdinalIgnoreCase)) ||
                               sourceName.Equals(targetName, StringComparison.OrdinalIgnoreCase);
            List<object> copied;
            if (sameDocument)
            {
                copied = new List<object>(sourceEntities.Count);
                foreach (dynamic sourceEntity in sourceEntities)
                    copied.Add((object)sourceEntity.Copy());
            }
            else
            {
                var nativeBefore = EntityCount(targetDoc);
                try
                {
                    object primaryObjects = AutoCadEntityArray(sourceDoc, sourceEntities);
                    object copyResult = sourceDoc.CopyObjects(primaryObjects, targetDoc.ModelSpace);
                    copied = ComObjects(copyResult);
                    if (copied.Count != sourceEntities.Count)
                        throw new InvalidOperationException(
                            $"ActiveX CopyObjects returned {copied.Count} primary objects for {sourceEntities.Count} source entities");
                    warnings.Add(
                        $"{copied.Count} cross-document entities were deep-cloned directly with ActiveX CopyObjects; AutoLISP/WBLOCK was not used");
                }
                catch (Exception nativeCopyError)
                {
                    if (nativeBefore >= 0)
                    {
                        foreach (var handle in EntityHandlesFromIndex(targetDoc, nativeBefore))
                        {
                            try { targetDoc.HandleToObject(handle).Delete(); } catch { }
                        }
                    }
                    warnings.Add($"batch ActiveX CopyObjects fallback: {nativeCopyError.Message}");
                    copied = new List<object>(sourceEntities.Count);
                    try
                    {
                        foreach (dynamic sourceEntity in sourceEntities)
                            copied.Add(CloneEntityAcrossDocuments(targetDoc, sourceEntity, warnings));
                    }
                    catch
                    {
                        // Cross-document recreation is a batch operation. If any entity is
                        // unsupported or invalid, remove every entity already recreated so a
                        // retry cannot leave duplicates or untranslated source-coordinate debris.
                        for (var i = copied.Count - 1; i >= 0; i--)
                        {
                            try { ((dynamic)copied[i]).Delete(); } catch { }
                        }
                        try { targetDoc.Regen(1); } catch { }
                        throw;
                    }
                    warnings.Add(
                        "cross-document entities were recreated directly with ActiveX ModelSpace Add* methods; AutoLISP/WBLOCK was not used");
                }
            }
            try
            {
                foreach (dynamic entity in copied)
                {
                    if (Math.Abs(scale - 1.0) > 1e-12)
                        entity.ScaleEntity(Point(sourceOrigin.X, sourceOrigin.Y), scale);
                    if (Math.Abs(angle) > 1e-12)
                        entity.Rotate(Point(sourceOrigin.X, sourceOrigin.Y), angle);
                    if (Math.Abs(targetOrigin.X - sourceOrigin.X) > 1e-12 || Math.Abs(targetOrigin.Y - sourceOrigin.Y) > 1e-12)
                        entity.Move(Point(sourceOrigin.X, sourceOrigin.Y), Point(targetOrigin.X, targetOrigin.Y));
                }
            }
            catch
            {
                for (var i = copied.Count - 1; i >= 0; i--)
                {
                    try { ((dynamic)copied[i]).Delete(); } catch { }
                }
                try { targetDoc.Regen(1); } catch { }
                throw;
            }
            try { targetDoc.Regen(1); } catch (Exception ex) { warnings.Add($"regen failed after copy: {ex.Message}"); }
            return copied;
        }
        finally
        {
            if (closeAfter)
            {
                try { sourceDoc.Close(false); } catch { }
            }
        }
    }

    private static object InsertXrefDirect(dynamic doc, JsonObject op, List<string> warnings)
    {
        var sourceFile = Path.GetFullPath(Json.GetString(op, "sourceFile")
            ?? throw new InvalidOperationException("insert_xref requires sourceFile"));
        if (!File.Exists(sourceFile)) throw new FileNotFoundException("xref source DWG not found", sourceFile);
        var point = Json.GetObj(op, "insertionPoint") ?? new JsonObject();
        var scale = op["scale"] is null ? 1.0 : Dbl(op["scale"]);
        var rotation = op["rotationDeg"] is null ? 0.0 : Dbl(op["rotationDeg"]) * Math.PI / 180.0;
        if (scale <= 0) throw new InvalidOperationException("xref scale must be positive");
        var requestedName = Json.GetString(op, "name") ?? Path.GetFileNameWithoutExtension(sourceFile);
        var reuseExistingDefinition = Json.GetBool(op, "reuseExistingDefinition");
        dynamic xref;
        if (reuseExistingDefinition)
        {
            var existingDefinition = Json.GetString(op, "existingDefinition") ?? requestedName;
            dynamic blockDefinition;
            try { blockDefinition = doc.Blocks.Item(existingDefinition); }
            catch { throw new InvalidOperationException($"existing XREF definition '{existingDefinition}' was not found"); }
            var isXref = false;
            try { isXref = (bool)blockDefinition.IsXRef; } catch { }
            if (!isXref)
                throw new InvalidOperationException($"existing block definition '{existingDefinition}' is not an XREF");

            xref = doc.ModelSpace.InsertBlock(
                Point(Dbl(point["x"]), Dbl(point["y"])), existingDefinition,
                scale, scale, scale, rotation);
            warnings.Add(
                $"reused existing XREF definition '{existingDefinition}' so its dependent-layer visibility matches completed frames");
        }
        else
        {
        var safeName = Regex.Replace(requestedName, "[^0-9A-Za-z가-힣_-]", "_");
        if (string.IsNullOrWhiteSpace(safeName)) safeName = "DocBridgeXref";
        var uniqueName = safeName;
        var suffix = 1;
        while (true)
        {
            try { _ = doc.Blocks.Item(uniqueName); uniqueName = $"{safeName}_{suffix++}"; }
            catch { break; }
        }

        xref = doc.ModelSpace.AttachExternalReference(
            sourceFile, uniqueName, Point(Dbl(point["x"]), Dbl(point["y"])), scale, scale, scale, rotation, false);
        }
        var layerName = Json.GetString(op, "layer");
        if (!string.IsNullOrWhiteSpace(layerName))
        {
            EnsureLayer(doc, layerName);
            xref.Layer = layerName;
        }

        var clip = Json.GetObj(op, "clipBounds");
        if (clip is not null)
        {
            if (!WaitCmdIdle(doc, 15)) throw new InvalidOperationException("AutoCAD command line was not idle before XCLIP");
            var command = string.Join("\n", new[]
            {
                "_.UCS", "_W",
                "_.XCLIP", "_L", "", "_N", "_R",
                $"{FNum(Dbl(clip["minX"]))},{FNum(Dbl(clip["minY"]))}",
                $"{FNum(Dbl(clip["maxX"]))},{FNum(Dbl(clip["maxY"]))}",
                "_.UCS", "_P", "",
            });
            doc.SendCommand(command + "\n");
            if (!WaitCmdIdle(doc, 60)) warnings.Add("native XCLIP command did not return to idle state within 60 seconds");
            warnings.Add("clip boundary applied with AutoCAD's native XCLIP command; no AutoLISP was used");
        }
        return (object)xref;
    }

    private static int AddFan(
        dynamic app, dynamic doc,
        (double X, double Y) center, double radius,
        double startAngle, double endAngle, int segments, double rotation,
        JsonObject color, List<string> warnings)
    {
        var count = 0;
        for (var i = 0; i < segments; i++)
        {
            var a0 = startAngle + (endAngle - startAngle) * i / segments;
            var a1 = startAngle + (endAngle - startAngle) * (i + 1) / segments;
            var p0 = (center.X + radius * Math.Cos(a0), center.Y + radius * Math.Sin(a0));
            var p1 = (center.X + radius * Math.Cos(a1), center.Y + radius * Math.Sin(a1));
            if (rotation != 0)
            {
                p0 = RotatePoint(p0, center, rotation);
                p1 = RotatePoint(p1, center, rotation);
            }
            count += AddSolidTriangle(app, doc, center, p0, p1, color, warnings);
        }
        return count;
    }

    private static int DrawTaegeukgiDirect(dynamic app, dynamic doc, List<string> warnings)
    {
        var white = new JsonObject { ["rgb"] = new JsonArray(255, 255, 255) };
        var black = new JsonObject { ["rgb"] = new JsonArray(0, 0, 0) };
        var red = new JsonObject { ["rgb"] = new JsonArray(205, 46, 58) };
        var blue = new JsonObject { ["rgb"] = new JsonArray(0, 71, 160) };
        var count = 0;

        // 3:2 white field, created first so every later symbol remains visible.
        count += AddSolidQuad(app, doc, (0.0, 0.0), (90.0, 0.0), (90.0, 60.0), (0.0, 60.0), white, warnings);

        var center = (X: 45.0, Y: 30.0);
        const double radius = 15.0;
        var rotation = -33.6900675259798 * Math.PI / 180.0;

        // Outer red/blue halves, then two half-radius discs form the S curve.
        count += AddFan(app, doc, center, radius, 0, Math.PI, 24, rotation, red, warnings);
        count += AddFan(app, doc, center, radius, Math.PI, 2 * Math.PI, 24, rotation, blue, warnings);
        var redCenter = RotatePoint((center.X - radius / 2, center.Y), center, rotation);
        var blueCenter = RotatePoint((center.X + radius / 2, center.Y), center, rotation);
        count += AddFan(app, doc, redCenter, radius / 2, 0, 2 * Math.PI, 24, 0, red, warnings);
        count += AddFan(app, doc, blueCenter, radius / 2, 0, 2 * Math.PI, 24, 0, blue, warnings);

        dynamic taegeukOutline = doc.ModelSpace.AddCircle(Point(center.X, center.Y), radius);
        SetEntityColor(app, taegeukOutline, black, warnings);
        count++;

        // Official-style rotated trigram bars. Each quad becomes two direct AcDbSolid objects.
        var bars = new (double X, double Y)[][]
        {
            new[] { (30.439120,48.721132),(22.118617,36.240377),(20.038491,37.627128),(28.358994,50.107882) },
            new[] { (27.318931,50.801257),(18.998428,38.320503),(16.918303,39.707253),(25.238806,52.188008) },
            new[] { (24.198743,52.881383),(15.878240,40.400629),(13.798114,41.787379),(22.118617,54.268134) },

            new[] { (67.881383,36.240377),(64.067819,41.960723),(66.147945,43.347473),(69.961509,37.627128) },
            new[] { (63.374444,43.000786),(59.560880,48.721132),(61.641006,50.107882),(65.454570,44.387536) },
            new[] { (71.001572,38.320503),(62.681069,50.801257),(64.761194,52.188008),(73.081697,39.707253) },
            new[] { (74.121760,40.400629),(70.308196,46.120974),(72.388322,47.507725),(76.201886,41.787379) },
            new[] { (69.614821,47.161037),(65.801257,52.881383),(67.881383,54.268134),(71.694947,48.547788) },

            new[] { (22.118617,23.759623),(30.439120,11.278868),(28.358994,9.892118),(20.038491,22.372872) },
            new[] { (18.998428,21.679497),(22.811992,15.959151),(20.731866,14.572401),(16.918303,20.292747) },
            new[] { (23.505367,14.919088),(27.318931,9.198743),(25.238806,7.811992),(21.425242,13.532338) },
            new[] { (15.878240,19.599371),(24.198743,7.118617),(22.118617,5.731866),(13.798114,18.212621) },

            new[] { (59.560880,11.278868),(63.374444,16.999214),(65.454570,15.612464),(61.641006,9.892118) },
            new[] { (64.067819,18.039277),(67.881383,23.759623),(69.961509,22.372872),(66.147945,16.652527) },
            new[] { (62.681069,9.198743),(66.494633,14.919088),(68.574758,13.532338),(64.761194,7.811992) },
            new[] { (67.188008,15.959151),(71.001572,21.679497),(73.081697,20.292747),(69.268134,14.572401) },
            new[] { (65.801257,7.118617),(69.614821,12.838963),(71.694947,11.452212),(67.881383,5.731866) },
            new[] { (70.308196,13.879026),(74.121760,19.599371),(76.201886,18.212621),(72.388322,12.492275) },
        };
        foreach (var bar in bars)
            count += AddSolidQuad(app, doc, bar[0], bar[1], bar[2], bar[3], black, warnings);

        var border = AddLwPolyline(doc,
            new JsonArray(new JsonArray(0, 0), new JsonArray(90, 0), new JsonArray(90, 60), new JsonArray(0, 60)),
            null, true);
        SetEntityColor(app, border, black, warnings);
        count++;

        try { doc.Regen(1); } catch { }
        try { app.ZoomExtents(); } catch { }
        return count;
    }

    private static int DrawUnionJackDirect(
        dynamic app, dynamic doc, List<string> warnings,
        double originX, double originY, double width, double height)
    {
        if (width <= 0 || height <= 0) throw new InvalidOperationException("Union Jack width and height must be positive");

        var blue = new JsonObject { ["rgb"] = new JsonArray(1, 33, 105) };
        var white = new JsonObject { ["rgb"] = new JsonArray(255, 255, 255) };
        var red = new JsonObject { ["rgb"] = new JsonArray(200, 16, 46) };
        var black = new JsonObject { ["rgb"] = new JsonArray(0, 0, 0) };
        var count = 0;

        (double X, double Y) P(double nx, double ny) =>
            (originX + width * nx / 120.0, originY + height * ny / 60.0);

        // Blue field.
        count += AddSolidQuad(app, doc, P(0, 0), P(120, 0), P(120, 60), P(0, 60), blue, warnings);

        // Broad white diagonals (St Andrew's saltire), clipped to the flag rectangle.
        count += AddSolidPolygon(app, doc,
            new[] { P(0, 0), P(12, 0), P(120, 54), P(120, 60), P(108, 60), P(0, 6) }, white, warnings);
        count += AddSolidPolygon(app, doc,
            new[] { P(0, 54), P(0, 60), P(12, 60), P(120, 6), P(120, 0), P(108, 0) }, white, warnings);

        // Narrow red diagonals (St Patrick's saltire).
        count += AddSolidPolygon(app, doc,
            new[] { P(0, 0), P(4, 0), P(120, 58), P(120, 60), P(116, 60), P(0, 2) }, red, warnings);
        count += AddSolidPolygon(app, doc,
            new[] { P(0, 58), P(0, 60), P(4, 60), P(120, 2), P(120, 0), P(116, 0) }, red, warnings);

        // White-edged central cross, then the narrower red St George's cross.
        count += AddSolidQuad(app, doc, P(0, 24), P(120, 24), P(120, 36), P(0, 36), white, warnings);
        count += AddSolidQuad(app, doc, P(54, 0), P(66, 0), P(66, 60), P(54, 60), white, warnings);
        count += AddSolidQuad(app, doc, P(0, 27), P(120, 27), P(120, 33), P(0, 33), red, warnings);
        count += AddSolidQuad(app, doc, P(57, 0), P(63, 0), P(63, 60), P(57, 60), red, warnings);

        var border = AddLwPolyline(doc,
            new JsonArray(
                new JsonArray(originX, originY),
                new JsonArray(originX + width, originY),
                new JsonArray(originX + width, originY + height),
                new JsonArray(originX, originY + height)),
            null, true);
        SetEntityColor(app, border, black, warnings);
        count++;

        try { doc.Regen(1); } catch { }
        try { app.ZoomExtents(); } catch { }
        return count;
    }

    private static int DrawBlockWallSchematicDirect(
        dynamic app, dynamic doc, List<string> warnings,
        double originX, double originY, double scale)
    {
        if (scale <= 0) throw new InvalidOperationException("block wall schematic scale must be positive");

        var white = new JsonObject { ["rgb"] = new JsonArray(250, 250, 250) };
        var black = new JsonObject { ["rgb"] = new JsonArray(35, 35, 35) };
        var header = new JsonObject { ["rgb"] = new JsonArray(58, 58, 58) };
        var grid = new JsonObject { ["rgb"] = new JsonArray(210, 214, 218) };
        var gray = new JsonObject { ["rgb"] = new JsonArray(165, 169, 173) };
        var darkBlue = new JsonObject { ["rgb"] = new JsonArray(39, 48, 92) };
        var brightBlue = new JsonObject { ["rgb"] = new JsonArray(53, 91, 170) };
        var yellow = new JsonObject { ["rgb"] = new JsonArray(243, 196, 48) };
        var red = new JsonObject { ["rgb"] = new JsonArray(197, 62, 70) };
        var count = 0;

        (double X, double Y) P(double x, double y) => (originX + x * scale, originY + y * scale);

        int FilledRect(double x1, double y1, double x2, double y2, JsonObject color) =>
            AddSolidQuad(app, doc, P(x1, y1), P(x2, y1), P(x2, y2), P(x1, y2), color, warnings);

        int Line(IReadOnlyList<(double X, double Y)> points, JsonObject color, bool closed = false)
        {
            var scaled = points.Select(p => P(p.X, p.Y)).ToArray();
            return AddDirectPolyline(app, doc, scaled, closed, color, warnings);
        }

        int Text(string value, double x, double y, double h, JsonObject color, double rotation = 0) =>
            AddDirectText(app, doc, value, P(x, y).X, P(x, y).Y, h * scale, rotation, color, warnings);

        int Block(double x, double y, double size, JsonObject fill, int subdivisions)
        {
            var c = FilledRect(x, y, x + size, y + size, fill);
            c += Line(new[] { (x, y), (x + size, y), (x + size, y + size), (x, y + size) }, black, true);
            for (var i = 1; i < subdivisions; i++)
            {
                var d = size * i / subdivisions;
                c += Line(new[] { (x + d, y), (x + d, y + size) }, grid);
                c += Line(new[] { (x, y + d), (x + size, y + d) }, grid);
            }
            return c;
        }

        // Dark title band and title text.
        count += FilledRect(-5, 140, 230, 155, header);
        count += Text("싸인블록과 짙은 색상의 일반블록을 활용한 징검다리 표현", 5, 145, 5.2, white);

        // Main 22.5 m x 5.4 m wall face and the angled approach wing.
        count += AddSolidPolygon(app, doc,
            new[] { P(0, 74), P(225, 74), P(219, 26), P(214, 20), P(15, 20), P(8, 27) }, white, warnings);
        count += AddSolidPolygon(app, doc,
            new[] { P(-15, 136), P(0, 74), P(80, 74) }, white, warnings);

        // Light construction grid on the wall face.
        for (var x = 6.0; x <= 219.0; x += 3.0)
            count += Line(new[] { (x, 20.0), (x, 74.0) }, grid);
        for (var y = 20.0; y <= 74.0; y += 3.0)
            count += Line(new[] { (6.0, y), (219.0, y) }, grid);

        // Yellow safety/boundary blocks.
        count += AddSolidQuad(app, doc, P(0, 72), P(5, 72), P(13, 27), P(8, 27), yellow, warnings);
        count += FilledRect(9, 24, 58, 27, yellow);
        count += FilledRect(8, 36, 219, 39, yellow);
        count += FilledRect(31, 24, 34, 39, yellow);
        count += AddSolidQuad(app, doc, P(219, 26), P(224, 28), P(229, 72), P(224, 72), yellow, warnings);

        // Repeating ordinary and sign blocks on the main face.
        foreach (var (x, y) in new (double X, double Y)[]
        {
            (28,62),(36,55),(44,48),(52,60),(60,52),(68,45),(76,61),(84,52),(92,45),
            (100,60),(108,50),(116,43),(124,59),(132,51),(140,45),(148,62),(156,53),
            (164,45),(172,60),(180,51),(188,44),(196,61),(204,52),(212,45),
        }) count += Block(x, y, 5, gray, 1);

        foreach (var (x, y) in new (double X, double Y)[]
        {
            (42,59),(56,59),(70,47),(96,48),(111,55),(126,49),(151,43),(166,55),(181,46),(199,48),
        }) count += Block(x, y, 6, darkBlue, 2);

        foreach (var (x, y) in new (double X, double Y)[]
        {
            (51,46),(63,63),(82,48),(118,45),(143,59),(160,48),(187,58),
        }) count += Block(x, y, 4, brightBlue, 2);

        // A simplified diagonal approach pattern.
        foreach (var (x, y, color) in new (double X, double Y, JsonObject Color)[]
        {
            (-3,105,gray),(5,100,gray),(13,95,darkBlue),(21,90,gray),(29,85,gray),(37,80,brightBlue),(45,76,gray),
        }) count += Block(x, y, 5, color, color == darkBlue || color == brightBlue ? 2 : 1);
        count += Line(new[] { (-15.0,136.0),(0.0,74.0),(80.0,74.0) }, black, true);

        // Detail callouts: 600, 400, and 200 mm blocks.
        count += Text("싸인블록", 69, 126, 3.5, black);
        count += Text("일반 블록", 99, 126, 3.5, black);
        count += Block(66, 102, 24, darkBlue, 2);
        count += Block(94, 102, 24, gray, 2);

        count += Text("싸인블록", 145, 126, 3.3, black);
        count += Text("일반 블록", 169, 126, 3.3, black);
        count += Block(145, 108, 18, brightBlue, 2);
        count += Block(167, 108, 18, gray, 2);

        count += Text("일반 블록", 204, 126, 3.3, black);
        count += Block(207, 114, 10, gray, 1);

        // Red leader lines from details to representative blocks.
        count += Line(new[] { (66.0,102.0),(104.0,56.0) }, red);
        count += Line(new[] { (90.0,102.0),(116.0,60.0) }, red);
        count += Line(new[] { (118.0,102.0),(132.0,62.0) }, red);
        count += Line(new[] { (145.0,108.0),(151.0,62.0) }, red);
        count += Line(new[] { (163.0,108.0),(177.0,60.0) }, red);
        count += Line(new[] { (185.0,108.0),(195.0,58.0) }, red);
        count += Line(new[] { (207.0,114.0),(212.0,68.0) }, red);

        // Detail dimensions and labels.
        count += Text("300", 59, 110, 3.0, black, Math.PI / 2);
        count += Text("300", 59, 102, 3.0, black, Math.PI / 2);
        count += Text("200", 138, 114, 3.0, black, Math.PI / 2);
        count += Text("200", 138, 108, 3.0, black, Math.PI / 2);
        count += Text("200", 200, 114, 3.0, black, Math.PI / 2);

        // Main dimensions: 22,500 mm x 5,400 mm and representative spacings.
        count += Line(new[] { (0.0,5.0),(225.0,5.0) }, black);
        count += Line(new[] { (0.0,2.0),(0.0,8.0) }, black);
        count += Line(new[] { (225.0,2.0),(225.0,8.0) }, black);
        count += Text("22,500mm", 98, 0, 4.5, black);

        count += Line(new[] { (-18.0,20.0),(-18.0,74.0) }, black);
        count += Line(new[] { (-21.0,20.0),(-15.0,20.0) }, black);
        count += Line(new[] { (-21.0,74.0),(-15.0,74.0) }, black);
        count += Text("5,400mm", -24, 40, 4.5, black, Math.PI / 2);

        count += Line(new[] { (0.0,12.0),(225.0,12.0) }, red);
        foreach (var x in new[] { 0.0, 132.0, 176.0, 198.0, 225.0 })
            count += Line(new[] { (x,10.0),(x,14.0) }, red);
        count += Text("200", 128, 14, 3.5, black);
        count += Text("400", 173, 14, 3.5, black);
        count += Text("600", 195, 14, 3.5, black);

        // Main exterior outline last so it remains crisp above the grid and fills.
        count += Line(new[] { (0.0,74.0),(225.0,74.0),(219.0,26.0),(214.0,20.0),(15.0,20.0),(8.0,27.0) }, black, true);

        try { doc.Regen(1); } catch { }
        try { app.ZoomExtents(); } catch { }
        return count;
    }

    private static int DrawBlockWallConstructionDirect(
        dynamic app, dynamic doc, List<string> warnings,
        double originX, double originY, double scale)
    {
        if (scale <= 0) throw new InvalidOperationException("construction drawing scale must be positive");

        const string outlineLayer = "DB-WALL-OUTLINE";
        const string gridLayer = "DB-WALL-GRID";
        const string boundaryLayer = "DB-BOUNDARY-BLOCK";
        const string ordinaryLayer = "DB-ORDINARY-BLOCK";
        const string signDarkLayer = "DB-SIGN-BLOCK-600";
        const string signBlueLayer = "DB-SIGN-BLOCK-400";
        const string dimLayer = "DB-DIMENSION";
        const string textLayer = "DB-TEXT";
        foreach (var layer in new[]
        {
            outlineLayer, gridLayer, boundaryLayer, ordinaryLayer,
            signDarkLayer, signBlueLayer, dimLayer, textLayer,
        }) EnsureLayer(doc, layer);

        var white = new JsonObject { ["rgb"] = new JsonArray(250, 250, 250) };
        var black = new JsonObject { ["rgb"] = new JsonArray(30, 30, 30) };
        var header = new JsonObject { ["rgb"] = new JsonArray(58, 58, 58) };
        var grid = new JsonObject { ["rgb"] = new JsonArray(210, 214, 218) };
        var gray = new JsonObject { ["rgb"] = new JsonArray(165, 169, 173) };
        var darkBlue = new JsonObject { ["rgb"] = new JsonArray(39, 48, 92) };
        var brightBlue = new JsonObject { ["rgb"] = new JsonArray(53, 91, 170) };
        var yellow = new JsonObject { ["rgb"] = new JsonArray(243, 196, 48) };
        var red = new JsonObject { ["rgb"] = new JsonArray(197, 62, 70) };
        var count = 0;

        (double X, double Y) P(double x, double y) => (originX + x * scale, originY + y * scale);

        int FilledRect(double x1, double y1, double x2, double y2, JsonObject color, string layer) =>
            AddSolidQuad(app, doc, P(x1, y1), P(x2, y1), P(x2, y2), P(x1, y2), color, warnings, layer);

        int Line(IReadOnlyList<(double X, double Y)> points, JsonObject color, string layer, bool closed = false)
        {
            var scaled = points.Select(p => P(p.X, p.Y)).ToArray();
            return AddDirectPolyline(app, doc, scaled, closed, color, warnings, layer);
        }

        int Text(string value, double x, double y, double h, JsonObject color, string layer, double rotation = 0)
        {
            var at = P(x, y);
            return AddDirectText(app, doc, value, at.X, at.Y, h * scale, rotation, color, warnings, layer);
        }

        int Module(double x, double y, double size, JsonObject fill, string layer, int subdivisions)
        {
            var c = FilledRect(x, y, x + size, y + size, fill, layer);
            c += Line(new[] { (x, y), (x + size, y), (x + size, y + size), (x, y + size) }, black, outlineLayer, true);
            for (var i = 1; i < subdivisions; i++)
            {
                var offset = size * i / subdivisions;
                c += Line(new[] { (x + offset, y), (x + offset, y + size) }, grid, gridLayer);
                c += Line(new[] { (x, y + offset), (x + size, y + offset) }, grid, gridLayer);
            }
            return c;
        }

        // Title and drawing scale note.
        count += FilledRect(-500, 7600, 23000, 8500, header, textLayer);
        count += Text("싸인블록과 일반블록을 활용한 징검다리 시공상세도", 500, 7920, 380, white, textLayer);
        count += Text("MODEL 1:1 / 단위:mm", 18500, 7720, 220, white, textLayer);

        // Wall face: true 22,500 x 5,400 mm model-space dimensions.
        count += AddSolidPolygon(app, doc,
            new[] { P(0, 5400), P(22500, 5400), P(22000, 800), P(21200, 0), P(1500, 0), P(700, 700) },
            white, warnings, ordinaryLayer);
        count += AddSolidPolygon(app, doc,
            new[] { P(-1200, 8500), P(0, 5400), P(7200, 5400) },
            white, warnings, ordinaryLayer);

        // 200 mm construction grid, matching the smallest ordinary block module.
        for (var x = 600.0; x <= 21900.0; x += 200.0)
            count += Line(new[] { (x, 400.0), (x, 5400.0) }, grid, gridLayer);
        for (var y = 400.0; y <= 5400.0; y += 200.0)
            count += Line(new[] { (600.0, y), (21900.0, y) }, grid, gridLayer);

        // 300 mm yellow boundary blocks and turning bands.
        count += AddSolidQuad(app, doc, P(0, 5200), P(300, 5200), P(1300, 700), P(700, 700), yellow, warnings, boundaryLayer);
        count += FilledRect(800, 300, 6200, 600, yellow, boundaryLayer);
        count += FilledRect(900, 1800, 21800, 2100, yellow, boundaryLayer);
        count += FilledRect(3500, 300, 3800, 2100, yellow, boundaryLayer);
        count += AddSolidQuad(app, doc, P(21700, 800), P(22000, 900), P(22500, 5100), P(22200, 5100), yellow, warnings, boundaryLayer);

        // 600 x 600 sign blocks (four 300 x 300 tiles).
        foreach (var (x, y) in new (double X, double Y)[]
        {
            (4300,4300),(6900,4300),(10100,3300),(12800,4200),(15700,3300),(18500,4200),
        }) count += Module(x, y, 600, darkBlue, signDarkLayer, 2);

        // 400 x 400 sign blocks (four 200 x 200 tiles).
        foreach (var (x, y) in new (double X, double Y)[]
        {
            (5600,2900),(8400,3100),(11400,4300),(14500,2800),(17600,4400),(20100,3100),
        }) count += Module(x, y, 400, brightBlue, signBlueLayer, 2);

        // 600 x 600 ordinary blocks (four 300 x 300 tiles).
        foreach (var (x, y) in new (double X, double Y)[]
        {
            (2500,4300),(5300,3500),(7900,4300),(9300,2500),(11200,2600),(11900,4300),
            (13700,3400),(14900,4300),(16800,2500),(19400,4300),
        }) count += Module(x, y, 600, gray, ordinaryLayer, 2);

        // 400 x 400 ordinary blocks (four 200 x 200 tiles).
        foreach (var (x, y) in new (double X, double Y)[]
        {
            (3300,3400),(6200,2600),(7400,3500),(8700,2600),(10600,4300),(12600,2500),
            (14200,2500),(15300,3500),(18100,3300),(20700,4300),
        }) count += Module(x, y, 400, gray, ordinaryLayer, 2);

        // Individual 200 x 200 ordinary blocks on the same 200 mm grid.
        foreach (var (x, y) in new (double X, double Y)[]
        {
            (3600,4600),(3900,3600),(5100,4700),(6500,3900),(7500,2700),(9000,4700),
            (9800,3900),(11000,3500),(12300,4700),(13300,2900),(15100,3900),(16200,4700),
            (17400,3700),(18800,3000),(19900,3900),(21400,4700),
        }) count += Module(x, y, 200, gray, ordinaryLayer, 1);

        // Modular pattern on the angled approach wing.
        foreach (var (x, y, size, fill, layer, sub) in new (double X, double Y, double Size, JsonObject Fill, string Layer, int Sub)[]
        {
            (-300,6900,600,darkBlue,signDarkLayer,2),(500,6600,400,gray,ordinaryLayer,2),
            (1300,6300,600,gray,ordinaryLayer,2),(2300,6000,400,brightBlue,signBlueLayer,2),
            (3200,5700,600,gray,ordinaryLayer,2),(4300,5500,400,gray,ordinaryLayer,2),
        }) count += Module(x, y, size, fill, layer, sub);

        // Enlarged details above the wall, using the exact module sizes.
        count += Text("싸인블록 600×600", 5900, 7160, 220, black, textLayer);
        count += Module(6000, 6200, 600, darkBlue, signDarkLayer, 2);
        count += Text("일반블록 600×600", 7300, 7160, 220, black, textLayer);
        count += Module(7400, 6200, 600, gray, ordinaryLayer, 2);

        count += Text("싸인블록 400×400", 11600, 7160, 220, black, textLayer);
        count += Module(11800, 6400, 400, brightBlue, signBlueLayer, 2);
        count += Text("일반블록 400×400", 13000, 7160, 220, black, textLayer);
        count += Module(13200, 6400, 400, gray, ordinaryLayer, 2);

        count += Text("일반블록 200×200", 17900, 7160, 220, black, textLayer);
        count += Module(18400, 6500, 200, gray, ordinaryLayer, 1);

        // Exact detail dimensions: 300+300, 200+200, and 200 mm.
        count += Text("300", 5750, 6250, 170, black, dimLayer, Math.PI / 2);
        count += Text("300", 5750, 6550, 170, black, dimLayer, Math.PI / 2);
        count += Text("300", 7250, 6250, 170, black, dimLayer, Math.PI / 2);
        count += Text("300", 7250, 6550, 170, black, dimLayer, Math.PI / 2);
        count += Text("200", 11620, 6410, 150, black, dimLayer, Math.PI / 2);
        count += Text("200", 11620, 6610, 150, black, dimLayer, Math.PI / 2);
        count += Text("200", 13020, 6410, 150, black, dimLayer, Math.PI / 2);
        count += Text("200", 13020, 6610, 150, black, dimLayer, Math.PI / 2);
        count += Text("200", 18150, 6510, 150, black, dimLayer, Math.PI / 2);

        // Leaders terminate at modules of the matching size and color.
        count += Line(new[] { (6000.0,6200.0),(4600.0,4900.0) }, red, dimLayer);
        count += Line(new[] { (6600.0,6200.0),(10400.0,3900.0) }, red, dimLayer);
        count += Line(new[] { (8000.0,6200.0),(12200.0,4900.0) }, red, dimLayer);
        count += Line(new[] { (11800.0,6400.0),(11600.0,4700.0) }, red, dimLayer);
        count += Line(new[] { (12200.0,6400.0),(14700.0,3200.0) }, red, dimLayer);
        count += Line(new[] { (13600.0,6400.0),(18300.0,3500.0) }, red, dimLayer);
        count += Line(new[] { (18600.0,6500.0),(21500.0,4800.0) }, red, dimLayer);

        // Overall 22,500 mm length and 5,400 mm height dimensions.
        count += Line(new[] { (0.0,-900.0),(22500.0,-900.0) }, black, dimLayer);
        count += Line(new[] { (0.0,-1100.0),(0.0,-700.0) }, black, dimLayer);
        count += Line(new[] { (22500.0,-1100.0),(22500.0,-700.0) }, black, dimLayer);
        count += Text("22,500mm", 9800, -1500, 320, black, dimLayer);

        count += Line(new[] { (-1400.0,0.0),(-1400.0,5400.0) }, black, dimLayer);
        count += Line(new[] { (-1600.0,0.0),(-1200.0,0.0) }, black, dimLayer);
        count += Line(new[] { (-1600.0,5400.0),(-1200.0,5400.0) }, black, dimLayer);
        count += Text("5,400mm", -2050, 2050, 320, black, dimLayer, Math.PI / 2);

        // Representative horizontal spacing dimensions aligned to the actual grid.
        count += Line(new[] { (0.0,-300.0),(22500.0,-300.0) }, red, dimLayer);
        foreach (var x in new[] { 0.0, 11800.0, 13200.0, 18400.0, 22500.0 })
            count += Line(new[] { (x,-450.0),(x,-150.0) }, red, dimLayer);
        count += Text("200", 11650, -100, 180, black, dimLayer);
        count += Text("400", 13000, -100, 180, black, dimLayer);
        count += Text("600", 18150, -100, 180, black, dimLayer);

        // Crisp exterior outlines last.
        count += Line(new[] { (0.0,5400.0),(22500.0,5400.0),(22000.0,800.0),(21200.0,0.0),(1500.0,0.0),(700.0,700.0) }, black, outlineLayer, true);
        count += Line(new[] { (-1200.0,8500.0),(0.0,5400.0),(7200.0,5400.0) }, black, outlineLayer, true);

        try { doc.Regen(1); } catch { }
        try { app.ZoomExtents(); } catch { }
        return count;
    }

    /// <summary>콘솔이 명령 대기 상태(CMDACTIVE=0)가 될 때까지 대기 — 유령 입력 큐 방지</summary>
    private static int DrawBlockWallInstallationDirect(
        dynamic app, dynamic doc, List<string> warnings,
        double originX, double originY, double scale)
    {
        if (scale <= 0) throw new InvalidOperationException("installation drawing scale must be positive");

        const string outlineLayer = "DB-WALL-OUTLINE";
        const string gridLayer = "DB-WALL-GRID";
        const string guideLayer = "DB-GUIDE-BLOCK";
        const string ordinaryLayer = "DB-ORDINARY-BLOCK";
        const string signDarkLayer = "DB-SIGN-BLOCK-600";
        const string signBlueLayer = "DB-SIGN-BLOCK-400";
        const string dimLayer = "DB-DIMENSION";
        const string textLayer = "DB-TEXT";
        foreach (var layer in new[]
        {
            outlineLayer, gridLayer, guideLayer, ordinaryLayer, signDarkLayer,
            signBlueLayer, dimLayer, textLayer,
        }) EnsureLayer(doc, layer);

        var black = new JsonObject { ["rgb"] = new JsonArray(30, 30, 30) };
        var grid = new JsonObject { ["rgb"] = new JsonArray(176, 181, 186) };
        var baseGray = new JsonObject { ["rgb"] = new JsonArray(222, 225, 228) };
        var darkGray = new JsonObject { ["rgb"] = new JsonArray(142, 146, 150) };
        var darkBlue = new JsonObject { ["rgb"] = new JsonArray(39, 48, 92) };
        var brightBlue = new JsonObject { ["rgb"] = new JsonArray(53, 91, 170) };
        var yellow = new JsonObject { ["rgb"] = new JsonArray(242, 197, 48) };
        var red = new JsonObject { ["rgb"] = new JsonArray(197, 62, 70) };
        var count = 0;

        (double X, double Y) P(double x, double y) =>
            (originX + x * scale, originY + y * scale);

        int Line(IReadOnlyList<(double X, double Y)> points, JsonObject color, string layer, bool closed = false)
        {
            var scaled = points.Select(p => P(p.X, p.Y)).ToArray();
            return AddDirectPolyline(app, doc, scaled, closed, color, warnings, layer);
        }

        int Text(string value, double x, double y, double height, JsonObject color, string layer, double rotation = 0)
        {
            var at = P(x, y);
            return AddDirectText(app, doc, value, at.X, at.Y, height * scale, rotation, color, warnings, layer);
        }

        IReadOnlyList<(double X, double Y)> AxisRect(double x, double y, double size) =>
            new[] { (x, y), (x + size, y), (x + size, y + size), (x, y + size) };

        int Module(IReadOnlyList<(double X, double Y)> corners, JsonObject fill, string layer, int subdivisions)
        {
            if (corners.Count != 4) throw new InvalidOperationException("module requires four corners");
            var localCount = AddSolidQuad(app, doc,
                P(corners[0].X, corners[0].Y), P(corners[1].X, corners[1].Y),
                P(corners[2].X, corners[2].Y), P(corners[3].X, corners[3].Y),
                fill, warnings, layer);
            localCount += Line(corners, black, outlineLayer, true);
            for (var i = 1; i < subdivisions; i++)
            {
                var t = (double)i / subdivisions;
                var a = (X: corners[0].X + (corners[1].X - corners[0].X) * t,
                         Y: corners[0].Y + (corners[1].Y - corners[0].Y) * t);
                var b = (X: corners[3].X + (corners[2].X - corners[3].X) * t,
                         Y: corners[3].Y + (corners[2].Y - corners[3].Y) * t);
                var c = (X: corners[0].X + (corners[3].X - corners[0].X) * t,
                         Y: corners[0].Y + (corners[3].Y - corners[0].Y) * t);
                var d = (X: corners[1].X + (corners[2].X - corners[1].X) * t,
                         Y: corners[1].Y + (corners[2].Y - corners[1].Y) * t);
                localCount += Line(new[] { a, b }, grid, gridLayer);
                localCount += Line(new[] { c, d }, grid, gridLayer);
            }
            return localCount;
        }

        int ClippedGrid(IReadOnlyList<(double X, double Y)> polygon, double spacing, double angle)
        {
            var dx = Math.Cos(angle);
            var dy = Math.Sin(angle);
            var nx = -dy;
            var ny = dx;

            int Family(double normalX, double normalY, double directionX, double directionY)
            {
                var familyCount = 0;
                var min = polygon.Min(p => p.X * normalX + p.Y * normalY);
                var max = polygon.Max(p => p.X * normalX + p.Y * normalY);
                var start = Math.Ceiling((min + 0.01) / spacing) * spacing;
                for (var offset = start; offset < max - 0.01; offset += spacing)
                {
                    var hits = new List<(double X, double Y)>();
                    for (var i = 0; i < polygon.Count; i++)
                    {
                        var a = polygon[i];
                        var b = polygon[(i + 1) % polygon.Count];
                        var fa = a.X * normalX + a.Y * normalY - offset;
                        var fb = b.X * normalX + b.Y * normalY - offset;
                        if ((fa < 0 && fb < 0) || (fa > 0 && fb > 0) || Math.Abs(fa - fb) < 1e-9) continue;
                        var t = fa / (fa - fb);
                        if (t < -1e-9 || t > 1 + 1e-9) continue;
                        var hit = (X: a.X + (b.X - a.X) * t, Y: a.Y + (b.Y - a.Y) * t);
                        if (!hits.Any(p => Math.Abs(p.X - hit.X) < 0.01 && Math.Abs(p.Y - hit.Y) < 0.01)) hits.Add(hit);
                    }
                    hits.Sort((a, b) => (a.X * directionX + a.Y * directionY)
                        .CompareTo(b.X * directionX + b.Y * directionY));
                    for (var i = 0; i + 1 < hits.Count; i += 2)
                        familyCount += Line(new[] { hits[i], hits[i + 1] }, grid, gridLayer);
                }
                return familyCount;
            }

            return Family(nx, ny, dx, dy) + Family(dx, dy, -nx, -ny);
        }

        // Only these exterior landmarks come from the raster reference. They are normalized
        // to the stated 22,500 mm length and 5,400 mm face height before any blocks are laid out.
        var mainFace = new (double X, double Y)[]
        {
            (0,5400), (22500,5400), (21775,1050), (21250,450), (20175,0),
            (2100,0), (1225,225), (700,1050),
        };
        var returnWing = new (double X, double Y)[]
        {
            (-1600,11400), (6900,5400), (0,5400),
        };

        count += AddSolidPolygon(app, doc, mainFace.Select(p => P(p.X, p.Y)).ToArray(),
            baseGray, warnings, ordinaryLayer);
        count += AddSolidPolygon(app, doc, returnWing.Select(p => P(p.X, p.Y)).ToArray(),
            baseGray, warnings, ordinaryLayer);

        // Exact 200 x 200 mm ordinary-paver grids, clipped at the scaled exterior cut line.
        count += ClippedGrid(mainFace, 200, 0);
        var wingAngle = Math.Atan2(-6000.0, 8500.0);
        count += ClippedGrid(returnWing, 200, wingAngle);

        // Fixed main-face schedule measured one module at a time from the reference raster.
        // No repeating placement algorithm is used: 20 sign modules + 68 dark-gray modules.
        foreach (var module in new (double X, double Y, double Size, bool Sign)[]
        {
            // Sign modules.
            (16800,4400,600,true),(6500,4600,400,true),(4300,4400,600,true),(5700,4400,600,true),
            (14900,4000,400,true),(16300,4000,400,true),(19000,4000,400,true),(9300,4000,400,true),
            (15500,3600,600,true),(19600,3600,600,true),(7100,3600,600,true),(9900,3600,600,true),
            (12700,3600,600,true),(11300,3600,600,true),(17700,3400,400,true),(7900,3400,400,true),
            (12100,3400,400,true),(5100,3400,400,true),(18300,2800,600,true),(14100,2800,600,true),

            // Dark-gray modules.
            (15400,4400,600,false),(16200,4600,400,false),(18200,4400,600,false),(19000,4600,400,false),
            (19600,4400,600,false),(20400,4600,400,false),(21000,4400,600,false),(5100,4600,400,false),
            (7100,4400,600,false),(7900,4600,400,false),(8500,4400,600,false),(9300,4600,400,false),
            (9800,4400,600,false),(10700,4600,400,false),(11200,4400,600,false),(12100,4600,400,false),
            (13400,4600,400,false),(14000,4400,600,false),(14900,4600,400,false),(2900,4400,600,false),
            (3700,4600,400,false),(12700,4400,600,false),(17700,4700,400,false),
            (17600,4000,400,false),(20400,4000,400,false),(10700,4000,400,false),(12100,4000,400,false),
            (13400,4000,400,false),(3700,4000,400,false),(5100,4000,400,false),(6500,4000,400,false),
            (7900,4000,400,false),(14000,3600,600,false),(16800,3600,600,false),(18200,3600,600,false),
            (21000,3600,600,false),(4300,3600,600,false),(5700,3600,600,false),(8500,3600,600,false),
            (19000,3400,400,false),(20400,3400,400,false),(10700,3400,400,false),(13400,3400,400,false),
            (14800,3400,400,false),(16200,3400,400,false),(6500,3400,400,false),(9300,3400,400,false),
            (15400,2800,600,false),(16800,2800,600,false),(19600,2800,600,false),(21000,2800,600,false),
            (5700,2800,600,false),(7100,2800,600,false),(8500,2800,600,false),(9800,2800,600,false),
            (11200,2800,600,false),(12700,2800,600,false),(17600,2800,400,false),(19000,2800,400,false),
            (20400,2800,400,false),(7900,2800,400,false),(9300,2800,400,false),(10700,2800,400,false),
            (12100,2800,400,false),(13400,2800,400,false),(14800,2800,400,false),(16200,2800,400,false),
            (6500,2800,400,false),
        })
        {
            var is600 = Math.Abs(module.Size - 600) < 0.01;
            count += Module(AxisRect(module.X, module.Y, module.Size),
                module.Sign ? (is600 ? darkBlue : brightBlue) : darkGray,
                module.Sign ? (is600 ? signDarkLayer : signBlueLayer) : ordinaryLayer, 2);
        }

        // Rotated modules on the return wing follow its own 200 mm setting-out axes.
        var wingDx = Math.Cos(wingAngle);
        var wingDy = Math.Sin(wingAngle);
        var wingNx = wingDy;
        var wingNy = -wingDx;
        (double X, double Y) WingPoint(double u, double v) =>
            (-1600 + u * wingDx + v * wingNx, 11400 + u * wingDy + v * wingNy);
        IReadOnlyList<(double X, double Y)> WingRect(double u, double v, double size) =>
            new[] { WingPoint(u, v), WingPoint(u + size, v), WingPoint(u + size, v + size), WingPoint(u, v + size) };

        // Fixed return-wing schedule measured from the rotated reference region:
        // 5 sign modules + 18 dark-gray modules.
        foreach (var module in new (double U, double V, double Size, bool Sign)[]
        {
            (4300,400,600,true),(3700,1500,400,true),(5100,1600,400,true),(5700,1900,600,true),
            (7100,1200,600,true),
            (2900,400,600,false),(3700,400,400,false),(3700,1000,400,false),(5000,400,400,false),
            (4300,1200,600,false),(5100,1000,400,false),(5700,500,600,false),(6400,500,400,false),
            (4300,1900,600,false),(5700,1200,600,false),(6500,1000,400,false),(7100,500,600,false),
            (5200,2100,400,false),(7900,500,400,false),(6500,1600,400,false),(8400,500,600,false),
            (7900,1100,400,false),(6500,2100,400,false),
        })
        {
            var corners = WingRect(module.U, module.V, module.Size);
            var is600 = Math.Abs(module.Size - 600) < 0.01;
            count += Module(corners,
                module.Sign ? (is600 ? darkBlue : brightBlue) : darkGray,
                module.Sign ? (is600 ? signDarkLayer : signBlueLayer) : ordinaryLayer, 2);
        }

        // 300 x 300 tactile guide blocks are continuous modules, never a single solid band.
        int GuideStrip((double X, double Y) start, (double X, double Y) end, bool interiorOnLeft, int blockCount)
        {
            const double guideSize = 300;
            var vx = end.X - start.X;
            var vy = end.Y - start.Y;
            var length = Math.Sqrt(vx * vx + vy * vy);
            if (blockCount <= 0 || blockCount * guideSize > length + 0.01)
                throw new InvalidOperationException("guide-block count does not fit the measured segment");
            var ux = vx / length;
            var uy = vy / length;
            var nx = interiorOnLeft ? -uy : uy;
            var ny = interiorOnLeft ? ux : -ux;
            var stripCount = 0;
            for (var block = 0; block < blockCount; block++)
            {
                var distance = block * guideSize;
                var a = (X: start.X + ux * distance, Y: start.Y + uy * distance);
                var b = (X: start.X + ux * (distance + guideSize), Y: start.Y + uy * (distance + guideSize));
                var corners = new[]
                {
                    a, b,
                    (X: b.X + nx * guideSize, Y: b.Y + ny * guideSize),
                    (X: a.X + nx * guideSize, Y: a.Y + ny * guideSize),
                };
                stripCount += Module(corners, yellow, guideLayer, 1);
            }
            return stripCount;
        }

        // Left boundary path continues from the return wing to the lower-left turn.
        count += GuideStrip((-1600,11400), (0,5400), true, 20);
        count += GuideStrip((0,5400), (700,1050), true, 14);
        count += GuideStrip((700,1050), (1225,225), true, 3);
        count += GuideStrip((1225,225), (2100,0), true, 3);
        // Right boundary guide follows the long sloping cut line.
        count += GuideStrip((22500,5400), (21775,1050), false, 14);

        // Counted schedules: horizontal 70, vertical branch 6, lower-left continuation 14.
        for (var block = 0; block < 70; block++)
            count += Module(AxisRect(900 + block * 300, 2100, 300), yellow, guideLayer, 1);
        for (var block = 0; block < 6; block++)
            count += Module(AxisRect(3300, 300 + block * 300, 300), yellow, guideLayer, 1);
        for (var block = 0; block < 14; block++)
            count += Module(AxisRect(2100 + block * 300, 0, 300), yellow, guideLayer, 1);

        // Compact construction legend; no decorative layout is copied from the reference image.
        count += Text("PAVER INSTALLATION PLAN", 9000, 11200, 420, black, textLayer);
        count += Text("MODEL 1:1  /  ALL DIMENSIONS IN mm", 9000, 10650, 230, black, textLayer);
        count += Module(AxisRect(9000, 9200, 600), darkBlue, signDarkLayer, 2);
        count += Text("SIGN 600x600 = 4 x 300x300", 9800, 9400, 220, black, textLayer);
        count += Module(AxisRect(15000, 9300, 400), brightBlue, signBlueLayer, 2);
        count += Text("SIGN 400x400 = 4 x 200x200", 15600, 9400, 220, black, textLayer);
        count += Module(AxisRect(9000, 8200, 600), darkGray, ordinaryLayer, 2);
        count += Text("DARK 600x600 = 4 x 300x300", 9800, 8400, 220, black, textLayer);
        count += Module(AxisRect(15000, 8300, 400), darkGray, ordinaryLayer, 2);
        count += Text("DARK 400x400 = 4 x 200x200", 15600, 8400, 220, black, textLayer);
        count += Module(AxisRect(9000, 7600, 200), baseGray, ordinaryLayer, 1);
        count += Text("BASE ORDINARY 200x200", 9400, 7600, 220, black, textLayer);
        count += Module(AxisRect(15000, 7500, 300), yellow, guideLayer, 1);
        count += Text("TACTILE GUIDE 300x300", 15500, 7600, 220, black, textLayer);

        // Overall dimensions and datum extension lines.
        count += Line(new[] { (0.0,-800.0),(22500.0,-800.0) }, black, dimLayer);
        count += Line(new[] { (0.0,-1000.0),(0.0,-600.0) }, black, dimLayer);
        count += Line(new[] { (22500.0,-1000.0),(22500.0,-600.0) }, black, dimLayer);
        count += Text("22,500", 10300, -1350, 320, black, dimLayer);
        count += Line(new[] { (0.0,-600.0),(0.0,0.0) }, red, dimLayer);
        count += Line(new[] { (22500.0,-600.0),(22500.0,5400.0) }, red, dimLayer);

        count += Line(new[] { (-1300.0,0.0),(-1300.0,5400.0) }, black, dimLayer);
        count += Line(new[] { (-1500.0,0.0),(-1100.0,0.0) }, black, dimLayer);
        count += Line(new[] { (-1500.0,5400.0),(-1100.0,5400.0) }, black, dimLayer);
        count += Text("5,400", -1850, 2050, 320, black, dimLayer, Math.PI / 2);

        // Exterior silhouette is the controlling cut line and is drawn last.
        count += Line(mainFace, black, outlineLayer, true);
        count += Line(returnWing, black, outlineLayer, true);

        try { doc.Regen(1); } catch { }
        try { app.ZoomExtents(); } catch { }
        return count;
    }

    private static bool WaitCmdIdle(dynamic doc, int timeoutSec)
    {
        var deadline = DateTime.UtcNow.AddSeconds(timeoutSec);
        while (DateTime.UtcNow < deadline)
        {
            int active;
            try { active = (int)doc.GetVariable("CMDACTIVE"); }
            catch { return false; }
            if (active == 0) return true;
            Thread.Sleep(300);
        }
        return false;
    }

    // ---------- IAppAdapter ----------

    public override JsonObject GetCapabilities() => new()
    {
        ["app"] = App,
        ["automation"] = "autocad-activex-com",
        ["directAppControl"] = true,
        ["connectsToExistingWindow"] = true,
        ["usesUiAutomation"] = false,
        ["usesAutoLisp"] = false,
        ["interactionPolicy"] = new JsonObject
        {
            ["mode"] = "preserve-foreground",
            ["backgroundInactiveWindow"] = true,
            ["restoresOriginalDocument"] = true,
            ["restoresLayoutAndSpace"] = true,
            ["restoresViewCenterAndSize"] = true,
            ["concurrentTargetInput"] = "stop-after-current-operation",
            ["sameDocumentConcurrentEditing"] = false,
        },
        ["readOps"] = new JsonArray("context", "entities", "layouts", "layers", "xrefs", "window", "regions", "dxf-fallback"),
        ["writeOps"] = new JsonArray(
            "activate_document", "regen_document", "set_layer_visibility", "set_layer_color", "move_entities",
            "rotate_entities", "set_text_value", "copy_entities_between_documents", "insert_xref",
            "zoom_window", "draw_entities", "copy_entities", "scale_entities", "mirror_entities",
            "offset_entities", "set_entity_properties", "set_block_attributes",
            "configure_layout", "create_viewport", "save_document", "plot_pdf",
            "delete_entities", "delete_entities_in_bounds", "delete_entities_from_index", "run_script_template"),
        ["drawEntityTypes"] = new JsonArray(
            "lwpolyline", "circle", "block", "text", "hatch", "line", "arc", "ellipse",
            "point", "mtext", "dim_aligned", "dim_rotated"),
        ["limits"] = new JsonObject
        {
            ["contextEntityScanDefault"] = 0,
            ["contextEntityScanSummary"] = MaxContextSummaryEntities,
            ["contextLayerPreview"] = MaxContextLayers,
            ["maxQueryResults"] = MaxQueryEntities,
            ["maxDrawEntitiesPerOp"] = MaxDrawEntities,
            ["queryContinuation"] = true,
        },
        ["safety"] = new JsonArray("dry-run", "snapshot", "confirm-token", "readback", "automatic-rollback"),
    };

    public override AdapterStatus GetStatus()
    {
        try
        {
            return ComInvokeWithRetry(() =>
            {
                var app = AttachCad();
                if (app is null)
                    return new AdapterStatus(false, false, "cad", null, null,
                        "AutoCAD 실행 인스턴스를 찾지 못했고 새 인스턴스 생성도 실패했습니다 (DXF fallback은 file 인자로 분석 가능)");
                dynamic d = app;
                string? version = null; string? doc = null;
                try { version = (string)d.Version; } catch { }
                try { doc = (string?)ActiveDocWait(d)?.FullName; } catch { }
                return new AdapterStatus(true, true, "cad", version, doc, null);
            });
        }
        catch (Exception ex) { return new AdapterStatus(false, false, "cad", null, null, ex.Message); }
    }

    public override ContextResult GetActiveContext() => GetActiveContext(new JsonObject());

    /// <summary>
    /// 대형 도면에서 상태 확인만으로 수백 개의 COM 엔티티를 읽지 않도록 기본값은 basic이다.
    /// summary는 기존 호환용으로 레이어 미리보기와 최대 500개 엔티티 유형 표본을 반환한다.
    /// 전체 레이어/영역/엔티티는 cad_query_entities의 명시적 scope로 조회한다.
    /// </summary>
    public ContextResult GetActiveContext(JsonObject? args)
    {
        var detailLevel = (Json.GetString(args, "detailLevel") ?? "basic").Trim().ToLowerInvariant();
        if (detailLevel is not ("basic" or "summary"))
        {
            var invalid = new ContextResult { App = App };
            invalid.Errors.Add("detailLevel은 'basic' 또는 'summary'여야 합니다.");
            return invalid;
        }

        return ComInvokeWithRetry(() =>
        {
            var r = new ContextResult { App = App };
            var foreground = new ForegroundInteractionGuard(App);
            try
            {
                var app = AttachCad();
                if (app is not null) TrackCadInteraction(app, foreground, state: null);
                if (app is null) { r.Errors.Add("AutoCAD가 실행 중이지 않습니다. DXF 분석은 cad_query_entities의 file 인자를 사용하세요."); return r; }
                dynamic d = app;
                var doc = ActiveDocWait(d);
                if (doc is null) { r.Errors.Add("열린 도면이 없습니다."); return r; }

                r.Ok = true;
                string fullName = "";
                try { fullName = (string)(doc.FullName ?? ""); } catch { }
                r.DocumentRef = string.IsNullOrEmpty(fullName) ? $"unsaved-{(string)doc.Name}" : fullName;
                r.Summary["drawing"] = (string)doc.Name;
                r.Summary["fullName"] = fullName;
                r.Summary["detailLevel"] = detailLevel;

                var openDocuments = new JsonArray();
                foreach (dynamic openDoc in d.Documents)
                {
                    string openName = "";
                    string openFullName = "";
                    var openCount = -1;
                    try { openName = (string)(openDoc.Name ?? ""); } catch { }
                    try { openFullName = (string)(openDoc.FullName ?? ""); } catch { }
                    try { openCount = (int)openDoc.ModelSpace.Count; } catch { }
                    openDocuments.Add(new JsonObject
                    {
                        ["name"] = openName,
                        ["fullName"] = openFullName,
                        ["entityCount"] = openCount,
                        ["active"] = openName.Equals((string)doc.Name, StringComparison.OrdinalIgnoreCase),
                    });
                }
                r.Summary["openDocuments"] = openDocuments;

                var layerCount = 0;
                try { layerCount = Convert.ToInt32(doc.Layers.Count, CultureInfo.InvariantCulture); } catch { }
                var layers = new JsonArray();
                string? currentLayer = CurrentLayerName(doc);
                r.Summary["currentLayer"] = currentLayer;
                r.Summary["layerStateSemantics"] = LayerStateSemantics();
                if (detailLevel == "summary")
                {
                    foreach (dynamic layer in doc.Layers)
                    {
                        var item = LayerState((object)layer, currentLayer);
                        layers.Add(item);
                        if (layers.Count >= MaxContextLayers) break;
                    }
                }
                r.Summary["layers"] = layers;
                r.Summary["layerCount"] = layerCount;
                r.Summary["layersTruncated"] = layerCount > layers.Count;
                r.Summary["layerPreviewIncluded"] = detailLevel == "summary";
                r.Summary["layerSummaryStatus"] = detailLevel == "basic" ? "omitted" :
                    layerCount > layers.Count ? "sampled" : "complete";
                if (detailLevel == "summary" && layerCount > layers.Count)
                    r.Warnings.Add($"활성 컨텍스트는 레이어 {layers.Count}/{layerCount}개만 반환했습니다. 전체 목록은 cad_query_entities(scope=layers)를 사용하세요.");

                var counts = new JsonObject();
                var countMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                var total = 0;
                var modelSpaceCount = (int)doc.ModelSpace.Count;
                if (detailLevel == "summary")
                {
                    foreach (dynamic ent in doc.ModelSpace)
                    {
                        total++;
                        string type;
                        try { type = (string)ent.EntityName; }
                        catch
                        {
                            if (total >= MaxContextSummaryEntities) break;
                            continue;
                        }
                        countMap[type] = countMap.GetValueOrDefault(type) + 1;
                        if (total >= MaxContextSummaryEntities) break;
                    }
                }
                foreach (var (t, c) in countMap) counts[t] = c;
                r.Summary["entityCount"] = modelSpaceCount;
                r.Summary["entitySummaryScanned"] = total;
                r.Summary["entitySummaryTruncated"] = total < modelSpaceCount;
                r.Summary["entitySummaryStatus"] = detailLevel == "basic"
                    ? "omitted"
                    : total < modelSpaceCount ? "sampled" : "complete";
                r.Summary["countsByType"] = counts;
                r.Summary["coverage"] = new JsonObject
                {
                    ["layers"] = new JsonObject
                    {
                        ["total"] = layerCount,
                        ["returned"] = layers.Count,
                        ["truncated"] = layerCount > layers.Count,
                        ["complete"] = layerCount == layers.Count,
                    },
                    ["entityTypeSummary"] = new JsonObject
                    {
                        ["totalEntities"] = modelSpaceCount,
                        ["scanned"] = total,
                        ["truncated"] = total < modelSpaceCount,
                        ["complete"] = total == modelSpaceCount,
                    },
                };
                r.Summary["nextActions"] = BuildContextNextActions(
                    string.IsNullOrWhiteSpace(fullName) ? (string)doc.Name : fullName,
                    includeLayerQuery: layerCount > layers.Count,
                    includeEntityQuery: total < modelSpaceCount);
                try { r.Summary["insunits"] = (int)doc.GetVariable("INSUNITS"); } catch { }
            }
            catch (Exception ex) when (IsCallRejected(ex)) { throw; }
            catch (Exception ex) { r.Errors.Add($"cad context failed: {ex.Message}"); }
            finally { r.Interaction = foreground.Complete(); }
            return r;
        });
    }

    private static JsonArray BuildContextNextActions(
        string documentRef, bool includeLayerQuery, bool includeEntityQuery)
    {
        var actions = new JsonArray();
        if (includeLayerQuery)
        {
            actions.Add(new JsonObject
            {
                ["tool"] = "cad_query_entities",
                ["reason"] = "전체 레이어 목록을 페이지 단위로 조회",
                ["arguments"] = new JsonObject
                {
                    ["scope"] = "layers",
                    ["document"] = documentRef,
                    ["startIndex"] = 0,
                    ["limit"] = 500,
                },
            });
        }
        if (includeEntityQuery)
        {
            actions.Add(new JsonObject
            {
                ["tool"] = "cad_query_entities",
                ["reason"] = "대형 도면은 전체 순차 표본 대신 도곽/작업 영역별로 조회",
                ["arguments"] = new JsonObject
                {
                    ["scope"] = "regions",
                    ["document"] = documentRef,
                },
                ["requiredArguments"] = new JsonArray("regions[].name", "regions[].bounds"),
            });
            actions.Add(new JsonObject
            {
                ["tool"] = "cad_query_entities",
                ["reason"] = "한 영역의 실제 엔티티가 필요하면 AutoCAD 네이티브 공간 선택 사용",
                ["arguments"] = new JsonObject
                {
                    ["scope"] = "window",
                    ["document"] = documentRef,
                    ["boundsMode"] = "intersect",
                    ["limit"] = 500,
                },
                ["requiredArguments"] = new JsonArray("bounds.minX", "bounds.minY", "bounds.maxX", "bounds.maxY"),
            });
        }
        return actions;
    }

    private static JsonObject CadQueryAction(
        string reason, JsonObject arguments, params string[] requiredArguments)
    {
        var action = new JsonObject
        {
            ["tool"] = "cad_query_entities",
            ["reason"] = reason,
            ["arguments"] = arguments,
        };
        if (requiredArguments.Length > 0)
            action["requiredArguments"] = new JsonArray(requiredArguments.Select(v => JsonValue.Create(v)).ToArray());
        return action;
    }

    private static JsonObject QueryArguments(JsonObject source, string scope)
    {
        var result = new JsonObject { ["scope"] = scope };
        foreach (var key in new[]
        {
            "document", "file", "layer", "entityType", "textContains", "blockName",
            "includeGeometry", "countOnly", "boundsMode", "bounds", "contains", "startsWith",
        })
        {
            if (source.TryGetPropertyValue(key, out var node) && node is not null)
                result[key] = node.DeepClone();
        }
        return result;
    }

    public override JsonObject Read(JsonObject args)
    {
        var file = Json.GetString(args, "file");
        if (!string.IsNullOrEmpty(file) && Path.GetExtension(file).Equals(".dxf", StringComparison.OrdinalIgnoreCase))
        {
            var st = GetStatus();
            if (!st.Available) return DxfReader.Analyze(file);
        }

        return ComInvokeWithRetry(() =>
        {
            dynamic? queryDoc = null;
            var closeAfterRead = false;
            var foreground = new ForegroundInteractionGuard(App);
            var documentState = new CadInteractionState();
            try
            {
                var app = AttachCad();
                if (app is null) return Json.ErrorResult("AutoCAD not running. use 'file' arg for DXF fallback analysis.", App);
                TrackCadInteraction(app, foreground, documentState);
                dynamic d = app;
                var documentSelector = Json.GetString(args, "document");
                if (!string.IsNullOrWhiteSpace(file) && Path.GetExtension(file).Equals(".dwg", StringComparison.OrdinalIgnoreCase))
                {
                    var fullPath = Path.GetFullPath(file);
                    queryDoc = FindOpenDocument(d, fullPath);
                    if (queryDoc is null)
                    {
                        if (!File.Exists(fullPath)) return Json.ErrorResult($"DWG file not found: {fullPath}", App);
                        queryDoc = d.Documents.Open(fullPath, true);
                        closeAfterRead = true;
                    }
                }
                else queryDoc = FindOpenDocument(d, documentSelector);
                if (queryDoc is null) return Json.ErrorResult($"CAD document not found: {documentSelector ?? file ?? "active"}", App);
                dynamic doc = queryDoc;

                if (string.Equals(Json.GetString(args, "scope"), "layouts", StringComparison.OrdinalIgnoreCase))
                    return InspectLayouts(doc);
                if (string.Equals(Json.GetString(args, "scope"), "layers", StringComparison.OrdinalIgnoreCase))
                    return InspectLayers(doc, args);
                if (string.Equals(Json.GetString(args, "scope"), "xrefs", StringComparison.OrdinalIgnoreCase))
                    return InspectXrefs(doc, args);
                if (string.Equals(Json.GetString(args, "scope"), "window", StringComparison.OrdinalIgnoreCase))
                    return InspectWindowSelection(doc, args);
                if (string.Equals(Json.GetString(args, "scope"), "regions", StringComparison.OrdinalIgnoreCase))
                    return InspectRegions(doc, Json.GetArr(args, "regions")
                        ?? throw new ArgumentException("scope=regions requires regions array"));

                var layerFilter = Json.GetString(args, "layer");
                var typeFilter = Json.GetString(args, "entityType");
                var textContains = Json.GetString(args, "textContains");
                var blockName = Json.GetString(args, "blockName");
                var includeGeometry = Json.GetBool(args, "includeGeometry");
                var countOnly = Json.GetBool(args, "countOnly");
                var limit = Math.Min(Json.GetInt(args, "limit") ?? 100, MaxQueryEntities);
                var bounds = Json.GetObj(args, "bounds");
                var boundsMode = Json.GetString(args, "boundsMode") ?? "center";
                var startIndex = Math.Max(0, Json.GetInt(args, "startIndex") ?? 0);
                var endIndex = Math.Min((int)doc.ModelSpace.Count - 1,
                    Json.GetInt(args, "endIndex") ?? ((int)doc.ModelSpace.Count - 1));

                var entities = new JsonArray();
                var count = 0;
                var scanned = 0;
                var lastScannedIndex = startIndex - 1;
                double? selectedMinX = null, selectedMinY = null, selectedMaxX = null, selectedMaxY = null;
                for (var index = startIndex; index <= endIndex; index++)
                {
                    dynamic ent = doc.ModelSpace.Item(index);
                    scanned++;
                    lastScannedIndex = index;
                    var type = (string)ent.EntityName;
                    string layer = "";
                    try { layer = (string)ent.Layer; } catch { }
                    if (layerFilter is not null && !layer.Equals(layerFilter, StringComparison.OrdinalIgnoreCase)) continue;
                    if (typeFilter is not null && !type.Equals(typeFilter, StringComparison.OrdinalIgnoreCase) &&
                        !type.Equals("AcDb" + typeFilter, StringComparison.OrdinalIgnoreCase)) continue;
                    if (textContains is not null && !TextOf(ent).Contains(textContains, StringComparison.OrdinalIgnoreCase)) continue;
                    if (blockName is not null)
                    {
                        string actualName = "";
                        try { actualName = (string)ent.EffectiveName; }
                        catch { try { actualName = (string)ent.Name; } catch { } }
                        if (!DocumentMatches(actualName, blockName)) continue;
                    }
                    if (bounds is not null && !BoundsMatch((object)ent,
                        Dbl(bounds["minX"]), Dbl(bounds["minY"]), Dbl(bounds["maxX"]), Dbl(bounds["maxY"]), boundsMode)) continue;

                    count++;
                    if (TryBoundingBox((object)ent, out var bx0, out var by0, out var bx1, out var by1))
                    {
                        selectedMinX = selectedMinX is null ? bx0 : Math.Min(selectedMinX.Value, bx0);
                        selectedMinY = selectedMinY is null ? by0 : Math.Min(selectedMinY.Value, by0);
                        selectedMaxX = selectedMaxX is null ? bx1 : Math.Max(selectedMaxX.Value, bx1);
                        selectedMaxY = selectedMaxY is null ? by1 : Math.Max(selectedMaxY.Value, by1);
                    }
                    if (!countOnly) entities.Add(EntityJson(ent, index, includeGeometry));
                    if (!countOnly && count >= limit) break;
                }

                string name = "";
                string fullName = "";
                try { name = (string)(doc.Name ?? ""); } catch { }
                try { fullName = (string)(doc.FullName ?? ""); } catch { }
                var hasMore = lastScannedIndex < endIndex;
                var nextActions = new JsonArray();
                if (hasMore)
                {
                    var continuation = QueryArguments(args, "entities");
                    continuation["startIndex"] = lastScannedIndex + 1;
                    continuation["endIndex"] = endIndex;
                    continuation["limit"] = limit;
                    nextActions.Add(CadQueryAction(
                        "현재 필터로 남은 ModelSpace 인덱스를 계속 조회", continuation));
                }
                var result = new JsonObject
                {
                    ["ok"] = true,
                    ["app"] = App,
                    ["documentRef"] = string.IsNullOrWhiteSpace(fullName) ? name : fullName,
                    ["entities"] = entities,
                    ["count"] = count,
                    ["scanned"] = scanned,
                    ["scanStartIndex"] = startIndex,
                    ["scanEndIndex"] = lastScannedIndex,
                    ["modelSpaceCount"] = (int)doc.ModelSpace.Count,
                    ["hasMore"] = hasMore,
                    ["nextStartIndex"] = hasMore ? lastScannedIndex + 1 : null,
                    ["truncated"] = hasMore,
                    ["coverage"] = new JsonObject
                    {
                        ["modelSpaceTotal"] = (int)doc.ModelSpace.Count,
                        ["requestedStartIndex"] = startIndex,
                        ["requestedEndIndex"] = endIndex,
                        ["scanned"] = scanned,
                        ["returned"] = countOnly ? 0 : entities.Count,
                        ["matched"] = count,
                        ["complete"] = !hasMore,
                    },
                    ["nextActions"] = nextActions,
                };
                if (selectedMinX is not null)
                    result["selectionBounds"] = new JsonObject
                    {
                        ["minX"] = selectedMinX,
                        ["minY"] = selectedMinY,
                        ["maxX"] = selectedMaxX,
                        ["maxY"] = selectedMaxY,
                    };
                return result;
            }
            catch (Exception ex) when (IsCallRejected(ex)) { throw; }
            catch (Exception ex) { return Json.ErrorResult($"cad_query_entities failed: {ex.Message}", App); }
            finally
            {
                if (closeAfterRead && queryDoc is not null)
                {
                    try { queryDoc.Close(false); } catch { }
                }
                _ = CompleteCadInteraction(foreground, documentState);
            }
        });
    }

    // ---------- preview ----------

    public override ApplyPreview Preview(IReadOnlyList<JsonObject> ops)
    {
        return ComInvokeWithRetry(() =>
        {
            var p = new ApplyPreview();
            var foreground = new ForegroundInteractionGuard(App);
            var documentState = new CadInteractionState();
            try
            {
                var app = AttachCad();
                if (app is null) { p.Errors.Add("AutoCAD not running"); return p; }
                TrackCadInteraction(app, foreground, documentState);
                dynamic d = app;
                var doc = ActiveDocWait(d);
                if (doc is null) { p.Errors.Add("열린 도면이 없습니다"); return p; }

                foreach (var op in ops)
                {
                    var name = Json.GetString(op, "op")!;
                    switch (name)
                    {
                        case "regen_document":
                            p.Affected.Add(new AffectedRef("display", "Regen all viewports; geometry unchanged"));
                            break;
                        case "activate_document":
                        {
                            var selector = Json.GetString(op, "document")!;
                            var target = FindOpenDocument(d, selector);
                            if (target is null)
                            {
                                p.Errors.Add($"CAD document not found: {selector}");
                                break;
                            }
                            doc = target;
                            string targetName = "";
                            try { targetName = (string)(doc.Name ?? selector); } catch { targetName = selector; }
                            p.Diff.Add(new DiffEntry { Ref = "active-document", Before = "preview selection", After = targetName });
                            p.Affected.Add(new AffectedRef("document", targetName));
                            break;
                        }
                        case "set_layer_visibility":
                        case "set_layer_color":
                        {
                            var layerName = Json.GetString(op, "layer")!;
                            dynamic layer;
                            try { layer = doc.Layers.Item(layerName); }
                            catch { p.Errors.Add($"layer '{layerName}' not found"); break; }
                            if (name == "set_layer_visibility")
                            {
                                var visible = Json.GetBool(op, "visible");
                                p.Diff.Add(new DiffEntry { Ref = $"layer:{layerName}", Before = $"LayerOn={layer.LayerOn}", After = $"LayerOn={visible}" });
                            }
                            else
                            {
                                var color = op["color"];
                                p.Diff.Add(new DiffEntry { Ref = $"layer:{layerName}", Before = $"Color={layer.Color}", After = $"Color={color}" });
                            }
                            p.Affected.Add(new AffectedRef("layer", layerName));
                            break;
                        }
                        case "move_entities":
                        case "rotate_entities":
                        case "delete_entities":
                        case "set_text_value":
                        {
                            var handles = new List<string>();
                            if (name == "set_text_value") handles.Add(Json.GetString(op, "handle")!);
                            else foreach (var hNode in Json.GetArr(op, "handles")!) handles.Add(hNode!.GetValue<string>());

                            foreach (var handle in handles.Take(MaxDiff))
                            {
                                dynamic ent;
                                try { ent = doc.HandleToObject(handle); }
                                catch { p.Errors.Add($"handle '{handle}' not found"); continue; }
                                var type = (string)ent.EntityName;
                                if (name == "set_text_value" && !IsTextLike(type))
                                {
                                    p.Errors.Add($"handle '{handle}' is {type} (not text-like)");
                                    continue;
                                }
                                if (name == "set_text_value")
                                    p.Diff.Add(new DiffEntry { Ref = $"entity:{handle}", Before = TextOf(ent), After = Json.GetString(op, "text") });
                                else
                                    p.Diff.Add(new DiffEntry { Ref = $"entity:{handle}", Before = type, After = name });
                            }
                            p.Affected.Add(new AffectedRef("entities", $"{handles.Count} handle(s)"));
                            if (name == "delete_entities") p.RequiresHighRiskApproval = true;
                            break;
                        }
                        case "delete_entities_in_bounds":
                        {
                            var bounds = Json.GetObj(op, "bounds")!;
                            var minX = Dbl(bounds["minX"]);
                            var minY = Dbl(bounds["minY"]);
                            var maxX = Dbl(bounds["maxX"]);
                            var maxY = Dbl(bounds["maxY"]);
                            if (maxX <= minX || maxY <= minY)
                            {
                                p.Errors.Add("delete_entities_in_bounds requires maxX>minX and maxY>minY");
                                break;
                            }
                            var handles = EntityHandlesInBounds(doc, minX, minY, maxX, maxY);
                            p.RequiresHighRiskApproval = true;
                            p.Affected.Add(new AffectedRef("entities", $"{handles.Count} inside bounds"));
                            p.Diff.Add(new DiffEntry
                            {
                                Ref = $"bounds:{minX},{minY}..{maxX},{maxY}",
                                Before = $"{handles.Count} entities",
                                After = "deleted",
                            });
                            break;
                        }
                        case "delete_entities_from_index":
                        {
                            var startIndex = Json.GetInt(op, "startIndex") ?? -1;
                            var handles = EntityHandlesFromIndex(doc, startIndex);
                            p.RequiresHighRiskApproval = true;
                            p.Affected.Add(new AffectedRef("entities", $"{handles.Count} from ModelSpace index {startIndex}"));
                            p.Diff.Add(new DiffEntry
                            {
                                Ref = $"modelspace-index:{startIndex}..{(int)doc.ModelSpace.Count - 1}",
                                Before = $"{handles.Count} entities",
                                After = "deleted",
                            });
                            break;
                        }
                        case "run_script_template":
                        {
                            p.RequiresHighRiskApproval = true;
                            var template = Json.GetString(op, "template")!;
                            var templatePath = ResolveTemplate(template);
                            if (templatePath is null)
                                p.Errors.Add($"template '{template}' not found in ops/script-templates (repo-registered templates only)");
                            else
                            {
                                try { _ = ExpandTemplate(templatePath, Json.GetObj(op, "params")); }
                                catch (Exception ex) { p.Errors.Add(ex.Message); break; }
                                p.Affected.Add(new AffectedRef("script", template));
                                p.Diff.Add(new DiffEntry { Ref = "template", Before = "", After = Path.GetFileName(templatePath) });
                            }
                            break;
                        }
                        case "copy_entities_between_documents":
                        {
                            object? sourceObject = null;
                            var closeAfter = false;
                            try
                            {
                                (sourceObject, closeAfter) = ResolveSourceDocument((object)d, (object)doc, op);
                                dynamic sourceDoc = sourceObject;
                                var selected = CollectCopyObjects(sourceDoc, op);
                                if (selected.Count == 0) p.Errors.Add("copy selection is empty");
                                var sourceName = "";
                                try { sourceName = (string)sourceDoc.Name; } catch { }
                                var targetName = "";
                                try { targetName = (string)doc.Name; } catch { }
                                p.Affected.Add(new AffectedRef("entities", $"{selected.Count} direct-COM copies: {sourceName} -> {targetName}"));
                                p.Diff.Add(new DiffEntry
                                {
                                    Ref = $"copy:{sourceName}",
                                    Before = $"{selected.Count} source entities",
                                    After = $"copied directly to {targetName}",
                                });
                            }
                            finally
                            {
                                if (closeAfter && sourceObject is not null)
                                {
                                    try { ((dynamic)sourceObject).Close(false); } catch { }
                                }
                            }
                            break;
                        }
                        case "insert_xref":
                        {
                            var sourceFile = Json.GetString(op, "sourceFile");
                            if (string.IsNullOrWhiteSpace(sourceFile) || !File.Exists(Path.GetFullPath(sourceFile)))
                                p.Errors.Add($"xref source DWG not found: {sourceFile}");
                            var point = Json.GetObj(op, "insertionPoint");
                            if (point is null) p.Errors.Add("insert_xref requires insertionPoint {x,y}");
                            var scale = op["scale"] is null ? 1.0 : Dbl(op["scale"]);
                            if (scale <= 0) p.Errors.Add("insert_xref scale must be positive");
                            p.Affected.Add(new AffectedRef("xref", sourceFile ?? ""));
                            p.Diff.Add(new DiffEntry
                            {
                                Ref = "direct-com-xref",
                                Before = "not attached",
                                After = $"{Path.GetFileName(sourceFile)} at ({Dbl(point?["x"])},{Dbl(point?["y"])})",
                            });
                            break;
                        }
                        case "zoom_window":
                        {
                            var bounds = Json.GetObj(op, "bounds")!;
                            var minX = Dbl(bounds["minX"]); var minY = Dbl(bounds["minY"]);
                            var maxX = Dbl(bounds["maxX"]); var maxY = Dbl(bounds["maxY"]);
                            if (maxX <= minX || maxY <= minY) p.Errors.Add("zoom_window requires maxX>minX and maxY>minY");
                            p.Affected.Add(new AffectedRef("view", $"{minX},{minY}..{maxX},{maxY}"));
                            p.Diff.Add(new DiffEntry { Ref = "active-view", Before = "current", After = "zoom window" });
                            break;
                        }
                        case "copy_entities":
                            if (Json.GetArr(op, "handles") is not { Count: > 0 }) p.Errors.Add("copy_entities requires handles");
                            p.Affected.Add(new AffectedRef("entities", $"copy by vector ({Dbl(op["dx"])},{Dbl(op["dy"])})"));
                            break;
                        case "scale_entities":
                            if (Json.GetArr(op, "handles") is not { Count: > 0 }) p.Errors.Add("scale_entities requires handles");
                            if (Json.GetArr(op, "basePoint") is not { Count: >= 2 }) p.Errors.Add("scale_entities requires basePoint");
                            if (Dbl(op["factor"]) <= 0) p.Errors.Add("scale_entities factor must be positive");
                            p.Affected.Add(new AffectedRef("entities", $"scale x{Dbl(op["factor"])}"));
                            break;
                        case "mirror_entities":
                            if (Json.GetArr(op, "handles") is not { Count: > 0 }) p.Errors.Add("mirror_entities requires handles");
                            if (Json.GetArr(op, "axisStart") is not { Count: >= 2 } || Json.GetArr(op, "axisEnd") is not { Count: >= 2 })
                                p.Errors.Add("mirror_entities requires axisStart and axisEnd");
                            p.Affected.Add(new AffectedRef("entities", "create mirrored copies"));
                            break;
                        case "offset_entities":
                            if (Json.GetArr(op, "handles") is not { Count: > 0 }) p.Errors.Add("offset_entities requires handles");
                            if (Math.Abs(Dbl(op["distance"])) < 1e-12) p.Errors.Add("offset_entities distance must be non-zero");
                            p.Affected.Add(new AffectedRef("entities", $"offset {Dbl(op["distance"])}"));
                            break;
                        case "set_entity_properties":
                            if (Json.GetArr(op, "handles") is not { Count: > 0 }) p.Errors.Add("set_entity_properties requires handles");
                            if (Json.GetObj(op, "properties") is null) p.Errors.Add("set_entity_properties requires properties");
                            p.Affected.Add(new AffectedRef("entities", "update common properties"));
                            break;
                        case "set_block_attributes":
                            if (string.IsNullOrWhiteSpace(Json.GetString(op, "handle"))) p.Errors.Add("set_block_attributes requires handle");
                            if (Json.GetObj(op, "attributes") is not { Count: > 0 }) p.Errors.Add("set_block_attributes requires attributes");
                            p.Affected.Add(new AffectedRef("block", Json.GetString(op, "handle") ?? ""));
                            break;
                        case "configure_layout":
                            if (string.IsNullOrWhiteSpace(Json.GetString(op, "name"))) p.Errors.Add("configure_layout requires name");
                            p.Affected.Add(new AffectedRef("layout", Json.GetString(op, "name") ?? ""));
                            break;
                        case "create_viewport":
                            if (string.IsNullOrWhiteSpace(Json.GetString(op, "layout"))) p.Errors.Add("create_viewport requires layout");
                            if (Json.GetArr(op, "center") is not { Count: >= 2 }) p.Errors.Add("create_viewport requires center");
                            if (Dbl(op["width"]) <= 0 || Dbl(op["height"]) <= 0 || Dbl(op["viewHeight"]) <= 0)
                                p.Errors.Add("create_viewport width, height and viewHeight must be positive");
                            p.Affected.Add(new AffectedRef("viewport", Json.GetString(op, "layout") ?? ""));
                            break;
                        case "save_document":
                            if (string.IsNullOrWhiteSpace(Json.GetString(op, "output")))
                            {
                                string currentPath = "";
                                try { currentPath = (string)(doc.FullName ?? ""); } catch { }
                                if (string.IsNullOrWhiteSpace(currentPath) || !Path.IsPathFullyQualified(currentPath))
                                    p.Errors.Add("unsaved drawing requires save_document.output");
                            }
                            p.Warnings.Add("save_document writes the active drawing; use a backup/output path and high-risk confirmation");
                            p.Affected.Add(new AffectedRef("drawing", Json.GetString(op, "output") ?? "active document"));
                            break;
                        case "plot_pdf":
                        {
                            var output = Json.GetString(op, "output");
                            if (string.IsNullOrWhiteSpace(output) || !Path.GetExtension(output).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
                                p.Errors.Add("plot_pdf requires a .pdf output path");
                            p.Warnings.Add("plot_pdf may replace an existing output file");
                            p.Affected.Add(new AffectedRef("pdf", output ?? ""));
                            break;
                        }
                        case "draw_entities":
                        {
                            var entities = Json.GetArr(op, "entities")!;
                            var byType = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
                            var ei = 0;
                            foreach (var eNode in entities)
                            {
                                ei++;
                                if (eNode is not JsonObject e) { p.Errors.Add($"entities[{ei}] is not an object"); continue; }
                                var etype = (Json.GetString(e, "type") ?? "").ToLowerInvariant();
                                switch (etype)
                                {
                                    case "lwpolyline":
                                        if (Json.GetArr(e, "points") is not { Count: >= 2 }) p.Errors.Add($"entities[{ei}] lwpolyline needs >= 2 points");
                                        break;
                                    case "circle":
                                        if (Json.GetArr(e, "center") is not { Count: >= 2 }) p.Errors.Add($"entities[{ei}] circle needs center [x,y]");
                                        break;
                                    case "block":
                                        if (Json.GetArr(e, "point") is not { Count: >= 2 }) p.Errors.Add($"entities[{ei}] block needs point [x,y]");
                                        if (string.IsNullOrWhiteSpace(Json.GetString(e, "name"))) p.Errors.Add($"entities[{ei}] block name is empty");
                                        break;
                                    case "text":
                                        if (Json.GetArr(e, "point") is not { Count: >= 2 }) p.Errors.Add($"entities[{ei}] text needs point [x,y]");
                                        if (string.IsNullOrEmpty(Json.GetString(e, "text"))) p.Errors.Add($"entities[{ei}] text value is empty");
                                        if (e["height"] is null || Dbl(e["height"]) <= 0) p.Errors.Add($"entities[{ei}] text height must be positive");
                                        break;
                                    case "hatch":
                                        if (Json.GetArr(Json.GetObj(e, "loop"), "points") is not { Count: >= 2 }) p.Errors.Add($"entities[{ei}] hatch needs loop.points >= 2");
                                        break;
                                    case "line":
                                        if (Json.GetArr(e, "start") is not { Count: >= 2 } || Json.GetArr(e, "end") is not { Count: >= 2 })
                                            p.Errors.Add($"entities[{ei}] line needs start/end");
                                        break;
                                    case "arc":
                                        if (Json.GetArr(e, "center") is not { Count: >= 2 } || Dbl(e["radius"]) <= 0)
                                            p.Errors.Add($"entities[{ei}] arc needs center and positive radius");
                                        break;
                                    case "ellipse":
                                    {
                                        var ratio = Dbl(e["radiusRatio"]);
                                        if (Json.GetArr(e, "center") is not { Count: >= 2 } || Json.GetArr(e, "majorAxis") is not { Count: >= 2 } || ratio <= 0 || ratio > 1)
                                            p.Errors.Add($"entities[{ei}] ellipse needs center, majorAxis and radiusRatio 0..1");
                                        break;
                                    }
                                    case "point":
                                        if (Json.GetArr(e, "point") is not { Count: >= 2 }) p.Errors.Add($"entities[{ei}] point needs point");
                                        break;
                                    case "mtext":
                                        if (Json.GetArr(e, "point") is not { Count: >= 2 } || Dbl(e["width"]) <= 0)
                                            p.Errors.Add($"entities[{ei}] mtext needs point and positive width");
                                        break;
                                    case "dim_aligned":
                                        if (Json.GetArr(e, "start") is not { Count: >= 2 } || Json.GetArr(e, "end") is not { Count: >= 2 } || Json.GetArr(e, "textPoint") is not { Count: >= 2 })
                                            p.Errors.Add($"entities[{ei}] dim_aligned needs start/end/textPoint");
                                        break;
                                    case "dim_rotated":
                                        if (Json.GetArr(e, "start") is not { Count: >= 2 } || Json.GetArr(e, "end") is not { Count: >= 2 } || Json.GetArr(e, "dimensionLinePoint") is not { Count: >= 2 })
                                            p.Errors.Add($"entities[{ei}] dim_rotated needs start/end/dimensionLinePoint");
                                        break;
                                    default:
                                        p.Errors.Add($"entities[{ei}] unknown type '{etype}'");
                                        break;
                                }
                                byType[etype] = byType.GetValueOrDefault(etype) + 1;
                            }
                            if (entities.Count > MaxDrawEntities) p.Errors.Add($"entities count {entities.Count} exceeds limit {MaxDrawEntities}");
                            p.Affected.Add(new AffectedRef("entities", $"{entities.Count} to draw ({string.Join(", ", byType.Select(kv => $"{kv.Key}:{kv.Value}"))})"));
                            p.Diff.Add(new DiffEntry { Ref = "modelspace", Before = "", After = $"+{entities.Count} entities" });
                            break;
                        }
                        case "draw_taegeukgi":
                        {
                            p.Affected.Add(new AffectedRef("modelspace", "direct COM Taegeukgi"));
                            p.Diff.Add(new DiffEntry
                            {
                                Ref = "modelspace",
                                Before = $"{EntityCount(doc)} entities",
                                After = "+136 direct COM entities (no AutoLISP)",
                            });
                            break;
                        }
                        case "draw_union_jack":
                        {
                            var width = op["width"] is null ? 120.0 : Dbl(op["width"]);
                            var height = op["height"] is null ? 60.0 : Dbl(op["height"]);
                            if (width <= 0 || height <= 0) p.Errors.Add("draw_union_jack width and height must be positive");
                            p.Affected.Add(new AffectedRef("modelspace", "direct COM Union Jack"));
                            p.Diff.Add(new DiffEntry
                            {
                                Ref = "modelspace",
                                Before = $"{EntityCount(doc)} entities",
                                After = "+27 direct COM entities (no AutoLISP)",
                            });
                            break;
                        }
                        case "draw_block_wall_schematic":
                        {
                            var scale = op["scale"] is null ? 1.0 : Dbl(op["scale"]);
                            if (scale <= 0) p.Errors.Add("draw_block_wall_schematic scale must be positive");
                            p.Affected.Add(new AffectedRef("modelspace", "1:1 mm modular block-wall construction drawing"));
                            p.Diff.Add(new DiffEntry
                            {
                                Ref = "modelspace",
                                Before = $"{EntityCount(doc)} entities",
                                After = "+22,500x5,400 mm wall; exact 600/400/200 mm modules; layered dimensions and details",
                            });
                            break;
                        }
                    }
                    if (!foreground.Checkpoint(stopOnConcurrentInput: true))
                    {
                        p.Errors.Add("[APP_USER_ACTIVITY_DETECTED] 사용자가 AutoCAD 창을 조작하여 미리보기를 중단했습니다. 해당 창 작업을 마친 뒤 다시 실행하세요.");
                        break;
                    }
                }
            }
            catch (Exception ex) when (IsCallRejected(ex)) { throw; }
            catch (Exception ex) { p.Errors.Add($"preview failed: {ex.Message}"); }
            finally { p.Interaction = CompleteCadInteraction(foreground, documentState); }
            return p;
        });
    }

    /// <summary>repo 등록 템플릿만 허용 (임의 스크립트 차단)</summary>
    internal static string? ResolveTemplate(string name)
    {
        if (name.Contains("..") || name.Contains('/') || name.Contains('\\')) return null;
        var dir = AppContext.BaseDirectory;
        for (var i = 0; i < 8 && dir is not null; i++)
        {
            var candidate = Path.Combine(dir, "ops", "script-templates", name.EndsWith(".scr") ? name : name + ".scr");
            if (File.Exists(candidate)) return candidate;
            dir = Directory.GetParent(dir)?.FullName;
        }
        return null;
    }

    private static string ExpandTemplate(string templatePath, JsonObject? parameters)
    {
        var script = File.ReadAllText(templatePath).Replace("{{DOCBRIDGE_TEMPLATE_DIR}}",
            (Path.GetDirectoryName(templatePath) ?? "").Replace('\\', '/'));
        if (parameters is not null)
            foreach (var (key, node) in parameters)
            {
                if (string.Equals(key, "DOCBRIDGE_TEMPLATE_DIR", StringComparison.Ordinal))
                    throw new InvalidOperationException("reserved template parameter: DOCBRIDGE_TEMPLATE_DIR");
                var value = node?.GetValue<string>() ?? "";
                if (value.Length > 256 || value.IndexOfAny(new[] { '\r', '\n', '\0', '\u001b' }) >= 0 ||
                    value.Contains("{{", StringComparison.Ordinal) || value.Contains("}}", StringComparison.Ordinal))
                    throw new InvalidOperationException(
                        $"unsafe template parameter '{key}': control characters, placeholders, and values over 256 chars are not allowed");
                script = script.Replace("{{" + key + "}}", value);
            }
        if (script.Contains("{{", StringComparison.Ordinal))
            throw new InvalidOperationException("unresolved template placeholders remain");
        return script;
    }

    // ---------- apply ----------

    public override ApplyExecution Apply(IReadOnlyList<JsonObject> ops, string snapshotId)
    {
        // A batch is not idempotent: retrying from the beginning can move/scale twice.
        return ComInvoke(() =>
        {
            var exec = new ApplyExecution { Ok = true };
            var displayRefresh = new CadDisplayRefresh();
            var checkedCount = 0;
            var mismatches = new List<string>();
            var foreground = new ForegroundInteractionGuard(App);
            var documentState = new CadInteractionState();
            var userActivityInterrupted = false;
            try
            {
                var app = AttachCad();
                if (app is null) { exec.Errors.Add("AutoCAD not running"); exec.Ok = false; return exec; }
                TrackCadInteraction(app, foreground, documentState);
                dynamic d = app;
                var doc = ActiveDocWait(d);
                if (doc is null) { exec.Errors.Add("열린 도면이 없습니다"); exec.Ok = false; return exec; }

                foreach (var op in ops)
                {
                    var name = Json.GetString(op, "op")!;
                    var opStarted = Stopwatch.StartNew();
                    var mismatchCountBefore = mismatches.Count;
                    string? opError = null;
                    var opOk = false;
                    try
                    {
                      if (name is not ("activate_document" or "save_document" or "plot_pdf" or "zoom_window"))
                          displayRefresh.Track((object)doc);
                      switch (name)
                      {
                        case "regen_document":
                            exec.Affected.Add(new AffectedRef("display", "queued all-viewport regeneration"));
                            break;
                        case "activate_document":
                        {
                            var selector = Json.GetString(op, "document")!;
                            var target = FindOpenDocument(d, selector)
                                ?? throw new InvalidOperationException($"CAD document not found: {selector}");
                            target.Activate();
                            doc = target;
                            string targetName = "";
                            try { targetName = (string)(doc.Name ?? selector); } catch { targetName = selector; }
                            checkedCount++;
                            string activeName = "";
                            try { activeName = (string)(d.ActiveDocument?.Name ?? ""); } catch { }
                            if (!DocumentMatches(activeName, targetName))
                                mismatches.Add($"activate_document: wanted '{targetName}', got '{activeName}'");
                            exec.Affected.Add(new AffectedRef("document", targetName));
                            break;
                        }
                        case "set_layer_visibility":
                        {
                            var layerName = Json.GetString(op, "layer")!;
                            var visible = Json.GetBool(op, "visible");
                            dynamic layer = doc.Layers.Item(layerName);
                            layer.LayerOn = visible;
                            checkedCount++;
                            if ((bool)layer.LayerOn != visible) mismatches.Add($"layer {layerName}: LayerOn readback failed");
                            exec.Affected.Add(new AffectedRef("layer", layerName));
                            break;
                        }
                        case "set_layer_color":
                        {
                            var layerName = Json.GetString(op, "layer")!;
                            dynamic layer = doc.Layers.Item(layerName);
                            var colorNode = op["color"];
                            var before = (int)layer.Color;
                            if (colorNode is JsonValue jv && jv.TryGetValue<int>(out var aci))
                                layer.Color = aci; // ACI 0-256
                            else { exec.Warnings.Add($"color must be ACI int in MVP: {colorNode}"); break; }
                            checkedCount++;
                            if ((int)layer.Color != aci) mismatches.Add($"layer {layerName}: Color readback failed");
                            exec.Diff.Add(new DiffEntry { Ref = $"layer:{layerName}", Before = before, After = aci });
                            exec.Affected.Add(new AffectedRef("layer", layerName));
                            break;
                        }
                        case "move_entities":
                        {
                            var dx = op["dx"]!.GetValue<double>();
                            var dy = op["dy"]!.GetValue<double>();
                            var moved = 0;
                            foreach (var hNode in Json.GetArr(op, "handles")!)
                            {
                                dynamic ent = doc.HandleToObject(hNode!.GetValue<string>());
                                ent.Move(Point(0, 0, 0), Point(dx, dy, 0));
                                moved++;
                                checkedCount++;
                            }
                            exec.Affected.Add(new AffectedRef("entities", $"{moved} moved by ({dx},{dy})"));
                            break;
                        }
                        case "rotate_entities":
                        {
                            var angleDeg = op["angleDeg"]!.GetValue<double>();
                            var angleRad = angleDeg * Math.PI / 180.0;
                            var bp = Json.GetObj(op, "basePoint");
                            var bx = bp is null ? 0.0 : Dbl(bp["x"]);
                            var by = bp is null ? 0.0 : Dbl(bp["y"]);
                            var rotated = 0;
                            foreach (var hNode in Json.GetArr(op, "handles")!)
                            {
                                dynamic ent = doc.HandleToObject(hNode!.GetValue<string>());
                                ent.Rotate(Point(bx, by, 0), angleRad);
                                rotated++;
                                checkedCount++;
                            }
                            exec.Affected.Add(new AffectedRef("entities", $"{rotated} rotated {angleDeg}deg"));
                            break;
                        }
                        case "set_text_value":
                        {
                            var handle = Json.GetString(op, "handle")!;
                            var text = Json.GetString(op, "text")!;
                            dynamic ent = doc.HandleToObject(handle);
                            var type = (string)ent.EntityName;
                            if (!IsTextLike(type)) { mismatches.Add($"handle '{handle}' is {type} (not text-like)"); break; }
                            var before = TextOf(ent);
                            ent.TextString = text;
                            checkedCount++;
                            var got = TextOf(ent);
                            if (got != text) mismatches.Add($"entity {handle}: want '{text}', got '{got}'");
                            exec.Diff.Add(new DiffEntry { Ref = $"entity:{handle}", Before = before, After = got });
                            exec.Affected.Add(new AffectedRef("entity", handle));
                            break;
                        }
                        case "delete_entities":
                        {
                            var deleted = 0;
                            foreach (var hNode in Json.GetArr(op, "handles")!)
                            {
                                try
                                {
                                    dynamic ent = doc.HandleToObject(hNode!.GetValue<string>());
                                    ent.Delete();
                                    deleted++;
                                    checkedCount++;
                                }
                                catch (Exception ex) { exec.Warnings.Add($"delete '{hNode}': {ex.Message}"); }
                            }
                            exec.Affected.Add(new AffectedRef("entities", $"{deleted} deleted"));
                            break;
                        }
                        case "delete_entities_in_bounds":
                        {
                            var bounds = Json.GetObj(op, "bounds")!;
                            var handles = EntityHandlesInBounds(doc,
                                Dbl(bounds["minX"]), Dbl(bounds["minY"]),
                                Dbl(bounds["maxX"]), Dbl(bounds["maxY"]));
                            var before = EntityCount(doc);
                            var deleted = 0;
                            foreach (var handle in handles)
                            {
                                dynamic entity = doc.HandleToObject(handle);
                                entity.Delete();
                                deleted++;
                            }
                            var after = EntityCount(doc);
                            checkedCount += deleted;
                            if (before >= 0 && after != before - deleted)
                                mismatches.Add($"delete_entities_in_bounds: expected {before - deleted} entities, got {after}");
                            exec.Affected.Add(new AffectedRef("entities", $"{deleted} deleted inside bounds"));
                            break;
                        }
                        case "delete_entities_from_index":
                        {
                            var startIndex = Json.GetInt(op, "startIndex") ?? -1;
                            var handles = EntityHandlesFromIndex(doc, startIndex);
                            var before = EntityCount(doc);
                            foreach (var handle in handles)
                            {
                                dynamic entity = doc.HandleToObject(handle);
                                entity.Delete();
                            }
                            var after = EntityCount(doc);
                            checkedCount += handles.Count;
                            if (before >= 0 && after != startIndex)
                                mismatches.Add($"delete_entities_from_index: expected {startIndex} entities, got {after}");
                            exec.Affected.Add(new AffectedRef("entities", $"{handles.Count} deleted from ModelSpace index {startIndex}"));
                            break;
                        }
                        case "run_script_template":
                        {
                            var template = Json.GetString(op, "template")!;
                            var templatePath = ResolveTemplate(template)
                                ?? throw new InvalidOperationException($"template '{template}' not registered");
                            var script = ExpandTemplate(templatePath, Json.GetObj(op, "params"));
                            if (!WaitCmdIdle(doc, 15))
                                throw new InvalidOperationException("AutoCAD command line was not idle before template execution");
                            doc.SendCommand(script + "\n");
                            checkedCount++;
                            if (!WaitCmdIdle(doc, 60))
                                mismatches.Add($"template '{template}' did not return AutoCAD to an idle command state");
                            exec.Affected.Add(new AffectedRef("script", template));
                            break;
                        }
                        case "copy_entities_between_documents":
                        {
                            var before = EntityCount(doc);
                            var copied = CopyEntitiesDirect((object)d, (object)doc, op, exec.Warnings);
                            var after = EntityCount(doc);
                            checkedCount += copied.Count;
                            if (before >= 0 && after != before + copied.Count)
                                mismatches.Add($"copy_entities_between_documents: expected {before + copied.Count} target entities, got {after}");
                            var copiedHandles = new JsonArray();
                            foreach (dynamic entity in copied.Take(100))
                            {
                                try { copiedHandles.Add((string)entity.Handle); } catch { }
                            }
                            exec.Diff.Add(new DiffEntry
                            {
                                Ref = "direct-com-copy",
                                Before = before,
                                After = after,
                            });
                            exec.Affected.Add(new AffectedRef("entities", $"{copied.Count} copied directly via COM"));
                            break;
                        }
                        case "insert_xref":
                        {
                            var before = EntityCount(doc);
                            dynamic xref = InsertXrefDirect(doc, op, exec.Warnings);
                            var after = EntityCount(doc);
                            checkedCount++;
                            if (before >= 0 && after != before + 1)
                                mismatches.Add($"insert_xref: expected {before + 1} target entities, got {after}");
                            string handle = "";
                            try { handle = (string)xref.Handle; } catch { }
                            exec.Affected.Add(new AffectedRef("xref", handle));
                            exec.Diff.Add(new DiffEntry { Ref = $"xref:{handle}", Before = "", After = Json.GetString(op, "sourceFile") });
                            break;
                        }
                        case "zoom_window":
                        {
                            var bounds = Json.GetObj(op, "bounds")!;
                            d.ZoomWindow(
                                Point(Dbl(bounds["minX"]), Dbl(bounds["minY"])),
                                Point(Dbl(bounds["maxX"]), Dbl(bounds["maxY"])));
                            checkedCount++;
                            exec.Affected.Add(new AffectedRef("view", "active AutoCAD window zoomed"));
                            break;
                        }
                        case "copy_entities":
                        {
                            var before = EntityCount(doc);
                            var copied = CopyEntitiesByVector(doc, op);
                            var after = EntityCount(doc);
                            checkedCount += copied.Count;
                            if (before >= 0 && after != before + copied.Count) mismatches.Add($"copy_entities: expected {before + copied.Count}, got {after}");
                            exec.Affected.Add(new AffectedRef("entities", $"{copied.Count} copied by vector"));
                            break;
                        }
                        case "scale_entities":
                        {
                            var changed = ScaleEntities(doc, op);
                            checkedCount += changed;
                            exec.Affected.Add(new AffectedRef("entities", $"{changed} scaled"));
                            break;
                        }
                        case "mirror_entities":
                        {
                            var before = EntityCount(doc);
                            var mirrored = MirrorEntities(doc, op);
                            var after = EntityCount(doc);
                            checkedCount += mirrored.Count;
                            if (before >= 0 && after != before + mirrored.Count) mismatches.Add($"mirror_entities: expected {before + mirrored.Count}, got {after}");
                            exec.Affected.Add(new AffectedRef("entities", $"{mirrored.Count} mirrored copies"));
                            break;
                        }
                        case "offset_entities":
                        {
                            var before = EntityCount(doc);
                            var offset = OffsetEntities(doc, op);
                            var after = EntityCount(doc);
                            checkedCount += offset.Count;
                            if (before >= 0 && after != before + offset.Count) mismatches.Add($"offset_entities: expected {before + offset.Count}, got {after}");
                            exec.Affected.Add(new AffectedRef("entities", $"{offset.Count} offset entities"));
                            break;
                        }
                        case "set_entity_properties":
                        {
                            var changed = SetEntityProperties(d, doc, op, exec.Warnings);
                            checkedCount += changed;
                            exec.Affected.Add(new AffectedRef("entities", $"{changed} properties updated"));
                            break;
                        }
                        case "set_block_attributes":
                        {
                            var changed = SetBlockAttributes(doc, op, exec.Warnings);
                            checkedCount += changed;
                            exec.Affected.Add(new AffectedRef("block", $"{changed} attributes updated"));
                            break;
                        }
                        case "configure_layout":
                        {
                            dynamic layout = ConfigureLayout(doc, op, exec.Warnings);
                            checkedCount++;
                            exec.Affected.Add(new AffectedRef("layout", (string)layout.Name));
                            break;
                        }
                        case "create_viewport":
                        {
                            dynamic viewport = CreateViewport(doc, op);
                            checkedCount++;
                            exec.Affected.Add(new AffectedRef("viewport", (string)viewport.Handle));
                            break;
                        }
                        case "save_document":
                        {
                            var output = SaveDocument(doc, op);
                            checkedCount++;
                            exec.Affected.Add(new AffectedRef("drawing", output));
                            break;
                        }
                        case "plot_pdf":
                        {
                            var output = PlotPdf(doc, op);
                            checkedCount++;
                            exec.Affected.Add(new AffectedRef("pdf", output));
                            break;
                        }
                        case "draw_entities":
                        {
                            var entities = Json.GetArr(op, "entities")!;
                            if (entities.Count > MaxDrawEntities)
                            {
                                mismatches.Add($"draw_entities: {entities.Count} entities exceeds limit {MaxDrawEntities}");
                                break;
                            }
                            var drawn = 0;
                            var idx = 0;
                            foreach (var eNode in entities)
                            {
                                idx++;
                                if (eNode is not JsonObject e) { mismatches.Add($"entities[{idx}] is not an object"); continue; }
                                var etype = (Json.GetString(e, "type") ?? "").ToLowerInvariant();
                                try
                                {
                                    switch (etype)
                                    {
                                        case "lwpolyline":
                                        {
                                            var pl = AddLwPolyline(doc, Json.GetArr(e, "points")!, null, Json.GetBool(e, "closed"));
                                            var layerName = Json.GetString(e, "layer");
                                            if (!string.IsNullOrWhiteSpace(layerName))
                                            {
                                                EnsureLayer(doc, layerName);
                                                ((dynamic)pl).Layer = layerName;
                                            }
                                            SetEntityColor(d, pl, Json.GetObj(e, "color"), exec.Warnings);
                                            checkedCount++;
                                            var got = (string)((dynamic)pl).EntityName;
                                            if (got != "AcDbPolyline") mismatches.Add($"entities[{idx}]: want AcDbPolyline, got {got}");
                                            drawn++;
                                            break;
                                        }
                                        case "circle":
                                        {
                                            var center = Json.GetArr(e, "center")!;
                                            var radius = Dbl(e["radius"]);
                                            if (radius <= 0) { mismatches.Add($"entities[{idx}] circle radius must be > 0"); break; }
                                            dynamic circle = doc.ModelSpace.AddCircle(Point(Dbl(center[0]), Dbl(center[1])), radius);
                                            var layerName = Json.GetString(e, "layer");
                                            if (!string.IsNullOrWhiteSpace(layerName))
                                            {
                                                EnsureLayer(doc, layerName);
                                                circle.Layer = layerName;
                                            }
                                            SetEntityColor(d, circle, Json.GetObj(e, "color"), exec.Warnings);
                                            checkedCount++;
                                            drawn++;
                                            break;
                                        }
                                        case "block":
                                        {
                                            var point = Json.GetArr(e, "point")!;
                                            var blockName = Json.GetString(e, "name")!;
                                            try { _ = doc.Blocks.Item(blockName); }
                                            catch { mismatches.Add($"entities[{idx}] block definition '{blockName}' not found"); break; }
                                            var xScale = e["xScale"] is null ? 1.0 : Dbl(e["xScale"]);
                                            var yScale = e["yScale"] is null ? xScale : Dbl(e["yScale"]);
                                            var zScale = e["zScale"] is null ? 1.0 : Dbl(e["zScale"]);
                                            if (xScale <= 0 || yScale <= 0 || zScale <= 0)
                                            {
                                                mismatches.Add($"entities[{idx}] block scales must be > 0");
                                                break;
                                            }
                                            var rotation = e["rotationDeg"] is null ? 0.0 : Dbl(e["rotationDeg"]) * Math.PI / 180.0;
                                            dynamic block = doc.ModelSpace.InsertBlock(
                                                Point(Dbl(point[0]), Dbl(point[1])), blockName,
                                                xScale, yScale, zScale, rotation);
                                            var layerName = Json.GetString(e, "layer");
                                            if (!string.IsNullOrWhiteSpace(layerName))
                                            {
                                                EnsureLayer(doc, layerName);
                                                block.Layer = layerName;
                                            }
                                            SetEntityColor(d, block, Json.GetObj(e, "color"), exec.Warnings);
                                            checkedCount++;
                                            var got = (string)block.EntityName;
                                            if (got != "AcDbBlockReference")
                                                mismatches.Add($"entities[{idx}]: want AcDbBlockReference, got {got}");
                                            drawn++;
                                            break;
                                        }
                                        case "text":
                                        {
                                            var point = Json.GetArr(e, "point")!;
                                            var text = Json.GetString(e, "text")!;
                                            var height = Dbl(e["height"]);
                                            if (height <= 0) { mismatches.Add($"entities[{idx}] text height must be > 0"); break; }
                                            dynamic textEntity = doc.ModelSpace.AddText(text, Point(Dbl(point[0]), Dbl(point[1])), height);
                                            var layerName = Json.GetString(e, "layer");
                                            if (!string.IsNullOrWhiteSpace(layerName))
                                            {
                                                EnsureLayer(doc, layerName);
                                                textEntity.Layer = layerName;
                                            }
                                            var styleName = Json.GetString(e, "style");
                                            if (!string.IsNullOrWhiteSpace(styleName))
                                            {
                                                try { textEntity.StyleName = styleName; }
                                                catch (Exception ex) { exec.Warnings.Add($"entities[{idx}] text style '{styleName}' not applied: {ex.Message}"); }
                                            }
                                            if (e["rotationDeg"] is not null) textEntity.Rotation = Dbl(e["rotationDeg"]) * Math.PI / 180.0;
                                            if (e["alignment"] is not null)
                                            {
                                                var alignment = (int)Dbl(e["alignment"]);
                                                textEntity.Alignment = alignment;
                                                if (alignment != 0) textEntity.TextAlignmentPoint = Point(Dbl(point[0]), Dbl(point[1]));
                                            }
                                            SetEntityColor(d, textEntity, Json.GetObj(e, "color"), exec.Warnings);
                                            checkedCount++;
                                            drawn++;
                                            break;
                                        }
                                        case "hatch":
                                        {
                                            var loop = Json.GetObj(e, "loop")!;
                                            var before = EntityCount(doc);
                                            if (before < 0) throw new InvalidOperationException("ModelSpace.Count failed");
                                            dynamic hatch = AddHatchDirect(
                                                d, doc, loop, Json.GetObj(e, "color"), exec.Warnings,
                                                Json.GetString(e, "layer"));
                                            checkedCount++;
                                            var after = EntityCount(doc);
                                            if (after != before + 1) mismatches.Add($"entities[{idx}] hatch: expected {before + 1} entities, got {after}");
                                            string gotType = "";
                                            try { gotType = (string)hatch.EntityName; } catch { }
                                            if (gotType != "AcDbHatch") mismatches.Add($"entities[{idx}]: want AcDbHatch, got {gotType}");
                                            else drawn++;
                                            break;
                                        }
                                        case "line":
                                        case "arc":
                                        case "ellipse":
                                        case "point":
                                        case "mtext":
                                        case "dim_aligned":
                                        case "dim_rotated":
                                        {
                                            dynamic created = AddProductionEntity(d, doc, e, exec.Warnings);
                                            string gotType = "";
                                            try { gotType = (string)created.EntityName; } catch { }
                                            if (string.IsNullOrWhiteSpace(gotType)) mismatches.Add($"entities[{idx}] {etype}: readback EntityName is empty");
                                            else { checkedCount++; drawn++; }
                                            break;
                                        }
                                        default:
                                            mismatches.Add($"entities[{idx}] unknown type '{etype}'");
                                            break;
                                    }
                                }
                                catch (Exception ex) when (IsCallRejected(ex)) { throw; }
                                catch (Exception ex) { mismatches.Add($"entities[{idx}] {etype}: {ex.Message}"); }
                            }
                            exec.Affected.Add(new AffectedRef("entities", $"{drawn} drawn via COM"));
                            break;
                        }
                        case "draw_taegeukgi":
                        {
                            var before = EntityCount(doc);
                            if (before < 0) throw new InvalidOperationException("ModelSpace.Count failed before direct Taegeukgi draw");
                            var drawn = DrawTaegeukgiDirect(d, doc, exec.Warnings);
                            var after = EntityCount(doc);
                            checkedCount += drawn;
                            if (after != before + drawn)
                                mismatches.Add($"draw_taegeukgi: expected {before + drawn} entities, got {after}");
                            exec.Affected.Add(new AffectedRef("entities", $"{drawn} Taegeukgi entities drawn directly via COM"));
                            break;
                        }
                        case "draw_union_jack":
                        {
                            var before = EntityCount(doc);
                            if (before < 0) throw new InvalidOperationException("ModelSpace.Count failed before direct Union Jack draw");
                            var originX = op["originX"] is null ? 105.0 : Dbl(op["originX"]);
                            var originY = op["originY"] is null ? 0.0 : Dbl(op["originY"]);
                            var width = op["width"] is null ? 120.0 : Dbl(op["width"]);
                            var height = op["height"] is null ? 60.0 : Dbl(op["height"]);
                            var drawn = DrawUnionJackDirect(d, doc, exec.Warnings, originX, originY, width, height);
                            var after = EntityCount(doc);
                            checkedCount += drawn;
                            if (after != before + drawn)
                                mismatches.Add($"draw_union_jack: expected {before + drawn} entities, got {after}");
                            exec.Affected.Add(new AffectedRef("entities", $"{drawn} Union Jack entities drawn directly via COM"));
                            break;
                        }
                        case "draw_block_wall_schematic":
                        {
                            var before = EntityCount(doc);
                            if (before < 0) throw new InvalidOperationException("ModelSpace.Count failed before block-wall schematic draw");
                            var originX = op["originX"] is null ? 2000.0 : Dbl(op["originX"]);
                            var originY = op["originY"] is null ? 0.0 : Dbl(op["originY"]);
                            var scale = op["scale"] is null ? 1.0 : Dbl(op["scale"]);
                            var drawn = DrawBlockWallInstallationDirect(d, doc, exec.Warnings, originX, originY, scale);
                            var after = EntityCount(doc);
                            checkedCount += drawn;
                            if (after != before + drawn)
                                mismatches.Add($"draw_block_wall_schematic: expected {before + drawn} entities, got {after}");
                            exec.Affected.Add(new AffectedRef("entities", $"{drawn} modular construction-drawing entities drawn directly via COM"));
                            break;
                        }
                    }
                }

                    catch (Exception ex)
                    {
                        opError = ex.Message;
                        mismatches.Add($"{name}: {ex.Message}");
                    }
                    finally
                    {
                        opStarted.Stop();
                        opOk = opError is null && mismatches.Count == mismatchCountBefore;
                        exec.OperationResults.Add(new JsonObject
                        {
                            ["index"] = exec.OperationResults.Count,
                            ["op"] = name,
                            ["ok"] = opOk,
                            ["elapsedMs"] = opStarted.ElapsedMilliseconds,
                            ["error"] = opError ?? (mismatches.Count > mismatchCountBefore
                                ? mismatches[mismatchCountBefore]
                                : null),
                        });
                    }
                    if (!opOk) break;
                    if (!foreground.Checkpoint(stopOnConcurrentInput: true))
                    {
                        userActivityInterrupted = true;
                        exec.Errors.Add("[APP_USER_ACTIVITY_DETECTED] 사용자가 AutoCAD 창을 조작하여 남은 작업을 안전하게 중단했습니다. 도면을 다시 읽은 뒤 이어서 실행하세요.");
                        break;
                    }
                }

                exec.Readback = new JsonObject
                {
                    ["verified"] = mismatches.Count == 0,
                    ["checked"] = checkedCount,
                    ["mismatches"] = Json.ToArray(mismatches),
                    ["snapshotId"] = snapshotId,
                };
                exec.Ok = mismatches.Count == 0 && !userActivityInterrupted;
            }
            catch (Exception ex) { exec.Ok = false; exec.Errors.Add($"apply failed: {ex.Message}"); }
            finally
            {
                // Restore the view first, then regenerate once per touched drawing before
                // the foreground guard completes. Never replay edits on a Regen failure.
                documentState.Restore();
                displayRefresh.Complete(exec);
                exec.Interaction = CompleteCadInteraction(foreground, documentState);
            }
            return exec;
        }, timeoutSec: 600);
    }

    // ---------- snapshot / restore ----------

    public override void CaptureSnapshot(string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject>? ops = null)
    {
        ComInvokeWithRetry(() =>
        {
            var app = AttachCad();
            if (app is null) { metadata["payload"] = "none (autocad not running)"; return; }
            dynamic d = app;
            var doc = ActiveDocWait(d);
            if (doc is null) { metadata["payload"] = "none (no drawing)"; return; }

            string fullName = "";
            try { fullName = (string)(doc.FullName ?? ""); } catch { }
            var savedAtSnapshot = false;
            try { savedAtSnapshot = (bool)doc.Saved; } catch { }
            if (!string.IsNullOrEmpty(fullName) && File.Exists(fullName))
            {
                var dest = Path.Combine(snapshotDir, "drawing-backup" + Path.GetExtension(fullName));
                try
                {
                    using var src = new FileStream(fullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    using var dst = new FileStream(dest, FileMode.Create, FileAccess.Write);
                    src.CopyTo(dst);
                    metadata["drawingBackup"] = Path.GetFileName(dest);
                    metadata["fileSha256"] = CadFileHash(dest);
                }
                catch (Exception ex) { metadata["drawingBackupError"] = ex.Message; }
            }

            var layers = new JsonObject();
            foreach (dynamic layer in doc.Layers)
            {
                layers[(string)layer.Name] = new JsonObject
                {
                    ["on"] = (bool)layer.LayerOn,
                    ["color"] = (int)layer.Color,
                };
            }
            var texts = new JsonObject();
            var tcount = 0;
            foreach (dynamic ent in doc.ModelSpace)
            {
                var type = (string)ent.EntityName;
                if (!IsTextLike(type)) continue;
                try { texts[(string)ent.Handle] = TextOf(ent); } catch { }
                if (++tcount >= 500) break;
            }

            File.WriteAllText(Path.Combine(snapshotDir, "state.json"),
                new JsonObject
                {
                    ["fullName"] = fullName,
                    ["layers"] = layers,
                    ["texts"] = texts,
                }.ToJsonString(Json.Pretty));
            metadata["payload"] = "drawing-backup + state.json";
            metadata["documentRef"] = string.IsNullOrEmpty(fullName) ? metadata["documentRef"] : fullName;
            metadata["savedAtSnapshot"] = savedAtSnapshot;
            if (ops is { Count: > 0 })
            {
                metadata["operationStateSha256"] = CadOperationStateHash(doc, ops);
                metadata["operationStateVersion"] = 2;
            }
        });
    }

    /// <summary>
    /// 저장 완료된 DWG는 전체 파일 SHA-256으로, 이미 dirty였던 열린 DWG는 작업 대상
    /// 핸들/레이어의 메모리 상태 SHA-256으로 preview/snapshot 재사용을 검증한다.
    /// dirty 도면을 무조건 거부하면 정상적인 dry-run -> apply가 영구히 불가능해지므로,
    /// 작업 범위 상태가 그대로인지 확인하는 보수적 폴백을 사용한다.
    /// </summary>
    public JsonObject ValidatePreviewReuse(
        string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject> ops)
    {
        return ComInvokeWithRetry(() =>
        {
            var app = AttachCad();
            if (app is null)
                return new JsonObject { ["ok"] = true, ["reusable"] = false, ["reason"] = "AutoCAD is not running" };
            dynamic cad = app;
            var doc = ActiveDocWait(cad);
            if (doc is null)
                return new JsonObject { ["ok"] = true, ["reusable"] = false, ["reason"] = "AutoCAD drawing is not open" };

            string fullName = "";
            try { fullName = (string)(doc.FullName ?? ""); } catch { }
            var expectedDocument = Json.GetString(metadata, "documentRef") ?? "";
            if (string.IsNullOrWhiteSpace(fullName) || string.IsNullOrWhiteSpace(expectedDocument) ||
                !string.Equals(Path.GetFullPath(fullName), Path.GetFullPath(expectedDocument), StringComparison.OrdinalIgnoreCase))
                return new JsonObject
                {
                    ["ok"] = true,
                    ["reusable"] = false,
                    ["fingerprintMethod"] = "cad-saved-file-sha256",
                    ["reason"] = "active AutoCAD document identity changed after dry-run",
                };

            var saved = false;
            try { saved = (bool)doc.Saved; } catch { }
            var savedAtSnapshot = Json.GetBool(metadata, "savedAtSnapshot");
            if (!savedAtSnapshot)
            {
                var expectedState = Json.GetString(metadata, "operationStateSha256");
                if (Json.GetInt(metadata, "operationStateVersion") != 2 || string.IsNullOrWhiteSpace(expectedState))
                    return new JsonObject
                    {
                        ["ok"] = true,
                        ["reusable"] = false,
                        ["fingerprintMethod"] = "cad-operation-state-sha256",
                        ["reason"] = "dirty AutoCAD operation cannot be safely fingerprinted; save the target drawing and run a new dry-run",
                    };

                var currentState = CadOperationStateHash(doc, ops);
                var stateReusable = string.Equals(expectedState, currentState, StringComparison.OrdinalIgnoreCase);
                return new JsonObject
                {
                    ["ok"] = true,
                    ["reusable"] = stateReusable,
                    ["fingerprintMethod"] = "cad-operation-state-sha256",
                    ["reason"] = stateReusable
                        ? "dirty DWG operation-state fingerprint matched"
                        : "dirty DWG operation target changed after dry-run",
                };
            }

            if (!saved)
                return new JsonObject
                {
                    ["ok"] = true,
                    ["reusable"] = false,
                    ["fingerprintMethod"] = "cad-saved-file-sha256",
                    ["reason"] = "saved DWG acquired unsaved changes after dry-run",
                };

            var expected = Json.GetString(metadata, "fileSha256");
            if (string.IsNullOrWhiteSpace(expected) || !File.Exists(fullName))
                return new JsonObject
                {
                    ["ok"] = true,
                    ["reusable"] = false,
                    ["fingerprintMethod"] = "cad-saved-file-sha256",
                    ["reason"] = "saved DWG fingerprint is unavailable",
                };

            var current = CadFileHash(fullName);
            var reusable = string.Equals(expected, current, StringComparison.OrdinalIgnoreCase);
            return new JsonObject
            {
                ["ok"] = true,
                ["reusable"] = reusable,
                ["fingerprintMethod"] = "cad-saved-file-sha256",
                ["reason"] = reusable ? "saved DWG fingerprint matched" : "saved DWG changed after dry-run",
            };
        });
    }

    private static string CadFileHash(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }

    private static string? CadOperationStateHash(dynamic doc, IReadOnlyList<JsonObject> ops)
    {
        var state = new StringBuilder(4096);
        string documentName = "";
        try { documentName = (string)(doc.FullName ?? doc.Name ?? ""); } catch { }
        state.Append("document=").Append(documentName).Append('\n');
        try { state.Append("modelSpaceCount=").Append((int)doc.ModelSpace.Count).Append('\n'); } catch { }

        var handles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var layers = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var op in ops)
        {
            var opName = Json.GetString(op, "op");
            // A handle-only digest cannot validate a region, external source, block,
            // layout or arbitrary geometry. Do not bless unsupported dirty previews.
            if (opName == "activate_document")
            {
                var selector = Json.GetString(op, "document");
                if (selector is null || !(DocumentMatches(documentName, selector) ||
                    DocumentMatches((string)doc.Name, selector))) return null;
            }
            else if (opName is not ("regen_document" or "set_layer_visibility" or "set_layer_color" or
                "move_entities" or "rotate_entities" or "scale_entities" or "set_text_value")) return null;
            state.Append("op=").Append(opName ?? "").Append('\n');
            var handle = Json.GetString(op, "handle");
            if (!string.IsNullOrWhiteSpace(handle)) handles.Add(handle);
            if (Json.GetArr(op, "handles") is { } handleArray)
                foreach (var node in handleArray)
                {
                    var value = node?.GetValue<string>();
                    if (!string.IsNullOrWhiteSpace(value)) handles.Add(value);
                }
            var layer = Json.GetString(op, "layer");
            if (!string.IsNullOrWhiteSpace(layer)) layers.Add(layer);
        }

        foreach (var layerName in layers.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
        {
            state.Append("layer=").Append(layerName).Append('|');
            try
            {
                dynamic layer = doc.Layers.Item(layerName);
                state.Append((bool)layer.LayerOn).Append('|')
                    .Append((bool)layer.Freeze).Append('|')
                    .Append((bool)layer.Lock).Append('|')
                    .Append((int)layer.Color);
            }
            catch { return null; }
            state.Append('\n');
        }

        foreach (var handle in handles.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
        {
            state.Append("entity=").Append(handle).Append('|');
            try
            {
                dynamic entity = doc.HandleToObject(handle);
                // Text-only fallback. Curves can change vertices without changing bbox.
                if ((string)entity.EntityName is not ("AcDbText" or "AcDbMText")) return null;
                var geometry = EntityJson(entity, -1, includeGeometry: true);
                if (geometry["bounds"] is null || geometry["height"] is null || geometry["rotation"] is null ||
                    geometry["insertionPoint"] is null) return null;
                state.Append(geometry.ToJsonString()).Append('|');
                try { state.Append((string)entity.EntityName); } catch { }
                state.Append('|');
                try { state.Append((string)entity.Layer); } catch { }
                state.Append('|').Append(TextOf(entity)).Append('|');
                try { state.Append(Convert.ToDouble(entity.Height, CultureInfo.InvariantCulture).ToString("R", CultureInfo.InvariantCulture)); } catch { }
                state.Append('|');
                try { state.Append(Convert.ToDouble(entity.Rotation, CultureInfo.InvariantCulture).ToString("R", CultureInfo.InvariantCulture)); } catch { }
                state.Append('|');
                try { state.Append(PointJson((object?)entity.InsertionPoint).ToJsonString()); } catch { }
                state.Append('|');
                if (TryBoundingBox((object)entity, out var minX, out var minY, out var maxX, out var maxY))
                    state.Append(minX.ToString("R", CultureInfo.InvariantCulture)).Append(',')
                        .Append(minY.ToString("R", CultureInfo.InvariantCulture)).Append(',')
                        .Append(maxX.ToString("R", CultureInfo.InvariantCulture)).Append(',')
                        .Append(maxY.ToString("R", CultureInfo.InvariantCulture));
            }
            catch { return null; }
            state.Append('\n');
        }

        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(state.ToString()))).ToLowerInvariant();
    }

    public override JsonObject RestoreSnapshot(string snapshotDir, JsonObject metadata)
    {
        return ComInvokeWithRetry(() =>
        {
            var statePath = Path.Combine(snapshotDir, "state.json");
            if (!File.Exists(statePath)) return Json.ErrorResult("state.json not found in snapshot", App);

            var app = AttachCad();
            if (app is null) return Json.ErrorResult("AutoCAD not running", App);
            dynamic d = app;
            var doc = ActiveDocWait(d);
            if (doc is null) return Json.ErrorResult("열린 도면이 없습니다", App);

            var state = JsonNode.Parse(File.ReadAllText(statePath)) as JsonObject ?? new JsonObject();
            var docRef = Json.GetString(state, "fullName");
            if (!string.IsNullOrEmpty(docRef))
            {
                string cur = "";
                try { cur = (string)(doc.FullName ?? ""); } catch { }
                if (!string.IsNullOrEmpty(cur) && !string.Equals(cur, docRef, StringComparison.OrdinalIgnoreCase))
                    return Json.ErrorResult($"현재 도면 '{cur}'가 스냅샷 도면 '{docRef}'와 다릅니다.", App);
            }

            var mismatches = new List<string>();
            var checkedCount = 0;

            // 레이어 상태 복원
            if (Json.GetObj(state, "layers") is { } layers)
                foreach (var (layerName, lNode) in layers)
                {
                    if (lNode is not JsonObject lo) continue;
                    try
                    {
                        dynamic layer = doc.Layers.Item(layerName);
                        var wantOn = Json.GetBool(lo, "on");
                        var wantColor = Json.GetInt(lo, "color") ?? 7;
                        layer.LayerOn = wantOn;
                        layer.Color = wantColor;
                        checkedCount++;
                        if ((bool)layer.LayerOn != wantOn || (int)layer.Color != wantColor)
                            mismatches.Add($"layer {layerName}: restore readback mismatch");
                    }
                    catch (Exception ex) { mismatches.Add($"layer {layerName}: {ex.Message}"); }
                }

            // 텍스트 값 복원
            if (Json.GetObj(state, "texts") is { } texts)
                foreach (var (handle, tNode) in texts)
                {
                    try
                    {
                        dynamic ent = doc.HandleToObject(handle);
                        var want = tNode!.GetValue<string>();
                        ent.TextString = want;
                        checkedCount++;
                        if (TextOf(ent) != want)
                            mismatches.Add($"entity {handle}: restore readback mismatch");
                    }
                    catch (Exception ex) { mismatches.Add($"entity {handle}: {ex.Message}"); }
                }

            return new JsonObject
            {
                ["ok"] = mismatches.Count == 0,
                ["restored"] = true,
                ["readback"] = new JsonObject
                {
                    ["verified"] = mismatches.Count == 0,
                    ["checked"] = checkedCount,
                    ["mismatches"] = Json.ToArray(mismatches),
                },
                ["warnings"] = Json.ToArray(new[]
                {
                    "레이어 상태/텍스트 값만 복원됩니다. 이동/회전/삭제된 엔티티는 drawing-backup 파일로만 보존됩니다.",
                }),
                ["errors"] = Json.ToArray(mismatches),
            };
        });
    }
}
