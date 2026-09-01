using System.Globalization;
using System.Reflection;
using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

/// <summary>AutoCAD 기본/전문 제작, 객체 수정, 블록 속성, 배치·출력 기능.</summary>
public sealed partial class CadAdapter
{
    private sealed class RegionAccumulator
    {
        public required string Name { get; init; }
        public required JsonObject Spec { get; init; }
        public required double MinX { get; init; }
        public required double MinY { get; init; }
        public required double MaxX { get; init; }
        public required double MaxY { get; init; }
        public required string BoundsMode { get; init; }
        public HashSet<string>? Types { get; init; }
        public IReadOnlyList<string>? RequestedTypes { get; init; }
        public string? Layer { get; init; }
        public string? TextContains { get; init; }
        public int Count { get; set; }
        public Dictionary<string, int> CountsByType { get; } = new(StringComparer.OrdinalIgnoreCase);
        public JsonArray Samples { get; } = new();
        public double? ActualMinX { get; set; }
        public double? ActualMinY { get; set; }
        public double? ActualMaxX { get; set; }
        public double? ActualMaxY { get; set; }
    }

    private static JsonArray RequiredPoint(JsonObject owner, string name, int min = 2)
    {
        var point = Json.GetArr(owner, name)
            ?? throw new ArgumentException($"{name} point is required");
        if (point.Count < min) throw new ArgumentException($"{name} must contain at least {min} coordinates");
        return point;
    }

    private static double[] PointFrom(JsonArray point)
        => Point(Dbl(point[0]), Dbl(point[1]), point.Count > 2 ? Dbl(point[2]) : 0);

    private static void ApplyEntityProperties(
        dynamic app, dynamic doc, dynamic entity, JsonObject source, List<string> warnings)
    {
        var layer = Json.GetString(source, "layer");
        if (!string.IsNullOrWhiteSpace(layer))
        {
            EnsureLayer(doc, layer);
            entity.Layer = layer;
        }
        SetEntityColor(app, entity, Json.GetObj(source, "color"), warnings);
        var linetype = Json.GetString(source, "linetype");
        if (!string.IsNullOrWhiteSpace(linetype))
        {
            try { entity.Linetype = linetype; }
            catch (Exception ex) { warnings.Add($"linetype '{linetype}' not applied: {ex.Message}"); }
        }
        if (source["linetypeScale"] is not null) entity.LinetypeScale = Dbl(source["linetypeScale"]);
        if (source["lineweight"] is not null) entity.Lineweight = (int)Dbl(source["lineweight"]);
        if (source["visible"] is not null) entity.Visible = Json.GetBool(source, "visible", true);
    }

    private static object AddProductionEntity(
        dynamic app, dynamic doc, JsonObject entity, List<string> warnings)
    {
        var type = (Json.GetString(entity, "type") ?? "").ToLowerInvariant();
        dynamic created;
        switch (type)
        {
            case "line":
                created = doc.ModelSpace.AddLine(
                    PointFrom(RequiredPoint(entity, "start")),
                    PointFrom(RequiredPoint(entity, "end")));
                break;
            case "arc":
            {
                var radius = Dbl(entity["radius"]);
                if (radius <= 0) throw new ArgumentException("arc.radius must be positive");
                created = doc.ModelSpace.AddArc(
                    PointFrom(RequiredPoint(entity, "center")), radius,
                    Dbl(entity["startAngleDeg"]) * Math.PI / 180.0,
                    Dbl(entity["endAngleDeg"]) * Math.PI / 180.0);
                break;
            }
            case "ellipse":
            {
                var ratio = Dbl(entity["radiusRatio"]);
                if (ratio <= 0 || ratio > 1) throw new ArgumentException("ellipse.radiusRatio must be >0 and <=1");
                created = doc.ModelSpace.AddEllipse(
                    PointFrom(RequiredPoint(entity, "center")),
                    PointFrom(RequiredPoint(entity, "majorAxis")), ratio);
                break;
            }
            case "point":
                created = doc.ModelSpace.AddPoint(PointFrom(RequiredPoint(entity, "point")));
                break;
            case "mtext":
            {
                var width = Dbl(entity["width"]);
                if (width <= 0) throw new ArgumentException("mtext.width must be positive");
                created = doc.ModelSpace.AddMText(
                    PointFrom(RequiredPoint(entity, "point")), width,
                    Json.GetString(entity, "text") ?? "");
                if (entity["height"] is not null) created.Height = Dbl(entity["height"]);
                if (entity["rotationDeg"] is not null) created.Rotation = Dbl(entity["rotationDeg"]) * Math.PI / 180.0;
                if (entity["attachmentPoint"] is not null) created.AttachmentPoint = (int)Dbl(entity["attachmentPoint"]);
                break;
            }
            case "dim_aligned":
                created = doc.ModelSpace.AddDimAligned(
                    PointFrom(RequiredPoint(entity, "start")),
                    PointFrom(RequiredPoint(entity, "end")),
                    PointFrom(RequiredPoint(entity, "textPoint")));
                break;
            case "dim_rotated":
                created = doc.ModelSpace.AddDimRotated(
                    PointFrom(RequiredPoint(entity, "start")),
                    PointFrom(RequiredPoint(entity, "end")),
                    PointFrom(RequiredPoint(entity, "dimensionLinePoint")),
                    Dbl(entity["rotationDeg"]) * Math.PI / 180.0);
                break;
            default:
                throw new ArgumentException($"unsupported production entity type: {type}");
        }
        ApplyEntityProperties(app, doc, created, entity, warnings);
        return (object)created;
    }

    private static List<dynamic> ResolveHandles(dynamic doc, JsonObject op)
    {
        var handles = Json.GetArr(op, "handles")
            ?? throw new ArgumentException("handles array is required");
        if (handles.Count == 0) throw new ArgumentException("handles array is empty");
        if (handles.Count > MaxQueryEntities) throw new ArgumentException($"handles exceeds {MaxQueryEntities}");
        var entities = new List<dynamic>(handles.Count);
        foreach (var node in handles)
        {
            var handle = node?.GetValue<string>() ?? "";
            if (string.IsNullOrWhiteSpace(handle)) throw new ArgumentException("handle is empty");
            entities.Add(doc.HandleToObject(handle));
        }
        return entities;
    }

    private static List<object> CopyEntitiesByVector(dynamic doc, JsonObject op)
    {
        var dx = Dbl(op["dx"]);
        var dy = Dbl(op["dy"]);
        var copied = new List<object>();
        try
        {
            foreach (dynamic entity in ResolveHandles(doc, op))
            {
                dynamic clone = entity.Copy();
                clone.Move(Point(0, 0), Point(dx, dy));
                copied.Add((object)clone);
            }
            return copied;
        }
        catch
        {
            for (var i = copied.Count - 1; i >= 0; i--) try { ((dynamic)copied[i]).Delete(); } catch { }
            throw;
        }
    }

    private static int ScaleEntities(dynamic doc, JsonObject op)
    {
        var factor = Dbl(op["factor"]);
        if (factor <= 0) throw new ArgumentException("factor must be positive");
        var basePoint = PointFrom(RequiredPoint(op, "basePoint"));
        var count = 0;
        foreach (dynamic entity in ResolveHandles(doc, op)) { entity.ScaleEntity(basePoint, factor); count++; }
        return count;
    }

    private static List<object> MirrorEntities(dynamic doc, JsonObject op)
    {
        var p1 = PointFrom(RequiredPoint(op, "axisStart"));
        var p2 = PointFrom(RequiredPoint(op, "axisEnd"));
        var mirrored = new List<object>();
        foreach (dynamic entity in ResolveHandles(doc, op)) mirrored.Add((object)entity.Mirror(p1, p2));
        return mirrored;
    }

    private static List<object> OffsetEntities(dynamic doc, JsonObject op)
    {
        var distance = Dbl(op["distance"]);
        if (Math.Abs(distance) < 1e-12) throw new ArgumentException("distance must be non-zero");
        var created = new List<object>();
        foreach (dynamic entity in ResolveHandles(doc, op)) created.AddRange(ComObjects(entity.Offset(distance)));
        return created;
    }

    private static int SetEntityProperties(
        dynamic app, dynamic doc, JsonObject op, List<string> warnings)
    {
        var properties = Json.GetObj(op, "properties")
            ?? throw new ArgumentException("properties object is required");
        var count = 0;
        foreach (dynamic entity in ResolveHandles(doc, op))
        {
            ApplyEntityProperties(app, doc, entity, properties, warnings);
            count++;
        }
        return count;
    }

    private static int SetBlockAttributes(dynamic doc, JsonObject op, List<string> warnings)
    {
        var handle = Json.GetString(op, "handle") ?? throw new ArgumentException("handle is required");
        var values = Json.GetObj(op, "attributes") ?? throw new ArgumentException("attributes object is required");
        dynamic block = doc.HandleToObject(handle);
        if (!(bool)block.HasAttributes) throw new InvalidOperationException($"block {handle} has no attributes");
        var requested = values.ToDictionary(kv => kv.Key, kv => kv.Value?.GetValue<string>() ?? "", StringComparer.OrdinalIgnoreCase);
        var changed = 0;
        foreach (dynamic attribute in ComObjects(block.GetAttributes()))
        {
            var tag = (string)attribute.TagString;
            if (!requested.TryGetValue(tag, out var value)) continue;
            attribute.TextString = value;
            changed++;
            requested.Remove(tag);
        }
        foreach (var missing in requested.Keys) warnings.Add($"block attribute tag not found: {missing}");
        if (changed == 0) throw new InvalidOperationException("no requested block attributes were found");
        return changed;
    }

    private static dynamic FindOrCreateLayout(dynamic doc, string name, bool create)
    {
        try { return doc.Layouts.Item(name); }
        catch when (create) { return doc.Layouts.Add(name); }
    }

    private static dynamic ConfigureLayout(dynamic doc, JsonObject op, List<string> warnings)
    {
        var name = Json.GetString(op, "name") ?? throw new ArgumentException("layout name is required");
        if (name.Equals("Model", StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("Model layout cannot be configured here");
        dynamic layout = FindOrCreateLayout(doc, name, Json.GetBool(op, "create", true));
        doc.ActiveLayout = layout;
        try { layout.RefreshPlotDeviceInfo(); } catch { }
        void TrySet(Action setter, string label)
        {
            try { setter(); } catch (Exception ex) { warnings.Add($"layout {label} not applied: {ex.Message}"); }
        }
        var config = Json.GetString(op, "configName");
        if (!string.IsNullOrWhiteSpace(config)) TrySet(() => layout.ConfigName = config, "configName");
        var media = Json.GetString(op, "canonicalMediaName");
        if (!string.IsNullOrWhiteSpace(media)) TrySet(() => layout.CanonicalMediaName = media, "canonicalMediaName");
        if (op["plotRotation"] is not null) TrySet(() => layout.PlotRotation = (int)Dbl(op["plotRotation"]), "plotRotation");
        if (op["centerPlot"] is not null) TrySet(() => layout.CenterPlot = Json.GetBool(op, "centerPlot"), "centerPlot");
        if (op["plotType"] is not null) TrySet(() => layout.PlotType = (int)Dbl(op["plotType"]), "plotType");
        if (op["standardScale"] is not null) TrySet(() => layout.StandardScale = (int)Dbl(op["standardScale"]), "standardScale");
        if (op["useStandardScale"] is not null) TrySet(() => layout.UseStandardScale = Json.GetBool(op, "useStandardScale"), "useStandardScale");
        return layout;
    }

    private static object CreateViewport(dynamic doc, JsonObject op)
    {
        var layoutName = Json.GetString(op, "layout") ?? throw new ArgumentException("layout is required");
        dynamic layout = FindOrCreateLayout(doc, layoutName, create: false);
        doc.ActiveLayout = layout;
        doc.MSpace = false;
        var center = PointFrom(RequiredPoint(op, "center"));
        var width = Dbl(op["width"]);
        var height = Dbl(op["height"]);
        var viewHeight = Dbl(op["viewHeight"]);
        if (width <= 0 || height <= 0 || viewHeight <= 0) throw new ArgumentException("viewport width, height and viewHeight must be positive");
        dynamic viewport = doc.PaperSpace.AddPViewport(center, width, height);
        viewport.Display(true);
        if (Json.GetArr(op, "viewCenter") is { Count: >= 2 } viewCenter)
            viewport.Target = Point(Dbl(viewCenter[0]), Dbl(viewCenter[1]));
        // AcadPViewport에는 ViewHeight/ViewCenter가 없으므로 종이공간 높이와
        // 원하는 모델 뷰 높이의 비를 공식 CustomScale 속성에 적용한다.
        viewport.CustomScale = height / viewHeight;
        if (op["twistAngleDeg"] is not null) viewport.TwistAngle = Dbl(op["twistAngleDeg"]) * Math.PI / 180.0;
        viewport.DisplayLocked = Json.GetBool(op, "displayLocked", true);
        return (object)viewport;
    }

    private static string SaveDocument(dynamic doc, JsonObject op)
    {
        var output = Json.GetString(op, "output");
        if (string.IsNullOrWhiteSpace(output))
        {
            string currentPath = "";
            try { currentPath = (string)(doc.FullName ?? ""); } catch { }
            if (string.IsNullOrWhiteSpace(currentPath) || !Path.IsPathFullyQualified(currentPath))
                throw new InvalidOperationException("unsaved drawing requires save_document.output to avoid a Save As dialog");
            doc.Save();
            return currentPath;
        }
        var full = Path.GetFullPath(output);
        Directory.CreateDirectory(Path.GetDirectoryName(full)!);
        doc.SaveAs(full);
        if (!File.Exists(full)) throw new IOException($"AutoCAD SaveAs output was not created: {full}");
        return full;
    }

    private static string PlotPdf(dynamic doc, JsonObject op)
    {
        var output = Path.GetFullPath(Json.GetString(op, "output") ?? throw new ArgumentException("output is required"));
        if (!Path.GetExtension(output).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
            throw new ArgumentException("plot_pdf output must end with .pdf");
        Directory.CreateDirectory(Path.GetDirectoryName(output)!);
        var temporary = Path.Combine(
            Path.GetDirectoryName(output)!,
            $".{Path.GetFileNameWithoutExtension(output)}.docbridge-{Guid.NewGuid():N}.tmp.pdf");
        try
        {
            doc.SetVariable("BACKGROUNDPLOT", 0);
            dynamic plot = doc.Plot;
            try { plot.QuietErrorMode = true; } catch { }
            var config = Json.GetString(op, "configName");
            bool ok = string.IsNullOrWhiteSpace(config)
                ? (bool)plot.PlotToFile(temporary)
                : (bool)plot.PlotToFile(temporary, config);
            if (!ok || !File.Exists(temporary) || new FileInfo(temporary).Length == 0)
                throw new IOException($"AutoCAD plot did not create a non-empty temporary PDF: {temporary}");
            File.Move(temporary, output, overwrite: true);
            return output;
        }
        finally
        {
            if (File.Exists(temporary)) try { File.Delete(temporary); } catch { }
        }
    }

    private static JsonObject InspectLayouts(dynamic doc)
    {
        var layouts = new JsonArray();
        foreach (dynamic layout in doc.Layouts)
        {
            var item = new JsonObject();
            try { item["name"] = (string)layout.Name; } catch { }
            try { item["modelType"] = (bool)layout.ModelType; } catch { }
            try { item["configName"] = (string)layout.ConfigName; } catch { }
            try { item["canonicalMediaName"] = (string)layout.CanonicalMediaName; } catch { }
            try { item["plotRotation"] = (int)layout.PlotRotation; } catch { }
            var viewports = new JsonArray();
            try
            {
                dynamic block = layout.Block;
                foreach (dynamic entity in block)
                {
                    string type = "";
                    try { type = (string)entity.EntityName; } catch { }
                    if (!type.Equals("AcDbViewport", StringComparison.OrdinalIgnoreCase)) continue;
                    viewports.Add(new JsonObject
                    {
                        ["handle"] = (string)entity.Handle,
                        ["center"] = PointJson((object?)entity.Center),
                        ["width"] = Convert.ToDouble(entity.Width, CultureInfo.InvariantCulture),
                        ["height"] = Convert.ToDouble(entity.Height, CultureInfo.InvariantCulture),
                        ["target"] = PointJson((object?)entity.Target),
                        ["customScale"] = Convert.ToDouble(entity.CustomScale, CultureInfo.InvariantCulture),
                    });
                }
            }
            catch { }
            item["viewports"] = viewports;
            layouts.Add(item);
        }
        string activeName = "";
        try { activeName = (string)doc.ActiveLayout.Name; } catch { }
        return new JsonObject
        {
            ["ok"] = true,
            ["app"] = "cad",
            ["scope"] = "layouts",
            ["activeLayout"] = activeName,
            ["layouts"] = layouts,
            ["count"] = layouts.Count,
        };
    }

    /// <summary>
    /// 여러 도곽/평면/종단/제목/키맵 영역을 ModelSpace 한 번의 순회로 집계한다.
    /// 반복 cad_query 호출을 줄이고 각 영역의 객체수·유형·실제 bbox를 함께 검증한다.
    /// </summary>
    private static JsonObject InspectRegions(dynamic doc, JsonArray specs)
    {
        if (specs.Count is < 1 or > 100) throw new ArgumentException("regions count must be 1..100");
        var regions = new List<RegionAccumulator>(specs.Count);
        var index = 0;
        foreach (var node in specs)
        {
            index++;
            if (node is not JsonObject spec) throw new ArgumentException($"regions[{index}] must be an object");
            var bounds = Json.GetObj(spec, "bounds") ?? throw new ArgumentException($"regions[{index}].bounds is required");
            var minX = Dbl(bounds["minX"]); var minY = Dbl(bounds["minY"]);
            var maxX = Dbl(bounds["maxX"]); var maxY = Dbl(bounds["maxY"]);
            if (maxX <= minX || maxY <= minY) throw new ArgumentException($"regions[{index}] has invalid bounds");
            HashSet<string>? types = null;
            IReadOnlyList<string>? requestedTypes = null;
            if (Json.GetArr(spec, "entityTypes") is { Count: > 0 } typeArray)
            {
                requestedTypes = typeArray.Select(v => v?.GetValue<string>()?.Trim() ?? "")
                    .Where(v => v.Length > 0)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToArray();
                types = requestedTypes
                    .Where(v => v.Length > 0)
                    .SelectMany(v => new[] { v, v.StartsWith("AcDb", StringComparison.OrdinalIgnoreCase) ? v[4..] : "AcDb" + v })
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);
            }
            var boundsMode = (Json.GetString(spec, "boundsMode") ?? "center").ToLowerInvariant();
            if (boundsMode is not ("center" or "inside" or "intersect"))
                throw new ArgumentException($"regions[{index}].boundsMode must be center, inside, or intersect");
            regions.Add(new RegionAccumulator
            {
                Name = Json.GetString(spec, "name") ?? $"region-{index}", Spec = spec,
                MinX = minX, MinY = minY, MaxX = maxX, MaxY = maxY, Types = types,
                BoundsMode = boundsMode, RequestedTypes = requestedTypes,
                Layer = Json.GetString(spec, "layer"), TextContains = Json.GetString(spec, "textContains"),
            });
        }

        var scanned = 0;
        foreach (dynamic entity in doc.ModelSpace)
        {
            scanned++;
            if (!TryBoundingBox((object)entity, out var ex0, out var ey0, out var ex1, out var ey1)) continue;
            string type = ""; string layer = ""; string handle = "";
            try { type = (string)entity.EntityName; } catch { }
            try { layer = (string)entity.Layer; } catch { }
            try { handle = (string)entity.Handle; } catch { }
            var centerX = (ex0 + ex1) / 2; var centerY = (ey0 + ey1) / 2;
            foreach (var region in regions)
            {
                var spatial = region.BoundsMode switch
                {
                    "inside" => ex0 >= region.MinX && ey0 >= region.MinY && ex1 <= region.MaxX && ey1 <= region.MaxY,
                    "intersect" => ex1 >= region.MinX && ex0 <= region.MaxX && ey1 >= region.MinY && ey0 <= region.MaxY,
                    _ => centerX >= region.MinX && centerX <= region.MaxX && centerY >= region.MinY && centerY <= region.MaxY,
                };
                if (!spatial) continue;
                if (region.Layer is not null && !layer.Equals(region.Layer, StringComparison.OrdinalIgnoreCase)) continue;
                if (region.Types is not null && !region.Types.Contains(type)) continue;
                var text = region.TextContains is null ? "" : TextOf(entity);
                if (region.TextContains is not null && !text.Contains(region.TextContains, StringComparison.OrdinalIgnoreCase)) continue;
                region.Count++;
                region.CountsByType[type] = region.CountsByType.GetValueOrDefault(type) + 1;
                region.ActualMinX = region.ActualMinX is null ? ex0 : Math.Min(region.ActualMinX.Value, ex0);
                region.ActualMinY = region.ActualMinY is null ? ey0 : Math.Min(region.ActualMinY.Value, ey0);
                region.ActualMaxX = region.ActualMaxX is null ? ex1 : Math.Max(region.ActualMaxX.Value, ex1);
                region.ActualMaxY = region.ActualMaxY is null ? ey1 : Math.Max(region.ActualMaxY.Value, ey1);
                if (region.Samples.Count < 20)
                {
                    var sample = new JsonObject { ["handle"] = handle, ["type"] = type, ["layer"] = layer };
                    if (IsTextLike(type)) sample["text"] = TextOf(entity);
                    region.Samples.Add(sample);
                }
            }
        }

        var output = new JsonArray();
        var errors = new JsonArray();
        foreach (var region in regions)
        {
            var counts = new JsonObject();
            foreach (var (type, count) in region.CountsByType) counts[type] = count;
            var minCount = Json.GetInt(region.Spec, "minCount") ?? 0;
            var maxCount = Json.GetInt(region.Spec, "maxCount");
            var verified = region.Count >= minCount && (maxCount is null || region.Count <= maxCount.Value);
            if (!verified) errors.Add($"{region.Name}: count {region.Count} outside expected {minCount}..{maxCount?.ToString() ?? "unbounded"}");
            var item = new JsonObject
            {
                ["name"] = region.Name, ["count"] = region.Count, ["verified"] = verified,
                ["countsByType"] = counts, ["samples"] = region.Samples,
                ["sampleCoverage"] = new JsonObject
                {
                    ["totalMatched"] = region.Count,
                    ["returned"] = region.Samples.Count,
                    ["truncated"] = region.Count > region.Samples.Count,
                    ["complete"] = region.Count == region.Samples.Count,
                },
            };
            if (region.ActualMinX is not null)
                item["actualBounds"] = new JsonObject
                {
                    ["minX"] = region.ActualMinX, ["minY"] = region.ActualMinY,
                    ["maxX"] = region.ActualMaxX, ["maxY"] = region.ActualMaxY,
                };
            if (region.Count > region.Samples.Count)
            {
                string documentSelector = "";
                try { documentSelector = (string)(doc.FullName ?? ""); } catch { }
                if (string.IsNullOrWhiteSpace(documentSelector))
                    try { documentSelector = (string)(doc.Name ?? ""); } catch { }
                var baseArguments = new JsonObject
                {
                    ["scope"] = "window",
                    ["document"] = documentSelector,
                    ["bounds"] = new JsonObject
                    {
                        ["minX"] = region.MinX, ["minY"] = region.MinY,
                        ["maxX"] = region.MaxX, ["maxY"] = region.MaxY,
                    },
                    ["boundsMode"] = region.BoundsMode,
                    ["limit"] = 500,
                };
                if (region.Layer is not null) baseArguments["layer"] = region.Layer;
                if (region.TextContains is not null) baseArguments["textContains"] = region.TextContains;
                var actions = new JsonArray();
                IEnumerable<string?> actionTypes = region.RequestedTypes is { Count: > 0 }
                    ? region.RequestedTypes.Select(value => (string?)value)
                    : new string?[] { null };
                foreach (var requestedType in actionTypes)
                {
                    var arguments = (JsonObject)baseArguments.DeepClone();
                    if (requestedType is not null) arguments["entityType"] = requestedType;
                    actions.Add(CadQueryAction(
                        requestedType is null
                            ? "이 영역의 표본 20개를 초과한 실제 엔티티를 같은 공간 필터로 조회"
                            : $"이 영역의 {requestedType} 엔티티를 같은 공간 필터로 조회",
                        arguments));
                }
                item["nextActions"] = actions;
            }
            output.Add(item);
        }
        return new JsonObject
        {
            ["ok"] = errors.Count == 0, ["app"] = "cad", ["scope"] = "regions",
            ["verified"] = errors.Count == 0, ["scanned"] = scanned,
            ["regions"] = output, ["errors"] = errors, ["warnings"] = new JsonArray(),
        };
    }

    private static JsonObject InspectLayers(dynamic doc, JsonObject args)
    {
        var contains = Json.GetString(args, "contains");
        var startsWith = Json.GetString(args, "startsWith");
        var startIndex = Math.Max(0, Json.GetInt(args, "startIndex") ?? 0);
        var limit = Math.Clamp(Json.GetInt(args, "limit") ?? 500, 1, 5000);
        var layers = new JsonArray();
        var matched = 0;
        var scanned = 0;
        var total = (int)doc.Layers.Count;
        for (var index = startIndex; index < total; index++)
        {
            dynamic layer = doc.Layers.Item(index);
            scanned++;
            var name = (string)(layer.Name ?? "");
            if (contains is not null && !name.Contains(contains, StringComparison.OrdinalIgnoreCase)) continue;
            if (startsWith is not null && !name.StartsWith(startsWith, StringComparison.OrdinalIgnoreCase)) continue;
            var item = new JsonObject { ["index"] = index, ["name"] = name };
            try { item["on"] = (bool)layer.LayerOn; } catch { }
            try { item["freeze"] = (bool)layer.Freeze; } catch { }
            try { item["locked"] = (bool)layer.Lock; } catch { }
            try { item["plottable"] = (bool)layer.Plottable; } catch { }
            try { item["color"] = (int)layer.Color; } catch { }
            try { item["linetype"] = (string)layer.Linetype; } catch { }
            layers.Add(item);
            matched++;
            if (matched >= limit) break;
        }
        var truncated = startIndex + scanned < total;
        var nextActions = new JsonArray();
        if (truncated)
        {
            var continuation = QueryArguments(args, "layers");
            continuation["startIndex"] = startIndex + scanned;
            continuation["limit"] = limit;
            nextActions.Add(CadQueryAction("남은 레이어 목록을 계속 조회", continuation));
        }
        return new JsonObject
        {
            ["ok"] = true, ["app"] = "cad", ["scope"] = "layers", ["layers"] = layers,
            ["count"] = matched, ["scanned"] = scanned, ["totalLayers"] = total,
            ["truncated"] = truncated,
            ["nextStartIndex"] = truncated ? startIndex + scanned : null,
            ["coverage"] = new JsonObject
            {
                ["total"] = total, ["scanned"] = scanned, ["returned"] = layers.Count,
                ["complete"] = !truncated,
            },
            ["nextActions"] = nextActions,
        };
    }

    private static JsonObject InspectXrefs(dynamic doc, JsonObject args)
    {
        var blockName = Json.GetString(args, "blockName");
        var startIndex = Math.Max(0, Json.GetInt(args, "startIndex") ?? 0);
        var endIndex = Math.Min((int)doc.ModelSpace.Count - 1,
            Json.GetInt(args, "endIndex") ?? ((int)doc.ModelSpace.Count - 1));
        var limit = Math.Clamp(Json.GetInt(args, "limit") ?? 100, 1, 1000);
        var xrefs = new JsonArray();
        var scanned = 0;
        for (var index = startIndex; index <= endIndex; index++)
        {
            dynamic entity = doc.ModelSpace.Item(index);
            scanned++;
            string type;
            try { type = (string)entity.EntityName; } catch { continue; }
            if (!type.Equals("AcDbBlockReference", StringComparison.OrdinalIgnoreCase)) continue;
            string name;
            try { name = (string)entity.EffectiveName; }
            catch { try { name = (string)entity.Name; } catch { continue; } }
            if (blockName is not null && !DocumentMatches(name, blockName)) continue;
            dynamic? definition = null;
            try { definition = doc.Blocks.Item(name); } catch { }
            var isXref = false;
            try { isXref = definition is not null && (bool)definition.IsXRef; } catch { }
            if (!isXref) continue;
            var item = EntityJson(entity, index, includeGeometry: true);
            item["isXref"] = true;
            if (definition is not null)
            {
                try { item["definitionPath"] = (string)definition.Path; } catch { }
                try { item["definitionEntityCount"] = (int)definition.Count; } catch { }
            }
            try
            {
                if ((bool)entity.HasExtensionDictionary)
                {
                    dynamic dictionary = entity.GetExtensionDictionary();
                    var extensionObjects = new JsonArray();
                    for (var dictionaryIndex = 0; dictionaryIndex < (int)dictionary.Count; dictionaryIndex++)
                    {
                        dynamic extensionObject = dictionary.Item(dictionaryIndex);
                        var extensionItem = new JsonObject();
                        try { extensionItem["name"] = (string)extensionObject.Name; } catch { }
                        try { extensionItem["objectName"] = (string)extensionObject.ObjectName; } catch { }
                        try
                        {
                            var children = new JsonArray();
                            for (var childIndex = 0; childIndex < (int)extensionObject.Count; childIndex++)
                            {
                                dynamic child = extensionObject.Item(childIndex);
                                var childItem = new JsonObject();
                                try { childItem["name"] = (string)child.Name; } catch { }
                                try { childItem["objectName"] = (string)child.ObjectName; } catch { }
                                children.Add(childItem);
                            }
                            extensionItem["children"] = children;
                        }
                        catch { }
                        extensionObjects.Add(extensionItem);
                    }
                    item["extensionObjects"] = extensionObjects;
                }
            }
            catch { }
            var dependentPrefix = name + "|";
            var dependentCount = 0;
            var onCount = 0;
            var thawedCount = 0;
            foreach (dynamic layer in doc.Layers)
            {
                string layerName;
                try { layerName = (string)layer.Name; } catch { continue; }
                if (!layerName.StartsWith(dependentPrefix, StringComparison.OrdinalIgnoreCase)) continue;
                dependentCount++;
                try { if ((bool)layer.LayerOn) onCount++; } catch { }
                try { if (!(bool)layer.Freeze) thawedCount++; } catch { }
            }
            item["dependentLayerCount"] = dependentCount;
            item["dependentLayersOn"] = onCount;
            item["dependentLayersThawed"] = thawedCount;
            xrefs.Add(item);
            if (xrefs.Count >= limit) break;
        }
        var scanEndIndex = startIndex + scanned - 1;
        var truncated = scanEndIndex < endIndex;
        var nextActions = new JsonArray();
        if (truncated)
        {
            var continuation = QueryArguments(args, "xrefs");
            continuation["startIndex"] = scanEndIndex + 1;
            continuation["endIndex"] = endIndex;
            continuation["limit"] = limit;
            nextActions.Add(CadQueryAction("남은 ModelSpace 범위에서 XREF 조회 계속", continuation));
        }
        return new JsonObject
        {
            ["ok"] = true, ["app"] = "cad", ["scope"] = "xrefs", ["xrefs"] = xrefs,
            ["count"] = xrefs.Count, ["scanned"] = scanned, ["scanStartIndex"] = startIndex,
            ["scanEndIndex"] = scanEndIndex, ["modelSpaceCount"] = (int)doc.ModelSpace.Count,
            ["truncated"] = truncated,
            ["nextStartIndex"] = truncated ? scanEndIndex + 1 : null,
            ["nextActions"] = nextActions,
        };
    }

    private static JsonObject InspectWindowSelection(dynamic doc, JsonObject args)
    {
        var bounds = Json.GetObj(args, "bounds")
            ?? throw new ArgumentException("scope=window requires bounds");
        var minX = Dbl(bounds["minX"]); var minY = Dbl(bounds["minY"]);
        var maxX = Dbl(bounds["maxX"]); var maxY = Dbl(bounds["maxY"]);
        if (maxX <= minX || maxY <= minY) throw new ArgumentException("scope=window has invalid bounds");
        var modeName = (Json.GetString(args, "boundsMode") ?? "intersect").ToLowerInvariant();
        if (modeName is not ("center" or "inside" or "intersect"))
            throw new ArgumentException("scope=window boundsMode must be center, inside, or intersect");
        // ActiveX has native inside/crossing modes but no center mode. Center
        // starts from the crossing superset and is reduced with BoundsMatch below.
        var selectionMode = modeName == "inside" ? 0 : 1;
        var layerFilter = Json.GetString(args, "layer");
        var typeFilter = Json.GetString(args, "entityType");
        var textContains = Json.GetString(args, "textContains");
        var blockName = Json.GetString(args, "blockName");
        var includeGeometry = Json.GetBool(args, "includeGeometry");
        var countOnly = Json.GetBool(args, "countOnly");
        var limit = Math.Clamp(Json.GetInt(args, "limit") ?? 100, 1, MaxQueryEntities);
        dynamic? selection = null;
        dynamic? originalDocument = null;
        var restoreOriginalDocument = false;
        int? originalActiveSpace = null;
        var selectionName = "DOCBRIDGE_WINDOW_" + Guid.NewGuid().ToString("N");
        try
        {
            try
            {
                dynamic application = doc.Application;
                originalDocument = application.ActiveDocument;
                var originalName = (string)(originalDocument.Name ?? "");
                var queryName = (string)(doc.Name ?? "");
                if (!originalName.Equals(queryName, StringComparison.OrdinalIgnoreCase))
                {
                    doc.Activate();
                    restoreOriginalDocument = true;
                }
            }
            catch { }
            // ActiveX SelectionSet.Select searches the active space only.  Source
            // drawings are frequently left on a paper-space layout, while the
            // project geometry lives in ModelSpace.  Switch temporarily so a
            // spatial query is deterministic, then restore the user's space.
            try
            {
                originalActiveSpace = (int)doc.ActiveSpace;
                if (originalActiveSpace != 1) doc.ActiveSpace = 1; // acModelSpace
            }
            catch { }
            selection = doc.SelectionSets.Add(selectionName);
            // Omit optional filter arguments entirely. Passing Type.Missing to
            // an IDispatch call is accepted by some AutoCAD versions but is
            // treated as an empty filter (and selects nothing) by AutoCAD 2027.
            var selectArgs = new object?[]
            {
                selectionMode, Point(minX, minY), Point(maxX, maxY),
            };
            ((object)selection).GetType().InvokeMember(
                "Select", BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                null, (object)selection, selectArgs, CultureInfo.InvariantCulture);
            var selectedByLayerFallback = false;
            var selected = (int)selection.Count;
            if (selected == 0 && !string.IsNullOrWhiteSpace(layerFilter))
            {
                // AutoCAD 2027 can return an empty Window/Crossing set for a
                // document activated only for automation.  Let AutoCAD's own
                // layer index narrow the candidates, then apply the requested
                // geometric bounds below.  This avoids a full ModelSpace scan.
                try
                {
                    selection.Clear();
                    var filterTypes = new short[] { 8 }; // DXF group: layer
                    var filterData = new object[] { layerFilter };
                    var allOnLayerArgs = new object?[]
                    {
                        5, Type.Missing, Type.Missing, filterTypes, filterData, // acSelectionSetAll
                    };
                    ((object)selection).GetType().InvokeMember(
                        "Select", BindingFlags.InvokeMethod | BindingFlags.Public | BindingFlags.Instance,
                        null, (object)selection, allOnLayerArgs, CultureInfo.InvariantCulture);
                    selected = (int)selection.Count;
                    selectedByLayerFallback = true;
                }
                catch { }
            }
            var entities = new JsonArray();
            var count = 0;
            for (var selectedIndex = 0; selectedIndex < selected; selectedIndex++)
            {
                dynamic entity = selection.Item(selectedIndex);
                string type = ""; string layer = "";
                try { type = (string)entity.EntityName; } catch { }
                try { layer = (string)entity.Layer; } catch { }
                if (layerFilter is not null && !layer.Equals(layerFilter, StringComparison.OrdinalIgnoreCase)) continue;
                if (typeFilter is not null && !type.Equals(typeFilter, StringComparison.OrdinalIgnoreCase) &&
                    !type.Equals("AcDb" + typeFilter, StringComparison.OrdinalIgnoreCase)) continue;
                if (textContains is not null && !TextOf(entity).Contains(textContains, StringComparison.OrdinalIgnoreCase)) continue;
                if ((selectedByLayerFallback || modeName == "center") &&
                    !BoundsMatch((object)entity, minX, minY, maxX, maxY, modeName)) continue;
                if (blockName is not null)
                {
                    string actualName = "";
                    try { actualName = (string)entity.EffectiveName; }
                    catch { try { actualName = (string)entity.Name; } catch { } }
                    if (!DocumentMatches(actualName, blockName)) continue;
                }
                count++;
                if (!countOnly) entities.Add(EntityJson(entity, -1, includeGeometry));
                if (!countOnly && count >= limit) break;
            }
            var truncated = !countOnly && count >= limit && selected > count;
            var nextActions = new JsonArray();
            if (truncated)
            {
                var countArguments = QueryArguments(args, "window");
                countArguments["countOnly"] = true;
                nextActions.Add(CadQueryAction(
                    "같은 공간 필터의 전체 일치 개수를 엔티티 목록 없이 확인", countArguments));
                if (limit < MaxQueryEntities)
                {
                    var expandedArguments = QueryArguments(args, "window");
                    expandedArguments["limit"] = Math.Min(MaxQueryEntities, Math.Max(limit * 2, 500));
                    nextActions.Add(CadQueryAction(
                        "같은 작업 영역의 반환 한도를 늘려 다시 조회", expandedArguments));
                }
            }
            return new JsonObject
            {
                ["ok"] = true, ["app"] = "cad", ["scope"] = "window", ["entities"] = entities,
                ["count"] = count, ["nativeSelected"] = selected, ["selectionMode"] = modeName,
                ["layerFallback"] = selectedByLayerFallback,
                ["truncated"] = truncated,
                ["coverage"] = new JsonObject
                {
                    ["nativeSelected"] = selected,
                    ["returned"] = countOnly ? 0 : entities.Count,
                    ["matchedBeforeLimit"] = count,
                    ["complete"] = !truncated,
                },
                ["nextActions"] = nextActions,
            };
        }
        finally
        {
            if (selection is not null) try { selection.Delete(); } catch { }
            if (originalActiveSpace is not null)
                try { doc.ActiveSpace = originalActiveSpace.Value; } catch { }
            if (restoreOriginalDocument && originalDocument is not null)
                try { originalDocument.Activate(); } catch { }
        }
    }
}
