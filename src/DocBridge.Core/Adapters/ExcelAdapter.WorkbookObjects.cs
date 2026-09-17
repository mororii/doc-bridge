using System.Globalization;
using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

/// <summary>
/// Sparkline, slicer, and cell-style ops plus their inspect inventories.
/// </summary>
public sealed partial class ExcelAdapter
{
    // ------------------------------------------------------------------
    // Sparkline helpers
    // ------------------------------------------------------------------

    private static readonly IReadOnlyDictionary<int, string> SparklineTypeNames =
        new Dictionary<int, string> { [1] = "line", [2] = "column" };

    // Worksheet.SparklineGroups can report an empty collection while live groups
    // exist (observed on Excel 16.0); the sheet-wide Cells range surfaces them.
    private static IEnumerable<object> EnumerateSparklineGroups(object sheet)
    {
        object? cells = null;
        object? groups = null;
        try
        {
            try { cells = (object)((dynamic)sheet).Cells; }
            catch { yield break; }
            try { groups = (object)((dynamic)cells).SparklineGroups; }
            catch { yield break; }
            var count = Convert.ToInt32(((dynamic)groups).Count, CultureInfo.InvariantCulture);
            var items = new List<object>();
            for (var i = 1; i <= count; i++)
                items.Add((object)((dynamic)groups).Item(i));
            foreach (var item in items) yield return item;
        }
        finally
        {
            RotHelper.ReleaseComReference(groups);
            RotHelper.ReleaseComReference(cells);
        }
    }

    private static object? GetSparkPoint(object points, string property)
    {
        try
        {
            return property switch
            {
                "Markers" => (object)((dynamic)points).Markers,
                "Highpoint" => (object)((dynamic)points).Highpoint,
                "Lowpoint" => (object)((dynamic)points).Lowpoint,
                "Negative" => (object)((dynamic)points).Negative,
                "Firstpoint" => (object)((dynamic)points).Firstpoint,
                "Lastpoint" => (object)((dynamic)points).Lastpoint,
                _ => null,
            };
        }
        catch { return null; }
    }

    private static bool? TrySparkVisible(object? points, string property)
    {
        if (points is null) return null;
        var point = GetSparkPoint(points, property);
        try { return point is null ? null : TryComBool(point, "Visible"); }
        finally { RotHelper.ReleaseComReference(point); }
    }

    private static JsonObject ReadSparklineGroupState(object group, string sheetName)
    {
        var state = new JsonObject { ["sheet"] = sheetName };
        object? location = null;
        object? points = null;
        object? seriesColor = null;
        try
        {
            var type = TryComInt(group, "Type");
            state["type"] = type;
            state["typeName"] = type is not null && SparklineTypeNames.TryGetValue(type.Value, out var name)
                ? name : "unknown";
            try { state["sourceData"] = TryComString(group, "SourceData"); }
            catch { state["sourceData"] = null; }
            try
            {
                location = (object)((dynamic)group).Location;
                state["location"] = Convert.ToString(((dynamic)location).Address(false, false), CultureInfo.InvariantCulture);
            }
            catch { state["location"] = null; }
            finally { RotHelper.ReleaseComReference(location); location = null; }
            try { state["count"] = Convert.ToInt32(((dynamic)group).Count, CultureInfo.InvariantCulture); }
            catch { state["count"] = null; }
            try
            {
                points = (object)((dynamic)group).Points;
                state["markers"] = Js(TrySparkVisible(points, "Markers"));
                state["showHigh"] = Js(TrySparkVisible(points, "Highpoint"));
                state["showLow"] = Js(TrySparkVisible(points, "Lowpoint"));
                state["showNegative"] = Js(TrySparkVisible(points, "Negative"));
                state["showFirst"] = Js(TrySparkVisible(points, "Firstpoint"));
                state["showLast"] = Js(TrySparkVisible(points, "Lastpoint"));
            }
            catch { /* points are optional detail */ }
            try
            {
                seriesColor = (object)((dynamic)group).SeriesColor;
                state["lineColor"] = Js(TryComDouble(seriesColor, "Color"));
            }
            catch { state["lineColor"] = null; }
            return state;
        }
        finally
        {
            RotHelper.ReleaseComReference(seriesColor);
            RotHelper.ReleaseComReference(points);
            RotHelper.ReleaseComReference(location);
        }
    }

    private static object? FindSparklineGroup(object sheet, string canonicalAddress, out JsonObject? state, string sheetName)
    {
        state = null;
        foreach (var group in EnumerateSparklineGroups(sheet))
        {
            var current = ReadSparklineGroupState(group, sheetName);
            if (string.Equals(Json.GetString(current, "location"), canonicalAddress, StringComparison.OrdinalIgnoreCase))
            {
                state = current;
                return group;
            }
            RotHelper.ReleaseComReference(group);
        }
        return null;
    }

    internal static List<JsonObject> ReadSparklineStates(object sheet, string sheetName)
    {
        var items = new List<JsonObject>();
        foreach (var group in EnumerateSparklineGroups(sheet))
        {
            try { items.Add(ReadSparklineGroupState(group, sheetName)); }
            finally { RotHelper.ReleaseComReference(group); }
        }
        return items;
    }

    private static void ValidateSparklineLocationShape(string address, string opName)
    {
        if (!ExcelDataOperationsContract.TryParseA1(address, out var box, allowUnion: false))
            throw new InvalidOperationException($"[EXCEL_SPARKLINE_LOCATION] {opName} location is not a rectangular A1 range");
        if (box.Rows != 1 && box.Columns != 1)
            throw new InvalidOperationException(
                $"[EXCEL_SPARKLINE_LOCATION] {opName} location must be one cell, one row, or one column");
    }

    private static void ApplySparklineOptions(object group, JsonObject op)
    {
        object? points = null;
        try
        {
            if (op.ContainsKey("sourceData") && Json.GetString(op, "sourceData") is string source &&
                !string.IsNullOrWhiteSpace(source))
                ((dynamic)group).SourceData = source;
            var hasPoints = op.ContainsKey("markers") || op.ContainsKey("showHigh") || op.ContainsKey("showLow") ||
                            op.ContainsKey("showNegative") || op.ContainsKey("showFirst") || op.ContainsKey("showLast");
            if (hasPoints)
            {
                try { points = (object)((dynamic)group).Points; }
                catch (Exception ex) { throw new InvalidOperationException($"[EXCEL_SPARKLINE_POINTS] points are unavailable: {ex.Message}"); }
                SetSparkVisible(points, "Markers", Json.GetBool(op, "markers"), op.ContainsKey("markers"));
                SetSparkVisible(points, "Highpoint", Json.GetBool(op, "showHigh"), op.ContainsKey("showHigh"));
                SetSparkVisible(points, "Lowpoint", Json.GetBool(op, "showLow"), op.ContainsKey("showLow"));
                SetSparkVisible(points, "Negative", Json.GetBool(op, "showNegative"), op.ContainsKey("showNegative"));
                SetSparkVisible(points, "Firstpoint", Json.GetBool(op, "showFirst"), op.ContainsKey("showFirst"));
                SetSparkVisible(points, "Lastpoint", Json.GetBool(op, "showLast"), op.ContainsKey("showLast"));
            }
            if (op.ContainsKey("lineColor"))
            {
                if (!ExcelStyleContract.TryParseColor(op["lineColor"], out var ole, errors: null, key: "lineColor"))
                    throw new InvalidOperationException("[EXCEL_SPARKLINE_COLOR] lineColor is not a valid color");
                object? seriesColor = null;
                try
                {
                    seriesColor = (object)((dynamic)group).SeriesColor;
                    ((dynamic)seriesColor).Color = ole;
                }
                finally { RotHelper.ReleaseComReference(seriesColor); }
            }
        }
        finally { RotHelper.ReleaseComReference(points); }
    }

    private static void SetSparkVisible(object points, string property, bool value, bool present)
    {
        if (!present) return;
        var point = GetSparkPoint(points, property);
        try
        {
            if (point is null)
                throw new InvalidOperationException($"[EXCEL_SPARKLINE_POINTS] {property} is unavailable");
            ((dynamic)point).Visible = value;
        }
        finally { RotHelper.ReleaseComReference(point); }
    }

    private static void PreviewCreateSparkline(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var location = BindRange(workbook, op, "location");
        EnsureSheetWritable(location.Sheet);
        ValidateSparklineLocationShape(location.Address, "create_sparkline");
        using var source = BindRange(workbook, op, "sourceData", independentSource: true);
        var existing = FindSparklineGroup(location.Sheet, location.Address, out _, location.SheetName);
        try
        {
            if (existing is not null)
                throw new InvalidOperationException(
                    $"[EXCEL_SPARKLINE_EXISTS] a sparkline group already covers {location.SheetName}!{location.Address}");
        }
        finally { RotHelper.ReleaseComReference(existing); }
        preview.Affected.Add(new AffectedRef("sparkline", $"{location.SheetName}!{location.Address}"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{location.SheetName}!{location.Address}",
            Before = null,
            After = op.DeepClone(),
        });
    }

    private static void ApplyCreateSparkline(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var location = BindRange(workbook, op, "location");
        EnsureSheetWritable(location.Sheet);
        ValidateSparklineLocationShape(location.Address, "create_sparkline");
        using var source = BindRange(workbook, op, "sourceData", independentSource: true);
        var existing = FindSparklineGroup(location.Sheet, location.Address, out _, location.SheetName);
        try
        {
            if (existing is not null)
                throw new InvalidOperationException(
                    $"[EXCEL_SPARKLINE_EXISTS] a sparkline group already covers {location.SheetName}!{location.Address}");
        }
        finally { RotHelper.ReleaseComReference(existing); }
        ExcelWorkbookOpsContract.TrySparklineType(Json.GetString(op, "type"), out var type);
        var sourceRef = ExcelDataRangeBinding.QualifiedA1(source.SheetName, source.Address);
        object? locationRange = null;
        object? groups = null;
        object? group = null;
        try
        {
            locationRange = (object)((dynamic)location.Sheet).Range(location.Address);
            groups = (object)((dynamic)locationRange).SparklineGroups;
            try { group = (object)((dynamic)groups).Add(type, sourceRef); }
            catch (Exception ex) { throw new InvalidOperationException($"[EXCEL_SPARKLINE_CREATE_FAILED] {ex.Message}"); }
            ApplySparklineOptions(group, op);
            checkedCells++;
            var actual = ReadSparklineGroupState(group, location.SheetName);
            VerifySparklineRequested(actual, op, location.Address, mismatches, $"{location.SheetName}!{location.Address}");
            execution.Affected.Add(new AffectedRef("sparkline", $"{location.SheetName}!{location.Address}"));
        }
        finally
        {
            RotHelper.ReleaseComReference(group);
            RotHelper.ReleaseComReference(groups);
            RotHelper.ReleaseComReference(locationRange);
        }
    }

    private static void VerifySparklineRequested(JsonObject actual, JsonObject op, string address,
        List<string> mismatches, string label)
    {
        if (!string.Equals(Json.GetString(actual, "location"), address, StringComparison.OrdinalIgnoreCase))
            mismatches.Add($"{label}: sparkline location readback mismatch");
        if (op.ContainsKey("type") && ExcelWorkbookOpsContract.TrySparklineType(Json.GetString(op, "type"), out var type) &&
            Json.GetInt(actual, "type") != type)
            mismatches.Add($"{label}: sparkline type readback mismatch");
        if (op.ContainsKey("sourceData") &&
            !string.Equals(NormalizeA1(Json.GetString(actual, "sourceData")), NormalizeA1(Json.GetString(op, "sourceData")), StringComparison.OrdinalIgnoreCase))
            mismatches.Add($"{label}: sparkline sourceData readback mismatch");
        foreach (var (field, state) in new[] { ("markers", "markers"), ("showHigh", "showHigh"), ("showLow", "showLow"),
                     ("showNegative", "showNegative"), ("showFirst", "showFirst"), ("showLast", "showLast") })
        {
            if (!op.ContainsKey(field)) continue;
            if (actual[state] is not JsonValue flag || !flag.TryGetValue<bool>(out var got) || got != Json.GetBool(op, field))
                mismatches.Add($"{label}: sparkline {field} readback mismatch");
        }
        if (op.ContainsKey("lineColor") &&
            ExcelStyleContract.TryParseColor(op["lineColor"], out var ole, errors: null, key: "lineColor") &&
            (!ExcelDataOperationsContract.TryGetFiniteNumber(actual["lineColor"], out var actualColor) ||
             Math.Abs(actualColor - ole) > 0.5))
            mismatches.Add($"{label}: sparkline lineColor readback mismatch");
    }

    private static string NormalizeA1(string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return "";
        var text = value.Trim().Trim('\'');
        var bang = text.IndexOf('!');
        var address = bang >= 0 ? text[(bang + 1)..] : text;
        return address.Replace("$", "", StringComparison.Ordinal).ToUpperInvariant();
    }

    private static void PreviewUpdateSparkline(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var location = BindRange(workbook, op, "location");
        EnsureSheetWritable(location.Sheet);
        if (!string.IsNullOrWhiteSpace(Json.GetString(op, "sourceData")))
        {
            using var source = BindRange(workbook, op, "sourceData", independentSource: true);
        }
        var group = FindSparklineGroup(location.Sheet, location.Address, out var state, location.SheetName);
        try
        {
            if (group is null)
                throw new InvalidOperationException(
                    $"[EXCEL_SPARKLINE_NOT_FOUND] no sparkline group covers {location.SheetName}!{location.Address}");
            preview.Affected.Add(new AffectedRef("sparkline", $"{location.SheetName}!{location.Address}"));
            preview.Diff.Add(new DiffEntry
            {
                Ref = $"{location.SheetName}!{location.Address}",
                Before = state?.DeepClone(),
                After = op.DeepClone(),
            });
        }
        finally { RotHelper.ReleaseComReference(group); }
    }

    private static void ApplyUpdateSparkline(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var location = BindRange(workbook, op, "location");
        EnsureSheetWritable(location.Sheet);
        var group = FindSparklineGroup(location.Sheet, location.Address, out var before, location.SheetName);
        if (group is null)
            throw new InvalidOperationException(
                $"[EXCEL_SPARKLINE_NOT_FOUND] no sparkline group covers {location.SheetName}!{location.Address}");
        try
        {
            var currentType = Json.GetInt(before, "type");
            if (op.ContainsKey("type") && ExcelWorkbookOpsContract.TrySparklineType(Json.GetString(op, "type"), out var wanted) &&
                currentType != wanted)
            {
                var recreated = (JsonObject)op.DeepClone();
                if (!string.IsNullOrWhiteSpace(Json.GetString(op, "sourceData")))
                {
                    using var source = BindRange(workbook, op, "sourceData", independentSource: true);
                    recreated["sourceData"] = ExcelDataRangeBinding.QualifiedA1(source.SheetName, source.Address);
                }
                RecreateSparklineGroup(location, recreated, wanted, before);
            }
            else
            {
                if (!string.IsNullOrWhiteSpace(Json.GetString(op, "sourceData")))
                {
                    using var source = BindRange(workbook, op, "sourceData", independentSource: true);
                    ((dynamic)group).SourceData = ExcelDataRangeBinding.QualifiedA1(source.SheetName, source.Address);
                }
                ApplySparklineOptions(group, op);
            }
            checkedCells++;
            var refreshed = FindSparklineGroup(location.Sheet, location.Address, out _, location.SheetName);
            try
            {
                if (refreshed is null)
                    throw new InvalidOperationException("[EXCEL_SPARKLINE_NOT_FOUND] sparkline group vanished during update");
                var actual = ReadSparklineGroupState(refreshed, location.SheetName);
                VerifySparklineRequested(actual, op, location.Address, mismatches, $"{location.SheetName}!{location.Address}");
            }
            finally { RotHelper.ReleaseComReference(refreshed); }
            execution.Affected.Add(new AffectedRef("sparkline", $"{location.SheetName}!{location.Address}"));
        }
        finally { RotHelper.ReleaseComReference(group); }
    }

    private static void RecreateSparklineGroup(DataRangeLease location, JsonObject op, int type, JsonObject? before)
    {
        object? locationRange = null;
        object? groups = null;
        object? group = null;
        try
        {
            var old = FindSparklineGroup(location.Sheet, location.Address, out _, location.SheetName);
            try
            {
                if (old is null)
                    throw new InvalidOperationException("[EXCEL_SPARKLINE_NOT_FOUND] sparkline group vanished during update");
                ((dynamic)old).Delete();
            }
            finally { RotHelper.ReleaseComReference(old); }
            string sourceRef;
            if (op.ContainsKey("sourceData") && !string.IsNullOrWhiteSpace(Json.GetString(op, "sourceData")))
            {
                sourceRef = Json.GetString(op, "sourceData")!;
            }
            else
            {
                sourceRef = Json.GetString(before, "sourceData")
                    ?? throw new InvalidOperationException("[EXCEL_SPARKLINE_SOURCE] previous sourceData is unreadable; pass sourceData explicitly");
            }
            locationRange = (object)((dynamic)location.Sheet).Range(location.Address);
            groups = (object)((dynamic)locationRange).SparklineGroups;
            group = (object)((dynamic)groups).Add(type, sourceRef);
            var merged = (JsonObject)op.DeepClone();
            merged.Remove("type");
            ApplySparklineOptions(group, merged);
        }
        finally
        {
            RotHelper.ReleaseComReference(group);
            RotHelper.ReleaseComReference(groups);
            RotHelper.ReleaseComReference(locationRange);
        }
    }

    private static void PreviewDeleteSparkline(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var location = BindRange(workbook, op, "location");
        EnsureSheetWritable(location.Sheet);
        var group = FindSparklineGroup(location.Sheet, location.Address, out var state, location.SheetName);
        try
        {
            if (group is null)
                throw new InvalidOperationException(
                    $"[EXCEL_SPARKLINE_NOT_FOUND] no sparkline group covers {location.SheetName}!{location.Address}");
            preview.Affected.Add(new AffectedRef("sparkline", $"{location.SheetName}!{location.Address}"));
            preview.Diff.Add(new DiffEntry
            {
                Ref = $"{location.SheetName}!{location.Address}",
                Before = state?.DeepClone(),
                After = JsonValue.Create("deleted"),
            });
        }
        finally { RotHelper.ReleaseComReference(group); }
    }

    private static void ApplyDeleteSparkline(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var location = BindRange(workbook, op, "location");
        EnsureSheetWritable(location.Sheet);
        var group = FindSparklineGroup(location.Sheet, location.Address, out _, location.SheetName);
        if (group is null)
            throw new InvalidOperationException(
                $"[EXCEL_SPARKLINE_NOT_FOUND] no sparkline group covers {location.SheetName}!{location.Address}");
        try
        {
            ((dynamic)group).Delete();
            checkedCells++;
            var remaining = FindSparklineGroup(location.Sheet, location.Address, out _, location.SheetName);
            try
            {
                if (remaining is not null)
                    mismatches.Add($"{location.SheetName}!{location.Address}: sparkline group still present after delete");
            }
            finally { RotHelper.ReleaseComReference(remaining); }
            execution.Affected.Add(new AffectedRef("sparkline", $"{location.SheetName}!{location.Address} deleted"));
        }
        finally { RotHelper.ReleaseComReference(group); }
    }

    // ------------------------------------------------------------------
    // Slicer helpers
    // ------------------------------------------------------------------

    private static IEnumerable<object> EnumerateSlicerCaches(object workbook)
    {
        object? caches = null;
        var items = new List<object>();
        try
        {
            try { caches = (object)((dynamic)workbook).SlicerCaches; }
            catch { return items; }
            var count = Convert.ToInt32(((dynamic)caches).Count, CultureInfo.InvariantCulture);
            for (var i = 1; i <= count; i++)
                items.Add((object)((dynamic)caches).Item(i));
            return items;
        }
        finally { RotHelper.ReleaseComReference(caches); }
    }

    private static JsonObject ReadSlicerCacheState(object cache)
    {
        var state = new JsonObject { ["name"] = TryComString(cache, "Name") };
        try { state["cacheType"] = Convert.ToInt32(((dynamic)cache).SlicerCacheType, CultureInfo.InvariantCulture); }
        catch { state["cacheType"] = null; }
        try
        {
            object? pivots = (object)((dynamic)cache).PivotTables;
            try
            {
                var count = Convert.ToInt32(((dynamic)pivots).Count, CultureInfo.InvariantCulture);
                if (count > 0)
                {
                    object? first = null;
                    try
                    {
                        first = (object)((dynamic)pivots).Item(1);
                        state["sourceKind"] = "pivot";
                        state["sourceName"] = TryComString(first, "Name");
                    }
                    finally { RotHelper.ReleaseComReference(first); }
                }
            }
            finally { RotHelper.ReleaseComReference(pivots); }
        }
        catch { /* pivot source is optional detail */ }
        if (Json.GetString(state, "sourceKind") is null)
        {
            try
            {
                object? listObject = (object)((dynamic)cache).ListObject;
                try
                {
                    state["sourceKind"] = "table";
                    state["sourceName"] = TryComString(listObject, "Name");
                }
                finally { RotHelper.ReleaseComReference(listObject); }
            }
            catch { state["sourceKind"] = "unknown"; }
        }
        var slicers = new JsonArray();
        object? owned = null;
        try
        {
            try { owned = (object)((dynamic)cache).Slicers; }
            catch { owned = null; }
            if (owned is not null)
            {
                var count = Convert.ToInt32(((dynamic)owned).Count, CultureInfo.InvariantCulture);
                for (var i = 1; i <= count; i++)
                {
                    object? slicer = null;
                    try
                    {
                        slicer = (object)((dynamic)owned).Item(i);
                        object? parent = null;
                        try
                        {
                            parent = (object)((dynamic)slicer).Parent;
                            slicers.Add(new JsonObject
                            {
                                ["name"] = TryComString(slicer, "Name"),
                                ["caption"] = TryComString(slicer, "Caption"),
                                ["sheet"] = TryComString(parent, "Name"),
                                ["top"] = Js(TryComDouble(slicer, "Top")),
                                ["left"] = Js(TryComDouble(slicer, "Left")),
                                ["width"] = Js(TryComDouble(slicer, "Width")),
                                ["height"] = Js(TryComDouble(slicer, "Height")),
                            });
                        }
                        finally { RotHelper.ReleaseComReference(parent); }
                    }
                    finally { RotHelper.ReleaseComReference(slicer); }
                }
            }
        }
        finally { RotHelper.ReleaseComReference(owned); }
        state["slicers"] = slicers;
        return state;
    }

    internal static List<JsonObject> ReadSlicerCacheStates(object workbook)
    {
        var items = new List<JsonObject>();
        foreach (var cache in EnumerateSlicerCaches(workbook))
        {
            try { items.Add(ReadSlicerCacheState(cache)); }
            finally { RotHelper.ReleaseComReference(cache); }
        }
        return items;
    }

    private static object? FindSlicerCache(object workbook, string name)
    {
        foreach (var cache in EnumerateSlicerCaches(workbook))
        {
            if (string.Equals(ReadComString(cache, "Name"), name, StringComparison.OrdinalIgnoreCase))
                return cache;
            RotHelper.ReleaseComReference(cache);
        }
        return null;
    }

    private static void PreviewCreateSlicer(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var sheet = BindSheet(workbook, op);
        EnsureSheetWritable(sheet.Sheet);
        var source = Json.GetString(op, "source")!;
        var field = Json.GetString(op, "field")!;
        var requestedName = Json.GetString(op, "name")!;
        var duplicate = FindSlicer(workbook, requestedName, out _);
        try
        {
            if (duplicate is not null)
                throw new InvalidOperationException($"[EXCEL_SLICER_NAME_DUPLICATE] slicer '{requestedName}' already exists");
        }
        finally { RotHelper.ReleaseComReference(duplicate); }
        var cache = FindSlicerCache(workbook, requestedName);
        try
        {
            if (cache is not null)
                throw new InvalidOperationException($"[EXCEL_SLICER_NAME_DUPLICATE] slicer cache '{requestedName}' already exists");
        }
        finally { RotHelper.ReleaseComReference(cache); }
        var resolved = ResolveSlicerSource(workbook, source, field, preview.Warnings);
        try
        {
            preview.Affected.Add(new AffectedRef("slicer", $"{sheet.SheetName}!{requestedName}"));
            preview.Diff.Add(new DiffEntry
            {
                Ref = $"{sheet.SheetName}:slicer",
                Before = null,
                After = op.DeepClone(),
            });
        }
        finally { ReleaseSlicerSource(resolved); }
    }

    private sealed record SlicerSource(object? Table, object? TableSheet, object? Pivot, object? PivotSheet, object? PivotField, string Kind);

    private static void ReleaseSlicerSource(SlicerSource? source)
    {
        if (source is null) return;
        RotHelper.ReleaseComReference(source.Table);
        RotHelper.ReleaseComReference(source.TableSheet);
        RotHelper.ReleaseComReference(source.Pivot);
        RotHelper.ReleaseComReference(source.PivotSheet);
        RotHelper.ReleaseComReference(source.PivotField);
    }

    private static SlicerSource ResolveSlicerSource(object workbook, string source, string field, List<string>? warnings)
    {
        var table = FindListObjectInWorkbook(workbook, source, out var tableSheet);
        if (table is not null)
        {
            try
            {
                VerifyTableColumn(table, source, field);
                return new SlicerSource(table, tableSheet, null, null, null, "table");
            }
            catch
            {
                RotHelper.ReleaseComReference(table);
                RotHelper.ReleaseComReference(tableSheet);
                throw;
            }
        }
        RotHelper.ReleaseComReference(tableSheet);
        var pivot = FindPivotInWorkbook(workbook, source, out var pivotSheet);
        if (pivot is null)
        {
            RotHelper.ReleaseComReference(pivotSheet);
            throw new InvalidOperationException($"[EXCEL_SLICER_SOURCE] '{source}' is neither a ListObject nor a PivotTable in this workbook");
        }
        object? pivotField = null;
        try
        {
            try { pivotField = (object)((dynamic)pivot).PivotFields(field); }
            catch
            {
                warnings?.Add($"pivot field '{field}' could not be enumerated; apply will fail closed if it is missing");
            }
            return new SlicerSource(null, null, pivot, pivotSheet, pivotField, "pivot");
        }
        catch
        {
            RotHelper.ReleaseComReference(pivot);
            RotHelper.ReleaseComReference(pivotSheet);
            RotHelper.ReleaseComReference(pivotField);
            throw;
        }
    }

    private static void VerifyTableColumn(object table, string source, string field)
    {
        object? columns = null;
        try
        {
            columns = (object)((dynamic)table).ListColumns;
            var count = Convert.ToInt32(((dynamic)columns).Count, CultureInfo.InvariantCulture);
            for (var i = 1; i <= count; i++)
            {
                object? column = null;
                try
                {
                    column = (object)((dynamic)columns).Item(i);
                    if (string.Equals(ReadComString(column, "Name"), field, StringComparison.OrdinalIgnoreCase))
                        return;
                }
                finally { RotHelper.ReleaseComReference(column); }
            }
            throw new InvalidOperationException($"[EXCEL_SLICER_FIELD] table '{source}' has no column '{field}'");
        }
        finally { RotHelper.ReleaseComReference(columns); }
    }

    private static void ApplyCreateSlicer(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var sheet = BindSheet(workbook, op);
        EnsureSheetWritable(sheet.Sheet);
        var source = Json.GetString(op, "source")!;
        var field = Json.GetString(op, "field")!;
        var requestedName = Json.GetString(op, "name")!;
        var caption = Json.GetString(op, "caption");
        var slicerSource = ResolveSlicerSource(workbook, source, field, execution.Warnings);
        object? caches = null;
        object? cache = null;
        object? owned = null;
        object? slicer = null;
        try
        {
            caches = (object)((dynamic)workbook).SlicerCaches;
            try
            {
                cache = slicerSource.Kind == "table"
                    ? (object)((dynamic)caches).Add2(slicerSource.Table!, field, requestedName)
                    : (object)((dynamic)caches).Add2(slicerSource.Pivot!,
                        slicerSource.PivotField ?? (object)field, requestedName);
            }
            catch (Exception ex) { throw new InvalidOperationException($"[EXCEL_SLICER_CACHE_FAILED] {ex.Message}"); }
            var actualCacheName = ReadComString(cache, "Name") ?? "";
            owned = (object)((dynamic)cache).Slicers;
            var position = Json.GetObj(op, "position");
            try
            {
                slicer = (object)((dynamic)owned).Add(
                    (object)sheet.Sheet,
                    Type.Missing,
                    requestedName,
                    string.IsNullOrWhiteSpace(caption) ? Type.Missing : (object)caption,
                    PositionNumber(position, "top"),
                    PositionNumber(position, "left"),
                    PositionNumber(position, "width"),
                    PositionNumber(position, "height"));
            }
            catch (Exception ex) { throw new InvalidOperationException($"[EXCEL_SLICER_CREATE_FAILED] {ex.Message}"); }
            checkedCells++;
            var actualSlicerName = ReadComString(slicer, "Name") ?? "";
            var found = FindSlicer(workbook, actualSlicerName, out _);
            try
            {
                if (found is null)
                    mismatches.Add($"slicer '{actualSlicerName}' was not read back after create");
            }
            finally { RotHelper.ReleaseComReference(found); }
            execution.Affected.Add(new AffectedRef("slicer", $"{sheet.SheetName}!{actualSlicerName} (cache {actualCacheName})"));
        }
        finally
        {
            RotHelper.ReleaseComReference(slicer);
            RotHelper.ReleaseComReference(owned);
            RotHelper.ReleaseComReference(cache);
            RotHelper.ReleaseComReference(caches);
            ReleaseSlicerSource(slicerSource);
        }
    }

    private static object PositionNumber(JsonObject? position, string field)
    {
        if (position is null) return Type.Missing;
        if (!position.ContainsKey(field)) return Type.Missing;
        return ExcelDataOperationsContract.TryGetFiniteNumber(position[field], out var number)
            ? number : Type.Missing;
    }

    private static void PreviewDeleteSlicer(object workbook, JsonObject op, ApplyPreview preview)
    {
        var slicer = FindSlicer(workbook, Json.GetString(op, "name")!, out _);
        try
        {
            if (slicer is null)
                throw new InvalidOperationException($"[EXCEL_SLICER_NOT_FOUND] slicer '{Json.GetString(op, "name")}' was not found");
            preview.Affected.Add(new AffectedRef("slicer", Json.GetString(op, "name")!));
            preview.Diff.Add(new DiffEntry
            {
                Ref = $"slicer:{Json.GetString(op, "name")}",
                Before = JsonValue.Create("present"),
                After = JsonValue.Create("deleted"),
            });
        }
        finally { RotHelper.ReleaseComReference(slicer); }
    }

    private static void ApplyDeleteSlicer(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        var name = Json.GetString(op, "name")!;
        object? owningSheet = null;
        var slicer = FindSlicer(workbook, name, out owningSheet);
        if (slicer is null)
            throw new InvalidOperationException($"[EXCEL_SLICER_NOT_FOUND] slicer '{name}' was not found");
        object? cache = null;
        try
        {
            try { cache = (object)((dynamic)slicer).SlicerCache; }
            catch { cache = null; }
            ((dynamic)slicer).Delete();
            checkedCells++;
            var remaining = FindSlicer(workbook, name, out _);
            try
            {
                if (remaining is not null)
                    mismatches.Add($"slicer '{name}' is still present after delete");
            }
            finally { RotHelper.ReleaseComReference(remaining); }
            if (cache is not null)
            {
                try
                {
                    object? owned = (object)((dynamic)cache).Slicers;
                    try
                    {
                        if (Convert.ToInt32(((dynamic)owned).Count, CultureInfo.InvariantCulture) == 0)
                            ((dynamic)cache).Delete();
                    }
                    finally { RotHelper.ReleaseComReference(owned); }
                }
                catch { /* orphan cache cleanup is best effort */ }
            }
            execution.Affected.Add(new AffectedRef("slicer", $"{name} deleted"));
        }
        finally
        {
            RotHelper.ReleaseComReference(slicer);
            RotHelper.ReleaseComReference(owningSheet);
            RotHelper.ReleaseComReference(cache);
        }
    }

    // ------------------------------------------------------------------
    // Cell styles
    // ------------------------------------------------------------------

    internal static List<JsonObject> ReadCellStyleStates(object workbook)
    {
        var items = new List<JsonObject>();
        object? styles = null;
        try
        {
            try { styles = (object)((dynamic)workbook).Styles; }
            catch { return items; }
            var count = 0;
            foreach (var style in EnumerateComCollection(styles))
            {
                if (count >= 500) break;
                try
                {
                    items.Add(new JsonObject
                    {
                        ["name"] = TryComString(style, "Name"),
                        ["builtIn"] = Js(TryComBool(style, "BuiltIn")),
                    });
                    count++;
                }
                finally { RotHelper.ReleaseComReference(style); }
            }
            return items;
        }
        finally { RotHelper.ReleaseComReference(styles); }
    }

    private static IEnumerable<object> EnumerateComCollection(object collection)
    {
        var count = Convert.ToInt32(((dynamic)collection).Count, CultureInfo.InvariantCulture);
        for (var i = 1; i <= count; i++)
            yield return (object)((dynamic)collection).Item(i);
    }

    private static bool WorkbookHasCellStyle(object workbook, string name)
    {
        object? styles = null;
        try
        {
            try { styles = (object)((dynamic)workbook).Styles; }
            catch { return false; }
            foreach (var style in EnumerateComCollection(styles))
            {
                try
                {
                    if (string.Equals(ReadComString(style, "Name"), name, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
                finally { RotHelper.ReleaseComReference(style); }
            }
            return false;
        }
        finally { RotHelper.ReleaseComReference(styles); }
    }

    private static JsonArray ReadStyleNameGrid(object sheet, string address)
    {
        object? range = null;
        var grid = new JsonArray();
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            foreach (var cell in EnumerateRangeCells(range))
            {
                object? style = null;
                try
                {
                    string? name = null;
                    try
                    {
                        style = (object)((dynamic)cell).Style;
                        name = ReadComString(style, "Name");
                    }
                    catch { name = null; }
                    grid.Add(name);
                }
                finally
                {
                    RotHelper.ReleaseComReference(style);
                    RotHelper.ReleaseComReference(cell);
                }
                if (grid.Count >= ExcelWorkbookOpsContract.MaxCellStyleCells) break;
            }
            return grid;
        }
        finally { RotHelper.ReleaseComReference(range); }
    }

    private static void PreviewApplyCellStyle(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var range = BindRange(workbook, op, "range");
        EnsureSheetWritable(range.Sheet);
        var styleName = Json.GetString(op, "styleName")!;
        if (!WorkbookHasCellStyle(workbook, styleName))
            throw new InvalidOperationException(
                $"[EXCEL_STYLE_NOT_FOUND] cell style '{styleName}' does not exist; inspect scope=cellStyles first");
        preview.Affected.Add(new AffectedRef("cellStyle", $"{range.SheetName}!{range.Address}"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{range.SheetName}!{range.Address}",
            Before = JsonValue.Create("current styles"),
            After = JsonValue.Create(styleName),
        });
    }

    private static void ApplyApplyCellStyle(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var range = BindRange(workbook, op, "range");
        EnsureSheetWritable(range.Sheet);
        var styleName = Json.GetString(op, "styleName")!;
        if (!WorkbookHasCellStyle(workbook, styleName))
            throw new InvalidOperationException(
                $"[EXCEL_STYLE_NOT_FOUND] cell style '{styleName}' does not exist; inspect scope=cellStyles first");
        object? comRange = null;
        try
        {
            comRange = (object)((dynamic)range.Sheet).Range(range.Address);
            try { ((dynamic)comRange).Style = styleName; }
            catch (Exception ex) { throw new InvalidOperationException($"[EXCEL_STYLE_APPLY_FAILED] {ex.Message}"); }
            var actual = ReadStyleNameGrid(range.Sheet, range.Address);
            checkedCells += actual.Count;
            foreach (var node in actual)
            {
                var name = node is JsonValue value && value.TryGetValue<string>(out var text) ? text : null;
                if (!string.Equals(name, styleName, StringComparison.OrdinalIgnoreCase))
                {
                    mismatches.Add($"{range.SheetName}!{range.Address}: style readback mismatch");
                    break;
                }
            }
            execution.Affected.Add(new AffectedRef("cellStyle", $"{range.SheetName}!{range.Address}={styleName}"));
        }
        finally { RotHelper.ReleaseComReference(comRange); }
    }
}
