using System.Globalization;
using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class ExcelAdapter
{
    private static void VerifyChartDetails(JsonObject actual, JsonObject op, string label, List<string> mismatches)
    {
        CompareLabels(Json.GetObj(actual, "dataLabels"), Json.GetObj(op, "dataLabels"),
            label + ": chart dataLabels", mismatches);
        var actualSeries = Json.GetArr(actual, "series");
        var requestedSeries = Json.GetArr(op, "series");
        if (requestedSeries is null) return;
        for (var i = 0; i < requestedSeries.Count; i++)
        {
            if (requestedSeries[i] is not JsonObject requested) continue;
            var seriesIndex = Json.GetInt(requested, "index") ?? i + 1;
            if (actualSeries is null || seriesIndex < 1 || seriesIndex > actualSeries.Count ||
                actualSeries[seriesIndex - 1] is not JsonObject observed)
            {
                mismatches.Add(label + $": series[{seriesIndex}] detail readback unavailable");
                continue;
            }
            CompareLabels(Json.GetObj(observed, "dataLabels"), Json.GetObj(requested, "dataLabels"),
                label + $": series[{i}].dataLabels", mismatches);
            ComparePoints(Json.GetArr(observed, "points"), Json.GetArr(requested, "points"),
                label + $": series[{i}]", mismatches);
            if (Json.GetObj(requested, "trendline") is JsonObject trend)
                CompareTrendline(Json.GetArr(observed, "trendlines"), trend, label + $": series[{i}]", mismatches);
        }
    }

    private static void CompareLabels(JsonObject? actual, JsonObject? requested, string label,
        List<string> mismatches)
    {
        if (requested is null) return;
        foreach (var key in new[] { "show", "value", "category", "series", "percentage" })
        {
            if (!requested.ContainsKey(key)) continue;
            if (!TryJsonBoolean(requested[key], out var expected) || actual is null ||
                !TryJsonBoolean(actual[key], out var observed) || observed != expected)
                mismatches.Add(label + "." + key + " readback mismatch");
        }
    }

    private static void ComparePoints(JsonArray? actual, JsonArray? requested, string label, List<string> mismatches)
    {
        if (requested is null) return;
        if (actual is null)
        {
            mismatches.Add(label + ": point detail readback unavailable");
            return;
        }
        foreach (var request in requested.OfType<JsonObject>())
        {
            var index = Json.GetInt(request, "index");
            if (index is null || index < 1 || index > actual.Count || actual[index.Value - 1] is not JsonObject observed)
            {
                mismatches.Add(label + $": point[{index}] readback unavailable");
                continue;
            }
            foreach (var color in new[] { "fillColor", "lineColor" })
            {
                if (!request.ContainsKey(color)) continue;
                var expected = ExcelStyleContract.ParseColor(request[color]!);
                if (!TryJsonInteger(observed[color], out var value) || value != expected)
                    mismatches.Add(label + $": point[{index}].{color} readback mismatch");
            }
            CompareLabels(Json.GetObj(observed, "dataLabels"), Json.GetObj(request, "dataLabels"),
                label + $": point[{index}].dataLabels", mismatches);
        }
    }

    private static void CompareTrendline(JsonArray? trends, JsonObject requested, string label,
        List<string> mismatches)
    {
        if (trends is null)
        {
            mismatches.Add(label + ": trendline readback unavailable");
            return;
        }
        var action = Json.GetString(requested, "action");
        var index = action == "add" ? trends.Count : Json.GetInt(requested, "index") ?? 1;
        if (action == "delete")
        {
            // A surviving later trendline may shift into this index.  It is only provably absent
            // without a before-state when the requested index is beyond the resulting collection.
            if (requested.ContainsKey("expectedRemainingCount") &&
                Json.GetInt(requested, "expectedRemainingCount") != trends.Count)
                mismatches.Add(label + $": trendline delete readback count mismatch");
            return;
        }
        if (index < 1 || index > trends.Count || trends[index - 1] is not JsonObject actual)
        {
            mismatches.Add(label + $": trendline[{index}] readback unavailable");
            return;
        }
        foreach (var key in new[] { "type", "order", "period", "name", "displayEquation", "displayRSquared" })
        {
            if (!requested.ContainsKey(key)) continue;
            if (!JsonNode.DeepEquals(actual[key], requested[key]))
                mismatches.Add(label + $": trendline[{index}].{key} readback mismatch");
        }
    }

    private static JsonObject? ReadChartDataLabels(object owner)
    {
        if (TryReadPointHasDataLabel(owner, out var pointHasLabel))
            return ReadPointDataLabels(owner, pointHasLabel);
        if (TryReadSeriesHasDataLabels(owner, out var seriesHasLabels))
            return ReadSeriesDataLabels(owner, seriesHasLabels);
        return ReadChartWideDataLabels(owner);
    }

    private static JsonObject? ReadChartWideDataLabels(object chart)
    {
        object? collection = null;
        try
        {
            collection = (object)((dynamic)chart).SeriesCollection();
            var count = ReadNativeInteger(((dynamic)collection).Count);
            if (count is null || count < 1) return null;
            var labels = new List<JsonObject>();
            for (var index = 1; index <= count; index++)
            {
                object? series = null;
                try
                {
                    series = (object)((dynamic)collection).Item(index);
                    if (!TryReadSeriesHasDataLabels(series, out var hasLabels)) return null;
                    var state = ReadSeriesDataLabels(series, hasLabels);
                    if (state is null) return null;
                    labels.Add(state);
                }
                finally { RotHelper.ReleaseComReference(series); }
            }
            return MergeLabelStates(labels);
        }
        catch { return null; }
        finally { RotHelper.ReleaseComReference(collection); }
    }

    private static JsonObject? ReadSeriesDataLabels(object series, bool? hasLabels)
    {
        if (hasLabels is null) return null;
        if (!hasLabels.Value) return LabelState(false, null, null, null, null);
        object? labels = null;
        try
        {
            labels = (object)((dynamic)series).DataLabels();
            return LabelState(true, ReadShowValue(labels), ReadShowCategoryName(labels), ReadShowSeriesName(labels),
                ReadShowPercentage(labels));
        }
        catch { return null; }
        finally { RotHelper.ReleaseComReference(labels); }
    }

    private static JsonObject? ReadPointDataLabels(object point, bool? hasLabel)
    {
        if (hasLabel is null) return null;
        if (!hasLabel.Value) return LabelState(false, null, null, null, null);
        object? label = null;
        try
        {
            label = (object)((dynamic)point).DataLabel;
            return LabelState(true, ReadShowValue(label), ReadShowCategoryName(label), ReadShowSeriesName(label),
                ReadShowPercentage(label));
        }
        catch { return null; }
        finally { RotHelper.ReleaseComReference(label); }
    }

    private static JsonObject LabelState(bool? show, bool? value, bool? category, bool? series, bool? percentage) =>
        new()
        {
            ["show"] = Js(show),
            ["value"] = Js(value),
            ["category"] = Js(category),
            ["series"] = Js(series),
            ["percentage"] = Js(percentage),
        };

    private static JsonObject MergeLabelStates(IReadOnlyList<JsonObject> states) =>
        LabelState(MergeLabelProperty(states, "show"), MergeLabelProperty(states, "value"),
            MergeLabelProperty(states, "category"), MergeLabelProperty(states, "series"),
            MergeLabelProperty(states, "percentage"));

    private static bool? MergeLabelProperty(IReadOnlyList<JsonObject> states, string property)
    {
        bool? value = null;
        var initialized = false;
        foreach (var state in states)
        {
            if (!TryJsonBoolean(state[property], out var current)) return null;
            if (initialized && value != current) return null;
            value = current;
            initialized = true;
        }
        return value;
    }

    private static JsonArray ReadTrendlines(object series)
    {
        var output = new JsonArray();
        object? lines = null;
        try
        {
            lines = (object)((dynamic)series).Trendlines();
            var count = ReadNativeInteger(((dynamic)lines).Count);
            if (count is null) return output;
            for (var index = 1; index <= count; index++)
            {
                object? line = null;
                try
                {
                    line = (object)((dynamic)lines).Item(index);
                    output.Add(new JsonObject
                    {
                        ["index"] = index,
                        ["type"] = TrendlineTypeToken(ReadTrendlineTypeValue(line)),
                        ["order"] = Js(ReadTrendlineOrder(line)),
                        ["period"] = Js(ReadTrendlinePeriod(line)),
                        ["name"] = ReadTrendlineName(line),
                        ["displayEquation"] = Js(ReadTrendlineDisplayEquation(line)),
                        ["displayRSquared"] = Js(ReadTrendlineDisplayRSquared(line)),
                    });
                }
                finally { RotHelper.ReleaseComReference(line); }
            }
        }
        catch { }
        finally { RotHelper.ReleaseComReference(lines); }
        return output;
    }

    // Called by the shared chart-series inspection and readback path.
    private static JsonArray? ReadChartPoints(object series)
    {
        object? points = null;
        try
        {
            points = (object)((dynamic)series).Points();
            var count = ReadNativeInteger(((dynamic)points).Count);
            if (count is null) return null;
            var output = new JsonArray();
            for (var index = 1; index <= count; index++)
            {
                object? point = null;
                try
                {
                    point = (object)((dynamic)points).Item(index);
                    output.Add(new JsonObject
                    {
                        ["index"] = index,
                        ["fillColor"] = Js(ReadPointColor(point, fill: true)),
                        ["lineColor"] = Js(ReadPointColor(point, fill: false)),
                        ["dataLabels"] = ReadChartDataLabels(point),
                    });
                }
                finally { RotHelper.ReleaseComReference(point); }
            }
            return output;
        }
        catch { return null; }
        finally { RotHelper.ReleaseComReference(points); }
    }

    private static int? ReadPointColor(object point, bool fill)
    {
        object? format = null;
        object? fillOrLine = null;
        object? foreColor = null;
        try
        {
            format = (object)((dynamic)point).Format;
            fillOrLine = fill ? (object)((dynamic)format).Fill : (object)((dynamic)format).Line;
            foreColor = (object)((dynamic)fillOrLine).ForeColor;
            return ReadNativeInteger(((dynamic)foreColor).RGB);
        }
        catch { return null; }
        finally
        {
            RotHelper.ReleaseComReference(foreColor);
            RotHelper.ReleaseComReference(fillOrLine);
            RotHelper.ReleaseComReference(format);
        }
    }

    private static bool TryReadPointHasDataLabel(object point, out bool? value)
    {
        try { value = ReadNativeBoolean(((dynamic)point).HasDataLabel); return true; }
        catch { value = null; return false; }
    }

    private static bool TryReadSeriesHasDataLabels(object series, out bool? value)
    {
        try { value = ReadNativeBoolean(((dynamic)series).HasDataLabels); return true; }
        catch { value = null; return false; }
    }

    private static bool? ReadShowValue(object labels)
    {
        try { return ReadNativeBoolean(((dynamic)labels).ShowValue); } catch { return null; }
    }
    private static bool? ReadShowCategoryName(object labels)
    {
        try { return ReadNativeBoolean(((dynamic)labels).ShowCategoryName); } catch { return null; }
    }
    private static bool? ReadShowSeriesName(object labels)
    {
        try { return ReadNativeBoolean(((dynamic)labels).ShowSeriesName); } catch { return null; }
    }
    private static bool? ReadShowPercentage(object labels)
    {
        try { return ReadNativeBoolean(((dynamic)labels).ShowPercentage); } catch { return null; }
    }
    private static int? ReadTrendlineTypeValue(object line)
    {
        try { return ReadNativeInteger(((dynamic)line).Type); } catch { return null; }
    }
    private static string? TrendlineTypeToken(int? type) => type switch
    {
        -4132 => "linear",
        5 => "exponential",
        -4133 => "logarithmic",
        3 => "polynomial",
        4 => "power",
        6 => "movingAverage",
        null => null,
        _ => "raw:" + type.Value.ToString(CultureInfo.InvariantCulture),
    };

    private static int? ReadTrendlineOrder(object line)
    {
        try { return ReadNativeInteger(((dynamic)line).Order); } catch { return null; }
    }
    private static int? ReadTrendlinePeriod(object line)
    {
        try { return ReadNativeInteger(((dynamic)line).Period); } catch { return null; }
    }
    private static string? ReadTrendlineName(object line)
    {
        try
        {
            object? value = ((dynamic)line).Name;
            return value is null or DBNull ? null : Convert.ToString(value, CultureInfo.InvariantCulture);
        }
        catch { return null; }
    }
    private static bool? ReadTrendlineDisplayEquation(object line)
    {
        try { return ReadNativeBoolean(((dynamic)line).DisplayEquation); } catch { return null; }
    }
    private static bool? ReadTrendlineDisplayRSquared(object line)
    {
        try { return ReadNativeBoolean(((dynamic)line).DisplayRSquared); } catch { return null; }
    }

    private static bool? ReadNativeBoolean(object? value)
    {
        if (value is null or DBNull) return null;
        if (value is bool boolean) return boolean;
        if (value is sbyte or byte or short or ushort or int or uint or long or ulong)
        {
            var integer = Convert.ToInt64(value, CultureInfo.InvariantCulture);
            return integer switch { 0 => false, -1 => true, 1 => true, _ => null };
        }
        return null;
    }

    private static int? ReadNativeInteger(object? value)
    {
        if (value is null or DBNull) return null;
        try { return Convert.ToInt32(value, CultureInfo.InvariantCulture); }
        catch { return null; }
    }

    private static bool TryJsonBoolean(JsonNode? node, out bool value)
    {
        value = false;
        return node is JsonValue jsonValue && jsonValue.TryGetValue(out value);
    }

    private static bool TryJsonInteger(JsonNode? node, out int value)
    {
        value = 0;
        return node is JsonValue jsonValue && jsonValue.TryGetValue(out value);
    }
}
