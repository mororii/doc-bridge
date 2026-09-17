using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

public static class ExcelChartDetailsApply
{
    public static void Apply(object chart, JsonObject op)
    {
        ArgumentNullException.ThrowIfNull(chart);
        ArgumentNullException.ThrowIfNull(op);

        var chartLabels = Json.GetObj(op, "dataLabels");
        var specs = Json.GetArr(op, "series");
        if (chartLabels is null && specs is null) return;

        object? collection = null;
        try
        {
            collection = (object)((dynamic)chart).SeriesCollection();
            if (chartLabels is not null) ApplyChartLabels(collection, chartLabels);
            if (specs is null) return;

            for (var i = 0; i < specs.Count; i++)
            {
                if (specs[i] is not JsonObject spec) continue;
                var index = Json.GetInt(spec, "index") ?? i + 1;
                EnsureCollectionIndex(collection, index, "series");
                object? series = null;
                try
                {
                    series = (object)((dynamic)collection).Item(index);
                    if (Json.GetObj(spec, "dataLabels") is JsonObject seriesLabels)
                        ApplySeriesLabels(series, seriesLabels);
                    if (Json.GetObj(spec, "trendline") is JsonObject trendline)
                        ApplyTrendline(series, trendline);
                    if (Json.GetArr(spec, "points") is JsonArray points)
                        ApplyPoints(series, points);
                }
                finally { ExcelOwnedCom.ReleaseOnce(series); }
            }
        }
        finally { ExcelOwnedCom.ReleaseOnce(collection); }
    }

    private static void ApplyChartLabels(object seriesCollection, JsonObject labels)
    {
        ValidateLabelRequest(labels);
        var count = CollectionCount(seriesCollection, "series");
        for (var index = 1; index <= count; index++)
        {
            object? series = null;
            try
            {
                series = (object)((dynamic)seriesCollection).Item(index);
                ApplySeriesLabels(series, labels);
            }
            finally { ExcelOwnedCom.ReleaseOnce(series); }
        }
    }

    private static void ApplySeriesLabels(object series, JsonObject labels)
    {
        ValidateLabelRequest(labels);
        var showRequested = labels.ContainsKey("show");
        var hasContent = HasLabelContent(labels);
        if (showRequested && !Json.GetBool(labels, "show"))
        {
            ((dynamic)series).HasDataLabels = false;
            return;
        }

        var hasLabels = Convert.ToBoolean(((dynamic)series).HasDataLabels);
        if ((showRequested || hasContent) && !hasLabels)
        {
            ((dynamic)series).ApplyDataLabels();
            hasLabels = true;
        }
        if (!hasContent) return;

        object? set = null;
        try
        {
            set = (object)((dynamic)series).DataLabels();
            ApplyLabelContent(set, labels);
        }
        finally { ExcelOwnedCom.ReleaseOnce(set); }
    }

    private static void ApplyPointLabels(object point, JsonObject labels)
    {
        ValidateLabelRequest(labels);
        var showRequested = labels.ContainsKey("show");
        var hasContent = HasLabelContent(labels);
        if (showRequested && !Json.GetBool(labels, "show"))
        {
            ((dynamic)point).HasDataLabel = false;
            return;
        }

        var hasLabel = Convert.ToBoolean(((dynamic)point).HasDataLabel);
        if ((showRequested || hasContent) && !hasLabel)
            ((dynamic)point).HasDataLabel = true;
        if (!hasContent) return;

        object? label = null;
        try
        {
            label = (object)((dynamic)point).DataLabel;
            ApplyLabelContent(label, labels);
        }
        finally { ExcelOwnedCom.ReleaseOnce(label); }
    }

    private static void ApplyLabelContent(object labels, JsonObject request)
    {
        if (request.ContainsKey("value")) ((dynamic)labels).ShowValue = Json.GetBool(request, "value");
        if (request.ContainsKey("category")) ((dynamic)labels).ShowCategoryName = Json.GetBool(request, "category");
        if (request.ContainsKey("series")) ((dynamic)labels).ShowSeriesName = Json.GetBool(request, "series");
        if (request.ContainsKey("percentage")) ((dynamic)labels).ShowPercentage = Json.GetBool(request, "percentage");
    }

    private static bool HasLabelContent(JsonObject labels) =>
        labels.ContainsKey("value") || labels.ContainsKey("category") ||
        labels.ContainsKey("series") || labels.ContainsKey("percentage");

    private static void ValidateLabelRequest(JsonObject labels)
    {
        if (labels.ContainsKey("show") && !Json.GetBool(labels, "show") && HasLabelContent(labels))
            throw new InvalidOperationException("[EXCEL_CHART_LABELS] dataLabels.show=false cannot be combined with label content options");
    }

    private static void ApplyTrendline(object series, JsonObject spec)
    {
        var action = Json.GetString(spec, "action");
        if (action is not ("add" or "update" or "delete"))
            throw new InvalidOperationException("[EXCEL_TRENDLINE] action must be add, update, or delete");

        var index = Json.GetInt(spec, "index") ?? 1;
        object? lines = null;
        object? line = null;
        try
        {
            lines = (object)((dynamic)series).Trendlines();
            var lineCount = CollectionCount(lines, "trendline");
            if (action is "update" or "delete") EnsureIndex(index, lineCount, "trendline");
            if (action == "delete")
            {
                line = (object)((dynamic)lines).Item(index);
                ((dynamic)line).Delete();
                if (CollectionCount(lines, "trendline") != lineCount - 1)
                    throw new InvalidOperationException("[EXCEL_TRENDLINE] delete did not reduce the native collection count");
                return;
            }

            if (action == "add")
            {
                var type = TrendlineType(Json.GetString(spec, "type"));
                ValidateTrendlineRequest(series, spec, type, isAdd: true);
                var order = OptionalComInteger(spec, "order");
                var period = OptionalComInteger(spec, "period");
                line = (object)((dynamic)lines).Add(type, order, period);
            }
            else
            {
                line = (object)((dynamic)lines).Item(index);
                var type = spec.ContainsKey("type")
                    ? TrendlineType(Json.GetString(spec, "type"))
                    : ReadTrendlineType(line);
                ValidateTrendlineRequest(series, spec, type, isAdd: false);
                if (spec.ContainsKey("type")) ((dynamic)line).Type = type;
                if (spec.ContainsKey("order")) ((dynamic)line).Order = RequiredInteger(spec, "order");
                if (spec.ContainsKey("period")) ((dynamic)line).Period = RequiredInteger(spec, "period");
            }

            if (spec.ContainsKey("name")) ((dynamic)line).Name = Json.GetString(spec, "name");
            if (spec.ContainsKey("displayEquation")) ((dynamic)line).DisplayEquation = Json.GetBool(spec, "displayEquation");
            if (spec.ContainsKey("displayRSquared")) ((dynamic)line).DisplayRSquared = Json.GetBool(spec, "displayRSquared");
        }
        finally
        {
            ExcelOwnedCom.ReleaseOnce(line);
            ExcelOwnedCom.ReleaseOnce(lines);
        }
    }

    private static void ValidateTrendlineRequest(object series, JsonObject spec, int effectiveType, bool isAdd)
    {
        var hasOrder = spec.ContainsKey("order");
        var hasPeriod = spec.ContainsKey("period");
        var setsPolynomial = spec.ContainsKey("type") && effectiveType == 3;
        var setsMovingAverage = spec.ContainsKey("type") && effectiveType == 6;

        if (effectiveType == 3)
        {
            if ((isAdd || setsPolynomial) && !hasOrder)
                throw new InvalidOperationException("[EXCEL_TRENDLINE_ORDER] polynomial trendline order must be 2..6");
            if (hasOrder)
            {
                var order = RequiredInteger(spec, "order");
                if (order is < 2 or > 6)
                    throw new InvalidOperationException("[EXCEL_TRENDLINE_ORDER] polynomial trendline order must be 2..6");
            }
        }
        else if (hasOrder)
        {
            throw new InvalidOperationException("[EXCEL_TRENDLINE_ORDER] order is valid only for polynomial trendlines");
        }

        if (effectiveType == 6)
        {
            if ((isAdd || setsMovingAverage) && !hasPeriod)
                throw new InvalidOperationException("[EXCEL_TRENDLINE_PERIOD] movingAverage period must be 2..255 and less than actual series point count");
            if (hasPeriod)
            {
                var period = RequiredInteger(spec, "period");
                if (period is < 2 or > 255)
                    throw new InvalidOperationException("[EXCEL_TRENDLINE_PERIOD] movingAverage period must be 2..255 and less than actual series point count");
                var pointCount = SeriesPointCount(series);
                if (period >= pointCount)
                    throw new InvalidOperationException("[EXCEL_TRENDLINE_PERIOD] movingAverage period must be >1 and less than actual series point count");
            }
        }
        else if (hasPeriod)
        {
            throw new InvalidOperationException("[EXCEL_TRENDLINE_PERIOD] period is valid only for movingAverage trendlines");
        }
    }

    private static object OptionalComInteger(JsonObject spec, string property) =>
        spec.ContainsKey(property) ? RequiredInteger(spec, property) : Type.Missing;

    private static int RequiredInteger(JsonObject spec, string property)
    {
        var value = Json.GetInt(spec, property);
        if (value is null) throw new InvalidOperationException($"[EXCEL_TRENDLINE] {property} must be an integer");
        return value.Value;
    }

    private static int ReadTrendlineType(object trendline)
    {
        object? value = ((dynamic)trendline).Type;
        if (value is null || value is DBNull)
            throw new InvalidOperationException("[EXCEL_TRENDLINE] existing trendline Type is unreadable");
        return Convert.ToInt32(value);
    }

    private static int SeriesPointCount(object series)
    {
        object? points = null;
        try
        {
            points = (object)((dynamic)series).Points();
            return CollectionCount(points, "point");
        }
        finally { ExcelOwnedCom.ReleaseOnce(points); }
    }

    private static int TrendlineType(string? type) => type switch
    {
        "linear" => -4132,
        "exponential" => 5,
        "logarithmic" => -4133,
        "polynomial" => 3,
        "power" => 4,
        "movingAverage" => 6,
        _ => throw new InvalidOperationException($"[EXCEL_TRENDLINE] unsupported trendline type '{type}'"),
    };

    private static void ApplyPoints(object series, JsonArray points)
    {
        object? collection = null;
        try
        {
            collection = (object)((dynamic)series).Points();
            var pointCount = CollectionCount(collection, "point");
            foreach (var spec in points.OfType<JsonObject>())
            {
                var index = Json.GetInt(spec, "index");
                if (index is null) throw new InvalidOperationException("[EXCEL_CHART_POINT] point index must be 1-based");
                EnsureIndex(index.Value, pointCount, "point");
                object? point = null;
                try
                {
                    point = (object)((dynamic)collection).Item(index.Value);
                    if (spec.ContainsKey("fillColor")) ApplyPointColor(point, "fillColor", spec["fillColor"]!);
                    if (spec.ContainsKey("lineColor")) ApplyPointColor(point, "lineColor", spec["lineColor"]!);
                    if (Json.GetObj(spec, "dataLabels") is JsonObject labels) ApplyPointLabels(point, labels);
                }
                finally { ExcelOwnedCom.ReleaseOnce(point); }
            }
        }
        finally { ExcelOwnedCom.ReleaseOnce(collection); }
    }

    private static void ApplyPointColor(object point, string property, JsonNode colorValue)
    {
        object? format = null;
        object? fillOrLine = null;
        object? color = null;
        try
        {
            format = (object)((dynamic)point).Format;
            fillOrLine = property == "fillColor"
                ? (object)((dynamic)format).Fill
                : (object)((dynamic)format).Line;
            color = (object)((dynamic)fillOrLine).ForeColor;
            ((dynamic)color).RGB = ExcelStyleContract.ParseColor(colorValue);
        }
        finally
        {
            ExcelOwnedCom.ReleaseOnce(color);
            ExcelOwnedCom.ReleaseOnce(fillOrLine);
            ExcelOwnedCom.ReleaseOnce(format);
        }
    }

    private static int CollectionCount(object collection, string kind)
    {
        object? value = ((dynamic)collection).Count;
        if (value is null || value is DBNull)
            throw new InvalidOperationException($"[EXCEL_CHART_{kind.ToUpperInvariant()}] collection count is unreadable");
        return Convert.ToInt32(value);
    }

    private static void EnsureCollectionIndex(object collection, int index, string kind) =>
        EnsureIndex(index, CollectionCount(collection, kind), kind);

    private static void EnsureIndex(int index, int count, string kind)
    {
        if (index < 1 || index > count)
            throw new InvalidOperationException($"[EXCEL_CHART_{kind.ToUpperInvariant()}] index {index} does not exist");
    }
}
