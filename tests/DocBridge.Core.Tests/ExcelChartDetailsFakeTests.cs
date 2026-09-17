using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public sealed class ExcelChartDetailsFakeTests
{
    [Fact]
    public void Readback_UsesRequestedSeriesAndNewlyAppendedTrendline()
    {
        var actual = JsonNode.Parse("""{"series":[{}, {"dataLabels":{"show":true},"trendlines":[{"type":"linear"},{"type":"polynomial","order":3}]}]}""")!.AsObject();
        var request = JsonNode.Parse("""{"series":[{"index":2,"dataLabels":{"show":true},"trendline":{"action":"add","type":"polynomial","order":3}}]}""")!.AsObject();
        Assert.Empty(VerifyDetails(actual, request));
        actual["series"]![1]!["trendlines"]![1]!["order"] = 2;
        Assert.Contains(VerifyDetails(actual, request), value => value.Contains("order"));
    }

    [Fact]
    public void Readback_RejectsMissingSeriesAndWrongDeleteCount()
    {
        Assert.NotEmpty(VerifyDetails(new JsonObject(), JsonNode.Parse("""{"series":[{"index":2,"dataLabels":{"show":true}}]}""")!.AsObject()));
        var actual = JsonNode.Parse("""{"series":[{"trendlines":[]}]}""")!.AsObject();
        var request = JsonNode.Parse("""{"series":[{"trendline":{"action":"delete","index":2,"expectedRemainingCount":1}}]}""")!.AsObject();
        Assert.Contains(VerifyDetails(actual, request), value => value.Contains("count mismatch"));
    }

    [Fact]
    public void DeleteTrendline_RejectsNativeNoOp()
    {
        var series = new FakeSeries(4);
        series.TrendlinesSurface.Items.Add(new FakeTrendline());
        var request = JsonNode.Parse("""{"series":[{"trendline":{"action":"delete","index":1}}]}""")!.AsObject();
        var error = Assert.Throws<InvalidOperationException>(() => ExcelChartDetailsApply.Apply(new FakeChart(series), request));
        Assert.Contains("did not reduce", error.Message);
    }

    private static List<string> VerifyDetails(JsonObject actual, JsonObject request)
    {
        var errors = new List<string>();
        var method = typeof(DocBridge.Core.Adapters.ExcelAdapter).GetMethod("VerifyChartDetails",
            System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static)!;
        method.Invoke(null, [actual, request, "chart", errors]);
        return errors;
    }

    [Fact]
    public void AddTrendline_PassesNativeIntegersAndMissingOptionals()
    {
        var chart = new FakeChart(new FakeSeries(pointCount: 5));

        ExcelChartDetailsApply.Apply(chart, new JsonObject
        {
            ["series"] = new JsonArray(new JsonObject
            {
                ["trendline"] = new JsonObject { ["action"] = "add", ["type"] = "linear" },
            }),
        });

        var add = Assert.Single(chart.Series[0].TrendlinesSurface.AddCalls);
        Assert.IsType<int>(add.Type);
        Assert.Equal(-4132, add.Type);
        Assert.Same(Type.Missing, add.Order);
        Assert.Same(Type.Missing, add.Period);
    }

    [Fact]
    public void UpdateTrendline_SetsRequestedTypeOrderPeriodAndPresentationOnly()
    {
        var series = new FakeSeries(pointCount: 8);
        var existing = new FakeTrendline { Type = -4132, Order = 2, Period = 2, Name = "before" };
        series.TrendlinesSurface.Items.Add(existing);
        var chart = new FakeChart(series);

        ExcelChartDetailsApply.Apply(chart, new JsonObject
        {
            ["series"] = new JsonArray(new JsonObject
            {
                ["trendline"] = new JsonObject
                {
                    ["action"] = "update", ["index"] = 1, ["type"] = "polynomial", ["order"] = 4,
                    ["name"] = "after", ["displayEquation"] = true, ["displayRSquared"] = true,
                },
            }),
        });

        Assert.Empty(series.TrendlinesSurface.AddCalls);
        Assert.Equal(3, existing.Type);
        Assert.Equal(4, existing.Order);
        Assert.Equal(2, existing.Period);
        Assert.Equal("after", existing.Name);
        Assert.True(existing.DisplayEquation);
        Assert.True(existing.DisplayRSquared);
    }

    [Fact]
    public void PointLabels_UseSingularNativeMembers()
    {
        var series = new FakeSeries(pointCount: 1);
        var chart = new FakeChart(series);

        ExcelChartDetailsApply.Apply(chart, new JsonObject
        {
            ["series"] = new JsonArray(new JsonObject
            {
                ["points"] = new JsonArray(new JsonObject
                {
                    ["index"] = 1,
                    ["dataLabels"] = new JsonObject { ["show"] = true, ["value"] = true },
                }),
            }),
        });

        var point = series.PointsSurface.Items[0];
        Assert.True(point.HasDataLabel);
        Assert.True(point.DataLabel.ShowValue);
    }

    [Fact]
    public void InvalidMovingAveragePeriod_IsRejectedBeforeAddMutation()
    {
        var series = new FakeSeries(pointCount: 4);
        var chart = new FakeChart(series);

        var exception = Assert.Throws<InvalidOperationException>(() => ExcelChartDetailsApply.Apply(chart, new JsonObject
        {
            ["series"] = new JsonArray(new JsonObject
            {
                ["trendline"] = new JsonObject
                {
                    ["action"] = "add", ["type"] = "movingAverage", ["period"] = 4,
                },
            }),
        }));

        Assert.Contains("EXCEL_TRENDLINE_PERIOD", exception.Message);
        Assert.Empty(series.TrendlinesSurface.AddCalls);
        Assert.Empty(series.TrendlinesSurface.Items);
    }

    [Fact]
    public void ChartWideLabels_DispatchToEverySeriesAndPreserveOmittedProperties()
    {
        var first = new FakeSeries(pointCount: 2);
        first.HasDataLabels = true;
        first.DataLabelsSurface.ShowValue = true;
        first.DataLabelsSurface.ShowSeriesName = true;
        var second = new FakeSeries(pointCount: 2);
        second.HasDataLabels = true;
        second.DataLabelsSurface.ShowValue = true;
        second.DataLabelsSurface.ShowSeriesName = true;
        var chart = new FakeChart(first, second);

        ExcelChartDetailsApply.Apply(chart, new JsonObject
        {
            ["dataLabels"] = new JsonObject { ["category"] = true },
        });

        foreach (var series in chart.Series)
        {
            Assert.Equal(0, series.ApplyDataLabelsCalls);
            Assert.True(series.DataLabelsSurface.ShowValue);
            Assert.True(series.DataLabelsSurface.ShowSeriesName);
            Assert.True(series.DataLabelsSurface.ShowCategoryName);
        }
    }

    public sealed class FakeChart
    {
        public FakeChart(params FakeSeries[] series) => Series = series.ToList();
        public List<FakeSeries> Series { get; }
        public FakeSeriesCollection SeriesCollection() => new(Series);
    }

    public sealed class FakeSeriesCollection
    {
        private readonly List<FakeSeries> _items;
        public FakeSeriesCollection(List<FakeSeries> items) => _items = items;
        public int Count => _items.Count;
        public FakeSeries Item(int index) => _items[index - 1];
    }

    public sealed class FakeSeries
    {
        public FakeSeries(int pointCount)
        {
            PointsSurface = new FakePoints(Enumerable.Range(0, pointCount).Select(_ => new FakePoint()).ToList());
        }

        public bool HasDataLabels { get; set; }
        public int ApplyDataLabelsCalls { get; private set; }
        public FakeDataLabels DataLabelsSurface { get; } = new();
        public FakeTrendlines TrendlinesSurface { get; } = new();
        public FakePoints PointsSurface { get; }
        public FakeDataLabels DataLabels() => DataLabelsSurface;
        public void ApplyDataLabels()
        {
            ApplyDataLabelsCalls++;
            HasDataLabels = true;
        }
        public FakeTrendlines Trendlines() => TrendlinesSurface;
        public FakePoints Points() => PointsSurface;
    }

    public sealed class FakePoints
    {
        public FakePoints(List<FakePoint> items) => Items = items;
        public List<FakePoint> Items { get; }
        public int Count => Items.Count;
        public FakePoint Item(int index) => Items[index - 1];
    }

    public sealed class FakePoint
    {
        public bool HasDataLabel { get; set; }
        public FakeDataLabels DataLabel { get; } = new();
        public FakeFormat Format { get; } = new();
    }

    public sealed class FakeDataLabels
    {
        public bool ShowValue { get; set; }
        public bool ShowCategoryName { get; set; }
        public bool ShowSeriesName { get; set; }
        public bool ShowPercentage { get; set; }
    }

    public sealed class FakeTrendlines
    {
        public List<FakeTrendline> Items { get; } = new();
        public List<FakeTrendlineAdd> AddCalls { get; } = new();
        public int Count => Items.Count;
        public FakeTrendline Item(int index) => Items[index - 1];
        public FakeTrendline Add(int type, object order, object period)
        {
            AddCalls.Add(new FakeTrendlineAdd(type, order, period));
            var line = new FakeTrendline { Type = type };
            if (order is int nativeOrder) line.Order = nativeOrder;
            if (period is int nativePeriod) line.Period = nativePeriod;
            Items.Add(line);
            return line;
        }
    }

    public sealed record FakeTrendlineAdd(int Type, object Order, object Period);

    public sealed class FakeTrendline
    {
        public int Type { get; set; }
        public int Order { get; set; }
        public int Period { get; set; }
        public string? Name { get; set; }
        public bool DisplayEquation { get; set; }
        public bool DisplayRSquared { get; set; }
        public void Delete() { }
    }

    public sealed class FakeFormat
    {
        public FakeFill Fill { get; } = new();
        public FakeFill Line { get; } = new();
    }

    public sealed class FakeFill
    {
        public FakeColor ForeColor { get; } = new();
    }

    public sealed class FakeColor
    {
        public int RGB { get; set; }
    }
}
