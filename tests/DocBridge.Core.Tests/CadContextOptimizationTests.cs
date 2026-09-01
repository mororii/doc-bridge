using System.Collections;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class CadContextOptimizationTests
{
    [Fact]
    public void Basic_context_does_not_enumerate_layers_or_modelspace()
    {
        var app = FakeCadApp.Create(entityCount: 260_000, layerCount: 751);
        using var adapter = new CadAdapter(() => app);

        var context = adapter.GetActiveContext(new JsonObject());

        Assert.True(context.Ok, string.Join("; ", context.Errors));
        Assert.Equal("basic", Json.GetString(context.Summary, "detailLevel"));
        Assert.Equal(260_000, Json.GetInt(context.Summary, "entityCount"));
        Assert.Equal(751, Json.GetInt(context.Summary, "layerCount"));
        Assert.Equal(0, app.ActiveDocument.ModelSpace.EnumerationCount);
        Assert.Equal(0, app.ActiveDocument.Layers.EnumerationCount);
        Assert.Equal("omitted", Json.GetString(context.Summary, "entitySummaryStatus"));
        Assert.Empty(Json.GetArr(context.Summary, "layers")!);
        Assert.NotEmpty(Json.GetArr(context.Summary, "nextActions")!);
    }

    [Fact]
    public void Summary_context_caps_samples_and_reports_explicit_coverage()
    {
        var app = FakeCadApp.Create(entityCount: 800, layerCount: 75);
        using var adapter = new CadAdapter(() => app);

        var context = adapter.GetActiveContext(new JsonObject { ["detailLevel"] = "summary" });

        Assert.True(context.Ok, string.Join("; ", context.Errors));
        Assert.Equal(500, Json.GetInt(context.Summary, "entitySummaryScanned"));
        Assert.True(Json.GetBool(context.Summary, "entitySummaryTruncated"));
        Assert.Equal("sampled", Json.GetString(context.Summary, "entitySummaryStatus"));
        Assert.Equal(500, app.ActiveDocument.ModelSpace.EnumerationCount);
        Assert.Equal(50, Json.GetArr(context.Summary, "layers")!.Count);
        Assert.Equal(50, app.ActiveDocument.Layers.EnumerationCount);
        var coverage = Json.GetObj(context.Summary, "coverage")!;
        Assert.False(Json.GetBool(Json.GetObj(coverage, "entityTypeSummary"), "complete"));
        var actions = Json.GetArr(context.Summary, "nextActions")!;
        Assert.Contains(actions, node =>
            Json.GetString(node as JsonObject, "tool") == "cad_query_entities" &&
            Json.GetString(Json.GetObj(node as JsonObject, "arguments"), "scope") == "regions");
    }

    [Fact]
    public void Invalid_context_detail_level_fails_without_touching_com()
    {
        var calls = 0;
        using var adapter = new CadAdapter(() => { calls++; return FakeCadApp.Create(1, 1); });

        var context = adapter.GetActiveContext(new JsonObject { ["detailLevel"] = "full" });

        Assert.False(context.Ok);
        Assert.NotEmpty(context.Errors);
        Assert.Equal(0, calls);
    }

    [Fact]
    public void Entity_query_exposes_executable_continuation_when_page_is_truncated()
    {
        var app = FakeCadApp.Create(entityCount: 5, layerCount: 1);
        using var adapter = new CadAdapter(() => app);

        var result = adapter.Read(new JsonObject
        {
            ["scope"] = "entities",
            ["layer"] = "BASE",
            ["limit"] = 2,
        });

        Assert.True(Json.GetBool(result, "ok"), result.ToJsonString());
        Assert.True(Json.GetBool(result, "truncated"));
        Assert.Equal(2, Json.GetInt(result, "nextStartIndex"));
        var action = Json.GetArr(result, "nextActions")!.Single() as JsonObject;
        var arguments = Json.GetObj(action, "arguments")!;
        Assert.Equal("entities", Json.GetString(arguments, "scope"));
        Assert.Equal("BASE", Json.GetString(arguments, "layer"));
        Assert.Equal(2, Json.GetInt(arguments, "startIndex"));
    }

    [Fact]
    public void Layer_query_reports_page_coverage_and_continuation()
    {
        var app = FakeCadApp.Create(entityCount: 0, layerCount: 5);
        using var adapter = new CadAdapter(() => app);

        var result = adapter.Read(new JsonObject
        {
            ["scope"] = "layers",
            ["limit"] = 2,
        });

        Assert.True(Json.GetBool(result, "truncated"));
        Assert.Equal(2, Json.GetInt(result, "nextStartIndex"));
        Assert.False(Json.GetBool(Json.GetObj(result, "coverage"), "complete"));
        var action = Json.GetArr(result, "nextActions")!.Single() as JsonObject;
        Assert.Equal(2, Json.GetInt(Json.GetObj(action, "arguments"), "startIndex"));
    }

    [Fact]
    public void Region_query_marks_twenty_item_samples_as_truncated_and_guides_window_read()
    {
        var app = FakeCadApp.Create(entityCount: 25, layerCount: 1);
        using var adapter = new CadAdapter(() => app);

        var result = adapter.Read(new JsonObject
        {
            ["scope"] = "regions",
            ["regions"] = new JsonArray
            {
                new JsonObject
                {
                    ["name"] = "sheet-1",
                    ["bounds"] = new JsonObject
                    {
                        ["minX"] = -1, ["minY"] = -1, ["maxX"] = 100, ["maxY"] = 100,
                    },
                    ["boundsMode"] = "inside",
                },
            },
        });

        Assert.True(Json.GetBool(result, "ok"), result.ToJsonString());
        var region = Json.GetArr(result, "regions")!.Single() as JsonObject;
        Assert.Equal(25, Json.GetInt(region, "count"));
        Assert.True(Json.GetBool(Json.GetObj(region, "sampleCoverage"), "truncated"));
        Assert.Equal(20, Json.GetInt(Json.GetObj(region, "sampleCoverage"), "returned"));
        var action = Json.GetArr(region, "nextActions")!.Single() as JsonObject;
        Assert.Equal("window", Json.GetString(Json.GetObj(action, "arguments"), "scope"));
    }

    [Fact]
    public void Region_follow_up_preserves_center_mode_and_each_requested_entity_type()
    {
        var app = FakeCadApp.Create(entityCount: 50, layerCount: 1);
        using var adapter = new CadAdapter(() => app);

        var result = adapter.Read(new JsonObject
        {
            ["scope"] = "regions",
            ["regions"] = new JsonArray
            {
                new JsonObject
                {
                    ["name"] = "typed-region",
                    ["bounds"] = new JsonObject
                    {
                        ["minX"] = -1, ["minY"] = -1, ["maxX"] = 100, ["maxY"] = 100,
                    },
                    ["boundsMode"] = "center",
                    ["entityTypes"] = new JsonArray("Line", "Text"),
                },
            },
        });

        Assert.True(Json.GetBool(result, "ok"), result.ToJsonString());
        var region = Json.GetArr(result, "regions")!.Single() as JsonObject;
        var actions = Json.GetArr(region, "nextActions")!;
        Assert.Equal(2, actions.Count);
        var arguments = actions.Select(node => Json.GetObj(node as JsonObject, "arguments")!).ToArray();
        Assert.All(arguments, item => Assert.Equal("center", Json.GetString(item, "boundsMode")));
        Assert.Equal(new[] { "Line", "Text" },
            arguments.Select(item => Json.GetString(item, "entityType")).OrderBy(value => value));
    }

    public sealed class FakeCadApp
    {
        public long HWND => 0;
        public required FakeDocument ActiveDocument { get; init; }
        public required List<FakeDocument> Documents { get; init; }

        public static FakeCadApp Create(int entityCount, int layerCount)
        {
            var document = new FakeDocument(entityCount, layerCount);
            return new FakeCadApp { ActiveDocument = document, Documents = new List<FakeDocument> { document } };
        }
    }

    public sealed class FakeDocument
    {
        public FakeDocument(int entityCount, int layerCount)
        {
            ModelSpace = new CountingCollection<FakeEntity>(entityCount,
                index => new FakeEntity
                {
                    EntityName = index % 2 == 0 ? "AcDbLine" : "AcDbText",
                    Handle = (index + 1).ToString("X"),
                    X = index,
                });
            Layers = new CountingCollection<FakeLayer>(layerCount,
                index => new FakeLayer { Name = $"L-{index:000}" });
        }

        public string Name => "large-test.dwg";
        public string FullName => "C:\\drawings\\large-test.dwg";
        public CountingCollection<FakeEntity> ModelSpace { get; }
        public CountingCollection<FakeLayer> Layers { get; }
        public int GetVariable(string name) => name == "INSUNITS" ? 4 : 0;
    }

    public sealed class FakeEntity
    {
        public string EntityName { get; init; } = "AcDbLine";
        public string Layer => "BASE";
        public string Handle { get; init; } = "1";
        public double X { get; init; }

        public void GetBoundingBox(out object minimum, out object maximum)
        {
            minimum = new[] { X, X, 0.0 };
            maximum = new[] { X + 0.5, X + 0.5, 0.0 };
        }
    }

    public sealed class FakeLayer
    {
        public string Name { get; init; } = "0";
        public bool LayerOn => true;
        public int Color => 7;
    }

    public sealed class CountingCollection<T> : IEnumerable<T>
    {
        private readonly int _count;
        private readonly Func<int, T> _factory;

        public CountingCollection(int count, Func<int, T> factory)
        {
            _count = count;
            _factory = factory;
        }

        public int Count => _count;
        public int EnumerationCount { get; private set; }
        public T Item(int index) => _factory(index);

        public IEnumerator<T> GetEnumerator()
        {
            for (var index = 0; index < _count; index++)
            {
                EnumerationCount++;
                yield return _factory(index);
            }
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }
}
