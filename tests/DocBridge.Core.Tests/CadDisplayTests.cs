using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class CadDisplayTests
{
    private static JsonObject Move(string handle = "1") => new()
    {
        ["op"] = "move_entities", ["handles"] = new JsonArray(handle), ["dx"] = 1.0, ["dy"] = 0.0,
    };

    [Fact]
    public void Batch_moves_regenerate_once_after_view_restoration()
    {
        var app = new App();
        using var adapter = new CadAdapter(() => app);
        var execution = adapter.Apply(new[] { Move(), Move() }, "test");
        Assert.True(execution.Ok, string.Join(";", execution.Errors));
        Assert.Equal(2, app.ActiveDocument.Entity.Moves);
        Assert.Equal(1, app.ActiveDocument.Regens);
        Assert.Equal(1, app.ActiveDocument.LastRegenType);
        Assert.Equal("completed", (string?)execution.Readback?["displayRefresh"]?["status"]);
        Assert.Equal("not-performed", (string?)execution.Readback?["displayRefresh"]?["visualVerification"]);
        Assert.Equal(0, app.ZoomCalls); // no redundant view write for an unchanged view
        Assert.Equal(0, app.ActiveDocument.StateWrites);
    }

    [Fact]
    public void Refresh_failure_is_a_warning_not_a_replay_or_data_failure()
    {
        var app = new App();
        app.ActiveDocument.RegenFails = true;
        using var adapter = new CadAdapter(() => app);
        var execution = adapter.Apply(new[] { Move() }, "test");
        Assert.True(execution.Ok);
        Assert.True((bool)execution.Readback!["verified"]!);
        Assert.Equal(1, app.ActiveDocument.Entity.Moves);
        Assert.Equal(1, app.ActiveDocument.Regens);
        Assert.Equal("failed", (string?)execution.Readback?["displayRefresh"]?["status"]);
        Assert.Contains(execution.Warnings, w => w.Contains("CAD_DISPLAY_REFRESH_FAILED"));
    }

    [Fact]
    public void Partial_COM_rejection_does_not_repeat_an_already_applied_move()
    {
        var app = new App();
        using var adapter = new CadAdapter(() => app);
        var execution = adapter.Apply(new[] { Move(), Move("rejected"), Move() }, "test");
        Assert.False(execution.Ok);
        Assert.Equal(1, app.ActiveDocument.Entity.Moves);
        Assert.Equal(2, execution.OperationResults.Count);
        Assert.Equal(1, app.ActiveDocument.Regens); // partially edited document still refreshed
    }

    [Fact]
    public void Every_modified_document_is_refreshed_and_original_document_restored()
    {
        var app = new App();
        using var adapter = new CadAdapter(() => app);
        var execution = adapter.Apply(new[] { Move(), new JsonObject
        {
            ["op"] = "activate_document", ["document"] = "second.dwg",
        }, Move() }, "test");
        Assert.True(execution.Ok, string.Join(";", execution.Errors));
        Assert.Same(app.Documents[0], app.ActiveDocument);
        Assert.All(app.Documents, d => Assert.Equal(1, d.Regens));
        Assert.Equal(2, ((JsonArray)execution.Readback!["displayRefresh"]!["documents"]!).Count);
    }

    [Fact]
    public void Explicit_regen_has_a_non_mutating_preview_and_no_geometry_edits()
    {
        var app = new App();
        using var adapter = new CadAdapter(() => app);
        var ops = new[] { new JsonObject { ["op"] = "regen_document" } };
        Assert.Empty(adapter.Preview(ops).Errors);
        Assert.Equal(0, app.ActiveDocument.Regens);
        Assert.True(adapter.Apply(ops, "test").Ok);
        Assert.Equal(1, app.ActiveDocument.Regens);
        Assert.Equal(0, app.ActiveDocument.Entity.Moves);
    }

    [Fact]
    public void Empty_batch_does_not_regenerate()
    {
        var app = new App();
        using var adapter = new CadAdapter(() => app);
        var execution = adapter.Apply(Array.Empty<JsonObject>(), "test");
        Assert.Equal("not-required", (string?)execution.Readback?["displayRefresh"]?["status"]);
        Assert.Equal(0, app.ActiveDocument.Regens);
    }

    [Fact]
    public void Layers_distinguish_current_on_frozen_locked_and_unknown()
    {
        var app = new App();
        using var adapter = new CadAdapter(() => app);
        var query = adapter.Read(new JsonObject { ["scope"] = "layers" });
        Assert.True(Json.GetBool(query, "ok"), query.ToJsonString());
        var layers = (JsonArray)query["layers"]!;
        Assert.Equal("BASE", (string?)query["currentLayer"]);
        Assert.True((bool)layers[0]!["current"]!);
        Assert.True((bool)layers[0]!["modelVisible"]!);
        Assert.True((bool)layers[1]!["locked"]!);
        Assert.True((bool)layers[1]!["modelVisible"]!); // locked does NOT mean hidden
        Assert.False((bool)layers[2]!["on"]!);
        Assert.False((bool)layers[2]!["modelVisible"]!);
        Assert.True((bool)layers[3]!["freeze"]!);
        Assert.False((bool)layers[3]!["modelVisible"]!);
        Assert.Null(layers[4]!["on"]);
        Assert.Null(layers[4]!["modelVisible"]);
        Assert.Contains("on", ((JsonArray)layers[4]!["unavailableProperties"]!).Select(n => (string?)n));
        Assert.Equal(0, app.ZoomCalls);
        Assert.Equal(0, app.ActiveDocument.StateWrites);
    }

    [Fact]
    public void Summary_includes_layer_states_basic_reports_omitted()
    {
        var app = new App();
        using var adapter = new CadAdapter(() => app);
        var basic = adapter.GetActiveContext(new JsonObject());
        Assert.Equal("omitted", (string?)basic.Summary["layerSummaryStatus"]);
        Assert.Equal("BASE", (string?)basic.Summary["currentLayer"]);
        Assert.Empty((JsonArray)basic.Summary["layers"]!);
        var summary = adapter.GetActiveContext(new JsonObject { ["detailLevel"] = "summary" });
        Assert.True(summary.Ok, string.Join(";", summary.Errors));
        Assert.Equal("complete", (string?)summary.Summary["layerSummaryStatus"]);
        Assert.True((bool)summary.Summary["layers"]![1]!["locked"]!);
    }

    [Fact]
    public void Entity_geometry_includes_display_state_without_editing_it()
    {
        var app = new App();
        using var adapter = new CadAdapter(() => app);
        var query = adapter.Read(new JsonObject { ["includeGeometry"] = true });
        Assert.True(Json.GetBool(query, "ok"), query.ToJsonString());
        var entity = query["entities"]![0]!;
        Assert.True((bool)entity["visible"]!);
        Assert.Equal(256, (int)entity["color"]!);
        Assert.Equal("ByLayer", (string?)entity["transparency"]);
    }

    [Fact]
    public void Dirty_text_preview_reuse_rejects_changed_target_state()
    {
        using var home = new TestHome();
        var app = new App();
        app.ActiveDocument.Saved = false;
        app.ActiveDocument.LayerCount = 4;
        using var adapter = new CadAdapter(() => app);
        var ops = new[] { Move() };
        var metadata = new JsonObject();
        adapter.CaptureSnapshot(home.Dir, metadata, ops);
        Assert.True(Json.GetBool(adapter.ValidatePreviewReuse(home.Dir, metadata, ops), "reusable"));
        app.ActiveDocument.Entity.Move(new double[3], new double[3]);
        Assert.False(Json.GetBool(adapter.ValidatePreviewReuse(home.Dir, metadata, ops), "reusable"));
    }

    [Theory]
    [InlineData("delete_entities_in_bounds")]
    [InlineData("copy_entities_between_documents")]
    [InlineData("configure_layout")]
    public void Dirty_preview_does_not_reuse_an_incomplete_scope_fingerprint(string op)
    {
        using var home = new TestHome();
        var app = new App();
        app.ActiveDocument.Saved = false;
        app.ActiveDocument.LayerCount = 4;
        using var adapter = new CadAdapter(() => app);
        var ops = new[] { new JsonObject { ["op"] = op } };
        var metadata = new JsonObject();
        adapter.CaptureSnapshot(home.Dir, metadata, ops);
        Assert.False(Json.GetBool(adapter.ValidatePreviewReuse(home.Dir, metadata, ops), "reusable"));
    }

    public sealed class App
    {
        public App()
        {
            Documents = new[] { new Document(this, "first.dwg"), new Document(this, "second.dwg") };
            ActiveDocument = Documents[0];
        }
        public long HWND => 0;
        public Document ActiveDocument { get; set; }
        public Document[] Documents { get; }
        public int ZoomCalls { get; private set; }
        public void ZoomCenter(object center, double size) => ZoomCalls++;
    }
    public sealed class Layout { public string Name => "Model"; }
    public sealed class Document
    {
        private readonly App _app;
        public Document(App app, string name) { _app = app; Name = name; }
        public string Name { get; }
        public string FullName => "C:\\drawings\\" + Name;
        public bool Saved { get; set; } = true;
        public Entity Entity { get; } = new();
        public CadContextOptimizationTests.CountingCollection<Entity> ModelSpace => new(1, _ => Entity);
        public int LayerCount { get; set; } = 5;
        public CadContextOptimizationTests.CountingCollection<object> Layers => new(LayerCount, i => i == 4 ?
            new UnknownLayer() : new Layer { Name = i == 0 ? "BASE" : $"L-{i}", LayerOn = i != 2, Freeze = i == 3, Lock = i == 1 });
        public Layer ActiveLayer => new() { Name = "BASE" };
        public int StateWrites { get; private set; }
        public Layout ActiveLayout { get => new(); set => StateWrites++; }
        public int ActiveSpace { get => 1; set => StateWrites++; }
        public bool MSpace { get => false; set => StateWrites++; }
        public object GetVariable(string name) => name switch
        {
            "VIEWCTR" => new[] { 0.0, 0.0 }, "VIEWSIZE" => 100.0, "INSUNITS" => 4, _ => 0,
        };
        public Entity HandleToObject(string handle) => handle == "rejected" ?
            throw new COMException("busy", unchecked((int)0x80010001)) : Entity;
        public void Activate() => _app.ActiveDocument = this;
        public bool RegenFails { get; set; }
        public int Regens { get; private set; }
        public int LastRegenType { get; private set; }
        public void Regen(int type)
        {
            Regens++; LastRegenType = type;
            if (RegenFails) throw new COMException("busy", unchecked((int)0x80010001));
        }
    }
    public sealed class Layer
    {
        public string Name { get; init; } = "BASE";
        public bool LayerOn { get; init; } = true;
        public bool Freeze { get; init; }
        public bool Lock { get; init; }
        public bool Plottable => true;
        public int Color => 7;
        public string Linetype => "Continuous";
    }
    public sealed class UnknownLayer { public string Name => "UNKNOWN"; }
    public sealed class Entity
    {
        public string EntityName => "AcDbText";
        public string Layer => "BASE";
        public string Handle => "1";
        public string TextString => "PROBE";
        public bool Visible => true;
        public int Color => 256;
        public string EntityTransparency => "ByLayer";
        public double Height => 2.5;
        public double Rotation => 0;
        public object InsertionPoint => new[] { (double)Moves, 0.0, 0.0 };
        public int Moves { get; private set; }
        public void Move(object from, object to) => Moves++;
        public void GetBoundingBox(out object min, out object max)
        {
            min = new[] { (double)Moves, 0.0, 0.0 }; max = new[] { Moves + 5.0, 2.5, 0.0 };
        }
    }
}
