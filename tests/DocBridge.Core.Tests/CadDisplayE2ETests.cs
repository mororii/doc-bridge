using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class CadDisplayE2ETests
{
    private readonly Xunit.Abstractions.ITestOutputHelper _output;
    public CadDisplayE2ETests(Xunit.Abstractions.ITestOutputHelper output) => _output = output;
    [Fact]
    public void Isolated_live_document_refresh_and_layer_readback()
    {
        if (Environment.GetEnvironmentVariable("DOCBRIDGE_CAD_DISPLAY_E2E") != "1")
        {
            _output.WriteLine("LIVE TEST NOT ENABLED");
            return;
        }
        _output.WriteLine("LIVE TEST ENABLED: isolated AutoCAD probe");
        using var home = new TestHome();
        using var adapter = new CadAdapter(() => RotHelper.GetActiveObject("AutoCAD.Application"));
        using var host = new DocBridgeHost(home.Options);
        host.Router.Register("cad", adapter);
        object? probe = null;
        object? original = null;
        object? application = null;
        string handle = "";
        string? userState = null;
        // Setup/teardown run on the adapter STA. Only this test's new probe may be closed.
        string UserState(dynamic app)
        {
            var states = new List<string>();
            foreach (dynamic doc in app.Documents)
            {
                if (ReferenceEquals((object)doc, probe)) continue;
                states.Add($"{doc.FullName}|{doc.Saved}|{doc.ModelSpace.Count}");
            }
            return string.Join("\n", states.OrderBy(s => s));
        }
        try
        {
            adapter.RunOnAdapterThread(() =>
            {
                var foreground = new ForegroundInteractionGuard("cad");
                try
                {
                dynamic app = RotHelper.GetActiveObject("AutoCAD.Application")
                    ?? throw new InvalidOperationException("Live test requires an already running AutoCAD; no application will be launched.");
                application = app;
                foreground.TrackTargetWindow(Convert.ToInt64(app.HWND));
                original = app.ActiveDocument;
                userState = UserState(app);
                _output.WriteLine("SETUP: application ready, creating isolated probe");
                dynamic doc = app.Documents.Add("acad.dwt");
                probe = doc;
                _output.WriteLine("SETUP: probe created");
                foreach (var name in new[] { "DB_ON", "DB_OFF", "DB_FROZEN", "DB_LOCKED" }) doc.Layers.Add(name);
                doc.Layers.Item("DB_OFF").LayerOn = false;
                doc.Layers.Item("DB_FROZEN").Freeze = true;
                doc.Layers.Item("DB_LOCKED").Lock = true;
                doc.ActiveLayer = doc.Layers.Item("DB_ON");
                dynamic text = doc.ModelSpace.AddText("DISPLAY PROBE", new[] { 0.0, 0.0, 0.0 }, 2.5);
                text.Layer = "DB_ON";
                handle = (string)text.Handle;
                doc.SaveAs(Path.Combine(home.Dir, "display-probe.dwg"));
                return true;
                }
                finally { foreground.Complete(); }
            });

            var layers = host.Read("cad", new JsonObject { ["scope"] = "layers", ["startsWith"] = "DB_" });
            Assert.True(Json.GetBool(layers, "ok"), layers.ToJsonString());
            var map = ((JsonArray)layers["layers"]!).ToDictionary(n => (string)n!["name"]!, n => n!);
            Assert.True((bool)map["DB_ON"]["current"]!);
            Assert.False((bool)map["DB_OFF"]["on"]!);
            Assert.False((bool)map["DB_OFF"]["modelVisible"]!);
            Assert.True((bool)map["DB_FROZEN"]["freeze"]!);
            Assert.False((bool)map["DB_FROZEN"]["modelVisible"]!);
            Assert.True((bool)map["DB_LOCKED"]["locked"]!);
            Assert.True((bool)map["DB_LOCKED"]["modelVisible"]!);

            var ops = new JsonArray
            {
                new JsonObject { ["op"] = "move_entities", ["handles"] = new JsonArray(handle), ["dx"] = 10.0, ["dy"] = 5.0 },
                new JsonObject { ["op"] = "scale_entities", ["handles"] = new JsonArray(handle), ["factor"] = 2.0, ["basePoint"] = new JsonArray(10.0, 5.0, 0.0) },
                new JsonObject { ["op"] = "set_layer_visibility", ["layer"] = "DB_OFF", ["visible"] = true },
            };
            JsonObject Apply(JsonArray batch)
            {
                var dry = host.ApplyOps("cad", new JsonObject { ["dryRun"] = true, ["ops"] = batch.DeepClone() });
                Assert.True(Json.GetBool(dry, "ok"), dry.ToJsonString());
                var applied = host.ApplyOps("cad", new JsonObject
                {
                    ["dryRun"] = false, ["ops"] = batch.DeepClone(), ["confirmToken"] = dry["confirmToken"]!.DeepClone(),
                });
                Assert.True(Json.GetBool(applied, "ok"), applied.ToJsonString());
                Assert.Equal("completed", (string?)applied["readback"]?["displayRefresh"]?["status"]);
                Assert.Single((JsonArray)applied["readback"]!["displayRefresh"]!["documents"]!);
                return applied;
            }
            Apply(ops);
            var query = host.Read("cad", new JsonObject { ["entityType"] = "Text", ["includeGeometry"] = true });
            var entity = ((JsonArray)query["entities"]!).Single(n => (string?)n?["handle"] == handle)!;
            Assert.Equal("DISPLAY PROBE", (string?)entity["text"]);
            Assert.Equal(5.0, (double)entity["height"]!, 5);
            Assert.Equal(10.0, (double)entity["insertionPoint"]![0]!, 5);
            Assert.Equal(5.0, (double)entity["insertionPoint"]![1]!, 5);
            Assert.True((bool)entity["visible"]!);
            Assert.NotNull(entity["transparency"]);
            var toggled = host.Read("cad", new JsonObject { ["scope"] = "layers", ["contains"] = "DB_OFF" });
            Assert.True((bool)toggled["layers"]![0]!["on"]!);
            Assert.True((bool)toggled["layers"]![0]!["modelVisible"]!);
            Apply(new JsonArray(new JsonObject { ["op"] = "regen_document" }));
            var after = host.Read("cad", new JsonObject { ["entityType"] = "Text", ["includeGeometry"] = true });
            Assert.Equal(query["entities"]!.ToJsonString(), after["entities"]!.ToJsonString());
            _output.WriteLine("PASS: live move/scale readback, automatic Regen, explicit Regen, layer current/on/off/frozen/locked states");
        }
        finally
        {
            adapter.RunOnAdapterThread(() =>
            {
                var foreground = new ForegroundInteractionGuard("cad");
                try
                {
                if (probe is null)
                {
                    _output.WriteLine("CLEANUP: no probe was created; no user documents changed by setup");
                    return true;
                }
                if (application is not null)
                {
                    try { foreground.TrackTargetWindow(Convert.ToInt64(((dynamic)application).HWND)); }
                    catch { /* Do not prevent own-probe cleanup just because HWND is unavailable. */ }
                }
                if (original is not null) ((dynamic)original).Activate();
                if (probe is not null) ((dynamic)probe).Close(false);
                probe = null;
                if (application is not null) Assert.Equal(userState, UserState((dynamic)application));
                _output.WriteLine("PASS: probe closed; existing document paths, saved flags and entity counts unchanged");
                return true;
                }
                finally { foreground.Complete(); }
            });
        }
    }
}
