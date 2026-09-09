using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class CadDisplayE2ETests
{
    private readonly Xunit.Abstractions.ITestOutputHelper _output;
    public CadDisplayE2ETests(Xunit.Abstractions.ITestOutputHelper output) => _output = output;
    [Theory]
    [InlineData(CadProduct.AutoCad, "DOCBRIDGE_CAD_DISPLAY_E2E")]
    [InlineData(CadProduct.GstarCad, "DOCBRIDGE_GSTARCAD_E2E")]
    public void Isolated_live_document_refresh_and_layer_readback(CadProduct product, string enableVariable)
    {
        if (Environment.GetEnvironmentVariable(enableVariable) != "1")
        {
            _output.WriteLine("LIVE TEST NOT ENABLED");
            return;
        }
        var appKey = product == CadProduct.GstarCad ? "gstarcad" : "cad";
        var progId = product == CadProduct.GstarCad ? "Gcad.Application" : "AutoCAD.Application";
        var template = product == CadProduct.GstarCad ? "gcadiso.dwt" : "acad.dwt";
        _output.WriteLine($"LIVE TEST ENABLED: isolated {appKey} probe");
        using var home = new TestHome();
        using var adapter = new CadAdapter(product: product);
        using var host = new DocBridgeHost(home.Options);
        host.Router.Register(appKey, adapter);
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
                var foreground = new ForegroundInteractionGuard(appKey);
                try
                {
                dynamic app = RotHelper.GetActiveObject(progId)
                    ?? throw new InvalidOperationException($"Live test requires an already running {appKey}; no application will be launched.");
                application = app;
                foreground.TrackTargetWindow(Convert.ToInt64(app.HWND));
                original = app.ActiveDocument;
                userState = UserState(app);
                _output.WriteLine("SETUP: application ready, creating isolated probe");
                dynamic doc = app.Documents.Add(template);
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

            var layers = host.Read(appKey, new JsonObject { ["scope"] = "layers", ["startsWith"] = "DB_" });
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
            JsonObject Apply(JsonArray batch, bool highRisk = false)
            {
                var elapsed = System.Diagnostics.Stopwatch.StartNew();
                var dry = host.ApplyOps(appKey, new JsonObject { ["dryRun"] = true, ["ops"] = batch.DeepClone() });
                Assert.True(Json.GetBool(dry, "ok"), dry.ToJsonString());
                var applied = host.ApplyOps(appKey, new JsonObject
                {
                    ["dryRun"] = false, ["ops"] = batch.DeepClone(), ["confirmToken"] = dry["confirmToken"]!.DeepClone(),
                    ["highRiskConfirm"] = highRisk,
                });
                Assert.True(Json.GetBool(applied, "ok"), applied.ToJsonString());
                Assert.True((bool)applied["readback"]!["verified"]!);
                var saveOnly = batch.All(n => (string?)n?["op"] == "save_document");
                Assert.Equal(saveOnly ? "not-required" : "completed", (string?)applied["readback"]?["displayRefresh"]?["status"]);
                if (!saveOnly) Assert.Single((JsonArray)applied["readback"]!["displayRefresh"]!["documents"]!);
                _output.WriteLine($"BATCH: {batch[0]?["op"]}, {batch.Count} ops, dry-run+apply {elapsed.ElapsedMilliseconds} ms");
                return applied;
            }
            Apply(ops);
            var query = host.Read(appKey, new JsonObject { ["entityType"] = "Text", ["includeGeometry"] = true });
            var entity = ((JsonArray)query["entities"]!).Single(n => (string?)n?["handle"] == handle)!;
            Assert.Equal("DISPLAY PROBE", (string?)entity["text"]);
            Assert.Equal(5.0, (double)entity["height"]!, 5);
            Assert.Equal(10.0, (double)entity["insertionPoint"]![0]!, 5);
            Assert.Equal(5.0, (double)entity["insertionPoint"]![1]!, 5);
            Assert.True((bool)entity["visible"]!);
            Assert.NotNull(entity["transparency"]);
            var toggled = host.Read(appKey, new JsonObject { ["scope"] = "layers", ["contains"] = "DB_OFF" });
            Assert.True((bool)toggled["layers"]![0]!["on"]!);
            Assert.True((bool)toggled["layers"]![0]!["modelVisible"]!);
            Apply(new JsonArray(new JsonObject { ["op"] = "regen_document" }));
            var after = host.Read(appKey, new JsonObject { ["entityType"] = "Text", ["includeGeometry"] = true });
            Assert.Equal(query["entities"]!.ToJsonString(), after["entities"]!.ToJsonString());
            _output.WriteLine("PASS: live move/scale readback, automatic Regen, explicit Regen, layer current/on/off/frozen/locked states");

            // Save only our own probe before a new broad geometry scope. A dirty-document
            // text-only fingerprint must never be used to authorize creation/deletion.
            void SaveProbe() => adapter.RunOnAdapterThread(() => { ((dynamic)probe!).Save(); return true; });
            SaveProbe();
            Apply(new JsonArray(JsonNode.Parse("""
                {"op":"draw_entities","entities":[
                  {"type":"line","start":[0,20,0],"end":[40,20,0],"layer":"DB_ON","color":{"aci":1}},
                  {"type":"circle","center":[60,20],"radius":5,"color":{"aci":3}},
                  {"type":"lwpolyline","points":[[0,40],[20,40],[20,50],[0,50]],"closed":true},
                  {"type":"text","point":[0,60],"height":2.5,"text":"GSTAR 연결 검증 − ㎜"},
                  {"type":"mtext","point":[0,80],"width":50,"height":2.5,"text":"FIRST\\PSECOND"}
                ]}
                """)));
            var drawn = host.Read(appKey, new JsonObject { ["includeGeometry"] = true });
            Assert.True(Json.GetBool(drawn, "ok"), drawn.ToJsonString());
            var entities = (JsonArray)drawn["entities"]!;
            Assert.Equal(6, entities.Count);
            Assert.Contains(entities, n => (string?)n?["text"] == "GSTAR 연결 검증 − ㎜");
            var line = entities.Single(n => (string?)n?["type"] == "AcDbLine")!;
            var circle = entities.Single(n => (string?)n?["type"] == "AcDbCircle")!;
            Assert.Equal(1, (int)line["color"]!);
            Assert.Equal(5, (double)circle["radius"]!, 5);
            SaveProbe();
            Apply(new JsonArray
            {
                new JsonObject { ["op"] = "copy_entities", ["handles"] = new JsonArray((string?)line["handle"]), ["dx"] = 0, ["dy"] = 10 },
                new JsonObject { ["op"] = "mirror_entities", ["handles"] = new JsonArray((string?)line["handle"]), ["axisStart"] = new JsonArray(0, 0, 0), ["axisEnd"] = new JsonArray(0, 100, 0) },
                new JsonObject { ["op"] = "offset_entities", ["handles"] = new JsonArray((string?)circle["handle"]), ["distance"] = 2 },
                new JsonObject { ["op"] = "set_text_value", ["handle"] = handle, ["text"] = "EDITED 한글 − ㎜" },
                new JsonObject { ["op"] = "set_layer_color", ["layer"] = "DB_ON", ["color"] = 2 },
            });
            var changed = host.Read(appKey, new JsonObject { ["includeGeometry"] = true });
            Assert.Equal(9, ((JsonArray)changed["entities"]!).Count);
            Assert.Contains((JsonArray)changed["entities"]!, n => (string?)n?["text"] == "EDITED 한글 − ㎜");
            SaveProbe();
            var savedFile = Path.Combine(home.Dir, "verified-copy.dwg");
            Apply(new JsonArray(new JsonObject { ["op"] = "save_document", ["output"] = savedFile }), highRisk: true);
            Assert.True(new FileInfo(savedFile).Length > 0);
            var selected = host.Read(appKey, new JsonObject { ["document"] = savedFile, ["entityType"] = "Circle", ["includeGeometry"] = true });
            Assert.Equal(2, ((JsonArray)selected["entities"]!).Count);
            _output.WriteLine("PASS: line/circle/polyline/Unicode text/MText creation, ACI, same-drawing copy/mirror/offset, text edit, layer color, DWG SaveAs and explicit document readback");
        }
        finally
        {
            adapter.RunOnAdapterThread(() =>
            {
                var foreground = new ForegroundInteractionGuard(appKey);
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
