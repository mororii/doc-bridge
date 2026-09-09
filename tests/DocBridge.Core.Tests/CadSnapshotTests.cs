using System.Collections;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class CadSnapshotTests
{
    [Fact]
    public void Scoped_text_and_layer_capture_uses_requested_handles_beyond_500_without_modelspace_scan()
    {
        using var home = new TestHome();
        var app = SnapshotApp.Create();
        using var adapter = new CadAdapter(() => app);
        var metadata = new JsonObject();
        adapter.CaptureSnapshot(home.Dir, metadata, new[]
        {
            TextOp("1F5"),
            LayerColorOp("DB_ON"),
            new JsonObject { ["op"] = "regen_document" },
        });

        Assert.Equal(0, app.ActiveDocument.ModelSpace.EnumerationCount);
        Assert.Equal("complete", Json.GetString(metadata, "snapshotCoverage"));
        var state = ReadState(home.Dir);
        Assert.Equal("complete", Json.GetString(state, "coverage"));
        Assert.Equal("PROBE-501", Json.GetString(Json.GetObj(state, "texts")?["1F5"] as JsonObject, "text"));
        Assert.Equal(3, Json.GetInt(Json.GetObj(state, "layers")?["DB_ON"] as JsonObject, "color"));
        Assert.False(Json.GetObj(state, "layers")?["DB_ON"] is JsonObject layer && layer.ContainsKey("on"));
        Assert.Null(Json.GetObj(state, "texts")?["1"]);
    }

    [Fact]
    public void Scoped_restore_verifies_native_text_layer_and_regen()
    {
        using var home = new TestHome();
        var app = SnapshotApp.Create();
        using var adapter = new CadAdapter(() => app);
        var ops = new[] { TextOp("1"), LayerOnOp("DB_OFF"), LayerColorOp("DB_ON") };
        var metadata = new JsonObject();
        adapter.CaptureSnapshot(home.Dir, metadata, ops);

        app.ActiveDocument.EntityByHandle("1").TextString = "CHANGED";
        app.ActiveDocument.LayerByName("DB_OFF").LayerOn = true;
        app.ActiveDocument.LayerByName("DB_ON").Color = 1;

        var restored = adapter.RestoreSnapshot(home.Dir, metadata);
        Assert.True(Json.GetBool(restored, "ok"), restored.ToJsonString());
        Assert.True(Json.GetBool(Json.GetObj(restored, "readback"), "verified"));
        Assert.True(Json.GetBool(restored, "regenerated"));
        Assert.Equal(1, app.ActiveDocument.Regens);
        Assert.Equal("PROBE-1", app.ActiveDocument.EntityByHandle("1").TextString);
        Assert.False(app.ActiveDocument.LayerByName("DB_OFF").LayerOn);
        Assert.Equal(3, app.ActiveDocument.LayerByName("DB_ON").Color);
        Assert.Equal(7, app.ActiveDocument.LayerByName("DB_OFF").Color);
    }

    [Fact]
    public void Geometry_snapshot_keeps_backup_metadata_but_restore_is_incomplete()
    {
        using var home = new TestHome();
        var app = SnapshotApp.Create();
        using var adapter = new CadAdapter(() => app);
        var metadata = new JsonObject();
        var ops = new[]
        {
            new JsonObject { ["op"] = "move_entities", ["handles"] = new JsonArray("1"), ["dx"] = 10.0, ["dy"] = 5.0 },
        };
        adapter.CaptureSnapshot(home.Dir, metadata, ops);
        Assert.Equal("incomplete", Json.GetString(metadata, "snapshotCoverage"));

        app.ActiveDocument.EntityByHandle("1").Move(new double[3], new[] { 10.0, 5.0, 0.0 });
        Assert.Equal(1, app.ActiveDocument.EntityByHandle("1").Moves);

        var restored = adapter.RestoreSnapshot(home.Dir, metadata);
        Assert.False(Json.GetBool(restored, "ok"));
        Assert.False(Json.GetBool(Json.GetObj(restored, "readback"), "verified"));
        Assert.Equal("incomplete", Json.GetString(restored, "coverage"));
        Assert.Equal(1, app.ActiveDocument.EntityByHandle("1").Moves);
        Assert.Contains("partial/incomplete", string.Join(" ", Json.GetArr(restored, "warnings")!.Select(n => n?.ToString())));
        Assert.Equal("PROBE-1", app.ActiveDocument.EntityByHandle("1").TextString);
        Assert.True(Json.GetBool(restored, "partialApplied") || Json.GetInt(Json.GetObj(restored, "readback"), "checked") >= 0);
    }

    [Fact]
    public void Missing_requested_handle_fails_capture_before_a_complete_snapshot()
    {
        using var home = new TestHome();
        var app = SnapshotApp.Create();
        using var adapter = new CadAdapter(() => app);
        var ex = Assert.ThrowsAny<Exception>(() =>
            adapter.CaptureSnapshot(home.Dir, new JsonObject(), new[] { TextOp("MISSING") }));
        Assert.Contains("MISSING", ex.ToString());
        Assert.False(File.Exists(Path.Combine(home.Dir, "state.json")));
    }

    [Fact]
    public void Restore_refuses_when_active_document_identity_changed()
    {
        using var home = new TestHome();
        var app = SnapshotApp.Create();
        using var adapter = new CadAdapter(() => app);
        var metadata = new JsonObject();
        adapter.CaptureSnapshot(home.Dir, metadata, new[] { TextOp("1") });
        app.ActiveDocument.FullName = "C:\\drawings\\other.dwg";
        var restored = adapter.RestoreSnapshot(home.Dir, metadata);
        Assert.False(Json.GetBool(restored, "ok"));
        Assert.False(Json.GetBool(Json.GetObj(restored, "readback"), "verified"));
        Assert.Contains("다릅니다", string.Join(" ", Json.GetArr(restored, "errors")!.Select(n => n?.ToString())));
    }

    [Fact]
    public void Regen_only_snapshot_is_complete_and_regenerates_on_restore()
    {
        using var home = new TestHome();
        var app = SnapshotApp.Create();
        using var adapter = new CadAdapter(() => app, CadProduct.GstarCad);
        var metadata = new JsonObject();
        adapter.CaptureSnapshot(home.Dir, metadata, new[] { new JsonObject { ["op"] = "regen_document" } });
        Assert.Equal("complete", Json.GetString(metadata, "snapshotCoverage"));
        var restored = adapter.RestoreSnapshot(home.Dir, metadata);
        Assert.True(Json.GetBool(restored, "ok"), restored.ToJsonString());
        Assert.True(Json.GetBool(restored, "regenerated"));
        Assert.Equal(1, app.ActiveDocument.Regens);
        Assert.Equal(0, app.ActiveDocument.ModelSpace.EnumerationCount);
    }

    [Fact]
    public void Legacy_state_without_version_never_claims_verified_rollback()
    {
        using var home = new TestHome();
        var app = SnapshotApp.Create();
        using var adapter = new CadAdapter(() => app);
        File.WriteAllText(Path.Combine(home.Dir, "state.json"), """
            {"fullName":"C:\\drawings\\probe.dwg","layers":{"DB_ON":{"on":true,"color":3}},"texts":{"1":"PROBE-1"}}
            """);
        app.ActiveDocument.EntityByHandle("1").TextString = "CHANGED";
        app.ActiveDocument.LayerByName("DB_ON").Color = 1;
        var restored = adapter.RestoreSnapshot(home.Dir, new JsonObject { ["drawingBackup"] = "drawing-backup.dwg" });
        Assert.False(Json.GetBool(restored, "ok"));
        Assert.False(Json.GetBool(Json.GetObj(restored, "readback"), "verified"));
        Assert.Equal("PROBE-1", app.ActiveDocument.EntityByHandle("1").TextString);
        Assert.Equal(3, app.ActiveDocument.LayerByName("DB_ON").Color);
        Assert.True(Json.GetBool(restored, "partialApplied"));
    }

    [Fact]
    public void Forged_complete_empty_coverage_does_not_write_or_verify()
    {
        using var home = new TestHome();
        var app = SnapshotApp.Create();
        using var adapter = new CadAdapter(() => app);
        File.WriteAllText(Path.Combine(home.Dir, "state.json"), """
            {"version":3,"coverage":"complete","kind":"operation-scoped","complete":true,
             "fullName":"C:\\drawings\\probe.dwg","ops":["set_text_value"],
             "requested":{"layers":[],"texts":[]},"layers":{},"texts":{}}
            """);
        app.ActiveDocument.EntityByHandle("1").TextString = "CHANGED";
        var restored = adapter.RestoreSnapshot(home.Dir, new JsonObject());
        Assert.False(Json.GetBool(restored, "ok"));
        Assert.False(Json.GetBool(Json.GetObj(restored, "readback"), "verified"));
        Assert.Equal("incomplete", Json.GetString(restored, "coverage"));
        Assert.False(Json.GetBool(restored, "complete"));
        Assert.Equal(0, Json.GetInt(Json.GetObj(restored, "readback"), "checked"));
        Assert.Equal("CHANGED", app.ActiveDocument.EntityByHandle("1").TextString);
    }

    [Fact]
    public void Scoped_restore_without_stable_identity_does_not_write()
    {
        using var home = new TestHome();
        var app = SnapshotApp.Create();
        using var adapter = new CadAdapter(() => app);
        File.WriteAllText(Path.Combine(home.Dir, "state.json"), """
            {"version":3,"coverage":"complete","kind":"operation-scoped","complete":true,
             "ops":["set_text_value"],"requested":{"layers":[],"texts":["1"]},
             "layers":{},"texts":{"1":{"text":"PROBE-1","fields":["text"],"entityName":"AcDbText"}}}
            """);
        app.ActiveDocument.EntityByHandle("1").TextString = "CHANGED";
        var restored = adapter.RestoreSnapshot(home.Dir, new JsonObject());
        Assert.False(Json.GetBool(restored, "ok"));
        Assert.False(Json.GetBool(Json.GetObj(restored, "readback"), "verified"));
        Assert.Equal("incomplete", Json.GetString(restored, "coverage"));
        Assert.Equal(0, Json.GetInt(Json.GetObj(restored, "readback"), "checked"));
        Assert.Equal("CHANGED", app.ActiveDocument.EntityByHandle("1").TextString);
        Assert.Contains("identity", string.Join(" ", Json.GetArr(restored, "errors")!.Select(n => n?.ToString())),
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Scoped_restore_with_name_only_fullname_does_not_write()
    {
        using var home = new TestHome();
        var app = SnapshotApp.Create();
        using var adapter = new CadAdapter(() => app);
        File.WriteAllText(Path.Combine(home.Dir, "state.json"), """
            {"version":3,"coverage":"complete","kind":"operation-scoped","complete":true,
             "fullName":"Drawing1.dwg",
             "ops":["set_text_value"],"requested":{"layers":[],"texts":["1"]},
             "layers":{},"texts":{"1":{"text":"PROBE-1","fields":["text"],"entityName":"AcDbText"}}}
            """);
        app.ActiveDocument.EntityByHandle("1").TextString = "CHANGED";
        var restored = adapter.RestoreSnapshot(home.Dir, new JsonObject());
        Assert.False(Json.GetBool(restored, "ok"));
        Assert.False(Json.GetBool(Json.GetObj(restored, "readback"), "verified"));
        Assert.Equal("incomplete", Json.GetString(restored, "coverage"));
        Assert.Equal(0, Json.GetInt(Json.GetObj(restored, "readback"), "checked"));
        Assert.Equal("CHANGED", app.ActiveDocument.EntityByHandle("1").TextString);
        Assert.Contains("ambiguous", string.Join(" ", Json.GetArr(restored, "errors")!.Select(n => n?.ToString())),
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Scoped_restore_with_drive_relative_fullname_does_not_write()
    {
        using var home = new TestHome();
        var app = SnapshotApp.Create();
        using var adapter = new CadAdapter(() => app);
        File.WriteAllText(Path.Combine(home.Dir, "state.json"), """
            {"version":3,"coverage":"complete","kind":"operation-scoped","complete":true,
             "fullName":"C:Drawing1.dwg",
             "ops":["set_text_value"],"requested":{"layers":[],"texts":["1"]},
             "layers":{},"texts":{"1":{"text":"PROBE-1","fields":["text"],"entityName":"AcDbText"}}}
            """);
        app.ActiveDocument.EntityByHandle("1").TextString = "CHANGED";
        var restored = adapter.RestoreSnapshot(home.Dir, new JsonObject());
        Assert.False(Json.GetBool(restored, "ok"));
        Assert.False(Json.GetBool(Json.GetObj(restored, "readback"), "verified"));
        Assert.Equal("incomplete", Json.GetString(restored, "coverage"));
        Assert.Equal(0, Json.GetInt(Json.GetObj(restored, "readback"), "checked"));
        Assert.Equal("CHANGED", app.ActiveDocument.EntityByHandle("1").TextString);
        Assert.Contains("ambiguous", string.Join(" ", Json.GetArr(restored, "errors")!.Select(n => n?.ToString())),
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Scoped_capture_denies_document_selector_that_does_not_match_used_rcw()
    {
        using var home = new TestHome();
        var app = SnapshotApp.Create();
        using var adapter = new CadAdapter(() => app);
        var op = TextOp("1");
        op["document"] = "other.dwg";
        var ex = Assert.ThrowsAny<Exception>(() => adapter.CaptureSnapshot(home.Dir, new JsonObject(), new[] { op }));
        Assert.Contains("does not match", ex.ToString(), StringComparison.OrdinalIgnoreCase);
        Assert.False(File.Exists(Path.Combine(home.Dir, "state.json")));
    }

    [Fact]
    public void Scoped_selector_does_not_match_backup_path_substring()
    {
        var app = SnapshotApp.Create();
        app.ActiveDocument.FullName = @"C:\drawings\a.dwg.backup.dwg";
        using var adapter = new CadAdapter(() => app);
        var op = new JsonObject
        {
            ["op"] = "set_text_value",
            ["handle"] = "1",
            ["text"] = "SUBSTRING-HIT",
            ["document"] = @"C:\drawings\a.dwg",
        };

        var preview = adapter.Preview(new[] { op });
        Assert.Contains(preview.Errors, error => error.Contains("does not match", StringComparison.OrdinalIgnoreCase));
        var apply = adapter.Apply(new[] { op }, "test");
        Assert.False(apply.Ok);
        Assert.Equal("PROBE-1", app.ActiveDocument.EntityByHandle("1").TextString);
    }

    [Fact]
    public void Preview_and_apply_deny_scoped_ops_when_document_does_not_match_used_rcw()
    {
        var app = new CadDisplayTests.App();
        using var adapter = new CadAdapter(() => app);
        var op = new JsonObject
        {
            ["op"] = "set_text_value",
            ["handle"] = "1",
            ["text"] = "CHANGED",
            ["document"] = "second.dwg",
        };
        var preview = adapter.Preview(new[] { op });
        Assert.Contains(preview.Errors, error => error.Contains("does not match", StringComparison.OrdinalIgnoreCase));
        var apply = adapter.Apply(new[] { op }, "test");
        Assert.False(apply.Ok);
        Assert.Equal("PROBE", app.ActiveDocument.Entity.TextString);
    }

    [Fact]
    public void Execute_does_not_let_bystander_document_header_mutate_active_drawing()
    {
        using var home = new TestHome();
        var app = new CadDisplayTests.App();
        using var adapter = new CadAdapter(() => app);
        using var host = new DocBridgeHost(home.Options);
        host.Router.Register("cad", adapter);
        host.Router.Register("gstarcad", adapter);

        var bystander = Path.GetFullPath(@"C:\drawings\second.dwg");
        var batch = new JsonObject
        {
            ["ops"] = new JsonArray(new JsonObject
            {
                ["op"] = "set_text_value",
                ["handle"] = "1",
                ["text"] = "BYSTANDER-HIT",
                ["document"] = bystander,
            }),
            ["executionMode"] = "execute",
            ["requestId"] = Guid.NewGuid().ToString("D"),
            ["expectedDocumentRef"] = bystander,
        };

        var denied = host.ApplyOps("cad", batch);
        Assert.False(Json.GetBool(denied, "ok"), denied.ToJsonString());
        Assert.Contains(Json.GetArr(denied, "errors")!, e =>
            e!.GetValue<string>().Contains("active drawing", StringComparison.OrdinalIgnoreCase) ||
            e!.GetValue<string>().Contains("does not match current document", StringComparison.OrdinalIgnoreCase));
        Assert.Equal("PROBE", app.ActiveDocument.Entity.TextString);
        Assert.Same(app.Documents[0], app.ActiveDocument);
        Assert.Equal("first.dwg", app.ActiveDocument.Name);
    }

    [Fact]
    public void Matching_document_selector_is_allowed_for_scoped_preview()
    {
        var app = new CadDisplayTests.App();
        using var adapter = new CadAdapter(() => app);
        var preview = adapter.Preview(new[]
        {
            new JsonObject
            {
                ["op"] = "set_text_value",
                ["handle"] = "1",
                ["text"] = "CHANGED",
                ["document"] = "first.dwg",
            },
        });
        Assert.Empty(preview.Errors);
    }

    private static JsonObject TextOp(string handle) =>
        new() { ["op"] = "set_text_value", ["handle"] = handle, ["text"] = "NEW" };

    private static JsonObject LayerOnOp(string layer) =>
        new() { ["op"] = "set_layer_visibility", ["layer"] = layer, ["visible"] = true };

    private static JsonObject LayerColorOp(string layer) =>
        new() { ["op"] = "set_layer_color", ["layer"] = layer, ["color"] = 1 };

    private static JsonObject ReadState(string dir) =>
        JsonNode.Parse(File.ReadAllText(Path.Combine(dir, "state.json"))) as JsonObject
        ?? throw new InvalidOperationException("state.json missing");

    public sealed class SnapshotApp
    {
        public static SnapshotApp Create()
        {
            var document = new SnapshotDocument();
            return new SnapshotApp { ActiveDocument = document, Documents = new[] { document } };
        }

        public long HWND => 0;
        public string Version => "24.0";
        public required SnapshotDocument ActiveDocument { get; set; }
        public required SnapshotDocument[] Documents { get; init; }
    }

    public sealed class SnapshotDocument
    {
        public SnapshotDocument()
        {
            Layers = new NamedLayers(
                new SnapshotLayer { Name = "DB_ON", LayerOn = true, Color = 3 },
                new SnapshotLayer { Name = "DB_OFF", LayerOn = false, Color = 7 });
            Entities = new Dictionary<string, SnapshotEntity>(StringComparer.OrdinalIgnoreCase)
            {
                ["1"] = new() { Handle = "1", TextString = "PROBE-1" },
                ["1F5"] = new() { Handle = "1F5", TextString = "PROBE-501" },
            };
            ModelSpace = new CountingModelSpace(Entities.Values);
        }

        public string Name => "probe.dwg";
        public string FullName { get; set; } = "C:\\drawings\\probe.dwg";
        public bool Saved { get; set; } = true;
        public NamedLayers Layers { get; }
        public CountingModelSpace ModelSpace { get; }
        public Dictionary<string, SnapshotEntity> Entities { get; }
        public int Regens { get; private set; }

        public SnapshotEntity HandleToObject(string handle) =>
            Entities.TryGetValue(handle, out var entity)
                ? entity
                : throw new InvalidOperationException($"handle not found: {handle}");

        public SnapshotLayer LayerByName(string name) => Layers.Item(name);
        public SnapshotEntity EntityByHandle(string handle) => HandleToObject(handle);
        public void Regen(int type) => Regens++;
    }

    public sealed class NamedLayers : IEnumerable<SnapshotLayer>
    {
        private readonly Dictionary<string, SnapshotLayer> _layers;

        public NamedLayers(params SnapshotLayer[] layers) =>
            _layers = layers.ToDictionary(layer => layer.Name, StringComparer.OrdinalIgnoreCase);

        public int Count => _layers.Count;
        public SnapshotLayer Item(string name) => _layers[name];
        public IEnumerator<SnapshotLayer> GetEnumerator() => _layers.Values.GetEnumerator();
        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    public sealed class CountingModelSpace : IEnumerable<SnapshotEntity>
    {
        private readonly IReadOnlyCollection<SnapshotEntity> _entities;
        public CountingModelSpace(IReadOnlyCollection<SnapshotEntity> entities) => _entities = entities;
        public int Count => _entities.Count;
        public int EnumerationCount { get; private set; }

        public SnapshotEntity Item(int index)
        {
            var i = 0;
            foreach (var entity in _entities)
            {
                if (i++ == index) return entity;
            }
            throw new ArgumentOutOfRangeException(nameof(index));
        }

        public IEnumerator<SnapshotEntity> GetEnumerator()
        {
            foreach (var entity in _entities)
            {
                EnumerationCount++;
                yield return entity;
            }
        }

        IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();
    }

    public sealed class SnapshotLayer
    {
        public string Name { get; init; } = "0";
        public bool LayerOn { get; set; } = true;
        public int Color { get; set; } = 7;
    }

    public sealed class SnapshotEntity
    {
        public string EntityName => "AcDbText";
        public string Layer => "DB_ON";
        public string Handle { get; init; } = "1";
        public string TextString { get; set; } = "PROBE";
        public int Moves { get; private set; }
        public void Move(object from, object to) => Moves++;
    }
}
