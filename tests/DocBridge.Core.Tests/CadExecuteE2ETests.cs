using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

/// <summary>
/// Guarded AutoCAD/GstarCAD execute fixture. Opt-in only.
/// DOCBRIDGE_CAD_EXECUTE_E2E=1 or DOCBRIDGE_GSTARCAD_EXECUTE_E2E=1.
/// No launch. Owns a temporary drawing only. Does not overwrite external files.
/// </summary>
public class CadExecuteE2ETests
{
    private readonly Xunit.Abstractions.ITestOutputHelper _output;
    public CadExecuteE2ETests(Xunit.Abstractions.ITestOutputHelper output) => _output = output;

    [Theory]
    [InlineData(CadProduct.AutoCad, "DOCBRIDGE_CAD_EXECUTE_E2E")]
    [InlineData(CadProduct.GstarCad, "DOCBRIDGE_GSTARCAD_EXECUTE_E2E")]
    public void Isolated_execute_text_layer_regen_and_geometry_policy(CadProduct product, string enableVariable)
    {
        if (Environment.GetEnvironmentVariable(enableVariable) != "1")
        {
            _output.WriteLine("LIVE TEST NOT ENABLED");
            return;
        }

        var appKey = product == CadProduct.GstarCad ? "gstarcad" : "cad";
        var progId = product == CadProduct.GstarCad ? "Gcad.Application" : "AutoCAD.Application";
        var template = product == CadProduct.GstarCad ? "gcadiso.dwt" : "acad.dwt";
        _output.WriteLine($"LIVE TEST ENABLED: isolated {appKey} execute probe");
        using var home = new TestHome();
        using var adapter = new CadAdapter(product: product);
        var executeAdapter = new CadInjectedApplyFaultAdapter(adapter);
        using var host = new DocBridgeHost(home.Options);
        host.Router.Register(appKey, executeAdapter);
        object? probe = null;
        object? original = null;
        object? application = null;
        string handle = "";
        string? userState = null;
        double startX = 0;
        double startY = 0;

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
                    dynamic doc = app.Documents.Add(template);
                    probe = doc;
                    foreach (var name in new[] { "DB_ON", "DB_OFF" }) doc.Layers.Add(name);
                    doc.Layers.Item("DB_OFF").LayerOn = false;
                    doc.ActiveLayer = doc.Layers.Item("DB_ON");
                    dynamic text = doc.ModelSpace.AddText("EXECUTE PROBE", new[] { 0.0, 0.0, 0.0 }, 2.5);
                    text.Layer = "DB_ON";
                    handle = (string)text.Handle;
                    doc.SaveAs(Path.Combine(home.Dir, "execute-probe.dwg"));
                    return true;
                }
                finally { foreground.Complete(); }
            });

            var ctx = host.GetActiveContext(appKey);
            Assert.True(Json.GetBool(ctx, "ok"), ctx.ToJsonString());
            var documentRef = Json.GetString(ctx, "documentRef");
            Assert.False(string.IsNullOrWhiteSpace(documentRef));

            JsonObject Execute(JsonArray ops, string? requestId = null)
            {
                var batch = new JsonObject
                {
                    ["executionMode"] = "execute",
                    ["requestId"] = requestId ?? Guid.NewGuid().ToString("D"),
                    ["expectedDocumentRef"] = documentRef,
                    ["ops"] = ops.DeepClone(),
                };
                return host.ApplyOps(appKey, batch);
            }

            void RequireExecute(JsonObject result, string step)
            {
                if (Json.GetBool(result, "dryRun") ||
                    JoinErrors(result).Contains("requires confirmToken", StringComparison.OrdinalIgnoreCase))
                    Assert.Fail($"{step}: executionMode execute is required; host treated the batch as legacy/token. {result.ToJsonString()}");
            }

            JsonObject ReadText() =>
                host.Read(appKey, new JsonObject { ["entityType"] = "Text", ["includeGeometry"] = true });

            JsonObject TextEntity(JsonObject query) =>
                ((JsonArray)query["entities"]!).Single(n => (string?)n?["handle"] == handle)!.AsObject();

            var before = ReadText();
            var beforeEntity = TextEntity(before);
            startX = (double)beforeEntity["insertionPoint"]![0]!;
            startY = (double)beforeEntity["insertionPoint"]![1]!;
            Assert.Equal("EXECUTE PROBE", (string?)beforeEntity["text"]);

            var textId = Guid.NewGuid().ToString("D");
            var textExec = Execute(new JsonArray(new JsonObject
            {
                ["op"] = "set_text_value", ["handle"] = handle, ["text"] = "EXECUTE TEXT",
                ["document"] = documentRef,
            }), textId);
            RequireExecute(textExec, "set_text_value");
            Assert.True(Json.GetBool(textExec, "ok"), textExec.ToJsonString());
            Assert.Equal("EXECUTE TEXT", (string?)TextEntity(ReadText())["text"]);
            _output.WriteLine("PASS: first set_text_value wrote native text");

            executeAdapter.ArmFaultAfterFirstSuccessfulOp();
            var fault = Execute(new JsonArray(
                new JsonObject { ["op"] = "set_text_value", ["handle"] = handle, ["text"] = "SHOULD-ROLLBACK", ["document"] = documentRef },
                new JsonObject { ["op"] = "set_layer_color", ["layer"] = "DB_ON", ["color"] = 3, ["document"] = documentRef }));
            RequireExecute(fault, "injected-apply-fault");
            Assert.False(Json.GetBool(fault, "ok"), fault.ToJsonString());
            Assert.True(executeAdapter.FirstSuccessfulOpApplied, "injected fault must run after a real first Apply write, not a preview deny");
            Assert.Equal("SHOULD-ROLLBACK", executeAdapter.ObservedNativeTextAfterFirstWrite);
            var afterFault = TextEntity(ReadText());
            var rollback = Json.GetObj(fault, "rollback");
            Assert.True(
                string.Equals("EXECUTE TEXT", (string?)afterFault["text"], StringComparison.Ordinal),
                $"native text must be restored; rollback.verified={Json.GetBool(rollback, "verified")} is not sufficient. {fault.ToJsonString()}");
            _output.WriteLine("PASS: injected Apply fault after first write; native rollback, not rollback boolean alone");

            var bystanderRef = Path.GetFullPath(Path.Combine(home.Dir, "bystander-other.dwg"));
            var bystander = host.ApplyOps(appKey, new JsonObject
            {
                ["executionMode"] = "execute",
                ["requestId"] = Guid.NewGuid().ToString("D"),
                ["expectedDocumentRef"] = bystanderRef,
                ["ops"] = new JsonArray(new JsonObject
                {
                    ["op"] = "set_text_value",
                    ["handle"] = handle,
                    ["text"] = "BYSTANDER-HIT",
                    ["document"] = bystanderRef,
                }),
            });
            RequireExecute(bystander, "bystander");
            Assert.False(Json.GetBool(bystander, "ok"), bystander.ToJsonString());
            Assert.Equal("EXECUTE TEXT", (string?)TextEntity(ReadText())["text"]);
            _output.WriteLine("PASS: bystander document header did not mutate owned native text");

            adapter.RunOnAdapterThread(() =>
            {
                dynamic entity = ((dynamic)probe!).HandleToObject(handle);
                entity.TextString = "CHALLENGE-MUTATED";
                return true;
            });
            Assert.Equal("CHALLENGE-MUTATED", (string?)TextEntity(ReadText())["text"]);

            var replay = Execute(new JsonArray(new JsonObject
            {
                ["op"] = "set_text_value", ["handle"] = handle, ["text"] = "EXECUTE TEXT",
                ["document"] = documentRef,
            }), textId);
            RequireExecute(replay, "set_text_value replay");
            Assert.True(Json.GetBool(replay, "ok"), replay.ToJsonString());
            Assert.True(Json.GetBool(replay, "idempotentReplay"), replay.ToJsonString());
            Assert.Equal("CHALLENGE-MUTATED", (string?)TextEntity(ReadText())["text"]);
            _output.WriteLine("PASS: intermediate replay challenge left native mutation; idempotentReplay=true");

            var layerExec = Execute(new JsonArray(
                new JsonObject { ["op"] = "set_layer_visibility", ["layer"] = "DB_OFF", ["visible"] = true, ["document"] = documentRef },
                new JsonObject { ["op"] = "set_layer_color", ["layer"] = "DB_ON", ["color"] = 2, ["document"] = documentRef },
                new JsonObject { ["op"] = "regen_document", ["document"] = documentRef }));
            RequireExecute(layerExec, "layer/regen");
            Assert.True(Json.GetBool(layerExec, "ok"), layerExec.ToJsonString());
            var layers = host.Read(appKey, new JsonObject { ["scope"] = "layers", ["startsWith"] = "DB_" });
            var map = ((JsonArray)layers["layers"]!).ToDictionary(n => (string)n!["name"]!, n => n!);
            Assert.True((bool)map["DB_OFF"]["on"]!);
            Assert.Equal(2, (int)map["DB_ON"]["color"]!);
            _output.WriteLine("PASS: execute set_layer_visibility/color and regen_document");

            var moveExec = Execute(new JsonArray(new JsonObject
            {
                ["op"] = "move_entities", ["handles"] = new JsonArray(handle), ["dx"] = 10.0, ["dy"] = 5.0,
            }));
            RequireExecute(moveExec, "geometry execute denial");
            Assert.False(Json.GetBool(moveExec, "ok"));
            var afterDenied = TextEntity(ReadText());
            Assert.Equal(startX, (double)afterDenied["insertionPoint"]![0]!, 5);
            Assert.Equal(startY, (double)afterDenied["insertionPoint"]![1]!, 5);
            Assert.Equal("CHALLENGE-MUTATED", (string?)afterDenied["text"]);
            _output.WriteLine("PASS: geometry execute denied with no native move");

            JsonObject LegacyApply(JsonArray batch, bool highRisk = false)
            {
                var dry = host.ApplyOps(appKey, new JsonObject { ["dryRun"] = true, ["ops"] = batch.DeepClone() });
                Assert.True(Json.GetBool(dry, "ok"), dry.ToJsonString());
                var applied = host.ApplyOps(appKey, new JsonObject
                {
                    ["dryRun"] = false,
                    ["ops"] = batch.DeepClone(),
                    ["confirmToken"] = dry["confirmToken"]!.DeepClone(),
                    ["highRiskConfirm"] = highRisk,
                });
                Assert.True(Json.GetBool(applied, "ok"), applied.ToJsonString());
                return applied;
            }

            LegacyApply(new JsonArray(new JsonObject
            {
                ["op"] = "move_entities", ["handles"] = new JsonArray(handle), ["dx"] = 10.0, ["dy"] = 5.0,
            }));
            var afterLegacyMove = TextEntity(ReadText());
            Assert.Equal(startX + 10.0, (double)afterLegacyMove["insertionPoint"]![0]!, 5);
            Assert.Equal(startY + 5.0, (double)afterLegacyMove["insertionPoint"]![1]!, 5);
            _output.WriteLine("PASS: geometry legacy dry-run/token path unchanged");

            var highRiskExec = Execute(new JsonArray(new JsonObject
            {
                ["op"] = "save_document", ["output"] = Path.Combine(home.Dir, "execute-rejected.dwg"),
            }));
            Assert.False(Json.GetBool(highRiskExec, "ok"));
            Assert.False(File.Exists(Path.Combine(home.Dir, "execute-rejected.dwg")));

            var dirtyDenied = Path.Combine(home.Dir, "execute-dirty-denied.dwg");
            var dirtySave = new JsonArray(new JsonObject { ["op"] = "save_document", ["output"] = dirtyDenied });
            var dirtyDry = host.ApplyOps(appKey, new JsonObject { ["dryRun"] = true, ["ops"] = dirtySave.DeepClone() });
            Assert.True(Json.GetBool(dirtyDry, "ok"), dirtyDry.ToJsonString());
            var dirtyApplied = host.ApplyOps(appKey, new JsonObject
            {
                ["dryRun"] = false,
                ["ops"] = dirtySave.DeepClone(),
                ["confirmToken"] = dirtyDry["confirmToken"]!.DeepClone(),
                ["highRiskConfirm"] = true,
            });
            Assert.False(Json.GetBool(dirtyApplied, "ok"), dirtyApplied.ToJsonString());
            Assert.Contains(
                "dirty AutoCAD operation cannot be safely fingerprinted",
                JoinErrors(dirtyApplied),
                StringComparison.OrdinalIgnoreCase);
            Assert.Equal("cad-operation-state-sha256", (string?)dirtyApplied["timings"]?["fingerprintMethod"]);
            Assert.False(File.Exists(dirtyDenied));
            _output.WriteLine("PASS: dirty legacy SaveAs denied by fingerprint guard; no output");

            var ownedHome = Path.GetFullPath(home.Dir);
            adapter.RunOnAdapterThread(() =>
            {
                var foreground = new ForegroundInteractionGuard(appKey);
                try
                {
                    if (application is not null)
                    {
                        try { foreground.TrackTargetWindow(Convert.ToInt64(((dynamic)application).HWND)); }
                        catch { /* Do not prevent own-probe Save just because HWND is unavailable. */ }
                    }
                    dynamic owned = (dynamic)probe!;
                    var fullName = Path.GetFullPath((string)owned.FullName);
                    Assert.True(
                        fullName.StartsWith(ownedHome + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
                        && fullName.EndsWith("execute-probe.dwg", StringComparison.OrdinalIgnoreCase),
                        $"RCW Save is allowed only on the owned probe, FullName={fullName}");
                    Assert.False((bool)owned.Saved, "owned probe must still be dirty before the RCW Save that establishes a clean baseline");
                    owned.Save();
                    Assert.True((bool)owned.Saved, "owned probe RCW Save must leave a clean saved baseline");
                    if (application is not null)
                        Assert.Equal(userState, UserState((dynamic)application));
                    return true;
                }
                finally { foreground.Complete(); }
            });

            var saved = Path.Combine(home.Dir, "execute-legacy-save.dwg");
            LegacyApply(new JsonArray(new JsonObject { ["op"] = "save_document", ["output"] = saved }), highRisk: true);
            Assert.True(File.Exists(saved) && new FileInfo(saved).Length > 0);
            Assert.False(File.Exists(dirtyDenied));
            _output.WriteLine("PASS: high-risk execute rejected; dirty token SaveAs denied; clean owned baseline then legacy token+highRiskConfirm SaveAs");
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

    private static string JoinErrors(JsonObject result)
    {
        var parts = new List<string>();
        void Walk(JsonNode? node)
        {
            if (node is JsonObject obj)
            {
                if (Json.GetArr(obj, "errors") is { } arr)
                    foreach (var item in arr)
                        if (item is JsonValue v && v.TryGetValue<string>(out var s) && !string.IsNullOrWhiteSpace(s))
                            parts.Add(s);
                foreach (var kv in obj)
                    if (kv.Key != "errors") Walk(kv.Value);
            }
            else if (node is JsonArray array)
            {
                foreach (var item in array) Walk(item);
            }
        }
        Walk(result);
        return string.Join(" | ", parts);
    }

    [Fact]
    public void Injected_apply_fault_after_first_text_write_restores_native_text()
    {
        using var home = new TestHome();
        var app = CadSnapshotTests.SnapshotApp.Create();
        using var cad = new CadAdapter(() => app);
        var faulting = new CadInjectedApplyFaultAdapter(cad);
        using var host = new DocBridgeHost(home.Options);
        host.Router.Register("cad", faulting);

        var documentRef = app.ActiveDocument.FullName;
        JsonObject Execute(string text, bool includeFollowOnOp)
        {
            var ops = new JsonArray(new JsonObject
            {
                ["op"] = "set_text_value",
                ["handle"] = "1",
                ["text"] = text,
                ["document"] = documentRef,
            });
            if (includeFollowOnOp)
            {
                ops.Add(new JsonObject
                {
                    ["op"] = "set_layer_color",
                    ["layer"] = "DB_ON",
                    ["color"] = 1,
                    ["document"] = documentRef,
                });
            }
            return host.ApplyOps("cad", new JsonObject
            {
                ["executionMode"] = "execute",
                ["requestId"] = Guid.NewGuid().ToString("D"),
                ["expectedDocumentRef"] = documentRef,
                ["ops"] = ops,
            });
        }

        var first = Execute("EXECUTE TEXT", includeFollowOnOp: false);
        Assert.True(Json.GetBool(first, "ok"), first.ToJsonString());
        Assert.False(Json.GetBool(first, "dryRun"));
        Assert.Equal("EXECUTE TEXT", app.ActiveDocument.EntityByHandle("1").TextString);

        faulting.ArmFaultAfterFirstSuccessfulOp();
        var fault = Execute("SHOULD-ROLLBACK", includeFollowOnOp: true);
        Assert.False(Json.GetBool(fault, "ok"), fault.ToJsonString());
        Assert.True(faulting.FirstSuccessfulOpApplied, fault.ToJsonString());
        Assert.Equal("SHOULD-ROLLBACK", faulting.ObservedNativeTextAfterFirstWrite);
        var rollback = Json.GetObj(fault, "rollback");
        Assert.True(
            string.Equals("EXECUTE TEXT", app.ActiveDocument.EntityByHandle("1").TextString, StringComparison.Ordinal),
            $"native text must be restored; rollback.verified={Json.GetBool(rollback, "verified")} is not sufficient. {fault.ToJsonString()}");
    }
}

/// <summary>
/// Forwards CAD IAppAdapter/IPreviewReuseAdapter to the real adapter. When armed,
/// Apply writes the first op through the inner adapter, reads that handle's native
/// text through inner.Read, records it, then throws so Host restore must undo the
/// write. first.Ok alone is not the changed-before-fault proof. Preview stays
/// forwarded so a valid follow-on op is not rejected before Apply.
/// </summary>
internal sealed class CadInjectedApplyFaultAdapter : IAppAdapter, IPreviewReuseAdapter
{
    private readonly CadAdapter _inner;
    private bool _faultAfterFirst;

    public CadInjectedApplyFaultAdapter(CadAdapter inner) => _inner = inner;

    public bool FirstSuccessfulOpApplied { get; private set; }
    public string? ObservedNativeTextAfterFirstWrite { get; private set; }

    public void ArmFaultAfterFirstSuccessfulOp()
    {
        _faultAfterFirst = true;
        FirstSuccessfulOpApplied = false;
        ObservedNativeTextAfterFirstWrite = null;
    }

    public string App => _inner.App;
    public AdapterStatus GetStatus() => _inner.GetStatus();
    public JsonObject GetCapabilities() => _inner.GetCapabilities();
    public ContextResult GetActiveContext() => _inner.GetActiveContext();
    public JsonObject Read(JsonObject args) => _inner.Read(args);
    public ApplyPreview Preview(IReadOnlyList<JsonObject> ops) => _inner.Preview(ops);
    public void CaptureSnapshot(string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject>? ops = null)
        => _inner.CaptureSnapshot(snapshotDir, metadata, ops);
    public JsonObject RestoreSnapshot(string snapshotDir, JsonObject metadata)
        => _inner.RestoreSnapshot(snapshotDir, metadata);
    public JsonObject ValidatePreviewReuse(string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject> ops)
        => _inner.ValidatePreviewReuse(snapshotDir, metadata, ops);

    public ApplyExecution Apply(IReadOnlyList<JsonObject> ops, string snapshotId)
    {
        if (!_faultAfterFirst || ops.Count < 2)
            return _inner.Apply(ops, snapshotId);

        _faultAfterFirst = false;
        var first = _inner.Apply(new[] { ops[0] }, snapshotId);
        FirstSuccessfulOpApplied = first.Ok;
        var handle = Json.GetString(ops[0], "handle");
        ObservedNativeTextAfterFirstWrite = ReadNativeTextByHandle(_inner, handle);
        if (!first.Ok)
            return first;
        throw new COMException(
            "injected CAD Apply fault after first successful write",
            unchecked((int)0x80004005));
    }

    private static string? ReadNativeTextByHandle(CadAdapter inner, string? handle)
    {
        if (string.IsNullOrWhiteSpace(handle)) return null;
        var query = inner.Read(new JsonObject
        {
            ["entityType"] = "Text",
            ["includeGeometry"] = true,
        });
        if (Json.GetArr(query, "entities") is not { } entities) return null;
        foreach (var node in entities)
        {
            if (node is not JsonObject entity) continue;
            if (!string.Equals(Json.GetString(entity, "handle"), handle, StringComparison.OrdinalIgnoreCase))
                continue;
            return Json.GetString(entity, "text");
        }
        return null;
    }

    public void Dispose()
    {
        // The test owns the inner CadAdapter lifetime.
    }
}
