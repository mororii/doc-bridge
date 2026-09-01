using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public sealed class ReliabilityTests : IDisposable
{
    private readonly TestHome _home = new();

    public void Dispose() => _home.Dispose();

    [Fact]
    public void Capabilities_include_runtime_status_and_limits()
    {
        using var host = new DocBridgeHost(_home.Options);
        host.Router.Register("hwp", new FakeAdapter());

        var result = host.CoreGetCapabilities(new JsonObject { ["app"] = "hwp" });

        Assert.True(Json.GetBool(result, "ok"));
        var hwp = Json.GetObj(Json.GetObj(result, "apps"), "hwp");
        Assert.NotNull(hwp);
        Assert.True(Json.GetBool(hwp, "available"));
        Assert.NotNull(Json.GetArr(hwp, "writeOps"));
    }

    [Fact]
    public void Failed_apply_is_automatically_rolled_back_and_timed()
    {
        using var host = new DocBridgeHost(_home.Options);
        var adapter = new MutateThenFailAdapter();
        host.Router.Register("fake", adapter);
        var ops = new JsonArray(new JsonObject { ["op"] = "insert_text", ["text"] = "changed" });

        var dry = host.ApplyOps("fake", new JsonObject { ["ops"] = ops.DeepClone(), ["dryRun"] = true });
        Assert.True(Json.GetBool(dry, "ok"));

        var applied = host.ApplyOps("fake", new JsonObject
        {
            ["ops"] = ops.DeepClone(),
            ["dryRun"] = false,
            ["confirmToken"] = Json.GetString(dry, "confirmToken"),
        });

        Assert.False(Json.GetBool(applied, "ok"));
        Assert.Equal("original", adapter.State);
        var rollback = Json.GetObj(applied, "rollback");
        Assert.True(Json.GetBool(rollback, "attempted"));
        Assert.True(Json.GetBool(rollback, "verified"));
        Assert.NotEmpty(Json.GetArr(applied, "operationResults")!);
        Assert.NotNull(applied["elapsedMs"]);
    }

    [Fact]
    public void Hwp_relative_insert_uses_unique_anchor_and_paragraph_boundaries()
    {
        using var host = new DocBridgeHost(_home.Options);
        var adapter = new FakeAdapter { BodyText = "제목\n기준 문단\n마지막" };
        host.Router.Register("hwp", adapter);
        var ops = new JsonArray
        {
            new JsonObject { ["op"] = "insert_before_text", ["anchor"] = "기준 문단", ["text"] = "앞 문단" },
            new JsonObject { ["op"] = "insert_after_text", ["anchor"] = "기준 문단", ["text"] = "뒤 문단" },
        };

        var dry = host.ApplyOps("hwp", new JsonObject { ["ops"] = ops.DeepClone(), ["dryRun"] = true });
        Assert.True(Json.GetBool(dry, "ok"), dry.ToJsonString());
        var applied = host.ApplyOps("hwp", new JsonObject
        {
            ["ops"] = ops.DeepClone(),
            ["dryRun"] = false,
            ["confirmToken"] = Json.GetString(dry, "confirmToken"),
        });

        Assert.True(Json.GetBool(applied, "ok"), applied.ToJsonString());
        Assert.Equal("제목\n앞 문단\n기준 문단\n뒤 문단\n마지막", adapter.BodyText);
    }

    [Fact]
    public void Hwp_relative_insert_rejects_ambiguous_anchor_without_occurrence()
    {
        using var host = new DocBridgeHost(_home.Options);
        var adapter = new FakeAdapter { BodyText = "반복\n중간\n반복" };
        host.Router.Register("hwp", adapter);
        var result = host.ApplyOps("hwp", new JsonObject
        {
            ["ops"] = new JsonArray
            {
                new JsonObject { ["op"] = "insert_after_text", ["anchor"] = "반복", ["text"] = "추가" },
            },
            ["dryRun"] = true,
        });

        Assert.False(Json.GetBool(result, "ok"));
        Assert.Contains(Json.GetArr(result, "errors")!, error =>
            error!.GetValue<string>().Contains("occurrence", StringComparison.OrdinalIgnoreCase));
        Assert.Equal("반복\n중간\n반복", adapter.BodyText);
    }

    private sealed class MutateThenFailAdapter : IAppAdapter
    {
        public string App => "fake";
        public string State { get; private set; } = "original";

        public AdapterStatus GetStatus() => new(true, true, App, "1", "atomic://doc", null);
        public JsonObject GetCapabilities() => new()
        {
            ["app"] = App,
            ["writeOps"] = new JsonArray("insert_text"),
            ["limits"] = new JsonObject(),
        };
        public ContextResult GetActiveContext() => new() { Ok = true, App = App, DocumentRef = "atomic://doc" };
        public JsonObject Read(JsonObject args) => new() { ["ok"] = true, ["state"] = State };
        public ApplyPreview Preview(IReadOnlyList<JsonObject> ops)
        {
            var preview = new ApplyPreview();
            preview.Affected.Add(new AffectedRef("state", "atomic://doc"));
            return preview;
        }
        public ApplyExecution Apply(IReadOnlyList<JsonObject> ops, string snapshotId)
        {
            State = "changed";
            var result = new ApplyExecution { Ok = false };
            result.Errors.Add("intentional failure after mutation");
            return result;
        }
        public void CaptureSnapshot(string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject>? ops = null)
        {
            File.WriteAllText(Path.Combine(snapshotDir, "state.txt"), State);
            metadata["payload"] = "state.txt";
        }
        public JsonObject RestoreSnapshot(string snapshotDir, JsonObject metadata)
        {
            State = File.ReadAllText(Path.Combine(snapshotDir, "state.txt"));
            return new JsonObject
            {
                ["ok"] = State == "original",
                ["readback"] = new JsonObject { ["verified"] = State == "original" },
            };
        }
        public void Dispose() { }
    }
}
