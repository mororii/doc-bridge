using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public sealed class HwpErrorPropagationTests : IDisposable
{
    private readonly TestHome _home = new();
    public void Dispose() => _home.Dispose();

    [Fact]
    public void Status_preserves_circuit_error_metadata()
    {
        using var host = Host(FailureStage.Status, "HWP_CIRCUIT_OPEN");
        var hwp = Json.GetObj(Json.GetObj(host.CoreGetStatus(), "apps"), "hwp")!;
        AssertStructured(hwp, "HWP_CIRCUIT_OPEN");
    }

    [Fact]
    public void Context_preserves_timeout_error_metadata()
    {
        using var host = Host(FailureStage.Context, "HWP_COM_TIMEOUT");
        AssertStructured(host.GetActiveContext("hwp"), "HWP_COM_TIMEOUT");
    }

    [Fact]
    public void Read_preserves_circuit_error_metadata()
    {
        using var host = Host(FailureStage.Read, "HWP_CIRCUIT_OPEN");
        AssertStructured(host.Read("hwp", new JsonObject()), "HWP_CIRCUIT_OPEN");
    }

    [Fact]
    public void Preview_preserves_timeout_error_metadata()
    {
        using var host = Host(FailureStage.Preview, "HWP_COM_TIMEOUT");
        var result = host.ApplyOps("hwp", Batch(dryRun: true));
        AssertStructured(result, "HWP_COM_TIMEOUT");
    }

    [Fact]
    public void Apply_preserves_circuit_error_metadata_after_safe_rollback()
    {
        var adapter = new ThrowingHwpAdapter(FailureStage.None, "HWP_CIRCUIT_OPEN");
        using var host = new DocBridgeHost(_home.Options);
        host.Router.Register("hwp", adapter);
        var dry = host.ApplyOps("hwp", Batch(dryRun: true));
        Assert.True(Json.GetBool(dry, "ok"), dry.ToJsonString());

        adapter.Stage = FailureStage.Apply;
        var result = host.ApplyOps("hwp", Batch(
            dryRun: false, confirmToken: Json.GetString(dry, "confirmToken")));

        AssertStructured(result, "HWP_CIRCUIT_OPEN");
        Assert.True(Json.GetBool(Json.GetObj(result, "rollback"), "verified"));
    }

    private DocBridgeHost Host(FailureStage stage, string code)
    {
        var host = new DocBridgeHost(_home.Options);
        host.Router.Register("hwp", new ThrowingHwpAdapter(stage, code));
        return host;
    }

    private static JsonObject Batch(bool dryRun, string? confirmToken = null)
    {
        var result = new JsonObject
        {
            ["ops"] = new JsonArray(new JsonObject { ["op"] = "insert_text", ["text"] = "test" }),
            ["dryRun"] = dryRun,
        };
        if (confirmToken is not null) result["confirmToken"] = confirmToken;
        return result;
    }

    private static void AssertStructured(JsonObject result, string code)
    {
        Assert.False(Json.GetBool(result, "ok"));
        Assert.Equal(code, Json.GetString(result, "errorCode"));
        Assert.Equal(4321, Json.GetInt(result, "retryAfterMs"));
        Assert.True(Json.GetBool(result, "retryable"));
        Assert.False(Json.GetBool(result, "automaticRetry"));
        Assert.Equal("after-delay", Json.GetString(Json.GetObj(result, "retryPolicy"), "mode"));
        Assert.False(string.IsNullOrWhiteSpace(Json.GetString(result, "userAction")));
    }

    private enum FailureStage { None, Status, Context, Read, Preview, Apply }

    private sealed class ThrowingHwpAdapter : IAppAdapter
    {
        private readonly string _code;
        internal ThrowingHwpAdapter(FailureStage stage, string code) { Stage = stage; _code = code; }
        internal FailureStage Stage { get; set; }
        public string App => "hwp";
        private HwpAutomationException Failure() => new(
            _code, "구조화 한글 오류", "팝업을 닫고 잠시 뒤 다시 시도하세요", retryAfterMs: 4321);
        public AdapterStatus GetStatus() => Stage == FailureStage.Status
            ? throw Failure()
            : new AdapterStatus(true, true, App, "test", "hwp://test-document", "test");
        public JsonObject GetCapabilities() => new()
        {
            ["app"] = App,
            ["writeOps"] = new JsonArray("insert_text"),
            ["limits"] = new JsonObject(),
        };
        public ContextResult GetActiveContext() => Stage == FailureStage.Context
            ? throw Failure()
            : new ContextResult { Ok = true, App = App, DocumentRef = "hwp://test-document" };
        public JsonObject Read(JsonObject args) => Stage == FailureStage.Read
            ? throw Failure()
            : new JsonObject { ["ok"] = true };
        public ApplyPreview Preview(IReadOnlyList<JsonObject> ops)
        {
            if (Stage == FailureStage.Preview) throw Failure();
            var result = new ApplyPreview();
            result.Affected.Add(new AffectedRef("document", "hwp://test-document"));
            return result;
        }
        public ApplyExecution Apply(IReadOnlyList<JsonObject> ops, string snapshotId)
        {
            if (Stage == FailureStage.Apply) throw Failure();
            return new ApplyExecution { Ok = true };
        }
        public void CaptureSnapshot(string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject>? ops = null) =>
            metadata["payload"] = "test";
        public JsonObject RestoreSnapshot(string snapshotDir, JsonObject metadata) => new()
        {
            ["ok"] = true,
            ["readback"] = new JsonObject { ["verified"] = true },
        };
        public void Dispose() { }
    }
}
