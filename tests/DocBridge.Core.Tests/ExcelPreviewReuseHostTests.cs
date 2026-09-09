using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelPreviewReuseHostTests
{
    [Fact]
    public void Format_only_fingerprint_change_denies_apply_before_token_consume()
    {
        using var home = new TestHome();
        using var host = new DocBridgeHost(home.Options);
        var adapter = new FreshnessStubAdapter
        {
            Reuse = new JsonObject
            {
                ["ok"] = true,
                ["reusable"] = true,
                ["freshPreviewAllowed"] = false,
                ["fingerprintMethod"] = "excel-format-target-sha256",
            },
        };
        host.Router.Register("excel", adapter);

        var dry = host.ApplyOps("excel", FormatBatch(dryRun: true));
        Assert.True(Json.GetBool(dry, "ok"), dry.ToJsonString());
        var token = Json.GetString(dry, "confirmToken");
        Assert.False(string.IsNullOrWhiteSpace(token));

        adapter.Reuse["reusable"] = false;
        adapter.Reuse["reason"] = "format-only target style fingerprint changed after dry-run";
        var denied = host.ApplyOps("excel", FormatBatch(dryRun: false, token: token));
        Assert.False(Json.GetBool(denied, "ok"));
        Assert.Contains(Json.GetArr(denied, "errors")!, item =>
            item!.GetValue<string>().Contains("changed after dry-run", StringComparison.Ordinal));
        Assert.Equal(0, adapter.ApplyCalls);
        Assert.True(adapter.ValidateCalls >= 1);
    }

    [Fact]
    public void Non_format_fresh_preview_is_allowed_and_does_not_deny_apply()
    {
        using var home = new TestHome();
        using var host = new DocBridgeHost(home.Options);
        var adapter = new FreshnessStubAdapter
        {
            Reuse = new JsonObject
            {
                ["ok"] = true,
                ["reusable"] = false,
                ["freshPreviewAllowed"] = true,
                ["fingerprintMethod"] = "excel-not-fingerprinted",
                ["reason"] = "Excel content, layout, and mixed batches are not fingerprinted; a fresh preview is required",
            },
        };
        host.Router.Register("excel", adapter);

        var dry = host.ApplyOps("excel", FormatBatch(dryRun: true));
        Assert.True(Json.GetBool(dry, "ok"), dry.ToJsonString());
        var previewAfterDry = adapter.PreviewCalls;

        var applied = host.ApplyOps("excel", FormatBatch(dryRun: false, token: Json.GetString(dry, "confirmToken")));
        Assert.True(Json.GetBool(applied, "ok"), applied.ToJsonString());
        Assert.False(Json.GetBool(Json.GetObj(applied, "timings"), "previewReused"));
        Assert.Equal("excel-not-fingerprinted", Json.GetString(Json.GetObj(applied, "timings"), "fingerprintMethod"));
        Assert.True(adapter.PreviewCalls > previewAfterDry);
        Assert.Equal(1, adapter.ApplyCalls);
    }

    private static JsonObject FormatBatch(bool dryRun, string? token = null) => new()
    {
        ["ops"] = new JsonArray
        {
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = new JsonObject { ["sheet"] = "Sheet1" },
                ["range"] = "A1",
                ["style"] = new JsonObject { ["bold"] = true },
            },
        },
        ["dryRun"] = dryRun,
        ["confirmToken"] = token,
    };

    private sealed class FreshnessStubAdapter : IAppAdapter, IPreviewReuseAdapter
    {
        public JsonObject Reuse { get; set; } = new();
        public int PreviewCalls { get; private set; }
        public int ApplyCalls { get; private set; }
        public int ValidateCalls { get; private set; }

        public string App => "excel";

        public AdapterStatus GetStatus() =>
            new(true, true, "excel", "test", @"C:\docbridge-fixtures\freshness.xlsx", "stub");

        public JsonObject GetCapabilities() => new();
        public ContextResult GetActiveContext() => new() { Ok = true, App = App, DocumentRef = @"C:\docbridge-fixtures\freshness.xlsx" };
        public JsonObject Read(JsonObject args) => new() { ["ok"] = true };

        public ApplyPreview Preview(IReadOnlyList<JsonObject> ops)
        {
            PreviewCalls++;
            var preview = new ApplyPreview();
            preview.Affected.Add(new AffectedRef("range", "Sheet1!A1"));
            preview.Diff.Add(new DiffEntry { Ref = "style", Before = JsonValue.Create("current"), After = JsonValue.Create(true) });
            return preview;
        }

        public ApplyExecution Apply(IReadOnlyList<JsonObject> ops, string snapshotId)
        {
            ApplyCalls++;
            return new ApplyExecution
            {
                Ok = true,
                Readback = new JsonObject { ["verified"] = true },
            };
        }

        public void CaptureSnapshot(string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject>? ops = null)
        {
            Directory.CreateDirectory(snapshotDir);
            File.WriteAllText(Path.Combine(snapshotDir, "state.json"), new JsonObject
            {
                ["restoreMode"] = "format-only",
                ["documentRef"] = @"C:\docbridge-fixtures\freshness.xlsx",
            }.ToJsonString());
            metadata["payload"] = "stub";
            metadata["documentRef"] = @"C:\docbridge-fixtures\freshness.xlsx";
        }

        public JsonObject RestoreSnapshot(string snapshotDir, JsonObject metadata) =>
            new() { ["ok"] = true, ["restored"] = true };

        public JsonObject ValidatePreviewReuse(string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject> ops)
        {
            ValidateCalls++;
            return Reuse.DeepClone().AsObject();
        }

        public void Dispose() { }
    }
}
