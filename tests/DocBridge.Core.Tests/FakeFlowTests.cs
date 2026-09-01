using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

/// <summary>
/// M0 핵심: FakeAdapter로 read → dry-run → confirm → apply → readback → restore 전체 흐름 검증.
/// 명령서 인수 조건:
///   - dryRun=true에서 diff와 confirmToken 반환
///   - dryRun=false는 confirmToken 없이 실패
///   - apply 후 readback verify 반환
///   - snapshot 복원 검증
///   - 금지 op 차단
/// </summary>
public class FakeFlowTests : IDisposable
{
    private readonly TestHome _home = new();
    private readonly DocBridgeHost _host;
    public FakeFlowTests() => _host = new DocBridgeHost(_home.Options);
    public void Dispose() { _host.Dispose(); _home.Dispose(); }

    private static JsonObject SetValuesBatch(bool dryRun, string? token = null) => new()
    {
        ["ops"] = new JsonArray
        {
            new JsonObject
            {
                ["op"] = "set_values",
                ["range"] = "Sheet1!B2",
                ["values"] = new JsonArray(new JsonArray("1500")),
            },
        },
        ["dryRun"] = dryRun,
        ["confirmToken"] = token,
    };

    [Fact]
    public void Core_ping_reports_adapters()
    {
        var ping = _host.CorePing();
        Assert.True(Json.GetBool(ping, "ok"));
        var adapters = Json.GetArr(ping, "adapters")!;
        Assert.Contains(adapters, a => a!.GetValue<string>() == "fake");
        Assert.Contains(adapters, a => a!.GetValue<string>() == "excel");
    }

    [Fact]
    public void Get_active_context_returns_structured_json()
    {
        var ctx = _host.GetActiveContext("fake");
        Assert.True(Json.GetBool(ctx, "ok"));
        Assert.Equal("fake", Json.GetString(ctx, "app"));
        Assert.NotNull(Json.GetObj(ctx, "summary"));
        Assert.NotNull(Json.GetObj(ctx, "selection"));
    }

    [Fact]
    public void Dryrun_returns_diff_and_confirm_token()
    {
        var result = _host.ApplyOps("fake", SetValuesBatch(dryRun: true));
        Assert.True(Json.GetBool(result, "ok"));
        Assert.True(Json.GetBool(result, "dryRun"));
        Assert.NotNull(Json.GetString(result, "confirmToken"));
        Assert.NotNull(Json.GetString(result, "snapshotId"));
        var diff = Json.GetArr(result, "diff")!;
        Assert.NotEmpty(diff);
        // B2 before=1000 after=1500
        var first = diff[0] as JsonObject;
        Assert.Equal("1000", first!["before"]!.GetValue<string>());
        Assert.Equal("1500", first["after"]!.GetValue<string>());
    }

    [Fact]
    public void Apply_without_confirm_token_fails()
    {
        var result = _host.ApplyOps("fake", SetValuesBatch(dryRun: false, token: null));
        Assert.False(Json.GetBool(result, "ok"));
        var errors = Json.GetArr(result, "errors")!;
        Assert.Contains(errors, e => e!.GetValue<string>().Contains("confirmToken"));
    }

    [Fact]
    public void Full_flow_dryrun_confirm_apply_readback()
    {
        // 1) dry-run
        var dry = _host.ApplyOps("fake", SetValuesBatch(dryRun: true));
        var token = Json.GetString(dry, "confirmToken")!;
        var snapshotId = Json.GetString(dry, "snapshotId")!;

        // 2) apply with token
        var applied = _host.ApplyOps("fake", SetValuesBatch(dryRun: false, token: token));
        Assert.True(Json.GetBool(applied, "ok"));
        Assert.False(Json.GetBool(applied, "dryRun"));

        // 3) readback verify
        var readback = Json.GetObj(applied, "readback")!;
        Assert.True(Json.GetBool(readback, "verified"));

        // 4) 실제 상태 변경 확인
        var ctx = _host.Read("fake", new JsonObject { ["range"] = "Sheet1!B2" });
        var values = Json.GetArr(ctx, "values")!;
        Assert.Equal("1500", values[0]![0]!.GetValue<string>());

        // 5) 스냅샷 목록에 존재
        var snaps = _host.CoreListSnapshots(new JsonObject { ["app"] = "fake" });
        var arr = Json.GetArr(snaps, "snapshots")!;
        Assert.Contains(arr, s => Json.GetString(s as JsonObject, "snapshotId") == snapshotId);
    }

    [Fact]
    public void Apply_reuses_dryrun_preview_only_after_full_fingerprint_match()
    {
        var adapter = Assert.IsType<DocBridge.Core.Adapters.FakeAdapter>(_host.Router.Get("fake"));
        var dry = _host.ApplyOps("fake", SetValuesBatch(dryRun: true));
        Assert.Equal(1, adapter.PreviewCallCount);

        var applied = _host.ApplyOps("fake", SetValuesBatch(
            dryRun: false, token: Json.GetString(dry, "confirmToken")));

        Assert.True(Json.GetBool(applied, "ok"));
        Assert.Equal(1, adapter.PreviewCallCount); // apply에서 Preview를 다시 계산하지 않음
        var timings = Json.GetObj(applied, "timings")!;
        Assert.True(Json.GetBool(timings, "previewReused"));
        Assert.Equal("fake-full-state", Json.GetString(timings, "fingerprintMethod"));
        Assert.NotNull(Json.GetLong(timings, "totalMs"));
    }

    [Fact]
    public void Changed_document_denies_cached_preview_without_consuming_token()
    {
        var adapter = Assert.IsType<DocBridge.Core.Adapters.FakeAdapter>(_host.Router.Get("fake"));
        var dry = _host.ApplyOps("fake", SetValuesBatch(dryRun: true));
        var token = Json.GetString(dry, "confirmToken")!;

        adapter.Sheets["Sheet1"]["B2"] = "external-change";
        var denied = _host.ApplyOps("fake", SetValuesBatch(false, token));
        Assert.False(Json.GetBool(denied, "ok"));
        Assert.Contains(Json.GetArr(denied, "errors")!, item =>
            item!.GetValue<string>().Contains("changed after dry-run", StringComparison.Ordinal));

        // fingerprint를 원상태로 돌리면 같은 토큰을 사용할 수 있다. 거부 단계에서 소비하지 않았다는 뜻이다.
        adapter.Sheets["Sheet1"]["B2"] = "1000";
        var applied = _host.ApplyOps("fake", SetValuesBatch(false, token));
        Assert.True(Json.GetBool(applied, "ok"));
    }

    [Fact]
    public void Dryrun_reports_phase_timings()
    {
        var dry = _host.ApplyOps("fake", SetValuesBatch(dryRun: true));
        var timings = Json.GetObj(dry, "timings")!;
        foreach (var key in new[] { "validationMs", "lockWaitMs", "statusMs", "previewMs", "snapshotMs", "tokenMs", "totalMs" })
            Assert.NotNull(Json.GetLong(timings, key));
    }

    [Fact]
    public void Token_reuse_after_apply_fails()
    {
        var dry = _host.ApplyOps("fake", SetValuesBatch(dryRun: true));
        var token = Json.GetString(dry, "confirmToken")!;
        Assert.True(Json.GetBool(_host.ApplyOps("fake", SetValuesBatch(false, token)), "ok"));
        var second = _host.ApplyOps("fake", SetValuesBatch(false, token));
        Assert.False(Json.GetBool(second, "ok"));
    }

    [Fact]
    public void Apply_is_denied_when_document_changed_after_dryrun()
    {
        var dry = _host.ApplyOps("fake", SetValuesBatch(dryRun: true));
        var token = Json.GetString(dry, "confirmToken")!;
        var adapter = Assert.IsType<DocBridge.Core.Adapters.FakeAdapter>(_host.Router.Get("fake"));
        adapter.DocumentRef = "another-document";

        var denied = _host.ApplyOps("fake", SetValuesBatch(dryRun: false, token: token));

        Assert.False(Json.GetBool(denied, "ok"));
        Assert.Contains(Json.GetArr(denied, "errors")!, e =>
            e!.GetValue<string>().Contains("document changed after dry-run", StringComparison.Ordinal));
        Assert.Equal("1000", adapter.Sheets["Sheet1"]["B2"]);
    }

    [Fact]
    public void Forbidden_op_is_blocked()
    {
        var batch = new JsonObject
        {
            ["ops"] = new JsonArray { new JsonObject { ["op"] = "run_macro", ["name"] = "x" } },
            ["dryRun"] = true,
        };
        var result = _host.ApplyOps("fake", batch);
        Assert.False(Json.GetBool(result, "ok"));
        var errors = Json.GetArr(result, "errors")!;
        Assert.Contains(errors, e => e!.GetValue<string>().Contains("FORBIDDEN"));
    }

    [Fact]
    public void Highrisk_op_requires_highrisk_confirm()
    {
        var batch = new JsonObject
        {
            ["ops"] = new JsonArray
            {
                new JsonObject { ["op"] = "delete_entities", ["handles"] = new JsonArray("H1") },
            },
            ["dryRun"] = true,
        };
        var dry = _host.ApplyOps("fake", batch);
        Assert.True(Json.GetBool(dry, "ok"));
        Assert.True(Json.GetBool(dry, "requiresHighRiskApproval"));

        var token = Json.GetString(dry, "confirmToken")!;
        // highRiskConfirm 없이 적용 → 실패
        var applyNoConfirm = new JsonObject
        {
            ["ops"] = batch["ops"]!.DeepClone(),
            ["dryRun"] = false,
            ["confirmToken"] = token,
        };
        var denied = _host.ApplyOps("fake", applyNoConfirm);
        Assert.False(Json.GetBool(denied, "ok"));
        Assert.Contains(Json.GetArr(denied, "errors")!, e => e!.GetValue<string>().Contains("high-risk"));
    }

    [Fact]
    public void Snapshot_restore_roundtrip_via_two_phase()
    {
        // 상태 변경
        var dry = _host.ApplyOps("fake", SetValuesBatch(dryRun: true));
        var token = Json.GetString(dry, "confirmToken")!;
        var snapshotId = Json.GetString(dry, "snapshotId")!;
        _host.ApplyOps("fake", SetValuesBatch(false, token));

        // 현재 B2=1500 확인
        var before = _host.Read("fake", new JsonObject { ["range"] = "Sheet1!B2" });
        Assert.Equal("1500", Json.GetArr(before, "values")![0]![0]!.GetValue<string>());

        // 복원 1단계: confirmToken 발급
        var restoreDry = _host.CoreRestoreSnapshot(new JsonObject { ["snapshotId"] = snapshotId });
        Assert.True(Json.GetBool(restoreDry, "ok"));
        Assert.True(Json.GetBool(restoreDry, "dryRun"));
        var restoreToken = Json.GetString(restoreDry, "confirmToken")!;

        // 복원 2단계: 실제 복원
        var restored = _host.CoreRestoreSnapshot(new JsonObject
        {
            ["snapshotId"] = snapshotId,
            ["confirmToken"] = restoreToken,
        });
        Assert.True(Json.GetBool(restored, "ok"));
        Assert.True(Json.GetBool(Json.GetObj(restored, "readback"), "verified"));

        // B2가 원래 값 1000으로 복귀
        var after = _host.Read("fake", new JsonObject { ["range"] = "Sheet1!B2" });
        Assert.Equal("1000", Json.GetArr(after, "values")![0]![0]!.GetValue<string>());
    }

    [Fact]
    public void Restore_without_token_fails_for_unknown_snapshot()
    {
        var result = _host.CoreRestoreSnapshot(new JsonObject { ["snapshotId"] = "nope" });
        Assert.False(Json.GetBool(result, "ok"));
    }

    [Fact]
    public void Audit_log_is_written_to_file_not_stdout()
    {
        _host.CorePing();
        _host.GetActiveContext("fake");
        var logsDir = Path.Combine(_home.Dir, "logs");
        Assert.True(Directory.Exists(logsDir));
        var files = Directory.GetFiles(logsDir, "audit-*.jsonl");
        Assert.NotEmpty(files);
        var lines = File.ReadAllLines(files[0]);
        Assert.NotEmpty(lines);
        foreach (var line in lines)
            Assert.NotNull(JsonNode.Parse(line)); // JSONL 형식 검증
    }

    [Fact]
    public void Apply_audit_contains_total_and_sanitized_per_operation_timings()
    {
        var dry = _host.ApplyOps("fake", SetValuesBatch(dryRun: true));
        var applied = _host.ApplyOps("fake", SetValuesBatch(
            dryRun: false, token: Json.GetString(dry, "confirmToken")));
        Assert.True(Json.GetBool(applied, "ok"));

        var logsDir = Path.Combine(_home.Dir, "logs");
        var entries = Directory.GetFiles(logsDir, "audit-*.jsonl")
            .SelectMany(File.ReadAllLines)
            .Select(line => JsonNode.Parse(line) as JsonObject)
            .Where(entry => entry is not null)
            .ToList();
        var applyEntry = entries.Last(entry =>
            Json.GetString(entry, "tool") == "fake_apply_ops" &&
            Json.GetString(entry, "action") == "apply");
        var detail = Json.GetObj(applyEntry, "detail")!;
        var auditTimings = Json.GetObj(detail, "timings")!;
        Assert.NotNull(Json.GetLong(auditTimings, "totalMs"));

        var operation = Assert.IsType<JsonObject>(Assert.Single(Json.GetArr(detail, "operationResults")!));
        Assert.Equal("set_values", Json.GetString(operation, "op"));
        Assert.NotNull(Json.GetLong(operation, "elapsedMs"));
        Assert.Null(operation["before"]);
        Assert.Null(operation["after"]);
    }
}
