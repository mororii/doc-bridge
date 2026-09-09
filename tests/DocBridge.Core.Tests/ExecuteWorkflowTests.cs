using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExecuteWorkflowTests : IDisposable
{
    internal static readonly string AbsExcel = Path.GetFullPath(@"C:\docbridge-identity\book.xlsx");
    internal static readonly string AbsCad = Path.GetFullPath(@"C:\docbridge-identity\drawing.dwg");
    internal const string HwpRef = "untitled-18636-2";

    private readonly TestHome _home = new();
    private readonly DocBridgeHost _host;
    private readonly FakeAdapter _fake;
    private readonly FakeAdapter _excel;
    private readonly FakeAdapter _hwp;
    private readonly FakeAdapter _cad;
    private readonly FakeAdapter _gstarcad;

    public ExecuteWorkflowTests()
    {
        _host = new DocBridgeHost(_home.Options);
        _fake = new FakeAdapter();
        _excel = new FakeAdapter { DocumentRef = AbsExcel };
        _hwp = new FakeAdapter { DocumentRef = HwpRef };
        _cad = new FakeAdapter { DocumentRef = AbsCad };
        _gstarcad = new FakeAdapter { DocumentRef = AbsCad };
        _host.Router.Register("fake", _fake);
        _host.Router.Register("excel", _excel);
        _host.Router.Register("hwp", _hwp);
        _host.Router.Register("cad", _cad);
        _host.Router.Register("gstarcad", _gstarcad);
    }

    public void Dispose()
    {
        _host.Dispose();
        _home.Dispose();
    }

    private static string NewId() => Guid.NewGuid().ToString("D");

    private static JsonArray LegacySetValuesOps(string value = "1500") => new(
        new JsonObject
        {
            ["op"] = "set_values",
            ["range"] = "Sheet1!B2",
            ["values"] = new JsonArray(new JsonArray(value)),
        });

    private static JsonObject ExcelValues(string requestId, string? expected = null, string value = "1500") => new()
    {
        ["ops"] = new JsonArray(new JsonObject
        {
            ["op"] = "set_values",
            ["range"] = "Sheet1!B2",
            ["values"] = new JsonArray(new JsonArray(value)),
        }),
        ["executionMode"] = "execute",
        ["requestId"] = requestId,
        ["expectedDocumentRef"] = expected ?? AbsExcel,
    };

    private static JsonObject CadLayer(string requestId, string? expected = null) => new()
    {
        ["ops"] = new JsonArray(new JsonObject
        {
            ["op"] = "set_layer_visibility",
            ["layer"] = "PLAN",
            ["visible"] = false,
        }),
        ["executionMode"] = "execute",
        ["requestId"] = requestId,
        ["expectedDocumentRef"] = expected ?? AbsCad,
    };

    [Fact]
    public void Legacy_dryrun_token_apply_still_works()
    {
        var requestOps = LegacySetValuesOps();
        var dry = _host.ApplyOps("fake", new JsonObject
        {
            ["ops"] = (JsonNode)requestOps.DeepClone(),
            ["dryRun"] = true,
        });
        Assert.True(Json.GetBool(dry, "ok"));
        Assert.NotNull(Json.GetString(dry, "confirmToken"));
        Assert.Null(dry["ops"]);

        var applied = _host.ApplyOps("fake", new JsonObject
        {
            ["ops"] = (JsonNode)requestOps.DeepClone(),
            ["dryRun"] = false,
            ["confirmToken"] = Json.GetString(dry, "confirmToken"),
        });
        Assert.True(Json.GetBool(applied, "ok"));
        Assert.Equal("1500", _fake.Sheets["Sheet1"]["B2"]);
    }

    [Fact]
    public void Legacy_excel_dryrun_still_accepts_book1_name()
    {
        _excel.DocumentRef = "Book1";
        var requestOps = new JsonArray(new JsonObject
        {
            ["op"] = "set_values",
            ["range"] = "Sheet1!B2",
            ["values"] = new JsonArray(new JsonArray("1500")),
            ["targetWorkbook"] = "Book1",
        });
        var dry = _host.ApplyOps("excel", new JsonObject
        {
            ["ops"] = (JsonNode)requestOps.DeepClone(),
            ["dryRun"] = true,
        });
        Assert.True(Json.GetBool(dry, "ok"), dry.ToJsonString());
        Assert.NotNull(Json.GetString(dry, "confirmToken"));
        Assert.Equal("Book1", _excel.DocumentRef);
    }

    [Fact]
    public void Execute_rejects_dryRun_token_and_highRisk_flags()
    {
        var withDry = ExcelValues(NewId());
        withDry["dryRun"] = false;
        var deniedDry = _host.ApplyOps("excel", withDry);
        Assert.False(Json.GetBool(deniedDry, "ok"));
        Assert.Equal(0, _excel.ApplyCallCount);
        Assert.Contains(Json.GetArr(deniedDry, "errors")!, e => e!.GetValue<string>().Contains("dryRun"));

        var withToken = ExcelValues(NewId());
        withToken["confirmToken"] = "conf_x";
        var deniedToken = _host.ApplyOps("excel", withToken);
        Assert.False(Json.GetBool(deniedToken, "ok"));
        Assert.Contains(Json.GetArr(deniedToken, "errors")!, e => e!.GetValue<string>().Contains("confirmToken"));

        var withRisk = ExcelValues(NewId());
        withRisk["highRiskConfirm"] = true;
        var deniedRisk = _host.ApplyOps("excel", withRisk);
        Assert.False(Json.GetBool(deniedRisk, "ok"));
        Assert.Contains(Json.GetArr(deniedRisk, "errors")!, e => e!.GetValue<string>().Contains("highRiskConfirm"));
    }

    [Fact]
    public void Execute_succeeds_with_one_preview_and_snapshot_and_no_token()
    {
        var beforeSnaps = Json.GetInt(_host.CoreListSnapshots(new JsonObject { ["app"] = "excel" }), "count") ?? 0;
        var result = _host.ApplyOps("excel", ExcelValues(NewId()));

        Assert.True(Json.GetBool(result, "ok"), result.ToJsonString());
        Assert.False(Json.GetBool(result, "dryRun"));
        Assert.Equal("execute", Json.GetString(result, "executionMode"));
        Assert.Null(Json.GetString(result, "confirmToken"));
        Assert.Equal(1, _excel.PreviewCallCount);
        Assert.Equal(1, _excel.CaptureSnapshotCallCount);
        Assert.Equal(1, _excel.ApplyCallCount);
        Assert.Equal(0, _excel.ValidatePreviewReuseCallCount);
        Assert.Equal("1500", _excel.Sheets["Sheet1"]["B2"]);
        Assert.Equal(beforeSnaps + 1, Json.GetInt(_host.CoreListSnapshots(new JsonObject { ["app"] = "excel" }), "count"));
        Assert.True(Json.GetBool(Json.GetObj(result, "readback"), "verified"));
        Assert.False(Json.GetBool(Json.GetObj(result, "rollbackCoverage"), "fullDocumentRestored"));
    }

    [Fact]
    public void Same_request_replays_without_adapter_calls_across_host_restart()
    {
        var id = NewId();
        var first = _host.ApplyOps("excel", ExcelValues(id));
        Assert.True(Json.GetBool(first, "ok"));
        var applyAfterFirst = _excel.ApplyCallCount;
        var previewAfterFirst = _excel.PreviewCallCount;
        var statusAfterFirst = _excel.StatusCallCount;

        var second = _host.ApplyOps("excel", ExcelValues(id));
        Assert.True(Json.GetBool(second, "ok"));
        Assert.True(Json.GetBool(second, "idempotentReplay"));
        Assert.Equal(applyAfterFirst, _excel.ApplyCallCount);
        Assert.Equal(previewAfterFirst, _excel.PreviewCallCount);
        Assert.Equal(statusAfterFirst, _excel.StatusCallCount);

        using var restarted = new DocBridgeHost(_home.Options);
        var replayed = restarted.ApplyOps("excel", ExcelValues(id));
        Assert.True(Json.GetBool(replayed, "ok"));
        Assert.True(Json.GetBool(replayed, "idempotentReplay"));
        Assert.Equal(Json.GetString(first, "snapshotId"), Json.GetString(replayed, "snapshotId"));
    }

    [Fact]
    public void Uuid_normalization_uses_the_same_journal_file()
    {
        var guid = Guid.NewGuid();
        var first = _host.ApplyOps("excel", ExcelValues(guid.ToString("D")));
        Assert.True(Json.GetBool(first, "ok"));

        var replayN = _host.ApplyOps("excel", ExcelValues(guid.ToString("N")));
        Assert.True(Json.GetBool(replayN, "idempotentReplay"));

        var replayB = _host.ApplyOps("excel", ExcelValues(guid.ToString("B").ToUpperInvariant()));
        Assert.True(Json.GetBool(replayB, "idempotentReplay"));
        Assert.Equal(1, _excel.ApplyCallCount);
    }

    [Fact]
    public void Mismatched_reuse_is_denied()
    {
        var id = NewId();
        Assert.True(Json.GetBool(_host.ApplyOps("excel", ExcelValues(id)), "ok"));

        var otherOps = ExcelValues(id, value: "9999");
        var deniedOps = _host.ApplyOps("excel", otherOps);
        Assert.False(Json.GetBool(deniedOps, "ok"));
        Assert.Contains(Json.GetArr(deniedOps, "errors")!, e => e!.GetValue<string>().Contains("different"));
        Assert.Equal("1500", _excel.Sheets["Sheet1"]["B2"]);

        var deniedApp = _host.ApplyOps("hwp", new JsonObject
        {
            ["ops"] = new JsonArray(new JsonObject { ["op"] = "append_text", ["text"] = "x" }),
            ["executionMode"] = "execute",
            ["requestId"] = id,
            ["expectedDocumentRef"] = HwpRef,
        });
        Assert.False(Json.GetBool(deniedApp, "ok"));
        Assert.DoesNotContain("x", _hwp.BodyText);
    }

    [Fact]
    public void Pending_journal_never_retries_or_calls_adapter()
    {
        var id = Guid.NewGuid().ToString("D");
        var journal = new ExecuteJournalService(_home.Options);
        Assert.True(journal.TryBegin(id, "excel", AbsExcel, "pending-hash"));

        var before = _excel.ApplyCallCount;
        var denied = _host.ApplyOps("excel", ExcelValues(id));
        Assert.False(Json.GetBool(denied, "ok"));
        Assert.Contains(Json.GetArr(denied, "errors")!, e =>
            e!.GetValue<string>().Contains("outcome unknown", StringComparison.OrdinalIgnoreCase));
        Assert.True(Json.GetBool(denied, "outcomeUnknown"));
        Assert.False(Json.GetBool(denied, "safeToRetry", true));
        Assert.Equal("started", Json.GetString(denied, "journalStatus"));
        Assert.Equal(before, _excel.ApplyCallCount);
        Assert.Equal(0, _excel.CaptureSnapshotCallCount);
    }

    [Fact]
    public void Failed_apply_replays_the_same_result_and_rolls_back()
    {
        _excel.FailNextApply = true;
        var id = NewId();
        var failed = _host.ApplyOps("excel", ExcelValues(id));
        Assert.False(Json.GetBool(failed, "ok"));
        Assert.True(Json.GetBool(Json.GetObj(failed, "rollback"), "attempted"));
        Assert.True(Json.GetBool(Json.GetObj(failed, "rollback"), "verified"));
        Assert.Equal("1000", _excel.Sheets["Sheet1"]["B2"]);

        var applyCount = _excel.ApplyCallCount;
        var replay = _host.ApplyOps("excel", ExcelValues(id));
        Assert.False(Json.GetBool(replay, "ok"));
        Assert.True(Json.GetBool(replay, "idempotentReplay"));
        Assert.Equal(applyCount, _excel.ApplyCallCount);
        Assert.Equal("1000", _excel.Sheets["Sheet1"]["B2"]);
    }

    [Fact]
    public void High_risk_preview_denies_without_snapshot()
    {
        _excel.ForceHighRiskPreview = true;
        var snaps = Json.GetInt(_host.CoreListSnapshots(new JsonObject { ["app"] = "excel" }), "count") ?? 0;
        var denied = _host.ApplyOps("excel", ExcelValues(NewId()));
        Assert.False(Json.GetBool(denied, "ok"));
        Assert.Contains(Json.GetArr(denied, "errors")!, e => e!.GetValue<string>().Contains("high-risk"));
        Assert.Equal(0, _excel.ApplyCallCount);
        Assert.Equal(0, _excel.CaptureSnapshotCallCount);
        Assert.Equal(snaps, Json.GetInt(_host.CoreListSnapshots(new JsonObject { ["app"] = "excel" }), "count"));
    }

    [Fact]
    public void Wrong_expected_document_denies_before_write()
    {
        var snaps = Json.GetInt(_host.CoreListSnapshots(new JsonObject { ["app"] = "excel" }), "count") ?? 0;
        var denied = _host.ApplyOps("excel", ExcelValues(NewId(), expected: Path.GetFullPath(@"C:\other\workbook.xlsx")));
        Assert.False(Json.GetBool(denied, "ok"));
        Assert.Equal(0, _excel.ApplyCallCount);
        Assert.Equal(0, _excel.CaptureSnapshotCallCount);
        Assert.Equal("1000", _excel.Sheets["Sheet1"]["B2"]);
        Assert.Equal(snaps, Json.GetInt(_host.CoreListSnapshots(new JsonObject { ["app"] = "excel" }), "count"));
    }

    [Theory]
    [InlineData("excel", "Book1")]
    [InlineData("excel", "Book1.xlsx")]
    [InlineData("cad", "Drawing1")]
    public void Ambiguous_unsaved_expected_ref_guides_legacy_preview(string app, string expected)
    {
        var adapter = app == "excel" ? _excel : _cad;
        var denied = _host.ApplyOps(app, app == "excel" ? ExcelValues(NewId(), expected) : CadLayer(NewId(), expected));
        Assert.False(Json.GetBool(denied, "ok"));
        Assert.Contains(Json.GetArr(denied, "errors")!, e =>
            e!.GetValue<string>().Contains("will not save", StringComparison.OrdinalIgnoreCase));
        Assert.Equal(0, adapter.ApplyCallCount);
        Assert.Equal(0, adapter.StatusCallCount);
    }

    [Fact]
    public void Explicit_workbook_is_not_resolved_from_active_status()
    {
        var batch = ExcelValues(NewId());
        var op = (JsonObject)((JsonArray)batch["ops"]!)[0]!;
        op["targetWorkbook"] = "Book1.xlsx";
        var denied = _host.ApplyOps("excel", batch);
        Assert.False(Json.GetBool(denied, "ok"));
        Assert.Contains(Json.GetArr(denied, "errors")!, e =>
            e!.GetValue<string>().Contains("fully qualified", StringComparison.OrdinalIgnoreCase));
        Assert.Equal(0, _excel.ApplyCallCount);
        Assert.Equal(0, _excel.StatusCallCount);
    }

    [Fact]
    public void Mixed_explicit_targets_are_rejected_before_adapter()
    {
        var batch = new JsonObject
        {
            ["ops"] = new JsonArray(
                new JsonObject
                {
                    ["op"] = "set_values",
                    ["range"] = "Sheet1!B2",
                    ["values"] = new JsonArray(new JsonArray("1")),
                    ["targetWorkbook"] = @"C:\data\one.xlsx",
                },
                new JsonObject
                {
                    ["op"] = "set_values",
                    ["range"] = "Sheet1!B3",
                    ["values"] = new JsonArray(new JsonArray("2")),
                }),
            ["executionMode"] = "execute",
            ["requestId"] = NewId(),
            ["expectedDocumentRef"] = @"C:\data\one.xlsx",
        };
        var denied = _host.ApplyOps("excel", batch);
        Assert.False(Json.GetBool(denied, "ok"));
        Assert.Contains(Json.GetArr(denied, "errors")!, e =>
            e!.GetValue<string>().Contains("mixed", StringComparison.OrdinalIgnoreCase));
        Assert.Equal(0, _excel.ApplyCallCount);
        Assert.Equal(0, _excel.StatusCallCount);
    }

    [Fact]
    public void Conflicting_explicit_workbooks_are_denied()
    {
        var batch = new JsonObject
        {
            ["ops"] = new JsonArray(
                new JsonObject
                {
                    ["op"] = "set_values",
                    ["range"] = "Sheet1!B2",
                    ["values"] = new JsonArray(new JsonArray("1")),
                    ["targetWorkbook"] = @"C:\data\one.xlsx",
                },
                new JsonObject
                {
                    ["op"] = "set_values",
                    ["range"] = "Sheet1!B3",
                    ["values"] = new JsonArray(new JsonArray("2")),
                    ["targetWorkbook"] = @"C:\data\two.xlsx",
                }),
            ["executionMode"] = "execute",
            ["requestId"] = NewId(),
            ["expectedDocumentRef"] = @"C:\data\one.xlsx",
        };
        var denied = _host.ApplyOps("excel", batch);
        Assert.False(Json.GetBool(denied, "ok"));
        Assert.Contains(Json.GetArr(denied, "errors")!, e =>
            e!.GetValue<string>().Contains("mixed explicit", StringComparison.OrdinalIgnoreCase));
        Assert.Equal(0, _excel.ApplyCallCount);
    }

    [Theory]
    [InlineData("cad")]
    [InlineData("gstarcad")]
    public void Cad_explicit_document_header_cannot_mutate_active_bystander(string app)
    {
        var adapter = app == "cad" ? _cad : _gstarcad;
        var active = AbsCad;
        var bystander = Path.GetFullPath(@"C:\docbridge-identity\other.dwg");
        adapter.DocumentRef = active;
        adapter.LayerVisibility["PLAN"] = true;

        var batch = CadLayer(NewId(), bystander);
        var op = (JsonObject)((JsonArray)batch["ops"]!)[0]!;
        op["document"] = bystander;
        var denied = _host.ApplyOps(app, batch);

        Assert.False(Json.GetBool(denied, "ok"), denied.ToJsonString());
        Assert.Contains(Json.GetArr(denied, "errors")!, e =>
            e!.GetValue<string>().Contains("active drawing", StringComparison.OrdinalIgnoreCase) ||
            e!.GetValue<string>().Contains("does not match current document", StringComparison.OrdinalIgnoreCase));
        Assert.Equal(0, adapter.PreviewCallCount);
        Assert.Equal(0, adapter.ApplyCallCount);
        Assert.Equal(0, adapter.CaptureSnapshotCallCount);
        Assert.Null(adapter.LastMutatedDocumentRef);
        Assert.True(adapter.LayerVisibility["PLAN"]);
        Assert.Equal(active, adapter.DocumentRef);
    }

    [Fact]
    public void Fake_apply_denies_explicit_target_that_does_not_match_active()
    {
        var active = AbsCad;
        var header = Path.GetFullPath(@"C:\docbridge-identity\other.dwg");
        var fake = new FakeAdapter { DocumentRef = active };
        var applied = fake.Apply(new[]
        {
            new JsonObject
            {
                ["op"] = "set_layer_visibility",
                ["layer"] = "PLAN",
                ["visible"] = false,
                ["document"] = header,
            },
        }, "snap");
        Assert.False(applied.Ok);
        Assert.Contains(applied.Errors, e => e.Contains("no auto activation", StringComparison.OrdinalIgnoreCase));
        Assert.False(fake.LayerVisibility.ContainsKey("PLAN"));
    }

    [Fact]
    public void Execute_bound_ops_keep_original_hash_and_caller_json()
    {
        var op = new JsonObject
        {
            ["op"] = "set_values",
            ["range"] = "Sheet1!B2",
            ["values"] = new JsonArray(new JsonArray("1500")),
        };
        var id = NewId();
        var batch = new JsonObject
        {
            ["ops"] = new JsonArray(op),
            ["executionMode"] = "execute",
            ["requestId"] = id,
            ["expectedDocumentRef"] = AbsExcel,
        };
        var originalHash = ConfirmTokenService.HashOps(new[] { op });
        Assert.True(Json.GetBool(_host.ApplyOps("excel", batch), "ok"));
        Assert.False(op.ContainsKey("targetWorkbook"));
        Assert.Equal(originalHash, ConfirmTokenService.HashOps(new[] { op }));

        var replay = _host.ApplyOps("excel", new JsonObject
        {
            ["ops"] = new JsonArray((JsonObject)op.DeepClone()),
            ["executionMode"] = "execute",
            ["requestId"] = id,
            ["expectedDocumentRef"] = AbsExcel,
        });
        Assert.True(Json.GetBool(replay, "idempotentReplay"));
        Assert.Equal(1, _excel.ApplyCallCount);
    }

    [Fact]
    public void Fake_target_switch_after_status_denies_bound_execute()
    {
        var switched = Path.GetFullPath(@"C:\docbridge-identity\switched.xlsx");
        _excel.SwitchDocumentRefAfterStatus = switched;
        var denied = _host.ApplyOps("excel", ExcelValues(NewId()));
        Assert.False(Json.GetBool(denied, "ok"), denied.ToJsonString());
        Assert.Contains(Json.GetArr(denied, "errors")!, e =>
            e!.GetValue<string>().Contains("no auto activation", StringComparison.OrdinalIgnoreCase) ||
            e!.GetValue<string>().Contains("does not match", StringComparison.OrdinalIgnoreCase));
        Assert.Equal(0, _excel.ApplyCallCount);
        Assert.Equal("1000", _excel.Sheets["Sheet1"]["B2"]);
        Assert.Equal(switched, _excel.DocumentRef);
    }

    [Fact]
    public void Fake_target_switch_after_preview_denies_before_write()
    {
        var switched = Path.GetFullPath(@"C:\docbridge-identity\switched.dwg");
        _cad.SwitchDocumentRefAfterPreview = switched;
        _cad.LayerVisibility["PLAN"] = true;
        var denied = _host.ApplyOps("cad", CadLayer(NewId()));
        Assert.False(Json.GetBool(denied, "ok"), denied.ToJsonString());
        Assert.Equal(0, _cad.ApplyCallCount);
        Assert.True(_cad.LayerVisibility["PLAN"]);
        Assert.Equal(switched, _cad.DocumentRef);
    }

    [Fact]
    public void Captured_snapshot_identity_mismatch_denies_before_apply()
    {
        _excel.SnapshotDocumentOverride = Path.GetFullPath(@"C:\other\captured.xlsx");
        var denied = _host.ApplyOps("excel", ExcelValues(NewId()));
        Assert.False(Json.GetBool(denied, "ok"));
        Assert.Contains(Json.GetArr(denied, "errors")!, e =>
            e!.GetValue<string>().Contains("captured snapshot document", StringComparison.OrdinalIgnoreCase));
        Assert.Equal(0, _excel.ApplyCallCount);
        Assert.Equal(1, _excel.CaptureSnapshotCallCount);
    }

    [Fact]
    public void Captured_snapshot_missing_identity_is_not_a_match()
    {
        var info = new SnapshotInfo("snap", "now", "excel", null, "execute", "dir");
        var metadata = new JsonObject();
        Assert.False(DocBridgeHost.CapturedSnapshotMatchesExpected(
            "excel", AbsExcel, info, metadata, out var error));
        Assert.Contains("missing documentRef", error, StringComparison.OrdinalIgnoreCase);

        var emptyInfo = info with { DocumentRef = "" };
        var emptyMeta = new JsonObject { ["documentRef"] = "" };
        Assert.False(DocBridgeHost.CapturedSnapshotMatchesExpected(
            "excel", AbsExcel, emptyInfo, emptyMeta, out var emptyError));
        Assert.Contains("missing documentRef", emptyError, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Captured_snapshot_empty_identity_denies_before_apply()
    {
        _excel.OmitSnapshotDocumentRef = true;
        var denied = _host.ApplyOps("excel", ExcelValues(NewId()));
        Assert.False(Json.GetBool(denied, "ok"), denied.ToJsonString());
        Assert.Contains(Json.GetArr(denied, "errors")!, e =>
            e!.GetValue<string>().Contains("missing documentRef", StringComparison.OrdinalIgnoreCase));
        Assert.Equal(0, _excel.ApplyCallCount);
        Assert.Equal(1, _excel.CaptureSnapshotCallCount);
    }

    [Theory]
    [InlineData("status")]
    [InlineData("preview")]
    [InlineData("snapshot")]
    public void Adapter_exception_after_journal_start_is_outcome_unknown_and_not_retried(string stage)
    {
        switch (stage)
        {
            case "status":
                _excel.ThrowOnGetStatus = true;
                break;
            case "preview":
                _excel.ThrowOnPreview = true;
                break;
            default:
                _excel.ThrowOnCaptureSnapshot = true;
                break;
        }

        var id = NewId();
        var denied = _host.ApplyOps("excel", ExcelValues(id));
        Assert.False(Json.GetBool(denied, "ok"), denied.ToJsonString());
        Assert.True(Json.GetBool(denied, "outcomeUnknown"));
        Assert.False(Json.GetBool(denied, "safeToRetry", true));
        Assert.Equal("started", Json.GetString(denied, "journalStatus"));
        Assert.Contains(Json.GetArr(denied, "errors")!, e =>
            e!.GetValue<string>().Contains("outcome unknown", StringComparison.OrdinalIgnoreCase));
        Assert.Equal(0, _excel.ApplyCallCount);
        Assert.Equal(1, _excel.StatusCallCount);
        Assert.Equal(stage == "status" ? 0 : 1, _excel.PreviewCallCount);
        Assert.Equal(stage == "snapshot" ? 1 : 0, _excel.CaptureSnapshotCallCount);

        var journal = new ExecuteJournalService(_home.Options).Lookup(id);
        Assert.Equal(ExecuteJournalState.Started, journal.State);

        var retry = _host.ApplyOps("excel", ExcelValues(id));
        Assert.False(Json.GetBool(retry, "ok"));
        Assert.True(Json.GetBool(retry, "outcomeUnknown"));
        Assert.False(Json.GetBool(retry, "safeToRetry", true));
        Assert.Equal("started", Json.GetString(retry, "journalStatus"));
        Assert.Equal(0, _excel.ApplyCallCount);
        Assert.Equal(1, _excel.StatusCallCount);
        Assert.Equal(stage == "status" ? 0 : 1, _excel.PreviewCallCount);
        Assert.Equal(stage == "snapshot" ? 1 : 0, _excel.CaptureSnapshotCallCount);
    }

    [Fact]
    public void Missing_apply_readback_proof_fails_legacy_and_execute()
    {
        _excel.OmitApplyReadback = true;
        var execute = _host.ApplyOps("excel", ExcelValues(NewId()));
        Assert.False(Json.GetBool(execute, "ok"));
        Assert.Contains(Json.GetArr(execute, "errors")!, e =>
            e!.GetValue<string>().Contains("readback.verified", StringComparison.OrdinalIgnoreCase));
        Assert.True(Json.GetBool(Json.GetObj(execute, "rollback"), "attempted"));

        _fake.OmitApplyReadback = true;
        var requestOps = LegacySetValuesOps();
        var dry = _host.ApplyOps("fake", new JsonObject
        {
            ["ops"] = (JsonNode)requestOps.DeepClone(),
            ["dryRun"] = true,
        });
        Assert.True(Json.GetBool(dry, "ok"), dry.ToJsonString());
        Assert.Null(dry["ops"]);
        var applied = _host.ApplyOps("fake", new JsonObject
        {
            ["ops"] = (JsonNode)requestOps.DeepClone(),
            ["dryRun"] = false,
            ["confirmToken"] = Json.GetString(dry, "confirmToken"),
        });
        Assert.False(Json.GetBool(applied, "ok"));
        Assert.Contains(Json.GetArr(applied, "errors")!, e =>
            e!.GetValue<string>().Contains("readback.verified", StringComparison.OrdinalIgnoreCase));
        Assert.True(Json.GetBool(Json.GetObj(applied, "rollback"), "attempted"));
    }

    [Theory]
    [InlineData("cad", "draw_entities")]
    [InlineData("cad", "move_entities")]
    [InlineData("cad", "zoom_window")]
    [InlineData("cad", "copy_entities")]
    [InlineData("cad", "activate_document")]
    [InlineData("gstarcad", "draw_entities")]
    [InlineData("gstarcad", "zoom_window")]
    public void Unsupported_cad_execute_denies_before_mutation(string app, string op)
    {
        JsonObject body = op switch
        {
            "draw_entities" => new() { ["op"] = op, ["entities"] = new JsonArray(new JsonObject { ["type"] = "line" }) },
            "move_entities" => new() { ["op"] = op, ["handles"] = new JsonArray("A"), ["dx"] = 1, ["dy"] = 1 },
            "copy_entities" => new() { ["op"] = op, ["handles"] = new JsonArray("A"), ["dx"] = 1, ["dy"] = 1 },
            "zoom_window" => new()
            {
                ["op"] = op,
                ["bounds"] = new JsonObject { ["minX"] = 0, ["minY"] = 0, ["maxX"] = 1, ["maxY"] = 1 },
            },
            _ => new() { ["op"] = op, ["document"] = AbsCad },
        };
        var adapter = app == "cad" ? _cad : _gstarcad;
        var denied = _host.ApplyOps(app, new JsonObject
        {
            ["ops"] = new JsonArray(body),
            ["executionMode"] = "execute",
            ["requestId"] = NewId(),
            ["expectedDocumentRef"] = AbsCad,
        });
        Assert.False(Json.GetBool(denied, "ok"));
        Assert.Equal(0, adapter.ApplyCallCount);
        Assert.Equal(0, adapter.PreviewCallCount);
        Assert.Equal(0, adapter.StatusCallCount);
        Assert.Equal(0, adapter.CaptureSnapshotCallCount);
    }

    [Fact]
    public void Cad_execute_rollback_uses_adapter_snapshot_coverage()
    {
        _cad.FailNextApply = true;
        var result = _host.ApplyOps("cad", CadLayer(NewId()));
        Assert.False(Json.GetBool(result, "ok"));
        var coverage = Json.GetObj(result, "rollbackCoverage");
        Assert.Equal("complete", Json.GetString(coverage, "coverage"));
        Assert.Equal("operation-scoped", Json.GetString(coverage, "snapshotKind"));
        Assert.False(Json.GetBool(coverage, "fullDocumentRestored"));
        var rollback = Json.GetObj(result, "rollback")!;
        Assert.True(Json.GetBool(rollback, "attempted"));
        Assert.True(Json.GetBool(rollback, "verified"));
        Assert.False(Json.GetBool(rollback, "fullDocumentRestored"));
        Assert.Equal("complete", Json.GetString(rollback, "coverage"));
        Assert.DoesNotContain(Json.GetArr(result, "warnings")!, e =>
            e!.GetValue<string>().Contains("first 500", StringComparison.OrdinalIgnoreCase) ||
            e!.GetValue<string>().Contains("first-500", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Cad_legacy_draw_still_uses_token_path()
    {
        var dry = _host.ApplyOps("cad", new JsonObject
        {
            ["ops"] = new JsonArray(new JsonObject
            {
                ["op"] = "draw_entities",
                ["entities"] = new JsonArray(new JsonObject { ["type"] = "line" }),
            }),
            ["dryRun"] = true,
        });
        Assert.True(Json.GetBool(dry, "ok"));
        Assert.NotNull(Json.GetString(dry, "confirmToken"));
    }

    [Fact]
    public void Status_app_filter_keeps_default_all_apps()
    {
        var all = _host.CoreGetStatus();
        Assert.True(Json.GetBool(all, "ok"));
        var apps = Json.GetObj(all, "apps")!;
        Assert.NotNull(Json.GetObj(apps, "excel"));
        Assert.NotNull(Json.GetObj(apps, "hwp"));
        Assert.NotNull(Json.GetObj(apps, "cad"));

        var excel = _host.CoreGetStatus(new JsonObject { ["app"] = "excel" });
        var excelApps = Json.GetObj(excel, "apps")!;
        Assert.NotNull(Json.GetObj(excelApps, "excel"));
        Assert.Null(excelApps["hwp"]);
        Assert.Null(excelApps["cad"]);

        var unknown = _host.CoreGetStatus(new JsonObject { ["app"] = "word" });
        Assert.False(Json.GetBool(unknown, "ok"));
    }

    [Fact]
    public void Capabilities_expose_execute_guidance()
    {
        var caps = _host.CoreGetCapabilities(new JsonObject { ["app"] = "cad" });
        var cad = Json.GetObj(Json.GetObj(caps, "apps"), "cad")!;
        var auto = Json.GetArr(cad, "autoExecuteOps")!.Select(n => n!.GetValue<string>()).ToHashSet();
        Assert.True(auto.SetEquals(new[] { "set_text_value", "set_layer_visibility", "set_layer_color", "regen_document" }));
        var coverage = Json.GetObj(Json.GetObj(cad, "execution"), "rollbackCoverage");
        Assert.False(Json.GetBool(coverage, "fullDocumentRestored"));
        Assert.Equal("adapter-snapshot-restore", Json.GetString(coverage, "coverage"));
        Assert.Equal("Adapter-scoped snapshot/restore for the requested CAD execute ops.", Json.GetString(coverage, "note"));
        Assert.DoesNotContain("first-500", coverage!.ToJsonString(), StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("first 500", coverage.ToJsonString(), StringComparison.OrdinalIgnoreCase);

        var excelCaps = _host.CoreGetCapabilities(new JsonObject { ["app"] = "excel" });
        var excelCoverage = Json.GetObj(Json.GetObj(Json.GetObj(Json.GetObj(excelCaps, "apps"), "excel"), "execution"), "rollbackCoverage");
        Assert.False(Json.GetBool(excelCoverage, "fullDocumentRestored"));
    }

    [Fact]
    public void Corrupt_completed_journal_fails_closed()
    {
        var id = Guid.NewGuid().ToString("D");
        var path = new ExecuteJournalService(_home.Options).PathFor(id);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, "{ \"status\": \"completed\", \"app\": \"excel\" }");

        var denied = _host.ApplyOps("excel", ExcelValues(id));
        Assert.False(Json.GetBool(denied, "ok"));
        Assert.Contains(Json.GetArr(denied, "errors")!, e =>
            e!.GetValue<string>().Contains("fail closed", StringComparison.OrdinalIgnoreCase));
        Assert.Equal(0, _excel.ApplyCallCount);
    }

    [Fact]
    public void Tampered_completed_empty_result_fails_closed_without_adapter_or_retry()
    {
        var id = NewId();
        var first = _host.ApplyOps("excel", ExcelValues(id));
        Assert.True(Json.GetBool(first, "ok"), first.ToJsonString());
        var applyAfterFirst = _excel.ApplyCallCount;

        var path = new ExecuteJournalService(_home.Options).PathFor(id);
        var node = JsonNode.Parse(File.ReadAllText(path))!.AsObject();
        node["result"] = new JsonObject();
        File.WriteAllText(path, node.ToJsonString());

        using var restarted = new DocBridgeHost(_home.Options);
        var denied = restarted.ApplyOps("excel", ExcelValues(id));
        Assert.False(Json.GetBool(denied, "ok"));
        Assert.True(Json.GetBool(denied, "outcomeUnknown"));
        Assert.False(Json.GetBool(denied, "safeToRetry", true));
        Assert.Equal("unreadable", Json.GetString(denied, "journalStatus"));
        Assert.Equal(applyAfterFirst, _excel.ApplyCallCount);
        Assert.Equal("1500", _excel.Sheets["Sheet1"]["B2"]);
    }

    [Fact]
    public void Journal_create_new_rejects_collision()
    {
        var journal = new ExecuteJournalService(_home.Options);
        var id = Guid.NewGuid().ToString("D");
        Assert.True(journal.TryBegin(id, "excel", "doc", "hash"));
        Assert.False(journal.TryBegin(id, "excel", "doc", "hash"));
    }
}

public class ExecuteJournalServiceTests : IDisposable
{
    private readonly TestHome _home = new();
    public void Dispose() => _home.Dispose();

    private static JsonObject HostResult(string requestId, bool ok = true) => new()
    {
        ["ok"] = ok,
        ["dryRun"] = false,
        ["executionMode"] = OperationValidator.ExecutionModeExecute,
        ["requestId"] = requestId,
    };

    private void WriteJournal(string id, JsonObject payload)
    {
        var journal = new ExecuteJournalService(_home.Options);
        var path = journal.PathFor(id);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, payload.ToJsonString());
    }

    [Fact]
    public void Completed_result_roundtrip_and_missing_result_is_unreadable()
    {
        var journal = new ExecuteJournalService(_home.Options);
        var id = Guid.NewGuid().ToString("D");
        Assert.Equal(ExecuteJournalState.Missing, journal.Lookup(id).State);
        Assert.True(journal.TryBegin(id, "excel", "doc", "hash"));
        var started = journal.Lookup(id);
        Assert.Equal(ExecuteJournalState.Started, started.State);
        Assert.Equal("unknown", started.Outcome);
        Assert.True(journal.TryComplete(id, "excel", "doc", "hash", HostResult(id)));
        var done = journal.Lookup(id);
        Assert.Equal(ExecuteJournalState.Completed, done.State);
        Assert.True(Json.GetBool(done.Result, "ok"));
        Assert.True(journal.Matches(done, "excel", "doc", "hash"));
        Assert.False(journal.Matches(done, "hwp", "doc", "hash"));
    }

    [Fact]
    public void Path_uuid_must_equal_normalized_stored_request_id()
    {
        var journal = new ExecuteJournalService(_home.Options);
        var pathId = Guid.NewGuid().ToString("D");
        var otherId = Guid.NewGuid().ToString("D");
        WriteJournal(pathId, new JsonObject
        {
            ["requestId"] = otherId,
            ["app"] = "excel",
            ["documentRef"] = "doc",
            ["opsHash"] = "hash",
            ["status"] = "completed",
            ["result"] = HostResult(otherId),
        });
        Assert.Equal(ExecuteJournalState.Unreadable, journal.Lookup(pathId).State);

        WriteJournal(pathId, new JsonObject
        {
            ["app"] = "excel",
            ["documentRef"] = "doc",
            ["opsHash"] = "hash",
            ["status"] = "started",
            ["outcome"] = "unknown",
        });
        Assert.Equal(ExecuteJournalState.Unreadable, journal.Lookup(pathId).State);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("{}")]
    [InlineData("{\"ok\":true}")]
    [InlineData("{\"ok\":true,\"executionMode\":\"execute\"}")]
    [InlineData("{\"ok\":1,\"executionMode\":\"execute\",\"requestId\":\"x\"}")]
    public void Completed_result_must_be_host_identity_contract(string? resultJson)
    {
        var journal = new ExecuteJournalService(_home.Options);
        var id = Guid.NewGuid().ToString("D");
        JsonNode? resultNode = resultJson is null ? null : JsonNode.Parse(resultJson);
        if (resultNode is JsonObject obj && Json.GetString(obj, "requestId") == "x")
            obj["requestId"] = id;
        WriteJournal(id, new JsonObject
        {
            ["requestId"] = id,
            ["app"] = "excel",
            ["documentRef"] = "doc",
            ["opsHash"] = "hash",
            ["status"] = "completed",
            ["result"] = resultNode?.DeepClone(),
        });
        Assert.Equal(ExecuteJournalState.Unreadable, journal.Lookup(id).State);
    }

    [Fact]
    public void Empty_documentRef_is_unreadable_and_cannot_complete()
    {
        var journal = new ExecuteJournalService(_home.Options);
        var id = Guid.NewGuid().ToString("D");
        Assert.False(journal.TryBegin(id, "excel", "", "hash"));
        WriteJournal(id, new JsonObject
        {
            ["requestId"] = id,
            ["app"] = "excel",
            ["documentRef"] = "",
            ["opsHash"] = "hash",
            ["status"] = "started",
            ["outcome"] = "unknown",
        });
        Assert.Equal(ExecuteJournalState.Unreadable, journal.Lookup(id).State);
        Assert.False(journal.TryComplete(id, "excel", "doc", "hash", HostResult(id)));
    }

    [Fact]
    public void Complete_requires_matching_started_record()
    {
        var journal = new ExecuteJournalService(_home.Options);
        var id = Guid.NewGuid().ToString("D");
        Assert.True(journal.TryBegin(id, "excel", "doc-a", "hash-a"));
        Assert.False(journal.TryComplete(id, "excel", "doc-b", "hash-a", HostResult(id)));
        Assert.False(journal.TryComplete(id, "hwp", "doc-a", "hash-a", HostResult(id)));
        Assert.False(journal.TryComplete(id, "excel", "doc-a", "hash-a", new JsonObject()));
        Assert.False(journal.TryComplete(id, "excel", "doc-a", "hash-a", new JsonObject { ["ok"] = true }));
        Assert.Equal(ExecuteJournalState.Started, journal.Lookup(id).State);
        Assert.True(journal.TryComplete(id, "excel", "doc-a", "hash-a", HostResult(id, ok: false)));
        Assert.Equal(ExecuteJournalState.Completed, journal.Lookup(id).State);
    }

    [Fact]
    public void Stored_uuid_alternate_format_matches_filename()
    {
        var journal = new ExecuteJournalService(_home.Options);
        var guid = Guid.NewGuid();
        var id = guid.ToString("D");
        WriteJournal(id, new JsonObject
        {
            ["requestId"] = guid.ToString("N"),
            ["app"] = "excel",
            ["documentRef"] = "doc",
            ["opsHash"] = "hash",
            ["status"] = "completed",
            ["result"] = HostResult(guid.ToString("B")),
        });
        var found = journal.Lookup(id);
        Assert.Equal(ExecuteJournalState.Completed, found.State);
        Assert.True(Json.GetBool(found.Result, "ok"));
    }

    [Fact]
    public void Tampered_result_requestId_is_unreadable()
    {
        var journal = new ExecuteJournalService(_home.Options);
        var id = Guid.NewGuid().ToString("D");
        var other = Guid.NewGuid().ToString("D");
        WriteJournal(id, new JsonObject
        {
            ["requestId"] = id,
            ["app"] = "excel",
            ["documentRef"] = "doc",
            ["opsHash"] = "hash",
            ["status"] = "completed",
            ["result"] = HostResult(other),
        });
        Assert.Equal(ExecuteJournalState.Unreadable, journal.Lookup(id).State);
    }

    [Fact]
    public void Normalize_collapses_uuid_formats()
    {
        var guid = Guid.Parse("aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee");
        Assert.True(ExecuteJournalService.TryNormalizeRequestId(guid.ToString("N"), out var n));
        Assert.True(ExecuteJournalService.TryNormalizeRequestId(guid.ToString("B"), out var b));
        Assert.Equal(n, b);
        Assert.Equal("aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee", n);
        Assert.False(ExecuteJournalService.TryNormalizeRequestId("not-a-uuid", out _));
        Assert.False(ExecuteJournalService.TryNormalizeRequestId(Guid.Empty.ToString("D"), out _));
    }
}
