using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

/// <summary>
/// Real CLI and MCP hosts reach Excel through <see cref="ExcelWorkerAdapter"/>, not through
/// <c>ExcelAdapter</c> directly. In-process, <c>CaptureSnapshot</c> mutates the caller's metadata
/// object; across the worker boundary the adapter mutates a deserialized copy inside the child
/// process. Before the fix the client discarded the worker's response and the worker returned only
/// <c>{captured:true}</c>, so on the production path every key the capture wrote — payload, the
/// authoritative documentRef the adapter resolved, restoreMode, formatFingerprint, the
/// auxiliary-backup keys — was silently lost, and <c>metadata.json</c> recorded only what the host
/// already knew.
///
/// That divergence reaches a safety decision: the host passes snapshot metadata to
/// <c>IsVerifiedRestore</c>, which refuses a verified rollback when the metadata marks partial
/// coverage. A capture that marked partial coverage in-process was reported as fully covered
/// through the worker.
///
/// These cases pin the merge: adapter keys arrive, host-owned keys stay host-owned, and anything
/// malformed fails closed rather than silently not merging — silent non-merge being the defect
/// itself. No worker process is started.
/// </summary>
public class ExcelWorkerSnapshotMetadataTests
{
    // What SnapshotService.Create writes before the adapter runs.
    private static JsonObject HostMetadata() => new()
    {
        ["snapshotId"] = "20260910-101112-abcdef01",
        ["createdAt"] = "2026-09-10T10:11:12.0000000+09:00",
        ["app"] = "excel",
        ["documentRef"] = @"C:\books\host-guess.xlsx",
        ["reason"] = "excel_apply_ops execute",
    };

    private static JsonObject Response(JsonObject captured) =>
        new() { ["captured"] = true, ["metadata"] = captured };

    // ------------------------------------------------------------------ merging

    [Fact]
    public void Adapter_written_keys_reach_the_callers_metadata()
    {
        var metadata = HostMetadata();
        var captured = HostMetadata();
        captured["payload"] = "format-only state.json";
        captured["restoreMode"] = "format-only";
        captured["formatFingerprint"] = "9f2b" + new string('0', 60);
        captured["workbookBackupSource"] = "last-saved-file";
        captured["workbookBackupFresh"] = false;

        ExcelWorkerAdapter.MergeCapturedSnapshotMetadata(metadata, Response(captured));

        Assert.Equal("format-only state.json", Json.GetString(metadata, "payload"));
        Assert.Equal("format-only", Json.GetString(metadata, "restoreMode"));
        Assert.Equal("9f2b" + new string('0', 60), Json.GetString(metadata, "formatFingerprint"));
        Assert.Equal("last-saved-file", Json.GetString(metadata, "workbookBackupSource"));
        Assert.False(Json.GetBool(metadata, "workbookBackupFresh"));
    }

    [Fact]
    public void Adapter_documentRef_overwrites_the_hosts_pre_seeded_guess()
    {
        // In-process the adapter's FullName wins, because it names the workbook actually
        // resolved — which ResolveTargetWorkbook may take from another Excel instance. The
        // worker path must not quietly keep the host's guess instead.
        var metadata = HostMetadata();
        var captured = HostMetadata();
        captured["documentRef"] = @"C:\books\actually-resolved.xlsx";

        ExcelWorkerAdapter.MergeCapturedSnapshotMetadata(metadata, Response(captured));

        Assert.Equal(@"C:\books\actually-resolved.xlsx", Json.GetString(metadata, "documentRef"));
    }

    [Fact]
    public void Partial_coverage_marking_survives_the_worker_boundary()
    {
        // The host feeds snapshot metadata to IsVerifiedRestore; losing this key is the
        // difference between refusing and claiming a verified rollback.
        var metadata = HostMetadata();
        var captured = HostMetadata();
        captured["snapshotCoverage"] = "partial";

        ExcelWorkerAdapter.MergeCapturedSnapshotMetadata(metadata, Response(captured));
        Assert.Equal("partial", Json.GetString(metadata, "snapshotCoverage"));

        // The consequence, not just the key: the same restore result flips from "verified" to
        // "not verified" purely on whether this metadata reached the host.
        var restored = new JsonObject
        {
            ["ok"] = true,
            ["readback"] = new JsonObject { ["verified"] = true },
        };
        Assert.True(DocBridgeHost.IsVerifiedRestore(restored, HostMetadata()));
        Assert.False(DocBridgeHost.IsVerifiedRestore(restored, metadata));
    }

    [Fact]
    public void Nested_values_are_cloned_not_aliased()
    {
        var metadata = HostMetadata();
        var captured = HostMetadata();
        captured["workbookBackupIdentity"] = new JsonObject
        {
            ["before"] = new JsonObject { ["saved"] = false },
        };

        var response = Response(captured);
        ExcelWorkerAdapter.MergeCapturedSnapshotMetadata(metadata, response);

        // Mutating the response afterwards must not reach the merged metadata.
        ((JsonObject)((JsonObject)captured["workbookBackupIdentity"]!)["before"]!)["saved"] = true;

        var merged = Json.GetObj(Json.GetObj(metadata, "workbookBackupIdentity"), "before")!;
        Assert.False(Json.GetBool(merged, "saved"));
    }

    [Fact]
    public void Keys_the_response_omits_are_left_alone()
    {
        var metadata = HostMetadata();
        metadata["hostOnlyNote"] = "keep me";
        var captured = HostMetadata();
        captured["payload"] = "format-only state.json";

        ExcelWorkerAdapter.MergeCapturedSnapshotMetadata(metadata, Response(captured));

        Assert.Equal("keep me", Json.GetString(metadata, "hostOnlyNote"));
        Assert.Equal("format-only state.json", Json.GetString(metadata, "payload"));
    }

    [Fact]
    public void A_null_valued_adapter_key_is_merged_as_json_null_not_dropped()
    {
        // "read and unreadable" is a distinct state from "not evaluated" in the auxiliary
        // backup contract; the boundary must not collapse the two.
        var metadata = HostMetadata();
        var captured = HostMetadata();
        captured["workbookBackupSavedFlag"] = null;

        ExcelWorkerAdapter.MergeCapturedSnapshotMetadata(metadata, Response(captured));

        Assert.True(metadata.ContainsKey("workbookBackupSavedFlag"));
        Assert.Null(metadata["workbookBackupSavedFlag"]);
    }

    // ------------------------------------------------------------- failing closed

    [Fact]
    public void A_response_without_an_explicit_captured_flag_fails_closed()
    {
        var metadata = HostMetadata();
        var response = new JsonObject { ["metadata"] = HostMetadata() };

        var ex = Assert.Throws<InvalidDataException>(() =>
            ExcelWorkerAdapter.MergeCapturedSnapshotMetadata(metadata, response));

        Assert.Contains("no boolean 'captured' flag", ex.Message);
    }

    [Fact]
    public void A_response_reporting_captured_false_fails_closed()
    {
        // Carrying a metadata object is not the same as having captured a snapshot.
        var metadata = HostMetadata();
        var captured = HostMetadata();
        captured["payload"] = "format-only state.json";
        var response = new JsonObject { ["captured"] = false, ["metadata"] = captured };

        var ex = Assert.Throws<InvalidDataException>(() =>
            ExcelWorkerAdapter.MergeCapturedSnapshotMetadata(metadata, response));

        Assert.Contains("captured=false", ex.Message);
        Assert.False(metadata.ContainsKey("payload"));
    }

    public static TheoryData<JsonNode?> NonBooleanCapturedFlags() => new()
    {
        { null },
        { JsonValue.Create("true") },
        { JsonValue.Create(1) },
        { new JsonObject() },
        { new JsonArray() },
    };

    [Theory]
    [MemberData(nameof(NonBooleanCapturedFlags))]
    public void A_captured_flag_that_is_not_a_boolean_fails_closed(JsonNode? flag)
    {
        var metadata = HostMetadata();
        var response = new JsonObject { ["captured"] = flag, ["metadata"] = HostMetadata() };

        Assert.Throws<InvalidDataException>(() =>
            ExcelWorkerAdapter.MergeCapturedSnapshotMetadata(metadata, response));
    }

    [Fact]
    public void A_response_without_metadata_fails_closed()
    {
        var metadata = HostMetadata();

        var ex = Assert.Throws<InvalidDataException>(() =>
            ExcelWorkerAdapter.MergeCapturedSnapshotMetadata(
                metadata, new JsonObject { ["captured"] = true }));

        Assert.Contains("did not return the captured metadata", ex.Message);
    }

    public static TheoryData<JsonNode?> MalformedMetadataMembers() => new()
    {
        { null },
        { JsonValue.Create("format-only state.json") },
        { JsonValue.Create(true) },
        { new JsonArray() },
    };

    [Theory]
    [MemberData(nameof(MalformedMetadataMembers))]
    public void A_metadata_member_that_is_not_an_object_fails_closed(JsonNode? malformed)
    {
        var metadata = HostMetadata();
        var response = new JsonObject { ["captured"] = true, ["metadata"] = malformed };

        Assert.Throws<InvalidDataException>(() =>
            ExcelWorkerAdapter.MergeCapturedSnapshotMetadata(metadata, response));
    }

    [Theory]
    [InlineData("snapshotId")]
    [InlineData("app")]
    [InlineData("createdAt")]
    [InlineData("reason")]
    public void A_worker_that_changes_a_host_owned_key_fails_closed(string key)
    {
        var metadata = HostMetadata();
        var captured = HostMetadata();
        captured[key] = "tampered";

        var ex = Assert.Throws<InvalidDataException>(() =>
            ExcelWorkerAdapter.MergeCapturedSnapshotMetadata(metadata, Response(captured)));

        Assert.Contains($"changed host-owned metadata '{key}'", ex.Message);
        // The caller's object must not be left half-merged with a tampered identity.
        Assert.NotEqual("tampered", Json.GetString(metadata, key));
    }

    [Theory]
    [InlineData("snapshotId")]
    [InlineData("app")]
    [InlineData("createdAt")]
    [InlineData("reason")]
    public void A_worker_that_drops_a_host_owned_key_fails_closed(string key)
    {
        var metadata = HostMetadata();
        var captured = HostMetadata();
        captured.Remove(key);

        var ex = Assert.Throws<InvalidDataException>(() =>
            ExcelWorkerAdapter.MergeCapturedSnapshotMetadata(metadata, Response(captured)));

        Assert.Contains($"dropped host-owned metadata '{key}'", ex.Message);
    }

    [Fact]
    public void A_host_owned_key_the_caller_never_set_is_not_required_back()
    {
        // SnapshotService writes documentRef only when the host resolved one; a key the caller
        // never had cannot be "dropped".
        var metadata = new JsonObject { ["snapshotId"] = "20260910-101112-abcdef01" };
        var captured = new JsonObject
        {
            ["snapshotId"] = "20260910-101112-abcdef01",
            ["payload"] = "format-only state.json",
        };

        ExcelWorkerAdapter.MergeCapturedSnapshotMetadata(metadata, Response(captured));

        Assert.Equal("format-only state.json", Json.GetString(metadata, "payload"));
        Assert.False(metadata.ContainsKey("reason"));
    }
}
