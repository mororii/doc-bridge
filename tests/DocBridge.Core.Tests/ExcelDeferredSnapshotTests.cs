using System.IO.Compression;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

/// <summary>
/// Contract tests for the opt-in Excel deferred format checkpoint.
///
/// Every case drives a real production predicate — the policy parser, the request and workbook
/// eligibility evaluators, the byte preflight over real ZIP/OLE2 bytes, the checkpoint path
/// resolver, the restore binding validator over a real state.json on disk, and the envelope
/// recognisers. Nothing here asserts a constant this file invented.
///
/// The wire envelope is `restoreMode=copy-sheet-topology` + `snapshotVersion=4` +
/// `payloadMode=deferred-format-only`. That combination is deliberate: 0.4.18 has no format-only
/// decoder at all, but both 0.4.18 and 0.4.20 do decode copy-sheet-topology and refuse a version
/// they do not know before writing anything.
///
/// No DocBridgeHost, no ExcelAdapter instance, no Excel: nothing here takes
/// Global\DocBridge.Automation. Native coverage lives in ExcelDeferredHostHarnessE2ETests.
/// </summary>
public class ExcelDeferredSnapshotTests
{
    // ------------------------------------------------------------- policy gating

    [Theory]
    [InlineData("1", true)]
    [InlineData(null, false)]
    [InlineData("", false)]
    [InlineData("0", false)]
    [InlineData("01", false)]
    [InlineData(" 1", false)]
    [InlineData("1 ", false)]
    [InlineData("true", false)]
    [InlineData("True", false)]
    [InlineData("yes", false)]
    [InlineData("on", false)]
    public void Only_the_exact_value_one_enables_deferred(string? raw, bool expected)
    {
        Assert.Equal(
            expected,
            ExcelAdapter.ExcelDeferredFormatSnapshotPolicy.FromSetting(raw).Enabled);
    }

    [Fact]
    public void The_policy_reads_the_documented_variable_without_mutating_the_environment()
    {
        var current = Environment.GetEnvironmentVariable(ExcelAdapter.DeferredFormatSnapshotVariable);
        Assert.Equal(
            ExcelAdapter.ExcelDeferredFormatSnapshotPolicy.FromSetting(current),
            ExcelAdapter.ExcelDeferredFormatSnapshotPolicy.FromEnvironment());
        Assert.Equal(
            "DOCBRIDGE_EXCEL_DEFERRED_FORMAT_SNAPSHOT", ExcelAdapter.DeferredFormatSnapshotVariable);
    }

    // -------------------------------------------------------- the wire envelope

    [Fact]
    public void The_wire_envelope_is_a_mode_old_binaries_decode_at_a_version_they_refuse()
    {
        // 0.4.18 has no format-only decoder, so a format-only envelope would fall through to its
        // legacy branch and could report a zero-work success. copy-sheet-topology is decoded by
        // both 0.4.18 and 0.4.20, and both compare the version before restoring anything.
        Assert.Equal("copy-sheet-topology", ExcelAdapter.DeferredFormatCompatibilityRestoreMode);
        Assert.Equal(4, ExcelAdapter.DeferredFormatSnapshotVersion);
        Assert.Equal("deferred-format-only", ExcelAdapter.DeferredFormatPayloadMode);
        Assert.Equal("deferred-format-only", ExcelAdapter.DeferredFormatRestoreMode);
    }

    [Fact]
    public void Only_the_complete_triple_is_recognised_as_a_deferred_envelope()
    {
        Assert.True(ExcelAdapter.IsDeferredFormatEnvelope(EnvelopeState()));

        foreach (var (key, value) in new (string, JsonNode?)[]
                 {
                     ("restoreMode", "format-only"),
                     ("restoreMode", "deferred-format-only"),
                     ("snapshotVersion", 2),
                     ("snapshotVersion", 3),
                     ("payloadMode", "something-else"),
                     ("payloadMode", null),
                 })
        {
            var state = EnvelopeState();
            state[key] = value;
            Assert.False(
                ExcelAdapter.IsDeferredFormatEnvelope(state),
                $"{key}={value?.ToJsonString() ?? "null"} must not be accepted as the envelope");
        }

        Assert.False(ExcelAdapter.IsDeferredFormatEnvelope(null));
    }

    [Fact]
    public void A_half_written_envelope_is_recognised_as_incomplete_rather_than_ordinary_topology()
    {
        // Fail closed: a state carrying the deferred version or payloadMode but not the full
        // triple must never be handed to the normal copy-sheet-topology restore.
        var versionOnly = EnvelopeState();
        versionOnly["payloadMode"] = null;
        Assert.False(ExcelAdapter.IsDeferredFormatEnvelope(versionOnly));
        Assert.True(ExcelAdapter.LooksLikeIncompleteDeferredFormatEnvelope(versionOnly));

        var payloadOnly = EnvelopeState();
        payloadOnly["snapshotVersion"] = 2;
        Assert.True(ExcelAdapter.LooksLikeIncompleteDeferredFormatEnvelope(payloadOnly));

        // A genuine, complete envelope is not "incomplete", and neither is ordinary topology v2.
        Assert.False(ExcelAdapter.LooksLikeIncompleteDeferredFormatEnvelope(EnvelopeState()));
        Assert.False(ExcelAdapter.LooksLikeIncompleteDeferredFormatEnvelope(new JsonObject
        {
            ["restoreMode"] = "copy-sheet-topology",
            ["snapshotVersion"] = 2,
        }));
    }

    // ----------------------------------------------- request eligibility (real SUT)

    [Fact]
    public void An_eligible_execute_request_is_accepted_and_reports_its_decision_inputs()
    {
        var result = ExcelAdapter.EvaluateDeferredFormatRequestEligibility(
            Enabled, ExecuteMetadata(), new[] { FillOp("A1:J100") });

        Assert.True(Json.GetBool(result, "eligible"), result.ToJsonString());
        Assert.Equal("ok", Json.GetString(result, "code"));
        Assert.Equal(1000, Json.GetInt(result, "targetCells"));
        Assert.True(Json.GetBool(result, "hasFillColor"));
        Assert.Equal("Sheet1", Json.GetString(result, "sheet"));
    }

    [Fact]
    public void Deferred_is_refused_when_the_policy_is_off()
    {
        AssertRefused(
            ExcelAdapter.EvaluateDeferredFormatRequestEligibility(
                Disabled, ExecuteMetadata(), new[] { FillOp("A1:J100") }),
            "policy");
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData("dry-run")]
    [InlineData("Execute")]
    [InlineData("EXECUTE")]
    public void Deferred_is_refused_unless_the_host_stamped_the_execute_context(string? context)
    {
        // The dry-run capture site must never produce a deferred snapshot: a later confirm calls
        // ValidatePreviewReuse, which answers freshPreviewAllowed=false on a mode mismatch, and
        // the user's confirmed apply would be denied. Matched ordinally: "Execute" is not it.
        var metadata = ExecuteMetadata();
        if (context is null) metadata.Remove(HostSnapshotCaptureContext.MetadataKey);
        else metadata[HostSnapshotCaptureContext.MetadataKey] = context;

        AssertRefused(
            ExcelAdapter.EvaluateDeferredFormatRequestEligibility(
                Enabled, metadata, new[] { FillOp("A1:J100") }),
            "host-context");
    }

    [Fact]
    public void The_execute_context_key_and_value_are_the_ones_the_host_stamps()
    {
        Assert.Equal("hostSnapshotContext", HostSnapshotCaptureContext.MetadataKey);
        Assert.Equal("execute", HostSnapshotCaptureContext.Execute);
    }

    [Fact]
    public void A_batch_that_is_not_a_single_format_range_is_refused()
    {
        AssertRefused(
            ExcelAdapter.EvaluateDeferredFormatRequestEligibility(Enabled, ExecuteMetadata(), null),
            "ops");
        AssertRefused(
            ExcelAdapter.EvaluateDeferredFormatRequestEligibility(
                Enabled, ExecuteMetadata(), Array.Empty<JsonObject>()),
            "ops");
        AssertRefused(
            ExcelAdapter.EvaluateDeferredFormatRequestEligibility(
                Enabled, ExecuteMetadata(), new[] { FillOp("A1:J100"), FillOp("A1:J100") }),
            "ops");
        AssertRefused(
            ExcelAdapter.EvaluateDeferredFormatRequestEligibility(
                Enabled, ExecuteMetadata(),
                new[] { FillOp("A1:J100"), new JsonObject { ["op"] = "set_values" } }),
            "ops");
    }

    [Theory]
    [InlineData("A1:J100,L1:L9")]
    [InlineData("A1:J100;L1:L9")]
    [InlineData("A1:J100 L1:L9")]
    [InlineData("not-a-range")]
    [InlineData("")]
    public void A_target_that_is_not_one_rectangle_is_refused(string range)
    {
        AssertRefused(
            ExcelAdapter.EvaluateDeferredFormatRequestEligibility(
                Enabled, ExecuteMetadata(), new[] { FillOp(range) }),
            "sheet-range");
    }

    [Fact]
    public void A_qualified_range_is_accepted_only_when_it_agrees_with_the_target_sheet()
    {
        var agreeing = ExcelAdapter.EvaluateDeferredFormatRequestEligibility(
            Enabled, ExecuteMetadata(), new[] { FillOp("Sheet1!A1:J100") });
        Assert.True(Json.GetBool(agreeing, "eligible"), agreeing.ToJsonString());

        // A qualification naming a different sheet than target.sheet is ambiguous about which
        // sheet is actually being formatted, so it is refused rather than silently preferred.
        AssertRefused(
            ExcelAdapter.EvaluateDeferredFormatRequestEligibility(
                Enabled, ExecuteMetadata(), new[] { FillOp("Other!A1:J100") }),
            "sheet-range");
    }

    [Fact]
    public void A_missing_target_sheet_is_refused()
    {
        var op = FillOp("A1:J100");
        op.Remove("target");
        AssertRefused(
            ExcelAdapter.EvaluateDeferredFormatRequestEligibility(Enabled, ExecuteMetadata(), new[] { op }),
            "sheet-range");
    }

    [Fact]
    public void A_batch_without_fillColor_stays_on_the_existing_eager_fast_path()
    {
        // Bold-only is already cheap through the uniform fast path; paying a whole-workbook copy
        // for it would be a straight loss.
        var boldOnly = FillOp("A1:J100");
        boldOnly["style"] = new JsonObject { ["bold"] = true };
        AssertRefused(
            ExcelAdapter.EvaluateDeferredFormatRequestEligibility(
                Enabled, ExecuteMetadata(), new[] { boldOnly }),
            "style");
    }

    [Fact]
    public void A_written_key_outside_the_supported_set_is_refused()
    {
        var op = FillOp("A1:J100");
        op["style"] = new JsonObject { ["fillColor"] = 16711680, ["fontColor"] = 255 };
        AssertRefused(
            ExcelAdapter.EvaluateDeferredFormatRequestEligibility(Enabled, ExecuteMetadata(), new[] { op }),
            "style");
    }

    [Theory]
    [InlineData("fontName", "돋움")]
    [InlineData("wrapText", true)]
    [InlineData("horizontalAlign", "center")]
    public void New_format_keys_are_not_deferred_eligible(string key, object value)
    {
        var op = FillOp("A1:J100");
        var style = new JsonObject { ["fillColor"] = 16711680 };
        style[key] = value switch
        {
            bool flag => JsonValue.Create(flag),
            string text => JsonValue.Create(text),
            _ => JsonValue.Create(value.ToString()),
        };
        op["style"] = style;
        AssertRefused(
            ExcelAdapter.EvaluateDeferredFormatRequestEligibility(Enabled, ExecuteMetadata(), new[] { op }),
            "style");
    }

    [Fact]
    public void Supported_extra_style_keys_are_still_accepted_alongside_fillColor()
    {
        var op = FillOp("A1:J100");
        op["style"] = new JsonObject
        {
            ["fillColor"] = 16711680,
            ["bold"] = true,
            ["italic"] = false,
            ["fontSize"] = 11.0,
            ["numberFormat"] = "General",
        };
        var result = ExcelAdapter.EvaluateDeferredFormatRequestEligibility(
            Enabled, ExecuteMetadata(), new[] { op });
        Assert.True(Json.GetBool(result, "eligible"), result.ToJsonString());
    }

    [Theory]
    [InlineData("A1:J99", 990)]
    [InlineData("A1:A1", 1)]
    [InlineData("A1:BZ100", 7800)]
    public void A_cell_count_outside_the_rollout_bounds_is_refused(string range, int cells)
    {
        var result = ExcelAdapter.EvaluateDeferredFormatRequestEligibility(
            Enabled, ExecuteMetadata(), new[] { FillOp(range) });
        AssertRefused(result, "cell-count");
        Assert.Equal(cells, Json.GetInt(result, "targetCells"));
    }

    [Fact]
    public void The_rollout_bounds_are_1000_to_5000_cells_and_16_MiB_inclusive()
    {
        Assert.Equal(1000, ExcelAdapter.DeferredFormatMinCells);
        Assert.Equal(5000, ExcelAdapter.DeferredFormatMaxCells);
        Assert.Equal(16L * 1024 * 1024, ExcelAdapter.DeferredFormatMaxFileBytes);
        Assert.Equal(16_777_216L, ExcelAdapter.DeferredFormatMaxFileBytes);

        // Both ends inclusive, exercised through the real evaluator.
        foreach (var cells in new[] { ExcelAdapter.DeferredFormatMinCells, ExcelAdapter.DeferredFormatMaxCells })
        {
            var result = ExcelAdapter.EvaluateDeferredFormatRequestEligibility(
                Enabled, ExecuteMetadata(), new[] { FillOp(RangeOfCells(cells)) });
            Assert.True(Json.GetBool(result, "eligible"), $"{cells} cells: {result.ToJsonString()}");
        }
    }

    [Fact]
    public void The_previously_measured_large_fixture_is_under_the_size_cap()
    {
        // 16,194,298 bytes is the 2026-09-09 fixture. It is under 16 MiB, so size is NOT what
        // excludes it — its formulas are, independently, under the v1 no-formula rule.
        Assert.True(16_194_298L <= ExcelAdapter.DeferredFormatMaxFileBytes);
    }

    // ----------------------------------------- workbook eligibility (real SUT, fake COM)

    [Fact]
    public void An_ordinary_local_xlsx_workbook_is_accepted()
    {
        using var home = new TestHome();
        var workbook = NewWorkbook(home);

        var result = ExcelAdapter.EvaluateDeferredFormatWorkbookEligibility(workbook, "Sheet1");

        Assert.True(Json.GetBool(result, "eligible"), result.ToJsonString());
        Assert.Equal(51, Json.GetInt(result, "fileFormat"));
        Assert.NotNull(result["workbookBytes"]);
    }

    [Theory]
    [InlineData("path", "notxlsx")]
    [InlineData("path", "missing")]
    [InlineData("file-format", "format")]
    [InlineData("macros", "vba")]
    [InlineData("protection", "structure")]
    [InlineData("protection", "windows")]
    [InlineData("password", "password")]
    [InlineData("password", "writereserved")]
    [InlineData("links", "links")]
    [InlineData("identity", "identity")]
    public void An_unsupported_workbook_is_refused_with_its_own_code(string code, string defect)
    {
        using var home = new TestHome();
        var workbook = NewWorkbook(home);
        switch (defect)
        {
            case "notxlsx": workbook.FullNameValue = Path.ChangeExtension(workbook.FullNameValue, ".xlsm"); break;
            case "missing": workbook.FullNameValue = Path.Combine(home.Dir, "gone.xlsx"); break;
            case "format": workbook.FileFormatValue = 52; break;
            case "vba": workbook.HasVBProjectValue = true; break;
            case "structure": workbook.ProtectStructureValue = true; break;
            case "windows": workbook.ProtectWindowsValue = true; break;
            case "password": workbook.HasPasswordValue = true; break;
            case "writereserved": workbook.WriteReservedValue = true; break;
            case "links": workbook.HasLinks = true; break;
            case "identity": workbook.SavedThrows = true; break;
        }

        AssertRefused(
            ExcelAdapter.EvaluateDeferredFormatWorkbookEligibility(workbook, "Sheet1"), code);
    }

    [Fact]
    public void An_unreadable_workbook_property_is_refused_rather_than_assumed_false()
    {
        using var home = new TestHome();
        var workbook = NewWorkbook(home);
        workbook.HasVBProjectThrows = true;

        AssertRefused(
            ExcelAdapter.EvaluateDeferredFormatWorkbookEligibility(workbook, "Sheet1"), "macros");
    }

    [Fact]
    public void An_oversized_workbook_is_refused_by_size()
    {
        using var home = new TestHome();
        var workbook = NewWorkbook(home);
        using (var stream = new FileStream(workbook.FullNameValue, FileMode.Open, FileAccess.Write))
            stream.SetLength(ExcelAdapter.DeferredFormatMaxFileBytes + 1);

        var result = ExcelAdapter.EvaluateDeferredFormatWorkbookEligibility(workbook, "Sheet1");
        AssertRefused(result, "file-size");
        Assert.Equal(
            ExcelAdapter.DeferredFormatMaxFileBytes + 1,
            result["workbookBytes"]!.GetValue<long>());
    }

    // ------------------------------------- checkpoint byte preflight (real SUT, real bytes)
    //
    // The preflight is now a bounded OPC/XmlReader walk, so a hand-rolled two-entry zip is no
    // longer a valid input: it would be refused for the wrong reason and prove nothing. Each case
    // starts from one canonical, valid, minimal package and mutates exactly one feature, so the
    // refusal code is attributable to that feature.

    [Fact]
    public void The_canonical_minimal_package_passes_the_byte_preflight()
    {
        using var home = new TestHome();
        var result = ExcelAdapter.PrefightDeferredFormatCheckpoint(WritePackage(home, "canonical.xlsx"));
        Assert.True(Json.GetBool(result, "ok"), result.ToJsonString());
        Assert.Equal("ok", Json.GetString(result, "code"));
    }

    public static TheoryData<string, string> PackageMutations() => new()
    {
        { "strict-ooxml", "strict-namespace" },
        { "formulas", "cell-formula" },
        { "merge-cells", "merge" },
        { "defined-names", "defined-name" },
        { "conditional-formatting", "conditional-format" },
        { "data-validation", "data-validation" },
        { "full-calc-on-load", "full-calc" },
        { "macro-enabled-main", "macro-main" },
        { "vba-project", "vba" },
        { "external-links", "external-rel" },
        { "malformed-xml", "malformed" },
        { "dtd-prohibited", "dtd" },
        { "missing-content-types", "drop-content-types" },
        { "missing-workbook-xml", "drop-workbook" },
    };

    [Theory]
    [MemberData(nameof(PackageMutations))]
    public void One_mutated_feature_is_refused_with_its_own_code(string code, string mutation)
    {
        using var home = new TestHome();
        var path = WritePackage(home, $"{mutation}.xlsx", mutation);

        AssertPreflightRefused(ExcelAdapter.PrefightDeferredFormatCheckpoint(path), code);
    }

    [Fact]
    public void A_missing_or_non_package_checkpoint_is_refused()
    {
        using var home = new TestHome();
        AssertPreflightRefused(
            ExcelAdapter.PrefightDeferredFormatCheckpoint(Path.Combine(home.Dir, "absent.xlsx")),
            "input-missing");

        var garbage = Path.Combine(home.Dir, "garbage.xlsx");
        File.WriteAllBytes(garbage, "not a workbook at all"u8.ToArray());
        AssertPreflightRefused(
            ExcelAdapter.PrefightDeferredFormatCheckpoint(garbage), "unknown-container");
    }

    [Fact]
    public void An_encrypted_checkpoint_is_refused_before_excel_could_prompt_for_a_password()
    {
        // SaveCopyAs of an encrypted workbook produces an encrypted copy, and reopening that in
        // the user's Excel raises a modal password prompt. An encrypted OOXML package is an OLE2
        // compound file, so this is decidable in bytes.
        using var home = new TestHome();
        var encrypted = Path.Combine(home.Dir, "encrypted.xlsx");
        var bytes = new List<byte> { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
        bytes.AddRange(new byte[512]);
        bytes.AddRange(Encoding.Unicode.GetBytes("EncryptedPackage"));
        bytes.AddRange(new byte[512]);
        File.WriteAllBytes(encrypted, bytes.ToArray());

        AssertPreflightRefused(
            ExcelAdapter.PrefightDeferredFormatCheckpoint(encrypted), "encrypted-ooxml");
    }

    [Fact]
    public void A_plain_ole2_compound_file_is_refused_separately_from_an_encrypted_one()
    {
        using var home = new TestHome();
        var legacy = Path.Combine(home.Dir, "legacy.xlsx");
        var bytes = new List<byte> { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
        bytes.AddRange(new byte[1024]);
        File.WriteAllBytes(legacy, bytes.ToArray());

        AssertPreflightRefused(
            ExcelAdapter.PrefightDeferredFormatCheckpoint(legacy), "ole2-compound-file");
    }

    [Fact]
    public void A_duplicate_package_part_is_refused()
    {
        using var home = new TestHome();
        var path = Path.Combine(home.Dir, "duplicate.xlsx");
        var parts = CanonicalPackage().ToList();
        parts.Add(("xl/workbook.xml", parts.First(part => part.Name == "xl/workbook.xml").Content));
        WriteZip(path, parts);

        AssertPreflightRefused(
            ExcelAdapter.PrefightDeferredFormatCheckpoint(path), "duplicate-package-part");
    }

    [Fact]
    public void An_oversized_checkpoint_is_refused_by_size()
    {
        using var home = new TestHome();
        var big = Path.Combine(home.Dir, "big.xlsx");
        using (var stream = new FileStream(big, FileMode.Create, FileAccess.Write))
            stream.SetLength(ExcelAdapter.DeferredFormatMaxFileBytes + 1);

        AssertPreflightRefused(
            ExcelAdapter.PrefightDeferredFormatCheckpoint(big), "file-size");
    }

    [Fact]
    public void The_native_harness_seed_is_plain_bytes_the_real_byte_gate_accepts()
    {
        // The native harness opens this seed instead of Workbooks.Add, because Excel's own
        // new-workbook template embeds the machine's Office Store add-ins into everything it
        // saves. Offline, without Excel, this pins the two properties that make the seed usable:
        // the product's own preflight accepts it, and it carries none of those add-in parts.
        using var home = new TestHome();
        var path = Path.Combine(home.Dir, "seed.xlsx");
        MinimalXlsxSeed.Write(path);

        var preflight = ExcelAdapter.PrefightDeferredFormatCheckpoint(path);
        Assert.True(Json.GetBool(preflight, "ok"), preflight.ToJsonString());

        using var zip = ZipFile.OpenRead(path);
        var names = zip.Entries.Select(entry => entry.FullName).ToList();
        Assert.DoesNotContain(names, name =>
            name.Contains("webextension", StringComparison.OrdinalIgnoreCase)
            || name.Contains("taskpanes", StringComparison.OrdinalIgnoreCase));
    }

    // ------------------------------------------------- checkpoint path resolution

    [Fact]
    public void The_checkpoint_path_is_reconstructed_from_the_snapshot_directory()
    {
        using var home = new TestHome();
        var snapshotDir = NewSnapshotDir(home, out var dirName);
        var expected = $"deferred-format-{dirName}.xlsx";

        Assert.True(ExcelAdapter.TryResolveDeferredFormatCheckpointPath(
            snapshotDir, expected, out var fullPath, out var error), error);
        Assert.Equal(Path.GetFullPath(Path.Combine(snapshotDir, expected)), fullPath);

        // A null recorded name still resolves: the directory alone determines the file.
        Assert.True(ExcelAdapter.TryResolveDeferredFormatCheckpointPath(
            snapshotDir, null, out var derived, out _));
        Assert.Equal(fullPath, derived);
    }

    [Theory]
    [InlineData(@"..\..\evil.xlsx")]
    [InlineData("../../evil.xlsx")]
    [InlineData(@"C:\Windows\Temp\evil.xlsx")]
    [InlineData(@"\\server\share\evil.xlsx")]
    [InlineData("sub/deferred-format-x.xlsx")]
    [InlineData("deferred-format-someone-elses-snapshot.xlsx")]
    public void A_recorded_name_that_is_not_the_reconstructed_stem_is_refused_never_followed(string recorded)
    {
        // Whoever can write state.json must not get to choose which workbook Excel opens.
        using var home = new TestHome();
        var snapshotDir = NewSnapshotDir(home, out _);

        Assert.False(ExcelAdapter.TryResolveDeferredFormatCheckpointPath(
            snapshotDir, recorded, out var fullPath, out var error));
        Assert.False(string.IsNullOrWhiteSpace(error));
        Assert.Equal("", fullPath);
    }

    [Fact]
    public void An_unsafe_or_missing_snapshot_directory_is_refused()
    {
        using var home = new TestHome();
        Assert.False(ExcelAdapter.TryResolveDeferredFormatCheckpointPath(
            Path.Combine(home.Dir, "does-not-exist"), null, out _, out var missing));
        Assert.False(string.IsNullOrWhiteSpace(missing));

        Assert.False(ExcelAdapter.TryResolveDeferredFormatCheckpointPath("", null, out _, out var empty));
        Assert.False(string.IsNullOrWhiteSpace(empty));
    }

    // ------------------------------------- restore binding integrity (real SUT, real files)

    [Fact]
    public void A_fully_bound_metadata_and_state_pair_validates()
    {
        using var home = new TestHome();
        var (metadata, _, statePath) = WriteBoundPair(home);

        var result = ExcelAdapter.ValidateDeferredFormatRestoreBinding(
            metadata, ReadState(statePath), statePath);

        Assert.True(Json.GetBool(result, "ok"), result.ToJsonString());
        Assert.Equal("ok", Json.GetString(result, "code"));
    }

    [Theory]
    [InlineData("restoreMode")]
    [InlineData("snapshotKind")]
    [InlineData("payloadMode")]
    [InlineData("documentRef")]
    [InlineData("deferredCheckpoint")]
    [InlineData("deferredCheckpointSha256")]
    [InlineData("deferredStateSha256")]
    public void Missing_bound_metadata_is_refused(string key)
    {
        using var home = new TestHome();
        var (metadata, _, statePath) = WriteBoundPair(home);
        metadata.Remove(key);

        AssertPreflightRefused(
            ExcelAdapter.ValidateDeferredFormatRestoreBinding(metadata, ReadState(statePath), statePath),
            "metadata-missing");
    }

    [Theory]
    [InlineData("documentRef", "documentRef")]
    [InlineData("deferredCheckpoint", "checkpoint-file")]
    [InlineData("deferredCheckpointSha256", "checkpoint-hash")]
    public void Metadata_that_disagrees_with_state_is_refused(string metadataKey, string code)
    {
        using var home = new TestHome();
        var (metadata, _, statePath) = WriteBoundPair(home);
        metadata[metadataKey] = metadataKey.EndsWith("Sha256", StringComparison.Ordinal)
            ? new string('b', 64)
            : @"C:\books\somewhere-else.xlsx";

        AssertPreflightRefused(
            ExcelAdapter.ValidateDeferredFormatRestoreBinding(metadata, ReadState(statePath), statePath),
            code);
    }

    [Fact]
    public void A_state_json_edited_after_capture_is_refused_by_its_bound_hash()
    {
        // The integrity property that matters: the state file cannot be retargeted after the fact,
        // because metadata carries its SHA-256 and the validator recomputes it from disk.
        using var home = new TestHome();
        var (metadata, state, statePath) = WriteBoundPair(home);

        state["documentRef"] = @"C:\books\attacker-chosen.xlsx";
        File.WriteAllText(statePath, state.ToJsonString());

        var result = ExcelAdapter.ValidateDeferredFormatRestoreBinding(
            metadata, ReadState(statePath), statePath);

        Assert.False(Json.GetBool(result, "ok"), result.ToJsonString());
        // documentRef disagreement is caught first; either way it never reaches "ok".
        Assert.Contains(Json.GetString(result, "code"), new[] { "documentRef", "state-hash" });
    }

    [Fact]
    public void A_state_json_whose_bytes_changed_at_all_is_refused()
    {
        using var home = new TestHome();
        var (metadata, _, statePath) = WriteBoundPair(home);
        File.WriteAllText(statePath, File.ReadAllText(statePath) + " ");

        AssertPreflightRefused(
            ExcelAdapter.ValidateDeferredFormatRestoreBinding(metadata, ReadState(statePath), statePath),
            "state-hash");
    }

    [Fact]
    public void A_missing_state_file_is_refused()
    {
        using var home = new TestHome();
        var (metadata, _, statePath) = WriteBoundPair(home);
        File.Delete(statePath);

        AssertPreflightRefused(
            ExcelAdapter.ValidateDeferredFormatRestoreBinding(metadata, EnvelopeState(), statePath),
            "state-hash");
    }

    [Fact]
    public void A_state_file_that_is_not_the_envelope_is_refused_as_envelope()
    {
        // The validator authenticates the file bytes and then parses THAT file; the caller's state
        // object is only cross-checked afterwards. So isolating the envelope rule means writing a
        // non-envelope state to the actual file and rebinding its hash, otherwise the hash check
        // fires first and the envelope rule is never reached.
        using var home = new TestHome();
        var (metadata, state, statePath) = WriteBoundPair(home);
        state["snapshotVersion"] = 2;
        RebindStateFile(metadata, state, statePath);

        AssertPreflightRefused(
            ExcelAdapter.ValidateDeferredFormatRestoreBinding(metadata, state, statePath),
            "envelope");
    }

    [Fact]
    public void A_caller_state_that_disagrees_with_the_authenticated_file_is_refused_as_state_hash()
    {
        // Complementary rule, deliberately kept separate: the file itself is a valid, correctly
        // hashed envelope, and only the caller's in-memory copy differs. The validator must trust
        // the authenticated bytes and refuse the mismatched caller state.
        using var home = new TestHome();
        var (metadata, _, statePath) = WriteBoundPair(home);
        var callerState = ReadState(statePath);
        callerState["snapshotVersion"] = 2;

        AssertPreflightRefused(
            ExcelAdapter.ValidateDeferredFormatRestoreBinding(metadata, callerState, statePath),
            "state-hash");
    }

    // --------------------------------------- restore verification stays unchanged

    [Fact]
    public void A_deferred_extraction_failure_can_never_be_reported_as_a_verified_rollback()
    {
        var failure = new JsonObject
        {
            ["ok"] = false,
            ["restored"] = false,
            ["restoreMode"] = ExcelAdapter.DeferredFormatRestoreMode,
            ["deferredExtraction"] = new JsonObject { ["attempted"] = true, ["ok"] = false },
        };

        Assert.False(DocBridgeHost.IsVerifiedRestore(failure));
        Assert.Null(Json.GetObj(failure, "readback"));
    }

    // -------------------------------------------------------------------- helpers

    private static ExcelAdapter.ExcelDeferredFormatSnapshotPolicy Enabled => new(true);

    private static ExcelAdapter.ExcelDeferredFormatSnapshotPolicy Disabled => new(false);

    private static JsonObject EnvelopeState() => new()
    {
        ["restoreMode"] = "copy-sheet-topology",
        ["snapshotVersion"] = 4,
        ["payloadMode"] = "deferred-format-only",
        ["documentRef"] = @"C:\books\live.xlsx",
    };

    private static JsonObject ExecuteMetadata() => new()
    {
        ["snapshotId"] = "20260910-101112-abcdef01",
        ["app"] = "excel",
        ["reason"] = "excel_apply_ops execute",
        [HostSnapshotCaptureContext.MetadataKey] = HostSnapshotCaptureContext.Execute,
    };

    private static JsonObject FillOp(string range) => new()
    {
        ["op"] = "format_range",
        ["target"] = new JsonObject { ["sheet"] = "Sheet1" },
        ["range"] = range,
        ["style"] = new JsonObject { ["fillColor"] = 16711680 },
    };

    private static string RangeOfCells(int cells)
    {
        const int rows = 100;
        var columns = cells / rows;
        Assert.Equal(cells, columns * rows);
        var name = "";
        var index = columns;
        while (index > 0)
        {
            index = Math.DivRem(index - 1, 26, out var rem);
            name = (char)('A' + rem) + name;
        }

        return $"A1:{name}{rows}";
    }

    private static string NewSnapshotDir(TestHome home, out string dirName)
    {
        dirName = "20260910-101112-abcdef01";
        var dir = Path.Combine(home.Dir, "snapshots", "excel", dirName);
        Directory.CreateDirectory(dir);
        return dir;
    }

    private static (JsonObject Metadata, JsonObject State, string StatePath) WriteBoundPair(TestHome home)
    {
        var snapshotDir = NewSnapshotDir(home, out var dirName);
        var checkpointName = $"deferred-format-{dirName}.xlsx";
        var checkpointPath = Path.Combine(snapshotDir, checkpointName);
        File.Copy(WritePackage(home, "source.xlsx"), checkpointPath, overwrite: true);
        var checkpointSha = Sha256Hex(checkpointPath);

        var state = EnvelopeState();
        state["checkpointFile"] = checkpointName;
        state["checkpointSha256"] = checkpointSha;
        var statePath = Path.Combine(snapshotDir, "state.json");
        File.WriteAllText(statePath, state.ToJsonString());

        var metadata = new JsonObject
        {
            ["snapshotId"] = dirName,
            ["app"] = "excel",
            ["restoreMode"] = ExcelAdapter.DeferredFormatRestoreMode,
            ["snapshotKind"] = ExcelAdapter.DeferredFormatRestoreMode,
            ["payloadMode"] = ExcelAdapter.DeferredFormatPayloadMode,
            ["documentRef"] = Json.GetString(state, "documentRef"),
            ["deferredCheckpoint"] = checkpointName,
            ["deferredCheckpointSha256"] = checkpointSha,
            ["deferredStateSha256"] = Sha256Hex(statePath),
        };

        return (metadata, state, statePath);
    }

    /// <summary>
    /// Writes <paramref name="state"/> to the real state.json and rebinds metadata to its new
    /// hash, so a later refusal is attributable to the state's content rather than to the hash
    /// check that would otherwise fire first.
    /// </summary>
    private static void RebindStateFile(JsonObject metadata, JsonObject state, string statePath)
    {
        File.WriteAllText(statePath, state.ToJsonString());
        metadata["deferredStateSha256"] = Sha256Hex(statePath);
    }

    private static JsonObject ReadState(string path) =>
        File.Exists(path)
            ? JsonNode.Parse(File.ReadAllText(path)) as JsonObject ?? new JsonObject()
            : EnvelopeState();

    private static string Sha256Hex(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }

    private const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string StrictNs = "http://purl.oclc.org/ooxml/spreadsheetml/main";
    private const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string ContentTypesNs = "http://schemas.openxmlformats.org/package/2006/content-types";
    private const string XlsxMain =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
    private const string WorksheetType =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";
    private const string PlainCell = "<c r=\"A1\" t=\"inlineStr\"><is><t>plain</t></is></c>";

    /// <summary>
    /// One canonical, valid, minimal xlsx: correct content types with the ordinary main Override,
    /// package and workbook rels that resolve inside the package, transitional namespaces, and a
    /// single inline-string cell. Every refusal case mutates exactly one feature of this, so the
    /// code it returns is attributable to that feature rather than to a malformed stub.
    /// </summary>
    private static List<(string Name, string Content)> CanonicalPackage() => new()
    {
        ("[Content_Types].xml",
            $"<Types xmlns=\"{ContentTypesNs}\">"
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + $"<Override PartName=\"/xl/workbook.xml\" ContentType=\"{XlsxMain}\"/>"
            + $"<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"{WorksheetType}\"/>"
            + "</Types>"),
        ("_rels/.rels",
            $"<Relationships xmlns=\"{PackageRelNs}\">"
            + $"<Relationship Id=\"rId1\" Type=\"{RelNs}/officeDocument\" Target=\"xl/workbook.xml\"/>"
            + "</Relationships>"),
        ("xl/_rels/workbook.xml.rels",
            $"<Relationships xmlns=\"{PackageRelNs}\">"
            + $"<Relationship Id=\"rId1\" Type=\"{RelNs}/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
            + "</Relationships>"),
        ("xl/workbook.xml",
            $"<workbook xmlns=\"{MainNs}\" xmlns:r=\"{RelNs}\">"
            + "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>"
            + "</workbook>"),
        ("xl/worksheets/sheet1.xml",
            $"<worksheet xmlns=\"{MainNs}\">"
            + $"<sheetData><row r=\"1\">{PlainCell}</row></sheetData>"
            + "</worksheet>"),
    };

    private static string WritePackage(TestHome home, string name, string? mutation = null)
    {
        var parts = CanonicalPackage();
        if (mutation is not null) Mutate(parts, mutation);
        var path = Path.Combine(home.Dir, name);
        WriteZip(path, parts);
        return path;
    }

    private static void Mutate(List<(string Name, string Content)> parts, string mutation)
    {
        void Replace(string name, Func<string, string> edit)
        {
            var index = parts.FindIndex(part => part.Name == name);
            Assert.True(index >= 0, $"canonical package has no part '{name}'");
            parts[index] = (name, edit(parts[index].Content));
        }

        string Append(string content, string closing, string inserted) =>
            content.Replace(closing, inserted + closing, StringComparison.Ordinal);

        switch (mutation)
        {
            case "strict-namespace":
                Replace("xl/workbook.xml", c => c.Replace(MainNs, StrictNs, StringComparison.Ordinal));
                Replace("xl/worksheets/sheet1.xml", c => c.Replace(MainNs, StrictNs, StringComparison.Ordinal));
                break;
            case "cell-formula":
                Replace("xl/worksheets/sheet1.xml",
                    c => c.Replace(PlainCell, "<c r=\"A1\"><f>1+1</f><v>2</v></c>", StringComparison.Ordinal));
                break;
            case "merge":
                Replace("xl/worksheets/sheet1.xml", c => Append(c, "</worksheet>",
                    "<mergeCells count=\"1\"><mergeCell ref=\"A1:B2\"/></mergeCells>"));
                break;
            case "defined-name":
                Replace("xl/workbook.xml", c => Append(c, "</workbook>",
                    "<definedNames><definedName name=\"MyName\">Sheet1!$A$1</definedName></definedNames>"));
                break;
            case "conditional-format":
                Replace("xl/worksheets/sheet1.xml", c => Append(c, "</worksheet>",
                    "<conditionalFormatting sqref=\"A1\">"
                    + "<cfRule type=\"expression\" priority=\"1\"><formula>TRUE</formula></cfRule>"
                    + "</conditionalFormatting>"));
                break;
            case "data-validation":
                Replace("xl/worksheets/sheet1.xml", c => Append(c, "</worksheet>",
                    "<dataValidations count=\"1\">"
                    + "<dataValidation type=\"list\" sqref=\"A1\"><formula1>1</formula1></dataValidation>"
                    + "</dataValidations>"));
                break;
            case "full-calc":
                Replace("xl/workbook.xml", c => Append(c, "</workbook>",
                    "<calcPr calcId=\"191029\" fullCalcOnLoad=\"1\"/>"));
                break;
            case "macro-main":
                Replace("[Content_Types].xml", c => c.Replace(
                    XlsxMain, "application/vnd.ms-excel.sheet.macroEnabled.main+xml", StringComparison.Ordinal));
                break;
            case "vba":
                parts.Add(("xl/vbaProject.bin", "placeholder; the part name is what is refused"));
                break;
            case "external-rel":
                Replace("xl/_rels/workbook.xml.rels", c => Append(c, "</Relationships>",
                    $"<Relationship Id=\"rId9\" Type=\"{RelNs}/externalLink\" "
                    + "Target=\"file:///C:/other.xlsx\" TargetMode=\"External\"/>"));
                break;
            case "malformed":
                // Keep the canonical namespace declarations and damage only the element closure,
                // so the refusal isolates XML syntax. Replacing the whole part with a bare
                // <workbook> also stripped the namespace, which the parser rejects as
                // strict-ooxml before it ever reaches a well-formedness error.
                Replace("xl/workbook.xml", c => c.Replace(
                    "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>",
                    "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/>",
                    StringComparison.Ordinal));
                break;
            case "dtd":
                Replace("xl/workbook.xml", c => "<!DOCTYPE workbook [<!ENTITY sample \"y\">]>" + c);
                break;
            case "drop-content-types":
                parts.RemoveAll(part => part.Name == "[Content_Types].xml");
                break;
            case "drop-workbook":
                parts.RemoveAll(part => part.Name == "xl/workbook.xml");
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(mutation), mutation, "unknown mutation");
        }
    }

    private static void WriteZip(string path, IEnumerable<(string Name, string Content)> entries) =>
        MinimalXlsxSeed.WriteZip(path, entries);

    private static FakeWorkbook NewWorkbook(TestHome home)
    {
        var path = Path.Combine(home.Dir, "book.xlsx");
        File.WriteAllText(path, "workbook bytes");
        return new FakeWorkbook { FullNameValue = path };
    }

    private static void AssertRefused(JsonObject result, string code)
    {
        Assert.False(Json.GetBool(result, "eligible"), result.ToJsonString());
        Assert.False(Json.GetBool(result, "used"));
        Assert.Equal(code, Json.GetString(result, "code"));
        Assert.False(string.IsNullOrWhiteSpace(Json.GetString(result, "reason")));
    }

    private static void AssertPreflightRefused(JsonObject result, string code)
    {
        Assert.False(Json.GetBool(result, "ok"), result.ToJsonString());
        Assert.Equal(code, Json.GetString(result, "code"));
    }

    /// <summary>Plain object reached through the same late-bound call sites production uses.</summary>
    public sealed class FakeWorkbook
    {
        public string FullNameValue { get; set; } = "";

        public int? FileFormatValue { get; set; } = 51;

        public bool HasVBProjectValue { get; set; }

        public bool HasVBProjectThrows { get; set; }

        public bool ProtectStructureValue { get; set; }

        public bool ProtectWindowsValue { get; set; }

        public bool HasPasswordValue { get; set; }

        public bool WriteReservedValue { get; set; }

        public bool HasLinks { get; set; }

        public bool SavedThrows { get; set; }

        public FakeApplication Owner { get; } = new();

        public string FullName => FullNameValue;

        public int FileFormat =>
            FileFormatValue ?? throw new InvalidOperationException("FileFormat is unavailable");

        public bool Saved =>
            SavedThrows ? throw new InvalidOperationException("Saved is unavailable") : false;

        public FakeApplication Application => Owner;

        public FakeSheets Worksheets { get; } = new();

        // Named exactly as the adapter reads them through `dynamic`.
        public bool HasVBProject =>
            HasVBProjectThrows ? throw new InvalidOperationException("unavailable") : HasVBProjectValue;

        public bool ProtectStructure => ProtectStructureValue;

        public bool ProtectWindows => ProtectWindowsValue;

        public bool HasPassword => HasPasswordValue;

        public bool WriteReserved => WriteReservedValue;

        public object? LinkSources(int type) => HasLinks ? new object[] { "C:\\other.xlsx" } : null;
    }

    public sealed class FakeApplication
    {
        public FakeWorkbooks Workbooks { get; } = new();

        public long Hwnd => 4242;
    }

    public sealed class FakeWorkbooks
    {
        public int Count => 1;
    }

    public sealed class FakeSheets
    {
        public FakeSheet Item(string name) => new();
    }

    public sealed class FakeSheet
    {
        public bool ProtectContents => false;
    }
}

/// <summary>
/// A genuinely minimal, plain <c>.xlsx</c> written offline, used by the native harness as the seed
/// it opens in its owned Excel instead of <c>Workbooks.Add</c>.
///
/// Why it exists: on a machine with Office Store add-ins installed, Excel's own new-workbook
/// template injects <c>xl/webextensions/taskpanes.xml</c> and <c>webextension1..3.xml</c> into
/// every package it saves — an explicit <c>Add(xlWBATWorksheet)</c> included. The deferred byte
/// gate refuses those parts as <c>complex-part</c>, which is the correct product behaviour and is
/// not to be relaxed, whitelisted, or worked around by stripping parts out of a checkpoint. The
/// fix belongs to the fixture: start from bytes the harness wrote itself, in a file it owns, with
/// no user setting, add-in, template, or registry key touched.
///
/// The package is the five parts an .xlsx cannot do without — content types, the package rels, the
/// workbook, the workbook rels, and one empty worksheet — each a well-formed transitional-namespace
/// document with an XML declaration. Styles, shared strings, theme, docProps and calc chain are all
/// optional and deliberately absent; Excel regenerates whatever it needs when the harness populates
/// the sheets and saves.
/// </summary>
internal static class MinimalXlsxSeed
{
    private const string Declaration = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>";
    private const string MainNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private const string PackageRelNs = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string ContentTypesNs = "http://schemas.openxmlformats.org/package/2006/content-types";
    private const string XlsxMain =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";
    private const string WorksheetType =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml";

    internal static IReadOnlyList<(string Name, string Content)> Parts() => new List<(string, string)>
    {
        ("[Content_Types].xml",
            Declaration
            + $"<Types xmlns=\"{ContentTypesNs}\">"
            + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
            + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
            + $"<Override PartName=\"/xl/workbook.xml\" ContentType=\"{XlsxMain}\"/>"
            + $"<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"{WorksheetType}\"/>"
            + "</Types>"),
        ("_rels/.rels",
            Declaration
            + $"<Relationships xmlns=\"{PackageRelNs}\">"
            + $"<Relationship Id=\"rId1\" Type=\"{RelNs}/officeDocument\" Target=\"xl/workbook.xml\"/>"
            + "</Relationships>"),
        ("xl/workbook.xml",
            Declaration
            + $"<workbook xmlns=\"{MainNs}\" xmlns:r=\"{RelNs}\">"
            + "<sheets><sheet name=\"Sheet1\" sheetId=\"1\" r:id=\"rId1\"/></sheets>"
            + "</workbook>"),
        ("xl/_rels/workbook.xml.rels",
            Declaration
            + $"<Relationships xmlns=\"{PackageRelNs}\">"
            + $"<Relationship Id=\"rId1\" Type=\"{RelNs}/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
            + "</Relationships>"),
        ("xl/worksheets/sheet1.xml",
            Declaration
            + $"<worksheet xmlns=\"{MainNs}\">"
            + "<dimension ref=\"A1\"/><sheetData/>"
            + "</worksheet>"),
    };

    internal static void Write(string path) => WriteZip(path, Parts());

    internal static void WriteZip(string path, IEnumerable<(string Name, string Content)> entries)
    {
        using var zip = ZipFile.Open(path, ZipArchiveMode.Create);
        foreach (var (name, content) in entries)
        {
            using var writer = new StreamWriter(zip.CreateEntry(name).Open());
            writer.Write(content);
        }
    }
}
