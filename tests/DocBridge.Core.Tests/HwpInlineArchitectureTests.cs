using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public sealed class HwpInlineArchitectureTests
{
    [Fact]
    public void Hwp_serialized_numeric_entities_are_decoded_before_text_length_and_matching()
    {
        const string serialized = "관경 &#8722; 300&#13212;";
        var decoded = HwpAdapter.DecodeHwpSerializedText(serialized);
        Assert.Equal("관경 − 300㎜", decoded);
        Assert.Equal(9, decoded.Length);
        Assert.DoesNotContain("&#", decoded);
    }

    [Fact]
    public void Preview_artifact_round_trips_diff_and_risk_metadata()
    {
        var preview = new ApplyPreview { DiffTruncated = true, RequiresHighRiskApproval = true };
        preview.Affected.Add(new AffectedRef("table", "table:0"));
        preview.Diff.Add(new DiffEntry { Ref = "table:0/cell:0", Before = "a", After = "b" });
        preview.Warnings.Add("warning");
        preview.Interaction = new JsonObject { ["policy"] = "preserve-foreground" };

        var metadata = new JsonObject();
        ApplyPreviewArtifact.StoreInMetadata(metadata, "ops-hash", preview);
        var restored = ApplyPreviewArtifact.FromMetadata(metadata, "ops-hash");

        Assert.NotNull(restored);
        Assert.True(restored!.DiffTruncated);
        Assert.True(restored.RequiresHighRiskApproval);
        Assert.Single(restored.Affected);
        Assert.Single(restored.Diff);
        Assert.Single(restored.Warnings);
        Assert.Equal("preserve-foreground", Json.GetString(restored.Interaction, "policy"));
        Assert.Null(ApplyPreviewArtifact.FromMetadata(metadata, "different-ops"));
    }

    [Fact]
    public void TypeLib_registration_evaluator_distinguishes_missing_and_version_mismatch()
    {
        var root = Path.Combine(Path.GetTempPath(), "docbridge-hwp-doctor-" + Guid.NewGuid().ToString("n"));
        var installedDir = Path.Combine(root, "current");
        var oldDir = Path.Combine(root, "old");
        Directory.CreateDirectory(installedDir);
        Directory.CreateDirectory(oldDir);
        var executable = Path.Combine(installedDir, "Hwp.exe");
        var currentTypeLib = Path.Combine(installedDir, "HwpObject.tlb");
        var oldTypeLib = Path.Combine(oldDir, "HwpObject.tlb");
        File.WriteAllText(executable, "test");
        File.WriteAllText(currentTypeLib, "test");
        File.WriteAllText(oldTypeLib, "test");
        try
        {
            Assert.Equal("HWP_NOT_INSTALLED", HwpEnvironmentDoctor.EvaluateRegistration(null, currentTypeLib));
            Assert.Equal("HWP_TYPELIB_NOT_REGISTERED", HwpEnvironmentDoctor.EvaluateRegistration(executable, null));
            Assert.Equal("HWP_TYPELIB_VERSION_MISMATCH", HwpEnvironmentDoctor.EvaluateRegistration(executable, oldTypeLib));
            Assert.Equal("CHECK_PASSED", HwpEnvironmentDoctor.EvaluateRegistration(executable, currentTypeLib));
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void Hwp_COM_startup_working_directory_is_the_installed_Bin_folder()
    {
        var root = Path.Combine(Path.GetTempPath(), "docbridge-hwp-cwd-" + Guid.NewGuid().ToString("n"));
        var bin = Path.Combine(root, "HOffice130", "Bin");
        Directory.CreateDirectory(bin);
        var executable = Path.Combine(bin, "Hwp.exe");
        File.WriteAllText(executable, "test");
        try
        {
            Assert.Equal(Path.GetFullPath(bin), HwpEnvironmentDoctor.GetAutomationWorkingDirectory(executable));
            Assert.Null(HwpEnvironmentDoctor.GetAutomationWorkingDirectory(Path.Combine(root, "missing", "Hwp.exe")));
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    [Fact]
    public void Automation_environment_detects_missing_or_malformed_windows_variables()
    {
        var windowsDirectory = HwpEnvironmentDoctor.GetAutomationWindowsDirectory();
        Assert.False(string.IsNullOrWhiteSpace(windowsDirectory));
        Assert.True(HwpEnvironmentDoctor.NeedsProcessEnvironmentRepair(null, null, windowsDirectory));
        Assert.True(HwpEnvironmentDoctor.NeedsProcessEnvironmentRepair(
            "not-a-windows-directory", windowsDirectory, windowsDirectory));
        Assert.False(HwpEnvironmentDoctor.NeedsProcessEnvironmentRepair(
            windowsDirectory, windowsDirectory, windowsDirectory));
    }

    [Fact]
    public void Worker_start_environment_repairs_windir_and_systemroot_for_WPF_FontCache()
    {
        var windowsDirectory = HwpEnvironmentDoctor.GetAutomationWindowsDirectory();
        Assert.False(string.IsNullOrWhiteSpace(windowsDirectory));
        var startInfo = new System.Diagnostics.ProcessStartInfo { UseShellExecute = false };
        startInfo.Environment.Remove("windir");
        startInfo.Environment.Remove("SystemRoot");

        HwpEnvironmentDoctor.ApplyAutomationEnvironment(startInfo);

        Assert.Equal(windowsDirectory, startInfo.Environment["windir"]);
        Assert.Equal(windowsDirectory, startInfo.Environment["SystemRoot"]);
        Assert.Throws<UriFormatException>(() => new Uri(@"\Fonts\", UriKind.Absolute));
        Assert.True(new Uri(Path.Combine(windowsDirectory!, "Fonts") + Path.DirectorySeparatorChar,
            UriKind.Absolute).IsAbsoluteUri);
    }

    [Theory]
    [InlineData("13.0.0.866", true)]
    [InlineData("13.0.0.3869", true)]
    [InlineData("13.0.0.3870", false)]
    [InlineData("13.0.0.4000", false)]
    [InlineData("12.0.0.1000", false)]
    public void Hwp_2024_update_recommendation_is_version_scoped(string version, bool expected)
    {
        Assert.Equal(expected,
            HwpEnvironmentDoctor.IsHwp2024VersionOlderThan(version, HwpEnvironmentDoctor.RecommendedHwp2024Version));
    }

    [Theory]
    [InlineData("Hnc.Controls.Native.PopupBorderImpl", "tour-popup-type-initializer")]
    [InlineData("Hnc.Controls.Widgets.TourPopup.ctor", "tour-popup-type-initializer")]
    [InlineData("MS.Internal.FontCache.Util", "font-cache-type-initializer")]
    [InlineData("CultureFontManager failed", "font-cache-type-initializer")]
    [InlineData("TypeInitializationException: Hnc.Controls.Widget", "hnc-controls-type-initializer")]
    [InlineData("ordinary HWP message", null)]
    public void Hwp_UI_failure_signatures_are_classified_without_retry(string text, string? expected)
    {
        Assert.Equal(expected, HwpUiFailureDetector.ClassifyText(text));
        Assert.Equal(expected is not null, HwpUiFailureDetector.IsDeterministicFailureMessage(text));
    }

    [Fact]
    public void Worker_protocol_round_trips_step_results_and_post_edit_reread()
    {
        var original = new ApplyExecution
        {
            Ok = true,
            Readback = new JsonObject
            {
                ["verified"] = true,
                ["postEditReread"] = new JsonObject { ["textSha256"] = "abc", ["textLength"] = 3 },
            },
            Interaction = new JsonObject
            {
                ["policy"] = "preserve-foreground",
                ["foregroundPreserved"] = true,
                ["originalStateRestored"] = true,
            },
        };
        original.Affected.Add(new AffectedRef("paragraph", "p-abc-1"));
        original.Diff.Add(new DiffEntry { Ref = "p-abc-1", Before = "a", After = "b" });
        original.OperationResults.Add(new JsonObject { ["index"] = 0, ["op"] = "insert_text", ["ok"] = true });

        var roundTrip = HwpWorkerProtocol.ExecutionFromJson(HwpWorkerProtocol.ExecutionToJson(original));

        Assert.True(roundTrip.Ok);
        Assert.Single(roundTrip.Affected);
        Assert.Single(roundTrip.Diff);
        Assert.Single(roundTrip.OperationResults);
        Assert.Equal("abc", Json.GetString(Json.GetObj(roundTrip.Readback, "postEditReread"), "textSha256"));
        Assert.Equal("preserve-foreground", Json.GetString(roundTrip.Interaction, "policy"));
        Assert.True(Json.GetBool(roundTrip.Interaction, "originalStateRestored"));
    }

    [Fact]
    public void Worker_protocol_detects_poisoned_COM_timeout_result()
    {
        Assert.True(HwpWorkerProtocol.ContainsComTimeout(new JsonObject
        {
            ["errors"] = new JsonArray("STA work item did not complete within 120s (possible COM modal dialog)"),
        }));
        Assert.False(HwpWorkerProtocol.ContainsComTimeout(new JsonObject { ["ok"] = true }));
    }

    [Fact]
    public void Worker_request_timeouts_are_scoped_by_method_and_workload()
    {
        Assert.Equal(15, HwpWorkerAdapter.ResolveRequestTimeout("getStatus", new JsonObject()).TotalSeconds);
        Assert.Equal(20, HwpWorkerAdapter.ResolveRequestTimeout("getActiveContext", new JsonObject()).TotalSeconds);
        Assert.Equal(30, HwpWorkerAdapter.ResolveRequestTimeout("read", new JsonObject { ["scope"] = "document" }).TotalSeconds);
        Assert.Equal(45, HwpWorkerAdapter.ResolveRequestTimeout("read", new JsonObject { ["scope"] = "bundle" }).TotalSeconds);
        Assert.Equal(45, HwpWorkerAdapter.ResolveRequestTimeout("read", new JsonObject { ["scope"] = "document_map" }).TotalSeconds);
        Assert.Equal(45, HwpWorkerAdapter.ResolveRequestTimeout("read", new JsonObject { ["scope"] = "structure" }).TotalSeconds);
        Assert.Equal(45, HwpWorkerAdapter.ResolveRequestTimeout("read", new JsonObject { ["scope"] = "fields" }).TotalSeconds);
        Assert.Equal(45, HwpWorkerAdapter.ResolveRequestTimeout("read", new JsonObject { ["scope"] = "tables" }).TotalSeconds);
        Assert.Equal(45, HwpWorkerAdapter.ResolveRequestTimeout("read", new JsonObject { ["mode"] = "tables" }).TotalSeconds);

        Assert.Equal(90, HwpWorkerAdapter.ResolveRequestTimeout("launch", new JsonObject
        {
            ["creationMode"] = "docx-first",
            ["sourceFile"] = @"C:\work\source.docx",
            ["outputFile"] = @"C:\work\result.hwpx",
        }).TotalSeconds);
        Assert.Equal(90, HwpWorkerAdapter.ResolveRequestTimeout("launch", new JsonObject
        {
            ["sourceFile"] = @"C:\work\source.docx",
        }).TotalSeconds);
        Assert.Equal(45, HwpWorkerAdapter.ResolveRequestTimeout("launch", new JsonObject
        {
            ["creationMode"] = "native-hwp",
            ["newDocument"] = true,
        }).TotalSeconds);

        var ordinary = new JsonObject
        {
            ["ops"] = new JsonArray(new JsonObject { ["op"] = "insert_text", ["text"] = "a" }),
        };
        var expensive = new JsonObject
        {
            ["ops"] = new JsonArray(new JsonObject { ["op"] = "export_pdf", ["output"] = "x.pdf" }),
        };
        Assert.Equal(45, HwpWorkerAdapter.ResolveRequestTimeout("apply", ordinary).TotalSeconds);
        Assert.Equal(90, HwpWorkerAdapter.ResolveRequestTimeout("apply", expensive).TotalSeconds);
        Assert.Equal(90, HwpWorkerAdapter.ResolveRequestTimeout("restoreSnapshot", new JsonObject()).TotalSeconds);
    }

    [Fact]
    public void Worker_retries_only_non_file_read_transport_failures()
    {
        Assert.True(HwpWorkerAdapter.CanRetryReadTransport("getStatus", new JsonObject(), readOnly: true));
        Assert.True(HwpWorkerAdapter.CanRetryReadTransport("read", new JsonObject(), readOnly: true));
        Assert.False(HwpWorkerAdapter.CanRetryReadTransport("read",
            new JsonObject { ["file"] = @"C:\work\a.hwp" }, readOnly: true));
        Assert.False(HwpWorkerAdapter.CanRetryReadTransport("preview", new JsonObject
        {
            ["ops"] = new JsonArray(new JsonObject { ["op"] = "insert_text", ["file"] = @"C:\work\a.hwp" }),
        }, readOnly: true));
        Assert.False(HwpWorkerAdapter.CanRetryReadTransport("apply", new JsonObject(), readOnly: false));
        Assert.False(HwpWorkerAdapter.IsRetryableTransportFailure(new TimeoutException()));
        Assert.True(HwpWorkerAdapter.IsRetryableTransportFailure(new IOException("worker exited")));
    }

    [Fact]
    public void Worker_circuit_blocks_until_cooldown_expires()
    {
        var clock = new ManualTimeProvider(new DateTimeOffset(2026, 8, 18, 0, 0, 0, TimeSpan.Zero));
        var circuit = new HwpWorkerCircuitBreaker(clock);
        circuit.Open("HWP_COM_TIMEOUT", TimeSpan.FromSeconds(45));

        Assert.True(circuit.TryGetOpen(out var opened));
        Assert.Equal("HWP_COM_TIMEOUT", opened.Code);
        Assert.InRange(opened.Remaining.TotalSeconds, 44.9, 45.0);

        clock.Advance(TimeSpan.FromSeconds(46));
        Assert.False(circuit.TryGetOpen(out _));
    }

    [Fact]
    public void Worker_circuit_is_not_reset_by_failed_repair_or_doctor_response()
    {
        Assert.False(HwpWorkerAdapter.ShouldResetCircuitAfterResponse(
            "repairTypeLib", new JsonObject { ["ok"] = false }));
        Assert.True(HwpWorkerAdapter.ShouldResetCircuitAfterResponse(
            "repairTypeLib", new JsonObject { ["ok"] = true }));
        Assert.False(HwpWorkerAdapter.ShouldResetCircuitAfterResponse(
            "doctor", new JsonObject { ["ok"] = true }));
        Assert.True(HwpWorkerAdapter.ShouldResetCircuitAfterResponse(
            "getStatus", new JsonObject { ["ok"] = true }));
    }

    [Fact]
    public void Worker_structured_failure_preserves_snapshot_error_contract()
    {
        var failure = HwpWorkerAdapter.StructuredFailure(new JsonObject
        {
            ["ok"] = false,
            ["errorCode"] = "HWP_DOCUMENT_NOT_FOUND",
            ["errors"] = new JsonArray("dry-run 대상 문서가 닫혔습니다"),
            ["userAction"] = "문서를 다시 열고 dry-run을 다시 실행하세요.",
            ["retryAfterMs"] = 2500,
        });

        Assert.Equal("HWP_DOCUMENT_NOT_FOUND", failure.Code);
        Assert.Equal("dry-run 대상 문서가 닫혔습니다", failure.Message);
        Assert.Equal("문서를 다시 열고 dry-run을 다시 실행하세요.", failure.UserAction);
        Assert.Equal(2500, failure.RetryAfterMs);
    }

    [Fact]
    public void Failed_snapshot_response_does_not_clear_existing_metadata()
    {
        var metadata = new JsonObject { ["documentRef"] = "hwp://still-open-before-call" };
        var result = new JsonObject
        {
            ["ok"] = false,
            ["errorCode"] = "HWP_DOCUMENT_NOT_FOUND",
            ["errors"] = new JsonArray("snapshot target closed"),
        };

        var failure = Assert.Throws<HwpAutomationException>(() =>
            HwpWorkerAdapter.ReplaceSnapshotMetadata(result, metadata));

        Assert.Equal("HWP_DOCUMENT_NOT_FOUND", failure.Code);
        Assert.Equal("hwp://still-open-before-call", Json.GetString(metadata, "documentRef"));
    }

    [Fact]
    public void Bulk_row_height_specs_validate_range_and_duplicates()
    {
        var valid = HwpAdapter.ParseRowHeightSpecs(Json.ParseObject("""
        { "rows": [{ "row": 0, "heightMm": 8.5 }, { "row": 3, "heightMm": 12 }] }
        """)!);
        Assert.Equal(2, valid.Count);
        Assert.Equal(3, valid[1].Row);

        Assert.Throws<ArgumentException>(() => HwpAdapter.ParseRowHeightSpecs(Json.ParseObject("""
        { "rows": [{ "row": 1, "heightMm": 8 }, { "row": 1, "heightMm": 9 }] }
        """)!));
        Assert.Throws<ArgumentOutOfRangeException>(() => HwpAdapter.ParseRowHeightSpecs(Json.ParseObject("""
        { "rows": [{ "row": 0, "heightMm": 3.9 }] }
        """)!));

        var tooMany = new JsonObject
        {
            ["rows"] = new JsonArray(Enumerable.Range(0, 101)
                .Select(row => (JsonNode)new JsonObject { ["row"] = row, ["heightMm"] = 8 })
                .ToArray()),
        };
        Assert.Throws<ArgumentOutOfRangeException>(() => HwpAdapter.ParseRowHeightSpecs(tooMany));
    }

    [Fact]
    public void Bulk_paragraph_format_requires_style_and_valid_scope()
    {
        HwpAdapter.ValidateFormatParagraphItems(Json.ParseObject("""
        {
          "items": [
            {
              "target": { "text": "제목" },
              "characterStyle": { "fontSize": 14, "bold": true },
              "paragraphStyle": { "align": "center", "lineSpacingPercent": 160 }
            }
          ]
        }
        """)!);

        Assert.Throws<ArgumentException>(() => HwpAdapter.ValidateFormatParagraphItems(Json.ParseObject("""
        { "items": [{ "target": { "scope": "document" } }] }
        """)!));
        Assert.Throws<ArgumentException>(() => HwpAdapter.ValidateFormatParagraphItems(Json.ParseObject("""
        { "items": [{ "target": { "scope": "unknown" }, "paragraphStyle": { "align": "left" } }] }
        """)!));
    }

    [Fact]
    public void Structured_HWP_error_includes_retry_delay()
    {
        var result = new HwpAutomationException(
            "HWP_CIRCUIT_OPEN", "보호 대기 중", "45초 뒤 다시 시도", retryAfterMs: 45000).ToResult();
        Assert.Equal("HWP_CIRCUIT_OPEN", Json.GetString(result, "errorCode"));
        Assert.Equal(45000, Json.GetInt(result, "retryAfterMs"));
        Assert.True(Json.GetBool(result, "retryable"));
        Assert.False(Json.GetBool(result, "automaticRetry"));
        Assert.Equal("after-delay", Json.GetString(Json.GetObj(result, "retryPolicy"), "mode"));
    }

    private sealed class ManualTimeProvider : TimeProvider
    {
        private DateTimeOffset _utcNow;
        internal ManualTimeProvider(DateTimeOffset utcNow) => _utcNow = utcNow;
        public override DateTimeOffset GetUtcNow() => _utcNow;
        internal void Advance(TimeSpan value) => _utcNow += value;
    }
}
