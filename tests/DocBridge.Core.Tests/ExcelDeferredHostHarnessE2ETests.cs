using System.Diagnostics;
using System.Globalization;
using System.IO.Compression;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Models;
using DocBridge.Core.Services;
using Xunit.Abstractions;

namespace DocBridge.Core.Tests;

/// <summary>
/// Native end-to-end harness for the opt-in Excel deferred format checkpoint. **Root runs this.**
///
/// Gated on <c>DOCBRIDGE_EXCEL_DEFERRED_NATIVE=1</c>. With the gate off every case returns at once,
/// so an ordinary suite run never launches Excel. With the gate **on**, a missing prerequisite is a
/// failure, never a silent skip — a green run that quietly did nothing is what this harness exists
/// to prevent.
///
/// Ownership follows the standalone probe rules. The Excel RCW is created **inside the adapter's
/// STA factory**, not on the xunit MTA thread, and ownership is proven by mapping the attached
/// application's own <c>Hwnd</c> to a PID — not by assuming any newly appeared Excel process is
/// ours. Only that instance and its generated fixture are ever closed.
///
/// The fixture is built from an offline minimal .xlsx seed the harness writes itself, not from
/// Workbooks.Add: on this machine some starts of Excel embed the user's Office Store add-ins as
/// xl/webextensions/* into every package it saves, which the product correctly refuses as
/// complex-part, and other starts embed nothing. Whichever it is on the day, the seed keeps the
/// fixture out of it, and is left in the output directory as evidence of what the fixture came
/// from. Case 7 covers the refusal side deterministically, with a formula this harness plants.
///
/// Evidence outlives the run: each case copies its whole DocBridge home (snapshots, metadata,
/// checkpoints) under <c>DOCBRIDGE_EXCEL_DEFERRED_OUTPUT</c> **before** the temp home is disposed,
/// under a name unique per case and per binary.
///
/// Environment (root supplies):
///   DOCBRIDGE_EXCEL_DEFERRED_NATIVE=1        enable
///   DOCBRIDGE_EXCEL_DEFERRED_OUTPUT          directory for fixture + evidence (required)
///   DOCBRIDGE_EXCEL_DEFERRED_SCALE           1000 (default) or 5000
///   DOCBRIDGE_EXCEL_DEFERRED_CLI             built doc-bridge-cli.exe (required for CLI cases)
///   DOCBRIDGE_EXCEL_DEFERRED_OLD_CLI         installed 0.4.18 / frozen 0.4.20 cli, ';'-separated
/// </summary>
[Trait("Category", "E2E")]
public sealed class ExcelDeferredHostHarnessE2ETests : IDisposable
{
    private const string TargetSheet = "Target";
    private const string KeepSheet = "Keep";
    private const string KeepMarker = "DO-NOT-TOUCH";
    private const int BlueOle = 16711680;

    /// <summary>
    /// The number format the post-checkpoint edits carry. A plain numeric custom format, which
    /// this Excel accepts through the late-bound Range.NumberFormat setter; the quoted-text
    /// formats used before did not survive that setter here, and neither did "General".
    /// </summary>
    private const string PostCheckpointNumberFormat = "0.0000";
    private const int CliTimeoutMs = 180_000;

    /// <summary>
    /// The formula case 7 plants on the Keep sheet. Literal arithmetic only: it references no other
    /// cell, sheet, workbook or external source, so it cannot reach anything but itself.
    /// </summary>
    private const string SafeFormula = "=1+2";

    /// <summary>How long teardown waits for our own Excel to exit before reporting it still runs.</summary>
    private const int OwnedExitWaitMs = 15_000;

    private readonly ITestOutputHelper _output;
    private readonly List<int> _preexistingExcelPids = ExcelPids();
    private readonly JsonArray _cases = new();

    private ExcelAdapter? _adapter;
    private object? _ownedApp;
    private int _ownedPid;
    private string _fixturePath = "";
    private string _seedPath = "";

    /// <summary>
    /// The owned workbook the cell/identity helpers currently look at. It is the fixture for every
    /// case except the environment-default one, which borrows the helpers for another workbook this
    /// harness created in the same owned instance, then puts this back.
    /// </summary>
    private string _activeBookPath = "";
    private long _ownedHwnd;
    private bool _ownershipProven;
    private bool _fixtureVerified;
    private bool _fixtureClosed;

    public ExcelDeferredHostHarnessE2ETests(ITestOutputHelper output) => _output = output;

    // ------------------------------------------------------------------- gating

    private static bool Enabled =>
        string.Equals(
            Environment.GetEnvironmentVariable("DOCBRIDGE_EXCEL_DEFERRED_NATIVE"), "1",
            StringComparison.Ordinal);

    private static string OutputDir =>
        Environment.GetEnvironmentVariable("DOCBRIDGE_EXCEL_DEFERRED_OUTPUT") ?? "";

    private static string CliPath =>
        Environment.GetEnvironmentVariable("DOCBRIDGE_EXCEL_DEFERRED_CLI") ?? "";

    private static IReadOnlyList<string> OldCliPaths =>
        (Environment.GetEnvironmentVariable("DOCBRIDGE_EXCEL_DEFERRED_OLD_CLI") ?? "")
        .Split(';', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

    private static int Scale =>
        int.TryParse(
            Environment.GetEnvironmentVariable("DOCBRIDGE_EXCEL_DEFERRED_SCALE"),
            NumberStyles.None, CultureInfo.InvariantCulture, out var value)
            ? value
            : 1000;

    private static void RequirePrerequisites()
    {
        Assert.False(string.IsNullOrWhiteSpace(OutputDir),
            "DOCBRIDGE_EXCEL_DEFERRED_OUTPUT is required when the native gate is on");
        Assert.True(Directory.Exists(OutputDir), $"output directory does not exist: {OutputDir}");
        Assert.True(Scale is 1000 or 5000, $"scale must be 1000 or 5000, got {Scale}");
        Assert.True(
            Scale >= ExcelAdapter.DeferredFormatMinCells && Scale <= ExcelAdapter.DeferredFormatMaxCells,
            $"scale {Scale} is outside the deferred rollout bounds");
    }

    // ------------------------------------------------------- 1. enabled, happy path

    [Fact]
    public void Enabled_execute_uses_the_deferred_checkpoint_and_explicit_restore_puts_everything_back()
    {
        if (!Enabled) return;
        RequirePrerequisites();

        using var scope = NewCase("enabled-happy-path");
        var home = scope.Home;
        var evidence = scope.Evidence;
        var host = OpenFixtureAndHost(home, deferred: true);

        // Dirty the workbook through a harmless owned marker BEFORE the baseline identity, so the
        // baseline itself is a dirty workbook. Comparing a Saved=true baseline against a workbook
        // that Apply legitimately dirtied would be testing the wrong thing.
        DirtyOwnedMarker("pre-checkpoint");
        var before = ScanTargetStyles();
        var identityBefore = ReadIdentity();
        Assert.False(
            Json.GetBool(identityBefore, "saved"),
            "baseline must already be dirty so the copy cannot appear to preserve dirtiness by luck");

        var execute = host.ApplyOps("excel", FormatBatch());
        Artifact(evidence, "execute", execute);
        Assert.True(Json.GetBool(execute, "ok"), execute.ToJsonString());

        var snapshotId = Json.GetString(execute, "snapshotId")!;
        AssertDeferredWasUsed(home, snapshotId);

        var applied = ScanTargetStyles();
        for (var i = 0; i < Scale; i++)
        {
            Assert.Equal(BlueOle, applied[i].Fill);
            Assert.NotEqual(before[i].Fill, applied[i].Fill);
        }

        // Newer, unrelated edits made after the checkpoint must survive the rollback.
        SetCell(TargetSheet, "A1", "post-checkpoint-value", PostCheckpointNumberFormat);
        SetFormula(TargetSheet, "B1", "=1+2");

        // Read them back natively before the restore as well. An assertion that only runs
        // afterwards cannot tell "the restore preserved these" from "they were never written".
        Assert.Equal("post-checkpoint-value", ReadCell(TargetSheet, "A1", "Value2"));
        Assert.Equal(PostCheckpointNumberFormat, ReadCell(TargetSheet, "A1", "NumberFormat"));
        Assert.Equal("=1+2", ReadCell(TargetSheet, "B1", "Formula"));

        var restored = RestoreExplicitly(home, snapshotId, evidence, "restore");
        Assert.True(DocBridgeHost.IsVerifiedRestore(restored), restored.ToJsonString());
        AssertDeferredRestoreShape(restored, "happy-path restore");

        var after = ScanTargetStyles();
        for (var i = 0; i < Scale; i++)
            AssertStyleEqual(before[i], after[i], i);

        Assert.Equal("post-checkpoint-value", ReadCell(TargetSheet, "A1", "Value2"));
        Assert.Equal(PostCheckpointNumberFormat, ReadCell(TargetSheet, "A1", "NumberFormat"));
        Assert.Equal("=1+2", ReadCell(TargetSheet, "B1", "Formula"));
        Assert.Equal(KeepMarker, ReadCell(KeepSheet, "A1", "Value2"));
        AssertSessionUnchanged(identityBefore);
    }

    // -------------------------------------- 2. real partial mutation, real rollback

    [Fact]
    public void A_real_partial_mutation_then_failure_is_rolled_back_and_verified()
    {
        if (!Enabled) return;
        RequirePrerequisites();

        using var scope = NewCase("partial-failure");
        var home = scope.Home;
        var evidence = scope.Evidence;
        // Deferred must be explicitly ON here: a partial-failure rollback that silently ran the
        // eager path would prove nothing about the deferred route.
        OpenFixture();
        SetDeferredEnv(true);
        var wrapper = new PrefixFailureAdapter(RequireAdapter()) { FailAfterRealPrefixWrites = 3 };
        var host = NewHost(home, wrapper);

        DirtyOwnedMarker("pre-partial");
        var before = ScanTargetStyles();
        var identityBefore = ReadIdentity();

        var executeWatch = Stopwatch.StartNew();
        var execute = host.ApplyOps("excel", FormatBatch());
        executeWatch.Stop();
        Artifact(evidence, "execute-partial-failure", execute);

        // One honest wall number for the whole failed apply: the real prefix writes, the injected
        // failure and the host's automatic rollback all happened inside it. The product's own
        // timings are cloned alongside, never edited, and no assertion reads a duration.
        Artifact(evidence, "execute-partial-failure-timing", new JsonObject
        {
            ["label"] = "execute-partial-failure",
            ["scale"] = Scale,
            ["applyWithAutoRollbackWallMs"] = executeWatch.ElapsedMilliseconds,
            ["timings"] = Json.GetObj(execute, "timings")?.DeepClone(),
            ["rollbackAttempted"] = Json.GetBool(Json.GetObj(execute, "rollback"), "attempted"),
            ["rollbackVerified"] = Json.GetBool(Json.GetObj(execute, "rollback"), "verified"),
            ["wallClockIsReportNotAssert"] = true,
        });

        Assert.False(Json.GetBool(execute, "ok"), execute.ToJsonString());
        Assert.True(wrapper.PrefixWritesPerformed >= 1, "no real mutation happened before the throw");

        var snapshotId = Json.GetString(execute, "snapshotId");
        Assert.False(string.IsNullOrWhiteSpace(snapshotId), execute.ToJsonString());
        AssertDeferredWasUsed(home, snapshotId!);

        var rollback = Json.GetObj(execute, "rollback");
        Assert.NotNull(rollback);
        Assert.True(Json.GetBool(rollback, "attempted"), rollback!.ToJsonString());
        Assert.True(Json.GetBool(rollback, "verified"), rollback.ToJsonString());
        AssertDeferredRestoreShape(
            Json.GetObj(rollback, "result") ?? rollback, "auto rollback");
        Assert.Contains(
            Json.GetArr(execute, "warnings")?.Select(node => node?.ToString() ?? "") ?? [],
            warning => warning.Contains("restored automatically", StringComparison.Ordinal));

        var after = ScanTargetStyles();
        for (var i = 0; i < Scale; i++)
            AssertStyleEqual(before[i], after[i], i);
        Assert.Equal(KeepMarker, ReadCell(KeepSheet, "A1", "Value2"));
        AssertSessionUnchanged(identityBefore);
    }

    // -------------------------------- 3. corrupted checkpoint refused, then rescanned

    [Fact]
    public void A_corrupted_checkpoint_is_refused_and_the_live_workbook_is_provably_untouched()
    {
        if (!Enabled) return;
        RequirePrerequisites();

        using var scope = NewCase("corrupted-checkpoint");
        var home = scope.Home;
        var evidence = scope.Evidence;
        var host = OpenFixtureAndHost(home, deferred: true);
        DirtyOwnedMarker("pre-corrupt");

        var execute = host.ApplyOps("excel", FormatBatch());
        Assert.True(Json.GetBool(execute, "ok"), execute.ToJsonString());
        var snapshotId = Json.GetString(execute, "snapshotId")!;
        AssertDeferredWasUsed(home, snapshotId);

        var checkpoint = RequireCheckpoint(home, snapshotId);
        var appliedBeforeRefusal = ScanTargetStyles();
        var identityBefore = ReadIdentity();

        var bytes = File.ReadAllBytes(checkpoint);
        bytes[^17] ^= 0x5A;
        File.WriteAllBytes(checkpoint, bytes);

        var dry = host.CoreRestoreSnapshot(new JsonObject { ["snapshotId"] = snapshotId });
        var attempted = host.CoreRestoreSnapshot(new JsonObject
        {
            ["snapshotId"] = snapshotId,
            ["confirmToken"] = Json.GetString(dry, "confirmToken"),
        });
        Artifact(evidence, "restore-corrupted", attempted);

        Assert.False(DocBridgeHost.IsVerifiedRestore(attempted), attempted.ToJsonString());
        Assert.True(File.Exists(checkpoint), "the checkpoint must be retained as evidence");

        // Fresh native rescan: the refusal must land before the checkpoint is opened, so the live
        // styles are exactly what the successful apply left.
        var afterRefusal = ScanTargetStyles();
        for (var i = 0; i < Scale; i++)
            AssertStyleEqual(appliedBeforeRefusal[i], afterRefusal[i], i);
        AssertSessionUnchanged(identityBefore);
    }

    // ------------------------------- 4. worker metadata round trip + second process

    [Fact]
    public void The_cli_worker_records_deferred_metadata_and_a_second_process_restores_it()
    {
        if (!Enabled) return;
        RequirePrerequisites();
        RequireCli();

        using var scope = NewCase("cli-round-trip");
        var home = scope.Home;
        var evidence = scope.Evidence;
        OpenFixture();
        DirtyOwnedMarker("pre-cli");
        var before = ScanTargetStyles();
        var identityBefore = ReadIdentity();

        var execute = RunCli(CliPath, home, "excel_apply_ops", FormatBatch(), deferred: true);
        Artifact(evidence, "cli-execute", execute);
        Assert.True(Json.GetBool(execute, "ok"), execute.ToJsonString());
        var snapshotId = Json.GetString(execute, "snapshotId")!;

        // Deferred metadata exists on the CLI path only because the worker round-trip returns and
        // merges the whole mutated metadata object.
        var metadata = AssertDeferredWasUsed(home, snapshotId);
        Assert.Equal("execute", Json.GetString(metadata, HostSnapshotCaptureContext.MetadataKey));

        // Newer, unrelated edits made after the checkpoint. The worker/CLI lifecycle is the path
        // that actually ships, so "unrelated values, formulas and other sheets survive" has to be
        // proven here and not only in the in-process case.
        SetCell(TargetSheet, "A1", "cli-post-checkpoint-value", PostCheckpointNumberFormat);
        SetFormula(TargetSheet, "B1", "=2+3");

        // Same reason as the in-process case: prove natively that these edits are really there
        // before a separate process restores over them.
        Assert.Equal("cli-post-checkpoint-value", ReadCell(TargetSheet, "A1", "Value2"));
        Assert.Equal(PostCheckpointNumberFormat, ReadCell(TargetSheet, "A1", "NumberFormat"));
        Assert.Equal("=2+3", ReadCell(TargetSheet, "B1", "Formula"));

        // A genuinely separate process consumes the persisted snapshot while our Excel and the
        // fixture stay alive in their original process.
        var dry = RunCli(CliPath, home, "core_restore_snapshot",
            new JsonObject { ["snapshotId"] = snapshotId }, deferred: true);
        Assert.True(Json.GetBool(dry, "ok"), dry.ToJsonString());
        var cliRestoreWatch = Stopwatch.StartNew();
        var restored = RunCli(CliPath, home, "core_restore_snapshot", new JsonObject
        {
            ["snapshotId"] = snapshotId,
            ["confirmToken"] = Json.GetString(dry, "confirmToken"),
        }, deferred: true);
        cliRestoreWatch.Stop();
        Artifact(evidence, "cli-restore", restored);

        // Process spawn, host start-up, COM attach and the restore itself are all inside this one
        // number. It is deliberately not comparable with the in-process restore wall clock, and is
        // recorded under its own key so nobody reads it as one.
        Artifact(evidence, "cli-restore-timing", new JsonObject
        {
            ["label"] = "cli-restore",
            ["scale"] = Scale,
            ["snapshotId"] = snapshotId,
            ["cliConfirmedRestoreProcessWallMs"] = cliRestoreWatch.ElapsedMilliseconds,
            ["extractMs"] = Json.GetObj(restored, "deferredExtraction")?["extractMs"]?.DeepClone(),
            ["wallClockIsReportNotAssert"] = true,
        });

        Assert.True(Json.GetBool(restored, "ok"), restored.ToJsonString());
        Assert.True(DocBridgeHost.IsVerifiedRestore(restored), restored.ToJsonString());
        AssertDeferredRestoreShape(restored, "cli restore");

        var after = ScanTargetStyles();
        for (var i = 0; i < Scale; i++)
            AssertStyleEqual(before[i], after[i], i);

        Assert.Equal("cli-post-checkpoint-value", ReadCell(TargetSheet, "A1", "Value2"));
        Assert.Equal(PostCheckpointNumberFormat, ReadCell(TargetSheet, "A1", "NumberFormat"));
        Assert.Equal("=2+3", ReadCell(TargetSheet, "B1", "Formula"));
        Assert.Equal(KeepMarker, ReadCell(KeepSheet, "A1", "Value2"));
        AssertSessionUnchanged(identityBefore);
    }

    // ----------------------------------- 5. old binaries refuse before mutation

    [Fact]
    public void Older_binaries_refuse_the_envelope_version_before_touching_the_workbook()
    {
        if (!Enabled) return;
        RequirePrerequisites();
        RequireCli();
        Assert.NotEmpty(OldCliPaths);

        using var scope = NewCase("old-cli-refusal");
        var home = scope.Home;
        var evidence = scope.Evidence;
        OpenFixture();

        // 0.4.18 / 0.4.20 restore copy-sheet-topology through ActiveWorkbook and require the
        // snapshot documentRef to equal that FullName *before* they reach the version check. Their
        // instance picker scores by workbook count, so any other Excel would win and the refusal
        // would come from an active-workbook mismatch instead of the topology version — proving
        // nothing about the envelope. Rather than activating a user's workbook to force the issue,
        // this fails outright when the identity cannot be proven.
        Assert.True(
            _preexistingExcelPids.Count == 0,
            "old-binary refusal needs the owned instance to be the only Excel running, otherwise "
            + $"the old CLI attaches elsewhere and the version refusal is not what is being tested; "
            + $"preexisting PIDs [{string.Join(",", _preexistingExcelPids)}]");
        ActivateOwnedFixture();

        DirtyOwnedMarker("pre-old-cli");

        var execute = RunCli(CliPath, home, "excel_apply_ops", FormatBatch(), deferred: true);
        Assert.True(Json.GetBool(execute, "ok"), execute.ToJsonString());
        var snapshotId = Json.GetString(execute, "snapshotId")!;
        AssertDeferredWasUsed(home, snapshotId);

        var appliedStyles = ScanTargetStyles();
        var identityBefore = ReadIdentity();

        foreach (var oldCli in OldCliPaths)
        {
            Assert.True(File.Exists(oldCli), $"old CLI not found: {oldCli}");
            var label = OldCliLabel(oldCli);

            // Re-activate before each old binary: the new CLI's own restore may have changed which
            // workbook is active in the owned instance.
            ActivateOwnedFixture();

            // The envelope is copy-sheet-topology v4. Both 0.4.18 and 0.4.20 decode topology, so
            // the dry-run must SUCCEED and hand back a token; the refusal must then come from the
            // version comparison, before any write — not from a generic failure earlier on.
            var dry = RunCli(oldCli, home, "core_restore_snapshot",
                new JsonObject { ["snapshotId"] = snapshotId }, deferred: false, allowFailure: true);
            Artifact(evidence, $"old-cli-{label}-dry", dry);
            Assert.True(Json.GetBool(dry, "ok"), $"{label} dry-run must succeed: {dry.ToJsonString()}");
            var token = Json.GetString(dry, "confirmToken");
            Assert.False(string.IsNullOrWhiteSpace(token), $"{label} dry-run returned no confirmToken");

            var attempted = RunCli(oldCli, home, "core_restore_snapshot", new JsonObject
            {
                ["snapshotId"] = snapshotId,
                ["confirmToken"] = token,
            }, deferred: false, allowFailure: true);
            Artifact(evidence, $"old-cli-{label}-restore", attempted);

            Assert.False(DocBridgeHost.IsVerifiedRestore(attempted), attempted.ToJsonString());
            Assert.Contains(
                "unsupported copy-sheet topology snapshot version",
                AllText(attempted),
                StringComparison.OrdinalIgnoreCase);

            var after = ScanTargetStyles();
            for (var i = 0; i < Scale; i++)
                AssertStyleEqual(appliedStyles[i], after[i], i);
        }

        AssertSessionUnchanged(identityBefore);
    }

    // ------------------------------------------ 6. paired off/on execute timing

    [Fact]
    public void Paired_off_and_on_execute_timings_are_recorded_on_the_same_fixture()
    {
        if (!Enabled) return;
        RequirePrerequisites();

        using var scope = NewCase("paired-off-on");
        var home = scope.Home;
        var evidence = scope.Evidence;
        // The fixture must exist before anything reads a cell from it.
        OpenFixture();
        DirtyOwnedMarker("pre-paired");
        var before = ScanTargetStyles();

        var off = TimedExecute(home, deferred: false, out var offSnapshot);
        AssertDeferredWasNotUsed(home, offSnapshot);
        RestoreExplicitly(home, offSnapshot, evidence, "paired-off-restore", out var offRestore);
        var afterOff = ScanTargetStyles();
        for (var i = 0; i < Scale; i++) AssertStyleEqual(before[i], afterOff[i], i);

        var on = TimedExecute(home, deferred: true, out var onSnapshot);
        AssertDeferredWasUsed(home, onSnapshot);
        RestoreExplicitly(home, onSnapshot, evidence, "paired-on-restore", out var onRestore);
        var afterOn = ScanTargetStyles();
        for (var i = 0; i < Scale; i++) AssertStyleEqual(before[i], afterOn[i], i);

        var paired = new JsonObject
        {
            ["scale"] = Scale,
            ["fixture"] = _fixturePath,
            ["deferredOff"] = off,
            ["deferredOn"] = on,
            ["deferredOffRestore"] = offRestore.DeepClone(),
            ["deferredOnRestore"] = onRestore.DeepClone(),
            ["sameFixture"] = true,
            ["timingIsReportNotAssert"] = true,
        };
        Artifact(evidence, "paired-off-on-timing", paired);
        _output.WriteLine($"[deferred-timing] {paired.ToJsonString()}");
    }

    // ------------------------ 7. a formula refuses the route; the eager path still restores

    /// <summary>
    /// The other side of the seed: a workbook the byte gate must refuse, and an eager snapshot that
    /// still protects the very same edit.
    ///
    /// The reason is a formula, deliberately. An earlier version of this case relied on the Office
    /// Store add-in parts Excel's own new-workbook template injects, and that turned out to depend
    /// on how the instance started — three webextension parts in some runs, none in others — so the
    /// case could pass without ever exercising a refusal. A formula is in the workbook because this
    /// harness put it there, in a cell it owns, on a sheet the target range never touches: the
    /// refusal is the same every run, on every machine.
    ///
    /// Nothing here claims the add-in (complex-part) refusal was proven natively. It was not.
    /// </summary>
    [Fact]
    public void A_workbook_containing_a_formula_is_refused_and_the_eager_path_still_restores_it()
    {
        if (!Enabled) return;
        RequirePrerequisites();

        using var scope = NewCase("formula-refusal");
        var home = scope.Home;
        var evidence = scope.Evidence;
        OpenFixture();

        var formulaPath = CreateFormulaWorkbook();
        using var borrowed = new ActiveWorkbookScope(this, formulaPath);

        // Precondition, read back natively rather than assumed: the formula really is in the saved
        // workbook, on the Keep sheet, outside everything the format op will write.
        Assert.Equal(SafeFormula, ReadCell(KeepSheet, "C1", "Formula"));

        // Before any apply: the real gate, over the real bytes, refuses them — and refuses them for
        // the formula, not for something incidental this fixture happens to carry.
        var preflight = PreflightOpenWorkbook(formulaPath);
        Artifact(evidence, "formula-refusal-preflight", new JsonObject
        {
            ["path"] = formulaPath,
            ["preflight"] = preflight.DeepClone(),
            ["parts"] = new JsonArray(
                PackagePartNames(formulaPath).Select(name => (JsonNode)name).ToArray()),
        });
        Assert.False(Json.GetBool(preflight, "ok"), preflight.ToJsonString());
        Assert.Equal("formulas", Json.GetString(preflight, "code"));
        Assert.Contains(
            "formula", Json.GetString(preflight, "reason") ?? "", StringComparison.OrdinalIgnoreCase);

        SetDeferredEnv(true);
        var host = NewHost(home, RequireAdapter());
        DirtyOwnedMarker("pre-formula-refusal");
        var before = ScanTargetStyles();
        var identityBefore = ReadIdentity();

        var execute = host.ApplyOps("excel", FormatBatch());
        Artifact(evidence, "formula-refusal-execute", execute);
        Assert.True(Json.GetBool(execute, "ok"), execute.ToJsonString());
        var snapshotId = Json.GetString(execute, "snapshotId")!;

        // Policy on, route refused, work still protected: an ordinary eager format-only snapshot,
        // and the recorded reason is the same formulas verdict seen above.
        AssertDeferredWasNotUsed(home, snapshotId);
        var eligibility = Json.GetObj(ReadSnapshotMetadata(home, snapshotId), "deferredEligibility");
        Assert.True(
            eligibility is not null,
            "the deferred policy was on, so the refusal must be recorded on the snapshot metadata");
        Assert.Equal("formulas", Json.GetString(eligibility, "code"));

        var applied = ScanTargetStyles();
        for (var i = 0; i < Scale; i++)
        {
            Assert.Equal(BlueOle, applied[i].Fill);
            Assert.NotEqual(before[i].Fill, applied[i].Fill);
        }

        var restored = RestoreExplicitly(home, snapshotId, evidence, "formula-refusal-restore");
        Assert.True(DocBridgeHost.IsVerifiedRestore(restored), restored.ToJsonString());

        var after = ScanTargetStyles();
        for (var i = 0; i < Scale; i++)
            AssertStyleEqual(before[i], after[i], i);
        Assert.Equal(SafeFormula, ReadCell(KeepSheet, "C1", "Formula"));
        Assert.Equal(KeepMarker, ReadCell(KeepSheet, "A1", "Value2"));
        AssertSessionUnchanged(identityBefore);
    }

    /// <summary>
    /// A second workbook this harness owns, built from the same offline seed as the fixture and
    /// carrying one safe local formula on the Keep sheet, saved as .xlsx (FileFormat 51) in the
    /// output directory and kept as evidence. It is created inside the instance we already proved
    /// is ours; no user workbook, template, add-in or setting is involved.
    /// </summary>
    private string CreateFormulaWorkbook()
    {
        var path = Path.Combine(OutputDir, $"deferred-formula-{Scale}-{Guid.NewGuid():N}.xlsx");
        var seedPath = _seedPath;
        RequireAdapter().RunOnAdapterThread<object?>(() =>
        {
            dynamic application = _ownedApp!;
            dynamic workbook = OpenSeedWorkbook(application, seedPath);
            PopulateFixtureSheets(workbook);

            // Written BEFORE the save, so the formula is in the bytes the gate reads. =1+2 refers
            // to nothing outside itself: no other cell, sheet, workbook or external source.
            workbook.Worksheets.Item(KeepSheet).Range("C1").Formula = SafeFormula;
            workbook.SaveAs(path, 51);
            workbook.Activate();
            return null;
        });

        _output.WriteLine($"[deferred] formula workbook {path} ({new FileInfo(path).Length} bytes)");
        return path;
    }

    // ------------------------------------------------------------------ assertions

    private JsonObject AssertDeferredWasUsed(TestHome home, string snapshotId)
    {
        var metadata = ReadSnapshotMetadata(home, snapshotId);
        var eligibility = Json.GetObj(metadata, "deferredEligibility");
        Assert.True(
            eligibility is not null && Json.GetBool(eligibility, "used"),
            $"deferred route was not used; metadata={metadata.ToJsonString()}");
        Assert.Equal(ExcelAdapter.DeferredFormatRestoreMode, Json.GetString(metadata, "restoreMode"));
        Assert.Equal(ExcelAdapter.DeferredFormatPayloadMode, Json.GetString(metadata, "payloadMode"));
        Assert.False(string.IsNullOrWhiteSpace(Json.GetString(metadata, "deferredCheckpointSha256")));
        Assert.False(metadata.ContainsKey("formatFingerprint"));
        Assert.False(metadata.ContainsKey("snapshotReuseVersion"));

        var statePath = Path.Combine(SnapshotDir(home, snapshotId), "state.json");
        Assert.True(File.Exists(statePath), $"state.json missing at {statePath}");
        var state = JsonNode.Parse(File.ReadAllText(statePath)) as JsonObject ?? new JsonObject();
        Assert.True(ExcelAdapter.IsDeferredFormatEnvelope(state), state.ToJsonString());

        var binding = ExcelAdapter.ValidateDeferredFormatRestoreBinding(metadata, state, statePath);
        Assert.True(Json.GetBool(binding, "ok"), binding.ToJsonString());
        Assert.True(File.Exists(RequireCheckpoint(home, snapshotId)));
        return metadata;
    }

    /// <summary>
    /// The other half of the gate: with the policy off, the same batch on the same fixture must
    /// still produce an ordinary eager format-only snapshot and no deferred envelope at all.
    /// </summary>
    private static void AssertDeferredWasNotUsed(TestHome home, string snapshotId)
    {
        var metadata = ReadSnapshotMetadata(home, snapshotId);
        var eligibility = Json.GetObj(metadata, "deferredEligibility");
        Assert.True(
            eligibility is null || !Json.GetBool(eligibility, "used"),
            $"deferred route was used with the policy off; metadata={metadata.ToJsonString()}");
        Assert.NotEqual(ExcelAdapter.DeferredFormatRestoreMode, Json.GetString(metadata, "restoreMode"));

        var statePath = Path.Combine(SnapshotDir(home, snapshotId), "state.json");
        Assert.True(File.Exists(statePath), $"state.json missing at {statePath}");
        var state = JsonNode.Parse(File.ReadAllText(statePath)) as JsonObject ?? new JsonObject();
        Assert.False(ExcelAdapter.IsDeferredFormatEnvelope(state), state.ToJsonString());
        Assert.False(ExcelAdapter.LooksLikeIncompleteDeferredFormatEnvelope(state), state.ToJsonString());
    }

    /// <summary>
    /// Reads the deferred shape the helper actually returned, rather than settling for a verified
    /// scoped restore. A restore that never ran the deferred extraction, or that failed to put the
    /// session back, would otherwise still pass IsVerifiedRestore.
    ///
    /// A successful restore reports restoreMode=format-only, because the scoped restore is what
    /// finally writes; deferred-format-only appears on metadata and on refusals. That is not
    /// asserted here on purpose.
    /// </summary>
    private static void AssertDeferredRestoreShape(JsonObject result, string where)
    {
        var extraction = Json.GetObj(result, "deferredExtraction");
        Assert.True(extraction is not null, $"{where}: no deferredExtraction on {result.ToJsonString()}");
        Assert.True(Json.GetBool(extraction, "attempted"), $"{where}: {extraction!.ToJsonString()}");
        Assert.True(Json.GetBool(extraction, "ok"), $"{where}: {extraction.ToJsonString()}");
        Assert.Null(Json.GetString(result, "deferredExtractionFailed"));
        Assert.True(Json.GetBool(result, "deferredRestoreRequiresLiveSession"), where);

        var session = Json.GetObj(result, "deferredSession") ?? Json.GetObj(extraction, "session");
        Assert.True(session is not null, $"{where}: no deferredSession on {result.ToJsonString()}");
        Assert.True(Json.GetBool(session, "enableEventsRestored"), $"{where}: {session!.ToJsonString()}");
        Assert.True(Json.GetBool(session, "automationSecurityRestored"), $"{where}: {session.ToJsonString()}");
        Assert.True(Json.GetBool(session, "workbookCountRestored"), $"{where}: {session.ToJsonString()}");
        Assert.False(Json.GetBool(session, "cleanupFailed"), $"{where}: {session.ToJsonString()}");
        Assert.Empty(Json.GetArr(session, "cleanupErrors") ?? new JsonArray());
    }

    private static void AssertStyleEqual(CellStyle expected, CellStyle actual, int index) =>
        Assert.True(expected == actual, $"cell {index}: expected {expected}, got {actual}");

    /// <summary>
    /// Session invariants that a checkpoint and a scoped restore must not disturb. The workbook is
    /// deliberately dirty on both sides: the copy reads memory without saving, so a restore that
    /// cleared the dirty flag would have destroyed the user's unsaved-state signal.
    /// </summary>
    private void AssertSessionUnchanged(JsonObject before)
    {
        var after = ReadIdentity();
        foreach (var key in new[]
                 {
                     "fullName", "workbookCount", "activeWorkbook",
                     "screenUpdating", "enableEvents", "automationSecurity", "displayAlerts",
                 })
        {
            Assert.Equal(Json.GetString(before, key) ?? before[key]?.ToJsonString(),
                Json.GetString(after, key) ?? after[key]?.ToJsonString());
        }

        Assert.False(Json.GetBool(before, "saved"), "baseline should have been dirty");
        Assert.False(
            Json.GetBool(after, "saved"),
            "the workbook must still be dirty: nothing here may save it or clear Saved");
    }

    // --------------------------------------------------------------- Excel ownership

    private static List<int> ExcelPids()
    {
        var ids = new List<int>();
        foreach (var process in Process.GetProcessesByName("EXCEL"))
        {
            using (process)
            {
                try { ids.Add(process.Id); } catch { }
            }
        }

        return ids;
    }

    private ExcelAdapter RequireAdapter()
    {
        if (_adapter is not null) return _adapter;

        // Everything — creating the instance, proving the PID, and building the fixture workbook —
        // happens inside the factory, on the adapter's own STA thread.
        //
        // The fixture cannot be built afterwards. The adapter's idle lifecycle check disconnects an
        // owned Excel that has zero workbooks, which is correct production behaviour: holding an
        // RCW on a workbook-less instance is exactly what leaves an EXCEL.EXE remnant. A factory
        // that returned a bare Application therefore had its RCW torn down moments after
        // GetStatus() returned, and the next property read failed with
        // InvalidComObjectException. The owned instance must never cross a status boundary with no
        // documents open, and production must not be bent to accommodate a test.
        _fixturePath = Path.Combine(OutputDir, $"deferred-fixture-{Scale}-{Guid.NewGuid():N}.xlsx");
        _activeBookPath = _fixturePath;
        var fixturePath = _fixturePath;

        // The fixture starts from bytes this harness wrote, not from Excel's new-workbook
        // template. Runs on this machine have seen that template embed the user's Office Store
        // add-ins (xl/webextensions/taskpanes.xml + webextension1..3.xml) into the saved package,
        // including through an explicit Workbooks.Add(xlWBATWorksheet), and other runs have seen
        // it embed none — it depends on how the instance started. The byte gate refuses those
        // parts as complex-part, which is correct and stays: nothing here weakens the gate, strips
        // parts from a checkpoint, or touches add-ins, templates or Excel settings. The seed makes
        // the fixture independent of which way that goes. It is written offline, before the owned
        // instance exists, and left in the output directory as evidence of exactly what the
        // fixture was made from.
        _seedPath = Path.Combine(OutputDir, $"deferred-seed-{Scale}-{Guid.NewGuid():N}.xlsx");
        MinimalXlsxSeed.Write(_seedPath);
        var seedPath = _seedPath;
        _output.WriteLine($"[deferred] seed {_seedPath} ({new FileInfo(_seedPath).Length} bytes)");

        _adapter = new ExcelAdapter(
            () =>
            {
                var type = Type.GetTypeFromProgID("Excel.Application")
                           ?? throw new InvalidOperationException("Excel is not installed");
                var app = Activator.CreateInstance(type)
                          ?? throw new InvalidOperationException("Excel.Application could not be created");

                object? workbook = null;

                // Nothing may be done TO the instance until it is proven ours. Until this flips,
                // the catch below only lets go of the reference: an instance that turned out to be
                // the user's must not be hidden, quit, or final-released by us.
                var ownedProven = false;
                try
                {
                    dynamic application = app;

                    // Prove ownership before touching anything. An unreadable handle, or one that
                    // maps to a process we did not start, is a hard failure: without a proven
                    // identity nothing later in this harness means anything.
                    var handle = Convert.ToInt64(application.Hwnd, CultureInfo.InvariantCulture);
                    if (handle == 0)
                        throw new InvalidOperationException(
                            "Excel reported no window handle immediately after creation; ownership unprovable");
                    var pid = RotHelper.ProcessIdFromWindowHandle(handle);
                    if (pid == 0)
                        throw new InvalidOperationException(
                            $"Excel Hwnd {handle} did not map to a process id; ownership unprovable");
                    if (_preexistingExcelPids.Contains(pid))
                        throw new InvalidOperationException(
                            $"factory attached preexisting Excel PID {pid}; refusing to operate on a user instance");

                    ownedProven = true;

                    // Only now is it safe to change the instance's state.
                    application.Visible = false;

                    workbook = BuildFixtureWorkbook(application, seedPath, fixturePath);
                    ((dynamic)workbook).Activate();

                    _ownedApp = app;
                    _ownedHwnd = handle;
                    _ownedPid = pid;
                    _ownershipProven = true;
                    return app;
                }
                catch
                {
                    if (ownedProven)
                    {
                        if (workbook is not null)
                        {
                            try { ((dynamic)workbook).Close(false); } catch { }
                        }

                        try { ((dynamic)app).Quit(); } catch { }
                        RotHelper.ReleaseComObject(app);
                    }
                    else
                    {
                        // Unknown or preexisting instance: balance the one reference we took and
                        // leave the process exactly as we found it.
                        RotHelper.ReleaseComReference(app);
                    }

                    throw;
                }
                finally { RotHelper.ReleaseComReference(workbook); }
            },
            appFactoryOwnsInstance: true);

        var status = _adapter.GetStatus();
        Assert.True(status.Available, $"Excel adapter is not available: {status.Detail}");
        Assert.True(_ownershipProven && _ownedApp is not null && _ownedPid != 0,
            "the adapter factory did not produce a proven-owned Excel instance");
        _output.WriteLine(
            $"[deferred] owned Excel PID {_ownedPid} (hwnd {_ownedHwnd}); " +
            $"preexisting [{string.Join(",", _preexistingExcelPids)}]");

        VerifyFixtureOnDisk();
        return _adapter;
    }

    /// <summary>
    /// Builds and saves the owned fixture. Runs inside the factory, on the STA thread, before the
    /// Application is handed back, so the instance is never workbook-less at a status boundary.
    /// </summary>
    private static object BuildFixtureWorkbook(dynamic application, string seedPath, string path)
    {
        dynamic workbook = OpenSeedWorkbook(application, seedPath);
        PopulateFixtureSheets(workbook);
        workbook.SaveAs(path, 51);
        return (object)workbook;
    }

    /// <summary>
    /// Opens the offline seed rather than Workbooks.Add. A bare Add(), and an explicit
    /// Add(xlWBATWorksheet) too, builds from whatever template the environment is configured with —
    /// which on some starts of this machine's Excel embeds Office Store add-in parts into the saved
    /// package and on others does not. The seed removes that variable entirely. UpdateLinks 0
    /// (never) because the seed has no links, ReadOnly false because we populate it and SaveAs from
    /// it. See https://learn.microsoft.com/en-us/office/vba/api/excel.workbooks.open
    ///
    /// Alerts are suppressed across the Open alone and restored immediately after. If Excel ever
    /// decided this package needed repair it would raise a modal prompt with nobody to dismiss it
    /// and hang the run. This is the instance we own, so no user setting or profile is involved,
    /// and a silently repaired seed is still caught downstream: VerifyFixtureOnDisk runs the real
    /// byte gate over what was saved.
    /// </summary>
    private static dynamic OpenSeedWorkbook(dynamic application, string seedPath)
    {
        var alerts = Convert.ToBoolean(application.DisplayAlerts, CultureInfo.InvariantCulture);
        application.DisplayAlerts = false;
        try { return application.Workbooks.Open(seedPath, 0, false); }
        finally { application.DisplayAlerts = alerts; }
    }

    /// <summary>
    /// Writes the Target and Keep sheets that both the seeded fixture and the environment-default
    /// workbook of case 7 carry, so the only difference between them is the package they came from.
    /// </summary>
    private static void PopulateFixtureSheets(dynamic workbook)
    {
        dynamic target = workbook.Worksheets.Item(1);
        target.Name = TargetSheet;

        // Plain data only, plain RGB only. v1 refuses checkpoints containing formulas, merges,
        // rich text, or non-theme RGB with tint, so the fixture must contain none of them.
        // Formulas are added later, after the checkpoint, on purpose.
        for (var i = 0; i < Scale; i++)
        {
            var (row, column) = Position(i);
            dynamic cell = target.Cells[row, column];
            cell.Value2 = $"s{i}";
            cell.Font.Bold = i % 2 == 0;
            cell.Interior.Pattern = i % 3 == 0 ? -4142 : 1;
            if (i % 3 != 0)
            {
                cell.Interior.Color = i % 3 == 1 ? 255 : 65280;
                cell.Interior.TintAndShade = 0;
            }
        }

        dynamic keep = workbook.Worksheets.Add();
        keep.Name = KeepSheet;
        keep.Range("A1").Value2 = KeepMarker;
    }

    private void VerifyFixtureOnDisk()
    {
        if (_fixtureVerified) return;
        Assert.True(File.Exists(_fixturePath), $"fixture was not written: {_fixturePath}");
        Assert.True(
            new FileInfo(_fixturePath).Length <= ExcelAdapter.DeferredFormatMaxFileBytes,
            "generated fixture exceeds the 16 MiB rollout cap");

        // Prove the fixture satisfies the real byte gate now, rather than discovering a
        // fullCalcOnLoad or style surprise as an opaque eligibility rejection later.
        var preflight = PreflightOpenWorkbook(_fixturePath);
        Assert.True(
            Json.GetBool(preflight, "ok"),
            $"generated fixture fails the deferred byte preflight: {preflight.ToJsonString()}");

        _fixtureVerified = true;
        _output.WriteLine($"[deferred] fixture {_fixturePath} ({new FileInfo(_fixturePath).Length} bytes)");
    }

    /// <summary>
    /// Runs the real byte gate over a workbook our own Excel currently has open for writing.
    ///
    /// The path overload opens FileShare.Read, which is right for production: a checkpoint is a
    /// copy nobody else holds, and refusing to read a file another process has open for writing is
    /// the safe answer there. That locking must not change. But these harness prechecks run against
    /// ORIGINAL live workbooks, so FileShare.Read is a guaranteed sharing violation and every
    /// verdict would come back "input-missing" before a single measurement. Read the same bytes
    /// with FileShare.ReadWrite and hand the stream to the existing Stream overload: identical
    /// gate, identical verdict, no lock we are not entitled to. These are fixture sanity checks
    /// only — never rollback evidence.
    /// </summary>
    private static JsonObject PreflightOpenWorkbook(string path)
    {
        using var bytes = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return ExcelAdapter.PrefightDeferredFormatCheckpoint(bytes);
    }

    /// <summary>Package part names of a workbook our own Excel still holds open.</summary>
    private static IReadOnlyList<string> PackagePartNames(string path)
    {
        using var bytes = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var zip = new ZipArchive(bytes, ZipArchiveMode.Read);
        return zip.Entries
            .Select(entry => entry.FullName)
            .OrderBy(name => name, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    /// <summary>The factory already created it; this only guarantees the adapter exists.</summary>
    private void OpenFixture() => RequireAdapter();

    /// <summary>
    /// Brings the owned fixture to the front of the owned instance. The old CLI attaches through
    /// ActiveWorkbook, so it must be looking at our workbook, and never at a user's.
    /// </summary>
    private void ActivateOwnedFixture() =>
        RequireAdapter().RunOnAdapterThread<object?>(() =>
        {
            dynamic workbook = WorkbookAt(_fixturePath);
            workbook.Activate();
            return null;
        });

    private static (int Row, int Column) Position(int index) => (index / 10 + 1, index % 10 + 1);

    /// <summary>
    /// Leaves the owned workbook dirty through a cell this harness owns on its own sheet, so a
    /// later "still dirty" assertion means something. Never writes Saved directly.
    /// </summary>
    private void DirtyOwnedMarker(string tag)
    {
        OpenFixture();

        // A value write is the whole job here: it dirties the workbook and touches only a cell
        // this harness owns. The marker used to reassign NumberFormat = "General" as well, which
        // this Excel refuses through the late-bound setter, failing the case before any host
        // Apply ran. Nothing about "the workbook is dirty" ever needed a format.
        SetValue(KeepSheet, "B1", $"dirty-{tag}-{Guid.NewGuid():N}");
        Assert.False(
            Json.GetBool(ReadIdentity(), "saved"),
            "writing an owned marker should have left the workbook dirty");
    }

    // --------------------------------------------------------------- host plumbing

    /// <summary>
    /// TestHome deletes its directory on dispose, which would take every snapshot, metadata.json
    /// and checkpoint with it. The scope copies the whole home into a uniquely named directory
    /// under the supplied output directory first, so evidence survives the run.
    /// </summary>
    private sealed class CaseScope : IDisposable
    {
        public CaseScope(TestHome home, EvidenceSink evidence)
        {
            Home = home;
            Evidence = evidence;
        }

        public TestHome Home { get; }

        public EvidenceSink Evidence { get; }

        public void Dispose()
        {
            Evidence.CaptureHome();
            Home.Dispose();
        }
    }

    /// <summary>
    /// Points the cell and identity helpers at another workbook this harness created inside the
    /// owned instance, then puts the fixture back and closes the borrowed workbook without saving.
    /// It only ever names workbooks the harness itself made; a user's workbook is never a candidate.
    /// </summary>
    private sealed class ActiveWorkbookScope : IDisposable
    {
        private readonly ExcelDeferredHostHarnessE2ETests _owner;
        private readonly string _previous;
        private readonly string _path;

        public ActiveWorkbookScope(ExcelDeferredHostHarnessE2ETests owner, string path)
        {
            _owner = owner;
            _previous = owner._activeBookPath;
            _path = path;
            owner._activeBookPath = path;
        }

        public void Dispose()
        {
            _owner._activeBookPath = _previous;
            try
            {
                _owner.RequireAdapter().RunOnAdapterThread<object?>(() =>
                {
                    // Close(false): the saved file stays on disk as evidence, unsaved edits go.
                    _owner.WorkbookAt(_path).Close(false);
                    _owner.WorkbookAt(_previous).Activate();
                    return null;
                });
            }
            catch (Exception ex)
            {
                _owner._output.WriteLine($"[deferred] releasing borrowed workbook {_path} failed: {ex.Message}");
            }
        }
    }

    private CaseScope NewCase(string caseName)
    {
        var home = new TestHome();
        var evidence = new EvidenceSink(caseName, home, OutputDir, _output);
        _cases.Add(new JsonObject { ["case"] = caseName, ["home"] = home.Dir, ["evidence"] = evidence.Dir });
        _output.WriteLine($"[deferred] case {caseName} home {home.Dir} evidence {evidence.Dir}");
        return new CaseScope(home, evidence);
    }

    private static void SetDeferredEnv(bool enabled) =>
        Environment.SetEnvironmentVariable(
            ExcelAdapter.DeferredFormatSnapshotVariable, enabled ? "1" : null);

    private DocBridgeHost OpenFixtureAndHost(TestHome home, bool deferred)
    {
        OpenFixture();
        SetDeferredEnv(deferred);
        return NewHost(home, RequireAdapter());
    }

    private DocBridgeHost NewHost(TestHome home, IAppAdapter adapter)
    {
        OpenFixture();
        var host = new DocBridgeHost(home.Options);
        host.Router.Register("excel", adapter);
        return host;
    }

    private JsonObject FormatBatch() => new()
    {
        ["ops"] = new JsonArray
        {
            new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = new JsonObject { ["sheet"] = TargetSheet },
                ["range"] = $"A1:J{Scale / 10}",
                ["style"] = new JsonObject { ["fillColor"] = BlueOle },
            },
        },
        // executionMode=execute must NOT carry dryRun at all: OperationValidator denies the
        // batch when both are present, even with dryRun=false.
        ["executionMode"] = "execute",
        ["requestId"] = Guid.NewGuid().ToString("N"),
        ["expectedDocumentRef"] = _activeBookPath,
    };

    private JsonObject TimedExecute(TestHome home, bool deferred, out string snapshotId)
    {
        var host = OpenFixtureAndHost(home, deferred);
        var watch = Stopwatch.StartNew();
        var result = host.ApplyOps("excel", FormatBatch());
        watch.Stop();
        Assert.True(Json.GetBool(result, "ok"), result.ToJsonString());
        snapshotId = Json.GetString(result, "snapshotId")!;
        return new JsonObject
        {
            ["deferredEnabled"] = deferred,
            ["wallMs"] = watch.ElapsedMilliseconds,
            ["timings"] = Json.GetObj(result, "timings")?.DeepClone(),
            ["deferredEligibility"] =
                Json.GetObj(ReadSnapshotMetadata(home, snapshotId), "deferredEligibility")?.DeepClone(),
        };
    }

    private JsonObject RestoreExplicitly(TestHome home, string snapshotId, EvidenceSink evidence, string label) =>
        RestoreExplicitly(home, snapshotId, evidence, label, out _);

    /// <summary>
    /// The dry-run and the confirmed call are timed separately on purpose: only the second one
    /// restores anything, and a single number folding token issuance into it would overstate the
    /// restore. Both are wall clock around the host call in this process — not a substitute for the
    /// product's own timings, which are cloned beside them and never edited. Nothing here asserts
    /// on a duration; the companion artifact is a report.
    /// </summary>
    private JsonObject RestoreExplicitly(
        TestHome home, string snapshotId, EvidenceSink evidence, string label, out JsonObject timing)
    {
        var host = NewHost(home, RequireAdapter());
        var dryWatch = Stopwatch.StartNew();
        var dry = host.CoreRestoreSnapshot(new JsonObject { ["snapshotId"] = snapshotId });
        dryWatch.Stop();
        Assert.True(Json.GetBool(dry, "ok"), dry.ToJsonString());

        var confirmedWatch = Stopwatch.StartNew();
        var restored = host.CoreRestoreSnapshot(new JsonObject
        {
            ["snapshotId"] = snapshotId,
            ["confirmToken"] = Json.GetString(dry, "confirmToken"),
        });
        confirmedWatch.Stop();

        Artifact(evidence, label, restored);
        var extraction = Json.GetObj(restored, "deferredExtraction");
        timing = new JsonObject
        {
            ["label"] = label,
            ["scale"] = Scale,
            ["snapshotId"] = snapshotId,
            ["dryRunWallMs"] = dryWatch.ElapsedMilliseconds,
            ["confirmedRestoreWallMs"] = confirmedWatch.ElapsedMilliseconds,
            ["extractMs"] = extraction?["extractMs"]?.DeepClone(),
            ["restoreMode"] = Json.GetString(restored, "restoreMode"),
            ["deferredExtractionAttempted"] = extraction is not null && Json.GetBool(extraction, "attempted"),
            ["wallClockIsReportNotAssert"] = true,
        };
        Artifact(evidence, $"{label}-timing", timing);
        _output.WriteLine($"[deferred-timing] {timing.ToJsonString()}");

        Assert.True(DocBridgeHost.IsVerifiedRestore(restored), restored.ToJsonString());
        return restored;
    }

    private static string SnapshotDir(TestHome home, string snapshotId) =>
        Path.Combine(home.Dir, "snapshots", "excel", snapshotId);

    private static JsonObject ReadSnapshotMetadata(TestHome home, string snapshotId)
    {
        var path = Path.Combine(SnapshotDir(home, snapshotId), "metadata.json");
        Assert.True(File.Exists(path), $"metadata.json missing at {path}");
        return JsonNode.Parse(File.ReadAllText(path)) as JsonObject ?? new JsonObject();
    }

    private static string RequireCheckpoint(TestHome home, string snapshotId)
    {
        var dir = SnapshotDir(home, snapshotId);
        Assert.True(
            ExcelAdapter.TryResolveDeferredFormatCheckpointPath(dir, null, out var path, out var error),
            $"checkpoint path could not be resolved: {error}");
        Assert.True(File.Exists(path), $"checkpoint missing at {path}");
        return path;
    }

    private static void RequireCli()
    {
        Assert.False(string.IsNullOrWhiteSpace(CliPath),
            "DOCBRIDGE_EXCEL_DEFERRED_CLI is required when the native gate is on");
        Assert.True(File.Exists(CliPath), $"CLI not found: {CliPath}");
    }

    private static string OldCliLabel(string path)
    {
        // .18 and .20 previously collided on one artifact name. Derive a distinct, stable label
        // from the version-bearing directory plus a short digest of the full path.
        var version = Path.GetFileName(Path.GetDirectoryName(Path.GetDirectoryName(path)) ?? "");
        if (string.IsNullOrWhiteSpace(version)) version = Path.GetFileNameWithoutExtension(path);
        var safe = new string(version.Select(ch => char.IsLetterOrDigit(ch) ? ch : '-').ToArray());
        var digest = Convert.ToHexString(
            System.Security.Cryptography.SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(path)));
        return $"{safe}-{digest.ToLowerInvariant()[..8]}";
    }

    private static string AllText(JsonObject result) => result.ToJsonString();

    private JsonObject RunCli(
        string cli, TestHome home, string tool, JsonObject args, bool deferred, bool allowFailure = false)
    {
        var requestPath = Path.Combine(home.Dir, $"request-{Guid.NewGuid():N}.json");
        File.WriteAllText(requestPath, args.ToJsonString());

        var info = new ProcessStartInfo(cli)
        {
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            UseShellExecute = false,
            WorkingDirectory = OutputDir,
        };
        info.ArgumentList.Add(tool);
        info.ArgumentList.Add("--json-file");
        info.ArgumentList.Add(requestPath);
        info.Environment["DOCBRIDGE_HOME"] = home.Dir;
        info.Environment[ExcelAdapter.DeferredFormatSnapshotVariable] = deferred ? "1" : "0";

        using var process = Process.Start(info)
                            ?? throw new InvalidOperationException($"CLI did not start: {cli}");

        // Serial ReadToEnd on both pipes deadlocks as soon as one fills, and WaitForExit after a
        // blocking read is already too late. Drain both asynchronously, then bound the wait.
        var stdoutTask = process.StandardOutput.ReadToEndAsync();
        var stderrTask = process.StandardError.ReadToEndAsync();
        var exited = process.WaitForExit(CliTimeoutMs);
        var drained = Task.WaitAll(new Task[] { stdoutTask, stderrTask }, CliTimeoutMs);
        var stdout = stdoutTask.IsCompletedSuccessfully ? stdoutTask.Result : "";
        var stderr = stderrTask.IsCompletedSuccessfully ? stderrTask.Result : "";

        if (!exited || !drained)
        {
            // Report; never kill. A stuck CLI may still own Excel RCWs, and terminating it here
            // could strand or damage the user's Office session.
            var timeout = new JsonObject
            {
                ["ok"] = false,
                ["errors"] = new JsonArray(
                    $"CLI {Path.GetFileName(cli)} {tool} did not finish within {CliTimeoutMs} ms " +
                    $"(exited={exited}, drained={drained}); left running deliberately, not killed"),
                ["stdout"] = stdout,
                ["stderr"] = stderr,
            };
            _output.WriteLine($"[deferred] {timeout.ToJsonString()}");
            if (allowFailure) return timeout;
            Assert.Fail(timeout.ToJsonString());
        }

        _output.WriteLine($"[deferred] cli {Path.GetFileName(cli)} {tool} exit={process.ExitCode}");
        if (!string.IsNullOrWhiteSpace(stderr)) _output.WriteLine($"[deferred] cli stderr {stderr}");

        var parsed = JsonNode.Parse(stdout) as JsonObject;
        if (parsed is null)
        {
            var nonJson = new JsonObject
            {
                ["ok"] = false,
                ["errors"] = new JsonArray($"CLI returned non-JSON (exit {process.ExitCode})"),
                ["stdout"] = stdout,
                ["stderr"] = stderr,
            };
            if (allowFailure) return nonJson;
            Assert.Fail(nonJson.ToJsonString());
        }

        return parsed!;
    }

    // ----------------------------------------------------------- native cell reads

    private readonly record struct CellStyle(double Fill, int Pattern, int ColorIndex, bool Bold)
    {
        public override string ToString() =>
            $"fill={Fill} pattern={Pattern} colorIndex={ColorIndex} bold={Bold}";
    }

    private CellStyle[] ScanTargetStyles()
    {
        Assert.False(
            string.IsNullOrEmpty(_activeBookPath),
            "ScanTargetStyles was called before the fixture existed");
        return RequireAdapter().RunOnAdapterThread(() =>
        {
            dynamic sheet = ActiveOwnedWorkbook().Worksheets.Item(TargetSheet);
            var styles = new CellStyle[Scale];
            for (var i = 0; i < Scale; i++)
            {
                var (row, column) = Position(i);
                dynamic cell = sheet.Cells[row, column];
                styles[i] = new CellStyle(
                    Convert.ToDouble(cell.Interior.Color, CultureInfo.InvariantCulture),
                    Convert.ToInt32(cell.Interior.Pattern, CultureInfo.InvariantCulture),
                    Convert.ToInt32(cell.Interior.ColorIndex, CultureInfo.InvariantCulture),
                    Convert.ToBoolean(cell.Font.Bold, CultureInfo.InvariantCulture));
            }

            return styles;
        });
    }

    /// <summary>One owned workbook by the file name it was saved under.</summary>
    private dynamic WorkbookAt(string path) =>
        ((dynamic)_ownedApp!).Workbooks.Item(Path.GetFileName(path));

    /// <summary>
    /// The owned workbook the cell and identity helpers read and write. Normally the fixture; case
    /// 7 points it at the environment-default workbook for that case only.
    /// </summary>
    private dynamic ActiveOwnedWorkbook() => WorkbookAt(_activeBookPath);

    private JsonObject ReadIdentity() =>
        RequireAdapter().RunOnAdapterThread(() =>
        {
            dynamic app = _ownedApp!;
            dynamic workbook = ActiveOwnedWorkbook();
            return new JsonObject
            {
                ["fullName"] = Convert.ToString((object?)workbook.FullName, CultureInfo.InvariantCulture),
                ["saved"] = Convert.ToBoolean(workbook.Saved, CultureInfo.InvariantCulture),
                ["workbookCount"] = Convert.ToInt32(app.Workbooks.Count, CultureInfo.InvariantCulture),
                ["activeWorkbook"] =
                    Convert.ToString((object?)app.ActiveWorkbook.FullName, CultureInfo.InvariantCulture),
                ["screenUpdating"] = Convert.ToBoolean(app.ScreenUpdating, CultureInfo.InvariantCulture),
                ["enableEvents"] = Convert.ToBoolean(app.EnableEvents, CultureInfo.InvariantCulture),
                ["automationSecurity"] = Convert.ToInt32(app.AutomationSecurity, CultureInfo.InvariantCulture),
                ["displayAlerts"] = Convert.ToBoolean(app.DisplayAlerts, CultureInfo.InvariantCulture),
            };
        });

    private string? ReadCell(string sheet, string address, string property) =>
        RequireAdapter().RunOnAdapterThread(() =>
        {
            dynamic cell = ActiveOwnedWorkbook().Worksheets.Item(sheet).Range(address);
            object? raw = property switch
            {
                "Value2" => cell.Value2,
                "Formula" => cell.Formula,
                "NumberFormat" => cell.NumberFormat,
                _ => throw new ArgumentOutOfRangeException(nameof(property), property, "unsupported"),
            };
            return Convert.ToString(raw, CultureInfo.InvariantCulture);
        });

    private void SetValue(string sheet, string address, string value) =>
        RequireAdapter().RunOnAdapterThread<object?>(() =>
        {
            ActiveOwnedWorkbook().Worksheets.Item(sheet).Range(address).Value2 = value;
            return null;
        });

    private void SetCell(string sheet, string address, string value, string numberFormat) =>
        RequireAdapter().RunOnAdapterThread<object?>(() =>
        {
            dynamic cell = ActiveOwnedWorkbook().Worksheets.Item(sheet).Range(address);
            cell.Value2 = value;
            cell.NumberFormat = numberFormat;
            return null;
        });

    private void SetFormula(string sheet, string address, string formula) =>
        RequireAdapter().RunOnAdapterThread<object?>(() =>
        {
            ActiveOwnedWorkbook().Worksheets.Item(sheet).Range(address).Formula = formula;
            return null;
        });

    // --------------------------------------------------------------- evidence

    /// <summary>
    /// TestHome deletes its directory on dispose, which would take every snapshot, metadata.json
    /// and checkpoint with it. Each case copies the whole home into a uniquely named directory
    /// under the supplied output directory before that happens.
    /// </summary>
    private sealed class EvidenceSink
    {
        private readonly TestHome _home;
        private readonly ITestOutputHelper _output;

        public EvidenceSink(string caseName, TestHome home, string outputDir, ITestOutputHelper output)
        {
            _home = home;
            _output = output;
            CaseName = caseName;
            Dir = string.IsNullOrWhiteSpace(outputDir)
                ? ""
                : Path.Combine(outputDir, $"deferred-{Scale}-{caseName}-{Guid.NewGuid():N}");
            if (!string.IsNullOrWhiteSpace(Dir)) Directory.CreateDirectory(Dir);
        }

        public string CaseName { get; }

        public string Dir { get; }

        public void CaptureHome()
        {
            if (string.IsNullOrWhiteSpace(Dir) || !Directory.Exists(_home.Dir)) return;
            var destinationRoot = Path.Combine(Dir, "home");
            foreach (var source in Directory.EnumerateFiles(_home.Dir, "*", SearchOption.AllDirectories))
            {
                var relative = Path.GetRelativePath(_home.Dir, source);
                var destination = Path.Combine(destinationRoot, relative);
                Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
                try { File.Copy(source, destination, overwrite: true); }
                catch (Exception ex) { _output.WriteLine($"[deferred] evidence copy failed {relative}: {ex.Message}"); }
            }

            _output.WriteLine($"[deferred] evidence for {CaseName} at {destinationRoot}");
        }
    }

    private void Artifact(EvidenceSink evidence, string name, JsonObject payload)
    {
        if (string.IsNullOrWhiteSpace(evidence.Dir)) return;
        var path = Path.Combine(evidence.Dir, $"{name}.json");
        Directory.CreateDirectory(evidence.Dir);
        File.WriteAllText(path, payload.ToJsonString(Json.Pretty));
        _output.WriteLine($"[deferred] artifact {path}");
    }

    // ---------------------------------------------------------------- teardown

    public void Dispose()
    {
        SetDeferredEnv(false);

        try
        {
            // Only touch the workbook when ownership was actually proven; the factory sets the
            // path before it runs, so a failed creation must not lead us to close anything.
            if (_ownershipProven && _adapter is not null && _ownedApp is not null && !_fixtureClosed
                && !string.IsNullOrEmpty(_fixturePath))
            {
                _adapter.RunOnAdapterThread<object?>(() =>
                {
                    dynamic workbook = WorkbookAt(_fixturePath);
                    if (string.Equals(
                            Convert.ToString((object?)workbook.FullName, CultureInfo.InvariantCulture),
                            _fixturePath, StringComparison.OrdinalIgnoreCase))
                    {
                        workbook.Close(false);
                    }

                    return null;
                });
                _fixtureClosed = true;
            }
        }
        catch (Exception ex) { _output.WriteLine($"[deferred] fixture close failed: {ex.Message}"); }

        try
        {
            if (_ownershipProven && _ownedApp is not null && _ownedPid != 0
                && !_preexistingExcelPids.Contains(_ownedPid))
                _adapter?.RunOnAdapterThread<object?>(() => { ((dynamic)_ownedApp!).Quit(); return null; });
        }
        catch (Exception ex) { _output.WriteLine($"[deferred] quit failed: {ex.Message}"); }

        try { _adapter?.Dispose(); } catch { }

        // Our own Excel can take a moment to go away after Quit. xunit builds a fresh instance of
        // this class for the next case, and its preexisting-PID census would otherwise record this
        // dying process as a user's instance and refuse to operate. So wait, bounded, on the PID we
        // proved is ours. A process still alive at the end of the wait is reported and left running
        // — nothing here kills or force-releases anything, least of all a process we did not start.
        if (_ownershipProven && _ownedPid != 0 && !_preexistingExcelPids.Contains(_ownedPid))
        {
            try
            {
                using var owned = Process.GetProcessById(_ownedPid);
                if (!owned.WaitForExit(OwnedExitWaitMs))
                {
                    _output.WriteLine(
                        $"[deferred] owned Excel PID {_ownedPid} still running after {OwnedExitWaitMs} ms; left alone");
                }
            }
            catch (ArgumentException)
            {
                // GetProcessById throws this once the process is gone, which is the outcome we want.
            }
            catch (Exception ex)
            {
                _output.WriteLine($"[deferred] waiting for owned PID {_ownedPid} failed: {ex.Message}");
            }
        }

        foreach (var pid in _preexistingExcelPids)
        {
            _output.WriteLine(
                Process.GetProcesses().Any(process => process.Id == pid)
                    ? $"[deferred] preexisting Excel PID {pid} preserved"
                    : $"[deferred] WARNING preexisting Excel PID {pid} is gone");
        }

        if (!string.IsNullOrWhiteSpace(OutputDir))
        {
            File.WriteAllText(
                Path.Combine(OutputDir, $"deferred-{Scale}-harness-summary-{Guid.NewGuid():N}.json"),
                new JsonObject
                {
                    ["scale"] = Scale,
                    ["fixture"] = _fixturePath,
                    ["seed"] = _seedPath,
                    ["ownedPid"] = _ownedPid,
                    ["preexistingPids"] =
                        new JsonArray(_preexistingExcelPids.Select(pid => (JsonNode)pid).ToArray()),
                    ["cases"] = _cases.DeepClone(),
                }.ToJsonString(Json.Pretty));
        }
    }

    /// <summary>
    /// Delegates to the real adapter but performs a genuine narrowed mutation and then throws, so
    /// the host's real automatic rollback runs against real damage. Never suppresses an error.
    /// </summary>
    private sealed class PrefixFailureAdapter : IAppAdapter, IPreviewReuseAdapter
    {
        private readonly ExcelAdapter _inner;

        public PrefixFailureAdapter(ExcelAdapter inner) => _inner = inner;

        public int FailAfterRealPrefixWrites { get; set; }

        public int PrefixWritesPerformed { get; private set; }

        public string App => _inner.App;

        public AdapterStatus GetStatus() => _inner.GetStatus();

        public JsonObject GetCapabilities() => _inner.GetCapabilities();

        public ContextResult GetActiveContext() => _inner.GetActiveContext();

        public JsonObject Read(JsonObject args) => _inner.Read(args);

        public ApplyPreview Preview(IReadOnlyList<JsonObject> ops) => _inner.Preview(ops);

        public ApplyExecution Apply(IReadOnlyList<JsonObject> ops, string snapshotId)
        {
            if (FailAfterRealPrefixWrites <= 0) return _inner.Apply(ops, snapshotId);

            var prefix = new List<JsonObject>();
            foreach (var op in ops)
            {
                var clone = op.DeepClone().AsObject();
                clone["range"] = $"A1:{(char)('A' + FailAfterRealPrefixWrites - 1)}1";
                prefix.Add(clone);
            }

            var partial = _inner.Apply(prefix, snapshotId);
            PrefixWritesPerformed = partial.Ok ? FailAfterRealPrefixWrites : 0;
            throw new InvalidOperationException(
                $"injected failure after {PrefixWritesPerformed} real cell writes " +
                "(harness only; those writes are genuine and must be rolled back)");
        }

        public void CaptureSnapshot(string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject>? ops = null) =>
            _inner.CaptureSnapshot(snapshotDir, metadata, ops);

        public JsonObject RestoreSnapshot(string snapshotDir, JsonObject metadata) =>
            _inner.RestoreSnapshot(snapshotDir, metadata);

        public JsonObject ValidatePreviewReuse(
            string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject> ops) =>
            _inner.ValidatePreviewReuse(snapshotDir, metadata, ops);

        // The adapter is owned by the harness, which disposes it once in Dispose.
        public void Dispose() { }
    }
}
