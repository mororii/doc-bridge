using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

/// <summary>
/// The auxiliary workbook copy that accompanies an Excel snapshot has no automated
/// restore consumer; it is operator-facing evidence beside the operation-scoped
/// state.json.
///
/// The default path copies the last-saved file, exactly as before, and must never
/// invoke SaveCopyAs — a small edit does not pay a whole-workbook serialization.
/// DOCBRIDGE_EXCEL_FRESH_WORKBOOK_BACKUP=1 opts in, and only the exact string "1"
/// counts. The policy is passed in rather than read from process state so every case
/// here is deterministic.
///
/// These tests pin that gate and, on the opt-in path, the freshness labelling and the
/// safety rules: only a verified native copy may be reported as fresh, the open
/// workbook count and the foreground guard come from the workbook's own Application
/// rather than from whatever instance the adapter happened to attach to, identity must
/// be fully readable on both sides before it counts as preserved, and an unverified
/// identity refuses the snapshot instead of letting the pending edit proceed.
///
/// No Excel process is involved. The fakes are plain objects reached through the same
/// late-bound `dynamic` call sites production uses.
/// </summary>
public class ExcelWorkbookBackupTests
{
    private const string DiskBytes = "LAST-SAVED-ON-DISK";
    private const string MemoryBytes = "CURRENT-MEMORY-WITH-UNSAVED-EDITS";
    private const string AppLabel = "excel";

    // ------------------------------------------------- default path (opt-in off)

    [Fact]
    public void Default_policy_never_invokes_savecopyas_and_labels_the_copy_stale()
    {
        using var home = new TestHome();
        var workbook = NewWorkbook(home, saved: false);
        var metadata = Capture(workbook, home, fresh: false);

        // Speed is the point: a small edit must not pay a whole-workbook serialization.
        Assert.Equal(0, workbook.SaveCopyAsCalls);
        Assert.Equal(0, workbook.Owner.ApplicationReads);
        Assert.Equal(0, workbook.FileFormatReads);
        Assert.Null(Json.GetObj(metadata, "workbookBackupInteraction"));

        Assert.False(Json.GetBool(metadata, "workbookBackupFreshCopyEnabled"));
        Assert.Equal("workbook-backup.xlsx", Json.GetString(metadata, "workbookBackup"));
        Assert.Equal("last-saved-file", Json.GetString(metadata, "workbookBackupSource"));
        Assert.False(Json.GetBool(metadata, "workbookBackupFresh"));
        Assert.True(Json.GetBool(metadata, "workbookBackupAvailable"));
        // The Saved flag is still reported, so a reader can tell the copy is missing
        // unsaved changes rather than having to guess.
        Assert.False(Json.GetBool(metadata, "workbookBackupSavedFlag"));
        Assert.Contains("DOCBRIDGE_EXCEL_FRESH_WORKBOOK_BACKUP is not set to 1",
            Json.GetString(metadata, "workbookBackupReason"));
        Assert.Contains("does not contain them", Json.GetString(metadata, "workbookBackupReason"));

        // Absent means "not evaluated"; it is not the null that means "unreadable".
        Assert.False(metadata.ContainsKey("workbookBackupFileFormat"));

        Assert.Equal(DiskBytes, File.ReadAllText(Path.Combine(home.Dir, "workbook-backup.xlsx")));
    }

    [Fact]
    public void Same_dirty_workbook_takes_the_native_copy_only_when_opted_in()
    {
        using var defaultHome = new TestHome();
        var defaultWorkbook = NewWorkbook(defaultHome, saved: false);
        var defaultMetadata = Capture(defaultWorkbook, defaultHome, fresh: false);

        using var optInHome = new TestHome();
        var optInWorkbook = NewWorkbook(optInHome, saved: false);
        var optInMetadata = Capture(optInWorkbook, optInHome, fresh: true);

        Assert.Equal(0, defaultWorkbook.SaveCopyAsCalls);
        Assert.Equal(1, optInWorkbook.SaveCopyAsCalls);
        Assert.Equal("last-saved-file", Json.GetString(defaultMetadata, "workbookBackupSource"));
        Assert.Equal("current-memory-savecopyas", Json.GetString(optInMetadata, "workbookBackupSource"));
        Assert.Equal(DiskBytes, File.ReadAllText(Path.Combine(defaultHome.Dir, "workbook-backup.xlsx")));
        Assert.Equal(MemoryBytes, File.ReadAllText(Path.Combine(optInHome.Dir, "workbook-backup.xlsx")));
    }

    [Fact]
    public void Default_policy_still_refuses_a_preexisting_target_and_unsupported_workbooks()
    {
        using var home = new TestHome();
        var workbook = NewWorkbook(home, saved: false);
        var target = Path.Combine(home.Dir, "workbook-backup.xlsx");
        File.WriteAllText(target, "DO-NOT-CLOBBER");

        var metadata = Capture(workbook, home, fresh: false);
        Assert.Equal("DO-NOT-CLOBBER", File.ReadAllText(target));
        Assert.False(Json.GetBool(metadata, "workbookBackupAvailable"));
        Assert.Contains("refusing to overwrite", Json.GetString(metadata, "workbookBackupReason"));

        using var other = new TestHome();
        var neverSaved = new FakeWorkbook { FullNameValue = "", SavedValue = false, FileFormatValue = 51 };
        var neverSavedMetadata = Capture(neverSaved, other, fresh: false);
        Assert.Equal("none", Json.GetString(neverSavedMetadata, "workbookBackupSource"));
        Assert.Contains("never been saved", Json.GetString(neverSavedMetadata, "workbookBackupReason"));
        Assert.Empty(Directory.GetFiles(other.Dir));
    }

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
    [InlineData("enabled", false)]
    [InlineData("1.0", false)]
    public void Only_the_exact_value_one_enables_the_fresh_copy(string? raw, bool expected)
    {
        Assert.Equal(
            expected,
            ExcelAdapter.ExcelWorkbookBackupPolicy.FromSetting(raw).FreshCopyEnabled);
    }

    [Fact]
    public void Environment_policy_is_wired_to_the_documented_variable()
    {
        // Reads the current process value rather than mutating it, so the assertion is
        // deterministic whatever the host environment happens to be.
        var current = Environment.GetEnvironmentVariable(ExcelAdapter.FreshWorkbookBackupVariable);
        Assert.Equal(
            ExcelAdapter.ExcelWorkbookBackupPolicy.FromSetting(current),
            ExcelAdapter.ExcelWorkbookBackupPolicy.FromEnvironment());
        Assert.Equal("DOCBRIDGE_EXCEL_FRESH_WORKBOOK_BACKUP", ExcelAdapter.FreshWorkbookBackupVariable);
    }

    // ----------------------------------------------------- opt-in path (fresh on)

    [Fact]
    public void Dirty_ordinary_xlsx_is_copied_from_memory_and_labelled_fresh()
    {
        using var home = new TestHome();
        var workbook = NewWorkbook(home, saved: false);
        var metadata = Capture(workbook, home);

        Assert.Equal("workbook-backup.xlsx", Json.GetString(metadata, "workbookBackup"));
        Assert.True(Json.GetBool(metadata, "workbookBackupFreshCopyEnabled"));
        Assert.Equal("current-memory-savecopyas", Json.GetString(metadata, "workbookBackupSource"));
        Assert.True(Json.GetBool(metadata, "workbookBackupFresh"));
        Assert.True(Json.GetBool(metadata, "workbookBackupAvailable"));
        Assert.False(Json.GetBool(metadata, "workbookBackupSavedFlag"));
        Assert.Equal(51, Json.GetInt(metadata, "workbookBackupFileFormat"));
        Assert.Null(Json.GetString(metadata, "workbookBackupNativeError"));
        Assert.Null(Json.GetString(metadata, "workbookBackupError"));
        Assert.NotNull(Json.GetObj(metadata, "workbookBackupInteraction"));

        // The point of the change: the copy holds memory, not the stale disk bytes.
        Assert.Equal(MemoryBytes, File.ReadAllText(Path.Combine(home.Dir, "workbook-backup.xlsx")));
        Assert.Equal(DiskBytes, File.ReadAllText(workbook.FullNameValue));
    }

    [Fact]
    public void Fresh_copy_preserves_workbook_identity_and_never_saves_closes_or_resets_state()
    {
        using var home = new TestHome();
        var workbook = NewWorkbook(home, saved: false);
        var metadata = Capture(workbook, home);

        Assert.True(Json.GetBool(metadata, "workbookBackupIdentityVerifiable"));
        Assert.True(Json.GetBool(metadata, "workbookBackupIdentityPreserved"));
        var identity = Json.GetObj(metadata, "workbookBackupIdentity")!;
        Assert.Equal(
            Json.GetObj(identity, "before")!.ToJsonString(),
            Json.GetObj(identity, "after")!.ToJsonString());

        Assert.False(workbook.SavedValue);
        Assert.Equal(Path.Combine(home.Dir, "book.xlsx"), workbook.FullNameValue);
        Assert.Equal(1, workbook.Owner.Workbooks.Count);
        Assert.Equal(0, workbook.SaveCalls);
        Assert.Equal(0, workbook.CloseCalls);
        Assert.Equal(0, workbook.ActivateCalls);
        Assert.Equal(0, workbook.SavedAssignments);
        Assert.Equal(0, workbook.Owner.DisplayAlertsAssignments);
        Assert.Equal(1, workbook.SaveCopyAsCalls);
    }

    // ------------------------------------- owning instance, not the attached one

    [Fact]
    public void Identity_uses_the_workbooks_own_application_not_the_attached_instance()
    {
        using var home = new TestHome();
        var workbook = NewWorkbook(home, saved: false);
        workbook.Owner.Workbooks.Count = 3;

        // ResolveTargetWorkbook can hand back a workbook from another Excel process,
        // so the attached instance's own count must never be reported.
        var attachedElsewhere = new FakeApplication();
        attachedElsewhere.Workbooks.Count = 99;

        var metadata = new JsonObject();
        ExcelAdapter.CaptureAuxiliaryWorkbookBackup(
            workbook, attachedElsewhere, home.Dir, metadata, AppLabel, Policy(fresh: true));

        Assert.Equal("workbook.Application", Json.GetString(metadata, "workbookBackupApplicationSource"));
        Assert.True(Json.GetBool(metadata, "workbookBackupAttachedApplicationProvided"));
        Assert.False(Json.GetBool(metadata, "workbookBackupAttachedApplicationUsedForCount"));

        var before = Json.GetObj(Json.GetObj(metadata, "workbookBackupIdentity"), "before")!;
        Assert.Equal(3, Json.GetInt(before, "workbookCount"));
        Assert.True(Json.GetBool(metadata, "workbookBackupIdentityPreserved"));
        Assert.Equal(1, workbook.Owner.ApplicationReads);
    }

    [Fact]
    public void Workbook_count_change_on_the_owning_instance_is_detected()
    {
        using var home = new TestHome();
        var workbook = NewWorkbook(home, saved: false);
        workbook.OwnerWorkbookCountAfterCopy = 2;

        var ex = Assert.Throws<InvalidOperationException>(() => Capture(workbook, home));
        Assert.Contains("identity is not verified after SaveCopyAs", ex.Message);
    }

    // ------------------------------------------------------------ unknown identity

    [Fact]
    public void Unreadable_owning_application_falls_back_conservatively_without_a_native_copy()
    {
        using var home = new TestHome();
        var workbook = NewWorkbook(home, saved: false);
        workbook.ApplicationThrows = true;
        var metadata = Capture(workbook, home);

        Assert.Equal("unreadable", Json.GetString(metadata, "workbookBackupApplicationSource"));
        Assert.False(Json.GetBool(metadata, "workbookBackupIdentityVerifiable"));
        Assert.False(Json.GetBool(metadata, "workbookBackupIdentityPreserved"));
        Assert.Equal("last-saved-file", Json.GetString(metadata, "workbookBackupSource"));
        Assert.False(Json.GetBool(metadata, "workbookBackupFresh"));
        Assert.Contains("could not be fully read before the copy",
            Json.GetString(metadata, "workbookBackupReason"));
        Assert.Equal(0, workbook.SaveCopyAsCalls);
        Assert.Equal(DiskBytes, File.ReadAllText(Path.Combine(home.Dir, "workbook-backup.xlsx")));
    }

    [Fact]
    public void Unreadable_workbook_count_falls_back_conservatively_without_a_native_copy()
    {
        using var home = new TestHome();
        var workbook = NewWorkbook(home, saved: false);
        workbook.Owner.WorkbooksThrows = true;
        var metadata = Capture(workbook, home);

        Assert.False(Json.GetBool(metadata, "workbookBackupIdentityVerifiable"));
        Assert.Equal("last-saved-file", Json.GetString(metadata, "workbookBackupSource"));
        Assert.False(Json.GetBool(metadata, "workbookBackupFresh"));
        Assert.Equal(0, workbook.SaveCopyAsCalls);

        var before = Json.GetObj(Json.GetObj(metadata, "workbookBackupIdentity"), "before")!;
        Assert.True(before.ContainsKey("workbookCount"));
        Assert.Null(before["workbookCount"]);
    }

    // -------------------------------------------------------- changed identity

    [Fact]
    public void Saved_flipped_during_the_native_copy_refuses_the_snapshot_and_does_not_reset_it()
    {
        using var home = new TestHome();
        var workbook = NewWorkbook(home, saved: false);
        workbook.FlipSavedDuringSaveCopyAs = true;
        var metadata = new JsonObject();

        var ex = Assert.Throws<InvalidOperationException>(() =>
            ExcelAdapter.CaptureAuxiliaryWorkbookBackup(
                workbook, workbook.Owner, home.Dir, metadata, AppLabel, Policy(fresh: true)));

        Assert.Contains("identity is not verified after SaveCopyAs", ex.Message);
        Assert.Contains("does not proceed against an unverified workbook", ex.Message);
        Assert.False(Json.GetBool(metadata, "workbookBackupIdentityPreserved"));
        Assert.NotNull(Json.GetString(metadata, "workbookBackupError"));

        // Reported, not corrected: the adapter must never write Saved or the path back.
        Assert.True(workbook.SavedValue);
        Assert.Equal(0, workbook.SavedAssignments);
        Assert.Equal(Path.Combine(home.Dir, "book.xlsx"), workbook.FullNameValue);
    }

    [Fact]
    public void Identity_that_becomes_unreadable_after_the_copy_is_not_treated_as_preserved()
    {
        using var home = new TestHome();
        var workbook = NewWorkbook(home, saved: false);
        // Both sides would serialize "saved": null and compare equal under a plain
        // JSON comparison; unreadable is not evidence of preservation.
        workbook.SavedThrowsAfterSaveCopyAs = true;
        var metadata = new JsonObject();

        Assert.Throws<InvalidOperationException>(() =>
            ExcelAdapter.CaptureAuxiliaryWorkbookBackup(
                workbook, workbook.Owner, home.Dir, metadata, AppLabel, Policy(fresh: true)));

        Assert.False(Json.GetBool(metadata, "workbookBackupIdentityPreserved"));
        var after = Json.GetObj(Json.GetObj(metadata, "workbookBackupIdentity"), "after")!;
        Assert.Null(after["saved"]);
    }

    [Fact]
    public void Changed_identity_refuses_even_when_the_native_copy_also_threw()
    {
        using var home = new TestHome();
        var workbook = NewWorkbook(home, saved: false);
        workbook.FullNameAfterSaveCopyAs = Path.Combine(home.Dir, "moved.xlsx");
        workbook.SaveCopyAsThrows = "Excel refused SaveCopyAs (1004)";
        var metadata = new JsonObject();

        var ex = Assert.Throws<InvalidOperationException>(() =>
            ExcelAdapter.CaptureAuxiliaryWorkbookBackup(
                workbook, workbook.Owner, home.Dir, metadata, AppLabel, Policy(fresh: true)));

        // The identity check must run before the native-failure branch, otherwise a
        // stale fallback would be written and the pending edit would proceed against
        // a workbook that moved.
        Assert.Contains("identity is not verified after SaveCopyAs", ex.Message);
        Assert.Contains("SaveCopyAs also failed with: Excel refused SaveCopyAs (1004)", ex.Message);
        Assert.False(Json.GetBool(metadata, "workbookBackupIdentityPreserved"));
        Assert.Equal("none", Json.GetString(metadata, "workbookBackupSource"));
        Assert.False(Json.GetBool(metadata, "workbookBackupFresh"));
        Assert.False(Json.GetBool(metadata, "workbookBackupAvailable"));
        Assert.Null(Json.GetString(metadata, "workbookBackup"));
        Assert.False(File.Exists(Path.Combine(home.Dir, "workbook-backup.xlsx")));

        var after = Json.GetObj(Json.GetObj(metadata, "workbookBackupIdentity"), "after")!;
        Assert.Equal(Path.Combine(home.Dir, "moved.xlsx"), Json.GetString(after, "fullName"));
        // Reported, never repaired.
        Assert.Equal(Path.Combine(home.Dir, "moved.xlsx"), workbook.FullNameValue);
        Assert.Equal(0, workbook.SavedAssignments);
    }

    [Fact]
    public void Changed_identity_refuses_even_when_the_native_copy_produced_no_file()
    {
        using var home = new TestHome();
        var workbook = NewWorkbook(home, saved: false);
        workbook.FullNameAfterSaveCopyAs = Path.Combine(home.Dir, "moved.xlsx");
        workbook.SaveCopyAsWritesNothing = true;
        var metadata = new JsonObject();

        var ex = Assert.Throws<InvalidOperationException>(() =>
            ExcelAdapter.CaptureAuxiliaryWorkbookBackup(
                workbook, workbook.Owner, home.Dir, metadata, AppLabel, Policy(fresh: true)));

        Assert.Contains("identity is not verified after SaveCopyAs", ex.Message);
        Assert.Contains("SaveCopyAs also produced no file", ex.Message);
        Assert.Equal("none", Json.GetString(metadata, "workbookBackupSource"));
        Assert.False(Json.GetBool(metadata, "workbookBackupAvailable"));
        Assert.False(File.Exists(Path.Combine(home.Dir, "workbook-backup.xlsx")));
    }

    // --------------------------------------------------------------- stale paths

    [Fact]
    public void Clean_workbook_keeps_the_file_copy_and_is_labelled_last_saved_file()
    {
        using var home = new TestHome();
        var workbook = NewWorkbook(home, saved: true);
        var metadata = Capture(workbook, home);

        Assert.Equal("workbook-backup.xlsx", Json.GetString(metadata, "workbookBackup"));
        Assert.Equal("last-saved-file", Json.GetString(metadata, "workbookBackupSource"));
        Assert.False(Json.GetBool(metadata, "workbookBackupFresh"));
        Assert.True(Json.GetBool(metadata, "workbookBackupAvailable"));
        Assert.True(Json.GetBool(metadata, "workbookBackupSavedFlag"));
        Assert.Contains("no unsaved changes", Json.GetString(metadata, "workbookBackupReason"));
        Assert.Equal(0, workbook.SaveCopyAsCalls);
        Assert.Equal(0, workbook.Owner.ApplicationReadsSeen);
        Assert.Equal(DiskBytes, File.ReadAllText(Path.Combine(home.Dir, "workbook-backup.xlsx")));
    }

    [Theory]
    [InlineData("book.xlsm", 52)]
    [InlineData("book.xlsb", 50)]
    [InlineData("book.xls", 56)]
    [InlineData("book.xlsx", 61)]
    [InlineData("book.xlsx", null)]
    public void Unvalidated_formats_keep_the_file_copy_and_say_why(string fileName, int? fileFormat)
    {
        using var home = new TestHome();
        var workbook = NewWorkbook(home, saved: false, fileName: fileName, fileFormat: fileFormat);
        var metadata = Capture(workbook, home);

        Assert.Equal("last-saved-file", Json.GetString(metadata, "workbookBackupSource"));
        Assert.False(Json.GetBool(metadata, "workbookBackupFresh"));
        Assert.False(Json.GetBool(metadata, "workbookBackupSavedFlag"));
        Assert.Contains("not a validated ordinary .xlsx", Json.GetString(metadata, "workbookBackupReason"));
        Assert.Equal(0, workbook.SaveCopyAsCalls);
        Assert.Equal(DiskBytes, File.ReadAllText(
            Path.Combine(home.Dir, "workbook-backup" + Path.GetExtension(fileName))));
    }

    [Fact]
    public void Unreadable_saved_flag_refuses_to_claim_a_current_memory_copy()
    {
        using var home = new TestHome();
        var workbook = NewWorkbook(home, saved: false);
        workbook.SavedThrows = true;
        var metadata = Capture(workbook, home);

        Assert.Equal("last-saved-file", Json.GetString(metadata, "workbookBackupSource"));
        Assert.False(Json.GetBool(metadata, "workbookBackupFresh"));
        // Unreadable is reported as null, never coerced into false.
        Assert.True(metadata.ContainsKey("workbookBackupSavedFlag"));
        Assert.Null(metadata["workbookBackupSavedFlag"]);
        Assert.Contains("Saved flag could not be read", Json.GetString(metadata, "workbookBackupReason"));
        Assert.Equal(0, workbook.SaveCopyAsCalls);
    }

    [Fact]
    public void Never_saved_workbook_is_explicitly_unsupported_and_writes_nothing()
    {
        using var home = new TestHome();
        var workbook = new FakeWorkbook { FullNameValue = "", SavedValue = false, FileFormatValue = 51 };
        var metadata = Capture(workbook, home);

        Assert.Equal("none", Json.GetString(metadata, "workbookBackupSource"));
        Assert.False(Json.GetBool(metadata, "workbookBackupFresh"));
        Assert.False(Json.GetBool(metadata, "workbookBackupAvailable"));
        Assert.Null(Json.GetString(metadata, "workbookBackup"));
        Assert.Contains("never been saved", Json.GetString(metadata, "workbookBackupReason"));
        Assert.Empty(Directory.GetFiles(home.Dir));
    }

    [Fact]
    public void Missing_on_disk_file_is_explicitly_unsupported()
    {
        using var home = new TestHome();
        var workbook = new FakeWorkbook
        {
            FullNameValue = Path.Combine(home.Dir, "gone.xlsx"),
            SavedValue = false,
            FileFormatValue = 51,
        };
        var metadata = Capture(workbook, home);

        Assert.False(Json.GetBool(metadata, "workbookBackupAvailable"));
        Assert.Contains("does not point at an existing file", Json.GetString(metadata, "workbookBackupReason"));
        Assert.Equal(0, workbook.SaveCopyAsCalls);
    }

    // ------------------------------------------------------------ failing native

    [Fact]
    public void Failed_native_copy_falls_back_but_is_never_relabelled_fresh()
    {
        using var home = new TestHome();
        var workbook = NewWorkbook(home, saved: false);
        workbook.SaveCopyAsThrows = "Excel refused SaveCopyAs (1004)";
        var metadata = Capture(workbook, home);

        Assert.Equal(1, workbook.SaveCopyAsCalls);
        Assert.Equal("Excel refused SaveCopyAs (1004)", Json.GetString(metadata, "workbookBackupNativeError"));
        Assert.Equal("last-saved-file", Json.GetString(metadata, "workbookBackupSource"));
        Assert.False(Json.GetBool(metadata, "workbookBackupFresh"));
        Assert.True(Json.GetBool(metadata, "workbookBackupAvailable"));
        Assert.Contains("native SaveCopyAs failed", Json.GetString(metadata, "workbookBackupReason"));
        Assert.NotNull(Json.GetObj(metadata, "workbookBackupInteraction"));
        Assert.Equal(DiskBytes, File.ReadAllText(Path.Combine(home.Dir, "workbook-backup.xlsx")));
    }

    [Fact]
    public void Partial_file_from_a_failed_native_copy_is_removed_before_the_fallback()
    {
        using var home = new TestHome();
        var workbook = NewWorkbook(home, saved: false);
        workbook.SaveCopyAsThrows = "disk full";
        workbook.WritePartialBeforeThrowing = true;
        var metadata = Capture(workbook, home);

        Assert.True(Json.GetBool(metadata, "workbookBackupPartialNativeArtifactRemoved"));
        Assert.Equal("last-saved-file", Json.GetString(metadata, "workbookBackupSource"));
        Assert.False(Json.GetBool(metadata, "workbookBackupFresh"));
        Assert.Equal(DiskBytes, File.ReadAllText(Path.Combine(home.Dir, "workbook-backup.xlsx")));
    }

    [Fact]
    public void Silent_native_copy_that_produces_no_file_is_treated_as_a_failure()
    {
        using var home = new TestHome();
        var workbook = NewWorkbook(home, saved: false);
        workbook.SaveCopyAsWritesNothing = true;
        var metadata = Capture(workbook, home);

        Assert.Contains("produced no file", Json.GetString(metadata, "workbookBackupNativeError"));
        Assert.Equal("last-saved-file", Json.GetString(metadata, "workbookBackupSource"));
        Assert.False(Json.GetBool(metadata, "workbookBackupFresh"));
        Assert.Equal(DiskBytes, File.ReadAllText(Path.Combine(home.Dir, "workbook-backup.xlsx")));
    }

    [Fact]
    public void Unreadable_source_file_reports_the_error_and_no_backup()
    {
        using var home = new TestHome();
        var workbook = NewWorkbook(home, saved: true);
        using var exclusive = new FileStream(
            workbook.FullNameValue, FileMode.Open, FileAccess.Read, FileShare.None);
        var metadata = Capture(workbook, home);

        Assert.False(Json.GetBool(metadata, "workbookBackupAvailable"));
        Assert.Equal("none", Json.GetString(metadata, "workbookBackupSource"));
        Assert.False(Json.GetBool(metadata, "workbookBackupFresh"));
        Assert.NotNull(Json.GetString(metadata, "workbookBackupError"));
        Assert.Null(Json.GetString(metadata, "workbookBackup"));
    }

    // ------------------------------------------------------------ no overwriting

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void Preexisting_target_is_never_overwritten(bool dirty)
    {
        using var home = new TestHome();
        var workbook = NewWorkbook(home, saved: !dirty);
        var target = Path.Combine(home.Dir, "workbook-backup.xlsx");
        File.WriteAllText(target, "DO-NOT-CLOBBER");
        var metadata = Capture(workbook, home);

        Assert.Equal("DO-NOT-CLOBBER", File.ReadAllText(target));
        Assert.False(Json.GetBool(metadata, "workbookBackupAvailable"));
        Assert.False(Json.GetBool(metadata, "workbookBackupFresh"));
        Assert.Null(Json.GetString(metadata, "workbookBackup"));
        Assert.Contains("refusing to overwrite", Json.GetString(metadata, "workbookBackupReason"));
        Assert.Equal(0, workbook.SaveCopyAsCalls);
    }

    // -------------------------------------------------------------------- fakes

    /// <summary>
    /// Defaults to the opt-in policy so the pre-existing cases keep exercising the
    /// native branch they were written for. The default (opt-out) path has its own
    /// dedicated cases above.
    /// </summary>
    private static JsonObject Capture(FakeWorkbook workbook, TestHome home, bool fresh = true)
    {
        var metadata = new JsonObject();
        ExcelAdapter.CaptureAuxiliaryWorkbookBackup(
            workbook, workbook.Owner, home.Dir, metadata, AppLabel, Policy(fresh));
        return metadata;
    }

    private static ExcelAdapter.ExcelWorkbookBackupPolicy Policy(bool fresh) => new(fresh);

    private static FakeWorkbook NewWorkbook(
        TestHome home, bool saved, string fileName = "book.xlsx", int? fileFormat = 51)
    {
        var path = Path.Combine(home.Dir, fileName);
        File.WriteAllText(path, DiskBytes);
        return new FakeWorkbook
        {
            FullNameValue = path,
            SavedValue = saved,
            FileFormatValue = fileFormat,
            MemoryBytes = MemoryBytes,
        };
    }

    public sealed class FakeWorkbook
    {
        public string FullNameValue { get; set; } = "";

        public bool SavedValue { get; set; }

        public int? FileFormatValue { get; set; } = 51;

        public string MemoryBytes { get; set; } = "";

        public bool SavedThrows { get; set; }

        public bool SavedThrowsAfterSaveCopyAs { get; set; }

        public bool ApplicationThrows { get; set; }

        public string? SaveCopyAsThrows { get; set; }

        public bool SaveCopyAsWritesNothing { get; set; }

        public bool WritePartialBeforeThrowing { get; set; }

        public bool FlipSavedDuringSaveCopyAs { get; set; }

        public int? OwnerWorkbookCountAfterCopy { get; set; }

        public string? FullNameAfterSaveCopyAs { get; set; }

        public int SaveCopyAsCalls { get; private set; }

        public int SaveCalls { get; private set; }

        public int CloseCalls { get; private set; }

        public int ActivateCalls { get; private set; }

        public int SavedAssignments { get; private set; }

        /// <summary>The instance that actually owns this workbook.</summary>
        public FakeApplication Owner { get; } = new();

        public FakeApplication Application =>
            ApplicationThrows
                ? throw new InvalidOperationException("Application is unavailable")
                : Owner.Borrow();

        public string FullName => FullNameValue;

        public bool Saved
        {
            get => SavedThrows || (SavedThrowsAfterSaveCopyAs && SaveCopyAsCalls > 0)
                ? throw new InvalidOperationException("Saved is unavailable")
                : SavedValue;
            set
            {
                SavedAssignments++;
                SavedValue = value;
            }
        }

        public int FileFormatReads { get; private set; }

        public int FileFormat
        {
            get
            {
                FileFormatReads++;
                return FileFormatValue ?? throw new InvalidOperationException("FileFormat is unavailable");
            }
        }

        public void SaveCopyAs(string fileName)
        {
            SaveCopyAsCalls++;
            // Identity mutations land before the failure branch on purpose: a native
            // copy that moves the workbook and *then* throws is exactly the case the
            // adapter must not paper over with a stale fallback.
            if (FlipSavedDuringSaveCopyAs) SavedValue = true;
            if (FullNameAfterSaveCopyAs is { } movedTo) FullNameValue = movedTo;
            if (OwnerWorkbookCountAfterCopy is int count) Owner.Workbooks.Count = count;

            if (SaveCopyAsThrows is { } message)
            {
                if (WritePartialBeforeThrowing) File.WriteAllText(fileName, "PARTIAL");
                throw new InvalidOperationException(message);
            }

            if (SaveCopyAsWritesNothing) return;
            File.WriteAllText(fileName, MemoryBytes);
        }

        public void Save() => SaveCalls++;

        public void Close(bool saveChanges) => CloseCalls++;

        public void Activate() => ActivateCalls++;
    }

    public sealed class FakeApplication
    {
        private bool _displayAlerts = true;

        public FakeWorkbooks Workbooks =>
            WorkbooksThrows
                ? throw new InvalidOperationException("Workbooks is unavailable")
                : WorkbooksValue;

        public FakeWorkbooks WorkbooksValue { get; } = new();

        public bool WorkbooksThrows { get; set; }

        /// <summary>How many times workbook.Application handed this instance out.</summary>
        public int ApplicationReads { get; private set; }

        public int ApplicationReadsSeen => ApplicationReads;

        public int DisplayAlertsAssignments { get; private set; }

        public long Hwnd => 0;

        public bool DisplayAlerts
        {
            get => _displayAlerts;
            set
            {
                DisplayAlertsAssignments++;
                _displayAlerts = value;
            }
        }

        public FakeApplication Borrow()
        {
            ApplicationReads++;
            return this;
        }
    }

    public sealed class FakeWorkbooks
    {
        public int Count { get; set; } = 1;
    }
}
