using System.Globalization;
using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class ExcelAdapter
{
    // xlOpenXMLWorkbook. https://learn.microsoft.com/en-us/office/vba/api/excel.xlfileformat
    internal const int XlOpenXmlWorkbookFileFormat = 51;

    internal const string WorkbookBackupStem = "workbook-backup";
    internal const string WorkbookBackupSourceCurrentMemory = "current-memory-savecopyas";
    internal const string WorkbookBackupSourceLastSavedFile = "last-saved-file";
    internal const string WorkbookBackupSourceNone = "none";

    internal const string FreshWorkbookBackupVariable = "DOCBRIDGE_EXCEL_FRESH_WORKBOOK_BACKUP";

    /// <summary>
    /// Whether the auxiliary workbook copy may be taken from the in-memory workbook
    /// with <c>SaveCopyAs</c>.
    ///
    /// Default is <c>false</c>. A native copy serializes the whole workbook regardless
    /// of how small the edit is, so it is not paid on the default path; the auxiliary
    /// copy keeps the historical last-saved file copy and says so in metadata. The
    /// operation-scoped rollback in state.json is a separate mechanism and is not
    /// affected either way.
    ///
    /// The value is read from the environment once per capture. Passing a policy
    /// explicitly keeps tests deterministic without mutating process state.
    /// </summary>
    internal readonly record struct ExcelWorkbookBackupPolicy(bool FreshCopyEnabled)
    {
        /// <summary>Exactly "1" enables. Anything else, including "true"/"yes"/"01"/whitespace, does not.</summary>
        internal static ExcelWorkbookBackupPolicy FromSetting(string? raw) =>
            new(string.Equals(raw, "1", StringComparison.Ordinal));

        internal static ExcelWorkbookBackupPolicy FromEnvironment() =>
            FromSetting(Environment.GetEnvironmentVariable(FreshWorkbookBackupVariable));
    }

    /// <summary>
    /// Writes the auxiliary workbook copy that accompanies a snapshot, and labels it
    /// honestly.
    ///
    /// This copy has no automated restore consumer; it is operator-facing evidence
    /// beside the operation-scoped state.json, and this method does not change the
    /// operation-scoped rollback semantics.
    ///
    /// **By default the copy is the last-saved file on disk**, exactly as before, and
    /// the metadata says so (<c>workbookBackupSource: last-saved-file</c>,
    /// <c>workbookBackupFresh: false</c>, with the workbook's <c>Saved</c> flag
    /// alongside). For a workbook with unsaved changes that copy does not contain them.
    /// That is a deliberate trade: a native copy serializes the entire workbook no
    /// matter how small the edit, so it is not put on the default path.
    ///
    /// Setting <c>DOCBRIDGE_EXCEL_FRESH_WORKBOOK_BACKUP=1</c> opts in. Only then, and
    /// only for an ordinary on-disk .xlsx (extension .xlsx and FileFormat 51) whose
    /// <c>Saved</c> flag is readable and false, is the copy produced by
    /// <c>Workbook.SaveCopyAs</c>, which per the Excel VBA reference "saves a copy of
    /// the workbook to a file but doesn't modify the open workbook in memory". Every
    /// other case keeps the file copy and is explicitly labelled
    /// <c>last-saved-file</c>. Format support is deliberately not expanded.
    ///
    /// <paramref name="attachedApplication"/> is the instance the adapter attached to,
    /// which is NOT necessarily the instance that owns this workbook:
    /// ResolveTargetWorkbook can select a workbook in another Excel process. The open
    /// workbook count and the foreground guard therefore come from
    /// <c>workbook.Application</c>, borrowed and released here, never final-released.
    /// The attached instance is recorded for diagnostics only.
    ///
    /// The workbook is never saved, closed, activated or had its <c>Saved</c> flag or
    /// path reset, and <c>DisplayAlerts</c> is never touched.
    /// </summary>
    internal static void CaptureAuxiliaryWorkbookBackup(
        object workbook,
        object? attachedApplication,
        string snapshotDir,
        JsonObject metadata,
        string appLabel,
        ExcelWorkbookBackupPolicy policy)
    {
        metadata["workbookBackupFreshCopyEnabled"] = policy.FreshCopyEnabled;
        metadata["workbookBackupAttachedApplicationProvided"] = attachedApplication is not null;
        // Count and foreground always come from the workbook's own owner, never from
        // the attached instance, so a cross-instance target cannot be misreported.
        metadata["workbookBackupAttachedApplicationUsedForCount"] = false;

        var savedFlag = TryReadWorkbookSavedFlag(workbook);
        metadata["workbookBackupSavedFlag"] = savedFlag is bool flag ? JsonValue.Create(flag) : null;

        var fullName = TryReadBackupFullName(workbook);
        if (string.IsNullOrWhiteSpace(fullName))
        {
            MarkBackupUnavailable(
                metadata,
                "workbook has never been saved to disk, so there is no last-saved file to copy and no " +
                "established file format for a native copy; auxiliary backup is not supported for it");
            return;
        }

        if (!File.Exists(fullName))
        {
            MarkBackupUnavailable(
                metadata,
                $"workbook FullName '{fullName}' does not point at an existing file; auxiliary backup is not supported for it");
            return;
        }

        var extension = Path.GetExtension(fullName);
        var destination = Path.Combine(snapshotDir, WorkbookBackupStem + extension);
        if (File.Exists(destination))
        {
            MarkBackupUnavailable(
                metadata,
                $"auxiliary backup target '{Path.GetFileName(destination)}' already exists; refusing to overwrite it");
            return;
        }

        if (!policy.FreshCopyEnabled)
        {
            // Default path. No FileFormat read and no SaveCopyAs: the whole point is
            // that a small edit does not pay a whole-workbook serialization. The copy
            // is the last-saved file and the metadata says so, including the Saved
            // flag, so a reader can tell whether unsaved changes are missing from it.
            // workbookBackupFileFormat is intentionally absent here; absent means
            // "not evaluated", which is not the same as the null that means
            // "read and unreadable".
            CopyLastSavedFile(fullName, destination, metadata,
                $"the fresh in-memory copy is opt-in and {FreshWorkbookBackupVariable} is not set to 1, " +
                "so the last-saved file on disk is copied; if the workbook has unsaved changes this copy " +
                "does not contain them. The operation-scoped rollback in state.json is unaffected");
            return;
        }

        var fileFormat = TryReadWorkbookFileFormat(workbook);
        metadata["workbookBackupFileFormat"] = fileFormat is int format ? JsonValue.Create(format) : null;

        if (savedFlag is null)
        {
            CopyLastSavedFile(fullName, destination, metadata,
                "workbook Saved flag could not be read, so a current-memory copy cannot be justified");
            return;
        }

        if (savedFlag == true)
        {
            CopyLastSavedFile(fullName, destination, metadata,
                "workbook reports no unsaved changes, so the last-saved file is used");
            return;
        }

        if (!string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase)
            || fileFormat != XlOpenXmlWorkbookFileFormat)
        {
            CopyLastSavedFile(fullName, destination, metadata,
                $"workbook has unsaved changes but is not a validated ordinary .xlsx " +
                $"(extension '{extension}', FileFormat {FormatFileFormat(fileFormat)}); " +
                "native copy support is deliberately not expanded to other formats");
            return;
        }

        CaptureNativeWorkbookCopy(workbook, fullName, destination, metadata, appLabel);
    }

    private static void CaptureNativeWorkbookCopy(
        object workbook, string fullName, string destination, JsonObject metadata, string appLabel)
    {
        object? owningApplication = null;
        try
        {
            owningApplication = TryBorrowWorkbookApplication(workbook);
            metadata["workbookBackupApplicationSource"] =
                owningApplication is null ? "unreadable" : "workbook.Application";

            var identityBefore = ReadWorkbookBackupIdentity(workbook, owningApplication);
            if (!IdentityFullyReadable(identityBefore))
            {
                // Unknown identity before the copy means a preserved-identity claim
                // afterwards would be unfalsifiable. Take the conservative labelled
                // fallback instead of a native copy nobody can verify.
                metadata["workbookBackupIdentity"] = new JsonObject { ["before"] = identityBefore.DeepClone() };
                metadata["workbookBackupIdentityPreserved"] = false;
                metadata["workbookBackupIdentityVerifiable"] = false;
                CopyLastSavedFile(fullName, destination, metadata,
                    "workbook identity (FullName, Saved, and the owning instance's open workbook count) " +
                    "could not be fully read before the copy, so a preserved-identity native copy cannot " +
                    "be claimed and the last-saved file is used instead");
                return;
            }

            metadata["workbookBackupIdentityVerifiable"] = true;

            Exception? nativeFailure = null;
            var interaction = new ForegroundInteractionGuard(appLabel);
            try
            {
                TrackWorkbookForeground(interaction, owningApplication);
                try { ((dynamic)workbook).SaveCopyAs(destination); }
                catch (Exception ex) { nativeFailure = ex; }
            }
            finally { metadata["workbookBackupInteraction"] = interaction.Complete(); }

            // Identity is checked immediately after the attempt, before any outcome
            // branch. A SaveCopyAs that moved the workbook and *then* threw, or that
            // moved it and produced no file, still leaves the workbook unverified;
            // returning a stale fallback there would let the pending edit proceed
            // against it. Only a proven-preserved identity may fall back.
            var identityAfter = ReadWorkbookBackupIdentity(workbook, owningApplication);
            metadata["workbookBackupIdentity"] = new JsonObject
            {
                ["before"] = identityBefore.DeepClone(),
                ["after"] = identityAfter.DeepClone(),
            };

            // Every field must be readable on both sides. Two unreadable fields are
            // both null and would otherwise compare equal, which is not evidence that
            // anything was preserved.
            var preserved = IdentityFullyReadable(identityAfter)
                && string.Equals(
                    identityBefore.ToJsonString(), identityAfter.ToJsonString(), StringComparison.Ordinal);
            metadata["workbookBackupIdentityPreserved"] = preserved;

            if (!preserved)
            {
                if (nativeFailure is not null)
                    metadata["workbookBackupNativeError"] = nativeFailure.Message;

                // Do not reset Saved or the path. Refuse the snapshot instead, so the
                // pending edit cannot proceed silently against a workbook whose
                // identity we can no longer vouch for.
                var failure =
                    "auxiliary backup: the target workbook's identity is not verified after SaveCopyAs " +
                    $"(before {identityBefore.ToJsonString()}, after {identityAfter.ToJsonString()}); " +
                    "the snapshot is refused so the pending edit does not proceed against an unverified workbook";
                if (nativeFailure is not null)
                    failure += $"; SaveCopyAs also failed with: {nativeFailure.Message}";
                else if (!File.Exists(destination))
                    failure += "; SaveCopyAs also produced no file";
                metadata["workbookBackupError"] = failure;
                metadata["workbookBackupSource"] = WorkbookBackupSourceNone;
                metadata["workbookBackupFresh"] = false;
                metadata["workbookBackupAvailable"] = File.Exists(destination);
                throw new InvalidOperationException(failure);
            }

            if (nativeFailure is not null)
            {
                metadata["workbookBackupNativeError"] = nativeFailure.Message;
                RemovePartialNativeArtifact(destination, metadata);
                CopyLastSavedFile(fullName, destination, metadata,
                    "native SaveCopyAs failed, so this copy is the last-saved file and not current memory");
                return;
            }

            if (!File.Exists(destination))
            {
                metadata["workbookBackupNativeError"] =
                    "SaveCopyAs returned without an error but produced no file";
                CopyLastSavedFile(fullName, destination, metadata,
                    "native SaveCopyAs produced no file, so this copy is the last-saved file and not current memory");
                return;
            }

            metadata["workbookBackup"] = Path.GetFileName(destination);
            metadata["workbookBackupSource"] = WorkbookBackupSourceCurrentMemory;
            metadata["workbookBackupFresh"] = true;
            metadata["workbookBackupAvailable"] = true;
            metadata["workbookBackupReason"] =
                "workbook has unsaved changes and is an ordinary on-disk .xlsx (FileFormat 51), " +
                "so the copy was taken from the in-memory workbook with SaveCopyAs";
        }
        finally
        {
            // Borrowed from workbook.Application: release exactly the one reference
            // taken here. Never FinalReleaseComObject, which would tear down an RCW
            // the caller and the rest of the adapter still hold.
            RotHelper.ReleaseComReference(owningApplication);
        }
    }

    private static void CopyLastSavedFile(
        string fullName, string destination, JsonObject metadata, string reason)
    {
        if (File.Exists(destination))
        {
            MarkBackupUnavailable(
                metadata,
                $"{reason}; the target '{Path.GetFileName(destination)}' already exists and will not be overwritten");
            return;
        }

        try
        {
            // Excel holds the file open, so the copy is taken with shared read/write.
            using var source = new FileStream(fullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var target = new FileStream(destination, FileMode.CreateNew, FileAccess.Write);
            source.CopyTo(target);
        }
        catch (Exception ex)
        {
            metadata["workbookBackupError"] = ex.Message;
            metadata["workbookBackupSource"] = WorkbookBackupSourceNone;
            metadata["workbookBackupFresh"] = false;
            metadata["workbookBackupAvailable"] = false;
            metadata["workbookBackupReason"] = reason;
            return;
        }

        metadata["workbookBackup"] = Path.GetFileName(destination);
        metadata["workbookBackupSource"] = WorkbookBackupSourceLastSavedFile;
        metadata["workbookBackupFresh"] = false;
        metadata["workbookBackupAvailable"] = true;
        metadata["workbookBackupReason"] = reason;
    }

    private static void MarkBackupUnavailable(JsonObject metadata, string reason)
    {
        metadata["workbookBackupSource"] = WorkbookBackupSourceNone;
        metadata["workbookBackupFresh"] = false;
        metadata["workbookBackupAvailable"] = false;
        metadata["workbookBackupReason"] = reason;
    }

    /// <summary>
    /// Removes a partial file left behind by a failed SaveCopyAs. Only a file this
    /// method has just seen created is removed; a target that already existed before
    /// the attempt is rejected earlier and never reaches here.
    /// </summary>
    private static void RemovePartialNativeArtifact(string destination, JsonObject metadata)
    {
        if (!File.Exists(destination)) return;
        try
        {
            File.Delete(destination);
            metadata["workbookBackupPartialNativeArtifactRemoved"] = true;
        }
        catch (Exception ex)
        {
            metadata["workbookBackupPartialNativeArtifactRemoved"] = false;
            metadata["workbookBackupPartialNativeArtifactError"] = ex.Message;
        }
    }

    private static void TrackWorkbookForeground(ForegroundInteractionGuard interaction, object? application)
    {
        if (application is null) return;
        try
        {
            interaction.TrackTargetWindow(
                Convert.ToInt64(((dynamic)application).Hwnd, CultureInfo.InvariantCulture));
        }
        catch { }
    }

    private static JsonObject ReadWorkbookBackupIdentity(object workbook, object? owningApplication)
    {
        var saved = TryReadWorkbookSavedFlag(workbook);
        var count = TryReadOpenWorkbookCount(owningApplication);
        return new JsonObject
        {
            ["fullName"] = TryReadBackupFullName(workbook),
            ["saved"] = saved is bool flag ? JsonValue.Create(flag) : null,
            ["workbookCount"] = count is int value ? JsonValue.Create(value) : null,
        };
    }

    private static bool IdentityFullyReadable(JsonObject identity) =>
        identity["fullName"] is JsonValue name
        && name.TryGetValue<string>(out var fullName)
        && !string.IsNullOrEmpty(fullName)
        && identity["saved"] is JsonValue saved
        && saved.TryGetValue<bool>(out _)
        && identity["workbookCount"] is JsonValue count
        && count.TryGetValue<int>(out _);

    private static object? TryBorrowWorkbookApplication(object workbook)
    {
        try { return (object?)((dynamic)workbook).Application; }
        catch { return null; }
    }

    private static string? TryReadBackupFullName(object workbook)
    {
        try { return Convert.ToString(((dynamic)workbook).FullName, CultureInfo.InvariantCulture); }
        catch { return null; }
    }

    private static bool? TryReadWorkbookSavedFlag(object workbook)
    {
        try
        {
            object? raw = ((dynamic)workbook).Saved;
            if (raw is null or DBNull) return null;
            return Convert.ToBoolean(raw, CultureInfo.InvariantCulture);
        }
        catch { return null; }
    }

    private static int? TryReadWorkbookFileFormat(object workbook)
    {
        try
        {
            object? raw = ((dynamic)workbook).FileFormat;
            if (raw is null or DBNull) return null;
            return Convert.ToInt32(raw, CultureInfo.InvariantCulture);
        }
        catch { return null; }
    }

    private static int? TryReadOpenWorkbookCount(object? application)
    {
        if (application is null) return null;
        object? workbooks = null;
        try
        {
            // Chaining application.Workbooks.Count leaks the borrowed collection RCW.
            workbooks = (object?)((dynamic)application).Workbooks;
            if (workbooks is null) return null;
            object? raw = ((dynamic)workbooks).Count;
            if (raw is null or DBNull) return null;
            return Convert.ToInt32(raw, CultureInfo.InvariantCulture);
        }
        catch { return null; }
        finally { RotHelper.ReleaseComReference(workbooks); }
    }

    private static string FormatFileFormat(int? fileFormat) =>
        fileFormat is int value ? value.ToString(CultureInfo.InvariantCulture) : "unreadable";
}
