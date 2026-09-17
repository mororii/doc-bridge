using System.Globalization;
using System.IO.Compression;
using System.Reflection;
using System.Security.Cryptography;
using System.Text.Json.Nodes;
using System.Xml;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class ExcelAdapter
{
    internal const string DeferredFormatSnapshotVariable = "DOCBRIDGE_EXCEL_DEFERRED_FORMAT_SNAPSHOT";
    internal const string DeferredFormatRestoreMode = "deferred-format-only";
    internal const string DeferredFormatPayloadMode = "deferred-format-only";
    internal const string DeferredFormatCompatibilityRestoreMode = "copy-sheet-topology";
    internal const int DeferredFormatSnapshotVersion = 4;
    internal const int DeferredFormatMinCells = 1000;
    internal const int DeferredFormatMaxCells = 5000;
    internal const long DeferredFormatMaxFileBytes = 16L * 1024 * 1024;
    internal const long DeferredFormatMaxUncompressedXmlBytes = 32L * 1024 * 1024;
    internal const long DeferredFormatMaxXmlPartBytes = 16L * 1024 * 1024;
    internal const int DeferredFormatMaxPackageParts = 48;

    private const int MsoAutomationSecurityForceDisable = 3;
    private const string DeferredFormatCheckpointPrefix = "deferred-format-";
    private const string SpreadsheetMlTransitionalNs = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private const string SpreadsheetMlStrictNs = "http://purl.oclc.org/ooxml/spreadsheetml/main";
    private const string PackageContentTypesNs = "http://schemas.openxmlformats.org/package/2006/content-types";
    private const string PackageRelsNs = "http://schemas.openxmlformats.org/package/2006/relationships";
    private const string XlsxMainContentType =
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml";

    private static readonly byte[] ZipMagic = [0x50, 0x4B, 0x03, 0x04];
    private static readonly byte[] Ole2Magic = [0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1];
    private static readonly HashSet<string> DeferredWrittenStyleKeys = new(StringComparer.Ordinal)
    {
        ExcelStyleContract.Bold,
        ExcelStyleContract.Italic,
        ExcelStyleContract.FontSize,
        ExcelStyleContract.NumberFormat,
        ExcelStyleContract.FillColor,
    };

    internal static bool IsDeferredWrittenStyleKey(string key) => DeferredWrittenStyleKeys.Contains(key);
    private static readonly HashSet<string> AllowedRelationshipLocals = new(StringComparer.OrdinalIgnoreCase)
    {
        "officeDocument",
        "worksheet",
        "styles",
        "sharedStrings",
        "theme",
        "extended-properties",
        "core-properties",
        "custom-properties",
        "printerSettings",
    };

    internal readonly record struct ExcelDeferredFormatSnapshotPolicy(bool Enabled)
    {
        internal static ExcelDeferredFormatSnapshotPolicy FromSetting(string? raw) =>
            new(string.Equals(raw, "1", StringComparison.Ordinal));

        internal static ExcelDeferredFormatSnapshotPolicy FromEnvironment() =>
            FromSetting(Environment.GetEnvironmentVariable(DeferredFormatSnapshotVariable));
    }

    internal static bool IsDeferredFormatEnvelope(JsonObject? state) =>
        state is not null
        && string.Equals(Json.GetString(state, "restoreMode"), DeferredFormatCompatibilityRestoreMode, StringComparison.Ordinal)
        && Json.GetInt(state, "snapshotVersion") == DeferredFormatSnapshotVersion
        && string.Equals(Json.GetString(state, "payloadMode"), DeferredFormatPayloadMode, StringComparison.Ordinal);

    internal static bool LooksLikeIncompleteDeferredFormatEnvelope(JsonObject? state) =>
        state is not null
        && !IsDeferredFormatEnvelope(state)
        && (Json.GetInt(state, "snapshotVersion") == DeferredFormatSnapshotVersion
            || string.Equals(Json.GetString(state, "payloadMode"), DeferredFormatPayloadMode, StringComparison.Ordinal));

    internal static JsonObject EvaluateDeferredFormatRequestEligibility(
        ExcelDeferredFormatSnapshotPolicy policy,
        JsonObject metadata,
        IReadOnlyList<JsonObject>? ops)
    {
        var result = EligibilityBase();
        if (!policy.Enabled)
            return RejectEligibility(result, "policy",
                $"{DeferredFormatSnapshotVariable} is not set to 1");

        if (!string.Equals(
                Json.GetString(metadata, HostSnapshotCaptureContext.MetadataKey),
                HostSnapshotCaptureContext.Execute,
                StringComparison.Ordinal))
        {
            return RejectEligibility(result, "host-context",
                "deferred format checkpoint requires hostSnapshotContext=execute");
        }

        if (!IsFormatOnlySnapshot(ops) || ops!.Count != 1)
            return RejectEligibility(result, "ops", "deferred format checkpoint requires a single format_range op");

        var op = ops[0];
        var sheet = Json.GetString(Json.GetObj(op, "target"), "sheet");
        var rangeText = Json.GetString(op, "range");
        result["sheet"] = sheet;
        result["range"] = rangeText;
        if (string.IsNullOrWhiteSpace(sheet) || string.IsNullOrWhiteSpace(rangeText)
            || !TryParseSingleRectangularA1(rangeText, sheet, out var address, out var cells))
        {
            return RejectEligibility(result, "sheet-range",
                "deferred format checkpoint requires target.sheet and one rectangular A1 range");
        }

        result["range"] = address;
        result["targetCells"] = cells;
        var errors = new List<string>();
        if (!ExcelStyleContract.TryNormalize(Json.GetObj(op, "style"), out var canonical, errors)
            || canonical.Count == 0)
        {
            return RejectEligibility(result, "style",
                errors.Count > 0 ? string.Join("; ", errors) : "format style is empty or invalid");
        }

        var written = new HashSet<string>(StringComparer.Ordinal);
        foreach (var (key, _) in canonical)
            written.Add(key);
        var hasFill = written.Contains(ExcelStyleContract.FillColor);
        result["hasFillColor"] = hasFill;
        if (!hasFill || written.Any(key => !DeferredWrittenStyleKeys.Contains(key)))
        {
            return RejectEligibility(result, "style",
                "deferred format checkpoint requires fillColor and only bold/italic/fontSize/numberFormat extras");
        }

        if (cells < DeferredFormatMinCells || cells > DeferredFormatMaxCells)
        {
            return RejectEligibility(result, "cell-count",
                $"target cells {cells} are outside {DeferredFormatMinCells}..{DeferredFormatMaxCells}");
        }

        result["eligible"] = true;
        result["code"] = "ok";
        result["reason"] = "request is inside the restricted deferred format route";
        return result;
    }

    internal static JsonObject EvaluateDeferredFormatWorkbookEligibility(object workbook, string? targetSheet)
    {
        var result = EligibilityBase();
        var fullName = TryReadBackupFullName(workbook);
        result["documentRef"] = fullName;
        if (!IsLocalExistingXlsx(fullName, out var pathReason))
            return RejectEligibility(result, "path", pathReason);

        long? bytes = null;
        try { bytes = new FileInfo(fullName!).Length; } catch { }
        result["workbookBytes"] = bytes is long length ? JsonValue.Create(length) : null;
        if (bytes is null)
            return RejectEligibility(result, "file-size", "workbook file size could not be read");
        if (bytes.Value > DeferredFormatMaxFileBytes)
        {
            return RejectEligibility(result, "file-size",
                $"workbook is {bytes.Value} bytes, above the {DeferredFormatMaxFileBytes} rollout cap");
        }

        var fileFormat = TryReadWorkbookFileFormat(workbook);
        result["fileFormat"] = fileFormat is int format ? JsonValue.Create(format) : null;
        if (fileFormat != XlOpenXmlWorkbookFileFormat)
            return RejectEligibility(result, "file-format", "workbook FileFormat is not ordinary xlsx 51");

        if (TryReadComBool(workbook, "HasVBProject") != false)
            return RejectEligibility(result, "macros", "HasVBProject is true or unreadable");
        if (TryReadComBool(workbook, "ProtectStructure") != false
            || TryReadComBool(workbook, "ProtectWindows") != false
            || TargetSheetProtected(workbook, targetSheet) != false)
        {
            return RejectEligibility(result, "protection",
                "workbook or target sheet protection is true or unreadable");
        }

        if (TryReadComBool(workbook, "HasPassword") != false
            || TryReadComBool(workbook, "WriteReserved") != false)
        {
            return RejectEligibility(result, "password", "password or write-reserved is true or unreadable");
        }

        if (TryWorkbookHasExcelLinks(workbook) != false)
            return RejectEligibility(result, "links", "Excel links are present or unreadable");

        object? owning = null;
        try
        {
            owning = TryBorrowWorkbookApplication(workbook);
            if (!DeferredIdentityFullyReadable(ReadDeferredOwnerIdentity(workbook, owning)))
            {
                return RejectEligibility(result, "identity",
                    "owner FullName, Saved, workbookCount, and hwnd must all be readable before SaveCopyAs");
            }
        }
        finally { RotHelper.ReleaseComReference(owning); }

        result["eligible"] = true;
        result["code"] = "ok";
        result["reason"] = "workbook is inside the restricted deferred format route";
        return result;
    }

    internal static JsonObject PrefightDeferredFormatCheckpoint(string path)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            return Prefight(false, "input-missing", "checkpoint file is missing");

        long size;
        try { size = new FileInfo(path).Length; }
        catch (Exception ex) { return Prefight(false, "input-missing", ex.Message); }
        if (size > DeferredFormatMaxFileBytes)
            return Prefight(false, "file-size", $"checkpoint is {size} bytes, above the 16 MiB cap");

        try
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
            return PrefightDeferredFormatCheckpoint(stream);
        }
        catch (InvalidDataException ex)
        {
            return Prefight(false, "bad-zip", ex.Message);
        }
        catch (Exception ex)
        {
            return Prefight(false, "input-missing", ex.Message);
        }
    }

    internal static JsonObject PrefightDeferredFormatCheckpoint(Stream stream)
    {
        if (stream is null)
            return Prefight(false, "input-missing", "checkpoint stream is missing");

        long size;
        try
        {
            if (!stream.CanSeek)
                return Prefight(false, "input-missing", "checkpoint stream must be seekable");
            stream.Position = 0;
            size = stream.Length;
        }
        catch (Exception ex)
        {
            return Prefight(false, "input-missing", ex.Message);
        }

        if (size > DeferredFormatMaxFileBytes)
            return Prefight(false, "file-size", $"checkpoint is {size} bytes, above the 16 MiB cap");

        byte[] head;
        try
        {
            head = new byte[8];
            var read = stream.Read(head, 0, head.Length);
            if (read < 4)
                return Prefight(false, "unknown-container", "checkpoint is too small to classify");
        }
        catch (Exception ex)
        {
            return Prefight(false, "input-missing", ex.Message);
        }

        if (head.AsSpan(0, Ole2Magic.Length).SequenceEqual(Ole2Magic))
        {
            return StreamContainsUtf16Marker(stream, "EncryptedPackage")
                ? Prefight(false, "encrypted-ooxml", "checkpoint is an encrypted OLE2 package")
                : Prefight(false, "ole2-compound-file", "checkpoint is an OLE2 compound file");
        }

        if (!head.AsSpan(0, ZipMagic.Length).SequenceEqual(ZipMagic))
            return Prefight(false, "unknown-container", "checkpoint is not a ZIP Open XML package");

        try
        {
            stream.Position = 0;
            return InspectDeferredFormatPackage(stream);
        }
        catch (InvalidDataException ex)
        {
            return Prefight(false, "bad-zip", ex.Message);
        }
        catch (XmlException ex) when (IsDtdProhibited(ex))
        {
            return Prefight(false, "dtd-prohibited", ex.Message);
        }
        catch (XmlException ex)
        {
            return Prefight(false, "malformed-xml", ex.Message);
        }
        catch (Exception ex)
        {
            return Prefight(false, "bad-zip", ex.Message);
        }
    }

    internal static bool TryResolveDeferredFormatCheckpointPath(
        string snapshotDir,
        string? recordedFileName,
        out string fullPath,
        out string error)
    {
        fullPath = "";
        error = "";
        if (string.IsNullOrWhiteSpace(snapshotDir) || !Directory.Exists(snapshotDir))
        {
            error = "snapshot directory is missing";
            return false;
        }

        var dirName = Path.GetFileName(snapshotDir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
        if (!SnapshotService.IsSafeSnapshotId(dirName))
        {
            error = "snapshot directory name is not a safe snapshot id";
            return false;
        }

        var expected = DeferredFormatCheckpointPrefix + dirName + ".xlsx";
        if (!string.IsNullOrWhiteSpace(recordedFileName)
            && !string.Equals(recordedFileName, expected, StringComparison.Ordinal))
        {
            error = $"recorded checkpoint file '{recordedFileName}' does not match reconstructed '{expected}'";
            return false;
        }

        var snapshotFull = Path.GetFullPath(snapshotDir);
        fullPath = Path.GetFullPath(Path.Combine(snapshotDir, expected));
        if (!SnapshotService.IsDirectoryUnder(snapshotFull, fullPath)
            && !string.Equals(
                Path.GetDirectoryName(fullPath),
                snapshotFull.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar),
                StringComparison.OrdinalIgnoreCase))
        {
            error = "reconstructed checkpoint path is outside the snapshot directory";
            fullPath = "";
            return false;
        }

        return true;
    }

    internal static JsonObject ValidateDeferredFormatRestoreBinding(
        JsonObject metadata,
        JsonObject state,
        string stateJsonPath) =>
        ValidateDeferredFormatRestoreBinding(metadata, state, stateJsonPath, out _);

    internal static JsonObject ValidateDeferredFormatRestoreBinding(
        JsonObject metadata,
        JsonObject? expectedState,
        string stateJsonPath,
        out JsonObject? boundState)
    {
        boundState = null;
        if (!string.Equals(Json.GetString(metadata, "restoreMode"), DeferredFormatRestoreMode, StringComparison.Ordinal)
            || !string.Equals(Json.GetString(metadata, "snapshotKind"), DeferredFormatRestoreMode, StringComparison.Ordinal)
            || !string.Equals(Json.GetString(metadata, "payloadMode"), DeferredFormatPayloadMode, StringComparison.Ordinal)
            || string.IsNullOrWhiteSpace(Json.GetString(metadata, "documentRef"))
            || string.IsNullOrWhiteSpace(Json.GetString(metadata, "deferredCheckpoint"))
            || string.IsNullOrWhiteSpace(Json.GetString(metadata, "deferredCheckpointSha256"))
            || string.IsNullOrWhiteSpace(Json.GetString(metadata, "deferredStateSha256")))
        {
            return Prefight(false, "metadata-missing",
                "deferred restore requires bound metadata restoreMode, snapshotKind, payloadMode, documentRef, checkpoint, and hashes");
        }

        if (!File.Exists(stateJsonPath))
            return Prefight(false, "state-hash", "state.json is missing");

        byte[] bytes;
        try
        {
            using var stream = new FileStream(stateJsonPath, FileMode.Open, FileAccess.Read, FileShare.Read);
            using var copy = new MemoryStream();
            stream.CopyTo(copy);
            bytes = copy.ToArray();
        }
        catch (Exception ex)
        {
            return Prefight(false, "state-hash", ex.Message);
        }

        var actualStateSha = Sha256Hex(bytes);
        if (!string.Equals(actualStateSha, Json.GetString(metadata, "deferredStateSha256"), StringComparison.OrdinalIgnoreCase))
            return Prefight(false, "state-hash", "state.json hash does not match metadata deferredStateSha256");

        JsonObject parsed;
        try
        {
            parsed = JsonNode.Parse(bytes) as JsonObject
                ?? throw new InvalidDataException("state.json root is not an object");
        }
        catch (Exception ex)
        {
            return Prefight(false, "envelope", $"state.json could not be parsed from the hash-checked bytes: {ex.Message}");
        }

        if (!IsDeferredFormatEnvelope(parsed))
            return Prefight(false, "envelope", "state.json is not a complete copy-sheet-topology v4 deferred envelope");

        if (expectedState is not null
            && !string.Equals(Json.ToCompact(expectedState), Json.ToCompact(parsed), StringComparison.Ordinal))
        {
            return Prefight(false, "state-hash",
                "caller state is not structurally equal to the hash-authenticated state.json bytes");
        }

        if (!string.Equals(
                Json.GetString(metadata, "documentRef"),
                Json.GetString(parsed, "documentRef"),
                StringComparison.OrdinalIgnoreCase))
        {
            return Prefight(false, "documentRef", "metadata documentRef does not match state documentRef");
        }

        if (!string.Equals(
                Json.GetString(metadata, "deferredCheckpoint"),
                Json.GetString(parsed, "checkpointFile"),
                StringComparison.Ordinal))
        {
            return Prefight(false, "checkpoint-file", "metadata checkpoint name does not match state checkpointFile");
        }

        if (!string.Equals(
                Json.GetString(metadata, "deferredCheckpointSha256"),
                Json.GetString(parsed, "checkpointSha256"),
                StringComparison.OrdinalIgnoreCase))
        {
            return Prefight(false, "checkpoint-hash", "metadata checkpoint hash does not match state checkpointSha256");
        }

        boundState = parsed;
        return Prefight(true, "ok", "metadata is bound to the deferred format envelope");
    }

    internal static bool TryCaptureDeferredFormatCheckpoint(
        object workbook,
        object? attachedApplication,
        string snapshotDir,
        JsonObject metadata,
        IReadOnlyList<JsonObject>? ops,
        ExcelDeferredFormatSnapshotPolicy policy)
    {
        _ = attachedApplication;
        var request = EvaluateDeferredFormatRequestEligibility(policy, metadata, ops);
        metadata["deferredEligibility"] = request.DeepClone();
        if (!Json.GetBool(request, "eligible"))
            return false;

        var sheet = Json.GetString(request, "sheet");
        var workbookGate = EvaluateDeferredFormatWorkbookEligibility(workbook, sheet);
        CopyDecisionInputs(request, workbookGate);
        metadata["deferredEligibility"] = workbookGate.DeepClone();
        if (!Json.GetBool(workbookGate, "eligible"))
            return false;

        if (!TryResolveDeferredFormatCheckpointPath(snapshotDir, recordedFileName: null, out var destination, out var pathError))
        {
            metadata["deferredEligibility"] = RejectEligibility(workbookGate, "path", pathError);
            return false;
        }

        if (File.Exists(destination))
        {
            metadata["deferredEligibility"] = RejectEligibility(workbookGate, "path",
                "deferred checkpoint target already exists; refusing to overwrite it");
            return false;
        }

        object? owningApplication = null;
        try
        {
            owningApplication = TryBorrowWorkbookApplication(workbook);
            var identityBefore = ReadDeferredOwnerIdentity(workbook, owningApplication);
            if (!DeferredIdentityFullyReadable(identityBefore))
            {
                metadata["deferredEligibility"] = RejectEligibility(workbookGate, "identity",
                    "owner identity could not be fully read before SaveCopyAs");
                return false;
            }

            Exception? nativeFailure = null;
            var interaction = new ForegroundInteractionGuard("excel");
            try
            {
                TrackWorkbookForeground(interaction, owningApplication);
                try { ((dynamic)workbook).SaveCopyAs(destination); }
                catch (Exception ex) { nativeFailure = ex; }
            }
            finally { metadata["deferredFormatInteraction"] = interaction.Complete(); }

            var identityAfter = ReadDeferredOwnerIdentity(workbook, owningApplication);
            metadata["deferredFormatIdentity"] = new JsonObject
            {
                ["before"] = identityBefore.DeepClone(),
                ["after"] = identityAfter.DeepClone(),
            };
            var preserved = DeferredIdentityFullyReadable(identityAfter)
                && string.Equals(identityBefore.ToJsonString(), identityAfter.ToJsonString(), StringComparison.Ordinal);
            metadata["deferredFormatIdentityPreserved"] = preserved;
            if (!preserved)
            {
                RemovePartialNativeArtifact(destination, metadata);
                var failure =
                    "deferred format checkpoint: the target workbook's identity is not verified after SaveCopyAs " +
                    $"(before {identityBefore.ToJsonString()}, after {identityAfter.ToJsonString()}); " +
                    "the snapshot is refused so the pending edit does not proceed against an unverified workbook";
                if (nativeFailure is not null)
                    failure += $"; SaveCopyAs also failed with: {nativeFailure.Message}";
                else if (!File.Exists(destination))
                    failure += "; SaveCopyAs also produced no file";
                throw new InvalidOperationException(failure);
            }

            if (nativeFailure is not null || !File.Exists(destination))
            {
                RemovePartialNativeArtifact(destination, metadata);
                metadata["deferredEligibility"] = RejectEligibility(workbookGate, "identity",
                    nativeFailure?.Message ?? "SaveCopyAs produced no file");
                return false;
            }

            var preflight = PrefightDeferredFormatCheckpoint(destination);
            if (!Json.GetBool(preflight, "ok"))
            {
                RemovePartialNativeArtifact(destination, metadata);
                metadata["deferredEligibility"] = RejectEligibility(
                    workbookGate, Json.GetString(preflight, "code") ?? "bad-zip",
                    Json.GetString(preflight, "reason") ?? "checkpoint preflight failed");
                return false;
            }

            var checkpointSha = Sha256FileHex(destination);
            var op = ops![0];
            var scoped = CollectWrittenFormatScope(op);
            var eligibility = workbookGate.DeepClone().AsObject();
            eligibility["eligible"] = true;
            eligibility["used"] = true;
            eligibility["code"] = "ok";
            eligibility["reason"] = "deferred format checkpoint written";
            CopyDecisionInputs(request, eligibility);

            var state = new JsonObject
            {
                ["snapshotVersion"] = DeferredFormatSnapshotVersion,
                ["restoreMode"] = DeferredFormatCompatibilityRestoreMode,
                ["payloadMode"] = DeferredFormatPayloadMode,
                ["styleScope"] = WrittenStyleScope,
                ["documentRef"] = TryReadBackupFullName(workbook),
                ["ops"] = CloneOps(ops),
                ["checkpointFile"] = Path.GetFileName(destination),
                ["checkpointSha256"] = checkpointSha,
                ["cellCount"] = Json.GetInt(request, "targetCells"),
                ["targetCells"] = Json.GetInt(request, "targetCells"),
                ["workbookBytes"] = workbookGate["workbookBytes"]?.DeepClone(),
                ["minTargetCells"] = DeferredFormatMinCells,
                ["maxTargetCells"] = DeferredFormatMaxCells,
                ["maxWorkbookBytes"] = DeferredFormatMaxFileBytes,
                ["sheet"] = sheet,
                ["range"] = Json.GetString(request, "range"),
                ["scopedProperties"] = ToScopedPropertyArray(scoped),
                ["fingerprintSource"] = "checkpoint",
                ["deferredRestoreRequiresLiveSession"] = true,
                ["eligibility"] = eligibility.DeepClone(),
            };

            var statePath = Path.Combine(snapshotDir, "state.json");
            File.WriteAllText(statePath, state.ToJsonString(Json.Pretty));
            var stateSha = Sha256FileHex(statePath);

            metadata["payload"] = "deferred-format-checkpoint";
            metadata["documentRef"] = Json.GetString(state, "documentRef");
            metadata["restoreMode"] = DeferredFormatRestoreMode;
            metadata["snapshotKind"] = DeferredFormatRestoreMode;
            metadata["payloadMode"] = DeferredFormatPayloadMode;
            metadata["deferredCheckpoint"] = Json.GetString(state, "checkpointFile");
            metadata["deferredCheckpointSha256"] = checkpointSha;
            metadata["deferredStateSha256"] = stateSha;
            metadata["deferredEligibility"] = eligibility;
            metadata["deferredRestoreRequiresLiveSession"] = true;
            return true;
        }
        finally
        {
            RotHelper.ReleaseComReference(owningApplication);
        }
    }

    private JsonObject RestoreDeferredFormatOnlySnapshot(string snapshotDir, JsonObject metadata, JsonObject callerState)
    {
        var statePath = Path.Combine(snapshotDir, "state.json");
        var binding = ValidateDeferredFormatRestoreBinding(metadata, callerState, statePath, out var state);
        if (!Json.GetBool(binding, "ok") || state is null)
            return RefuseDeferredExtraction(Json.GetString(binding, "reason") ?? "deferred restore binding failed");

        if (!TryResolveDeferredFormatCheckpointPath(
                snapshotDir, Json.GetString(state, "checkpointFile"), out var checkpointPath, out var pathError))
        {
            return RefuseDeferredExtraction(pathError);
        }

        var ops = new List<JsonObject>();
        foreach (var node in Json.GetArr(state, "ops") ?? new JsonArray())
            if (node is JsonObject op) ops.Add(op);
        var documentRef = Json.GetString(state, "documentRef");
        var extractStarted = System.Diagnostics.Stopwatch.StartNew();

        FileStream? hold = null;
        WorkbookLease? lease = null;
        object? owningApplication = null;
        object? workbooks = null;
        object? checkpoint = null;
        object? preActive = null;
        object? previousSecurity = null;
        object? previousEvents = null;
        object? original = null;
        ForegroundInteractionGuard? interaction = null;
        var securitySet = false;
        var eventsSet = false;
        var activateOriginal = false;
        int? countBefore = null;
        var cleanupErrors = new List<string>();
        JsonObject? interactionReport = null;
        bool? eventsRestored = null;
        bool? securityRestored = null;
        bool? countRestored = null;
        JsonObject? result = null;
        try
        {
            var app = AttachExcel();
            if (app is null)
            {
                result = RefuseDeferredExtraction("Excel not running");
            }
            else
            {
                lease = ResolveSnapshotWorkbook((object)app, documentRef, ops, allowFileOpen: false);
                original = lease.Workbook;
                owningApplication = TryBorrowWorkbookApplication(original);
                if (owningApplication is null)
                {
                    result = RefuseDeferredExtraction("owning Application could not be borrowed");
                }
                else
                {
                    var identityBefore = ReadDeferredOwnerIdentity(original, owningApplication);
                    if (!DeferredIdentityFullyReadable(identityBefore))
                    {
                        result = RefuseDeferredExtraction("original workbook identity is unreadable before checkpoint open");
                    }
                    else
                    {
                        countBefore = TryReadOpenWorkbookCount(owningApplication);
                        preActive = TryBorrowActiveWorkbook(owningApplication);
                        activateOriginal = preActive is null || SameWorkbookFullName(preActive, original);
                        if (activateOriginal && preActive is not null)
                        {
                            RotHelper.ReleaseComReference(preActive);
                            preActive = null;
                        }

                        hold = new FileStream(checkpointPath, FileMode.Open, FileAccess.Read, FileShare.Read);
                        var heldSha = Sha256Hex(hold);
                        if (!string.Equals(heldSha, Json.GetString(metadata, "deferredCheckpointSha256"), StringComparison.OrdinalIgnoreCase))
                        {
                            result = RefuseDeferredExtraction("checkpoint hash does not match the bound metadata hash");
                        }
                        else
                        {
                            var preflight = PrefightDeferredFormatCheckpoint(hold);
                            if (!Json.GetBool(preflight, "ok"))
                            {
                                result = RefuseDeferredExtraction(Json.GetString(preflight, "reason") ?? "checkpoint preflight failed");
                            }
                            else
                            {
                                previousSecurity = TryReadComProperty(owningApplication, "AutomationSecurity");
                                previousEvents = TryReadComProperty(owningApplication, "EnableEvents");
                                if (previousSecurity is null || previousEvents is null)
                                {
                                    result = RefuseDeferredExtraction("AutomationSecurity or EnableEvents could not be read");
                                }
                                else
                                {
                                    ((dynamic)owningApplication).AutomationSecurity = MsoAutomationSecurityForceDisable;
                                    securitySet = true;
                                    ((dynamic)owningApplication).EnableEvents = false;
                                    eventsSet = true;

                                    interaction = new ForegroundInteractionGuard(App);
                                    TrackWorkbookForeground(interaction, owningApplication);
                                    workbooks = (object?)((dynamic)owningApplication).Workbooks;
                                    if (workbooks is null)
                                    {
                                        result = RefuseDeferredExtraction("Workbooks collection could not be borrowed");
                                    }
                                    else
                                    {
                                        // Borrowed RCWs from Application/Workbooks/Item are released with
                                        // ReleaseComReference only. The checkpoint we Open is ours: Close(false)
                                        // then ReleaseComReference — never FinalRelease, including on failure.
                                        checkpoint = (object)((dynamic)workbooks).Open(
                                            Filename: checkpointPath,
                                            UpdateLinks: 0,
                                            ReadOnly: true,
                                            IgnoreReadOnlyRecommended: true,
                                            Notify: false,
                                            AddToMru: false);
                                        var extracted = CaptureFormatOnlyState(checkpoint, ops, documentRef);
                                        extracted["fingerprintSource"] = "checkpoint";
                                        // Close before the full identity compare. identityBefore
                                        // includes workbookCount; comparing while this checkpoint
                                        // is still open makes count +1 and refuses every extract.
                                        // FileShare.Read and the foreground guard stay held.
                                        if (!TryCloseOwnedCheckpoint(ref checkpoint, cleanupErrors))
                                        {
                                            result = RefuseDeferredExtraction(
                                                "owned checkpoint could not be closed before identity comparison");
                                        }
                                        else
                                        {
                                            var identityAfter = ReadDeferredOwnerIdentity(original, owningApplication);
                                            if (!DeferredIdentityFullyReadable(identityAfter)
                                                || !string.Equals(
                                                    identityBefore.ToJsonString(),
                                                    identityAfter.ToJsonString(),
                                                    StringComparison.Ordinal))
                                            {
                                                result = RefuseDeferredExtraction(
                                                    "original FullName, Saved, or workbook count changed across checkpoint extract");
                                            }
                                            else
                                            {
                                                extractStarted.Stop();
                                                var restored = RestoreFormatOnlyState(original, extracted);
                                                restored["deferredRestoreRequiresLiveSession"] = true;
                                                restored["deferredExtraction"] = new JsonObject
                                                {
                                                    ["attempted"] = true,
                                                    ["ok"] = Json.GetBool(restored, "ok"),
                                                    ["extractMs"] = extractStarted.ElapsedMilliseconds,
                                                };
                                                result = restored;
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            result = RefuseDeferredExtraction(ex.Message, extractStarted.ElapsedMilliseconds);
        }
        finally
        {
            TryCloseOwnedCheckpoint(ref checkpoint, cleanupErrors);

            try
            {
                // Activate only after the checkpoint-open guard exists. Hash/preflight
                // refusals leave interaction null; Activate would steal foreground.
                if (interaction is not null)
                {
                    if (preActive is not null)
                        ((dynamic)preActive).Activate();
                    else if (activateOriginal && original is not null)
                        ((dynamic)original).Activate();
                }
            }
            catch (Exception ex)
            {
                cleanupErrors.Add($"pre-open ActiveWorkbook restore failed: {ex.Message}");
            }

            try
            {
                if (eventsSet && owningApplication is not null && previousEvents is not null)
                    ((dynamic)owningApplication).EnableEvents = previousEvents;
            }
            catch (Exception ex)
            {
                cleanupErrors.Add($"EnableEvents restore failed: {ex.Message}");
            }

            try
            {
                if (securitySet && owningApplication is not null && previousSecurity is not null)
                    ((dynamic)owningApplication).AutomationSecurity = previousSecurity;
            }
            catch (Exception ex)
            {
                cleanupErrors.Add($"AutomationSecurity restore failed: {ex.Message}");
            }

            if (eventsSet)
            {
                var currentEvents = owningApplication is null
                    ? null
                    : TryReadComProperty(owningApplication, "EnableEvents");
                var eventsMatch = DeferredSettingsEqual(currentEvents, previousEvents);
                eventsRestored = eventsMatch;
                if (eventsMatch != true)
                    cleanupErrors.Add("EnableEvents readback is not the original value");
            }

            if (securitySet)
            {
                var currentSecurity = owningApplication is null
                    ? null
                    : TryReadComProperty(owningApplication, "AutomationSecurity");
                var securityMatch = DeferredSettingsEqual(currentSecurity, previousSecurity);
                securityRestored = securityMatch;
                if (securityMatch != true)
                    cleanupErrors.Add("AutomationSecurity readback is not the original value");
            }

            if (countBefore is not null)
            {
                var countAfter = TryReadOpenWorkbookCount(owningApplication);
                var countMatch = countAfter is not null && countAfter == countBefore;
                countRestored = countMatch;
                if (!countMatch)
                    cleanupErrors.Add("open workbook count readback does not match the pre-open count");
            }

            if (interaction is not null)
            {
                try { interactionReport = interaction.Complete(); }
                catch (Exception ex) { cleanupErrors.Add($"foreground guard complete failed: {ex.Message}"); }
            }

            RotHelper.ReleaseComReference(workbooks);
            RotHelper.ReleaseComReference(preActive);
            RotHelper.ReleaseComReference(owningApplication);
            try { lease?.Dispose(); }
            catch (Exception ex) { cleanupErrors.Add($"workbook lease dispose failed: {ex.Message}"); }
            hold?.Dispose();
        }

        result ??= RefuseDeferredExtraction("deferred restore produced no result", extractStarted.ElapsedMilliseconds);
        if (cleanupErrors.Count > 0)
        {
            result = UnverifiedDeferredCleanup(
                "deferred restore session cleanup or settings readback failed",
                extractStarted.ElapsedMilliseconds,
                result,
                cleanupErrors);
        }

        // Restore RPC returns this object only. Child-local metadata mutations never
        // reach CLI/MCP or the host audit record, so session evidence lives here.
        AttachDeferredSessionEvidence(
            result, interactionReport, eventsRestored, securityRestored, countRestored, cleanupErrors);
        return result;
    }

    private static JsonObject UnverifiedDeferredCleanup(
        string reason,
        long extractMs,
        JsonObject restoreAttempt,
        IReadOnlyList<string> cleanupErrors)
    {
        var errors = new JsonArray { reason };
        var cleanup = new JsonArray();
        foreach (var item in cleanupErrors)
        {
            errors.Add(item);
            cleanup.Add(item);
        }

        return new JsonObject
        {
            ["ok"] = false,
            ["restored"] = false,
            ["restoreMode"] = DeferredFormatRestoreMode,
            ["deferredRestoreRequiresLiveSession"] = true,
            ["deferredExtraction"] = new JsonObject
            {
                ["attempted"] = true,
                ["ok"] = false,
                ["extractMs"] = extractMs,
                ["reason"] = reason,
                ["cleanupFailed"] = true,
                ["cleanupErrors"] = cleanup,
                ["restoreAttempt"] = restoreAttempt.DeepClone(),
            },
            ["deferredExtractionFailed"] = reason,
            ["errors"] = errors,
        };
    }

    private static void AttachDeferredSessionEvidence(
        JsonObject result,
        JsonObject? interaction,
        bool? eventsRestored,
        bool? securityRestored,
        bool? countRestored,
        IReadOnlyList<string> cleanupErrors)
    {
        if (interaction is not null)
            result["deferredFormatRestoreInteraction"] = interaction.DeepClone();

        var cleanup = new JsonArray();
        foreach (var item in cleanupErrors)
            cleanup.Add(item);

        var session = new JsonObject
        {
            ["enableEventsRestored"] = eventsRestored is bool events ? JsonValue.Create(events) : null,
            ["automationSecurityRestored"] = securityRestored is bool security ? JsonValue.Create(security) : null,
            ["workbookCountRestored"] = countRestored is bool count ? JsonValue.Create(count) : null,
            ["cleanupFailed"] = cleanupErrors.Count > 0,
            ["cleanupErrors"] = cleanup,
        };
        if (interaction is not null)
            session["foreground"] = interaction.DeepClone();

        result["deferredSession"] = session;
        if (Json.GetObj(result, "deferredExtraction") is JsonObject extraction)
        {
            extraction["session"] = session.DeepClone();
            if (interaction is not null)
                extraction["foreground"] = interaction.DeepClone();
            if (cleanupErrors.Count > 0)
            {
                extraction["cleanupFailed"] = true;
                extraction["cleanupErrors"] = cleanup.DeepClone();
            }
        }
    }

    private static bool TryCloseOwnedCheckpoint(ref object? checkpoint, List<string> errors)
    {
        if (checkpoint is null) return true;
        try
        {
            ((dynamic)checkpoint).Close(false);
        }
        catch (Exception ex)
        {
            errors.Add($"checkpoint close failed: {ex.Message}");
            RotHelper.ReleaseComReference(checkpoint);
            checkpoint = null;
            return false;
        }

        RotHelper.ReleaseComReference(checkpoint);
        checkpoint = null;
        return true;
    }

    private static JsonObject RefuseDeferredExtraction(string reason, long extractMs = 0) => new()
    {
        ["ok"] = false,
        ["restored"] = false,
        ["restoreMode"] = DeferredFormatRestoreMode,
        ["deferredRestoreRequiresLiveSession"] = true,
        ["deferredExtraction"] = new JsonObject
        {
            ["attempted"] = true,
            ["ok"] = false,
            ["extractMs"] = extractMs,
            ["reason"] = reason,
        },
        ["deferredExtractionFailed"] = reason,
        ["errors"] = new JsonArray(reason),
    };

    private static JsonObject EligibilityBase() => new()
    {
        ["eligible"] = false,
        ["used"] = false,
        ["code"] = "",
        ["reason"] = "",
        ["minTargetCells"] = DeferredFormatMinCells,
        ["maxTargetCells"] = DeferredFormatMaxCells,
        ["maxWorkbookBytes"] = DeferredFormatMaxFileBytes,
    };

    private static JsonObject RejectEligibility(JsonObject current, string code, string reason)
    {
        current["eligible"] = false;
        current["used"] = false;
        current["code"] = code;
        current["reason"] = reason;
        return current;
    }

    private static JsonObject Prefight(bool ok, string code, string reason) => new()
    {
        ["ok"] = ok,
        ["code"] = code,
        ["reason"] = reason,
    };

    private static void CopyDecisionInputs(JsonObject source, JsonObject destination)
    {
        foreach (var key in new[] { "targetCells", "hasFillColor", "sheet", "range", "minTargetCells", "maxTargetCells", "maxWorkbookBytes" })
        {
            if (source[key] is not null && destination[key] is null)
                destination[key] = source[key]!.DeepClone();
        }
    }

    private static bool TryParseSingleRectangularA1(
        string rangeText, string expectedSheet, out string address, out int cells)
    {
        address = "";
        cells = 0;
        try
        {
            var parsed = ExcelRangeReference.Parse(rangeText);
            if (!string.IsNullOrWhiteSpace(parsed.SheetName)
                && !string.Equals(parsed.SheetName, expectedSheet, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            if (!TryParseA1Range(parsed.Address, out _, out _, out var rows, out var columns))
                return false;
            address = parsed.Address;
            cells = rows * columns;
            return cells >= 1;
        }
        catch (FormatException)
        {
            return false;
        }
    }

    private static bool IsLocalExistingXlsx(string? fullName, out string reason)
    {
        reason = "";
        if (string.IsNullOrWhiteSpace(fullName))
        {
            reason = "workbook FullName is empty";
            return false;
        }

        if (!Path.IsPathFullyQualified(fullName)
            || fullName.StartsWith(@"\\", StringComparison.Ordinal)
            || fullName.Length < 3
            || !char.IsAsciiLetter(fullName[0])
            || fullName[1] != ':')
        {
            reason = "workbook path is not a local drive-letter file";
            return false;
        }

        if (!string.Equals(Path.GetExtension(fullName), ".xlsx", StringComparison.OrdinalIgnoreCase))
        {
            reason = "workbook extension is not .xlsx";
            return false;
        }

        if (!File.Exists(fullName))
        {
            reason = "workbook FullName does not point at an existing file";
            return false;
        }

        return true;
    }

    private static JsonObject ReadDeferredOwnerIdentity(object workbook, object? owningApplication)
    {
        var identity = ReadWorkbookBackupIdentity(workbook, owningApplication);
        var hwnd = TryReadApplicationHwnd(owningApplication);
        identity["hwnd"] = hwnd is long value ? JsonValue.Create(value) : null;
        return identity;
    }

    private static bool DeferredIdentityFullyReadable(JsonObject identity) =>
        IdentityFullyReadable(identity)
        && identity["hwnd"] is JsonValue hwnd
        && hwnd.TryGetValue<long>(out _);

    private static long? TryReadApplicationHwnd(object? application)
    {
        if (application is null) return null;
        try { return Convert.ToInt64(((dynamic)application).Hwnd, CultureInfo.InvariantCulture); }
        catch { return null; }
    }

    private static bool? TryReadComBool(object target, string propertyName)
    {
        try
        {
            object? raw = propertyName switch
            {
                "HasVBProject" => ((dynamic)target).HasVBProject,
                "ProtectStructure" => ((dynamic)target).ProtectStructure,
                "ProtectWindows" => ((dynamic)target).ProtectWindows,
                "HasPassword" => ((dynamic)target).HasPassword,
                "WriteReserved" => ((dynamic)target).WriteReserved,
                "ProtectContents" => ((dynamic)target).ProtectContents,
                _ => null,
            };
            return TryCoerceComBool(raw);
        }
        catch
        {
            return null;
        }
    }

    private static bool? TryCoerceComBool(object? raw)
    {
        if (raw is null || raw is DBNull || ReferenceEquals(raw, Missing.Value))
            return null;
        if (raw is bool already) return already;
        try { return Convert.ToBoolean(raw, CultureInfo.InvariantCulture); }
        catch { return null; }
    }

    private static object? TryReadComProperty(object target, string propertyName)
    {
        try
        {
            object? raw = propertyName switch
            {
                "AutomationSecurity" => (object?)((dynamic)target).AutomationSecurity,
                "EnableEvents" => (object?)((dynamic)target).EnableEvents,
                _ => null,
            };
            if (raw is null || raw is DBNull || ReferenceEquals(raw, Missing.Value))
                return null;
            return raw;
        }
        catch
        {
            return null;
        }
    }

    private static bool? DeferredSettingsEqual(object? left, object? right)
    {
        if (left is null || right is null) return null;
        var leftBool = TryCoerceComBool(left);
        var rightBool = TryCoerceComBool(right);
        if (leftBool is not null && rightBool is not null)
            return leftBool == rightBool;
        try
        {
            return Convert.ToInt32(left, CultureInfo.InvariantCulture)
                   == Convert.ToInt32(right, CultureInfo.InvariantCulture);
        }
        catch
        {
            return null;
        }
    }

    private static bool? TryWorkbookHasExcelLinks(object workbook)
    {
        try
        {
            object? raw = ((dynamic)workbook).LinkSources(1);
            if (raw is null || ReferenceEquals(raw, Missing.Value)) return false;
            if (raw is string text) return !string.IsNullOrWhiteSpace(text);
            if (raw is Array { Length: 0 }) return false;
            if (raw is Array) return true;
            return null;
        }
        catch
        {
            return null;
        }
    }

    private static bool? TargetSheetProtected(object workbook, string? targetSheet)
    {
        if (string.IsNullOrWhiteSpace(targetSheet)) return null;
        try
        {
            dynamic sheet = GetSheet(workbook, targetSheet);
            return TryReadComBool((object)sheet, "ProtectContents");
        }
        catch
        {
            return null;
        }
    }

    private static object? TryBorrowActiveWorkbook(object application)
    {
        try { return (object?)((dynamic)application).ActiveWorkbook; }
        catch { return null; }
    }

    private static bool SameWorkbookFullName(object left, object right) =>
        string.Equals(TryReadBackupFullName(left), TryReadBackupFullName(right), StringComparison.OrdinalIgnoreCase);

    private static string Sha256FileHex(string path)
    {
        using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read);
        return Sha256Hex(stream);
    }

    private static string Sha256Hex(byte[] bytes) =>
        Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static string Sha256Hex(Stream stream)
    {
        if (stream.CanSeek) stream.Position = 0;
        var hash = SHA256.HashData(stream);
        if (stream.CanSeek) stream.Position = 0;
        return Convert.ToHexString(hash).ToLowerInvariant();
    }

    private static bool StreamContainsUtf16Marker(Stream stream, string ascii)
    {
        var marker = System.Text.Encoding.Unicode.GetBytes(ascii);
        if (stream.CanSeek) stream.Position = 0;
        var previous = Array.Empty<byte>();
        var buffer = new byte[1 << 20];
        while (true)
        {
            var read = stream.Read(buffer, 0, buffer.Length);
            if (read <= 0)
            {
                if (stream.CanSeek) stream.Position = 0;
                return false;
            }

            var chunk = read == buffer.Length ? buffer : buffer.AsSpan(0, read).ToArray();
            var combined = new byte[previous.Length + chunk.Length];
            previous.CopyTo(combined, 0);
            chunk.CopyTo(combined.AsSpan(previous.Length));
            if (combined.AsSpan().IndexOf(marker) >= 0)
            {
                if (stream.CanSeek) stream.Position = 0;
                return true;
            }

            previous = combined.Length >= marker.Length
                ? combined[^marker.Length..]
                : combined;
        }
    }

    private static JsonObject InspectDeferredFormatPackage(Stream stream)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        var parts = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
        foreach (var entry in zip.Entries)
        {
            var name = NormalizePartName(entry.FullName);
            if (name.Length == 0 || name.EndsWith('/')) continue;
            if (!parts.TryAdd(name, entry))
                return Prefight(false, "duplicate-package-part", $"duplicate package part '{name}'");
        }

        if (parts.Count > DeferredFormatMaxPackageParts)
        {
            return Prefight(false, "complex-part",
                $"checkpoint has {parts.Count} parts, above the {DeferredFormatMaxPackageParts} plain-data cap");
        }

        if (!parts.ContainsKey("[Content_Types].xml"))
            return Prefight(false, "missing-content-types", "[Content_Types].xml is missing");
        if (!parts.ContainsKey("xl/workbook.xml"))
            return Prefight(false, "missing-workbook-xml", "xl/workbook.xml is missing");

        foreach (var name in parts.Keys)
        {
            var pathError = ClassifyRestrictedPartPath(name);
            if (pathError is not null) return pathError;
        }

        var xmlBudget = DeferredFormatMaxUncompressedXmlBytes;
        var typeError = InspectContentTypes(parts["[Content_Types].xml"], ref xmlBudget);
        if (typeError is not null) return typeError;

        foreach (var name in parts.Keys.Where(IsRelationshipPart).OrderBy(item => item, StringComparer.OrdinalIgnoreCase))
        {
            var relError = InspectRelationships(parts[name], name, parts.Keys, ref xmlBudget);
            if (relError is not null) return relError;
        }

        foreach (var (name, entry) in parts)
        {
            if (!name.EndsWith(".xml", StringComparison.OrdinalIgnoreCase)
                && !name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            if (IsRelationshipPart(name) || name.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase))
                continue;

            var walkError = WalkSpreadsheetXml(entry, name, ref xmlBudget);
            if (walkError is not null) return walkError;
        }

        return Prefight(true, "ok", "checkpoint bytes are inside the restricted deferred format route");
    }

    private static JsonObject? ClassifyRestrictedPartPath(string name)
    {
        var normalized = name.ToLowerInvariant();
        if (normalized.EndsWith("vbaproject.bin", StringComparison.Ordinal))
            return Prefight(false, "vba-project", "checkpoint contains vbaProject.bin");
        if (normalized.Contains("macrosheet", StringComparison.Ordinal)
            || normalized.Contains("xlmacrosheet", StringComparison.Ordinal))
        {
            return Prefight(false, "xlm-macrosheet", "checkpoint contains an Excel 4.0 macrosheet part");
        }

        if (normalized.StartsWith("xl/externallinks/", StringComparison.Ordinal)
            || normalized.Contains("/externallink", StringComparison.Ordinal))
        {
            return Prefight(false, "external-links", "checkpoint contains external links");
        }

        if (normalized is "xl/connections.xml" || normalized.StartsWith("xl/connections/", StringComparison.Ordinal)
            || normalized.StartsWith("xl/querytables/", StringComparison.Ordinal))
        {
            return Prefight(false, "connections", "checkpoint contains connections");
        }

        if (normalized.Contains("oleobject", StringComparison.Ordinal)
            || normalized.Contains("activex", StringComparison.Ordinal)
            || normalized.Contains("/embeddings/", StringComparison.Ordinal))
        {
            return Prefight(false, "ole-or-activex", "checkpoint contains OLE or ActiveX parts");
        }

        if (normalized is "xl/calcchain.xml"
            || normalized.StartsWith("xl/tables/", StringComparison.Ordinal)
            || normalized is "xl/volatiledependencies.xml")
        {
            return Prefight(false, "formulas", "checkpoint contains formula-capable parts");
        }

        if (IsAllowedPlainDataPart(name))
            return null;

        return Prefight(false, "complex-part",
            $"checkpoint contains '{name}', outside the plain-data vocabulary");
    }

    private static bool IsAllowedPlainDataPart(string name)
    {
        if (name.Equals("[Content_Types].xml", StringComparison.OrdinalIgnoreCase)) return true;
        if (name.Equals("_rels/.rels", StringComparison.OrdinalIgnoreCase)) return true;
        if (name.Equals("docProps/core.xml", StringComparison.OrdinalIgnoreCase)) return true;
        if (name.Equals("docProps/app.xml", StringComparison.OrdinalIgnoreCase)) return true;
        if (name.Equals("docProps/custom.xml", StringComparison.OrdinalIgnoreCase)) return true;
        if (name.Equals("xl/workbook.xml", StringComparison.OrdinalIgnoreCase)) return true;
        if (name.Equals("xl/_rels/workbook.xml.rels", StringComparison.OrdinalIgnoreCase)) return true;
        if (name.Equals("xl/styles.xml", StringComparison.OrdinalIgnoreCase)) return true;
        if (name.Equals("xl/sharedStrings.xml", StringComparison.OrdinalIgnoreCase)) return true;
        if (name.Equals("xl/theme/theme1.xml", StringComparison.OrdinalIgnoreCase)) return true;
        if (IsWorksheetPart(name) || IsWorksheetRelsPart(name) || IsPrinterSettingsPart(name))
            return true;
        return false;
    }

    private static bool IsWorksheetPart(string name)
    {
        const string prefix = "xl/worksheets/sheet";
        const string suffix = ".xml";
        return HasNumericStem(name, prefix, suffix);
    }

    private static bool IsWorksheetRelsPart(string name)
    {
        const string prefix = "xl/worksheets/_rels/sheet";
        const string suffix = ".xml.rels";
        return HasNumericStem(name, prefix, suffix);
    }

    private static bool IsPrinterSettingsPart(string name)
    {
        const string prefix = "xl/printerSettings/printerSettings";
        const string suffix = ".bin";
        return HasNumericStem(name, prefix, suffix);
    }

    private static bool HasNumericStem(string name, string prefix, string suffix)
    {
        if (!name.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)
            || !name.EndsWith(suffix, StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        var stem = name[prefix.Length..^suffix.Length];
        return stem.Length > 0 && stem.All(char.IsAsciiDigit);
    }

    private static bool IsRelationshipPart(string name) =>
        name.Equals("_rels/.rels", StringComparison.OrdinalIgnoreCase)
        || name.EndsWith(".rels", StringComparison.OrdinalIgnoreCase);

    private static JsonObject? InspectContentTypes(ZipArchiveEntry entry, ref long xmlBudget)
    {
        string? mainType = null;
        var error = WalkXml(entry, "[Content_Types].xml", ref xmlBudget, (reader, local, ns) =>
        {
            if (reader.NodeType != XmlNodeType.Element) return null;
            if (local is "Types" or "Default" or "Override"
                && !IsKnownNamespace(ns, PackageContentTypesNs))
            {
                return Prefight(false, "missing-content-types",
                    "[Content_Types].xml uses an unexpected namespace");
            }

            if (local is not ("Default" or "Override"))
                return null;

            var attrs = ReadAttributes(reader);
            var contentType = attrs.GetValueOrDefault("ContentType") ?? "";
            if (ContainsInsensitive(contentType, "binary.macroEnabled")
                || string.Equals(contentType,
                    "application/vnd.ms-excel.sheet.binary.macroEnabled.main",
                    StringComparison.OrdinalIgnoreCase))
            {
                return Prefight(false, "xlsb-binary", "content types declare a binary workbook main part");
            }

            if (ContainsInsensitive(contentType, "macroEnabled"))
                return Prefight(false, "macro-enabled-main", "content types declare a macro-enabled main part");
            if (ContainsInsensitive(contentType, "macrosheet"))
                return Prefight(false, "xlm-macrosheet", "content types declare an Excel 4.0 macrosheet");

            if (local == "Override")
            {
                var partName = NormalizePartName(attrs.GetValueOrDefault("PartName") ?? "");
                if (partName.Equals("xl/workbook.xml", StringComparison.OrdinalIgnoreCase))
                    mainType = contentType;
            }

            return null;
        });
        if (error is not null) return error;
        if (string.IsNullOrWhiteSpace(mainType))
            return Prefight(false, "missing-content-types", "workbook main Override is missing");
        if (!string.Equals(mainType, XlsxMainContentType, StringComparison.OrdinalIgnoreCase))
        {
            return Prefight(false, "complex-part",
                "main content type is not an ordinary xlsx workbook");
        }

        return null;
    }

    private static JsonObject? InspectRelationships(
        ZipArchiveEntry entry,
        string partName,
        IReadOnlyCollection<string> packageNames,
        ref long xmlBudget)
    {
        return WalkXml(entry, partName, ref xmlBudget, (reader, local, ns) =>
        {
            if (reader.NodeType != XmlNodeType.Element) return null;
            if (local is "Relationships" or "Relationship"
                && !IsKnownNamespace(ns, PackageRelsNs))
            {
                return Prefight(false, "relationship-target",
                    $"{partName} uses an unexpected relationships namespace");
            }

            if (local != "Relationship")
                return null;

            var attrs = ReadAttributes(reader);
            var type = attrs.GetValueOrDefault("Type") ?? "";
            var target = attrs.GetValueOrDefault("Target") ?? "";
            var mode = attrs.GetValueOrDefault("TargetMode") ?? "Internal";
            if (string.Equals(mode, "External", StringComparison.OrdinalIgnoreCase))
                return Prefight(false, "external-links", $"{partName} has an External relationship");

            var typed = ClassifyRelationshipType(type);
            if (typed is not null) return typed;

            if (!TryResolveInternalTarget(partName, target, out var resolved, out var resolveError))
            {
                return Prefight(false, "relationship-target",
                    $"{partName} target '{target}' is invalid ({resolveError})");
            }

            if (!packageNames.Contains(resolved, StringComparer.OrdinalIgnoreCase))
            {
                return Prefight(false, "relationship-target",
                    $"{partName} target '{target}' does not resolve to a package part");
            }

            return null;
        });
    }

    private static JsonObject? ClassifyRelationshipType(string typeUri)
    {
        var local = RelationshipLocalName(typeUri);
        if (AllowedRelationshipLocals.Contains(local))
            return null;
        if (local.Equals("vbaProject", StringComparison.OrdinalIgnoreCase))
            return Prefight(false, "vba-project", "relationships name a VBA project");
        if (local.Contains("macrosheet", StringComparison.OrdinalIgnoreCase))
            return Prefight(false, "xlm-macrosheet", "relationships name an Excel 4.0 macrosheet");
        if (local.Equals("externalLink", StringComparison.OrdinalIgnoreCase)
            || local.Equals("hyperlink", StringComparison.OrdinalIgnoreCase))
        {
            return Prefight(false, "external-links", "relationships name an external link");
        }

        if (local.Contains("ole", StringComparison.OrdinalIgnoreCase)
            || local.Contains("activeX", StringComparison.OrdinalIgnoreCase)
            || local.Equals("control", StringComparison.OrdinalIgnoreCase)
            || local.Equals("ctrlProp", StringComparison.OrdinalIgnoreCase))
        {
            return Prefight(false, "ole-or-activex", "relationships name OLE or ActiveX");
        }

        if (local.Equals("connections", StringComparison.OrdinalIgnoreCase)
            || local.Equals("queryTable", StringComparison.OrdinalIgnoreCase)
            || local.Equals("model", StringComparison.OrdinalIgnoreCase))
        {
            return Prefight(false, "connections", "relationships name connections");
        }

        if (local.Equals("calcChain", StringComparison.OrdinalIgnoreCase)
            || local.Equals("volatileDependencies", StringComparison.OrdinalIgnoreCase)
            || local.Equals("table", StringComparison.OrdinalIgnoreCase))
        {
            return Prefight(false, "formulas", "relationships name a formula-capable part");
        }

        return Prefight(false, "complex-part", $"relationships name '{local}', outside the plain-data vocabulary");
    }

    private static JsonObject? WalkSpreadsheetXml(ZipArchiveEntry entry, string partName, ref long xmlBudget)
    {
        var workbook = partName.Equals("xl/workbook.xml", StringComparison.OrdinalIgnoreCase);
        var sharedStrings = partName.Equals("xl/sharedStrings.xml", StringComparison.OrdinalIgnoreCase);
        var inlineStringDepth = 0;
        return WalkXml(entry, partName, ref xmlBudget, (reader, local, ns) =>
        {
            if (reader.NodeType == XmlNodeType.EndElement)
            {
                if (local == "is" && inlineStringDepth > 0)
                    inlineStringDepth--;
                return null;
            }

            if (reader.NodeType != XmlNodeType.Element)
                return null;

            if (workbook && local == "workbook")
            {
                if (string.Equals(ns, SpreadsheetMlStrictNs, StringComparison.OrdinalIgnoreCase)
                    || ContainsInsensitive(ns, "purl.oclc.org/ooxml/spreadsheetml"))
                {
                    return Prefight(false, "strict-ooxml", "workbook uses Strict OOXML");
                }

                if (!string.Equals(ns, SpreadsheetMlTransitionalNs, StringComparison.OrdinalIgnoreCase))
                    return Prefight(false, "strict-ooxml", "workbook is not transitional SpreadsheetML");
            }

            if (local is "xlMacrosheet" or "xlIntlMacrosheet" or "macrosheet")
                return Prefight(false, "xlm-macrosheet", "workbook names an Excel 4.0 macrosheet");
            if (local == "definedName")
                return Prefight(false, "defined-names", "checkpoint contains defined names");
            if (local is "conditionalFormatting" or "cfRule")
                return Prefight(false, "conditional-formatting", "checkpoint contains conditional formatting");
            if (local is "dataValidations" or "dataValidation")
                return Prefight(false, "data-validation", "checkpoint contains data validation");
            if (local is "mergeCells" or "mergeCell")
                return Prefight(false, "merge-cells", "checkpoint contains merged cells");
            if (local is "f" or "calculatedColumn" or "calculatedColumnFormula")
                return Prefight(false, "formulas", "checkpoint contains formulas; initial route is plain data only");

            if (local == "calcPr")
            {
                var attrs = ReadAttributes(reader);
                if (attrs.TryGetValue("fullCalcOnLoad", out var raw))
                {
                    var flag = TryParseOoxmlBool(raw);
                    if (flag != false)
                    {
                        return Prefight(false, "full-calc-on-load",
                            flag is null
                                ? "workbook fullCalcOnLoad is unreadable"
                                : "workbook requests fullCalcOnLoad");
                    }
                }
            }

            if (local == "is")
                inlineStringDepth++;
            if (local == "rPh"
                || (local == "r" && (sharedStrings || inlineStringDepth > 0)))
            {
                return Prefight(false, "rich-text", "checkpoint contains rich-text runs");
            }

            if (local is "color" or "fgColor" or "bgColor" or "srgbClr")
            {
                var attrs = ReadAttributes(reader);
                if (HasUnsupportedRgbTint(local, attrs))
                    return Prefight(false, "rgb-tint", "checkpoint has non-theme RGB color with nonzero tint");
            }

            return null;
        });
    }

    private static JsonObject? WalkXml(
        ZipArchiveEntry entry,
        string partName,
        ref long xmlBudget,
        Func<XmlReader, string, string, JsonObject?> onNode)
    {
        if (entry.Length > DeferredFormatMaxXmlPartBytes)
        {
            return Prefight(false, "uncompressed-xml",
                $"{partName} declares {entry.Length} uncompressed bytes, above the XML part cap");
        }

        if (entry.Length > xmlBudget)
        {
            return Prefight(false, "uncompressed-xml",
                $"checkpoint XML exceeds the {DeferredFormatMaxUncompressedXmlBytes} uncompressed budget");
        }

        var partBudget = Math.Min(DeferredFormatMaxXmlPartBytes, xmlBudget);
        using var raw = entry.Open();
        using var bounded = new BudgetStream(raw, partBudget);
        var settings = new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            IgnoreComments = true,
            IgnoreProcessingInstructions = true,
            IgnoreWhitespace = true,
            CheckCharacters = true,
            CloseInput = false,
            MaxCharactersFromEntities = 0,
            MaxCharactersInDocument = partBudget,
            ConformanceLevel = ConformanceLevel.Document,
        };

        try
        {
            using var reader = XmlReader.Create(bounded, settings);
            while (reader.Read())
            {
                if (reader.NodeType is not (XmlNodeType.Element or XmlNodeType.EndElement))
                    continue;
                var error = onNode(reader, reader.LocalName, reader.NamespaceURI);
                if (error is not null) return error;
            }
        }
        catch (XmlException ex) when (IsDtdProhibited(ex))
        {
            return Prefight(false, "dtd-prohibited", $"{partName}: {ex.Message}");
        }
        catch (XmlException ex)
        {
            return Prefight(false, "malformed-xml", $"{partName}: {ex.Message}");
        }
        catch (InvalidDataException ex)
        {
            return Prefight(false, "uncompressed-xml", $"{partName}: {ex.Message}");
        }

        xmlBudget -= bounded.BytesRead;
        if (xmlBudget < 0)
        {
            return Prefight(false, "uncompressed-xml",
                $"checkpoint XML exceeds the {DeferredFormatMaxUncompressedXmlBytes} uncompressed budget");
        }

        return null;
    }

    private static bool HasUnsupportedRgbTint(string local, IReadOnlyDictionary<string, string> attrs)
    {
        var hasTheme = attrs.ContainsKey("theme")
                       || local.Equals("schemeClr", StringComparison.OrdinalIgnoreCase);
        var hasRgb = attrs.ContainsKey("rgb")
                     || (local.Equals("srgbClr", StringComparison.OrdinalIgnoreCase) && attrs.ContainsKey("val"));
        if (!hasRgb || hasTheme || !attrs.TryGetValue("tint", out var tintRaw))
            return false;
        if (!double.TryParse(tintRaw, NumberStyles.Float, CultureInfo.InvariantCulture, out var tint))
            return true;
        return Math.Abs(tint) > 1e-9;
    }

    private static bool TryResolveInternalTarget(string relsPart, string target, out string resolved, out string error)
    {
        resolved = "";
        error = "";
        if (string.IsNullOrWhiteSpace(target))
        {
            error = "empty";
            return false;
        }

        var value = target.Replace('\\', '/');
        if (value.Contains("://", StringComparison.Ordinal) || value.Contains(':'))
        {
            error = "scheme";
            return false;
        }

        var rels = NormalizePartName(relsPart);
        var relsDir = rels.Contains('/', StringComparison.Ordinal)
            ? rels[..rels.LastIndexOf('/')]
            : "";
        var baseDir = relsDir.Equals("_rels", StringComparison.OrdinalIgnoreCase)
            ? ""
            : relsDir.EndsWith("/_rels", StringComparison.OrdinalIgnoreCase)
                ? relsDir[..^6]
                : relsDir;

        resolved = value.StartsWith('/')
            ? NormalizePartName(value)
            : string.IsNullOrEmpty(baseDir) ? NormalizePartName(value) : NormalizePartName($"{baseDir}/{value}");
        resolved = resolved.Replace("/./", "/", StringComparison.Ordinal);
        if (resolved.StartsWith("./", StringComparison.Ordinal))
            resolved = resolved[2..];
        if (resolved.Contains("..", StringComparison.Ordinal) || resolved.StartsWith('/'))
        {
            error = "escape";
            resolved = "";
            return false;
        }

        return true;
    }

    private static string RelationshipLocalName(string typeUri)
    {
        var trimmed = typeUri.Trim().TrimEnd('/');
        var slash = trimmed.LastIndexOf('/');
        return slash >= 0 ? trimmed[(slash + 1)..] : trimmed;
    }

    private static Dictionary<string, string> ReadAttributes(XmlReader reader)
    {
        var attrs = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (!reader.MoveToFirstAttribute())
            return attrs;
        do { attrs[reader.LocalName] = reader.Value; }
        while (reader.MoveToNextAttribute());
        reader.MoveToElement();
        return attrs;
    }

    private static bool IsKnownNamespace(string actual, string expected) =>
        string.Equals(actual, expected, StringComparison.OrdinalIgnoreCase);

    private static bool? TryParseOoxmlBool(string? raw)
    {
        if (raw is null) return null;
        var value = raw.Trim();
        if (value.Equals("1", StringComparison.OrdinalIgnoreCase)
            || value.Equals("true", StringComparison.OrdinalIgnoreCase)
            || value.Equals("on", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        if (value.Equals("0", StringComparison.OrdinalIgnoreCase)
            || value.Equals("false", StringComparison.OrdinalIgnoreCase)
            || value.Equals("off", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        return null;
    }

    private static bool IsDtdProhibited(XmlException ex) =>
        ex.Message.Contains("DTD", StringComparison.OrdinalIgnoreCase);

    private static string NormalizePartName(string name) =>
        name.Replace('\\', '/').TrimStart('/');

    private static bool ContainsInsensitive(string text, string value) =>
        text.Contains(value, StringComparison.OrdinalIgnoreCase);

    private sealed class BudgetStream : Stream
    {
        private readonly Stream _inner;
        private readonly long _max;

        public BudgetStream(Stream inner, long max)
        {
            _inner = inner;
            _max = max;
        }

        public long BytesRead { get; private set; }
        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override void Flush() => _inner.Flush();
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        public override int Read(byte[] buffer, int offset, int count)
        {
            var read = _inner.Read(buffer, offset, count);
            BytesRead += read;
            if (BytesRead > _max)
                throw new InvalidDataException("uncompressed XML budget exceeded");
            return read;
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing) _inner.Dispose();
            base.Dispose(disposing);
        }
    }
}
