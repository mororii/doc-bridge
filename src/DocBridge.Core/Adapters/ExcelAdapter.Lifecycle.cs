using System.Globalization;
using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class ExcelAdapter
{
    private const int XlTypePdf = 0;

    internal static bool IsCreateOrOpenOnly(IReadOnlyList<JsonObject>? ops) =>
        ops is { Count: 1 } &&
        Json.GetString(ops[0], "op") is "create_workbook" or "open_workbook";

    internal static bool IsCloseOnly(IReadOnlyList<JsonObject>? ops) =>
        ops is { Count: 1 } &&
        string.Equals(Json.GetString(ops[0], "op"), "close_workbook", StringComparison.OrdinalIgnoreCase);

    private static void PreviewLifecycleOperation(object application, object? workbook, JsonObject op,
        ApplyPreview preview)
    {
        var name = Json.GetString(op, "op")!;
        switch (name)
        {
            case "create_workbook":
                preview.Affected.Add(new AffectedRef("workbook", "new"));
                preview.Diff.Add(new DiffEntry
                {
                    Ref = "workbook",
                    Before = null,
                    After = Json.GetString(op, "sheetName") ?? "Sheet1",
                });
                return;
            case "open_workbook":
            {
                var path = RequireAbsoluteOutput(Json.GetString(op, "path"), "open_workbook");
                if (!File.Exists(path))
                    preview.Errors.Add($"open_workbook path does not exist: {path}");
                preview.Affected.Add(new AffectedRef("workbook", path));
                preview.Diff.Add(new DiffEntry { Ref = "workbook", Before = null, After = path });
                return;
            }
            case "close_workbook":
                if (string.IsNullOrWhiteSpace(ExcelAuthoringSchema.ReadExplicitWorkbookTarget(op)))
                    preview.Errors.Add("close_workbook requires target.workbook (or targetWorkbook); active workbook close is not allowed");
                else if (workbook is null)
                    preview.Errors.Add("close_workbook requires an explicit open workbook target");
                else
                {
                    var identity = ReadWorkbookIdentity(workbook);
                    preview.Affected.Add(new AffectedRef("workbook", identity.FullName ?? identity.Name ?? "open"));
                    preview.Diff.Add(new DiffEntry { Ref = "workbook", Before = identity.FullName, After = null });
                }
                return;
            case "save_workbook":
                PreviewSaveOrExport(workbook, op, preview, pdf: false);
                return;
            case "export_pdf":
                PreviewSaveOrExport(workbook, op, preview, pdf: true);
                return;
        }
    }

    private static void ApplyLifecycleOperation(object application, object? workbook, JsonObject op,
        ApplyExecution execution, List<string> mismatches, ref int checkedItems)
    {
        var name = Json.GetString(op, "op")!;
        switch (name)
        {
            case "create_workbook":
                ApplyCreateWorkbook(application, op, execution, mismatches, ref checkedItems);
                return;
            case "open_workbook":
                ApplyOpenWorkbook(application, op, execution, mismatches, ref checkedItems);
                return;
            case "close_workbook":
                ApplyCloseWorkbook(application, workbook ?? throw MissingWorkbook(), op, execution, mismatches, ref checkedItems);
                return;
            case "save_workbook":
                ApplySaveWorkbook(workbook ?? throw MissingWorkbook(), op, execution, mismatches, ref checkedItems);
                return;
            case "export_pdf":
                ApplyExportPdf(workbook ?? throw MissingWorkbook(), op, execution, mismatches, ref checkedItems);
                return;
        }
    }

    private static JsonObject CaptureLifecycleState(object application, object? workbook,
        IReadOnlyList<JsonObject> ops, string? documentRef, string? snapshotDir = null)
    {
        var currentFullName = workbook is null ? null : ReadWorkbookIdentity(workbook).FullName;
        var outputs = new JsonArray();
        var openSources = new JsonArray();
        var index = 0;
        foreach (var op in ops)
        {
            var name = Json.GetString(op, "op");
            if (string.Equals(name, "open_workbook", StringComparison.OrdinalIgnoreCase))
            {
                var openPath = Json.GetString(op, "path");
                if (!string.IsNullOrWhiteSpace(openPath))
                {
                    string openFull;
                    try { openFull = Path.GetFullPath(openPath); } catch { openFull = openPath; }
                    openSources.Add(new JsonObject
                    {
                        ["op"] = name,
                        ["path"] = openFull,
                        ["role"] = ExcelAuthoringSchema.OpenSourceRole,
                        ["writable"] = false,
                    });
                }

                index++;
                continue;
            }

            var output = string.Equals(name, "save_workbook", StringComparison.OrdinalIgnoreCase)
                ? ExcelAuthoringSchema.ResolveSaveCapturePath(Json.GetString(op, "output"), currentFullName)
                : Json.GetString(op, "output");
            if (string.IsNullOrWhiteSpace(output))
            {
                index++;
                continue;
            }

            string full;
            try { full = Path.GetFullPath(output); } catch { full = output; }
            var existed = File.Exists(full);
            string? backupFile = null;
            long? length = null;
            if (existed && !ExcelAuthoringPaths.IsProtectedSource(full) && !string.IsNullOrWhiteSpace(snapshotDir))
            {
                backupFile = TryCopyOutputBackup(snapshotDir, index, full);
                try { length = new FileInfo(full).Length; } catch { }
            }

            outputs.Add(new JsonObject
            {
                ["op"] = name,
                ["path"] = full,
                ["existed"] = existed,
                ["backupFile"] = backupFile,
                ["length"] = length,
                ["writable"] = true,
                ["samePathSave"] = ExcelAuthoringSchema.IsSamePathSave(full, currentFullName),
            });
            index++;
        }

        var inventory = ListOpenWorkbooks(application);
        return new JsonObject
        {
            ["snapshotVersion"] = ExcelLayoutSnapshotVersion,
            ["restoreMode"] = LifecycleRestoreMode,
            ["closePolicy"] = ExcelAuthoringSchema.ClosePolicyOwnedOnly,
            ["documentRef"] = documentRef,
            ["existingWorkbooks"] = inventory,
            ["preexistingWorkbooks"] = inventory.DeepClone(),
            ["ownedWorkbooks"] = new JsonArray(),
            ["openSources"] = openSources,
            ["outputs"] = outputs,
            ["ops"] = CloneOps(ops),
        };
    }

    private static JsonObject RestoreLifecycleState(object application, JsonObject state)
    {
        var mismatches = new RestoreMismatchCollector();
        var owned = Json.GetArr(state, "ownedWorkbooks");
        var checkedItems = 0;
        foreach (var identity in EnumerateOpenWorkbooks(application))
        {
            try
            {
                if (!ExcelAuthoringSchema.ShouldCloseOnLifecycleRestore(identity.Name, identity.FullName, owned))
                    continue;
                CloseWorkbook(identity.Workbook, saveChanges: false);
                checkedItems++;
            }
            catch (Exception ex) { mismatches.Add($"lifecycle restore close failed: {ex.Message}"); }
            finally { RotHelper.ReleaseComReference(identity.Workbook); }
        }

        foreach (var node in Json.GetArr(state, "outputs") ?? new JsonArray())
        {
            if (node is not JsonObject output) continue;
            if (!ExcelAuthoringSchema.IsWritableLifecycleOutput(output))
                continue;
            var path = Json.GetString(output, "path");
            if (string.IsNullOrWhiteSpace(path) || ExcelAuthoringPaths.IsProtectedSource(path))
                continue;
            var backup = Json.GetString(output, "backupFile");
            try
            {
                if (Json.GetBool(output, "existed"))
                {
                    if (string.IsNullOrWhiteSpace(backup) || !File.Exists(backup))
                    {
                        mismatches.Add($"lifecycle restore missing overwrite backup for '{path}'");
                        continue;
                    }

                    File.Copy(backup, path, overwrite: true);
                    checkedItems++;
                }
                else if (File.Exists(path))
                {
                    File.Delete(path);
                    checkedItems++;
                }
            }
            catch (Exception ex) { mismatches.Add($"lifecycle restore could not restore '{path}': {ex.Message}"); }

            if (Json.GetBool(output, "samePathSave") &&
                FindOpenWorkbookByPath(application, path) is { } stillOpen)
            {
                RotHelper.ReleaseComReference(stillOpen);
                mismatches.Add(
                    $"lifecycle restore of same-path save left workbook identity/content incomplete for '{path}'");
            }
        }

        var preexisting = (Json.GetArr(state, "preexistingWorkbooks") ?? Json.GetArr(state, "existingWorkbooks") ?? new JsonArray())
            .OfType<JsonObject>()
            .ToList();
        var after = ListOpenWorkbooks(application).OfType<JsonObject>().ToList();
        var lost = ExcelLifecycleRestoreContract.LostPreexisting(preexisting, after, owned);
        foreach (var name in lost)
            mismatches.Add($"lifecycle restore lost preexisting workbook '{name}'; not waived and not recreated");
        checkedItems += ExcelLifecycleRestoreContract.VerifiedPreexistingCount(preexisting, after, owned);
        var completeness = ExcelLifecycleRestoreContract.Completeness(checkedItems, mismatches.Count);
        var verified = ExcelLifecycleRestoreContract.IsVerified(checkedItems, mismatches.Count);
        if (completeness == ExcelLifecycleRestoreContract.CompletenessUnproven)
            mismatches.Add(ExcelLifecycleRestoreContract.UnprovenMessage);

        var result = BuildRestoreResult(verified, 0, checkedItems, LifecycleRestoreMode, mismatches);
        result["restoreCompleteness"] = completeness;
        result["verified"] = verified;
        return result;
    }

    private static void PreviewSaveOrExport(object? workbook, JsonObject op, ApplyPreview preview, bool pdf)
    {
        if (workbook is null)
        {
            preview.Errors.Add($"{Json.GetString(op, "op")} requires an open workbook");
            return;
        }

        var output = Json.GetString(op, "output");
        if (pdf && string.IsNullOrWhiteSpace(output))
        {
            preview.Errors.Add("export_pdf requires output");
            return;
        }

        string current;
        try { current = Convert.ToString(((dynamic)workbook).FullName, CultureInfo.InvariantCulture) ?? ""; }
        catch { current = ""; }

        var target = string.IsNullOrWhiteSpace(output) ? current : RequireAbsoluteOutput(output, Json.GetString(op, "op")!);
        if (ExcelAuthoringPaths.IsProtectedSource(target))
        {
            preview.Errors.Add($"{Json.GetString(op, "op")} refuses to write the protected source workbook");
            return;
        }

        var exists = !string.IsNullOrWhiteSpace(output) && File.Exists(target);
        if (ExcelAuthoringSchema.SaveConflictsWithExistingFile(string.IsNullOrWhiteSpace(output) ? null : target,
                current, Json.GetBool(op, "overwrite"), exists))
            preview.Errors.Add($"{Json.GetString(op, "op")} '{target}' exists; set overwrite:true to replace it");

        preview.Affected.Add(new AffectedRef(pdf ? "pdf" : "workbook", target));
        preview.Diff.Add(new DiffEntry { Ref = pdf ? "pdf" : "file", Before = current, After = target });
    }

    private static void ApplyCreateWorkbook(object application, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        object? books = null;
        object? created = null;
        try
        {
            books = (object)((dynamic)application).Workbooks;
            created = (object)((dynamic)books).Add();
            RecordOwnedWorkbookIdentity(created, "Workbooks.Add");
            var sheetName = Json.GetString(op, "sheetName");
            if (!string.IsNullOrWhiteSpace(sheetName))
            {
                var errors = new List<string>();
                if (!ExcelSheetNameContract.TryNormalize(sheetName, 1, "create_workbook.sheetName", errors, out var normalized))
                    throw new InvalidOperationException(string.Join("; ", errors));
                object? sheet = null;
                try
                {
                    sheet = (object)((dynamic)created).Worksheets.Item(1);
                    ((dynamic)sheet).Name = normalized;
                }
                finally { RotHelper.ReleaseComReference(sheet); }
            }

            var name = Convert.ToString(((dynamic)created).Name, CultureInfo.InvariantCulture);
            checkedItems++;
            execution.Affected.Add(new AffectedRef("workbook", name ?? "new"));
        }
        finally
        {
            RotHelper.ReleaseComReference(created);
            RotHelper.ReleaseComReference(books);
        }
    }

    private static void ApplyOpenWorkbook(object application, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        var path = RequireAbsoluteOutput(Json.GetString(op, "path"), "open_workbook");
        if (!File.Exists(path))
            throw new InvalidOperationException($"open_workbook path does not exist: {path}");
        var before = CapturePassiveWorkbookInventory(application);
        var ownerPid = ReadApplicationProcessId(application);
        if (FindOpenWorkbookByPath(application, path) is { } already)
        {
            RotHelper.ReleaseComReference(already);
            checkedItems++;
            execution.Affected.Add(new AffectedRef("workbook", path));
            execution.Warnings.Add("open_workbook: workbook was already open");
            if (!WorkbookStillOpenAcrossInstances(application, path))
                mismatches.Add($"open_workbook readback: '{path}' not found across running Excel instances");
            ReportLostBystanders(application, before, path, ownerPid, mismatches);
            return;
        }

        object? books = null;
        object? opened = null;
        try
        {
            books = (object)((dynamic)application).Workbooks;
            opened = (object)((dynamic)books).Open(path, 0, false);
            RecordOwnedWorkbookIdentity(opened, "Workbooks.Open");
            var full = Convert.ToString(((dynamic)opened).FullName, CultureInfo.InvariantCulture);
            checkedItems++;
            if (!string.Equals(full, path, StringComparison.OrdinalIgnoreCase))
                mismatches.Add("open_workbook readback path mismatch");
            if (!WorkbookStillOpenAcrossInstances(application, path))
                mismatches.Add($"open_workbook readback: '{path}' not found across running Excel instances");
            execution.Affected.Add(new AffectedRef("workbook", path));
            ReportLostBystanders(application, before, path, ownerPid, mismatches);
        }
        finally
        {
            RotHelper.ReleaseComReference(opened);
            RotHelper.ReleaseComReference(books);
        }
    }

    private static void ApplyCloseWorkbook(object application, object workbook, JsonObject op,
        ApplyExecution execution, List<string> mismatches, ref int checkedItems)
    {
        if (string.IsNullOrWhiteSpace(ExcelAuthoringSchema.ReadExplicitWorkbookTarget(op)))
            throw new InvalidOperationException(
                "close_workbook requires target.workbook (or targetWorkbook); active workbook close is not allowed");
        var identity = ReadWorkbookIdentity(workbook);
        var before = CapturePassiveWorkbookInventory(application);
        var ownerPid = ReadApplicationProcessId(application);
        var saveChanges = Json.GetBool(op, "saveChanges");
        CloseWorkbook(workbook, saveChanges);
        checkedItems++;
        if (WorkbookStillOpenAcrossInstances(application, identity.FullName ?? identity.Name ?? ""))
            mismatches.Add("close_workbook readback: workbook is still open");
        ReportLostBystanders(application, before, identity.FullName ?? identity.Name ?? "", ownerPid, mismatches);
        execution.Affected.Add(new AffectedRef("workbook", identity.FullName ?? identity.Name ?? "closed"));
    }

    private static void ApplySaveWorkbook(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        var output = Json.GetString(op, "output");
        string current;
        try { current = Convert.ToString(((dynamic)workbook).FullName, CultureInfo.InvariantCulture) ?? ""; }
        catch { current = ""; }

        if (string.IsNullOrWhiteSpace(output) && ExcelAuthoringPaths.IsLikelyUnsavedWorkbookName(current))
            throw new InvalidOperationException("save_workbook without output cannot save an untitled workbook; provide output");

        var path = string.IsNullOrWhiteSpace(output) ? current : RequireAbsoluteOutput(output, "save_workbook");
        if (ExcelAuthoringPaths.IsProtectedSource(path))
            throw new InvalidOperationException("save_workbook refuses to write the protected source workbook");

        if (string.IsNullOrWhiteSpace(output) || ExcelAuthoringSchema.IsSamePathSave(path, current))
        {
            ((dynamic)workbook).Save();
            checkedItems++;
            var saved = false;
            try { saved = Convert.ToBoolean(((dynamic)workbook).Saved, CultureInfo.InvariantCulture); }
            catch { saved = false; }
            if (!saved)
                mismatches.Add("save_workbook Saved is still false after Save");
            if (!File.Exists(path))
                mismatches.Add($"save_workbook did not persist '{path}'");
            execution.Affected.Add(new AffectedRef("workbook", path));
            return;
        }

        if (ExcelAuthoringSchema.SaveConflictsWithExistingFile(path, current, Json.GetBool(op, "overwrite"), File.Exists(path)))
            throw new InvalidOperationException($"save_workbook '{path}' exists; set overwrite:true to replace it");
        Directory.CreateDirectory(Path.GetDirectoryName(path) ?? ".");
        ((dynamic)workbook).SaveAs(path);
        checkedItems++;
        var actual = Convert.ToString(((dynamic)workbook).FullName, CultureInfo.InvariantCulture);
        if (!string.Equals(actual, path, StringComparison.OrdinalIgnoreCase))
            mismatches.Add("save_workbook readback path mismatch");
        if (!File.Exists(path))
            mismatches.Add($"save_workbook did not create '{path}'");
        execution.Affected.Add(new AffectedRef("workbook", path));
    }

    private static void ApplyExportPdf(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        var path = RequireAbsoluteOutput(Json.GetString(op, "output"), "export_pdf");
        if (ExcelAuthoringPaths.IsProtectedSource(path))
            throw new InvalidOperationException("export_pdf refuses to write the protected source workbook");
        if (ExcelAuthoringSchema.SaveConflictsWithExistingFile(path, currentFullName: null, Json.GetBool(op, "overwrite"), File.Exists(path)))
            throw new InvalidOperationException($"export_pdf '{path}' exists; set overwrite:true to replace it");
        Directory.CreateDirectory(Path.GetDirectoryName(path) ?? ".");
        object? sheet = null;
        try
        {
            var sheetName = Json.GetString(op, "sheet") ?? Json.GetString(Json.GetObj(op, "target"), "sheet");
            if (!string.IsNullOrWhiteSpace(sheetName))
            {
                sheet = GetExplicitTargetSheetReference(workbook, SheetTargetOp(sheetName));
                ((dynamic)sheet).ExportAsFixedFormat(XlTypePdf, path);
            }
            else
            {
                ((dynamic)workbook).ExportAsFixedFormat(XlTypePdf, path);
            }

            checkedItems++;
            if (!File.Exists(path))
                mismatches.Add($"export_pdf did not create '{path}'");
            execution.Affected.Add(new AffectedRef("pdf", path));
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static string RequireAbsoluteOutput(string? path, string opName)
    {
        if (string.IsNullOrWhiteSpace(path) || !Path.IsPathRooted(path))
            throw new InvalidOperationException($"{opName} requires an absolute path");
        return Path.GetFullPath(path);
    }

    private static InvalidOperationException MissingWorkbook() =>
        new("lifecycle save/export/close requires a connected workbook");

    private static List<JsonObject> CapturePassiveWorkbookInventory(object application) =>
        ListOpenWorkbooks(application).OfType<JsonObject>().ToList();

    private static void ReportLostBystanders(
        object application, IReadOnlyList<JsonObject> before, string openedPath, int ownerPid,
        List<string> mismatches)
    {
        var after = CapturePassiveWorkbookInventory(application);
        var lost = ExcelOpenWorkbookInventoryContract.LostBystanders(before, after, openedPath);
        if (!ExcelOpenWorkbookInventoryContract.PreserveAggregateFailure(lost))
            return;
        var instancePids = before.Concat(after)
            .Select(book => Json.GetInt(book, "processId") ?? 0)
            .Where(pid => pid > 0)
            .Distinct()
            .ToList();
        mismatches.Add(ExcelOpenWorkbookInventoryContract.FailureMessage(lost, ownerPid, instancePids));
    }

    private static JsonArray ListWorkbookIdentities(object application)
    {
        var result = new JsonArray();
        foreach (var item in EnumerateOpenWorkbooks(application))
        {
            try
            {
                result.Add(new JsonObject
                {
                    ["name"] = item.Name,
                    ["fullName"] = item.FullName,
                });
            }
            finally { RotHelper.ReleaseComReference(item.Workbook); }
        }

        return result;
    }

    private static (string Name, string FullName) ReadWorkbookIdentity(object workbook)
    {
        try
        {
            return (
                Convert.ToString(((dynamic)workbook).Name, CultureInfo.InvariantCulture) ?? "",
                Convert.ToString(((dynamic)workbook).FullName, CultureInfo.InvariantCulture) ?? "");
        }
        catch
        {
            return ("", "");
        }
    }

    private static List<(object Workbook, string Name, string FullName)> EnumerateOpenWorkbooks(object application)
    {
        var result = new List<(object, string, string)>();
        object? books = null;
        try
        {
            books = (object)((dynamic)application).Workbooks;
            var count = Convert.ToInt32(((dynamic)books).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                object? workbook = null;
                try
                {
                    workbook = (object)((dynamic)books).Item(index);
                    result.Add((
                        workbook,
                        Convert.ToString(((dynamic)workbook).Name, CultureInfo.InvariantCulture) ?? "",
                        Convert.ToString(((dynamic)workbook).FullName, CultureInfo.InvariantCulture) ?? ""));
                    workbook = null;
                }
                finally { RotHelper.ReleaseComReference(workbook); }
            }
        }
        finally { RotHelper.ReleaseComReference(books); }
        return result;
    }

    private static object? FindOpenWorkbookByPath(object application, string path)
    {
        if (string.IsNullOrWhiteSpace(path)) return null;
        foreach (var item in EnumerateOpenWorkbooks(application))
        {
            if (string.Equals(item.FullName, path, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(item.Name, path, StringComparison.OrdinalIgnoreCase))
                return item.Workbook;
            RotHelper.ReleaseComReference(item.Workbook);
        }

        return null;
    }

    private static bool WorkbookStillOpenAcrossInstances(object attachedApplication, string reference)
    {
        if (string.IsNullOrWhiteSpace(reference))
            return false;
        try
        {
            using var found = FindOpenWorkbook(attachedApplication, reference, allowFileOpenFallback: false);
            return found.Workbook is not null;
        }
        catch
        {
            return false;
        }
    }

    private static void CloseWorkbook(object workbook, bool saveChanges)
    {
        ((dynamic)workbook).Close(saveChanges);
    }

    private static string? TryCopyOutputBackup(string snapshotDir, int index, string source)
    {
        try
        {
            var dest = Path.Combine(snapshotDir, $"output-backup-{index}{Path.GetExtension(source)}");
            if (File.Exists(dest)) return dest;
            using var input = new FileStream(source, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var output = new FileStream(dest, FileMode.CreateNew, FileAccess.Write);
            input.CopyTo(output);
            return dest;
        }
        catch
        {
            return null;
        }
    }

    private static void RecordOwnedWorkbookIdentity(object workbook, string source)
    {
        var identity = ReadWorkbookIdentity(workbook);
        PatchSnapshotState(_lastSnapshotDir, state =>
            ExcelAuthoringSchema.AppendOwnedWorkbook(state, identity.Name, identity.FullName, source));
    }

    internal static void PatchSnapshotState(string? snapshotDir, Action<JsonObject> mutate)
    {
        if (string.IsNullOrWhiteSpace(snapshotDir)) return;
        var path = Path.Combine(snapshotDir, "state.json");
        if (!File.Exists(path)) return;
        if (JsonNode.Parse(File.ReadAllText(path)) is not JsonObject state) return;
        mutate(state);
        File.WriteAllText(path, state.ToJsonString(Json.Pretty));
    }

    internal JsonObject LaunchOwnedExcelInstance() => LaunchExcelInstance(null);

    internal JsonObject LaunchExcelInstance(JsonObject? args)
    {
        var argErrors = new List<string>();
        if (!ExcelInstancePin.TryValidateSelectors(args, argErrors))
            return Json.ErrorResult(string.Join("; ", argErrors), App);
        return ComInvoke(() =>
        {
            var dedicated = Json.GetBool(args, "dedicatedInstance");
            if (Json.GetBool(args, "clearPin"))
                ExcelInstancePin.ClearPin();
            object? app;
            var pinned = false;
            var pinnedPid = 0;
            long pinnedHwnd = 0;
            if (dedicated)
            {
                app = AttachDedicatedExcel();
            }
            else if (HasLaunchSelector(args))
            {
                app = FindSelectedLaunchInstance(args);
                dynamic selected = app;
                try { selected.Visible = true; } catch { }
                pinnedHwnd = Convert.ToInt64(selected.Hwnd, CultureInfo.InvariantCulture);
                if (pinnedHwnd == 0)
                {
                    RotHelper.ReleaseDiscoveredApplication(app, null);
                    throw new InvalidOperationException(
                        "[EXCEL_PIN_NO_WINDOW_HANDLE] the selected Excel has no window handle " +
                        "(minimized?); restore its window first, then pin again");
                }
                pinnedPid = RotHelper.ProcessIdFromWindowHandle(pinnedHwnd);
                if (pinnedPid < 1)
                {
                    RotHelper.ReleaseDiscoveredApplication(app, null);
                    throw new InvalidOperationException(
                        "[EXCEL_PIN_NO_PROCESS] the selected Excel window has no owning process");
                }
                if (_attached is not null)
                    DetachExcelReference();
                if (!ExcelInstancePin.TrySavePin(pinnedPid, pinnedHwnd, "excel_launch", out var saveError))
                {
                    RotHelper.ReleaseDiscoveredApplication(app, null);
                    throw new InvalidOperationException($"[EXCEL_PIN_SAVE_FAILED] {saveError}");
                }
                _attached = app;
                _ownsInstance = false;
                _ownedProcessId = pinnedPid;
                pinned = true;
            }
            else
            {
                app = AttachExcel(allowCreate: true);
                var pin = ExcelInstancePin.TryLoadPin();
                if (app is not null && pin is not null && AttachedMatchesPin(app, pin))
                {
                    pinned = true;
                    pinnedPid = pin.ProcessId;
                    pinnedHwnd = pin.Hwnd;
                }
            }
            if (app is null)
                return Json.ErrorResult(
                    dedicated
                        ? "excel_launch dedicatedInstance could not create an Excel.Application"
                        : "excel_launch could not create or attach an Excel instance",
                    App);
            dynamic d = app;
            try { d.Visible = true; } catch { }
            object? books = null;
            object? created = null;
            try
            {
                books = (object)d.Workbooks;
                var count = Convert.ToInt32(((dynamic)books).Count, CultureInfo.InvariantCulture);
                var createdWorkbook = false;
                if (count == 0)
                {
                    created = (object)((dynamic)books).Add();
                    createdWorkbook = true;
                }

                var identity = created is not null
                    ? ReadWorkbookIdentity(created)
                    : EnumerateOpenWorkbooks(app).Select(item =>
                    {
                        RotHelper.ReleaseComReference(item.Workbook);
                        return (item.Name, item.FullName);
                    }).FirstOrDefault();
                var flags = ExcelAuthoringSchema.DescribeLaunchOwnership(_ownsInstance, createdWorkbook, dedicated);
                flags["ok"] = true;
                flags["app"] = App;
                flags["workbook"] = identity.Name;
                flags["documentRef"] = string.IsNullOrWhiteSpace(identity.FullName) ? identity.Name : identity.FullName;
                flags["pinned"] = pinned;
                if (pinned)
                {
                    flags["pinnedProcessId"] = pinnedPid;
                    flags["pinnedHwnd"] = pinnedHwnd;
                }
                return flags;
            }
            finally
            {
                RotHelper.ReleaseComReference(created);
                RotHelper.ReleaseComReference(books);
            }
        });
    }

    private static bool HasLaunchSelector(JsonObject? args)
    {
        if (args is null) return false;
        return args.ContainsKey("processId") || args.ContainsKey("hwnd") || Json.GetBool(args, "activeWindow");
    }

    /// <summary>
    /// Resolve one ROT Excel application by explicit selector. Throws with an
    /// actionable code when nothing matches; never falls back silently.
    /// </summary>
    private static object FindSelectedLaunchInstance(JsonObject? args)
    {
        long? wantHwnd = null;
        int? wantPid = null;
        if (Json.GetBool(args, "activeWindow"))
        {
            var foreground = RotHelper.ForegroundWindowHandle();
            if (foreground == 0)
                throw new InvalidOperationException(
                    "[EXCEL_NO_FOREGROUND_WINDOW] no foreground window is available; " +
                    "focus the Excel window first or pin by processId");
            wantHwnd = foreground;
        }
        if (args?.ContainsKey("processId") == true &&
            ExcelDataOperationsContract.TryGetFiniteNumber(args["processId"], out var pid))
            wantPid = (int)pid;
        if (!wantHwnd.HasValue && args?.ContainsKey("hwnd") == true &&
            ExcelDataOperationsContract.TryGetFiniteNumber(args["hwnd"], out var hwnd))
            wantHwnd = (long)hwnd;
        foreach (var candidate in RotHelper.GetExcelApplications())
        {
            var keep = false;
            try
            {
                var candidateHwnd = Convert.ToInt64(((dynamic)candidate).Hwnd, CultureInfo.InvariantCulture);
                if (wantHwnd.HasValue && candidateHwnd != wantHwnd.Value) continue;
                if (wantPid.HasValue &&
                    RotHelper.ProcessIdFromWindowHandle(candidateHwnd) != wantPid.Value) continue;
                if (!RotHelper.IsWindowAlive(candidateHwnd)) continue;
                keep = true;
                return candidate;
            }
            catch
            {
                // Unreadable candidates never match a selector.
            }
            finally
            {
                if (!keep) RotHelper.ReleaseDiscoveredApplication(candidate, null);
            }
        }
        throw new InvalidOperationException(
            "[EXCEL_PIN_TARGET_MISSING] no running Excel matches the requested selector; " +
            "verify the window is open (not minimized) and the processId/hwnd is current");
    }
}
