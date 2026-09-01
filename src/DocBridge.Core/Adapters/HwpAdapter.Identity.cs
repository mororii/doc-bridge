using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class HwpAdapter
{
    private sealed record OpenHwpMatch(object Application, object Document, string FullName,
        string DocumentId, long WindowHandle, int ProcessId);

    private static string HwpDocumentRef(
        string fullName, string documentId, long windowHandle, int processId)
    {
        if (!string.IsNullOrWhiteSpace(fullName))
        {
            try { return CanonicalHwpPath(fullName); }
            catch { return fullName; }
        }
        // MainWindowHandle can change while HWP merely activates another tab.
        // Keep it as observation metadata, never as part of a persistent document key.
        return $"untitled-{processId}-{documentId}";
    }

    private static string HwpInstanceRef(string documentId, long windowHandle, int processId) =>
        $"hwp:{processId}:{documentId}";

    private static bool HwpDocumentRefMatches(
        string requested, string fullName, string documentId, long windowHandle, int processId)
    {
        if (string.Equals(requested, HwpDocumentRef(fullName, documentId, windowHandle, processId),
                StringComparison.OrdinalIgnoreCase) ||
            string.Equals(requested, HwpInstanceRef(documentId, windowHandle, processId),
                StringComparison.OrdinalIgnoreCase))
            return true;

        // Accept the short-lived development format that contained a window handle.
        // Compare only the stable process/document portions because the handle can change.
        if ((requested.StartsWith($"hwp:{processId}:", StringComparison.OrdinalIgnoreCase) &&
             requested.EndsWith($":{documentId}", StringComparison.OrdinalIgnoreCase)) ||
            (requested.StartsWith($"untitled-{processId}-", StringComparison.OrdinalIgnoreCase) &&
             requested.EndsWith($"-{documentId}", StringComparison.OrdinalIgnoreCase)))
            return true;

        // 0.4.8 이전의 미저장 문서 ref도 유일한 경우에는 계속 인식한다.
        if (string.IsNullOrWhiteSpace(fullName) &&
            string.Equals(requested, $"untitled-{documentId}", StringComparison.OrdinalIgnoreCase))
            return true;

        if (string.IsNullOrWhiteSpace(fullName)) return false;
        try
        {
            return Path.IsPathFullyQualified(requested) &&
                   string.Equals(CanonicalHwpPath(requested), CanonicalHwpPath(fullName),
                       StringComparison.OrdinalIgnoreCase);
        }
        catch { return false; }
    }

    /// <summary>
    /// HWP can republish the same application under a new ROT moniker when a document tab is
    /// activated. Reacquire the dispatch whose active document is the requested one instead of
    /// retaining the pre-activation RCW, which can report an empty XHwpDocuments collection.
    /// </summary>
    private object? ReacquireActivatedHwpApplication(OpenHwpMatch selected)
    {
        if (_appFactory is not null) return selected.Application;

        object? fallback = null;
        foreach (var application in RotHelper.GetHwpApplications())
        {
            if (!RotHelper.HwpWindowVisible(application))
            {
                RotHelper.ReleaseComObject(application);
                continue;
            }

            var hwnd = RotHelper.HwpWindowHandle(application);
            var processId = RotHelper.ProcessIdFromWindowHandle(hwnd);
            if (selected.ProcessId != 0 && processId != selected.ProcessId)
            {
                RotHelper.ReleaseComObject(application);
                continue;
            }

            var containsTarget = false;
            var targetIsActive = false;
            try
            {
                dynamic hwp = application;
                dynamic documents = hwp.XHwpDocuments;
                object? activeDocument = null;
                try
                {
                    activeDocument = (object)documents.Active_XHwpDocument;
                    dynamic active = activeDocument;
                    targetIsActive = string.Equals(
                        Convert.ToString(active.DocumentID) ?? "", selected.DocumentId,
                        StringComparison.Ordinal);
                }
                catch { }
                finally { RotHelper.ReleaseComObject(activeDocument); }

                var count = Convert.ToInt32(documents.Count);
                for (var index = 0; index < count && !containsTarget; index++)
                {
                    object? document = null;
                    try
                    {
                        document = (object)documents.Item(index);
                        dynamic d = document;
                        containsTarget = string.Equals(
                            Convert.ToString(d.DocumentID) ?? "", selected.DocumentId,
                            StringComparison.Ordinal);
                    }
                    catch { }
                    finally { RotHelper.ReleaseComObject(document); }
                }
            }
            catch { }

            if (!containsTarget)
            {
                RotHelper.ReleaseComObject(application);
                continue;
            }

            if (targetIsActive)
            {
                if (fallback is not null) RotHelper.ReleaseComObject(fallback);
                return application;
            }

            if (fallback is null) fallback = application;
            else RotHelper.ReleaseComObject(application);
        }
        return fallback;
    }

    /// <summary>
    /// 모든 표시 HWP ROT 인스턴스와 그 안의 모든 탭을 가벼운 메타데이터로 열거한다.
    /// 본문은 연결된 활성 문서에서만 별도로 읽어 다중 문서 확인 비용을 제한한다.
    /// </summary>
    private JsonArray InspectOpenHwpDocuments(long connectedWindowHandle)
    {
        var applications = _appFactory is not null && _attached is not null
            ? new List<object> { _attached }
            : RotHelper.GetHwpApplications().ToList();
        var entries = new List<JsonObject>();
        try
        {
            for (var applicationIndex = 0; applicationIndex < applications.Count; applicationIndex++)
            {
                var application = applications[applicationIndex];
                if (_appFactory is null && !RotHelper.HwpWindowVisible(application)) continue;
                var windowHandle = RotHelper.HwpWindowHandle(application);
                var processId = RotHelper.ProcessIdFromWindowHandle(windowHandle);
                try
                {
                    dynamic hwp = application;
                    dynamic documents = hwp.XHwpDocuments;
                    var count = Convert.ToInt32(documents.Count);
                    var activeDocumentId = "";
                    object? activeDocument = null;
                    try
                    {
                        activeDocument = (object)documents.Active_XHwpDocument;
                        dynamic active = activeDocument;
                        activeDocumentId = Convert.ToString(active.DocumentID) ?? "";
                    }
                    catch { }
                    finally { RotHelper.ReleaseComObject(activeDocument); }

                    for (var documentIndex = 0; documentIndex < count; documentIndex++)
                    {
                        object? document = null;
                        try
                        {
                            document = (object)documents.Item(documentIndex);
                            dynamic d = document;
                            var fullName = Convert.ToString(d.FullName) ?? "";
                            var documentId = Convert.ToString(d.DocumentID) ?? "";
                            var activeInWindow = string.Equals(
                                documentId, activeDocumentId, StringComparison.Ordinal);
                            var item = new JsonObject
                            {
                                ["documentRef"] = HwpDocumentRef(
                                    fullName, documentId, windowHandle, processId),
                                ["instanceRef"] = HwpInstanceRef(
                                    documentId, windowHandle, processId),
                                ["documentId"] = documentId,
                                ["fullName"] = fullName,
                                ["name"] = string.IsNullOrWhiteSpace(fullName)
                                    ? $"빈 문서 {documentId}"
                                    : Path.GetFileName(fullName),
                                ["saved"] = !string.IsNullOrWhiteSpace(fullName),
                                ["windowHandle"] = windowHandle.ToString(),
                                ["processId"] = processId,
                                ["windowIndex"] = applicationIndex,
                                ["documentIndex"] = documentIndex,
                                ["activeInWindow"] = activeInWindow,
                                ["active"] = windowHandle == connectedWindowHandle && activeInWindow,
                            };
                            try { item["modified"] = Convert.ToBoolean(d.Modified); } catch { }
                            try { item["format"] = Convert.ToString(d.Format); } catch { }
                            try { item["editMode"] = Convert.ToInt32(d.EditMode); } catch { }
                            entries.Add(item);
                        }
                        catch
                        {
                            // 닫히는 탭은 건너뛰고 나머지 문서를 계속 열거한다.
                        }
                        finally { RotHelper.ReleaseComObject(document); }
                    }
                }
                catch
                {
                    // 모달 또는 종료 중인 창은 다른 창의 목록을 막지 않는다.
                }
            }

            // Some HWP builds publish more than one ROT moniker for the same document window.
            // Collapse those aliases before reporting counts or deciding that a path is duplicated.
            entries = entries
                .GroupBy(entry => Json.GetString(entry, "instanceRef") ?? "", StringComparer.OrdinalIgnoreCase)
                .Select(group => group
                    .OrderByDescending(entry => Json.GetBool(entry, "active"))
                    .ThenBy(entry => Json.GetInt(entry, "windowIndex"))
                    .First())
                .ToList();

            var duplicatePaths = entries
                .Where(entry => Json.GetBool(entry, "saved"))
                .GroupBy(entry =>
                {
                    var fullName = Json.GetString(entry, "fullName") ?? "";
                    try { return CanonicalHwpPath(fullName); }
                    catch { return fullName; }
                }, StringComparer.OrdinalIgnoreCase)
                .Where(group => group.Count() > 1)
                .SelectMany(group => group)
                .ToHashSet();
            foreach (var entry in entries)
                entry["duplicatePath"] = duplicatePaths.Contains(entry);

            return new JsonArray(entries.Select(entry => (JsonNode?)entry).ToArray());
        }
        finally
        {
            foreach (var application in applications)
            {
                if (!ReferenceEquals(application, _attached))
                    RotHelper.ReleaseComObject(application);
            }
        }
    }

    /// <summary>
    /// file이 지정되면 모든 표시 중인 한글 ROT 인스턴스와 모든 탭을 먼저 조사한다.
    /// 정확히 한 문서만 일치할 때 그 문서를 활성화하고, 중복이면 임의 선택하지 않는다.
    /// </summary>
    private object? AttachHwpForTarget(
        string? file, string? documentRef, bool allowCreate,
        ForegroundInteractionGuard? foreground = null,
        HwpInteractionState? documentState = null)
    {
        if (!string.IsNullOrWhiteSpace(file) && !string.IsNullOrWhiteSpace(documentRef))
            throw new HwpAutomationException(
                "HWP_TARGET_CONFLICT",
                "file과 documentRef를 동시에 지정할 수 없습니다.",
                "hwp_get_active_context의 openDocuments에서 하나의 documentRef를 고르거나 절대 file 경로 하나만 사용하세요.");

        if (string.IsNullOrWhiteSpace(file) && string.IsNullOrWhiteSpace(documentRef))
        {
            var activeApplication = AttachHwp(allowCreate);
            if (activeApplication is not null)
                TrackHwpInteraction(activeApplication, foreground, documentState, captureTarget: true);
            return activeApplication;
        }

        var target = string.IsNullOrWhiteSpace(file) ? null : CanonicalHwpPath(file);
        var applications = new List<object>();
        if (_appFactory is not null)
        {
            var injected = AttachHwp(allowCreate);
            if (injected is not null) applications.Add(injected);
        }
        else
        {
            applications.AddRange(RotHelper.GetHwpApplications());
        }
        var matches = new List<OpenHwpMatch>();
        try
        {
            foreach (var application in applications)
            {
                if (_appFactory is null && !RotHelper.HwpWindowVisible(application)) continue;
                try
                {
                    dynamic hwp = application;
                    dynamic documents = hwp.XHwpDocuments;
                    var count = Convert.ToInt32(documents.Count);
                    for (var index = 0; index < count; index++)
                    {
                        object? document = null;
                        try
                        {
                            document = (object)documents.Item(index);
                            dynamic d = document;
                            var fullName = Convert.ToString(d.FullName) ?? "";
                            var documentId = "";
                            try { documentId = Convert.ToString(d.DocumentID) ?? ""; } catch { }
                            var hwnd = RotHelper.HwpWindowHandle(application);
                            var processId = RotHelper.ProcessIdFromWindowHandle(hwnd);
                            var matched = target is not null
                                ? !string.IsNullOrWhiteSpace(fullName) &&
                                  string.Equals(CanonicalHwpPath(fullName), target,
                                      StringComparison.OrdinalIgnoreCase)
                                : HwpDocumentRefMatches(
                                    documentRef!, fullName, documentId, hwnd, processId);
                            if (!matched)
                            {
                                RotHelper.ReleaseComObject(document);
                                continue;
                            }
                            matches.Add(new OpenHwpMatch(application, document, fullName, documentId,
                                hwnd, processId));
                        }
                        catch
                        {
                            RotHelper.ReleaseComObject(document);
                        }
                    }
                }
                catch
                {
                    // 종료 중이거나 모달 상태인 인스턴스는 다른 인스턴스를 계속 조사한다.
                }
            }

            // A single document can appear through multiple ROT aliases. Keep one physical
            // process/document pair so exact targeting does not become falsely ambiguous.
            var duplicateMatches = matches
                .GroupBy(match => HwpInstanceRef(match.DocumentId, match.WindowHandle, match.ProcessId),
                    StringComparer.OrdinalIgnoreCase)
                .SelectMany(group => group.Skip(1))
                .ToList();
            foreach (var duplicate in duplicateMatches)
            {
                matches.Remove(duplicate);
                RotHelper.ReleaseComObject(duplicate.Document);
            }

            if (matches.Count > 1)
            {
                var locations = string.Join(", ", matches.Select(match =>
                    $"PID {match.ProcessId}/문서 {match.DocumentId}"));
                if (target is not null)
                    throw new HwpAutomationException(
                        "HWP_DUPLICATE_LOCAL_PATH",
                        $"동일한 한글 파일이 {matches.Count}개 창 또는 탭에 열려 있어 임의 편집을 거부했습니다: {target} ({locations})",
                        "중복 창을 닫거나 openDocuments의 instanceRef로 정확한 탭을 지정하세요.");
                throw new HwpAutomationException(
                    "HWP_AMBIGUOUS_DOCUMENT_REF",
                    $"documentRef '{documentRef}'가 {matches.Count}개 한글 문서와 일치해 임의 편집을 거부했습니다: {locations}",
                    "hwp_get_active_context의 openDocuments에서 고유한 instanceRef를 사용하세요.");
            }

            if (matches.Count == 1)
            {
                var selected = matches[0];
                if (selected.WindowHandle != 0) foreground?.TrackTargetWindow(selected.WindowHandle);
                else foreground?.TrackTargetProcess(selected.ProcessId);
                documentState?.CaptureOriginal(selected.Application);
                try { ((dynamic)selected.Document).SetActive_XHwpDocument(); }
                catch (Exception ex)
                {
                    throw new HwpAutomationException(
                        "HWP_DOCUMENT_ACTIVATION_FAILED",
                        $"일치하는 한글 문서를 찾았지만 활성화하지 못했습니다: {target ?? documentRef}",
                        "한글의 모달 대화상자를 닫고 다시 시도하세요.", ex);
                }
                // 내부 탭 활성화가 Windows 전경 창까지 가져오는 빌드가 있으므로
                // 긴 후속 COM 호출 전에 즉시 사용자의 원래 창을 복구한다.
                foreground?.Checkpoint(stopOnConcurrentInput: false);

                var selectedApplication = ReacquireActivatedHwpApplication(selected)
                    ?? selected.Application;
                if (_attached is not null && !ReferenceEquals(_attached, selectedApplication))
                    RotHelper.ReleaseComObject(_attached);
                _attached = selectedApplication;
                _ownsAttached = false;
                _ownedProcessId = 0;
                _connectionMode = target is not null
                    ? "existing-window-exact-path"
                    : "existing-window-document-ref";
                _closeTargetWhenDone = false;
                documentState?.CaptureTarget(selectedApplication);

                foreach (var match in matches) RotHelper.ReleaseComObject(match.Document);
                foreach (var application in applications)
                    if (!ReferenceEquals(application, selectedApplication))
                        RotHelper.ReleaseComObject(application);
                return _attached;
            }

            if (!string.IsNullOrWhiteSpace(documentRef))
                throw new HwpAutomationException(
                    "HWP_DOCUMENT_NOT_FOUND",
                    $"열린 한글 창과 탭에서 documentRef '{documentRef}'를 찾지 못했습니다.",
                    "hwp_get_active_context를 다시 호출해 최신 openDocuments 목록에서 대상을 선택하세요.");
        }
        catch
        {
            foreach (var match in matches) RotHelper.ReleaseComObject(match.Document);
            foreach (var application in applications)
                if (!ReferenceEquals(application, _attached))
                    RotHelper.ReleaseComObject(application);
            throw;
        }

        foreach (var application in applications)
            if (!ReferenceEquals(application, _attached))
                RotHelper.ReleaseComObject(application);
        _closeTargetWhenDone = true;
        var createdApplication = AttachHwp(allowCreate);
        if (createdApplication is not null)
            TrackHwpInteraction(createdApplication, foreground, documentState: null, captureTarget: false);
        return createdApplication;
    }

    private static string CanonicalHwpPath(string path)
    {
        var expanded = Environment.ExpandEnvironmentVariables(path.Trim().Trim('"'));
        return Path.GetFullPath(expanded)
            .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
    }

    private static JsonObject CaptureDocumentIdentity(dynamic hwp)
    {
        var identity = new JsonObject();
        var doc = ActiveDoc(hwp);
        if (doc is null)
        {
            identity["saved"] = false;
            identity["documentRef"] = null;
            return identity;
        }

        string fullName = "";
        string documentId = "";
        try { fullName = Convert.ToString(doc.FullName) ?? ""; } catch { }
        try { documentId = Convert.ToString(doc.DocumentID) ?? ""; } catch { }
        identity["saved"] = !string.IsNullOrWhiteSpace(fullName);
        identity["fullName"] = fullName;
        identity["documentId"] = documentId;
        var windowHandle = RotHelper.HwpWindowHandle((object)hwp);
        var processId = RotHelper.ProcessIdFromWindowHandle(windowHandle);
        identity["documentRef"] = HwpDocumentRef(fullName, documentId, windowHandle, processId);
        identity["instanceRef"] = HwpInstanceRef(documentId, windowHandle, processId);
        identity["windowHandle"] = windowHandle.ToString();
        identity["processId"] = processId;
        try { identity["modified"] = Convert.ToBoolean(doc.Modified); } catch { }
        try { identity["format"] = Convert.ToString(doc.Format); } catch { }
        return identity;
    }

    public JsonObject Doctor(JsonObject args) => HwpEnvironmentDoctor.Diagnose();

    public JsonObject RepairTypeLib(JsonObject args) => HwpEnvironmentDoctor.Repair(
        Json.GetString(args, "hwpExecutable"), Json.GetBool(args, "elevate", true));

    private void CloseTargetIfNeeded(dynamic hwp, string? file)
    {
        if (!string.IsNullOrWhiteSpace(file) && _closeTargetWhenDone) CloseActiveDoc(hwp);
        _closeTargetWhenDone = false;
    }
}
