using Microsoft.Win32;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

/// <summary>
/// 한글(HWP) 어댑터: 실행 중인 한글에 HWPFrame.HwpObject COM으로 연결하거나 전용 인스턴스를 생성한다.
/// - 읽기: 문서 컨텍스트, 텍스트/선택, 표·그림 등 문서 구조, 기존 양식 필드
/// - 쓰기: 텍스트/찾아바꾸기, 글자·문단·용지 서식, 표/셀/행/열, 그림, 쪽번호·머리말/꼬리말, PDF
/// - 공개 op는 이 PC의 한글 2024에서 무인 실행과 readback을 통과한 기능으로 제한한다.
/// - 금지(정책): save_overwrite_without_backup, run_external_macro
/// - snapshot: 문서 파일 복사(저장된 경우) + 전체 텍스트 state.json
/// - restore: 전체 선택 후 텍스트 복원(서식/개체는 복원되지 않음 — 경고)
///
/// ※ 실측 기반: 이 PC의 한글 자동화 API에서 InsertText/AllReplace/SelectAll/
///   GetTextFile('TEXT'|'selection') 동작을 확인하고 구현했다.
/// </summary>
public sealed partial class HwpAdapter : ComAdapterBase, IHwpAutomationAdapter, IPreviewReuseAdapter
{
    private const int MaxChars = 20000;
    private const int MaxDiff = 100;
    private static readonly TimeSpan ComTimeout = TimeSpan.FromSeconds(120);
    private const string SecurityModuleDll = "FilePathCheckerModuleExample.dll";
    private const string SecurityModuleName = "DocBridgeFilePathChecker";
    private const string SecurityRegistryPath = @"SOFTWARE\HNC\HwpAutomation\Modules";

    /// <summary>테스트/특수 용도: 어댑터 STA 스레드 안에서 평가되는 HwpObject 팩토리</summary>
    private readonly Func<object?>? _appFactory;
    private object? _attached;
    private bool _ownsAttached;
    private int _ownedProcessId;
    private bool _securityRegistrationActive;
    private string _connectionMode = "none";
    private bool _closeTargetWhenDone;

    public HwpAdapter(Func<object?>? appFactory = null) : base("hwp", "HWPFrame.HwpObject")
    {
        _appFactory = appFactory;
    }

    private static HashSet<int> HwpProcessIds()
    {
        var ids = new HashSet<int>();
        foreach (var process in Process.GetProcessesByName("Hwp"))
        {
            using (process)
            {
                try { ids.Add(process.Id); } catch { }
            }
        }
        return ids;
    }

    private static int FindNewHwpProcessId(HashSet<int> before)
    {
        var candidates = new List<(int Id, DateTime Started)>();
        foreach (var process in Process.GetProcessesByName("Hwp"))
        {
            using (process)
            {
                try
                {
                    if (!before.Contains(process.Id)) candidates.Add((process.Id, process.StartTime));
                }
                catch { }
            }
        }
        return candidates.OrderByDescending(x => x.Started).Select(x => x.Id).FirstOrDefault();
    }

    /// <summary>반드시 STA 스레드 안에서 호출. RCW 생성/사용을 같은 아파트먼트로 통일.
    /// 팩토리가 주입된 경우 그 결과가 최종(fallback 없음) — 테스트 결정성 확보.</summary>
    private object? AttachHwp(bool allowCreate)
    {
        if (_appFactory is not null)
        {
            HashSet<int>? before = null;
            if (_attached is null) before = HwpProcessIds();
            _attached ??= _appFactory();
            _ownsAttached = _attached is not null;
            _connectionMode = _attached is null ? "none" : "injected-owned";
            if (_attached is not null && _ownedProcessId == 0)
                _ownedProcessId = RotHelper.ProcessIdFromWindowHandle(RotHelper.HwpWindowHandle(_attached));
            if (_attached is not null && _ownedProcessId == 0 && before is not null)
                _ownedProcessId = FindNewHwpProcessId(before);
            return _attached;
        }

        if (_ownsAttached && _attached is not null && allowCreate) return _attached;

        // 한글은 일반 ProgID GetActiveObject 대신 !HwpObject.* ROT 모니커를 게시한다.
        // 사용자가 열어 둔 표시 창을 우선하고, 연결한 사용자 인스턴스는 Dispose에서 종료하지 않는다.
        var ownedHwnd = _ownsAttached && _attached is not null
            ? RotHelper.HwpWindowHandle(_attached)
            : 0;
        var running = new List<object>();
        foreach (var app in RotHelper.GetHwpApplications())
        {
            var hwnd = RotHelper.HwpWindowHandle(app);
            if (!RotHelper.HwpWindowVisible(app) || (ownedHwnd != 0 && hwnd == ownedHwnd))
            {
                // 다른 도구/구버전이 남긴 숨은 -Automation -Embedding 프로세스는
                // 사용자가 연 문서가 아니며, 연결하면 소유권과 종료 처리를 잃는다.
                RotHelper.ReleaseComObject(app);
                continue;
            }
            running.Add(app);
        }
        if (running.Count > 0)
        {
            var preferred = running[0];
            for (var i = 1; i < running.Count; i++) RotHelper.ReleaseComObject(running[i]);

            if (_attached is not null && !ReferenceEquals(_attached, preferred))
                RotHelper.ReleaseComObject(_attached);
            _attached = preferred;
            _ownsAttached = false;
            _ownedProcessId = 0;
            _connectionMode = "existing-window";
            return _attached;
        }

        // Live-document tools must never launch a blank HWP process. Creating a private
        // automation instance is reserved for explicit file operations only.
        if (!allowCreate)
        {
            _connectionMode = "none";
            return null;
        }

        if (HwpUiFailureDetector.Detect() is { } activeFailure)
            throw HwpUiInitializationException(activeFailure);

        if (HwpEnvironmentDoctor.GetAutomationWindowsDirectory() is null)
            throw new HwpAutomationException(
                "HWP_AUTOMATION_ENVIRONMENT_INVALID",
                "한글이 WPF 글꼴 URI를 만들 때 필요한 windir/SystemRoot를 복구할 수 없습니다.",
                "Windows 환경 변수와 설치 폴더를 복구한 뒤 hwp_doctor를 다시 실행하세요.");

        if (_attached is not null)
        {
            RotHelper.ReleaseComObject(_attached);
            _attached = null;
        }

        var existingHwpProcesses = HwpProcessIds();
        _attached = HwpEnvironmentDoctor.RunWithAutomationWorkingDirectory(
            () => RotHelper.CreateInstance("HWPFrame.HwpObject"));
        _ownsAttached = _attached is not null;
        _connectionMode = _attached is null ? "none" : "owned-instance";
        if (_attached is not null)
            _ownedProcessId = RotHelper.ProcessIdFromWindowHandle(RotHelper.HwpWindowHandle(_attached));
        if (_attached is not null && _ownedProcessId == 0)
            _ownedProcessId = FindNewHwpProcessId(existingHwpProcesses);
        return _attached;
    }

    /// <summary>
    /// 새 문서 작성 흐름을 명시적으로 시작한다. 읽기/편집 도구는 부수 효과로
    /// 한글이나 빈 문서를 만들지 않으며, 새 문서가 필요할 때 이 진입점만 사용한다.
    /// </summary>
    public JsonObject Launch(JsonObject args)
    {
        return ComInvoke(() =>
        {
            var foreground = new ForegroundInteractionGuard(App);
            try
            {
                var importRequest = ParseDocxImportRequest(args);
                var newDocument = Json.GetBool(args, "newDocument");
                var app = AttachHwp(allowCreate: true);
                if (app is not null)
                    TrackHwpInteraction(app, foreground, documentState: null, captureTarget: false);
                if (app is null)
                    return Json.ErrorResult("한글 자동화 인스턴스를 시작할 수 없습니다", App);

                dynamic hwp = app;
                if (importRequest is not null)
                    return ImportDocxAsNativeHwp(hwp, app, importRequest, foreground);

                var createdDocument = false;
                var active = ActiveDoc(hwp);
                if (newDocument || active is null)
                {
                    if (!(bool)hwp.HAction.Run("FileNew"))
                        return Json.ErrorResult("한글 새 문서 생성(FileNew)이 실패했습니다", App);
                    createdDocument = true;
                    active = ActiveDoc(hwp);
                }
                if (active is null)
                    return Json.ErrorResult("한글은 실행됐지만 편집 가능한 문서를 만들지 못했습니다", App);

                try { hwp.XHwpWindows.Active_XHwpWindow.Visible = true; } catch { }
                if (_ownsAttached && HwpUiFailureDetector.WaitForFailure(TimeSpan.FromMilliseconds(750)) is { } uiFailure)
                    throw HwpUiInitializationException(uiFailure);
                KeepOwnedLiveDocumentOpen(hwp);

                string fullName = "";
                string documentId = "";
                try { fullName = (string)(active.FullName ?? ""); } catch { }
                try { documentId = active.DocumentID?.ToString() ?? ""; } catch { }
                var windowHandle = RotHelper.HwpWindowHandle(app);
                var processId = RotHelper.ProcessIdFromWindowHandle(windowHandle);
                var documentRef = HwpDocumentRef(
                    fullName, documentId, windowHandle, processId);

                return new JsonObject
                {
                    ["ok"] = true,
                    ["app"] = App,
                    ["documentRef"] = documentRef,
                    ["summary"] = new JsonObject
                    {
                        ["createdDocument"] = createdDocument,
                        ["creationMode"] = "native-hwp",
                        ["creationPolicyVersion"] = HwpCreationPolicy.PolicyVersion,
                        ["connectionMode"] = _connectionMode,
                        ["instanceRef"] = HwpInstanceRef(documentId, windowHandle, processId),
                        ["visible"] = true,
                        ["instruction"] = "이 문서를 계속 재사용하고 hwp_launch를 문단마다 다시 호출하지 마세요.",
                    },
                };
            }
            catch (HwpAutomationException ex)
            {
                return ex.ToResult(App);
            }
            catch (Exception ex)
            {
                return Json.ErrorResult($"hwp launch failed: {ex.Message}", App);
            }
            finally { _ = foreground.Complete(); }
        });
    }

    private static string? ResolveSecurityModulePath()
    {
        var candidates = new[]
        {
            Environment.GetEnvironmentVariable("DOCBRIDGE_HWP_SECURITY_MODULE"),
            Path.Combine(AppContext.BaseDirectory, "hwp-security", SecurityModuleDll),
            Path.Combine(AppContext.BaseDirectory, "assets", "hwp-security", SecurityModuleDll),
        };
        return candidates.FirstOrDefault(path => !string.IsNullOrWhiteSpace(path) && File.Exists(path));
    }

    /// <summary>
    /// 한컴 공식 자동화 보안 모듈을 doc-bridge 전용 이름으로 HKCU에 등록한다.
    /// 설치 위치가 바뀌어도 파일 작업 시 현재 배포 경로로 안전하게 갱신한다.
    /// </summary>
    private void EnsureFileAutomationSecurity(dynamic hwp)
    {
        if (_securityRegistrationActive) return;

        var modulePath = ResolveSecurityModulePath()
            ?? throw new InvalidOperationException(
                $"한글 자동화 보안 모듈을 찾을 수 없습니다. hwp-security\\{SecurityModuleDll}을 배포하거나 " +
                "DOCBRIDGE_HWP_SECURITY_MODULE 환경 변수를 설정하세요.");

        // 한글 2024가 32비트 COM 서버이므로 64비트 doc-bridge에서도 반드시
        // 32비트 레지스트리 뷰에 값을 등록해야 한다.
        try
        {
            var fullModulePath = Path.GetFullPath(modulePath);
            using var baseKey = RegistryKey.OpenBaseKey(RegistryHive.CurrentUser, RegistryView.Registry32);
            using var readKey = baseKey.OpenSubKey(SecurityRegistryPath, writable: false);
            var configuredPath = readKey?.GetValue(
                SecurityModuleName, null, RegistryValueOptions.DoNotExpandEnvironmentNames) as string;

            if (!string.Equals(configuredPath, fullModulePath, StringComparison.OrdinalIgnoreCase))
            {
                using var writeKey = baseKey.CreateSubKey(SecurityRegistryPath, writable: true)
                    ?? throw new InvalidOperationException(
                        $"한글 자동화 보안 모듈 레지스트리를 열 수 없습니다: HKCU\\{SecurityRegistryPath}");
                writeKey.SetValue(SecurityModuleName, fullModulePath, RegistryValueKind.String);
                writeKey.Flush();
            }
            // 일부 한글 2024 빌드는 성공해도 false를 반환하므로 예외 발생 여부로 판단한다.
            _ = hwp.RegisterModule("FilePathCheckDLL", SecurityModuleName);
            _securityRegistrationActive = true;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"한글 자동화 보안 모듈 등록에 실패했습니다: {ex.Message}", ex);
        }
    }

    private static string GetDocText(dynamic hwp)
    {
        try { return DecodeHwpSerializedText((string)(hwp.GetTextFile("TEXT", "") ?? "")); }
        catch { return ""; }
    }

    private static HwpAutomationException HwpUiInitializationException(HwpUiFailure failure) =>
        new(
            "HWP_UI_INITIALIZATION_FAILED",
            $"한글 UI 초기화 오류가 감지되었습니다 ({failure.Signature}). 같은 실행을 자동 재시도하지 않습니다.",
            HwpUiFailureDetector.UpdateAction);

    /// <summary>한글 InsertText가 개행을 제거하는 특성에 맞춘 비교용 정규화</summary>
    private static string NormalizeNewlines(string s) => s.Replace("\r\n", "\n").Replace('\r', '\n');

    private static string GetSelectionText(dynamic hwp)
    {
        try { return DecodeHwpSerializedText((string)(hwp.GetTextFile("TEXT", "saveblock:true") ?? "")); }
        catch { return ""; }
    }

    /// <summary>
    /// 현재 문서를 한글 네이티브 포맷(Base64 문자열)으로 메모리에 추출한다.
    /// 저장 전 변경 내용과 표/서식/개체를 포함하므로 라이브 창 편집의 전체 백업에 사용한다.
    /// </summary>
    private static string GetNativeDocumentSnapshot(dynamic hwp)
    {
        var data = (string)(hwp.GetTextFile("HWP", "") ?? "");
        if (string.IsNullOrWhiteSpace(data))
            throw new InvalidOperationException("한글 네이티브 문서 스냅샷을 만들 수 없습니다");
        return data;
    }

    private static bool RestoreNativeDocumentSnapshot(dynamic hwp, string data)
    {
        // 표 셀/머리말/개체 편집 안에서 실패한 경우에도 문서 본문으로 확실히 빠져나온다.
        for (var i = 0; i < 4; i++)
        {
            try { hwp.HAction.Run("Cancel"); } catch { }
            try { hwp.HAction.Run("CloseEx"); } catch { }
        }
        try { hwp.HAction.Run("MoveDocBegin"); } catch { }
        if (!(bool)hwp.HAction.Run("SelectAll")) return false;
        if (!(bool)hwp.HAction.Run("Delete")) return false;
        try { return Convert.ToInt32(hwp.SetTextFile(data, "HWP", "insertfile")) != 0; }
        catch { return false; }
    }

    private static dynamic? ActiveDoc(dynamic hwp)
    {
        try
        {
            dynamic active = hwp.XHwpDocuments.Active_XHwpDocument;
            if (active is not null) return active;
        }
        catch { }
        try
        {
            dynamic docs = hwp.XHwpDocuments;
            if ((int)docs.Count > 0) return docs.Item(0);
        }
        catch { }
        return null;
    }

    /// <summary>ops 첫 항목의 명시적 파일 대상.</summary>
    private static string? FileArgOf(IReadOnlyList<JsonObject> ops) =>
        ops.Count > 0 ? Json.GetString(ops[0], "file") : null;

    /// <summary>hwp_get_active_context.openDocuments에서 선택한 라이브 문서 대상.</summary>
    private static string? DocumentRefArgOf(IReadOnlyList<JsonObject> ops) =>
        ops.Count > 0 ? Json.GetString(ops[0], "documentRef") : null;

    private static string? HwpTargetSelectorError(IReadOnlyList<JsonObject> ops)
    {
        var files = ops.Select(op => Json.GetString(op, "file"))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => Path.GetFullPath(value!))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        var documentRefs = ops.Select(op => Json.GetString(op, "documentRef"))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        var targetedOps = ops.Count(op =>
            !string.IsNullOrWhiteSpace(Json.GetString(op, "file")) ||
            !string.IsNullOrWhiteSpace(Json.GetString(op, "documentRef")));

        if (ops.Any(op =>
                !string.IsNullOrWhiteSpace(Json.GetString(op, "file")) &&
                !string.IsNullOrWhiteSpace(Json.GetString(op, "documentRef"))))
            return "한글 op 하나에 file과 documentRef를 동시에 지정할 수 없습니다";
        if (files.Count > 0 && documentRefs.Count > 0)
            return "한 배치에서 file 대상과 documentRef 대상을 섞을 수 없습니다";
        if (files.Count > 1)
            return "한 배치의 모든 한글 op는 같은 file을 지정해야 합니다";
        if (documentRefs.Count > 1)
            return "한 배치의 모든 한글 op는 같은 documentRef를 지정해야 합니다";
        if (targetedOps > 0 && targetedOps != ops.Count)
            return "한 배치의 모든 한글 op는 대상을 모두 생략하거나 같은 file/documentRef를 지정해야 합니다";
        return null;
    }

    private static string FileHash(string path)
    {
        using var stream = File.OpenRead(path);
        return Convert.ToHexString(SHA256.HashData(stream)).ToLowerInvariant();
    }

    /// <summary>
    /// file 인터가 있으면 해당 파일을 열어 활성 문서로 만든다.
    /// 한글 COM 인스턴스는 헤드리스 전용 + ROT 미등록이므로,
    /// 사용자 문서와 주고받는 유일한 안전한 경로는 파일이다.
    /// </summary>
    private dynamic? OpenOrGetDoc(dynamic hwp, string? file)
    {
        if (string.IsNullOrEmpty(file)) return ActiveDoc(hwp);
        if (!File.Exists(file)) return null;
        var target = CanonicalHwpPath(file);

        // 이미 같은 파일이 열려 있으면 재사용
        try
        {
            dynamic docs = hwp.XHwpDocuments;
            for (var i = 0; i < (int)docs.Count; i++)
            {
                dynamic d = docs.Item(i);
                object docObject = (object)d;
                string fn = "";
                try { fn = (string)(d.FullName ?? ""); } catch { }
                if (!string.IsNullOrWhiteSpace(fn) &&
                    string.Equals(CanonicalHwpPath(fn), target, StringComparison.OrdinalIgnoreCase))
                {
                    d.SetActive_XHwpDocument();
                    _closeTargetWhenDone = false;
                    RotHelper.ReleaseComObject(docObject);
                    return ActiveDoc(hwp);
                }
                RotHelper.ReleaseComObject(docObject);
            }
        }
        catch { }

        // 보안 모듈이 없으면 FileOpen이 숨은 승인 창에서 무기한 대기할 수 있으므로
        // 새 파일을 열기 전에 등록을 완료하거나 명시적으로 실패시킨다.
        EnsureFileAutomationSecurity(hwp);
        _closeTargetWhenDone = true;

        // hwp.Open(path)는 인자 오버로드 불일치, 3인자 형태는 모달을 띄워 블록될 수 있다 (실측).
        // FileOpen HAction + 파라미터 명시 방식이 헤드리스에서 안정적이다.
        try
        {
            var ok = OpenDocumentWithFormat(hwp, target, HwpAutomationFormatForPath(target));
            if (!ok) { _closeTargetWhenDone = false; return null; }
        }
        catch { _closeTargetWhenDone = false; return null; }
        return ActiveDoc(hwp);
    }

    /// <summary>
    /// 활성 문서를 원래 파일에 저장 (호스트가 스냅샷 백업을 선행하므로 덮어쓰기 아님).
    /// hwp.Save()는 이 버전(13.0.0.866)에서 COM 오류를 던지므로
    /// FileSaveAs + SaveOverWrite 방식을 사용한다 (실측으로 검증됨).
    /// </summary>
    private static void SaveActiveDoc(dynamic hwp, string file, bool overwrite = true)
    {
        if (!overwrite && File.Exists(file))
            throw new HwpAutomationException(
                "HWP_OUTPUT_EXISTS",
                $"출력 파일이 이미 있어 덮어쓰지 않았습니다: {file}",
                "다른 outputFile 이름을 지정하세요.");

        object? actionObject = null;
        object? parameterObject = null;
        object? hSetObject = null;
        try
        {
            dynamic act = hwp.HAction;
            dynamic sa = hwp.HParameterSet.HFileSaveAs;
            dynamic hSet = sa.HSet;
            actionObject = (object)act;
            parameterObject = (object)sa;
            hSetObject = (object)hSet;
            act.GetDefault("FileSaveAs", hSet);
            sa.SaveFileName = file;
            try { sa.SaveFormat = HwpAutomationFormatForPath(file); } catch { }
            try { sa.SaveOverWrite = overwrite; } catch { }
            if (!(bool)act.Execute("FileSaveAs", hSet))
                throw new InvalidOperationException($"한글 FileSaveAs가 실패했습니다: {file}");
        }
        finally
        {
            RotHelper.ReleaseComObject(hSetObject);
            RotHelper.ReleaseComObject(parameterObject);
            RotHelper.ReleaseComObject(actionObject);
        }
    }

    /// <summary>파일 잠금 해제를 위해 활성 문서 닫기 (실패해도 무시)</summary>
    private static void CloseActiveDoc(dynamic hwp)
    {
        try { hwp.HAction.Run("FileClose"); } catch { }
    }

    private static bool ExecInsertText(dynamic hwp, string text)
    {
        var parts = NormalizeNewlines(text).Split('\n');
        for (var i = 0; i < parts.Length; i++)
        {
            if (parts[i].Length > 0)
            {
                dynamic act = hwp.HAction;
                dynamic ps = hwp.HParameterSet.HInsertText;
                act.GetDefault("InsertText", ps.HSet);
                ps.Text = parts[i];
                if (!(bool)act.Execute("InsertText", ps.HSet)) return false;
            }
            if (i < parts.Length - 1 && !(bool)hwp.HAction.Run("BreakPara")) return false;
        }
        return true;
    }

    private static bool ExecReplaceDocumentText(dynamic hwp, string text, JsonObject op)
    {
        try { hwp.HAction.Run("MoveDocBegin"); } catch { }
        var context = CaptureCurrentNativeStyle(hwp, "document-first-paragraph");
        if (!(bool)hwp.HAction.Run("SelectAll")) return false;
        if (text.Length == 0) return (bool)hwp.HAction.Run("Delete");
        _ = hwp.HAction.Run("Delete");
        if (!PrepareContextualWriteStyle(hwp, op, context)) return false;
        return ExecInsertText(hwp, text);
    }

    private static bool ExecAppendText(dynamic hwp, string text, bool startNewParagraph, JsonObject op)
    {
        var before = NormalizeNewlines(GetDocText(hwp));
        try { hwp.HAction.Run("MoveDocEnd"); }
        catch { return false; }
        var context = CaptureCurrentNativeStyle(hwp, "previous-document-end");

        var normalized = NormalizeNewlines(text);
        if (startNewParagraph && before.TrimEnd('\n').Length > 0 &&
            !before.EndsWith('\n') && !normalized.StartsWith('\n'))
        {
            if (!(bool)hwp.HAction.Run("BreakPara")) return false;
        }
        if (!PrepareContextualWriteStyle(hwp, op, context)) return false;
        return ExecInsertText(hwp, text);
    }

    private static int CountTextOccurrences(string document, string anchor, bool matchCase)
    {
        if (string.IsNullOrEmpty(anchor)) return 0;
        var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var count = 0;
        var index = 0;
        while ((index = document.IndexOf(anchor, index, comparison)) >= 0)
        {
            count++;
            index += anchor.Length;
        }
        return count;
    }

    private static int IndexOfTextOccurrence(string document, string anchor, int occurrence, bool matchCase)
    {
        if (string.IsNullOrEmpty(anchor) || occurrence < 1) return -1;
        var comparison = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var index = 0;
        for (var i = 0; i < occurrence; i++)
        {
            index = document.IndexOf(anchor, index, comparison);
            if (index < 0) return -1;
            if (i < occurrence - 1) index += anchor.Length;
        }
        return index;
    }

    /// <summary>
    /// 문서 처음부터 anchor의 n번째 항목을 찾아 선택 상태로 둔다.
    /// IgnoreMessage를 강제해 "문서 끝" 모달을 막고, 호출자가 MoveLeft/MoveRight로
    /// 선택 시작/끝에 캐럿을 정확히 접을 수 있게 한다.
    /// </summary>
    private static bool SelectTextOccurrence(dynamic hwp, string anchor, int occurrence, bool matchCase)
    {
        if (string.IsNullOrEmpty(anchor) || occurrence < 1) return false;
        int? previousMessageMode = null;
        try
        {
            try { previousMessageMode = Convert.ToInt32(hwp.GetMessageBoxMode()); } catch { }
            try { hwp.SetMessageBoxMode(0x2FFF1); } catch { }
            if (!(bool)hwp.HAction.Run("MoveDocBegin")) return false;

            dynamic act = hwp.HAction;
            dynamic find = hwp.HParameterSet.HFindReplace;
            act.GetDefault("RepeatFind", find.HSet);
            find.FindString = anchor;
            try { find.Direction = hwp.FindDir("Forward"); } catch { }
            try { find.ReplaceMode = 0; } catch { }
            try { find.IgnoreMessage = 1; } catch { }
            try { find.MatchCase = matchCase ? 1 : 0; } catch { }
            try { find.SeveralWords = 0; } catch { }
            try { find.UseWildCards = 0; } catch { }
            try { find.WholeWordOnly = 0; } catch { }
            try { find.AllWordForms = 0; } catch { }
            try { find.FindRegExp = 0; } catch { }
            try { find.FindJaso = 0; } catch { }
            try { find.HanjaFromHangul = 0; } catch { }
            try { find.FindType = 1; } catch { }

            for (var i = 0; i < occurrence; i++)
                if (!(bool)act.Execute("RepeatFind", find.HSet)) return false;
            return true;
        }
        finally
        {
            if (previousMessageMode is not null)
                try { hwp.SetMessageBoxMode(previousMessageMode.Value); } catch { }
        }
    }

    private static bool ExecInsertRelativeToText(dynamic hwp, JsonObject op, bool before)
    {
        var anchor = Json.GetString(op, "anchor")!;
        var text = Json.GetString(op, "text")!;
        var occurrence = Json.GetInt(op, "occurrence") ?? 1;
        var matchCase = Json.GetBool(op, "matchCase", true);
        var mode = (Json.GetString(op, "mode") ?? "paragraph").ToLowerInvariant();
        if (mode is not ("inline" or "paragraph"))
            throw new ArgumentException("mode must be 'inline' or 'paragraph'");
        if (!SelectTextOccurrence(hwp, anchor, occurrence, matchCase)) return false;
        var anchorContext = CaptureCurrentNativeStyle(hwp, $"anchor:{anchor}#{occurrence}");

        // 찾기 결과의 선택 시작/끝을 직접 받아 캐럿을 배치한다. MoveLeft/MoveRight는
        // 한글 버전과 선택 방향에 따라 한 글자 더 이동할 수 있어 사용하지 않는다.
        dynamic start;
        dynamic end;
        try
        {
            start = hwp.CreateSet("ListParaPos");
            end = hwp.CreateSet("ListParaPos");
        }
        catch (Exception ex) { throw new InvalidOperationException($"CreateSet(ListParaPos) failed: {ex.Message}", ex); }
        try
        {
            if (!(bool)hwp.GetSelectedPosBySet(start, end)) return false;
        }
        catch (Exception ex) { throw new InvalidOperationException($"GetSelectedPosBySet failed: {ex.Message}", ex); }
        var context = mode == "paragraph"
            ? ResolveParagraphContextStyle((object)hwp, (object)start, (object)end, anchorContext)
            : anchorContext;
        try
        {
            object position = before ? (object)start : (object)end;
            if (!(bool)hwp.SetPosBySet(position)) return false;
        }
        catch (Exception ex) { throw new InvalidOperationException($"SetPosBySet failed: {ex.Message}", ex); }
        if (mode == "inline")
        {
            if (!PrepareContextualWriteStyle(hwp, op, context)) return false;
            return ExecInsertText(hwp, text);
        }

        if (before)
        {
            if (!(bool)hwp.HAction.Run("MoveParaBegin")) return false;
            if (!PrepareContextualWriteStyle(hwp, op, context)) return false;
            if (!ExecInsertText(hwp, text)) return false;
            return NormalizeNewlines(text).EndsWith('\n') || (bool)hwp.HAction.Run("BreakPara");
        }

        if (!(bool)hwp.HAction.Run("MoveParaEnd")) return false;
        if (!NormalizeNewlines(text).StartsWith('\n') && !(bool)hwp.HAction.Run("BreakPara")) return false;
        if (!PrepareContextualWriteStyle(hwp, op, context)) return false;
        return ExecInsertText(hwp, text);
    }

    private sealed record TableInsertResult(
        bool Ok,
        int Rows,
        int Cols,
        int Cells,
        int StyledCells,
        int TableCountBefore,
        int TableCountAfter,
        IReadOnlyList<string> ExpectedTexts);

    private static List<List<string>> ParseTableRows(JsonObject op)
    {
        var rowsNode = Json.GetArr(op, "rows")
            ?? throw new ArgumentException("insert_table.rows must be a non-empty 2D array");
        if (rowsNode.Count is < 1 or > 100)
            throw new ArgumentException("insert_table supports 1 to 100 rows");

        var rows = new List<List<string>>(rowsNode.Count);
        int? expectedCols = null;
        foreach (var rowNode in rowsNode)
        {
            if (rowNode is not JsonArray row || row.Count == 0)
                throw new ArgumentException("each insert_table row must be a non-empty array");
            expectedCols ??= row.Count;
            if (row.Count != expectedCols)
                throw new ArgumentException("all insert_table rows must have the same number of columns");
            if (row.Count > 10)
                throw new ArgumentException("insert_table supports at most 10 columns");

            var values = new List<string>(row.Count);
            foreach (var cell in row)
            {
                if (cell is null) values.Add("");
                else if (cell is JsonValue value && value.TryGetValue<string>(out var text)) values.Add(text);
                else throw new ArgumentException("insert_table cell values must be strings or null");
            }
            rows.Add(values);
        }

        if (rows.Count * expectedCols!.Value > 500)
            throw new ArgumentException("insert_table supports at most 500 cells per operation");
        return rows;
    }

    private static double JsonNumber(JsonObject obj, string key, double fallback)
    {
        if (!obj.TryGetPropertyValue(key, out var node) || node is not JsonValue value) return fallback;
        if (value.TryGetValue<double>(out var d)) return d;
        if (value.TryGetValue<int>(out var i)) return i;
        return fallback;
    }

    private static List<double> ParseColumnWidths(JsonObject op, int cols)
    {
        var widths = new List<double>(cols);
        var node = Json.GetArr(op, "columnWidths");
        if (node is not null && node.Count != cols)
            throw new ArgumentException("insert_table.columnWidths must match the number of columns");

        if (node is not null)
        {
            foreach (var item in node)
            {
                if (item is not JsonValue value ||
                    !(value.TryGetValue<double>(out var width) ||
                      (value.TryGetValue<int>(out var integerWidth) && (width = integerWidth) >= 0)) ||
                    width <= 0)
                    throw new ArgumentException("insert_table.columnWidths entries must be positive numbers");
                widths.Add(width);
            }
        }
        else
        {
            for (var i = 0; i < cols; i++) widths.Add(1.0);
        }

        var sum = widths.Sum();
        return widths.Select(width => width / sum).ToList();
    }

    private static (int R, int G, int B)? ParseHexColor(string? color)
    {
        if (string.IsNullOrWhiteSpace(color)) return null;
        var hex = color.Trim().TrimStart('#');
        if (hex.Length != 6 || !int.TryParse(hex, System.Globalization.NumberStyles.HexNumber,
                System.Globalization.CultureInfo.InvariantCulture, out var rgb))
            throw new ArgumentException($"invalid table fill color '{color}' (expected #RRGGBB)");
        return ((rgb >> 16) & 255, (rgb >> 8) & 255, rgb & 255);
    }

    private sealed record HorizontalMerge(int StartRow, int StartCol, int EndCol);

    private static Dictionary<(int Row, int Col), JsonObject> ParseCellStyles(JsonObject op, int rows, int cols)
    {
        var result = new Dictionary<(int, int), JsonObject>();
        var styles = Json.GetArr(op, "cellStyles");
        if (styles is null) return result;

        foreach (var node in styles)
        {
            if (node is not JsonObject style)
                throw new ArgumentException("insert_table.cellStyles entries must be objects");
            var row = Json.GetInt(style, "row")
                ?? throw new ArgumentException("insert_table.cellStyles[].row is required");
            var col = Json.GetInt(style, "col")
                ?? throw new ArgumentException("insert_table.cellStyles[].col is required");
            if (row < 0 || row >= rows || col < 0 || col >= cols)
                throw new ArgumentException($"insert_table cell style ({row},{col}) is outside the table");
            _ = ParseHexColor(Json.GetString(style, "fill"));
            if (style.TryGetPropertyValue("borders", out var borderNode) && borderNode is not JsonObject)
                throw new ArgumentException("insert_table cell style borders must be an object");
            if (borderNode is JsonObject borders)
            {
                foreach (var side in new[] { "left", "right", "top", "bottom" })
                {
                    if (!borders.TryGetPropertyValue(side, out var sideNode)) continue;
                    if (sideNode is not JsonValue sideValue || !sideValue.TryGetValue<bool>(out _))
                        throw new ArgumentException($"insert_table cell border '{side}' must be true or false");
                }
            }
            var size = JsonNumber(style, "fontSize", 9.5);
            if (size is < 6 or > 72)
                throw new ArgumentException("insert_table cell style fontSize must be between 6 and 72 points");
            result[(row, col)] = style;
        }
        return result;
    }

    private static List<HorizontalMerge> ParseMergeCells(JsonObject op, int rows, int cols)
    {
        var result = new List<HorizontalMerge>();
        var merges = Json.GetArr(op, "mergeCells");
        if (merges is null) return result;

        foreach (var node in merges)
        {
            if (node is not JsonObject merge)
                throw new ArgumentException("insert_table.mergeCells entries must be objects");
            var startRow = Json.GetInt(merge, "startRow")
                ?? throw new ArgumentException("insert_table.mergeCells[].startRow is required");
            var endRow = Json.GetInt(merge, "endRow") ?? startRow;
            var startCol = Json.GetInt(merge, "startCol")
                ?? throw new ArgumentException("insert_table.mergeCells[].startCol is required");
            var endCol = Json.GetInt(merge, "endCol")
                ?? throw new ArgumentException("insert_table.mergeCells[].endCol is required");
            if (startRow != endRow)
                throw new ArgumentException("insert_table currently supports horizontal merges within one row only");
            if (startRow < 0 || startRow >= rows || startCol < 0 || endCol >= cols || endCol <= startCol)
                throw new ArgumentException("insert_table.mergeCells contains an invalid range");
            result.Add(new HorizontalMerge(startRow, startCol, endCol));
        }

        return result
            .OrderByDescending(merge => merge.StartRow)
            .ThenByDescending(merge => merge.StartCol)
            .ToList();
    }

    private static bool ApplyCellFill(dynamic hwp, (int R, int G, int B) color)
    {
        try
        {
            try { hwp.HAction.Run("TableCellBlock"); } catch { }
            dynamic fill = hwp.HParameterSet.HCellBorderFill;
            hwp.HAction.GetDefault("CellFill", fill.HSet);
            fill.FillAttr.type = hwp.BrushType("NullBrush|WinBrush");
            fill.FillAttr.WinBrushFaceColor = hwp.RGBColor(color.R, color.G, color.B);
            fill.FillAttr.WinBrushHatchColor = hwp.RGBColor(153, 153, 153);
            fill.FillAttr.WinBrushFaceStyle = hwp.HatchStyle("None");
            fill.FillAttr.WindowsBrush = 1;
            return (bool)hwp.HAction.Execute("CellFill", fill.HSet);
        }
        finally
        {
            try { hwp.HAction.Run("Cancel"); } catch { }
        }
    }

    private static bool ApplyCurrentCellBorders(dynamic hwp, JsonObject borders)
    {
        try
        {
            if (!(bool)hwp.HAction.Run("TableCellBlock")) return false;
            dynamic border = hwp.HParameterSet.HCellBorderFill;
            hwp.HAction.GetDefault("CellBorderFill", border.HSet);
            var anyVisible = false;
            foreach (var (key, property) in new[]
            {
                ("left", "BorderTypeLeft"),
                ("right", "BorderTypeRight"),
                ("top", "BorderTypeTop"),
                ("bottom", "BorderTypeBottom"),
            })
            {
                if (!borders.TryGetPropertyValue(key, out var node) || node is not JsonValue value ||
                    !value.TryGetValue<bool>(out var visible)) continue;
                anyVisible |= visible;
                var lineType = visible ? Convert.ToInt32(hwp.HwpLineType("Solid")) : 0;
                switch (property)
                {
                    case "BorderTypeLeft":
                        border.BorderTypeLeft = lineType;
                        if (visible) border.BorderWidthLeft = hwp.HwpLineWidth("0.12mm");
                        break;
                    case "BorderTypeRight":
                        border.BorderTypeRight = lineType;
                        if (visible) border.BorderWidthRight = hwp.HwpLineWidth("0.12mm");
                        break;
                    case "BorderTypeTop":
                        border.BorderTypeTop = lineType;
                        if (visible) border.BorderWidthTop = hwp.HwpLineWidth("0.12mm");
                        break;
                    case "BorderTypeBottom":
                        border.BorderTypeBottom = lineType;
                        if (visible) border.BorderWidthBottom = hwp.HwpLineWidth("0.12mm");
                        break;
                }
            }
            if (!anyVisible)
            {
                try { border.TypeVert = 0; } catch { }
                try { border.TypeHorz = 0; } catch { }
            }
            return (bool)hwp.HAction.Execute("CellBorderFill", border.HSet);
        }
        finally
        {
            try { hwp.HAction.Run("Cancel"); } catch { }
        }
    }

    private static bool HideCurrentCellBorders(dynamic hwp) => ApplyCurrentCellBorders(hwp, new JsonObject
    {
        ["left"] = false,
        ["right"] = false,
        ["top"] = false,
        ["bottom"] = false,
    });

    private static bool HasVisibleBorder(JsonObject borders)
    {
        foreach (var side in new[] { "left", "right", "top", "bottom" })
            if (borders.TryGetPropertyValue(side, out var node) && node is JsonValue value &&
                value.TryGetValue<bool>(out var visible) && visible)
                return true;
        return false;
    }

    private static JsonObject InvisibleVersionOfBorders(JsonObject borders)
    {
        var result = new JsonObject();
        foreach (var side in new[] { "left", "right", "top", "bottom" })
            if (borders.ContainsKey(side)) result[side] = false;
        return result;
    }

    private static bool ApplyDeferredCellBorders(
        dynamic hwp,
        IReadOnlyList<(int Row, int Col, JsonObject Borders)> deferred,
        int originalCols)
    {
        foreach (var item in deferred)
        {
            if (!(bool)hwp.HAction.Run("TableColBegin")) return false;
            if (!(bool)hwp.HAction.Run("TableColPageUp")) return false;
            var cellsBefore = item.Row * originalCols + item.Col;
            for (var i = 0; i < cellsBefore; i++)
                if (!(bool)hwp.HAction.Run("TableRightCell")) return false;
            if (!ApplyCurrentCellBorders(hwp, item.Borders)) return false;
        }
        return true;
    }

    private static bool CenterCurrentCell(dynamic hwp)
    {
        try
        {
            if (!(bool)hwp.HAction.Run("TableCellBlock")) return false;
            return (bool)hwp.HAction.Run("TableCellAlignCenterCenter");
        }
        finally
        {
            try { hwp.HAction.Run("Cancel"); } catch { }
        }
    }

    private static bool ApplyHorizontalMerges(dynamic hwp, IReadOnlyList<HorizontalMerge> merges, int originalCols)
    {
        foreach (var merge in merges)
        {
            if (!(bool)hwp.HAction.Run("TableColBegin")) return false;
            if (!(bool)hwp.HAction.Run("TableColPageUp")) return false;
            var cellsBefore = merge.StartRow * originalCols + merge.StartCol;
            for (var i = 0; i < cellsBefore; i++)
                if (!(bool)hwp.HAction.Run("TableRightCell")) return false;

            if (!(bool)hwp.HAction.Run("TableCellBlockExtendAbs")) return false;
            for (var i = merge.StartCol; i < merge.EndCol; i++)
                if (!(bool)hwp.HAction.Run("TableRightCell")) return false;
            if (!(bool)hwp.HAction.Run("TableMergeCell")) return false;
            try { hwp.HAction.Run("Cancel"); } catch { }
        }
        return true;
    }

    private void KeepOwnedLiveDocumentOpen(dynamic hwp)
    {
        if (!_ownsAttached) return;
        try { hwp.XHwpWindows.Active_XHwpWindow.Visible = true; } catch { }
        // A live, untitled workflow intentionally hands the newly created window to the user.
        // Releasing the RCW must not call Quit when the CLI process exits.
        _ownsAttached = false;
        _connectionMode = "new-visible-window";
    }

    private static int CountTableControls(dynamic hwp)
    {
        var count = 0;
        try
        {
            dynamic? ctrl = hwp.HeadCtrl;
            for (var guard = 0; ctrl is not null && guard < 10000; guard++)
            {
                try
                {
                    if (string.Equals((string)(ctrl.CtrlID ?? ""), "tbl", StringComparison.OrdinalIgnoreCase)) count++;
                }
                catch { }
                try { ctrl = ctrl.Next; } catch { break; }
            }
        }
        catch { }
        return count;
    }

    private static JsonObject InspectStructure(dynamic hwp, int maxControls, bool includePageCount)
    {
        var controls = new JsonArray();
        var counts = new JsonObject();
        var countMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var scanned = 0;
        dynamic? ctrl = null;
        try { ctrl = hwp.HeadCtrl; } catch { }
        while (ctrl is not null && scanned < maxControls)
        {
            string id = "unknown";
            string description = "";
            try { id = (string)(ctrl.CtrlID ?? "unknown"); } catch { }
            try { description = (string)(ctrl.UserDesc ?? ""); } catch { }
            countMap[id] = countMap.GetValueOrDefault(id) + 1;
            controls.Add(new JsonObject
            {
                ["index"] = scanned,
                ["controlId"] = id,
                ["description"] = description,
            });
            scanned++;
            try { ctrl = ctrl.Next; } catch { ctrl = null; }
        }
        foreach (var (id, count) in countMap) counts[id] = count;

        var result = new JsonObject
        {
            ["controls"] = controls,
            ["controlCountScanned"] = scanned,
            ["controlsTruncated"] = ctrl is not null,
            ["countsByControlId"] = counts,
            ["pageCountIncluded"] = includePageCount,
            ["paginationPerformed"] = includePageCount,
        };
        if (includePageCount)
        {
            try { result["pageCount"] = Convert.ToInt32(hwp.PageCount); }
            catch (Exception ex) { result["pageCountError"] = ex.Message; }
        }
        return result;
    }

    private static TableInsertResult ExecInsertTable(dynamic hwp, JsonObject op)
    {
        var rows = ParseTableRows(op);
        var rowCount = rows.Count;
        var colCount = rows[0].Count;
        var hasExplicitColumnWidths = Json.GetArr(op, "columnWidths") is not null;
        var ratios = ParseColumnWidths(op, colCount);
        var header = Json.GetBool(op, "header", true);
        var headerFill = ParseHexColor(Json.GetString(op, "headerFill"));
        var firstColumnFill = ParseHexColor(Json.GetString(op, "firstColumnFill"));
        var fontSize = JsonNumber(op, "fontSize", 9.5);
        var cellStyles = ParseCellStyles(op, rowCount, colCount);
        var merges = ParseMergeCells(op, rowCount, colCount);
        var verticalCenter = Json.GetBool(op, "verticalCenter");
        var hideAllBorders = Json.GetBool(op, "hideAllBorders");
        if (fontSize is < 6 or > 72) throw new ArgumentException("insert_table.fontSize must be between 6 and 72 points");

        var tableCountBefore = CountTableControls(hwp);
        try { hwp.HAction.Run("MoveDocEnd"); } catch { }
        if (!(bool)hwp.HAction.Run("BreakPara"))
            return new TableInsertResult(false, rowCount, colCount, 0, 0, tableCountBefore,
                CountTableControls(hwp), Array.Empty<string>());

        dynamic create = hwp.HParameterSet.HTableCreation;
        hwp.HAction.GetDefault("TableCreate", create.HSet);
        create.Rows = rowCount;
        create.Cols = colCount;
        create.WidthType = hasExplicitColumnWidths ? 2 : 0;
        create.HeightType = 0;

        dynamic section = hwp.HParameterSet.HSecDef;
        hwp.HAction.GetDefault("PageSetup", section.HSet);
        var totalWidth = Convert.ToInt32(section.PageDef.PaperWidth)
            - Convert.ToInt32(section.PageDef.LeftMargin)
            - Convert.ToInt32(section.PageDef.RightMargin)
            - Convert.ToInt32(section.PageDef.GutterLen)
            - Convert.ToInt32(hwp.MiliToHwpUnit(2.0));
        create.WidthValue = totalWidth;
        create.CreateItemArray("ColWidth", colCount);
        dynamic colWidths = create.ColWidth;
        var usableWidth = totalWidth - Convert.ToInt32(hwp.MiliToHwpUnit(3.6 * colCount));
        var assigned = 0;
        for (var col = 0; col < colCount; col++)
        {
            var width = col == colCount - 1
                ? usableWidth - assigned
                : Convert.ToInt32(Math.Round(usableWidth * ratios[col]));
            colWidths.Item[col] = width;
            assigned += width;
        }
        create.TableProperties.Width = totalWidth;
        try { create.TableProperties.TreatAsChar = true; } catch { }
        if (!(bool)hwp.HAction.Execute("TableCreate", create.HSet))
            return new TableInsertResult(false, rowCount, colCount, 0, 0, tableCountBefore,
                CountTableControls(hwp), Array.Empty<string>());

        try
        {
            dynamic ctrl = hwp.CurSelectedCtrl ?? hwp.ParentCtrl;
            dynamic props = hwp.CreateSet("Table");
            props.SetItem("TreatAsChar", true);
            ctrl.Properties = props;
        }
        catch { }

        if (header)
        {
            try
            {
                dynamic shape = hwp.HParameterSet.HShapeObject;
                hwp.HAction.GetDefault("TablePropertyDialog", shape.HSet);
                shape.ShapeTableCell.Header = true;
                _ = hwp.HAction.Execute("TablePropertyDialog", shape.HSet);
            }
            catch { }
        }

        var styledCells = 0;
        var expectedTexts = new List<string>(rowCount * colCount);
        var deferredBorders = new List<(int Row, int Col, JsonObject Borders)>();
        for (var row = 0; row < rowCount; row++)
        {
            for (var col = 0; col < colCount; col++)
            {
                var isHeader = header && row == 0;
                var isFirstColumn = col == 0;
                cellStyles.TryGetValue((row, col), out var cellStyle);
                var cellBold = cellStyle is null
                    ? isHeader || isFirstColumn
                    : Json.GetBool(cellStyle, "bold", isHeader || isFirstColumn);
                var cellItalic = cellStyle is not null && Json.GetBool(cellStyle, "italic");
                var cellFontSize = cellStyle is null ? fontSize : JsonNumber(cellStyle, "fontSize", fontSize);
                var defaultAlign = isHeader || col < 2 ? "center" : "left";
                var cellAlign = Json.GetString(cellStyle, "align") ?? defaultAlign;
                var style = new JsonObject
                {
                    ["bold"] = cellBold,
                    ["italic"] = cellItalic,
                    ["fontSize"] = cellFontSize,
                    ["align"] = cellAlign,
                };
                if (!ApplyCharShape(hwp, style) || !ApplyParagraphAlignment(hwp, style))
                    return new TableInsertResult(false, rowCount, colCount, expectedTexts.Count, styledCells,
                        tableCountBefore, CountTableControls(hwp), expectedTexts);

                if (hideAllBorders && !HideCurrentCellBorders(hwp))
                    return new TableInsertResult(false, rowCount, colCount, expectedTexts.Count, styledCells,
                        tableCountBefore, CountTableControls(hwp), expectedTexts);

                if (cellStyle is not null && Json.GetBool(cellStyle, "hideBorders") && !HideCurrentCellBorders(hwp))
                    return new TableInsertResult(false, rowCount, colCount, expectedTexts.Count, styledCells,
                        tableCountBefore, CountTableControls(hwp), expectedTexts);

                var cellBorders = Json.GetObj(cellStyle, "borders");
                if (cellBorders is not null)
                {
                    var invisibleBorders = InvisibleVersionOfBorders(cellBorders);
                    if (invisibleBorders.Count > 0 && !ApplyCurrentCellBorders(hwp, invisibleBorders))
                        return new TableInsertResult(false, rowCount, colCount, expectedTexts.Count, styledCells,
                            tableCountBefore, CountTableControls(hwp), expectedTexts);
                    if (HasVisibleBorder(cellBorders))
                        deferredBorders.Add((row, col, (JsonObject)cellBorders.DeepClone()));
                }

                var fill = cellStyle is not null && cellStyle.ContainsKey("fill")
                    ? ParseHexColor(Json.GetString(cellStyle, "fill"))
                    : isHeader ? headerFill : isFirstColumn ? firstColumnFill : null;
                if (fill is not null && !ApplyCellFill(hwp, fill.Value))
                    return new TableInsertResult(false, rowCount, colCount, expectedTexts.Count, styledCells,
                        tableCountBefore, CountTableControls(hwp), expectedTexts);

                if ((verticalCenter || (cellStyle is not null && Json.GetBool(cellStyle, "verticalCenter"))) &&
                    !CenterCurrentCell(hwp))
                    return new TableInsertResult(false, rowCount, colCount, expectedTexts.Count, styledCells,
                        tableCountBefore, CountTableControls(hwp), expectedTexts);

                if (!ExecInsertText(hwp, rows[row][col]))
                    return new TableInsertResult(false, rowCount, colCount, expectedTexts.Count, styledCells,
                        tableCountBefore, CountTableControls(hwp), expectedTexts);
                if (!string.IsNullOrEmpty(rows[row][col])) expectedTexts.Add(rows[row][col]);
                styledCells++;

                var isLast = row == rowCount - 1 && col == colCount - 1;
                if (!isLast && !(bool)hwp.HAction.Run("TableRightCell"))
                    return new TableInsertResult(false, rowCount, colCount, expectedTexts.Count, styledCells,
                        tableCountBefore, CountTableControls(hwp), expectedTexts);
            }
        }

        if (!ApplyDeferredCellBorders(hwp, deferredBorders, colCount))
            return new TableInsertResult(false, rowCount, colCount, expectedTexts.Count, styledCells,
                tableCountBefore, CountTableControls(hwp), expectedTexts);

        if (!ApplyHorizontalMerges(hwp, merges, colCount))
            return new TableInsertResult(false, rowCount, colCount, expectedTexts.Count, styledCells,
                tableCountBefore, CountTableControls(hwp), expectedTexts);

        try { hwp.HAction.Run("Cancel"); } catch { }
        try { hwp.HAction.Run("MoveDocEnd"); } catch { }
        return new TableInsertResult(true, rowCount, colCount, rowCount * colCount, styledCells,
            tableCountBefore, CountTableControls(hwp), expectedTexts);
    }

    private static bool ApplyCharShape(dynamic hwp, JsonObject style)
    {
        dynamic act = hwp.HAction;
        dynamic ps = hwp.HParameterSet.HCharShape;
        act.GetDefault("CharShape", ps.HSet);
        if (style.TryGetPropertyValue("bold", out var b)) ps.Bold = b!.GetValue<bool>();
        if (style.TryGetPropertyValue("italic", out var it)) ps.Italic = it!.GetValue<bool>();
        if (TryJsonNumber(style, "fontSize", out var fontSize))
        {
            if (fontSize is < 1 or > 4096) throw new ArgumentOutOfRangeException(nameof(fontSize), "fontSize는 1~4096pt입니다");
            ps.Height = Convert.ToInt32(Math.Round(fontSize * 100.0)); // 한글 Height 단위: 1pt = 100
        }
        if (Json.GetString(style, "fontName") is { Length: > 0 } fontName)
        {
            ps.FaceNameHangul = fontName;
            ps.FaceNameLatin = fontName;
            ps.FaceNameHanja = fontName;
            ps.FaceNameJapanese = fontName;
            ps.FaceNameOther = fontName;
            ps.FaceNameSymbol = fontName;
            ps.FaceNameUser = fontName;
        }
        if (Json.GetString(style, "textColor") is { Length: > 0 } textColor)
            ps.TextColor = ToHwpColorRef(textColor);
        if (Json.GetString(style, "shadeColor") is { Length: > 0 } shadeColor)
            ps.ShadeColor = ToHwpColorRef(shadeColor);
        if (style.TryGetPropertyValue("underline", out var underline) && underline is not null)
        {
            ps.UnderlineType = underline is JsonValue uv && uv.TryGetValue<bool>(out var enabled)
                ? enabled ? 1 : 0
                : UnderlineType(Json.GetString(style, "underline"));
        }
        if (Json.GetString(style, "underlineColor") is { Length: > 0 } underlineColor)
            ps.UnderlineColor = ToHwpColorRef(underlineColor);
        if (style.TryGetPropertyValue("strikeout", out var strikeout) && strikeout is not null)
        {
            ps.StrikeOutType = strikeout is JsonValue sv && sv.TryGetValue<bool>(out var enabled)
                ? enabled ? 3 : 0
                : StrikeOutType(Json.GetString(style, "strikeout"));
        }
        if (Json.GetString(style, "strikeoutColor") is { Length: > 0 } strikeoutColor)
            ps.StrikeOutColor = ToHwpColorRef(strikeoutColor);
        if (style.TryGetPropertyValue("letterSpacing", out var spacing) && spacing is not null)
            SetAllLanguageCharShapeFields(ps, "Spacing", spacing.GetValue<int>());
        if (style.TryGetPropertyValue("widthRatio", out var ratio) && ratio is not null)
            SetAllLanguageCharShapeFields(ps, "Ratio", ratio.GetValue<int>());
        if (style.TryGetPropertyValue("offset", out var offset) && offset is not null)
            SetAllLanguageCharShapeFields(ps, "Offset", offset.GetValue<int>());
        if (style.TryGetPropertyValue("superscript", out var superscript) && superscript is not null)
            ps.SuperScript = superscript.GetValue<bool>();
        if (style.TryGetPropertyValue("subscript", out var subscript) && subscript is not null)
            ps.SubScript = subscript.GetValue<bool>();
        return (bool)act.Execute("CharShape", ps.HSet);
    }

    private static bool ApplyParagraphAlignment(dynamic hwp, JsonObject style)
    {
        var align = Json.GetString(style, "align")?.ToLowerInvariant();
        if (string.IsNullOrEmpty(align)) return true;
        var action = align switch
        {
            "left" => "ParagraphShapeAlignLeft",
            "center" => "ParagraphShapeAlignCenter",
            "right" => "ParagraphShapeAlignRight",
            "justify" => "ParagraphShapeAlignJustify",
            "distribute" => "ParagraphShapeAlignDistribute",
            "division" => "ParagraphShapeAlignDivision",
            _ => null,
        };
        return action is not null && (bool)hwp.HAction.Run(action);
    }

    /// <summary>
    /// 특정 문자열을 아래쪽으로 찾아 선택한 뒤 글자 모양을 적용한다.
    /// RepeatFind는 찾은 문자열을 선택 상태로 두므로 CharShape를 정확한 범위에 적용할 수 있다.
    /// </summary>
    private static int ApplyCharShapeToTextMatches(dynamic hwp, string targetText, JsonObject style)
    {
        if (string.IsNullOrEmpty(targetText)) return 0;
        var count = 0;
        int? previousMessageMode = null;
        try
        {
            try { previousMessageMode = Convert.ToInt32(hwp.GetMessageBoxMode()); } catch { }
            try { hwp.SetMessageBoxMode(0x2FFF1); } catch { }
            hwp.HAction.Run("MoveDocBegin");

            dynamic act = hwp.HAction;
            dynamic find = hwp.HParameterSet.HFindReplace;
            act.GetDefault("FindDlg", find.HSet);
            _ = act.Execute("FindDlg", find.HSet);
            find = hwp.HParameterSet.HFindReplace;
            try { find.MatchCase = 1; } catch { }
            try { find.SeveralWords = 0; } catch { }
            try { find.UseWildCards = 0; } catch { }
            try { find.WholeWordOnly = 0; } catch { }
            try { find.AutoSpell = 0; } catch { }
            try { find.Direction = hwp.FindDir("Forward"); } catch { }
            find.FindString = targetText;
            try { find.IgnoreMessage = 1; } catch { }
            try { find.HanjaFromHangul = 0; } catch { }
            try { find.AllWordForms = 0; } catch { }
            try { find.FindJaso = 0; } catch { }
            try { find.FindRegExp = 0; } catch { }
            try { find.FindType = 1; } catch { }

            while (count < 1000 && (bool)act.Execute("RepeatFind", find.HSet))
            {
                if (!ApplyCharShape(hwp, style) || !ApplyParagraphAlignment(hwp, style)) break;
                count++;
            }
        }
        finally
        {
            try { hwp.HAction.Run("Cancel"); } catch { }
            try { hwp.HAction.Run("MoveDocBegin"); } catch { }
            try { hwp.SetMessageBoxMode(previousMessageMode ?? 0xFFFFF); } catch { }
        }
        return count;
    }

    // ---------- IAppAdapter ----------

    public override JsonObject GetCapabilities() => new()
    {
        ["app"] = App,
        ["automation"] = "hwp-automation-com",
        ["directAppControl"] = true,
        ["connectsToExistingWindow"] = true,
        ["enumeratesOpenDocuments"] = true,
        ["liveDocumentTarget"] = "documentRef",
        ["explicitNewDocumentLaunch"] = true,
        ["fileAutomation"] = true,
        ["usesUiAutomation"] = false,
        ["usesExternalMacro"] = false,
        ["creationPolicy"] = new JsonObject
        {
            ["version"] = HwpCreationPolicy.PolicyVersion,
            ["planningTool"] = "hwp_plan_creation",
            ["defaultNewDocumentMode"] = "docx-first",
            ["nativeModeConditions"] = new JsonArray(
                "existing-hwp-or-hwpx",
                "existing-hwp-template",
                "native-fields",
                "hwp-only-objects",
                "complex-merged-tables",
                "preserve-original-layout",
                "docx-generator-unavailable"),
            ["wordComRequired"] = false,
        },
        ["interactionPolicy"] = new JsonObject
        {
            ["mode"] = "preserve-foreground",
            ["backgroundInactiveWindow"] = true,
            ["restoresOriginalDocument"] = true,
            ["restoresCaretAndSelection"] = true,
            ["concurrentTargetInput"] = "stop-after-current-operation",
            ["sameDocumentConcurrentEditing"] = false,
        },
        ["readOps"] = new JsonArray("launch", "context", "text", "selection", "bundle", "document_map", "structure", "fields", "tables", "doctor"),
        ["writeOps"] = new JsonArray(
            "insert_text", "append_text", "insert_before_text", "insert_after_text",
            "replace_document_text", "replace_selection", "find_replace",
            "set_paragraph_style_basic", "set_paragraph_format", "format_paragraphs", "set_page_setup", "insert_break",
            "insert_table", "table_cell_set_text", "table_set_cells", "insert_picture", "insert_page_number",
            "set_header_footer_text", "table_insert_rows", "table_insert_columns",
            "table_delete_rows", "table_delete_columns", "table_merge_cells", "table_set_row_height", "table_set_row_heights",
            "set_field_text", "export_pdf"),
        ["limits"] = new JsonObject
        {
            ["maxReadChars"] = MaxChars,
            ["maxDiffEntries"] = MaxDiff,
            ["comTimeoutSec"] = (int)ComTimeout.TotalSeconds,
            ["expensivePageCountIsOptIn"] = true,
        },
        ["safety"] = new JsonArray("dry-run", "snapshot", "confirm-token", "readback", "automatic-rollback"),
    };

    public override AdapterStatus GetStatus()
    {
        var installed = _appFactory is not null || Type.GetTypeFromProgID("HWPFrame.HwpObject") is not null;
        if (!installed)
            return new AdapterStatus(false, false, "hwp", null, null,
                "HWPFrame.HwpObject가 등록되어 있지 않습니다");

        try
        {
            return ComInvoke(() =>
            {
                var app = AttachHwp(allowCreate: false);
                if (app is null)
                    return new AdapterStatus(true, false, "hwp", null, null,
                        "한글 자동화 API가 설치되어 있지만 실행 중인 한글 창에는 연결되지 않았습니다");
                dynamic d = app;
                string? doc = null;
                try
                {
                    dynamic? active = ActiveDoc(d);
                    if (active is not null)
                    {
                        var fullName = (string)(active.FullName ?? "");
                        var documentId = active.DocumentID?.ToString() ?? "";
                        var windowHandle = RotHelper.HwpWindowHandle(app);
                        doc = HwpDocumentRef(fullName, documentId, windowHandle,
                            RotHelper.ProcessIdFromWindowHandle(windowHandle));
                    }
                }
                catch { }
                return new AdapterStatus(true, true, "hwp", null, doc,
                    _connectionMode == "existing-window"
                        ? "사용자가 열어 둔 한글 창에 연결됨"
                        : "doc-bridge 전용 한글 인스턴스에 연결됨");
            });
        }
        catch (Exception ex) { return new AdapterStatus(true, false, "hwp", null, null, ex.Message); }
    }

    public override ContextResult GetActiveContext()
    {
        return ComInvoke(() =>
        {
            var r = new ContextResult { App = App };
            var foreground = new ForegroundInteractionGuard(App);
            try
            {
                var app = AttachHwp(allowCreate: false);
                if (app is not null)
                    TrackHwpInteraction(app, foreground, documentState: null, captureTarget: false);
                if (app is null) { r.Errors.Add("한글이 실행 중이지 않습니다. 한글을 열고 문서를 표시한 뒤 다시 시도하세요."); return r; }
                dynamic hwp = app;
                var doc = ActiveDoc(hwp);
                if (doc is null) { r.Errors.Add("열린 한글 문서가 없습니다."); return r; }

                r.Ok = true;
                string fullName = "";
                try { fullName = (string)(doc.FullName ?? ""); } catch { }
                var docId = "";
                try { docId = doc.DocumentID.ToString(); } catch { }
                var windowHandle = RotHelper.HwpWindowHandle(app);
                var processId = RotHelper.ProcessIdFromWindowHandle(windowHandle);
                r.DocumentRef = HwpDocumentRef(fullName, docId, windowHandle, processId);

                string text = GetDocText(hwp);
                r.Summary["documentId"] = docId;
                r.Summary["fullName"] = fullName;
                r.Summary["connectionMode"] = _connectionMode;
                r.Summary["windowHandle"] = windowHandle.ToString();
                r.Summary["processId"] = processId;
                r.Summary["instanceRef"] = HwpInstanceRef(docId, windowHandle, processId);
                var openDocuments = InspectOpenHwpDocuments(windowHandle);
                r.Summary["openDocuments"] = openDocuments;
                r.Summary["openDocumentCount"] = openDocuments.Count;
                r.Summary["duplicatePathCount"] = openDocuments.Count(node =>
                    node is JsonObject item && Json.GetBool(item, "duplicatePath"));
                try { r.Summary["format"] = (string?)doc.Format?.ToString(); } catch { }
                try { r.Summary["modified"] = (bool)doc.Modified; } catch { }
                try { r.Summary["editMode"] = (int)doc.EditMode; } catch { }
                r.Summary["textLength"] = text.Length;
                r.Summary["textPreview"] = text[..Math.Min(200, text.Length)];

                string selText = GetSelectionText(hwp);
                r.Selection = new JsonObject
                {
                    ["hasSelection"] = selText.Length > 0,
                    ["selectionLength"] = selText.Length,
                    ["selectionPreview"] = selText[..Math.Min(200, selText.Length)],
                };
            }
            catch (Exception ex) { r.Errors.Add($"hwp context failed: {ex.Message}"); }
            finally { r.Interaction = foreground.Complete(); }
            return r;
        });
    }

    public override JsonObject Read(JsonObject args)
    {
        return ComInvoke(() =>
        {
            var foreground = new ForegroundInteractionGuard(App);
            var documentState = new HwpInteractionState();
            try
            {
                var file = Json.GetString(args, "file");
                var documentRef = Json.GetString(args, "documentRef");
                var app = AttachHwpForTarget(file, documentRef, allowCreate: file is not null,
                    foreground, documentState);
                if (app is null) return Json.ErrorResult(
                    file is null && documentRef is null
                        ? "사용자가 열어 둔 한글 문서를 찾지 못했습니다. 한글에서 문서를 연 뒤 다시 시도하세요. DocBridge는 빈 한글 창을 자동 실행하지 않습니다."
                        : "한글 자동화 인스턴스를 시작할 수 없습니다",
                    App);
                dynamic hwp = app;
                if (OpenOrGetDoc(hwp, file) is null)
                    return Json.ErrorResult(file is not null ? $"파일을 열 수 없습니다: {file}" : "열린 한글 문서가 없습니다", App);

                if (!_closeTargetWhenDone) documentState.CaptureTarget(app);
                var scope = Json.GetString(args, "scope") ?? "selection";
                var maxChars = Math.Min(Json.GetInt(args, "maxChars") ?? MaxChars, MaxChars);

                if (scope == "bundle")
                {
                    var requestedSections = Json.GetArr(args, "sections")?
                        .Select(node => node?.GetValue<string>() ?? "")
                        .Where(value => value.Length > 0)
                        .ToHashSet(StringComparer.OrdinalIgnoreCase)
                        ?? new HashSet<string>(new[] { "text", "document_map", "structure" }, StringComparer.OrdinalIgnoreCase);
                    var allowedSections = new HashSet<string>(
                        new[] { "text", "document_map", "structure", "fields", "tables" },
                        StringComparer.OrdinalIgnoreCase);
                    var unknown = requestedSections.Where(section => !allowedSections.Contains(section)).ToArray();
                    if (unknown.Length > 0)
                        throw new ArgumentException($"unknown bundle section(s): {string.Join(", ", unknown)}");

                    var bundle = new JsonObject();
                    var sectionTimings = new JsonObject();
                    string? normalizedText = null;
                    string BundleText()
                    {
                        if (normalizedText is null)
                        {
                            var started = Stopwatch.StartNew();
                            normalizedText = NormalizeNewlines(GetDocText(hwp));
                            started.Stop();
                            sectionTimings["textReadMs"] = started.ElapsedMilliseconds;
                        }
                        return normalizedText;
                    }

                    if (requestedSections.Contains("text"))
                    {
                        var bundleText = BundleText();
                        bundle["text"] = bundleText[..Math.Min(maxChars, bundleText.Length)];
                        bundle["textLength"] = bundleText.Length;
                        bundle["textTruncated"] = bundleText.Length > maxChars;
                    }
                    if (requestedSections.Contains("document_map"))
                    {
                        var started = Stopwatch.StartNew();
                        bundle["documentMap"] = BuildDocumentMapFromText(
                            hwp, BundleText(),
                            Math.Max(0, Json.GetInt(args, "startParagraph") ?? 0),
                            Math.Clamp(Json.GetInt(args, "maxParagraphs") ?? 500, 1, 2000));
                        started.Stop();
                        sectionTimings["documentMapMs"] = started.ElapsedMilliseconds;
                    }
                    if (requestedSections.Contains("structure"))
                    {
                        var started = Stopwatch.StartNew();
                        bundle["structure"] = InspectStructure(
                            hwp,
                            Math.Clamp(Json.GetInt(args, "maxControls") ?? 200, 1, 1000),
                            Json.GetBool(args, "includePageCount"));
                        started.Stop();
                        sectionTimings["structureMs"] = started.ElapsedMilliseconds;
                    }
                    if (requestedSections.Contains("fields"))
                    {
                        var started = Stopwatch.StartNew();
                        bundle["fieldInventory"] = InspectFields(
                            hwp,
                            Math.Clamp(Json.GetInt(args, "maxFields") ?? 100, 1, 500),
                            Json.GetBool(args, "includeValues", true));
                        started.Stop();
                        sectionTimings["fieldsMs"] = started.ElapsedMilliseconds;
                    }
                    if (requestedSections.Contains("tables"))
                    {
                        var started = Stopwatch.StartNew();
                        bundle["tableInventory"] = InspectTables(
                            (object)hwp,
                            Json.GetInt(args, "tableIndex"),
                            Math.Clamp(Json.GetInt(args, "maxCells") ?? 100, 1, 1000),
                            Json.GetBool(args, "includeStyles", false));
                        started.Stop();
                        sectionTimings["tablesMs"] = started.ElapsedMilliseconds;
                    }
                    CloseTargetIfNeeded(hwp, file);
                    return new JsonObject
                    {
                        ["ok"] = true,
                        ["app"] = App,
                        ["scope"] = scope,
                        ["file"] = file,
                        ["documentRef"] = Json.GetString(CaptureDocumentIdentity(hwp), "documentRef"),
                        ["sections"] = new JsonArray(requestedSections.OrderBy(value => value).Select(value => (JsonNode?)value).ToArray()),
                        ["bundle"] = bundle,
                        ["timings"] = sectionTimings,
                    };
                }

                if (scope == "document_map")
                {
                    var map = BuildDocumentMap(hwp,
                        Math.Max(0, Json.GetInt(args, "startParagraph") ?? 0),
                        Math.Clamp(Json.GetInt(args, "maxParagraphs") ?? 500, 1, 2000));
                    CloseTargetIfNeeded(hwp, file);
                    return new JsonObject
                    {
                        ["ok"] = true,
                        ["app"] = App,
                        ["scope"] = scope,
                        ["file"] = file,
                        ["documentRef"] = Json.GetString(CaptureDocumentIdentity(hwp), "documentRef"),
                        ["documentMap"] = map,
                    };
                }

                if (scope == "structure")
                {
                    var maxControls = Math.Clamp(Json.GetInt(args, "maxControls") ?? 200, 1, 1000);
                    var structure = InspectStructure(hwp, maxControls, Json.GetBool(args, "includePageCount"));
                    CloseTargetIfNeeded(hwp, file);
                    return new JsonObject
                    {
                        ["ok"] = true,
                        ["app"] = App,
                        ["scope"] = scope,
                        ["file"] = file,
                        ["documentRef"] = Json.GetString(CaptureDocumentIdentity(hwp), "documentRef"),
                        ["structure"] = structure,
                    };
                }

                if (scope == "fields")
                {
                    var fields = InspectFields(hwp, Math.Clamp(Json.GetInt(args, "maxFields") ?? 100, 1, 500),
                        Json.GetBool(args, "includeValues", true));
                    CloseTargetIfNeeded(hwp, file);
                    return new JsonObject
                    {
                        ["ok"] = true,
                        ["app"] = App,
                        ["scope"] = scope,
                        ["file"] = file,
                        ["documentRef"] = Json.GetString(CaptureDocumentIdentity(hwp), "documentRef"),
                        ["fieldInventory"] = fields,
                    };
                }

                if (scope == "tables")
                {
                    var tableIndex = Json.GetInt(args, "tableIndex");
                    var maxCells = Math.Clamp(Json.GetInt(args, "maxCells") ?? 100, 1, 1000);
                    var tables = InspectTables((object)hwp, tableIndex, maxCells, Json.GetBool(args, "includeStyles", true));
                    CloseTargetIfNeeded(hwp, file);
                    return new JsonObject
                    {
                        ["ok"] = true,
                        ["app"] = App,
                        ["scope"] = scope,
                        ["file"] = file,
                        ["documentRef"] = Json.GetString(CaptureDocumentIdentity(hwp), "documentRef"),
                        ["tableInventory"] = tables,
                    };
                }

                string text = scope switch
                {
                    "selection" => GetSelectionText(hwp),
                    "document" => GetDocText(hwp),
                    "paragraph" => throw new NotSupportedException(
                        "paragraph scope는 MVP에서 지원하지 않습니다 (현재 문단 위치 API 미확정). scope=selection|document를 사용하세요."),
                    _ => throw new ArgumentException($"unknown scope '{scope}' (selection|document|bundle|document_map|structure|fields|tables)"),
                };

                CloseTargetIfNeeded(hwp, file); // DocBridge가 직접 연 파일만 잠금 해제
                return new JsonObject
                {
                    ["ok"] = true,
                    ["app"] = App,
                    ["scope"] = scope,
                    ["file"] = file,
                    ["documentRef"] = Json.GetString(CaptureDocumentIdentity(hwp), "documentRef"),
                    ["text"] = text[..Math.Min(maxChars, text.Length)],
                    ["length"] = text.Length,
                    ["truncated"] = text.Length > maxChars,
                };
            }
            catch (HwpAutomationException ex) { return ex.ToResult(App); }
            catch (Exception ex) { return Json.ErrorResult($"hwp_read_text failed: {ex.Message}", App); }
            finally { _ = CompleteHwpInteraction(foreground, documentState); }
        });
    }

    // ---------- preview ----------

    public override ApplyPreview Preview(IReadOnlyList<JsonObject> ops)
    {
        return ComInvoke(() =>
        {
            var p = new ApplyPreview();
            var foreground = new ForegroundInteractionGuard(App);
            var documentState = new HwpInteractionState();
            try
            {
                var targetFile = FileArgOf(ops);
                var targetDocumentRef = DocumentRefArgOf(ops);
                var selectorError = HwpTargetSelectorError(ops);
                if (selectorError is not null)
                {
                    p.Errors.Add(selectorError);
                    return p;
                }
                var app = AttachHwpForTarget(
                    targetFile, targetDocumentRef, allowCreate: targetFile is not null,
                    foreground, documentState);
                if (app is null)
                {
                    p.Errors.Add(targetFile is null && targetDocumentRef is null
                        ? "사용자가 열어 둔 한글 문서를 찾지 못했습니다. DocBridge는 빈 한글 창을 자동 실행하지 않습니다."
                        : "한글 자동화 인스턴스를 시작할 수 없습니다");
                    return p;
                }
                dynamic hwp = app;
                if (OpenOrGetDoc(hwp, targetFile) is null)
                {
                    p.Errors.Add(targetFile is null
                        ? "연결된 한글 창에 열린 문서가 없습니다"
                        : "file 지정 문서를 열 수 없습니다");
                    return p;
                }
                if (!_closeTargetWhenDone) documentState.CaptureTarget(app);
                if (targetFile is null)
                {
                    KeepOwnedLiveDocumentOpen(hwp);
                    p.Warnings.Add(targetDocumentRef is null
                        ? "사용자가 열어 둔 현재 한글 문서를 직접 편집합니다. 적용 후 문서는 열린 상태로 유지되며 자동 저장하지 않습니다."
                        : $"openDocuments의 documentRef '{targetDocumentRef}' 문서를 직접 편집합니다. 적용 후 열린 상태로 유지되며 자동 저장하지 않습니다.");
                }

                // 한 batch의 preview에서 전체 TEXT와 표 control 목록을 반복 조회하지 않는다.
                // GetTextFile("TEXT")와 HeadCtrl 순회는 긴 문서에서 가장 큰 preview 병목이다.
                var previewDocumentText = NormalizeNewlines(GetDocText(hwp));
                string PreviewDocumentText() => previewDocumentText;
                var previewTableCount = CountControlId((object)hwp, "tbl");
                var previewTableDimensions = new Dictionary<int, HwpTableDimensions?>();
                bool PreviewTableExists(int index) => index >= 0 && index < previewTableCount;
                HwpTableDimensions? PreviewTableDimensions(int index)
                {
                    if (!PreviewTableExists(index)) return null;
                    if (!previewTableDimensions.TryGetValue(index, out var dimensions))
                    {
                        dimensions = TryReadTableDimensions((object)hwp, index);
                        previewTableDimensions[index] = dimensions;
                    }
                    return dimensions;
                }

                foreach (var op in ops)
                {
                    var name = Json.GetString(op, "op")!;
                    switch (name)
                    {
                        case "insert_text":
                        {
                            var text = Json.GetString(op, "text")!;
                            p.Affected.Add(new AffectedRef("cursor", $"insert {text.Length} chars"));
                            p.Diff.Add(new DiffEntry { Ref = "cursor", Before = "", After = text[..Math.Min(100, text.Length)] });
                            // 실제 커서 위치는 가상 문서에서 알 수 없지만 후속 op가 방금 쓴
                            // 문구를 기준으로 삼는 일반적인 batch는 순차 검증할 수 있어야 한다.
                            previewDocumentText = SimulatePreviewAppend(previewDocumentText, text, false);
                            break;
                        }
                        case "append_text":
                        {
                            var text = Json.GetString(op, "text")!;
                            p.Affected.Add(new AffectedRef("document-end", $"append {text.Length} chars as real paragraphs"));
                            p.Diff.Add(new DiffEntry { Ref = "document:end", Before = "current end", After = text[..Math.Min(100, text.Length)] });
                            previewDocumentText = SimulatePreviewAppend(
                                previewDocumentText, text, Json.GetBool(op, "startNewParagraph", true));
                            break;
                        }
                        case "insert_before_text":
                        case "insert_after_text":
                        {
                            var anchor = Json.GetString(op, "anchor")!;
                            var text = Json.GetString(op, "text")!;
                            var matchCase = Json.GetBool(op, "matchCase", true);
                            var mode = (Json.GetString(op, "mode") ?? "paragraph").ToLowerInvariant();
                            if (mode is not ("inline" or "paragraph"))
                            {
                                p.Errors.Add($"{name}.mode must be 'inline' or 'paragraph'");
                                break;
                            }
                            var document = NormalizeNewlines(PreviewDocumentText());
                            var matchCount = CountTextOccurrences(document, anchor, matchCase);
                            var occurrence = Json.GetInt(op, "occurrence") ?? 1;
                            if (matchCount == 0)
                            {
                                p.Errors.Add($"{name}: 기준 문구를 찾지 못했습니다: '{anchor}'");
                                break;
                            }
                            if (!op.ContainsKey("occurrence") && matchCount != 1)
                            {
                                p.Errors.Add($"{name}: 기준 문구가 {matchCount}개입니다. occurrence(1부터 시작)를 명시하세요.");
                                break;
                            }
                            if (occurrence < 1 || occurrence > matchCount)
                            {
                                p.Errors.Add($"{name}: occurrence={occurrence}, 유효 범위=1..{matchCount}");
                                break;
                            }
                            var relation = name == "insert_before_text" ? "before" : "after";
                            p.Affected.Add(new AffectedRef($"anchor:{anchor}",
                                $"insert {relation}, occurrence {occurrence}/{matchCount}, mode={mode}, compare anchor+previous+next paragraph styles"));
                            p.Diff.Add(new DiffEntry
                            {
                                Ref = $"anchor:{anchor}",
                                Before = anchor,
                                After = name == "insert_before_text" ? $"{text}{anchor}" : $"{anchor}{text}",
                            });
                            previewDocumentText = SimulatePreviewRelativeInsert(
                                previewDocumentText, anchor, text, occurrence, matchCase,
                                name == "insert_before_text", mode);
                            break;
                        }
                        case "replace_document_text":
                        {
                            string before = PreviewDocumentText();
                            var text = Json.GetString(op, "text")!;
                            p.Affected.Add(new AffectedRef("document", $"replace {before.Length} chars with {text.Length} chars"));
                            p.Diff.Add(new DiffEntry
                            {
                                Ref = "document",
                                Before = before[..Math.Min(100, before.Length)],
                                After = text[..Math.Min(100, text.Length)],
                            });
                            previewDocumentText = NormalizeNewlines(text);
                            break;
                        }
                        case "replace_selection":
                        {
                            string sel = GetSelectionText(hwp);
                            var text = Json.GetString(op, "text")!;
                            if (sel.Length == 0)
                                p.Warnings.Add("현재 선택 영역이 없습니다. replace_selection은 커서 위치 삽입으로 동작합니다.");
                            p.Affected.Add(new AffectedRef("selection", $"{sel.Length} chars"));
                            p.Diff.Add(new DiffEntry
                            {
                                Ref = "selection",
                                Before = sel[..Math.Min(100, sel.Length)],
                                After = text[..Math.Min(100, text.Length)],
                            });
                            if (sel.Length > 0)
                            {
                                var selectionIndex = previewDocumentText.IndexOf(sel, StringComparison.Ordinal);
                                previewDocumentText = selectionIndex >= 0
                                    ? previewDocumentText.Remove(selectionIndex, sel.Length).Insert(selectionIndex, text)
                                    : SimulatePreviewAppend(previewDocumentText, text, false);
                            }
                            else previewDocumentText = SimulatePreviewAppend(previewDocumentText, text, false);
                            break;
                        }
                        case "find_replace":
                        {
                            var simulation = PreviewFindReplace(hwp, op, PreviewDocumentText(), p);
                            var scope = FindReplaceScope(op);
                            if (scope is null || !scope.ContainsKey("tableIndex"))
                                previewDocumentText = simulation.After;
                            break;
                        }
                        case "set_paragraph_style_basic":
                        {
                            var style = Json.GetObj(op, "style") ?? new JsonObject();
                            ValidateCharacterStyle(style);
                            var target = Json.GetObj(op, "target");
                            var targetText = Json.GetString(target, "text");
                            var scope = Json.GetString(target, "scope") ?? "selection";
                            if (!string.IsNullOrEmpty(targetText))
                            {
                                string doc = PreviewDocumentText();
                                var (count, _) = CountMatches(doc, targetText, targetText, 0);
                                if (count == 0) p.Warnings.Add($"서식 대상 문구가 없습니다: '{targetText}'");
                                p.Affected.Add(new AffectedRef("charShape", $"{count} match(es): {targetText}"));
                                p.Diff.Add(new DiffEntry { Ref = $"style:{targetText}", Before = "current", After = style.DeepClone() });
                            }
                            else
                            {
                                p.Affected.Add(new AffectedRef("charShape", scope == "document" ? "whole document" : "current selection/position"));
                                p.Diff.Add(new DiffEntry { Ref = $"style:{scope}", Before = "current", After = style.DeepClone() });
                            }
                            break;
                        }
                        case "set_paragraph_format":
                        {
                            var style = Json.GetObj(op, "style") ?? throw new ArgumentException("set_paragraph_format.style이 필요합니다");
                            ValidateParagraphStyle(style);
                            var target = Json.GetObj(op, "target");
                            var targetText = Json.GetString(target, "text");
                            var scope = Json.GetString(target, "scope") ?? "selection";
                            if (!string.IsNullOrEmpty(targetText))
                            {
                                string docText = PreviewDocumentText();
                                var (count, _) = CountMatches(docText, targetText, targetText, 0);
                                if (count == 0) p.Warnings.Add($"문단 서식 대상 문구가 없습니다: '{targetText}'");
                                p.Affected.Add(new AffectedRef("paragraphShape", $"{count} match(es): {targetText}"));
                            }
                            else
                            {
                                p.Affected.Add(new AffectedRef("paragraphShape", scope));
                            }
                            p.Diff.Add(new DiffEntry { Ref = $"paragraph:{targetText ?? scope}", Before = "current", After = style.DeepClone() });
                            break;
                        }
                        case "format_paragraphs":
                        {
                            ValidateFormatParagraphItems(op);
                            foreach (var node in Json.GetArr(op, "items")!)
                            {
                                var item = (JsonObject)node!;
                                var target = Json.GetObj(item, "target");
                                var targetText = Json.GetString(target, "text");
                                var scope = Json.GetString(target, "scope") ?? "selection";
                                var count = string.IsNullOrEmpty(targetText)
                                    ? 1
                                    : CountMatches(PreviewDocumentText(), targetText, targetText, 0).Count;
                                if (count == 0) p.Warnings.Add($"묶음 서식 대상 문구가 없습니다: '{targetText}'");
                                var reference = !string.IsNullOrEmpty(targetText) ? $"format:{targetText}" : $"format:{scope}";
                                p.Affected.Add(new AffectedRef(reference, $"combined character+paragraph format, {count} target(s)"));
                                p.Diff.Add(new DiffEntry
                                {
                                    Ref = reference,
                                    Before = "current",
                                    After = new JsonObject
                                    {
                                        ["characterStyle"] = Json.GetObj(item, "characterStyle")?.DeepClone(),
                                        ["paragraphStyle"] = Json.GetObj(item, "paragraphStyle")?.DeepClone(),
                                    },
                                });
                            }
                            break;
                        }
                        case "set_page_setup":
                        {
                            ValidatePageSetup(op);
                            var page = Json.GetObj(op, "page") ?? throw new ArgumentException("set_page_setup.page가 필요합니다");
                            p.Affected.Add(new AffectedRef("page-setup", Json.GetString(op, "applyTo") ?? "current-section"));
                            p.Diff.Add(new DiffEntry { Ref = "page-setup", Before = "current", After = page.DeepClone() });
                            break;
                        }
                        case "insert_break":
                        {
                            ValidateBreak(op);
                            var type = Json.GetString(op, "type") ?? "page";
                            p.Affected.Add(new AffectedRef($"break:{type}", "current cursor"));
                            p.Diff.Add(new DiffEntry { Ref = $"break:{type}", Before = "none", After = type });
                            break;
                        }
                        case "insert_table":
                        {
                            var rows = ParseTableRows(op);
                            _ = ParseColumnWidths(op, rows[0].Count);
                            _ = ParseHexColor(Json.GetString(op, "headerFill"));
                            _ = ParseHexColor(Json.GetString(op, "firstColumnFill"));
                            _ = ParseCellStyles(op, rows.Count, rows[0].Count);
                            var merges = ParseMergeCells(op, rows.Count, rows[0].Count);
                            var previewRows = new JsonArray();
                            foreach (var row in rows.Take(4))
                            {
                                var previewRow = new JsonArray();
                                foreach (var cell in row) previewRow.Add(cell);
                                previewRows.Add(previewRow);
                            }
                            p.Affected.Add(new AffectedRef("table",
                                $"{rows.Count}x{rows[0].Count}, {rows.Count * rows[0].Count} cells, {merges.Count} merge(s)"));
                            p.Diff.Add(new DiffEntry
                            {
                                Ref = "table:new",
                                Before = "no table",
                                After = previewRows,
                            });
                            var newTableIndex = previewTableCount++;
                            previewTableDimensions[newTableIndex] =
                                new HwpTableDimensions(rows.Count, rows[0].Count);
                            foreach (var cellText in rows.SelectMany(row => row)
                                         .Where(cellText => !string.IsNullOrEmpty(cellText)))
                                previewDocumentText = SimulatePreviewAppend(
                                    previewDocumentText, cellText, true);
                            break;
                        }
                        case "table_cell_set_text":
                        {
                            var tableIndex = Json.GetInt(op, "tableIndex") ?? 0;
                            var cellIndex = Json.GetInt(op, "cellIndex");
                            var row = Json.GetInt(op, "row") ?? 0;
                            var col = Json.GetInt(op, "col") ?? 0;
                            var text = Json.GetString(op, "text") ?? "";
                            if (tableIndex < 0 || row < 0 || col < 0 || cellIndex < 0)
                                throw new ArgumentException("tableIndex, row, col, cellIndex는 0 이상이어야 합니다");
                            if (cellIndex is null && (!op.ContainsKey("row") || !op.ContainsKey("col")))
                                throw new ArgumentException("table_cell_set_text에는 cellIndex 또는 row+col이 필요합니다");
                            if (!PreviewTableExists(tableIndex)) p.Errors.Add($"표 {tableIndex}을 찾을 수 없습니다");
                            else if (PreviewTableDimensions(tableIndex) is { } dimensions && cellIndex is null &&
                                     (row >= dimensions.Rows || col >= dimensions.Columns))
                                p.Errors.Add($"표 {tableIndex}의 유효 셀 범위는 row 0..{dimensions.Rows - 1}, col 0..{dimensions.Columns - 1}입니다");
                            var locator = cellIndex is null ? $"cell:{row},{col}" : $"cellIndex:{cellIndex}";
                            var styleMode = PreserveStyle(op) ? "contextual char+paragraph style" : "style preservation disabled";
                            p.Affected.Add(new AffectedRef($"table:{tableIndex}/{locator}", $"replace text; {styleMode}"));
                            p.Diff.Add(new DiffEntry { Ref = $"table:{tableIndex}/{locator}", Before = "current cell text/style", After = text });
                            previewDocumentText = SimulatePreviewAppend(previewDocumentText, text, true);
                            break;
                        }
                        case "table_set_cells":
                        {
                            var tableIndex = Json.GetInt(op, "tableIndex") ?? 0;
                            var cells = Json.GetArr(op, "cells") ?? throw new ArgumentException("table_set_cells.cells 배열이 필요합니다");
                            if (cells.Count is < 1 or > 500) throw new ArgumentException("table_set_cells.cells는 1~500개입니다");
                            if (!PreviewTableExists(tableIndex)) p.Errors.Add($"표 {tableIndex}을 찾을 수 없습니다");
                            foreach (var node in cells)
                            {
                                if (node is not JsonObject cell) throw new ArgumentException("cells 항목은 객체여야 합니다");
                                var cellIndex = Json.GetInt(cell, "cellIndex");
                                var row = Json.GetInt(cell, "row") ?? 0;
                                var col = Json.GetInt(cell, "col") ?? 0;
                                if (cellIndex is null && (!cell.ContainsKey("row") || !cell.ContainsKey("col")))
                                    throw new ArgumentException("각 cells 항목에는 cellIndex 또는 row+col이 필요합니다");
                                var text = Json.GetString(cell, "text") ?? throw new ArgumentException("각 cells 항목에는 text가 필요합니다");
                                var locator = cellIndex is null ? $"cell:{row},{col}" : $"cellIndex:{cellIndex}";
                                p.Affected.Add(new AffectedRef($"table:{tableIndex}/{locator}", "batched text replacement"));
                                p.Diff.Add(new DiffEntry { Ref = $"table:{tableIndex}/{locator}", Before = "current cell text", After = text });
                                previewDocumentText = SimulatePreviewAppend(previewDocumentText, text, true);
                            }
                            break;
                        }
                        case "table_insert_rows":
                        case "table_insert_columns":
                        {
                            var tableIndex = Json.GetInt(op, "tableIndex") ?? 0;
                            var row = Json.GetInt(op, "row") ?? 0;
                            var col = Json.GetInt(op, "col") ?? 0;
                            var count = Json.GetInt(op, "count") ?? 1;
                            if (count is < 1 or > 20) throw new ArgumentException("count는 1~20입니다");
                            if ((Json.GetString(op, "position") ?? "after").ToLowerInvariant() is not ("before" or "after"))
                                throw new ArgumentException("position은 before|after 중 하나여야 합니다");
                            if (!PreviewTableExists(tableIndex)) p.Errors.Add($"표 {tableIndex}을 찾을 수 없습니다");
                            var unit = name == "table_insert_rows" ? "row" : "column";
                            p.Affected.Add(new AffectedRef($"table:{tableIndex}", $"insert {count} {unit}(s)"));
                            p.Diff.Add(new DiffEntry { Ref = $"table:{tableIndex}", Before = "current structure", After = $"+{count} {unit}(s)" });
                            if (PreviewTableDimensions(tableIndex) is { } dimensions)
                            {
                                if (row >= dimensions.Rows || col >= dimensions.Columns)
                                    p.Errors.Add($"표 {tableIndex}의 유효 셀 범위는 row 0..{dimensions.Rows - 1}, col 0..{dimensions.Columns - 1}입니다");
                                else previewTableDimensions[tableIndex] = name == "table_insert_rows"
                                    ? dimensions with { Rows = dimensions.Rows + count }
                                    : dimensions with { Columns = dimensions.Columns + count };
                            }
                            break;
                        }
                        case "table_delete_rows":
                        case "table_delete_columns":
                        {
                            var tableIndex = Json.GetInt(op, "tableIndex") ?? 0;
                            var row = Json.GetInt(op, "row") ?? 0;
                            var col = Json.GetInt(op, "col") ?? 0;
                            var count = Json.GetInt(op, "count") ?? 1;
                            if (count is < 1 or > 20) throw new ArgumentException("count는 1~20입니다");
                            if (!PreviewTableExists(tableIndex)) p.Errors.Add($"표 {tableIndex}을 찾을 수 없습니다");
                            var unit = name == "table_delete_rows" ? "row" : "column";
                            p.Warnings.Add($"{unit} 삭제는 셀 내용과 서식을 제거합니다. 스냅샷 복원 대상입니다.");
                            p.Affected.Add(new AffectedRef($"table:{tableIndex}", $"delete {count} {unit}(s)"));
                            p.Diff.Add(new DiffEntry { Ref = $"table:{tableIndex}", Before = $"existing {unit}(s)", After = $"-{count} {unit}(s)" });
                            if (PreviewTableDimensions(tableIndex) is { } dimensions)
                            {
                                var available = name == "table_delete_rows"
                                    ? dimensions.Rows - row
                                    : dimensions.Columns - col;
                                var remaining = (name == "table_delete_rows" ? dimensions.Rows : dimensions.Columns) - count;
                                if (row >= dimensions.Rows || col >= dimensions.Columns)
                                    p.Errors.Add($"표 {tableIndex}의 유효 셀 범위는 row 0..{dimensions.Rows - 1}, col 0..{dimensions.Columns - 1}입니다");
                                else if (available < count)
                                    p.Errors.Add($"삭제 시작 위치부터 남은 {unit}은 {available}개입니다");
                                else if (remaining < 1) p.Errors.Add("표에는 최소 한 행과 한 열이 남아야 합니다");
                                else previewTableDimensions[tableIndex] = name == "table_delete_rows"
                                    ? dimensions with { Rows = remaining }
                                    : dimensions with { Columns = remaining };
                            }
                            break;
                        }
                        case "table_merge_cells":
                        {
                            var tableIndex = Json.GetInt(op, "tableIndex") ?? 0;
                            var sr = Json.GetInt(op, "startRow") ?? throw new ArgumentException("startRow가 필요합니다");
                            var sc = Json.GetInt(op, "startCol") ?? throw new ArgumentException("startCol이 필요합니다");
                            var er = Json.GetInt(op, "endRow") ?? throw new ArgumentException("endRow가 필요합니다");
                            var ec = Json.GetInt(op, "endCol") ?? throw new ArgumentException("endCol이 필요합니다");
                            if (er < sr || ec < sc || (er == sr && ec == sc)) throw new ArgumentException("올바른 병합 범위가 필요합니다");
                            if (!PreviewTableExists(tableIndex)) p.Errors.Add($"표 {tableIndex}을 찾을 수 없습니다");
                            p.Affected.Add(new AffectedRef($"table:{tableIndex}", $"merge ({sr},{sc})-({er},{ec})"));
                            break;
                        }
                        case "table_set_row_height":
                        {
                            var tableIndex = Json.GetInt(op, "tableIndex") ?? 0;
                            var row = Json.GetInt(op, "row") ?? throw new ArgumentException("row가 필요합니다");
                            if (!TryJsonNumber(op, "heightMm", out var heightMm) || heightMm is < 4 or > 50)
                                throw new ArgumentOutOfRangeException("heightMm", "heightMm는 4~50mm입니다");
                            if (!PreviewTableExists(tableIndex)) p.Errors.Add($"표 {tableIndex}을 찾을 수 없습니다");
                            else if (PreviewTableDimensions(tableIndex) is { } dimensions && row >= dimensions.Rows)
                                p.Errors.Add($"표 {tableIndex}의 유효 행 범위는 0..{dimensions.Rows - 1}입니다");
                            p.Affected.Add(new AffectedRef($"table:{tableIndex}/row:{row}", $"set height {heightMm:0.00}mm"));
                            p.Diff.Add(new DiffEntry { Ref = $"table:{tableIndex}/row:{row}", Before = "current height", After = $"{heightMm:0.00}mm" });
                            break;
                        }
                        case "table_set_row_heights":
                        {
                            var tableIndex = Json.GetInt(op, "tableIndex") ?? 0;
                            var specs = ParseRowHeightSpecs(op);
                            if (!PreviewTableExists(tableIndex)) p.Errors.Add($"표 {tableIndex}을 찾을 수 없습니다");
                            var dimensions = PreviewTableDimensions(tableIndex);
                            foreach (var spec in specs)
                            {
                                if (dimensions is not null && spec.Row >= dimensions.Rows)
                                    p.Errors.Add($"표 {tableIndex}의 유효 행 범위는 0..{dimensions.Rows - 1}입니다");
                                p.Affected.Add(new AffectedRef($"table:{tableIndex}/row:{spec.Row}",
                                    $"set height {spec.HeightMm:0.00}mm (bulk)"));
                                p.Diff.Add(new DiffEntry
                                {
                                    Ref = $"table:{tableIndex}/row:{spec.Row}",
                                    Before = "current height",
                                    After = $"{spec.HeightMm:0.00}mm",
                                });
                            }
                            break;
                        }
                        case "set_field_text":
                        {
                            var fieldName = Json.GetString(op, "name") ?? throw new ArgumentException("set_field_text.name이 필요합니다");
                            if (!(bool)hwp.FieldExist(fieldName)) p.Errors.Add($"필드를 찾을 수 없습니다: {fieldName}");
                            p.Affected.Add(new AffectedRef($"field:{fieldName}", "replace field text"));
                            p.Diff.Add(new DiffEntry { Ref = $"field:{fieldName}", Before = "current", After = Json.GetString(op, "text") ?? "" });
                            break;
                        }
                        case "insert_picture":
                        {
                            ValidatePicture(op);
                            var path = Path.GetFullPath(Json.GetString(op, "path") ?? throw new ArgumentException("insert_picture.path가 필요합니다"));
                            var tableIndex = Json.GetInt(op, "tableIndex");
                            var row = Json.GetInt(op, "row") ?? 0;
                            var col = Json.GetInt(op, "col") ?? 0;
                            var cellIndex = Json.GetInt(op, "cellIndex");
                            var reference = "picture:new";
                            if (tableIndex is not null)
                            {
                                if (!PreviewTableExists(tableIndex.Value))
                                    p.Errors.Add($"표 {tableIndex}을 찾을 수 없습니다");
                                else if (PreviewTableDimensions(tableIndex.Value) is { } dimensions &&
                                         cellIndex is null && (row >= dimensions.Rows || col >= dimensions.Columns))
                                    p.Errors.Add($"표 {tableIndex}의 유효 셀 범위는 row 0..{dimensions.Rows - 1}, col 0..{dimensions.Columns - 1}입니다");
                                reference = cellIndex is null
                                    ? $"table:{tableIndex}/cell:{row},{col}/picture"
                                    : $"table:{tableIndex}/cellIndex:{cellIndex}/picture";
                            }
                            p.Affected.Add(new AffectedRef(reference, path));
                            p.Diff.Add(new DiffEntry
                            {
                                Ref = reference,
                                Before = Json.GetBool(op, "clearCell") ? "cell content (cleared)" : "existing cell content preserved",
                                After = path,
                            });
                            break;
                        }
                        case "insert_page_number":
                            ValidatePageNumber(op);
                            p.Affected.Add(new AffectedRef("page-number", Json.GetString(op, "position") ?? "bottom-center"));
                            p.Diff.Add(new DiffEntry { Ref = "page-number", Before = "current", After = op.DeepClone() });
                            break;
                        case "set_header_footer_text":
                            ValidateHeaderFooter(op);
                            p.Affected.Add(new AffectedRef(Json.GetString(op, "kind") ?? "header", Json.GetString(op, "pages") ?? "both"));
                            p.Diff.Add(new DiffEntry { Ref = Json.GetString(op, "kind") ?? "header", Before = "current", After = Json.GetString(op, "text") ?? "" });
                            break;
                        case "export_pdf":
                        {
                            ValidateExportPdf(op);
                            var output = Path.GetFullPath(Json.GetString(op, "output") ?? throw new ArgumentException("export_pdf.output이 필요합니다"));
                            if (File.Exists(output)) p.Warnings.Add($"기존 PDF가 교체됩니다: {output}");
                            p.Affected.Add(new AffectedRef("pdf:export", output));
                            p.Diff.Add(new DiffEntry { Ref = "pdf:export", Before = File.Exists(output) ? "existing file" : "none", After = output });
                            break;
                        }
                    }
                    if (!foreground.Checkpoint(stopOnConcurrentInput: true))
                    {
                        p.Errors.Add("[APP_USER_ACTIVITY_DETECTED] 사용자가 한글 창을 조작하여 미리보기를 중단했습니다. 해당 창 작업을 마친 뒤 다시 실행하세요.");
                        break;
                    }
                }
                CloseTargetIfNeeded(hwp, FileArgOf(ops)); // DocBridge가 직접 연 파일만 닫기
            }
            catch (HwpAutomationException ex) { p.Errors.Add($"[{ex.Code}] {ex.Message} {ex.UserAction}"); }
            catch (Exception ex) { p.Errors.Add($"preview failed: {ex.Message}"); }
            finally { p.Interaction = CompleteHwpInteraction(foreground, documentState); }
            return p;
        });
    }

    private static (int Count, List<(string Before, string After)> Samples) CountMatches(
        string doc, string find, string replace, int maxSamples)
    {
        var samples = new List<(string, string)>();
        var count = 0; var idx = 0;
        while ((idx = doc.IndexOf(find, idx, StringComparison.Ordinal)) >= 0)
        {
            count++;
            if (samples.Count < maxSamples)
            {
                var start = Math.Max(0, idx - 10);
                var ctx = doc[start..Math.Min(doc.Length, idx + find.Length + 10)];
                samples.Add((ctx, ctx.Replace(find, replace, StringComparison.Ordinal)));
            }
            idx += find.Length;
        }
        return (count, samples);
    }

    internal static bool FindReplaceReadbackVerified(
        string before, string after, string find, string replace, int expectedMatches, int remainingMatches)
    {
        if (expectedMatches == 0) return string.Equals(before, after, StringComparison.Ordinal);
        if (!replace.Contains(find, StringComparison.Ordinal)) return remainingMatches == 0;
        if (string.Equals(before, after, StringComparison.Ordinal)) return false;
        var beforeReplacementCount = CountMatches(before, replace, replace, 0).Count;
        var afterReplacementCount = CountMatches(after, replace, replace, 0).Count;
        return afterReplacementCount >= beforeReplacementCount + expectedMatches;
    }

    // ---------- apply ----------

    public override ApplyExecution Apply(IReadOnlyList<JsonObject> ops, string snapshotId)
    {
        return ComInvoke(() =>
        {
            var exec = new ApplyExecution { Ok = true };
            var sessionId = Guid.NewGuid().ToString("n");
            var sessionStartedAt = DateTimeOffset.UtcNow;
            var checkedCount = 0;
            var mismatches = new List<string>();
            var foreground = new ForegroundInteractionGuard(App);
            var documentState = new HwpInteractionState();
            var userActivityInterrupted = false;
            try
            {
                var targetFile = FileArgOf(ops);
                var targetDocumentRef = DocumentRefArgOf(ops);
                var selectorError = HwpTargetSelectorError(ops);
                if (selectorError is not null)
                {
                    exec.Errors.Add(selectorError);
                    exec.Ok = false;
                    return exec;
                }
                var app = AttachHwpForTarget(
                    targetFile, targetDocumentRef, allowCreate: targetFile is not null,
                    foreground, documentState);
                if (app is null)
                {
                    exec.Errors.Add(targetFile is null && targetDocumentRef is null
                        ? "사용자가 열어 둔 한글 문서를 찾지 못했습니다. DocBridge는 빈 한글 창을 자동 실행하지 않습니다."
                        : "한글 자동화 인스턴스를 시작할 수 없습니다");
                    exec.Ok = false;
                    return exec;
                }
                dynamic hwp = app;
                if (OpenOrGetDoc(hwp, targetFile) is null) { exec.Errors.Add("열린 한글 문서가 없습니다 (file 지정 시 파일 열기 실패)"); exec.Ok = false; return exec; }
                if (!_closeTargetWhenDone) documentState.CaptureTarget(app);
                if (targetFile is null) KeepOwnedLiveDocumentOpen(hwp);

                foreach (var op in ops)
                {
                    var name = Json.GetString(op, "op")!;
                    var opStarted = Stopwatch.StartNew();
                    var mismatchCountBefore = mismatches.Count;
                    string? opError = null;
                    var opOk = false;
                    try
                    {
                      switch (name)
                      {
                        case "insert_text":
                        {
                            var text = Json.GetString(op, "text")!;
                            var context = GetSelectionText(hwp).Length > 0
                                ? CaptureCurrentNativeStyle(hwp, "current-selection")
                                : CaptureCaretContextStyle((object)hwp, "current-caret");
                            if (!PrepareContextualWriteStyle(hwp, op, context))
                            {
                                mismatches.Add("insert_text style application failed");
                                break;
                            }
                            var ok = ExecInsertText(hwp, text);
                            if (!ok) { mismatches.Add("InsertText action returned false"); break; }
                            checkedCount++;
                            if (!NormalizeNewlines(GetDocText(hwp)).Contains(NormalizeNewlines(text), StringComparison.Ordinal))
                                mismatches.Add("insert_text readback failed (text not found)");
                            exec.Affected.Add(new AffectedRef("cursor", $"inserted {text.Length} chars"));
                            break;
                        }
                        case "append_text":
                        {
                            var text = Json.GetString(op, "text")!;
                            var startNewParagraph = Json.GetBool(op, "startNewParagraph", true);
                            var ok = ExecAppendText(hwp, text, startNewParagraph, op);
                            if (!ok) { mismatches.Add("append_text action returned false"); break; }
                            checkedCount++;
                            var after = NormalizeNewlines(GetDocText(hwp));
                            if (!after.Contains(NormalizeNewlines(text), StringComparison.Ordinal))
                                mismatches.Add("append_text readback failed (text not found)");
                            exec.Affected.Add(new AffectedRef("document-end", $"appended {text.Length} chars"));
                            break;
                        }
                        case "insert_before_text":
                        case "insert_after_text":
                        {
                            var anchor = Json.GetString(op, "anchor")!;
                            var text = Json.GetString(op, "text")!;
                            var matchCase = Json.GetBool(op, "matchCase", true);
                            var occurrence = Json.GetInt(op, "occurrence") ?? 1;
                            var beforeText = NormalizeNewlines(GetDocText(hwp));
                            var matchCount = CountTextOccurrences(beforeText, anchor, matchCase);
                            if (matchCount == 0)
                            {
                                mismatches.Add($"{name}: anchor not found: {anchor}");
                                break;
                            }
                            if (!op.ContainsKey("occurrence") && matchCount != 1)
                            {
                                mismatches.Add($"{name}: anchor has {matchCount} matches; occurrence is required");
                                break;
                            }
                            if (occurrence < 1 || occurrence > matchCount)
                            {
                                mismatches.Add($"{name}: occurrence {occurrence} is outside 1..{matchCount}");
                                break;
                            }
                            bool ok;
                            try { ok = ExecInsertRelativeToText(hwp, op, name == "insert_before_text"); }
                            catch (Exception ex) { throw new InvalidOperationException($"relative insertion positioning failed: {ex.Message}", ex); }
                            if (!ok) { mismatches.Add($"{name} action returned false"); break; }

                            checkedCount++;
                            try
                            {
                                var afterText = NormalizeNewlines(GetDocText(hwp));
                                var anchorIndex = IndexOfTextOccurrence(afterText, anchor, occurrence, matchCase);
                                var previousAnchorIndex = occurrence > 1
                                    ? IndexOfTextOccurrence(afterText, anchor, occurrence - 1, matchCase)
                                    : -anchor.Length;
                                var nextAnchorIndex = occurrence < matchCount
                                    ? IndexOfTextOccurrence(afterText, anchor, occurrence + 1, matchCase)
                                    : afterText.Length;
                                var lowerBound = Math.Max(0, previousAnchorIndex + anchor.Length);
                                var insertedText = NormalizeNewlines(text);
                                var correctRegion = false;
                                if (anchorIndex >= 0 && nextAnchorIndex >= anchorIndex)
                                {
                                    int regionStart;
                                    int regionLength;
                                    if (name == "insert_before_text")
                                    {
                                        regionStart = lowerBound;
                                        regionLength = anchorIndex - lowerBound;
                                    }
                                    else
                                    {
                                        regionStart = anchorIndex + anchor.Length;
                                        regionLength = nextAnchorIndex - regionStart;
                                    }
                                    var region = regionLength >= 0 && regionStart >= 0 &&
                                                 regionStart + regionLength <= afterText.Length
                                        ? afterText.Substring(regionStart, regionLength)
                                        : string.Empty;
                                    correctRegion = HwpReadbackContainsEquivalent(region, insertedText);
                                }
                                if (NormalizeHwpReadbackComparable(afterText).Length <=
                                        NormalizeHwpReadbackComparable(beforeText).Length ||
                                    CountTextOccurrences(afterText, anchor, matchCase) != matchCount ||
                                    !correctRegion)
                                    mismatches.Add($"{name} readback failed (inserted text/anchor count mismatch)");
                            }
                            catch (Exception ex) { throw new InvalidOperationException($"relative insertion readback failed: {ex.Message}", ex); }
                            exec.Affected.Add(new AffectedRef($"anchor:{anchor}",
                                $"inserted {(name == "insert_before_text" ? "before" : "after")} occurrence {occurrence}/{matchCount}; surrounding paragraph style resolved"));
                            break;
                        }
                        case "replace_document_text":
                        {
                            string before = GetDocText(hwp);
                            var text = Json.GetString(op, "text")!;
                            var ok = ExecReplaceDocumentText(hwp, text, op);
                            if (!ok) { mismatches.Add("replace_document_text action returned false"); break; }
                            checkedCount++;
                            string after = GetDocText(hwp);
                            if (!NormalizeNewlines(after).Contains(NormalizeNewlines(text), StringComparison.Ordinal))
                                mismatches.Add("replace_document_text readback failed");
                            exec.Diff.Add(new DiffEntry
                            {
                                Ref = "document",
                                Before = before[..Math.Min(100, before.Length)],
                                After = text[..Math.Min(100, text.Length)],
                            });
                            exec.Affected.Add(new AffectedRef("document", $"replaced with {text.Length} chars"));
                            break;
                        }
                        case "replace_selection":
                        {
                            string before = GetSelectionText(hwp);
                            var text = Json.GetString(op, "text")!;
                            if (before.Length == 0)
                                exec.Warnings.Add("선택 영역 없음 — 커서 위치에 삽입했습니다");
                            var context = CaptureCurrentNativeStyle(hwp,
                                before.Length == 0 ? "current-caret" : "replaced-selection");
                            if (!PrepareContextualWriteStyle(hwp, op, context))
                            {
                                mismatches.Add("replace_selection style application failed");
                                break;
                            }
                            var ok = ExecInsertText(hwp, text); // 한글은 선택 상태에서 입력 시 선택 대체
                            if (!ok) { mismatches.Add("InsertText(replace) action returned false"); break; }
                            checkedCount++;
                            if (!NormalizeNewlines(GetDocText(hwp)).Contains(NormalizeNewlines(text), StringComparison.Ordinal))
                                mismatches.Add("replace_selection readback failed");
                            exec.Diff.Add(new DiffEntry { Ref = "selection", Before = before, After = text });
                            exec.Affected.Add(new AffectedRef("selection", "replaced"));
                            break;
                        }
                        case "find_replace":
                        {
                            var scope = FindReplaceScope(op);
                            HwpWriteResult result = scope is not null && scope.ContainsKey("tableIndex")
                                ? ApplyFindReplaceInTableCell(hwp, op)
                                : ApplyFindReplaceInDocument(hwp, op);
                            checkedCount++;
                            if (!result.Ok) mismatches.Add(result.Detail);
                            exec.Affected.Add(new AffectedRef(result.Ref, result.Detail));
                            if (result.Before is not null || result.After is not null)
                                exec.Diff.Add(new DiffEntry { Ref = result.Ref, Before = result.Before, After = result.After });
                            break;
                        }
                        case "set_paragraph_style_basic":
                        {
                            var style = Json.GetObj(op, "style") ?? new JsonObject();
                            var target = Json.GetObj(op, "target");
                            var targetText = Json.GetString(target, "text");
                            var scope = Json.GetString(target, "scope") ?? "selection";
                            var appliedCount = 0;
                            if (!string.IsNullOrEmpty(targetText))
                            {
                                appliedCount = ApplyCharShapeToTextMatches(hwp, targetText, style);
                            }
                            else if (scope == "document")
                            {
                                try { hwp.HAction.Run("MoveDocBegin"); } catch { }
                                if (!(bool)hwp.HAction.Run("SelectAll"))
                                {
                                    mismatches.Add("SelectAll for document style returned false");
                                    break;
                                }
                                if (ApplyCharShape(hwp, style) && ApplyParagraphAlignment(hwp, style)) appliedCount = 1;
                                try { hwp.HAction.Run("Cancel"); } catch { }
                                try { hwp.HAction.Run("MoveDocBegin"); } catch { }
                            }
                            else
                            {
                                if (ApplyCharShape(hwp, style) && ApplyParagraphAlignment(hwp, style)) appliedCount = 1;
                            }

                            if (appliedCount == 0) { mismatches.Add("CharShape target was not found or action returned false"); break; }
                            checkedCount++;
                            exec.Affected.Add(new AffectedRef("charShape",
                                !string.IsNullOrEmpty(targetText) ? $"applied to {appliedCount} match(es): {targetText}" : $"applied to {scope}"));
                            break;
                        }
                        case "set_paragraph_format":
                        {
                            var appliedCount = ApplyParagraphFormatTarget(hwp, op);
                            if (appliedCount == 0) { mismatches.Add("ParagraphShape target was not found or action returned false"); break; }
                            checkedCount++;
                            exec.Affected.Add(new AffectedRef("paragraphShape", $"applied to {appliedCount} target(s)"));
                            break;
                        }
                        case "format_paragraphs":
                        {
                            IReadOnlyList<HwpFormatResult> results = ExecFormatParagraphs(hwp, op);
                            var failed = results.FirstOrDefault(result => result.AppliedCount == 0);
                            if (failed is not null)
                            {
                                mismatches.Add($"format_paragraphs target was not found or action returned false: {failed.Ref}");
                                break;
                            }
                            checkedCount += results.Count;
                            foreach (var result in results)
                                exec.Affected.Add(new AffectedRef("paragraphFormat", $"{result.Ref}: {result.AppliedCount} target(s)"));
                            break;
                        }
                        case "set_page_setup":
                        {
                            var result = ExecSetPageSetup(hwp, op);
                            if (!result.Ok) { mismatches.Add(result.Detail); break; }
                            checkedCount++;
                            exec.Affected.Add(new AffectedRef(result.Ref, result.Detail));
                            exec.Diff.Add(new DiffEntry { Ref = result.Ref, Before = result.Before, After = result.After });
                            break;
                        }
                        case "insert_break":
                        {
                            var result = ExecInsertBreak(hwp, op);
                            if (!result.Ok) { mismatches.Add(result.Detail); break; }
                            checkedCount++;
                            exec.Affected.Add(new AffectedRef(result.Ref, result.Detail));
                            break;
                        }
                        case "insert_table":
                        {
                            TableInsertResult result = ExecInsertTable((object)hwp, op);
                            if (!result.Ok)
                            {
                                mismatches.Add($"insert_table failed after {result.Cells}/{result.Rows * result.Cols} cells");
                                break;
                            }

                            checkedCount++;
                            string docAfter = NormalizeNewlines(GetDocText(hwp));
                            var missing = result.ExpectedTexts
                                .Distinct(StringComparer.Ordinal)
                                .Where(cell => !docAfter.Contains(NormalizeNewlines(cell), StringComparison.Ordinal))
                                .Take(10)
                                .ToList();
                            if (missing.Count > 0)
                                mismatches.Add($"insert_table readback missing {missing.Count} cell value(s): {string.Join(", ", missing)}");
                            if (result.TableCountAfter <= result.TableCountBefore)
                                mismatches.Add($"insert_table control readback failed ({result.TableCountBefore} -> {result.TableCountAfter})");

                            exec.Affected.Add(new AffectedRef("table",
                                $"inserted {result.Rows}x{result.Cols}, {result.Cells} cells, {result.StyledCells} styled; controls {result.TableCountBefore}->{result.TableCountAfter}"));
                            exec.Diff.Add(new DiffEntry
                            {
                                Ref = "table:new",
                                Before = result.TableCountBefore,
                                After = result.TableCountAfter,
                            });
                            break;
                        }
                        case "table_cell_set_text":
                        {
                            var result = ExecTableCellSetText(hwp, op);
                            if (!result.Ok) { mismatches.Add(result.Detail); break; }
                            checkedCount++;
                            exec.Affected.Add(new AffectedRef(result.Ref, result.Detail));
                            exec.Diff.Add(new DiffEntry { Ref = result.Ref, Before = result.Before, After = result.After });
                            break;
                        }
                        case "table_set_cells":
                        {
                            IReadOnlyList<HwpWriteResult> results = ExecTableSetCells(hwp, op);
                            var failed = results.FirstOrDefault(result => !result.Ok);
                            if (failed is not null) { mismatches.Add(failed.Detail); break; }
                            checkedCount += results.Count;
                            foreach (var result in results)
                            {
                                exec.Affected.Add(new AffectedRef(result.Ref, result.Detail));
                                exec.Diff.Add(new DiffEntry { Ref = result.Ref, Before = result.Before, After = result.After });
                            }
                            break;
                        }
                        case "table_insert_rows":
                        case "table_insert_columns":
                        {
                            var result = ExecTableInsertLine(hwp, op, name == "table_insert_rows");
                            if (!result.Ok) { mismatches.Add(result.Detail); break; }
                            checkedCount++;
                            exec.Affected.Add(new AffectedRef(result.Ref, result.Detail));
                            break;
                        }
                        case "table_delete_rows":
                        case "table_delete_columns":
                        {
                            var result = ExecTableDeleteLine(hwp, op, name == "table_delete_rows");
                            if (!result.Ok) { mismatches.Add(result.Detail); break; }
                            checkedCount++;
                            exec.Affected.Add(new AffectedRef(result.Ref, result.Detail));
                            break;
                        }
                        case "table_merge_cells":
                        {
                            var result = ExecTableMergeCells(hwp, op);
                            if (!result.Ok) { mismatches.Add(result.Detail); break; }
                            checkedCount++;
                            exec.Affected.Add(new AffectedRef(result.Ref, result.Detail));
                            break;
                        }
                        case "table_set_row_height":
                        {
                            var result = ExecTableSetRowHeight(hwp, op);
                            if (!result.Ok) { mismatches.Add(result.Detail); break; }
                            checkedCount++;
                            exec.Affected.Add(new AffectedRef(result.Ref, result.Detail));
                            exec.Diff.Add(new DiffEntry { Ref = result.Ref, Before = result.Before, After = result.After });
                            break;
                        }
                        case "table_set_row_heights":
                        {
                            IReadOnlyList<HwpWriteResult> results = ExecTableSetRowHeights(hwp, op);
                            var failed = results.FirstOrDefault(result => !result.Ok);
                            if (failed is not null) { mismatches.Add(failed.Detail); break; }
                            checkedCount += results.Count;
                            foreach (var result in results)
                            {
                                exec.Affected.Add(new AffectedRef(result.Ref, result.Detail));
                                exec.Diff.Add(new DiffEntry { Ref = result.Ref, Before = result.Before, After = result.After });
                            }
                            break;
                        }
                        case "set_field_text":
                        {
                            var result = ExecSetFieldText(hwp, op);
                            if (!result.Ok) { mismatches.Add(result.Detail); break; }
                            checkedCount++;
                            exec.Affected.Add(new AffectedRef(result.Ref, result.Detail));
                            exec.Diff.Add(new DiffEntry { Ref = result.Ref, Before = result.Before, After = result.After });
                            break;
                        }
                        case "insert_picture":
                        {
                            EnsureFileAutomationSecurity(hwp);
                            var result = ExecInsertPicture(hwp, op);
                            if (!result.Ok) { mismatches.Add(result.Detail); break; }
                            checkedCount++;
                            exec.Affected.Add(new AffectedRef(result.Ref, result.Detail));
                            exec.Diff.Add(new DiffEntry { Ref = result.Ref, Before = result.Before, After = result.After });
                            break;
                        }
                        case "insert_page_number":
                        {
                            var result = ExecInsertPageNumber(hwp, op);
                            if (!result.Ok) { mismatches.Add(result.Detail); break; }
                            checkedCount++;
                            exec.Affected.Add(new AffectedRef(result.Ref, result.Detail));
                            exec.Diff.Add(new DiffEntry { Ref = result.Ref, Before = result.Before, After = result.After });
                            break;
                        }
                        case "set_header_footer_text":
                        {
                            var result = ExecSetHeaderFooterText(hwp, op);
                            if (!result.Ok) { mismatches.Add(result.Detail); break; }
                            checkedCount++;
                            exec.Affected.Add(new AffectedRef(result.Ref, result.Detail));
                            exec.Diff.Add(new DiffEntry { Ref = result.Ref, Before = result.Before, After = result.After });
                            break;
                        }
                        case "export_pdf":
                        {
                            EnsureFileAutomationSecurity(hwp);
                            var result = ExecExportPdf(hwp, op);
                            if (!result.Ok) { mismatches.Add(result.Detail); break; }
                            checkedCount++;
                            exec.Affected.Add(new AffectedRef(result.Ref, result.Detail));
                            exec.Diff.Add(new DiffEntry { Ref = result.Ref, Before = result.Before, After = result.After });
                            break;
                        }
                    }
                }

                // 파일 기반 워크플로: 적용 결과를 파일에 저장하고 닫는다 (잠금 해제)
                    catch (Exception ex)
                    {
                        opError = ex.Message;
                        mismatches.Add($"{name}: {ex.Message}");
                    }
                    finally
                    {
                        opStarted.Stop();
                        opOk = opError is null && mismatches.Count == mismatchCountBefore;
                        exec.OperationResults.Add(new JsonObject
                        {
                            ["index"] = exec.OperationResults.Count,
                            ["op"] = name,
                            ["ok"] = opOk,
                            ["elapsedMs"] = opStarted.ElapsedMilliseconds,
                            ["error"] = opError ?? (mismatches.Count > mismatchCountBefore
                                ? mismatches[mismatchCountBefore]
                                : null),
                        });
                    }
                    if (!opOk) break;
                    if (!foreground.Checkpoint(stopOnConcurrentInput: true))
                    {
                        userActivityInterrupted = true;
                        exec.Errors.Add("[APP_USER_ACTIVITY_DETECTED] 사용자가 한글 창을 조작하여 남은 작업을 안전하게 중단했습니다. 문서를 다시 읽은 뒤 이어서 실행하세요.");
                        break;
                    }
                }

                if (targetFile is not null)
                {
                    try
                    {
                        SaveActiveDoc(hwp, targetFile);
                        exec.Warnings.Add($"파일에 저장했습니다: {targetFile} (원본은 스냅샷에 백업 보존)");
                    }
                    catch (Exception ex) { mismatches.Add($"file save failed: {ex.Message}"); }
                }

                JsonObject? postEditReread = null;
                if (!userActivityInterrupted)
                {
                    try { postEditReread = CapturePostEditReread(hwp); }
                    catch (Exception ex) { mismatches.Add($"post-edit reread failed: {ex.Message}"); }
                }
                CloseTargetIfNeeded(hwp, targetFile);

                exec.Readback = new JsonObject
                {
                    ["verified"] = mismatches.Count == 0,
                    ["checked"] = checkedCount,
                    ["mismatches"] = Json.ToArray(mismatches),
                    ["snapshotId"] = snapshotId,
                    ["postEditReread"] = postEditReread,
                    ["session"] = new JsonObject
                    {
                        ["sessionId"] = sessionId,
                        ["state"] = mismatches.Count == 0 ? "ended" : "failed",
                        ["startedAt"] = sessionStartedAt.ToString("o"),
                        ["endedAt"] = DateTimeOffset.UtcNow.ToString("o"),
                        ["requestedSteps"] = ops.Count,
                        ["completedSteps"] = exec.OperationResults.Count(result => result is JsonObject item && Json.GetBool(item, "ok")),
                        ["failedStep"] = exec.OperationResults.FirstOrDefault(result => result is JsonObject item && !Json.GetBool(item, "ok"))?.DeepClone(),
                        ["stoppedEarly"] = exec.OperationResults.Count < ops.Count || mismatches.Count > 0,
                        ["autoSave"] = targetFile is not null,
                    },
                };
                exec.Ok = mismatches.Count == 0 && !userActivityInterrupted;
            }
            catch (Exception ex) { exec.Ok = false; exec.Errors.Add($"apply failed: {ex.Message}"); }
            finally { exec.Interaction = CompleteHwpInteraction(foreground, documentState); }
            return exec;
        });
    }

    // ---------- snapshot / restore ----------

    public override void CaptureSnapshot(string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject>? ops = null)
    {
        var targetFile = ops is null ? null : FileArgOf(ops);
        var targetDocumentRef = ops is null ? null : DocumentRefArgOf(ops);
        if (!string.IsNullOrWhiteSpace(targetFile))
        {
            var fullName = Path.GetFullPath(targetFile);
            if (!File.Exists(fullName))
                throw new FileNotFoundException($"한글 스냅샷 대상 파일이 없습니다: {fullName}", fullName);

            var dest = Path.Combine(snapshotDir, "document-backup" + Path.GetExtension(fullName));
            using (var src = new FileStream(fullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            using (var dst = new FileStream(dest, FileMode.CreateNew, FileAccess.Write, FileShare.None))
                src.CopyTo(dst);

            var hash = FileHash(dest);
            File.WriteAllText(Path.Combine(snapshotDir, "state.json"),
                new JsonObject
                {
                    ["fullName"] = fullName,
                    ["fileSha256"] = hash,
                    ["fileLength"] = new FileInfo(dest).Length,
                }.ToJsonString(Json.Pretty));
            metadata["documentRef"] = fullName;
            metadata["documentBackup"] = Path.GetFileName(dest);
            metadata["fileSha256"] = hash;
            metadata["payload"] = "full document file backup";
            metadata["restoreMode"] = "file-copy";
            return;
        }

        ComInvoke(() =>
        {
            var foreground = new ForegroundInteractionGuard(App);
            var documentState = new HwpInteractionState();
            try
            {
            var app = AttachHwpForTarget(
                null, targetDocumentRef, allowCreate: false, foreground, documentState);
            if (app is null) { metadata["payload"] = "none (hwp not running)"; return; }
            dynamic hwp = app;
            var doc = ActiveDoc(hwp);
            if (doc is null) { metadata["payload"] = "none (no document)"; return; }
            documentState.CaptureTarget(app);

            string fullName = "";
            try { fullName = (string)(doc.FullName ?? ""); } catch { }

            // 1) 디스크 파일도 보조 백업으로 복사하되, 라이브 복원에는 저장 전 상태를 담은
            //    아래 네이티브 메모리 스냅샷을 사용한다.
            if (!string.IsNullOrEmpty(fullName) && File.Exists(fullName))
            {
                var dest = Path.Combine(snapshotDir, "document-backup" + Path.GetExtension(fullName));
                try
                {
                    using var src = new FileStream(fullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    using var dst = new FileStream(dest, FileMode.Create, FileAccess.Write);
                    src.CopyTo(dst);
                    metadata["diskDocumentBackup"] = Path.GetFileName(dest);
                }
                catch (Exception ex) { metadata["diskDocumentBackupError"] = ex.Message; }
            }

            // 2) 현재 화면의 저장 전 상태까지 포함한 네이티브 HWP 전체 백업
            string native = GetNativeDocumentSnapshot(hwp);
            var nativeName = "document-live.hwp.base64.txt";
            var nativePath = Path.Combine(snapshotDir, nativeName);
            File.WriteAllText(nativePath, native);
            var nativeHash = Convert.ToHexString(SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(native)))
                .ToLowerInvariant();

            // 3) 빠른 검증/오류 메시지용 텍스트 state.json
            string text = GetDocText(hwp);
            var docId = "";
            try { docId = doc.DocumentID.ToString(); } catch { }
            var previewFingerprint = CapturePreviewFingerprint(hwp, text);
            File.WriteAllText(Path.Combine(snapshotDir, "state.json"),
                new JsonObject
                {
                    ["fullName"] = fullName,
                    ["documentId"] = docId,
                    ["text"] = text,
                    ["textLength"] = text.Length,
                    ["nativeSha256"] = nativeHash,
                    ["previewFingerprint"] = previewFingerprint,
                }.ToJsonString(Json.Pretty));
            metadata["payload"] = "native HWP live backup + state.json";
            metadata["liveDocumentBackup"] = nativeName;
            metadata["nativeSha256"] = nativeHash;
            metadata["previewFingerprint"] = previewFingerprint;
            metadata["restoreMode"] = "hwp-native-memory";
            metadata["documentId"] = docId;
            var identity = CaptureDocumentIdentity(hwp);
            metadata["documentRef"] = Json.GetString(identity, "documentRef") ?? metadata["documentRef"];
            metadata["instanceRef"] = Json.GetString(identity, "instanceRef");
            }
            finally { _ = CompleteHwpInteraction(foreground, documentState); }
        });
    }

    public JsonObject ValidatePreviewReuse(
        string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject> ops)
    {
        var targetFile = FileArgOf(ops);
        var targetDocumentRef = DocumentRefArgOf(ops);
        if (!string.IsNullOrWhiteSpace(targetFile))
        {
            var fullName = Path.GetFullPath(targetFile);
            var expected = Json.GetString(metadata, "fileSha256");
            if (string.IsNullOrWhiteSpace(expected) || !File.Exists(fullName))
                return new JsonObject
                {
                    ["ok"] = true, ["reusable"] = false,
                    ["reason"] = "file snapshot fingerprint is unavailable",
                };
            var current = FileHash(fullName);
            var reusable = string.Equals(expected, current, StringComparison.OrdinalIgnoreCase);
            return new JsonObject
            {
                ["ok"] = true,
                ["reusable"] = reusable,
                ["fingerprintMethod"] = "file-sha256",
                ["expected"] = expected,
                ["current"] = current,
                ["reason"] = reusable ? "snapshot fingerprint matched" : "file changed after dry-run",
            };
        }

        return ComInvoke(() =>
        {
            var foreground = new ForegroundInteractionGuard(App);
            var documentState = new HwpInteractionState();
            try
            {
            var app = AttachHwpForTarget(
                null, targetDocumentRef, allowCreate: false, foreground, documentState);
            if (app is null)
                return new JsonObject { ["ok"] = true, ["reusable"] = false, ["reason"] = "HWP is not running" };
            dynamic hwp = app;
            var doc = ActiveDoc(hwp);
            if (doc is null)
                return new JsonObject { ["ok"] = true, ["reusable"] = false, ["reason"] = "HWP document is not open" };
            documentState.CaptureTarget(app);

            var expectedDocumentId = Json.GetString(metadata, "documentId") ?? "";
            var currentDocumentId = "";
            try { currentDocumentId = doc.DocumentID.ToString(); } catch { }
            if (!string.Equals(expectedDocumentId, currentDocumentId, StringComparison.Ordinal))
                return new JsonObject
                {
                    ["ok"] = true, ["reusable"] = false,
                    ["fingerprintMethod"] = "hwp-text+selection+position+controls+fields+caret-style-sha256",
                    ["reason"] = "active HWP document identity changed after dry-run",
                };

            var expected = Json.GetString(metadata, "previewFingerprint") ?? "";
            var current = CapturePreviewFingerprint(hwp);
            var reusable = expected.Length > 0 && string.Equals(expected, current, StringComparison.OrdinalIgnoreCase);
            return new JsonObject
            {
                ["ok"] = true,
                ["reusable"] = reusable,
                ["fingerprintMethod"] = "hwp-text+selection+position+controls+fields+caret-style-sha256",
                ["expected"] = expected,
                ["current"] = current,
                ["reason"] = reusable ? "snapshot fingerprint matched" : "HWP preview-dependent state changed after dry-run",
            };
            }
            finally { _ = CompleteHwpInteraction(foreground, documentState); }
        });
    }

    public override JsonObject RestoreSnapshot(string snapshotDir, JsonObject metadata)
    {
        var backupName = Json.GetString(metadata, "documentBackup");
        var targetFile = Json.GetString(metadata, "documentRef");
        if (!string.IsNullOrWhiteSpace(backupName) && !string.IsNullOrWhiteSpace(targetFile))
        {
            try
            {
                var snapshotRoot = Path.GetFullPath(snapshotDir)
                    .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
                var backupPath = Path.GetFullPath(Path.Combine(snapshotDir, backupName));
                if (!backupPath.StartsWith(snapshotRoot, StringComparison.OrdinalIgnoreCase) || !File.Exists(backupPath))
                    return Json.ErrorResult("invalid or missing document backup in snapshot", App);

                var destination = Path.GetFullPath(targetFile);
                var parent = Path.GetDirectoryName(destination)
                    ?? throw new InvalidOperationException("한글 복원 대상 디렉터리를 확인할 수 없습니다");
                Directory.CreateDirectory(parent);
                var temp = Path.Combine(parent, $".{Path.GetFileName(destination)}.docbridge-{Guid.NewGuid():N}.tmp");
                try
                {
                    File.Copy(backupPath, temp, overwrite: false);
                    if (File.Exists(destination)) File.Replace(temp, destination, null, ignoreMetadataErrors: true);
                    else File.Move(temp, destination);
                }
                finally
                {
                    if (File.Exists(temp)) File.Delete(temp);
                }

                var wantHash = Json.GetString(metadata, "fileSha256") ?? FileHash(backupPath);
                var gotHash = FileHash(destination);
                var verified = string.Equals(wantHash, gotHash, StringComparison.OrdinalIgnoreCase);
                return new JsonObject
                {
                    ["ok"] = verified,
                    ["restored"] = verified,
                    ["documentRef"] = destination,
                    ["readback"] = new JsonObject
                    {
                        ["verified"] = verified,
                        ["sha256"] = gotHash,
                        ["checked"] = 1,
                    },
                    ["warnings"] = Json.ToArray(new[] { "파일 전체 백업으로 복원했습니다. 열려 있던 한글 창에서는 문서를 다시 여세요." }),
                    ["errors"] = verified ? Json.ToArray(Array.Empty<string>()) : Json.ToArray(new[] { "restored file hash mismatch" }),
                };
            }
            catch (Exception ex)
            {
                return Json.ErrorResult($"한글 파일 복원 실패: {ex.Message}. 대상 파일을 한글에서 닫고 다시 시도하세요.", App);
            }
        }

        var liveBackupName = Json.GetString(metadata, "liveDocumentBackup");
        if (string.Equals(Json.GetString(metadata, "restoreMode"), "hwp-native-memory", StringComparison.Ordinal) &&
            !string.IsNullOrWhiteSpace(liveBackupName))
        {
            return ComInvoke(() =>
            {
                var foreground = new ForegroundInteractionGuard(App);
                var documentState = new HwpInteractionState();
                try
                {
                var snapshotRoot = Path.GetFullPath(snapshotDir)
                    .TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar) + Path.DirectorySeparatorChar;
                var liveBackupPath = Path.GetFullPath(Path.Combine(snapshotDir, liveBackupName));
                if (!liveBackupPath.StartsWith(snapshotRoot, StringComparison.OrdinalIgnoreCase) || !File.Exists(liveBackupPath))
                    return Json.ErrorResult("invalid or missing native HWP backup in snapshot", App);

                var expectedRef = Json.GetString(metadata, "documentRef");
                var app = AttachHwpForTarget(null, expectedRef, allowCreate: false,
                    foreground, documentState);
                if (app is null) return Json.ErrorResult("한글이 실행 중이지 않습니다", App);
                dynamic hwp = app;
                var doc = ActiveDoc(hwp);
                if (doc is null) return Json.ErrorResult("열린 한글 문서가 없습니다", App);

                documentState.CaptureTarget(app);
                var expectedId = Json.GetString(metadata, "documentId");
                string currentName = "";
                string currentId = "";
                try { currentName = (string)(doc.FullName ?? ""); } catch { }
                try { currentId = doc.DocumentID.ToString(); } catch { }
                if (!string.IsNullOrWhiteSpace(expectedRef) && Path.IsPathFullyQualified(expectedRef) &&
                    !string.Equals(Path.GetFullPath(expectedRef), Path.GetFullPath(currentName), StringComparison.OrdinalIgnoreCase))
                    return Json.ErrorResult("현재 한글 문서가 라이브 스냅샷 대상과 다릅니다", App);
                if (string.IsNullOrWhiteSpace(currentName) && !string.IsNullOrWhiteSpace(expectedId) &&
                    !string.Equals(expectedId, currentId, StringComparison.Ordinal))
                    return Json.ErrorResult("현재 빈 문서가 라이브 스냅샷 대상과 다릅니다", App);

                var native = File.ReadAllText(liveBackupPath);
                var wantHash = Json.GetString(metadata, "nativeSha256");
                var gotHash = Convert.ToHexString(SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(native)))
                    .ToLowerInvariant();
                if (!string.IsNullOrWhiteSpace(wantHash) && !string.Equals(wantHash, gotHash, StringComparison.OrdinalIgnoreCase))
                    return Json.ErrorResult("native HWP backup hash mismatch", App);

                bool ok = RestoreNativeDocumentSnapshot(hwp, native);
                var statePath = Path.Combine(snapshotDir, "state.json");
                var state = File.Exists(statePath)
                    ? JsonNode.Parse(File.ReadAllText(statePath)) as JsonObject
                    : null;
                var wantText = Json.GetString(state, "text") ?? "";
                string after = GetDocText(hwp);
                var verified = ok && string.Equals(
                    NormalizeNewlines(after).TrimEnd('\n'),
                    NormalizeNewlines(wantText).TrimEnd('\n'),
                    StringComparison.Ordinal);
                return new JsonObject
                {
                    ["ok"] = verified,
                    ["restored"] = verified,
                    ["documentRef"] = Json.GetString(CaptureDocumentIdentity(hwp), "documentRef"),
                    ["readback"] = new JsonObject
                    {
                        ["verified"] = verified,
                        ["checked"] = 1,
                        ["textLength"] = after.Length,
                        ["nativeSha256"] = gotHash,
                    },
                    ["warnings"] = Json.ToArray(Array.Empty<string>()),
                    ["errors"] = verified ? Json.ToArray(Array.Empty<string>())
                        : Json.ToArray(new[] { "native HWP restore readback mismatch" }),
                };
                }
                finally { _ = CompleteHwpInteraction(foreground, documentState); }
            });
        }

        return ComInvoke(() =>
        {
            var foreground = new ForegroundInteractionGuard(App);
            var documentState = new HwpInteractionState();
            try
            {
            var statePath = Path.Combine(snapshotDir, "state.json");
            if (!File.Exists(statePath)) return Json.ErrorResult("state.json not found in snapshot", App);

            var snapshotDocumentRef = Json.GetString(metadata, "documentRef");
            var app = AttachHwpForTarget(null, snapshotDocumentRef, allowCreate: false,
                foreground, documentState);
            if (app is null) return Json.ErrorResult("한글이 실행 중이지 않습니다", App);
            dynamic hwp = app;
            if (ActiveDoc(hwp) is null) return Json.ErrorResult("열린 한글 문서가 없습니다", App);

            documentState.CaptureTarget(app);
            var state = JsonNode.Parse(File.ReadAllText(statePath)) as JsonObject ?? new JsonObject();
            var wantText = Json.GetString(state, "text") ?? "";
            var docRef = Json.GetString(state, "fullName");

            // 현재 문서 확인
            try
            {
                var cur = ActiveDoc(hwp);
                string curName = "";
                try { curName = (string)(cur.FullName ?? ""); } catch { }
                if (!string.IsNullOrEmpty(docRef) && !string.IsNullOrEmpty(curName) &&
                    !string.Equals(curName, docRef, StringComparison.OrdinalIgnoreCase))
                    return Json.ErrorResult(
                        $"현재 문서 '{curName}'가 스냅샷 문서 '{docRef}'와 다릅니다. 해당 문서를 먼저 여세요.", App);
            }
            catch { }

            // 텍스트 전체 복원: 전체 선택 → InsertText
            dynamic act = hwp.HAction;
            act.Run("SelectAll");
            var ok = ExecInsertText(hwp, wantText);

            string after = GetDocText(hwp);
            var verified = ok && string.Equals(
                NormalizeNewlines(after).TrimEnd('\n'),
                NormalizeNewlines(wantText).TrimEnd('\n'),
                StringComparison.Ordinal);
            return new JsonObject
            {
                ["ok"] = verified,
                ["restored"] = true,
                ["readback"] = new JsonObject
                {
                    ["verified"] = verified,
                    ["checked"] = 1,
                    ["textLength"] = after.Length,
                },
                ["warnings"] = Json.ToArray(new[]
                {
                    "한글 복원은 텍스트 기준입니다. 서식/표/개체는 스냅샷 파일 백업(document-backup)으로만 보존됩니다.",
                }),
                ["errors"] = verified ? Json.ToArray(Array.Empty<string>())
                    : Json.ToArray(new[] { "restore readback mismatch" }),
            };
            }
            finally { _ = CompleteHwpInteraction(foreground, documentState); }
        });
    }

    public override void Dispose()
    {
        var ownedProcessId = _ownsAttached ? _ownedProcessId : 0;
        if (_attached is not null)
        {
            try
            {
                Sta.Invoke<object?>(() =>
                {
                    if (_ownsAttached && ownedProcessId == 0)
                        ownedProcessId = RotHelper.ProcessIdFromWindowHandle(RotHelper.HwpWindowHandle(_attached));
                    if (_ownsAttached)
                    {
                        try
                        {
                            dynamic hwp = _attached;
                            // 전용 자동화 문서는 저장 질문 없이 닫아 숨은 모달과 프로세스 누수를 막는다.
                            try { hwp.Clear(1); } catch { }
                            try
                            {
                                dynamic windows = hwp.XHwpWindows;
                                windows.Close(false);
                                RotHelper.ReleaseComObject(windows);
                            }
                            catch { }
                        }
                        catch { }
                    }
                    RotHelper.ReleaseComObject(_attached);
                    return null;
                }, TimeSpan.FromSeconds(3));
            }
            catch { }
        }
        _attached = null;
        base.Dispose();
        if (ownedProcessId > 0)
        {
            try
            {
                using var process = Process.GetProcessById(ownedProcessId);
                if (!process.WaitForExit(1500) &&
                    string.Equals(process.ProcessName, "Hwp", StringComparison.OrdinalIgnoreCase))
                {
                    process.Kill(entireProcessTree: false);
                    process.WaitForExit(5000);
                }
            }
            catch (ArgumentException) { }
            catch (InvalidOperationException) { }
        }
        _ownedProcessId = 0;
    }
}
