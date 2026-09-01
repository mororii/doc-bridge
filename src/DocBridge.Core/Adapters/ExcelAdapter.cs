using System.Globalization;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

/// <summary>
/// Excel 어댑터 (M1): 실행 중인 Microsoft Excel desktop에 COM Interop(late binding)으로 연결.
/// - 모든 COM 호출은 STA 디스패처 스레드에서 직렬화
/// - 읽기: workbook/worksheet/selection/used range/sheet 목록
/// - 쓰기 op: set_values, set_formulas, insert_rows, insert_cols, format_range, find_replace, copy_sheet,
///   merge_cells, unmerge_cells, set_rows_hidden, set_cols_hidden, set_sheet_visibility
/// - 금지(정책): delete_sheet, overwrite_workbook_without_backup, run_macro
/// - snapshot: workbook 파일 복사(저장된 경우) + 시트 값 state.json
/// </summary>
public sealed partial class ExcelAdapter : ComAdapterBase, IConnectionLifecycleAdapter
{
    private const int MaxCells = 10000;   // 대용량 읽기 상한 (정책 maxReadCells)
    private const int MaxSnapshotCells = 1000000; // 로컬 복원용 상한 (AI 응답에는 노출하지 않음)
    private const int MaxFormatSnapshotCells = 100000; // 셀별 서식 복원 상한
    private const int MaxDiff = 100;      // diff 항목 상한 (정책 maxDiffEntries)
    private static readonly TimeSpan ComTimeout = TimeSpan.FromSeconds(120);

    /// <summary>테스트/특수 용도: 어댑터 STA 스레드 안에서 평가되는 Application 팩토리</summary>
    private readonly Func<object?>? _appFactory;
    private readonly bool _appFactoryOwnsInstance;
    private readonly Timer _idleLifecycleTimer;
    private object? _attached;
    private bool _ownsInstance;
    private int _ownedProcessId;
    private ExcelOwnerWatchdog.Lease? _ownerWatchdog;
    private int _lifecycleTickRunning;
    private int _excelDisposed;

    public ExcelAdapter(Func<object?>? appFactory = null, bool appFactoryOwnsInstance = false)
        : base("excel", "Excel.Application")
    {
        _appFactory = appFactory;
        _appFactoryOwnsInstance = appFactoryOwnsInstance;
        _idleLifecycleTimer = new Timer(
            _ => RunIdleLifecycleCheck(),
            null,
            TimeSpan.FromSeconds(1),
            TimeSpan.FromSeconds(1));
    }

    /// <summary>반드시 STA 스레드 안에서 호출. RCW 생성/사용을 같은 아파트먼트로 통일.
    /// 팩토리가 주입된 경우 그 결과가 최종(fallback 없음) — 테스트 결정성 확보.</summary>
    private object? AttachExcel(bool allowCreate = false)
    {
        if (_attached is not null)
        {
            try
            {
                // A cached RCW survives after EXCEL.EXE exits, but any real property access then
                // raises RPC_E_DISCONNECTED (or an equivalent server-death HRESULT). Probe the
                // application before every top-level adapter call so the same MCP process can
                // follow a newly started Excel instance.
                dynamic current = _attached;
                var hwnd = Convert.ToInt64(current.Hwnd, CultureInfo.InvariantCulture);
                if (!Marshal.IsComObject(_attached))
                    return _attached;

                if (!RotHelper.IsWindowAlive(hwnd))
                {
                    DetachExcelReference();
                }
                else
                {
                    var workbookCount = GetWorkbookCount(_attached);
                    if (workbookCount > 0) return _attached;

                    // Closing the last workbook can leave EXCEL.EXE alive because this adapter
                    // still owns an RCW. If the user has already started a fresh Excel window,
                    // prefer that instance. Injected factories (tests/embedded hosts) create a
                    // replacement immediately. Production only creates a visible instance when
                    // an Excel operation actually needs one; status probing never launches Excel.
                    var replacement = _appFactory is null ? FindPreferredRunningExcel(hwnd, requireWorkbook: true) : null;
                    if (replacement is not null)
                    {
                        _ = DisconnectExcelCore("replaced-by-user-instance");
                        _attached = replacement;
                        _ownsInstance = false;
                        return _attached;
                    }
                    if (_appFactory is not null) DetachExcelReference();
                    else return _attached;
                }
            }
            catch (Exception ex) when (IsComDisconnected(ex))
            {
                DetachExcelReference();
            }
        }

        if (_appFactory is not null)
        {
            _attached = _appFactory();
            _ownsInstance = _attached is not null && _appFactoryOwnsInstance;
            if (_ownsInstance)
            {
                MakeOwnedInstanceVisible(_attached!);
                _ownedProcessId = ReadApplicationProcessId(_attached!);
                _ownerWatchdog = ExcelOwnerWatchdog.Start(_ownedProcessId);
            }
            return _attached;
        }

        _attached = FindPreferredRunningExcel();
        _ownsInstance = false;
        _ownedProcessId = 0;
        if (_attached is not null || !allowCreate) return _attached;

        _attached = RotHelper.CreateInstance("Excel.Application");
        _ownsInstance = _attached is not null;
        if (_attached is not null)
        {
            MakeOwnedInstanceVisible(_attached);
            _ownedProcessId = ReadApplicationProcessId(_attached);
            _ownerWatchdog = ExcelOwnerWatchdog.Start(_ownedProcessId);
        }
        return _attached;
    }

    private static object? FindPreferredRunningExcel(long excludedHwnd = 0, bool requireWorkbook = false)
    {
        object? selected = null;
        var selectedScore = int.MinValue;
        foreach (var candidate in RotHelper.GetExcelApplications())
        {
            var keep = false;
            try
            {
                dynamic app = candidate;
                var hwnd = Convert.ToInt64(app.Hwnd, CultureInfo.InvariantCulture);
                if (hwnd == excludedHwnd || !RotHelper.IsWindowAlive(hwnd)) continue;
                var workbookCount = GetWorkbookCount(candidate);
                if (requireWorkbook && workbookCount == 0) continue;

                // A visible instance with a real workbook is a user session. Prefer it over a
                // zero-workbook automation remnant returned by the ProgID fallback.
                var score = workbookCount > 0 ? 1000 + workbookCount : 1;
                if (score <= selectedScore) continue;
                if (selected is not null) RotHelper.ReleaseComObject(selected);
                selected = candidate;
                selectedScore = score;
                keep = true;
            }
            catch
            {
                // Ignore Excel instances that are closing, busy, or no longer connected.
            }
            finally
            {
                if (!keep && !ReferenceEquals(candidate, selected))
                    RotHelper.ReleaseComObject(candidate);
            }
        }
        return selected;
    }

    private static int GetWorkbookCount(object application)
    {
        object? workbooks = null;
        try
        {
            dynamic app = application;
            workbooks = (object)app.Workbooks;
            return Convert.ToInt32(((dynamic)workbooks).Count, CultureInfo.InvariantCulture);
        }
        finally
        {
            RotHelper.ReleaseComObject(workbooks);
        }
    }

    private static void MakeOwnedInstanceVisible(object application)
    {
        // Never create a hidden Excel process. Do not change DisplayAlerts: suppressing alerts
        // would turn a later Quit into a possible data-loss path.
        dynamic app = application;
        app.Visible = true;
    }

    private static int ReadApplicationProcessId(object application)
    {
        try
        {
            var hwnd = Convert.ToInt64(((dynamic)application).Hwnd, CultureInfo.InvariantCulture);
            return RotHelper.ProcessIdFromWindowHandle(hwnd);
        }
        catch { return 0; }
    }

    private static string? ReadActiveWorkbookFullName(object application)
    {
        object? workbook = null;
        try
        {
            workbook = (object?)((dynamic)application).ActiveWorkbook;
            return workbook is null
                ? null
                : Convert.ToString(((dynamic)workbook).FullName, CultureInfo.InvariantCulture);
        }
        catch { return null; }
        finally { RotHelper.ReleaseComObject(workbook); }
    }

    private void DetachExcelReference()
    {
        var ownerWatchdog = _ownerWatchdog;
        _ownerWatchdog = null;
        if (_attached is null)
        {
            _ownsInstance = false;
            _ownedProcessId = 0;
            ownerWatchdog?.Dispose();
            return;
        }
        RotHelper.ReleaseComObject(_attached);
        _attached = null;
        _ownsInstance = false;
        _ownedProcessId = 0;
        ForceComReferenceCleanup();
        ownerWatchdog?.Dispose();
    }

    /// <summary>
    /// A file-scoped read may create Excel solely to open the requested workbook.  Once that
    /// workbook has been closed (or opening it failed), retaining a zero-workbook application
    /// only leaves a disabled grey Excel shell behind.  Clean up only an instance created by
    /// DocBridge; a user's zero-workbook Excel shell is never quit.
    /// </summary>
    private void DisconnectOwnedExcelIfEmpty(string reason)
    {
        if (!_ownsInstance || _attached is null) return;
        try
        {
            if (GetWorkbookCount(_attached) == 0)
                _ = DisconnectExcelCore(reason);
        }
        catch (Exception ex) when (IsComDisconnected(ex))
        {
            DetachExcelReference();
        }
        catch
        {
            // Busy/modal Excel is left intact.  The idle lifecycle timer retries without ever
            // force-killing Excel or suppressing save prompts.
        }
    }

    private static List<string> ReadUnsavedWorkbookNames(object application)
    {
        var unsaved = new List<string>();
        object? workbooks = null;
        try
        {
            dynamic app = application;
            workbooks = (object)app.Workbooks;
            var count = Convert.ToInt32(((dynamic)workbooks).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                object? workbook = null;
                try
                {
                    workbook = (object)((dynamic)workbooks).Item(index);
                    var saved = Convert.ToBoolean(((dynamic)workbook).Saved, CultureInfo.InvariantCulture);
                    if (!saved)
                        unsaved.Add(Convert.ToString(((dynamic)workbook).Name, CultureInfo.InvariantCulture)
                            ?? $"workbook #{index}");
                }
                catch
                {
                    // If save state cannot be proved, Quit is not safe.
                    unsaved.Add($"workbook #{index} (save state unknown)");
                }
                finally
                {
                    RotHelper.ReleaseComObject(workbook);
                }
            }
        }
        finally
        {
            RotHelper.ReleaseComObject(workbooks);
        }
        return unsaved;
    }

    private JsonObject DisconnectExcelCore(string reason)
    {
        var attached = _attached;
        var owned = _ownsInstance;
        var ownedProcessId = owned ? _ownedProcessId : 0;
        var ownerWatchdog = _ownerWatchdog;
        _attached = null;
        _ownsInstance = false;
        _ownedProcessId = 0;
        _ownerWatchdog = null;

        var quitCalled = false;
        var unsaved = new List<string>();
        var warnings = new List<string>();
        if (attached is not null && owned)
        {
            try { unsaved = ReadUnsavedWorkbookNames(attached); }
            catch (Exception ex)
            {
                unsaved.Add("save state could not be verified");
                warnings.Add($"Excel 저장 상태 확인 실패: {ex.Message}");
            }

            if (unsaved.Count == 0)
            {
                try
                {
                    ((dynamic)attached).Quit();
                    quitCalled = true;
                }
                catch (Exception ex)
                {
                    warnings.Add($"DocBridge 소유 Excel 종료 요청 실패: {ex.Message}");
                }
            }
            else
            {
                warnings.Add("저장되지 않은 통합문서가 있어 Excel을 종료하지 않고 COM 연결만 해제했습니다.");
            }
        }

        RotHelper.ReleaseComObject(attached);
        ForceComReferenceCleanup();
        // Signal only after this process has released its root RCW. On a clean Quit the watchdog
        // holds the final COM reference, then releases it; on an unsaved workbook it only detaches.
        ownerWatchdog?.Dispose();
        if (quitCalled && ownedProcessId > 0 && !WaitForProcessExit(ownedProcessId, TimeSpan.FromSeconds(10)))
            warnings.Add($"DocBridge 소유 Excel PID {ownedProcessId}가 정상 종료 대기 시간 안에 끝나지 않았습니다. 강제 종료하지 않았습니다.");
        return new JsonObject
        {
            ["ok"] = true,
            ["app"] = App,
            ["disconnected"] = attached is not null,
            ["reason"] = reason,
            ["ownedInstance"] = owned,
            ["quitCalled"] = quitCalled,
            ["unsavedWorkbooks"] = Json.ToArray(unsaved),
            ["warnings"] = Json.ToArray(warnings),
        };
    }

    private static void ForceComReferenceCleanup()
    {
        // All top-level adapter calls have returned before disconnect/idle cleanup runs, so
        // temporary Workbook/Worksheet/Range RCWs are no longer live locals. Two passes release
        // both those RCWs and finalizers created by the first pass.
        GC.Collect();
        GC.WaitForPendingFinalizers();
        GC.Collect();
        GC.WaitForPendingFinalizers();
    }

    private static bool WaitForProcessExit(int processId, TimeSpan timeout)
    {
        try
        {
            using var process = Process.GetProcessById(processId);
            return process.WaitForExit((int)timeout.TotalMilliseconds);
        }
        catch (ArgumentException) { return true; }
        catch (InvalidOperationException) { return true; }
        catch { return false; }
    }

    private void RunIdleLifecycleCheck()
    {
        if (Volatile.Read(ref _excelDisposed) != 0 ||
            Interlocked.Exchange(ref _lifecycleTickRunning, 1) != 0)
            return;
        try
        {
            Sta.Invoke<object?>(() =>
            {
                if (_attached is null || !Marshal.IsComObject(_attached)) return null;
                try
                {
                    // Once the user closes the last workbook/window, keeping our RCW is the exact
                    // condition that creates the add-in-less EXCEL.EXE remnant.
                    if (GetWorkbookCount(_attached) == 0)
                        _ = DisconnectExcelCore("last-workbook-closed");
                }
                catch (Exception ex) when (IsComDisconnected(ex))
                {
                    DetachExcelReference();
                }
                catch
                {
                    // Busy/modal Excel is retried on the next tick; never force-kill it.
                }
                return null;
            }, TimeSpan.FromSeconds(3));
        }
        catch { }
        finally { Volatile.Write(ref _lifecycleTickRunning, 0); }
    }

    // ---------- 공통 COM 유틸 (모두 STA 스레드에서 실행) ----------

    private static dynamic RequireWorkbook(dynamic app)
    {
        dynamic? wb = app.ActiveWorkbook;
        if (wb is null)
            throw new InvalidOperationException("Excel에 열린 workbook이 없습니다. 문서를 먼저 여세요.");
        return wb;
    }

    private static dynamic GetSheet(dynamic wb, string? sheetName)
    {
        if (string.IsNullOrWhiteSpace(sheetName)) return wb.ActiveSheet;
        dynamic sheets = wb.Worksheets;
        try { return sheets.Item(sheetName); }
        catch { throw new InvalidOperationException($"sheet '{sheetName}' not found"); }
    }

    private static dynamic GetRequiredTargetSheet(dynamic wb, JsonObject op)
    {
        var sheetName = Json.GetString(Json.GetObj(op, "target"), "sheet");
        if (string.IsNullOrWhiteSpace(sheetName))
            throw new InvalidOperationException(
                $"Excel write op '{Json.GetString(op, "op")}' requires target.sheet; the active sheet is never assumed for writes");
        return GetSheet(wb, sheetName);
    }

    private static (object Sheet, string Address) ResolveRangeTarget(
        object workbook, string? explicitSheetName, string rangeReference, bool requireExplicitSheet)
    {
        dynamic wb = workbook;
        var parsed = ExcelRangeReference.Parse(rangeReference);
        if (!string.IsNullOrWhiteSpace(explicitSheetName) &&
            !string.IsNullOrWhiteSpace(parsed.SheetName) &&
            !string.Equals(explicitSheetName, parsed.SheetName, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                $"Excel target.sheet '{explicitSheetName}' does not match range sheet '{parsed.SheetName}'");

        var sheetName = explicitSheetName ?? parsed.SheetName;
        if (requireExplicitSheet && string.IsNullOrWhiteSpace(sheetName))
            throw new InvalidOperationException(
                "Excel write range must use target.sheet or a sheet-qualified range such as '매출'!B2");
        dynamic sheet = GetSheet(wb, sheetName);
        return ((object)sheet, parsed.Address);
    }

    private sealed class OpenWorkbook : IDisposable
    {
        private readonly bool _releaseApplication;
        private int _disposed;

        public OpenWorkbook(object application, object workbook, bool closeWhenDone = false,
            bool releaseApplication = false)
        {
            Application = application;
            Workbook = workbook;
            CloseWhenDone = closeWhenDone;
            _releaseApplication = releaseApplication;
        }

        public object Application { get; }
        public object Workbook { get; }
        public bool CloseWhenDone { get; }

        public void Dispose()
        {
            if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
            if (CloseWhenDone)
                try { ((dynamic)Workbook).Close(false); } catch { }
            // Workbook/Application can be aliases of RCWs still held by the destination
            // lease or the attached application. Balance only this acquisition; a final
            // release here can separate those live aliases from their native objects.
            RotHelper.ReleaseComReference(Workbook);
            if (_releaseApplication) RotHelper.ReleaseComReference(Application);
        }
    }

    private sealed class WorkbookLease : IDisposable
    {
        private readonly OpenWorkbook? _opened;
        private int _disposed;

        public WorkbookLease(object workbook, OpenWorkbook? opened = null)
        {
            Workbook = workbook;
            _opened = opened;
        }

        public object Workbook { get; }

        public void Dispose()
        {
            if (Interlocked.Exchange(ref _disposed, 1) != 0) return;
            if (_opened is not null) _opened.Dispose();
            else RotHelper.ReleaseComReference(Workbook);
        }
    }

    private static OpenWorkbook FindOpenWorkbook(object attachedApplication, string reference, bool allowFileOpenFallback = true)
    {
        var applications = new List<(object Application, bool ReleaseWhenDone)> { (attachedApplication, false) };
        foreach (var candidate in RotHelper.GetExcelApplications())
        {
            if (ReferenceEquals(candidate, attachedApplication))
            {
                RotHelper.ReleaseComReference(candidate);
                continue;
            }
            applications.Add((candidate, true));
        }
        var seenApps = new HashSet<long>();
        var matches = new List<OpenWorkbook>();

        foreach (var candidate in applications)
        {
            var applicationTransferred = false;
            object? workbooks = null;
            try
            {
                var appObj = candidate.Application;
                dynamic app = appObj;
                var hwnd = Convert.ToInt64(app.Hwnd, CultureInfo.InvariantCulture);
                if (!seenApps.Add(hwnd)) continue;
                workbooks = (object)app.Workbooks;
                var count = Convert.ToInt32(((dynamic)workbooks).Count, CultureInfo.InvariantCulture);
                for (var index = 1; index <= count; index++)
                {
                    object? workbook = null;
                    var matched = false;
                    try
                    {
                        workbook = (object)((dynamic)workbooks).Item(index);
                        var name = Convert.ToString(((dynamic)workbook).Name, CultureInfo.InvariantCulture) ?? "";
                        var fullName = Convert.ToString(((dynamic)workbook).FullName, CultureInfo.InvariantCulture) ?? "";
                        matched = string.Equals(reference, name, StringComparison.OrdinalIgnoreCase) ||
                            PathsEqualWhenQualified(reference, fullName);
                        if (matched)
                        {
                            matches.Add(new OpenWorkbook(appObj, workbook,
                                releaseApplication: candidate.ReleaseWhenDone));
                            applicationTransferred = candidate.ReleaseWhenDone;
                            workbook = null;
                        }
                    }
                    finally
                    {
                        if (!matched) RotHelper.ReleaseComObject(workbook);
                    }
                }
            }
            catch
            {
                // 종료 중이거나 모달 상태인 Excel 인스턴스는 다음 후보를 계속 확인한다.
            }
            finally
            {
                RotHelper.ReleaseComObject(workbooks);
                if (candidate.ReleaseWhenDone && !applicationTransferred)
                    RotHelper.ReleaseComObject(candidate.Application);
            }
        }

        if (allowFileOpenFallback && matches.Count == 0 && Path.IsPathFullyQualified(reference) && File.Exists(reference))
        {
            try
            {
                dynamic app = attachedApplication;
                object? workbooks = null;
                try
                {
                    workbooks = (object)app.Workbooks;
                    object workbook = ((dynamic)workbooks).Open(reference, 0, true);
                    return new OpenWorkbook(attachedApplication, workbook, closeWhenDone: true);
                }
                finally { RotHelper.ReleaseComObject(workbooks); }
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"source workbook '{reference}' could not be opened read-only: {ex.Message}", ex);
            }
        }
        if (matches.Count == 0)
            throw new InvalidOperationException(allowFileOpenFallback
                ? $"open source workbook '{reference}' not found across running Excel instances and the file could not be opened"
                : $"open source workbook '{reference}' not found across running Excel instances; opening a disk file requires allowOpenFile=true with an absolute existing path");
        if (matches.Count > 1)
        {
            foreach (var match in matches) match.Dispose();
            throw new InvalidOperationException($"more than one open workbook matches '{reference}'; use the absolute sourceWorkbook path");
        }
        return matches[0];
    }

    private static bool PathsEqualWhenQualified(string reference, string fullName)
    {
        if (!Path.IsPathFullyQualified(reference) || !Path.IsPathFullyQualified(fullName)) return false;
        try { return string.Equals(Path.GetFullPath(reference), Path.GetFullPath(fullName), StringComparison.OrdinalIgnoreCase); }
        catch { return false; }
    }

    private static WorkbookLease ResolveTargetWorkbook(dynamic attachedApplication, dynamic defaultWorkbook,
        IReadOnlyList<JsonObject>? ops)
    {
        var references = (ops ?? Array.Empty<JsonObject>())
            .Select(op => Json.GetString(op, "targetWorkbook") ?? Json.GetString(Json.GetObj(op, "target"), "workbook"))
            .Where(reference => !string.IsNullOrWhiteSpace(reference))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (references.Count == 0) return new WorkbookLease((object)defaultWorkbook);
        if (references.Count > 1)
        {
            RotHelper.ReleaseComReference((object)defaultWorkbook);
            throw new InvalidOperationException("one Excel operation batch cannot target more than one destination workbook");
        }
        try
        {
            var opened = FindOpenWorkbook((object)attachedApplication, references[0]!, allowFileOpenFallback: false);
            return new WorkbookLease(opened.Workbook, opened);
        }
        finally { RotHelper.ReleaseComReference((object)defaultWorkbook); }
    }

    private static bool SheetExists(dynamic workbook, string sheetName)
    {
        try { _ = workbook.Worksheets.Item(sheetName); return true; }
        catch { return false; }
    }

    private static JsonNode? ToJsonValue(object? v) => v switch
    {
        null => null,
        string s => JsonValue.Create(s),
        bool b => JsonValue.Create(b),
        double d => JsonValue.Create(d),
        float f => JsonValue.Create((double)f),
        int i => JsonValue.Create(i),
        long l => JsonValue.Create(l),
        decimal m => JsonValue.Create((double)m),
        DateTime dt => JsonValue.Create(dt.ToString("o")),
        _ => JsonValue.Create(Convert.ToString(v, CultureInfo.InvariantCulture) ?? ""),
    };

    private static string ToDisp(object? v) => v switch
    {
        null => "",
        string s => s,
        bool b => b ? "TRUE" : "FALSE",
        DateTime dt => dt.ToString("yyyy-MM-dd HH:mm:ss"),
        IFormattable f => f.ToString(null, CultureInfo.InvariantCulture),
        _ => v.ToString() ?? "",
    };

    /// <summary>Range.Value2를 2차원 JsonArray로 (상한 적용). dynamic 전염 방지를 위해 object로 받는다.</summary>
    private static JsonArray RangeToJson(object rangeObj, out int cellCount, bool formulas = false, int maxCells = MaxCells)
    {
        dynamic range = rangeObj;
        cellCount = 0;
        object? raw = formulas ? range.Formula : range.Value2;
        var rows = new JsonArray();
        if (raw is object[,] arr)
        {
            var r1 = arr.GetLowerBound(0); var r2 = arr.GetUpperBound(0);
            var c1 = arr.GetLowerBound(1); var c2 = arr.GetUpperBound(1);
            for (var r = r1; r <= r2; r++)
            {
                if (cellCount >= maxCells) break;
                var jr = new JsonArray();
                for (var c = c1; c <= c2; c++)
                {
                    if (cellCount >= maxCells) break;
                    jr.Add(ToJsonValue(arr[r, c]));
                    cellCount++;
                }
                rows.Add(jr);
            }
        }
        else
        {
            rows.Add(new JsonArray { ToJsonValue(raw) });
            cellCount = 1;
        }
        return rows;
    }

    // ---------- IAppAdapter ----------

    public override AdapterStatus GetStatus()
    {
        try
        {
            return ComInvoke(() =>
            {
                var app = AttachExcel(allowCreate: false);
                if (app is null)
                    return new AdapterStatus(false, false, "excel", null, null,
                        "실행 중인 Excel 인스턴스가 없습니다");
                dynamic d = app;
                string? version = null; string? doc = null;
                try { version = (string)d.Version; } catch { }
                doc = ReadActiveWorkbookFullName(app);
                var detail = _ownsInstance
                    ? "DocBridge가 생성한 인스턴스"
                    : "사용자가 열어 둔 엑셀 창에 연결됨";
                return new AdapterStatus(true, true, "excel", version, doc, detail);
            });
        }
        catch (Exception ex)
        {
            return new AdapterStatus(false, false, "excel", null, null, ex.Message);
        }
    }

    public override ContextResult GetActiveContext()
    {
        return ComInvoke(() =>
        {
            var r = new ContextResult { App = App };
            object? workbook = null;
            object? activeSheet = null;
            object? worksheets = null;
            object? usedRange = null;
            object? selection = null;
            try
            {
                var app = AttachExcel();
                if (app is null)
                {
                    r.Errors.Add("Excel이 실행 중이지 않습니다. Excel을 열고 문서를 표시한 뒤 다시 시도하세요.");
                    return r;
                }
                dynamic d = app;
                workbook = (object)RequireWorkbook(d);
                activeSheet = (object)((dynamic)workbook).ActiveSheet;
                dynamic wb = workbook;
                dynamic sheet = activeSheet;

                r.DocumentRef = (string)wb.FullName;

                var sheetNames = new JsonArray();
                var sheetStates = new JsonArray();
                worksheets = (object)wb.Worksheets;
                var sheetCount = Convert.ToInt32(((dynamic)worksheets).Count, CultureInfo.InvariantCulture);
                for (var index = 1; index <= sheetCount; index++)
                {
                    object? candidateSheet = null;
                    try
                    {
                        candidateSheet = (object)((dynamic)worksheets).Item(index);
                        var candidateName = Convert.ToString(((dynamic)candidateSheet).Name, CultureInfo.InvariantCulture) ?? "";
                        var candidateVisibility = Convert.ToInt32(((dynamic)candidateSheet).Visible, CultureInfo.InvariantCulture);
                        sheetNames.Add(candidateName);
                        sheetStates.Add(new JsonObject
                        {
                            ["name"] = candidateName,
                            ["visibility"] = SheetVisibilityName(candidateVisibility),
                        });
                    }
                    finally { RotHelper.ReleaseComReference(candidateSheet); }
                }
                r.Summary["workbook"] = (string)wb.Name;
                r.Summary["sheets"] = sheetNames;
                r.Summary["sheetStates"] = sheetStates;
                r.Summary["activeSheet"] = (string)sheet.Name;
                usedRange = (object)sheet.UsedRange;
                dynamic used = usedRange;
                r.Summary["usedRange"] = $"{sheet.Name}!{(string)used.Address(false, false)}";
                r.Summary["saved"] = (bool)wb.Saved;
                r.Summary["openWorkbooks"] = ListOpenWorkbooks((object)d);

                selection = (object)d.Selection;
                dynamic sel = selection;
                var selAddr = (string)sel.Address(false, false);
                var selObj = new JsonObject { ["ref"] = $"{sheet.Name}!{selAddr}" };
                selObj["values"] = RangeToJson((object)sel, out _, formulas: false);
                selObj["formulas"] = RangeToJson((object)sel, out _, formulas: true);
                r.Selection = selObj;
                // Do not expose a partially populated context as success. Every required
                // field above must have completed without a COM/RCW error first.
                r.Ok = true;
            }
            catch (Exception ex)
            {
                r.Ok = false;
                r.Errors.Add($"excel context failed: {ex.Message}");
            }
            finally
            {
                RotHelper.ReleaseComObject(selection);
                RotHelper.ReleaseComObject(usedRange);
                RotHelper.ReleaseComObject(worksheets);
                RotHelper.ReleaseComObject(activeSheet);
                RotHelper.ReleaseComObject(workbook);
                // Also recover a zero-workbook owned shell left by an interrupted older call.
                // This branch can never quit a user-owned Excel instance.
                DisconnectOwnedExcelIfEmpty("context-no-workbook");
            }
            return r;
        });
    }

    public override JsonObject Read(JsonObject args)
    {
        return ComInvoke(() =>
        {
            try
            {
                var scope = (Json.GetString(args, "scope") ?? "range").ToLowerInvariant();
                var workbookRef = Json.GetString(args, "workbook");
                var allowOpenFile = Json.GetBool(args, "allowOpenFile");
                var explicitFile = !string.IsNullOrWhiteSpace(workbookRef) &&
                    Path.IsPathFullyQualified(workbookRef);
                if (explicitFile && !File.Exists(workbookRef!))
                    return Json.ErrorResult($"Excel workbook file not found: {workbookRef}", App);

                // Discovery, diagnostics, and active-workbook reads must never launch Excel.
                // A path alone is not consent to launch/open anything: AI clients often pass a
                // guessed path during discovery.  File opening requires the explicit opt-in and
                // an existing absolute file.  Any instance created for it is cleaned up below.
                var canOpenFile = scope != "diagnostics" && allowOpenFile && explicitFile;
                var app = AttachExcel(allowCreate: canOpenFile);
                if (app is null)
                    return Json.ErrorResult(
                        string.IsNullOrWhiteSpace(workbookRef)
                            ? "Excel not running; open a workbook first"
                            : explicitFile && !allowOpenFile
                                ? "Excel not running; open the workbook first or retry with allowOpenFile=true"
                                : "Excel not running; open the named workbook or provide an absolute existing file path with allowOpenFile=true",
                        App);
                dynamic d = app;
                if (scope == "diagnostics") return InspectExcelDiagnostics(d);
                OpenWorkbook? opened = null;
                object? workbookObject = null;
                if (string.IsNullOrWhiteSpace(workbookRef)) workbookObject = (object)RequireWorkbook(d);
                else
                {
                    opened = FindOpenWorkbook((object)d, workbookRef, allowFileOpenFallback: allowOpenFile);
                    workbookObject = opened.Workbook;
                }
                dynamic wb = workbookObject;
                try
                {
                    if (scope != "range") return InspectWorkbook(d, wb, args, scope);
                var sheetName = Json.GetString(args, "sheet");
                var rangeReference = Json.GetString(args, "range");
                if (string.IsNullOrWhiteSpace(rangeReference))
                    return Json.ErrorResult("excel_read_range requires 'range'", App);
                var resolvedRange = ResolveRangeTarget((object)wb, sheetName, rangeReference, requireExplicitSheet: false);
                object? sheetObject = resolvedRange.Sheet;
                object? rangeObject = null;
                try
                {
                    dynamic sheet = sheetObject;
                    var rangeAddr = resolvedRange.Address;
                    rangeObject = (object)sheet.Range(rangeAddr);
                    dynamic range = rangeObject;

                    var includeFormulas = Json.GetBool(args, "includeFormulas");
                    var includeStyles = Json.GetBool(args, "includeStyles");
                    var includeLayout = Json.GetBool(args, "includeLayout");

                    var result = new JsonObject
                    {
                        ["ok"] = true,
                        ["app"] = App,
                        ["workbook"] = (string)wb.FullName,
                        ["sheet"] = (string)sheet.Name,
                        ["range"] = rangeAddr,
                        ["values"] = RangeToJson(rangeObject, out var cells),
                    };
                    long totalCells;
                    try { totalCells = Convert.ToInt64(range.CountLarge, CultureInfo.InvariantCulture); }
                    catch { totalCells = Convert.ToInt64(range.Count, CultureInfo.InvariantCulture); }
                    result["truncated"] = totalCells > cells;
                    if (includeFormulas) result["formulas"] = RangeToJson(rangeObject, out _, formulas: true);
                    if (includeStyles)
                    {
                        object? font = null;
                        object? interior = null;
                        try
                        {
                            font = (object)range.Font;
                            interior = (object)range.Interior;
                            result["styles"] = new JsonObject
                            {
                                ["numberFormat"] = (string?)range.NumberFormat?.ToString(),
                                ["fontBold"] = (bool?)((dynamic)font).Bold,
                                ["fontItalic"] = (bool?)((dynamic)font).Italic,
                                ["interiorColor"] = (double?)((dynamic)interior).Color,
                            };
                        }
                        finally
                        {
                            RotHelper.ReleaseComReference(interior);
                            RotHelper.ReleaseComReference(font);
                        }
                    }
                    if (includeLayout) result["layout"] = ReadRangeLayout(sheetObject, rangeObject);
                    return result;
                }
                finally
                {
                    RotHelper.ReleaseComReference(rangeObject);
                    RotHelper.ReleaseComReference(sheetObject);
                }
                }
                finally
                {
                    if (opened is not null) opened.Dispose();
                    else RotHelper.ReleaseComObject(workbookObject);
                }
            }
            catch (Exception ex) { return ExcelErrorResult(ex); }
            finally
            {
                // FindOpenWorkbook closes file-fallback workbooks before this point.  If opening
                // failed, the owned application is also still empty.  In both cases remove the
                // grey shell immediately instead of waiting for an idle timer or server exit.
                DisconnectOwnedExcelIfEmpty("file-read-complete-or-failed");
            }
        });
    }

    private static JsonArray ListOpenWorkbooks(object attachedApplication)
    {
        var result = new JsonArray();
        var applications = new List<(object Application, bool ReleaseWhenDone)> { (attachedApplication, false) };
        foreach (var candidate in RotHelper.GetExcelApplications())
        {
            if (ReferenceEquals(candidate, attachedApplication))
            {
                RotHelper.ReleaseComReference(candidate);
                continue;
            }
            applications.Add((candidate, true));
        }
        var seenApps = new HashSet<long>();
        foreach (var candidate in applications)
        {
            object? activeWorkbook = null;
            object? workbooks = null;
            try
            {
                var appObj = candidate.Application;
                dynamic app = appObj;
                var hwnd = Convert.ToInt64(app.Hwnd, CultureInfo.InvariantCulture);
                if (!seenApps.Add(hwnd)) continue;
                var activeName = "";
                try
                {
                    activeWorkbook = (object?)app.ActiveWorkbook;
                    if (activeWorkbook is not null)
                        activeName = Convert.ToString(((dynamic)activeWorkbook).Name, CultureInfo.InvariantCulture) ?? "";
                }
                catch { }
                workbooks = (object)app.Workbooks;
                var count = Convert.ToInt32(((dynamic)workbooks).Count, CultureInfo.InvariantCulture);
                for (var index = 1; index <= count; index++)
                {
                    object? workbook = null;
                    object? worksheets = null;
                    try
                    {
                        workbook = (object)((dynamic)workbooks).Item(index);
                        worksheets = (object)((dynamic)workbook).Worksheets;
                        var sheetNames = new JsonArray();
                        var sheetCount = Convert.ToInt32(((dynamic)worksheets).Count, CultureInfo.InvariantCulture);
                        for (var sheetIndex = 1; sheetIndex <= sheetCount; sheetIndex++)
                        {
                            object? sheet = null;
                            try
                            {
                                sheet = (object)((dynamic)worksheets).Item(sheetIndex);
                                sheetNames.Add(Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture));
                            }
                            finally { RotHelper.ReleaseComReference(sheet); }
                        }
                        var workbookName = Convert.ToString(((dynamic)workbook).Name, CultureInfo.InvariantCulture);
                        result.Add(new JsonObject
                        {
                            ["excelHwnd"] = hwnd,
                            ["name"] = workbookName,
                            ["fullName"] = Convert.ToString(((dynamic)workbook).FullName, CultureInfo.InvariantCulture),
                            ["activeInInstance"] = string.Equals(activeName, workbookName, StringComparison.OrdinalIgnoreCase),
                            ["sheets"] = sheetNames,
                            ["saved"] = Convert.ToBoolean(((dynamic)workbook).Saved, CultureInfo.InvariantCulture),
                        });
                    }
                    finally
                    {
                        RotHelper.ReleaseComObject(worksheets);
                        RotHelper.ReleaseComReference(workbook);
                    }
                }
            }
            catch
            {
                // 닫히는 중이거나 모달 상태인 인스턴스는 목록에서 생략한다.
            }
            finally
            {
                RotHelper.ReleaseComObject(workbooks);
                RotHelper.ReleaseComReference(activeWorkbook);
                if (candidate.ReleaseWhenDone) RotHelper.ReleaseComObject(candidate.Application);
            }
        }
        return result;
    }

    // ---------- preview ----------

    public override ApplyPreview Preview(IReadOnlyList<JsonObject> ops)
    {
        return ComInvoke(() =>
        {
            var p = new ApplyPreview();
            var interaction = new ForegroundInteractionGuard(App);
            try
            {
                var app = AttachExcel();
                if (app is null) { p.Errors.Add("Excel not running"); return p; }
                dynamic d = app;
                using var workbookLease = ResolveTargetWorkbook(d, RequireWorkbook(d), ops);
                dynamic wb = workbookLease.Workbook;
                try { interaction.TrackTargetWindow(Convert.ToInt64(wb.Application.Hwnd, CultureInfo.InvariantCulture)); }
                catch { try { interaction.TrackTargetWindow(Convert.ToInt64(d.Hwnd, CultureInfo.InvariantCulture)); } catch { } }

                if (IsVisibilityOnlySnapshot(ops))
                {
                    foreach (var error in ValidateVisibilityBatch((object)wb, ops)) p.Errors.Add(error);
                    if (p.Errors.Count > 0) return p;
                }
                var visibilityPreviewState = IsVisibilityOnlySnapshot(ops)
                    ? new VisibilityPreviewState((object)wb)
                    : null;

                foreach (var op in ops)
                {
                    var name = Json.GetString(op, "op")!;
                    switch (name)
                    {
                        case "set_values": PreviewSetValues(wb, op, p, formulas: false); break;
                        case "set_formulas": PreviewSetValues(wb, op, p, formulas: true); break;
                        case "insert_rows":
                        {
                            var row = Json.GetInt(op, "row")!.Value;
                            var count = Json.GetInt(op, "count")!.Value;
                            dynamic sheet = GetRequiredTargetSheet(wb, op);
                            p.Affected.Add(new AffectedRef("rows", $"{sheet.Name}!{row}:{row + count - 1}"));
                            p.Diff.Add(new DiffEntry { Ref = "insert", Before = $"row {row}", After = $"{count} new row(s)" });
                            break;
                        }
                        case "insert_cols":
                        {
                            var count = Json.GetInt(op, "count")!.Value;
                            var colNode = op["col"];
                            dynamic sheet = GetRequiredTargetSheet(wb, op);
                            p.Affected.Add(new AffectedRef("cols", $"{sheet.Name}!col {colNode}"));
                            p.Diff.Add(new DiffEntry { Ref = "insert", Before = $"col {colNode}", After = $"{count} new col(s)" });
                            break;
                        }
                        case "format_range":
                        {
                            var resolvedRange = ResolveRangeTarget(
                                (object)wb,
                                Json.GetString(Json.GetObj(op, "target"), "sheet"),
                                Json.GetString(op, "range")!,
                                requireExplicitSheet: true);
                            dynamic sheet = resolvedRange.Sheet;
                            var rangeAddr = resolvedRange.Address;
                            p.Affected.Add(new AffectedRef("range", $"{sheet.Name}!{rangeAddr}"));
                            var style = Json.GetObj(op, "style") ?? new JsonObject();
                            p.Diff.Add(new DiffEntry { Ref = "style", Before = "current", After = style.DeepClone() });
                            break;
                        }
                        case "find_replace": PreviewFindReplace(d, wb, op, p); break;
                        case "copy_sheet": PreviewCopySheet(d, wb, op, p); break;
                        case "merge_cells":
                        case "unmerge_cells":
                            PreviewMergeOperation((object)wb, op, p);
                            break;
                        case "set_rows_hidden":
                        case "set_cols_hidden":
                        case "set_sheet_visibility":
                            PreviewVisibilityOperation((object)wb, op, p,
                                visibilityPreviewState ?? throw new InvalidOperationException(
                                    "visibility preview state was not initialized"));
                            break;
                    }
                    if (!interaction.Checkpoint(stopOnConcurrentInput: true))
                    {
                        p.Errors.Add("[APP_USER_ACTIVITY_DETECTED] 사용자가 Excel 창을 조작하여 미리보기를 중단했습니다. 해당 창 작업을 마친 뒤 다시 실행하세요.");
                        break;
                    }
                }
            }
            catch (Exception ex) { p.Errors.Add($"preview failed: {ex.Message}"); }
            finally { p.Interaction = interaction.Complete(); }
            return p;
        });
    }

    private static void PreviewSetValues(dynamic wb, JsonObject op, ApplyPreview p, bool formulas)
    {
        var rangeReference = Json.GetString(op, "range")!;
        var values = Json.GetArr(op, formulas ? "formulas" : "values")!;
        var sheetName = Json.GetString(Json.GetObj(op, "target"), "sheet");
        var resolvedRange = ResolveRangeTarget((object)wb, sheetName, rangeReference, requireExplicitSheet: true);
        dynamic sheet = resolvedRange.Sheet;
        var rangeAddr = resolvedRange.Address;
        dynamic range = sheet.Range(rangeAddr);

        if (values.Count == 0)
        {
            p.Errors.Add($"{(formulas ? "formulas" : "values")} must be a non-empty 2D array");
            return;
        }
        var rows = values.Count;
        var cols = values[0] is JsonArray firstRow ? firstRow.Count : 0;
        if (cols == 0 || values.Any(node => node is not JsonArray row || row.Count != cols))
        {
            p.Errors.Add($"{(formulas ? "formulas" : "values")} must be a rectangular non-empty 2D array");
            return;
        }
        var targetRows = (int)range.Rows.Count;
        var targetCols = (int)range.Columns.Count;
        if (rows != targetRows || cols != targetCols)
        {
            p.Errors.Add($"{sheet.Name}!{rangeAddr} is {targetRows}x{targetCols}, but input is {rows}x{cols}");
            return;
        }

        // 현재 값을 bulk read
        object? raw = formulas ? range.Formula : range.Value2;
        p.Affected.Add(new AffectedRef("range", $"{sheet.Name}!{rangeAddr}"));

        var r0 = range.Row; var c0 = range.Column;
        for (var i = 0; i < rows; i++)
        {
            if (values[i] is not JsonArray rowArr) continue;
            for (var j = 0; j < rowArr.Count; j++)
            {
                // raw는 1-based object[,] (또는 단일 값)
                object? before = raw is object[,] a2 ? a2[i + 1, j + 1] : raw;
                var afterNode = rowArr[j];
                if (p.Diff.Count >= MaxDiff) { p.DiffTruncated = true; continue; }
                p.Diff.Add(new DiffEntry
                {
                    Ref = $"{sheet.Name}!{CellName(c0 + j, r0 + i)}",
                    Before = ToJsonValue(before),
                    After = afterNode?.DeepClone(),
                });
            }
        }
    }

    private static string CellName(int col, int row) => $"{ColName(col)}{row}";

    private static void PreviewFindReplace(dynamic app, dynamic wb, JsonObject op, ApplyPreview p)
    {
        var find = Json.GetString(op, "find")!;
        var replace = Json.GetString(op, "replace")!;
        var matchCase = Json.GetBool(Json.GetObj(op, "options"), "matchCase");
        var cmp = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var scope = Json.GetString(Json.GetObj(op, "target"), "scope") ?? "sheet";

        var sheets = new List<dynamic>();
        if (scope == "workbook")
            foreach (dynamic s in wb.Worksheets) sheets.Add(s);
        else
            sheets.Add(GetRequiredTargetSheet(wb, op));

        var matches = 0;
        foreach (var sheet in sheets)
        {
            dynamic used = sheet.UsedRange;
            object? raw = used.Value2;
            var usedRow = (int)used.Row;
            var usedCol = (int)used.Column;
            if (raw is object[,] arr)
            {
                var lr = arr.GetLowerBound(0);
                var lc = arr.GetLowerBound(1);
                for (var r = arr.GetLowerBound(0); r <= arr.GetUpperBound(0); r++)
                    for (var c = arr.GetLowerBound(1); c <= arr.GetUpperBound(1); c++)
                    {
                        if (arr[r, c] is string s && s.Contains(find, cmp))
                        {
                            matches++;
                            if (p.Diff.Count < MaxDiff)
                                p.Diff.Add(new DiffEntry
                                {
                                    Ref = $"{sheet.Name}!{CellName(usedCol + c - lc, usedRow + r - lr)}",
                                    Before = s,
                                    After = s.Replace(find, replace, cmp),
                                });
                        }
                    }
            }
            else if (raw is string single && single.Contains(find, cmp))
            {
                matches++;
                p.Diff.Add(new DiffEntry { Ref = $"{sheet.Name}!{CellName(usedCol, usedRow)}", Before = single, After = single.Replace(find, replace, cmp) });
            }
        }
        p.Affected.Add(new AffectedRef("matches", $"{matches} cell(s), scope={scope}"));
        if (matches == 0) p.Warnings.Add($"no matches for '{find}' (scope={scope})");
    }

    private static void PreviewCopySheet(dynamic destinationApp, dynamic destinationWorkbook,
        JsonObject op, ApplyPreview preview)
    {
        var sourceWorkbookRef = Json.GetString(op, "sourceWorkbook")!;
        var sourceSheetName = Json.GetString(op, "sourceSheet")!;
        var targetSheetName = Json.GetString(op, "targetSheet") ?? sourceSheetName;
        var source = FindOpenWorkbook((object)destinationApp, sourceWorkbookRef);
        try
        {
            dynamic sourceWorkbook = source.Workbook;
            dynamic sourceSheet = GetSheet(sourceWorkbook, sourceSheetName);
            object? sourceUsedRange = null;

            try
            {
                if (SheetExists(destinationWorkbook, targetSheetName))
                {
                    preview.Errors.Add($"destination workbook already contains sheet '{targetSheetName}'");
                    return;
                }

                var sourceFullName = Convert.ToString(sourceWorkbook.FullName, CultureInfo.InvariantCulture) ?? sourceWorkbookRef;
                sourceUsedRange = (object)sourceSheet.UsedRange;
                var usedAddress = Convert.ToString(((dynamic)sourceUsedRange).Address(false, false), CultureInfo.InvariantCulture) ?? "";
                preview.Affected.Add(new AffectedRef("source-sheet", $"{sourceFullName}::{sourceSheetName}"));
                preview.Affected.Add(new AffectedRef("target-sheet", $"{destinationWorkbook.Name}::{targetSheetName}"));
                preview.Diff.Add(new DiffEntry
                {
                    Ref = $"sheet:{targetSheetName}",
                    Before = null,
                    After = JsonValue.Create($"copied from {sourceWorkbook.Name}::{sourceSheetName} ({usedAddress})"),
                });
            }
            finally
            {
                RotHelper.ReleaseComObject(sourceUsedRange);
                RotHelper.ReleaseComObject((object)sourceSheet);
            }
        }
        finally
        {
            source.Dispose();
        }
    }

    // ---------- apply ----------

    public override ApplyExecution Apply(IReadOnlyList<JsonObject> ops, string snapshotId)
    {
        return ComInvoke(() =>
        {
            var exec = new ApplyExecution { Ok = true };
            var mismatches = new List<string>();
            var checkedCells = 0;
            var interaction = new ForegroundInteractionGuard(App);
            dynamic? targetApplication = null;
            dynamic? originalWorkbook = null;
            dynamic? originalSheet = null;
            var internalDocumentSwitched = false;
            var originalStateRestored = true;
            try
            {
                var app = AttachExcel();
                if (app is null) { exec.Errors.Add("Excel not running"); exec.Ok = false; return exec; }
                dynamic d = app;
                using var workbookLease = ResolveTargetWorkbook(d, RequireWorkbook(d), ops);
                dynamic wb = workbookLease.Workbook;
                targetApplication = wb.Application;
                try { interaction.TrackTargetWindow(Convert.ToInt64(targetApplication.Hwnd, CultureInfo.InvariantCulture)); } catch { }
                try { originalWorkbook = targetApplication.ActiveWorkbook; } catch { }
                try { originalSheet = originalWorkbook?.ActiveSheet; } catch { }
                if (IsVisibilityOnlySnapshot(ops))
                {
                    var visibilityErrors = ValidateVisibilityBatch((object)wb, ops);
                    if (visibilityErrors.Count > 0)
                        throw new InvalidOperationException(string.Join("; ", visibilityErrors));
                }
                targetApplication.ScreenUpdating = false;
                try
                {
                    foreach (var op in ops)
                    {
                        var name = Json.GetString(op, "op")!;
                        switch (name)
                        {
                            case "set_values": ApplySetValues(wb, op, exec, mismatches, ref checkedCells, formulas: false); break;
                            case "set_formulas": ApplySetValues(wb, op, exec, mismatches, ref checkedCells, formulas: true); break;
                            case "insert_rows":
                            {
                                var row = Json.GetInt(op, "row")!.Value;
                                var count = Json.GetInt(op, "count")!.Value;
                                dynamic sheet = GetRequiredTargetSheet(wb, op);
                                sheet.Rows[$"{row}:{row + count - 1}"].Insert();
                                exec.Affected.Add(new AffectedRef("rows", $"{sheet.Name}!{row}:{row + count - 1}"));
                                break;
                            }
                            case "insert_cols":
                            {
                                var count = Json.GetInt(op, "count")!.Value;
                                var col = op["col"]!;
                                dynamic sheet = GetRequiredTargetSheet(wb, op);
                                string colName = col is JsonValue jv && jv.TryGetValue<int>(out var ci) ? ColName(ci) : col.GetValue<string>();
                                var endCol = ColName(ColIndex(colName) + count - 1);
                                sheet.Columns[$"{colName}:{endCol}"].Insert();
                                exec.Affected.Add(new AffectedRef("cols", $"{sheet.Name}!{colName}:{endCol}"));
                                break;
                            }
                            case "format_range": ApplyFormat(wb, op, exec, mismatches, ref checkedCells); break;
                            case "find_replace": ApplyFindReplace(wb, op, exec, mismatches, ref checkedCells); break;
                            case "copy_sheet": ApplyCopySheet(targetApplication, wb, op, exec, mismatches, ref checkedCells, interaction); break;
                            case "merge_cells":
                            case "unmerge_cells":
                                ApplyMergeOperation((object)wb, op, exec, mismatches, ref checkedCells);
                                break;
                            case "set_rows_hidden":
                            case "set_cols_hidden":
                            case "set_sheet_visibility":
                                ApplyVisibilityOperation((object)wb, op, exec, mismatches, ref checkedCells);
                                break;
                        }

                        if (!interaction.Checkpoint(stopOnConcurrentInput: true))
                            throw new InvalidOperationException(
                                "[APP_USER_ACTIVITY_DETECTED] Excel이 작업 중 전경으로 전환되어 다음 작업을 중단했습니다");
                    }
                }
                finally { try { targetApplication.ScreenUpdating = true; } catch { } }

                exec.Readback = new JsonObject
                {
                    ["verified"] = mismatches.Count == 0,
                    ["checked"] = checkedCells,
                    ["mismatches"] = Json.ToArray(mismatches),
                    ["snapshotId"] = snapshotId,
                };
                exec.Ok = mismatches.Count == 0;
            }
            catch (Exception ex)
            {
                exec.Ok = false;
                exec.Errors.Add($"apply failed: {ex.Message}");
            }
            finally
            {
                if (targetApplication is not null && originalWorkbook is not null)
                {
                    try
                    {
                        dynamic? activeWorkbook = targetApplication.ActiveWorkbook;
                        dynamic? activeSheet = activeWorkbook?.ActiveSheet;
                        var workbookChanged = !string.Equals(
                            Convert.ToString(activeWorkbook?.Name, CultureInfo.InvariantCulture),
                            Convert.ToString(originalWorkbook.Name, CultureInfo.InvariantCulture),
                            StringComparison.OrdinalIgnoreCase);
                        var sheetChanged = !string.Equals(
                            Convert.ToString(activeSheet?.Name, CultureInfo.InvariantCulture),
                            Convert.ToString(originalSheet?.Name, CultureInfo.InvariantCulture),
                            StringComparison.OrdinalIgnoreCase);
                        internalDocumentSwitched = workbookChanged || sheetChanged;
                        if (internalDocumentSwitched)
                        {
                            originalWorkbook.Activate();
                            if (originalSheet is not null) originalSheet.Activate();
                        }
                        originalStateRestored = string.Equals(
                            Convert.ToString(targetApplication.ActiveWorkbook?.Name, CultureInfo.InvariantCulture),
                            Convert.ToString(originalWorkbook.Name, CultureInfo.InvariantCulture),
                            StringComparison.OrdinalIgnoreCase) &&
                            (originalSheet is null || string.Equals(
                                Convert.ToString(targetApplication.ActiveWorkbook?.ActiveSheet?.Name, CultureInfo.InvariantCulture),
                                Convert.ToString(originalSheet.Name, CultureInfo.InvariantCulture),
                                StringComparison.OrdinalIgnoreCase));
                    }
                    catch
                    {
                        originalStateRestored = false;
                    }
                }
                var telemetry = interaction.Complete();
                telemetry["internalDocumentSwitched"] = internalDocumentSwitched;
                telemetry["originalStateRestored"] = originalStateRestored;
                exec.Interaction = telemetry;
            }
            return exec;
        });
    }

    private static void ApplySetValues(dynamic wb, JsonObject op, ApplyExecution exec,
        List<string> mismatches, ref int checkedCells, bool formulas)
    {
        var rangeReference = Json.GetString(op, "range")!;
        var values = Json.GetArr(op, formulas ? "formulas" : "values")!;
        var resolvedRange = ResolveRangeTarget(
            (object)wb,
            Json.GetString(Json.GetObj(op, "target"), "sheet"),
            rangeReference,
            requireExplicitSheet: true);
        dynamic sheet = resolvedRange.Sheet;
        var rangeAddr = resolvedRange.Address;
        dynamic range = sheet.Range(rangeAddr);

        var rows = values.Count;
        var cols = values[0] is JsonArray ja0 ? ja0.Count : 1;
        var data = new object?[rows, cols];
        for (var i = 0; i < rows; i++)
        {
            if (values[i] is not JsonArray rowArr) continue;
            for (var j = 0; j < rowArr.Count; j++)
            {
                data[i, j] = NodeToComValue(rowArr[j]);
            }
        }

        if (formulas) range.Formula = data;
        else range.Value2 = data;
        exec.Affected.Add(new AffectedRef("range", $"{sheet.Name}!{rangeAddr}"));

        // readback
        object? raw = formulas ? range.Formula : range.Value2;
        for (var i = 0; i < rows; i++)
            for (var j = 0; j < cols; j++)
            {
                checkedCells++;
                var want = data[i, j];
                object? got = raw is object[,] arr ? arr[i + 1, j + 1] : raw;
                if (!ComValuesEqual(want, got))
                    mismatches.Add($"{CellName(range.Column + j, range.Row + i)}: want '{ToDisp(want)}', got '{ToDisp(got)}'");
            }
    }

    private static void ApplyCopySheet(dynamic destinationApp, dynamic destinationWorkbook,
        JsonObject op, ApplyExecution exec, List<string> mismatches, ref int checkedCells,
        ForegroundInteractionGuard interaction)
    {
        var sourceWorkbookRef = Json.GetString(op, "sourceWorkbook")!;
        var sourceSheetName = Json.GetString(op, "sourceSheet")!;
        var targetSheetName = Json.GetString(op, "targetSheet") ?? sourceSheetName;
        var source = FindOpenWorkbook((object)destinationApp, sourceWorkbookRef);
        dynamic? sourceAppForRestore = null;
        dynamic? sourceWorkbookForRestore = null;
        dynamic? sourceSheetForRestore = null;
        var restoreSeparateSource = false;
        try
        {
            dynamic sourceApp = source.Application;
            dynamic sourceWorkbook = source.Workbook;
            dynamic sourceSheet = GetSheet(sourceWorkbook, sourceSheetName);
            try { interaction.TrackTargetWindow(Convert.ToInt64(sourceApp.Hwnd, CultureInfo.InvariantCulture)); } catch { }

            try
            {
                restoreSeparateSource = Convert.ToInt64(sourceApp.Hwnd, CultureInfo.InvariantCulture) !=
                    Convert.ToInt64(destinationApp.Hwnd, CultureInfo.InvariantCulture);
                if (restoreSeparateSource)
                {
                    sourceAppForRestore = sourceApp;
                    sourceWorkbookForRestore = sourceApp.ActiveWorkbook;
                    sourceSheetForRestore = sourceWorkbookForRestore?.ActiveSheet;
                }
            }
            catch { restoreSeparateSource = false; }

            if (SheetExists(destinationWorkbook, targetSheetName))
                throw new InvalidOperationException($"destination workbook already contains sheet '{targetSheetName}'");

            dynamic copiedSheet;
            var sourceHwnd = Convert.ToInt64(sourceApp.Hwnd, CultureInfo.InvariantCulture);
            var destinationHwnd = Convert.ToInt64(destinationApp.Hwnd, CultureInfo.InvariantCulture);
            if (sourceHwnd == destinationHwnd)
            {
                dynamic after = destinationWorkbook.Worksheets.Item(destinationWorkbook.Worksheets.Count);
                sourceSheet.Copy(After: after);
                copiedSheet = destinationWorkbook.ActiveSheet;
            }
            else
            {
                copiedSheet = CopySheetBetweenExcelInstances(
                    sourceApp, sourceSheet, destinationApp, destinationWorkbook);
            }

            if (!string.Equals((string)copiedSheet.Name, targetSheetName, StringComparison.Ordinal))
                copiedSheet.Name = targetSheetName;

            VerifyCopiedSheet(sourceSheet, copiedSheet, mismatches, ref checkedCells);

            exec.Affected.Add(new AffectedRef("sheet", $"{destinationWorkbook.Name}::{targetSheetName}"));
            exec.Diff.Add(new DiffEntry
            {
                Ref = $"sheet:{targetSheetName}",
                Before = null,
                After = JsonValue.Create($"copied from {sourceWorkbook.Name}::{sourceSheetName}"),
            });
        }
        finally
        {
            if (restoreSeparateSource && sourceAppForRestore is not null && sourceWorkbookForRestore is not null)
            {
                try
                {
                    sourceWorkbookForRestore.Activate();
                    if (sourceSheetForRestore is not null) sourceSheetForRestore.Activate();
                }
                catch { }
            }
            RotHelper.ReleaseComObject((object?)sourceSheetForRestore);
            RotHelper.ReleaseComObject((object?)sourceWorkbookForRestore);
            source.Dispose();
        }
    }

    private static dynamic CopySheetBetweenExcelInstances(dynamic sourceApp, dynamic sourceSheet,
        dynamic destinationApp, dynamic destinationWorkbook)
    {
        var tempDir = Path.Combine(Path.GetTempPath(), "DocBridge", "sheet-transfer");
        Directory.CreateDirectory(tempDir);
        var tempPath = Path.Combine(tempDir, $"sheet-{Guid.NewGuid():N}.xlsx");
        dynamic? exportWorkbook = null;
        dynamic? importWorkbook = null;
        var sourceAlerts = true;
        var destinationAlerts = true;
        try
        {
            try { sourceAlerts = Convert.ToBoolean(sourceApp.DisplayAlerts, CultureInfo.InvariantCulture); } catch { }
            try { destinationAlerts = Convert.ToBoolean(destinationApp.DisplayAlerts, CultureInfo.InvariantCulture); } catch { }
            sourceApp.DisplayAlerts = false;
            destinationApp.DisplayAlerts = false;

            sourceSheet.Copy();
            exportWorkbook = sourceApp.ActiveWorkbook;
            exportWorkbook.SaveAs(tempPath, 51); // xlOpenXMLWorkbook (.xlsx)
            exportWorkbook.Close(false);
            exportWorkbook = null;

            importWorkbook = destinationApp.Workbooks.Open(tempPath, 0, true);
            dynamic importedSheet = importWorkbook.Worksheets.Item(1);
            dynamic after = destinationWorkbook.Worksheets.Item(destinationWorkbook.Worksheets.Count);
            importedSheet.Copy(After: after);
            dynamic copied = destinationWorkbook.ActiveSheet;
            importWorkbook.Close(false);
            importWorkbook = null;
            return copied;
        }
        finally
        {
            try { if (exportWorkbook is not null) exportWorkbook.Close(false); } catch { }
            try { if (importWorkbook is not null) importWorkbook.Close(false); } catch { }
            try { sourceApp.DisplayAlerts = sourceAlerts; } catch { }
            try { destinationApp.DisplayAlerts = destinationAlerts; } catch { }
            try { if (File.Exists(tempPath)) File.Delete(tempPath); } catch { }
        }
    }

    private static void VerifyCopiedSheet(dynamic sourceSheet, dynamic copiedSheet,
        List<string> mismatches, ref int checkedCells)
    {
        dynamic sourceUsed = sourceSheet.UsedRange;
        dynamic copiedUsed = copiedSheet.UsedRange;
        var sourceAddress = Convert.ToString(sourceUsed.Address(false, false), CultureInfo.InvariantCulture) ?? "";
        var copiedAddress = Convert.ToString(copiedUsed.Address(false, false), CultureInfo.InvariantCulture) ?? "";
        checkedCells++;
        if (!string.Equals(sourceAddress, copiedAddress, StringComparison.OrdinalIgnoreCase))
            mismatches.Add($"sheet used range mismatch: source={sourceAddress}, copied={copiedAddress}");

        var rows = Convert.ToInt32(sourceUsed.Rows.Count, CultureInfo.InvariantCulture);
        var cols = Convert.ToInt32(sourceUsed.Columns.Count, CultureInfo.InvariantCulture);
        var total = (long)rows * cols;
        if (total > MaxSnapshotCells)
        {
            mismatches.Add($"copied sheet verification exceeds {MaxSnapshotCells} cells");
            return;
        }

        object? sourceFormulas = sourceUsed.Formula;
        object? copiedFormulas = copiedUsed.Formula;
        for (var row = 0; row < rows; row++)
        {
            for (var col = 0; col < cols; col++)
            {
                var expected = sourceFormulas is object[,] sa ? sa[row + 1, col + 1] : sourceFormulas;
                var actual = copiedFormulas is object[,] ca ? ca[row + 1, col + 1] : copiedFormulas;
                checkedCells++;
                if (!ComValuesEqual(expected, actual))
                    mismatches.Add($"{copiedSheet.Name}!{CellName(copiedUsed.Column + col, copiedUsed.Row + row)}: copied value/formula mismatch");
            }
        }

        for (var col = 1; col <= cols; col++)
        {
            var expectedWidth = Convert.ToDouble(sourceUsed.Columns.Item(col).ColumnWidth, CultureInfo.InvariantCulture);
            var actualWidth = Convert.ToDouble(copiedUsed.Columns.Item(col).ColumnWidth, CultureInfo.InvariantCulture);
            checkedCells++;
            if (Math.Abs(expectedWidth - actualWidth) > 0.01)
                mismatches.Add($"{copiedSheet.Name}: column {col} width mismatch ({expectedWidth} != {actualWidth})");
        }
    }

    private static void ApplyFormat(dynamic wb, JsonObject op, ApplyExecution exec,
        List<string> mismatches, ref int checkedCells)
    {
        var resolvedRange = ResolveRangeTarget(
            (object)wb,
            Json.GetString(Json.GetObj(op, "target"), "sheet"),
            Json.GetString(op, "range")!,
            requireExplicitSheet: true);
        dynamic sheet = resolvedRange.Sheet;
        var rangeAddr = resolvedRange.Address;
        dynamic range = sheet.Range(rangeAddr);
        var style = Json.GetObj(op, "style") ?? new JsonObject();

        if (style.TryGetPropertyValue("bold", out var b)) range.Font.Bold = b!.GetValue<bool>();
        if (style.TryGetPropertyValue("italic", out var it)) range.Font.Italic = it!.GetValue<bool>();
        if (style.TryGetPropertyValue("fontSize", out var fs)) range.Font.Size = fs!.GetValue<double>();
        if (style.TryGetPropertyValue("numberFormat", out var nf)) range.NumberFormat = nf!.GetValue<string>();
        if (style.TryGetPropertyValue("fontColor", out var fc) && fc is not null)
            range.Font.Color = ParseColor(fc);
        if (style.TryGetPropertyValue("fillColor", out var fill) && fill is not null)
            range.Interior.Color = ParseColor(fill);

        exec.Affected.Add(new AffectedRef("range", $"{sheet.Name}!{rangeAddr}"));

        // 적용한 각 속성을 즉시 다시 읽어 성공 여부를 확인한다.
        foreach (var (key, node) in style)
        {
            if (node is null) continue;
            checkedCells++;
            try
            {
                var matched = key switch
                {
                    "bold" => Convert.ToBoolean(range.Font.Bold, CultureInfo.InvariantCulture) == node.GetValue<bool>(),
                    "italic" => Convert.ToBoolean(range.Font.Italic, CultureInfo.InvariantCulture) == node.GetValue<bool>(),
                    "fontSize" => Math.Abs(Convert.ToDouble(range.Font.Size, CultureInfo.InvariantCulture) - node.GetValue<double>()) < 1e-9,
                    "numberFormat" => string.Equals(Convert.ToString(range.NumberFormat, CultureInfo.InvariantCulture), node.GetValue<string>(), StringComparison.Ordinal),
                    "fontColor" => Math.Abs(Convert.ToDouble(range.Font.Color, CultureInfo.InvariantCulture) - ParseColor(node)) < 1e-9,
                    "fillColor" => Math.Abs(Convert.ToDouble(range.Interior.Color, CultureInfo.InvariantCulture) - ParseColor(node)) < 1e-9,
                    _ => false,
                };
                if (!matched) mismatches.Add($"{sheet.Name}!{rangeAddr}: style '{key}' readback mismatch");
            }
            catch (Exception ex)
            {
                mismatches.Add($"{sheet.Name}!{rangeAddr}: style '{key}' readback failed: {ex.Message}");
            }
        }
    }

    private static double ParseColor(JsonNode node)
    {
        if (node is JsonValue jv)
        {
            if (jv.TryGetValue<int>(out var ole)) return ole;
            if (jv.TryGetValue<string>(out var hex) && hex.StartsWith('#') && hex.Length == 7)
            {
                var r = Convert.ToInt32(hex[1..3], 16);
                var g = Convert.ToInt32(hex[3..5], 16);
                var b = Convert.ToInt32(hex[5..7], 16);
                return r | (g << 8) | (b << 16); // OLE COLORREF (BGR 순서 아님: RGB 패킹)
            }
        }
        throw new ArgumentException("color must be OLE int or '#RRGGBB'");
    }

    private static void ApplyFindReplace(dynamic wb, JsonObject op, ApplyExecution exec,
        List<string> mismatches, ref int checkedCells)
    {
        var find = Json.GetString(op, "find")!;
        var replace = Json.GetString(op, "replace")!;
        var matchCase = Json.GetBool(Json.GetObj(op, "options"), "matchCase");
        var cmp = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
        var scope = Json.GetString(Json.GetObj(op, "target"), "scope") ?? "sheet";

        var sheets = new List<dynamic>();
        if (scope == "workbook")
            foreach (dynamic s in wb.Worksheets) sheets.Add(s);
        else
            sheets.Add(GetRequiredTargetSheet(wb, op));

        var replaced = 0;
        foreach (var sheet in sheets)
        {
            dynamic used = sheet.UsedRange;
            object? raw = used.Value2;
            var usedRow = (int)used.Row;
            var usedCol = (int)used.Column;
            if (raw is object[,] arr)
            {
                var lr = arr.GetLowerBound(0);
                var lc = arr.GetLowerBound(1);

                // 수식 셀 보호: HasFormula 영역은 건드리지 않는다
                for (var r = arr.GetLowerBound(0); r <= arr.GetUpperBound(0); r++)
                    for (var c = arr.GetLowerBound(1); c <= arr.GetUpperBound(1); c++)
                    {
                        if (arr[r, c] is not string s || !s.Contains(find, cmp)) continue;
                        dynamic cell = used.Cells[r, c];
                        bool hasFormula = false;
                        try { hasFormula = (bool)cell.HasFormula; } catch { }
                        if (hasFormula)
                        {
                            exec.Warnings.Add($"skipped formula cell {sheet.Name}!{CellName(usedCol + c - lc, usedRow + r - lr)}");
                            continue;
                        }
                        var nv = s.Replace(find, replace, cmp);
                        cell.Value2 = nv;
                        replaced++;
                        checkedCells++;
                        var cellRef = $"{sheet.Name}!{CellName(usedCol + c - lc, usedRow + r - lr)}";
                        if (exec.Diff.Count < MaxDiff)
                            exec.Diff.Add(new DiffEntry { Ref = cellRef, Before = s, After = nv });

                        var got = (string?)cell.Value2?.ToString() ?? "";
                        if (got != nv) mismatches.Add($"{cellRef}: want '{nv}', got '{got}'");
                    }
            }
            else if (raw is string single && single.Contains(find, cmp))
            {
                bool hasFormula = false;
                try { hasFormula = (bool)used.HasFormula; } catch { }
                var cellRef = $"{sheet.Name}!{CellName(usedCol, usedRow)}";
                if (hasFormula)
                {
                    exec.Warnings.Add($"skipped formula cell {cellRef}");
                }
                else
                {
                    var nv = single.Replace(find, replace, cmp);
                    used.Value2 = nv;
                    replaced++;
                    checkedCells++;
                    if (exec.Diff.Count < MaxDiff)
                        exec.Diff.Add(new DiffEntry { Ref = cellRef, Before = single, After = nv });
                    var got = (string?)used.Value2?.ToString() ?? "";
                    if (got != nv) mismatches.Add($"{cellRef}: want '{nv}', got '{got}'");
                }
            }
        }
        exec.Affected.Add(new AffectedRef("matches", $"{replaced} cell(s) replaced"));
        if (replaced == 0) exec.Warnings.Add($"no matches for '{find}' (scope={scope})");
    }

    // ---------- snapshot / restore ----------

    private const int CurrentExcelSnapshotVersion = 2;
    private const int MaxRestoreMismatchSamples = 100;
    private const string CopySheetTopologyRestoreMode = "copy-sheet-topology";
    private const string LegacyFullRangeRestoreMode = "legacy-full-range";

    private sealed class RestoreMismatchCollector
    {
        private readonly List<string> _samples = new(MaxRestoreMismatchSamples);

        public int Count { get; private set; }
        public IReadOnlyList<string> Samples => _samples;
        public bool Truncated => Count > _samples.Count;

        public void Add(string value)
        {
            Count++;
            if (_samples.Count < MaxRestoreMismatchSamples) _samples.Add(value);
        }
    }

    private static bool IsCopySheetOnlySnapshot(IReadOnlyList<JsonObject>? ops) =>
        ops is { Count: > 0 } &&
        ops.All(op => string.Equals(Json.GetString(op, "op"), "copy_sheet", StringComparison.OrdinalIgnoreCase));

    private static List<string> ReadOrderedWorksheetNames(object workbook)
    {
        var result = new List<string>();
        object? worksheets = null;
        try
        {
            dynamic wb = workbook;
            worksheets = (object)wb.Worksheets;
            var count = Convert.ToInt32(((dynamic)worksheets).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                object? sheet = null;
                try
                {
                    sheet = (object)((dynamic)worksheets).Item(index);
                    result.Add(Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture) ?? "");
                }
                finally { RotHelper.ReleaseComReference(sheet); }
            }
        }
        finally { RotHelper.ReleaseComReference(worksheets); }
        return result;
    }

    private static string ReadActiveWorksheetName(object workbook)
    {
        object? activeSheet = null;
        try
        {
            activeSheet = (object)((dynamic)workbook).ActiveSheet;
            return Convert.ToString(((dynamic)activeSheet).Name, CultureInfo.InvariantCulture) ?? "";
        }
        finally { RotHelper.ReleaseComReference(activeSheet); }
    }

    private static JsonArray StringsToJsonArray(IEnumerable<string> values)
    {
        var result = new JsonArray();
        foreach (var value in values) result.Add(value);
        return result;
    }

    private static JsonObject CaptureCopySheetTopologyState(object workbook, IReadOnlyList<JsonObject> ops,
        string? documentRef)
    {
        var originalSheets = ReadOrderedWorksheetNames(workbook);
        var originalActiveSheet = ReadActiveWorksheetName(workbook);
        var originalSet = originalSheets.ToHashSet(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(originalActiveSheet) || !originalSet.Contains(originalActiveSheet))
            throw new InvalidOperationException("copy_sheet snapshot requires an active worksheet in the target workbook");
        var targetSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var targetSheets = new List<string>();
        var capturedOps = new JsonArray();

        foreach (var op in ops)
        {
            capturedOps.Add(op.DeepClone());
            var target = Json.GetString(op, "targetSheet") ?? Json.GetString(op, "sourceSheet");
            if (string.IsNullOrWhiteSpace(target))
                throw new InvalidOperationException("copy_sheet snapshot requires sourceSheet or targetSheet");
            if (originalSet.Contains(target))
                throw new InvalidOperationException(
                    $"copy_sheet snapshot refused because target sheet '{target}' already exists");
            if (!targetSet.Add(target))
                throw new InvalidOperationException(
                    $"copy_sheet snapshot refused because target sheet '{target}' appears more than once");
            targetSheets.Add(target);
        }

        return new JsonObject
        {
            ["snapshotVersion"] = CurrentExcelSnapshotVersion,
            ["restoreMode"] = CopySheetTopologyRestoreMode,
            ["documentRef"] = documentRef,
            ["originalSheets"] = StringsToJsonArray(originalSheets),
            ["originalActiveSheet"] = originalActiveSheet,
            ["targetSheets"] = StringsToJsonArray(targetSheets),
            ["ops"] = capturedOps,
        };
    }

    private static List<string>? ReadRequiredStringArray(JsonObject state, string propertyName)
    {
        if (state[propertyName] is not JsonArray values) return null;
        var result = new List<string>(values.Count);
        foreach (var value in values)
        {
            if (value is not JsonValue item || !item.TryGetValue<string>(out var text) ||
                string.IsNullOrWhiteSpace(text))
                return null;
            result.Add(text);
        }
        return result;
    }

    private static bool TryDeleteWorksheet(object application, object workbook, string sheetName,
        RestoreMismatchCollector mismatches)
    {
        object? worksheets = null;
        object? sheet = null;
        try
        {
            dynamic wb = workbook;
            worksheets = (object)wb.Worksheets;
            try { sheet = (object)((dynamic)worksheets).Item(sheetName); }
            catch { return false; }

            var alerts = true;
            try { alerts = Convert.ToBoolean(((dynamic)application).DisplayAlerts, CultureInfo.InvariantCulture); }
            catch { }
            try
            {
                ((dynamic)application).DisplayAlerts = false;
                ((dynamic)sheet).Delete();
                return true;
            }
            catch (Exception ex)
            {
                mismatches.Add($"target sheet '{sheetName}' could not be deleted: {ex.Message}");
                return false;
            }
            finally { try { ((dynamic)application).DisplayAlerts = alerts; } catch { } }
        }
        finally
        {
            RotHelper.ReleaseComReference(sheet);
            RotHelper.ReleaseComReference(worksheets);
        }
    }

    private static void ActivateWorksheet(object workbook, string sheetName, RestoreMismatchCollector mismatches)
    {
        object? worksheets = null;
        object? sheet = null;
        try
        {
            worksheets = (object)((dynamic)workbook).Worksheets;
            sheet = (object)((dynamic)worksheets).Item(sheetName);
            ((dynamic)sheet).Activate();
        }
        catch (Exception ex)
        {
            mismatches.Add($"original active sheet '{sheetName}' could not be activated: {ex.Message}");
        }
        finally
        {
            RotHelper.ReleaseComReference(sheet);
            RotHelper.ReleaseComReference(worksheets);
        }
    }

    private static JsonObject RestoreCopySheetTopology(object application, object workbook, JsonObject state)
    {
        var mismatches = new RestoreMismatchCollector();
        var originalSheets = ReadRequiredStringArray(state, "originalSheets");
        var targetSheets = ReadRequiredStringArray(state, "targetSheets");
        var originalActiveSheet = Json.GetString(state, "originalActiveSheet");
        if (originalSheets is null || targetSheets is null || targetSheets.Count == 0 ||
            string.IsNullOrWhiteSpace(originalActiveSheet))
        {
            mismatches.Add(
                "copy-sheet topology snapshot is missing valid originalSheets, originalActiveSheet, or targetSheets");
            return BuildRestoreResult(false, 0, 0, CopySheetTopologyRestoreMode, mismatches);
        }

        var originalSet = originalSheets.ToHashSet(StringComparer.OrdinalIgnoreCase);
        var distinctTargets = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var target in targetSheets)
        {
            if (originalSet.Contains(target))
                mismatches.Add($"unsafe topology snapshot: target sheet '{target}' was part of the original workbook");
            if (!distinctTargets.Add(target))
                mismatches.Add($"invalid topology snapshot: duplicate target sheet '{target}'");
        }

        if (mismatches.Count == 0)
        {
            for (var index = targetSheets.Count - 1; index >= 0; index--)
                _ = TryDeleteWorksheet(application, workbook, targetSheets[index], mismatches);
        }

        ActivateWorksheet(workbook, originalActiveSheet, mismatches);

        var actualSheets = ReadOrderedWorksheetNames(workbook);
        var actualActiveSheet = ReadActiveWorksheetName(workbook);
        if (actualSheets.Count != originalSheets.Count)
            mismatches.Add($"worksheet count mismatch: expected {originalSheets.Count}, actual {actualSheets.Count}");

        var compared = Math.Max(originalSheets.Count, actualSheets.Count);
        for (var index = 0; index < compared; index++)
        {
            var expected = index < originalSheets.Count ? originalSheets[index] : "<none>";
            var actual = index < actualSheets.Count ? actualSheets[index] : "<missing>";
            if (!string.Equals(expected, actual, StringComparison.OrdinalIgnoreCase))
                mismatches.Add($"worksheet order mismatch at {index + 1}: expected '{expected}', actual '{actual}'");
        }

        foreach (var target in targetSheets)
            if (actualSheets.Contains(target, StringComparer.OrdinalIgnoreCase))
                mismatches.Add($"target sheet '{target}' still exists after restore");
        if (!string.Equals(originalActiveSheet, actualActiveSheet, StringComparison.OrdinalIgnoreCase))
            mismatches.Add(
                $"active worksheet mismatch: expected '{originalActiveSheet}', actual '{actualActiveSheet}'");

        return BuildRestoreResult(
            mismatches.Count == 0,
            restoredCells: 0,
            checkedItems: compared + targetSheets.Count + 1,
            CopySheetTopologyRestoreMode,
            mismatches);
    }

    private static JsonObject BuildRestoreResult(bool verified, int restoredCells, int checkedItems,
        string restoreMode, RestoreMismatchCollector mismatches)
    {
        var errors = Json.ToArray(mismatches.Samples);
        return new JsonObject
        {
            ["ok"] = verified,
            ["restored"] = verified,
            ["restoreMode"] = restoreMode,
            ["restoredCells"] = restoredCells,
            ["readback"] = new JsonObject
            {
                ["verified"] = verified,
                ["checked"] = checkedItems,
                ["totalMismatchCount"] = mismatches.Count,
                ["mismatchSampleCount"] = mismatches.Samples.Count,
                ["mismatchesTruncated"] = mismatches.Truncated,
                ["mismatches"] = errors.DeepClone(),
            },
            ["errors"] = errors,
        };
    }

    public override void CaptureSnapshot(string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject>? ops = null)
    {
        ComInvoke(() =>
        {
            var app = AttachExcel();
            if (app is null) { metadata["payload"] = "none (excel not running)"; return; }
            dynamic d = app;
            dynamic? defaultWorkbook = d.ActiveWorkbook;
            if (defaultWorkbook is null) { metadata["payload"] = "none (no workbook)"; return; }
            using var workbookLease = ResolveTargetWorkbook(d, defaultWorkbook, ops);
            dynamic wb = workbookLease.Workbook;

            // 1) workbook 파일 복사 (저장된 경우) — Excel이 잠그고 있으므로 공유 읽기로 복사
            string? fullName = null;
            try { fullName = (string)wb.FullName; } catch { }
            if (!string.IsNullOrEmpty(fullName) && File.Exists(fullName))
            {
                var dest = Path.Combine(snapshotDir, "workbook-backup" + Path.GetExtension(fullName));
                try
                {
                    using var src = new FileStream(fullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                    using var dst = new FileStream(dest, FileMode.Create, FileAccess.Write);
                    src.CopyTo(dst);
                    metadata["workbookBackup"] = Path.GetFileName(dest);
                }
                catch (Exception ex) { metadata["workbookBackupError"] = ex.Message; }
            }

            // copy_sheet-only batches have a complete, operation-scoped inverse: remove the
            // newly created target worksheets in reverse order. Capturing or rewriting every
            // pre-existing Formula cell is both unnecessary and unsafe for external, shared,
            // array, and Formula2 expressions that Excel may canonicalize on assignment.
            if (IsCopySheetOnlySnapshot(ops))
            {
                var topologyState = CaptureCopySheetTopologyState(
                    (object)wb,
                    ops ?? throw new InvalidOperationException("copy_sheet snapshot operations are missing"),
                    fullName);
                File.WriteAllText(Path.Combine(snapshotDir, "state.json"), topologyState.ToJsonString(Json.Pretty));
                metadata["payload"] = "workbook-backup + structural state.json";
                metadata["documentRef"] = fullName;
                return;
            }

            if (IsVisibilityOnlySnapshot(ops))
            {
                var visibilityState = CaptureVisibilityState(
                    (object)wb,
                    ops ?? throw new InvalidOperationException("visibility snapshot operations are missing"),
                    fullName);
                File.WriteAllText(Path.Combine(snapshotDir, "state.json"), visibilityState.ToJsonString(Json.Pretty));
                metadata["payload"] = "workbook-backup + visibility state.json";
                metadata["documentRef"] = fullName;
                return;
            }

            if (IsMergeOnlySnapshot(ops))
            {
                var mergeState = CaptureMergeState(
                    (object)wb,
                    ops![0],
                    fullName);
                File.WriteAllText(Path.Combine(snapshotDir, "state.json"), mergeState.ToJsonString(Json.Pretty));
                metadata["payload"] = "workbook-backup + merge state.json";
                metadata["documentRef"] = fullName;
                return;
            }

            // 2) 시트 값 state.json (범위 상한 적용)
            var sheets = new JsonObject();
            foreach (dynamic s in wb.Worksheets)
            {
                dynamic used = s.UsedRange;
                long totalCells;
                try { totalCells = Convert.ToInt64(used.CountLarge, CultureInfo.InvariantCulture); }
                catch { totalCells = Convert.ToInt64(used.Count, CultureInfo.InvariantCulture); }
                var values = RangeToJson((object)used, out var cells, maxCells: MaxSnapshotCells);
                var formulas = RangeToJson((object)used, out var formulaCells, formulas: true, maxCells: MaxSnapshotCells);
                sheets[(string)s.Name] = new JsonObject
                {
                    ["address"] = (string)used.Address(false, false),
                    ["values"] = values,
                    ["formulas"] = formulas,
                    ["truncated"] = totalCells > cells || totalCells > formulaCells,
                };
                if (totalCells > cells || totalCells > formulaCells)
                    throw new InvalidOperationException(
                        $"sheet '{(string)s.Name}' snapshot exceeds {MaxSnapshotCells} cells; write was blocked because a complete automatic restore cannot be guaranteed");
            }
            var capturedOps = new JsonArray();
            var formatStates = new JsonArray();
            var formatCellCount = 0L;
            foreach (var op in ops ?? Array.Empty<JsonObject>())
            {
                capturedOps.Add(op.DeepClone());
                if (Json.GetString(op, "op") != "format_range") continue;

                var resolvedRange = ResolveRangeTarget(
                    (object)wb,
                    Json.GetString(Json.GetObj(op, "target"), "sheet"),
                    Json.GetString(op, "range")!,
                    requireExplicitSheet: true);
                dynamic sheet = resolvedRange.Sheet;
                var rangeAddr = resolvedRange.Address;
                dynamic range = sheet.Range(rangeAddr);
                var rows = Convert.ToInt32(range.Rows.Count, CultureInfo.InvariantCulture);
                var cols = Convert.ToInt32(range.Columns.Count, CultureInfo.InvariantCulture);
                formatCellCount += (long)rows * cols;
                if (formatCellCount > MaxFormatSnapshotCells)
                    throw new InvalidOperationException(
                        $"format snapshot exceeds {MaxFormatSnapshotCells} cells; write was blocked because cell formatting could not be restored safely");

                var cellStyles = new JsonArray();
                for (var row = 1; row <= rows; row++)
                {
                    var styleRow = new JsonArray();
                    for (var col = 1; col <= cols; col++)
                    {
                        dynamic cell = range.Cells.Item(row, col);
                        styleRow.Add(CaptureCellStyle(cell));
                    }
                    cellStyles.Add(styleRow);
                }
                formatStates.Add(new JsonObject
                {
                    ["sheet"] = (string)sheet.Name,
                    ["range"] = rangeAddr,
                    ["styles"] = cellStyles,
                });
            }

            File.WriteAllText(Path.Combine(snapshotDir, "state.json"),
                new JsonObject
                {
                    ["snapshotVersion"] = CurrentExcelSnapshotVersion,
                    ["restoreMode"] = LegacyFullRangeRestoreMode,
                    ["sheets"] = sheets,
                    ["ops"] = capturedOps,
                    ["formatStates"] = formatStates,
                }.ToJsonString(Json.Pretty));
            metadata["payload"] = "workbook-backup + state.json";
            metadata["documentRef"] = fullName;
        });
    }

    public override JsonObject RestoreSnapshot(string snapshotDir, JsonObject metadata)
    {
        return ComInvoke(() =>
        {
            var statePath = Path.Combine(snapshotDir, "state.json");
            if (!File.Exists(statePath))
                return Json.ErrorResult("state.json not found in snapshot", App);

            var app = AttachExcel();
            if (app is null) return Json.ErrorResult("Excel not running", App);
            dynamic d = app;
            dynamic wb;
            try { wb = RequireWorkbook(d); }
            catch (Exception ex) { return Json.ErrorResult(ex.Message, App); }

            var docRef = Json.GetString(metadata, "documentRef");
            if (!string.IsNullOrEmpty(docRef) &&
                !string.Equals((string)wb.FullName, docRef, StringComparison.OrdinalIgnoreCase))
                return Json.ErrorResult(
                    $"active workbook '{(string)wb.FullName}' does not match snapshot document '{docRef}'. " +
                    "스냅샷 시점의 workbook을 먼저 여세요.", App);

            var state = JsonNode.Parse(File.ReadAllText(statePath)) as JsonObject ?? new JsonObject();
            var stateDocumentRef = Json.GetString(state, "documentRef");
            if (!string.IsNullOrEmpty(stateDocumentRef) &&
                !string.Equals((string)wb.FullName, stateDocumentRef, StringComparison.OrdinalIgnoreCase))
                return Json.ErrorResult(
                    $"active workbook '{(string)wb.FullName}' does not match snapshot state document '{stateDocumentRef}'. " +
                    "스냅샷 시점의 workbook을 먼저 여세요.", App);
            var restoreMode = Json.GetString(state, "restoreMode");
            if (string.Equals(restoreMode, CopySheetTopologyRestoreMode, StringComparison.Ordinal))
            {
                if (Json.GetInt(state, "snapshotVersion") != CurrentExcelSnapshotVersion)
                {
                    var invalid = new RestoreMismatchCollector();
                    invalid.Add("unsupported copy-sheet topology snapshot version");
                    return BuildRestoreResult(false, 0, 0, CopySheetTopologyRestoreMode, invalid);
                }
                return RestoreCopySheetTopology((object)d, (object)wb, state);
            }
            if (string.Equals(restoreMode, VisibilityRestoreMode, StringComparison.Ordinal))
            {
                if (Json.GetInt(state, "snapshotVersion") != ExcelLayoutSnapshotVersion)
                    return Json.ErrorResult("unsupported Excel visibility snapshot version", App);
                return RestoreVisibilityState((object)wb, state);
            }
            if (string.Equals(restoreMode, MergeRestoreMode, StringComparison.Ordinal))
            {
                if (Json.GetInt(state, "snapshotVersion") != ExcelLayoutSnapshotVersion)
                    return Json.ErrorResult("unsupported Excel merge snapshot version", App);
                return RestoreMergeState((object)wb, state);
            }

            // Versionless snapshots from 0.4.14 and earlier retain their original full-range
            // restore behavior. Do not attempt formula-string normalization here: preserving
            // legacy compatibility is safer than guessing at Formula/Formula2 equivalence.
            var sheets = Json.GetObj(state, "sheets") ?? new JsonObject();
            var mismatches = new RestoreMismatchCollector();
            var checkedCells = 0;
            var restoredCells = 0;

            d.ScreenUpdating = false;
            try
            {
                // 삽입 op는 역순 삭제해야 원래 셀 좌표와 UsedRange가 돌아온다.
                var capturedOps = Json.GetArr(state, "ops") ?? new JsonArray();
                for (var index = capturedOps.Count - 1; index >= 0; index--)
                {
                    if (capturedOps[index] is not JsonObject op) continue;
                    var name = Json.GetString(op, "op");
                    if (name == "insert_rows")
                    {
                        dynamic sheet = GetRequiredTargetSheet(wb, op);
                        var row = Json.GetInt(op, "row")!.Value;
                        var count = Json.GetInt(op, "count")!.Value;
                        sheet.Rows[$"{row}:{row + count - 1}"].Delete();
                    }
                    else if (name == "insert_cols")
                    {
                        dynamic sheet = GetRequiredTargetSheet(wb, op);
                        var count = Json.GetInt(op, "count")!.Value;
                        var colNode = op["col"]!;
                        var colName = colNode is JsonValue jv && jv.TryGetValue<int>(out var ci)
                            ? ColName(ci)
                            : colNode.GetValue<string>();
                        var endCol = ColName(ColIndex(colName) + count - 1);
                        sheet.Columns[$"{colName}:{endCol}"].Delete();
                    }
                    else if (name == "copy_sheet")
                    {
                        var copiedSheetName = Json.GetString(op, "targetSheet") ?? Json.GetString(op, "sourceSheet");
                        if (!string.IsNullOrEmpty(copiedSheetName) && SheetExists(wb, copiedSheetName))
                        {
                            var alerts = true;
                            try { alerts = Convert.ToBoolean(d.DisplayAlerts, CultureInfo.InvariantCulture); } catch { }
                            try
                            {
                                d.DisplayAlerts = false;
                                dynamic copiedSheet = wb.Worksheets.Item(copiedSheetName);
                                copiedSheet.Delete();
                            }
                            finally { try { d.DisplayAlerts = alerts; } catch { } }
                        }
                    }
                }

                foreach (var (sheetName, sheetNode) in sheets)
                {
                    if (sheetNode is not JsonObject so) continue;
                    dynamic sheet;
                    try { sheet = GetSheet(wb, sheetName); }
                    catch { mismatches.Add($"sheet '{sheetName}' missing in current workbook"); continue; }
                    var address = Json.GetString(so, "address");
                    var values = Json.GetArr(so, "values");
                    var formulas = Json.GetArr(so, "formulas");
                    if (Json.GetBool(so, "truncated"))
                        return Json.ErrorResult($"snapshot for sheet '{sheetName}' is truncated; automatic restore refused", App);
                    if (address is null || values is null) continue;
                    dynamic range = sheet.Range(address);

                    var source = formulas ?? values;
                    var rows = source.Count;
                    var cols = rows > 0 && source[0] is JsonArray j0 ? j0.Count : 0;
                    if (rows == 0 || cols == 0) continue;
                    var data = new object?[rows, cols];
                    for (var i = 0; i < rows; i++)
                        for (var j = 0; j < cols; j++)
                            data[i, j] = NodeToComValue(source[i]![j]);
                    if (formulas is not null) range.Formula = data;
                    else range.Value2 = data;
                    restoredCells += rows * cols;

                    // readback
                    object? raw = formulas is not null ? range.Formula : range.Value2;
                    for (var i = 0; i < rows; i++)
                        for (var j = 0; j < cols; j++)
                        {
                            checkedCells++;
                            object? got = raw is object[,] arr ? arr[i + 1, j + 1] : raw;
                            if (!ComValuesEqual(data[i, j], got))
                                mismatches.Add($"{sheetName}!{CellName(range.Column + j, range.Row + i)}: restore mismatch");
                        }
                }

                // 값/수식 복원 뒤, format_range가 바꾼 셀별 서식을 원상 복구한다.
                foreach (var formatNode in Json.GetArr(state, "formatStates") ?? new JsonArray())
                {
                    if (formatNode is not JsonObject formatState) continue;
                    var sheetName = Json.GetString(formatState, "sheet")!;
                    var address = Json.GetString(formatState, "range")!;
                    var styleRows = Json.GetArr(formatState, "styles") ?? new JsonArray();
                    dynamic sheet = GetSheet(wb, sheetName);
                    dynamic range = sheet.Range(address);
                    for (var row = 0; row < styleRows.Count; row++)
                    {
                        if (styleRows[row] is not JsonArray styleCols) continue;
                        for (var col = 0; col < styleCols.Count; col++)
                        {
                            if (styleCols[col] is not JsonObject style) continue;
                            dynamic cell = range.Cells.Item(row + 1, col + 1);
                            RestoreCellStyle(cell, style);
                            checkedCells++;
                            if (!CellStyleMatches(cell, style))
                                mismatches.Add($"{sheetName}!{CellName(range.Column + col, range.Row + row)}: style restore mismatch");
                        }
                    }
                }
            }
            finally { d.ScreenUpdating = true; }

            return BuildRestoreResult(
                mismatches.Count == 0,
                restoredCells,
                checkedCells,
                string.IsNullOrWhiteSpace(restoreMode) ? LegacyFullRangeRestoreMode : restoreMode,
                mismatches);
        });
    }

    private static readonly int[] ExcelCellBorderIndexes = { 5, 6, 7, 8, 9, 10, 11, 12 };

    private static JsonArray CaptureCellBorders(object cell)
    {
        var result = new JsonArray();
        object? borders = null;
        try
        {
            borders = (object)((dynamic)cell).Borders;
            foreach (var index in ExcelCellBorderIndexes)
            {
                object? border = null;
                try
                {
                    border = (object)((dynamic)borders).Item(index);
                    result.Add(new JsonObject
                    {
                        ["index"] = index,
                        ["lineStyle"] = Convert.ToInt32(((dynamic)border).LineStyle, CultureInfo.InvariantCulture),
                        ["weight"] = Convert.ToInt32(((dynamic)border).Weight, CultureInfo.InvariantCulture),
                        ["color"] = Convert.ToDouble(((dynamic)border).Color, CultureInfo.InvariantCulture),
                    });
                }
                catch
                {
                    // Some Excel builds do not expose inside borders for a one-cell Range.
                    // Edge/diagonal borders that are exposed are still captured exactly.
                }
                finally { RotHelper.ReleaseComReference(border); }
            }
        }
        finally { RotHelper.ReleaseComReference(borders); }
        return result;
    }

    private static void RestoreCellBorders(object cell, JsonArray states)
    {
        object? borders = null;
        try
        {
            borders = (object)((dynamic)cell).Borders;
            foreach (var node in states)
            {
                if (node is not JsonObject state || Json.GetInt(state, "index") is not int index) continue;
                object? border = null;
                try
                {
                    border = (object)((dynamic)borders).Item(index);
                    // Set LineStyle last because assigning color/weight to an empty border
                    // can temporarily create a visible line in Excel.
                    ((dynamic)border).Color = state["color"]!.GetValue<double>();
                    ((dynamic)border).Weight = state["weight"]!.GetValue<int>();
                    ((dynamic)border).LineStyle = state["lineStyle"]!.GetValue<int>();
                }
                finally { RotHelper.ReleaseComReference(border); }
            }
        }
        finally { RotHelper.ReleaseComReference(borders); }
    }

    private static bool CellBordersMatch(object cell, JsonArray states)
    {
        object? borders = null;
        try
        {
            borders = (object)((dynamic)cell).Borders;
            foreach (var node in states)
            {
                if (node is not JsonObject state || Json.GetInt(state, "index") is not int index) continue;
                object? border = null;
                try
                {
                    border = (object)((dynamic)borders).Item(index);
                    if (Convert.ToInt32(((dynamic)border).LineStyle, CultureInfo.InvariantCulture) !=
                            state["lineStyle"]!.GetValue<int>() ||
                        Convert.ToInt32(((dynamic)border).Weight, CultureInfo.InvariantCulture) !=
                            state["weight"]!.GetValue<int>() ||
                        Math.Abs(Convert.ToDouble(((dynamic)border).Color, CultureInfo.InvariantCulture) -
                                 state["color"]!.GetValue<double>()) > 1e-9)
                        return false;
                }
                finally { RotHelper.ReleaseComReference(border); }
            }
            return true;
        }
        finally { RotHelper.ReleaseComReference(borders); }
    }

    private static JsonObject CaptureMergeCellStyle(object cell)
    {
        object? font = null;
        object? interior = null;
        try
        {
            dynamic dynamicCell = cell;
            font = (object)dynamicCell.Font;
            interior = (object)dynamicCell.Interior;
            dynamic dynamicFont = font;
            dynamic dynamicInterior = interior;
            return new JsonObject
            {
                ["bold"] = Convert.ToBoolean(dynamicFont.Bold, CultureInfo.InvariantCulture),
                ["italic"] = Convert.ToBoolean(dynamicFont.Italic, CultureInfo.InvariantCulture),
                ["fontName"] = Convert.ToString(dynamicFont.Name, CultureInfo.InvariantCulture),
                ["fontSize"] = Convert.ToDouble(dynamicFont.Size, CultureInfo.InvariantCulture),
                ["underline"] = Convert.ToInt32(dynamicFont.Underline, CultureInfo.InvariantCulture),
                ["strikethrough"] = Convert.ToBoolean(dynamicFont.Strikethrough, CultureInfo.InvariantCulture),
                ["numberFormat"] = Convert.ToString(dynamicCell.NumberFormat, CultureInfo.InvariantCulture),
                ["fontColor"] = Convert.ToDouble(dynamicFont.Color, CultureInfo.InvariantCulture),
                ["fillColor"] = Convert.ToDouble(dynamicInterior.Color, CultureInfo.InvariantCulture),
                ["fillPattern"] = Convert.ToInt32(dynamicInterior.Pattern, CultureInfo.InvariantCulture),
                ["fillPatternColor"] = Convert.ToDouble(dynamicInterior.PatternColor, CultureInfo.InvariantCulture),
                ["horizontalAlignment"] = Convert.ToInt32(dynamicCell.HorizontalAlignment, CultureInfo.InvariantCulture),
                ["verticalAlignment"] = Convert.ToInt32(dynamicCell.VerticalAlignment, CultureInfo.InvariantCulture),
                ["wrapText"] = Convert.ToBoolean(dynamicCell.WrapText, CultureInfo.InvariantCulture),
                ["shrinkToFit"] = Convert.ToBoolean(dynamicCell.ShrinkToFit, CultureInfo.InvariantCulture),
                ["indentLevel"] = Convert.ToInt32(dynamicCell.IndentLevel, CultureInfo.InvariantCulture),
                ["orientation"] = Convert.ToInt32(dynamicCell.Orientation, CultureInfo.InvariantCulture),
                ["locked"] = Convert.ToBoolean(dynamicCell.Locked, CultureInfo.InvariantCulture),
                ["formulaHidden"] = Convert.ToBoolean(dynamicCell.FormulaHidden, CultureInfo.InvariantCulture),
                ["borders"] = CaptureCellBorders(cell),
            };
        }
        finally
        {
            RotHelper.ReleaseComReference(interior);
            RotHelper.ReleaseComReference(font);
        }
    }

    private static void RestoreMergeCellStyle(object cell, JsonObject style)
    {
        object? font = null;
        object? interior = null;
        try
        {
            dynamic dynamicCell = cell;
            font = (object)dynamicCell.Font;
            interior = (object)dynamicCell.Interior;
            dynamic dynamicFont = font;
            dynamic dynamicInterior = interior;
            dynamicFont.Bold = Json.GetBool(style, "bold");
            dynamicFont.Italic = Json.GetBool(style, "italic");
            dynamicFont.Name = Json.GetString(style, "fontName");
            dynamicFont.Size = style["fontSize"]!.GetValue<double>();
            dynamicFont.Underline = style["underline"]!.GetValue<int>();
            dynamicFont.Strikethrough = Json.GetBool(style, "strikethrough");
            dynamicCell.NumberFormat = Json.GetString(style, "numberFormat") ?? "General";
            dynamicFont.Color = style["fontColor"]!.GetValue<double>();
            dynamicInterior.PatternColor = style["fillPatternColor"]!.GetValue<double>();
            dynamicInterior.Color = style["fillColor"]!.GetValue<double>();
            // Color/PatternColor assignments can implicitly switch Pattern to solid.
            // Apply the captured pattern last so "no fill" is restored exactly.
            dynamicInterior.Pattern = style["fillPattern"]!.GetValue<int>();
            dynamicCell.HorizontalAlignment = style["horizontalAlignment"]!.GetValue<int>();
            dynamicCell.VerticalAlignment = style["verticalAlignment"]!.GetValue<int>();
            dynamicCell.WrapText = Json.GetBool(style, "wrapText");
            dynamicCell.ShrinkToFit = Json.GetBool(style, "shrinkToFit");
            dynamicCell.IndentLevel = style["indentLevel"]!.GetValue<int>();
            dynamicCell.Orientation = style["orientation"]!.GetValue<int>();
            dynamicCell.Locked = Json.GetBool(style, "locked");
            dynamicCell.FormulaHidden = Json.GetBool(style, "formulaHidden");
            RestoreCellBorders(cell, Json.GetArr(style, "borders") ?? new JsonArray());
        }
        finally
        {
            RotHelper.ReleaseComReference(interior);
            RotHelper.ReleaseComReference(font);
        }
    }

    private static bool MergeCellStyleMatches(object cell, JsonObject style)
    {
        object? font = null;
        object? interior = null;
        try
        {
            dynamic dynamicCell = cell;
            font = (object)dynamicCell.Font;
            interior = (object)dynamicCell.Interior;
            dynamic dynamicFont = font;
            dynamic dynamicInterior = interior;
            return Convert.ToBoolean(dynamicFont.Bold, CultureInfo.InvariantCulture) == Json.GetBool(style, "bold")
                   && Convert.ToBoolean(dynamicFont.Italic, CultureInfo.InvariantCulture) == Json.GetBool(style, "italic")
                   && string.Equals(Convert.ToString(dynamicFont.Name, CultureInfo.InvariantCulture), Json.GetString(style, "fontName"), StringComparison.Ordinal)
                   && Math.Abs(Convert.ToDouble(dynamicFont.Size, CultureInfo.InvariantCulture) - style["fontSize"]!.GetValue<double>()) < 1e-9
                   && Convert.ToInt32(dynamicFont.Underline, CultureInfo.InvariantCulture) == style["underline"]!.GetValue<int>()
                   && Convert.ToBoolean(dynamicFont.Strikethrough, CultureInfo.InvariantCulture) == Json.GetBool(style, "strikethrough")
                   && string.Equals(Convert.ToString(dynamicCell.NumberFormat, CultureInfo.InvariantCulture), Json.GetString(style, "numberFormat"), StringComparison.Ordinal)
                   && Math.Abs(Convert.ToDouble(dynamicFont.Color, CultureInfo.InvariantCulture) - style["fontColor"]!.GetValue<double>()) < 1e-9
                   && Math.Abs(Convert.ToDouble(dynamicInterior.Color, CultureInfo.InvariantCulture) - style["fillColor"]!.GetValue<double>()) < 1e-9
                   && Convert.ToInt32(dynamicInterior.Pattern, CultureInfo.InvariantCulture) == style["fillPattern"]!.GetValue<int>()
                   && Math.Abs(Convert.ToDouble(dynamicInterior.PatternColor, CultureInfo.InvariantCulture) - style["fillPatternColor"]!.GetValue<double>()) < 1e-9
                   && Convert.ToInt32(dynamicCell.HorizontalAlignment, CultureInfo.InvariantCulture) == style["horizontalAlignment"]!.GetValue<int>()
                   && Convert.ToInt32(dynamicCell.VerticalAlignment, CultureInfo.InvariantCulture) == style["verticalAlignment"]!.GetValue<int>()
                   && Convert.ToBoolean(dynamicCell.WrapText, CultureInfo.InvariantCulture) == Json.GetBool(style, "wrapText")
                   && Convert.ToBoolean(dynamicCell.ShrinkToFit, CultureInfo.InvariantCulture) == Json.GetBool(style, "shrinkToFit")
                   && Convert.ToInt32(dynamicCell.IndentLevel, CultureInfo.InvariantCulture) == style["indentLevel"]!.GetValue<int>()
                   && Convert.ToInt32(dynamicCell.Orientation, CultureInfo.InvariantCulture) == style["orientation"]!.GetValue<int>()
                   && Convert.ToBoolean(dynamicCell.Locked, CultureInfo.InvariantCulture) == Json.GetBool(style, "locked")
                   && Convert.ToBoolean(dynamicCell.FormulaHidden, CultureInfo.InvariantCulture) == Json.GetBool(style, "formulaHidden")
                   && CellBordersMatch(cell, Json.GetArr(style, "borders") ?? new JsonArray());
        }
        finally
        {
            RotHelper.ReleaseComReference(interior);
            RotHelper.ReleaseComReference(font);
        }
    }

    // Keep the legacy format_range snapshot compact and backward-compatible.
    // Merge snapshots use the broader helpers above because Excel can normalize
    // many more properties when cells are merged.
    private static JsonObject CaptureCellStyle(object cell)
    {
        object? font = null;
        object? interior = null;
        try
        {
            dynamic dynamicCell = cell;
            font = (object)dynamicCell.Font;
            interior = (object)dynamicCell.Interior;
            dynamic dynamicFont = font;
            dynamic dynamicInterior = interior;
            return new JsonObject
            {
                ["bold"] = Convert.ToBoolean(dynamicFont.Bold, CultureInfo.InvariantCulture),
                ["italic"] = Convert.ToBoolean(dynamicFont.Italic, CultureInfo.InvariantCulture),
                ["fontSize"] = Convert.ToDouble(dynamicFont.Size, CultureInfo.InvariantCulture),
                ["numberFormat"] = Convert.ToString(dynamicCell.NumberFormat, CultureInfo.InvariantCulture),
                ["fontColor"] = Convert.ToDouble(dynamicFont.Color, CultureInfo.InvariantCulture),
                ["fillColor"] = Convert.ToDouble(dynamicInterior.Color, CultureInfo.InvariantCulture),
            };
        }
        finally
        {
            RotHelper.ReleaseComReference(interior);
            RotHelper.ReleaseComReference(font);
        }
    }

    private static void RestoreCellStyle(object cell, JsonObject style)
    {
        object? font = null;
        object? interior = null;
        try
        {
            dynamic dynamicCell = cell;
            font = (object)dynamicCell.Font;
            interior = (object)dynamicCell.Interior;
            dynamic dynamicFont = font;
            dynamic dynamicInterior = interior;
            dynamicFont.Bold = Json.GetBool(style, "bold");
            dynamicFont.Italic = Json.GetBool(style, "italic");
            dynamicFont.Size = style["fontSize"]!.GetValue<double>();
            dynamicCell.NumberFormat = Json.GetString(style, "numberFormat") ?? "General";
            dynamicFont.Color = style["fontColor"]!.GetValue<double>();
            dynamicInterior.Color = style["fillColor"]!.GetValue<double>();
        }
        finally
        {
            RotHelper.ReleaseComReference(interior);
            RotHelper.ReleaseComReference(font);
        }
    }

    private static bool CellStyleMatches(object cell, JsonObject style)
    {
        object? font = null;
        object? interior = null;
        try
        {
            dynamic dynamicCell = cell;
            font = (object)dynamicCell.Font;
            interior = (object)dynamicCell.Interior;
            dynamic dynamicFont = font;
            dynamic dynamicInterior = interior;
            return Convert.ToBoolean(dynamicFont.Bold, CultureInfo.InvariantCulture) == Json.GetBool(style, "bold")
                   && Convert.ToBoolean(dynamicFont.Italic, CultureInfo.InvariantCulture) == Json.GetBool(style, "italic")
                   && Math.Abs(Convert.ToDouble(dynamicFont.Size, CultureInfo.InvariantCulture) - style["fontSize"]!.GetValue<double>()) < 1e-9
                   && string.Equals(Convert.ToString(dynamicCell.NumberFormat, CultureInfo.InvariantCulture), Json.GetString(style, "numberFormat"), StringComparison.Ordinal)
                   && Math.Abs(Convert.ToDouble(dynamicFont.Color, CultureInfo.InvariantCulture) - style["fontColor"]!.GetValue<double>()) < 1e-9
                   && Math.Abs(Convert.ToDouble(dynamicInterior.Color, CultureInfo.InvariantCulture) - style["fillColor"]!.GetValue<double>()) < 1e-9;
        }
        finally
        {
            RotHelper.ReleaseComReference(interior);
            RotHelper.ReleaseComReference(font);
        }
    }

    // ---------- range 주소 유틸 ----------

    private const long MaxExactExcelInteger = 9_007_199_254_740_991L;

    /// <summary>
    /// Converts JSON values to the native CLR types expected by Excel Range.Value2.
    /// Excel exposes every numeric cell through COM as Double, so integral JSON
    /// values are normalized to Double instead of leaking Int32/Int64 into a
    /// multi-cell SAFEARRAY. Integers outside IEEE-754's exact range are refused.
    /// </summary>
    internal static object? NodeToComValue(JsonNode? node)
    {
        if (node is null) return null;
        if (node is not JsonValue value) return node.ToJsonString();
        if (value.TryGetValue<string>(out var text)) return text;
        if (value.TryGetValue<bool>(out var boolean)) return boolean;
        if (value.TryGetValue<int>(out var int32)) return Convert.ToDouble(int32, CultureInfo.InvariantCulture);
        if (value.TryGetValue<long>(out var int64))
        {
            if (int64 is > MaxExactExcelInteger or < -MaxExactExcelInteger)
                throw new ArgumentOutOfRangeException(nameof(node), int64,
                    $"Excel cannot store integer {int64} exactly; use text when all digits must be preserved");
            return Convert.ToDouble(int64, CultureInfo.InvariantCulture);
        }
        if (value.TryGetValue<uint>(out var uint32)) return Convert.ToDouble(uint32, CultureInfo.InvariantCulture);
        if (value.TryGetValue<ulong>(out var uint64))
        {
            if (uint64 > (ulong)MaxExactExcelInteger)
                throw new ArgumentOutOfRangeException(nameof(node), uint64,
                    $"Excel cannot store integer {uint64} exactly; use text when all digits must be preserved");
            return Convert.ToDouble(uint64, CultureInfo.InvariantCulture);
        }
        if (value.TryGetValue<decimal>(out var decimalValue))
            return Convert.ToDouble(decimalValue, CultureInfo.InvariantCulture);
        if (value.TryGetValue<float>(out var single)) return Convert.ToDouble(single, CultureInfo.InvariantCulture);
        if (value.TryGetValue<double>(out var number))
        {
            if (double.IsNaN(number) || double.IsInfinity(number))
                throw new ArgumentOutOfRangeException(nameof(node), number, "Excel Value2 requires a finite number");
            return number;
        }
        return node.ToJsonString();
    }

    private static bool ComValuesEqual(object? expected, object? actual)
    {
        if (expected is null || actual is null) return expected is null && actual is null;
        if (IsNumber(expected) && IsNumber(actual))
            return Math.Abs(Convert.ToDouble(expected, CultureInfo.InvariantCulture) -
                            Convert.ToDouble(actual, CultureInfo.InvariantCulture)) < 1e-9;
        if (expected is bool eb && actual is bool ab) return eb == ab;
        return string.Equals(ToDisp(expected), ToDisp(actual), StringComparison.Ordinal);
    }

    private static bool IsNumber(object value) => value is byte or sbyte or short or ushort or int or uint or long or ulong or float or double or decimal;

    private static string ColName(int c)
    {
        var name = "";
        while (c > 0) { var m = (c - 1) % 26; name = (char)('A' + m) + name; c = (c - 1) / 26; }
        return name;
    }

    private static int ColIndex(string name)
    {
        var c = 0;
        foreach (var ch in name.ToUpperInvariant()) c = c * 26 + (ch - 'A' + 1);
        return c;
    }

    public override void Dispose()
    {
        if (Interlocked.Exchange(ref _excelDisposed, 1) != 0) return;
        _idleLifecycleTimer.Dispose();
        try { RunOnAdapterThread<object?>(() => { _ = DisconnectExcelCore("adapter-dispose"); return null; }); } catch { }
        base.Dispose();
    }

    public JsonObject Disconnect()
    {
        // Quit is asynchronous in Microsoft 365 and the owned-process wait is up to 10 seconds.
        // Keep the outer STA timeout larger so a successful normal shutdown is not mislabeled as
        // a modal-dialog timeout at the same boundary.
        try { return ComInvoke(() => DisconnectExcelCore("explicit-disconnect"), timeoutSec: 20); }
        catch (Exception ex) { return Json.ErrorResult($"Excel disconnect failed: {ex.Message}", App); }
    }
}
