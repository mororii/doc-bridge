using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

/// <summary>
/// COM 기반 어댑터 공통 베이스. STA 디스패처 + ROT 연결을 제공한다.
/// M0 단계에서는 자식 클래스가 미구현 stub 상태일 수 있다.
/// </summary>
public abstract class ComAdapterBase : IAppAdapter
{
    protected readonly StaThreadRunner Sta;
    private readonly string _progId;
    private bool _disposed;

    protected ComAdapterBase(string app, string progId)
    {
        App = app;
        _progId = progId;
        Sta = new StaThreadRunner($"doc-bridge-{app}");
    }

    public string App { get; }

    /// <summary>
    /// 실행 중 인스턴스 우선, 없으면 새 인스턴스 생성 시도.
    /// 새 인스턴스가 필요한 파일 기반 HWP에서만 사용한다. Excel/CAD는 상태 조회나 읽기가
    /// 프로그램을 몰래 시작하지 않도록 ROT의 실행 중 인스턴스에만 연결한다.
    /// </summary>
    protected object? Attach()
    {
        var running = RotHelper.GetActiveObject(_progId);
        if (running is not null) return running;
        return RotHelper.CreateInstance(_progId);
    }

    /// <summary>
    /// 테스트/고급 용도: 어댑터 STA 스레드에서 직접 작업 실행.
    /// COM 객체는 생성된 아파트먼트에서만 안전하게 다룰 수 있으므로
    /// 정리(Quit) 같은 작업도 이 경로로 수행한다.
    /// </summary>
    public T RunOnAdapterThread<T>(Func<T> f) => Sta.Invoke(f);

    /// <summary>
    /// COM 호출 표준 래퍼: 타임아웃(기본 120초)을 두어
    /// 모달 대화상자 등으로 인한 영구 블록을 오류로 표면화한다.
    /// (교훈: 한글 AllReplace는 IgnoreMessage 미설정 시 모달 대기로 블록됨)
    /// </summary>
    protected T ComInvoke<T>(Func<T> f, int timeoutSec = 120) =>
        Sta.Invoke(f, TimeSpan.FromSeconds(timeoutSec));

    /// <summary>Action 오버로드 (반환 없는 COM 작업)</summary>
    protected void ComInvoke(Action f, int timeoutSec = 120) =>
        Sta.Invoke<object?>(() => { f(); return null; }, TimeSpan.FromSeconds(timeoutSec));

    /// <summary>AutoCAD 시작/바쁨 상태에서 COM이 호출을 거부하는 경우 판별 (AggregateException 래핑 포함)</summary>
    protected static bool IsCallRejected(Exception ex)
    {
        if (ex is COMException com &&
            (com.HResult == unchecked((int)0x80010001) ||   // RPC_E_CALL_REJECTED
             com.HResult == unchecked((int)0x8001010A)))    // RPC_E_SERVERCALL_RETRYLATER
            return true;
        if (ex is AggregateException agg)
            return agg.InnerExceptions.Any(IsCallRejected);
        return ex.InnerException is not null && IsCallRejected(ex.InnerException);
    }

    /// <summary>
    /// Returns true when a cached COM proxy is no longer connected to its server process.
    /// These errors are different from RPC_E_CALL_REJECTED: retrying the same proxy cannot
    /// recover after the Office application has exited and restarted.
    /// </summary>
    protected static bool IsComDisconnected(Exception ex)
    {
        if (ex is InvalidComObjectException) return true;
        if (ex is COMException com &&
            (com.HResult == unchecked((int)0x80010108) ||   // RPC_E_DISCONNECTED
             com.HResult == unchecked((int)0x800706BA) ||   // RPC_S_SERVER_UNAVAILABLE
             com.HResult == unchecked((int)0x800401FD) ||   // CO_E_OBJNOTCONNECTED
             com.HResult == unchecked((int)0x80010007) ||   // RPC_E_SERVER_DIED
             com.HResult == unchecked((int)0x80010012)))    // RPC_E_SERVER_DIED_DNE
            return true;
        if (ex is AggregateException agg)
            return agg.InnerExceptions.Any(IsComDisconnected);
        return ex.InnerException is not null && IsComDisconnected(ex.InnerException);
    }

    /// <summary>
    /// COM 호출 재시도 래퍼: AutoCAD처럼 시작 중/모달/저장 중 호출을 거부하는
    /// 앱 대응용. 거부 HRESULT에 한해 지연 후 재시도한다 (기본 최대 60초).
    /// </summary>
    protected T ComInvokeWithRetry<T>(Func<T> f, int timeoutSec = 120, int maxAttempts = 60, int delayMs = 1000)
    {
        for (var attempt = 1; ; attempt++)
        {
            try { return ComInvoke(f, timeoutSec); }
            catch (Exception ex) when (IsCallRejected(ex) && attempt < maxAttempts)
            {
                Thread.Sleep(delayMs);
            }
        }
    }

    /// <summary>Action 오버로드 (반환 없는 COM 작업, 재시도 포함)</summary>
    protected void ComInvokeWithRetry(Action f, int timeoutSec = 120, int maxAttempts = 60, int delayMs = 1000) =>
        ComInvokeWithRetry<object?>(() => { f(); return null; }, timeoutSec, maxAttempts, delayMs);

    public abstract AdapterStatus GetStatus();
    public virtual JsonObject GetCapabilities() => new()
    {
        ["app"] = App,
        ["automation"] = "com",
        ["directAppControl"] = true,
        ["readOps"] = new JsonArray(),
        ["writeOps"] = new JsonArray(),
        ["limits"] = new JsonObject(),
        ["interactionPolicy"] = new JsonObject
        {
            ["mode"] = "preserve-foreground",
            ["backgroundInactiveWindow"] = true,
            ["restoresOriginalDocument"] = true,
            ["concurrentTargetInput"] = "stop-after-current-operation",
            ["sameDocumentConcurrentEditing"] = false,
        },
    };
    public abstract ContextResult GetActiveContext();
    public abstract JsonObject Read(JsonObject args);
    public abstract ApplyPreview Preview(IReadOnlyList<JsonObject> ops);
    public abstract ApplyExecution Apply(IReadOnlyList<JsonObject> ops, string snapshotId);
    public abstract void CaptureSnapshot(string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject>? ops = null);
    public abstract JsonObject RestoreSnapshot(string snapshotDir, JsonObject metadata);

    public virtual void Dispose()
    {
        if (_disposed) return;
        _disposed = true;
        Sta.Dispose();
    }
}

/// <summary>어댑터가 아직 구현되지 않았을 때의 stub 응답</summary>
public abstract class NotYetAdapter : IAppAdapter
{
    protected NotYetAdapter(string app, string milestone) { App = app; Milestone = milestone; }
    public string App { get; }
    public string Milestone { get; }

    private string Msg => $"{App} adapter not implemented yet (planned at {Milestone})";

    public AdapterStatus GetStatus() => new(false, false, App, null, null, Msg);
    public JsonObject GetCapabilities() => new()
    {
        ["app"] = App,
        ["automation"] = "unavailable",
        ["directAppControl"] = false,
        ["readOps"] = new JsonArray(),
        ["writeOps"] = new JsonArray(),
        ["limits"] = new JsonObject(),
    };
    public ContextResult GetActiveContext()
    {
        var r = new ContextResult { Ok = false, App = App };
        r.Errors.Add(Msg);
        return r;
    }
    public JsonObject Read(JsonObject args) => Json.ErrorResult(Msg, App);
    public ApplyPreview Preview(IReadOnlyList<JsonObject> ops)
    {
        var p = new ApplyPreview();
        p.Errors.Add(Msg);
        return p;
    }
    public ApplyExecution Apply(IReadOnlyList<JsonObject> ops, string snapshotId)
    {
        var e = new ApplyExecution { Ok = false };
        e.Errors.Add(Msg);
        return e;
    }
    public void CaptureSnapshot(string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject>? ops = null)
    {
        metadata["stub"] = true;
    }
    public JsonObject RestoreSnapshot(string snapshotDir, JsonObject metadata) => Json.ErrorResult(Msg, App);
    public void Dispose() { }
}
