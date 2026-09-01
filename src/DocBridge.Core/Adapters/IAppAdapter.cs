using System.Text.Json.Nodes;
using DocBridge.Core.Models;

namespace DocBridge.Core.Adapters;

/// <summary>
/// 문서 프로그램 어댑터 공통 인터페이스 (명령서 11.3: 모든 어댑터는 IAppAdapter를 구현).
/// COM 호출이 필요한 어댑터는 내부적으로 StaThreadRunner를 사용한다.
/// </summary>
public interface IAppAdapter : IDisposable
{
    /// <summary>"excel" | "hwp" | "cad" | "fake"</summary>
    string App { get; }

    AdapterStatus GetStatus();

    /// <summary>Machine-readable supported operations, limits, and automation mode.</summary>
    JsonObject GetCapabilities();

    /// <summary>*_get_active_context 의 앱별 구현</summary>
    ContextResult GetActiveContext();

    /// <summary>*_read_* 의 앱별 구현 (읽기는 자유롭게 호출 가능)</summary>
    JsonObject Read(JsonObject args);

    /// <summary>dry-run: 쓰지 않고 affected/diff 계산</summary>
    ApplyPreview Preview(IReadOnlyList<JsonObject> ops);

    /// <summary>실제 적용 + readback 검증 (스냅샷은 호출 전에 생성되어 있어야 함)</summary>
    ApplyExecution Apply(IReadOnlyList<JsonObject> ops, string snapshotId);

    /// <summary>스냅샷 페이로드 캡처 (snapshotDir 안에 파일 작성, metadata 확장 가능)</summary>
    void CaptureSnapshot(string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject>? ops = null);

    /// <summary>스냅샷 복원 (고위험: host가 confirmToken 검증 후 호출)</summary>
    JsonObject RestoreSnapshot(string snapshotDir, JsonObject metadata);
}

/// <summary>
/// dry-run에서 만든 preview를 실제 적용 때 안전하게 재사용할 수 있는 어댑터의 선택 계약.
/// 구현체는 스냅샷 당시의 전체 문서 fingerprint와 현재 상태를 비교해야 하며,
/// 경로/이름만 같은 경우에는 reusable=true를 반환하면 안 된다.
/// </summary>
public interface IPreviewReuseAdapter
{
    JsonObject ValidatePreviewReuse(
        string snapshotDir,
        JsonObject metadata,
        IReadOnlyList<JsonObject> ops);
}

/// <summary>
/// Long-lived MCP 서버를 종료하지 않고 앱 COM 참조만 명시적으로 놓는 선택 계약.
/// 구현체는 사용자가 소유한 앱을 종료하지 않아야 한다.
/// </summary>
public interface IConnectionLifecycleAdapter
{
    JsonObject Disconnect();
}

/// <summary>
/// 한글 전용 확장 진입점. 직접 COM 어댑터와 외부 worker 프록시가 같은 계약을 구현해
/// 호스트가 실행 방식에 의존하지 않도록 한다.
/// </summary>
public interface IHwpAutomationAdapter
{
    JsonObject Launch(JsonObject args);
    JsonObject Doctor(JsonObject args);
    JsonObject RepairTypeLib(JsonObject args);
}
