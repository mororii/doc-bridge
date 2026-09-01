using DocBridge.Core.Adapters;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Services;

/// <summary>
/// 앱 이름 → 어댑터 라우팅 (명령서 아키텍처의 SessionRouter).
/// 어댑터는 지연 생성되며 프로세스 내 싱글턴이다.
/// COM 기반 자동화 호출은 host가 named mutex로 크로스 프로세스 직렬화한다.
/// </summary>
public sealed class SessionRouter : IDisposable
{
    private readonly Dictionary<string, Func<IAppAdapter>> _factories;
    private readonly Dictionary<string, IAppAdapter> _instances = new(StringComparer.OrdinalIgnoreCase);
    private readonly object _gate = new();
    private bool _disposed;

    public SessionRouter()
    {
        _factories = new Dictionary<string, Func<IAppAdapter>>(StringComparer.OrdinalIgnoreCase)
        {
            ["fake"] = () => new FakeAdapter(),
            ["excel"] = () => ExcelWorkerAdapter.CanUseCurrentHost
                ? new ExcelWorkerAdapter()
                : new ExcelAdapter(),
            ["hwp"] = () => new HwpWorkerAdapter(),
            ["cad"] = () => new CadAdapter(),
        };
    }

    public IReadOnlyCollection<string> Apps => _factories.Keys;

    public IAppAdapter Get(string app)
    {
        lock (_gate)
        {
            if (_disposed) throw new ObjectDisposedException(nameof(SessionRouter));
            if (_instances.TryGetValue(app, out var existing)) return existing;
            if (!_factories.TryGetValue(app, out var factory))
                throw new ArgumentException($"unknown app '{app}'. supported: {string.Join(", ", _factories.Keys)}");
            var created = factory();
            _instances[app] = created;
            return created;
        }
    }

    public bool TryGet(string app, out IAppAdapter? adapter)
    {
        try { adapter = Get(app); return true; }
        catch { adapter = null; return false; }
    }

    /// <summary>테스트/외부 주입용: 이미 연결된 어댑터 인스턴스를 등록한다</summary>
    public void Register(string app, IAppAdapter adapter)
    {
        lock (_gate)
        {
            if (_instances.TryGetValue(app, out var old) && !ReferenceEquals(old, adapter))
                old.Dispose();
            _instances[app] = adapter;
        }
    }

    public void Dispose()
    {
        lock (_gate)
        {
            if (_disposed) return;
            _disposed = true;
            foreach (var a in _instances.Values) a.Dispose();
            _instances.Clear();
        }
    }
}
