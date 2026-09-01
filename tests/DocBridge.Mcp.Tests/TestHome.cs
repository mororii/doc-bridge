using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Mcp.Tests;

/// <summary>테스트 전용 임시 홈 (Core.Tests의 TestHome과 동일 역할)</summary>
public sealed class TestHome : IDisposable
{
    public string Dir { get; }

    public TestHome()
    {
        Dir = Path.Combine(Path.GetTempPath(), "docbridge-mcp-test-" + Guid.NewGuid().ToString("N")[..8]);
        Directory.CreateDirectory(Dir);
    }

    public DocBridgeOptions Options => new(Dir);

    public void Dispose()
    {
        try { Directory.Delete(Dir, recursive: true); } catch { }
    }
}
