namespace DocBridge.Core.Models;

/// <summary>
/// doc-bridge 런타임 옵션.
/// RootDir 기본값: %LOCALAPPDATA%\DocBridge, DOCBRIDGE_HOME 환경변수로 재정의 가능(테스트용).
/// </summary>
public sealed class DocBridgeOptions
{
    public string RootDir { get; }
    public string? PolicyPath { get; set; }

    public DocBridgeOptions(string? rootDir = null)
    {
        RootDir = rootDir
            ?? Environment.GetEnvironmentVariable("DOCBRIDGE_HOME")
            ?? Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "DocBridge");
    }

    public string SnapshotsDir => Path.Combine(RootDir, "snapshots");
    public string TokensDir => Path.Combine(RootDir, "tokens");
    public string LogsDir => Path.Combine(RootDir, "logs");

    public void EnsureDirectories()
    {
        Directory.CreateDirectory(RootDir);
        Directory.CreateDirectory(SnapshotsDir);
        Directory.CreateDirectory(TokensDir);
        Directory.CreateDirectory(LogsDir);
    }
}
