using System.Text.Json.Nodes;
using DocBridge.Core.Models;

namespace DocBridge.Core.Services;

/// <summary>
/// 스냅샷 서비스 (보안 원칙 4: 쓰기 전에는 반드시 snapshot/backup).
/// 스냅샷 = {RootDir}/snapshots/{app}/{snapshotId}/ 디렉터리
///   - metadata.json : 공통 메타
///   - 그 외 파일    : 어댑터가 캡처한 백업 페이로드 (workbook copy, state.json 등)
/// </summary>
public sealed class SnapshotService
{
    private readonly DocBridgeOptions _options;

    public SnapshotService(DocBridgeOptions options)
    {
        _options = options;
        options.EnsureDirectories();
    }

    public const string MetadataFile = "metadata.json";

    public SnapshotInfo Create(string app, string reason, string? documentRef, Action<string, JsonObject> capture)
    {
        var id = $"{DateTimeOffset.Now:yyyyMMdd-HHmmss}-{Guid.NewGuid().ToString("N")[..8]}";
        var dir = Path.Combine(_options.SnapshotsDir, app, id);
        Directory.CreateDirectory(dir);

        var meta = new JsonObject
        {
            ["snapshotId"] = id,
            ["createdAt"] = DateTimeOffset.Now.ToString("o"),
            ["app"] = app,
            ["documentRef"] = documentRef,
            ["reason"] = reason,
        };

        capture(dir, meta); // 어댑터가 백업 페이로드 작성 + metadata 확장

        File.WriteAllText(Path.Combine(dir, MetadataFile), meta.ToJsonString(Json.Pretty));
        return new SnapshotInfo(id, meta["createdAt"]!.GetValue<string>(), app,
            Json.GetString(meta, "documentRef"), reason, dir);
    }

    public IReadOnlyList<SnapshotInfo> List(string? app, int limit = 20)
    {
        var result = new List<SnapshotInfo>();
        var root = _options.SnapshotsDir;
        if (!Directory.Exists(root)) return result;

        var appDirs = app is null
            ? Directory.GetDirectories(root)
            : new[] { Path.Combine(root, app) };

        foreach (var appDir in appDirs)
        {
            if (!Directory.Exists(appDir)) continue;
            foreach (var snapDir in Directory.GetDirectories(appDir))
            {
                var metaPath = Path.Combine(snapDir, MetadataFile);
                if (!File.Exists(metaPath)) continue;
                try
                {
                    var meta = JsonNode.Parse(File.ReadAllText(metaPath)) as JsonObject;
                    if (meta is null) continue;
                    result.Add(new SnapshotInfo(
                        Json.GetString(meta, "snapshotId") ?? Path.GetFileName(snapDir),
                        Json.GetString(meta, "createdAt") ?? "",
                        Json.GetString(meta, "app") ?? Path.GetFileName(appDir),
                        Json.GetString(meta, "documentRef"),
                        Json.GetString(meta, "reason") ?? "",
                        snapDir));
                }
                catch { /* 손상된 스냅샷은 건너뜀 */ }
            }
        }
        return result.OrderByDescending(s => s.CreatedAt).Take(Math.Max(1, limit)).ToList();
    }

    public (SnapshotInfo Info, JsonObject Metadata)? Get(string snapshotId)
    {
        foreach (var s in List(null, int.MaxValue))
            if (string.Equals(s.SnapshotId, snapshotId, StringComparison.Ordinal))
            {
                var meta = JsonNode.Parse(File.ReadAllText(Path.Combine(s.Dir, MetadataFile))) as JsonObject
                           ?? new JsonObject();
                return (s, meta);
            }
        return null;
    }
}
