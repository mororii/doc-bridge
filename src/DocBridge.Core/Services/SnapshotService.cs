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

    /// <summary>Test hook: how many metadata.json files this instance has read.</summary>
    internal int MetadataFilesRead { get; private set; }

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

        capture(dir, meta);

        File.WriteAllText(Path.Combine(dir, MetadataFile), meta.ToJsonString(Json.Pretty));
        return new SnapshotInfo(id, meta["createdAt"]!.GetValue<string>(), app,
            Json.GetString(meta, "documentRef"), reason, dir);
    }

    public IReadOnlyList<SnapshotInfo> List(string? app, int limit = 20)
    {
        var result = new List<SnapshotInfo>();
        var take = Math.Max(1, limit);
        foreach (var dir in EnumerateNewestSnapshotDirectories(app))
        {
            if (!TryGetDirectoryIdentity(dir, out var folderApp, out var folderId))
                continue;
            if (app is not null && !string.Equals(folderApp, app, StringComparison.Ordinal))
                continue;
            if (!TryReadSnapshot(dir, folderApp, folderId, out var info, out _))
                continue;
            result.Add(info);
            if (result.Count >= take) break;
        }
        return result;
    }

    public (SnapshotInfo Info, JsonObject Metadata)? Get(string snapshotId) => Get(snapshotId, app: null);

    /// <summary>
    /// Exact snapshot-id lookup. Probes <c>snapshots/{app}/{snapshotId}/</c> when
    /// <paramref name="app"/> is a safe known app; otherwise probes that id under
    /// each app folder. Never lists or parses sibling snapshot history.
    /// </summary>
    public (SnapshotInfo Info, JsonObject Metadata)? Get(string snapshotId, string? app)
    {
        if (!IsSafeSnapshotId(snapshotId)) return null;
        var root = _options.SnapshotsDir;
        if (!Directory.Exists(root)) return null;

        if (!string.IsNullOrWhiteSpace(app))
        {
            if (!IsSafeSnapshotId(app)) return null;
            return TryGetExact(root, app, snapshotId);
        }

        foreach (var appDir in Directory.GetDirectories(root))
        {
            var folderApp = Path.GetFileName(appDir);
            if (!IsSafeSnapshotId(folderApp)) continue;
            var found = TryGetExact(root, folderApp, snapshotId);
            if (found is not null) return found;
        }

        return null;
    }

    /// <summary>
    /// 동일 문서·동일 ops의 반복 dry-run에 사용할 수 있는 가장 최근 스냅샷 후보를 찾는다.
    /// 여기서는 파일 메타데이터 키만 비교한다. 실제 문서 fingerprint 일치는 호출자가
    /// IPreviewReuseAdapter.ValidatePreviewReuse로 다시 검증해야 한다.
    ///
    /// opsHash까지 키에 포함하는 이유는 일부 어댑터의 rollback payload가 operation-scoped이기
    /// 때문이다. 다른 ops 사이에서 스냅샷을 공유해 성능을 얻는 대신 복원 정확성을 잃지 않는다.
    /// </summary>
    public (SnapshotInfo Info, JsonObject Metadata)? FindLatestReusableCandidate(
        string app,
        string? documentRef,
        string opsHash,
        Func<string?, string?, bool> sameDocument,
        int searchLimit = 20)
    {
        if (!IsSafeSnapshotId(app)) return null;
        var remaining = Math.Clamp(searchLimit, 1, 100);
        foreach (var dir in EnumerateNewestSnapshotDirectories(app))
        {
            if (remaining <= 0) break;
            if (!TryGetDirectoryIdentity(dir, out var folderApp, out var folderId)
                || !string.Equals(folderApp, app, StringComparison.Ordinal))
            {
                continue;
            }

            if (!TryReadSnapshot(dir, folderApp, folderId, out var info, out var metadata, out var metadataRead))
            {
                if (metadataRead) remaining--;
                continue;
            }

            remaining--;
            if (Json.GetInt(metadata, "snapshotReuseVersion") == 1
                && string.Equals(Json.GetString(metadata, "opsHash"), opsHash, StringComparison.Ordinal)
                && sameDocument(info.DocumentRef, documentRef)
                && ApplyPreviewArtifact.FromMetadata(metadata, opsHash) is not null)
            {
                return (info, metadata);
            }
        }

        return null;
    }

    internal static bool IsSafeSnapshotId(string? snapshotId)
    {
        if (string.IsNullOrWhiteSpace(snapshotId)) return false;
        if (snapshotId is "." or "..") return false;
        if (snapshotId.Contains('/') || snapshotId.Contains('\\') || snapshotId.Contains(':'))
            return false;
        if (snapshotId.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
            return false;
        return snapshotId.IndexOfAny(Path.GetInvalidPathChars()) < 0;
    }

    private (SnapshotInfo Info, JsonObject Metadata)? TryGetExact(string root, string app, string snapshotId)
    {
        var dir = Path.Combine(root, app, snapshotId);
        if (!Directory.Exists(dir) || !IsDirectoryUnder(root, dir))
            return null;
        if (!TryGetDirectoryIdentity(dir, out var folderApp, out var folderId)
            || !string.Equals(folderApp, app, StringComparison.Ordinal)
            || !string.Equals(folderId, snapshotId, StringComparison.Ordinal))
        {
            return null;
        }

        if (!TryReadSnapshot(dir, folderApp, folderId, out var info, out var metadata))
            return null;
        return (info, metadata);
    }

    private static bool TryGetDirectoryIdentity(string dir, out string folderApp, out string folderId)
    {
        folderApp = "";
        folderId = Path.GetFileName(dir.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
        var parent = Directory.GetParent(dir);
        if (parent is null)
            return false;
        folderApp = parent.Name;
        return IsSafeSnapshotId(folderApp) && IsSafeSnapshotId(folderId);
    }

    private IEnumerable<string> EnumerateNewestSnapshotDirectories(string? app)
    {
        var root = _options.SnapshotsDir;
        if (!Directory.Exists(root)) return Array.Empty<string>();

        var dirs = new List<string>();
        IEnumerable<string> appDirs;
        if (app is null)
        {
            appDirs = Directory.GetDirectories(root);
        }
        else
        {
            if (!IsSafeSnapshotId(app)) return Array.Empty<string>();
            appDirs = new[] { Path.Combine(root, app) };
        }

        foreach (var appDir in appDirs)
        {
            if (!Directory.Exists(appDir)) continue;
            foreach (var snapDir in Directory.GetDirectories(appDir))
            {
                if (!IsSafeSnapshotId(Path.GetFileName(snapDir))) continue;
                if (IsDirectoryUnder(root, snapDir))
                    dirs.Add(snapDir);
            }
        }

        return dirs.OrderByDescending(Path.GetFileName, StringComparer.Ordinal);
    }

    private bool TryReadSnapshot(
        string dir, string folderApp, string folderId, out SnapshotInfo info, out JsonObject metadata)
        => TryReadSnapshot(dir, folderApp, folderId, out info, out metadata, out _);

    private bool TryReadSnapshot(
        string dir,
        string folderApp,
        string folderId,
        out SnapshotInfo info,
        out JsonObject metadata,
        out bool metadataRead)
    {
        info = new SnapshotInfo("", "", folderApp, null, "", dir);
        metadata = new JsonObject();
        metadataRead = false;
        if (!IsSafeSnapshotId(folderApp) || !IsSafeSnapshotId(folderId))
            return false;
        if (!TryGetDirectoryIdentity(dir, out var actualApp, out var actualId)
            || !string.Equals(actualApp, folderApp, StringComparison.Ordinal)
            || !string.Equals(actualId, folderId, StringComparison.Ordinal))
        {
            return false;
        }

        var metaPath = Path.Combine(dir, MetadataFile);
        if (!File.Exists(metaPath)) return false;
        try
        {
            MetadataFilesRead++;
            metadataRead = true;
            if (JsonNode.Parse(File.ReadAllText(metaPath)) is not JsonObject meta)
                return false;
            if (!MetadataIdentityMatchesDirectory(meta, actualApp, actualId))
                return false;

            metadata = meta;
            info = new SnapshotInfo(
                actualId,
                Json.GetString(meta, "createdAt") ?? "",
                actualApp,
                Json.GetString(meta, "documentRef"),
                Json.GetString(meta, "reason") ?? "",
                dir);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static bool MetadataIdentityMatchesDirectory(JsonObject meta, string folderApp, string folderId)
    {
        if (meta.ContainsKey("snapshotId")
            && !string.Equals(Json.GetString(meta, "snapshotId"), folderId, StringComparison.Ordinal))
        {
            return false;
        }

        if (meta.ContainsKey("app")
            && !string.Equals(Json.GetString(meta, "app"), folderApp, StringComparison.Ordinal))
        {
            return false;
        }

        return true;
    }

    private static bool IsDirectoryUnder(string root, string candidate)
    {
        var rootFull = Path.GetFullPath(root);
        var candidateFull = Path.GetFullPath(candidate);
        var prefix = rootFull.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)
                     + Path.DirectorySeparatorChar;
        return candidateFull.StartsWith(prefix, StringComparison.OrdinalIgnoreCase);
    }
}
