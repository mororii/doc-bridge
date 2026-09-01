using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// confirmToken 발급/검증 서비스 (보안 원칙 3: 실제 적용은 confirmToken이 있어야 한다).
/// - 토큰 형식: conf_{id}.{hmac-base64url}
/// - HMAC 키는 사용자 로컬 파일에만 보관
/// - 토큰은 scope+opsHash+snapshotId에 바인딩, 만료(TTL), 단일 사용
/// - 크로스 프로세스 일관성: 파일 기반 저장 + named mutex
/// </summary>
public sealed class ConfirmTokenService
{
    private const string MutexName = @"Global\DocBridge.Tokens";
    private readonly string _dir;
    private readonly string _storePath;
    private readonly string _keyPath;
    private readonly int _ttlSeconds;

    public ConfirmTokenService(Models.DocBridgeOptions options, int ttlSeconds = 300)
    {
        _dir = options.TokensDir;
        Directory.CreateDirectory(_dir);
        _storePath = Path.Combine(_dir, "pending.json");
        _keyPath = Path.Combine(_dir, "token-key.bin");
        _ttlSeconds = ttlSeconds;
    }

    private sealed class PendingEntry
    {
        public string Scope { get; set; } = "";
        public string OpsHash { get; set; } = "";
        public string? SnapshotId { get; set; }
        public DateTimeOffset ExpiresAt { get; set; }
        public bool Used { get; set; }
    }

    private byte[] LoadOrCreateKey()
    {
        if (File.Exists(_keyPath))
        {
            var k = File.ReadAllBytes(_keyPath);
            if (k.Length >= 32) return k;
        }
        var key = RandomNumberGenerator.GetBytes(32);
        File.WriteAllBytes(_keyPath, key);
        return key;
    }

    private string Sign(byte[] key, string payload)
    {
        using var h = new HMACSHA256(key);
        var sig = h.ComputeHash(Encoding.UTF8.GetBytes(payload));
        return Convert.ToBase64String(sig).Replace('+', '-').Replace('/', '_').TrimEnd('=');
    }

    public static string HashOps(IEnumerable<JsonObject> ops)
    {
        var sb = new StringBuilder();
        foreach (var op in ops) sb.Append(Json.Canonical(op)).Append('\n');
        return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(sb.ToString()))).ToLowerInvariant();
    }

    public static string HashScope(string scope) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(scope))).ToLowerInvariant();

    private Dictionary<string, PendingEntry> LoadStore()
    {
        if (!File.Exists(_storePath)) return new();
        try
        {
            var node = JsonNode.Parse(File.ReadAllText(_storePath)) as JsonObject;
            var dict = new Dictionary<string, PendingEntry>();
            if (node is null) return dict;
            foreach (var kv in node)
            {
                if (kv.Value is not JsonObject e) continue;
                dict[kv.Key] = new PendingEntry
                {
                    Scope = Json.GetString(e, "scope") ?? "",
                    OpsHash = Json.GetString(e, "opsHash") ?? "",
                    SnapshotId = Json.GetString(e, "snapshotId"),
                    ExpiresAt = DateTimeOffset.TryParse(Json.GetString(e, "expiresAt"), out var dt)
                        ? dt : DateTimeOffset.MinValue,
                    Used = Json.GetBool(e, "used"),
                };
            }
            return dict;
        }
        catch { return new(); }
    }

    private void SaveStore(Dictionary<string, PendingEntry> store)
    {
        var o = new JsonObject();
        foreach (var kv in store)
        {
            o[kv.Key] = new JsonObject
            {
                ["scope"] = kv.Value.Scope,
                ["opsHash"] = kv.Value.OpsHash,
                ["snapshotId"] = kv.Value.SnapshotId,
                ["expiresAt"] = kv.Value.ExpiresAt.ToString("o"),
                ["used"] = kv.Value.Used,
            };
        }
        File.WriteAllText(_storePath, o.ToJsonString(Json.Pretty));
    }

    /// <summary>토큰 발급. 반환: (token, expiresInSec)</summary>
    public (string Token, int ExpiresInSec) Create(string scope, string opsHash, string? snapshotId)
    {
        using var mutex = new Mutex(false, MutexName);
        mutex.WaitOne(TimeSpan.FromSeconds(10));
        try
        {
            var key = LoadOrCreateKey();
            var id = Guid.NewGuid().ToString("N")[..16];
            var expires = DateTimeOffset.UtcNow.AddSeconds(_ttlSeconds);
            var payload = $"{id}|{scope}|{opsHash}|{snapshotId}|{expires:o}";
            var token = $"conf_{id}.{Sign(key, payload)}";

            var store = LoadStore();
            // 만료 항목 정리
            foreach (var k in store.Where(kv => kv.Value.ExpiresAt < DateTimeOffset.UtcNow).Select(kv => kv.Key).ToList())
                store.Remove(k);
            store[id] = new PendingEntry
            {
                Scope = scope, OpsHash = opsHash, SnapshotId = snapshotId,
                ExpiresAt = expires, Used = false,
            };
            SaveStore(store);
            return (token, _ttlSeconds);
        }
        finally { mutex.ReleaseMutex(); }
    }

    /// <summary>토큰을 소비하지 않고 검증한다. apply 전 preview artifact/문서 상태 확인에 사용한다.</summary>
    public (bool Ok, string? SnapshotId, string? Error) Validate(string token, string scope, string opsHash) =>
        ValidateCore(token, scope, opsHash, consume: false);

    /// <summary>토큰 검증 + 소비(단일 사용). 성공 시 snapshotId 반환.</summary>
    public (bool Ok, string? SnapshotId, string? Error) ValidateAndConsume(string token, string scope, string opsHash) =>
        ValidateCore(token, scope, opsHash, consume: true);

    private (bool Ok, string? SnapshotId, string? Error) ValidateCore(
        string token, string scope, string opsHash, bool consume)
    {
        using var mutex = new Mutex(false, MutexName);
        mutex.WaitOne(TimeSpan.FromSeconds(10));
        try
        {
            if (string.IsNullOrWhiteSpace(token) || !token.StartsWith("conf_", StringComparison.Ordinal))
                return (false, null, "invalid confirmToken format");

            var body = token[5..];
            var dot = body.IndexOf('.');
            if (dot <= 0) return (false, null, "invalid confirmToken format");
            var id = body[..dot];
            var sig = body[(dot + 1)..];

            var store = LoadStore();
            if (!store.TryGetValue(id, out var entry))
                return (false, null, "confirmToken not found (unknown or already purged)");
            if (entry.Used)
                return (false, null, "confirmToken already used (single-use)");
            if (entry.ExpiresAt < DateTimeOffset.UtcNow)
                return (false, null, "confirmToken expired");
            if (!string.Equals(entry.Scope, scope, StringComparison.Ordinal))
                return (false, null, "confirmToken scope mismatch");
            if (!string.Equals(entry.OpsHash, opsHash, StringComparison.Ordinal))
                return (false, null, "confirmToken does not match these ops (ops changed after dry-run)");

            var key = LoadOrCreateKey();
            var payload = $"{id}|{entry.Scope}|{entry.OpsHash}|{entry.SnapshotId}|{entry.ExpiresAt:o}";
            var expected = Sign(key, payload);
            if (!CryptographicOperations.FixedTimeEquals(
                    Encoding.UTF8.GetBytes(expected), Encoding.UTF8.GetBytes(sig)))
                return (false, null, "confirmToken signature invalid");

            if (consume)
            {
                entry.Used = true;
                SaveStore(store);
            }
            return (true, entry.SnapshotId, null);
        }
        finally { mutex.ReleaseMutex(); }
    }
}
