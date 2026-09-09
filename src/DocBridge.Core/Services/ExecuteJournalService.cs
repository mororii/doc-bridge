using System.Text;
using System.Text.Json.Nodes;
using DocBridge.Core.Models;

namespace DocBridge.Core.Services;

public enum ExecuteJournalState
{
    Missing,
    Started,
    Completed,
    Unreadable,
}

public sealed record ExecuteJournalRecord(
    string RequestId,
    string App,
    string DocumentRef,
    string OpsHash,
    ExecuteJournalState State,
    JsonObject? Result,
    string? Outcome = null);

/// <summary>
/// Durable execute idempotency journal. Lookup never talks to an app.
/// Started records use FileMode.CreateNew + flush so a crash before completion
/// stays pending and is never auto-retried. Completed records atomically replace
/// the started file. Missing or corrupt completed payloads fail closed.
/// </summary>
public sealed class ExecuteJournalService
{
    public const string StatusStarted = "started";
    public const string StatusCompleted = "completed";

    private readonly string _dir;

    public ExecuteJournalService(DocBridgeOptions options)
    {
        _dir = options.ExecuteJournalsDir;
        Directory.CreateDirectory(_dir);
    }

    public static bool TryNormalizeRequestId(string? raw, out string normalized)
    {
        normalized = "";
        if (string.IsNullOrWhiteSpace(raw) || !Guid.TryParse(raw.Trim(), out var id) || id == Guid.Empty)
            return false;
        normalized = id.ToString("D");
        return true;
    }

    public string PathFor(string normalizedRequestId) =>
        Path.Combine(_dir, normalizedRequestId + ".json");

    public ExecuteJournalRecord Lookup(string normalizedRequestId)
    {
        var path = PathFor(normalizedRequestId);
        if (!File.Exists(path))
            return Empty(normalizedRequestId, ExecuteJournalState.Missing);

        try
        {
            var text = File.ReadAllText(path);
            if (string.IsNullOrWhiteSpace(text))
                return Empty(normalizedRequestId, ExecuteJournalState.Unreadable);

            if (JsonNode.Parse(text) is not JsonObject node)
                return Empty(normalizedRequestId, ExecuteJournalState.Unreadable);

            var status = Json.GetString(node, "status");
            var app = Json.GetString(node, "app") ?? "";
            var documentRef = Json.GetString(node, "documentRef") ?? "";
            var opsHash = Json.GetString(node, "opsHash") ?? "";
            var storedId = Json.GetString(node, "requestId");
            if (!IdsMatch(storedId, normalizedRequestId) || !HasRequiredEnvelope(app, documentRef, opsHash))
                return Empty(normalizedRequestId, ExecuteJournalState.Unreadable);

            if (string.Equals(status, StatusStarted, StringComparison.OrdinalIgnoreCase))
            {
                var outcome = Json.GetString(node, "outcome") ?? "unknown";
                return new ExecuteJournalRecord(
                    normalizedRequestId, app, documentRef, opsHash, ExecuteJournalState.Started, null, outcome);
            }

            if (string.Equals(status, StatusCompleted, StringComparison.OrdinalIgnoreCase))
            {
                if (!TryCompleteResult(node["result"], normalizedRequestId, out var result))
                    return Empty(normalizedRequestId, ExecuteJournalState.Unreadable);
                return new ExecuteJournalRecord(
                    normalizedRequestId, app, documentRef, opsHash, ExecuteJournalState.Completed, result);
            }

            return Empty(normalizedRequestId, ExecuteJournalState.Unreadable);
        }
        catch
        {
            return Empty(normalizedRequestId, ExecuteJournalState.Unreadable);
        }
    }

    public bool Matches(ExecuteJournalRecord record, string app, string documentRef, string opsHash) =>
        string.Equals(record.App, app, StringComparison.OrdinalIgnoreCase) &&
        string.Equals(record.OpsHash, opsHash, StringComparison.Ordinal) &&
        DocBridgeHost.SameDocumentRef(app, record.DocumentRef, documentRef);

    /// <summary>Atomically create a started record. False if the id already exists or IO fails.</summary>
    public bool TryBegin(string normalizedRequestId, string app, string documentRef, string opsHash)
    {
        try
        {
            if (!TryNormalizeRequestId(normalizedRequestId, out var id) ||
                !HasRequiredEnvelope(app, documentRef, opsHash))
                return false;
            normalizedRequestId = id;
            Directory.CreateDirectory(_dir);
            var path = PathFor(normalizedRequestId);
            var payload = new JsonObject
            {
                ["requestId"] = normalizedRequestId,
                ["app"] = app,
                ["documentRef"] = documentRef,
                ["opsHash"] = opsHash,
                ["status"] = StatusStarted,
                ["outcome"] = "unknown",
                ["startedAt"] = DateTimeOffset.UtcNow.ToString("o"),
            };
            return WriteNew(path, payload);
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Atomically replace a matching started record with the completed result.
    /// The path UUID, stored requestId, app, document, and ops hash must match
    /// the started row. Corrupt or mismatched rows are left untouched.
    /// </summary>
    public bool TryComplete(string normalizedRequestId, string app, string documentRef, string opsHash, JsonObject result)
    {
        try
        {
            if (!TryNormalizeRequestId(normalizedRequestId, out var id) ||
                !HasRequiredEnvelope(app, documentRef, opsHash) ||
                !TryCompleteResult(result, id, out var completeResult))
                return false;
            normalizedRequestId = id;
            var existing = Lookup(normalizedRequestId);
            if (existing.State != ExecuteJournalState.Started ||
                !Matches(existing, app, documentRef, opsHash))
                return false;

            Directory.CreateDirectory(_dir);
            var path = PathFor(normalizedRequestId);
            var tmp = path + ".tmp";
            var payload = new JsonObject
            {
                ["requestId"] = normalizedRequestId,
                ["app"] = app,
                ["documentRef"] = documentRef,
                ["opsHash"] = opsHash,
                ["status"] = StatusCompleted,
                ["completedAt"] = DateTimeOffset.UtcNow.ToString("o"),
                ["result"] = completeResult,
            };
            if (!WriteReplaceable(tmp, payload))
                return false;
            if (!File.Exists(path))
            {
                try { File.Delete(tmp); } catch { /* best effort */ }
                return false;
            }
            File.Replace(tmp, path, destinationBackupFileName: null);
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static bool HasRequiredEnvelope(string app, string documentRef, string opsHash) =>
        !string.IsNullOrWhiteSpace(app) &&
        !string.IsNullOrWhiteSpace(documentRef) &&
        !string.IsNullOrWhiteSpace(opsHash);

    private static bool IdsMatch(string? storedId, string normalizedRequestId) =>
        !string.IsNullOrWhiteSpace(storedId) &&
        TryNormalizeRequestId(storedId, out var storedNorm) &&
        TryNormalizeRequestId(normalizedRequestId, out var requestedNorm) &&
        string.Equals(storedNorm, requestedNorm, StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Host completed payloads always store explicit ok plus execute identity
    /// (executionMode, requestId). result={} or a bare ok flag is not complete.
    /// </summary>
    internal static bool TryCompleteResult(JsonNode? node, string expectedRequestId, out JsonObject result)
    {
        result = new JsonObject();
        if (node is not JsonObject obj) return false;
        if (!obj.TryGetPropertyValue("ok", out var okNode) ||
            okNode is not JsonValue okValue ||
            !okValue.TryGetValue<bool>(out _))
            return false;
        if (!string.Equals(
                Json.GetString(obj, "executionMode"),
                OperationValidator.ExecutionModeExecute,
                StringComparison.OrdinalIgnoreCase))
            return false;
        if (!IdsMatch(Json.GetString(obj, "requestId"), expectedRequestId))
            return false;
        result = (JsonObject)obj.DeepClone();
        return true;
    }

    private static bool WriteNew(string path, JsonObject payload)
    {
        using var stream = new FileStream(
            path, FileMode.CreateNew, FileAccess.Write, FileShare.None, 4096,
            FileOptions.None);
        var bytes = Encoding.UTF8.GetBytes(payload.ToJsonString(Json.Pretty));
        stream.Write(bytes, 0, bytes.Length);
        stream.Flush(flushToDisk: true);
        return true;
    }

    private static bool WriteReplaceable(string tmpPath, JsonObject payload)
    {
        using (var stream = new FileStream(
            tmpPath, FileMode.Create, FileAccess.Write, FileShare.None, 4096,
            FileOptions.None))
        {
            var bytes = Encoding.UTF8.GetBytes(payload.ToJsonString(Json.Pretty));
            stream.Write(bytes, 0, bytes.Length);
            stream.Flush(flushToDisk: true);
        }
        return true;
    }

    private static ExecuteJournalRecord Empty(string requestId, ExecuteJournalState state) =>
        new(requestId, "", "", "", state, null);
}
