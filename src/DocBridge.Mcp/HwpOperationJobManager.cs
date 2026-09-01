using System.Collections.Concurrent;
using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Mcp;

/// <summary>긴 HWP 쓰기를 MCP 요청 수명과 분리해 client timeout 후 중복 적용을 막는다.</summary>
internal sealed class HwpOperationJobManager
{
    private sealed class Job
    {
        public required string Id { get; init; }
        public required DateTimeOffset CreatedAt { get; init; }
        public DateTimeOffset? StartedAt { get; set; }
        public DateTimeOffset? CompletedAt { get; set; }
        public string Status { get; set; } = "queued";
        public JsonObject? Result { get; set; }
        public string? Error { get; set; }
        public readonly object Gate = new();
    }

    private readonly DocBridgeHost _host;
    private readonly ConcurrentDictionary<string, Job> _jobs = new(StringComparer.Ordinal);
    public HwpOperationJobManager(DocBridgeHost host) => _host = host;

    public JsonObject Submit(JsonObject args)
    {
        Cleanup();
        var batch = args.DeepClone() as JsonObject ?? new JsonObject();
        var id = $"hwp-{Guid.NewGuid():N}";
        var job = new Job { Id = id, CreatedAt = DateTimeOffset.UtcNow };
        _jobs[id] = job;
        _ = Task.Run(() => Run(job, batch));
        return new JsonObject
        {
            ["ok"] = true, ["jobId"] = id, ["status"] = "queued",
            ["createdAt"] = job.CreatedAt, ["pollTool"] = "hwp_get_job",
            ["safeToRetrySubmit"] = false,
        };
    }

    public JsonObject Get(JsonObject args)
    {
        var id = Json.GetString(args, "jobId");
        if (string.IsNullOrWhiteSpace(id)) return Json.ErrorResult("hwp_get_job requires jobId", "hwp");
        if (!_jobs.TryGetValue(id, out var job))
            return Json.ErrorResult($"HWP job '{id}' not found or expired", "hwp");
        lock (job.Gate)
        {
            return new JsonObject
            {
                ["ok"] = job.Status is not "failed", ["jobId"] = job.Id,
                ["status"] = job.Status, ["createdAt"] = job.CreatedAt,
                ["startedAt"] = job.StartedAt, ["completedAt"] = job.CompletedAt,
                ["terminal"] = job.Status is "succeeded" or "failed",
                ["result"] = job.Result?.DeepClone(), ["error"] = job.Error,
            };
        }
    }

    private void Run(Job job, JsonObject batch)
    {
        lock (job.Gate) { job.Status = "running"; job.StartedAt = DateTimeOffset.UtcNow; }
        try
        {
            var result = _host.ApplyOps("hwp", batch);
            lock (job.Gate)
            {
                job.Result = result.DeepClone() as JsonObject;
                job.Status = "succeeded";
                job.CompletedAt = DateTimeOffset.UtcNow;
            }
        }
        catch (Exception ex)
        {
            lock (job.Gate)
            {
                job.Error = ex.Message; job.Status = "failed";
                job.CompletedAt = DateTimeOffset.UtcNow;
            }
        }
    }

    private void Cleanup()
    {
        var cutoff = DateTimeOffset.UtcNow.AddHours(-24);
        foreach (var pair in _jobs)
            if (pair.Value.CompletedAt is { } completed && completed < cutoff)
                _jobs.TryRemove(pair.Key, out _);
    }
}
