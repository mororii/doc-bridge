using System.Text.Json.Nodes;

namespace DocBridge.Core.Adapters;

/// <summary>AI와 설치 진단기가 재시도 여부를 판단할 수 있는 안정적인 한글 오류 코드.</summary>
public sealed class HwpAutomationException : InvalidOperationException
{
    public HwpAutomationException(string code, string message, string? userAction = null, Exception? inner = null,
        int? retryAfterMs = null)
        : base(message, inner)
    {
        Code = code;
        UserAction = userAction;
        RetryAfterMs = retryAfterMs;
    }

    public string Code { get; }
    public string? UserAction { get; }
    public int? RetryAfterMs { get; }

    public JsonObject ToResult(string app = "hwp")
    {
        var automaticRetry = Code is "HWP_COM_BUSY" or "HWP_WORKER_RESTARTED";
        var delayedRetry = RetryAfterMs is > 0;
        var result = new JsonObject
        {
            ["ok"] = false,
            ["app"] = app,
            ["errorCode"] = Code,
            ["retryable"] = automaticRetry || delayedRetry,
            ["automaticRetry"] = automaticRetry,
            ["retryPolicy"] = new JsonObject
            {
                ["mode"] = automaticRetry ? "immediate" : delayedRetry ? "after-delay" : "manual-after-remediation",
                ["retryAfterMs"] = RetryAfterMs,
            },
            ["errors"] = new JsonArray(Message),
        };
        if (!string.IsNullOrWhiteSpace(UserAction)) result["userAction"] = UserAction;
        if (RetryAfterMs is not null) result["retryAfterMs"] = RetryAfterMs.Value;
        return result;
    }
}
