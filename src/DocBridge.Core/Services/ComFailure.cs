using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>Structured COM/write/readback failure details including HRESULT.</summary>
public static class ComFailure
{
    public static JsonObject Describe(Exception error, string stage)
    {
        var hr = error.HResult;
        return new JsonObject
        {
            ["stage"] = stage,
            ["exceptionType"] = error.GetType().FullName,
            ["message"] = error.Message,
            ["hresult"] = FormatHResult(hr),
            ["hresultValue"] = hr,
            ["errorCode"] = error is COMException com ? com.ErrorCode : hr,
        };
    }

    public static string FormatHResult(int hresult) =>
        string.Create(CultureInfo.InvariantCulture, $"0x{unchecked((uint)hresult):X8}");

    public static string FormatMessage(Exception error, string stage, string op, int index)
    {
        var hr = FormatHResult(error.HResult);
        return $"ops[{index}] '{op}' {stage} failed ({error.GetType().Name} {hr}): {error.Message}";
    }
}
