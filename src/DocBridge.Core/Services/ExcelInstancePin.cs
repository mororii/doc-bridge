using System.Globalization;
using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// COM-free persistence and validation for Excel instance pinning: the user
/// designates one Excel window (processId/HWND) and later calls attach only
/// to it instead of the auto-selected instance. The pin file lives next to
/// the runtime root so CLI (fresh process per call) and MCP share it.
/// </summary>
public static class ExcelInstancePin
{
    public const int SchemaVersion = 1;

    public sealed record PinState(int ProcessId, long Hwnd, string PinnedAtUtc, string Origin, string ProcessStartUtc);

    public static string RootDir() =>
        Environment.GetEnvironmentVariable("DOCBRIDGE_HOME")
        ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DocBridge");

    public static string PinPath() => Path.Combine(RootDir(), "excel-instance-pin.json");

    public static string? ProcessStartTimestamp(int processId)
    {
        try
        {
            using var process = System.Diagnostics.Process.GetProcessById(processId);
            return process.StartTime.ToUniversalTime().ToString("o", CultureInfo.InvariantCulture);
        }
        catch
        {
            return null;
        }
    }

    public static bool TrySavePin(int processId, long hwnd, string origin, out string? error)
    {
        error = null;
        if (processId < 1)
        {
            error = "processId must be a positive integer";
            return false;
        }
        if (hwnd < 1)
        {
            error = "hwnd must be a positive integer";
            return false;
        }
        try
        {
            Directory.CreateDirectory(RootDir());
            var state = new JsonObject
            {
                ["version"] = SchemaVersion,
                ["processId"] = processId,
                ["hwnd"] = hwnd,
                ["pinnedAtUtc"] = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture),
                ["origin"] = string.IsNullOrWhiteSpace(origin) ? "excel_launch" : origin,
                ["processStartUtc"] = ProcessStartTimestamp(processId),
            };
            var temp = PinPath() + ".tmp";
            File.WriteAllText(temp, state.ToJsonString(Json.Pretty));
            File.Move(temp, PinPath(), overwrite: true);
            return true;
        }
        catch (Exception ex)
        {
            error = ex.Message;
            return false;
        }
    }

    public static PinState? TryLoadPin()
    {
        try
        {
            var path = PinPath();
            if (!File.Exists(path)) return null;
            if (JsonNode.Parse(File.ReadAllText(path)) is not JsonObject state) return null;
            if (Json.GetInt(state, "version") != SchemaVersion) return null;
            var pid = Json.GetInt(state, "processId") ?? 0;
            var hwnd = Json.GetLong(state, "hwnd") ?? 0;
            if (pid < 1 || hwnd < 1) return null;
            return new PinState(pid, hwnd,
                Json.GetString(state, "pinnedAtUtc") ?? "",
                Json.GetString(state, "origin") ?? "",
                Json.GetString(state, "processStartUtc") ?? "");
        }
        catch
        {
            return null;
        }
    }

    public static void ClearPin()
    {
        try { File.Delete(PinPath()); }
        catch { /* best effort */ }
    }

    /// <summary>
    /// COM-free validation for excel_launch instance selectors. At most one
    /// selector per call; processId+hwnd may be combined as a pair.
    /// dedicatedInstance never pins and rejects selectors.
    /// </summary>
    public static bool TryValidateSelectors(JsonObject? args, ICollection<string> errors)
    {
        var before = errors.Count;
        if (args is null) return true;
        var hasPid = args.ContainsKey("processId");
        var hasHwnd = args.ContainsKey("hwnd");
        var hasActive = Json.GetBool(args, "activeWindow");
        var hasClear = Json.GetBool(args, "clearPin");
        var dedicated = Json.GetBool(args, "dedicatedInstance");

        if (hasPid && (!ExcelDataOperationsContract.TryGetFiniteNumber(args["processId"], out var pid) ||
                       pid != Math.Truncate(pid) || pid < 1 || pid > int.MaxValue))
            errors.Add("excel_launch processId must be a positive integer");
        if (hasHwnd && (!ExcelDataOperationsContract.TryGetFiniteNumber(args["hwnd"], out var hwnd) ||
                        hwnd != Math.Truncate(hwnd) || hwnd < 1))
            errors.Add("excel_launch hwnd must be a positive integer");
        if (args.ContainsKey("activeWindow") && !IsBool(args["activeWindow"]))
            errors.Add("excel_launch activeWindow must be boolean");
        if (args.ContainsKey("clearPin") && !IsBool(args["clearPin"]))
            errors.Add("excel_launch clearPin must be boolean");

        var selectorCount = (hasPid || hasHwnd ? 1 : 0) + (hasActive ? 1 : 0);
        if (selectorCount > 1)
            errors.Add("excel_launch accepts one instance selector: a processId/hwnd pair, hwnd alone, or activeWindow");
        if (dedicated && (hasPid || hasHwnd || hasActive))
            errors.Add("excel_launch dedicatedInstance cannot be combined with processId, hwnd, or activeWindow");
        if (hasClear && (hasPid || hasHwnd || hasActive || dedicated))
            errors.Add("excel_launch clearPin cannot be combined with selectors or dedicatedInstance");
        return errors.Count == before;
    }

    private static bool IsBool(JsonNode? node) =>
        node is JsonValue value && value.TryGetValue<bool>(out _);
}
