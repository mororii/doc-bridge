using System.Globalization;
using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// COM-free persistence for the owned HWP automation instance: the PID recorded
/// at owned launch lets later processes rebind the exact instance through ROT
/// and restore its window instead of relying on the visible-window heuristic.
/// Without the pin, an automation instance whose holder exited is skipped as
/// "not visible" and becomes unreachable even though it is alive.
/// </summary>
public static class HwpInstancePin
{
    public const int SchemaVersion = 1;

    public sealed record PinState(int ProcessId, long Hwnd, string PinnedAtUtc, string Origin, string ProcessStartUtc);

    public static string RootDir() =>
        Environment.GetEnvironmentVariable("DOCBRIDGE_HOME")
        ?? Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "DocBridge");

    public static string PinPath() => Path.Combine(RootDir(), "hwp-instance-pin.json");

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
        try
        {
            Directory.CreateDirectory(RootDir());
            var state = new JsonObject
            {
                ["version"] = SchemaVersion,
                ["processId"] = processId,
                ["hwnd"] = hwnd,
                ["pinnedAtUtc"] = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture),
                ["origin"] = string.IsNullOrWhiteSpace(origin) ? "hwp_launch" : origin,
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
            if (pid < 1) return null;
            return new PinState(pid,
                Json.GetLong(state, "hwnd") ?? 0,
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
    /// The pin is usable only while the same Hwp.exe process is alive:
    /// PID alive, image name Hwp, and start timestamp matches (PID reuse guard).
    /// </summary>
    public static bool TryValidatePin(PinState pin)
    {
        if (pin.ProcessId < 1) return false;
        try
        {
            using var process = System.Diagnostics.Process.GetProcessById(pin.ProcessId);
            if (process.HasExited) return false;
            if (!string.Equals(process.ProcessName, "Hwp", StringComparison.OrdinalIgnoreCase))
                return false;
            if (!string.IsNullOrWhiteSpace(pin.ProcessStartUtc))
            {
                var actual = process.StartTime.ToUniversalTime().ToString("o", CultureInfo.InvariantCulture);
                if (!string.Equals(actual, pin.ProcessStartUtc, StringComparison.Ordinal)) return false;
            }
            return true;
        }
        catch
        {
            return false;
        }
    }
}
