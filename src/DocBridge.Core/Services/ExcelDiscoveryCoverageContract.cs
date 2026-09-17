using System.Runtime.InteropServices;
using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// A rejected or busy visible Excel instance is unknown coverage, not an
/// empty Workbooks collection. Attaching to a proven-empty remnant must not
/// imply that visible documents vanished.
/// </summary>
public static class ExcelDiscoveryCoverageContract
{
    public const string CoverageUnknown = "unknown";
    public const string CoverageEmpty = "empty";
    public const string CoverageNonEmpty = "nonempty";
    public const string RejectedOrBusy = "RPC_E_CALL_REJECTED_OR_BUSY";

    public static bool IsBusyOrRejected(Exception ex)
    {
        for (var current = ex; current is not null; current = current.InnerException)
        {
            if (current is COMException com &&
                (com.HResult == unchecked((int)0x80010001) ||
                 com.HResult == unchecked((int)0x8001010A)))
                return true;
            if (current is AggregateException agg && agg.InnerExceptions.Any(IsBusyOrRejected))
                return true;
        }

        return false;
    }

    public static string DescribeFailure(Exception ex) =>
        IsBusyOrRejected(ex) ? RejectedOrBusy : $"{ex.GetType().Name}: {ex.Message}";

    public static string Coverage(int? workbookCount, bool readFailed) =>
        readFailed || workbookCount is null
            ? CoverageUnknown
            : workbookCount == 0
                ? CoverageEmpty
                : CoverageNonEmpty;

    public static bool IsProvenEmpty(int? workbookCount, bool readFailed) =>
        !readFailed && workbookCount == 0;

    public static bool SkipWhenRequireWorkbook(int? workbookCount, bool readFailed) =>
        IsProvenEmpty(workbookCount, readFailed);

    public static int DiscoveryScore(int? workbookCount, bool readFailed) =>
        Coverage(workbookCount, readFailed) switch
        {
            CoverageNonEmpty => 1000 + workbookCount.GetValueOrDefault(),
            CoverageUnknown => 500,
            _ => 1,
        };

    public static bool ShouldRetainBusyCandidate(int selectedScore) =>
        DiscoveryScore(null, true) > selectedScore;

    public static JsonObject UnknownInstance(long hwnd, int processId, string failure) =>
        new()
        {
            ["excelHwnd"] = hwnd,
            ["processId"] = processId,
            ["readFailed"] = true,
            ["coverage"] = CoverageUnknown,
            ["failure"] = failure,
        };

    public static bool SameInstanceUnknownCoverage(JsonObject before, IReadOnlyList<JsonObject> after)
    {
        var pid = Json.GetInt(before, "processId") ?? 0;
        var hwnd = Json.GetLong(before, "excelHwnd") ?? 0;
        return after.Any(item =>
            (Json.GetBool(item, "readFailed") ||
             string.Equals(Json.GetString(item, "coverage"), CoverageUnknown, StringComparison.OrdinalIgnoreCase)) &&
            ((pid > 0 && Json.GetInt(item, "processId") == pid) ||
             (hwnd != 0 && Json.GetLong(item, "excelHwnd") == hwnd)));
    }
}
