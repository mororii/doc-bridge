using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// open_workbook must not hide a lost bystander. A blank Saved=true Book1 /
/// 통합 문서1 present in the passive all-instance inventory before Open can
/// be consumed by Excel. That is an aggregate failure: report it, do not
/// waive, and do not recreate the untitled book.
/// </summary>
public static class ExcelOpenWorkbookInventoryContract
{
    public static bool IsBlankUntitledName(string? name) =>
        ExcelAuthoringPaths.IsLikelyUnsavedWorkbookName(name);

    public static string IdentityKey(JsonObject book)
    {
        var pid = Json.GetInt(book, "processId") ?? 0;
        var hwnd = Json.GetLong(book, "excelHwnd") ?? 0;
        var name = Json.GetString(book, "name") ?? "";
        var full = Json.GetString(book, "fullName") ?? name;
        return $"{pid}|{hwnd}|{full}|{name}";
    }

    public static bool SameBook(JsonObject left, JsonObject right) =>
        string.Equals(IdentityKey(left), IdentityKey(right), StringComparison.OrdinalIgnoreCase) ||
        (string.Equals(Json.GetString(left, "name"), Json.GetString(right, "name"), StringComparison.OrdinalIgnoreCase) &&
         string.Equals(Json.GetString(left, "fullName"), Json.GetString(right, "fullName"), StringComparison.OrdinalIgnoreCase) &&
         (Json.GetInt(left, "processId") ?? 0) == (Json.GetInt(right, "processId") ?? 0));

    public static bool IsOpenedTarget(JsonObject book, string openedPath)
    {
        var full = Json.GetString(book, "fullName") ?? "";
        var name = Json.GetString(book, "name") ?? "";
        return string.Equals(full, openedPath, StringComparison.OrdinalIgnoreCase) ||
               string.Equals(name, Path.GetFileName(openedPath), StringComparison.OrdinalIgnoreCase);
    }

    public static List<string> LostBystanders(
        IReadOnlyList<JsonObject> before, IReadOnlyList<JsonObject> after, string openedPath)
    {
        var lost = new List<string>();
        foreach (var book in before)
        {
            if (IsOpenedTarget(book, openedPath))
                continue;
            if (after.Any(item => SameBook(book, item)))
                continue;
            var name = Json.GetString(book, "name") ?? Json.GetString(book, "fullName") ?? "unknown";
            lost.Add(name);
        }

        return lost;
    }

    public static bool PreserveAggregateFailure(IReadOnlyList<string> lostBystanders) =>
        lostBystanders.Count > 0;

    public static string FailureMessage(IReadOnlyList<string> lostBystanders, int ownerPid, IReadOnlyList<int> instancePids)
    {
        var instances = instancePids.Count == 0
            ? "none"
            : string.Join(",", instancePids.Distinct());
        return
            "open_workbook lost bystander workbook(s) " +
            string.Join(", ", lostBystanders) +
            $"; target-owner PID {ownerPid}; all-instance PIDs [{instances}]; " +
            "aggregate failure; not waived and not recreated";
    }
}
