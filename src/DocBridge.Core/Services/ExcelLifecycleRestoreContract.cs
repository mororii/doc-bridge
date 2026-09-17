using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// Lifecycle rollback is document/session proof only when identities are
/// actually checked. Zero checks cannot be <c>complete</c> or verified.
/// Preexisting bystanders must still be present unless they were owned
/// books this restore is allowed to close. Do not recreate lost books.
/// </summary>
public static class ExcelLifecycleRestoreContract
{
    public const string CompletenessComplete = "complete";
    public const string CompletenessIncomplete = "incomplete";
    public const string CompletenessUnproven = "unproven";

    public static bool ZeroChecksIsComplete => false;

    public static bool IsVerified(int checkedItems, int mismatchCount) =>
        mismatchCount == 0 && checkedItems > 0;

    public static string Completeness(int checkedItems, int mismatchCount)
    {
        if (mismatchCount > 0)
            return CompletenessIncomplete;
        if (checkedItems <= 0)
            return CompletenessUnproven;
        return CompletenessComplete;
    }

    public static string UnprovenMessage =>
        "lifecycle restore checked 0 identities; document/session restoration is unproven";

    public static bool IsOwnedForClose(JsonObject book, JsonArray? owned) =>
        ExcelAuthoringSchema.ShouldCloseOnLifecycleRestore(
            Json.GetString(book, "name"),
            Json.GetString(book, "fullName"),
            owned);

    public static bool SameIdentity(JsonObject left, JsonObject right)
    {
        if (ExcelOpenWorkbookInventoryContract.SameBook(left, right))
            return true;
        var leftFull = Json.GetString(left, "fullName");
        var rightFull = Json.GetString(right, "fullName");
        return !string.IsNullOrWhiteSpace(leftFull) &&
               string.Equals(leftFull, rightFull, StringComparison.OrdinalIgnoreCase);
    }

    public static List<string> LostPreexisting(
        IReadOnlyList<JsonObject> before, IReadOnlyList<JsonObject> after, JsonArray? owned)
    {
        var lost = new List<string>();
        foreach (var book in before)
        {
            if (IsOwnedForClose(book, owned))
                continue;
            if (after.Any(item => SameIdentity(book, item)))
                continue;
            if (ExcelDiscoveryCoverageContract.SameInstanceUnknownCoverage(book, after))
                continue;
            lost.Add(Json.GetString(book, "fullName") ?? Json.GetString(book, "name") ?? "unknown");
        }

        return lost;
    }

    public static int VerifiedPreexistingCount(
        IReadOnlyList<JsonObject> before, IReadOnlyList<JsonObject> after, JsonArray? owned)
    {
        var count = 0;
        foreach (var book in before)
        {
            if (IsOwnedForClose(book, owned))
                continue;
            if (after.Any(item => SameIdentity(book, item)))
                count++;
        }

        return count;
    }
}
