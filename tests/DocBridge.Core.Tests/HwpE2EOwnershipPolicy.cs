namespace DocBridge.Core.Tests;

/// <summary>
/// Pure HWP E2E ownership rules. No COM, no process launch — only the fail-closed
/// decisions that must hold before a real-app fixture mutates or cleans a document.
/// </summary>
public enum HwpE2EOwnershipMode
{
    None = 0,
    ExclusiveNewProcess = 1,
    IsolatedTab = 2,
}

public sealed record HwpE2EDocumentSnapshot(
    IReadOnlyList<string> DocumentIds,
    string? ActiveDocumentId,
    bool? ActiveIsEmpty,
    int DocumentCount)
{
    public static HwpE2EDocumentSnapshot Empty { get; } = new([], null, null, 0);

    public static HwpE2EDocumentSnapshot Create(
        IReadOnlyList<string> documentIds,
        string? activeDocumentId,
        bool? activeIsEmpty)
        => new(
            documentIds,
            activeDocumentId,
            activeIsEmpty,
            documentIds.Count);
}

public sealed record HwpE2EOwnershipClaim(
    HwpE2EOwnershipMode Mode,
    int ProcessId,
    IReadOnlyList<string> OwnedDocumentIds,
    IReadOnlyList<string> PreexistingDocumentIds,
    string Reason,
    string? SeededDocumentId = null)
{
    public bool CanMutate => Mode is HwpE2EOwnershipMode.ExclusiveNewProcess
        or HwpE2EOwnershipMode.IsolatedTab;
}

public sealed record HwpE2ECleanupPlan(
    IReadOnlyList<string> DocumentIdsToClose,
    IReadOnlyList<string> RetainedUnknownDocumentIds,
    bool AllowCloseAllWindows,
    bool AllowProcessKill,
    bool RetainUnknownState,
    string Reason);

public static class HwpE2EOwnershipPolicy
{
    public const string SeedText = "사과 가격은 1000원, 배 가격은 2000원입니다. 사과 재고 확인 필요.";

    /// <summary>
    /// Window-handle PID wins when present. Otherwise exactly one new HWP PID
    /// may prove exclusive ownership. Anything else is unknown (0).
    /// </summary>
    public static int ResolveProcessId(
        int windowProcessId,
        IReadOnlyCollection<int> existingProcessIds,
        IReadOnlyCollection<int> currentProcessIds)
    {
        if (windowProcessId > 0) return windowProcessId;
        var existing = existingProcessIds as ISet<int> ?? existingProcessIds.ToHashSet();
        var newcomers = currentProcessIds
            .Where(id => id > 0 && !existing.Contains(id))
            .Distinct()
            .ToList();
        return newcomers.Count == 1 ? newcomers[0] : 0;
    }

    public static bool TryIdentifyCreatedTab(
        HwpE2EDocumentSnapshot? before,
        HwpE2EDocumentSnapshot after,
        out string createdDocumentId)
    {
        createdDocumentId = "";
        var beforeIds = before?.DocumentIds ?? Array.Empty<string>();
        var created = after.DocumentIds
            .Where(id => !string.IsNullOrWhiteSpace(id) &&
                         !beforeIds.Contains(id, StringComparer.Ordinal))
            .Distinct(StringComparer.Ordinal)
            .ToList();
        if (created.Count != 1) return false;
        createdDocumentId = created[0];
        return true;
    }

    /// <summary>
    /// Accepts only an exact before/after delta. Later non-preexisting tabs are not
    /// inferred as owned unless they appear in this captured set.
    /// </summary>
    public static bool TryCaptureCreatedDocumentIds(
        HwpE2EDocumentSnapshot before,
        HwpE2EDocumentSnapshot after,
        out IReadOnlyList<string> created)
    {
        created = [];
        if (before.DocumentIds.Count != before.DocumentCount ||
            after.DocumentIds.Count != after.DocumentCount)
            return false;

        var beforeSet = before.DocumentIds.ToHashSet(StringComparer.Ordinal);
        var added = after.DocumentIds
            .Where(id => !beforeSet.Contains(id))
            .Distinct(StringComparer.Ordinal)
            .ToArray();
        if (added.Length == 0) return false;
        if (after.DocumentCount != before.DocumentCount + added.Length) return false;
        created = added;
        return true;
    }

    public static HwpE2EOwnershipClaim WithAdditionalOwnedDocuments(
        HwpE2EOwnershipClaim claim,
        IReadOnlyList<string> created)
    {
        var owned = claim.OwnedDocumentIds
            .Concat(created.Where(id => !string.IsNullOrWhiteSpace(id)))
            .Distinct(StringComparer.Ordinal)
            .ToArray();
        return claim with { OwnedDocumentIds = owned };
    }

    public static bool TryValidateInventory(
        int reportedCount,
        IReadOnlyList<string?> documentIds,
        string? activeDocumentId,
        bool? activeIsEmpty,
        out HwpE2EDocumentSnapshot snapshot,
        out string error)
    {
        snapshot = HwpE2EDocumentSnapshot.Empty;
        error = "";
        if (reportedCount < 0)
        {
            error = "HWP inventory reported a negative document count.";
            return false;
        }

        if (documentIds.Count != reportedCount)
        {
            error = "HWP inventory failed: collected document IDs differ from XHwpDocuments.Count.";
            return false;
        }

        var seen = new HashSet<string>(StringComparer.Ordinal);
        var ids = new List<string>(reportedCount);
        foreach (var id in documentIds)
        {
            if (string.IsNullOrWhiteSpace(id))
            {
                error = "HWP inventory failed: a document ID could not be read or is empty.";
                return false;
            }

            if (!seen.Add(id))
            {
                error = "HWP inventory failed: duplicate document ID.";
                return false;
            }

            ids.Add(id);
        }

        if (reportedCount > 0 && string.IsNullOrWhiteSpace(activeDocumentId))
        {
            error = "HWP inventory failed: active document ID is empty.";
            return false;
        }

        if (reportedCount > 0 && !seen.Contains(activeDocumentId!))
        {
            error = "HWP inventory failed: active document ID is not in the collected set.";
            return false;
        }

        snapshot = new HwpE2EDocumentSnapshot(ids, activeDocumentId, activeIsEmpty, reportedCount);
        return true;
    }

    public static bool TryProve(
        IReadOnlyCollection<int> existingProcessIds,
        int createdProcessId,
        bool fileNewSucceeded,
        HwpE2EDocumentSnapshot? before,
        HwpE2EDocumentSnapshot after,
        out HwpE2EOwnershipClaim claim,
        out string error)
    {
        claim = new(HwpE2EOwnershipMode.None, 0, [], [], "unproven");
        error = "";

        if (createdProcessId <= 0)
        {
            error = "HWP E2E refused to mutate: process identity is unknown (window PID missing and no unique new Hwp PID).";
            return false;
        }

        var preexisting = before?.DocumentIds.Where(id => !string.IsNullOrWhiteSpace(id)).ToArray()
                          ?? Array.Empty<string>();
        var existing = existingProcessIds as ISet<int> ?? existingProcessIds.ToHashSet();
        var reusedExisting = existing.Contains(createdProcessId);

        if (!reusedExisting)
            return TryProveExclusiveProcess(createdProcessId, fileNewSucceeded, before, after, preexisting, out claim, out error);

        return TryProveIsolatedTab(createdProcessId, fileNewSucceeded, before, after, preexisting, out claim, out error);
    }

    public static bool AllowsDocumentMutation(HwpE2EOwnershipClaim? claim)
        => claim is { CanMutate: true };

    /// <summary>
    /// SelectAll/Delete/Clear are never a fallback for a failed FileNew.
    /// A proven empty owned tab is seeded with InsertText only.
    /// </summary>
    public static bool AllowsClearOrSelectAllDelete(HwpE2EOwnershipClaim? claim) => false;

    public static HwpE2ECleanupPlan PlanCleanup(
        HwpE2EOwnershipClaim? claim,
        IReadOnlyList<string>? currentDocumentIds)
    {
        if (claim is null || !claim.CanMutate)
        {
            return new(
                [],
                currentDocumentIds ?? [],
                AllowCloseAllWindows: false,
                AllowProcessKill: false,
                RetainUnknownState: true,
                Reason: "No proven ownership. Retain unknown HWP state; do not close windows or kill the process.");
        }

        if (currentDocumentIds is null)
        {
            return new(
                [],
                [],
                AllowCloseAllWindows: false,
                AllowProcessKill: false,
                RetainUnknownState: true,
                Reason: "Could not enumerate documents after a proven session. Retain unknown state rather than guessing.");
        }

        var owned = claim.OwnedDocumentIds
            .Where(id => !string.IsNullOrWhiteSpace(id))
            .ToHashSet(StringComparer.Ordinal);
        var toClose = currentDocumentIds
            .Where(owned.Contains)
            .Distinct(StringComparer.Ordinal)
            .ToArray();
        var retainedUnknown = currentDocumentIds
            .Where(id => !string.IsNullOrWhiteSpace(id) && !owned.Contains(id))
            .Distinct(StringComparer.Ordinal)
            .ToArray();

        return new(
            toClose,
            retainedUnknown,
            AllowCloseAllWindows: false,
            AllowProcessKill: false,
            RetainUnknownState: false,
            Reason: retainedUnknown.Length == 0
                ? "FileClose only the explicitly recorded OwnedDocumentIds."
                : "FileClose only explicitly recorded OwnedDocumentIds; retaining unknown tabs: "
                  + string.Join(", ", retainedUnknown));
    }

    public static bool ShouldRetainArtifacts(string? artifactDirectory)
        => !string.IsNullOrWhiteSpace(artifactDirectory);

    public static string UniqueOwnedHwpArtifactPath(string artifactDirectory, string seededDocumentId)
    {
        var safeId = new string(seededDocumentId
            .Select(ch => char.IsLetterOrDigit(ch) ? ch : '-')
            .ToArray());
        if (string.IsNullOrWhiteSpace(safeId)) safeId = "seeded";
        return Path.Combine(
            artifactDirectory,
            $"hwp-e2e-owned-{DateTime.UtcNow:yyyyMMddTHHmmssfff}-{safeId}-{Guid.NewGuid():N}.hwp");
    }

    private static bool TryProveExclusiveProcess(
        int processId,
        bool fileNewSucceeded,
        HwpE2EDocumentSnapshot? before,
        HwpE2EDocumentSnapshot after,
        IReadOnlyList<string> preexisting,
        out HwpE2EOwnershipClaim claim,
        out string error)
    {
        claim = new(HwpE2EOwnershipMode.None, processId, [], preexisting, "unproven");
        error = "";

        if (fileNewSucceeded)
        {
            if (string.IsNullOrWhiteSpace(after.ActiveDocumentId))
            {
                error = "HWP E2E refused to mutate: FileNew succeeded on a new process but the active tab identity is empty.";
                return false;
            }

            if (after.ActiveIsEmpty != true)
            {
                error = "HWP E2E refused to mutate: FileNew tab on a new process is not proven empty; will not SelectAll/Delete.";
                return false;
            }

            var owned = after.DocumentIds
                .Where(id => !string.IsNullOrWhiteSpace(id))
                .Distinct(StringComparer.Ordinal)
                .ToArray();
            if (owned.Length == 0) owned = [after.ActiveDocumentId];
            claim = new(
                HwpE2EOwnershipMode.ExclusiveNewProcess,
                processId,
                owned,
                preexisting,
                "New HWP process PID is not in the pre-existing set; FileNew produced an owned empty tab.",
                after.ActiveDocumentId);
            return true;
        }

        if (after.DocumentCount == 1 &&
            !string.IsNullOrWhiteSpace(after.ActiveDocumentId) &&
            after.ActiveIsEmpty == true)
        {
            claim = new(
                HwpE2EOwnershipMode.ExclusiveNewProcess,
                processId,
                [after.ActiveDocumentId],
                preexisting,
                "New HWP process PID is exclusive; FileNew failed but the single default tab is empty and identified.",
                after.ActiveDocumentId);
            return true;
        }

        error = "HWP E2E refused to mutate: new process is exclusive but FileNew failed and the default document is missing, unknown, or not empty.";
        return false;
    }

    private static bool TryProveIsolatedTab(
        int processId,
        bool fileNewSucceeded,
        HwpE2EDocumentSnapshot? before,
        HwpE2EDocumentSnapshot after,
        IReadOnlyList<string> preexisting,
        out HwpE2EOwnershipClaim claim,
        out string error)
    {
        claim = new(HwpE2EOwnershipMode.None, processId, [], preexisting, "unproven");
        error = "";

        if (!fileNewSucceeded)
        {
            error = "HWP E2E refused to mutate: COM reused an existing HWP process and FileNew failed. Will not clear the current document.";
            return false;
        }

        if (before is null)
        {
            error = "HWP E2E refused to mutate: existing HWP process reused but preexisting tabs could not be inventoried.";
            return false;
        }

        if (!TryIdentifyCreatedTab(before, after, out var createdId))
        {
            error = "HWP E2E refused to mutate: existing HWP process reused and FileNew did not produce exactly one new tab identity.";
            return false;
        }

        if (!string.Equals(after.ActiveDocumentId, createdId, StringComparison.Ordinal))
        {
            error = "HWP E2E refused to mutate: FileNew created a tab but it is not the active document.";
            return false;
        }

        if (after.DocumentCount != before.DocumentCount + 1)
        {
            error = "HWP E2E refused to mutate: FileNew did not increase the tab count by exactly one.";
            return false;
        }

        if (after.ActiveIsEmpty != true)
        {
            error = "HWP E2E refused to mutate: isolated FileNew tab is not proven empty; will not SelectAll/Delete.";
            return false;
        }

        claim = new(
            HwpE2EOwnershipMode.IsolatedTab,
            processId,
            [createdId],
            preexisting,
            "Existing HWP process reused, but FileNew produced exactly one new empty tab with a stable DocumentID.",
            createdId);
        return true;
    }
}
