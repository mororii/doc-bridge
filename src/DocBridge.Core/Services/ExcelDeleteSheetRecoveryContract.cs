using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// delete_sheet recovery is a whole-sheet copy from an owned current workbook
/// backup. Used-range values/numberFormat/tabColor replay is not a full restore:
/// it drops formulas, merges, charts, tables, pivots, validation, page setup,
/// view, and row/column geometry.
/// </summary>
public static class ExcelDeleteSheetRecoveryContract
{
    public const string ModeWorkbookCopySheet = "workbook-copy-sheet";
    public const string ModeUnavailable = "unavailable";
    public const string ModeUsedRangeIncomplete = "used-range-incomplete";

    public static readonly string[] LostByUsedRangeReplay =
    [
        "formulas", "merges", "charts", "tables", "pivots", "validation",
        "pageSetup", "view", "rowHeights", "columnWidths",
    ];

    public static bool IsFullRecoveryAvailable(JsonObject? backupMetadata, string? backupPath)
    {
        if (backupMetadata is null || string.IsNullOrWhiteSpace(backupPath) || !File.Exists(backupPath))
            return false;
        if (!Json.GetBool(backupMetadata, "workbookBackupAvailable"))
            return false;
        var source = Json.GetString(backupMetadata, "workbookBackupSource");
        if (string.Equals(source, "current-memory-savecopyas", StringComparison.Ordinal))
            return true;
        if (string.Equals(source, "last-saved-file", StringComparison.Ordinal) &&
            backupMetadata["workbookBackupSavedFlag"] is JsonValue flag &&
            flag.TryGetValue<bool>(out var saved) && saved)
            return true;
        return false;
    }

    public static string RejectReason(JsonObject? backupMetadata, string? backupPath)
    {
        if (string.IsNullOrWhiteSpace(backupPath) || !File.Exists(backupPath))
            return "delete_sheet requires a current workbook copy before write; no backup file exists";
        if (backupMetadata is null || !Json.GetBool(backupMetadata, "workbookBackupAvailable"))
            return "delete_sheet requires a current workbook copy before write; auxiliary backup is unavailable";
        var source = Json.GetString(backupMetadata, "workbookBackupSource");
        if (string.Equals(source, "last-saved-file", StringComparison.Ordinal) &&
            backupMetadata["workbookBackupSavedFlag"] is JsonValue flag &&
            flag.TryGetValue<bool>(out var saved) && !saved)
            return "delete_sheet refuses last-saved-file recovery while the workbook has unsaved changes; need SaveCopyAs current-memory copy";
        return $"delete_sheet recovery is not available from backup source '{source}'";
    }

    public static JsonObject RecoveryEnvelope(
        string mode,
        string? backupPath,
        string? backupSource,
        string? targetWorkbook,
        IReadOnlyList<JsonObject> sheets) =>
        new()
        {
            ["mode"] = mode,
            ["backupPath"] = backupPath,
            ["backupSource"] = backupSource,
            ["targetWorkbook"] = targetWorkbook,
            ["sheets"] = new JsonArray(sheets.Select(sheet => sheet.DeepClone()).ToArray()),
            ["usedRangeReplayIsFullRestore"] = false,
            ["lostIfUsedRangeReplay"] = new JsonArray(LostByUsedRangeReplay.Select(item => (JsonNode)item).ToArray()),
        };
}
