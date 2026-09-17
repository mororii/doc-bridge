using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// Honesty rules for data-objects recovery artifacts. A last-saved-file copy
/// is not a current-memory backup. A string marker is not proof the file exists.
/// </summary>
public static class ExcelDataRecoveryContract
{
    public const string FreshNativeSource = "current-memory-savecopyas";
    public const string LastSavedSource = "last-saved-file";

    public static bool TryFreshNativeCopy(JsonObject? backupMetadata, out string artifactName)
    {
        artifactName = "";
        if (backupMetadata is null) return false;
        if (!Json.GetBool(backupMetadata, "workbookBackupFresh")) return false;
        if (!Json.GetBool(backupMetadata, "workbookBackupAvailable")) return false;
        if (!string.Equals(Json.GetString(backupMetadata, "workbookBackupSource"), FreshNativeSource,
                StringComparison.Ordinal))
            return false;
        var name = Json.GetString(backupMetadata, "workbookBackup");
        if (string.IsNullOrWhiteSpace(name)) return false;
        artifactName = name;
        return true;
    }

    public static JsonObject SurfaceRecoveryFields(JsonObject? backupMetadata)
    {
        if (TryFreshNativeCopy(backupMetadata, out var name))
        {
            return new JsonObject
            {
                ["recoveryAvailable"] = true,
                ["recoveryArtifact"] = name,
                ["recoverySource"] = FreshNativeSource,
            };
        }

        var source = Json.GetString(backupMetadata, "workbookBackupSource");
        return new JsonObject
        {
            ["recoveryAvailable"] = false,
            ["recoveryArtifact"] = null,
            ["recoverySource"] = string.IsNullOrWhiteSpace(source) ? "none" : source,
            ["recoveryReason"] =
                "auxiliary workbook backup is last-saved-file unless a fresh in-memory SaveCopyAs was actually taken; " +
                "COM surface restore is the supported inverse and this marker is not a recovery artifact",
        };
    }
}
