using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelDataOperationsRecoveryTests
{
    [Fact]
    public void Missing_metadata_is_not_a_native_copy_artifact()
    {
        var surface = ExcelDataRecoveryContract.SurfaceRecoveryFields(null);
        Assert.False(Json.GetBool(surface, "recoveryAvailable"));
        Assert.Null(Json.GetString(surface, "recoveryArtifact"));
        Assert.Equal("none", Json.GetString(surface, "recoverySource"));
        Assert.False(ExcelDataRecoveryContract.TryFreshNativeCopy(null, out _));
    }

    [Fact]
    public void Last_saved_file_backup_is_not_a_recovery_artifact()
    {
        var metadata = new JsonObject
        {
            ["workbookBackupSource"] = ExcelDataRecoveryContract.LastSavedSource,
            ["workbookBackupFresh"] = false,
            ["workbookBackupAvailable"] = true,
            ["workbookBackup"] = "last-saved.xlsx",
        };
        var surface = ExcelDataRecoveryContract.SurfaceRecoveryFields(metadata);
        Assert.False(Json.GetBool(surface, "recoveryAvailable"));
        Assert.Null(Json.GetString(surface, "recoveryArtifact"));
        Assert.Equal(ExcelDataRecoveryContract.LastSavedSource, Json.GetString(surface, "recoverySource"));
        Assert.Contains("last-saved-file", Json.GetString(surface, "recoveryReason") ?? "");
        Assert.False(ExcelDataRecoveryContract.TryFreshNativeCopy(metadata, out _));
    }

    [Fact]
    public void String_marker_alone_is_not_proof_of_a_fresh_copy()
    {
        var marker = new JsonObject
        {
            ["recoveryArtifact"] = "native-copy",
            ["workbookBackup"] = "native-copy",
            ["workbookBackupSource"] = "native-copy",
            ["workbookBackupAvailable"] = true,
            ["workbookBackupFresh"] = true,
        };
        Assert.False(ExcelDataRecoveryContract.TryFreshNativeCopy(marker, out _));
        var surface = ExcelDataRecoveryContract.SurfaceRecoveryFields(marker);
        Assert.False(Json.GetBool(surface, "recoveryAvailable"));
        Assert.Null(Json.GetString(surface, "recoveryArtifact"));
    }

    [Fact]
    public void Proven_savecopyas_exposes_the_actual_filename()
    {
        var metadata = new JsonObject
        {
            ["workbookBackupSource"] = ExcelDataRecoveryContract.FreshNativeSource,
            ["workbookBackupFresh"] = true,
            ["workbookBackupAvailable"] = true,
            ["workbookBackup"] = "workbook-fresh-20260911.xlsx",
        };
        Assert.True(ExcelDataRecoveryContract.TryFreshNativeCopy(metadata, out var name));
        Assert.Equal("workbook-fresh-20260911.xlsx", name);
        var surface = ExcelDataRecoveryContract.SurfaceRecoveryFields(metadata);
        Assert.True(Json.GetBool(surface, "recoveryAvailable"));
        Assert.Equal("workbook-fresh-20260911.xlsx", Json.GetString(surface, "recoveryArtifact"));
        Assert.Equal(ExcelDataRecoveryContract.FreshNativeSource, Json.GetString(surface, "recoverySource"));
    }

    [Fact]
    public void Fresh_flag_without_filename_is_not_an_artifact()
    {
        var metadata = new JsonObject
        {
            ["workbookBackupSource"] = ExcelDataRecoveryContract.FreshNativeSource,
            ["workbookBackupFresh"] = true,
            ["workbookBackupAvailable"] = true,
            ["workbookBackup"] = "",
        };
        Assert.False(ExcelDataRecoveryContract.TryFreshNativeCopy(metadata, out _));
        Assert.False(Json.GetBool(ExcelDataRecoveryContract.SurfaceRecoveryFields(metadata), "recoveryAvailable"));
    }
}
