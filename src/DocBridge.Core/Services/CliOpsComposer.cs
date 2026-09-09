using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// CLI --ops merge. An ops object with executionMode=execute keeps that
/// batch as-is (no injected dryRun). Explicit legacy flags are still applied
/// so Host/validator can deny the conflict. Legacy --ops arrays still default
/// to dry-run. No extra execute flag family.
/// </summary>
public static class CliOpsComposer
{
    public sealed record Flags(
        bool OpsLoaded,
        bool DryRun = false,
        string? ConfirmToken = null,
        bool HighRiskConfirm = false);

    public static void Apply(JsonObject toolArgs, Flags flags)
    {
        var execute = string.Equals(
            Json.GetString(toolArgs, "executionMode"),
            OperationValidator.ExecutionModeExecute,
            StringComparison.OrdinalIgnoreCase);

        if (execute)
        {
            if (flags.DryRun) toolArgs["dryRun"] = true;
            if (flags.ConfirmToken is not null)
            {
                toolArgs["dryRun"] = false;
                toolArgs["confirmToken"] = flags.ConfirmToken;
            }
            if (flags.HighRiskConfirm) toolArgs["highRiskConfirm"] = true;
            return;
        }

        if (flags.OpsLoaded && flags.ConfirmToken is null && !flags.DryRun)
            toolArgs["dryRun"] = true;
        if (flags.DryRun) toolArgs["dryRun"] = true;
        if (flags.ConfirmToken is not null)
        {
            toolArgs["dryRun"] = false;
            toolArgs["confirmToken"] = flags.ConfirmToken;
        }
        if (flags.HighRiskConfirm) toolArgs["highRiskConfirm"] = true;
    }
}
