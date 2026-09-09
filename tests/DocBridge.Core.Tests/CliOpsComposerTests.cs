using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public sealed class CliOpsComposerTests
{
    [Fact]
    public void Legacy_ops_file_defaults_to_dry_run()
    {
        var args = new JsonObject { ["ops"] = new JsonArray() };
        CliOpsComposer.Apply(args, new CliOpsComposer.Flags(OpsLoaded: true));
        Assert.True(Json.GetBool(args, "dryRun"));
        Assert.Null(Json.GetString(args, "executionMode"));
    }

    [Fact]
    public void Execute_ops_object_does_not_inject_dry_run()
    {
        var args = new JsonObject
        {
            ["ops"] = new JsonArray(),
            ["executionMode"] = "execute",
            ["requestId"] = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee",
            ["expectedDocumentRef"] = @"C:\doc\book.xlsx",
        };
        CliOpsComposer.Apply(args, new CliOpsComposer.Flags(OpsLoaded: true));
        Assert.False(args.ContainsKey("dryRun"));
        Assert.Equal("execute", Json.GetString(args, "executionMode"));
        Assert.Equal("aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee", Json.GetString(args, "requestId"));
    }

    [Fact]
    public void Explicit_legacy_flags_on_execute_ops_are_kept_for_validator_deny()
    {
        var args = new JsonObject
        {
            ["ops"] = new JsonArray(),
            ["executionMode"] = "execute",
            ["requestId"] = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee",
            ["expectedDocumentRef"] = @"C:\doc\book.xlsx",
        };
        CliOpsComposer.Apply(args, new CliOpsComposer.Flags(OpsLoaded: true, DryRun: true));
        Assert.True(Json.GetBool(args, "dryRun"));
        Assert.Equal("execute", Json.GetString(args, "executionMode"));

        var withToken = new JsonObject
        {
            ["ops"] = new JsonArray(),
            ["executionMode"] = "execute",
            ["requestId"] = "aaaaaaaa-bbbb-cccc-dddd-eeeeeeeeeeee",
        };
        CliOpsComposer.Apply(withToken, new CliOpsComposer.Flags(OpsLoaded: true, ConfirmToken: "conf_x"));
        Assert.Equal("conf_x", Json.GetString(withToken, "confirmToken"));
        Assert.False(Json.GetBool(withToken, "dryRun"));

        var withRisk = new JsonObject
        {
            ["ops"] = new JsonArray(),
            ["executionMode"] = "execute",
        };
        CliOpsComposer.Apply(withRisk, new CliOpsComposer.Flags(OpsLoaded: true, HighRiskConfirm: true));
        Assert.True(Json.GetBool(withRisk, "highRiskConfirm"));
    }

    [Fact]
    public void Confirm_token_still_forces_legacy_apply()
    {
        var args = new JsonObject { ["ops"] = new JsonArray() };
        CliOpsComposer.Apply(args, new CliOpsComposer.Flags(OpsLoaded: true, ConfirmToken: "conf_x"));
        Assert.False(Json.GetBool(args, "dryRun"));
        Assert.Equal("conf_x", Json.GetString(args, "confirmToken"));
    }
}
