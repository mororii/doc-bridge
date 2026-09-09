using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class HostSafetyTests
{
    [Fact]
    public void Abandoned_mutex_is_treated_as_acquired_ownership()
    {
        var name = @"Local\DocBridge.AbandonedMutex." + Guid.NewGuid().ToString("N");
        using var held = new Mutex(false, name);
        var owner = new Thread(() =>
        {
            Assert.True(held.WaitOne(TimeSpan.FromSeconds(5)));
        });
        owner.IsBackground = true;
        owner.Start();
        owner.Join();

        using var waiter = new Mutex(false, name);
        Assert.True(DocBridgeHost.TryAcquireMutex(waiter, TimeSpan.FromSeconds(5), out var abandoned));
        Assert.True(abandoned);
        waiter.ReleaseMutex();
    }

    [Fact]
    public void Restore_requires_explicit_readback_verified()
    {
        var missingReadback = new JsonObject { ["ok"] = true };
        Assert.False(DocBridgeHost.IsVerifiedRestore(missingReadback));
        Assert.False(DocBridgeHost.HasExplicitReadbackProof(null));

        var missingVerified = new JsonObject
        {
            ["ok"] = true,
            ["readback"] = new JsonObject { ["checked"] = 1 },
        };
        Assert.False(DocBridgeHost.IsVerifiedRestore(missingVerified));

        var explicitTrue = new JsonObject
        {
            ["ok"] = true,
            ["readback"] = new JsonObject { ["verified"] = true },
        };
        Assert.True(DocBridgeHost.IsVerifiedRestore(explicitTrue));
        Assert.True(DocBridgeHost.HasExplicitReadbackProof(Json.GetObj(explicitTrue, "readback")));
    }

    [Fact]
    public void Partial_or_incomplete_coverage_is_never_full_rollback_success()
    {
        var lyingReadback = new JsonObject
        {
            ["ok"] = true,
            ["coverage"] = "incomplete",
            ["readback"] = new JsonObject { ["verified"] = true },
        };
        Assert.False(DocBridgeHost.IsVerifiedRestore(lyingReadback));

        var metadataIncomplete = new JsonObject
        {
            ["ok"] = true,
            ["readback"] = new JsonObject { ["verified"] = true },
        };
        Assert.False(DocBridgeHost.IsVerifiedRestore(
            metadataIncomplete, new JsonObject { ["snapshotCoverage"] = "partial" }));
    }

    [Fact]
    public void Cad_complete_scoped_restore_can_verify_without_claiming_full_document()
    {
        var scoped = new JsonObject
        {
            ["ok"] = true,
            ["coverage"] = "complete",
            ["readback"] = new JsonObject { ["verified"] = true },
        };
        Assert.True(DocBridgeHost.IsVerifiedRestore(scoped));
        Assert.False(Json.GetBool(scoped, "fullDocumentRestored"));
    }

    [Fact]
    public void Complete_coverage_without_readback_is_not_verified()
    {
        Assert.False(DocBridgeHost.IsVerifiedRestore(new JsonObject
        {
            ["ok"] = true,
            ["coverage"] = "complete",
        }));
    }
}
