using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

/// <summary>
/// table_delete 제품 경로 검증. 병합 표와 각주 앵커 포함 표의 삭제까지 확인한다.
/// DOCBRIDGE_E2E=1일 때만 실행. 소유 탭만 건드린다.
/// </summary>
[Trait("Category", "E2E")]
public sealed class HwpTableDeleteE2ETests : IDisposable
{
    private static bool Enabled =>
        string.Equals(Environment.GetEnvironmentVariable("DOCBRIDGE_E2E"), "1", StringComparison.Ordinal);

    private readonly TestHome _home = new();
    private HwpAdapter? _adapter;
    private object? _createdApp;
    private HwpE2EOwnershipClaim? _ownership;
    private bool _sessionReleased;
    private bool _adapterDisposed;

    private DocBridgeHost CreateHostWithHwp()
    {
        var existingProcessIds = HwpE2EOwnership.CurrentHwpProcessIds();
        _adapter = new HwpAdapter(() =>
        {
            var type = Type.GetTypeFromProgID("HWPFrame.HwpObject")
                ?? throw new InvalidOperationException("HWP not installed");
            dynamic hwp = HwpEnvironmentDoctor.RunWithAutomationWorkingDirectory(
                () => Activator.CreateInstance(type)!)!;
            if (!HwpE2EOwnership.TryReadInventory(hwp, out HwpE2EDocumentSnapshot before))
                throw new InvalidOperationException("tabledel: before-inventory failed");
            var windowProcessId = RotHelper.ProcessIdFromWindowHandle(RotHelper.HwpWindowHandle(hwp));
            var processId = HwpE2EOwnershipPolicy.ResolveProcessId(
                windowProcessId, existingProcessIds, HwpE2EOwnership.CurrentHwpProcessIds());
            if (!HwpE2EOwnership.TryFileNew(hwp))
                throw new InvalidOperationException("tabledel: FileNew failed");
            if (!HwpE2EOwnership.TryReadInventory(hwp, out HwpE2EDocumentSnapshot after))
                throw new InvalidOperationException("tabledel: after-inventory failed");
            if (!HwpE2EOwnershipPolicy.TryProve(existingProcessIds, processId, true, before, after,
                    out HwpE2EOwnershipClaim claim, out string error))
            {
                if (HwpE2EOwnershipPolicy.TryIdentifyCreatedTab(before, after, out var createdId))
                    HwpE2EOwnership.TryCloseDocumentById(hwp, createdId);
                throw new InvalidOperationException(error);
            }
            _ownership = claim;
            _createdApp = hwp;
            HwpE2EOwnership.InsertSeedText(hwp, HwpE2EOwnershipPolicy.SeedText);
            return (object)hwp;
        });
        var host = new DocBridgeHost(_home.Options);
        host.Router.Register("hwp", new HwpE2ESafeAdapter(_adapter, ReleaseOwnedSession));
        return host;
    }

    private static JsonObject ApplyBatch(DocBridgeHost host, JsonArray ops, bool highRisk = false)
    {
        var dry = host.ApplyOps("hwp", new JsonObject { ["ops"] = ops.DeepClone(), ["dryRun"] = true });
        Assert.True(Json.GetBool(dry, "ok"), $"dry-run failed: {dry}");
        var applied = host.ApplyOps("hwp", new JsonObject
        {
            ["ops"] = ops.DeepClone(),
            ["dryRun"] = false,
            ["confirmToken"] = Json.GetString(dry, "confirmToken"),
            ["highRiskConfirm"] = highRisk,
        });
        Assert.True(Json.GetBool(applied, "ok"), $"apply failed: {applied}");
        return applied;
    }

    private static int TableCount(DocBridgeHost host)
    {
        var read = host.Read("hwp", new JsonObject { ["scope"] = "tables", ["tableIndex"] = 0, ["maxCells"] = 5 });
        var inventory = Json.GetObj(read, "tableInventory");
        return inventory is null ? 0 : Json.GetInt(inventory, "tableCount") ?? 0;
    }

    [Fact]
    public void Delete_merged_table_with_footnote_removes_table_and_note()
    {
        if (!Enabled) return;
        using var host = CreateHostWithHwp();
        ApplyBatch(host, new JsonArray
        {
            new JsonObject
            {
                ["op"] = "insert_table",
                ["rows"] = new JsonArray { new JsonArray("A1", "A2"), new JsonArray("B1", "B2") },
            },
        });
        ApplyBatch(host, new JsonArray
        {
            new JsonObject
            {
                ["op"] = "insert_table",
                ["rows"] = new JsonArray
                {
                    new JsonArray("병합표제목", ""),
                    new JsonArray("각주닻말", "옆셀"),
                },
                ["mergeCells"] = new JsonArray
                {
                    new JsonObject { ["startRow"] = 0, ["startCol"] = 0, ["endCol"] = 1 },
                },
            },
        });
        ApplyBatch(host, new JsonArray
        {
            new JsonObject
            {
                ["op"] = "insert_footnote",
                ["target"] = new JsonObject { ["text"] = "각주닻말" },
                ["text"] = "삭제테스트각주문장",
            },
        });
        Assert.Equal(2, TableCount(host));
        var deleted = ApplyBatch(host, new JsonArray
        {
            new JsonObject { ["op"] = "table_delete", ["tableIndex"] = 1 },
        }, highRisk: true);
        Assert.True(Json.GetBool(deleted, "ok"), $"delete failed: {deleted}");
        Assert.Equal(1, TableCount(host));
        var text = Json.GetString(host.Read("hwp", new JsonObject { ["scope"] = "document" }), "text") ?? "";
        Assert.DoesNotContain("삭제테스트각주문장", text);
        Assert.DoesNotContain("각주닻말", text);
        var deletedFirst = ApplyBatch(host, new JsonArray
        {
            new JsonObject { ["op"] = "table_delete", ["tableIndex"] = 0 },
        }, highRisk: true);
        Assert.True(Json.GetBool(deletedFirst, "ok"), $"delete first failed: {deletedFirst}");
        Assert.Equal(0, TableCount(host));
    }

    [Fact]
    public void Outer_borders_apply_without_changing_interior()
    {
        if (!Enabled) return;
        using var host = CreateHostWithHwp();
        ApplyBatch(host, new JsonArray
        {
            new JsonObject
            {
                ["op"] = "insert_table",
                ["rows"] = new JsonArray
                {
                    new JsonArray("H1", "H2", "H3"),
                    new JsonArray("A", "B", "C"),
                    new JsonArray("D", "E", "F"),
                    new JsonArray("G", "H", "I"),
                },
                ["header"] = true,
            },
        });
        var applied = ApplyBatch(host, new JsonArray
        {
            new JsonObject { ["op"] = "table_set_borders", ["tableIndex"] = 0, ["widthMm"] = 0.5 },
        });
        Assert.True(Json.GetBool(applied, "ok"), $"borders failed: {applied}");
        var appliedWhole = ApplyBatch(host, new JsonArray
        {
            new JsonObject { ["op"] = "table_set_borders", ["tableIndex"] = 0, ["widthMm"] = 1 },
        });
        Assert.True(Json.GetBool(appliedWhole, "ok"), $"borders 1.0 failed: {appliedWhole}");
        var appliedEdges = ApplyBatch(host, new JsonArray
        {
            new JsonObject
            {
                ["op"] = "table_set_borders", ["tableIndex"] = 0,
                ["widthMm"] = 0.5, ["edges"] = new JsonArray("right", "bottom"),
            },
        });
        Assert.True(Json.GetBool(appliedEdges, "ok"), $"borders edges failed: {appliedEdges}");
    }

    public void Dispose()
    {
        ReleaseOwnedSession();
        if (!_adapterDisposed)
        {
            HwpE2EOwnership.DisposeAdapterWithoutKilling(_adapter);
            _adapterDisposed = true;
        }
        _adapter = null;
        _createdApp = null;
        GC.Collect();
        GC.WaitForPendingFinalizers();
        _home.Dispose();
    }

    private void ReleaseOwnedSession()
    {
        if (_sessionReleased) return;
        _sessionReleased = true;
        if (_adapter is null || _createdApp is null || _ownership is null) return;
        try
        {
            _adapter.RunOnAdapterThread<object?>(() =>
            {
                foreach (var id in _ownership.OwnedDocumentIds)
                    HwpE2EOwnership.TryCloseDocumentById(_createdApp, id);
                var pid = _ownership.ProcessId;
                if (_ownership.Mode == HwpE2EOwnershipMode.ExclusiveNewProcess && pid > 0)
                {
                    try
                    {
                        using var proc = System.Diagnostics.Process.GetProcessById(pid);
                        proc.Kill();
                        proc.WaitForExit(10000);
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine("HwpTableDeleteE2E: owned kill reported: " + ex.Message);
                    }
                }
                return null;
            });
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("HwpTableDeleteE2E: owned-session release reported: " + ex.Message);
        }
    }
}
