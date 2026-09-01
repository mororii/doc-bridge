using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

/// <summary>
/// 실AutoCAD E2E (M3 인수 조건): DOCBRIDGE_E2E=1 일 때만 실행.
/// 테스트 전용 AutoCAD 인스턴스(비가시)에서 도면/레이어/텍스트를 만들고
/// layer color/visibility, set_text_value, delete_entities(고위험), snapshot/restore를 검증한다.
/// </summary>
public class CadE2ETests : IDisposable
{
    private static bool Enabled =>
        string.Equals(Environment.GetEnvironmentVariable("DOCBRIDGE_E2E"), "1", StringComparison.Ordinal);

    private readonly TestHome _home = new();
    private CadAdapter? _adapter;
    private object? _app;
    private bool _ownsApp;
    private string? _probeDocPath;

    private DocBridgeHost CreateHostWithCad()
    {
        var dwgPath = Path.Combine(_home.Dir, "probe.dwg");
        _adapter = new CadAdapter(() =>
        {
            // 어댑터 STA 스레드 안에서 실행됨.
            // 사용자가 열어둔(로그인된) 인스턴스가 있으면 그것에 붙는다 — 프로덕션 어댑터와 같은
            // ROT 우선 정책. COM으로 새로 띄운 인스턴스는 라이센스 로그인 세션 없이 시작되어
            // 환경에 따라 수 분 내 종료될 수 있으므로, 실행 중 인스턴스를 우선 사용한다.
            const int bootRetries = 90;
            dynamic app;
            var running = RotHelper.GetActiveObject("AutoCAD.Application");
            if (running is not null)
            {
                app = running;
                _ownsApp = false;
            }
            else
            {
                var type = Type.GetTypeFromProgID("AutoCAD.Application")
                    ?? throw new InvalidOperationException("AutoCAD not installed");
                app = Activator.CreateInstance(type)!;
                _ownsApp = true;

                // AutoCAD 부팅 중에는 COM 호출이 거부(RPC_E_CALL_REJECTED)될 수 있어 단계별로 대기한다
                for (var i = 0; ; i++)
                {
                    try { app.Visible = false; break; }
                    catch (System.Runtime.InteropServices.COMException) when (i < bootRetries) { Thread.Sleep(1000); }
                }
            }
            _app = app; // 설정 실패 시에도 Dispose에서 정리 가능하도록 즉시 보관
            dynamic doc;
            for (var i = 0; ; i++)
            {
                try { doc = app.Documents.Add("acad.dwt"); break; }
                catch (System.Runtime.InteropServices.COMException) when (i < bootRetries) { Thread.Sleep(1000); }
            }
            dynamic layer = doc.Layers.Add("DOCBRIDGE");
            layer.LayerOn = true;

            dynamic pt = new double[3];
            dynamic txt = doc.ModelSpace.AddText("초기 텍스트 PROBE", pt, 2.5);
            txt.Layer = "DOCBRIDGE";

            dynamic titleBlock = doc.Blocks.Add(new double[] { 0, 0, 0 }, "DOCBRIDGE-TITLE");
            titleBlock.AddAttribute(2.5, 0, "TITLE", new double[] { 0, 0, 0 }, "TITLE", "초기제목");
            doc.ModelSpace.InsertBlock(new double[] { 80, 0, 0 }, "DOCBRIDGE-TITLE", 1, 1, 1, 0);

            doc.SaveAs(dwgPath);
            // SaveAs가 문서 이름/경로를 바꾸므로 저장 후의 FullName을 보관해야 Dispose가 정확히 닫는다
            _probeDocPath = dwgPath;
            return (object)app;
        });
        var host = new DocBridgeHost(_home.Options);
        host.Router.Register("cad", _adapter);
        return host;
    }

    [Fact]
    public void Cad_full_flow()
    {
        if (!Enabled) return;
        using var host = CreateHostWithCad();

        // 1) get_active_context — 구조화된 JSON
        var ctx = host.GetActiveContext("cad", new JsonObject { ["detailLevel"] = "summary" });
        Assert.True(Json.GetBool(ctx, "ok"), $"context failed: {ctx}");
        Assert.Equal("cad", Json.GetString(ctx, "app"));
        var layers = Json.GetArr(Json.GetObj(ctx, "summary"), "layers")!;
        Assert.Contains(layers, l => Json.GetString(l as JsonObject, "name") == "DOCBRIDGE");

        // 2) cad_query_entities — 텍스트 엔티티 핸들 확보
        var query = host.Read("cad", new JsonObject { ["entityType"] = "Text" });
        Assert.True(Json.GetBool(query, "ok"), $"query failed: {query}");
        var entities = Json.GetArr(query, "entities")!;
        Assert.NotEmpty(entities);
        var handle = Json.GetString(entities[0] as JsonObject, "handle")!;
        Assert.Contains("PROBE", Json.GetString(entities[0] as JsonObject, "text"));

        // 3) set_layer_color + set_text_value dry-run → apply → readback
        var batch = new JsonObject
        {
            ["ops"] = new JsonArray
            {
                new JsonObject { ["op"] = "set_layer_color", ["layer"] = "DOCBRIDGE", ["color"] = 1 },
                new JsonObject { ["op"] = "set_text_value", ["handle"] = handle, ["text"] = "변경된 텍스트 OK" },
            },
            ["dryRun"] = true,
        };
        var dry = host.ApplyOps("cad", batch);
        Assert.True(Json.GetBool(dry, "ok"), $"dry-run failed: {dry}");
        var token = Json.GetString(dry, "confirmToken")!;
        var snapshotId = Json.GetString(dry, "snapshotId")!;

        // 3b) token 없이 apply → 실패
        var noToken = new JsonObject { ["ops"] = batch["ops"]!.DeepClone(), ["dryRun"] = false };
        Assert.False(Json.GetBool(host.ApplyOps("cad", noToken), "ok"));

        var apply = new JsonObject
        {
            ["ops"] = batch["ops"]!.DeepClone(),
            ["dryRun"] = false,
            ["confirmToken"] = token,
        };
        var applied = host.ApplyOps("cad", apply);
        Assert.True(Json.GetBool(applied, "ok"), $"apply failed: {applied}");
        Assert.True(Json.GetBool(Json.GetObj(applied, "readback"), "verified"));

        // 4) 변경 확인
        var after = host.Read("cad", new JsonObject { ["entityType"] = "Text" });
        Assert.Equal("변경된 텍스트 OK", Json.GetString(Json.GetArr(after, "entities")![0] as JsonObject, "text"));

        // 4b) draw_entities — 스크립트 콘솔 없이 COM으로 직접 그리기 (일반 write op)
        var drawBatch = new JsonObject
        {
            ["ops"] = new JsonArray
            {
                new JsonObject
                {
                    ["op"] = "draw_entities",
                    ["entities"] = new JsonArray
                    {
                        new JsonObject
                        {
                            ["type"] = "lwpolyline",
                            ["points"] = new JsonArray(new JsonArray(0, 0), new JsonArray(20, 0), new JsonArray(20, 10)),
                            ["closed"] = true,
                            ["color"] = new JsonObject { ["aci"] = 7 },
                        },
                        new JsonObject
                        {
                            ["type"] = "circle",
                            ["center"] = new JsonArray(5, 5),
                            ["radius"] = 3.0,
                            ["color"] = new JsonObject { ["rgb"] = new JsonArray(255, 0, 0) },
                        },
                        new JsonObject
                        {
                            ["type"] = "hatch",
                            ["loop"] = new JsonObject
                            {
                                ["points"] = new JsonArray(new JsonArray(30, 0), new JsonArray(50, 0), new JsonArray(40, 15)),
                                ["bulges"] = new JsonArray(0, 0, 0),
                            },
                            ["color"] = new JsonObject { ["rgb"] = new JsonArray(0, 0, 255) },
                        },
                        new JsonObject { ["type"] = "line", ["start"] = new JsonArray(0, 20), ["end"] = new JsonArray(20, 20) },
                        new JsonObject { ["type"] = "arc", ["center"] = new JsonArray(30, 25), ["radius"] = 5, ["startAngleDeg"] = 0, ["endAngleDeg"] = 180 },
                        new JsonObject { ["type"] = "ellipse", ["center"] = new JsonArray(50, 25), ["majorAxis"] = new JsonArray(8, 0), ["radiusRatio"] = 0.5 },
                        new JsonObject { ["type"] = "point", ["point"] = new JsonArray(65, 25) },
                        new JsonObject { ["type"] = "mtext", ["point"] = new JsonArray(0, 35), ["width"] = 30, ["height"] = 2.5, ["text"] = "직접 ActiveX MText" },
                        new JsonObject { ["type"] = "dim_aligned", ["start"] = new JsonArray(0, 20), ["end"] = new JsonArray(20, 20), ["textPoint"] = new JsonArray(10, 25) },
                        new JsonObject { ["type"] = "dim_rotated", ["start"] = new JsonArray(30, 20), ["end"] = new JsonArray(50, 20), ["dimensionLinePoint"] = new JsonArray(40, 15), ["rotationDeg"] = 0 },
                    },
                },
            },
            ["dryRun"] = true,
        };
        var drawDry = host.ApplyOps("cad", drawBatch);
        Assert.True(Json.GetBool(drawDry, "ok"), $"draw dry-run failed: {drawDry}");
        Assert.False(Json.GetBool(drawDry, "requiresHighRiskApproval"));

        var drawApply = new JsonObject
        {
            ["ops"] = drawBatch["ops"]!.DeepClone(),
            ["dryRun"] = false,
            ["confirmToken"] = Json.GetString(drawDry, "confirmToken"),
        };
        var drawn = host.ApplyOps("cad", drawApply);
        Assert.True(Json.GetBool(drawn, "ok"), $"draw apply failed: {drawn}");
        Assert.True(Json.GetBool(Json.GetObj(drawn, "readback"), "verified"), $"draw readback: {drawn}");

        // 생성 확인: polyline 1 + circle 1 + hatch 1 (hatch 경계 polyline은 삭제되어야 함)
        var polys = host.Read("cad", new JsonObject { ["entityType"] = "Polyline" });
        Assert.Single(Json.GetArr(polys, "entities")!);
        var circles = host.Read("cad", new JsonObject { ["entityType"] = "Circle" });
        Assert.Single(Json.GetArr(circles, "entities")!);
        var hatches = host.Read("cad", new JsonObject { ["entityType"] = "Hatch" });
        Assert.Single(Json.GetArr(hatches, "entities")!);

        var lines = host.Read("cad", new JsonObject { ["entityType"] = "Line", ["includeGeometry"] = true });
        Assert.Single(Json.GetArr(lines, "entities")!);
        var lineHandle = Json.GetString(Json.GetArr(lines, "entities")![0] as JsonObject, "handle")!;
        Assert.Single(Json.GetArr(host.Read("cad", new JsonObject { ["entityType"] = "Arc" }), "entities")!);
        Assert.Single(Json.GetArr(host.Read("cad", new JsonObject { ["entityType"] = "Ellipse" }), "entities")!);
        Assert.Single(Json.GetArr(host.Read("cad", new JsonObject { ["entityType"] = "Point" }), "entities")!);
        Assert.Single(Json.GetArr(host.Read("cad", new JsonObject { ["entityType"] = "MText" }), "entities")!);
        Assert.Single(Json.GetArr(host.Read("cad", new JsonObject { ["entityType"] = "AlignedDimension" }), "entities")!);
        Assert.Single(Json.GetArr(host.Read("cad", new JsonObject { ["entityType"] = "RotatedDimension" }), "entities")!);

        var modifyOps = new JsonArray
        {
            new JsonObject { ["op"] = "copy_entities", ["handles"] = new JsonArray(lineHandle), ["dx"] = 0, ["dy"] = 10 },
            new JsonObject { ["op"] = "scale_entities", ["handles"] = new JsonArray(lineHandle), ["basePoint"] = new JsonArray(0, 20), ["factor"] = 1.5 },
            new JsonObject { ["op"] = "mirror_entities", ["handles"] = new JsonArray(lineHandle), ["axisStart"] = new JsonArray(0, 0), ["axisEnd"] = new JsonArray(0, 50) },
            new JsonObject { ["op"] = "offset_entities", ["handles"] = new JsonArray(lineHandle), ["distance"] = 2.0 },
            new JsonObject { ["op"] = "set_entity_properties", ["handles"] = new JsonArray(lineHandle), ["properties"] = new JsonObject { ["layer"] = "DOCBRIDGE", ["color"] = new JsonObject { ["aci"] = 3 }, ["linetypeScale"] = 2.0 } },
        };
        var modifyDry = host.ApplyOps("cad", new JsonObject { ["ops"] = modifyOps.DeepClone(), ["dryRun"] = true });
        Assert.True(Json.GetBool(modifyDry, "ok"), $"modify dry-run failed: {modifyDry}");
        var modified = host.ApplyOps("cad", new JsonObject
        {
            ["ops"] = modifyOps.DeepClone(), ["dryRun"] = false,
            ["confirmToken"] = Json.GetString(modifyDry, "confirmToken"),
        });
        Assert.True(Json.GetBool(modified, "ok"), $"modify apply failed: {modified}");
        Assert.Equal(4, Json.GetArr(host.Read("cad", new JsonObject { ["entityType"] = "Line" }), "entities")!.Count);
        Assert.Single(Json.GetArr(host.Read("cad", new JsonObject { ["entityType"] = "Line", ["layer"] = "DOCBRIDGE" }), "entities")!);

        var titleBlocks = host.Read("cad", new JsonObject { ["blockName"] = "DOCBRIDGE-TITLE", ["includeGeometry"] = true });
        var titleBlockEntity = Json.GetArr(titleBlocks, "entities")!.Single() as JsonObject;
        var titleBlockHandle = Json.GetString(titleBlockEntity, "handle")!;
        var attributeOps = new JsonArray
        {
            new JsonObject
            {
                ["op"] = "set_block_attributes", ["handle"] = titleBlockHandle,
                ["attributes"] = new JsonObject { ["TITLE"] = "종평면도(1/45)" },
            },
        };
        var attributeDry = host.ApplyOps("cad", new JsonObject { ["ops"] = attributeOps.DeepClone(), ["dryRun"] = true });
        Assert.True(Json.GetBool(attributeDry, "ok"), $"attribute dry-run failed: {attributeDry}");
        var attributeApply = host.ApplyOps("cad", new JsonObject
        {
            ["ops"] = attributeOps.DeepClone(), ["dryRun"] = false,
            ["confirmToken"] = Json.GetString(attributeDry, "confirmToken"),
        });
        Assert.True(Json.GetBool(attributeApply, "ok"), $"attribute apply failed: {attributeApply}");
        var updatedTitleBlock = (JsonObject)Json.GetArr(host.Read("cad", new JsonObject { ["blockName"] = "DOCBRIDGE-TITLE", ["includeGeometry"] = true }), "entities")!.Single()!;
        Assert.Equal("종평면도(1/45)", Json.GetString(Json.GetObj(updatedTitleBlock, "attributes"), "TITLE"));

        var regionCheck = host.Read("cad", new JsonObject
        {
            ["scope"] = "regions",
            ["regions"] = new JsonArray
            {
                new JsonObject
                {
                    ["name"] = "profile-panel", ["bounds"] = new JsonObject { ["minX"] = -1, ["minY"] = 18, ["maxX"] = 72, ["maxY"] = 42 },
                    ["minCount"] = 8,
                },
                new JsonObject
                {
                    ["name"] = "title-block", ["bounds"] = new JsonObject { ["minX"] = 75, ["minY"] = -5, ["maxX"] = 100, ["maxY"] = 10 },
                    ["entityTypes"] = new JsonArray("BlockReference"), ["minCount"] = 1, ["maxCount"] = 1,
                },
            },
        });
        Assert.True(Json.GetBool(regionCheck, "ok"), $"region verification failed: {regionCheck}");
        Assert.Equal(2, Json.GetArr(regionCheck, "regions")!.Count);

        var layoutOps = new JsonArray
        {
            new JsonObject { ["op"] = "configure_layout", ["name"] = "DOCBRIDGE-E2E", ["create"] = true },
            new JsonObject
            {
                ["op"] = "create_viewport", ["layout"] = "DOCBRIDGE-E2E",
                ["center"] = new JsonArray(100, 70), ["width"] = 160, ["height"] = 100,
                ["viewCenter"] = new JsonArray(30, 20), ["viewHeight"] = 60, ["displayLocked"] = true,
            },
        };
        var layoutDry = host.ApplyOps("cad", new JsonObject { ["ops"] = layoutOps.DeepClone(), ["dryRun"] = true });
        Assert.True(Json.GetBool(layoutDry, "ok"), $"layout dry-run failed: {layoutDry}");
        var layoutApply = host.ApplyOps("cad", new JsonObject
        {
            ["ops"] = layoutOps.DeepClone(), ["dryRun"] = false,
            ["confirmToken"] = Json.GetString(layoutDry, "confirmToken"),
        });
        Assert.True(Json.GetBool(layoutApply, "ok"), $"layout apply failed: {layoutApply}");
        var layoutsRead = host.Read("cad", new JsonObject { ["scope"] = "layouts" });
        Assert.True(Json.GetBool(layoutsRead, "ok"));
        Assert.Contains(Json.GetArr(layoutsRead, "layouts")!, item => Json.GetString(item as JsonObject, "name") == "DOCBRIDGE-E2E");

        // 5) 고위험 op: highRiskConfirm 없이는 실패, 함께면 성공
        var delBatch = new JsonObject
        {
            ["ops"] = new JsonArray
            {
                new JsonObject { ["op"] = "delete_entities", ["handles"] = new JsonArray(handle) },
            },
            ["dryRun"] = true,
        };
        var delDry = host.ApplyOps("cad", delBatch);
        Assert.True(Json.GetBool(delDry, "ok"));
        Assert.True(Json.GetBool(delDry, "requiresHighRiskApproval"));

        var delNoConfirm = new JsonObject
        {
            ["ops"] = delBatch["ops"]!.DeepClone(),
            ["dryRun"] = false,
            ["confirmToken"] = Json.GetString(delDry, "confirmToken"),
        };
        Assert.False(Json.GetBool(host.ApplyOps("cad", delNoConfirm), "ok"));

        // 6) 스냅샷 복원 (레이어 색상/텍스트 원복)
        var restoreDry = host.CoreRestoreSnapshot(new JsonObject { ["snapshotId"] = snapshotId });
        Assert.True(Json.GetBool(restoreDry, "ok"));
        var restored = host.CoreRestoreSnapshot(new JsonObject
        {
            ["snapshotId"] = snapshotId,
            ["confirmToken"] = Json.GetString(restoreDry, "confirmToken"),
        });
        Assert.True(Json.GetBool(restored, "ok"), $"restore failed: {restored}");
        var restoredText = host.Read("cad", new JsonObject { ["entityType"] = "Text" });
        Assert.Contains("PROBE", Json.GetString(Json.GetArr(restoredText, "entities")![0] as JsonObject, "text"));
    }

    [Fact]
    public void Cad_layout_plot_and_save_flow()
    {
        if (!Enabled) return;
        using var host = CreateHostWithCad();
        var pdf = Path.Combine(_home.Dir, "cad-production-e2e.pdf");
        var saved = Path.Combine(_home.Dir, "cad-production-saved.dwg");
        var operations = new JsonArray
        {
            new JsonObject
            {
                ["op"] = "configure_layout", ["name"] = "DOCBRIDGE-PLOT", ["create"] = true,
                ["configName"] = "DWG To PDF.pc3",
            },
            new JsonObject { ["op"] = "plot_pdf", ["output"] = pdf, ["configName"] = "DWG To PDF.pc3" },
            new JsonObject { ["op"] = "save_document", ["output"] = saved },
        };
        var dry = host.ApplyOps("cad", new JsonObject { ["ops"] = operations.DeepClone(), ["dryRun"] = true });
        Assert.True(Json.GetBool(dry, "ok"), $"plot/save dry-run failed: {dry}");
        Assert.True(Json.GetBool(dry, "requiresHighRiskApproval"));
        var applied = host.ApplyOps("cad", new JsonObject
        {
            ["ops"] = operations.DeepClone(), ["dryRun"] = false,
            ["confirmToken"] = Json.GetString(dry, "confirmToken"), ["highRiskConfirm"] = true,
        });
        Assert.True(Json.GetBool(applied, "ok"), $"plot/save apply failed: {applied}");
        Assert.True(File.Exists(pdf) && new FileInfo(pdf).Length > 0);
        Assert.True(File.Exists(saved) && new FileInfo(saved).Length > 0);
        _probeDocPath = saved;
        if (Environment.GetEnvironmentVariable("DOCBRIDGE_E2E_ARTIFACTS") is { Length: > 0 } artifactDir)
        {
            Directory.CreateDirectory(artifactDir);
            File.Copy(pdf, Path.Combine(artifactDir, "cad-production-e2e.pdf"), overwrite: true);
        }
    }

    private static void DLog(string msg)
    {
        // COM 정리 문제 진단용. DOCBRIDGE_E2E_DEBUG=1 일 때만 %TEMP%\cad-e2e-dispose.log에 기록.
        if (Environment.GetEnvironmentVariable("DOCBRIDGE_E2E_DEBUG") != "1") return;
        try
        {
            File.AppendAllText(
                Path.Combine(Path.GetTempPath(), "cad-e2e-dispose.log"),
                $"{DateTime.Now:HH:mm:ss.fff} [pid {Environment.ProcessId}] {msg}{Environment.NewLine}");
        }
        catch { }
    }

    public void Dispose()
    {
        // _app RCW는 host(using) 해제 → 어댑터 정리 과정에서 FinalReleaseComObject로
        // 분리(separated)되어 InvalidComObjectException이 된다. 저장된 RCW를 재사용하지 않고
        // ROT에서 새 RCW를 얻어 정리한다 (PowerShell 등 외부 클라이언트와 동일하게 안전).
        DLog($"Dispose enter: _ownsApp={_ownsApp} _probeDocPath={_probeDocPath ?? "null"}");
        try
        {
            var fresh = RotHelper.GetActiveObject("AutoCAD.Application");
            if (fresh is null) { DLog("ROT에 AutoCAD 없음 — 정리 불필요"); }
            else
            {
                dynamic app = fresh;
                if (_ownsApp)
                {
                    // 우리가 만든 인스턴스: 문서 전부 닫고 종료
                    try { foreach (dynamic doc in app.Documents) doc.Close(false); } catch (Exception ex) { DLog($"ownsApp close docs: {ex.GetType().Name} {ex.Message}"); }
                    try { app.Quit(); } catch (Exception ex) { DLog($"ownsApp quit: {ex.GetType().Name} {ex.Message}"); }
                }
                else if (_probeDocPath is not null)
                {
                    // 사용자 인스턴스에 붙은 경우: probe 문서만 닫고 인스턴스는 그대로 둔다.
                    // FullName으로 정확히 식별한다 (이름만으로는 동명 문서 오식별 위험).
                    // 활성 문서는 Close가 거부될 수 있으므로, 다른 문서를 먼저 활성화하고 닫는다.
                    // AutoCAD이 바쁘면 RPC_E_SERVERCALL_RETRYLATER로 Close가 실패하므로,
                    // 성공이 확인될 때까지 재시도한다 (done은 성공 시에만 true).
                    for (var attempt = 0; attempt < 5; attempt++)
                    {
                        var done = false;
                        try
                        {
                            dynamic? target = null;
                            dynamic? other = null;
                            var seen = new List<string>();
                            foreach (dynamic doc in app.Documents)
                            {
                                string fn;
                                try { fn = (string)doc.FullName; } catch (Exception ex) { DLog($"attempt{attempt} FullName read fail: {ex.Message}"); continue; }
                                seen.Add(fn);
                                if (string.Equals(fn, _probeDocPath, StringComparison.OrdinalIgnoreCase)) target = doc;
                                else other ??= doc;
                            }
                            DLog($"attempt{attempt}: docs=[{string.Join(" | ", seen)}] target={(target is null ? "null" : "found")} other={(other is null ? "null" : "found")}");
                            if (target is null) { done = true; }
                            else
                            {
                                if (other is not null) { try { other.Activate(); DLog($"attempt{attempt}: other activated"); } catch (Exception ex) { DLog($"attempt{attempt}: other.Activate fail: {ex.GetType().Name} {ex.Message}"); } }
                                try { target.Close(false); DLog($"attempt{attempt}: target closed"); done = true; }
                                catch (Exception ex) { DLog($"attempt{attempt}: target.Close fail: {ex.GetType().Name} {ex.Message}"); }
                            }
                        }
                        catch (Exception ex) { DLog($"attempt{attempt} outer: {ex.GetType().Name} {ex.Message}"); }
                        if (done) break;
                        Thread.Sleep(1000);
                    }
                }
            }
        }
        catch (Exception ex) { DLog($"Dispose outer: {ex.GetType().Name} {ex.Message}"); }
        _adapter?.Dispose();
        _home.Dispose();
        DLog("Dispose exit");
    }
}
