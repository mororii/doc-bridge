using System.Diagnostics;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

/// <summary>
/// 실한글 E2E (M2 인수 조건): DOCBRIDGE_E2E=1 일 때만 실행.
/// 한글 인스턴스는 어댑터 STA 스레드 안에서 생성(팩토리 패턴)해
/// 크로스-아파트먼트 COM 마샬링을 원천 배제한다.
/// </summary>
public class HwpE2ETests : IDisposable
{
    private static bool Enabled =>
        string.Equals(Environment.GetEnvironmentVariable("DOCBRIDGE_E2E"), "1", StringComparison.Ordinal);

    private readonly TestHome _home = new();
    private HwpAdapter? _adapter;
    private object? _createdApp;
    private int _createdProcessId;

    private DocBridgeHost CreateHostWithHwp()
    {
        var existingProcessIds = Process.GetProcessesByName("Hwp")
            .Select(process =>
            {
                using (process) return process.Id;
            })
            .ToHashSet();
        _adapter = new HwpAdapter(() =>
        {
            // 어댑터 STA 스레드 안에서 실행됨 (RCW 생성/사용 아파트먼트 일치)
            var type = Type.GetTypeFromProgID("HWPFrame.HwpObject")
                ?? throw new InvalidOperationException("한글 not installed");
            // Production HwpAdapter repairs process windir/SystemRoot and pins the HWP Bin
            // directory while activating COM.  The injected E2E factory must use the same
            // activation envelope; otherwise older HWP 2024 builds can hang in TourPopup or
            // FontCache initialization even though the production path is healthy.
            dynamic hwp = HwpEnvironmentDoctor.RunWithAutomationWorkingDirectory(
                () => Activator.CreateInstance(type)!)!;
            // 기존 한글 프로세스가 남아 있어도 E2E는 항상 별도의 빈 문서 탭에서 시작한다.
            try { hwp.HAction.Run("FileNew"); } catch { }
            try
            {
                hwp.HAction.Run("MoveDocBegin");
                if (!(bool)hwp.IsEmpty)
                {
                    hwp.HAction.Run("SelectAll");
                    hwp.HAction.Run("Delete");
                }
            }
            catch { }
            dynamic act = hwp.HAction;
            dynamic ps = hwp.HParameterSet.HInsertText;
            act.GetDefault("InsertText", ps.HSet);
            ps.Text = "사과 가격은 1000원, 배 가격은 2000원입니다. 사과 재고 확인 필요.";
            act.Execute("InsertText", ps.HSet);
            _createdApp = hwp;
            _createdProcessId = RotHelper.ProcessIdFromWindowHandle(RotHelper.HwpWindowHandle(hwp));
            if (_createdProcessId == 0)
            {
                _createdProcessId = Process.GetProcessesByName("Hwp")
                    .Select(process =>
                    {
                        using (process) return process.Id;
                    })
                    .FirstOrDefault(id => !existingProcessIds.Contains(id));
            }
            return (object)hwp;
        });
        var host = new DocBridgeHost(_home.Options);
        host.Router.Register("hwp", _adapter);
        return host;
    }

    [Fact]
    public void Hwp_full_flow()
    {
        if (!Enabled) return;
        using var host = CreateHostWithHwp();

        // 1) get_active_context — 구조화된 JSON
        var ctx = host.GetActiveContext("hwp");
        Assert.True(Json.GetBool(ctx, "ok"), $"context failed: {ctx}");
        Assert.Equal("hwp", Json.GetString(ctx, "app"));
        Assert.NotNull(Json.GetString(ctx, "documentRef"));
        var summary = Json.GetObj(ctx, "summary")!;
        Assert.True(Json.GetInt(summary, "textLength") > 10);

        // 새 CLI/MCP 프로세스도 ROT의 기존 한글 창을 status 단계에서 즉시 식별해야 한다.
        // 그렇지 않으면 dry-run과 apply가 서로 다른 프로세스일 때 untitled 문서 ref가 빈 값이 된다.
        using (var freshStatusAdapter = new HwpAdapter())
        {
            var freshStatus = freshStatusAdapter.GetStatus();
            Assert.True(freshStatus.Available);
            Assert.True(freshStatus.Connected, freshStatus.Detail);
            Assert.False(string.IsNullOrWhiteSpace(freshStatus.Document));
        }

        // 2) hwp_read_text (document scope)
        var read = host.Read("hwp", new JsonObject { ["scope"] = "document" });
        Assert.True(Json.GetBool(read, "ok"));
        Assert.Contains("사과", Json.GetString(read, "text"));

        // 2b) 통합 읽기 — 한 COM 호출에서 본문·문단 지도·구조를 함께 반환
        var bundleRead = host.Read("hwp", new JsonObject
        {
            ["scope"] = "bundle",
            ["sections"] = new JsonArray("text", "document_map", "structure"),
        });
        Assert.True(Json.GetBool(bundleRead, "ok"), $"bundle read failed: {bundleRead}");
        var bundle = Json.GetObj(bundleRead, "bundle")!;
        Assert.Contains("사과", Json.GetString(bundle, "text"));
        Assert.NotNull(Json.GetObj(bundle, "documentMap"));
        Assert.NotNull(Json.GetObj(bundle, "structure"));
        Assert.NotNull(Json.GetObj(bundleRead, "timings"));

        // 3) find_replace dry-run → diff + token
        var frBatch = new JsonObject
        {
            ["ops"] = new JsonArray
            {
                new JsonObject
                {
                    ["op"] = "find_replace",
                    ["find"] = "사과",
                    ["replace"] = "청사과",
                    ["options"] = new JsonObject { ["matchCase"] = false },
                },
            },
            ["dryRun"] = true,
        };
        var dry = host.ApplyOps("hwp", frBatch);
        Assert.True(Json.GetBool(dry, "ok"), $"dry-run failed: {dry}");
        var token = Json.GetString(dry, "confirmToken");
        var snapshotId = Json.GetString(dry, "snapshotId");
        Assert.NotNull(token);
        Assert.NotNull(snapshotId);
        Assert.NotEmpty(Json.GetArr(dry, "diff")!);

        // 3b) token 없이 apply → 실패
        var noToken = new JsonObject { ["ops"] = frBatch["ops"]!.DeepClone(), ["dryRun"] = false };
        Assert.False(Json.GetBool(host.ApplyOps("hwp", noToken), "ok"));

        // 4) apply + readback
        var apply = new JsonObject
        {
            ["ops"] = frBatch["ops"]!.DeepClone(),
            ["dryRun"] = false,
            ["confirmToken"] = token,
        };
        var applied = host.ApplyOps("hwp", apply);
        Assert.True(Json.GetBool(applied, "ok"), $"apply failed: {applied}");
        Assert.True(Json.GetBool(Json.GetObj(applied, "readback"), "verified"));
        Assert.True(Json.GetBool(Json.GetObj(applied, "timings"), "previewReused"));

        // 5) 실제 문서 확인 ("청사과"에 "사과"가 부분문자열로 포함되므로 완전 문자열로 검증)
        var after = host.Read("hwp", new JsonObject { ["scope"] = "document" });
        var afterText = Json.GetString(after, "text")!;
        Assert.Contains("청사과 가격은 1000원", afterText);
        Assert.Contains("청사과 재고 확인", afterText);
        Assert.DoesNotContain("배 가격은 1000원", afterText); // 미치환 항목 없음 확인용 위조 문자열

        // 6) insert_text
        var insBatch = new JsonObject
        {
            ["ops"] = new JsonArray { new JsonObject { ["op"] = "insert_text", ["text"] = " DOC-SENTINEL" } },
            ["dryRun"] = true,
        };
        var insDry = host.ApplyOps("hwp", insBatch);
        var insApply = new JsonObject
        {
            ["ops"] = insBatch["ops"]!.DeepClone(),
            ["dryRun"] = false,
            ["confirmToken"] = Json.GetString(insDry, "confirmToken"),
        };
        var insApplied = host.ApplyOps("hwp", insApply);
        Assert.True(Json.GetBool(insApplied, "ok"), $"insert_text failed: {insApplied}");
        var docAfterIns = host.Read("hwp", new JsonObject { ["scope"] = "document" });
        Assert.Contains("DOC-SENTINEL", Json.GetString(docAfterIns, "text"));

        // 7) append_text — 문서 끝으로 이동하고 여러 문단을 한 op로 추가
        var appendBatch = new JsonObject
        {
            ["ops"] = new JsonArray
            {
                new JsonObject
                {
                    ["op"] = "append_text",
                    ["text"] = "첫째 추가 문단\n둘째 추가 문단",
                },
            },
            ["dryRun"] = true,
        };
        var appendDry = host.ApplyOps("hwp", appendBatch);
        Assert.True(Json.GetBool(appendDry, "ok"), $"append dry-run failed: {appendDry}");
        var appendApplied = host.ApplyOps("hwp", new JsonObject
        {
            ["ops"] = appendBatch["ops"]!.DeepClone(),
            ["dryRun"] = false,
            ["confirmToken"] = Json.GetString(appendDry, "confirmToken"),
        });
        Assert.True(Json.GetBool(appendApplied, "ok"), $"append_text failed: {appendApplied}");
        Assert.True(Json.GetBool(Json.GetObj(appendApplied, "readback"), "verified"));
        var docAfterAppend = Json.GetString(host.Read("hwp", new JsonObject { ["scope"] = "document" }), "text")!;
        var normalizedAppend = docAfterAppend.Replace("\r\n", "\n").Replace('\r', '\n');
        Assert.Contains("첫째 추가 문단\n둘째 추가 문단", normalizedAppend);

        // 8) 문서 중간의 고유 기준 문구 앞/뒤에 새 문단 삽입
        var relativeOps = new JsonArray
        {
            new JsonObject
            {
                ["op"] = "insert_before_text",
                ["anchor"] = "배 가격은 2000원입니다.",
                ["text"] = "기준 문단 앞에 삽입",
                ["mode"] = "paragraph",
            },
            new JsonObject
            {
                ["op"] = "insert_after_text",
                ["anchor"] = "배 가격은 2000원입니다.",
                ["text"] = "기준 문단 뒤에 삽입",
                ["mode"] = "paragraph",
            },
        };
        var relativeDry = host.ApplyOps("hwp", new JsonObject { ["ops"] = relativeOps.DeepClone(), ["dryRun"] = true });
        Assert.True(Json.GetBool(relativeDry, "ok"), $"relative insert dry-run failed: {relativeDry}");
        var relativeApplied = host.ApplyOps("hwp", new JsonObject
        {
            ["ops"] = relativeOps.DeepClone(),
            ["dryRun"] = false,
            ["confirmToken"] = Json.GetString(relativeDry, "confirmToken"),
        });
        Assert.True(Json.GetBool(relativeApplied, "ok"), $"relative insert failed: {relativeApplied}");
        Assert.True(Json.GetBool(Json.GetObj(relativeApplied, "readback"), "verified"));
        var relativeText = Json.GetString(host.Read("hwp", new JsonObject { ["scope"] = "document" }), "text")!
            .Replace("\r\n", "\n").Replace('\r', '\n');
        var beforeIndex = relativeText.IndexOf("기준 문단 앞에 삽입", StringComparison.Ordinal);
        var anchorIndex = relativeText.IndexOf("배 가격은 2000원입니다.", StringComparison.Ordinal);
        var afterIndex = relativeText.IndexOf("기준 문단 뒤에 삽입", StringComparison.Ordinal);
        Assert.True(beforeIndex >= 0 && beforeIndex < anchorIndex && anchorIndex < afterIndex,
            $"relative order mismatch: {relativeText}");

        // 9) 중복 기준 문구의 두 번째 항목만 인라인 삽입
        var occurrenceOps = new JsonArray
        {
            new JsonObject
            {
                ["op"] = "insert_after_text",
                ["anchor"] = "청사과",
                ["occurrence"] = 2,
                ["text"] = "[두 번째 항목]",
                ["mode"] = "inline",
            },
        };
        var occurrenceDry = host.ApplyOps("hwp", new JsonObject { ["ops"] = occurrenceOps.DeepClone(), ["dryRun"] = true });
        Assert.True(Json.GetBool(occurrenceDry, "ok"), $"occurrence dry-run failed: {occurrenceDry}");
        var occurrenceApplied = host.ApplyOps("hwp", new JsonObject
        {
            ["ops"] = occurrenceOps.DeepClone(),
            ["dryRun"] = false,
            ["confirmToken"] = Json.GetString(occurrenceDry, "confirmToken"),
        });
        Assert.True(Json.GetBool(occurrenceApplied, "ok"), $"occurrence insert failed: {occurrenceApplied}");
        var occurrenceText = Json.GetString(host.Read("hwp", new JsonObject { ["scope"] = "document" }), "text")!;
        Assert.Equal(1, occurrenceText.Split("청사과[두 번째 항목]", StringSplitOptions.None).Length - 1);
        Assert.DoesNotContain("청사과[두 번째 항목] 가격은", occurrenceText);

        // 10) 스냅샷 복원 (find_replace 시점 상태로)
        var restoreDry = host.CoreRestoreSnapshot(new JsonObject { ["snapshotId"] = snapshotId });
        Assert.True(Json.GetBool(restoreDry, "ok"));
        var restored = host.CoreRestoreSnapshot(new JsonObject
        {
            ["snapshotId"] = snapshotId,
            ["confirmToken"] = Json.GetString(restoreDry, "confirmToken"),
        });
        Assert.True(Json.GetBool(restored, "ok"), $"restore failed: {restored}");
        var restoredText = host.Read("hwp", new JsonObject { ["scope"] = "document" });
        Assert.Contains("사과 가격은 1000원", Json.GetString(restoredText, "text"));
    }

    [Fact]
    public void Hwp_launch_creates_one_document_then_reuses_it()
    {
        if (!Enabled) return;
        using var host = CreateHostWithHwp();
        var first = host.HwpLaunch(new JsonObject { ["newDocument"] = true });
        Assert.True(Json.GetBool(first, "ok"), $"first hwp_launch failed: {first}");
        Assert.True(Json.GetBool(Json.GetObj(first, "summary"), "createdDocument"));
        var firstRef = Json.GetString(first, "documentRef");
        Assert.False(string.IsNullOrWhiteSpace(firstRef));

        var second = host.HwpLaunch(new JsonObject { ["newDocument"] = false });
        Assert.True(Json.GetBool(second, "ok"), $"second hwp_launch failed: {second}");
        Assert.False(Json.GetBool(Json.GetObj(second, "summary"), "createdDocument"));
        Assert.Equal(firstRef, Json.GetString(second, "documentRef"));
    }

    [Fact]
    public void Hwp_multi_document_inventory_and_document_ref_targeting()
    {
        if (!Enabled) return;
        using var host = CreateHostWithHwp();

        var firstContext = host.GetActiveContext("hwp");
        Assert.True(Json.GetBool(firstContext, "ok"), $"first context failed: {firstContext}");
        var firstRef = Json.GetString(firstContext, "documentRef");
        Assert.False(string.IsNullOrWhiteSpace(firstRef));

        var launch = host.HwpLaunch(new JsonObject { ["newDocument"] = true });
        Assert.True(Json.GetBool(launch, "ok"), $"second document launch failed: {launch}");
        var secondRef = Json.GetString(launch, "documentRef");
        Assert.False(string.IsNullOrWhiteSpace(secondRef));
        Assert.NotEqual(firstRef, secondRef);

        var context = host.GetActiveContext("hwp");
        Assert.True(Json.GetBool(context, "ok"), $"multi context failed: {context}");
        var summary = Json.GetObj(context, "summary")!;
        var openDocuments = Json.GetArr(summary, "openDocuments")!;
        Assert.True(openDocuments.Count >= 2, $"expected at least two HWP tabs: {context}");
        var refs = openDocuments
            .Select(node => Json.GetString(node as JsonObject, "documentRef"))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
        Assert.Contains(firstRef, refs);
        Assert.Contains(secondRef, refs);
        Assert.Equal(openDocuments.Count, refs.Count);

        var firstRead = host.Read("hwp", new JsonObject
        {
            ["documentRef"] = firstRef,
            ["scope"] = "document",
        });
        Assert.True(Json.GetBool(firstRead, "ok"), $"first targeted read failed: {firstRead}");
        Assert.Equal(firstRef, Json.GetString(firstRead, "documentRef"));
        Assert.Contains("사과 가격", Json.GetString(firstRead, "text"));

        const string marker = "DOCBRIDGE_MULTI_DOCUMENT_TARGET";
        var operations = new JsonArray
        {
            new JsonObject
            {
                ["op"] = "insert_text",
                ["documentRef"] = secondRef,
                ["text"] = marker,
            },
        };
        var dry = host.ApplyOps("hwp", new JsonObject
        {
            ["ops"] = operations.DeepClone(),
            ["dryRun"] = true,
        });
        Assert.True(Json.GetBool(dry, "ok"), $"targeted dry-run failed: {dry}");
        var applied = host.ApplyOps("hwp", new JsonObject
        {
            ["ops"] = operations.DeepClone(),
            ["dryRun"] = false,
            ["confirmToken"] = Json.GetString(dry, "confirmToken"),
        });
        Assert.True(Json.GetBool(applied, "ok"), $"targeted apply failed: {applied}");
        Assert.True(Json.GetBool(Json.GetObj(applied, "readback"), "verified"));

        var secondRead = host.Read("hwp", new JsonObject
        {
            ["documentRef"] = secondRef,
            ["scope"] = "document",
        });
        Assert.True(Json.GetBool(secondRead, "ok"), $"second targeted read failed: {secondRead}");
        Assert.Contains(marker, Json.GetString(secondRead, "text"));
        var firstReadAfter = host.Read("hwp", new JsonObject
        {
            ["documentRef"] = firstRef,
            ["scope"] = "document",
        });
        Assert.DoesNotContain(marker, Json.GetString(firstReadAfter, "text"));
    }

    [Fact]
    public void Hwp_production_format_table_page_and_pdf_flow()
    {
        if (!Enabled) return;
        using var host = CreateHostWithHwp();

        JsonObject Apply(JsonArray operations, bool highRisk = false)
        {
            var dry = host.ApplyOps("hwp", new JsonObject { ["ops"] = operations.DeepClone(), ["dryRun"] = true });
            Assert.True(Json.GetBool(dry, "ok"), $"dry-run failed: {dry}");
            var result = host.ApplyOps("hwp", new JsonObject
            {
                ["ops"] = operations.DeepClone(),
                ["dryRun"] = false,
                ["confirmToken"] = Json.GetString(dry, "confirmToken"),
                ["highRiskConfirm"] = highRisk,
            });
            Assert.True(Json.GetBool(result, "ok"), $"apply failed: {result}");
            Assert.True(Json.GetBool(Json.GetObj(result, "readback"), "verified"));
            return result;
        }

        Apply(new JsonArray
        {
            new JsonObject
            {
                ["op"] = "set_paragraph_style_basic",
                ["target"] = new JsonObject { ["text"] = "사과 가격" },
                ["style"] = new JsonObject
                {
                    ["fontName"] = "맑은 고딕", ["fontSize"] = 12,
                    ["bold"] = true, ["textColor"] = "#1F4E78", ["letterSpacing"] = 2,
                },
            },
            new JsonObject
            {
                ["op"] = "set_paragraph_format",
                ["target"] = new JsonObject { ["scope"] = "document" },
                ["style"] = new JsonObject
                {
                    ["align"] = "left", ["lineSpacingPercent"] = 160,
                    ["spaceAfterPt"] = 4, ["widowOrphan"] = true,
                },
            },
            new JsonObject
            {
                ["op"] = "format_paragraphs",
                ["items"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["target"] = new JsonObject { ["text"] = "재고 확인" },
                        ["characterStyle"] = new JsonObject { ["bold"] = true, ["textColor"] = "#C00000" },
                        ["paragraphStyle"] = new JsonObject { ["lineSpacingPercent"] = 160 },
                    },
                },
            },
            new JsonObject
            {
                ["op"] = "set_page_setup",
                ["applyTo"] = "document",
                ["page"] = new JsonObject
                {
                    ["widthMm"] = 210, ["heightMm"] = 297, ["orientation"] = "portrait",
                    ["leftMarginMm"] = 20, ["rightMarginMm"] = 20,
                    ["topMarginMm"] = 15, ["bottomMarginMm"] = 15,
                },
            },
        });

        Apply(new JsonArray
        {
            new JsonObject
            {
                ["op"] = "insert_table",
                ["rows"] = new JsonArray
                {
                    new JsonArray("구분", "담당", "상태"),
                    new JsonArray("안전점검", "홍길동", "예정"),
                },
                ["header"] = true,
                ["headerFill"] = "#D9EAF7",
                ["cellStyles"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["row"] = 1, ["col"] = 2,
                        ["fontName"] = "맑은 고딕", ["fontSize"] = 13,
                        ["align"] = "center",
                    },
                },
            },
        });

        var beforeTableRead = host.Read("hwp", new JsonObject
        {
            ["scope"] = "tables", ["tableIndex"] = 0,
            ["maxCells"] = 20, ["includeStyles"] = true,
        });
        Assert.True(Json.GetBool(beforeTableRead, "ok"), $"pre-write table/style read failed: {beforeTableRead}");
        var beforeCells = Json.GetArr(
            Json.GetArr(Json.GetObj(beforeTableRead, "tableInventory"), "tables")![0] as JsonObject, "cells")!;
        var planned = beforeCells.Select(node => node as JsonObject)
            .First(cell => Json.GetString(cell, "text") == "예정")!;
        var plannedCharacter = Json.GetObj(Json.GetObj(planned, "style"), "character")!;
        var plannedParagraph = Json.GetObj(Json.GetObj(planned, "style"), "paragraph")!;

        Apply(new JsonArray
        {
            new JsonObject
            {
                ["op"] = "table_cell_set_text", ["tableIndex"] = 0,
                ["row"] = 1, ["col"] = 2, ["text"] = "완료",
            },
            new JsonObject
            {
                ["op"] = "table_set_row_heights", ["tableIndex"] = 0,
                ["rows"] = new JsonArray
                {
                    new JsonObject { ["row"] = 0, ["heightMm"] = 8.0 },
                    new JsonObject { ["row"] = 1, ["heightMm"] = 10.0 },
                },
            },
            new JsonObject { ["op"] = "insert_page_number", ["position"] = "bottom-center", ["format"] = "arabic" },
        });

        var tableRead = host.Read("hwp", new JsonObject
        {
            ["scope"] = "tables", ["tableIndex"] = 0,
            ["maxCells"] = 20, ["includeStyles"] = true,
        });
        Assert.True(Json.GetBool(tableRead, "ok"), $"table/style read failed: {tableRead}");
        var tables = Json.GetArr(Json.GetObj(tableRead, "tableInventory"), "tables")!;
        var cells = Json.GetArr(tables[0] as JsonObject, "cells")!;
        var completed = cells.Select(node => node as JsonObject)
            .First(cell => Json.GetString(cell, "text") == "완료")!;
        var completedStyle = Json.GetObj(completed, "style")!;
        var completedCharacter = Json.GetObj(completedStyle, "character")!;
        var completedParagraph = Json.GetObj(completedStyle, "paragraph")!;
        Assert.Equal(Json.GetString(plannedCharacter, "fontName"), Json.GetString(completedCharacter, "fontName"));
        Assert.Equal(plannedCharacter["fontSizePt"]!.GetValue<double>(), completedCharacter["fontSizePt"]!.GetValue<double>());
        Assert.Equal(Json.GetInt(plannedParagraph, "alignType"), Json.GetInt(completedParagraph, "alignType"));

        var png = Path.Combine(_home.Dir, "one-pixel.png");
        File.WriteAllBytes(png, Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="));
        Apply(new JsonArray
        {
            new JsonObject { ["op"] = "insert_break", ["type"] = "paragraph" },
            new JsonObject
            {
                ["op"] = "insert_picture", ["path"] = png, ["embedded"] = true,
                ["sizeOption"] = "specific", ["widthMm"] = 10, ["heightMm"] = 10,
            },
            new JsonObject
            {
                ["op"] = "set_header_footer_text", ["kind"] = "footer",
                ["pages"] = "both", ["text"] = "DocBridge 통합 테스트",
            },
        });

        var structure = host.Read("hwp", new JsonObject { ["scope"] = "structure", ["maxControls"] = 100 });
        Assert.True(Json.GetBool(structure, "ok"), $"structure read failed: {structure}");
        var counts = Json.GetObj(Json.GetObj(structure, "structure"), "countsByControlId");
        Assert.True(Json.GetInt(counts, "tbl") >= 1);
        Assert.True((Json.GetInt(counts, "$pic") ?? 0) + (Json.GetInt(counts, "gso") ?? 0) >= 1);
        Assert.True(Json.GetInt(counts, "foot") >= 1);
        Assert.Contains("완료", Json.GetString(host.Read("hwp", new JsonObject { ["scope"] = "document" }), "text"));

        var pdf = Path.Combine(_home.Dir, "hwp-production-e2e.pdf");
        Apply(new JsonArray { new JsonObject { ["op"] = "export_pdf", ["output"] = pdf } }, highRisk: true);
        Assert.True(File.Exists(pdf));
        Assert.True(new FileInfo(pdf).Length > 0);
        if (Environment.GetEnvironmentVariable("DOCBRIDGE_E2E_ARTIFACTS") is { Length: > 0 } artifactDir)
        {
            Directory.CreateDirectory(artifactDir);
            File.Copy(pdf, Path.Combine(artifactDir, "hwp-production-e2e.pdf"), overwrite: true);
        }
    }

    [Fact]
    public void Hwp_table_structure_flow()
    {
        if (!Enabled) return;
        if (!int.TryParse(Environment.GetEnvironmentVariable("DOCBRIDGE_HWP_TABLE_STAGE"), out var stage) || stage < 1) return;
        using var host = CreateHostWithHwp();

        void Apply(JsonArray operations, bool highRisk = false)
        {
            var dry = host.ApplyOps("hwp", new JsonObject { ["ops"] = operations.DeepClone(), ["dryRun"] = true });
            Assert.True(Json.GetBool(dry, "ok"), $"dry-run failed: {dry}");
            var result = host.ApplyOps("hwp", new JsonObject
            {
                ["ops"] = operations.DeepClone(), ["dryRun"] = false,
                ["confirmToken"] = Json.GetString(dry, "confirmToken"), ["highRiskConfirm"] = highRisk,
            });
            Assert.True(Json.GetBool(result, "ok"), $"apply failed: {result}");
        }

        Apply(new JsonArray { new JsonObject
        {
            ["op"] = "insert_table",
            ["rows"] = new JsonArray { new JsonArray("A", "B"), new JsonArray("C", "D") },
        }});
        Apply(new JsonArray { new JsonObject
        {
            ["op"] = "table_insert_rows", ["tableIndex"] = 0, ["row"] = 1, ["col"] = 0,
            ["count"] = 1, ["position"] = "after",
        }});
        if (stage == 1) return;
        Apply(new JsonArray { new JsonObject
        {
            ["op"] = "table_cell_set_text", ["tableIndex"] = 0, ["row"] = 2, ["col"] = 0, ["text"] = "ROW-SENTINEL",
        }});
        Assert.Contains("ROW-SENTINEL", Json.GetString(host.Read("hwp", new JsonObject { ["scope"] = "document" }), "text"));
        if (stage == 2) return;
        Apply(new JsonArray { new JsonObject
        {
            ["op"] = "table_delete_rows", ["tableIndex"] = 0, ["row"] = 2, ["col"] = 0,
        }}, highRisk: true);
        Assert.DoesNotContain("ROW-SENTINEL", Json.GetString(host.Read("hwp", new JsonObject { ["scope"] = "document" }), "text"));
        if (stage == 3) return;
        Apply(new JsonArray { new JsonObject
        {
            ["op"] = "table_insert_columns", ["tableIndex"] = 0, ["row"] = 0, ["col"] = 1,
            ["count"] = 1, ["position"] = "after",
        }});
        Apply(new JsonArray { new JsonObject
        {
            ["op"] = "table_cell_set_text", ["tableIndex"] = 0, ["row"] = 0, ["col"] = 2, ["text"] = "COL-SENTINEL",
        }});
        Assert.Contains("COL-SENTINEL", Json.GetString(host.Read("hwp", new JsonObject { ["scope"] = "document" }), "text"));
        if (stage == 4) return;
        Apply(new JsonArray { new JsonObject
        {
            ["op"] = "table_delete_columns", ["tableIndex"] = 0, ["row"] = 0, ["col"] = 2,
        }}, highRisk: true);
        Assert.DoesNotContain("COL-SENTINEL", Json.GetString(host.Read("hwp", new JsonObject { ["scope"] = "document" }), "text"));
        if (stage == 5) return;
        Apply(new JsonArray { new JsonObject
        {
            ["op"] = "table_merge_cells", ["tableIndex"] = 0,
            ["startRow"] = 0, ["startCol"] = 0, ["endRow"] = 0, ["endCol"] = 1,
        }});
        return;
    }

    [Fact]
    public void Hwp_table_insert_and_delete_honor_count_exactly()
    {
        if (!Enabled || Environment.GetEnvironmentVariable("DOCBRIDGE_HWP_TABLE_COUNT_E2E") != "1") return;
        using var host = CreateHostWithHwp();

        void Apply(JsonArray operations, bool highRisk = false)
        {
            var dry = host.ApplyOps("hwp", new JsonObject { ["ops"] = operations.DeepClone(), ["dryRun"] = true });
            Assert.True(Json.GetBool(dry, "ok"), $"dry-run failed: {dry}");
            var result = host.ApplyOps("hwp", new JsonObject
            {
                ["ops"] = operations.DeepClone(), ["dryRun"] = false,
                ["confirmToken"] = Json.GetString(dry, "confirmToken"), ["highRiskConfirm"] = highRisk,
            });
            Assert.True(Json.GetBool(result, "ok"), $"apply failed: {result}");
        }

        Apply(new JsonArray { new JsonObject
        {
            ["op"] = "insert_table",
            ["rows"] = new JsonArray { new JsonArray("A", "B"), new JsonArray("C", "D") },
        }});
        Apply(new JsonArray { new JsonObject
        {
            ["op"] = "table_insert_rows", ["tableIndex"] = 0, ["row"] = 1, ["col"] = 0,
            ["count"] = 4, ["position"] = "after",
        }});
        Apply(new JsonArray
        {
            new JsonObject
            {
                ["op"] = "table_cell_set_text", ["tableIndex"] = 0,
                ["row"] = 4, ["col"] = 0, ["text"] = "DELETE-ME-ROW-4",
            },
            new JsonObject
            {
                ["op"] = "table_cell_set_text", ["tableIndex"] = 0,
                ["row"] = 5, ["col"] = 0, ["text"] = "KEEP-ROW-5",
            },
        });
        var afterInsert = Json.GetString(host.Read("hwp", new JsonObject { ["scope"] = "document" }), "text");
        Assert.Contains("DELETE-ME-ROW-4", afterInsert);
        Assert.Contains("KEEP-ROW-5", afterInsert);

        Apply(new JsonArray { new JsonObject
        {
            ["op"] = "table_delete_rows", ["tableIndex"] = 0, ["row"] = 1, ["col"] = 0,
            ["count"] = 4,
        }}, highRisk: true);
        var afterDelete = Json.GetString(host.Read("hwp", new JsonObject { ["scope"] = "document" }), "text");
        Assert.DoesNotContain("DELETE-ME-ROW-4", afterDelete);
        Assert.Contains("KEEP-ROW-5", afterDelete);
    }

    [Fact]
    public void Hwp_relative_insert_accepts_minus_and_millimeter_unicode_readback()
    {
        if (!Enabled || Environment.GetEnvironmentVariable("DOCBRIDGE_HWP_UNICODE_E2E") != "1") return;
        using var host = CreateHostWithHwp();

        JsonObject Apply(JsonArray operations)
        {
            var dry = host.ApplyOps("hwp", new JsonObject { ["ops"] = operations.DeepClone(), ["dryRun"] = true });
            Assert.True(Json.GetBool(dry, "ok"), $"dry-run failed: {dry}");
            var result = host.ApplyOps("hwp", new JsonObject
            {
                ["ops"] = operations.DeepClone(), ["dryRun"] = false,
                ["confirmToken"] = Json.GetString(dry, "confirmToken"),
            });
            Assert.True(Json.GetBool(result, "ok"), $"apply failed: {result}");
            return result;
        }

        Apply(new JsonArray { new JsonObject
        {
            ["op"] = "replace_document_text", ["text"] = "제목\n기준 문단\n끝 문단",
        }});
        var longParagraph = string.Join(" · ", Enumerable.Range(1, 12)
            .Select(index => $"구간 {index}: 관경 − {index * 100}㎜, 연장 {index * 125}㎜"));
        Apply(new JsonArray { new JsonObject
        {
            ["op"] = "insert_before_text", ["anchor"] = "기준 문단",
            ["text"] = longParagraph, ["mode"] = "paragraph", ["matchCase"] = true,
        }});

        var after = Json.GetString(host.Read("hwp", new JsonObject { ["scope"] = "document" }), "text") ?? "";
        Assert.True(HwpAdapter.HwpReadbackContainsEquivalent(after, longParagraph), after);
        Assert.Contains("기준 문단", after);
    }

    [Fact]
    public void Hwp_picture_can_be_inserted_into_an_exact_table_cell()
    {
        if (!Enabled || Environment.GetEnvironmentVariable("DOCBRIDGE_HWP_TABLE_PICTURE_E2E") != "1") return;
        using var host = CreateHostWithHwp();
        var picture = Path.GetFullPath(Path.Combine(AppContext.BaseDirectory,
            "..", "..", "..", "..", "..", "reports", "hwp-e2e", "rendered", "page-1.png"));
        Assert.True(File.Exists(picture), picture);

        JsonObject Apply(JsonArray operations)
        {
            var dry = host.ApplyOps("hwp", new JsonObject { ["ops"] = operations.DeepClone(), ["dryRun"] = true });
            Assert.True(Json.GetBool(dry, "ok"), $"dry-run failed: {dry}");
            var result = host.ApplyOps("hwp", new JsonObject
            {
                ["ops"] = operations.DeepClone(), ["dryRun"] = false,
                ["confirmToken"] = Json.GetString(dry, "confirmToken"),
            });
            Assert.True(Json.GetBool(result, "ok"), $"apply failed: {result}");
            return result;
        }

        Apply(new JsonArray { new JsonObject
        {
            ["op"] = "insert_table",
            ["rows"] = new JsonArray { new JsonArray("사진", "설명"), new JsonArray("", "현장 사진") },
        }});
        var result = Apply(new JsonArray { new JsonObject
        {
            ["op"] = "insert_picture", ["path"] = picture,
            ["tableIndex"] = 0, ["row"] = 1, ["col"] = 0,
            ["sizeOption"] = "cell-ratio", ["embedded"] = true, ["clearCell"] = false,
        }});
        Assert.Contains(result["affected"]!.AsArray(), item =>
            Json.GetString(item as JsonObject, "type") == "table:0/cell:1,0/picture");

        var structure = host.Read("hwp", new JsonObject { ["scope"] = "structure" });
        var counts = Json.GetObj(Json.GetObj(structure, "structure"), "countsByControlId");
        Assert.True((Json.GetInt(counts, "$pic") ?? 0) + (Json.GetInt(counts, "gso") ?? 0) >= 1,
            structure.ToJsonString());
    }

    [Fact]
    public void Hwp_dry_run_simulates_text_and_table_operations_in_order()
    {
        if (!Enabled || Environment.GetEnvironmentVariable("DOCBRIDGE_HWP_SEQUENTIAL_PREVIEW_E2E") != "1") return;
        using var host = CreateHostWithHwp();
        var operations = new JsonArray
        {
            new JsonObject
            {
                ["op"] = "replace_document_text", ["text"] = "순차 기준 문단",
            },
            new JsonObject
            {
                ["op"] = "insert_after_text", ["anchor"] = "순차 기준 문단",
                ["text"] = "앞 op가 만든 기준을 찾은 문단", ["mode"] = "paragraph",
            },
            new JsonObject
            {
                ["op"] = "insert_table",
                ["rows"] = new JsonArray { new JsonArray("항목", "내용"), new JsonArray("검증", "순차 표") },
            },
            new JsonObject
            {
                ["op"] = "table_set_row_height", ["tableIndex"] = 0,
                ["row"] = 0, ["heightMm"] = 8.0,
            },
        };

        var dry = host.ApplyOps("hwp", new JsonObject { ["ops"] = operations.DeepClone(), ["dryRun"] = true });
        Assert.True(Json.GetBool(dry, "ok"), $"sequential dry-run failed: {dry}");
        Assert.Empty(Json.GetArr(dry, "errors") ?? new JsonArray());
        Assert.Equal(4, Json.GetArr(dry, "affected")?.Count);
        Assert.False(Json.GetBool(dry, "requiresHighRiskApproval"));
    }

    [Fact]
    public void Hwp_blank_repeated_form_cell_inherits_matching_label_style()
    {
        if (!Enabled) return;
        using var host = CreateHostWithHwp();

        JsonObject Apply(JsonArray operations)
        {
            var dry = host.ApplyOps("hwp", new JsonObject { ["ops"] = operations.DeepClone(), ["dryRun"] = true });
            Assert.True(Json.GetBool(dry, "ok"), $"dry-run failed: {dry}");
            var result = host.ApplyOps("hwp", new JsonObject
            {
                ["ops"] = operations.DeepClone(), ["dryRun"] = false,
                ["confirmToken"] = Json.GetString(dry, "confirmToken"),
            });
            Assert.True(Json.GetBool(result, "ok"), $"apply failed: {result}");
            return result;
        }

        Apply(new JsonArray
        {
            new JsonObject
            {
                ["op"] = "insert_break", ["type"] = "paragraph",
            },
            new JsonObject
            {
                ["op"] = "insert_table",
                ["rows"] = new JsonArray { new JsonArray("담당자", "홍길동") },
                ["cellStyles"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["row"] = 0, ["col"] = 1,
                        ["fontName"] = "맑은 고딕", ["fontSize"] = 12,
                        ["align"] = "center",
                    },
                },
            },
        });
        Apply(new JsonArray
        {
            new JsonObject { ["op"] = "insert_break", ["type"] = "paragraph" },
            new JsonObject
            {
                ["op"] = "insert_table",
                ["rows"] = new JsonArray { new JsonArray("담당자", "") },
            },
        });

        JsonObject ReadCell(int tableIndex, int cellIndex)
        {
            var read = host.Read("hwp", new JsonObject
            {
                ["scope"] = "tables", ["tableIndex"] = tableIndex,
                ["maxCells"] = 10, ["includeStyles"] = true,
            });
            Assert.True(Json.GetBool(read, "ok"), $"table read failed: {read}");
            var tables = Json.GetArr(Json.GetObj(read, "tableInventory"), "tables")!;
            return (Json.GetArr(tables[0] as JsonObject, "cells")![cellIndex] as JsonObject)!;
        }

        var source = ReadCell(0, 1);
        var sourceStyle = Json.GetObj(source, "style")!;
        var write = Apply(new JsonArray
        {
            new JsonObject
            {
                ["op"] = "table_cell_set_text", ["tableIndex"] = 1,
                ["cellIndex"] = 1, ["text"] = "김코덱스",
            },
        });
        Assert.Contains("repeated-label", write.ToJsonString());

        var target = ReadCell(1, 1);
        Assert.Equal("김코덱스", Json.GetString(target, "text"));
        var targetStyle = Json.GetObj(target, "style")!;
        var sourceCharacter = Json.GetObj(sourceStyle, "character")!;
        var targetCharacter = Json.GetObj(targetStyle, "character")!;
        var sourceParagraph = Json.GetObj(sourceStyle, "paragraph")!;
        var targetParagraph = Json.GetObj(targetStyle, "paragraph")!;
        Assert.Equal(Json.GetString(sourceCharacter, "fontName"), Json.GetString(targetCharacter, "fontName"));
        Assert.Equal(sourceCharacter["fontSizePt"]!.GetValue<double>(), targetCharacter["fontSizePt"]!.GetValue<double>());
        Assert.Equal(Json.GetInt(sourceParagraph, "alignType"), Json.GetInt(targetParagraph, "alignType"));
    }

    [Fact]
    public void Hwp_issue_1_unicode_scoped_replace_large_table_inventory_and_count_contracts()
    {
        if (!Enabled || Environment.GetEnvironmentVariable("DOCBRIDGE_HWP_ISSUE1_E2E") != "1") return;
        using var host = CreateHostWithHwp();

        JsonObject Apply(JsonArray operations, bool highRisk = false)
        {
            var dry = host.ApplyOps("hwp", new JsonObject { ["ops"] = operations.DeepClone(), ["dryRun"] = true });
            Assert.True(Json.GetBool(dry, "ok"), $"dry-run failed: {dry}");
            var result = host.ApplyOps("hwp", new JsonObject
            {
                ["ops"] = operations.DeepClone(), ["dryRun"] = false,
                ["confirmToken"] = Json.GetString(dry, "confirmToken"),
                ["highRiskConfirm"] = highRisk,
            });
            Assert.True(Json.GetBool(result, "ok"), $"apply failed: {result}");
            return result;
        }

        Apply(new JsonArray(new JsonObject
        {
            ["op"] = "replace_document_text", ["text"] = "관경 − 300㎜\nA A\nA",
        }));
        var rawRead = host.Read("hwp", new JsonObject { ["scope"] = "document" });
        Assert.Contains("−", Json.GetString(rawRead, "text"));
        Assert.Contains("㎜", Json.GetString(rawRead, "text"));
        Assert.DoesNotContain("&#8722;", Json.GetString(rawRead, "text"));

        Apply(new JsonArray(new JsonObject
        {
            ["op"] = "find_replace", ["find"] = "&#8722;", ["replace"] = "-",
            ["occurrence"] = 1,
        }));
        Apply(new JsonArray(new JsonObject
        {
            ["op"] = "find_replace", ["find"] = "A", ["replace"] = "B",
            ["occurrence"] = 2,
        }));
        Apply(new JsonArray(new JsonObject
        {
            ["op"] = "find_replace", ["find"] = "A", ["replace"] = "C",
            ["scope"] = new JsonObject { ["startParagraph"] = 2, ["endParagraph"] = 2 },
        }));
        var scopedRead = Json.GetString(host.Read("hwp", new JsonObject { ["scope"] = "document" }), "text")!;
        Assert.Contains("관경 - 300㎜", scopedRead);
        Assert.Contains("A B\nC", scopedRead.Replace("\r\n", "\n"));

        var largeRows = new JsonArray();
        for (var row = 0; row < 9; row++)
        {
            var cells = new JsonArray();
            for (var col = 0; col < 10; col++) cells.Add($"L{row:00}-{col:00}");
            largeRows.Add(cells);
        }
        Apply(new JsonArray(
            new JsonObject { ["op"] = "insert_break", ["type"] = "paragraph" },
            new JsonObject { ["op"] = "insert_table", ["rows"] = largeRows },
            new JsonObject
            {
                ["op"] = "table_insert_columns", ["tableIndex"] = 0,
                ["row"] = 0, ["col"] = 9, ["count"] = 20, ["position"] = "after",
            },
            new JsonObject
            {
                ["op"] = "table_insert_columns", ["tableIndex"] = 0,
                ["row"] = 0, ["col"] = 29, ["count"] = 14, ["position"] = "after",
            },
            new JsonObject { ["op"] = "insert_break", ["type"] = "paragraph" },
            new JsonObject
            {
                ["op"] = "insert_table",
                ["rows"] = new JsonArray(new JsonArray("사진", "설명"), new JsonArray("", "현장")),
            }));

        var inventory = host.Read("hwp", new JsonObject
        {
            ["scope"] = "tables", ["maxCells"] = 500, ["includeStyles"] = false,
        });
        var tables = Json.GetArr(Json.GetObj(inventory, "tableInventory"), "tables")!;
        Assert.Equal(2, tables.Count);
        Assert.Equal(396, Json.GetInt(tables[0] as JsonObject, "cellCountRead"));
        Assert.Equal(4, Json.GetInt(tables[1] as JsonObject, "cellCountRead"));
        Assert.NotEqual(Json.GetString(tables[0] as JsonObject, "controlRef"),
            Json.GetString(tables[1] as JsonObject, "controlRef"));

        Apply(new JsonArray(new JsonObject
        {
            ["op"] = "table_set_cells", ["tableIndex"] = 1, ["preserveStyle"] = false,
            ["cells"] = new JsonArray(
                new JsonObject { ["row"] = 0, ["col"] = 0, ["text"] = "사진번호" },
                new JsonObject { ["row"] = 0, ["col"] = 1, ["text"] = "설명내용" },
                new JsonObject { ["row"] = 1, ["col"] = 0, ["text"] = "P-01" },
                new JsonObject { ["row"] = 1, ["col"] = 1, ["text"] = "검증완료" }),
        }));
        Apply(new JsonArray(new JsonObject
        {
            ["op"] = "table_insert_rows", ["tableIndex"] = 1,
            ["row"] = 1, ["col"] = 0, ["count"] = 4, ["position"] = "after",
        }));
        var afterInsert = host.Read("hwp", new JsonObject
        {
            ["scope"] = "tables", ["tableIndex"] = 1, ["maxCells"] = 100, ["includeStyles"] = false,
        });
        Assert.Equal(12, Json.GetInt(Json.GetArr(Json.GetObj(afterInsert, "tableInventory"), "tables")![0] as JsonObject, "cellCountRead"));
        Apply(new JsonArray(new JsonObject
        {
            ["op"] = "table_delete_rows", ["tableIndex"] = 1,
            ["row"] = 2, ["col"] = 0, ["count"] = 4,
        }), highRisk: true);
        var afterDelete = host.Read("hwp", new JsonObject
        {
            ["scope"] = "tables", ["tableIndex"] = 1, ["maxCells"] = 100, ["includeStyles"] = false,
        });
        Assert.Equal(4, Json.GetInt(Json.GetArr(Json.GetObj(afterDelete, "tableInventory"), "tables")![0] as JsonObject, "cellCountRead"));
    }

    [Fact]
    public void Hwp_blank_cell_uses_above_and_below_same_role_consensus()
    {
        if (!Enabled) return;
        using var host = CreateHostWithHwp();

        JsonObject Apply(JsonArray operations)
        {
            var dry = host.ApplyOps("hwp", new JsonObject { ["ops"] = operations.DeepClone(), ["dryRun"] = true });
            Assert.True(Json.GetBool(dry, "ok"), $"dry-run failed: {dry}");
            var result = host.ApplyOps("hwp", new JsonObject
            {
                ["ops"] = operations.DeepClone(), ["dryRun"] = false,
                ["confirmToken"] = Json.GetString(dry, "confirmToken"),
            });
            Assert.True(Json.GetBool(result, "ok"), $"apply failed: {result}");
            return result;
        }

        Apply(new JsonArray
        {
            new JsonObject
            {
                ["op"] = "insert_table",
                ["rows"] = new JsonArray
                {
                    new JsonArray("항목 1", "위 값"),
                    new JsonArray("항목 2", ""),
                    new JsonArray("항목 3", "아래 값"),
                },
                ["cellStyles"] = new JsonArray
                {
                    new JsonObject
                    {
                        ["row"] = 0, ["col"] = 1,
                        ["fontName"] = "맑은 고딕", ["fontSize"] = 13,
                        ["bold"] = true, ["align"] = "center",
                    },
                    new JsonObject
                    {
                        ["row"] = 2, ["col"] = 1,
                        ["fontName"] = "맑은 고딕", ["fontSize"] = 13,
                        ["bold"] = true, ["align"] = "center",
                    },
                },
            },
        });

        var write = Apply(new JsonArray
        {
            new JsonObject
            {
                ["op"] = "table_cell_set_text", ["tableIndex"] = 0,
                ["cellIndex"] = 3, ["text"] = "중간 값",
            },
        });
        Assert.Contains("surrounding-same-role", write.ToJsonString());
        Assert.Contains("consensus", write.ToJsonString());

        var read = host.Read("hwp", new JsonObject
        {
            ["scope"] = "tables", ["tableIndex"] = 0,
            ["maxCells"] = 10, ["includeStyles"] = true,
        });
        Assert.True(Json.GetBool(read, "ok"), $"table read failed: {read}");
        var tables = Json.GetArr(Json.GetObj(read, "tableInventory"), "tables")!;
        var cells = Json.GetArr(tables[0] as JsonObject, "cells")!;
        var top = Json.GetObj(cells[1] as JsonObject, "style")!;
        var middle = Json.GetObj(cells[3] as JsonObject, "style")!;
        Assert.Equal("중간 값", Json.GetString(cells[3] as JsonObject, "text"));
        Assert.Equal(
            Json.GetObj(top, "character")!["fontSizePt"]!.GetValue<double>(),
            Json.GetObj(middle, "character")!["fontSizePt"]!.GetValue<double>());
        Assert.Equal(
            Json.GetInt(Json.GetObj(top, "paragraph"), "alignType"),
            Json.GetInt(Json.GetObj(middle, "paragraph"), "alignType"));
    }

    [Fact]
    public void Hwp_existing_form_middle_insert_preserves_structure_and_reopens()
    {
        if (!Enabled) return;
        var sampleFile = Environment.GetEnvironmentVariable("DOCBRIDGE_HWP_SAMPLE_FILE");
        if (string.IsNullOrWhiteSpace(sampleFile)) return;
        Assert.True(File.Exists(sampleFile), $"sample HWP not found: {sampleFile}");

        using var adapter = new HwpAdapter();
        using var host = new DocBridgeHost(_home.Options);
        host.Router.Register("hwp", adapter);

        var beforeStructure = host.Read("hwp", new JsonObject
        {
            ["file"] = sampleFile,
            ["scope"] = "structure",
            ["maxControls"] = 1000,
        });
        Assert.True(Json.GetBool(beforeStructure, "ok"), $"structure read failed: {beforeStructure}");
        var beforeCounts = Json.GetObj(Json.GetObj(beforeStructure, "structure"), "countsByControlId")!;

        const string anchor = "관련근거 : 현장확인";
        const string inserted = "[DocBridge 중간 삽입 검증] 검토 결과: 추가 보완사항 없음";
        var operations = new JsonArray
        {
            new JsonObject
            {
                ["op"] = "insert_after_text",
                ["file"] = sampleFile,
                ["anchor"] = anchor,
                ["text"] = inserted,
                ["occurrence"] = 1,
                ["matchCase"] = true,
                ["mode"] = "paragraph",
            },
        };

        var dry = host.ApplyOps("hwp", new JsonObject
        {
            ["ops"] = operations.DeepClone(),
            ["dryRun"] = true,
        });
        Assert.True(Json.GetBool(dry, "ok"), $"dry-run failed: {dry}");
        Assert.Contains("inherit adjacent style", Json.GetString(Json.GetArr(dry, "affected")![0] as JsonObject, "ref"));

        var applied = host.ApplyOps("hwp", new JsonObject
        {
            ["ops"] = operations.DeepClone(),
            ["dryRun"] = false,
            ["confirmToken"] = Json.GetString(dry, "confirmToken"),
        });
        Assert.True(Json.GetBool(applied, "ok"), $"apply failed: {applied}");
        Assert.True(Json.GetBool(Json.GetObj(applied, "readback"), "verified"), $"readback failed: {applied}");

        // 파일을 닫은 뒤 다시 열어 실제 저장 결과와 기준 문구 직후 위치를 검증한다.
        var reopened = host.Read("hwp", new JsonObject
        {
            ["file"] = sampleFile,
            ["scope"] = "document",
            ["maxChars"] = 20000,
        });
        Assert.True(Json.GetBool(reopened, "ok"), $"reopen failed: {reopened}");
        var reopenedText = Json.GetString(reopened, "text")!;
        var anchorIndex = reopenedText.IndexOf(anchor, StringComparison.Ordinal);
        var insertedIndex = reopenedText.IndexOf(inserted, StringComparison.Ordinal);
        Assert.True(anchorIndex >= 0 && insertedIndex > anchorIndex,
            $"inserted paragraph is not after the requested anchor: anchor={anchorIndex}, inserted={insertedIndex}");
        Assert.True(insertedIndex - anchorIndex < 120,
            $"inserted paragraph is too far from the requested anchor: distance={insertedIndex - anchorIndex}");

        var afterStructure = host.Read("hwp", new JsonObject
        {
            ["file"] = sampleFile,
            ["scope"] = "structure",
            ["maxControls"] = 1000,
        });
        Assert.True(Json.GetBool(afterStructure, "ok"), $"post-structure read failed: {afterStructure}");
        var afterCounts = Json.GetObj(Json.GetObj(afterStructure, "structure"), "countsByControlId")!;
        Assert.Equal(Json.GetInt(beforeCounts, "tbl"), Json.GetInt(afterCounts, "tbl"));
        Assert.Equal(Json.GetInt(beforeCounts, "gso"), Json.GetInt(afterCounts, "gso"));
        Assert.Equal(Json.GetInt(beforeCounts, "$pic"), Json.GetInt(afterCounts, "$pic"));

        var samplePdf = Environment.GetEnvironmentVariable("DOCBRIDGE_HWP_SAMPLE_PDF");
        if (!string.IsNullOrWhiteSpace(samplePdf))
        {
            var pdfOps = new JsonArray
            {
                new JsonObject
                {
                    ["op"] = "export_pdf",
                    ["file"] = sampleFile,
                    ["output"] = samplePdf,
                },
            };
            var pdfDry = host.ApplyOps("hwp", new JsonObject
            {
                ["ops"] = pdfOps.DeepClone(),
                ["dryRun"] = true,
            });
            Assert.True(Json.GetBool(pdfDry, "ok"), $"pdf dry-run failed: {pdfDry}");
            var pdfApply = host.ApplyOps("hwp", new JsonObject
            {
                ["ops"] = pdfOps.DeepClone(),
                ["dryRun"] = false,
                ["confirmToken"] = Json.GetString(pdfDry, "confirmToken"),
                ["highRiskConfirm"] = true,
            });
            Assert.True(Json.GetBool(pdfApply, "ok"), $"pdf export failed: {pdfApply}");
            Assert.True(File.Exists(samplePdf) && new FileInfo(samplePdf).Length > 0,
                $"sample PDF was not created: {samplePdf}");
        }
    }

    public void Dispose()
    {
        if (_adapter is not null && _createdApp is not null)
        {
            try
            {
                _adapter.RunOnAdapterThread<object?>(() =>
                {
                    try
                    {
                        dynamic hwp = _createdApp;
                        try { hwp.Clear(1); } catch { }
                        try { hwp.XHwpWindows.Close(false); } catch { }
                    }
                    catch { }
                    return null;
                });
            }
            catch { }
        }
        _adapter?.Dispose();
        _createdApp = null;
        if (_createdProcessId > 0)
        {
            try
            {
                using var process = Process.GetProcessById(_createdProcessId);
                if (!process.WaitForExit(1500) &&
                    string.Equals(process.ProcessName, "Hwp", StringComparison.OrdinalIgnoreCase))
                {
                    process.Kill(entireProcessTree: false);
                    process.WaitForExit(5000);
                }
            }
            catch (ArgumentException) { }
            catch (InvalidOperationException) { }
        }
        _createdProcessId = 0;
        GC.Collect();
        GC.WaitForPendingFinalizers();
        _home.Dispose();
    }
}
