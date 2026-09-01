using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

/// <summary>
/// 실Excel E2E (M1 인수 조건): DOCBRIDGE_E2E=1 일 때만 실행.
/// Excel 인스턴스는 어댑터 STA 스레드 안에서 생성(팩토리 패턴)해
/// 크로스-아파트먼트 COM 마샬링을 원천 배제한다.
/// </summary>
public class ExcelE2ETests : IDisposable
{
    private static bool Enabled =>
        string.Equals(Environment.GetEnvironmentVariable("DOCBRIDGE_E2E"), "1", StringComparison.Ordinal);

    private readonly TestHome _home = new();
    private ExcelAdapter? _adapter;
    private object? _createdApp;
    private string? _workbookPath;

    private DocBridgeHost CreateHostWithExcel()
    {
        var dir = _home.Dir;
        _adapter = new ExcelAdapter(() =>
        {
            // 어댑터 STA 스레드 안에서 실행됨 (RCW 생성/사용 아파트먼트 일치)
            var type = Type.GetTypeFromProgID("Excel.Application")
                ?? throw new InvalidOperationException("Excel not installed");
            dynamic app = Activator.CreateInstance(type)!;
            app.Visible = false;
            object? workbooks = null;
            object? workbook = null;
            object? sheet = null;
            object? worksheets = null;
            object? helperSheet = null;
            try
            {
                workbooks = (object)app.Workbooks;
                workbook = (object)((dynamic)workbooks).Add();
                sheet = (object)((dynamic)workbook).ActiveSheet;
                ((dynamic)sheet).Name = "매출";
                SetCellValue(sheet, "A1", "항목");
                SetCellValue(sheet, "B1", "금액");
                SetCellValue(sheet, "A2", "사과");
                SetCellValue(sheet, "B2", 1000d);
                SetCellValue(sheet, "A3", "배");
                SetCellValue(sheet, "B3", 2000d);
                SetCellValue(sheet, "A4", "사과주스");
                SetCellValue(sheet, "B4", 3000d);
                SetCellValue(sheet, "A5", "합계");
                SetCellFormula(sheet, "B5", "=SUM(B2:B4)");
                SetCellValue(sheet, "A7", "병합 제목");
                SetCellFillColor(sheet, "B7", 255d);
                SetCellFillColor(sheet, "C7", 65280d);
                SetCellExtendedStyle(sheet, "B7");
                SetCellValue(sheet, "B8", "삭제되면 안 됨");
                SetCellValue(sheet, "D10", "코드");
                SetCellValue(sheet, "E10", "수량");
                SetCellValue(sheet, "D11", "A");
                SetCellValue(sheet, "E11", 1d);
                CreateTestTable(sheet, "D10:E11");

                worksheets = (object)((dynamic)workbook).Worksheets;
                helperSheet = (object)((dynamic)worksheets).Add(After: sheet);
                ((dynamic)helperSheet).Name = "보조";
                ((dynamic)sheet).Activate();

                _workbookPath = Path.Combine(dir, "sample.xlsx");
                ((dynamic)workbook).SaveAs(_workbookPath);
            }
            finally
            {
                RotHelper.ReleaseComObject(sheet);
                RotHelper.ReleaseComObject(helperSheet);
                RotHelper.ReleaseComObject(worksheets);
                RotHelper.ReleaseComObject(workbook);
                RotHelper.ReleaseComObject(workbooks);
            }
            _createdApp = app;
            return (object)app;
        }, appFactoryOwnsInstance: true);
        var host = new DocBridgeHost(_home.Options);
        host.Router.Register("excel", _adapter);
        return host;
    }

    [Fact]
    public void Excel_full_flow()
    {
        if (!Enabled) return;
        var host = CreateHostWithExcel();
        try
        {

        // 1) get_active_context — 구조화된 JSON 반환
        var ctx = host.GetActiveContext("excel");
        Assert.True(Json.GetBool(ctx, "ok"), $"context failed: {ctx}");
        Assert.Empty(Json.GetArr(ctx, "errors")!);
        Assert.Equal("excel", Json.GetString(ctx, "app"));
        Assert.Equal(_workbookPath, Json.GetString(ctx, "documentRef"));
        var summary = Json.GetObj(ctx, "summary")!;
        var sheets = Json.GetArr(summary, "sheets")!;
        Assert.Contains(sheets, s => s!.GetValue<string>() == "매출");
        Assert.Equal("매출", Json.GetString(summary, "activeSheet"));
        Assert.Equal("매출!A1:E11", Json.GetString(summary, "usedRange"));
        Assert.True(Json.GetBool(summary, "saved"));
        Assert.NotEmpty(Json.GetArr(summary, "openWorkbooks")!);
        Assert.NotNull(Json.GetObj(ctx, "selection"));
        var sheetStates = Json.GetArr(summary, "sheetStates")!;
        Assert.Contains(sheetStates, state => Json.GetString(state as JsonObject, "name") == "보조" &&
                                             Json.GetString(state as JsonObject, "visibility") == "visible");

        // The first enumeration must not final-release an RCW that the adapter still needs.
        var repeatedContext = host.GetActiveContext("excel");
        Assert.True(Json.GetBool(repeatedContext, "ok"), $"repeated context failed: {repeatedContext}");
        Assert.Empty(Json.GetArr(repeatedContext, "errors")!);
        Assert.Equal("매출", Json.GetString(Json.GetObj(repeatedContext, "summary"), "activeSheet"));

        // 2) excel_read_range
        var read = host.Read("excel", new JsonObject { ["range"] = "A1:B5" });
        Assert.True(Json.GetBool(read, "ok"), $"read failed: {read}");
        var values = Json.GetArr(read, "values")!;
        Assert.Equal(5, values.Count);
        Assert.Equal("사과", values[1]![0]!.GetValue<string>());

        // 2a) 행·열 숨김/표시: dry-run/apply/read layout/operation-scoped restore.
        var visibilityOps = new JsonArray
        {
            new JsonObject
            {
                ["op"] = "set_rows_hidden", ["row"] = 3, ["count"] = 2, ["hidden"] = true,
                ["target"] = new JsonObject { ["sheet"] = "매출" },
            },
            new JsonObject
            {
                ["op"] = "set_cols_hidden", ["col"] = "B", ["count"] = 1, ["hidden"] = true,
                ["target"] = new JsonObject { ["sheet"] = "매출" },
            },
            new JsonObject
            {
                ["op"] = "set_sheet_visibility", ["visibility"] = "hidden",
                ["target"] = new JsonObject { ["sheet"] = "보조" },
            },
        };
        var visibilityDry = host.ApplyOps("excel", new JsonObject
        {
            ["ops"] = visibilityOps.DeepClone(), ["dryRun"] = true,
        });
        Assert.True(Json.GetBool(visibilityDry, "ok"), $"visibility dry-run failed: {visibilityDry}");
        var visibilityApplied = host.ApplyOps("excel", new JsonObject
        {
            ["ops"] = visibilityOps.DeepClone(), ["dryRun"] = false,
            ["confirmToken"] = Json.GetString(visibilityDry, "confirmToken"),
        });
        Assert.True(Json.GetBool(visibilityApplied, "ok"), $"visibility apply failed: {visibilityApplied}");
        Assert.True(Json.GetBool(Json.GetObj(visibilityApplied, "readback"), "verified"));

        var hiddenLayoutRead = host.Read("excel", new JsonObject
        {
            ["sheet"] = "매출", ["range"] = "A1:B5", ["includeLayout"] = true,
        });
        Assert.True(Json.GetBool(hiddenLayoutRead, "ok"), $"layout read failed: {hiddenLayoutRead}");
        var hiddenLayout = Json.GetObj(hiddenLayoutRead, "layout")!;
        Assert.Contains(Json.GetArr(hiddenLayout, "rowStates")!, state =>
            Json.GetInt(state as JsonObject, "row") == 3 && Json.GetBool(state as JsonObject, "hidden"));
        Assert.Contains(Json.GetArr(hiddenLayout, "columnStates")!, state =>
            Json.GetString(state as JsonObject, "col") == "B" && Json.GetBool(state as JsonObject, "hidden"));
        var hiddenContext = host.GetActiveContext("excel");
        Assert.Contains(Json.GetArr(Json.GetObj(hiddenContext, "summary"), "sheetStates")!, state =>
            Json.GetString(state as JsonObject, "name") == "보조" &&
            Json.GetString(state as JsonObject, "visibility") == "hidden");

        var visibilityRestoreDry = host.CoreRestoreSnapshot(new JsonObject
        {
            ["snapshotId"] = Json.GetString(visibilityDry, "snapshotId"),
        });
        var visibilityRestored = host.CoreRestoreSnapshot(new JsonObject
        {
            ["snapshotId"] = Json.GetString(visibilityDry, "snapshotId"),
            ["confirmToken"] = Json.GetString(visibilityRestoreDry, "confirmToken"),
        });
        Assert.True(Json.GetBool(visibilityRestored, "ok"), $"visibility restore failed: {visibilityRestored}");
        Assert.Equal("visibility-state", Json.GetString(visibilityRestored, "restoreMode"));
        var restoredLayout = Json.GetObj(host.Read("excel", new JsonObject
        {
            ["sheet"] = "매출", ["range"] = "A1:B5", ["includeLayout"] = true,
        }), "layout")!;
        Assert.DoesNotContain(Json.GetArr(restoredLayout, "rowStates")!, state => Json.GetBool(state as JsonObject, "hidden"));
        Assert.DoesNotContain(Json.GetArr(restoredLayout, "columnStates")!, state => Json.GetBool(state as JsonObject, "hidden"));

        // Dry-run must simulate operations in sequence rather than re-reading the
        // unchanged workbook for every diff entry.
        var sequentialVisibilityDry = host.ApplyOps("excel", new JsonObject
        {
            ["ops"] = new JsonArray(
                new JsonObject
                {
                    ["op"] = "set_rows_hidden", ["row"] = 6, ["count"] = 1, ["hidden"] = true,
                    ["target"] = new JsonObject { ["sheet"] = "매출" },
                },
                new JsonObject
                {
                    ["op"] = "set_rows_hidden", ["row"] = 6, ["count"] = 1, ["hidden"] = false,
                    ["target"] = new JsonObject { ["sheet"] = "매출" },
                },
                new JsonObject
                {
                    ["op"] = "set_sheet_visibility", ["visibility"] = "hidden",
                    ["target"] = new JsonObject { ["sheet"] = "보조" },
                },
                new JsonObject
                {
                    ["op"] = "set_sheet_visibility", ["visibility"] = "visible",
                    ["target"] = new JsonObject { ["sheet"] = "보조" },
                }),
            ["dryRun"] = true,
        });
        Assert.True(Json.GetBool(sequentialVisibilityDry, "ok"),
            $"sequential visibility dry-run failed: {sequentialVisibilityDry}");
        var sequentialDiff = Json.GetArr(sequentialVisibilityDry, "diff")!;
        Assert.False(Json.GetBool(sequentialDiff[0] as JsonObject, "before"));
        Assert.True(Json.GetBool(sequentialDiff[1] as JsonObject, "before"));
        Assert.Equal("visible", Json.GetString(sequentialDiff[2] as JsonObject, "before"));
        Assert.Equal("hidden", Json.GetString(sequentialDiff[3] as JsonObject, "before"));

        var activeHideBlocked = host.ApplyOps("excel", new JsonObject
        {
            ["ops"] = new JsonArray(new JsonObject
            {
                ["op"] = "set_sheet_visibility", ["visibility"] = "hidden",
                ["target"] = new JsonObject { ["sheet"] = "매출" },
            }),
            ["dryRun"] = true,
        });
        Assert.False(Json.GetBool(activeHideBlocked, "ok"));
        Assert.Contains(Json.GetArr(activeHideBlocked, "errors")!, error =>
            error!.GetValue<string>().Contains("EXCEL_ACTIVE_SHEET_HIDE_BLOCKED", StringComparison.Ordinal));

        // 2aa) 병합/병합해제와 두 방향의 operation-scoped restore.
        var mergeOps = new JsonArray(new JsonObject
        {
            ["op"] = "merge_cells", ["target"] = new JsonObject { ["sheet"] = "매출" }, ["range"] = "A7:C7",
        });
        var mergeDry = host.ApplyOps("excel", new JsonObject { ["ops"] = mergeOps.DeepClone(), ["dryRun"] = true });
        Assert.True(Json.GetBool(mergeDry, "ok"), $"merge dry-run failed: {mergeDry}");
        var mergeApplied = host.ApplyOps("excel", new JsonObject
        {
            ["ops"] = mergeOps.DeepClone(), ["dryRun"] = false,
            ["confirmToken"] = Json.GetString(mergeDry, "confirmToken"),
        });
        Assert.True(Json.GetBool(mergeApplied, "ok"), $"merge apply failed: {mergeApplied}");
        var mergedLayout = Json.GetObj(host.Read("excel", new JsonObject
        {
            ["sheet"] = "매출", ["range"] = "A7:C7", ["includeLayout"] = true,
        }), "layout")!;
        Assert.Contains(Json.GetArr(mergedLayout, "mergedAreas")!, node => node!.GetValue<string>() == "A7:C7");

        var destructiveMergeBlocked = host.ApplyOps("excel", new JsonObject
        {
            ["ops"] = new JsonArray(new JsonObject
            {
                ["op"] = "merge_cells", ["target"] = new JsonObject { ["sheet"] = "매출" }, ["range"] = "A8:C8",
            }),
            ["dryRun"] = true,
        });
        Assert.False(Json.GetBool(destructiveMergeBlocked, "ok"));
        Assert.Contains(Json.GetArr(destructiveMergeBlocked, "errors")!, error =>
            error!.GetValue<string>().Contains("EXCEL_MERGE_WOULD_DELETE_CONTENT", StringComparison.Ordinal));

        var nonContiguousMergeBlocked = host.ApplyOps("excel", new JsonObject
        {
            ["ops"] = new JsonArray(new JsonObject
            {
                ["op"] = "merge_cells", ["target"] = new JsonObject { ["sheet"] = "매출" },
                ["range"] = "A9,C9",
            }),
            ["dryRun"] = true,
        });
        Assert.False(Json.GetBool(nonContiguousMergeBlocked, "ok"));
        Assert.Contains(Json.GetArr(nonContiguousMergeBlocked, "errors")!, error =>
            error!.GetValue<string>().Contains("EXCEL_MERGE_NONCONTIGUOUS_RANGE", StringComparison.Ordinal));

        var partialTableMergeBlocked = host.ApplyOps("excel", new JsonObject
        {
            ["ops"] = new JsonArray(new JsonObject
            {
                ["op"] = "merge_cells", ["target"] = new JsonObject { ["sheet"] = "매출" },
                ["range"] = "C10:D10",
            }),
            ["dryRun"] = true,
        });
        Assert.False(Json.GetBool(partialTableMergeBlocked, "ok"));
        Assert.Contains(Json.GetArr(partialTableMergeBlocked, "errors")!, error =>
            error!.GetValue<string>().Contains("EXCEL_TABLE_MERGE_BLOCKED", StringComparison.Ordinal));

        var unmergeOps = new JsonArray(new JsonObject
        {
            ["op"] = "unmerge_cells", ["target"] = new JsonObject { ["sheet"] = "매출" }, ["range"] = "A7:C7",
        });
        var unmergeDry = host.ApplyOps("excel", new JsonObject { ["ops"] = unmergeOps.DeepClone(), ["dryRun"] = true });
        Assert.True(Json.GetBool(unmergeDry, "ok"), $"unmerge dry-run failed: {unmergeDry}");
        var unmergeApplied = host.ApplyOps("excel", new JsonObject
        {
            ["ops"] = unmergeOps.DeepClone(), ["dryRun"] = false,
            ["confirmToken"] = Json.GetString(unmergeDry, "confirmToken"),
        });
        Assert.True(Json.GetBool(unmergeApplied, "ok"), $"unmerge apply failed: {unmergeApplied}");
        Assert.Empty(Json.GetArr(Json.GetObj(host.Read("excel", new JsonObject
        {
            ["sheet"] = "매출", ["range"] = "A7:C7", ["includeLayout"] = true,
        }), "layout"), "mergedAreas")!);

        var unmergeRestoreDry = host.CoreRestoreSnapshot(new JsonObject { ["snapshotId"] = Json.GetString(unmergeDry, "snapshotId") });
        var unmergeRestored = host.CoreRestoreSnapshot(new JsonObject
        {
            ["snapshotId"] = Json.GetString(unmergeDry, "snapshotId"),
            ["confirmToken"] = Json.GetString(unmergeRestoreDry, "confirmToken"),
        });
        Assert.True(Json.GetBool(unmergeRestored, "ok"), $"unmerge restore failed: {unmergeRestored}");
        Assert.Equal("merge-state", Json.GetString(unmergeRestored, "restoreMode"));

        var mergeRestoreDry = host.CoreRestoreSnapshot(new JsonObject { ["snapshotId"] = Json.GetString(mergeDry, "snapshotId") });
        var mergeRestored = host.CoreRestoreSnapshot(new JsonObject
        {
            ["snapshotId"] = Json.GetString(mergeDry, "snapshotId"),
            ["confirmToken"] = Json.GetString(mergeRestoreDry, "confirmToken"),
        });
        Assert.True(Json.GetBool(mergeRestored, "ok"), $"merge restore failed: {mergeRestored}");
        var unmergedRead = host.Read("excel", new JsonObject
        {
            ["sheet"] = "매출", ["range"] = "A7:C7", ["includeLayout"] = true,
        });
        Assert.Empty(Json.GetArr(Json.GetObj(unmergedRead, "layout"), "mergedAreas")!);
        Assert.Equal("병합 제목", Json.GetArr(unmergedRead, "values")![0]![0]!.GetValue<string>());
        Assert.Equal(255d, Json.GetObj(host.Read("excel", new JsonObject
        {
            ["sheet"] = "매출", ["range"] = "B7", ["includeStyles"] = true,
        }), "styles")!["interiorColor"]!.GetValue<double>());
        Assert.Equal(65280d, Json.GetObj(host.Read("excel", new JsonObject
        {
            ["sheet"] = "매출", ["range"] = "C7", ["includeStyles"] = true,
        }), "styles")!["interiorColor"]!.GetValue<double>());

        // 2b) An explicit source and destination referring to the same workbook used to
        // separate the shared Workbook RCW during preview/apply. Exercise both phases and
        // then issue another preview through the same adapter to prove the lease remains live.
        var copyOps = new JsonArray
        {
            new JsonObject
            {
                ["op"] = "copy_sheet",
                ["sourceWorkbook"] = _workbookPath,
                ["sourceSheet"] = "매출",
                ["targetWorkbook"] = _workbookPath,
                ["targetSheet"] = "매출_복사",
            },
        };
        var copyDry = host.ApplyOps("excel", new JsonObject
        {
            ["ops"] = copyOps.DeepClone(),
            ["dryRun"] = true,
        });
        Assert.True(Json.GetBool(copyDry, "ok"), $"same-workbook copy dry-run failed: {copyDry}");
        Assert.Empty(Json.GetArr(copyDry, "errors")!);
        Assert.NotNull(Json.GetString(copyDry, "confirmToken"));

        var copyApplied = host.ApplyOps("excel", new JsonObject
        {
            ["ops"] = copyOps.DeepClone(),
            ["dryRun"] = false,
            ["confirmToken"] = Json.GetString(copyDry, "confirmToken"),
        });
        Assert.True(Json.GetBool(copyApplied, "ok"), $"same-workbook copy apply failed: {copyApplied}");
        Assert.True(Json.GetBool(Json.GetObj(copyApplied, "readback"), "verified"));
        Assert.True(Json.GetBool(Json.GetObj(copyApplied, "interaction"), "originalStateRestored"));

        var copied = host.Read("excel", new JsonObject
        {
            ["workbook"] = _workbookPath,
            ["sheet"] = "매출_복사",
            ["range"] = "A1:B5",
        });
        Assert.True(Json.GetBool(copied, "ok"), $"copied sheet read failed: {copied}");
        Assert.Equal("사과", Json.GetArr(copied, "values")![1]![0]!.GetValue<string>());

        var postCopyPreview = host.ApplyOps("excel", new JsonObject
        {
            ["ops"] = new JsonArray
            {
                new JsonObject
                {
                    ["op"] = "set_values",
                    ["targetWorkbook"] = _workbookPath,
                    ["target"] = new JsonObject { ["sheet"] = "매출_복사" },
                    ["range"] = "B2:B2",
                    ["values"] = new JsonArray(new JsonArray(1111d)),
                },
            },
            ["dryRun"] = true,
        });
        Assert.True(Json.GetBool(postCopyPreview, "ok"), $"post-copy preview failed: {postCopyPreview}");
        Assert.Empty(Json.GetArr(postCopyPreview, "errors")!);

        // Restore the copy-only snapshot through the real Excel COM path. The operation-
        // scoped inverse must remove only the copied sheet and restore the original active
        // sheet without rewriting any pre-existing Formula/Value2 cells.
        var copyRestoreDry = host.CoreRestoreSnapshot(new JsonObject
        {
            ["snapshotId"] = Json.GetString(copyDry, "snapshotId"),
        });
        Assert.True(Json.GetBool(copyRestoreDry, "ok"), $"copy restore dry-run failed: {copyRestoreDry}");
        var copyRestored = host.CoreRestoreSnapshot(new JsonObject
        {
            ["snapshotId"] = Json.GetString(copyDry, "snapshotId"),
            ["confirmToken"] = Json.GetString(copyRestoreDry, "confirmToken"),
        });
        Assert.True(Json.GetBool(copyRestored, "ok"), $"copy restore failed: {copyRestored}");
        Assert.Equal("copy-sheet-topology", Json.GetString(copyRestored, "restoreMode"));
        Assert.True(Json.GetBool(Json.GetObj(copyRestored, "readback"), "verified"));

        var contextAfterCopyRestore = host.GetActiveContext("excel");
        Assert.True(Json.GetBool(contextAfterCopyRestore, "ok"), $"context after copy restore failed: {contextAfterCopyRestore}");
        var restoredSummary = Json.GetObj(contextAfterCopyRestore, "summary")!;
        Assert.Equal("매출", Json.GetString(restoredSummary, "activeSheet"));
        Assert.DoesNotContain(Json.GetArr(restoredSummary, "sheets")!,
            value => string.Equals(value?.GetValue<string>(), "매출_복사", StringComparison.Ordinal));

        var originalFormulaAfterCopyRestore = host.Read("excel", new JsonObject
        {
            ["workbook"] = _workbookPath,
            ["sheet"] = "매출",
            ["range"] = "B5",
            ["includeFormulas"] = true,
        });
        Assert.True(Json.GetBool(originalFormulaAfterCopyRestore, "ok"));
        Assert.Equal("=SUM(B2:B4)",
            Json.GetArr(originalFormulaAfterCopyRestore, "formulas")![0]![0]!.GetValue<string>());

        // 3) find_replace dry-run → diff + confirmToken
        var frBatch = new JsonObject
        {
            ["ops"] = new JsonArray
            {
                new JsonObject
                {
                    ["op"] = "find_replace",
                    ["find"] = "사과",
                    ["replace"] = "청사과",
                    ["target"] = new JsonObject { ["scope"] = "sheet", ["sheet"] = "매출" },
                    ["options"] = new JsonObject { ["matchCase"] = false },
                },
            },
            ["dryRun"] = true,
        };
        var dry = host.ApplyOps("excel", frBatch);
        Assert.True(Json.GetBool(dry, "ok"), $"dry-run failed: {dry}");
        var token = Json.GetString(dry, "confirmToken");
        var snapshotId = Json.GetString(dry, "snapshotId");
        Assert.NotNull(token);
        Assert.NotNull(snapshotId);
        var diff = Json.GetArr(dry, "diff")!;
        Assert.Equal(2, diff.Count); // A2 "사과", A4 "사과주스"

        // 3b) confirmToken 없이 apply → 실패 (인수 조건)
        var noToken = new JsonObject
        {
            ["ops"] = frBatch["ops"]!.DeepClone(),
            ["dryRun"] = false,
        };
        Assert.False(Json.GetBool(host.ApplyOps("excel", noToken), "ok"));

        // 4) apply with token → readback verified
        var apply = new JsonObject
        {
            ["ops"] = frBatch["ops"]!.DeepClone(),
            ["dryRun"] = false,
            ["confirmToken"] = token,
        };
        var applied = host.ApplyOps("excel", apply);
        Assert.True(Json.GetBool(applied, "ok"), $"apply failed: {applied}");
        Assert.True(Json.GetBool(Json.GetObj(applied, "readback"), "verified"));

        // 5) 실제 값 확인
        var after = host.Read("excel", new JsonObject { ["range"] = "A2:A4" });
        var av = Json.GetArr(after, "values")!;
        Assert.Equal("청사과", av[0]![0]!.GetValue<string>());
        Assert.Equal("청사과주스", av[2]![0]!.GetValue<string>());

        // 6) set_values + format_range dry-run/apply
        var styleBefore = host.Read("excel", new JsonObject { ["range"] = "A1", ["includeStyles"] = true });
        var beforeStyles = Json.GetObj(styleBefore, "styles")!.DeepClone() as JsonObject;
        var svBatch = new JsonObject
        {
            ["ops"] = new JsonArray
            {
                new JsonObject { ["op"] = "set_values", ["range"] = "B2:B2",
                    ["target"] = new JsonObject { ["sheet"] = "매출" },
                    ["values"] = new JsonArray(new JsonArray(1500)) },
                new JsonObject { ["op"] = "set_formulas", ["range"] = "'매출'!B5:B5",
                    ["formulas"] = new JsonArray(new JsonArray("=B2+B3+B4")) },
                new JsonObject { ["op"] = "format_range", ["range"] = "A1:B1",
                    ["target"] = new JsonObject { ["sheet"] = "매출" },
                    ["style"] = new JsonObject { ["bold"] = true, ["fillColor"] = "#FFE699" } },
            },
            ["dryRun"] = true,
        };
        var svDry = host.ApplyOps("excel", svBatch);
        Assert.True(Json.GetBool(svDry, "ok"));
        var svSnapshotId = Json.GetString(svDry, "snapshotId")!;
        var svApply = new JsonObject
        {
            ["ops"] = svBatch["ops"]!.DeepClone(),
            ["dryRun"] = false,
            ["confirmToken"] = Json.GetString(svDry, "confirmToken"),
        };
        var svApplied = host.ApplyOps("excel", svApply);
        Assert.True(Json.GetBool(svApplied, "ok"), $"set_values apply failed: {svApplied}");

        var b2 = host.Read("excel", new JsonObject { ["range"] = "B2" });
        Assert.Equal(1500.0, Json.GetArr(b2, "values")![0]![0]!.GetValue<double>());
        var b5 = host.Read("excel", new JsonObject { ["range"] = "B5", ["includeFormulas"] = true });
        Assert.Equal("=B2+B3+B4", Json.GetArr(b5, "formulas")![0]![0]!.GetValue<string>());
        var styled = host.Read("excel", new JsonObject { ["range"] = "A1", ["includeStyles"] = true });
        Assert.True(Json.GetBool(Json.GetObj(styled, "styles"), "fontBold"));

        // 수식/숫자 타입까지 두 번째 스냅샷으로 복원
        var svRestoreDry = host.CoreRestoreSnapshot(new JsonObject { ["snapshotId"] = svSnapshotId });
        var svRestored = host.CoreRestoreSnapshot(new JsonObject
        {
            ["snapshotId"] = svSnapshotId,
            ["confirmToken"] = Json.GetString(svRestoreDry, "confirmToken"),
        });
        Assert.True(Json.GetBool(svRestored, "ok"), $"typed restore failed: {svRestored}");
        var typedRestored = host.Read("excel", new JsonObject { ["range"] = "B2:B5", ["includeFormulas"] = true });
        Assert.Equal(1000.0, Json.GetArr(typedRestored, "values")![0]![0]!.GetValue<double>());
        Assert.Equal("=SUM(B2:B4)", Json.GetArr(typedRestored, "formulas")![3]![0]!.GetValue<string>());
        var styleRestored = Json.GetObj(host.Read("excel", new JsonObject { ["range"] = "A1", ["includeStyles"] = true }), "styles")!;
        Assert.Equal(Json.GetBool(beforeStyles, "fontBold"), Json.GetBool(styleRestored, "fontBold"));
        Assert.Equal(beforeStyles!["interiorColor"]!.GetValue<double>(), styleRestored["interiorColor"]!.GetValue<double>());

        // 6b) 삽입 행은 복원 시 역순 삭제되어 원래 구조/수식 좌표로 돌아온다.
        var insertBatch = new JsonObject
        {
            ["ops"] = new JsonArray
            {
                new JsonObject
                {
                    ["op"] = "insert_rows", ["row"] = 3, ["count"] = 2,
                    ["target"] = new JsonObject { ["sheet"] = "매출" },
                },
            },
            ["dryRun"] = true,
        };
        var insertDry = host.ApplyOps("excel", insertBatch);
        Assert.True(Json.GetBool(insertDry, "ok"), $"insert dry-run failed: {insertDry}");
        var insertApplied = host.ApplyOps("excel", new JsonObject
        {
            ["ops"] = insertBatch["ops"]!.DeepClone(),
            ["dryRun"] = false,
            ["confirmToken"] = Json.GetString(insertDry, "confirmToken"),
        });
        Assert.True(Json.GetBool(insertApplied, "ok"), $"insert apply failed: {insertApplied}");
        var shifted = host.Read("excel", new JsonObject { ["range"] = "B7", ["includeFormulas"] = true });
        Assert.Equal("=SUM(B2:B6)", Json.GetArr(shifted, "formulas")![0]![0]!.GetValue<string>());

        var insertRestoreDry = host.CoreRestoreSnapshot(new JsonObject { ["snapshotId"] = Json.GetString(insertDry, "snapshotId") });
        var insertRestored = host.CoreRestoreSnapshot(new JsonObject
        {
            ["snapshotId"] = Json.GetString(insertDry, "snapshotId"),
            ["confirmToken"] = Json.GetString(insertRestoreDry, "confirmToken"),
        });
        Assert.True(Json.GetBool(insertRestored, "ok"), $"insert restore failed: {insertRestored}");
        var structureRestored = host.Read("excel", new JsonObject { ["range"] = "A3:B5", ["includeFormulas"] = true });
        Assert.Equal("배", Json.GetArr(structureRestored, "values")![0]![0]!.GetValue<string>());
        Assert.Equal("=SUM(B2:B4)", Json.GetArr(structureRestored, "formulas")![2]![1]!.GetValue<string>());

        // 7) 스냅샷 복원 (find_replace 이전 상태로)
        var restoreDry = host.CoreRestoreSnapshot(new JsonObject { ["snapshotId"] = snapshotId });
        Assert.True(Json.GetBool(restoreDry, "ok"));
        var restored = host.CoreRestoreSnapshot(new JsonObject
        {
            ["snapshotId"] = snapshotId,
            ["confirmToken"] = Json.GetString(restoreDry, "confirmToken"),
        });
        Assert.True(Json.GetBool(restored, "ok"), $"restore failed: {restored}");

        var restoredVal = host.Read("excel", new JsonObject { ["range"] = "A2" });
        Assert.Equal("사과", Json.GetArr(restoredVal, "values")![0]![0]!.GetValue<string>());
        }
        finally
        {
            CloseOwnedTestExcel();
            host.Dispose();
        }
    }

    [Fact]
    public void Excel_adapter_reports_unavailable_when_factory_returns_null()
    {
        using var adapter = new ExcelAdapter(() => null);
        var st = adapter.GetStatus();
        Assert.False(st.Available);
        Assert.NotNull(st.Detail);
    }

    private void CloseOwnedTestExcel()
    {
        if (_adapter is not null && _createdApp is not null)
        {
            _adapter.RunOnAdapterThread<object?>(() =>
            {
                object? workbooks = null;
                try
                {
                    dynamic app = _createdApp;
                    workbooks = (object)app.Workbooks;
                    for (var index = Convert.ToInt32(((dynamic)workbooks).Count); index >= 1; index--)
                    {
                        object? workbook = null;
                        try
                        {
                            workbook = (object)((dynamic)workbooks).Item(index);
                            ((dynamic)workbook).Close(false);
                        }
                        finally { RotHelper.ReleaseComObject(workbook); }
                    }
                    if (Convert.ToInt32(((dynamic)workbooks).Count) != 0)
                        throw new InvalidOperationException("Excel E2E cleanup did not close every test workbook");
                }
                finally { RotHelper.ReleaseComObject(workbooks); }
                return null;
            });
        }
        _createdApp = null;
        var disconnect = _adapter?.Disconnect();
        if (disconnect is not null && !Json.GetBool(disconnect, "ok"))
            throw new InvalidOperationException($"Excel E2E disconnect failed: {disconnect}");
        _adapter = null;
    }

    public void Dispose()
    {
        CloseOwnedTestExcel();
        _home.Dispose();
    }

    private static void SetCellValue(object sheet, string address, object value)
    {
        object? range = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            ((dynamic)range).Value2 = value;
        }
        finally { RotHelper.ReleaseComObject(range); }
    }

    private static void SetCellFormula(object sheet, string address, string formula)
    {
        object? range = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            ((dynamic)range).Formula = formula;
        }
        finally { RotHelper.ReleaseComObject(range); }
    }

    private static void SetCellFillColor(object sheet, string address, double color)
    {
        object? range = null;
        object? interior = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            interior = (object)((dynamic)range).Interior;
            ((dynamic)interior).Color = color;
        }
        finally
        {
            RotHelper.ReleaseComReference(interior);
            RotHelper.ReleaseComReference(range);
        }
    }

    private static void SetCellExtendedStyle(object sheet, string address)
    {
        object? range = null;
        object? font = null;
        object? interior = null;
        object? borders = null;
        object? rightBorder = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            font = (object)((dynamic)range).Font;
            interior = (object)((dynamic)range).Interior;
            borders = (object)((dynamic)range).Borders;
            rightBorder = (object)((dynamic)borders).Item(10);

            ((dynamic)font).Name = "Arial";
            ((dynamic)font).Underline = 2;
            ((dynamic)font).Strikethrough = true;
            ((dynamic)range).NumberFormat = "0.00";
            ((dynamic)range).HorizontalAlignment = -4152;
            ((dynamic)range).VerticalAlignment = -4160;
            ((dynamic)range).WrapText = true;
            ((dynamic)range).ShrinkToFit = true;
            ((dynamic)range).IndentLevel = 2;
            ((dynamic)range).Orientation = 15;
            ((dynamic)interior).PatternColor = 65535d;
            ((dynamic)interior).Pattern = 1;
            ((dynamic)rightBorder).Color = 16711680d;
            ((dynamic)rightBorder).Weight = 4;
            ((dynamic)rightBorder).LineStyle = 1;
            ((dynamic)range).Locked = false;
            ((dynamic)range).FormulaHidden = true;
        }
        finally
        {
            RotHelper.ReleaseComReference(rightBorder);
            RotHelper.ReleaseComReference(borders);
            RotHelper.ReleaseComReference(interior);
            RotHelper.ReleaseComReference(font);
            RotHelper.ReleaseComReference(range);
        }
    }

    private static void CreateTestTable(object sheet, string address)
    {
        object? range = null;
        object? listObjects = null;
        object? table = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            listObjects = (object)((dynamic)sheet).ListObjects;
            table = (object)((dynamic)listObjects).Add(1, range, Type.Missing, 1);
            ((dynamic)table).Name = "DocBridgeMergeSafetyTable";
        }
        finally
        {
            RotHelper.ReleaseComReference(table);
            RotHelper.ReleaseComReference(listObjects);
            RotHelper.ReleaseComReference(range);
        }
    }
}
