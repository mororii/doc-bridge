namespace DocBridge.Development.ExcelProductionWorkflowProbe;

internal static class ScenarioExtended
{
    public static List<PlannedBatch> E7(string artifactDir, string sourceXlsx)
    {
        const string sheet = "원장";
        var xlsx = Path.Combine(artifactDir, "e7-pivot-ledger.xlsx");
        var values = new JsonArray
        {
            Arr("품목", "월", "금액"),
            Arr("PVC관", "1월", 1_200_000),
            Arr("맨홀", "1월", 420_000),
            Arr("PVC관", "2월", 800_000),
            Arr("맨홀", "2월", 210_000),
        };
        return
        [
            ApplyPlanner.Batch("e7", "e7-create", OpFamily.Lifecycle, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "create_workbook",
                ["sheetName"] = sheet,
            }), "token"),
            ApplyPlanner.Batch("e7", "e7-save-as", OpFamily.Lifecycle, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "save_workbook",
                ["output"] = xlsx,
            }), "token"),
            ApplyPlanner.Batch("e7", "e7-values", OpFamily.Values, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "set_values",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "A1:C5",
                ["values"] = values,
            }), "execute"),
            ApplyPlanner.Batch("e7", "e7-table", OpFamily.Data, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "create_table",
                ["target"] = JsonUtil.Target(sheet),
                ["name"] = "원장표",
                ["range"] = "A1:C5",
                ["hasHeaders"] = true,
            }), "token", "native table is the expanding source"),
            ApplyPlanner.Batch("e7", "e7-add-pivot-sheet", OpFamily.Structure, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "add_sheet",
                ["name"] = "피벗",
                ["afterSheet"] = sheet,
            }), "token"),
            ApplyPlanner.Batch("e7", "e7-pivot", OpFamily.Data, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "create_pivot",
                ["target"] = JsonUtil.Target("피벗"),
                ["name"] = "원장피벗",
                ["sourceRange"] = "'원장'!A1:C5",
                ["destination"] = "A3",
                ["rows"] = JsonUtil.Arr("품목"),
                ["columns"] = JsonUtil.Arr("월"),
                ["values"] = JsonUtil.Arr(new JsonObject { ["field"] = "금액", ["function"] = "sum" }),
            }), "token", "target.sheet is destination 피벗; sourceRange is sheet-qualified 원장!A1:C5"),
            ApplyPlanner.Batch("e7", "e7-add-row", OpFamily.Values, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "set_values",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "A6:C6",
                ["values"] = Arr(Arr("PVC관", "3월", 500_000)),
            }), "execute"),
            ApplyPlanner.Batch("e7", "e7-resize-table", OpFamily.Data, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "resize_table",
                ["target"] = JsonUtil.Target(sheet),
                ["name"] = "원장표",
                ["range"] = "A1:C6",
            }), "token", "table expands to include row 6"),
            ApplyPlanner.Batch("e7", "e7-update-pivot-source", OpFamily.Data, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "update_pivot",
                ["target"] = JsonUtil.Target("피벗"),
                ["name"] = "원장피벗",
                ["sourceRange"] = "원장표",
            }), "token", "data pin: sourceRange 원장표; equivalent '원장'!A1:C6; expected 3130000 / PVC 2500000 / manhole 630000"),
            ApplyPlanner.Batch("e7", "e7-refresh-pivot", OpFamily.Data, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "refresh_pivot",
                ["target"] = JsonUtil.Target("피벗"),
                ["name"] = "원장피벗",
            }), "token", "refresh after sourceRange update; not a claim that refresh-alone of A1:C5 includes row 6"),
            .. RemainingOpWorkflows.E7(sheet),
            .. ScenarioCore.SaveCloseReopen("e7", xlsx, Path.ChangeExtension(xlsx, ".pdf"), sheet),
        ];
    }

    public static List<PlannedBatch> E8(string artifactDir, string sourceXlsx)
    {
        const string sheet = "입력양식";
        var xlsx = Path.Combine(artifactDir, "e8-protected-form.xlsx");
        return
        [
            ApplyPlanner.Batch("e8", "e8-create", OpFamily.Lifecycle, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "create_workbook",
                ["sheetName"] = sheet,
            }), "token"),
            ApplyPlanner.Batch("e8", "e8-save-as", OpFamily.Lifecycle, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "save_workbook",
                ["output"] = xlsx,
            }), "token"),
            ApplyPlanner.Batch("e8", "e8-values", OpFamily.Values, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "set_values",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "A1:B3",
                ["values"] = new JsonArray
                {
                    Arr("입력 (잠금 해제)", 10),
                    Arr("단가 (잠금 해제)", 1000),
                    Arr("금액 (수식 잠금)", null),
                },
            }), "execute"),
            ApplyPlanner.Batch("e8", "e8-formula", OpFamily.Values, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "set_formulas",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "B3",
                ["formulas"] = new JsonArray { new JsonArray { "=B1*B2" } },
            }), "execute"),
            ApplyPlanner.Batch("e8", "e8-unlock-inputs", OpFamily.Format, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "A1:B2",
                ["style"] = new JsonObject { ["locked"] = false },
            }), "execute"),
            ApplyPlanner.Batch("e8", "e8-lock-formula", OpFamily.Format, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "format_range",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "B3",
                ["style"] = new JsonObject { ["locked"] = true },
            }), "execute"),
            ApplyPlanner.Batch("e8", "e8-protect", OpFamily.Protect, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "protect_sheet",
                ["target"] = JsonUtil.Target(sheet),
            }), "token", "no password field; secrets never logged"),
            .. RemainingOpWorkflows.E8(sheet),
            .. ScenarioCore.SaveCloseReopen("e8", xlsx, Path.ChangeExtension(xlsx, ".pdf"), sheet),
        ];
    }

    public static List<PlannedBatch> E9(string artifactDir, string sourceXlsx, string fixtureDir)
    {
        const string sheet = "정리";
        var xlsx = Path.Combine(artifactDir, "e9-csv-cleanup.xlsx");
        var csv = Path.Combine(fixtureDir, "e9-sample.csv");
        Directory.CreateDirectory(fixtureDir);
        File.WriteAllText(csv,
            "\"id\",\"name\",\"qty\"\r\n\"1\",\"PVC, pipe\",\"10\"\r\n\"1\",\"PVC, pipe\",\"10\"\r\n\"2\",\"맨홀\",\"3\"\r\n",
            new UTF8Encoding(false));
        return
        [
            ApplyPlanner.Batch("e9", "e9-create", OpFamily.Lifecycle, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "create_workbook",
                ["sheetName"] = sheet,
            }), "token"),
            ApplyPlanner.Batch("e9", "e9-save-as", OpFamily.Lifecycle, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "save_workbook",
                ["output"] = xlsx,
            }), "token"),
            ApplyPlanner.Batch("e9", "e9-import", OpFamily.Lifecycle, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "import_csv",
                ["path"] = csv,
                ["target"] = JsonUtil.Target(sheet),
            }), "token", "blocked until import_csv is advertised"),
            ApplyPlanner.Batch("e9", "e9-dedupe", OpFamily.Data, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "remove_duplicates",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = "A1:C4",
            }), "token"),
            .. RemainingOpWorkflows.E9(sheet),
            ApplyPlanner.Batch("e9", "e9-export", OpFamily.Lifecycle, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "export_csv",
                ["output"] = Path.Combine(artifactDir, "e9-clean.csv"),
            }), "token"),
            .. ScenarioCore.SaveCloseReopen("e9", xlsx, Path.ChangeExtension(xlsx, ".pdf"), sheet),
        ];
    }

    public static JsonObject ExpectedE7() => new()
    {
        ["failedRequirementUnless"] = "create_table + resize_table + create_pivot + update_pivot sourceRange advertised",
        ["crossSheetDestination"] = "target.sheet=피벗 destination A3; sourceRange='원장'!A1:C5",
        ["expandingSource"] = "set_values A6:C6 PVC/3월/500000 + resize_table 원장표 A1:C6; update_pivot sourceRange=원장표 (equivalent '원장'!A1:C6)",
        ["refreshAloneA1C5IsNotRow6"] = true,
        ["dataContract"] = "cursor-data/pivot-chart-20260911/REPORT.md",
        ["pivotTotals"] = new JsonObject { ["pvc"] = 2_500_000, ["manhole"] = 630_000, ["grand"] = 3_130_000 },
        ["afterAdd"] = new JsonObject { ["pvcMarch"] = 500_000 },
    };

    public static JsonObject ExpectedE8() => new()
    {
        ["amountInitial"] = 10_000,
        ["amountAfterB1"] = 11_000,
        ["protect"] = "locked B3 write refused; unlocked B1=11 recalculates B3 to 11000; no password field",
    };

    public static JsonObject ExpectedE9() => new()
    {
        ["uniqueRows"] = 2,
        ["failedRequirementUnless"] = "import_csv / export_csv / remove_duplicates advertised by data dispatch ctx_279fd10e42c9",
    };

    private static JsonArray Arr(params object?[] cells)
    {
        var a = new JsonArray();
        foreach (var c in cells) a.Add(JsonUtil.FromClr(c));
        return a;
    }
}
