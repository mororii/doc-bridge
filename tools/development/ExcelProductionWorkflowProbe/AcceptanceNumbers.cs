namespace DocBridge.Development.ExcelProductionWorkflowProbe;

internal static class AcceptanceNumbers
{
    public const double Bp95 = 100d;
    public const double Bp96 = 22656d;
    public const double Ab95 = 17.659d;
    public const int MergeCount = 281;
    public const int MonthHeaderMerges = 31;
    public const int BarCount = 941;

    public static readonly CostLine[] CostLines =
    [
        new("관로 부설", 125.5, 48000, 6_024_000),
        new("맨홀 설치", 8, 420_000, 3_360_000),
        new("굴착", 215.75, 18_000, 3_883_500),
        new("되메우기", 198.25, 14_500, 2_874_625),
    ];

    public const double CostSubtotal = 16_142_125;
    public const double CostVat = 1_614_213;
    public const double CostGrand = 17_756_338;
    public const double CostQtyChanged = 130;
    public const double CostAmountChanged = 6_240_000;
    public const double CostSubtotalChanged = 16_358_125;
    public const double CostVatChanged = 1_635_813;
    public const double CostGrandChanged = 17_993_938;

    public const int DailyPeople = 13;
    public const int DailyEquipment = 3;

    public const double PvcRemain = 65;
    public const double ManholeRemain = 7;
    public const double PvcRemainAfterAdd = 80;

    public const double ContractAmount = 100_000_000;
    public const double Jan = 12_000_000;
    public const double Feb = 18_000_000;
    public const double Mar = 15_000_000;
    public const double Cum = 45_000_000;
    public const double CumPct = 45;
    public const double Remain = 55_000_000;
    public const double FebChanged = 20_000_000;
    public const double CumChanged = 47_000_000;
    public const double CumPctChanged = 47;
    public const double RemainChanged = 53_000_000;

    public static readonly double[] PlanRates = [10, 25, 40, 60, 80, 100];
    public static readonly double[] ActualRates = [8, 23, 43, 57, 82, 98];
    public const double MarActualChanged = 45;
    public const double MarDiffChanged = 5;

    public static JsonObject SelfCheck()
    {
        var errors = new JsonArray();
        double Round0(double x) => Math.Round(x, 0, MidpointRounding.AwayFromZero);

        var sub = 0d;
        foreach (var line in CostLines)
        {
            var amount = Round0(line.Qty * line.Unit);
            if (Math.Abs(amount - line.Amount) > 0.0)
                errors.Add($"{line.Name} ROUND(qty*unit,0) {amount} != {line.Amount}");
            sub += amount;
        }
        if (Math.Abs(sub - CostSubtotal) > 0) errors.Add($"subtotal {sub} != {CostSubtotal}");
        var vat = Round0(sub * 0.1);
        if (Math.Abs(vat - CostVat) > 0) errors.Add($"vat {vat} != {CostVat}");
        if (Math.Abs(sub + vat - CostGrand) > 0) errors.Add("grand mismatch");

        var changed = Round0(CostQtyChanged * CostLines[0].Unit);
        if (Math.Abs(changed - CostAmountChanged) > 0) errors.Add("changed amount");
        var sub2 = sub - CostLines[0].Amount + changed;
        if (Math.Abs(sub2 - CostSubtotalChanged) > 0) errors.Add("changed subtotal");
        if (Math.Abs(Round0(sub2 * 0.1) - CostVatChanged) > 0) errors.Add("changed vat");
        if (Math.Abs(sub2 + Round0(sub2 * 0.1) - CostGrandChanged) > 0) errors.Add("changed grand");

        Check(errors, Add(6, 4, 3), DailyPeople, "daily people");
        Check(errors, Add(2, 1), DailyEquipment, "daily equipment");
        Check(errors, Add(120, -40, 50, -65), PvcRemain, "pvc remain");
        Check(errors, Add(10, -3), ManholeRemain, "manhole remain");
        Check(errors, Add(PvcRemain, 20, -5), PvcRemainAfterAdd, "pvc after add");
        Check(errors, Add(Jan, Feb, Mar), Cum, "cum");
        if (Math.Abs(Cum / ContractAmount * 100 - CumPct) > 1e-9) errors.Add("cum pct");
        Check(errors, ContractAmount - Cum, Remain, "remain");
        Check(errors, Add(Jan, FebChanged, Mar), CumChanged, "cum changed");
        Check(errors, MarActualChanged - 40, MarDiffChanged, "mar diff");
        Check(errors, Add(1_200_000, 800_000, 500_000), 2_500_000, "e7 pvc");
        Check(errors, Add(420_000, 210_000), 630_000, "e7 manhole");
        Check(errors, Add(2_500_000, 630_000), 3_130_000, "e7 grand");
        if (GoldenXml.NormalizePartPath("/xl/worksheets/sheet1.xml") != "xl/worksheets/sheet1.xml")
            errors.Add("package path absolute");
        if (GoldenXml.NormalizePartPath("worksheets/sheet1.xml") != "xl/worksheets/sheet1.xml")
            errors.Add("package path relative");
        if (GoldenXml.NormalizePartPath("xl/worksheets/sheet1.xml") != "xl/worksheets/sheet1.xml")
            errors.Add("package path already-xl");

        var fixtures = Path.Combine(AppContext.BaseDirectory, "Fixtures", "inventory");
        var inventory = IdentityGuard.VerifyPublicFixtures(fixtures);
        if (JsonUtil.Bool(inventory, "ok") != true)
        {
            foreach (var err in (JsonUtil.Get(inventory, "errors") as JsonArray) ?? [])
                errors.Add(err?.ToString() ?? "inventory fixture");
        }

        var rect = StyleReconstructor.CoalesceRectangles([(1, 1), (1, 2), (2, 1), (2, 2)]);
        if (rect.Count != 1 || rect[0] != "A1:B2")
            errors.Add("style rectangles must coalesce a 2x2 block to A1:B2, not per-row runs");
        var broken = StyleReconstructor.CoalesceRectangles([(1, 1), (1, 2), (2, 1)]);
        if (broken.Count != 2)
            errors.Add("L-shaped equal-style cells must stay two rectangles");
        var planner = StyleReconstructor.PlannerSelfCheck();
        if (JsonUtil.Bool(planner, "ok") != true)
        {
            foreach (var err in (JsonUtil.Get(planner, "errors") as JsonArray) ?? [])
                errors.Add(err?.ToString() ?? "property-layer planner");
        }

        var goodPage = PagePayload.E1Acceptance(new JsonObject
        {
            ["paperSize"] = "8",
            ["orientation"] = "landscape",
            ["centerFooter"] = "&\"돋움\"&8 수원시 하수관로",
            ["marginsMm"] = new JsonObject { ["left"] = 10.0 },
        });
        foreach (var err in PagePayload.HygieneErrors(goodPage))
            errors.Add("typed page hygiene rejected a valid payload: " + err);
        if (goodPage.ContainsKey("oddHeader") || goodPage.ContainsKey("oddFooter")
            || goodPage.ContainsKey("fitToWidth") || goodPage.ContainsKey("fitToHeight"))
            errors.Add("E1 page payload must omit oddHeader/oddFooter and fit keys");
        if (goodPage["scale"] is not JsonValue sv || !sv.TryGetValue<int>(out var scale) || scale != 55)
            errors.Add("E1 page.scale must be integer 55");

        var dirtyPage = new JsonObject
        {
            ["scale"] = 55.0,
            ["fitToWidth"] = 1,
            ["oddFooter"] = "x",
            ["centerFooter"] = "&Cfooter",
        };
        if (PagePayload.HygieneErrors(dirtyPage).Count < 3)
            errors.Add("page hygiene must refuse double scale, fit keys, oddFooter, and &C footer");

        var dirtyLayout = new JsonArray
        {
            new JsonObject { ["col"] = "A", ["widthChars"] = 3.4, ["unit"] = "com-columnwidth", ["note"] = "oracle" },
        };
        if (PublicLayout.HasOracleMetadata(dirtyLayout))
        {
            var stripped = PublicLayout.Columns(dirtyLayout);
            if (PublicLayout.HasOracleMetadata(stripped) || stripped.Count != 1
                || ((JsonObject)stripped[0]!).ContainsKey("unit") || ((JsonObject)stripped[0]!).ContainsKey("note"))
                errors.Add("public column items must drop oracle unit/note");
        }

        var blank = BlankCheckpoint.SelfCheck();
        if (JsonUtil.Bool(blank, "ok") != true)
        {
            foreach (var err in (JsonUtil.Get(blank, "errors") as JsonArray) ?? [])
                errors.Add(err?.ToString() ?? "blank checkpoint");
        }

        var layout = PracticalLayoutMetrics.SelfCheck();
        if (JsonUtil.Bool(layout, "ok") != true)
        {
            foreach (var err in (JsonUtil.Get(layout, "errors") as JsonArray) ?? [])
                errors.Add(err?.ToString() ?? "practical layout");
        }

        var csv = PracticalCsvRecordParser.SelfCheck();
        if (JsonUtil.Bool(csv, "ok") != true)
        {
            foreach (var err in (JsonUtil.Get(csv, "errors") as JsonArray) ?? [])
                errors.Add(err?.ToString() ?? "csv oracle");
        }

        Check(errors, Add(120, 10, 50, 20), 200, "e4 inbound after TX-004");
        Check(errors, Add(40, 3, 65, 5), 113, "e4 outbound after TX-004");
        Check(errors, 200 - 113, 87, "e4 net after TX-004");
        Check(errors, Add(200, 30), 230, "e4 inbound after TX-005");
        Check(errors, Add(113, 8), 121, "e4 outbound after TX-005");
        Check(errors, Add(230, -50), 180, "e4 inbound after delete TX-003");
        Check(errors, Add(121, -65), 56, "e4 outbound after delete TX-003");

        try
        {
            ApplyPlanner.ValidateObjectLifecycle("self-create-delete", JsonUtil.Arr(
                new JsonObject { ["op"] = "create_chart", ["name"] = "ScratchChart", ["target"] = JsonUtil.Target("S") },
                new JsonObject { ["op"] = "delete_chart", ["name"] = "ScratchChart", ["target"] = JsonUtil.Target("S") }));
            errors.Add("lifecycle must refuse create+delete of the same object in one batch");
        }
        catch (InvalidOperationException)
        {
            /* expected */
        }

        if (ContractCatalog.FamilyOf("set_view") != OpFamily.SheetLayout)
            errors.Add("set_view must be SheetLayout, not Visibility");
        if (ContractCatalog.FamilyOf("set_outline") != OpFamily.RangeEdit)
            errors.Add("set_outline must be RangeEdit, not Visibility");
        try
        {
            ApplyPlanner.ValidateFamily(OpFamily.Visibility, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "set_view",
                ["target"] = JsonUtil.Target("S"),
                ["zoom"] = 100,
            }));
            errors.Add("visibility-only batch must refuse set_view");
        }
        catch (InvalidOperationException)
        {
            /* expected */
        }

        return new JsonObject
        {
            ["ok"] = errors.Count == 0,
            ["errors"] = errors,
        };
    }

    internal readonly record struct CostLine(string Name, double Qty, double Unit, double Amount);

    private static double Add(params double[] xs)
    {
        var t = 0d;
        foreach (var x in xs) t += x;
        return t;
    }

    private static void Check(JsonArray errors, double actual, double expected, string name)
    {
        if (Math.Abs(actual - expected) > 1e-9)
            errors.Add($"{name}: {actual} != {expected}");
    }
}
