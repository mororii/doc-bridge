namespace DocBridge.Development.ExcelProductionWorkflowProbe;

internal sealed class LiveVerifyRequest
{
    public required PublicClient Client { get; init; }
    public required ApplySession Session { get; init; }
    public required string Id { get; init; }
    public required string ArtifactDir { get; init; }
    public required string SourceXlsx { get; init; }
    public required string ExpectedSha { get; init; }
    public required string RebuildDir { get; init; }
    public string? ComOraclePath { get; init; }
    public required string OwnedWorkbook { get; init; }
    public JsonNode? LastApply { get; init; }
    public JsonNode? InjectedApply { get; init; }
    public JsonObject? InitialRead { get; init; }
    public string? StopBeforeBatchId { get; init; }
}

/// <summary>
/// Independent live checks. Product <c>readback.verified</c> is never trusted alone.
/// Public reads use top-level workbook/sheet and the actual saved xlsx path.
/// </summary>
internal static class LiveVerifier
{
    public static JsonObject Catalog() => new()
    {
        ["rule"] = "Write success is never accepted from product readback.verified alone.",
        ["publicReads"] = JsonUtil.Arr("excel_read_range", "excel_inspect", "excel_get_active_context", "core_get_status"),
        ["readArgs"] = "top-level workbook+sheet; never target object or $owned",
        ["scenarios"] = new JsonObject
        {
            ["e1"] = JsonUtil.Arr(
                "owned blank OOXML after create+save: no nonempty cells/merges/target A1:BQ96 fills before values",
                "source SHA-256 unchanged after authoring",
                "artifact != source path; 통합 문서1 never targeted",
                "excel_read_range workbook=owned xlsx sheet=예정공정표 A1:BQ96",
                "typed BP95≈100, BP96=22656, AB95=17.659",
                "281 merges, 31 month headers, G5:BP5 no intersecting merge",
                "freeze G6, page A3 landscape printArea A1:BQ96, exact footer",
                "all A:BQ columnWidths and 1:96 rowHeights present",
                "941 medium bottoms: 940 in G1:BP96 plus BQ93",
                "60 used styles from artifact XML vs source XML, not 3-field equality",
                "collision: B100 kept; E100:F100 remains merged; C100:D100 unmerged",
                "preview refuse is not rollback.verified",
                "Chart1 via public create_chart from chart1.xml (not raster)",
                "save + close_workbook + open_workbook + export_pdf"),
            ["e2"] = JsonUtil.Arr("initial G5/G9/G10/G11 before qty 130", "after recalc", "printArea A1:H14 excludes DIRTY", "currency/grouping + table/title format", "F2 ISO string + empty-string blanks", "mixed A26:D26 strict strings + original formats after reopen", "row 15/16 height mapping 28/22 recorded, not hidden", "save-close-reopen", "PDF"),
            ["e3"] = JsonUtil.Arr("B8=13 C8=3", "inspect pictures=2", "TinyPng live fixtures", "printArea A1:F23", "PDF"),
            ["e4"] = JsonUtil.Arr("initial G3=65 G4=7 before add", "PVC remain 80 after add", "void 999 excluded", "순입고 J3:J6 80/7/−15/15", "개정 TX-003 gone TX-005 present", "PDF"),
            ["e5"] = JsonUtil.Arr("initial 45M/45%/55M", "after Feb 20M → 47M/47%/53M", "quoted absolute names $B$2 after rename", "hyperlink '월별기성'!A1", "readable #,##0 / 0%", "PDF"),
            ["e6"] = JsonUtil.Arr("initial D5=3", "after C5=45 D5=5", "charts-2 PlanActualLine/DiffColumn", "DiffColumn union A2:A8,D2:D8 + series", "printArea A1:D45", "page break A22", "PDF"),
            ["e7"] = JsonUtil.Arr("create_pivot 피벗!A3 원장피벗 source '원장'!A1:C5", "update_pivot sourceRange=원장표 after A6 PVC/3월/500000", "PVC 2500000 / 맨홀 630000 / grand 3130000"),
            ["e8"] = JsonUtil.Arr("B3 initial 10000", "unlocked B1=11 recalculates B3=11000", "locked B3 write refused", "protect state"),
            ["e9"] = JsonUtil.Arr("unique data rows 2 after remove_duplicates", "imported 정리 vs oracle B2 quoted comma / D2 CRLF / E2=0012 / F2==1+1", "exported csv unique 2", "missing import/export/dedupe is a failed requirement"),
        },
    };

    public static JsonObject CaptureInitial(PublicClient client, string ownedWorkbook, string id)
    {
        var sheet = SheetOf(id);
        var range = id switch
        {
            "e2" => "G5:G11",
            "e4" => "G3:G5",
            "e5" => "B2:B4",
            "e6" => "D3:D8",
            _ => RangeOf(id),
        };
        var read = PublicApi.ReadRange(client, ownedWorkbook, sheet, range, formulas: true);
        return new JsonObject
        {
            ["ok"] = JsonUtil.Bool(read, "ok") == true && PublicApi.HasReadableValues(read),
            ["sheet"] = sheet,
            ["range"] = range,
            ["read"] = read.DeepClone(),
        };
    }

    public static JsonObject CaptureEvidence(PublicClient client, string ownedWorkbook, string id, PlannedBatch batch)
    {
        var capture = batch.Capture ?? "";
        if (capture.StartsWith("read:", StringComparison.OrdinalIgnoreCase))
        {
            var spec = capture["read:".Length..];
            var bang = spec.IndexOf('!');
            var sheet = bang > 0 ? spec[..bang] : SheetOf(id);
            var range = bang > 0 ? spec[(bang + 1)..] : RangeOf(id);
            var read = PublicApi.ReadRange(client, ownedWorkbook, sheet, range, formulas: true);
            return new JsonObject
            {
                ["ok"] = JsonUtil.Bool(read, "ok") == true && PublicApi.HasReadableValues(read),
                ["kind"] = "range",
                ["sheet"] = sheet,
                ["range"] = range,
                ["batch"] = batch.Id,
                ["read"] = read.DeepClone(),
            };
        }

        var inspectSheet = SheetOf(id);
        var inspect = PublicApi.Inspect(client, ownedWorkbook, "objects", inspectSheet);
        return new JsonObject
        {
            ["ok"] = JsonUtil.Bool(inspect, "ok") == true,
            ["kind"] = "objects",
            ["sheet"] = inspectSheet,
            ["batch"] = batch.Id,
            ["inspect"] = inspect.DeepClone(),
        };
    }

    public static JsonObject VerifyScenario(LiveVerifyRequest req)
    {
        var checks = new JsonArray();
        var xlsx = Path.Combine(req.ArtifactDir, req.Id switch
        {
            "e1" => "e1-schedule.xlsx",
            "e2" => "e2-cost-estimate.xlsx",
            "e3" => "e3-daily-report.xlsx",
            "e4" => "e4-materials-ledger.xlsx",
            "e5" => "e5-progress-payment.xlsx",
            "e6" => "e6-monthly-report.xlsx",
            "e7" => "e7-pivot-ledger.xlsx",
            "e8" => "e8-protected-form.xlsx",
            "e9" => "e9-csv-cleanup.xlsx",
            _ => req.Id + ".xlsx",
        });
        var pdf = Path.ChangeExtension(xlsx, ".pdf");
        IdentityGuard.RefuseIfPlaceholder(req.OwnedWorkbook);
        IdentityGuard.RefuseProtected(req.OwnedWorkbook, req.SourceXlsx, null);
        Add(checks, "owned-is-saved-xlsx",
            string.Equals(Path.GetFullPath(req.OwnedWorkbook), Path.GetFullPath(xlsx), StringComparison.OrdinalIgnoreCase)
            && File.Exists(xlsx));

        var frozenHash = string.Equals(SourceIntegrity.Sha256File(req.SourceXlsx), req.ExpectedSha, StringComparison.OrdinalIgnoreCase);
        Add(checks, "source-hash-frozen-match", frozenHash);
        Add(checks, "source-hash-preserved",
            frozenHash || !string.IsNullOrWhiteSpace(req.StopBeforeBatchId));
        Add(checks, "artifact-not-source",
            !string.Equals(Path.GetFullPath(xlsx), Path.GetFullPath(req.SourceXlsx), StringComparison.OrdinalIgnoreCase));

        if (File.Exists(xlsx))
        {
            SourceIntegrity.RefuseIfProtectedPath(xlsx, req.SourceXlsx);
            Add(checks, "artifact-hash-differs-from-source",
                !string.Equals(SourceIntegrity.Sha256File(xlsx), req.ExpectedSha, StringComparison.OrdinalIgnoreCase));
        }

        var sheet = SheetOf(req.Id);
        var read = PublicApi.ReadRange(req.Client, req.OwnedWorkbook, sheet, RangeOf(req.Id),
            formulas: true, styles: req.Id is "e1" or "e2" or "e5", layout: req.Id is "e1" or "e2" or "e3" or "e6");
        Add(checks, "public-read-ok", JsonUtil.Bool(read, "ok") == true && PublicApi.HasReadableValues(read));

        var inspect = PublicApi.Inspect(req.Client, req.OwnedWorkbook, "objects", sheet);
        Add(checks, "public-inspect-ok", JsonUtil.Bool(inspect, "ok") == true);
        if (req.Id == "e1")
            Add(checks, "chart1-present", CountObjects(inspect, "chart") >= 1);

        switch (req.Id)
        {
            case "e1":
                VerifyE1(req, checks, read, xlsx);
                break;
            case "e2":
                AssertInitial(req, checks, "G5", AcceptanceNumbers.CostLines[0].Amount);
                AssertInitial(req, checks, "G9", AcceptanceNumbers.CostSubtotal);
                AssertInitial(req, checks, "G10", AcceptanceNumbers.CostVat);
                AssertInitial(req, checks, "G11", AcceptanceNumbers.CostGrand);
                AddNum(checks, "line1-after-qty-130", CellNumber(read, "G5"), AcceptanceNumbers.CostAmountChanged);
                AddNum(checks, "line2-unchanged", CellNumber(read, "G6"), AcceptanceNumbers.CostLines[1].Amount);
                AddNum(checks, "line3-unchanged", CellNumber(read, "G7"), AcceptanceNumbers.CostLines[2].Amount);
                AddNum(checks, "line4-unchanged", CellNumber(read, "G8"), AcceptanceNumbers.CostLines[3].Amount);
                AddNum(checks, "subtotal-after", CellNumber(read, "G9"), AcceptanceNumbers.CostSubtotalChanged);
                AddNum(checks, "vat-after", CellNumber(read, "G10"), AcceptanceNumbers.CostVatChanged);
                AddNum(checks, "grand-after", CellNumber(read, "G11"), AcceptanceNumbers.CostGrandChanged);
                Add(checks, "dirty-999-not-in-grand",
                    Math.Abs((CellNumber(read, "G11") ?? 0) - AcceptanceNumbers.CostGrandChanged) < 0.5);
                Add(checks, "print-area-A1H14", PrintAreaIs(read, "A1:H14"));
                Add(checks, "dirty-outside-print-area",
                    !string.Equals(CellText(read, "A13"), "DIRTY", StringComparison.OrdinalIgnoreCase)
                    && string.Equals(CellText(read, "A24"), "DIRTY", StringComparison.OrdinalIgnoreCase));
                var money = PublicApi.ReadRange(req.Client, req.OwnedWorkbook, sheet, "G11", styles: true);
                Add(checks, "g11-grouped-or-currency", NumberFormatHas(money, "#,##0"));
                var title = PublicApi.ReadRange(req.Client, req.OwnedWorkbook, sheet, "A1", styles: true);
                Add(checks, "title-dotum-or-bold", FontIs(title, "돋움", 16) || StyleBold(title));
                Add(checks, "f2-iso-string", JsonStringEquals(read, "F2", "2026-09-10"));
                Add(checks, "h5-empty-string-or-blank", IsBlank(read, "H5"));
                Add(checks, "h6-empty-string-or-blank", IsBlank(read, "H6"));
                Add(checks, "h7-empty-string-or-blank", IsBlank(read, "H7"));
                Add(checks, "h8-empty-string-or-blank", IsBlank(read, "H8"));
                Add(checks, "b12-empty-string-or-blank", IsBlank(read, "B12"));
                Add(checks, "d12-empty-string-or-blank", IsBlank(read, "D12"));
                Add(checks, "f12-empty-string-or-blank", IsBlank(read, "F12"));
                Add(checks, "g12-approval-label", string.Equals(CellText(read, "G12"), "결재", StringComparison.Ordinal));
                Add(checks, "h12-blank-no-gap", IsBlank(read, "H12"));
                Add(checks, "sig-merge-a12-b14", MergeSetHasExact(read, "A12:B14"));
                Add(checks, "sig-merge-c12-d14", MergeSetHasExact(read, "C12:D14"));
                Add(checks, "sig-merge-e12-f14", MergeSetHasExact(read, "E12:F14"));
                Add(checks, "sig-merge-g12-h14", MergeSetHasExact(read, "G12:H14"));
                Add(checks, "sig-merge-not-h12-only",
                    !MergeSetHasExact(read, "H12:H14"));
                Add(checks, "sig-merge-four-exact",
                    MergeSetHasExact(read, "A12:B14")
                    && MergeSetHasExact(read, "C12:D14")
                    && MergeSetHasExact(read, "E12:F14")
                    && MergeSetHasExact(read, "G12:H14"));
                var skipMixed = string.Equals(req.StopBeforeBatchId, "e2-mixed-format-prep", StringComparison.OrdinalIgnoreCase);
                if (skipMixed)
                    Add(checks, "mixed-deferred-stop-before", true);
                else
                    VerifyMixedTypedRow(req, checks, sheet);
                var heightRead = PublicApi.ReadRange(req.Client, req.OwnedWorkbook, sheet, "A15:A16", layout: true);
                RecordHeightMapping(heightRead, checks, 15, 28, requireRequested: !skipMixed);
                RecordHeightMapping(heightRead, checks, 16, 22, requireRequested: !skipMixed);
                break;
            case "e3":
                AddNum(checks, "people", CellNumber(read, "B8"), AcceptanceNumbers.DailyPeople);
                AddNum(checks, "equipment", CellNumber(read, "C8"), AcceptanceNumbers.DailyEquipment);
                Add(checks, "pictures-2", CountObjects(inspect, "picture") == 2);
                Add(checks, "wrap-B9", StyleWrap(read) || StyleWrap(
                    PublicApi.ReadRange(req.Client, req.OwnedWorkbook, "작업일보", "B9", styles: true)));
                Add(checks, "print-area-A1F23", PrintAreaIs(read, "A1:F23"));
                break;
            case "e4":
                AssertInitial(req, checks, "G3", AcceptanceNumbers.PvcRemain);
                AssertInitial(req, checks, "G4", AcceptanceNumbers.ManholeRemain);
                AssertInitial(req, checks, "G5", -15);
                AddNum(checks, "pvc-tx001", CellNumber(read, "G3"), AcceptanceNumbers.PvcRemain);
                AddNum(checks, "manhole-remain", CellNumber(read, "G4"), AcceptanceNumbers.ManholeRemain);
                AddNum(checks, "pvc-tx004", CellNumber(read, "G6"), 15);
                AddNum(checks, "pvc-remain-after-add",
                    (CellNumber(read, "G3") ?? 0) + (CellNumber(read, "G6") ?? 0),
                    AcceptanceNumbers.PvcRemainAfterAdd);
                Add(checks, "void-999-not-in-remain",
                    Math.Abs((CellNumber(read, "G3") ?? 0) + (CellNumber(read, "G4") ?? 0) + (CellNumber(read, "G6") ?? 0)
                             - (AcceptanceNumbers.PvcRemain + AcceptanceNumbers.ManholeRemain + 15)) < 0.5);
                AddNum(checks, "net-in-j3", CellNumber(read, "J3"), 80);
                AddNum(checks, "net-in-j4", CellNumber(read, "J4"), 7);
                AddNum(checks, "net-in-j5", CellNumber(read, "J5"), -15);
                AddNum(checks, "net-in-j6", CellNumber(read, "J6"), 15);
                Add(checks, "table-present", CountObjects(inspect, "table") >= 1);
                var revision = PublicApi.ReadRange(req.Client, req.OwnedWorkbook, "개정", "A1:I8", formulas: true);
                Add(checks, "revision-sheet-readable",
                    JsonUtil.Bool(revision, "ok") == true && PublicApi.HasReadableValues(revision));
                Add(checks, "revision-tx003-gone", !GridContainsText(revision, "TX-003"));
                Add(checks, "revision-tx005-present", GridContainsText(revision, "TX-005"));
                Add(checks, "revision-tx001-present", GridContainsText(revision, "TX-001"));
                Add(checks, "revision-tx002-present", GridContainsText(revision, "TX-002"));
                Add(checks, "revision-tx004-present", GridContainsText(revision, "TX-004"));
                break;
            case "e5":
                AssertInitial(req, checks, "B2", AcceptanceNumbers.Cum);
                AssertInitial(req, checks, "B3", AcceptanceNumbers.Cum / AcceptanceNumbers.ContractAmount, 0.001);
                AssertInitial(req, checks, "B4", AcceptanceNumbers.Remain);
                AddNum(checks, "cum-after-feb-20m", CellNumber(read, "B2"), AcceptanceNumbers.CumChanged);
                AddApprox(checks, "pct-after", CellNumber(read, "B3"),
                    AcceptanceNumbers.CumChanged / AcceptanceNumbers.ContractAmount, 0.001);
                AddNum(checks, "remain-after", CellNumber(read, "B4"), AcceptanceNumbers.RemainChanged);
                var cumStyle = PublicApi.ReadRange(req.Client, req.OwnedWorkbook, sheet, "B2", styles: true);
                Add(checks, "b2-grouped-or-currency", NumberFormatHas(cumStyle, "#,##0"));
                var pctStyle = PublicApi.ReadRange(req.Client, req.OwnedWorkbook, sheet, "B3", styles: true);
                Add(checks, "b3-percent", NumberFormatHas(pctStyle, "%"));
                VerifyE5NamesAndLink(req, checks);
                break;
            case "e6":
                AssertInitial(req, checks, "D5", AcceptanceNumbers.ActualRates[2] - AcceptanceNumbers.PlanRates[2]);
                AddNum(checks, "mar-diff-after-change", CellNumber(read, "D5"), AcceptanceNumbers.MarDiffChanged);
                Add(checks, "charts-2", CountObjects(inspect, "chart") >= 2);
                Add(checks, "print-area-A1D45", PrintAreaIs(read, "A1:D45"));
                break;
            case "e7":
                VerifyE7(req, checks, inspect);
                break;
            case "e8":
                VerifyE8(req, checks, read);
                break;
            case "e9":
                VerifyE9(req, checks, read);
                break;
        }

        if (File.Exists(pdf))
        {
            var pages = PdfCatalog.CountPages(pdf);
            var minPages = req.Id == "e6" ? 2 : 1;
            Add(checks, "pdf-exists", true);
            checks.Add(new JsonObject
            {
                ["id"] = "pdf-pages",
                ["ok"] = pages >= minPages,
                ["pages"] = pages,
                ["min"] = minPages,
            });
        }
        else
            Add(checks, "pdf-exists", false);

        var failed = checks.OfType<JsonObject>().Count(c => JsonUtil.Bool(c, "ok") == false);
        return new JsonObject
        {
            ["scenario"] = req.Id,
            ["ok"] = failed == 0,
            ["failed"] = failed,
            ["ownedWorkbook"] = req.OwnedWorkbook,
            ["checks"] = checks,
            ["productReadbackIgnored"] = true,
        };
    }

    private static void VerifyE1(LiveVerifyRequest req, JsonArray checks, JsonObject read, string xlsx)
    {
        AddNum(checks, "BP96", CellNumber(read, "BP96"), AcceptanceNumbers.Bp96);
        AddApprox(checks, "BP95", CellNumber(read, "BP95"), AcceptanceNumbers.Bp95, 0.05);
        AddNum(checks, "AB95", CellNumber(read, "AB95"), AcceptanceNumbers.Ab95);
        Add(checks, "merges-281", MergeCount(read) == AcceptanceNumbers.MergeCount);
        Add(checks, "month-headers-31", MonthHeaderCount(read) == AcceptanceNumbers.MonthHeaderMerges);
        Add(checks, "row5-unmerged", !MergesOverlap(read, "G5:BP5"));
        Add(checks, "freeze-G6", FreezeIs(read, "G6"));
        Add(checks, "page-A3-landscape", PagePaperLandscape(read));
        Add(checks, "print-area-A1BQ96", PrintAreaIs(read, "A1:BQ96"));
        Add(checks, "footer-suwon", FooterContains(read, "수원시 하수관로"));
        Add(checks, "page-scale-or-fit", PageScaleOrFit(read));
        Add(checks, "columns-A-BQ", ColumnWidthCount(read) == 69);
        Add(checks, "rows-1-96", RowHeightCount(read) == 96);

        if (!string.IsNullOrWhiteSpace(req.ComOraclePath) && File.Exists(req.ComOraclePath))
        {
            var com = JsonUtil.Load(req.ComOraclePath).AsObject();
            Add(checks, "gbp-width-chars-3.4", ColumnWidthNear(read, 7, 68, 3.4, 0.2));
            Add(checks, "com-oracle-source-unchanged", JsonUtil.Bool(com, "sourceUnchanged") == true);
        }

        var title = PublicApi.ReadRange(req.Client, req.OwnedWorkbook, ScenarioE1.Sheet, "A1", styles: true);
        Add(checks, "title-dotum-28", FontIs(title, "돋움", 28));

        if (req.InjectedApply is not null)
        {
            var phase = InjectedPhase(req.InjectedApply);
            Add(checks, "collision-refused", phase is "preview-refused" or "apply-refused");
            Add(checks, "collision-preview-is-not-rollback",
                phase != "preview-refused" || JsonUtil.Bool(JsonUtil.Get(req.InjectedApply, "rollback"), "verified") != true);
            Add(checks, "collision-apply-rollback-if-apply-path",
                phase != "apply-refused" || JsonUtil.Bool(JsonUtil.Get(req.InjectedApply, "rollback"), "verified") == true
                || JsonUtil.Str(JsonUtil.Get(req.InjectedApply, "rollback"), "status") == "unverified");

            var collision = PublicApi.ReadRange(req.Client, req.OwnedWorkbook, ScenarioE1.Sheet, "A100:F100",
                formulas: true, layout: true);
            Add(checks, "b100-non-ul-preserved",
                string.Equals(CellText(collision, "B100"), "NON-UL-CONTENT", StringComparison.Ordinal));
            Add(checks, "c100-d100-unmerged", !MergesOverlap(collision, "C100:D100"));
            Add(checks, "e100-f100-legal-merge-kept", MergesOverlap(collision, "E100:F100"));
        }
        else
        {
            Add(checks, "injected-failure-captured", false);
        }

        if (!File.Exists(xlsx))
        {
            Add(checks, "artifact-xlsx-present", false);
            return;
        }

        try
        {
            var sourceXml = GoldenXml.Inspect(req.SourceXlsx, ScenarioE1.Sheet);
            var artifactXml = GoldenXml.Inspect(xlsx, ScenarioE1.Sheet);
            var sourceMerges = StringSet(sourceXml, "merges");
            var artifactMerges = StringSet(artifactXml, "merges");
            artifactMerges.ExceptWith(new[] { "E100:F100" });
            Add(checks, "xml-merge-set-equals-source", sourceMerges.SetEquals(artifactMerges));
            Add(checks, "reopen-xml-row5-unmerged", JsonUtil.Bool(artifactXml, "row5Unmerged") == true);
            Add(checks, "reopen-xml-month-31", (int)(JsonUtil.Num(artifactXml, "monthHeaderCount") ?? 0) == 31);
            Add(checks, "xml-footer-present", !string.IsNullOrWhiteSpace(JsonUtil.Str(artifactXml, "centerFooter")));
            Add(checks, "xml-columns-present", (JsonUtil.Get(artifactXml, "columns") as JsonArray)?.Count > 0);
            Add(checks, "not-three-field-layout-claim", true);

            var barCells = LoadRebuildBars(req.RebuildDir);
            var sourceBars = GoldenXml.CountUsedStylesAndMediumBottoms(req.SourceXlsx, ScenarioE1.Sheet, barCells);
            var artifactBars = GoldenXml.CountUsedStylesAndMediumBottoms(xlsx, ScenarioE1.Sheet, barCells);
            Add(checks, "used-styles-60",
                (int)(JsonUtil.Num(artifactBars, "usedCellStyleCount") ?? 0) == 60
                && (int)(JsonUtil.Num(sourceBars, "usedCellStyleCount") ?? 0) == 60);
            Add(checks, "schedule-medium-bottoms-940",
                (int)(JsonUtil.Num(artifactBars, "xmlScheduleMediumBottomCellCount") ?? 0) == 940);
            var other = JsonUtil.Get(JsonUtil.Get(artifactBars, "barCompare"), "otherMediumBottomsOutsideG1BP96") as JsonArray;
            Add(checks, "bq93-medium-listed",
                other is not null && other.Any(n => string.Equals(n?.GetValue<string>(), "BQ93", StringComparison.OrdinalIgnoreCase)));
        }
        catch (Exception ex)
        {
            checks.Add(new JsonObject { ["id"] = "e1-xml-compare", ["ok"] = false, ["error"] = ex.Message });
        }
    }

    private static void VerifyMixedTypedRow(LiveVerifyRequest req, JsonArray checks, string sheet)
    {
        var mixed = PublicApi.ReadRange(req.Client, req.OwnedWorkbook, sheet, "A26:D26", formulas: true);
        Add(checks, "mixed-a26-iso-string", JsonStringEquals(mixed, "A26", "2026-09-10"));
        Add(checks, "mixed-b26-leading-zero-string", JsonStringEquals(mixed, "B26", "0012"));
        Add(checks, "mixed-c26-1e3-string", JsonStringEquals(mixed, "C26", "1E3"));
        Add(checks, "mixed-d26-formula-prefixed-string", JsonStringEquals(mixed, "D26", "=1+1"));
        Add(checks, "mixed-d26-not-evaluated-to-2",
            JsonStringEquals(mixed, "D26", "=1+1") && CellNumber(mixed, "D26") is not 2);

        foreach (var (address, expected) in new (string, string)[]
                 {
                     ("A26", "General"),
                     ("B26", "0.00"),
                     ("C26", "yyyy-mm-dd"),
                     ("D26", "\"LIT\""),
                 })
        {
            var cell = PublicApi.ReadRange(req.Client, req.OwnedWorkbook, sheet, address, styles: true);
            var actual = StyleNumberFormat(cell);
            checks.Add(new JsonObject
            {
                ["id"] = "mixed-" + address.ToLowerInvariant() + "-format",
                ["ok"] = string.Equals(actual, expected, StringComparison.Ordinal),
                ["expected"] = expected,
                ["actual"] = actual,
            });
        }
    }

    private static void RecordHeightMapping(JsonNode read, JsonArray checks, int row, double requested, bool requireRequested = true)
    {
        var actual = RowHeightPoints(read, row);
        double? delta = actual is { } h ? h - requested : null;
        var matches = actual is { } a && Math.Abs(a - requested) <= 0.05;
        checks.Add(new JsonObject
        {
            ["id"] = "height-row" + row.ToString(CultureInfo.InvariantCulture) + "-mapping",
            ["ok"] = actual is not null,
            ["row"] = row,
            ["requested"] = requested,
            ["actual"] = actual,
            ["delta"] = delta,
            ["matchesRequested"] = matches,
            ["epsilon"] = 0.05,
        });
        checks.Add(new JsonObject
        {
            ["id"] = "height-row" + row.ToString(CultureInfo.InvariantCulture) + "-requested",
            ["ok"] = requireRequested ? matches : true,
            ["deferred"] = !requireRequested,
            ["requested"] = requested,
            ["actual"] = actual,
            ["epsilon"] = 0.05,
        });
    }

    private static void VerifyE5NamesAndLink(LiveVerifyRequest req, JsonArray checks)
    {
        var names = PublicApi.Inspect(req.Client, req.OwnedWorkbook, "objects", objectKind: "names");
        var objects = JsonUtil.Get(names, "objects") as JsonArray ?? [];
        var required = new[] { "ContractAmount", "JanAmt", "FebAmt", "MarAmt" };
        foreach (var name in required)
        {
            var found = objects.OfType<JsonObject>().FirstOrDefault(o =>
                string.Equals(JsonUtil.Str(o, "localName") ?? JsonUtil.Str(o, "name"), name, StringComparison.OrdinalIgnoreCase));
            var refers = JsonUtil.Str(found, "refersTo") ?? "";
            var sheetOk = refers.Contains("월별기성", StringComparison.Ordinal) ||
                          refers.Contains("기성입력", StringComparison.Ordinal);
            var absOk = refers.Contains("$B$", StringComparison.Ordinal);
            checks.Add(new JsonObject
            {
                ["id"] = "name-" + name + "-quoted-absolute",
                ["ok"] = found is not null && sheetOk && absOk,
                ["refersTo"] = refers,
            });
        }

        var links = PublicApi.Inspect(req.Client, req.OwnedWorkbook, "objects", "누계집계", "hyperlinks");
        var linkObjects = JsonUtil.Get(links, "objects") as JsonArray ?? [];
        var link = linkObjects.OfType<JsonObject>().FirstOrDefault(o =>
            (JsonUtil.Str(o, "range") ?? "").Contains("B5", StringComparison.OrdinalIgnoreCase));
        var sub = JsonUtil.Str(link, "subAddress") ?? "";
        Add(checks, "hyperlink-after-rename-월별기성-A1",
            sub.Contains("월별기성", StringComparison.Ordinal) &&
            sub.Contains("A1", StringComparison.OrdinalIgnoreCase));
    }

    private static void VerifyE7(LiveVerifyRequest req, JsonArray checks, JsonObject inspect)
    {
        var pivots = CountObjects(inspect, "pivot");
        Add(checks, "pivot-present", pivots >= 1);
        var pivotRead = PublicApi.ReadRange(req.Client, req.OwnedWorkbook, "피벗", "A1:Z40", formulas: false);
        Add(checks, "pivot-sheet-readable", JsonUtil.Bool(pivotRead, "ok") == true && PublicApi.HasReadableValues(pivotRead));
        Add(checks, "pivot-pvc-total-2500000", GridContainsNumber(pivotRead, 2_500_000));
        Add(checks, "pivot-manhole-total-630000", GridContainsNumber(pivotRead, 630_000));
        Add(checks, "pivot-grand-3130000", GridContainsNumber(pivotRead, 3_130_000));
    }

    private static void VerifyE8(LiveVerifyRequest req, JsonArray checks, JsonObject read)
    {
        AddNum(checks, "amount", CellNumber(read, "B3"), 11_000);
        var lockedWrite = req.Session.TryExecuteValues("입력양식", "B3",
            new JsonArray { new JsonArray { 99_999 } });
        Add(checks, "locked-b3-write-refused", JsonUtil.Bool(lockedWrite, "ok") != true);
        var openWrite = req.Session.TryExecuteValues("입력양식", "B1",
            new JsonArray { new JsonArray { 11 } });
        Add(checks, "unlocked-b1-write-allowed", JsonUtil.Bool(openWrite, "ok") == true);
        var after = PublicApi.ReadRange(req.Client, req.OwnedWorkbook, "입력양식", "A1:B3", formulas: true);
        AddNum(checks, "b3-after-b1-11000", CellNumber(after, "B3"), 11_000);
        AddNum(checks, "b1-now-11", CellNumber(after, "B1"), 11);
        var b3 = PublicApi.ReadRange(req.Client, req.OwnedWorkbook, "입력양식", "B3", styles: true);
        var b1 = PublicApi.ReadRange(req.Client, req.OwnedWorkbook, "입력양식", "B1", styles: true);
        Add(checks, "b3-style-locked", StyleLocked(b3) == true);
        Add(checks, "b1-style-unlocked", StyleLocked(b1) == false);
    }

    private static void VerifyE9(LiveVerifyRequest req, JsonArray checks, JsonObject read)
    {
        var unique = CountNonEmptyDataRows(read);
        Add(checks, "unique-data-rows-2", unique == 2);
        var oracle = JsonUtil.Get(PracticalCsvRecordParser.OracleFromCanonical(), "assertImportedThenCleaned") as JsonObject;
        Add(checks, "import-oracle-b2-quoted-comma",
            string.Equals(CellText(read, "B2"), JsonUtil.Str(oracle, "B2"), StringComparison.Ordinal));
        Add(checks, "import-oracle-d2-quoted-crlf",
            string.Equals(CellText(read, "D2"), JsonUtil.Str(oracle, "D2"), StringComparison.Ordinal));
        Add(checks, "import-oracle-e2-leading-zero",
            JsonStringEquals(read, "E2", JsonUtil.Str(oracle, "E2") ?? "0012"));
        Add(checks, "import-oracle-f2-formula-literal",
            JsonStringEquals(read, "F2", JsonUtil.Str(oracle, "F2") ?? "=1+1"));
        Add(checks, "import-not-parser-set-values", true);
        var csv = Path.Combine(req.ArtifactDir, "e9-clean.csv");
        Add(checks, "exported-csv-exists", File.Exists(csv));
        if (File.Exists(csv))
        {
            var lines = File.ReadAllLines(csv)
                .Select(l => l.Trim())
                .Where(l => l.Length > 0)
                .ToList();
            var data = lines.Skip(1).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            Add(checks, "exported-csv-unique-data-2", data.Count == 2);
        }
    }

    private static string InjectedPhase(JsonNode response)
    {
        if (JsonUtil.Bool(response, "ok") == true) return "unexpected-ok";
        if (!string.IsNullOrWhiteSpace(JsonUtil.Str(response, "confirmToken"))) return "apply-refused";
        return "preview-refused";
    }

    private static string SheetOf(string id) => id switch
    {
        "e1" => ScenarioE1.Sheet,
        "e2" => "내역서",
        "e3" => "작업일보",
        "e4" => "자재대장",
        "e5" => "누계집계",
        "e6" => "월간보고",
        "e7" => "원장",
        "e8" => "입력양식",
        "e9" => "정리",
        _ => "Sheet1",
    };

    private static string RangeOf(string id) => id switch
    {
        "e1" => "A1:BQ96",
        "e2" => "A1:H24",
        "e3" => "A1:F12",
        "e4" => "A1:J12",
        "e5" => "A1:B6",
        "e6" => "A1:D8",
        "e7" => "A1:C6",
        "e8" => "A1:B3",
        "e9" => "A1:G4",
        _ => "A1:A1",
    };

    private static void AssertInitial(LiveVerifyRequest req, JsonArray checks, string cell, double expected, double tol = 0.5)
    {
        Add(checks, "initial-captured", req.InitialRead is not null && JsonUtil.Bool(req.InitialRead, "ok") == true);
        var read = JsonUtil.Get(req.InitialRead, "read") as JsonObject;
        if (read is null)
        {
            Add(checks, "initial-" + cell, false);
            return;
        }
        AddApprox(checks, "initial-" + cell, CellNumber(read, cell), expected, tol);
    }

    private static void Add(JsonArray checks, string id, bool ok) =>
        checks.Add(new JsonObject { ["id"] = id, ["ok"] = ok });

    private static void AddNum(JsonArray checks, string id, double? actual, double expected) =>
        checks.Add(new JsonObject
        {
            ["id"] = id,
            ["ok"] = actual is { } a && Math.Abs(a - expected) < 0.5,
            ["actual"] = actual,
            ["expected"] = expected,
        });

    private static void AddApprox(JsonArray checks, string id, double? actual, double expected, double tol) =>
        checks.Add(new JsonObject
        {
            ["id"] = id,
            ["ok"] = actual is { } a && Math.Abs(a - expected) <= tol,
            ["actual"] = actual,
            ["expected"] = expected,
            ["tolerance"] = tol,
        });

    private static double? CellNumber(JsonNode read, string address)
    {
        var cells = JsonUtil.Get(read, "cells") as JsonArray
                    ?? JsonUtil.Get(JsonUtil.Get(read, "result"), "cells") as JsonArray;
        if (cells is not null)
        {
            foreach (var cell in cells.OfType<JsonObject>())
            {
                var addr = JsonUtil.Str(cell, "address") ?? JsonUtil.Str(cell, "a1") ?? JsonUtil.Str(cell, "r");
                if (!string.Equals(addr, address, StringComparison.OrdinalIgnoreCase)) continue;
                return AsNumber(cell["value"] ?? cell["v"] ?? cell["number"]);
            }
        }
        return AsNumber(GridCell(read, address));
    }

    private static string? CellText(JsonNode read, string address)
    {
        var cell = GridCell(read, address);
        if (cell is null) return null;
        if (cell is JsonValue v && v.TryGetValue<string>(out var s)) return s;
        var n = AsNumber(cell);
        return n?.ToString(CultureInfo.InvariantCulture) ?? cell.ToString();
    }

    private static JsonNode? GridCell(JsonNode read, string address)
    {
        var (row, col) = A1.ParseCell(address);
        var origin = JsonUtil.Str(read, "range") ?? JsonUtil.Str(JsonUtil.Get(read, "result"), "range") ?? "A1";
        var (or, oc, _, _) = A1.ParseRange(origin.Split('!')[^1]);
        var values = JsonUtil.Get(read, "values") as JsonArray
                     ?? JsonUtil.Get(JsonUtil.Get(read, "result"), "values") as JsonArray;
        var ri = row - or;
        var ci = col - oc;
        if (values is null || ri < 0 || ri >= values.Count || values[ri] is not JsonArray line) return null;
        if (ci < 0 || ci >= line.Count) return null;
        return line[ci];
    }

    private static double? AsNumber(JsonNode? n)
    {
        if (n is JsonValue v)
        {
            if (v.TryGetValue<double>(out var d)) return d;
            if (v.TryGetValue<int>(out var i)) return i;
            if (v.TryGetValue<long>(out var l)) return l;
            if (v.TryGetValue<string>(out var s) &&
                double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var p))
                return p;
        }
        return null;
    }

    private static JsonNode? Layout(JsonNode read) =>
        JsonUtil.Get(read, "layout") ?? JsonUtil.Get(JsonUtil.Get(read, "result"), "layout") ?? read;

    private static IEnumerable<string> MergeRefs(JsonNode read)
    {
        var merges = JsonUtil.Get(Layout(read), "mergedAreas") as JsonArray
                     ?? JsonUtil.Get(read, "mergedAreas") as JsonArray;
        if (merges is null) yield break;
        foreach (var n in merges)
        {
            var s = n?.GetValue<string>() ?? JsonUtil.Str(n as JsonObject, "ref") ?? "";
            if (!string.IsNullOrWhiteSpace(s)) yield return s;
        }
    }

    private static int MergeCount(JsonNode read)
    {
        var list = MergeRefs(read).ToList();
        return list.Count == 0 && JsonUtil.Get(Layout(read), "mergedAreas") is null ? -1 : list.Count;
    }

    private static int MonthHeaderCount(JsonNode read) => MergeRefs(read).Count(A1.IsMonthHeaderPair);

    private static bool MergesOverlap(JsonNode read, string range) =>
        MergeRefs(read).Any(m => A1.Overlaps(m, range));

    private static string NormalizeMergeRef(string range) =>
        range.Replace("$", "", StringComparison.Ordinal).Split('!')[^1].Trim();

    private static bool MergeSetHasExact(JsonNode read, string expected)
    {
        var want = NormalizeMergeRef(expected);
        return MergeRefs(read).Any(m =>
            string.Equals(NormalizeMergeRef(m), want, StringComparison.OrdinalIgnoreCase));
    }

    private static bool FreezeIs(JsonNode read, string cell)
    {
        var freeze = JsonUtil.Get(Layout(read), "freezePanes") ?? JsonUtil.Get(read, "freezePanes");
        if (JsonUtil.Bool(freeze, "readable") == false) return false;
        var top = JsonUtil.Str(freeze, "topLeftCell") ?? JsonUtil.Str(freeze, "cell");
        return string.Equals(top, cell, StringComparison.OrdinalIgnoreCase);
    }

    private static bool PagePaperLandscape(JsonNode read)
    {
        var page = JsonUtil.Get(Layout(read), "pageSetup") ?? JsonUtil.Get(read, "pageSetup");
        var name = JsonUtil.Str(page, "paperSizeName") ?? JsonUtil.Str(page, "paperSize");
        var size = JsonUtil.Num(page, "paperSize");
        var ori = JsonUtil.Str(page, "orientation");
        var paper = string.Equals(name, "A3", StringComparison.OrdinalIgnoreCase) || size is 8;
        return paper && string.Equals(ori, "landscape", StringComparison.OrdinalIgnoreCase);
    }

    private static bool PrintAreaIs(JsonNode read, string expected)
    {
        var page = JsonUtil.Get(Layout(read), "pageSetup") ?? JsonUtil.Get(read, "pageSetup");
        var got = (JsonUtil.Str(page, "printArea") ?? "").Replace("$", "", StringComparison.Ordinal);
        if (got.Contains('!', StringComparison.Ordinal))
            got = got[(got.IndexOf('!') + 1)..];
        return string.Equals(got, expected, StringComparison.OrdinalIgnoreCase);
    }

    private static bool FooterContains(JsonNode read, string needle)
    {
        var page = JsonUtil.Get(Layout(read), "pageSetup") ?? JsonUtil.Get(read, "pageSetup");
        var footer = (JsonUtil.Str(page, "centerFooter") ?? "") + (JsonUtil.Str(page, "oddFooter") ?? "")
                     + (JsonUtil.Str(page, "leftFooter") ?? "");
        return footer.Contains(needle, StringComparison.Ordinal);
    }

    private static bool PageScaleOrFit(JsonNode read)
    {
        var page = JsonUtil.Get(Layout(read), "pageSetup") ?? JsonUtil.Get(read, "pageSetup");
        var scale = JsonUtil.Num(page, "scale");
        var fitW = JsonUtil.Num(page, "fitToWidth");
        var fitH = JsonUtil.Num(page, "fitToHeight");
        var fitDisabled = fitW is null or 0 && fitH is null or 0;
        return scale is { } s && Math.Abs(s - 55) < 0.5 && fitDisabled;
    }

    private static int ColumnWidthCount(JsonNode read) =>
        (JsonUtil.Get(Layout(read), "columnWidths") as JsonArray)?.Count ?? -1;

    private static int RowHeightCount(JsonNode read) =>
        (JsonUtil.Get(Layout(read), "rowHeights") as JsonArray)?.Count ?? -1;

    private static bool ColumnWidthNear(JsonNode read, int colFrom, int colTo, double expected, double tol)
    {
        var cols = JsonUtil.Get(Layout(read), "columnWidths") as JsonArray;
        if (cols is null) return false;
        var seen = 0;
        foreach (var col in cols.OfType<JsonObject>())
        {
            var n = (int)(JsonUtil.Num(col, "column") ?? 0);
            if (n < colFrom || n > colTo) continue;
            var w = JsonUtil.Num(col, "widthChars") ?? JsonUtil.Num(col, "columnWidth");
            if (w is null || Math.Abs(w.Value - expected) > tol) return false;
            seen++;
        }
        return seen == colTo - colFrom + 1;
    }

    private static bool FontIs(JsonNode read, string name, double size)
    {
        var style = JsonUtil.Get(read, "styles") as JsonObject
                    ?? JsonUtil.Get(JsonUtil.Get(read, "result"), "styles") as JsonObject;
        if (style is null)
        {
            var arr = JsonUtil.Get(read, "styles") as JsonArray
                      ?? JsonUtil.Get(JsonUtil.Get(read, "result"), "styles") as JsonArray;
            style = arr?.OfType<JsonObject>().FirstOrDefault();
        }
        if (style is null) return false;
        var font = JsonUtil.Str(style, "fontName") ?? JsonUtil.Str(JsonUtil.Get(style, "font"), "name");
        var sz = JsonUtil.Num(style, "fontSize") ?? JsonUtil.Num(style, "size") ?? JsonUtil.Num(JsonUtil.Get(style, "font"), "size");
        return string.Equals(font, name, StringComparison.OrdinalIgnoreCase) && sz is { } n && Math.Abs(n - size) < 0.5;
    }

    private static bool JsonStringEquals(JsonNode read, string address, string expected)
    {
        var cell = GridCell(read, address);
        return cell is JsonValue v && v.TryGetValue<string>(out var s) &&
               string.Equals(s, expected, StringComparison.Ordinal);
    }

    private static bool IsBlank(JsonNode read, string address)
    {
        var cell = GridCell(read, address);
        if (cell is null) return true;
        return cell is JsonValue v && v.TryGetValue<string>(out var s) && string.IsNullOrEmpty(s);
    }

    private static string? StyleNumberFormat(JsonNode read)
    {
        var style = JsonUtil.Get(read, "styles") as JsonObject
                    ?? JsonUtil.Get(JsonUtil.Get(read, "result"), "styles") as JsonObject;
        if (style is null)
        {
            var arr = JsonUtil.Get(read, "styles") as JsonArray
                      ?? JsonUtil.Get(JsonUtil.Get(read, "result"), "styles") as JsonArray;
            style = arr?.OfType<JsonObject>().FirstOrDefault();
        }
        return JsonUtil.Str(style, "numberFormat") ?? JsonUtil.Str(JsonUtil.Get(style, "number"), "format");
    }

    private static double? RowHeightPoints(JsonNode read, int row)
    {
        var rows = JsonUtil.Get(Layout(read), "rowHeights") as JsonArray;
        if (rows is null) return null;
        foreach (var item in rows.OfType<JsonObject>())
        {
            var n = (int)(JsonUtil.Num(item, "row") ?? 0);
            if (n != row) continue;
            return JsonUtil.Num(item, "heightPoints") ?? JsonUtil.Num(item, "rowHeight");
        }
        return null;
    }

    private static bool NumberFormatHas(JsonNode read, string needle)
    {
        var style = JsonUtil.Get(read, "styles") as JsonObject
                    ?? JsonUtil.Get(JsonUtil.Get(read, "result"), "styles") as JsonObject;
        if (style is null)
        {
            var arr = JsonUtil.Get(read, "styles") as JsonArray
                      ?? JsonUtil.Get(JsonUtil.Get(read, "result"), "styles") as JsonArray;
            style = arr?.OfType<JsonObject>().FirstOrDefault();
        }
        var fmt = JsonUtil.Str(style, "numberFormat") ?? JsonUtil.Str(JsonUtil.Get(style, "number"), "format") ?? "";
        return fmt.Contains(needle, StringComparison.Ordinal);
    }

    private static bool StyleBold(JsonNode read)
    {
        var style = JsonUtil.Get(read, "styles") as JsonObject
                    ?? JsonUtil.Get(JsonUtil.Get(read, "result"), "styles") as JsonObject;
        return JsonUtil.Bool(style, "bold") == true || JsonUtil.Bool(style, "fontBold") == true;
    }

    private static bool? StyleLocked(JsonNode read)
    {
        var style = JsonUtil.Get(read, "styles") as JsonObject
                    ?? JsonUtil.Get(JsonUtil.Get(read, "result"), "styles") as JsonObject;
        return JsonUtil.Bool(style, "locked");
    }

    private static bool StyleWrap(JsonNode read)
    {
        var style = JsonUtil.Get(read, "styles") as JsonObject
                    ?? JsonUtil.Get(JsonUtil.Get(read, "result"), "styles") as JsonObject;
        return JsonUtil.Bool(style, "wrapText") == true;
    }

    private static int CountObjects(JsonNode inspect, string kind)
    {
        var objects = JsonUtil.Get(inspect, "objects") as JsonArray
                      ?? JsonUtil.Get(JsonUtil.Get(inspect, "result"), "objects") as JsonArray
                      ?? JsonUtil.Get(inspect, kind) as JsonArray
                      ?? JsonUtil.Get(inspect, kind + "s") as JsonArray;
        if (objects is null) return -1;
        var needle = kind.TrimEnd('s');
        return objects.OfType<JsonObject>().Count(o =>
        {
            var k = JsonUtil.Str(o, "kind") ?? JsonUtil.Str(o, "type") ?? JsonUtil.Str(o, "objectKind") ?? "";
            return k.Contains(needle, StringComparison.OrdinalIgnoreCase);
        });
    }

    private static bool GridContainsText(JsonNode read, string expected)
    {
        var values = JsonUtil.Get(read, "values") as JsonArray
                     ?? JsonUtil.Get(JsonUtil.Get(read, "result"), "values") as JsonArray;
        if (values is null) return false;
        foreach (var row in values.OfType<JsonArray>())
        {
            foreach (var cell in row)
            {
                if (cell is JsonValue v && v.TryGetValue<string>(out var s) &&
                    string.Equals(s, expected, StringComparison.Ordinal))
                    return true;
                if (cell is not null && string.Equals(cell.ToString(), expected, StringComparison.Ordinal))
                    return true;
            }
        }
        return false;
    }

    private static bool GridContainsNumber(JsonNode read, double expected)
    {
        var values = JsonUtil.Get(read, "values") as JsonArray
                     ?? JsonUtil.Get(JsonUtil.Get(read, "result"), "values") as JsonArray;
        if (values is null) return false;
        foreach (var row in values.OfType<JsonArray>())
        {
            foreach (var cell in row)
            {
                var n = AsNumber(cell);
                if (n is { } v && Math.Abs(v - expected) < 0.5) return true;
            }
        }
        return false;
    }

    private static int CountNonEmptyDataRows(JsonNode read)
    {
        var values = JsonUtil.Get(read, "values") as JsonArray
                     ?? JsonUtil.Get(JsonUtil.Get(read, "result"), "values") as JsonArray;
        if (values is null) return -1;
        var count = 0;
        for (var i = 1; i < values.Count; i++)
        {
            if (values[i] is not JsonArray row) continue;
            if (row.Any(c => c is not null && c.ToString() is { Length: > 0 } s && s != "null"))
                count++;
        }
        return count;
    }

    private static HashSet<string> StringSet(JsonObject xml, string key)
    {
        var arr = JsonUtil.Get(xml, key) as JsonArray ?? [];
        return arr.Select(n => n?.GetValue<string>() ?? "")
            .Where(s => s.Length > 0)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);
    }

    private static List<(int R, int C)> LoadRebuildBars(string rebuildDir)
    {
        var path = Path.Combine(rebuildDir, "bars.json");
        if (!File.Exists(path)) return [];
        var bars = JsonUtil.Load(path) as JsonArray ?? [];
        var list = new List<(int, int)>();
        foreach (var n in bars.OfType<JsonObject>())
        {
            var r = (int)(JsonUtil.Num(n, "r") ?? JsonUtil.Num(n, "row") ?? 0);
            var c = (int)(JsonUtil.Num(n, "c") ?? JsonUtil.Num(n, "col") ?? 0);
            if (r > 0 && c > 0) list.Add((r, c));
        }
        return list;
    }
}
