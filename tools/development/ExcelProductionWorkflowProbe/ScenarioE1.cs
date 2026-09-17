namespace DocBridge.Development.ExcelProductionWorkflowProbe;

internal static class ScenarioE1
{
    public const string Id = "e1";
    public const string Sheet = "예정공정표";

    public static List<PlannedBatch> Plan(ProbeOptions options, string artifactDir)
    {
        var xlsx = Path.Combine(artifactDir, "e1-schedule.xlsx");
        var pdf = Path.Combine(artifactDir, "e1-schedule.pdf");
        SourceIntegrity.RefuseIfProtectedPath(xlsx, options.SourceXlsx);
        SourceIntegrity.RefuseIfProtectedPath(pdf, options.SourceXlsx);

        var oracle = ScheduleOracle.Export(options);
        var month = ((JsonArray)oracle["e1MergePlan"]!["batch1_monthHeaders"]!).Select(n => n!.GetValue<string>()).ToList();
        if (month.Count != 31)
            throw new InvalidOperationException($"e1-merge-month-31 must stay one token batch of 31 pairs, got {month.Count}");
        var remaining = ((JsonArray)oracle["e1MergePlan"]!["batch2_remaining"]!).Select(n => n!.GetValue<string>()).ToList();
        var values = ClipGrid(ScheduleOracle.LoadValues(options.RebuildDir), 96, 69);
        var reconstruction = StyleReconstructor.Reconstruct(options.SourceXlsx, Sheet);
        OverlayComLayout(reconstruction, options.ComOraclePath);
        var page = JsonUtil.Get(reconstruction, "page") as JsonObject ?? new JsonObject();
        var typedPage = PagePayload.E1Acceptance(page);
        var pageHygiene = PagePayload.HygieneErrors(typedPage);
        if (pageHygiene.Count > 0)
            throw new InvalidOperationException("E1 set_page_setup payload failed typed schema hygiene: " + pageHygiene);

        var batches = new List<PlannedBatch>
        {
            ApplyPlanner.Batch(Id, "e1-create", OpFamily.Lifecycle, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "create_workbook",
                ["sheetName"] = Sheet,
            }), "token", "blank owned workbook; identity must be new vs inventory"),
            ApplyPlanner.Batch(Id, "e1-save-as", OpFamily.Lifecycle, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "save_workbook",
                ["output"] = xlsx,
            }), "token", "SaveAs before execute; DocumentRef set only after public identity matches this path"),
            ApplyPlanner.Batch(Id, BlankCheckpoint.BatchId, OpFamily.Lifecycle, new JsonArray(), "checkpoint",
                "read-only OOXML of the owned save; fail before values if cells/merges/target fills exist. Planner blankSheetNoFillAssumed is not this proof.",
                capture: "blank-nofill"),
            ApplyPlanner.Batch(Id, "e1-values", OpFamily.Values, JsonUtil.Arr(new JsonObject
            {
                ["op"] = "set_values",
                ["target"] = JsonUtil.Target(Sheet),
                ["range"] = "A1:BQ96",
                ["values"] = values.DeepClone(),
            }), "execute"),
        };

        batches.AddRange(FormulaBatches(ScheduleOracle.LoadFormulas(options.RebuildDir)));
        batches.AddRange(StyleReconstructor.FormatBatches(Id, Sheet, (JsonArray)reconstruction["formatOps"]!));
        batches.Add(ApplyPlanner.Batch(Id, "e1-merge-month-31", OpFamily.Merge, ApplyPlanner.Merges(Sheet, month), "token",
            "31 month-header pairs after property styles; row 5 stays unmerged"));
        const int mergeChunk = 40;
        for (var i = 0; i < remaining.Count; i += mergeChunk)
        {
            var part = remaining.Skip(i).Take(mergeChunk).ToList();
            batches.Add(ApplyPlanner.Batch(Id, $"e1-merge-remaining-{i / mergeChunk + 1}", OpFamily.Merge, ApplyPlanner.Merges(Sheet, part), "token",
                $"remaining merges chunk {i / mergeChunk + 1}; total {month.Count + remaining.Count} must stay 281; styles already applied; product STA preview budget is 120s"));
        }

        batches.Add(ApplyPlanner.Batch(Id, "e1-sheet-layout", OpFamily.SheetLayout, JsonUtil.Arr(
            new JsonObject
            {
                ["op"] = "set_column_widths",
                ["target"] = JsonUtil.Target(Sheet),
                ["columns"] = PublicLayout.Columns(reconstruction["columns"]),
            },
            new JsonObject
            {
                ["op"] = "set_row_heights",
                ["target"] = JsonUtil.Target(Sheet),
                ["rows"] = PublicLayout.Rows(reconstruction["rows"]),
            },
            new JsonObject
            {
                ["op"] = "freeze_panes",
                ["target"] = JsonUtil.Target(Sheet),
                ["cell"] = "G6",
            },
            new JsonObject
            {
                ["op"] = "set_page_setup",
                ["target"] = JsonUtil.Target(Sheet),
                ["page"] = typedPage,
            }), "token", "ALL A:BQ native widths and 1:96 heights; requested integer scale 55 / freeze G6; centerFooter only (no oddHeader/oddFooter, no null fit). Public column/row items strip oracle unit/note."));

        var chart = ChartReconstructor.Reconstruct(options.SourceXlsx, options.ComOraclePath);
        var createOp = (JsonObject)chart["createOp"]!.DeepClone();
        var chartHygiene = PagePayload.ChartHygieneErrors(createOp);
        if (chartHygiene.Count > 0)
            throw new InvalidOperationException("E1 create_chart failed published data contract: " + chartHygiene);
        batches.Add(ApplyPlanner.Batch(Id, "e1-chart1", OpFamily.Data, JsonUtil.Arr(createOp), "token",
            "Native Chart1 via create_chart from xl/charts/chart1.xml + drawing anchors. Not a raster."));

        batches.Add(ApplyPlanner.Batch(Id, "e1-collision-seed", OpFamily.Values, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "set_values",
            ["target"] = JsonUtil.Target(Sheet),
            ["range"] = "A100:D100",
            ["values"] = new JsonArray
            {
                new JsonArray { null, "NON-UL-CONTENT", null, null },
            },
        }), "execute", "B100 is non-top-left of A100:B100. Top-left-only content would be legal."));

        batches.Add(ApplyPlanner.Batch(Id, "e1-legal-new-merge", OpFamily.Merge, ApplyPlanner.Merges(Sheet, ["E100:F100"]), "token",
            "Genuinely new empty merge that must remain after the later collision batch is refused."));

        batches.Add(ApplyPlanner.Batch(Id, "e1-merge-collision-refuse", OpFamily.Merge, ApplyPlanner.Merges(Sheet, ["A100:B100", "C100:D100"]), "token",
            "A100:B100 refuses (B100 non-UL content). C100:D100 is a new merge in the same refused batch and must not remain. Preview refuse is not rollback.",
            injectedFailure: true));

        batches.Add(ApplyPlanner.Batch(Id, "e1-save", OpFamily.Lifecycle, JsonUtil.Arr(
            new JsonObject { ["op"] = "save_workbook", ["output"] = xlsx, ["overwrite"] = true }), "token"));
        batches.Add(ApplyPlanner.Batch(Id, "e1-close", OpFamily.Lifecycle, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "close_workbook",
            ["saveChanges"] = false,
        }), "token", "required for save-close-reopen; missing-op is a failed requirement"));
        batches.Add(ApplyPlanner.Batch(Id, "e1-reopen", OpFamily.Lifecycle, JsonUtil.Arr(new JsonObject
        {
            ["op"] = "open_workbook",
            ["path"] = xlsx,
        }), "token"));
        batches.Add(ApplyPlanner.Batch(Id, "e1-export-pdf", OpFamily.Lifecycle, JsonUtil.Arr(
            new JsonObject { ["op"] = "export_pdf", ["output"] = pdf, ["sheet"] = Sheet, ["overwrite"] = true }), "token",
            "visual-ready PDF after reopen"));

        return batches;
    }

    public static JsonObject Expected() => new()
    {
        ["sheet"] = Sheet,
        ["range"] = "A1:BQ96",
        ["merges"] = AcceptanceNumbers.MergeCount,
        ["monthHeaderMerges"] = AcceptanceNumbers.MonthHeaderMerges,
        ["bars"] = AcceptanceNumbers.BarCount,
        ["BP95"] = AcceptanceNumbers.Bp95,
        ["BP96"] = AcceptanceNumbers.Bp96,
        ["AB95"] = AcceptanceNumbers.Ab95,
        ["freeze"] = new JsonObject
        {
            ["requested"] = "G6",
            ["sourceXmlPane"] = "6/5 G6",
            ["sourceNativeSplit"] = "5/4",
            ["note"] = "Requested G6 is acceptance. Native 5/4 is recorded, not treated as equal.",
        },
        ["page"] = new JsonObject
        {
            ["requestedScale"] = 55,
            ["sourceXmlScale"] = 55,
            ["sourceXmlFitToPage"] = 1,
            ["sourceNativeZoom"] = 0,
            ["sourceNativeFitToWidth"] = 1,
            ["note"] = "Requested scale 55 is acceptance. Native Zoom 0 / FitWidth 1 is recorded separately.",
        },
        ["chart1"] = "create_chart from chart1.xml + drawing1.xml anchors; red 2.25pt line G95:BP95; not raster",
        ["usedStyles"] = 60,
        ["blankCheckpoint"] = "e1-blank-checkpoint reads owned OOXML after create+save; no nonempty cells/merges; effective XF0 and applyFill=0→cellStyleXfs fill cover omitted/un-s cells; unused gray125 allowed; planner blankSheetNoFillAssumed is not proof",
        ["stylesBeforeMerges"] = "format_range property batches after values/formulas and before the 281 merges",
        ["layout"] = "all A:BQ widths and 1:96 heights from XML; exact footer; not 3-field equality",
        ["formulas"] = "only actual formula cells/groups; null grid is forbidden",
        ["collision"] = "legal E100:F100 stays merged; B100 kept; C100:D100 unmerged; preview refuse != rollback",
        ["reopen"] = "save + close_workbook + open_workbook + export_pdf",
        ["sourceHashFromConfig"] = true,
        ["row5Unmerged"] = true,
        ["xlsx"] = "e1-schedule.xlsx",
        ["pdf"] = "e1-schedule.pdf",
        ["pdfPagesExpected"] = 1,
    };

    private static IEnumerable<PlannedBatch> FormulaBatches(JsonArray sparse)
    {
        var cells = sparse.OfType<JsonObject>()
            .Select(o => (
                R: (int)(JsonUtil.Num(o, "r") ?? 0),
                C: (int)(JsonUtil.Num(o, "c") ?? 0),
                F: JsonUtil.Str(o, "f")))
            .Where(t => t.R >= 1 && t.C >= 1 && !string.IsNullOrWhiteSpace(t.F))
            .OrderBy(t => t.R).ThenBy(t => t.C)
            .ToList();

        var ops = new JsonArray();
        foreach (var row in cells.GroupBy(t => t.R))
        {
            var list = row.ToList();
            var start = 0;
            while (start < list.Count)
            {
                var end = start;
                while (end + 1 < list.Count && list[end + 1].C == list[end].C + 1) end++;
                var formulas = new JsonArray();
                var line = new JsonArray();
                for (var i = start; i <= end; i++) line.Add(list[i].F);
                formulas.Add(line);
                ops.Add(new JsonObject
                {
                    ["op"] = "set_formulas",
                    ["target"] = JsonUtil.Target(Sheet),
                    ["range"] = A1.Range(row.Key, list[start].C, row.Key, list[end].C),
                    ["formulas"] = formulas,
                });
                start = end + 1;
            }
        }

        const int chunk = 80;
        for (var i = 0; i < ops.Count; i += chunk)
        {
            var part = new JsonArray();
            for (var j = i; j < Math.Min(i + chunk, ops.Count); j++)
                part.Add(ops[j]!.DeepClone());
            yield return ApplyPlanner.Batch(Id, $"e1-formulas-{i / chunk + 1}", OpFamily.Values, part, "execute",
                "only actual formula cells; no null sparse matrix");
        }
    }

    private static void OverlayComLayout(JsonObject reconstruction, string? comOraclePath)
    {
        if (string.IsNullOrWhiteSpace(comOraclePath) || !File.Exists(comOraclePath)) return;
        var com = JsonUtil.Load(comOraclePath).AsObject();
        if (com["columnsABQ"] is JsonArray cols)
        {
            var grouped = new JsonArray();
            string? start = null;
            double? w = null;
            var count = 0;
            void Flush()
            {
                if (start is null || w is null || count == 0) return;
                grouped.Add(new JsonObject
                {
                    ["col"] = start,
                    ["count"] = count,
                    ["widthChars"] = w.Value,
                });
            }
            foreach (var col in cols.OfType<JsonObject>())
            {
                var cw = JsonUtil.Num(col, "columnWidth");
                var letter = JsonUtil.Str(col, "col");
                if (cw is null || letter is null) continue;
                if (w is { } prev && Math.Abs(prev - cw.Value) < 1e-6 && start is not null) count++;
                else
                {
                    Flush();
                    start = letter;
                    w = cw;
                    count = 1;
                }
            }
            Flush();
            reconstruction["columns"] = grouped;
        }
        if (com["rows1to96"] is JsonArray rows)
        {
            var grouped = new JsonArray();
            foreach (var row in rows.OfType<JsonObject>())
                grouped.Add(new JsonObject
                {
                    ["row"] = JsonUtil.Num(row, "row"),
                    ["count"] = 1,
                    ["heightPoints"] = JsonUtil.Num(row, "heightPoints"),
                });
            reconstruction["rows"] = grouped;
        }
        if (JsonUtil.Get(com, "print") is JsonObject print && reconstruction["page"] is JsonObject page)
        {
            page["widthSource"] = "com-oracle";
            page["comZoom"] = print["zoom"]?.DeepClone();
            page["comFitToPagesWide"] = JsonUtil.Num(print, "fitToPagesWide");
            page["comFitToPagesTall"] = JsonUtil.Num(print, "fitToPagesTall");
            page["centerFooter"] = PrintFooter.StripCenterMarker(JsonUtil.Str(print, "centerFooter"));
        }
    }

    private static JsonArray ClipGrid(JsonArray source, int rows, int cols)
    {
        var grid = new JsonArray();
        for (var r = 0; r < rows; r++)
        {
            var row = new JsonArray();
            var src = r < source.Count ? source[r] as JsonArray : null;
            for (var c = 0; c < cols; c++)
                row.Add(src is not null && c < src.Count ? src[c]?.DeepClone() : null);
            grid.Add(row);
        }
        return grid;
    }
}
