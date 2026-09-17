using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelExtendedOpsContractTests
{
    [Fact]
    public void Csv_round_trips_quoted_crlf_empty_and_leading_zeros()
    {
        var csv = "001,\"line1\r\nline2, still\",\r\n002,\"say \"\"hi\"\"\",x\r\n";
        var parsed = ExcelCsvContract.NormalizeRectangle(ExcelCsvContract.Parse(csv));
        Assert.Equal(2, parsed.Count);
        Assert.Equal("001", parsed[0][0]);
        Assert.Equal("line1\r\nline2, still", parsed[0][1]);
        Assert.Equal("", parsed[0][2]);
        Assert.Equal("say \"hi\"", parsed[1][1]);
        Assert.True(ExcelCsvContract.LooksLikeIdentifierOrFormulaText("001"));
        Assert.True(ExcelCsvContract.LooksLikeIdentifierOrFormulaText("=A1"));
        var written = ExcelCsvContract.Write(parsed);
        var again = ExcelCsvContract.NormalizeRectangle(ExcelCsvContract.Parse(written));
        Assert.True(ExcelCsvContract.TablesEqual(parsed, again));
    }

    [Fact]
    public void Csv_keeps_empty_rows_and_ragged_columns()
    {
        var parsed = ExcelCsvContract.NormalizeRectangle(ExcelCsvContract.Parse("a,b\r\n\r\nc"));
        Assert.Equal(3, parsed.Count);
        Assert.Equal(2, parsed[0].Count);
        Assert.Equal("", parsed[1][0]);
        Assert.Equal("", parsed[1][1]);
        Assert.Equal("c", parsed[2][0]);
        Assert.Equal("", parsed[2][1]);
    }

    [Fact]
    public void Export_and_import_respect_the_same_delimiter()
    {
        var table = ExcelCsvContract.NormalizeRectangle(new[]
        {
            new[] { "001", "a;b", "" },
            new[] { "=A1", "line1\r\nline2", "x" },
        });
        var written = ExcelCsvContract.Write(table, ';');
        Assert.Contains(";", written, StringComparison.Ordinal);
        Assert.DoesNotContain(",", written.Split('"')[0], StringComparison.Ordinal);
        var again = ExcelCsvContract.NormalizeRectangle(ExcelCsvContract.Parse(written, ';'));
        Assert.True(ExcelCsvContract.TablesEqual(table, again));
        Assert.Equal("001", again[0][0]);
        Assert.Equal("=A1", again[1][0]);
        Assert.Equal("line1\r\nline2", again[1][1]);
        Assert.Equal("", again[0][2]);
    }

    [Fact]
    public void Import_and_export_limits_are_explicit_before_allocation()
    {
        Assert.Contains("exceeds", ExcelCsvContract.ExportLimitError(1001, 1000), StringComparison.OrdinalIgnoreCase);
        Assert.Null(ExcelCsvContract.ExportLimitError(10, 10));
        Assert.Contains("exceeds", ExcelCsvContract.ImportLimitError(1001, 1000), StringComparison.OrdinalIgnoreCase);
        Assert.Null(ExcelCsvContract.ImportLimitError(10, 10));
        Assert.Equal(',', ExcelCsvContract.ResolveDelimiter(new JsonObject()));
        Assert.Equal(';', ExcelCsvContract.ResolveDelimiter(new JsonObject { ["delimiter"] = ";" }));
        Assert.Throws<InvalidOperationException>(() =>
            ExcelCsvContract.ResolveDelimiter(new JsonObject { ["delimiter"] = ";;" }));
        Assert.False(ExcelCsvContract.TablesEqual(
            new[] { new[] { "001" } },
            new[] { new[] { "1" } }));
    }

    [Fact]
    public void Outline_show_false_is_collapse_not_clear()
    {
        var collapse = Json.ParseObject("""{ "op": "set_outline", "show": false, "range": "A2:A5" }""")!;
        Assert.Equal(ExcelOutlineContract.OutlineAction.Collapse, ExcelOutlineContract.ResolveAction(collapse));
        Assert.Contains("collapse", ExcelOutlineContract.Describe(collapse), StringComparison.OrdinalIgnoreCase);
        Assert.True(ExcelOutlineContract.MatchesReadback(
            ExcelOutlineContract.OutlineAction.Collapse, null, actualLevel: 2, showDetail: false, hasOutline: true));
        Assert.False(ExcelOutlineContract.MatchesReadback(
            ExcelOutlineContract.OutlineAction.Collapse, null, actualLevel: 2, showDetail: true, hasOutline: true));

        var clear = Json.ParseObject("""{ "op": "set_outline", "clear": true, "range": "A2:A5" }""")!;
        Assert.Equal(ExcelOutlineContract.OutlineAction.Ungroup, ExcelOutlineContract.ResolveAction(clear));
        var group = Json.ParseObject("""{ "op": "set_outline", "level": 2, "range": "A2:A5", "summaryBelow": true }""")!;
        Assert.Equal(ExcelOutlineContract.OutlineAction.Group, ExcelOutlineContract.ResolveAction(group));
        Assert.Equal(2, ExcelOutlineContract.RequestedLevel(group));
        Assert.Equal(ExcelOutlineContract.XlSummaryBelow, ExcelOutlineContract.SummaryRow(group));
        Assert.Equal(-4131, ExcelOutlineContract.XlSummaryOnLeft);
        Assert.Equal(-4152, ExcelOutlineContract.XlSummaryOnRight);
        var right = Json.ParseObject("""{ "op": "set_outline", "range": "B1:D1", "summaryRight": true }""")!;
        Assert.Equal(-4152, ExcelOutlineContract.SummaryColumn(right));
    }

    [Fact]
    public void Outline_axis_is_explicit_row_or_column()
    {
        var rows = Json.ParseObject("""{ "op": "set_outline", "range": "A2:A5" }""")!;
        Assert.Equal(ExcelOutlineContract.OutlineAxis.Rows, ExcelOutlineContract.ResolveAxis(rows));
        var columns = Json.ParseObject("""{ "op": "set_outline", "range": "B1:D1", "axis": "column" }""")!;
        Assert.Equal(ExcelOutlineContract.OutlineAxis.Columns, ExcelOutlineContract.ResolveAxis(columns));
        Assert.Contains("columns", ExcelOutlineContract.Describe(columns), StringComparison.OrdinalIgnoreCase);
        Assert.Throws<InvalidOperationException>(() =>
            ExcelOutlineContract.ResolveAxis(Json.ParseObject("""{ "axis": "diagonal" }""")!));
    }

    [Fact]
    public void Set_formulas_defaults_to_legacy_formula_unless_explicit_formula2()
    {
        Assert.Equal(ExcelFormula2Contract.EngineFormula, ExcelFormula2Contract.ResolveEngine(
            Json.ParseObject("""{ "op": "set_formulas", "formulas": [["=A1+1"]] }""")!));
        Assert.Equal(ExcelFormula2Contract.EngineFormula, ExcelFormula2Contract.ResolveEngine(
            Json.ParseObject("""{ "op": "set_formulas", "engine": "formula" }""")!));
        Assert.Equal(ExcelFormula2Contract.EngineFormula2, ExcelFormula2Contract.ResolveEngine(
            Json.ParseObject("""{ "op": "set_formulas", "engine": "formula2" }""")!));
        Assert.Equal(ExcelFormula2Contract.EngineFormula2, ExcelFormula2Contract.ResolveEngine(
            Json.ParseObject("""{ "op": "set_formulas", "formulaEngine": "formula2" }""")!));
        Assert.Equal(ExcelFormula2Contract.EngineFormula2, ExcelFormula2Contract.ResolveEngine(
            Json.ParseObject("""{ "op": "set_formulas", "formula2": true }""")!));
        Assert.Throws<InvalidOperationException>(() =>
            ExcelFormula2Contract.ResolveEngine(Json.ParseObject("""{ "engine": "legacy" }""")!));
        Assert.False(ExcelFormula2Contract.WritesFormula2(
            Json.ParseObject("""{ "op": "set_formulas", "formulas": [["=A1"]] }""")!));
        Assert.True(ExcelFormula2Contract.WritesFormula2(ExcelFormula2Contract.Sequence32Fixture));
        Assert.True(ExcelFormula2Contract.WritesFormula2(ExcelFormula2Contract.FilterFixture));
        Assert.Null(ExcelFormula2Contract.HasSpillFromCom(null));
        Assert.True(ExcelFormula2Contract.HasSpillFromCom(true));
        Assert.False(ExcelFormula2Contract.HasSpillFromCom(false));
        var mixed = ExcelFormula2Contract.SpillReadback("B2", ExcelFormula2Contract.EngineFormula2, null, null,
            formula2: "=SEQUENCE(3,2)");
        Assert.Null(mixed["hasSpill"]);
        Assert.Equal("formula2", Json.GetString(mixed, "engine"));
        Assert.Equal("=SEQUENCE(3,2)", Json.GetString(mixed, "formula2"));
        Assert.False(ExcelFormula2Contract.ReadbacksMatch(
            mixed,
            ExcelFormula2Contract.SpillReadback("B2", ExcelFormula2Contract.EngineFormula2, true, "B2:C4")));
    }

    [Fact]
    public void Formula2_estimates_SEQUENCE_spill_and_keeps_unknown_filter_on_dest()
    {
        Assert.True(ExcelFormula2Contract.TryEstimateSequenceFootprint("=SEQUENCE(3,2)", out var rows, out var columns));
        Assert.Equal(3, rows);
        Assert.Equal(2, columns);
        Assert.False(ExcelFormula2Contract.TryEstimateSequenceFootprint("=FILTER(A2:B10,LEN(A2:A10)>0)", out _, out _));
        var sequence = ExcelFormula2Contract.ResolveCaptureFootprint(ExcelFormula2Contract.Sequence32Fixture, 1, 1);
        Assert.Equal((3, 2), sequence);
        var filter = ExcelFormula2Contract.ResolveCaptureFootprint(ExcelFormula2Contract.FilterFixture, 1, 1);
        Assert.Equal((1, 1), filter);
        Assert.True(ExcelFormula2Contract.SpillFitsCapture("B2", "B2:C4", 3, 2));
        Assert.False(ExcelFormula2Contract.SpillFitsCapture("B2", "B2:C4", 1, 1));
        Assert.False(ExcelFormula2Contract.SpillFitsCapture("B2", "E2:F10", 9, 2));
        Assert.True(ExcelFormula2Contract.TryParseA1Cell("$E$2", out var dollarRow, out var dollarCol));
        Assert.Equal(2, dollarRow);
        Assert.Equal(5, dollarCol);
        Assert.Equal("E2:F10", ExcelFormula2Contract.CleanA1("$E$2:$F$10"));
        Assert.True(ExcelFormula2Contract.IsOpenDynamicArrayFormula(ExcelFormula2Contract.FirstFormula(ExcelFormula2Contract.FilterFixture)));
        Assert.True(ExcelFormula2Contract.IsOpenDynamicArrayFormula(ExcelFormula2Contract.FirstFormula(ExcelFormula2Contract.SortFixture)));
        Assert.True(ExcelFormula2Contract.IsOpenDynamicArrayFormula(ExcelFormula2Contract.FirstFormula(ExcelFormula2Contract.UniqueFixture)));
        Assert.False(ExcelFormula2Contract.RejectsOversizedActualSpill(ExcelFormula2Contract.FilterFixture));
        Assert.True(ExcelFormula2Contract.AcceptsActualSpill(ExcelFormula2Contract.FilterFixture, "E2", "E2:F10", 1, 1));
        Assert.True(ExcelFormula2Contract.AcceptsActualSpill(ExcelFormula2Contract.FilterFixture, "$E$2", "$E$2:$F$10", 1, 1));
        Assert.False(ExcelFormula2Contract.AcceptsActualSpill(ExcelFormula2Contract.FilterFixture, "E2", "G2:H10", 1, 1));
        Assert.False(ExcelFormula2Contract.AcceptsActualSpill(ExcelFormula2Contract.Sequence32Fixture, "B2", "B2:C10", 3, 2));
        Assert.Equal(new[] { "E3:F10", "F2" }, ExcelFormula2Contract.NewlyOwnedSpillRanges("E2", "E2:F10"));
        Assert.Empty(ExcelFormula2Contract.NewlyOwnedSpillRanges("B2", "E2:F10"));
        Assert.Equal("B2:C4", ExcelFormula2Contract.ResizeA1("B2", 3, 2));
        Assert.Empty(ExcelFormula2Contract.NewlyOwnedSpillRanges(ExcelFormula2Contract.ResizeA1("B2", 3, 2), "B2:C4"));
        var publicReadback = ExcelFormula2Contract.SpillReadback(
            "B2", "formula2", true, "B2:C4", "=SEQUENCE(3,2)",
            formulas: new JsonArray { new JsonArray { "=SEQUENCE(3,2)" } },
            spillValues: new JsonArray { new JsonArray { 1, 2 } });
        Assert.Equal("B2:C4", Json.GetString(publicReadback, "spillRange"));
        Assert.NotNull(publicReadback["spillValues"]);
        Assert.NotNull(publicReadback["formula2"]);
    }

    [Fact]
    public void Calculate_is_workbook_or_range_never_application_and_not_formula2_write()
    {
        var book = Json.ParseObject("""{ "op": "calculate" }""")!;
        Assert.Equal(ExcelCalculateContract.CalculateScope.Workbook, ExcelCalculateContract.ResolveScope(book));
        Assert.Contains("not Application.Calculate", ExcelCalculateContract.Describe(book), StringComparison.Ordinal);
        Assert.Contains("worksheet.Calculate", ExcelCalculateContract.Describe(book), StringComparison.Ordinal);
        Assert.True(ExcelCalculateContract.RecalculationComplete(0));
        Assert.False(ExcelCalculateContract.RecalculationComplete(1));
        Assert.False(ExcelCalculateContract.RequestsFormula2Rewrite(book));

        var range = Json.ParseObject("""{ "op": "calculate", "range": "B2:B10", "formula2": true }""")!;
        Assert.Equal(ExcelCalculateContract.CalculateScope.Range, ExcelCalculateContract.ResolveScope(range));
        Assert.True(ExcelCalculateContract.RequestsFormula2Rewrite(range));
        Assert.Contains("set_formulas", ExcelCalculateContract.Describe(range), StringComparison.Ordinal);
    }

    [Fact]
    public void Parser_bounds_cells_during_allocation_not_after_the_whole_table()
    {
        var ex = Assert.Throws<InvalidOperationException>(() =>
            ExcelCsvContract.Parse("a,b,c\r\nd,e,f", ',', maxCells: 5));
        Assert.Contains("while parsing", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Equal(2, ExcelCsvContract.Parse("a,b\r\nc,d", ',', maxCells: 4).Count);
    }

    [Fact]
    public void Delete_sheet_used_range_replay_is_not_full_restore()
    {
        Assert.Contains("formulas", ExcelDeleteSheetRecoveryContract.LostByUsedRangeReplay);
        Assert.Contains("charts", ExcelDeleteSheetRecoveryContract.LostByUsedRangeReplay);
        var missing = new JsonObject { ["workbookBackupAvailable"] = false };
        Assert.False(ExcelDeleteSheetRecoveryContract.IsFullRecoveryAvailable(missing, @"C:\tmp\none.xlsx"));
        var backup = Path.Combine(Path.GetTempPath(), "docbridge-delete-recovery-" + Guid.NewGuid().ToString("N")[..8] + ".xlsx");
        File.WriteAllBytes(backup, new byte[] { 0 });
        try
        {
            var stale = new JsonObject
            {
                ["workbookBackupAvailable"] = true,
                ["workbookBackupSource"] = "last-saved-file",
                ["workbookBackupSavedFlag"] = false,
            };
            Assert.False(ExcelDeleteSheetRecoveryContract.IsFullRecoveryAvailable(stale, backup));
            Assert.Contains("unsaved", ExcelDeleteSheetRecoveryContract.RejectReason(stale, backup),
                StringComparison.OrdinalIgnoreCase);
            var current = new JsonObject
            {
                ["workbookBackupAvailable"] = true,
                ["workbookBackupSource"] = "current-memory-savecopyas",
            };
            Assert.True(ExcelDeleteSheetRecoveryContract.IsFullRecoveryAvailable(current, backup));
        }
        finally { try { File.Delete(backup); } catch { /* temp */ } }

        var envelope = ExcelDeleteSheetRecoveryContract.RecoveryEnvelope(
            ExcelDeleteSheetRecoveryContract.ModeWorkbookCopySheet, @"C:\tmp\backup.xlsx",
            "current-memory-savecopyas", @"C:\tmp\book.xlsx",
            [new JsonObject { ["name"] = "예정공정표", ["index"] = 1 }]);
        Assert.False(Json.GetBool(envelope, "usedRangeReplayIsFullRestore"));
    }

    [Fact]
    public void Auto_fill_and_copy_range_require_filled_band_and_relative_formulas()
    {
        Assert.Equal(new[] { (2, 0), (3, 0) }, ExcelFillCopyContract.FilledBand(2, 1, 4, 1));
        Assert.Equal("=B3", ExcelFillCopyContract.ShiftA1Formula("=B1", 2, 0));
        Assert.Equal("=$B$1", ExcelFillCopyContract.ShiftA1Formula("=$B$1", 2, 3));
        Assert.Equal("=D1", ExcelFillCopyContract.ShiftA1Formula("=B1", 0, 2));
        var source = new[] { new[] { "=B1" } };
        var destShifted = new[] { new[] { "=D1" } };
        Assert.True(ExcelFillCopyContract.FormulasMatchWithRelativeShift(source, destShifted, 0, 2));
        Assert.False(ExcelFillCopyContract.FormulasMatchWithRelativeShift(source, source, 0, 2));
        var fillSource = new[] { new[] { "=A1" }, new[] { "=A2" } };
        var fillDest = new[] { new[] { "=A1" }, new[] { "=A2" }, new[] { "=A3" } };
        Assert.True(ExcelFillCopyContract.FilledFormulasContinue(fillSource, fillDest));
        var seqSource = new[] { new[] { "1" }, new[] { "2" } };
        var seqDest = new[] { new[] { "1" }, new[] { "2" }, new[] { "3" }, new[] { "4" } };
        Assert.True(ExcelFillCopyContract.FilledSequenceContinues(seqSource, seqDest));
        Assert.False(ExcelFillCopyContract.FilledSequenceContinues(seqSource, new[] { new[] { "1" }, new[] { "2" }, new[] { "9" } }));
        Assert.True(ExcelFillCopyContract.FilledAreaHasContent(seqSource, seqDest));
        Assert.False(ExcelFillCopyContract.FilledAreaHasContent(seqSource, seqSource));
    }

    [Fact]
    public void Formula2_does_not_record_blocked_spill_and_keeps_mixed_engines()
    {
        Assert.False(ExcelFormula2Contract.ShouldRecordSpill(false, "B2:C4"));
        Assert.False(ExcelFormula2Contract.ShouldRecordSpill(null, "B2:C4"));
        Assert.False(ExcelFormula2Contract.ShouldRecordSpill(true, null));
        Assert.True(ExcelFormula2Contract.ShouldRecordSpill(true, "B2:C4"));
        var mixed = new[]
        {
            Json.ParseObject("""{ "op": "set_formulas", "engine": "formula", "formulas": [["=A1"]] }""")!,
            Json.ParseObject("""{ "op": "set_formulas", "engine": "formula2", "formulas": [["=SEQUENCE(3,2)"]] }""")!,
        };
        Assert.True(ExcelFormula2Contract.MixedBatchKeepsPerOpEngine(mixed));
        Assert.Equal(ExcelFormula2Contract.EngineFormula, ExcelFormula2Contract.ResolveEngine(mixed[0]));
        Assert.Equal(ExcelFormula2Contract.EngineFormula2, ExcelFormula2Contract.ResolveEngine(mixed[1]));
    }

    [Fact]
    public void Automation_lock_is_isolated_by_home_path()
    {
        var previous = Environment.GetEnvironmentVariable("DOCBRIDGE_AUTOMATION_LOCK");
        try
        {
            Environment.SetEnvironmentVariable("DOCBRIDGE_AUTOMATION_LOCK", null);
            var first = DocBridgeHost.AutomationMutexName(@"C:\tmp\docbridge-a");
            var second = DocBridgeHost.AutomationMutexName(@"C:\tmp\docbridge-b");
            Assert.StartsWith(DocBridgeHost.AutomationMutexPrefix + ".", first);
            Assert.NotEqual(first, second);
            Assert.Equal(first, DocBridgeHost.AutomationMutexName(@"C:\tmp\docbridge-a\"));
            Environment.SetEnvironmentVariable("DOCBRIDGE_AUTOMATION_LOCK", @"Global\DocBridge.Automation.isolated-override");
            Assert.Equal(@"Global\DocBridge.Automation.isolated-override",
                DocBridgeHost.AutomationMutexName(@"C:\tmp\docbridge-a"));
        }
        finally
        {
            Environment.SetEnvironmentVariable("DOCBRIDGE_AUTOMATION_LOCK", previous);
        }
    }

    [Fact]
    public void Authoring_schema_publishes_explicit_formula_engine_and_outline_axis()
    {
        var props = ExcelAuthoringSchema.DescribeApplyOpProperties();
        var engine = Json.GetObj(props, "engine");
        Assert.Contains("formula", engine!["enum"]!.AsArray().Select(node => node!.GetValue<string>()));
        Assert.Contains("formula2", engine["enum"]!.AsArray().Select(node => node!.GetValue<string>()));
        Assert.Contains("legacy Range.Formula", engine["description"]!.GetValue<string>(), StringComparison.Ordinal);
        Assert.NotNull(Json.GetObj(props, "formulaEngine"));
        Assert.NotNull(Json.GetObj(props, "axis"));
        Assert.Contains("collapse", Json.GetString(Json.GetObj(props, "show"), "description") ?? "",
            StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Formula2_uses_SpillingToRange_and_clears_new_owned_before_old_restore()
    {
        Assert.Equal("SpillingToRange", ExcelFormula2Contract.SpillPropertySpillingToRange);
        Assert.Equal("B2:C4", ExcelFormula2Contract.PreferSpillingToRange("B2:C4", "B2"));
        Assert.Null(ExcelFormula2Contract.PreferSpillingToRange(null, "$B$2:$C$4"));
        Assert.False(ExcelFormula2Contract.TreatsFailedSpillLookupAsDestScalar);
        Assert.False(ExcelFormula2Contract.WriteCountIsRestoredValueProof);
        Assert.Null(ExcelFormula2Contract.AcceptSpillAddress("B2", "B2", spillingToRangeOk: false));
        Assert.Equal("B2:D4", ExcelFormula2Contract.AcceptSpillAddress("B2", "B2:D4", spillingToRangeOk: true));
        Assert.True(ExcelFormula2Contract.RestoreClearsNewlyOwnedBeforeOldRestore);
        Assert.True(ExcelFormula2Contract.MustNotClearNeighbors(false));
        Assert.True(ExcelFormula2Contract.MustNotClearNeighbors(null));
        Assert.False(ExcelFormula2Contract.MustNotClearNeighbors(true));
        Assert.Equal((3, 3), ExcelFormula2Contract.UnionFootprint("B2", 2, 2, "B2:D4"));
        Assert.Equal(new[] { "B3:C4", "C2" }, ExcelFormula2Contract.NeighborAndOldSpillRanges("B2", "B2:C4"));
        Assert.Equal(new[] { "B2", "B4:D4", "D2:D3" },
            ExcelFormula2Contract.ClearOrderBeforeOldRestore("B2", "B2:C3", "B2:D4", true));
        Assert.Equal(new[] { "B2" }, ExcelFormula2Contract.ClearOrderBeforeOldRestore("B2", "B2:D4", "B2:D4", false));
        Assert.True(ExcelFormula2Contract.RestoredAnchorMatches("=SEQUENCE(2,2)", "=SEQUENCE(2,2)"));
        Assert.False(ExcelFormula2Contract.RestoredAnchorMatches("=SEQUENCE(2,2)", "=SEQUENCE(3,3)"));
        Assert.True(ExcelFormula2Contract.RestoredValuesMatch(
            ExcelFormula2Contract.Sequence22Values, ExcelFormula2Contract.Sequence22Values.DeepClone()));
        Assert.False(ExcelFormula2Contract.RestoredValuesMatch(
            ExcelFormula2Contract.Sequence22Values, ExcelFormula2Contract.Sequence33Values));
        Assert.True(ExcelFormula2Contract.RestoredValuesMatch(
            ExcelFormula2Contract.FilterSpillValues, ExcelFormula2Contract.FilterSpillValues.DeepClone()));
    }

    [Fact]
    public void Delete_sheet_dependencies_are_in_workbook_not_backup_links()
    {
        Assert.True(ExcelDeleteSheetDependencyContract.FormulaReferencesDeletedSheet("=Data!A1", "Data"));
        Assert.False(ExcelDeleteSheetDependencyContract.FormulaReferencesDeletedSheet("=\"Data\"", "Data"));
        Assert.True(ExcelDeleteSheetDependencyContract.IsBackupFileExternalLink("=[backup.xlsx]Data!A1"));
        Assert.False(ExcelDeleteSheetDependencyContract.FormulaReferencesDeletedSheet("=[backup.xlsx]Data!A1", "Data"));
        Assert.True(ExcelDeleteSheetDependencyContract.FormulaReferencesDeletedSheet(
            "=[owned.xlsx]Data!A1", "Data", "owned.xlsx"));
        Assert.True(ExcelDeleteSheetDependencyContract.IsBrokenRef("=#REF!"));
        Assert.False(ExcelDeleteSheetDependencyContract.RestoredDependencyMatches("=Data!A1", "=#REF!"));
        Assert.True(ExcelDeleteSheetDependencyContract.RestoredDependencyMatches("=Data!A1", "=Data!A1"));
        Assert.True(ExcelDeleteSheetDependencyContract.NameReferencesDeletedSheet("=Data!$A$1", "Data"));
        Assert.True(ExcelDeleteSheetDependencyContract.ChartFormulaReferencesDeletedSheet(
            "=SERIES(,Data!$A$1:$A$3,Data!$B$1:$B$3,1)", "Data"));
        Assert.Equal("pivots", ExcelDeleteSheetDependencyContract.CaptureKeyPivots);
        Assert.True(ExcelDeleteSheetDependencyContract.PivotSourceReferencesDeletedSheet("Data!A1:C5", "Data"));
        Assert.True(ExcelDeleteSheetDependencyContract.PivotSourceReferencesDeletedSheet("'월별기성'!$A$1:$C$5", "월별기성"));
        Assert.False(ExcelDeleteSheetDependencyContract.PivotSourceReferencesDeletedSheet("Report!A1:C5", "Data"));
        Assert.False(ExcelDeleteSheetDependencyContract.RestoredDependencyMatches("Data!A1:C5", "#REF!"));
        Assert.False(ExcelDeleteSheetDependencyContract.IsBackupFileExternalLink("=Data!A1+SUM(Items[Amount])"));
        Assert.True(ExcelDeleteSheetDependencyContract.FormulaReferencesDeletedSheet(
            "=Data!A1+SUM(Items[Amount])", "Data"));
        Assert.True(ExcelDeleteSheetDependencyContract.FormulaReferencesDeletedSheet(
            "='[owned.xlsx]원장'!R1C1:R6C3", "원장", "owned.xlsx"));
        Assert.False(ExcelDeleteSheetDependencyContract.IsBackupFileExternalLink(
            "='[owned.xlsx]원장'!A1", "owned.xlsx"));
        Assert.True(ExcelDeleteSheetDependencyContract.IsBackupFileExternalLink(
            "='[owned.xlsx]원장'!A1", "other.xlsx"));
        Assert.True(ExcelDeleteSheetDependencyContract.PivotSourceReferencesDeletedSheet(
            "'[owned.xlsx]원장'!A1:C6", "원장", "owned.xlsx"));
        Assert.True(ExcelDeleteSheetDependencyContract.FormulaReferencesDeletedSheet("=Data:Summary!A1", "Data"));
        Assert.True(ExcelDeleteSheetDependencyContract.FormulaReferencesDeletedSheet("='월별기성:요약'!A1", "월별기성"));
        Assert.True(ExcelDeleteSheetDependencyContract.CanFaithfullyCapture(
            "=Data!A1+SUM(Items[Amount])", "owned.xlsx", out _));
        Assert.False(ExcelDeleteSheetDependencyContract.CanFaithfullyCapture("=[", null, out var limitation));
        Assert.Contains(ExcelDeleteSheetDependencyContract.UncapturedError, limitation);
        var refused = Assert.Throws<InvalidOperationException>(() =>
            ExcelDeleteSheetDependencyContract.RequireFaithfulCapture("=[", null));
        Assert.True(ExcelDeleteSheetDependencyContract.IsUncaptured(refused));
        ExcelDeleteSheetDependencyContract.RequireFaithfulCapture("=Data!A1+SUM(Items[Amount])", "owned.xlsx");
        var sheetOrder = new[] { "Sheet1", "Sheet2", "Sheet3", "Sheet4" };
        Assert.True(ExcelDeleteSheetDependencyContract.FormulaReferencesDeletedSheet(
            "=SUM(Sheet1:Sheet3!A1)", "Sheet2", null, sheetOrder));
        Assert.False(ExcelDeleteSheetDependencyContract.FormulaReferencesDeletedSheet(
            "=SUM(Sheet1:Sheet3!A1)", "Sheet4", null, sheetOrder));
        Assert.True(ExcelDeleteSheetDependencyContract.FormulaReferencesDeletedSheet("='123'!A1", "123"));
        Assert.True(ExcelDeleteSheetDependencyContract.FormulaReferencesDeletedSheet(
            "=Data!A1+'[external.xlsx]Sheet1'!A1", "Data"));
        Assert.True(ExcelDeleteSheetDependencyContract.CanFaithfullyCapture(
            "=SUM(Items[[#Data],[Amount]])", null, out _));
        Assert.False(ExcelDeleteSheetDependencyContract.FormulaReferencesDeletedSheet(
            "=SUM(Items[[#Data],[Amount]])", "Data"));
        Assert.True(ExcelDeleteSheetDependencyContract.FormulaReferencesDeletedSheet(
            "=Data!A1+SUM(Items[[#Data],[Amount]])", "Data"));
        var captureFailed = ExcelDeleteSheetDependencyContract.CaptureFailed("formula read failed");
        Assert.True(ExcelDeleteSheetDependencyContract.IsUncaptured(captureFailed));
    }

    [Fact]
    public void Formula2_stateful_restore_clears_new_spill_then_restores_old_values_not_write_count()
    {
        var grid = new FakeSpillGrid();
        grid.Write("B2", "=SEQUENCE(2,2)", ExcelFormula2Contract.Sequence22Values);
        grid.Write("B2", "=SEQUENCE(3,3)", ExcelFormula2Contract.Sequence33Values);
        grid.Neighbor["C5"] = "keep";
        var order = ExcelFormula2Contract.ClearOrderBeforeOldRestore("B2", "B2:C3", "B2:D4", true);
        Assert.Equal(new[] { "B2", "B4:D4", "D2:D3" }, order);
        foreach (var extra in order)
            grid.Clear(extra);
        grid.Write("B2", "=SEQUENCE(2,2)", ExcelFormula2Contract.Sequence22Values);
        Assert.True(ExcelFormula2Contract.RestoredAnchorMatches("=SEQUENCE(2,2)", grid.Formula["B2"]));
        Assert.True(ExcelFormula2Contract.RestoredValuesMatch(
            ExcelFormula2Contract.Sequence22Values, grid.Values["B2"]));
        Assert.Equal("keep", grid.Neighbor["C5"]);
        Assert.False(ExcelFormula2Contract.WriteCountIsRestoredValueProof);
        Assert.True(ExcelFormula2Contract.RestoredValuesMatch(
            ExcelFormula2Contract.FilterSpillValues, ExcelFormula2Contract.FilterSpillValues.DeepClone()));
        Assert.True(ExcelFormula2Contract.MustNotClearNeighbors(hasSpill: false));
    }

    private sealed class FakeSpillGrid
    {
        public Dictionary<string, string> Formula { get; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, JsonNode> Values { get; } = new(StringComparer.OrdinalIgnoreCase);
        public Dictionary<string, string> Neighbor { get; } = new(StringComparer.OrdinalIgnoreCase);

        public void Write(string dest, string formula, JsonNode values)
        {
            Formula[dest] = formula;
            Values[dest] = values.DeepClone();
        }

        public void Clear(string range)
        {
            Formula.Remove(range);
            Values.Remove(range);
        }
    }
}
