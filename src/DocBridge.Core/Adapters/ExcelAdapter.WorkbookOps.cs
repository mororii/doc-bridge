using System.Globalization;
using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

/// <summary>
/// Excel workbook-level and advanced-object ops: external links, calculation
/// mode, freeze/paste-special, goal seek, workbook protection, split panes,
/// sparklines, slicers, and cell styles. No VBA or macro execution.
/// </summary>
public sealed partial class ExcelAdapter
{
    // ------------------------------------------------------------------
    // Shared readers
    // ------------------------------------------------------------------

    internal static List<JsonObject> ReadExcelLinks(object workbook)
    {
        object? sources = null;
        var items = new List<JsonObject>();
        try
        {
            try
            {
                sources = (object)((dynamic)workbook).LinkSources(ExcelWorkbookOpsContract.XlExcelLinks);
            }
            catch
            {
                return items;
            }
            if (sources is not Array arr) return items;
            var count = 0;
            foreach (var item in arr)
            {
                if (count >= ExcelWorkbookOpsContract.MaxLinkSources) break;
                var source = Convert.ToString(item, CultureInfo.InvariantCulture) ?? "";
                if (string.IsNullOrWhiteSpace(source)) continue;
                var status = ReadLinkStatus(workbook, source);
                items.Add(new JsonObject
                {
                    ["source"] = source,
                    ["leaf"] = ExcelWorkbookOpsContract.LeafOf(source),
                    ["type"] = "excelLink",
                    ["status"] = status is null ? null : JsonValue.Create(status.Value),
                    ["statusName"] = status is null ? "unreadable" : ExcelWorkbookOpsContract.LinkStatusName(status.Value),
                });
                count++;
            }
            return items;
        }
        finally { RotHelper.ReleaseComReference(sources); }
    }

    private static int? ReadLinkStatus(object workbook, string source)
    {
        try
        {
            return Convert.ToInt32(
                ((dynamic)workbook).LinkInfo(source, ExcelWorkbookOpsContract.XlLinkInfoStatus),
                CultureInfo.InvariantCulture);
        }
        catch { return null; }
    }

    private static string ResolveUniqueLink(object workbook, string requested)
    {
        var sources = ReadExcelLinks(workbook)
            .Select(node => Json.GetString(node, "source") ?? "")
            .Where(source => !string.IsNullOrWhiteSpace(source))
            .ToList();
        var exact = sources.FirstOrDefault(source =>
            string.Equals(source, requested, StringComparison.OrdinalIgnoreCase));
        if (exact is not null) return exact;
        var leaf = ExcelWorkbookOpsContract.LeafOf(requested);
        var matches = sources.Where(source =>
                string.Equals(ExcelWorkbookOpsContract.LeafOf(source), string.IsNullOrWhiteSpace(leaf) ? requested : leaf, StringComparison.OrdinalIgnoreCase) ||
                source.IndexOf(requested, StringComparison.OrdinalIgnoreCase) >= 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (matches.Count == 1) return matches[0];
        if (matches.Count == 0)
            throw new InvalidOperationException(
                $"[EXCEL_LINK_NOT_FOUND] link '{requested}' was not found; known: {string.Join("; ", sources.Take(10))}");
        throw new InvalidOperationException(
            $"[EXCEL_LINK_AMBIGUOUS] link '{requested}' matches {matches.Count}; use the full source: {string.Join("; ", matches.Take(10))}");
    }

    private sealed record LinkDependent(string Sheet, string Address, string Formula);

    private static List<LinkDependent> ScanLinkDependents(object workbook, string leaf)
    {
        var result = new List<LinkDependent>();
        foreach (var sheet in EnumerateWorksheets(workbook, null))
        {
            object? used = null;
            try
            {
                try { used = (object)((dynamic)sheet).UsedRange; }
                catch { continue; }
                if (used is null) continue;
                object? raw = null;
                try { raw = ((dynamic)used).Formula; }
                catch { continue; }
                if (raw is object[,] arr)
                {
                    var r1 = arr.GetLowerBound(0); var r2 = arr.GetUpperBound(0);
                    var c1 = arr.GetLowerBound(1); var c2 = arr.GetUpperBound(1);
                    object? usedAddress = null;
                    string? baseAddress = null;
                    try
                    {
                        usedAddress = (object)((dynamic)used).Address(false, false);
                        baseAddress = Convert.ToString(usedAddress, CultureInfo.InvariantCulture);
                    }
                    finally { RotHelper.ReleaseComReference(usedAddress); }
                    if (string.IsNullOrWhiteSpace(baseAddress)) continue;
                    if (!ExcelDataOperationsContract.TryParseA1(baseAddress, out var baseBox)) continue;
                    for (var r = r1; r <= r2; r++)
                    {
                        for (var c = c1; c <= c2; c++)
                        {
                            if (result.Count >= ExcelWorkbookOpsContract.MaxLinkDependents) return result;
                            var formula = arr[r, c] as string;
                            if (!ExcelWorkbookOpsContract.LinkRefersTo(formula, leaf)) continue;
                            var address = $"{ColName(baseBox.Column + (c - c1))}{baseBox.Row + (r - r1)}";
                            result.Add(new LinkDependent(
                                ReadComString(sheet, "Name") ?? "", address, formula!));
                        }
                    }
                }
                else if (raw is string single &&
                         ExcelWorkbookOpsContract.LinkRefersTo(single, leaf))
                {
                    if (result.Count >= ExcelWorkbookOpsContract.MaxLinkDependents) return result;
                    object? addr = null;
                    try
                    {
                        addr = (object)((dynamic)used).Address(false, false);
                        result.Add(new LinkDependent(
                            ReadComString(sheet, "Name") ?? "",
                            Convert.ToString(addr, CultureInfo.InvariantCulture) ?? "",
                            single));
                    }
                    finally { RotHelper.ReleaseComReference(addr); }
                }
            }
            finally
            {
                RotHelper.ReleaseComReference(used);
                RotHelper.ReleaseComReference(sheet);
            }
        }
        return result;
    }

    private static object GetWorkbookApplication(object workbook)
    {
        try { return (object)((dynamic)workbook).Application; }
        catch (Exception ex) { throw new InvalidOperationException("[EXCEL_APP_UNAVAILABLE] workbook application is unavailable: " + ex.Message); }
    }

    private static int ReadCalculationMode(object application) =>
        Convert.ToInt32(((dynamic)application).Calculation, CultureInfo.InvariantCulture);

    private static void RecalculateApplication(object application, string? recalculate)
    {
        var token = (recalculate ?? "none").Trim().ToLowerInvariant();
        if (token is "full") ((dynamic)application).CalculateFull();
        else if (token is "fullrebuild") ((dynamic)application).CalculateFullRebuild();
    }

    private static void VerifyCalculationIdle(object application, string label, List<string> mismatches)
    {
        try
        {
            var state = Convert.ToInt32(((dynamic)application).CalculationState, CultureInfo.InvariantCulture);
            if (!ExcelCalculateContract.RecalculationComplete(state))
                mismatches.Add($"{label}: CalculationState {state} is not xlDone");
        }
        catch { /* CalculationState is optional verification */ }
    }

    private static bool WorkbookHasPassword(object workbook)
    {
        try { return Convert.ToBoolean(((dynamic)workbook).HasPassword, CultureInfo.InvariantCulture); }
        catch { return false; }
    }

    private static JsonObject ReadWorkbookProtection(object workbook) => new()
    {
        ["structure"] = TryComBool(workbook, "ProtectStructure") ?? false,
        ["windows"] = TryComBool(workbook, "ProtectWindows") ?? false,
    };

    private static object? FindPivotInWorkbook(object workbook, string name, out object? owningSheet)
    {
        owningSheet = null;
        foreach (var sheet in EnumerateWorksheets(workbook, null))
        {
            foreach (var pivot in EnumeratePivotTables(sheet))
            {
                var actual = ReadComString(pivot, "Name");
                if (string.Equals(actual, name, StringComparison.OrdinalIgnoreCase))
                {
                    owningSheet = sheet;
                    return pivot;
                }
                RotHelper.ReleaseComReference(pivot);
            }
            RotHelper.ReleaseComReference(sheet);
        }
        return null;
    }

    // Worksheet.Slicers is unreliable (it can report an empty collection while
    // the slicer shape is live on the sheet), so slicers are always resolved
    // through their SlicerCache, which is also where inspect reads them.
    private static object? FindSlicer(object workbook, string name, out object? owningSheet)
    {
        owningSheet = null;
        foreach (var cache in EnumerateSlicerCaches(workbook))
        {
            object? owned = null;
            var claimed = false;
            try
            {
                try { owned = (object)((dynamic)cache).Slicers; }
                catch { continue; }
                var count = Convert.ToInt32(((dynamic)owned).Count, CultureInfo.InvariantCulture);
                for (var i = 1; i <= count; i++)
                {
                    object? slicer = null;
                    try
                    {
                        slicer = (object)((dynamic)owned).Item(i);
                        if (string.Equals(ReadComString(slicer, "Name"), name, StringComparison.OrdinalIgnoreCase))
                        {
                            try { owningSheet = (object)((dynamic)slicer).Parent; }
                            catch { owningSheet = null; }
                            claimed = true;
                            return slicer;
                        }
                    }
                    finally
                    {
                        if (!claimed) RotHelper.ReleaseComReference(slicer);
                    }
                }
            }
            finally
            {
                RotHelper.ReleaseComReference(owned);
                RotHelper.ReleaseComReference(cache);
            }
        }
        return null;
    }

    // ------------------------------------------------------------------
    // update_external_links
    // ------------------------------------------------------------------

    private static void PreviewUpdateExternalLinks(object workbook, JsonObject op, ApplyPreview preview)
    {
        var links = ReadExcelLinks(workbook);
        var matched = FilterLinks(links, op);
        if (matched.Count == 0)
            throw new InvalidOperationException("[EXCEL_LINK_NOT_FOUND] no Excel links match the requested filter");
        foreach (var link in matched)
        {
            var status = Json.GetInt(link, "status");
            if (status is not null && !ExcelWorkbookOpsContract.IsUpdatableLinkStatus(status.Value))
                throw new InvalidOperationException(
                    $"[EXCEL_LINK_UNREACHABLE] link '{Json.GetString(link, "source")}' status is {ExcelWorkbookOpsContract.LinkStatusName(status.Value)}; update is refused");
        }
        preview.Affected.Add(new AffectedRef("links", $"{matched.Count} link(s)"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = "workbook:links",
            Before = new JsonArray(matched.Select(node => node.DeepClone()).ToArray()),
            After = JsonValue.Create("refresh cached values"),
        });
    }

    private static List<JsonObject> FilterLinks(IEnumerable<JsonObject> links, JsonObject op)
    {
        var source = Json.GetString(op, "source");
        var contains = Json.GetString(op, "sourceContains");
        var result = new List<JsonObject>();
        foreach (var node in links)
        {
            var candidate = Json.GetString(node, "source") ?? "";
            if (!string.IsNullOrWhiteSpace(source) &&
                !string.Equals(candidate, source, StringComparison.OrdinalIgnoreCase)) continue;
            if (!string.IsNullOrWhiteSpace(contains) &&
                candidate.IndexOf(contains, StringComparison.OrdinalIgnoreCase) < 0) continue;
            result.Add(node);
        }
        return result;
    }

    private static void ApplyUpdateExternalLinks(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        var matched = FilterLinks(ReadExcelLinks(workbook), op);
        if (matched.Count == 0)
            throw new InvalidOperationException("[EXCEL_LINK_NOT_FOUND] no Excel links match the requested filter");
        foreach (var link in matched)
        {
            var source = Json.GetString(link, "source")!;
            try { ((dynamic)workbook).UpdateLink(source, ExcelWorkbookOpsContract.XlExcelLinks); }
            catch (Exception ex) { throw new InvalidOperationException($"[EXCEL_LINK_UPDATE_FAILED] '{source}': {ex.Message}"); }
            checkedCells++;
            execution.Affected.Add(new AffectedRef("link", source));
        }
        object? application = null;
        try
        {
            application = GetWorkbookApplication(workbook);
            VerifyCalculationIdle(application, "update_external_links", mismatches);
        }
        finally { RotHelper.ReleaseComReference(application); }
        var after = ReadExcelLinks(workbook);
        foreach (var link in matched)
        {
            var source = Json.GetString(link, "source")!;
            var current = after.FirstOrDefault(node =>
                string.Equals(Json.GetString(node, "source"), source, StringComparison.OrdinalIgnoreCase));
            if (current is null) { mismatches.Add($"{source}: link disappeared after update"); continue; }
            var status = Json.GetInt(current, "status");
            if (status is not null && status is 1 or 2 or 7)
                mismatches.Add($"{source}: status {ExcelWorkbookOpsContract.LinkStatusName(status.Value)} after update");
        }
    }

    // ------------------------------------------------------------------
    // change_link_source / break_external_link
    // ------------------------------------------------------------------

    private static void PreviewChangeLinkSource(object workbook, JsonObject op, ApplyPreview preview)
    {
        var resolved = ResolveUniqueLink(workbook, Json.GetString(op, "source")!);
        preview.Affected.Add(new AffectedRef("link", resolved));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"link:{resolved}",
            Before = JsonValue.Create(resolved),
            After = JsonValue.Create(Json.GetString(op, "newSource")),
        });
    }

    private static void ApplyChangeLinkSource(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        var resolved = ResolveUniqueLink(workbook, Json.GetString(op, "source")!);
        var target = Json.GetString(op, "newSource")!;
        try { ((dynamic)workbook).ChangeLink(resolved, target, ExcelWorkbookOpsContract.XlExcelLinks); }
        catch (Exception ex) { throw new InvalidOperationException($"[EXCEL_LINK_CHANGE_FAILED] '{resolved}' -> '{target}': {ex.Message}"); }
        object? application = null;
        try
        {
            application = GetWorkbookApplication(workbook);
            try { RecalculateApplication(application, "full"); } catch { /* recalc is best effort */ }
            VerifyCalculationIdle(application, "change_link_source", mismatches);
        }
        finally { RotHelper.ReleaseComReference(application); }
        var sources = ReadExcelLinks(workbook)
            .Select(node => Json.GetString(node, "source") ?? "").ToList();
        if (!sources.Any(source => string.Equals(source, target, StringComparison.OrdinalIgnoreCase)))
            mismatches.Add($"new source '{target}' was not read back after ChangeLink");
        checkedCells++;
        execution.Affected.Add(new AffectedRef("link", $"{resolved} -> {target}"));
    }

    private static void PreviewBreakExternalLink(object workbook, JsonObject op, ApplyPreview preview)
    {
        var resolved = ResolveUniqueLink(workbook, Json.GetString(op, "source")!);
        var dependents = ScanLinkDependents(workbook, ExcelWorkbookOpsContract.LeafOf(resolved));
        if (dependents.Count >= ExcelWorkbookOpsContract.MaxLinkDependents)
            throw new InvalidOperationException(
                $"[EXCEL_LINK_DEPENDENTS_TOO_LARGE] {dependents.Count} dependent cells reach the {ExcelWorkbookOpsContract.MaxLinkDependents} snapshot bound; narrow the workbook first");
        preview.Affected.Add(new AffectedRef("link", resolved));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"link:{resolved}",
            Before = JsonValue.Create($"bound with {dependents.Count} dependent cell(s)"),
            After = JsonValue.Create("broken; dependents keep cached values"),
        });
        if (dependents.Count == 0)
            preview.Warnings.Add($"link '{resolved}' has no dependent formulas in this workbook");
    }

    private static void ApplyBreakExternalLink(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        var resolved = ResolveUniqueLink(workbook, Json.GetString(op, "source")!);
        var dependents = ScanLinkDependents(workbook, ExcelWorkbookOpsContract.LeafOf(resolved));
        if (dependents.Count >= ExcelWorkbookOpsContract.MaxLinkDependents)
            throw new InvalidOperationException(
                $"[EXCEL_LINK_DEPENDENTS_TOO_LARGE] dependents changed since preview and reach the snapshot bound");
        try { ((dynamic)workbook).BreakLink(resolved, ExcelWorkbookOpsContract.XlExcelLinks); }
        catch (Exception ex) { throw new InvalidOperationException($"[EXCEL_LINK_BREAK_FAILED] '{resolved}': {ex.Message}"); }
        var remaining = ReadExcelLinks(workbook)
            .Any(node => string.Equals(Json.GetString(node, "source"), resolved, StringComparison.OrdinalIgnoreCase));
        if (remaining)
            mismatches.Add($"link '{resolved}' is still present after BreakLink");
        checkedCells += Math.Max(dependents.Count, 1);
        execution.Affected.Add(new AffectedRef("link", $"{resolved} broken ({dependents.Count} dependent(s))"));
    }

    // ------------------------------------------------------------------
    // set_calculation_mode
    // ------------------------------------------------------------------

    private static void PreviewSetCalculationMode(object workbook, JsonObject op, ApplyPreview preview)
    {
        object? application = null;
        try
        {
            application = GetWorkbookApplication(workbook);
            var before = ReadCalculationMode(application);
            ExcelWorkbookOpsContract.TryCalculationMode(Json.GetString(op, "mode"), out var after);
            preview.Affected.Add(new AffectedRef("calculation", $"mode {before} -> {after}"));
            preview.Diff.Add(new DiffEntry
            {
                Ref = "application:calculation",
                Before = JsonValue.Create(before),
                After = JsonValue.Create(after),
            });
        }
        finally { RotHelper.ReleaseComReference(application); }
    }

    private static void ApplySetCalculationMode(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        object? application = null;
        try
        {
            application = GetWorkbookApplication(workbook);
            ExcelWorkbookOpsContract.TryCalculationMode(Json.GetString(op, "mode"), out var mode);
            ((dynamic)application).Calculation = mode;
            try { RecalculateApplication(application, Json.GetString(op, "recalculate")); }
            catch (Exception ex) { throw new InvalidOperationException($"[EXCEL_CALCULATE_FAILED] recalculate failed: {ex.Message}"); }
            var actual = ReadCalculationMode(application);
            checkedCells++;
            if (actual != mode)
                mismatches.Add($"calculation mode readback {actual} != requested {mode}");
            VerifyCalculationIdle(application, "set_calculation_mode", mismatches);
            execution.Affected.Add(new AffectedRef("calculation", $"mode {mode}"));
        }
        finally { RotHelper.ReleaseComReference(application); }
    }

    // ------------------------------------------------------------------
    // freeze_values
    // ------------------------------------------------------------------

    private static void PreviewFreezeValues(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var range = BindRange(workbook, op, "range");
        EnsureSheetWritable(range.Sheet);
        RejectMergedSurface(range.Sheet, range.Address, "freeze_values");
        var info = DescribeFormulaSurface(range.Sheet, range.Address);
        preview.Affected.Add(new AffectedRef("range", $"{range.SheetName}!{range.Address}"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{range.SheetName}!{range.Address}",
            Before = JsonValue.Create($"{info.FormulaCells} formula cell(s) of {info.TotalCells}"),
            After = JsonValue.Create("cached values"),
        });
        if (info.ErrorCells > 0)
            preview.Warnings.Add($"{info.ErrorCells} error cell(s) keep their formulas; errors cannot be written as values");
        if (info.FormulaCells == 0)
            preview.Warnings.Add("range contains no formulas; freeze is a no-op");
    }

    private sealed record FormulaSurfaceInfo(int TotalCells, int FormulaCells, int ErrorCells);

    private static void RejectMergedSurface(object sheet, string address, string opName)
    {
        object? range = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            object? merged;
            try { merged = ((dynamic)range).MergeCells; }
            catch (Exception ex) { throw new InvalidOperationException($"[EXCEL_MERGE_CHECK] {opName} could not read MergeCells: {ex.Message}"); }
            if (merged is null || (merged is bool flag ? flag : true))
                throw new InvalidOperationException($"[EXCEL_MERGED_RANGE] {opName} refuses merged ranges: {address}");
        }
        finally { RotHelper.ReleaseComReference(range); }
    }

    private static FormulaSurfaceInfo DescribeFormulaSurface(object sheet, string address)
    {
        object? range = null;
        object? formulas = null;
        object? errorCells = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            var total = Convert.ToInt32(((dynamic)range).Count, CultureInfo.InvariantCulture);
            formulas = (object)((dynamic)range).Formula;
            var formulaCount = 0;
            if (formulas is object[,] arr)
            {
                foreach (var item in arr)
                {
                    if (item is string text && text.StartsWith('=')) formulaCount++;
                }
            }
            else if (formulas is string single && single.StartsWith('=')) formulaCount = 1;
            var errorCount = 0;
            try
            {
                errorCells = (object)((dynamic)range).SpecialCells(-4123, 16);
                errorCount = Convert.ToInt32(((dynamic)errorCells).Count, CultureInfo.InvariantCulture);
            }
            catch { errorCount = 0; }
            return new FormulaSurfaceInfo(total, formulaCount, errorCount);
        }
        finally
        {
            RotHelper.ReleaseComReference(errorCells);
            RotHelper.ReleaseComReference(formulas);
            RotHelper.ReleaseComReference(range);
        }
    }

    private static void ApplyFreezeValues(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var range = BindRange(workbook, op, "range");
        EnsureSheetWritable(range.Sheet);
        RejectMergedSurface(range.Sheet, range.Address, "freeze_values");
        object? comRange = null;
        object? formulaRaw = null;
        object? valueRaw = null;
        object? errorCells = null;
        try
        {
            comRange = (object)((dynamic)range.Sheet).Range(range.Address);
            formulaRaw = (object)((dynamic)comRange).Formula;
            valueRaw = (object)((dynamic)comRange).Value2;
            var errorKeys = new HashSet<(int Row, int Col)>();
            try
            {
                errorCells = (object)((dynamic)comRange).SpecialCells(-4123, 16);
                foreach (var cell in EnumerateRangeCells(errorCells))
                {
                    object? addr = null;
                    try
                    {
                        var row = Convert.ToInt32(((dynamic)cell).Row, CultureInfo.InvariantCulture);
                        var col = Convert.ToInt32(((dynamic)cell).Column, CultureInfo.InvariantCulture);
                        errorKeys.Add((row, col));
                    }
                    finally
                    {
                        RotHelper.ReleaseComReference(addr);
                        RotHelper.ReleaseComReference(cell);
                    }
                }
            }
            catch { /* no error cells */ }

            if (formulaRaw is not object[,] formulaArr || valueRaw is not object[,] valueArr)
            {
                FreezeSingleCell(comRange, formulaRaw as string, valueRaw, mismatches, range, execution, ref checkedCells);
                return;
            }
            var r1 = formulaArr.GetLowerBound(0); var r2 = formulaArr.GetUpperBound(0);
            var c1 = formulaArr.GetLowerBound(1); var c2 = formulaArr.GetUpperBound(1);
            if (errorKeys.Count == 0)
            {
                var grid = new object[r2 - r1 + 1, c2 - c1 + 1];
                for (var r = r1; r <= r2; r++)
                    for (var c = c1; c <= c2; c++)
                        grid[r - r1, c - c1] = valueArr[r, c] ?? "";
                ((dynamic)comRange).Value2 = grid;
            }
            else
            {
                var skipped = new List<string>();
                var frozen = 0;
                for (var r = r1; r <= r2; r++)
                {
                    for (var c = c1; c <= c2; c++)
                    {
                        if (formulaArr[r, c] is not string formula || !formula.StartsWith('=')) continue;
                        object? cell = null;
                        try
                        {
                            cell = (object)((dynamic)comRange).Item(r - r1 + 1, c - c1 + 1);
                            var absRow = Convert.ToInt32(((dynamic)cell).Row, CultureInfo.InvariantCulture);
                            var absCol = Convert.ToInt32(((dynamic)cell).Column, CultureInfo.InvariantCulture);
                            if (errorKeys.Contains((absRow, absCol)))
                            {
                                if (skipped.Count < 50)
                                    skipped.Add(Convert.ToString(((dynamic)cell).Address(false, false), CultureInfo.InvariantCulture) ?? "");
                                continue;
                            }
                            ((dynamic)cell).Value2 = valueArr[r, c] ?? "";
                            frozen++;
                        }
                        finally { RotHelper.ReleaseComReference(cell); }
                    }
                }
                checkedCells += frozen;
                if (skipped.Count > 0)
                    execution.Warnings.Add($"freeze skipped {skipped.Count} error cell(s), formulas kept: {string.Join(", ", skipped)}");
                VerifyNoFormulasExcept(comRange, errorKeys, mismatches, $"{range.SheetName}!{range.Address}");
                execution.Affected.Add(new AffectedRef("range", $"{range.SheetName}!{range.Address}"));
                return;
            }
            checkedCells += (r2 - r1 + 1) * (c2 - c1 + 1);
            VerifyNoFormulas(comRange, mismatches, $"{range.SheetName}!{range.Address}");
            execution.Affected.Add(new AffectedRef("range", $"{range.SheetName}!{range.Address}"));
        }
        finally
        {
            RotHelper.ReleaseComReference(errorCells);
            RotHelper.ReleaseComReference(valueRaw);
            RotHelper.ReleaseComReference(formulaRaw);
            RotHelper.ReleaseComReference(comRange);
        }
    }

    private static void FreezeSingleCell(object comRange, string? formula, object? cached,
        List<string> mismatches, DataRangeLease range, ApplyExecution execution, ref int checkedCells)
    {
        if (formula is not null && formula.StartsWith('='))
        {
            bool isError;
            try { isError = IsCachedError(comRange); }
            catch { isError = false; }
            if (isError)
            {
                execution.Warnings.Add("single error cell keeps its formula; errors cannot be written as values");
                execution.Affected.Add(new AffectedRef("range", $"{range.SheetName}!{range.Address}"));
                return;
            }
            ((dynamic)comRange).Value2 = cached ?? "";
        }
        checkedCells++;
        VerifyNoFormulas(comRange, mismatches, $"{range.SheetName}!{range.Address}");
        execution.Affected.Add(new AffectedRef("range", $"{range.SheetName}!{range.Address}"));
    }

    private static bool IsCachedError(object comRange)
    {
        object? errors = null;
        try
        {
            errors = (object)((dynamic)comRange).SpecialCells(-4123, 16);
            return Convert.ToInt32(((dynamic)errors).Count, CultureInfo.InvariantCulture) > 0;
        }
        catch { return false; }
        finally { RotHelper.ReleaseComReference(errors); }
    }

    private static IEnumerable<object> EnumerateRangeCells(object range)
    {
        object? cells = null;
        try
        {
            cells = (object)((dynamic)range).Cells;
            var count = Convert.ToInt32(((dynamic)cells).Count, CultureInfo.InvariantCulture);
            for (var i = 1; i <= count; i++)
                yield return (object)((dynamic)cells).Item(i);
        }
        finally { RotHelper.ReleaseComReference(cells); }
    }

    private static void VerifyNoFormulas(object comRange, List<string> mismatches, string label)
    {
        object? formulas = null;
        try
        {
            formulas = (object)((dynamic)comRange).Formula;
            if (formulas is object[,] arr)
            {
                foreach (var item in arr)
                {
                    if (item is string text && text.StartsWith('='))
                    {
                        mismatches.Add($"{label}: formulas remain after freeze");
                        return;
                    }
                }
            }
            else if (formulas is string single && single.StartsWith('='))
                mismatches.Add($"{label}: formula remains after freeze");
        }
        finally { RotHelper.ReleaseComReference(formulas); }
    }

    private static void VerifyNoFormulasExcept(object comRange, HashSet<(int Row, int Col)> allowed,
        List<string> mismatches, string label)
    {
        object? formulas = null;
        try
        {
            formulas = (object)((dynamic)comRange).Formula;
            if (formulas is not object[,] arr) return;
            var r1 = arr.GetLowerBound(0); var r2 = arr.GetUpperBound(0);
            var c1 = arr.GetLowerBound(1); var c2 = arr.GetUpperBound(1);
            for (var r = r1; r <= r2; r++)
            {
                for (var c = c1; c <= c2; c++)
                {
                    if (arr[r, c] is not string text || !text.StartsWith('=')) continue;
                    object? cell = null;
                    try
                    {
                        cell = (object)((dynamic)comRange).Item(r - r1 + 1, c - c1 + 1);
                        var absRow = Convert.ToInt32(((dynamic)cell).Row, CultureInfo.InvariantCulture);
                        var absCol = Convert.ToInt32(((dynamic)cell).Column, CultureInfo.InvariantCulture);
                        if (!allowed.Contains((absRow, absCol)))
                        {
                            mismatches.Add($"{label}: unexpected formula remains after freeze");
                            return;
                        }
                    }
                    finally { RotHelper.ReleaseComReference(cell); }
                }
            }
        }
        finally { RotHelper.ReleaseComReference(formulas); }
    }

    // ------------------------------------------------------------------
    // paste_special
    // ------------------------------------------------------------------

    private static void PreviewPasteSpecial(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var source = BindRange(workbook, op, "sourceRange", independentSource: true);
        using var dest = BindRange(workbook, op, "destination");
        EnsureSheetWritable(dest.Sheet);
        var plan = PlanPasteSpecial(op, source.Address, dest.Address);
        RejectMergedSurface(dest.Sheet, plan.WriteAddress, "paste_special");
        preview.Affected.Add(new AffectedRef("range", $"{dest.SheetName}!{plan.WriteAddress}"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{dest.SheetName}!{plan.WriteAddress}",
            Before = JsonValue.Create("current destination"),
            After = JsonValue.Create($"{plan.Paste} from {source.SheetName}!{source.Address}"),
        });
    }

    private sealed record PastePlan(string WriteAddress, int Rows, int Cols, string Paste, bool Transpose,
        bool SkipBlanks, int Operation);

    private static PastePlan PlanPasteSpecial(JsonObject op, string sourceAddress, string destAddress)
    {
        if (!ExcelDataOperationsContract.TryParseA1(sourceAddress, out var sourceBox, allowUnion: false) ||
            !ExcelDataOperationsContract.TryParseA1(destAddress, out var destBox, allowUnion: false))
            throw new InvalidOperationException("[EXCEL_PASTE_ADDRESS] source or destination is not a rectangular A1 range");
        ExcelWorkbookOpsContract.TryPasteType(Json.GetString(op, "paste"), out var paste);
        var transpose = Json.GetBool(op, "transpose");
        ExcelWorkbookOpsContract.TryPasteOperation(Json.GetString(op, "operation"), out var operation);
        if (transpose && paste is "formats" or "all")
            throw new InvalidOperationException("[EXCEL_PASTE_TRANSPOSE_FORMATS] transpose supports paste=values|formulas only; use format_range for transposed formats");
        if (operation != 0 && paste is "formulas" or "formats")
            throw new InvalidOperationException("[EXCEL_PASTE_OPERATION] operation requires paste=values|all");
        var srcRows = sourceBox.Rows;
        var srcCols = sourceBox.Columns;
        var needRows = transpose ? srcCols : srcRows;
        var needCols = transpose ? srcRows : srcCols;
        string writeAddress;
        if (destBox.Rows == 1 && destBox.Columns == 1)
        {
            writeAddress = $"{ColName(destBox.Column)}{destBox.Row}:{ColName(destBox.Column + needCols - 1)}{destBox.Row + needRows - 1}";
        }
        else
        {
            if (destBox.Rows < needRows || destBox.Columns < needCols)
                throw new InvalidOperationException(
                    $"[EXCEL_PASTE_FIT] destination {destAddress} is smaller than the {(transpose ? "transposed " : "")}source {sourceAddress}");
            writeAddress = destAddress;
        }
        return new PastePlan(writeAddress, needRows, needCols, paste, transpose,
            Json.GetBool(op, "skipBlanks"), operation);
    }

    private static void ApplyPasteSpecial(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var source = BindRange(workbook, op, "sourceRange", independentSource: true);
        using var dest = BindRange(workbook, op, "destination");
        EnsureSheetWritable(dest.Sheet);
        var plan = PlanPasteSpecial(op, source.Address, dest.Address);
        RejectMergedSurface(dest.Sheet, plan.WriteAddress, "paste_special");
        object? srcRange = null;
        object? dstRange = null;
        object? srcFormulas = null;
        object? srcValues = null;
        object? dstFormulas = null;
        object? dstValues = null;
        try
        {
            srcRange = (object)((dynamic)source.Sheet).Range(source.Address);
            dstRange = (object)((dynamic)dest.Sheet).Range(plan.WriteAddress);
            srcFormulas = (object)((dynamic)srcRange).Formula;
            srcValues = (object)((dynamic)srcRange).Value2;
            dstFormulas = (object)((dynamic)dstRange).Formula;
            dstValues = (object)((dynamic)dstRange).Value2;
            var srcF = ToGrid(srcFormulas);
            var srcV = ToGrid(srcValues);
            var dstF = ToGrid(dstFormulas);
            var dstV = ToGrid(dstValues);
            List<List<object?>>? expectedValues = null;
            List<List<object?>>? expectedFormulas = null;
            var expectNoFormulas = false;
            if (plan.Paste is "all")
            {
                ((dynamic)srcRange).Copy(dstRange);
                expectedValues = TransposedGrid(srcV, plan);
                if (plan.SkipBlanks || plan.Operation != 0)
                {
                    var adjusted = BuildAdjustedGrid(srcV, dstV, plan, pickFormula: false);
                    WritePasteGrid(dstRange, ToComGrid(adjusted), isFormula: false);
                    expectedValues = adjusted;
                }
            }
            else if (plan.Paste is "values")
            {
                expectedValues = BuildAdjustedGrid(srcV, dstV, plan, pickFormula: false);
                WritePasteGrid(dstRange, ToComGrid(expectedValues), isFormula: false);
                expectNoFormulas = true;
            }
            else if (plan.Paste is "formulas")
            {
                expectedFormulas = BuildAdjustedGrid(srcF, dstF, plan, pickFormula: true);
                WritePasteGrid(dstRange, ToComGrid(expectedFormulas), isFormula: true);
            }
            else
            {
                ((dynamic)srcRange).Copy(dstRange);
                WritePasteGrid(dstRange, ToComGrid(dstF), isFormula: true);
                WritePasteGrid(dstRange, ToComGrid(dstV), isFormula: false);
                expectedValues = dstV;
                expectedFormulas = dstF;
            }
            checkedCells += plan.Rows * plan.Cols;
            VerifyPasteSpecial(dstRange, srcRange, plan, expectedValues, expectedFormulas, expectNoFormulas,
                mismatches, $"{dest.SheetName}!{plan.WriteAddress}");
            execution.Affected.Add(new AffectedRef("range", $"{dest.SheetName}!{plan.WriteAddress}"));
        }
        finally
        {
            RotHelper.ReleaseComReference(dstValues);
            RotHelper.ReleaseComReference(dstFormulas);
            RotHelper.ReleaseComReference(srcValues);
            RotHelper.ReleaseComReference(srcFormulas);
            RotHelper.ReleaseComReference(dstRange);
            RotHelper.ReleaseComReference(srcRange);
        }
    }

    private static List<List<object?>> ToGrid(object? raw)
    {
        var grid = new List<List<object?>>();
        if (raw is object[,] arr)
        {
            var r1 = arr.GetLowerBound(0); var r2 = arr.GetUpperBound(0);
            var c1 = arr.GetLowerBound(1); var c2 = arr.GetUpperBound(1);
            for (var r = r1; r <= r2; r++)
            {
                var row = new List<object?>();
                for (var c = c1; c <= c2; c++) row.Add(arr[r, c]);
                grid.Add(row);
            }
        }
        else
        {
            grid.Add(new List<object?> { raw });
        }
        return grid;
    }

    private static bool IsBlankCell(object? formula, object? value)
    {
        if (formula is string text && !string.IsNullOrEmpty(text)) return false;
        return value is null || value is DBNull || (value is string s && s.Length == 0);
    }

    private static List<List<object?>> TransposedGrid(List<List<object?>> src, PastePlan plan)
    {
        var grid = new List<List<object?>>();
        for (var r = 0; r < plan.Rows; r++)
        {
            var row = new List<object?>();
            for (var c = 0; c < plan.Cols; c++)
            {
                var sr = plan.Transpose ? c : r;
                var sc = plan.Transpose ? r : c;
                row.Add(sr < src.Count && sc < src[sr].Count ? src[sr][sc] : null);
            }
            grid.Add(row);
        }
        return grid;
    }

    private static object[,] ToComGrid(List<List<object?>> grid)
    {
        var rows = grid.Count;
        var cols = grid.Count == 0 ? 0 : grid.Max(row => row.Count);
        var result = new object[Math.Max(rows, 1), Math.Max(cols, 1)];
        for (var r = 0; r < rows; r++)
            for (var c = 0; c < grid[r].Count; c++)
                result[r, c] = grid[r][c] ?? "";
        return result;
    }

    private static List<List<object?>> BuildAdjustedGrid(List<List<object?>> src, List<List<object?>> dst,
        PastePlan plan, bool pickFormula)
    {
        var grid = new List<List<object?>>();
        for (var r = 0; r < plan.Rows; r++)
        {
            var row = new List<object?>();
            for (var c = 0; c < plan.Cols; c++)
            {
                var sr = plan.Transpose ? c : r;
                var sc = plan.Transpose ? r : c;
                var srcFormula = sr < src.Count && sc < src[sr].Count ? src[sr][sc] : null;
                var srcValue = sr < src.Count && sc < src[sr].Count ? src[sr][sc] : null;
                var dstValue = r < dst.Count && c < dst[r].Count ? dst[r][c] : null;
                object? picked = pickFormula ? srcFormula : srcValue;
                if (plan.SkipBlanks && IsBlankCell(srcFormula, srcValue))
                {
                    row.Add(dstValue);
                    continue;
                }
                if (plan.Operation != 0 && !pickFormula)
                    picked = ApplyPasteOperation(dstValue, picked, plan.Operation);
                row.Add(picked);
            }
            grid.Add(row);
        }
        return grid;
    }

    private static object? ApplyPasteOperation(object? dst, object? src, int operation)
    {
        if (!TryComNumber(dst, out var left) || !TryComNumber(src, out var right)) return src;
        return operation switch
        {
            1 => left + right,
            2 => left - right,
            3 => left * right,
            4 => right == 0
                ? throw new InvalidOperationException("[EXCEL_PASTE_DIVIDE_BY_ZERO] paste operation divide by zero")
                : left / right,
            _ => src,
        };
    }

    private static bool TryComNumber(object? value, out double number)
    {
        number = 0;
        try
        {
            if (value is null || value is DBNull) return false;
            number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
            return double.IsFinite(number);
        }
        catch { return false; }
    }

    private static void WritePasteGrid(object dstRange, object[,] grid, bool isFormula)
    {
        if (isFormula) ((dynamic)dstRange).Formula = grid;
        else ((dynamic)dstRange).Value2 = grid;
    }

    private static void VerifyPasteSpecial(object dstRange, object srcRange, PastePlan plan,
        List<List<object?>>? expectedValues, List<List<object?>>? expectedFormulas, bool expectNoFormulas,
        List<string> mismatches, string label)
    {
        object? dstV = null;
        object? dstF = null;
        object? srcNF = null;
        object? dstNF = null;
        try
        {
            if (expectedValues is not null)
            {
                dstV = (object)((dynamic)dstRange).Value2;
                var actual = ToGrid(dstV);
                for (var r = 0; r < plan.Rows; r++)
                {
                    for (var c = 0; c < plan.Cols; c++)
                    {
                        var expected = r < expectedValues.Count && c < expectedValues[r].Count ? expectedValues[r][c] : null;
                        var got = r < actual.Count && c < actual[r].Count ? actual[r][c] : null;
                        if (!ComValuesEqual(NormalizePasteValue(expected), NormalizePasteValue(got)))
                        {
                            mismatches.Add($"{label}: value readback mismatch at offset ({r},{c})");
                            return;
                        }
                    }
                }
            }
            if (expectedFormulas is not null)
            {
                dstF = (object)((dynamic)dstRange).Formula;
                var actual = ToGrid(dstF);
                for (var r = 0; r < plan.Rows; r++)
                {
                    for (var c = 0; c < plan.Cols; c++)
                    {
                        var expected = r < expectedFormulas.Count && c < expectedFormulas[r].Count
                            ? expectedFormulas[r][c] as string : null;
                        var got = r < actual.Count && c < actual[r].Count ? actual[r][c] as string : null;
                        if (!string.Equals(expected ?? "", got ?? "", StringComparison.Ordinal))
                        {
                            mismatches.Add($"{label}: formula readback mismatch at offset ({r},{c})");
                            return;
                        }
                    }
                }
            }
            if (expectNoFormulas)
            {
                dstF ??= (object)((dynamic)dstRange).Formula;
                var actual = ToGrid(dstF);
                for (var r = 0; r < Math.Min(plan.Rows, actual.Count); r++)
                {
                    for (var c = 0; c < Math.Min(plan.Cols, actual[r].Count); c++)
                    {
                        if (actual[r][c] is string text && text.StartsWith('='))
                        {
                            mismatches.Add($"{label}: formulas remain after values paste");
                            return;
                        }
                    }
                }
            }
            if (plan.Paste is "all" or "formats")
            {
                srcNF = (object)((dynamic)srcRange).NumberFormat;
                dstNF = (object)((dynamic)dstRange).NumberFormat;
                var srcFormats = FlattenFormats(srcNF);
                var dstFormats = FlattenFormats(dstNF);
                for (var r = 0; r < plan.Rows; r++)
                {
                    for (var c = 0; c < plan.Cols; c++)
                    {
                        var sr = plan.Transpose ? c : r;
                        var sc = plan.Transpose ? r : c;
                        var expected = sr < srcFormats.Count && sc < srcFormats[sr].Count ? srcFormats[sr][sc] : null;
                        var got = r < dstFormats.Count && c < dstFormats[r].Count ? dstFormats[r][c] : null;
                        if (!string.Equals(expected ?? "", got ?? "", StringComparison.Ordinal))
                        {
                            mismatches.Add($"{label}: numberFormat readback mismatch at offset ({r},{c})");
                            return;
                        }
                    }
                }
            }
        }
        finally
        {
            RotHelper.ReleaseComReference(dstNF);
            RotHelper.ReleaseComReference(srcNF);
            RotHelper.ReleaseComReference(dstF);
            RotHelper.ReleaseComReference(dstV);
        }
    }

    private static List<List<string?>> FlattenFormats(object? raw)
    {
        var grid = new List<List<string?>>();
        if (raw is object[,] arr)
        {
            var r1 = arr.GetLowerBound(0); var r2 = arr.GetUpperBound(0);
            var c1 = arr.GetLowerBound(1); var c2 = arr.GetUpperBound(1);
            for (var r = r1; r <= r2; r++)
            {
                var row = new List<string?>();
                for (var c = c1; c <= c2; c++)
                    row.Add(Convert.ToString(arr[r, c], CultureInfo.InvariantCulture));
                grid.Add(row);
            }
        }
        else
        {
            grid.Add(new List<string?> { Convert.ToString(raw, CultureInfo.InvariantCulture) });
        }
        return grid;
    }

    private static object? NormalizePasteValue(object? value) =>
        value is DBNull ? null : value;

    // ------------------------------------------------------------------
    // goal_seek
    // ------------------------------------------------------------------

    private static void PreviewGoalSeek(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var changing = BindRange(workbook, op, "cell");
        using var goal = BindRange(workbook, op, "goalCell");
        EnsureSheetWritable(changing.Sheet);
        if (!string.Equals(changing.SheetName, goal.SheetName, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("[EXCEL_GOAL_SEEK_SHEET] cell and goalCell must be on the same worksheet");
        object? goalRange = null;
        object? formula = null;
        try
        {
            goalRange = (object)((dynamic)goal.Sheet).Range(goal.Address);
            formula = (object)((dynamic)goalRange).Formula;
            string? text = null;
            if (formula is object[,] arr)
                text = arr[arr.GetLowerBound(0), arr.GetLowerBound(1)] as string;
            else
                text = formula as string;
            if (string.IsNullOrWhiteSpace(text) || !text.StartsWith('='))
                throw new InvalidOperationException($"[EXCEL_GOAL_SEEK_FORMULA] goalCell {goal.SheetName}!{goal.Address} has no formula");
        }
        finally
        {
            RotHelper.ReleaseComReference(formula);
            RotHelper.ReleaseComReference(goalRange);
        }
        preview.Affected.Add(new AffectedRef("goalSeek", $"{goal.SheetName}!{goal.Address}"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{goal.SheetName}!{goal.Address}",
            Before = JsonValue.Create("current value"),
            After = op["goal"]?.DeepClone(),
        });
    }

    private static void ApplyGoalSeek(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var changing = BindRange(workbook, op, "cell");
        using var goal = BindRange(workbook, op, "goalCell");
        EnsureSheetWritable(changing.Sheet);
        if (!string.Equals(changing.SheetName, goal.SheetName, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("[EXCEL_GOAL_SEEK_SHEET] cell and goalCell must be on the same worksheet");
        var target = ExcelDataOperationsContract.TryGetFiniteNumber(op["goal"], out var goalValue)
            ? goalValue
            : throw new InvalidOperationException("[EXCEL_GOAL_SEEK_GOAL] goal must be a finite number");
        object? goalRange = null;
        object? changingRange = null;
        try
        {
            goalRange = (object)((dynamic)goal.Sheet).Range(goal.Address);
            changingRange = (object)((dynamic)changing.Sheet).Range(changing.Address);
            bool reached;
            try { reached = Convert.ToBoolean(((dynamic)goalRange).GoalSeek(target, changingRange), CultureInfo.InvariantCulture); }
            catch (Exception ex) { throw new InvalidOperationException($"[EXCEL_GOAL_SEEK_FAILED] GoalSeek did not converge: {ex.Message}"); }
            if (!reached)
                throw new InvalidOperationException("[EXCEL_GOAL_SEEK_FAILED] GoalSeek reported it may not have found a solution");
            checkedCells += 2;
            var tolerance = ExcelDataOperationsContract.TryGetFiniteNumber(op["tolerance"], out var customTolerance)
                ? customTolerance
                : 1e-4;
            object? actualRaw = null;
            try
            {
                actualRaw = (object)((dynamic)goalRange).Value2;
                if (TryComNumber(actualRaw, out var actual))
                {
                    var allowed = Math.Max(tolerance * Math.Max(1.0, Math.Abs(target)), 1e-9);
                    if (Math.Abs(actual - target) > allowed)
                        mismatches.Add($"{goal.SheetName}!{goal.Address}: value {actual} is outside tolerance of goal {target}");
                }
                else
                    mismatches.Add($"{goal.SheetName}!{goal.Address}: goal value is not numeric after GoalSeek");
            }
            finally { RotHelper.ReleaseComReference(actualRaw); }
            execution.Affected.Add(new AffectedRef("goalSeek", $"{goal.SheetName}!{goal.Address}={target} via {changing.SheetName}!{changing.Address}"));
        }
        finally
        {
            RotHelper.ReleaseComReference(changingRange);
            RotHelper.ReleaseComReference(goalRange);
        }
    }

    // ------------------------------------------------------------------
    // protect_workbook / unprotect_workbook
    // ------------------------------------------------------------------

    private static void PreviewProtectWorkbook(object workbook, JsonObject op, ApplyPreview preview)
    {
        if (WorkbookHasPassword(workbook))
            throw new InvalidOperationException("[EXCEL_WORKBOOK_PASSWORD] password workbooks are not supported by workbook protection ops");
        var before = ReadWorkbookProtection(workbook);
        var structure = op.ContainsKey("structure") ? Json.GetBool(op, "structure") : true;
        var windows = op.ContainsKey("windows") && Json.GetBool(op, "windows");
        preview.Affected.Add(new AffectedRef("workbook", "protection"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = "workbook:protection",
            Before = before.DeepClone(),
            After = new JsonObject { ["structure"] = structure, ["windows"] = windows },
        });
    }

    private static void ApplyProtectWorkbook(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        if (WorkbookHasPassword(workbook))
            throw new InvalidOperationException("[EXCEL_WORKBOOK_PASSWORD] password workbooks are not supported by workbook protection ops");
        var structure = op.ContainsKey("structure") ? Json.GetBool(op, "structure") : true;
        var windows = op.ContainsKey("windows") && Json.GetBool(op, "windows");
        if (TryComBool(workbook, "ProtectStructure") == true || TryComBool(workbook, "ProtectWindows") == true)
        {
            try { ((dynamic)workbook).Unprotect(); }
            catch (Exception ex) { throw new InvalidOperationException($"[EXCEL_WORKBOOK_UNPROTECT_FAILED] {ex.Message}"); }
        }
        try { ((dynamic)workbook).Protect(Type.Missing, structure, windows); }
        catch (Exception ex) { throw new InvalidOperationException($"[EXCEL_WORKBOOK_PROTECT_FAILED] {ex.Message}"); }
        checkedCells++;
        var actual = ReadWorkbookProtection(workbook);
        if (Json.GetBool(actual, "structure") != structure || Json.GetBool(actual, "windows") != windows)
            mismatches.Add("workbook protection readback mismatch");
        execution.Affected.Add(new AffectedRef("workbook", $"protected structure={structure} windows={windows}"));
    }

    private static void PreviewUnprotectWorkbook(object workbook, JsonObject op, ApplyPreview preview)
    {
        if (WorkbookHasPassword(workbook))
            throw new InvalidOperationException("[EXCEL_WORKBOOK_PASSWORD] password workbooks are not supported by workbook protection ops");
        preview.Affected.Add(new AffectedRef("workbook", "protection"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = "workbook:protection",
            Before = ReadWorkbookProtection(workbook),
            After = new JsonObject { ["structure"] = false, ["windows"] = false },
        });
    }

    private static void ApplyUnprotectWorkbook(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        if (WorkbookHasPassword(workbook))
            throw new InvalidOperationException("[EXCEL_WORKBOOK_PASSWORD] password workbooks are not supported by workbook protection ops");
        try { ((dynamic)workbook).Unprotect(); }
        catch (Exception ex) { throw new InvalidOperationException($"[EXCEL_WORKBOOK_UNPROTECT_FAILED] {ex.Message}"); }
        checkedCells++;
        var actual = ReadWorkbookProtection(workbook);
        if (Json.GetBool(actual, "structure") || Json.GetBool(actual, "windows"))
            mismatches.Add("workbook is still protected after Unprotect");
        execution.Affected.Add(new AffectedRef("workbook", "unprotected"));
    }

    // ------------------------------------------------------------------
    // set_split_panes
    // ------------------------------------------------------------------

    private static void PreviewSetSplitPanes(object workbook, JsonObject op, ApplyPreview preview)
    {
        using var sheet = BindSheet(workbook, op);
        var before = CaptureFreezePanes(sheet.Sheet);
        var planned = PlannedSplitPanes(op);
        preview.Affected.Add(new AffectedRef("window", $"{sheet.SheetName}:split"));
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{sheet.SheetName}:split",
            Before = before.DeepClone(),
            After = planned.DeepClone(),
        });
    }

    private static JsonObject PlannedSplitPanes(JsonObject op)
    {
        if (Json.GetBool(op, "remove"))
            return new JsonObject { ["frozen"] = false, ["splitRow"] = 0, ["splitColumn"] = 0 };
        ExcelWorkbookOpsContract.TryResolveSplit(op, out var rows, out var cols, out _);
        return new JsonObject
        {
            ["frozen"] = Json.GetBool(op, "freeze"),
            ["splitRow"] = rows,
            ["splitColumn"] = cols,
        };
    }

    private static void ApplySetSplitPanes(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedCells)
    {
        using var sheet = BindSheet(workbook, op);
        var window = CaptureWindowState(workbook);
        try
        {
            object? sheetWindow = null;
            try
            {
                if (!TryGetWorkbookWindow(workbook, out sheetWindow, out var windowError) || sheetWindow is null)
                    throw new InvalidOperationException(windowError ?? "workbook window is unavailable");
                var visible = Convert.ToInt32(((dynamic)sheet.Sheet).Visible, CultureInfo.InvariantCulture);
                if (visible != XlSheetVisible)
                    throw new InvalidOperationException($"set_split_panes cannot activate hidden sheet '{sheet.SheetName}'");
                ActivateWorksheet(workbook, sheet.SheetName, new RestoreMismatchCollector());
                var planned = PlannedSplitPanes(op);
                if (Json.GetBool(op, "remove"))
                {
                    ((dynamic)sheetWindow).FreezePanes = false;
                    ((dynamic)sheetWindow).Split = false;
                    ((dynamic)sheetWindow).SplitRow = 0;
                    ((dynamic)sheetWindow).SplitColumn = 0;
                }
                else if (Json.GetBool(planned, "frozen"))
                {
                    ((dynamic)sheetWindow).FreezePanes = false;
                    ((dynamic)sheetWindow).SplitRow = Json.GetInt(planned, "splitRow") ?? 0;
                    ((dynamic)sheetWindow).SplitColumn = Json.GetInt(planned, "splitColumn") ?? 0;
                    ((dynamic)sheetWindow).FreezePanes = true;
                }
                else
                {
                    ((dynamic)sheetWindow).FreezePanes = false;
                    ((dynamic)sheetWindow).SplitRow = Json.GetInt(planned, "splitRow") ?? 0;
                    ((dynamic)sheetWindow).SplitColumn = Json.GetInt(planned, "splitColumn") ?? 0;
                    ((dynamic)sheetWindow).Split =
                        (Json.GetInt(planned, "splitRow") ?? 0) > 0 || (Json.GetInt(planned, "splitColumn") ?? 0) > 0;
                }
            }
            finally { RotHelper.ReleaseComReference(sheetWindow); }
            var actual = CaptureFreezePanes(sheet.Sheet);
            var expected = PlannedSplitPanes(op);
            checkedCells++;
            if (Json.GetBool(actual, "readable") == false)
                mismatches.Add($"{sheet.SheetName}: split read failed: {Json.GetString(actual, "error")}");
            else if (!SplitEquals(actual, expected))
                mismatches.Add($"{sheet.SheetName}: split readback mismatch");
            execution.Affected.Add(new AffectedRef("window", $"{sheet.SheetName}:split"));
        }
        finally { RestoreWindowState(workbook, window, new RestoreMismatchCollector()); }
    }

    private static bool SplitEquals(JsonObject actual, JsonObject expected) =>
        Json.GetBool(actual, "frozen") == Json.GetBool(expected, "frozen") &&
        (Json.GetInt(actual, "splitRow") ?? -1) == (Json.GetInt(expected, "splitRow") ?? -2) &&
        (Json.GetInt(actual, "splitColumn") ?? -1) == (Json.GetInt(expected, "splitColumn") ?? -2);
}
