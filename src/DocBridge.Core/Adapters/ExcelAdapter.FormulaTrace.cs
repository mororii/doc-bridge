using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class ExcelAdapter
{
    private const int FormulaTraceDefaultDepth = 4;
    private const int FormulaTraceDefaultCells = 500;
    private const int FormulaTraceMaxDepth = 20;
    private const int FormulaTraceMaxCells = 5_000;

    private static JsonObject InspectFormulaTrace(dynamic workbook, JsonObject args)
    {
        var requestedSheet = Json.GetString(args, "sheet");
        var requestedRange = Json.GetString(args, "range");
        var requestedWorkbook = Json.GetString(args, "workbook");
        if (string.IsNullOrWhiteSpace(requestedWorkbook))
            throw new ArgumentException("excel_inspect scope=formula_trace requires explicit 'workbook'");
        if (string.IsNullOrWhiteSpace(requestedSheet))
            throw new ArgumentException("excel_inspect scope=formula_trace requires explicit 'sheet'");
        if (string.IsNullOrWhiteSpace(requestedRange))
            throw new ArgumentException("excel_inspect scope=formula_trace requires 'range'");

        var maxDepth = Math.Clamp(Json.GetInt(args, "maxDepth") ?? FormulaTraceDefaultDepth, 0, FormulaTraceMaxDepth);
        var maxCells = Math.Clamp(Json.GetInt(args, "maxCells") ?? FormulaTraceDefaultCells, 1, FormulaTraceMaxCells);
        var resolved = ResolveRangeTarget((object)workbook, requestedSheet, requestedRange, requireExplicitSheet: true);
        object? rootSheet = resolved.Sheet;
        object? rootRange = null;
        try
        {
            dynamic sheet = rootSheet;
            rootRange = (object)sheet.Range(resolved.Address);
            dynamic range = rootRange;
            object? areas = null;
            try
            {
                areas = (object)range.Areas;
                if (Convert.ToInt32(((dynamic)areas).Count, CultureInfo.InvariantCulture) != 1)
                    throw new ArgumentException("formula_trace requires one contiguous A1 range");
            }
            finally { RotHelper.ReleaseComReference(areas); }

            string rootAddress = Convert.ToString(range.Address(false, false), CultureInfo.InvariantCulture) ?? resolved.Address;
            ExcelA1Box rootBox;
            if (!ExcelA1Box.TryParse(rootAddress, out rootBox))
                throw new ArgumentException("formula_trace range must be a rectangular A1 range");
            var rootSheetName = Convert.ToString(sheet.Name, CultureInfo.InvariantCulture) ?? requestedSheet;
            var workbookPath = Convert.ToString(workbook.FullName, CultureInfo.InvariantCulture) ?? "";
            var rootTruncated = rootBox.CellCount > maxCells;
            var rootLimit = Math.Min(rootBox.CellCount, maxCells);
            FormulaTraceReadContext readContext = new((object)workbook, workbookPath);
            IReadOnlyList<ExcelFormulaTraceContract.Cell> roots = readContext.Read(rootSheetName, rootBox, rootLimit);

            var traversal = ExcelFormulaTraceContract.Trace(
                roots,
                maxDepth,
                maxCells,
                (source, reference, remaining) => ResolveFormulaTraceReference(workbook, workbookPath, readContext, source, reference, remaining));
            JsonObject result = FormulaTraceResult(workbookPath, rootSheetName, rootAddress, maxDepth, maxCells, rootBox.CellCount, rootTruncated, traversal);
            return args["compact"]?.GetValue<bool>() == true ? ExcelFormulaTraceOutputContract.Compact(result) : result;
        }
        finally
        {
            RotHelper.ReleaseComReference(rootRange);
            RotHelper.ReleaseComReference(rootSheet);
        }
    }

    private static ExcelFormulaTraceContract.Cell ReadFormulaTraceCell(
        dynamic workbook, string workbookPath, string sheetName, string address)
    {
        object? sheet = null;
        object? range = null;
        try
        {
            sheet = GetSheet(workbook, sheetName);
            dynamic targetSheet = sheet;
            range = (object)targetSheet.Range(address);
            dynamic cell = range;
            var actualAddress = Convert.ToString(cell.Address(false, false), CultureInfo.InvariantCulture) ?? address;
            string? formula = null;
            string formulaRead;
            var formulaReadSucceeded = false;
            try
            {
                var raw = cell.Formula2;
                formula = raw is string formulaText && formulaText.StartsWith("=", StringComparison.Ordinal) ? formulaText : null;
                formulaRead = "formula2";
                formulaReadSucceeded = true;
            }
            catch (Exception formula2Error)
            {
                try
                {
                    var raw = cell.Formula;
                    formula = raw is string formulaText && formulaText.StartsWith("=", StringComparison.Ordinal) ? formulaText : null;
                    formulaRead = "formula (Formula2 fallback: " + formula2Error.GetType().Name + ")";
                    formulaReadSucceeded = true;
                }
                catch (Exception formulaError)
                {
                    formulaRead = "unavailable (Formula2: " + formula2Error.GetType().Name + "; Formula: " + formulaError.GetType().Name + ")";
                }
            }

            object? value = null;
            var valueReadSucceeded = false;
            try
            {
                value = cell.Value2;
                valueReadSucceeded = true;
            }
            catch { }
            var display = TryString(() => cell.Text);
            var isExcelError = TryClassifyFormulaTraceExcelError(cell);
            var evidence = new JsonObject
            {
                ["workbook"] = workbookPath,
                ["sheet"] = sheetName,
                ["address"] = actualAddress,
                ["value"] = FormulaTraceValue(value),
                ["formula"] = formula,
                ["formulaRead"] = formulaRead,
                ["errorClassification"] = isExcelError switch
                {
                    true => "excel_error",
                    false => "not_excel_error",
                    _ => "unknown",
                },
                ["readEvidenceComplete"] = formulaReadSucceeded && valueReadSucceeded && isExcelError is not null,
            };
            if (!string.IsNullOrWhiteSpace(display) && (formula is not null || isExcelError is true))
                evidence["displayText"] = display;
            if (isExcelError is true && !string.IsNullOrWhiteSpace(display))
                evidence["errorDisplay"] = display;
            return new ExcelFormulaTraceContract.Cell(sheetName, actualAddress, formula, evidence);
        }
        finally
        {
            RotHelper.ReleaseComReference(range);
            RotHelper.ReleaseComReference(sheet);
        }
    }

    private static JsonNode? FormulaTraceValue(object? value) => value switch
    {
        null or DBNull => null,
        ErrorWrapper error => JsonValue.Create(error.ErrorCode),
        string text => JsonValue.Create(text),
        bool flag => JsonValue.Create(flag),
        byte number => JsonValue.Create(number),
        short number => JsonValue.Create(number),
        int number => JsonValue.Create(number),
        long number => JsonValue.Create(number),
        float number => JsonValue.Create((double)number),
        double number => JsonValue.Create(number),
        decimal number => JsonValue.Create(number),
        DateTime date => JsonValue.Create(date.ToString("o", CultureInfo.InvariantCulture)),
        _ => JsonValue.Create(Convert.ToString(value, CultureInfo.InvariantCulture)),
    };

    private static bool? TryClassifyFormulaTraceExcelError(dynamic range)
    {
        object? application = null;
        object? worksheetFunction = null;
        try
        {
            application = (object)range.Application;
            worksheetFunction = (object)((dynamic)application).WorksheetFunction;
            return Convert.ToBoolean(((dynamic)worksheetFunction).IsError(range), CultureInfo.InvariantCulture);
        }
        catch
        {
            return null;
        }
        finally
        {
            RotHelper.ReleaseComReference(worksheetFunction);
            RotHelper.ReleaseComReference(application);
        }
    }

    private static ExcelFormulaTraceContract.Resolution ResolveFormulaTraceReference(
        dynamic workbook,
        string workbookPath,
        FormulaTraceReadContext readContext,
        ExcelFormulaTraceContract.Cell source,
        ExcelFormulaTraceContract.Reference reference,
        int remaining)
    {
        if (reference.Kind is "cell" or "axis")
            return ResolveFormulaTraceRange(workbook, workbookPath, readContext, source, reference,
                reference.Sheet ?? source.Sheet, reference.Address, remaining);
        if (reference.Kind == "spill")
            return ResolveFormulaTraceSpill(workbook, workbookPath, readContext, source, reference, remaining);
        if (reference.Kind == "structured")
            return ResolveFormulaTraceStructuredSelection(workbook, workbookPath, readContext, source, reference, remaining);
        if (reference.Kind == "name")
            return ResolveFormulaTraceName(workbook, workbookPath, readContext, source, reference, remaining);
        return FormulaTraceIssue(source, reference, reference.Reason ?? "unsupported_reference");
    }

    private static ExcelFormulaTraceContract.Resolution ResolveFormulaTraceSpill(dynamic workbook, string workbookPath, FormulaTraceReadContext readContext,
        ExcelFormulaTraceContract.Cell source, ExcelFormulaTraceContract.Reference reference, int remaining)
    {
        object? sheet = null; object? anchor = null; object? spill = null; object? spillSheet = null; object? spillBook = null; object? areas = null;
        try
        {
            sheet = GetSheet(workbook, reference.Sheet ?? source.Sheet);
            anchor = (object)((dynamic)sheet).Range(reference.Address);
            spill = (object)((dynamic)anchor).SpillingToRange;
            spillSheet = (object)((dynamic)spill).Worksheet;
            spillBook = (object)((dynamic)spillSheet).Parent;
            if (!string.Equals(TryString(() => ((dynamic)spillBook).FullName), workbookPath, StringComparison.OrdinalIgnoreCase))
                return FormulaTraceIssue(source, reference, "external_workbook_reference");
            areas = (object)((dynamic)spill).Areas;
            if (Convert.ToInt32(((dynamic)areas).Count, CultureInfo.InvariantCulture) != 1)
                return FormulaTraceIssue(source, reference, "spill_multi_area_range");
            var address = Convert.ToString(((dynamic)spill).Address(false, false), CultureInfo.InvariantCulture);
            var name = Convert.ToString(((dynamic)spillSheet).Name, CultureInfo.InvariantCulture);
            return string.IsNullOrWhiteSpace(address) || string.IsNullOrWhiteSpace(name)
                ? FormulaTraceIssue(source, reference, "spill_range_unavailable")
                : ResolveFormulaTraceRange(workbook, workbookPath, readContext, source, reference, name, address, remaining);
        }
        catch { return FormulaTraceIssue(source, reference, "spill_range_unavailable"); }
        finally { RotHelper.ReleaseComReference(areas); RotHelper.ReleaseComReference(spillBook); RotHelper.ReleaseComReference(spillSheet); RotHelper.ReleaseComReference(spill); RotHelper.ReleaseComReference(anchor); RotHelper.ReleaseComReference(sheet); }
    }

    private static ExcelFormulaTraceContract.Resolution ResolveFormulaTraceStructured(dynamic workbook, string workbookPath, FormulaTraceReadContext readContext,
        ExcelFormulaTraceContract.Cell source, ExcelFormulaTraceContract.Reference reference, int remaining)
    {
        var open = reference.Token.IndexOf('[');
        if (open <= 0 || !reference.Token.EndsWith(']')) return FormulaTraceIssue(source, reference, "structured_reference_ambiguous");
        var tableName = reference.Token[..open];
        var selector = reference.Token[open..];
        if (selector.Contains("@", StringComparison.Ordinal) || selector.Contains("#This Row", StringComparison.OrdinalIgnoreCase))
            return FormulaTraceIssue(source, reference, "structured_current_row_requires_native_table_context");
        var column = selector.Trim('[', ']');
        if (column.Contains("#", StringComparison.Ordinal) || column.Contains(",", StringComparison.Ordinal) || column.Contains("][", StringComparison.Ordinal))
            return FormulaTraceIssue(source, reference, "structured_selector_unsupported");
        object? table = null; object? columns = null; object? listColumn = null; object? range = null; object? rangeSheet = null; object? rangeBook = null;
        try
        {
            table = FindFormulaTraceTable((object)workbook, tableName);
            if (table is null) return FormulaTraceIssue(source, reference, "structured_table_not_found");
            columns = (object)((dynamic)table).ListColumns;
            listColumn = (object)((dynamic)columns).Item(column);
            range = (object?)((dynamic)listColumn).DataBodyRange;
            if (range is null)
                return new ExcelFormulaTraceContract.Resolution(Array.Empty<ExcelFormulaTraceContract.Cell>(), Array.Empty<ExcelFormulaTraceContract.Issue>());
            rangeSheet = (object)((dynamic)range).Worksheet; rangeBook = (object)((dynamic)rangeSheet).Parent;
            if (!string.Equals(TryString(() => ((dynamic)rangeBook).FullName), workbookPath, StringComparison.OrdinalIgnoreCase)) return FormulaTraceIssue(source, reference, "external_workbook_reference");
            var address = Convert.ToString(((dynamic)range).Address(false, false), CultureInfo.InvariantCulture);
            var name = Convert.ToString(((dynamic)rangeSheet).Name, CultureInfo.InvariantCulture);
            return string.IsNullOrWhiteSpace(address) || string.IsNullOrWhiteSpace(name) ? FormulaTraceIssue(source, reference, "structured_range_unavailable") : ResolveFormulaTraceRange(workbook, workbookPath, readContext, source, reference, name, address, remaining);
        }
        catch { return FormulaTraceIssue(source, reference, "structured_table_or_column_not_found"); }
        finally { RotHelper.ReleaseComReference(rangeBook); RotHelper.ReleaseComReference(rangeSheet); RotHelper.ReleaseComReference(range); RotHelper.ReleaseComReference(listColumn); RotHelper.ReleaseComReference(columns); RotHelper.ReleaseComReference(table); }
    }

    private static object? FindFormulaTraceTable(object workbook, string tableName)
    {
        object? sheets = null;
        try
        {
            sheets = (object)((dynamic)workbook).Worksheets;
            int count = Convert.ToInt32(((dynamic)sheets).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                object? sheet = null;
                object? tables = null;
                try
                {
                    sheet = (object)((dynamic)sheets).Item(index);
                    tables = (object)((dynamic)sheet).ListObjects;
                    try { return (object)((dynamic)tables).Item(tableName); }
                    catch (COMException) { }
                }
                finally
                {
                    RotHelper.ReleaseComReference(tables);
                    RotHelper.ReleaseComReference(sheet);
                }
            }
            return null;
        }
        finally { RotHelper.ReleaseComReference(sheets); }
    }

    private static ExcelFormulaTraceContract.Resolution ResolveFormulaTraceRange(
        dynamic workbook,
        string workbookPath,
        FormulaTraceReadContext readContext,
        ExcelFormulaTraceContract.Cell source,
        ExcelFormulaTraceContract.Reference reference,
        string sheetName,
        string? address,
        int remaining)
    {
        if (string.IsNullOrWhiteSpace(address) || !ExcelA1Box.TryParse(address, out var box))
            return FormulaTraceIssue(source, reference, "invalid_a1_reference");
        object? sheet = null;
        try
        {
            sheet = GetSheet(workbook, sheetName);
        }
        catch
        {
            return FormulaTraceIssue(source, reference, "missing_sheet");
        }
        finally { RotHelper.ReleaseComReference(sheet); }

        var issues = new List<ExcelFormulaTraceContract.Issue>();
        var toRead = Math.Min(box.CellCount, remaining);
        List<ExcelFormulaTraceContract.Cell> targets;
        try { targets = readContext.Read(sheetName, box, toRead).ToList(); }
        catch { issues.Add(new ExcelFormulaTraceContract.Issue(source.Sheet, source.Address, reference, "missing_or_invalid_reference")); return new ExcelFormulaTraceContract.Resolution(Array.Empty<ExcelFormulaTraceContract.Cell>(), issues); }
        return new ExcelFormulaTraceContract.Resolution(
            targets,
            issues,
            box.CellCount > remaining,
            box.CellCount > remaining ? "max_cells" : null);
    }

    private static ExcelFormulaTraceContract.Resolution ResolveFormulaTraceName(
        dynamic workbook,
        string workbookPath,
        FormulaTraceReadContext readContext,
        ExcelFormulaTraceContract.Cell source,
        ExcelFormulaTraceContract.Reference reference,
        int remaining)
    {
        object? nativeName = null;
        object? referredRange = null;
        object? targetSheet = null;
        object? targetWorkbook = null;
        object? areas = null;
        try
        {
            nativeName = FindFormulaTraceName(workbook, source.Sheet, reference.Token);
            if (nativeName is null)
                return FormulaTraceIssue(source, reference, "defined_name_not_found");
            try
            {
                referredRange = (object)((dynamic)nativeName).RefersToRange;
            }
            catch
            {
                return FormulaTraceIssue(source, reference, "defined_name_expression_or_constant");
            }
            targetSheet = (object)((dynamic)referredRange).Worksheet;
            targetWorkbook = (object)((dynamic)targetSheet).Parent;
            var targetWorkbookPath = TryString(() => ((dynamic)targetWorkbook).FullName);
            if (!string.Equals(targetWorkbookPath, workbookPath, StringComparison.OrdinalIgnoreCase))
                return FormulaTraceIssue(source, reference, "external_workbook_reference");
            areas = (object)((dynamic)referredRange).Areas;
            if (Convert.ToInt32(((dynamic)areas).Count, CultureInfo.InvariantCulture) != 1)
                return FormulaTraceIssue(source, reference, "defined_name_multi_area_range");
            var targetSheetName = Convert.ToString(((dynamic)targetSheet).Name, CultureInfo.InvariantCulture) ?? "";
            var targetAddress = Convert.ToString(((dynamic)referredRange).Address(false, false), CultureInfo.InvariantCulture);
            if (string.IsNullOrWhiteSpace(targetSheetName) || string.IsNullOrWhiteSpace(targetAddress))
                return FormulaTraceIssue(source, reference, "defined_name_range_unavailable");
            return ResolveFormulaTraceRange(workbook, workbookPath, readContext, source, reference, targetSheetName, targetAddress, remaining);
        }
        finally
        {
            RotHelper.ReleaseComReference(areas);
            RotHelper.ReleaseComReference(targetWorkbook);
            RotHelper.ReleaseComReference(targetSheet);
            RotHelper.ReleaseComReference(referredRange);
            RotHelper.ReleaseComReference(nativeName);
        }
    }

    private static object? FindFormulaTraceName(dynamic workbook, string sourceSheet, string token)
    {
        object? sheet = null;
        object? names = null;
        try
        {
            sheet = GetSheet(workbook, sourceSheet);
            names = (object)((dynamic)sheet).Names;
            try { return (object)((dynamic)names).Item(token); } catch { }
        }
        finally
        {
            RotHelper.ReleaseComReference(names);
            RotHelper.ReleaseComReference(sheet);
        }

        object? workbookNames = null;
        try
        {
            workbookNames = (object)workbook.Names;
            try { return (object)((dynamic)workbookNames).Item(token); } catch { }
            var localName = ExcelFormulaReference.QuoteSheetName(sourceSheet) + "!" + token;
            try { return (object)((dynamic)workbookNames).Item(localName); } catch { }
        }
        finally { RotHelper.ReleaseComReference(workbookNames); }
        return null;
    }

    private static ExcelFormulaTraceContract.Resolution FormulaTraceIssue(
        ExcelFormulaTraceContract.Cell source,
        ExcelFormulaTraceContract.Reference reference,
        string reason) => new(
        Array.Empty<ExcelFormulaTraceContract.Cell>(),
        new[] { new ExcelFormulaTraceContract.Issue(source.Sheet, source.Address, reference, reason) });

    private static JsonObject FormulaTraceResult(
        string workbookPath,
        string rootSheet,
        string rootRange,
        int maxDepth,
        int maxCells,
        long requestedRootCells,
        bool rootTruncated,
        ExcelFormulaTraceContract.Traversal traversal)
    {
        var nodes = new JsonArray(traversal.Nodes.Select(node => (JsonNode)node.Evidence.DeepClone()).ToArray());
        var edges = new JsonArray(traversal.Edges.Select(edge => (JsonNode)new JsonObject
        {
            ["from"] = new JsonObject { ["sheet"] = edge.SourceSheet, ["address"] = edge.SourceAddress },
            ["to"] = new JsonObject { ["sheet"] = edge.TargetSheet, ["address"] = edge.TargetAddress },
            ["via"] = edge.Via,
        }).ToArray());
        var unresolved = new JsonArray(traversal.Unresolved.Select(issue => (JsonNode)new JsonObject
        {
            ["from"] = new JsonObject { ["sheet"] = issue.SourceSheet, ["address"] = issue.SourceAddress },
            ["reference"] = issue.Reference.Token,
            ["kind"] = issue.Reference.Kind,
            ["reason"] = issue.Reason,
        }).ToArray());
        var frontier = new JsonArray(traversal.Frontier.Select(item => (JsonNode)JsonValue.Create(item)!).ToArray());
        if (rootTruncated)
        {
            unresolved.Add(new JsonObject
            {
                ["from"] = new JsonObject { ["sheet"] = rootSheet, ["address"] = rootRange },
                ["reference"] = rootRange,
                ["kind"] = "range",
                ["reason"] = "root_range_truncated_by_max_cells",
            });
            frontier.Add(rootSheet + "!" + rootRange);
        }
        var reasons = new JsonArray(traversal.TruncationReasons.Select(reason => (JsonNode)JsonValue.Create(reason)!).ToArray());
        if (rootTruncated && !traversal.TruncationReasons.Contains("max_cells", StringComparer.OrdinalIgnoreCase))
            reasons.Add("max_cells");
        var cycles = new JsonArray(traversal.Cycles.Select(cycle => (JsonNode)new JsonObject
        {
            ["from"] = new JsonObject { ["sheet"] = cycle.SourceSheet, ["address"] = cycle.SourceAddress },
            ["to"] = new JsonObject { ["sheet"] = cycle.TargetSheet, ["address"] = cycle.TargetAddress },
        }).ToArray());
        var readEvidenceComplete = traversal.Nodes.All(node =>
            node.Evidence["readEvidenceComplete"] is not JsonValue state || state.GetValue<bool>());
        var complete = !rootTruncated && traversal.Complete && readEvidenceComplete;
        return new JsonObject
        {
            ["ok"] = true,
            ["app"] = "excel",
            ["scope"] = "formula_trace",
            ["workbook"] = workbookPath,
            ["root"] = new JsonObject { ["sheet"] = rootSheet, ["range"] = rootRange },
            ["nodes"] = nodes,
            ["edges"] = edges,
            ["unresolved"] = unresolved,
            ["cycles"] = cycles,
            ["coverage"] = new JsonObject
            {
                ["maxDepth"] = maxDepth,
                ["maxCells"] = maxCells,
                ["requestedRootCells"] = requestedRootCells,
                ["nodesReturned"] = nodes.Count,
                ["edgesReturned"] = edges.Count,
                ["complete"] = complete,
                ["readEvidenceComplete"] = readEvidenceComplete,
                ["truncated"] = reasons.Count > 0,
                ["truncationReasons"] = reasons,
                ["frontier"] = frontier,
            },
        };
    }
}
