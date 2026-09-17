using System.Globalization;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class ExcelAdapter
{
    private static ExcelFormulaTraceContract.Resolution ResolveFormulaTraceStructuredSelection(
        dynamic workbook, string workbookPath, FormulaTraceReadContext readContext, ExcelFormulaTraceContract.Cell source,
        ExcelFormulaTraceContract.Reference reference, int remaining)
    {
        if (!ExcelStructuredReferenceContract.TryParse(reference.Token, out var parsed, out var parseReason))
            return FormulaTraceIssue(source, reference, parseReason ?? "structured_reference_malformed");

        object? table = null;
        object? columns = null;
        object? tableRange = null;
        object? header = null;
        object? body = null;
        object? totals = null;
        object? sourceSheet = null;
        object? sourceRange = null;
        try
        {
            if (parsed.TableName is null)
            {
                sourceSheet = GetSheet(workbook, source.Sheet);
                sourceRange = (object)((dynamic)sourceSheet).Range(source.Address);
                try { table = (object)((dynamic)sourceRange).ListObject; }
                catch { return FormulaTraceIssue(source, reference, "structured_local_table_not_found"); }
            }
            else table = FindFormulaTraceTable((object)workbook, parsed.TableName);
            if (table is null) return FormulaTraceIssue(source, reference, "structured_table_not_found");

            columns = (object)((dynamic)table).ListColumns;
            var count = Convert.ToInt32(((dynamic)columns).Count, CultureInfo.InvariantCulture);
            var names = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (var index = 1; index <= count; index++)
            {
                object? column = null;
                try
                {
                    column = (object)((dynamic)columns).Item(index);
                    var name = Convert.ToString(((dynamic)column).Name, CultureInfo.InvariantCulture);
                    if (string.IsNullOrWhiteSpace(name) || !names.TryAdd(name, index))
                        return FormulaTraceIssue(source, reference, "structured_column_ambiguous");
                }
                finally { RotHelper.ReleaseComReference(column); }
            }

            tableRange = (object)((dynamic)table).Range;
            var tableSheet = (object)((dynamic)tableRange).Worksheet;
            try
            {
                var tableBook = (object)((dynamic)tableSheet).Parent;
                try
                {
                    if (!string.Equals(TryString(() => ((dynamic)tableBook).FullName), workbookPath, StringComparison.OrdinalIgnoreCase))
                        return FormulaTraceIssue(source, reference, "external_workbook_reference");
                }
                finally { RotHelper.ReleaseComReference(tableBook); }

                header = (object?)((dynamic)table).HeaderRowRange;
                body = (object?)((dynamic)table).DataBodyRange;
                totals = (object?)((dynamic)table).TotalsRowRange;
                var firstColumn = Convert.ToInt32(((dynamic)tableRange).Column, CultureInfo.InvariantCulture);
                var plan = ExcelStructuredReferenceContract.Plan(parsed,
                    new ExcelStructuredReferenceContract.TableBounds(firstColumn, count,
                        RangeRow(header), RangeRow(body), RangeEndRow(body), RangeRow(totals)), names,
                    SourceRow(source));
                if (!plan.Resolved) return FormulaTraceIssue(source, reference, plan.Reason!);
                var sheetName = Convert.ToString(((dynamic)tableSheet).Name, CultureInfo.InvariantCulture);
                if (string.IsNullOrWhiteSpace(sheetName)) return FormulaTraceIssue(source, reference, "structured_range_unavailable");
                return ResolvePlannedStructuredBoxes(workbook, workbookPath, readContext, source, reference, sheetName, plan.Addresses, remaining);
            }
            finally { RotHelper.ReleaseComReference(tableSheet); }
        }
        catch { return FormulaTraceIssue(source, reference, "structured_table_or_column_not_found"); }
        finally
        {
            RotHelper.ReleaseComReference(totals);
            RotHelper.ReleaseComReference(body);
            RotHelper.ReleaseComReference(header);
            RotHelper.ReleaseComReference(tableRange);
            RotHelper.ReleaseComReference(columns);
            RotHelper.ReleaseComReference(table);
            RotHelper.ReleaseComReference(sourceRange);
            RotHelper.ReleaseComReference(sourceSheet);
        }
    }

    private static ExcelFormulaTraceContract.Resolution ResolvePlannedStructuredBoxes(dynamic workbook, string workbookPath, FormulaTraceReadContext readContext,
        ExcelFormulaTraceContract.Cell source, ExcelFormulaTraceContract.Reference reference, string sheetName,
        IReadOnlyList<string> addresses, int remaining)
    {
        var targets = new List<ExcelFormulaTraceContract.Cell>();
        var issues = new List<ExcelFormulaTraceContract.Issue>();
        var known = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var address in addresses)
        {
            if (targets.Count >= remaining)
                return new ExcelFormulaTraceContract.Resolution(targets, issues, true, "max_cells");
            var resolution = ResolveFormulaTraceRange(workbook, workbookPath, readContext, source, reference, sheetName, address, remaining - targets.Count);
            issues.AddRange(resolution.Unresolved);
            foreach (var target in resolution.Targets)
                if (known.Add(target.Sheet + "!" + target.Address)) targets.Add(target);
            if (resolution.Truncated) return new ExcelFormulaTraceContract.Resolution(targets, issues, true, resolution.TruncationReason);
        }
        return new ExcelFormulaTraceContract.Resolution(targets, issues);
    }

    private static int? RangeRow(object? range) => range is null ? null : Convert.ToInt32(((dynamic)range).Row, CultureInfo.InvariantCulture);
    private static int? RangeEndRow(object? range) => range is null ? null : Convert.ToInt32(((dynamic)range).Row, CultureInfo.InvariantCulture) + Convert.ToInt32(((dynamic)range).Rows.Count, CultureInfo.InvariantCulture) - 1;
    private static int SourceRow(ExcelFormulaTraceContract.Cell source) => ExcelA1Box.TryParseCell(source.Address, out var row, out _) ? row : 0;
}
