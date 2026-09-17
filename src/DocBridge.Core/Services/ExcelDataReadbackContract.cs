using System.Globalization;
using System.Linq;
using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// COM-free comparison of requested data-object fields against a captured
/// native snapshot. Adapter readback calls these after live COM capture.
/// Sort uses all requested keys plus row-association permutation.
/// T2C compares every destination row, not only the first cell.
/// </summary>
public static class ExcelDataReadbackContract
{
    public static void CompareRequestedTotals(JsonObject actual, JsonObject op, string label,
        ICollection<string> mismatches)
    {
        ArgumentNullException.ThrowIfNull(actual);
        ArgumentNullException.ThrowIfNull(op);
        ArgumentNullException.ThrowIfNull(mismatches);
        var requested = Json.GetArr(op, "columns") ?? Json.GetArr(op, "totals");
        if (requested is null || requested.Count == 0) return;
        var actualTotals = Json.GetArr(actual, "totals") ?? new JsonArray();
        foreach (var req in requested.OfType<JsonObject>())
        {
            var column = ColumnKey(req["column"]);
            var function = Json.GetString(req, "function");
            var match = actualTotals.OfType<JsonObject>().FirstOrDefault(item =>
                ColumnMatches(item, column) &&
                string.Equals(Json.GetString(item, "function"), function, StringComparison.OrdinalIgnoreCase));
            if (match is null)
            {
                mismatches.Add($"{label}: totals column '{column}' function '{function}' not present in readback");
                continue;
            }
            if (IsAggregateFunction(function) && !HasFiniteAggregate(match))
                mismatches.Add($"{label}: totals column '{column}' function '{function}' has no finite aggregate value");
        }
    }

    public static void CompareRequestedPivot(JsonObject actual, JsonObject op, string label,
        ICollection<string> mismatches, bool requireAggregates)
    {
        ArgumentNullException.ThrowIfNull(actual);
        ArgumentNullException.ThrowIfNull(op);
        ArgumentNullException.ThrowIfNull(mismatches);

        var requestedName = Json.GetString(op, "name");
        var actualName = Json.GetString(actual, "name");
        if (!string.IsNullOrWhiteSpace(requestedName) &&
            !string.Equals(actualName, requestedName, StringComparison.OrdinalIgnoreCase))
            mismatches.Add($"{label}: pivot name '{actualName}' != '{requestedName}'");

        if (op.ContainsKey("sourceRange") &&
            !ExcelDataRangeBinding.SourceReadbackMatches(
                Json.GetString(op, "sourceRange"),
                Json.GetString(actual, "source"),
                defaultSheet: null,
                ownerWorkbook: ReadOwnerWorkbook(actual)))
        {
            mismatches.Add(
                $"{label}: source '{Json.GetString(actual, "source")}' does not include '{Json.GetString(op, "sourceRange")}'");
        }

        CompareFieldList(actual, "rowFields", op, "rows", label, mismatches);
        CompareFieldList(actual, "columnFields", op, "columns", label, mismatches);
        CompareFieldList(actual, "pageFields", op, "filters", label, mismatches);

        var requestedValues = Json.GetArr(op, "values");
        if (requestedValues is null || requestedValues.Count == 0)
        {
            if (requireAggregates && !HasAnyAggregate(actual))
                mismatches.Add($"{label}: refresh/readback has no numeric pivot aggregates (name-only is not enough)");
            return;
        }

        var actualValues = Json.GetArr(actual, "dataFields") ?? new JsonArray();
        var aggregates = Json.GetArr(actual, "aggregates") ?? new JsonArray();
        foreach (var node in requestedValues.OfType<JsonObject>())
        {
            var field = Json.GetString(node, "field");
            var function = Json.GetString(node, "function") ?? "sum";
            var match = actualValues.OfType<JsonObject>().Any(actualField =>
                DataFieldMentions(actualField, field) &&
                string.Equals(Json.GetString(actualField, "function") ?? "sum", function,
                    StringComparison.OrdinalIgnoreCase));
            if (!match)
                mismatches.Add($"{label}: missing data field '{field}' function '{function}'");

            var actualField = actualValues.OfType<JsonObject>().FirstOrDefault(item =>
                DataFieldMentions(item, field) &&
                string.Equals(Json.GetString(item, "function") ?? "sum", function,
                    StringComparison.OrdinalIgnoreCase));
            var caption = Json.GetString(node, "caption");
            if (!string.IsNullOrWhiteSpace(caption) && actualField is not null &&
                !string.Equals(Json.GetString(actualField, "caption") ?? Json.GetString(actualField, "name"),
                    caption, StringComparison.OrdinalIgnoreCase))
                mismatches.Add($"{label}: data field '{field}' caption '{Json.GetString(actualField, "caption")}' != '{caption}'");
            var numberFormat = Json.GetString(node, "numberFormat");
            if (!string.IsNullOrWhiteSpace(numberFormat) && actualField is not null &&
                !string.Equals(Json.GetString(actualField, "numberFormat"), numberFormat, StringComparison.OrdinalIgnoreCase))
                mismatches.Add($"{label}: data field '{field}' numberFormat '{Json.GetString(actualField, "numberFormat")}' != '{numberFormat}'");

            var aggregate = aggregates.OfType<JsonObject>().FirstOrDefault(item =>
                (DataFieldMentions(item, field) || FieldMentions(Json.GetString(item, "field"), field)) &&
                string.Equals(Json.GetString(item, "function") ?? "sum", function,
                    StringComparison.OrdinalIgnoreCase));
            if (requireAggregates || requestedValues.Count > 0)
            {
                if (aggregate is null || !HasFiniteAggregate(aggregate))
                    mismatches.Add($"{label}: data field '{field}' function '{function}' has no numeric aggregate readback");
            }
        }
    }

    public static void CompareRequestedChartSeries(JsonObject actual, JsonObject op, string label,
        ICollection<string> mismatches)
    {
        ArgumentNullException.ThrowIfNull(actual);
        ArgumentNullException.ThrowIfNull(op);
        ArgumentNullException.ThrowIfNull(mismatches);
        var requested = ExcelChartSeriesContract.ReadSeries(Json.GetArr(op, "series"));
        var chartLevelType = Json.GetString(op, "chartType");
        if (requested.Count == 0 && string.IsNullOrWhiteSpace(chartLevelType)) return;
        var actualSeries = Json.GetArr(actual, "series") ?? new JsonArray();
        var byIndex = new Dictionary<int, ExcelChartSeriesContract.SeriesSpec>();
        foreach (var spec in requested)
            byIndex[spec.Index] = spec;
        var maxRequested = requested.Count == 0 ? 0 : requested.Max(item => item.Index);
        if (maxRequested > actualSeries.Count)
            mismatches.Add($"{label}: series count {actualSeries.Count} < requested index {maxRequested}");
        var limit = Math.Max(actualSeries.Count, maxRequested);
        for (var index = 1; index <= limit; index++)
        {
            var got = index <= actualSeries.Count ? actualSeries[index - 1] as JsonObject : null;
            if (got is null)
            {
                mismatches.Add($"{label}: series[{index}] is missing");
                continue;
            }
            var hasSpec = byIndex.TryGetValue(index, out var spec);
            var formula = Json.GetString(got, "formula") ?? "";
            var defaultSheet = Json.GetString(actual, "sheet");
            var ownerWorkbook = ReadOwnerWorkbook(actual);
            if (hasSpec && spec.Values is not null &&
                !ExcelChartSeriesContract.FormulaMentionsValues(formula, spec.Values, defaultSheet, ownerWorkbook) &&
                !ExcelDataRangeBinding.SourceReadbackMatches(
                    spec.Values, Json.GetString(got, "values"), defaultSheet, ownerWorkbook))
                mismatches.Add($"{label}: series[{index}].values '{spec.Values}' not in formula '{formula}'");
            if (hasSpec && spec.Categories is not null &&
                !ExcelChartSeriesContract.FormulaMentionsCategories(formula, spec.Categories, defaultSheet, ownerWorkbook) &&
                !ExcelDataRangeBinding.SourceReadbackMatches(
                    spec.Categories, Json.GetString(got, "categories"), defaultSheet, ownerWorkbook))
                mismatches.Add($"{label}: series[{index}].categories '{spec.Categories}' not in formula '{formula}'");
            if (hasSpec && spec.Range is not null &&
                !ExcelChartSeriesContract.FormulaMentionsValues(formula, spec.Range, defaultSheet, ownerWorkbook) &&
                !ExcelDataRangeBinding.SourceReadbackMatches(
                    spec.Range, Json.GetString(got, "values"), defaultSheet, ownerWorkbook) &&
                !ExcelDataRangeBinding.SourceReadbackMatches(
                    spec.Range, Json.GetString(got, "range"), defaultSheet, ownerWorkbook))
                mismatches.Add($"{label}: series[{index}].range '{spec.Range}' not applied as values");
            if (hasSpec && spec.Name is not null && spec.NameIsRange)
            {
                if (!ExcelReferenceIdentity.SeriesNameReferenceMatches(formula, spec.Name, defaultSheet, ownerWorkbook) &&
                    !ExcelChartSeriesContract.FormulaMentionsName(formula, spec.Name, defaultSheet, ownerWorkbook) &&
                    !ExcelDataRangeBinding.SourceReadbackMatches(
                        spec.Name.TrimStart('='), Json.GetString(got, "name"), defaultSheet, ownerWorkbook))
                    mismatches.Add($"{label}: series[{index}].name reference '{spec.Name}' not in formula '{formula}'");
            }
            else if (hasSpec && spec.Name is not null &&
                !string.Equals(Json.GetString(got, "name"), spec.Name, StringComparison.Ordinal) &&
                !ExcelReferenceIdentity.SeriesNameEquals(formula, spec.Name))
            {
                mismatches.Add($"{label}: series[{index}].name '{Json.GetString(got, "name")}' != '{spec.Name}'");
            }
            var expectedType = hasSpec
                ? ExcelChartSeriesContract.EffectiveChartType(spec, chartLevelType)
                : chartLevelType;
            if (expectedType is not null &&
                !string.Equals(Json.GetString(got, "chartType"), expectedType, StringComparison.OrdinalIgnoreCase))
                mismatches.Add($"{label}: series[{index}].chartType '{Json.GetString(got, "chartType")}' != '{expectedType}'");
            if (hasSpec && spec.AxisGroup is not null &&
                !string.Equals(Json.GetString(got, "axisGroup"), spec.AxisGroup, StringComparison.OrdinalIgnoreCase))
                mismatches.Add($"{label}: series[{index}].axisGroup '{Json.GetString(got, "axisGroup")}' != '{spec.AxisGroup}'");
        }
    }

    public static void CompareRequestedChartAxes(JsonObject actual, JsonObject op, string label,
        ICollection<string> mismatches) =>
        ExcelChartAxesContract.CompareRequested(actual, op, label, mismatches);

    public static void CompareSortReadback(JsonArray? before, JsonArray? after,
        IReadOnlyList<(int GridColumn, bool Descending)> keys, bool hasHeaders,
        string label, ICollection<string> mismatches)
    {
        ArgumentNullException.ThrowIfNull(keys);
        ArgumentNullException.ThrowIfNull(mismatches);
        if (before is null || after is null)
        {
            mismatches.Add($"{label}: sort readback missing before/after grids");
            return;
        }
        if (keys.Count == 0)
        {
            mismatches.Add($"{label}: sort readback has no keys");
            return;
        }

        var beforeRows = MaterializeRows(before);
        var afterRows = MaterializeRows(after);
        if (beforeRows.Count != afterRows.Count)
        {
            mismatches.Add($"{label}: sort row count {afterRows.Count} != before {beforeRows.Count}");
            return;
        }

        var dataStart = hasHeaders ? 1 : 0;
        if (hasHeaders && beforeRows.Count > 0 && !RowEquals(beforeRows[0], afterRows[0]))
            mismatches.Add($"{label}: sort header row was not preserved");

        var beforeData = beforeRows.Skip(dataStart).ToList();
        var afterData = afterRows.Skip(dataStart).ToList();
        var beforeBag = beforeData.Select(RowKey).OrderBy(item => item, StringComparer.Ordinal).ToList();
        var afterBag = afterData.Select(RowKey).OrderBy(item => item, StringComparer.Ordinal).ToList();
        if (!beforeBag.SequenceEqual(afterBag, StringComparer.Ordinal))
            mismatches.Add($"{label}: sort did not preserve row associations (not a permutation of input rows)");

        for (var i = 1; i < afterData.Count; i++)
        {
            if (CompareKeyedRows(afterData[i - 1], afterData[i], keys) > 0)
            {
                mismatches.Add($"{label}: multi-key sort order broken at data row {i + dataStart + 1}");
                return;
            }
        }
    }

    public static void CompareTextToColumnsReadback(JsonArray? sourceBefore, JsonArray? destinationAfter,
        JsonObject op, string label, ICollection<string> mismatches)
    {
        ArgumentNullException.ThrowIfNull(op);
        ArgumentNullException.ThrowIfNull(mismatches);
        if (sourceBefore is null || destinationAfter is null)
        {
            mismatches.Add($"{label}: text_to_columns readback missing source/destination grids");
            return;
        }

        var sourceRows = MaterializeRows(sourceBefore);
        var destRows = MaterializeRows(destinationAfter);
        var delimiters = TextToColumnsDelimiters(op);
        if (delimiters.Length == 0)
        {
            mismatches.Add($"{label}: text_to_columns has no delimiter to verify");
            return;
        }

        var rows = Math.Max(sourceRows.Count, destRows.Count);
        if (destRows.Count < sourceRows.Count)
            mismatches.Add($"{label}: destination has {destRows.Count} rows, source has {sourceRows.Count}");

        for (var r = 0; r < sourceRows.Count; r++)
        {
            var text = sourceRows[r].Count > 0 ? sourceRows[r][0] : "";
            if (string.IsNullOrEmpty(text)) continue;
            var expected = ExcelTextToColumnsParse.ExpectedFields(text, op);
            if (r >= destRows.Count)
            {
                mismatches.Add($"{label}: destination missing row {r + 1} (expected {expected.Length} fields)");
                continue;
            }
            var actual = destRows[r];
            for (var c = 0; c < expected.Length; c++)
            {
                var got = c < actual.Count ? actual[c] : "";
                if (!string.Equals(got, expected[c], StringComparison.Ordinal))
                {
                    mismatches.Add(
                        $"{label}: destination row {r + 1} col {c + 1} is '{got}', expected '{expected[c]}' from '{text}'");
                    return;
                }
            }
            for (var c = expected.Length; c < actual.Count; c++)
            {
                if (!string.IsNullOrEmpty(actual[c]))
                {
                    mismatches.Add($"{label}: destination row {r + 1} col {c + 1} is leftover '{actual[c]}'");
                    return;
                }
            }
        }
    }

    public static string[] TextToColumnsDelimiters(JsonObject op)
    {
        var delimiters = new List<string>();
        if (Json.GetBool(op, "other") && !string.IsNullOrWhiteSpace(Json.GetString(op, "otherChar")))
            delimiters.Add(Json.GetString(op, "otherChar")!);
        if (Json.GetBool(op, "semicolon")) delimiters.Add(";");
        if (Json.GetBool(op, "tab")) delimiters.Add("\t");
        if (Json.GetBool(op, "space")) delimiters.Add(" ");
        if (!op.ContainsKey("comma") || Json.GetBool(op, "comma"))
            delimiters.Add(",");
        return delimiters.ToArray();
    }

    public static bool IsNameOnlyPivot(JsonObject actual) =>
        !HasAnyAggregate(actual) &&
        (Json.GetArr(actual, "dataFields") is null || Json.GetArr(actual, "dataFields")!.Count == 0);

    private static void CompareFieldList(JsonObject actual, string actualKey, JsonObject op, string requestedKey,
        string label, ICollection<string> mismatches)
    {
        var requested = Json.GetArr(op, requestedKey);
        if (requested is null) return;
        var got = Json.GetArr(actual, actualKey) ?? new JsonArray();
        var names = got.Select(node => node?.ToString() ?? "").ToList();
        for (var i = 0; i < requested.Count; i++)
        {
            var want = requested[i]?.ToString();
            if (string.IsNullOrWhiteSpace(want)) continue;
            if (!names.Any(name => name.Contains(want, StringComparison.OrdinalIgnoreCase)))
                mismatches.Add($"{label}: {requestedKey}[{i}] '{want}' missing from {actualKey}");
        }
    }

    private static bool HasAnyAggregate(JsonObject actual)
    {
        if (Json.GetArr(actual, "aggregates") is JsonArray aggregates &&
            aggregates.OfType<JsonObject>().Any(HasFiniteAggregate))
            return true;
        return Json.GetArr(actual, "dataFields") is JsonArray fields &&
               fields.OfType<JsonObject>().Any(HasFiniteAggregate);
    }

    private static bool HasFiniteAggregate(JsonObject item)
    {
        if (ExcelDataOperationsContract.TryGetFiniteNumber(item["value"], out _))
            return true;
        if (item["values"] is JsonArray values)
        {
            foreach (var node in values)
            {
                if (ExcelDataOperationsContract.TryGetFiniteNumber(node, out _))
                    return true;
            }
        }
        return Json.GetBool(item, "hasAggregate");
    }

    private static bool IsAggregateFunction(string? function) =>
        function is "sum" or "count" or "average" or "max" or "min" or "countNums" or
            "product" or "stdDev" or "var";

    private static string? ReadOwnerWorkbook(JsonObject actual)
    {
        var owner = Json.GetString(actual, "ownerWorkbook");
        if (!string.IsNullOrWhiteSpace(owner)) return owner.Trim();
        var workbook = Json.GetString(actual, "workbook");
        return string.IsNullOrWhiteSpace(workbook) ? null : workbook.Trim();
    }

    private static bool DataFieldMentions(JsonObject actualField, string? requested) =>
        FieldMentions(Json.GetString(actualField, "sourceName"), requested) ||
        FieldMentions(Json.GetString(actualField, "name"), requested) ||
        FieldMentions(Json.GetString(actualField, "caption"), requested) ||
        FieldMentions(Json.GetString(actualField, "field"), requested);

    private static bool FieldMentions(string? actual, string? requested) =>
        !string.IsNullOrWhiteSpace(requested) &&
        !string.IsNullOrWhiteSpace(actual) &&
        actual.Contains(requested, StringComparison.OrdinalIgnoreCase);

    private static bool ColumnMatches(JsonObject actual, string? column)
    {
        if (string.IsNullOrWhiteSpace(column)) return false;
        var name = Json.GetString(actual, "column") ?? "";
        if (name.Equals(column, StringComparison.OrdinalIgnoreCase)) return true;
        if (ExcelDataOperationsContract.TryGetFiniteNumber(actual["index"], out var index) &&
            column == index.ToString(CultureInfo.InvariantCulture))
            return true;
        return name.Contains(column, StringComparison.OrdinalIgnoreCase);
    }

    private static List<List<string>> MaterializeRows(JsonArray grid)
    {
        var rows = new List<List<string>>();
        foreach (var node in grid)
        {
            if (node is JsonArray row)
                rows.Add(row.Select(CellText).ToList());
            else
                rows.Add(new List<string> { CellText(node) });
        }
        return rows;
    }

    private static string RowKey(List<string> row) => string.Join('\u001f', row);

    private static bool RowEquals(List<string> left, List<string> right) =>
        left.Count == right.Count && left.SequenceEqual(right, StringComparer.Ordinal);

    private static int CompareKeyedRows(List<string> left, List<string> right,
        IReadOnlyList<(int GridColumn, bool Descending)> keys)
    {
        foreach (var (column, descending) in keys)
        {
            var cmp = CompareSortCells(CellAt(left, column), CellAt(right, column));
            if (cmp == 0) continue;
            return descending ? -cmp : cmp;
        }
        return 0;
    }

    private static string CellAt(List<string> row, int column) =>
        column >= 0 && column < row.Count ? row[column] : "";

    private static int CompareSortCells(string left, string right)
    {
        if (ExcelDataOperationsContract.TryGetFiniteNumber(JsonValue.Create(left), out var ln) &&
            ExcelDataOperationsContract.TryGetFiniteNumber(JsonValue.Create(right), out var rn))
            return ln.CompareTo(rn);
        if (double.TryParse(left, NumberStyles.Float, CultureInfo.InvariantCulture, out ln) &&
            double.TryParse(right, NumberStyles.Float, CultureInfo.InvariantCulture, out rn))
            return ln.CompareTo(rn);
        return string.Compare(left, right, StringComparison.OrdinalIgnoreCase);
    }

    private static string CellText(JsonNode? node) => node is null ? "" : node.ToString();

    private static string? ColumnKey(JsonNode? node)
    {
        if (node is JsonValue value)
        {
            if (value.TryGetValue<string>(out var text)) return text;
            if (ExcelDataOperationsContract.TryGetFiniteNumber(value, out var number))
                return number.ToString(CultureInfo.InvariantCulture);
        }
        return node?.ToString();
    }
}
