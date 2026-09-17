using System.Globalization;
using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// COM-free completeness and equality for data-objects snapshots.
/// Restore may return verified=true only after this comparison runs
/// against a live re-read. Missing actuals or incomplete payload
/// cannot be verified.
/// </summary>
public static class ExcelDataSnapshotEquality
{
    public const string TableCreateRollbackMethod = "Unlist";
    public const int ComparableSnapshotVersion = 2;

    public static readonly IReadOnlySet<string> SupportedKinds = new HashSet<string>(StringComparer.Ordinal)
    {
        "table", "sortRange", "autoFilter", "definedName", "validation",
        "conditionalFormat", "chart", "picture", "note", "hyperlink",
        "pivot", "duplicates", "textToColumns", "shape", "textbox", "richText",
        "connector", "externalLinks", "calculationMode", "freezeValues",
        "pasteSpecial", "goalSeek", "workbookProtection", "splitPanes",
        "sparkline", "slicer", "cellStyle",
    };

    public static IReadOnlyList<string> ExplainIncomplete(JsonObject? entry)
    {
        var errors = new List<string>();
        if (entry is null)
        {
            errors.Add("snapshot entry is missing");
            return errors;
        }

        var kind = Json.GetString(entry, "kind");
        if (string.IsNullOrWhiteSpace(kind) || !SupportedKinds.Contains(kind))
        {
            errors.Add($"snapshot kind '{kind}' is missing or not comparable");
            return errors;
        }

        var prefix = EntryPrefix(entry);
        switch (kind)
        {
            case "table":
                if (Json.GetBool(entry, "existed"))
                {
                    Require(entry, "state", errors, prefix + " existing table needs state");
                    var state = Json.GetObj(entry, "state");
                    if (string.IsNullOrWhiteSpace(Json.GetString(state, "name")))
                        errors.Add(prefix + " table state.name is required");
                    if (string.IsNullOrWhiteSpace(Json.GetString(state, "range")))
                        errors.Add(prefix + " table state.range is required");
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(Json.GetString(entry, "range")) &&
                        string.IsNullOrWhiteSpace(Json.GetString(entry, "name")))
                        errors.Add(prefix + " created table needs name or range");
                    if (!HasGrid(entry, "values") && !HasGrid(entry, "formulas"))
                        errors.Add(prefix + " created table needs values or formulas so Unlist can be verified");
                }
                break;
            case "sortRange":
            case "duplicates":
            case "textToColumns":
            case "freezeValues":
            case "pasteSpecial":
                RequireGrid(entry, "formulas", errors, prefix + " formulas");
                RequireGrid(entry, "values", errors, prefix + " values");
                RequireGrid(entry, "numberFormats", errors, prefix + " numberFormats");
                break;
            case "goalSeek":
                RequireGrid(entry, "formulas", errors, prefix + " formulas");
                RequireGrid(entry, "values", errors, prefix + " values");
                RequireGrid(entry, "numberFormats", errors, prefix + " numberFormats");
                if (Json.GetObj(entry, "changing") is null)
                    errors.Add(prefix + " changing cell is required");
                break;
            case "externalLinks":
                if (Json.GetArr(entry, "sources") is null)
                    errors.Add(prefix + " link sources are required");
                if ((Json.GetString(entry, "mode") ?? "") == "break" && Json.GetArr(entry, "dependents") is null)
                    errors.Add(prefix + " break dependents are required");
                break;
            case "calculationMode":
                if (Json.GetInt(entry, "mode") is null)
                    errors.Add(prefix + " calculation mode is required");
                break;
            case "workbookProtection":
                if (!entry.ContainsKey("structure") || !entry.ContainsKey("windows"))
                    errors.Add(prefix + " structure/windows flags are required");
                break;
            case "splitPanes":
                if (Json.GetObj(entry, "state") is null)
                    errors.Add(prefix + " split state is required");
                break;
            case "sparkline":
            case "slicer":
                if (Json.GetBool(entry, "existed") && Json.GetObj(entry, "state") is null)
                    errors.Add(prefix + $" existing {kind} needs state");
                break;
            case "cellStyle":
                if (Json.GetArr(entry, "styles") is null)
                    errors.Add(prefix + " style grid is required");
                break;
            case "autoFilter":
                Require(entry, "state", errors, prefix + " filter state");
                var filter = Json.GetObj(entry, "state");
                if (filter is not null && !filter.ContainsKey("filters"))
                    errors.Add(prefix + " filter criteria array is required");
                if (filter is not null && !filter.ContainsKey("enabled"))
                    errors.Add(prefix + " filter enabled flag is required");
                break;
            case "definedName":
                if (string.IsNullOrWhiteSpace(Json.GetString(entry, "name")))
                    errors.Add(prefix + " defined name is missing");
                if (Json.GetBool(entry, "existed"))
                {
                    var nameState = Json.GetObj(entry, "state");
                    if (string.IsNullOrWhiteSpace(Json.GetString(nameState, "refersTo")))
                        errors.Add(prefix + " deleted/updated name needs state.refersTo so it can be recreated");
                }
                break;
            case "validation":
                Require(entry, "state", errors, prefix + " validation state");
                var validation = Json.GetObj(entry, "state");
                if (validation is not null && Json.GetBool(validation, "present"))
                {
                    if (Json.GetInt(validation, "validationType") is null && Json.GetInt(validation, "type") is null)
                        errors.Add(prefix + " validation type is required");
                    foreach (var field in new[] { "ignoreBlank", "inCellDropdown", "showInput", "showError" })
                    {
                        if (!validation.ContainsKey(field))
                            errors.Add(prefix + $" validation.{field} is required");
                    }
                }
                break;
            case "conditionalFormat":
                Require(entry, "state", errors, prefix + " conditional format state");
                var cf = Json.GetObj(entry, "state");
                var rules = Json.GetArr(cf, "rules") ?? new JsonArray();
                if (cf is not null && !cf.ContainsKey("count") && rules.Count == 0 && !cf.ContainsKey("rules"))
                    errors.Add(prefix + " conditional format rules are required");
                for (var i = 0; i < rules.Count; i++)
                {
                    if (rules[i] is not JsonObject rule)
                    {
                        errors.Add(prefix + $" rules[{i}] must be an object");
                        continue;
                    }
                    if (Json.GetBool(rule, "unreadable"))
                        errors.Add(prefix + $" rules[{i}] is unreadable and cannot be verified");
                }
                break;
            case "richText":
                Require(entry, "state", errors, prefix + " rich text state");
                if (!Json.GetBool(entry, "coverageComplete")) errors.Add(prefix + " rich text snapshot coverage is incomplete");
                break;
            case "chart":
                if (Json.GetBool(entry, "existed"))
                {
                    var chart = Json.GetObj(entry, "state");
                    if (chart is null)
                        errors.Add(prefix + " existing chart needs state");
                    else if ((Json.GetInt(chart, "seriesCount") ?? 0) <= 0 &&
                             Json.GetArr(chart, "series") is not { Count: > 0 } &&
                             string.IsNullOrWhiteSpace(Json.GetString(chart, "series1Formula")) &&
                             string.IsNullOrWhiteSpace(Json.GetString(chart, "sourceRange")))
                        errors.Add(prefix + " chart series/source is required");
                }
                break;
            case "picture":
            case "shape":
            case "textbox":
            case "pivot":
                if (Json.GetBool(entry, "existed") && Json.GetObj(entry, "state") is null)
                    errors.Add(prefix + $" existing {kind} needs state");
                break;
            case "note":
                Require(entry, "state", errors, prefix + " note state");
                break;
            case "hyperlink":
                Require(entry, "state", errors, prefix + " hyperlink state");
                var link = Json.GetObj(entry, "state");
                if (link is not null && !link.ContainsKey("cellFormula") && !link.ContainsKey("cellValue"))
                    errors.Add(prefix + " hyperlink needs cellFormula or cellValue");
                if (link is not null && Json.GetBool(link, "present") && !link.ContainsKey("font"))
                    errors.Add(prefix + " hyperlink needs font snapshot");
                break;
        }

        return errors;
    }

    public static IReadOnlyList<string> Compare(JsonObject expected, JsonObject? actual)
    {
        var errors = new List<string>(ExplainIncomplete(expected));
        if (actual is null)
        {
            errors.Add($"{EntryPrefix(expected)} live readback is missing; equality was not performed");
            return errors;
        }
        if (Json.GetBool(actual, "readFailed"))
        {
            errors.Add($"{EntryPrefix(expected)} live readback failed: {Json.GetString(actual, "error") ?? "unknown"}");
            return errors;
        }

        var kind = Json.GetString(expected, "kind");
        var prefix = EntryPrefix(expected);
        switch (kind)
        {
            case "table":
                CompareTable(expected, actual, prefix, errors);
                break;
            case "sortRange":
            case "duplicates":
            case "textToColumns":
            case "freezeValues":
            case "pasteSpecial":
                CompareGrids(expected, actual, "formulas", prefix, errors);
                CompareGrids(expected, actual, "values", prefix, errors);
                CompareNotesAndLinks(expected, actual, prefix, errors);
                break;
            case "goalSeek":
                CompareGrids(expected, actual, "formulas", prefix, errors);
                CompareGrids(expected, actual, "values", prefix, errors);
                CompareNotesAndLinks(expected, actual, prefix, errors);
                CompareGoalChanging(Json.GetObj(expected, "changing"), Json.GetObj(actual, "changing"), prefix, errors);
                break;
            case "externalLinks":
                CompareLinkSets(expected, actual, prefix, errors);
                break;
            case "calculationMode":
                if ((Json.GetInt(expected, "mode") ?? -1) != (Json.GetInt(actual, "mode") ?? -2))
                    errors.Add(prefix + " calculation mode mismatch");
                break;
            case "workbookProtection":
                CompareBool(expected, actual, "structure", prefix + " structure", errors);
                CompareBool(expected, actual, "windows", prefix + " windows", errors);
                break;
            case "splitPanes":
                CompareSplit(Json.GetObj(expected, "state"), Json.GetObj(actual, "state") ?? actual, prefix, errors);
                break;
            case "sparkline":
                CompareSparkline(expected, actual, prefix, errors);
                break;
            case "slicer":
                CompareSlicer(expected, actual, prefix, errors);
                break;
            case "cellStyle":
                if (!GridsEqual(expected["styles"], actual["styles"]))
                    errors.Add(prefix + " styles mismatch");
                break;
            case "autoFilter":
                CompareFilter(Json.GetObj(expected, "state"), Json.GetObj(actual, "state") ?? actual, prefix, errors);
                break;
            case "definedName":
                CompareName(expected, actual, prefix, errors);
                break;
            case "validation":
                CompareValidation(Json.GetObj(expected, "state"), Json.GetObj(actual, "state") ?? actual, prefix, errors);
                break;
            case "conditionalFormat":
                CompareConditional(Json.GetObj(expected, "state"), Json.GetObj(actual, "state") ?? actual, prefix, errors);
                break;
            case "richText":
                if (!string.Equals(Json.Canonical(Json.GetObj(expected, "state")), Json.Canonical(Json.GetObj(actual, "state") ?? actual), StringComparison.Ordinal))
                    errors.Add(prefix + " rich text differs after restore");
                break;
            case "chart":
                CompareChart(expected, actual, prefix, errors);
                break;
            case "picture":
            case "shape":
            case "textbox":
            case "connector":
                CompareNamedObject(expected, actual, prefix, errors);
                break;
            case "pivot":
                ComparePivot(expected, actual, prefix, errors);
                break;
            case "note":
                CompareNote(Json.GetObj(expected, "state"), Json.GetObj(actual, "state") ?? actual, prefix, errors);
                break;
            case "hyperlink":
                CompareHyperlink(Json.GetObj(expected, "state"), Json.GetObj(actual, "state") ?? actual, prefix, errors);
                break;
        }

        return errors;
    }

    /// <summary>
    /// Evaluates a whole snapshot. <paramref name="liveActualsByEntryIndex"/> must
    /// align with <c>entries</c>. Null actuals mean equality was not run.
    /// </summary>
    public static JsonObject Evaluate(JsonObject? state, JsonArray? liveActualsByEntryIndex)
    {
        var mismatches = new List<string>();
        var version = Json.GetInt(state, "snapshotVersion");
        var restoreMode = Json.GetString(state, "restoreMode");
        if (!string.Equals(restoreMode, ExcelDataOperationsContract.RestoreMode, StringComparison.Ordinal))
            mismatches.Add("restoreMode is not data-objects");
        if (version is null || version < 1)
            mismatches.Add("snapshotVersion is missing or invalid");
        if (version is not null && version < ComparableSnapshotVersion)
            mismatches.Add($"snapshotVersion {version} is incomplete; equality requires version {ComparableSnapshotVersion}");

        var entries = Json.GetArr(state, "entries") ?? new JsonArray();
        var compared = 0;
        if (liveActualsByEntryIndex is null)
        {
            mismatches.Add("equality comparison was not performed; verified cannot be true");
            foreach (var node in entries.OfType<JsonObject>())
                mismatches.AddRange(ExplainIncomplete(node));
        }
        else
        {
            for (var i = 0; i < entries.Count; i++)
            {
                if (entries[i] is not JsonObject entry) continue;
                var actual = i < liveActualsByEntryIndex.Count ? liveActualsByEntryIndex[i] as JsonObject : null;
                var found = Compare(entry, actual);
                mismatches.AddRange(found);
                compared++;
            }
        }

        var verified = mismatches.Count == 0 && liveActualsByEntryIndex is not null;
        return new JsonObject
        {
            ["verified"] = verified,
            ["compared"] = compared,
            ["tableCreateRollback"] = TableCreateRollbackMethod,
            ["equalityRequired"] = true,
            ["mismatches"] = new JsonArray(mismatches.Select(item => JsonValue.Create(item)).ToArray()),
        };
    }

    public static bool GridsEqual(JsonNode? left, JsonNode? right)
    {
        if (left is null && right is null) return true;
        if (left is null || right is null) return false;
        if (left is JsonArray leftRows && right is JsonArray rightRows)
        {
            if (leftRows.Count != rightRows.Count) return false;
            for (var r = 0; r < leftRows.Count; r++)
            {
                if (!GridsEqual(leftRows[r], rightRows[r])) return false;
            }
            return true;
        }
        return ScalarEqual(left, right);
    }

    private static void CompareTable(JsonObject expected, JsonObject actual, string prefix, List<string> errors)
    {
        if (!Json.GetBool(expected, "existed"))
        {
            if (actual.ContainsKey("present") && Json.GetBool(actual, "present"))
                errors.Add(prefix + " created table is still present after Unlist rollback");
            CompareGrids(expected, actual, "formulas", prefix, errors);
            CompareGrids(expected, actual, "values", prefix, errors);
            return;
        }

        var expectedState = Json.GetObj(expected, "state") ?? expected;
        var actualState = Json.GetObj(actual, "state") ?? actual;
        CompareString(expectedState, actualState, "name", prefix + " table name", errors);
        CompareString(expectedState, actualState, "range", prefix + " table range", errors, ignoreCase: true);
        CompareBool(expectedState, actualState, "showTotals", prefix + " showTotals", errors);
        CompareBool(expectedState, actualState, "showHeaders", prefix + " showHeaders", errors);
        CompareGrids(expected, actual, "formulas", prefix, errors);
        CompareGrids(expected, actual, "values", prefix, errors);
    }

    private static void CompareFilter(JsonObject? expected, JsonObject? actual, string prefix, List<string> errors)
    {
        if (expected is null || actual is null)
        {
            errors.Add(prefix + " filter state missing on compare");
            return;
        }
        CompareBool(expected, actual, "enabled", prefix + " AutoFilterMode", errors);
        CompareBool(expected, actual, "filtered", prefix + " FilterMode", errors);
        var expectedFilters = Json.GetArr(expected, "filters") ?? new JsonArray();
        var actualFilters = Json.GetArr(actual, "filters") ?? new JsonArray();
        if (expectedFilters.Count != actualFilters.Count)
            errors.Add($"{prefix} filter criteria count {actualFilters.Count} != {expectedFilters.Count}");
        var count = Math.Min(expectedFilters.Count, actualFilters.Count);
        for (var i = 0; i < count; i++)
        {
            if (expectedFilters[i] is not JsonObject exp || actualFilters[i] is not JsonObject act) continue;
            CompareBool(exp, act, "on", prefix + $" filters[{i}].on", errors);
            if (!ScalarEqual(exp["criteria1"], act["criteria1"]))
                errors.Add($"{prefix} filters[{i}].criteria1 mismatch");
            if (!ScalarEqual(exp["criteria2"], act["criteria2"]))
                errors.Add($"{prefix} filters[{i}].criteria2 mismatch");
        }
    }

    private static void CompareName(JsonObject expected, JsonObject actual, string prefix, List<string> errors)
    {
        var existed = Json.GetBool(expected, "existed");
        var present = actual.ContainsKey("present") ? Json.GetBool(actual, "present") : Json.GetObj(actual, "state") is not null
            || !string.IsNullOrWhiteSpace(Json.GetString(actual, "refersTo"));
        if (!existed)
        {
            if (present) errors.Add(prefix + " defined name still exists after rollback");
            return;
        }
        if (!present)
        {
            errors.Add(prefix + " defined name was not recreated");
            return;
        }
        var expectedState = Json.GetObj(expected, "state") ?? expected;
        var actualState = Json.GetObj(actual, "state") ?? actual;
        var expectedRef = Json.GetString(expectedState, "refersTo") ?? "";
        var actualRef = Json.GetString(actualState, "refersTo") ?? "";
        if (!RefersToEqual(expectedRef, actualRef))
            errors.Add($"{prefix} RefersTo '{actualRef}' != '{expectedRef}'");
    }

    private static void CompareValidation(JsonObject? expected, JsonObject? actual, string prefix, List<string> errors)
    {
        if (expected is null || actual is null)
        {
            errors.Add(prefix + " validation state missing on compare");
            return;
        }
        CompareBool(expected, actual, "present", prefix + " validation.present", errors);
        if (!Json.GetBool(expected, "present")) return;
        var expectedType = Json.GetInt(expected, "validationType") ?? Json.GetInt(expected, "type");
        var actualType = Json.GetInt(actual, "validationType") ?? Json.GetInt(actual, "type");
        if (expectedType != actualType)
            errors.Add($"{prefix} validation type {actualType} != {expectedType}");
        CompareString(expected, actual, "formula1", prefix + " formula1", errors, ignoreCase: true);
        CompareString(expected, actual, "formula2", prefix + " formula2", errors, ignoreCase: true);
        CompareBool(expected, actual, "ignoreBlank", prefix + " ignoreBlank", errors);
        CompareBool(expected, actual, "inCellDropdown", prefix + " inCellDropdown", errors);
        CompareBool(expected, actual, "showInput", prefix + " showInput", errors);
        CompareBool(expected, actual, "showError", prefix + " showError", errors);
    }

    private static void CompareConditional(JsonObject? expected, JsonObject? actual, string prefix, List<string> errors)
    {
        if (expected is null || actual is null)
        {
            errors.Add(prefix + " conditional format state missing on compare");
            return;
        }
        var expectedCount = Json.GetInt(expected, "count") ?? Json.GetArr(expected, "rules")?.Count ?? 0;
        var actualCount = Json.GetInt(actual, "count") ?? Json.GetArr(actual, "rules")?.Count ?? 0;
        if (expectedCount != actualCount)
            errors.Add($"{prefix} FormatConditions.Count {actualCount} != {expectedCount}");
        var expectedRules = Json.GetArr(expected, "rules") ?? new JsonArray();
        var actualRules = Json.GetArr(actual, "rules") ?? new JsonArray();
        var count = Math.Min(expectedRules.Count, actualRules.Count);
        for (var i = 0; i < count; i++)
        {
            if (expectedRules[i] is not JsonObject exp || actualRules[i] is not JsonObject act) continue;
            var expType = Json.GetInt(exp, "conditionType") ?? Json.GetInt(exp, "type");
            var actType = Json.GetInt(act, "conditionType") ?? Json.GetInt(act, "type");
            if (expType != actType)
                errors.Add($"{prefix} rules[{i}].type {actType} != {expType}");
            if (!string.Equals(Json.GetString(exp, "formula1") ?? "", Json.GetString(act, "formula1") ?? "",
                    StringComparison.OrdinalIgnoreCase))
                errors.Add($"{prefix} rules[{i}].formula1 mismatch");
            CompareStyle(Json.GetObj(exp, "style"), Json.GetObj(act, "style"), prefix + $" rules[{i}].style", errors);
        }
    }

    private static void CompareChart(JsonObject expected, JsonObject actual, string prefix, List<string> errors)
    {
        if (!Json.GetBool(expected, "existed"))
        {
            if (Json.GetBool(actual, "present"))
                errors.Add(prefix + " created chart is still present after rollback");
            return;
        }
        if (actual.ContainsKey("present") && !Json.GetBool(actual, "present"))
        {
            errors.Add(prefix + " chart is missing after restore");
            return;
        }
        var expectedState = Json.GetObj(expected, "state") ?? expected;
        var actualState = Json.GetObj(actual, "state") ?? actual;
        var expectedSeries = Json.GetArr(expectedState, "series");
        var skipChartType = ExcelChartSeriesContract.DistinctRequestedChartTypes(expectedSeries, null).Count > 1;
        if (!skipChartType)
            CompareString(expectedState, actualState, "chartType", prefix + " chartType", errors, ignoreCase: true);
        CompareString(expectedState, actualState, "title", prefix + " title", errors);
        var actualSeries = Json.GetArr(actualState, "series");
        if (expectedSeries is { Count: > 0 })
        {
            if (actualSeries is null || actualSeries.Count != expectedSeries.Count)
                errors.Add($"{prefix} series count {actualSeries?.Count ?? 0} != {expectedSeries.Count}");
            else
            {
                for (var i = 0; i < expectedSeries.Count; i++)
                {
                    var expObj = expectedSeries[i] as JsonObject;
                    var actObj = actualSeries[i] as JsonObject;
                    var exp = Json.GetString(expObj, "formula");
                    var act = Json.GetString(actObj, "formula");
                    if (!string.Equals(exp ?? "", act ?? "", StringComparison.OrdinalIgnoreCase))
                        errors.Add($"{prefix} series[{i}].formula mismatch");
                    if (expObj is null || actObj is null) continue;
                    if (expObj.ContainsKey("chartType") &&
                        !string.Equals(Json.GetString(expObj, "chartType"), Json.GetString(actObj, "chartType"),
                            StringComparison.OrdinalIgnoreCase))
                        errors.Add($"{prefix} series[{i}].chartType mismatch");
                    if (expObj.ContainsKey("axisGroup") &&
                        !string.Equals(Json.GetString(expObj, "axisGroup"), Json.GetString(actObj, "axisGroup"),
                            StringComparison.OrdinalIgnoreCase))
                        errors.Add($"{prefix} series[{i}].axisGroup mismatch");
                }
            }
        }
        else
        {
            var expFormula = Json.GetString(expectedState, "series1Formula");
            var actFormula = Json.GetString(actualState, "series1Formula");
            if (!string.IsNullOrWhiteSpace(expFormula) &&
                !string.Equals(expFormula, actFormula ?? "", StringComparison.OrdinalIgnoreCase))
                errors.Add(prefix + " series1Formula mismatch");
            var expCount = Json.GetInt(expectedState, "seriesCount") ?? 0;
            var actCount = Json.GetInt(actualState, "seriesCount") ?? 0;
            if (expCount != actCount)
                errors.Add($"{prefix} seriesCount {actCount} != {expCount}");
        }
        if (Json.GetObj(expectedState, "axes") is JsonObject expectedAxes)
        {
            ExcelChartAxesContract.CompareRequested(
                actualState,
                new JsonObject { ["axes"] = JsonNode.Parse(expectedAxes.ToJsonString())!.DeepClone() },
                prefix,
                errors);
        }
    }

    private static void CompareNamedObject(JsonObject expected, JsonObject actual, string prefix, List<string> errors)
    {
        if (!Json.GetBool(expected, "existed"))
        {
            if (Json.GetBool(actual, "present"))
                errors.Add(prefix + " created object is still present after rollback");
            return;
        }
        if (actual.ContainsKey("present") && !Json.GetBool(actual, "present"))
        {
            errors.Add(prefix + " object is missing after restore");
            return;
        }
        var expectedState = Json.GetObj(expected, "state") ?? expected;
        var actualState = Json.GetObj(actual, "state") ?? actual;
        CompareString(expectedState, actualState, "name", prefix + " name", errors, ignoreCase: true);
        CompareNumber(expectedState, actualState, "left", prefix + " left", errors);
        CompareNumber(expectedState, actualState, "top", prefix + " top", errors);
        CompareNumber(expectedState, actualState, "width", prefix + " width", errors);
        CompareNumber(expectedState, actualState, "height", prefix + " height", errors);
    }

    private static void ComparePivot(JsonObject expected, JsonObject actual, string prefix, List<string> errors)
    {
        if (!Json.GetBool(expected, "existed"))
        {
            if (Json.GetBool(actual, "present"))
                errors.Add(prefix + " created pivot is still present after rollback");
            return;
        }
        if (actual.ContainsKey("present") && !Json.GetBool(actual, "present"))
        {
            errors.Add(prefix + " pivot is missing after restore");
            return;
        }
        var expectedState = Json.GetObj(expected, "state") ?? expected;
        var actualState = Json.GetObj(actual, "state") ?? actual;
        CompareString(expectedState, actualState, "name", prefix + " pivot name", errors, ignoreCase: true);
        CompareStringArray(expectedState, actualState, "rowFields", prefix + " rowFields", errors);
        CompareStringArray(expectedState, actualState, "columnFields", prefix + " columnFields", errors);
        CompareStringArray(expectedState, actualState, "pageFields", prefix + " pageFields", errors);
    }

    private static void CompareNote(JsonObject? expected, JsonObject? actual, string prefix, List<string> errors)
    {
        if (expected is null || actual is null)
        {
            errors.Add(prefix + " note state missing on compare");
            return;
        }
        CompareBool(expected, actual, "present", prefix + " note.present", errors);
        if (Json.GetBool(expected, "present"))
            CompareString(expected, actual, "text", prefix + " note.text", errors);
    }

    private static void CompareHyperlink(JsonObject? expected, JsonObject? actual, string prefix, List<string> errors)
    {
        if (expected is null || actual is null)
        {
            errors.Add(prefix + " hyperlink state missing on compare");
            return;
        }
        CompareBool(expected, actual, "present", prefix + " hyperlink.present", errors);
        if (Json.GetBool(expected, "present"))
        {
            CompareString(expected, actual, "address", prefix + " address", errors, ignoreCase: true);
            CompareString(expected, actual, "subAddress", prefix + " subAddress", errors, ignoreCase: true);
        }
        if (expected.ContainsKey("cellFormula") &&
            !string.Equals(Json.GetString(expected, "cellFormula") ?? "", Json.GetString(actual, "cellFormula") ?? "",
                StringComparison.OrdinalIgnoreCase))
            errors.Add(prefix + " cellFormula mismatch");
        if (expected.ContainsKey("cellValue") && !ScalarEqual(expected["cellValue"], actual["cellValue"]))
            errors.Add(prefix + " cellValue mismatch");
        CompareStyle(Json.GetObj(expected, "font"), Json.GetObj(actual, "font"), prefix + " font", errors);
    }

    private static void CompareGoalChanging(JsonObject? expected, JsonObject? actual, string prefix, List<string> errors)
    {
        if (expected is null) return;
        if (actual is null)
        {
            errors.Add(prefix + " changing cell readback is missing");
            return;
        }
        CompareString(expected, actual, "address", prefix + " changing address", errors, ignoreCase: true);
        if (expected.ContainsKey("formula") &&
            !string.Equals(Json.GetString(expected, "formula") ?? "", Json.GetString(actual, "formula") ?? "",
                StringComparison.OrdinalIgnoreCase))
            errors.Add(prefix + " changing formula mismatch");
        if (expected.ContainsKey("value") && !ScalarEqual(expected["value"], actual["value"]))
            errors.Add(prefix + " changing value mismatch");
    }

    private static void CompareLinkSets(JsonObject expected, JsonObject actual, string prefix, List<string> errors)
    {
        var expSources = Json.GetArr(expected, "sources")?.Select(node => node?.ToString() ?? "")
            .ToHashSet(StringComparer.OrdinalIgnoreCase) ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var actSources = Json.GetArr(actual, "sources")?.Select(node => node?.ToString() ?? "")
            .ToHashSet(StringComparer.OrdinalIgnoreCase) ?? new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (!expSources.SetEquals(actSources))
            errors.Add(prefix + " link sources differ after restore");
        if ((Json.GetString(expected, "mode") ?? "") != "break") return;
        var expDeps = IndexDependents(Json.GetArr(expected, "dependents"));
        var actDeps = IndexDependents(Json.GetArr(actual, "dependents"));
        if (expDeps.Count != actDeps.Count)
        {
            errors.Add(prefix + $" dependent count {actDeps.Count} != {expDeps.Count}");
            return;
        }
        foreach (var (key, formula) in expDeps)
        {
            if (!actDeps.TryGetValue(key, out var actualFormula) ||
                !string.Equals(formula, actualFormula, StringComparison.OrdinalIgnoreCase))
            {
                errors.Add(prefix + $" dependent {key} formula mismatch");
                return;
            }
        }
    }

    private static Dictionary<string, string> IndexDependents(JsonArray? items)
    {
        var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        if (items is null) return result;
        foreach (var node in items.OfType<JsonObject>())
        {
            var key = $"{Json.GetString(node, "sheet")}!{Json.GetString(node, "address")}";
            result[key] = Json.GetString(node, "formula") ?? "";
        }
        return result;
    }

    private static void CompareSplit(JsonObject? expected, JsonObject? actual, string prefix, List<string> errors)
    {
        if (expected is null || actual is null)
        {
            errors.Add(prefix + " split state missing on compare");
            return;
        }
        CompareBool(expected, actual, "frozen", prefix + " frozen", errors);
        var expRow = Json.GetInt(expected, "splitRow");
        var actRow = Json.GetInt(actual, "splitRow");
        if (expRow is not null && expRow != actRow)
            errors.Add(prefix + $" splitRow {actRow} != {expRow}");
        var expCol = Json.GetInt(expected, "splitColumn");
        var actCol = Json.GetInt(actual, "splitColumn");
        if (expCol is not null && expCol != actCol)
            errors.Add(prefix + $" splitColumn {actCol} != {expCol}");
    }

    private static void CompareSparkline(JsonObject expected, JsonObject actual, string prefix, List<string> errors)
    {
        if (!Json.GetBool(expected, "existed"))
        {
            if (Json.GetBool(actual, "present"))
                errors.Add(prefix + " created sparkline is still present after rollback");
            return;
        }
        if (actual.ContainsKey("present") && !Json.GetBool(actual, "present"))
        {
            errors.Add(prefix + " sparkline is missing after restore");
            return;
        }
        var expectedState = Json.GetObj(expected, "state") ?? expected;
        var actualState = Json.GetObj(actual, "state") ?? actual;
        CompareString(expectedState, actualState, "sourceData", prefix + " sourceData", errors, ignoreCase: true);
        var expType = Json.GetInt(expectedState, "type");
        var actType = Json.GetInt(actualState, "type");
        if (expType is not null && expType != actType)
            errors.Add(prefix + $" type {actType} != {expType}");
    }

    private static void CompareSlicer(JsonObject expected, JsonObject actual, string prefix, List<string> errors)
    {
        if (!Json.GetBool(expected, "existed"))
        {
            if (Json.GetBool(actual, "present"))
                errors.Add(prefix + " created slicer is still present after rollback");
            return;
        }
        if (actual.ContainsKey("present") && !Json.GetBool(actual, "present"))
        {
            errors.Add(prefix + " slicer is missing after restore");
            return;
        }
        var expectedState = Json.GetObj(expected, "state") ?? expected;
        var actualState = Json.GetObj(actual, "state") ?? actual;
        CompareString(expectedState, actualState, "cacheName", prefix + " cacheName", errors, ignoreCase: true);
        CompareString(expectedState, actualState, "caption", prefix + " caption", errors);
        CompareString(expectedState, actualState, "sheet", prefix + " sheet", errors, ignoreCase: true);
    }

    private static void CompareNotesAndLinks(JsonObject expected, JsonObject actual, string prefix, List<string> errors)
    {
        if (expected.ContainsKey("notes") && !GridsEqual(expected["notes"], actual["notes"]))
            errors.Add(prefix + " notes mismatch");
        if (expected.ContainsKey("hyperlinks") && !GridsEqual(expected["hyperlinks"], actual["hyperlinks"]))
            errors.Add(prefix + " hyperlinks mismatch");
        if (expected.ContainsKey("numberFormats") && !GridsEqual(expected["numberFormats"], actual["numberFormats"]))
            errors.Add(prefix + " numberFormats mismatch");
    }

    private static void CompareGrids(JsonObject expected, JsonObject actual, string field, string prefix, List<string> errors)
    {
        if (!expected.ContainsKey(field)) return;
        if (!GridsEqual(expected[field], actual[field]))
            errors.Add($"{prefix} {field} mismatch");
    }

    private static void CompareStyle(JsonObject? expected, JsonObject? actual, string prefix, List<string> errors)
    {
        if (expected is null) return;
        if (actual is null)
        {
            errors.Add(prefix + " missing");
            return;
        }
        foreach (var key in new[] { "bold", "italic", "fontColor", "fillColor", "color" })
        {
            if (!expected.ContainsKey(key)) continue;
            if (!ScalarEqual(expected[key], actual[key]))
                errors.Add($"{prefix}.{key} mismatch");
        }
    }

    private static void CompareString(JsonObject expected, JsonObject actual, string field, string prefix,
        List<string> errors, bool ignoreCase = false)
    {
        if (!expected.ContainsKey(field)) return;
        var comparison = ignoreCase ? StringComparison.OrdinalIgnoreCase : StringComparison.Ordinal;
        if (!string.Equals(Json.GetString(expected, field) ?? "", Json.GetString(actual, field) ?? "", comparison))
            errors.Add($"{prefix} mismatch");
    }

    private static void CompareBool(JsonObject expected, JsonObject actual, string field, string prefix, List<string> errors)
    {
        if (!expected.ContainsKey(field)) return;
        if (Json.GetBool(expected, field) != Json.GetBool(actual, field))
            errors.Add($"{prefix} mismatch");
    }

    private static void CompareNumber(JsonObject expected, JsonObject actual, string field, string prefix, List<string> errors)
    {
        if (!expected.ContainsKey(field)) return;
        var hasExpected = ExcelDataOperationsContract.TryGetFiniteNumber(expected[field], out var exp);
        var hasActual = ExcelDataOperationsContract.TryGetFiniteNumber(actual[field], out var act);
        if (!hasExpected) return;
        if (!hasActual || Math.Abs(exp - act) > 1.5)
            errors.Add($"{prefix} {act} != {exp}");
    }

    private static void CompareStringArray(JsonObject expected, JsonObject actual, string field, string prefix,
        List<string> errors)
    {
        var left = Json.GetArr(expected, field);
        if (left is null) return;
        var right = Json.GetArr(actual, field) ?? new JsonArray();
        if (left.Count != right.Count)
        {
            errors.Add($"{prefix} count {right.Count} != {left.Count}");
            return;
        }
        for (var i = 0; i < left.Count; i++)
        {
            if (!string.Equals(left[i]?.ToString(), right[i]?.ToString(), StringComparison.OrdinalIgnoreCase))
                errors.Add($"{prefix}[{i}] mismatch");
        }
    }

    private static bool ScalarEqual(JsonNode? left, JsonNode? right)
    {
        if (left is null && right is null) return true;
        if (left is null || right is null) return string.IsNullOrEmpty(left?.ToString()) && string.IsNullOrEmpty(right?.ToString());
        if (left is JsonArray || right is JsonArray) return GridsEqual(left, right);
        if (ExcelDataOperationsContract.TryGetFiniteNumber(left, out var ln) &&
            ExcelDataOperationsContract.TryGetFiniteNumber(right, out var rn))
            return Math.Abs(ln - rn) < 0.0000001;
        return string.Equals(NormalizeScalar(left), NormalizeScalar(right), StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeScalar(JsonNode node)
    {
        if (node is JsonValue value)
        {
            if (value.TryGetValue<bool>(out var flag)) return flag ? "true" : "false";
            if (value.TryGetValue<string>(out var text)) return text ?? "";
        }
        return node.ToJsonString();
    }

    private static bool RefersToEqual(string left, string right) =>
        ExcelFormulaReference.SemanticEquals(left, right);

    private static bool HasGrid(JsonObject entry, string field) =>
        entry[field] is JsonArray { Count: > 0 };

    private static void Require(JsonObject entry, string field, List<string> errors, string message)
    {
        if (!entry.ContainsKey(field) || entry[field] is null)
            errors.Add(message);
    }

    private static void RequireGrid(JsonObject entry, string field, List<string> errors, string label)
    {
        if (!HasGrid(entry, field))
            errors.Add(label + " grid is required");
    }

    private static string EntryPrefix(JsonObject entry)
    {
        var kind = Json.GetString(entry, "kind") ?? "entry";
        var sheet = Json.GetString(entry, "sheet");
        var name = Json.GetString(entry, "name") ?? Json.GetString(entry, "range");
        return string.IsNullOrWhiteSpace(sheet) ? $"{kind} '{name}':" : $"{kind} {sheet}!{name}:";
    }
}
