using System.Globalization;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;

namespace DocBridge.Core.Services;

/// <summary>
/// Formula2 is the dynamic-array dialect. Range.Formula applies implicit
/// intersection and cannot spill. HasSpill mixed-null must stay null.
/// Existing set_formulas clients keep Range.Formula unless they opt in.
/// </summary>
public static class ExcelFormula2Contract
{
    public const string EngineFormula2 = "formula2";
    public const string EngineFormula = "formula";
    public const string RestoreMode = "formula2-range";
    public const string SpillPropertySpillingToRange = "SpillingToRange";
    public const string SpillPropertySpillRangeLegacy = "SpillRange";

    private static readonly Regex SequencePattern = new(
        @"^=\s*SEQUENCE\s*\(\s*(\d+)\s*(?:,\s*(\d+))?",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

    private static readonly Regex OpenDynamicArrayPattern = new(
        @"^=\s*(FILTER|SORT|SORTBY|UNIQUE|RANDARRAY)\s*\(",
        RegexOptions.IgnoreCase | RegexOptions.CultureInvariant | RegexOptions.Compiled);

    public static string ResolveEngine(JsonObject op)
    {
        var explicitEngine = Json.GetString(op, "engine")
            ?? Json.GetString(op, "formulaEngine")
            ?? Json.GetString(op, "dialect");
        if (!string.IsNullOrWhiteSpace(explicitEngine))
        {
            if (explicitEngine.Equals(EngineFormula, StringComparison.OrdinalIgnoreCase))
                return EngineFormula;
            if (explicitEngine.Equals(EngineFormula2, StringComparison.OrdinalIgnoreCase))
                return EngineFormula2;
            throw new InvalidOperationException("set_formulas engine must be formula or formula2");
        }

        if (op.ContainsKey("formula2") && Json.GetBool(op, "formula2"))
            return EngineFormula2;

        // Preserve existing scalar Range.Formula evaluation. Formula2 is opt-in.
        return EngineFormula;
    }

    public static bool WritesFormula2(JsonObject op) =>
        string.Equals(ResolveEngine(op), EngineFormula2, StringComparison.OrdinalIgnoreCase);

    public static bool ShouldRecordSpill(bool? hasSpill, string? spillRange) =>
        hasSpill == true && !string.IsNullOrWhiteSpace(spillRange);

    public static bool MustNotClearNeighbors(bool? hasSpill) => hasSpill != true;

    public static bool RestoreClearsNewlyOwnedBeforeOldRestore => true;
    public static bool TreatsFailedSpillLookupAsDestScalar => false;
    public static bool WriteCountIsRestoredValueProof => false;

    /// <summary>
    /// Native property is <see cref="SpillPropertySpillingToRange"/>. A failed
    /// lookup is unknown, not a dest-sized scalar. <paramref name="spillRangeFallback"/>
    /// is ignored so legacy SpillRange cannot stand in for a failed read.
    /// </summary>
    public static string? PreferSpillingToRange(string? spillingToRange, string? spillRangeFallback)
    {
        _ = spillRangeFallback;
        return string.IsNullOrWhiteSpace(spillingToRange) ? null : CleanA1(spillingToRange);
    }

    public static string? AcceptSpillAddress(string dest, string? spillingToRange, bool spillingToRangeOk)
    {
        _ = dest;
        if (TreatsFailedSpillLookupAsDestScalar)
            return CleanA1(dest);
        if (!spillingToRangeOk || string.IsNullOrWhiteSpace(spillingToRange))
            return null;
        return CleanA1(spillingToRange);
    }

    public static IReadOnlyList<string> ClearOrderBeforeOldRestore(
        string dest, string captured, string? newSpill, bool? hasSpill)
    {
        var order = new List<string>();
        dest = CleanA1(dest);
        if (!string.IsNullOrWhiteSpace(dest))
            order.Add(dest);
        if (MustNotClearNeighbors(hasSpill) || string.IsNullOrWhiteSpace(newSpill))
            return order;
        captured = string.IsNullOrWhiteSpace(captured) ? dest : CleanA1(captured);
        foreach (var extra in NewlyOwnedSpillRanges(captured, newSpill))
            order.Add(extra);
        return order;
    }

    public static bool RestoredValuesMatch(JsonNode? expected, JsonNode? actual)
    {
        if (expected is null && actual is null)
            return true;
        if (expected is not JsonArray expectedRows || actual is not JsonArray actualRows)
            return false;
        if (expectedRows.Count != actualRows.Count)
            return false;
        for (var row = 0; row < expectedRows.Count; row++)
        {
            if (expectedRows[row] is not JsonArray expectedRow ||
                actualRows[row] is not JsonArray actualRow ||
                expectedRow.Count != actualRow.Count)
                return false;
            for (var column = 0; column < expectedRow.Count; column++)
            {
                if (!JsonCellsTypedEqual(expectedRow[column], actualRow[column]))
                    return false;
            }
        }

        return true;
    }

    private static bool JsonCellsTypedEqual(JsonNode? expected, JsonNode? actual)
    {
        var want = ExcelValueWriteContract.Classify(expected);
        if (actual is null)
            return ExcelValueWriteContract.TypedEqual(want, null);
        if (actual is not JsonValue value)
            return false;
        if (value.GetValueKind() == System.Text.Json.JsonValueKind.Null)
            return ExcelValueWriteContract.TypedEqual(want, null);
        if (value.TryGetValue<string>(out var text))
            return ExcelValueWriteContract.TypedEqual(want, text);
        if (value.TryGetValue<bool>(out var flag))
            return ExcelValueWriteContract.TypedEqual(want, flag);
        return ExcelValueWriteContract.TypedEqual(want, ExcelValueWriteContract.Classify(actual).ComValue);
    }

    public static IReadOnlyList<string> NeighborAndOldSpillRanges(string destAddress, string capturedOrOldSpill)
    {
        if (string.IsNullOrWhiteSpace(capturedOrOldSpill))
            return [];
        return NewlyOwnedSpillRanges(destAddress, capturedOrOldSpill);
    }

    public static bool RestoredAnchorMatches(string? expectedFormula2, string? actualFormula2) =>
        string.Equals(expectedFormula2 ?? "", actualFormula2 ?? "", StringComparison.Ordinal);

    public static (int Rows, int Columns) UnionFootprint(
        string destAddress, int capturedRows, int capturedColumns, string? oldSpill)
    {
        var rows = Math.Max(1, capturedRows);
        var columns = Math.Max(1, capturedColumns);
        if (string.IsNullOrWhiteSpace(oldSpill) || !SpillAnchoredAtDest(destAddress, oldSpill))
            return (rows, columns);
        if (!TryParseA1Size(oldSpill, out var spillRows, out var spillCols))
            return (rows, columns);
        return (Math.Max(rows, spillRows), Math.Max(columns, spillCols));
    }

    public static bool MixedBatchKeepsPerOpEngine(IReadOnlyList<JsonObject> ops)
    {
        var engines = ops
            .Where(op => string.Equals(Json.GetString(op, "op"), "set_formulas", StringComparison.OrdinalIgnoreCase))
            .Select(ResolveEngine)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        return engines.Count > 1;
    }

    public static bool? NormalizeHasSpill(bool? value) => value;

    public static bool? HasSpillFromCom(object? raw) =>
        raw switch
        {
            null => null,
            bool flag => flag,
            _ => null,
        };

    public static JsonObject SpillReadback(
        string address,
        string engine,
        bool? hasSpill,
        string? spillRange,
        string? formula2 = null,
        JsonNode? formulas = null,
        JsonNode? spillValues = null)
    {
        var readback = new JsonObject
        {
            ["address"] = address,
            ["engine"] = engine,
            ["hasSpill"] = hasSpill is bool flag ? JsonValue.Create(flag) : null,
            ["spillRange"] = spillRange,
        };
        if (formula2 is not null)
            readback["formula2"] = formula2;
        if (formulas is not null)
            readback["formulas"] = formulas.DeepClone();
        if (spillValues is not null)
            readback["spillValues"] = spillValues.DeepClone();
        return readback;
    }

    public static bool TryEstimateSequenceFootprint(string? formula, out int rows, out int columns)
    {
        rows = 0;
        columns = 0;
        if (string.IsNullOrWhiteSpace(formula))
            return false;
        var match = SequencePattern.Match(formula.Trim());
        if (!match.Success)
            return false;
        if (!int.TryParse(match.Groups[1].Value, NumberStyles.None, CultureInfo.InvariantCulture, out rows) ||
            rows < 1)
            return false;
        columns = 1;
        if (match.Groups[2].Success &&
            int.TryParse(match.Groups[2].Value, NumberStyles.None, CultureInfo.InvariantCulture, out var parsedColumns) &&
            parsedColumns >= 1)
            columns = parsedColumns;
        return true;
    }

    public static string? FirstFormula(JsonObject op)
    {
        var formulas = Json.GetArr(op, "formulas");
        return formulas is { Count: > 0 } && formulas[0] is JsonArray row && row.Count > 0
            ? row[0]?.ToString()
            : null;
    }

    public static bool IsOpenDynamicArrayFormula(string? formula) =>
        !string.IsNullOrWhiteSpace(formula) && OpenDynamicArrayPattern.IsMatch(formula.Trim());

    public static bool HasExplicitSpillHint(JsonObject op) =>
        Json.GetInt(op, "spillRows") is > 0 && Json.GetInt(op, "spillColumns") is > 0;

    /// <summary>
    /// SEQUENCE with constant size, or an explicit spillRows/Columns hint, has a known
    /// footprint. FILTER/SORT/UNIQUE do not; callers must not require those hints.
    /// </summary>
    public static bool RejectsOversizedActualSpill(JsonObject op) =>
        HasExplicitSpillHint(op) || TryEstimateSequenceFootprint(FirstFormula(op), out _, out _);

    public static (int Rows, int Columns) ResolveCaptureFootprint(JsonObject op, int destRows, int destColumns)
    {
        if (HasExplicitSpillHint(op))
        {
            var spillRows = Json.GetInt(op, "spillRows")!.Value;
            var spillColumns = Json.GetInt(op, "spillColumns")!.Value;
            return (Math.Max(destRows, spillRows), Math.Max(destColumns, spillColumns));
        }

        if (TryEstimateSequenceFootprint(FirstFormula(op), out var rows, out var columns))
            return (Math.Max(destRows, rows), Math.Max(destColumns, columns));

        return (destRows, destColumns);
    }

    public static string CleanA1(string? address) =>
        string.IsNullOrWhiteSpace(address) ? "" : address.Replace("$", "", StringComparison.Ordinal);

    public static bool SpillAnchoredAtDest(string destAddress, string? spillRange)
    {
        if (string.IsNullOrWhiteSpace(spillRange))
            return false;
        if (!TryParseA1Rect(destAddress, out var destRow, out var destCol, out _, out _))
            return false;
        if (!TryParseA1Rect(spillRange, out var spillRow, out var spillCol, out _, out _))
            return false;
        return destRow == spillRow && destCol == spillCol;
    }

    public static bool AcceptsActualSpill(
        JsonObject op, string destAddress, string? spillRange, int capturedRows, int capturedColumns)
    {
        if (string.IsNullOrWhiteSpace(spillRange))
            return false;
        if (!SpillAnchoredAtDest(destAddress, spillRange))
            return false;
        if (!RejectsOversizedActualSpill(op))
            return true;
        return SpillFitsCapture(destAddress, spillRange, capturedRows, capturedColumns);
    }

    public static bool SpillFitsCapture(string destAddress, string? spillRange, int capturedRows, int capturedColumns)
    {
        if (string.IsNullOrWhiteSpace(spillRange))
            return true;
        if (!SpillAnchoredAtDest(destAddress, spillRange))
            return false;
        if (!TryParseA1Size(destAddress, out var destRows, out var destCols) ||
            !TryParseA1Size(spillRange, out var spillRows, out var spillCols))
            return false;
        return spillRows <= capturedRows && spillCols <= capturedColumns &&
               destRows <= capturedRows && destCols <= capturedColumns;
    }

    /// <summary>
    /// Spill cells that were empty (or newly written) outside the captured dest.
    /// Occupied neighbors block the spill in Excel, so they are never owned here.
    /// </summary>
    public static string ResizeA1(string origin, int rows, int columns)
    {
        if (rows < 1 || columns < 1 || !TryParseA1Rect(origin, out var row, out var column, out _, out _))
            return CleanA1(origin);
        return ToA1Range(row, column, row + rows - 1, column + columns - 1);
    }

    public static IReadOnlyList<string> NewlyOwnedSpillRanges(string destAddress, string spillRange)
    {
        if (!TryParseA1Rect(destAddress, out var destRow, out var destCol, out var destRows, out var destCols))
            return [];
        if (!TryParseA1Rect(spillRange, out var spillRow, out var spillCol, out var spillRows, out var spillCols))
            return [];
        if (destRow != spillRow || destCol != spillCol)
            return [];
        var ranges = new List<string>();
        if (spillRows > destRows)
            ranges.Add(ToA1Range(spillRow + destRows, spillCol, spillRow + spillRows - 1, spillCol + spillCols - 1));
        if (spillCols > destCols)
            ranges.Add(ToA1Range(spillRow, spillCol + destCols, spillRow + destRows - 1, spillCol + spillCols - 1));
        return ranges;
    }

    public static bool TryParseA1Size(string address, out int rows, out int columns)
    {
        rows = 0;
        columns = 0;
        return TryParseA1Rect(address, out _, out _, out rows, out columns);
    }

    public static bool TryParseA1Rect(string address, out int row, out int column, out int rows, out int columns)
    {
        row = 0;
        column = 0;
        rows = 0;
        columns = 0;
        var local = CleanA1(address);
        var bang = local.LastIndexOf('!');
        if (bang >= 0)
            local = local[(bang + 1)..];
        var parts = local.Split(':');
        if (!TryParseA1Cell(parts[0], out var r1, out var c1))
            return false;
        if (parts.Length == 1)
        {
            row = r1;
            column = c1;
            rows = 1;
            columns = 1;
            return true;
        }

        if (!TryParseA1Cell(parts[1], out var r2, out var c2))
            return false;
        row = Math.Min(r1, r2);
        column = Math.Min(c1, c2);
        rows = Math.Abs(r2 - r1) + 1;
        columns = Math.Abs(c2 - c1) + 1;
        return true;
    }

    public static bool TryParseA1Cell(string cell, out int row, out int column)
    {
        row = 0;
        column = 0;
        cell = CleanA1(cell);
        var i = 0;
        while (i < cell.Length && char.IsLetter(cell[i]))
        {
            column = column * 26 + (char.ToUpperInvariant(cell[i]) - 'A' + 1);
            i++;
        }

        if (i == 0 || i >= cell.Length)
            return false;
        return int.TryParse(cell[i..], NumberStyles.None, CultureInfo.InvariantCulture, out row) && row > 0 && column > 0;
    }

    private static string ToA1Range(int r1, int c1, int r2, int c2)
    {
        var start = ColName(c1) + r1.ToString(CultureInfo.InvariantCulture);
        if (r1 == r2 && c1 == c2)
            return start;
        return start + ":" + ColName(c2) + r2.ToString(CultureInfo.InvariantCulture);
    }

    private static string ColName(int column)
    {
        var chars = new Stack<char>();
        var c = column;
        while (c > 0)
        {
            var m = (c - 1) % 26;
            chars.Push((char)('A' + m));
            c = (c - 1) / 26;
        }

        return new string(chars.ToArray());
    }

    public static bool ReadbacksMatch(JsonObject expected, JsonObject actual)
    {
        if (!string.Equals(Json.GetString(expected, "engine"), Json.GetString(actual, "engine"), StringComparison.Ordinal))
            return false;
        if (!string.Equals(Json.GetString(expected, "address"), Json.GetString(actual, "address"), StringComparison.OrdinalIgnoreCase))
            return false;
        var expectedSpill = expected.ContainsKey("hasSpill") && expected["hasSpill"] is JsonValue ev && ev.TryGetValue<bool>(out var eb)
            ? eb
            : (bool?)null;
        var actualSpill = actual.ContainsKey("hasSpill") && actual["hasSpill"] is JsonValue av && av.TryGetValue<bool>(out var ab)
            ? ab
            : (bool?)null;
        if (expectedSpill != actualSpill)
            return false;
        return string.Equals(
            Json.GetString(expected, "spillRange") ?? "",
            Json.GetString(actual, "spillRange") ?? "",
            StringComparison.OrdinalIgnoreCase);
    }

    public static readonly JsonObject Sequence32Fixture = Json.ParseObject("""
        { "op": "set_formulas", "engine": "formula2", "range": "B2",
          "formulas": [["=SEQUENCE(3,2)"]], "target": { "sheet": "계산" } }
        """)!;

    public static readonly JsonObject FilterFixture = Json.ParseObject("""
        { "op": "set_formulas", "engine": "formula2", "range": "E2",
          "formulas": [["=FILTER(A2:B10,LEN(A2:A10)>0)"]], "target": { "sheet": "계산" } }
        """)!;

    public static readonly JsonObject SortFixture = Json.ParseObject("""
        { "op": "set_formulas", "engine": "formula2", "range": "G2",
          "formulas": [["=SORT(A2:B10,1,1)"]], "target": { "sheet": "계산" } }
        """)!;

    public static readonly JsonObject UniqueFixture = Json.ParseObject("""
        { "op": "set_formulas", "engine": "formula2", "range": "I2",
          "formulas": [["=UNIQUE(A2:A10)"]], "target": { "sheet": "계산" } }
        """)!;

    public static readonly JsonObject Sequence22Fixture = Json.ParseObject("""
        { "op": "set_formulas", "engine": "formula2", "range": "B2",
          "formulas": [["=SEQUENCE(2,2)"]], "target": { "sheet": "계산" } }
        """)!;

    public static readonly JsonArray Sequence22Values = (JsonArray)JsonNode.Parse("[[1,2],[3,4]]")!;
    public static readonly JsonArray Sequence33Values = (JsonArray)JsonNode.Parse("[[1,2,3],[4,5,6],[7,8,9]]")!;
    public static readonly JsonArray FilterSpillValues = (JsonArray)JsonNode.Parse("""[["kept",1],["kept",2]]""")!;
}
