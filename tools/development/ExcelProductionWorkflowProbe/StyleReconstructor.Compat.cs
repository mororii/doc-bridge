namespace DocBridge.Development.ExcelProductionWorkflowProbe;

/// <summary>
/// Compiled only when <c>SkipE1Planner=true</c>. Default builds compile the real
/// StyleReconstructor.cs so data can test the planner in an isolated output.
/// </summary>
internal static class StyleReconstructor
{
    public static JsonObject Reconstruct(string xlsxPath, string sheetName, int maxRow = 96, int maxCol = 69) =>
        throw new InvalidOperationException("SkipE1Planner=true; real StyleReconstructor.cs was not compiled.");

    public static IEnumerable<PlannedBatch> FormatBatches(string scenario, string sheet, JsonArray formatOps, int chunk = 15) =>
        Array.Empty<PlannedBatch>();

    internal static JsonObject PlannerSelfCheck() => new()
    {
        ["ok"] = false,
        ["errors"] = JsonUtil.Arr("SkipE1Planner=true; data ctx_07e3b0188992 owns StyleReconstructor.cs"),
    };

    internal static IReadOnlyList<string> CoalesceRectangles(IEnumerable<(int Row, int Col)> cells)
    {
        var set = cells.ToHashSet();
        var used = new HashSet<(int Row, int Col)>();
        var ranges = new List<string>();
        foreach (var start in set.OrderBy(c => c.Row).ThenBy(c => c.Col))
        {
            if (!used.Add(start)) continue;
            var maxCol = start.Col;
            while (set.Contains((start.Row, maxCol + 1)) && !used.Contains((start.Row, maxCol + 1)))
                maxCol++;
            var maxRow = start.Row;
            var grow = true;
            while (grow)
            {
                var next = maxRow + 1;
                for (var c = start.Col; c <= maxCol; c++)
                {
                    if (!set.Contains((next, c)) || used.Contains((next, c)))
                    {
                        grow = false;
                        break;
                    }
                }
                if (grow) maxRow = next;
            }
            for (var r = start.Row; r <= maxRow; r++)
                for (var c = start.Col; c <= maxCol; c++)
                    used.Add((r, c));
            ranges.Add(A1.Range(start.Row, start.Col, maxRow, maxCol));
        }
        return ranges;
    }
}
