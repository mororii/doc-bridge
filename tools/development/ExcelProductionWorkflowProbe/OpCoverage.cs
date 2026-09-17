namespace DocBridge.Development.ExcelProductionWorkflowProbe;

internal static class OpCoverage
{
    public static JsonObject FromPlans(IReadOnlyDictionary<string, List<PlannedBatch>> plans)
    {
        var planned = new Dictionary<string, JsonArray>(StringComparer.OrdinalIgnoreCase);
        foreach (var (scenario, batches) in plans)
        {
            foreach (var batch in batches)
            {
                foreach (var op in batch.Ops.OfType<JsonObject>())
                {
                    var name = JsonUtil.Str(op, "op");
                    if (string.IsNullOrWhiteSpace(name)) continue;
                    if (!planned.TryGetValue(name, out var arr))
                    {
                        arr = new JsonArray();
                        planned[name] = arr;
                    }
                    arr.Add($"{scenario}/{batch.Id}");
                }
            }
        }

        var advertised = ContractCatalog.RootAdvertisedOps;
        var withoutPlan = advertised.Where(op => !planned.ContainsKey(op)).OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToList();
        var plannedList = planned.Keys.OrderBy(x => x, StringComparer.OrdinalIgnoreCase).ToList();

        return new JsonObject
        {
            ["officeCalls"] = 0,
            ["liveComRun"] = false,
            ["stagedFixturesIntegrated"] = true,
            ["note"] =
                "Offline planned-op inventory only. A planned batch is not apply+readback, " +
                "persisted/visual/recovery acceptance, or every parameter. Practical E3/E6/E8/E9 " +
                "fixtures are hooked; RemainingOp spliced before save. E2 resume IDs stay except " +
                "the G12:H14 gap-fix batches after layout-page. e2-sig-gap-relabel writes only G12:H12. " +
                "Candidate 0842 lacks empty-string merge fix.",
            ["advertisedOps"] = advertised.Count,
            ["plannedDistinctOps"] = plannedList.Count,
            ["advertisedWithoutPlan"] = new JsonArray(withoutPlan.Select(o => JsonValue.Create(o)).ToArray()),
            ["advertisedWithoutPlanCount"] = withoutPlan.Count,
            ["plannedOps"] = new JsonArray(plannedList.Select(o => JsonValue.Create(o)).ToArray()),
            ["plannedByOp"] = new JsonObject(planned.OrderBy(p => p.Key, StringComparer.OrdinalIgnoreCase)
                .Select(p => new KeyValuePair<string, JsonNode?>(p.Key, p.Value.DeepClone()))),
        };
    }
}
