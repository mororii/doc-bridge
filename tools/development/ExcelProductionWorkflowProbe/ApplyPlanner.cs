namespace DocBridge.Development.ExcelProductionWorkflowProbe;

internal sealed class PlannedBatch
{
    public required string Id { get; init; }
    public required string Scenario { get; init; }
    public required OpFamily Family { get; init; }
    public required string Path { get; init; } // execute | token
    public required JsonArray Ops { get; init; }
    public string? Note { get; init; }
    public string? Capture { get; init; }
    public bool HighRisk { get; init; }
    public bool InjectedFailure { get; init; }

    public JsonObject ToJson() => new()
    {
        ["id"] = Id,
        ["scenario"] = Scenario,
        ["family"] = Family.ToString(),
        ["path"] = Path,
        ["highRisk"] = HighRisk,
        ["injectedFailure"] = InjectedFailure,
        ["capture"] = Capture,
        ["opCount"] = Ops.Count,
        ["note"] = Note,
        ["ops"] = Ops.DeepClone(),
    };
}

internal static class ApplyPlanner
{
    public static PlannedBatch Batch(
        string scenario,
        string id,
        OpFamily family,
        JsonArray ops,
        string path,
        string? note = null,
        bool injectedFailure = false,
        string? capture = null)
    {
        ValidateFamily(family, ops);
        ValidateObjectLifecycle(id, ops);
        if (path == "execute")
        {
            foreach (var op in ops.OfType<JsonObject>())
            {
                var name = JsonUtil.Str(op, "op") ?? "";
                if (name is not ("set_values" or "set_formulas" or "format_range"))
                    throw new InvalidOperationException($"{id}: execute path is only for set_values/set_formulas/format_range");
            }
        }

        return new PlannedBatch
        {
            Id = id,
            Scenario = scenario,
            Family = family,
            Path = path,
            Ops = ops,
            Note = note,
            Capture = capture,
            HighRisk = ops.OfType<JsonObject>().Any(o =>
            {
                var n = JsonUtil.Str(o, "op") ?? "";
                return n is "export_pdf" or "delete_rows" or "delete_cols" or "delete_sheet";
            }),
            InjectedFailure = injectedFailure,
        };
    }

    public static JsonArray Merges(string sheet, IEnumerable<string> ranges)
    {
        var ops = new JsonArray();
        foreach (var range in ranges)
        {
            ops.Add(new JsonObject
            {
                ["op"] = "merge_cells",
                ["target"] = JsonUtil.Target(sheet),
                ["range"] = range,
            });
        }
        return ops;
    }

    public static void ValidateFamily(OpFamily family, JsonArray ops)
    {
        foreach (var op in ops.OfType<JsonObject>())
        {
            var name = JsonUtil.Str(op, "op") ?? "";
            var got = ContractCatalog.FamilyOf(name);
            if (got != family && !(family == OpFamily.Values && got == OpFamily.Format))
                throw new InvalidOperationException($"op {name} family {got} cannot sit in {family} batch");
        }

        if (family == OpFamily.Merge)
        {
            var kinds = ops.OfType<JsonObject>().Select(o => JsonUtil.Str(o, "op")).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            if (kinds.Count > 1)
                throw new InvalidOperationException("do not mix merge_cells with unmerge_cells");
        }

        if (family == OpFamily.Lifecycle)
        {
            var names = ops.OfType<JsonObject>().Select(o => JsonUtil.Str(o, "op") ?? "").ToList();
            if (names.Any(n => n is "create_workbook" or "open_workbook" or "close_workbook") && names.Count != 1)
                throw new InvalidOperationException("create_workbook/open_workbook/close_workbook must be alone");
            if (names.Any(n => ContractCatalog.FamilyOf(n) != OpFamily.Lifecycle))
                throw new InvalidOperationException("lifecycle batch has a non-lifecycle op");
        }

        if (family == OpFamily.Structure && ops.OfType<JsonObject>().Any(o => JsonUtil.Str(o, "op") == "copy_sheet") && ops.Count != 1)
            throw new InvalidOperationException("copy_sheet remains alone");

        if (family == OpFamily.Visibility)
        {
            foreach (var name in ops.OfType<JsonObject>().Select(o => JsonUtil.Str(o, "op") ?? ""))
            {
                if (!ContractCatalog.VisibilityOnlyOps.Contains(name))
                    throw new InvalidOperationException(
                        $"visibility-only batch cannot include '{name}'; set_view is SheetLayout and set_outline is RangeEdit");
            }
        }
    }

    public static void ValidateObjectLifecycle(string id, JsonArray ops)
    {
        var created = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var mutated = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var op in ops.OfType<JsonObject>())
        {
            var name = JsonUtil.Str(op, "op") ?? "";
            var key = ObjectKey(op);
            if (key is null) continue;
            if (IsCreateOp(name)) created[key] = name;
            if (IsMutateOrDeleteOp(name)) mutated[key] = name;
        }

        foreach (var item in created)
        {
            if (mutated.TryGetValue(item.Key, out var later))
            {
                throw new InvalidOperationException(
                    $"{id}: public preview resolves {later} against the current workbook; " +
                    $"do not {item.Value} then {later} '{item.Key}' in one batch");
            }
        }
    }

    private static bool IsCreateOp(string name) => name is
        "insert_sheet_picture" or "insert_shape" or "insert_textbox" or "create_table"
        or "add_table_column" or "define_name" or "create_chart" or "set_hyperlink"
        or "create_pivot" or "set_cell_note" or "add_conditional_format"
        or "set_data_validation" or "set_auto_filter";

    private static bool IsMutateOrDeleteOp(string name) => name is
        "update_picture" or "delete_picture" or "update_shape" or "delete_shape"
        or "update_textbox" or "delete_textbox" or "style_table" or "set_table_totals"
        or "resize_table" or "sort_table" or "delete_table" or "update_name" or "delete_name"
        or "update_chart" or "delete_chart" or "clear_hyperlink" or "update_pivot"
        or "refresh_pivot" or "delete_pivot" or "clear_cell_note"
        or "clear_conditional_formats" or "clear_data_validation" or "clear_auto_filter";

    private static string? ObjectKey(JsonObject op)
    {
        var name = JsonUtil.Str(op, "op") ?? "";
        var objectName = JsonUtil.Str(op, "name");
        var range = JsonUtil.Str(op, "range") ?? JsonUtil.Str(op, "cell");
        return name switch
        {
            "insert_sheet_picture" or "update_picture" or "delete_picture" => "picture:" + objectName,
            "insert_shape" or "update_shape" or "delete_shape" => "shape:" + objectName,
            "insert_textbox" or "update_textbox" or "delete_textbox" => "textbox:" + objectName,
            "create_table" or "style_table" or "set_table_totals" or "resize_table" or "sort_table"
                or "add_table_column" or "delete_table" or "set_auto_filter" or "clear_auto_filter"
                => "table:" + objectName,
            "define_name" or "update_name" or "delete_name" => "name:" + objectName,
            "create_chart" or "update_chart" or "delete_chart" => "chart:" + objectName,
            "create_pivot" or "update_pivot" or "refresh_pivot" or "delete_pivot" => "pivot:" + objectName,
            "set_hyperlink" or "clear_hyperlink" => "link:" + range,
            "set_cell_note" or "clear_cell_note" => "note:" + (JsonUtil.Str(op, "cell") ?? range),
            "add_conditional_format" or "clear_conditional_formats" => "cf:" + range,
            "set_data_validation" or "clear_data_validation" => "dv:" + range,
            _ => null,
        };
    }
}
