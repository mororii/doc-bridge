namespace DocBridge.Development.ExcelProductionWorkflowProbe;

internal enum OpFamily
{
    Lifecycle,
    Values,
    Format,
    Merge,
    SheetLayout,
    Rename,
    Structure,
    Delete,
    RangeEdit,
    Protect,
    Visibility,
    Data,
    Unknown,
}

internal sealed record OpName(string Canonical, string[] Aliases, OpFamily Family, bool AutoExecute = false, bool HighRisk = false);

internal static class ContractCatalog
{
    public static readonly IReadOnlyList<OpName> Phase1 =
    [
        new("create_workbook", [], OpFamily.Lifecycle),
        new("open_workbook", [], OpFamily.Lifecycle),
        new("close_workbook", [], OpFamily.Lifecycle),
        new("save_workbook", [], OpFamily.Lifecycle),
        new("export_pdf", [], OpFamily.Lifecycle, HighRisk: true),
        new("set_values", [], OpFamily.Values, AutoExecute: true),
        new("set_formulas", [], OpFamily.Values, AutoExecute: true),
        new("calculate", [], OpFamily.Values),
        new("format_range", [], OpFamily.Format, AutoExecute: true),
        new("merge_cells", [], OpFamily.Merge),
        new("unmerge_cells", [], OpFamily.Merge),
        new("set_row_heights", [], OpFamily.SheetLayout),
        new("set_column_widths", [], OpFamily.SheetLayout),
        new("freeze_panes", [], OpFamily.SheetLayout),
        new("set_page_setup", [], OpFamily.SheetLayout),
        new("set_view", [], OpFamily.SheetLayout),
        new("set_page_breaks", [], OpFamily.RangeEdit),
        new("set_outline", [], OpFamily.RangeEdit),
        new("rename_sheet", [], OpFamily.Rename),
        new("add_sheet", [], OpFamily.Structure),
        new("move_sheet", [], OpFamily.Structure),
        new("copy_sheet", [], OpFamily.Structure),
        new("set_tab_color", [], OpFamily.Structure),
        new("insert_rows", [], OpFamily.Structure),
        new("insert_cols", [], OpFamily.Structure),
        new("delete_rows", [], OpFamily.Delete, HighRisk: true),
        new("delete_cols", [], OpFamily.Delete, HighRisk: true),
        new("delete_sheet", [], OpFamily.Delete, HighRisk: true),
        new("clear_range", [], OpFamily.RangeEdit),
        new("copy_range", [], OpFamily.RangeEdit),
        new("find_replace", [], OpFamily.RangeEdit),
        new("fill_range", [], OpFamily.RangeEdit),
        new("auto_fill", [], OpFamily.RangeEdit),
        new("protect_sheet", [], OpFamily.Protect),
        new("unprotect_sheet", [], OpFamily.Protect),
        // Product visibility-only batches admit only these three names.
        new("set_rows_hidden", [], OpFamily.Visibility),
        new("set_cols_hidden", [], OpFamily.Visibility),
        new("set_sheet_visibility", [], OpFamily.Visibility),
    ];

    public static readonly IReadOnlySet<string> VisibilityOnlyOps =
        new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "set_rows_hidden", "set_cols_hidden", "set_sheet_visibility",
        };

    public static readonly IReadOnlyList<OpName> DataOps =
    [
        new("create_table", [], OpFamily.Data),
        new("resize_table", [], OpFamily.Data),
        new("style_table", ["set_table_style"], OpFamily.Data),
        new("set_table_totals", [], OpFamily.Data),
        new("add_table_column", [], OpFamily.Data),
        new("delete_table", [], OpFamily.Data),
        new("sort_range", [], OpFamily.Data),
        new("sort_table", [], OpFamily.Data),
        new("set_auto_filter", [], OpFamily.Data),
        new("clear_auto_filter", [], OpFamily.Data),
        new("define_name", [], OpFamily.Data),
        new("update_name", [], OpFamily.Data),
        new("delete_name", [], OpFamily.Data),
        new("set_data_validation", ["set_validation"], OpFamily.Data),
        new("clear_data_validation", ["clear_validation"], OpFamily.Data),
        new("add_conditional_format", ["add_conditional_format", "add_conditional_formats"], OpFamily.Data),
        new("clear_conditional_formats", ["clear_conditional_format"], OpFamily.Data),
        new("create_chart", ["add_chart"], OpFamily.Data),
        new("update_chart", ["set_chart"], OpFamily.Data),
        new("delete_chart", [], OpFamily.Data),
        new("insert_sheet_picture", ["insert_picture"], OpFamily.Data),
        new("update_picture", ["set_picture"], OpFamily.Data),
        new("delete_picture", [], OpFamily.Data),
        new("set_cell_note", ["set_comment"], OpFamily.Data),
        new("clear_cell_note", ["clear_comment"], OpFamily.Data),
        new("set_hyperlink", [], OpFamily.Data),
        new("clear_hyperlink", [], OpFamily.Data),
        new("insert_shape", [], OpFamily.Data),
        new("update_shape", [], OpFamily.Data),
        new("delete_shape", [], OpFamily.Data),
        new("insert_textbox", [], OpFamily.Data),
        new("update_textbox", [], OpFamily.Data),
        new("delete_textbox", [], OpFamily.Data),
    ];

    public static readonly IReadOnlyList<OpName> Extended =
    [
        new("create_pivot", ["add_pivot"], OpFamily.Data),
        new("update_pivot", [], OpFamily.Data),
        new("refresh_pivot", [], OpFamily.Data),
        new("delete_pivot", [], OpFamily.Data),
        new("export_csv", [], OpFamily.Lifecycle),
        new("import_csv", ["open_csv"], OpFamily.Lifecycle),
        new("remove_duplicates", [], OpFamily.Data),
        new("text_to_columns", [], OpFamily.Data),
    ];

    public static IEnumerable<OpName> All => Phase1.Concat(DataOps).Concat(Extended);

    public static readonly IReadOnlyList<string> RootAdvertisedOps =
    [
        "find_replace", "copy_sheet", "set_tab_color", "insert_sheet_picture", "calculate",
        "set_rows_hidden", "set_page_breaks", "delete_chart", "resize_table", "update_shape",
        "set_page_setup", "style_table", "export_pdf", "create_chart", "freeze_panes",
        "delete_name", "delete_pivot", "clear_hyperlink", "remove_duplicates", "delete_cols",
        "unprotect_sheet", "set_cell_note", "move_sheet", "format_range", "unmerge_cells",
        "clear_range", "insert_textbox", "update_picture", "insert_cols", "set_auto_filter",
        "copy_range", "open_workbook", "delete_rows", "set_row_heights", "rename_sheet",
        "refresh_pivot", "insert_shape", "set_formulas", "sort_table", "clear_auto_filter",
        "update_name", "insert_rows", "clear_conditional_formats", "close_workbook",
        "import_csv", "set_hyperlink", "create_table", "sort_range", "add_conditional_format",
        "set_column_widths", "text_to_columns", "set_cols_hidden", "set_sheet_visibility",
        "set_view", "clear_cell_note", "delete_textbox", "update_pivot", "merge_cells",
        "create_workbook", "delete_shape", "delete_table", "create_pivot", "protect_sheet",
        "save_workbook", "set_outline", "export_csv", "add_sheet", "set_table_totals",
        "auto_fill", "set_data_validation", "clear_data_validation", "update_chart",
        "fill_range", "delete_picture", "define_name", "add_table_column", "set_values",
        "update_textbox", "delete_sheet",
    ];

    public static OpFamily FamilyOf(string op)
    {
        foreach (var n in Phase1.Concat(DataOps).Concat(Extended))
        {
            if (Matches(n, op)) return n.Family;
        }
        return OpFamily.Unknown;
    }

    public static bool Matches(OpName name, string op) =>
        string.Equals(name.Canonical, op, StringComparison.OrdinalIgnoreCase) ||
        name.Aliases.Any(a => string.Equals(a, op, StringComparison.OrdinalIgnoreCase));

    public static string Resolve(string canonical, IReadOnlySet<string> advertised)
    {
        var all = Phase1.Concat(DataOps).Concat(Extended);
        var spec = all.FirstOrDefault(n => string.Equals(n.Canonical, canonical, StringComparison.OrdinalIgnoreCase));
        if (spec is null) return canonical;
        if (advertised.Contains(spec.Canonical)) return spec.Canonical;
        foreach (var alias in spec.Aliases)
            if (advertised.Contains(alias)) return alias;
        return spec.Canonical;
    }

    public static IReadOnlySet<string> AdvertisedWriteOps(JsonNode? capabilities)
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var write = JsonUtil.Get(capabilities, "writeOps") as JsonArray;
        if (write is null)
        {
            var apps = JsonUtil.Get(capabilities, "apps") as JsonObject
                       ?? JsonUtil.Get(JsonUtil.Get(capabilities, "result"), "apps") as JsonObject;
            var excel = apps is null ? null : JsonUtil.Get(apps, "excel");
            write = JsonUtil.Get(excel, "writeOps") as JsonArray
                    ?? JsonUtil.Get(capabilities, "excel") as JsonArray;
        }
        if (write is not null)
        {
            foreach (var n in write)
                if (n is JsonValue v && v.TryGetValue<string>(out var s))
                    set.Add(s);
        }
        return set;
    }

    public static JsonObject Describe() => new()
    {
        ["phase1"] = new JsonArray(Phase1.Select(DescribeOne).ToArray()),
        ["data"] = new JsonArray(DataOps.Select(DescribeOne).ToArray()),
        ["extended"] = new JsonArray(Extended.Select(DescribeOne).ToArray()),
        ["executeAllowlist"] = JsonUtil.Arr("set_values", "set_formulas", "format_range"),
        ["sources"] = new JsonObject
        {
            ["phase1"] = "output/docbridge-production-20260910/cursor/EXCEL-PHASE1-API-CONTRACT.md",
            ["data"] = "output/docbridge-production-20260910/cursor-data/OP-SCHEMA.md",
        },
    };

    private static JsonObject DescribeOne(OpName n) => new()
    {
        ["canonical"] = n.Canonical,
        ["aliases"] = new JsonArray(n.Aliases.Select(a => JsonValue.Create(a)).ToArray()),
        ["family"] = n.Family.ToString(),
        ["autoExecute"] = n.AutoExecute,
        ["highRisk"] = n.HighRisk,
    };
}
