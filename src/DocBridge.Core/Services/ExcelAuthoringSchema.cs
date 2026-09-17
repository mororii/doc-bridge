using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// Public Excel authoring input schema shared by MCP, capabilities, and validator tests.
/// Layout/lifecycle/range-sheet parameters live here so op names are never published alone.
/// </summary>
public static class ExcelAuthoringSchema
{
    public const string ClosePolicyOwnedOnly = "owned-only";
    public const string OpenSourceRole = "open-source";
    public const string AddOwnershipSource = "Workbooks.Add";
    public const string OpenOwnershipSource = "Workbooks.Open";

    public static readonly string[] InspectReadScopes = ExcelDataOperationsContract.ReadScopes.ToArray();

    public static JsonArray InspectScopeEnum()
    {
        var scopes = new JsonArray("scan", "objects", "errors", "formula_trace", "diagnostics");
        foreach (var scope in InspectReadScopes)
            scopes.Add(scope);
        return scopes;
    }

    public static JsonObject DescribeInspectObjectKind() => new()
    {
        ["type"] = "string",
        ["enum"] = new JsonArray(InspectReadScopes.Select(scope => (JsonNode)JsonValue.Create(scope)!).ToArray()),
        ["description"] =
            "preferred when scope=objects. One of tables|charts|pictures|names|validations|" +
            "conditionalFormats|notes|hyperlinks|filters|pivots|shapes. " +
            "scope may also be that same read scope. objectKind and scope are normalized together.",
    };

    public static JsonObject DescribeApplyOpProperties() => new()
    {
        ["rows"] = new JsonObject
        {
            ["type"] = "array",
            ["description"] = "set_row_heights: [{row, count?, heightPoints?, autoFit?}]. heightPoints is Excel points 0.1..409.5.",
            ["items"] = DescribeRowHeightItem(),
        },
        ["columns"] = new JsonObject
        {
            ["type"] = "array",
            ["description"] = "set_column_widths: [{col, count?, widthChars?, autoFit?}]. widthChars is 0..255.",
            ["items"] = DescribeColumnWidthItem(),
        },
        ["page"] = new JsonObject
        {
            ["type"] = "object",
            ["description"] =
                "set_page_setup. Excel example: {paperSize:\"A3\",orientation:\"landscape\",scale:55," +
                "printArea:\"A1:BQ96\",leftMarginMm:10,headerMarginMm:8}. scale wins over fitToWidth/fitToHeight. Margins are millimetres.",
            ["properties"] = new JsonObject
            {
                ["paperSize"] = new JsonObject { ["description"] = "A3|A4|letter|legal|tabloid|B4|B5 or Excel paper integer" },
                ["orientation"] = new JsonObject { ["type"] = "string", ["enum"] = new JsonArray("portrait", "landscape") },
                ["scale"] = new JsonObject { ["type"] = "integer", ["minimum"] = 10, ["maximum"] = 400 },
                ["fitToWidth"] = new JsonObject { ["type"] = "integer", ["minimum"] = 0, ["maximum"] = 32767 },
                ["fitToHeight"] = new JsonObject { ["type"] = "integer", ["minimum"] = 0, ["maximum"] = 32767 },
                ["printArea"] = new JsonObject { ["type"] = "string" },
                ["printTitleRows"] = new JsonObject { ["type"] = "string" },
                ["printTitleColumns"] = new JsonObject { ["type"] = "string" },
                ["leftMarginMm"] = new JsonObject { ["type"] = "number" },
                ["rightMarginMm"] = new JsonObject { ["type"] = "number" },
                ["topMarginMm"] = new JsonObject { ["type"] = "number" },
                ["bottomMarginMm"] = new JsonObject { ["type"] = "number" },
                ["headerMarginMm"] = new JsonObject { ["type"] = "number" },
                ["footerMarginMm"] = new JsonObject { ["type"] = "number" },
                ["leftHeader"] = new JsonObject { ["type"] = "string" },
                ["centerHeader"] = new JsonObject { ["type"] = "string" },
                ["rightHeader"] = new JsonObject { ["type"] = "string" },
                ["leftFooter"] = new JsonObject { ["type"] = "string" },
                ["centerFooter"] = new JsonObject { ["type"] = "string" },
                ["rightFooter"] = new JsonObject { ["type"] = "string" },
            },
        },
        ["cell"] = new JsonObject
        {
            ["type"] = "string",
            ["description"] = "freeze_panes top-left cell, e.g. G6 → xSplit=6, ySplit=5, topLeftCell=G6.",
        },
        ["unfreeze"] = new JsonObject { ["type"] = "boolean" },
        ["output"] = new JsonObject
        {
            ["type"] = "string",
            ["description"] = "save_workbook/export_pdf absolute output path. Existing files need overwrite:true unless saving the current workbook onto itself.",
        },
        ["path"] = new JsonObject
        {
            ["type"] = "string",
            ["description"] = "open_workbook absolute existing path. Does not launch Excel.",
        },
        ["overwrite"] = new JsonObject { ["type"] = "boolean" },
        ["sheetName"] = new JsonObject { ["type"] = "string", ["description"] = "create_workbook first-sheet name." },
        ["newName"] = new JsonObject { ["type"] = "string", ["description"] = "rename_sheet new worksheet name." },
        ["name"] = new JsonObject { ["type"] = "string", ["description"] = "add_sheet worksheet name." },
        ["position"] = new JsonObject
        {
            ["type"] = "string",
            ["enum"] = new JsonArray("before", "after", "first", "last"),
            ["description"] = "move_sheet position.",
        },
        ["destRange"] = new JsonObject { ["type"] = "string", ["description"] = "copy_range destination A1 range." },
        ["destSheet"] = new JsonObject { ["type"] = "string" },
        ["what"] = new JsonObject
        {
            ["type"] = "string",
            ["enum"] = new JsonArray("all", "contents", "formats", "formulas"),
        },
        ["mode"] = new JsonObject
        {
            ["type"] = "string",
            ["enum"] = new JsonArray("all", "values", "formulas", "formats"),
        },
        ["sheet"] = new JsonObject { ["type"] = "string", ["description"] = "export_pdf optional single-sheet name." },
        ["saveChanges"] = new JsonObject { ["type"] = "boolean", ["description"] = "close_workbook. Default false." },
        ["color"] = new JsonObject { ["description"] = "set_tab_color OLE 0..16777215 or #RRGGBB." },
        ["value"] = new JsonObject { ["description"] = "fill_range scalar when values is omitted." },
        ["destination"] = new JsonObject { ["type"] = "string", ["description"] = "import_csv destination A1 cell." },
        ["delimiter"] = new JsonObject { ["type"] = "string" },
        ["formula2"] = new JsonObject
        {
            ["type"] = "boolean",
            ["description"] =
                "set_formulas: opt in to Range.Formula2. calculate.formula2 is rejected. " +
                "Prefer engine/formulaEngine.",
        },
        ["engine"] = new JsonObject
        {
            ["type"] = "string",
            ["enum"] = new JsonArray("formula", "formula2"),
            ["description"] =
                "set_formulas dialect. Default formula (legacy Range.Formula, no spill). " +
                "formula2 is Range.Formula2 and can spill.",
        },
        ["formulaEngine"] = new JsonObject
        {
            ["type"] = "string",
            ["enum"] = new JsonArray("formula", "formula2"),
            ["description"] = "Alias of engine.",
        },
        ["dialect"] = new JsonObject
        {
            ["type"] = "string",
            ["enum"] = new JsonArray("formula", "formula2"),
        },
        ["spillRows"] = new JsonObject { ["type"] = "integer", ["minimum"] = 1, ["description"] = "Formula2 snapshot footprint rows when SEQUENCE cannot be parsed." },
        ["spillColumns"] = new JsonObject { ["type"] = "integer", ["minimum"] = 1 },
        ["axis"] = new JsonObject
        {
            ["type"] = "string",
            ["enum"] = new JsonArray("row", "rows", "column", "columns", "col"),
            ["description"] = "set_outline axis. Default rows. columns uses EntireColumn/Columns.Group.",
        },
        ["orientation"] = new JsonObject { ["type"] = "string", ["enum"] = new JsonArray("row", "column") },
        ["clear"] = new JsonObject { ["type"] = "boolean", ["description"] = "set_page_breaks: reset all breaks." },
        ["level"] = new JsonObject { ["type"] = "integer", ["minimum"] = 1, ["maximum"] = 8 },
        ["summaryBelow"] = new JsonObject { ["type"] = "boolean" },
        ["summaryRight"] = new JsonObject { ["type"] = "boolean" },
        ["show"] = new JsonObject { ["type"] = "boolean", ["description"] = "set_outline false collapses detail; it does not ClearOutline." },
        ["zoom"] = new JsonObject { ["type"] = "integer", ["minimum"] = 10, ["maximum"] = 400, ["description"] = "set_view window zoom percent." },
        ["displayGridlines"] = new JsonObject { ["type"] = "boolean" },
        ["displayHeadings"] = new JsonObject { ["type"] = "boolean" },
        ["displayZeros"] = new JsonObject { ["type"] = "boolean" },
        ["view"] = new JsonObject
        {
            ["type"] = "string",
            ["enum"] = new JsonArray("normal", "pageLayout", "pageBreakPreview"),
            ["description"] = "set_view Excel window view.",
        },
        ["options"] = new JsonObject
        {
            ["type"] = "object",
            ["description"] =
                "protect_sheet practical flags. Password is rejected. Example: " +
                "{allowFormattingCells:true,allowSorting:true,allowFiltering:true,userInterfaceOnly:true}.",
            ["properties"] = DescribeProtectionOptions(),
        },
        ["allowFormattingCells"] = new JsonObject { ["type"] = "boolean" },
        ["allowFormattingColumns"] = new JsonObject { ["type"] = "boolean" },
        ["allowFormattingRows"] = new JsonObject { ["type"] = "boolean" },
        ["allowInsertingColumns"] = new JsonObject { ["type"] = "boolean" },
        ["allowInsertingRows"] = new JsonObject { ["type"] = "boolean" },
        ["allowInsertingHyperlinks"] = new JsonObject { ["type"] = "boolean" },
        ["allowDeletingColumns"] = new JsonObject { ["type"] = "boolean" },
        ["allowDeletingRows"] = new JsonObject { ["type"] = "boolean" },
        ["allowSorting"] = new JsonObject { ["type"] = "boolean" },
        ["allowFiltering"] = new JsonObject { ["type"] = "boolean" },
        ["allowUsingPivotTables"] = new JsonObject { ["type"] = "boolean" },
        ["drawingObjects"] = new JsonObject { ["type"] = "boolean" },
        ["contents"] = new JsonObject { ["type"] = "boolean" },
        ["scenarios"] = new JsonObject { ["type"] = "boolean" },
        ["userInterfaceOnly"] = new JsonObject { ["type"] = "boolean" },
    };

    public static JsonObject DescribeRowHeightItem() => new()
    {
        ["type"] = "object",
        ["additionalProperties"] = false,
        ["required"] = new JsonArray("row"),
        ["properties"] = new JsonObject
        {
            ["row"] = new JsonObject { ["type"] = "integer", ["minimum"] = 1, ["maximum"] = 1048576 },
            ["count"] = new JsonObject { ["type"] = "integer", ["minimum"] = 1 },
            ["heightPoints"] = new JsonObject { ["type"] = "number", ["minimum"] = 0.1, ["maximum"] = 409.5 },
            ["autoFit"] = new JsonObject { ["type"] = "boolean" },
        },
    };

    public static JsonObject DescribeColumnWidthItem() => new()
    {
        ["type"] = "object",
        ["additionalProperties"] = false,
        ["required"] = new JsonArray("col"),
        ["properties"] = new JsonObject
        {
            ["col"] = new JsonObject { ["description"] = "column letter or 1-based index" },
            ["count"] = new JsonObject { ["type"] = "integer", ["minimum"] = 1 },
            ["widthChars"] = new JsonObject { ["type"] = "number", ["minimum"] = 0, ["maximum"] = 255 },
            ["autoFit"] = new JsonObject { ["type"] = "boolean" },
        },
    };

    public static JsonObject DescribeProtectionOptions() => new()
    {
        ["drawingObjects"] = new JsonObject { ["type"] = "boolean" },
        ["contents"] = new JsonObject { ["type"] = "boolean" },
        ["scenarios"] = new JsonObject { ["type"] = "boolean" },
        ["userInterfaceOnly"] = new JsonObject { ["type"] = "boolean" },
        ["allowFormattingCells"] = new JsonObject { ["type"] = "boolean" },
        ["allowFormattingColumns"] = new JsonObject { ["type"] = "boolean" },
        ["allowFormattingRows"] = new JsonObject { ["type"] = "boolean" },
        ["allowInsertingColumns"] = new JsonObject { ["type"] = "boolean" },
        ["allowInsertingRows"] = new JsonObject { ["type"] = "boolean" },
        ["allowInsertingHyperlinks"] = new JsonObject { ["type"] = "boolean" },
        ["allowDeletingColumns"] = new JsonObject { ["type"] = "boolean" },
        ["allowDeletingRows"] = new JsonObject { ["type"] = "boolean" },
        ["allowSorting"] = new JsonObject { ["type"] = "boolean" },
        ["allowFiltering"] = new JsonObject { ["type"] = "boolean" },
        ["allowUsingPivotTables"] = new JsonObject { ["type"] = "boolean" },
    };

    public static readonly string[] ProtectionOptionKeys =
    {
        "drawingObjects", "contents", "scenarios", "userInterfaceOnly",
        "allowFormattingCells", "allowFormattingColumns", "allowFormattingRows",
        "allowInsertingColumns", "allowInsertingRows", "allowInsertingHyperlinks",
        "allowDeletingColumns", "allowDeletingRows", "allowSorting", "allowFiltering",
        "allowUsingPivotTables",
    };

    public static JsonObject NormalizeProtectionOptions(JsonObject op)
    {
        var merged = new JsonObject();
        var nested = Json.GetObj(op, "options") ?? Json.GetObj(op, "protection");
        foreach (var key in ProtectionOptionKeys)
        {
            if (nested is not null && nested.ContainsKey(key))
                merged[key] = nested[key]!.DeepClone();
            else if (op.ContainsKey(key))
                merged[key] = op[key]!.DeepClone();
        }

        return merged;
    }

    public static bool OptionalBool(JsonObject options, string key, bool fallback) =>
        options.ContainsKey(key) ? Json.GetBool(options, key) : fallback;

    public static bool SaveConflictsWithExistingFile(
        string? output, string? currentFullName, bool overwrite, bool targetExists)
    {
        if (overwrite || !targetExists || string.IsNullOrWhiteSpace(output))
            return false;
        return !string.Equals(output, currentFullName, StringComparison.OrdinalIgnoreCase);
    }

    public static string? ReadExplicitWorkbookTarget(JsonObject op) =>
        Json.GetString(op, "targetWorkbook") ?? Json.GetString(Json.GetObj(op, "target"), "workbook");

    public static bool IsSamePathSave(string? output, string? currentFullName)
    {
        if (string.IsNullOrWhiteSpace(output) || string.IsNullOrWhiteSpace(currentFullName))
            return false;
        if (ExcelAuthoringPaths.IsLikelyUnsavedWorkbookName(currentFullName))
            return false;
        try
        {
            return string.Equals(
                Path.GetFullPath(output),
                Path.GetFullPath(currentFullName),
                StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return string.Equals(output, currentFullName, StringComparison.OrdinalIgnoreCase);
        }
    }

    public static string? ResolveSaveCapturePath(string? output, string? currentFullName)
    {
        if (!string.IsNullOrWhiteSpace(output))
            return output;
        if (string.IsNullOrWhiteSpace(currentFullName) ||
            ExcelAuthoringPaths.IsLikelyUnsavedWorkbookName(currentFullName))
            return null;
        return currentFullName;
    }

    public static bool IsWritableLifecycleOutput(JsonObject output)
    {
        if (string.Equals(Json.GetString(output, "role"), OpenSourceRole, StringComparison.OrdinalIgnoreCase))
            return false;
        if (output.ContainsKey("writable") && !Json.GetBool(output, "writable"))
            return false;
        if (string.Equals(Json.GetString(output, "op"), "open_workbook", StringComparison.OrdinalIgnoreCase))
            return false;
        return !string.IsNullOrWhiteSpace(Json.GetString(output, "path"));
    }

    public static void AppendOwnedWorkbook(JsonObject state, string? name, string? fullName, string? source = null)
    {
        if (Json.GetArr(state, "ownedWorkbooks") is not JsonArray owned)
        {
            owned = new JsonArray();
            state["ownedWorkbooks"] = owned;
        }

        if (IdentityListed(owned, name, fullName))
            return;
        var item = new JsonObject
        {
            ["name"] = name,
            ["fullName"] = fullName,
        };
        if (!string.IsNullOrWhiteSpace(source))
            item["source"] = source;
        owned.Add(item);
        state["closePolicy"] = ClosePolicyOwnedOnly;
    }

    public static JsonObject DescribeLaunchOwnership(bool ownsInstance, bool createdWorkbook, bool dedicatedRequested) =>
        new()
        {
            ["ownedInstance"] = ownsInstance,
            ["createdInstance"] = ownsInstance,
            ["attachedExisting"] = !ownsInstance,
            ["createdWorkbook"] = createdWorkbook,
            ["dedicatedInstance"] = dedicatedRequested,
        };

    public static bool IdentityListed(JsonArray? list, string? name, string? fullName)
    {
        if (list is null) return false;
        foreach (var node in list)
        {
            if (node is not JsonObject item) continue;
            var listedName = Json.GetString(item, "name");
            var listedFull = Json.GetString(item, "fullName");
            if (!string.IsNullOrWhiteSpace(fullName) &&
                string.Equals(listedFull, fullName, StringComparison.OrdinalIgnoreCase))
                return true;
            if (!string.IsNullOrWhiteSpace(name) &&
                string.Equals(listedName, name, StringComparison.OrdinalIgnoreCase) &&
                (string.IsNullOrWhiteSpace(listedFull) || string.IsNullOrWhiteSpace(fullName) ||
                 string.Equals(listedFull, fullName, StringComparison.OrdinalIgnoreCase)))
                return true;
        }

        return false;
    }

    /// <summary>
    /// Close only workbooks this batch created or opened. Never invert a preexisting list.
    /// </summary>
    public static bool ShouldCloseOnLifecycleRestore(string? name, string? fullName, JsonArray? ownedWorkbooks)
    {
        if (ownedWorkbooks is null || ownedWorkbooks.Count == 0)
            return false;
        return IdentityListed(ownedWorkbooks, name, fullName);
    }

    public static bool ShouldRestoreDeleteEntry(JsonObject entry)
    {
        if (!entry.ContainsKey("applied"))
            return true;
        return Json.GetBool(entry, "applied");
    }

    public static bool MarkDeleteEntryApplied(JsonObject state, int index)
    {
        var entries = Json.GetArr(state, "entries");
        if (entries is null || index < 0 || index >= entries.Count || entries[index] is not JsonObject entry)
            return false;
        entry["applied"] = true;
        return true;
    }

    public static JsonObject NormalizeInspectArgs(JsonObject args)
    {
        var normalized = args.DeepClone() as JsonObject ?? new JsonObject();
        var objectKind = Json.GetString(normalized, "objectKind")
                         ?? Json.GetString(normalized, "objectScope")
                         ?? Json.GetString(normalized, "kind");
        var scope = Json.GetString(normalized, "scope") ?? "objects";
        if (ExcelDataOperationsContract.IsReadScope(scope) && string.IsNullOrWhiteSpace(objectKind))
        {
            objectKind = scope;
            scope = "objects";
        }

        if (!string.IsNullOrWhiteSpace(objectKind))
            normalized["objectKind"] = objectKind;
        normalized["scope"] = scope;
        return normalized;
    }

    public static bool IsDataObjectInspect(string? scope, string? objectKind) =>
        string.Equals(scope, "objects", StringComparison.OrdinalIgnoreCase) ||
        ExcelDataOperationsContract.IsReadScope(scope) ||
        ExcelDataOperationsContract.IsReadScope(objectKind);
}
