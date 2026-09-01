using System.Globalization;
using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

/// <summary>
/// Excel의 기본 배치 편집(병합/병합해제, 행·열·시트 숨김/표시)과
/// operation-scoped 복구 상태를 구현한다. 모든 COM 객체는 획득 단계별로
/// 균형 해제해 ActiveWorkbook/Worksheet RCW 별칭을 분리하지 않는다.
/// </summary>
public sealed partial class ExcelAdapter
{
    private const int XlSheetVisible = -1;
    private const int XlSheetHidden = 0;
    private const int XlSheetVeryHidden = 2;
    private const int ExcelLayoutSnapshotVersion = 1;
    // Merge snapshots preserve per-cell formatting (including borders), which is
    // COM-call intensive. Keep the operation bounded to a practical rollback time.
    private const int MaxMergeOperationCells = 2_000;
    private const string VisibilityRestoreMode = "visibility-state";
    private const string MergeRestoreMode = "merge-state";

    private static readonly HashSet<string> VisibilityOperationNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "set_rows_hidden", "set_cols_hidden", "set_sheet_visibility",
    };

    private static readonly HashSet<string> MergeOperationNames = new(StringComparer.OrdinalIgnoreCase)
    {
        "merge_cells", "unmerge_cells",
    };

    public override JsonObject GetCapabilities() => new()
    {
        ["app"] = App,
        ["automation"] = "excel-activex-com",
        ["directAppControl"] = true,
        ["connectsToExistingWindow"] = true,
        ["usesUiAutomation"] = false,
        ["usesExternalMacro"] = false,
        ["interactionPolicy"] = new JsonObject
        {
            ["mode"] = "preserve-foreground",
            ["backgroundInactiveWindow"] = true,
            ["restoresOriginalDocument"] = true,
            ["concurrentTargetInput"] = "stop-after-current-operation",
            ["sameDocumentConcurrentEditing"] = false,
        },
        ["readOps"] = new JsonArray("context", "range", "scan", "objects", "errors", "diagnostics", "layout"),
        ["writeOps"] = new JsonArray(
            "set_values", "set_formulas", "insert_rows", "insert_cols", "format_range",
            "find_replace", "copy_sheet", "merge_cells", "unmerge_cells",
            "set_rows_hidden", "set_cols_hidden", "set_sheet_visibility"),
        ["limits"] = new JsonObject
        {
            ["maxReadCells"] = MaxCells,
            ["maxMergeCells"] = MaxMergeOperationCells,
            ["maxSnapshotCells"] = MaxSnapshotCells,
            ["maxDiffEntries"] = MaxDiff,
            ["maxRows"] = 1_048_576,
            ["maxColumns"] = 16_384,
        },
        ["safety"] = new JsonArray(
            "dry-run", "snapshot", "confirm-token", "readback", "automatic-rollback",
            "merge-content-loss-block", "last-visible-sheet-block", "active-sheet-hide-block"),
    };

    private static bool IsVisibilityOnlySnapshot(IReadOnlyList<JsonObject>? ops) =>
        ops is { Count: > 0 } &&
        ops.All(op => VisibilityOperationNames.Contains(Json.GetString(op, "op") ?? ""));

    private static bool IsMergeOnlySnapshot(IReadOnlyList<JsonObject>? ops) =>
        ops is { Count: 1 } && MergeOperationNames.Contains(Json.GetString(ops[0], "op") ?? "");

    private static int ParseColumnNumber(JsonNode? node)
    {
        if (node is JsonValue value && value.TryGetValue<int>(out var number))
        {
            if (number is >= 1 and <= 16_384) return number;
            throw new ArgumentOutOfRangeException(nameof(node), "Excel column number must be 1..16384");
        }

        if (node is not JsonValue stringValue || !stringValue.TryGetValue<string>(out var text) ||
            string.IsNullOrWhiteSpace(text))
            throw new ArgumentException("Excel column must be a number or A..XFD");
        var column = ColIndex(text.Trim());
        if (column is < 1 or > 16_384 || !string.Equals(ColName(column), text.Trim(), StringComparison.OrdinalIgnoreCase))
            throw new ArgumentOutOfRangeException(nameof(node), "Excel column must be A..XFD");
        return column;
    }

    private static object GetExplicitTargetSheetReference(object workbook, JsonObject op)
    {
        var sheetName = Json.GetString(Json.GetObj(op, "target"), "sheet");
        if (string.IsNullOrWhiteSpace(sheetName))
            throw new InvalidOperationException(
                $"Excel write op '{Json.GetString(op, "op")}' requires target.sheet; the active sheet is never assumed for writes");

        object? worksheets = null;
        try
        {
            worksheets = (object)((dynamic)workbook).Worksheets;
            try { return (object)((dynamic)worksheets).Item(sheetName); }
            catch { throw new InvalidOperationException($"sheet '{sheetName}' not found"); }
        }
        finally { RotHelper.ReleaseComReference(worksheets); }
    }

    private static (object Sheet, object Range, string Address) ResolveBasicRange(object workbook, JsonObject op)
    {
        var resolved = ResolveRangeTarget(
            workbook,
            Json.GetString(Json.GetObj(op, "target"), "sheet"),
            Json.GetString(op, "range")!,
            requireExplicitSheet: true);
        object? range = null;
        try
        {
            range = (object)((dynamic)resolved.Sheet).Range(resolved.Address);
            var canonical = Convert.ToString(((dynamic)range).Address(false, false), CultureInfo.InvariantCulture)
                            ?? resolved.Address;
            return (resolved.Sheet, range, canonical);
        }
        catch
        {
            RotHelper.ReleaseComReference(range);
            RotHelper.ReleaseComReference(resolved.Sheet);
            throw;
        }
    }

    private static JsonArray CaptureHiddenStates(object sheet, bool rows, int start, int count)
    {
        object? collection = null;
        var result = new JsonArray();
        try
        {
            collection = rows ? (object)((dynamic)sheet).Rows : (object)((dynamic)sheet).Columns;
            for (var offset = 0; offset < count; offset++)
            {
                object? item = null;
                try
                {
                    item = (object)((dynamic)collection).Item(start + offset);
                    result.Add(Convert.ToBoolean(((dynamic)item).Hidden, CultureInfo.InvariantCulture));
                }
                finally { RotHelper.ReleaseComReference(item); }
            }
        }
        finally { RotHelper.ReleaseComReference(collection); }
        return result;
    }

    private static bool HiddenStatesMatch(JsonArray states, bool expected)
    {
        foreach (var state in states)
            if (state is not JsonValue value || !value.TryGetValue<bool>(out var actual) || actual != expected)
                return false;
        return true;
    }

    private static JsonNode HiddenStateSummary(JsonArray states)
    {
        if (states.Count == 0) return JsonValue.Create("empty");
        var first = states[0]!.GetValue<bool>();
        return states.All(state => state!.GetValue<bool>() == first)
            ? JsonValue.Create(first)
            : JsonValue.Create("mixed");
    }

    private static string SheetVisibilityName(int value) => value switch
    {
        XlSheetVisible => "visible",
        XlSheetHidden => "hidden",
        XlSheetVeryHidden => "veryHidden",
        _ => $"unknown({value})",
    };

    private static int CountVisibleSheets(object workbook)
    {
        object? sheets = null;
        try
        {
            sheets = (object)((dynamic)workbook).Sheets;
            var count = Convert.ToInt32(((dynamic)sheets).Count, CultureInfo.InvariantCulture);
            var visible = 0;
            for (var index = 1; index <= count; index++)
            {
                object? sheet = null;
                try
                {
                    sheet = (object)((dynamic)sheets).Item(index);
                    if (Convert.ToInt32(((dynamic)sheet).Visible, CultureInfo.InvariantCulture) == XlSheetVisible)
                        visible++;
                }
                finally { RotHelper.ReleaseComReference(sheet); }
            }
            return visible;
        }
        finally { RotHelper.ReleaseComReference(sheets); }
    }

    private static string ReadActiveSheetName(object workbook)
    {
        object? sheet = null;
        try
        {
            sheet = (object)((dynamic)workbook).ActiveSheet;
            return Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture) ?? "";
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static void ValidateSheetVisibilityChange(object workbook, object sheet, int desired)
    {
        var current = Convert.ToInt32(((dynamic)sheet).Visible, CultureInfo.InvariantCulture);
        if (current == desired) return;
        if (Convert.ToBoolean(((dynamic)workbook).ProtectStructure, CultureInfo.InvariantCulture))
            throw new InvalidOperationException("[EXCEL_WORKBOOK_STRUCTURE_PROTECTED] workbook structure is protected");
        if (desired != XlSheetHidden) return;

        var sheetName = Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture) ?? "";
        if (string.Equals(sheetName, ReadActiveSheetName(workbook), StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException(
                "[EXCEL_ACTIVE_SHEET_HIDE_BLOCKED] the active worksheet cannot be hidden; activate another sheet and retry");
        if (CountVisibleSheets(workbook) <= 1)
            throw new InvalidOperationException(
                "[EXCEL_LAST_VISIBLE_SHEET] at least one workbook sheet must remain visible");
    }

    private static List<string> ValidateVisibilityBatch(object workbook, IReadOnlyList<JsonObject> ops)
    {
        var errors = new List<string>();
        var states = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        object? sheets = null;
        try
        {
            sheets = (object)((dynamic)workbook).Sheets;
            var count = Convert.ToInt32(((dynamic)sheets).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                object? sheet = null;
                try
                {
                    sheet = (object)((dynamic)sheets).Item(index);
                    states[Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture) ?? ""] =
                        Convert.ToInt32(((dynamic)sheet).Visible, CultureInfo.InvariantCulture);
                }
                finally { RotHelper.ReleaseComReference(sheet); }
            }
        }
        finally { RotHelper.ReleaseComReference(sheets); }

        var active = ReadActiveSheetName(workbook);
        var protectedStructure = Convert.ToBoolean(((dynamic)workbook).ProtectStructure, CultureInfo.InvariantCulture);
        foreach (var op in ops.Where(op => Json.GetString(op, "op") == "set_sheet_visibility"))
        {
            var sheetName = Json.GetString(Json.GetObj(op, "target"), "sheet") ?? "";
            if (!states.TryGetValue(sheetName, out var current))
            {
                errors.Add($"sheet '{sheetName}' not found");
                continue;
            }
            var desired = string.Equals(Json.GetString(op, "visibility"), "visible", StringComparison.OrdinalIgnoreCase)
                ? XlSheetVisible
                : XlSheetHidden;
            if (current == desired) continue;
            if (protectedStructure)
            {
                errors.Add("[EXCEL_WORKBOOK_STRUCTURE_PROTECTED] workbook structure is protected");
                continue;
            }
            if (desired == XlSheetHidden)
            {
                if (string.Equals(sheetName, active, StringComparison.OrdinalIgnoreCase))
                {
                    errors.Add("[EXCEL_ACTIVE_SHEET_HIDE_BLOCKED] the active worksheet cannot be hidden; activate another sheet and retry");
                    continue;
                }
                if (states.Values.Count(value => value == XlSheetVisible) <= 1)
                {
                    errors.Add("[EXCEL_LAST_VISIBLE_SHEET] at least one workbook sheet must remain visible");
                    continue;
                }
            }
            states[sheetName] = desired;
        }
        return errors.Distinct(StringComparer.Ordinal).ToList();
    }

    private sealed class VisibilityPreviewState
    {
        private readonly Dictionary<string, int> _sheetVisibility = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<(string Sheet, int Index), bool> _rowHidden = new();
        private readonly Dictionary<(string Sheet, int Index), bool> _columnHidden = new();

        public VisibilityPreviewState(object workbook)
        {
            object? sheets = null;
            try
            {
                sheets = (object)((dynamic)workbook).Sheets;
                var count = Convert.ToInt32(((dynamic)sheets).Count, CultureInfo.InvariantCulture);
                for (var index = 1; index <= count; index++)
                {
                    object? sheet = null;
                    try
                    {
                        sheet = (object)((dynamic)sheets).Item(index);
                        var name = Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture) ?? "";
                        _sheetVisibility[name] =
                            Convert.ToInt32(((dynamic)sheet).Visible, CultureInfo.InvariantCulture);
                    }
                    finally { RotHelper.ReleaseComReference(sheet); }
                }
            }
            finally { RotHelper.ReleaseComReference(sheets); }
        }

        public JsonArray ReadHidden(object sheet, string sheetName, bool rows, int start, int count)
        {
            var actual = CaptureHiddenStates(sheet, rows, start, count);
            var simulated = new JsonArray();
            var overrides = rows ? _rowHidden : _columnHidden;
            for (var offset = 0; offset < count; offset++)
            {
                var key = (sheetName, start + offset);
                simulated.Add(overrides.TryGetValue(key, out var hidden)
                    ? hidden
                    : actual[offset]!.GetValue<bool>());
            }
            return simulated;
        }

        public void SetHidden(string sheetName, bool rows, int start, int count, bool hidden)
        {
            var overrides = rows ? _rowHidden : _columnHidden;
            for (var offset = 0; offset < count; offset++) overrides[(sheetName, start + offset)] = hidden;
        }

        public int GetSheetVisibility(string sheetName) =>
            _sheetVisibility.TryGetValue(sheetName, out var visibility)
                ? visibility
                : throw new InvalidOperationException($"sheet '{sheetName}' not found");

        public void SetSheetVisibility(string sheetName, int visibility) =>
            _sheetVisibility[sheetName] = visibility;
    }

    private static void PreviewVisibilityOperation(object workbook, JsonObject op, ApplyPreview preview,
        VisibilityPreviewState state)
    {
        var name = Json.GetString(op, "op")!;
        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, op);
            var sheetName = Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture) ?? "";
            if (name == "set_rows_hidden")
            {
                var row = Json.GetInt(op, "row")!.Value;
                var count = Json.GetInt(op, "count")!.Value;
                var hidden = Json.GetBool(op, "hidden");
                var before = state.ReadHidden(sheet, sheetName, true, row, count);
                preview.Affected.Add(new AffectedRef("rows", $"{sheetName}!{row}:{row + count - 1}"));
                preview.Diff.Add(new DiffEntry { Ref = "hidden", Before = HiddenStateSummary(before), After = JsonValue.Create(hidden) });
                state.SetHidden(sheetName, true, row, count, hidden);
                return;
            }
            if (name == "set_cols_hidden")
            {
                var col = ParseColumnNumber(op["col"]);
                var count = Json.GetInt(op, "count")!.Value;
                var hidden = Json.GetBool(op, "hidden");
                var before = state.ReadHidden(sheet, sheetName, false, col, count);
                preview.Affected.Add(new AffectedRef("cols", $"{sheetName}!{ColName(col)}:{ColName(col + count - 1)}"));
                preview.Diff.Add(new DiffEntry { Ref = "hidden", Before = HiddenStateSummary(before), After = JsonValue.Create(hidden) });
                state.SetHidden(sheetName, false, col, count, hidden);
                return;
            }

            var desired = string.Equals(Json.GetString(op, "visibility"), "visible", StringComparison.OrdinalIgnoreCase)
                ? XlSheetVisible
                : XlSheetHidden;
            // ValidateVisibilityBatch has already checked the complete sequential sheet
            // transition. Reading the live workbook again here would make dry-run diffs
            // disagree with apply when two operations target the same sheet.
            var current = state.GetSheetVisibility(sheetName);
            preview.Affected.Add(new AffectedRef("sheet", sheetName));
            preview.Diff.Add(new DiffEntry
            {
                Ref = $"sheet:{sheetName}:visibility",
                Before = JsonValue.Create(SheetVisibilityName(current)),
                After = JsonValue.Create(SheetVisibilityName(desired)),
            });
            state.SetSheetVisibility(sheetName, desired);
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static void ApplyVisibilityOperation(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        var name = Json.GetString(op, "op")!;
        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, op);
            var sheetName = Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture) ?? "";
            if (name is "set_rows_hidden" or "set_cols_hidden")
            {
                var rows = name == "set_rows_hidden";
                var start = rows ? Json.GetInt(op, "row")!.Value : ParseColumnNumber(op["col"]);
                var count = Json.GetInt(op, "count")!.Value;
                var hidden = Json.GetBool(op, "hidden");
                object? collection = null;
                object? targetRange = null;
                try
                {
                    collection = rows ? (object)((dynamic)sheet).Rows : (object)((dynamic)sheet).Columns;
                    var reference = rows
                        ? $"{start}:{start + count - 1}"
                        : $"{ColName(start)}:{ColName(start + count - 1)}";
                    targetRange = (object)((dynamic)collection)[reference];
                    ((dynamic)targetRange).Hidden = hidden;
                }
                finally
                {
                    RotHelper.ReleaseComReference(targetRange);
                    RotHelper.ReleaseComReference(collection);
                }

                var actual = CaptureHiddenStates(sheet, rows, start, count);
                checkedItems += count;
                if (!HiddenStatesMatch(actual, hidden))
                    mismatches.Add($"{sheetName}!{(rows ? $"rows {start}:{start + count - 1}" : $"cols {ColName(start)}:{ColName(start + count - 1)}")}: hidden readback mismatch");
                execution.Affected.Add(new AffectedRef(rows ? "rows" : "cols",
                    rows
                        ? $"{sheetName}!{start}:{start + count - 1}"
                        : $"{sheetName}!{ColName(start)}:{ColName(start + count - 1)}"));
                return;
            }

            var desired = string.Equals(Json.GetString(op, "visibility"), "visible", StringComparison.OrdinalIgnoreCase)
                ? XlSheetVisible
                : XlSheetHidden;
            ValidateSheetVisibilityChange(workbook, sheet, desired);
            ((dynamic)sheet).Visible = desired;
            var actualVisibility = Convert.ToInt32(((dynamic)sheet).Visible, CultureInfo.InvariantCulture);
            checkedItems++;
            if (actualVisibility != desired)
                mismatches.Add($"sheet '{sheetName}' visibility: expected {SheetVisibilityName(desired)}, actual {SheetVisibilityName(actualVisibility)}");
            execution.Affected.Add(new AffectedRef("sheet", $"{sheetName} ({SheetVisibilityName(desired)})"));
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private sealed record MergeAreaInfo(string Address, int Row, int Column, int Rows, int Columns);

    private static List<MergeAreaInfo> EnumerateMergeAreas(object range, int maxCells = MaxMergeOperationCells)
    {
        var result = new Dictionary<string, MergeAreaInfo>(StringComparer.OrdinalIgnoreCase);
        object? cells = null;
        try
        {
            cells = (object)((dynamic)range).Cells;
            var count = Convert.ToInt64(((dynamic)cells).CountLarge, CultureInfo.InvariantCulture);
            if (count > maxCells)
                throw new InvalidOperationException($"merge operation is limited to {maxCells} cells (requested {count})");
            for (var index = 1; index <= (int)count; index++)
            {
                object? cell = null;
                object? area = null;
                object? areaRows = null;
                object? areaColumns = null;
                try
                {
                    cell = (object)((dynamic)cells).Item(index);
                    if (!Convert.ToBoolean(((dynamic)cell).MergeCells, CultureInfo.InvariantCulture)) continue;
                    area = (object)((dynamic)cell).MergeArea;
                    var address = Convert.ToString(((dynamic)area).Address(false, false), CultureInfo.InvariantCulture) ?? "";
                    if (result.ContainsKey(address)) continue;
                    areaRows = (object)((dynamic)area).Rows;
                    areaColumns = (object)((dynamic)area).Columns;
                    result[address] = new MergeAreaInfo(
                        address,
                        Convert.ToInt32(((dynamic)area).Row, CultureInfo.InvariantCulture),
                        Convert.ToInt32(((dynamic)area).Column, CultureInfo.InvariantCulture),
                        Convert.ToInt32(((dynamic)areaRows).Count, CultureInfo.InvariantCulture),
                        Convert.ToInt32(((dynamic)areaColumns).Count, CultureInfo.InvariantCulture));
                }
                finally
                {
                    RotHelper.ReleaseComReference(areaColumns);
                    RotHelper.ReleaseComReference(areaRows);
                    RotHelper.ReleaseComReference(area);
                    RotHelper.ReleaseComReference(cell);
                }
            }
        }
        finally { RotHelper.ReleaseComReference(cells); }
        return result.Values.OrderBy(area => area.Row).ThenBy(area => area.Column).ToList();
    }

    private static bool IsWithin(MergeAreaInfo area, int row, int column, int rows, int columns) =>
        area.Row >= row && area.Column >= column &&
        area.Row + area.Rows <= row + rows && area.Column + area.Columns <= column + columns;

    private static object? ReadFormulaOrValue(object cell)
    {
        try { return ((dynamic)cell).Formula2; }
        catch { return ((dynamic)cell).Formula; }
    }

    private static bool IsNonEmptyExcelValue(object? value) =>
        value is not null && (value is not string text || text.Length != 0);

    private static void EnsureMergeWillNotDeleteContent(object range)
    {
        object? cells = null;
        try
        {
            cells = (object)((dynamic)range).Cells;
            var count = Convert.ToInt64(((dynamic)cells).CountLarge, CultureInfo.InvariantCulture);
            if (count is < 2 or > MaxMergeOperationCells)
                throw new InvalidOperationException(
                    $"merge_cells requires 2..{MaxMergeOperationCells} cells (requested {count})");
            for (var index = 2; index <= (int)count; index++)
            {
                object? cell = null;
                try
                {
                    cell = (object)((dynamic)cells).Item(index);
                    var content = ReadFormulaOrValue(cell);
                    if (IsNonEmptyExcelValue(content))
                        throw new InvalidOperationException(
                            "[EXCEL_MERGE_WOULD_DELETE_CONTENT] merge was blocked because a non-upper-left cell contains a value or formula");
                }
                finally { RotHelper.ReleaseComReference(cell); }
            }
        }
        finally { RotHelper.ReleaseComReference(cells); }
    }

    private static bool IntersectsExcelTable(object range)
    {
        object? cells = null;
        try
        {
            cells = (object)((dynamic)range).Cells;
            var count = Convert.ToInt64(((dynamic)cells).CountLarge, CultureInfo.InvariantCulture);
            for (var index = 1; index <= (int)count; index++)
            {
                object? cell = null;
                object? listObject = null;
                try
                {
                    cell = (object)((dynamic)cells).Item(index);
                    try { listObject = (object?)((dynamic)cell).ListObject; }
                    catch { /* a normal cell outside a table has no ListObject */ }
                    if (listObject is not null) return true;
                }
                finally
                {
                    RotHelper.ReleaseComReference(listObject);
                    RotHelper.ReleaseComReference(cell);
                }
            }
            return false;
        }
        finally { RotHelper.ReleaseComReference(cells); }
    }

    private sealed record MergePlan(
        string Operation,
        string SheetName,
        string RequestedAddress,
        IReadOnlyList<MergeAreaInfo> BeforeAreas,
        IReadOnlyList<string> EffectiveAddresses,
        bool NoOp,
        JsonNode? AnchorContent);

    private static MergePlan AnalyzeMergeOperation(object workbook, JsonObject op)
    {
        var operation = Json.GetString(op, "op")!;
        var resolved = ResolveBasicRange(workbook, op);
        try
        {
            dynamic sheet = resolved.Sheet;
            dynamic range = resolved.Range;
            if (Convert.ToBoolean(sheet.ProtectContents, CultureInfo.InvariantCulture))
                throw new InvalidOperationException("[EXCEL_SHEET_PROTECTED] worksheet contents are protected");

            object? rangeAreas = null;
            try
            {
                rangeAreas = (object)range.Areas;
                if (Convert.ToInt32(((dynamic)rangeAreas).Count, CultureInfo.InvariantCulture) != 1)
                    throw new InvalidOperationException(
                        "[EXCEL_MERGE_NONCONTIGUOUS_RANGE] merge and unmerge require one contiguous rectangular range");
            }
            finally { RotHelper.ReleaseComReference(rangeAreas); }

            object? rangeRows = null;
            object? rangeColumns = null;
            int row;
            int column;
            int rows;
            int columns;
            try
            {
                row = Convert.ToInt32(range.Row, CultureInfo.InvariantCulture);
                column = Convert.ToInt32(range.Column, CultureInfo.InvariantCulture);
                rangeRows = (object)range.Rows;
                rangeColumns = (object)range.Columns;
                rows = Convert.ToInt32(((dynamic)rangeRows).Count, CultureInfo.InvariantCulture);
                columns = Convert.ToInt32(((dynamic)rangeColumns).Count, CultureInfo.InvariantCulture);
            }
            finally
            {
                RotHelper.ReleaseComReference(rangeColumns);
                RotHelper.ReleaseComReference(rangeRows);
            }
            var count = (long)rows * columns;
            if (count > MaxMergeOperationCells)
                throw new InvalidOperationException($"merge operation is limited to {MaxMergeOperationCells} cells (requested {count})");

            var areas = EnumerateMergeAreas(resolved.Range);
            object? anchor = null;
            JsonNode? anchorContent;
            try
            {
                object? cells = null;
                try
                {
                    cells = (object)range.Cells;
                    anchor = (object)((dynamic)cells).Item(1);
                }
                finally { RotHelper.ReleaseComReference(cells); }
                anchorContent = ToJsonValue(ReadFormulaOrValue(anchor));
            }
            finally { RotHelper.ReleaseComReference(anchor); }

            if (operation == "merge_cells")
            {
                if (count < 2)
                    throw new InvalidOperationException("merge_cells requires a range containing at least two cells");
                if (IntersectsExcelTable(resolved.Range))
                    throw new InvalidOperationException("[EXCEL_TABLE_MERGE_BLOCKED] cells inside an Excel table cannot be merged");
                if (areas.Count == 1 && string.Equals(areas[0].Address, resolved.Address, StringComparison.OrdinalIgnoreCase))
                    return new MergePlan(operation, Convert.ToString(sheet.Name, CultureInfo.InvariantCulture) ?? "",
                        resolved.Address, areas, new[] { resolved.Address }, true, anchorContent);
                if (areas.Count > 0)
                    throw new InvalidOperationException(
                        "[EXCEL_MERGE_OVERLAP] target range intersects an existing merged area; unmerge it in a separate batch first");
                EnsureMergeWillNotDeleteContent(resolved.Range);
                return new MergePlan(operation, Convert.ToString(sheet.Name, CultureInfo.InvariantCulture) ?? "",
                    resolved.Address, areas, new[] { resolved.Address }, false, anchorContent);
            }

            if (areas.Count == 0)
                return new MergePlan(operation, Convert.ToString(sheet.Name, CultureInfo.InvariantCulture) ?? "",
                    resolved.Address, areas, Array.Empty<string>(), true, anchorContent);
            var effectiveCellCount = areas.Sum(area => (long)area.Rows * area.Columns);
            if (effectiveCellCount > MaxMergeOperationCells)
                throw new InvalidOperationException(
                    $"merge operation is limited to {MaxMergeOperationCells} effective cells (requested merged areas contain {effectiveCellCount})");
            foreach (var area in areas)
                if (!IsWithin(area, row, column, rows, columns) && count != 1)
                    throw new InvalidOperationException(
                        "[EXCEL_UNMERGE_PARTIAL_OVERLAP] target must be a merged-cell anchor or fully contain every merged area");
            return new MergePlan(operation, Convert.ToString(sheet.Name, CultureInfo.InvariantCulture) ?? "",
                resolved.Address, areas, areas.Select(area => area.Address).ToList(), false, anchorContent);
        }
        finally
        {
            RotHelper.ReleaseComReference(resolved.Range);
            RotHelper.ReleaseComReference(resolved.Sheet);
        }
    }

    private static void PreviewMergeOperation(object workbook, JsonObject op, ApplyPreview preview)
    {
        var plan = AnalyzeMergeOperation(workbook, op);
        preview.Affected.Add(new AffectedRef("range", $"{plan.SheetName}!{plan.RequestedAddress}"));
        var before = plan.BeforeAreas.Count == 0
            ? "unmerged"
            : string.Join(", ", plan.BeforeAreas.Select(area => area.Address));
        var after = plan.Operation == "merge_cells"
            ? plan.RequestedAddress
            : "unmerged";
        preview.Diff.Add(new DiffEntry
        {
            Ref = $"{plan.SheetName}!{plan.RequestedAddress}:mergedAreas",
            Before = JsonValue.Create(before),
            After = JsonValue.Create(after),
        });
        if (plan.NoOp) preview.Warnings.Add($"{plan.Operation} is already satisfied for {plan.SheetName}!{plan.RequestedAddress}");
    }

    private static void ApplyMergeOperation(object workbook, JsonObject op, ApplyExecution execution,
        List<string> mismatches, ref int checkedItems)
    {
        var plan = AnalyzeMergeOperation(workbook, op);
        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, op);
            if (!plan.NoOp)
            {
                if (plan.Operation == "merge_cells")
                {
                    object? range = null;
                    try
                    {
                        range = (object)((dynamic)sheet).Range(plan.RequestedAddress);
                        ((dynamic)range).Merge(false);
                    }
                    finally { RotHelper.ReleaseComReference(range); }
                }
                else
                {
                    foreach (var address in plan.EffectiveAddresses)
                    {
                        object? area = null;
                        try
                        {
                            area = (object)((dynamic)sheet).Range(address);
                            ((dynamic)area).UnMerge();
                        }
                        finally { RotHelper.ReleaseComReference(area); }
                    }
                }
            }

            object? verifyRange = null;
            try
            {
                verifyRange = (object)((dynamic)sheet).Range(plan.RequestedAddress);
                var actualAreas = EnumerateMergeAreas(verifyRange);
                checkedItems += Math.Max(1, plan.EffectiveAddresses.Count);
                if (plan.Operation == "merge_cells")
                {
                    if (actualAreas.Count != 1 ||
                        !string.Equals(actualAreas[0].Address, plan.RequestedAddress, StringComparison.OrdinalIgnoreCase))
                        mismatches.Add($"{plan.SheetName}!{plan.RequestedAddress}: merged-area readback mismatch");
                }
                else if (actualAreas.Count != 0)
                {
                    mismatches.Add($"{plan.SheetName}!{plan.RequestedAddress}: expected unmerged cells, found {string.Join(", ", actualAreas.Select(area => area.Address))}");
                }

                object? cells = null;
                object? anchor = null;
                try
                {
                    cells = (object)((dynamic)verifyRange).Cells;
                    anchor = (object)((dynamic)cells).Item(1);
                    var actualAnchor = ToJsonValue(ReadFormulaOrValue(anchor));
                    checkedItems++;
                    if (!JsonNode.DeepEquals(plan.AnchorContent, actualAnchor))
                        mismatches.Add($"{plan.SheetName}!{plan.RequestedAddress}: upper-left value/formula changed during {plan.Operation}");
                }
                finally
                {
                    RotHelper.ReleaseComReference(anchor);
                    RotHelper.ReleaseComReference(cells);
                }
            }
            finally { RotHelper.ReleaseComReference(verifyRange); }
            execution.Affected.Add(new AffectedRef("range", $"{plan.SheetName}!{plan.RequestedAddress}"));
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static JsonObject CaptureVisibilityState(object workbook, IReadOnlyList<JsonObject> ops, string? documentRef)
    {
        var entries = new JsonArray();
        foreach (var op in ops)
        {
            object? sheet = null;
            try
            {
                sheet = GetExplicitTargetSheetReference(workbook, op);
                var name = Json.GetString(op, "op")!;
                var sheetName = Convert.ToString(((dynamic)sheet).Name, CultureInfo.InvariantCulture) ?? "";
                if (name == "set_rows_hidden")
                {
                    var row = Json.GetInt(op, "row")!.Value;
                    var count = Json.GetInt(op, "count")!.Value;
                    entries.Add(new JsonObject
                    {
                        ["kind"] = "rows", ["sheet"] = sheetName, ["start"] = row, ["count"] = count,
                        ["hiddenStates"] = CaptureHiddenStates(sheet, true, row, count),
                    });
                }
                else if (name == "set_cols_hidden")
                {
                    var col = ParseColumnNumber(op["col"]);
                    var count = Json.GetInt(op, "count")!.Value;
                    entries.Add(new JsonObject
                    {
                        ["kind"] = "cols", ["sheet"] = sheetName, ["start"] = col, ["count"] = count,
                        ["hiddenStates"] = CaptureHiddenStates(sheet, false, col, count),
                    });
                }
                else
                {
                    entries.Add(new JsonObject
                    {
                        ["kind"] = "sheet", ["sheet"] = sheetName,
                        ["visibility"] = Convert.ToInt32(((dynamic)sheet).Visible, CultureInfo.InvariantCulture),
                    });
                }
            }
            finally { RotHelper.ReleaseComReference(sheet); }
        }
        return new JsonObject
        {
            ["snapshotVersion"] = ExcelLayoutSnapshotVersion,
            ["restoreMode"] = VisibilityRestoreMode,
            ["documentRef"] = documentRef,
            ["originalActiveSheet"] = ReadActiveSheetName(workbook),
            ["entries"] = entries,
        };
    }

    private static JsonObject CaptureMergeStyleRange(object sheet, string address, ref long capturedCells)
    {
        object? range = null;
        object? rowsObject = null;
        object? columnsObject = null;
        object? cells = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            rowsObject = (object)((dynamic)range).Rows;
            columnsObject = (object)((dynamic)range).Columns;
            var rows = Convert.ToInt32(((dynamic)rowsObject).Count, CultureInfo.InvariantCulture);
            var columns = Convert.ToInt32(((dynamic)columnsObject).Count, CultureInfo.InvariantCulture);
            capturedCells += (long)rows * columns;
            if (capturedCells > MaxFormatSnapshotCells)
                throw new InvalidOperationException(
                    $"merge format snapshot exceeds {MaxFormatSnapshotCells} cells; write was blocked because cell formatting could not be restored safely");

            cells = (object)((dynamic)range).Cells;
            var styles = new JsonArray();
            for (var row = 1; row <= rows; row++)
            {
                var styleRow = new JsonArray();
                for (var column = 1; column <= columns; column++)
                {
                    object? cell = null;
                    try
                    {
                        cell = (object)((dynamic)cells).Item(row, column);
                        styleRow.Add(CaptureMergeCellStyle(cell));
                    }
                    finally { RotHelper.ReleaseComReference(cell); }
                }
                styles.Add(styleRow);
            }
            return new JsonObject { ["range"] = address, ["styles"] = styles };
        }
        finally
        {
            RotHelper.ReleaseComReference(cells);
            RotHelper.ReleaseComReference(columnsObject);
            RotHelper.ReleaseComReference(rowsObject);
            RotHelper.ReleaseComReference(range);
        }
    }

    private static int RestoreMergeStyleRanges(object sheet, string sheetName, JsonArray styleRanges,
        RestoreMismatchCollector mismatches)
    {
        var restoredCells = 0;
        foreach (var node in styleRanges)
        {
            if (node is not JsonObject styleRange) continue;
            var styleAddress = Json.GetString(styleRange, "range");
            var styleRows = Json.GetArr(styleRange, "styles");
            if (string.IsNullOrWhiteSpace(styleAddress) || styleRows is null)
            {
                mismatches.Add("merge style snapshot is missing range or styles");
                continue;
            }

            object? range = null;
            object? cells = null;
            try
            {
                range = (object)((dynamic)sheet).Range(styleAddress);
                cells = (object)((dynamic)range).Cells;
                for (var row = 0; row < styleRows.Count; row++)
                {
                    if (styleRows[row] is not JsonArray styleColumns) continue;
                    for (var column = 0; column < styleColumns.Count; column++)
                    {
                        if (styleColumns[column] is not JsonObject style) continue;
                        object? cell = null;
                        try
                        {
                            cell = (object)((dynamic)cells).Item(row + 1, column + 1);
                            RestoreMergeCellStyle(cell, style);
                            restoredCells++;
                            if (!MergeCellStyleMatches(cell, style))
                            {
                                var actualStyle = CaptureMergeCellStyle(cell);
                                var fields = style
                                    .Where(pair => !JsonNode.DeepEquals(pair.Value, actualStyle[pair.Key]))
                                    .Select(pair => pair.Key);
                                mismatches.Add(
                                    $"{sheetName}!{styleAddress}[{row + 1},{column + 1}]: style restore mismatch ({string.Join(", ", fields)})");
                            }
                        }
                        catch (Exception ex)
                        {
                            mismatches.Add(
                                $"{sheetName}!{styleAddress}[{row + 1},{column + 1}]: style restore failed: {ex.Message}");
                        }
                        finally { RotHelper.ReleaseComReference(cell); }
                    }
                }
            }
            finally
            {
                RotHelper.ReleaseComReference(cells);
                RotHelper.ReleaseComReference(range);
            }
        }
        return restoredCells;
    }

    private static JsonObject CaptureMergeState(object workbook, JsonObject op, string? documentRef)
    {
        var plan = AnalyzeMergeOperation(workbook, op);
        var areas = new JsonArray();
        foreach (var area in plan.BeforeAreas) areas.Add(area.Address);
        var styleRanges = new JsonArray();
        if (!plan.NoOp)
        {
            object? sheet = null;
            try
            {
                sheet = GetExplicitTargetSheetReference(workbook, op);
                var styleAddresses = plan.Operation == "merge_cells"
                    ? new[] { plan.RequestedAddress }
                    : plan.BeforeAreas.Select(area => area.Address).Distinct(StringComparer.OrdinalIgnoreCase);
                var capturedCells = 0L;
                foreach (var address in styleAddresses)
                    styleRanges.Add(CaptureMergeStyleRange(sheet, address, ref capturedCells));
            }
            finally { RotHelper.ReleaseComReference(sheet); }
        }
        return new JsonObject
        {
            ["snapshotVersion"] = ExcelLayoutSnapshotVersion,
            ["restoreMode"] = MergeRestoreMode,
            ["documentRef"] = documentRef,
            ["operation"] = plan.Operation,
            ["sheet"] = plan.SheetName,
            ["range"] = plan.RequestedAddress,
            ["beforeMergedAreas"] = areas,
            ["anchorContent"] = plan.AnchorContent?.DeepClone(),
            ["noOp"] = plan.NoOp,
            ["styleRanges"] = styleRanges,
        };
    }

    private static void SetIndividualHiddenStates(object sheet, bool rows, int start, JsonArray states,
        RestoreMismatchCollector mismatches)
    {
        object? collection = null;
        try
        {
            collection = rows ? (object)((dynamic)sheet).Rows : (object)((dynamic)sheet).Columns;
            for (var offset = 0; offset < states.Count; offset++)
            {
                var expected = states[offset]!.GetValue<bool>();
                object? item = null;
                try
                {
                    item = (object)((dynamic)collection).Item(start + offset);
                    ((dynamic)item).Hidden = expected;
                    var actual = Convert.ToBoolean(((dynamic)item).Hidden, CultureInfo.InvariantCulture);
                    if (actual != expected)
                        mismatches.Add($"{(rows ? "row" : "column")} {start + offset} hidden restore mismatch");
                }
                catch (Exception ex)
                {
                    mismatches.Add($"{(rows ? "row" : "column")} {start + offset} hidden restore failed: {ex.Message}");
                }
                finally { RotHelper.ReleaseComReference(item); }
            }
        }
        finally { RotHelper.ReleaseComReference(collection); }
    }

    private static JsonObject RestoreVisibilityState(object workbook, JsonObject state)
    {
        var mismatches = new RestoreMismatchCollector();
        var entries = Json.GetArr(state, "entries") ?? new JsonArray();
        var checkedItems = 0;

        // 표시 상태를 먼저 복구해야 뒤에서 원래 hidden/veryHidden 상태를 적용할 때
        // Excel의 "마지막 표시 시트" 제약에 걸리지 않는다.
        foreach (var node in entries)
        {
            if (node is not JsonObject entry || Json.GetString(entry, "kind") != "sheet" ||
                Json.GetInt(entry, "visibility") != XlSheetVisible) continue;
            object? sheet = null;
            try
            {
                sheet = GetExplicitTargetSheetReference(workbook, new JsonObject
                {
                    ["op"] = "restore_sheet_visibility",
                    ["target"] = new JsonObject { ["sheet"] = Json.GetString(entry, "sheet") },
                });
                ((dynamic)sheet).Visible = XlSheetVisible;
            }
            catch (Exception ex) { mismatches.Add($"sheet '{Json.GetString(entry, "sheet")}' show restore failed: {ex.Message}"); }
            finally { RotHelper.ReleaseComReference(sheet); }
        }

        for (var index = entries.Count - 1; index >= 0; index--)
        {
            if (entries[index] is not JsonObject entry) continue;
            object? sheet = null;
            try
            {
                sheet = GetExplicitTargetSheetReference(workbook, new JsonObject
                {
                    ["op"] = "restore_visibility",
                    ["target"] = new JsonObject { ["sheet"] = Json.GetString(entry, "sheet") },
                });
                var kind = Json.GetString(entry, "kind");
                if (kind is "rows" or "cols")
                {
                    var states = Json.GetArr(entry, "hiddenStates") ?? new JsonArray();
                    SetIndividualHiddenStates(sheet, kind == "rows", Json.GetInt(entry, "start")!.Value, states, mismatches);
                    checkedItems += states.Count;
                }
                else if (kind == "sheet")
                {
                    var expected = Json.GetInt(entry, "visibility")!.Value;
                    ((dynamic)sheet).Visible = expected;
                    var actual = Convert.ToInt32(((dynamic)sheet).Visible, CultureInfo.InvariantCulture);
                    checkedItems++;
                    if (actual != expected)
                        mismatches.Add($"sheet '{Json.GetString(entry, "sheet")}' visibility restore mismatch");
                }
            }
            catch (Exception ex) { mismatches.Add($"visibility restore failed: {ex.Message}"); }
            finally { RotHelper.ReleaseComReference(sheet); }
        }

        var activeName = Json.GetString(state, "originalActiveSheet");
        if (!string.IsNullOrWhiteSpace(activeName)) ActivateWorksheet(workbook, activeName, mismatches);
        return BuildRestoreResult(mismatches.Count == 0, 0, checkedItems, VisibilityRestoreMode, mismatches);
    }

    private static JsonObject RestoreMergeState(object workbook, JsonObject state)
    {
        var mismatches = new RestoreMismatchCollector();
        var operation = Json.GetString(state, "operation");
        var sheetName = Json.GetString(state, "sheet");
        var address = Json.GetString(state, "range");
        var before = Json.GetArr(state, "beforeMergedAreas") ?? new JsonArray();
        var styleRanges = Json.GetArr(state, "styleRanges") ?? new JsonArray();
        var noOp = Json.GetBool(state, "noOp");
        if (operation is not ("merge_cells" or "unmerge_cells") || string.IsNullOrWhiteSpace(sheetName) ||
            string.IsNullOrWhiteSpace(address))
        {
            mismatches.Add("merge snapshot is missing operation, sheet, or range");
            return BuildRestoreResult(false, 0, 0, MergeRestoreMode, mismatches);
        }

        object? sheet = null;
        try
        {
            sheet = GetExplicitTargetSheetReference(workbook, new JsonObject
            {
                ["op"] = "restore_merge",
                ["target"] = new JsonObject { ["sheet"] = sheetName },
            });
            object? requested = null;
            try
            {
                requested = (object)((dynamic)sheet).Range(address);
                var restoredStyleCells = 0;
                if (!noOp)
                {
                    var currentAreas = EnumerateMergeAreas(requested);
                    foreach (var current in currentAreas)
                    {
                        object? area = null;
                        try
                        {
                            area = (object)((dynamic)sheet).Range(current.Address);
                            ((dynamic)area).UnMerge();
                        }
                        finally { RotHelper.ReleaseComReference(area); }
                    }

                    // Merging normalizes non-anchor cell formatting. Restore the original
                    // per-cell styles while cells are unmerged, then recreate the original
                    // merge topology.
                    restoredStyleCells = RestoreMergeStyleRanges(sheet, sheetName, styleRanges, mismatches);

                    foreach (var node in before)
                    {
                        var originalAddress = node?.GetValue<string>();
                        if (string.IsNullOrWhiteSpace(originalAddress)) continue;
                        object? area = null;
                        try
                        {
                            area = (object)((dynamic)sheet).Range(originalAddress);
                            ((dynamic)area).Merge(false);
                        }
                        catch (Exception ex) { mismatches.Add($"merge area '{originalAddress}' restore failed: {ex.Message}"); }
                        finally { RotHelper.ReleaseComReference(area); }
                    }
                }

                var actual = EnumerateMergeAreas(requested).Select(item => item.Address).ToHashSet(StringComparer.OrdinalIgnoreCase);
                var expected = before.Select(node => node?.GetValue<string>() ?? "").Where(text => text.Length > 0)
                    .ToHashSet(StringComparer.OrdinalIgnoreCase);
                foreach (var missing in expected.Except(actual, StringComparer.OrdinalIgnoreCase))
                    mismatches.Add($"merge area '{missing}' missing after restore");
                foreach (var extra in actual.Except(expected, StringComparer.OrdinalIgnoreCase))
                    mismatches.Add($"unexpected merge area '{extra}' after restore");

                object? cells = null;
                object? anchor = null;
                try
                {
                    cells = (object)((dynamic)requested).Cells;
                    anchor = (object)((dynamic)cells).Item(1);
                    var actualAnchor = ToJsonValue(ReadFormulaOrValue(anchor));
                    if (!JsonNode.DeepEquals(state["anchorContent"], actualAnchor))
                        mismatches.Add($"{sheetName}!{address}: upper-left value/formula restore mismatch");
                }
                finally
                {
                    RotHelper.ReleaseComReference(anchor);
                    RotHelper.ReleaseComReference(cells);
                }
                return BuildRestoreResult(
                    mismatches.Count == 0,
                    restoredStyleCells,
                    expected.Count + restoredStyleCells + 1,
                    MergeRestoreMode,
                    mismatches);
            }
            finally { RotHelper.ReleaseComReference(requested); }
        }
        catch (Exception ex)
        {
            mismatches.Add($"merge restore failed: {ex.Message}");
            return BuildRestoreResult(false, 0, 0, MergeRestoreMode, mismatches);
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static JsonObject ReadRangeLayout(object sheet, object range)
    {
        object? rangeRows = null;
        object? rangeColumns = null;
        try
        {
            rangeRows = (object)((dynamic)range).Rows;
            rangeColumns = (object)((dynamic)range).Columns;
            var rowCount = Convert.ToInt32(((dynamic)rangeRows).Count, CultureInfo.InvariantCulture);
            var columnCount = Convert.ToInt32(((dynamic)rangeColumns).Count, CultureInfo.InvariantCulture);
            var firstRow = Convert.ToInt32(((dynamic)range).Row, CultureInfo.InvariantCulture);
            var firstColumn = Convert.ToInt32(((dynamic)range).Column, CultureInfo.InvariantCulture);
            var limitedRows = Math.Min(rowCount, MaxCells);
            var limitedColumns = Math.Min(columnCount, MaxCells);
            var rowStates = CaptureHiddenStates(sheet, true, firstRow, limitedRows);
            var columnStates = CaptureHiddenStates(sheet, false, firstColumn, limitedColumns);
            var rowHidden = new JsonArray();
            for (var index = 0; index < rowStates.Count; index++)
                rowHidden.Add(new JsonObject { ["row"] = firstRow + index, ["hidden"] = rowStates[index]!.GetValue<bool>() });
            var colHidden = new JsonArray();
            for (var index = 0; index < columnStates.Count; index++)
                colHidden.Add(new JsonObject
                {
                    ["col"] = ColName(firstColumn + index),
                    ["column"] = firstColumn + index,
                    ["hidden"] = columnStates[index]!.GetValue<bool>(),
                });

            var mergedAreas = new JsonArray();
            var totalCells = (long)rowCount * columnCount;
            if (totalCells <= MaxCells)
                foreach (var area in EnumerateMergeAreas(range, MaxCells)) mergedAreas.Add(area.Address);

            var visibility = Convert.ToInt32(((dynamic)sheet).Visible, CultureInfo.InvariantCulture);
            return new JsonObject
            {
                ["sheetVisibility"] = SheetVisibilityName(visibility),
                ["rowStates"] = rowHidden,
                ["columnStates"] = colHidden,
                ["mergedAreas"] = mergedAreas,
                ["coverage"] = new JsonObject
                {
                    ["complete"] = rowCount <= MaxCells && columnCount <= MaxCells && totalCells <= MaxCells,
                    ["requestedRows"] = rowCount,
                    ["returnedRows"] = limitedRows,
                    ["requestedColumns"] = columnCount,
                    ["returnedColumns"] = limitedColumns,
                    ["mergedAreaScanCells"] = totalCells <= MaxCells ? totalCells : 0,
                },
            };
        }
        finally
        {
            RotHelper.ReleaseComReference(rangeColumns);
            RotHelper.ReleaseComReference(rangeRows);
        }
    }
}
