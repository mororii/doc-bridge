using System.Globalization;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class ExcelAdapter
{
    internal const string FormatOnlyRestoreMode = "format-only";

    public JsonObject ValidatePreviewReuse(
        string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject> ops)
    {
        return ComInvoke(() => ValidateFormatPreviewReuse(snapshotDir, metadata, ops));
    }

    private JsonObject ValidateFormatPreviewReuse(
        string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject> ops)
    {
        if (!IsFormatOnlySnapshot(ops))
        {
            return new JsonObject
            {
                ["ok"] = true,
                ["reusable"] = false,
                ["freshPreviewAllowed"] = true,
                ["fingerprintAvailable"] = false,
                ["fingerprintMethod"] = "excel-not-fingerprinted",
                ["reason"] = "Excel content, layout, and mixed batches are not fingerprinted; a fresh preview is required",
            };
        }

        var statePath = Path.Combine(snapshotDir, "state.json");
        if (!File.Exists(statePath))
        {
            return new JsonObject
            {
                ["ok"] = true,
                ["reusable"] = false,
                ["freshPreviewAllowed"] = false,
                ["fingerprintAvailable"] = true,
                ["fingerprintMethod"] = "excel-format-target-sha256",
                ["reason"] = "format-only snapshot state.json is missing",
            };
        }

        var state = JsonNode.Parse(File.ReadAllText(statePath)) as JsonObject ?? new JsonObject();
        if (!string.Equals(Json.GetString(state, "restoreMode"), FormatOnlyRestoreMode, StringComparison.Ordinal))
        {
            return new JsonObject
            {
                ["ok"] = true,
                ["reusable"] = false,
                ["freshPreviewAllowed"] = false,
                ["fingerprintAvailable"] = true,
                ["fingerprintMethod"] = "excel-format-target-sha256",
                ["reason"] = "snapshot restoreMode is not format-only",
            };
        }

        var expected = Json.GetString(state, "fingerprint") ?? Json.GetString(metadata, "formatFingerprint");
        var app = AttachExcel();
        if (app is null)
        {
            return new JsonObject
            {
                ["ok"] = true,
                ["reusable"] = false,
                ["freshPreviewAllowed"] = false,
                ["fingerprintAvailable"] = true,
                ["fingerprintMethod"] = "excel-format-target-sha256",
                ["reason"] = "Excel not running",
            };
        }

        var documentRef = Json.GetString(state, "documentRef")
            ?? Json.GetString(metadata, "documentRef");
        using var lease = ResolveSnapshotWorkbook((object)app, documentRef, ops, allowFileOpen: false);
        var current = CaptureFormatOnlyState(lease.Workbook, ops, documentRef ?? ReadWorkbookFullName(lease.Workbook));
        var actual = Json.GetString(current, "fingerprint");
        var reusable = !string.IsNullOrWhiteSpace(expected) &&
                       string.Equals(expected, actual, StringComparison.OrdinalIgnoreCase);
        return new JsonObject
        {
            ["ok"] = true,
            ["reusable"] = reusable,
            ["freshPreviewAllowed"] = false,
            ["fingerprintAvailable"] = true,
            ["fingerprintMethod"] = "excel-format-target-sha256",
            ["reason"] = reusable
                ? "format-only target style fingerprint matched"
                : "format-only target style fingerprint changed after dry-run",
            ["expectedFingerprint"] = expected,
            ["currentFingerprint"] = actual,
        };
    }

    private static WorkbookLease ResolveSnapshotWorkbook(
        object attachedApplication, string? documentRef, IReadOnlyList<JsonObject>? ops, bool allowFileOpen)
    {
        if (!string.IsNullOrWhiteSpace(documentRef))
        {
            var opened = FindOpenWorkbook(attachedApplication, documentRef, allowFileOpenFallback: allowFileOpen);
            return new WorkbookLease(opened.Workbook, opened);
        }

        dynamic app = attachedApplication;
        dynamic defaultWorkbook = RequireWorkbook(app);
        return ResolveTargetWorkbook(app, defaultWorkbook, ops);
    }

    private static string? ReadWorkbookFullName(object workbook)
    {
        try { return Convert.ToString(((dynamic)workbook).FullName, CultureInfo.InvariantCulture); }
        catch { return null; }
    }

    private JsonObject CaptureFormatOnlyState(object workbook, IReadOnlyList<JsonObject> ops, string? documentRef)
    {
        var formatStates = CaptureFormatStates(workbook, ops, out var cellCount, out var coverage);
        var fingerprint = ComputeFormatStatesFingerprint(formatStates);
        return new JsonObject
        {
            ["snapshotVersion"] = FormatOnlyScopedSnapshotVersion,
            ["restoreMode"] = FormatOnlyRestoreMode,
            ["styleScope"] = WrittenStyleScope,
            ["documentRef"] = documentRef,
            ["ops"] = CloneOps(ops),
            ["formatStates"] = formatStates,
            ["coverage"] = coverage,
            ["cellCount"] = cellCount,
            ["fingerprint"] = fingerprint,
        };
    }

    private JsonArray CaptureFormatStates(
        object workbook, IReadOnlyList<JsonObject> ops, out long cellCount, out JsonObject coverage)
    {
        var formatStates = new JsonArray();
        var occupied = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        cellCount = 0;
        var overlapping = 0;
        var ranges = new JsonArray();

        foreach (var op in ops)
        {
            if (!string.Equals(Json.GetString(op, "op"), "format_range", StringComparison.OrdinalIgnoreCase))
                continue;

            var scoped = CollectWrittenFormatScope(op);
            var resolvedRange = ResolveRangeTarget(
                workbook,
                Json.GetString(Json.GetObj(op, "target"), "sheet"),
                Json.GetString(op, "range")!,
                requireExplicitSheet: true);
            dynamic sheet = resolvedRange.Sheet;
            var requestedAddress = resolvedRange.Address;
            var sheetName = Convert.ToString(sheet.Name, CultureInfo.InvariantCulture) ?? "";
            object? rangeObject = null;
            object? areasObject = null;
            try
            {
                rangeObject = (object)sheet.Range(requestedAddress);
                dynamic range = rangeObject;
                try { areasObject = (object)range.Areas; }
                catch (Exception ex)
                {
                    throw new InvalidOperationException(
                        $"[EXCEL_FORMAT_MULTI_AREA] {sheetName}!{requestedAddress}: Areas could not be read; format-only snapshot refuses a partial first-area capture ({ex.Message})");
                }

                dynamic areas = areasObject;
                var areaCount = Convert.ToInt32(areas.Count, CultureInfo.InvariantCulture);
                if (areaCount < 1)
                    throw new InvalidOperationException(
                        $"[EXCEL_FORMAT_EMPTY_RANGE] {sheetName}!{requestedAddress} has no areas; exact rollback cannot be proved");

                for (var areaIndex = 1; areaIndex <= areaCount; areaIndex++)
                {
                    object? areaObject = null;
                    try
                    {
                        areaObject = (object)areas.Item(areaIndex);
                        CaptureOneFormatArea(
                            sheetName,
                            requestedAddress,
                            areaObject,
                            scoped,
                            formatStates,
                            ranges,
                            occupied,
                            ref cellCount,
                            ref overlapping);
                    }
                    finally { RotHelper.ReleaseComReference(areaObject); }
                }
            }
            finally
            {
                RotHelper.ReleaseComReference(areasObject);
                RotHelper.ReleaseComReference(rangeObject);
            }
        }

        coverage = new JsonObject
        {
            ["ranges"] = ranges,
            ["cellCount"] = cellCount,
            ["uniqueCells"] = occupied.Count,
            ["overlappingCells"] = overlapping,
            ["complete"] = true,
        };
        return formatStates;
    }

    private static void CaptureOneFormatArea(
        string sheetName,
        string requestedAddress,
        object areaObject,
        IReadOnlySet<string> scoped,
        JsonArray formatStates,
        JsonArray ranges,
        Dictionary<string, int> occupied,
        ref long cellCount,
        ref int overlapping)
    {
        dynamic area = areaObject;
        object? areaRows = null;
        object? areaColumns = null;
        object? cells = null;
        try
        {
            areaRows = (object)area.Rows;
            areaColumns = (object)area.Columns;
            var rows = Convert.ToInt32(((dynamic)areaRows).Count, CultureInfo.InvariantCulture);
            var cols = Convert.ToInt32(((dynamic)areaColumns).Count, CultureInfo.InvariantCulture);
            var startRow = Convert.ToInt32(area.Row, CultureInfo.InvariantCulture);
            var startCol = Convert.ToInt32(area.Column, CultureInfo.InvariantCulture);
            var areaAddress = ReadAreaAddress(area, requestedAddress);
            if (rows < 1 || cols < 1)
                throw new InvalidOperationException(
                    $"[EXCEL_FORMAT_EMPTY_RANGE] {sheetName}!{areaAddress} has empty dimensions; exact rollback cannot be proved");

            cellCount += (long)rows * cols;
            if (cellCount > MaxFormatSnapshotCells)
                throw new InvalidOperationException(
                    $"format snapshot exceeds {MaxFormatSnapshotCells} cells; write was blocked because cell formatting could not be restored safely");

            RejectPartialMergesInArea(areaObject, sheetName, areaAddress, startRow, startCol, rows, cols);
            MarkFormatCoverage(sheetName, startRow, startCol, rows, cols, occupied, ref overlapping);

            var areaRef = $"{sheetName}!{areaAddress}";
            var canonicalAddress = FormatRangeAddress(startRow, startCol, rows, cols);
            if (CanUseUniformRangeFastPath(scoped)
                && TryCaptureUniformFastPath(areaObject, scoped, areaRef, out var uniformStyle))
            {
                formatStates.Add(CreateScopedFormatState(
                    sheetName,
                    canonicalAddress,
                    startRow,
                    startCol,
                    rows,
                    cols,
                    scoped,
                    FormatStyleModeUniform,
                    uniformStyle,
                    null,
                    null));
            }
            else
            {
                cells = (object)area.Cells;
                var captured = new JsonObject[rows, cols];
                for (var row = 1; row <= rows; row++)
                {
                    for (var col = 1; col <= cols; col++)
                    {
                        object? cell = null;
                        try
                        {
                            cell = (object)((dynamic)cells).Item(row, col);
                            var cellRef = $"{sheetName}!{CellName(startCol + col - 1, startRow + row - 1)}";
                            captured[row - 1, col - 1] = CaptureScopedFormatStyle(cell, cellRef, scoped);
                        }
                        finally { RotHelper.ReleaseComReference(cell); }
                    }
                }

                var groups = GroupUniformStyleRectangles(captured, rows, cols, startRow, startCol);
                formatStates.Add(CreateScopedFormatState(
                    sheetName,
                    canonicalAddress,
                    startRow,
                    startCol,
                    rows,
                    cols,
                    scoped,
                    FormatStyleModeGroups,
                    null,
                    groups,
                    null));
            }

            ranges.Add(new JsonObject
            {
                ["sheet"] = sheetName,
                ["range"] = areaAddress,
                ["cells"] = rows * cols,
            });
        }
        finally
        {
            RotHelper.ReleaseComReference(cells);
            RotHelper.ReleaseComReference(areaColumns);
            RotHelper.ReleaseComReference(areaRows);
        }
    }

    private static string ReadAreaAddress(dynamic area, string fallback)
    {
        try
        {
            var address = Convert.ToString(area.Address(false, false), CultureInfo.InvariantCulture);
            return string.IsNullOrWhiteSpace(address) ? fallback : address;
        }
        catch
        {
            return fallback;
        }
    }

    private static void RejectPartialMergeCoverage(
        object cell, string cellRef, int startRow, int startCol, int rows, int cols)
    {
        object? mergeArea = null;
        object? mergeRows = null;
        object? mergeColumns = null;
        try
        {
            var mergeCells = RequireScalar(((dynamic)cell).MergeCells, "MergeCells", cellRef);
            if (!Convert.ToBoolean(mergeCells, CultureInfo.InvariantCulture)) return;

            mergeArea = (object)((dynamic)cell).MergeArea;
            dynamic area = mergeArea;
            mergeRows = (object)area.Rows;
            mergeColumns = (object)area.Columns;
            var mergeInfo = new MergeAreaInfo(
                ReadAreaAddress(area, cellRef),
                Convert.ToInt32(area.Row, CultureInfo.InvariantCulture),
                Convert.ToInt32(area.Column, CultureInfo.InvariantCulture),
                Convert.ToInt32(((dynamic)mergeRows).Count, CultureInfo.InvariantCulture),
                Convert.ToInt32(((dynamic)mergeColumns).Count, CultureInfo.InvariantCulture));
            if (!IsWithin(mergeInfo, startRow, startCol, rows, cols))
                throw new InvalidOperationException(
                    $"[EXCEL_FORMAT_PARTIAL_MERGE] {cellRef} is part of merged range '{mergeInfo.Address}' that extends outside the format target; exact rollback cannot be proved");
        }
        finally
        {
            RotHelper.ReleaseComReference(mergeColumns);
            RotHelper.ReleaseComReference(mergeRows);
            RotHelper.ReleaseComReference(mergeArea);
        }
    }

    private static string ComputeFormatStatesFingerprint(JsonArray formatStates) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(Json.Canonical(formatStates)))).ToLowerInvariant();

    private static JsonArray CloneOps(IReadOnlyList<JsonObject> ops)
    {
        var captured = new JsonArray();
        foreach (var op in ops) captured.Add(op.DeepClone());
        return captured;
    }

    private JsonObject RestoreFormatOnlyState(object workbook, JsonObject state)
    {
        var version = Json.GetInt(state, "snapshotVersion");
        var scoped = version == FormatOnlyScopedSnapshotVersion;
        if (version != CurrentExcelSnapshotVersion && !scoped)
            return Json.ErrorResult("unsupported Excel format-only snapshot version", App);
        if (scoped && !string.Equals(Json.GetString(state, "styleScope"), WrittenStyleScope, StringComparison.Ordinal))
            return Json.ErrorResult("unsupported Excel format-only styleScope; refusing restore", App);

        if (!TryValidateFormatOnlyStateShape(state, scoped, out var formatStates, out var expectedCells, out var shapeError))
            return Json.ErrorResult(shapeError, App);

        var mismatches = new RestoreMismatchCollector();
        var checkedCells = 0;
        if (scoped)
            ApplyScopedFormatStates(workbook, formatStates, mismatches, ref checkedCells);
        else
            ApplyFormatOnlyStates(workbook, formatStates, mismatches, ref checkedCells);
        if (checkedCells == 0)
        {
            mismatches.Add("format-only restore checked zero cells; refusing to claim a verified rollback");
        }
        else if (checkedCells != expectedCells)
        {
            mismatches.Add(
                $"format-only restore coverage mismatch: expected {expectedCells} cells, checked {checkedCells}");
        }

        return BuildRestoreResult(
            mismatches.Count == 0,
            restoredCells: checkedCells,
            checkedItems: checkedCells,
            FormatOnlyRestoreMode,
            mismatches);
    }

    private static bool TryValidateFormatOnlyStateShape(
        JsonObject state, bool scoped, out JsonArray formatStates, out int expectedCells, out string error)
    {
        formatStates = new JsonArray();
        expectedCells = 0;
        error = "";

        if (state["formatStates"] is not JsonArray states)
        {
            error = "format-only snapshot is missing formatStates; refusing restore";
            return false;
        }
        if (Json.GetObj(state, "coverage") is not JsonObject coverage)
        {
            error = "format-only snapshot is missing coverage; refusing restore";
            return false;
        }
        if (!Json.GetBool(coverage, "complete"))
        {
            error = "format-only snapshot coverage is incomplete; refusing restore";
            return false;
        }
        if (states.Count == 0)
        {
            error = "format-only snapshot formatStates is empty; refusing to claim restore of zero cells";
            return false;
        }

        if (!TryReadNonNegativeInt(coverage, "cellCount", out var declaredCells))
        {
            error = "format-only snapshot coverage.cellCount is missing or invalid; refusing restore";
            return false;
        }
        if (declaredCells > MaxFormatSnapshotCells)
        {
            error = $"format-only snapshot coverage.cellCount {declaredCells} exceeds {MaxFormatSnapshotCells}; refusing restore";
            return false;
        }
        var summed = 0;
        for (var index = 0; index < states.Count; index++)
        {
            if (states[index] is not JsonObject formatState)
            {
                error = $"format-only snapshot formatStates[{index}] must be an object; refusing restore";
                return false;
            }
            if (scoped)
            {
                if (!TryValidateScopedFormatStateEntry(formatState, index, out var scopedCells, out error))
                    return false;
                summed += scopedCells;
                continue;
            }

            if (!TryValidateFormatStateEntry(formatState, index, out var cells, out error))
                return false;
            summed += cells;
        }

        if (declaredCells != summed)
        {
            error = $"format-only snapshot coverage.cellCount {declaredCells} does not match styles ({summed}); refusing restore";
            return false;
        }

        formatStates = states;
        expectedCells = summed;
        return true;
    }

    private static bool TryValidateFormatStateEntry(
        JsonObject formatState, int index, out int cells, out string error)
    {
        cells = 0;
        error = "";
        var sheet = Json.GetString(formatState, "sheet");
        var address = Json.GetString(formatState, "range");
        if (string.IsNullOrWhiteSpace(sheet) || string.IsNullOrWhiteSpace(address))
        {
            error = $"format-only snapshot formatStates[{index}] is missing sheet/range; refusing restore";
            return false;
        }
        if (!TryReadPositiveInt(formatState, "row", out _) ||
            !TryReadPositiveInt(formatState, "column", out _) ||
            !TryReadPositiveInt(formatState, "rows", out var rows) ||
            !TryReadPositiveInt(formatState, "columns", out var columns))
        {
            error = $"format-only snapshot formatStates[{index}] ({sheet}!{address}) is missing dimensions; refusing restore";
            return false;
        }
        if (formatState["styles"] is not JsonArray styleRows)
        {
            error = $"format-only snapshot formatStates[{index}] ({sheet}!{address}) is missing styles; refusing restore";
            return false;
        }
        if (styleRows.Count != rows)
        {
            error = $"format-only snapshot formatStates[{index}] ({sheet}!{address}) styles rows {styleRows.Count} != {rows}; refusing restore";
            return false;
        }
        for (var row = 0; row < styleRows.Count; row++)
        {
            if (styleRows[row] is not JsonArray styleCols || styleCols.Count != columns)
            {
                error = $"format-only snapshot formatStates[{index}] ({sheet}!{address}) styles[{row}] does not match columns {columns}; refusing restore";
                return false;
            }
            for (var col = 0; col < styleCols.Count; col++)
            {
                if (styleCols[col] is not JsonObject style)
                {
                    error = $"format-only snapshot formatStates[{index}] ({sheet}!{address}) styles[{row}][{col}] is not an object; refusing restore";
                    return false;
                }
                if (!HasRequiredFormatOnlyStyleKeys(style, out var missing))
                {
                    error = $"format-only snapshot formatStates[{index}] ({sheet}!{address}) styles[{row}][{col}] is missing '{missing}'; refusing restore";
                    return false;
                }
            }
        }

        cells = rows * columns;
        return true;
    }

    private static readonly string[] RequiredFormatOnlyStyleKeys =
    {
        "bold", "italic", "fontSize", "numberFormat",
        "fontColor", "fontColorIndex", "fontTintAndShade",
        "fillColor", "fillColorIndex", "fillTintAndShade",
        "fillPattern", "fillPatternColor", "fillPatternColorIndex", "fillPatternTintAndShade",
    };

    private static bool HasRequiredFormatOnlyStyleKeys(JsonObject style, out string missing)
    {
        foreach (var key in RequiredFormatOnlyStyleKeys)
        {
            if (!style.ContainsKey(key) || style[key] is null)
            {
                missing = key;
                return false;
            }
        }
        missing = "";
        return true;
    }

    private static void ApplyFormatOnlyStates(
        object workbook, JsonArray formatStates, RestoreMismatchCollector mismatches, ref int checkedCells)
    {
        foreach (var formatNode in formatStates)
        {
            var formatState = (JsonObject)formatNode!;
            var sheetName = Json.GetString(formatState, "sheet")!;
            var address = Json.GetString(formatState, "range")!;
            var styleRows = (JsonArray)formatState["styles"]!;
            dynamic sheet = GetSheet(workbook, sheetName);
            object? rangeObject = null;
            object? cells = null;
            try
            {
                rangeObject = (object)sheet.Range(address);
                dynamic range = rangeObject;
                cells = (object)range.Cells;
                var startRow = Convert.ToInt32(range.Row, CultureInfo.InvariantCulture);
                var startCol = Convert.ToInt32(range.Column, CultureInfo.InvariantCulture);
                for (var row = 0; row < styleRows.Count; row++)
                {
                    var styleCols = (JsonArray)styleRows[row]!;
                    for (var col = 0; col < styleCols.Count; col++)
                    {
                        var style = (JsonObject)styleCols[col]!;
                        object? cell = null;
                        try
                        {
                            cell = (object)((dynamic)cells).Item(row + 1, col + 1);
                            RestoreFormatOnlyCellStyle(cell, style);
                            checkedCells++;
                            if (!FormatOnlyCellStyleMatches(cell, style))
                                mismatches.Add($"{sheetName}!{CellName(startCol + col, startRow + row)}: style restore mismatch");
                        }
                        finally { RotHelper.ReleaseComReference(cell); }
                    }
                }
            }
            finally
            {
                RotHelper.ReleaseComReference(cells);
                RotHelper.ReleaseComReference(rangeObject);
            }
        }
    }

    private static void ApplyFormat(dynamic wb, JsonObject op, ApplyExecution exec,
        List<string> mismatches, ref int checkedCells)
    {
        var style = Json.GetObj(op, "style") ?? new JsonObject();
        var styleErrors = new List<string>();
        if (!ExcelStyleContract.TryNormalize(style, out var canonical, styleErrors))
            throw new InvalidOperationException(string.Join("; ", styleErrors));

        var resolvedRange = ResolveRangeTarget(
            (object)wb,
            Json.GetString(Json.GetObj(op, "target"), "sheet"),
            Json.GetString(op, "range")!,
            requireExplicitSheet: true);
        dynamic sheet = resolvedRange.Sheet;
        var rangeAddr = resolvedRange.Address;
        object? rangeObject = null;
        object? font = null;
        object? interior = null;
        try
        {
            rangeObject = (object)sheet.Range(rangeAddr);
            dynamic range = rangeObject;
            font = (object)range.Font;
            interior = (object)range.Interior;
            dynamic dynamicFont = font;
            dynamic dynamicInterior = interior;

            if (canonical.ContainsKey(ExcelStyleContract.Bold))
                dynamicFont.Bold = Json.GetBool(canonical, ExcelStyleContract.Bold);
            if (canonical.ContainsKey(ExcelStyleContract.Italic))
                dynamicFont.Italic = Json.GetBool(canonical, ExcelStyleContract.Italic);
            if (canonical.ContainsKey(ExcelStyleContract.FontSize))
                dynamicFont.Size = canonical[ExcelStyleContract.FontSize]!.GetValue<double>();
            if (canonical.ContainsKey(ExcelStyleContract.NumberFormat))
                AssignRangeNumberFormat(range, Json.GetString(canonical, ExcelStyleContract.NumberFormat) ?? "");
            if (canonical.ContainsKey(ExcelStyleContract.FontColor))
                dynamicFont.Color = canonical[ExcelStyleContract.FontColor]!.GetValue<double>();
            if (canonical.ContainsKey(ExcelStyleContract.NoFill) &&
                Json.GetBool(canonical, ExcelStyleContract.NoFill))
                dynamicInterior.Pattern = ExcelStyleContract.XlPatternNone;
            if (canonical.ContainsKey(ExcelStyleContract.FillColor))
            {
                dynamicInterior.Color = canonical[ExcelStyleContract.FillColor]!.GetValue<double>();
                dynamicInterior.Pattern = ExcelStyleContract.XlPatternSolid;
            }
            if (canonical.ContainsKey(ExcelStyleContract.FillPattern))
                dynamicInterior.Pattern = ExcelStyleContract.FillPatternValue(
                    Json.GetString(canonical, ExcelStyleContract.FillPattern)!);
            if (canonical.ContainsKey(ExcelStyleContract.FontName))
                dynamicFont.Name = Json.GetString(canonical, ExcelStyleContract.FontName);
            if (canonical.ContainsKey(ExcelStyleContract.HorizontalAlign))
                range.HorizontalAlignment = ExcelStyleContract.HorizontalAlignValue(
                    Json.GetString(canonical, ExcelStyleContract.HorizontalAlign)!);
            if (canonical.ContainsKey(ExcelStyleContract.VerticalAlign))
                range.VerticalAlignment = ExcelStyleContract.VerticalAlignValue(
                    Json.GetString(canonical, ExcelStyleContract.VerticalAlign)!);
            if (canonical.ContainsKey(ExcelStyleContract.WrapText))
                range.WrapText = Json.GetBool(canonical, ExcelStyleContract.WrapText);
            if (canonical.ContainsKey(ExcelStyleContract.ShrinkToFit))
                range.ShrinkToFit = Json.GetBool(canonical, ExcelStyleContract.ShrinkToFit);
            if (canonical.ContainsKey(ExcelStyleContract.Underline))
                dynamicFont.Underline = ExcelStyleContract.UnderlineValue(
                    Json.GetString(canonical, ExcelStyleContract.Underline)!);
            if (canonical.ContainsKey(ExcelStyleContract.Strikethrough))
                dynamicFont.Strikethrough = Json.GetBool(canonical, ExcelStyleContract.Strikethrough);
            if (canonical.ContainsKey(ExcelStyleContract.Indent))
                range.IndentLevel = Json.GetInt(canonical, ExcelStyleContract.Indent)!.Value;
            if (canonical.ContainsKey(ExcelStyleContract.Orientation))
                range.Orientation = ExcelStyleContract.OrientationValue(canonical[ExcelStyleContract.Orientation]!);
            if (canonical.ContainsKey(ExcelStyleContract.Locked))
                range.Locked = Json.GetBool(canonical, ExcelStyleContract.Locked);
            if (canonical.ContainsKey(ExcelStyleContract.Borders) &&
                Json.GetObj(canonical, ExcelStyleContract.Borders) is { } borders)
                ApplyCanonicalBorders(rangeObject, borders);

            exec.Affected.Add(new AffectedRef("range", $"{sheet.Name}!{rangeAddr}"));

            foreach (var (key, node) in canonical)
            {
                if (node is null) continue;
                checkedCells++;
                try
                {
                    var matched = key switch
                    {
                        ExcelStyleContract.Borders => AppliedBordersMatch(
                            rangeObject, Json.GetObj(canonical, ExcelStyleContract.Borders)),
                        ExcelStyleContract.Bold => ExcelFormatReadback.AppliedBoldMatches(
                            rangeObject, node.GetValue<bool>()),
                        _ => ExcelFormatAreas.AllAreas(
                            rangeObject, area => VerifyNonBorderStyleOnArea(area, key, node)),
                    };
                    if (!matched)
                    {
                        if (key == ExcelStyleContract.Borders)
                        {
                            var why = ExcelBorderReadback.ExplainAppliedBorders(
                                rangeObject, Json.GetObj(canonical, ExcelStyleContract.Borders));
                            mismatches.Add(
                                $"{sheet.Name}!{rangeAddr}: style 'borders' readback mismatch; {string.Join("; ", why)}");
                        }
                        else
                            mismatches.Add($"{sheet.Name}!{rangeAddr}: style '{key}' readback mismatch");
                    }
                }
                catch (Exception ex)
                {
                    mismatches.Add($"{sheet.Name}!{rangeAddr}: style '{key}' readback failed: {ex.Message}");
                }
            }
        }
        finally
        {
            RotHelper.ReleaseComReference(interior);
            RotHelper.ReleaseComReference(font);
            RotHelper.ReleaseComReference(rangeObject);
        }
    }

    private static bool VerifyNonBorderStyleOnArea(object areaObject, string key, JsonNode node)
    {
        object? font = null;
        object? interior = null;
        try
        {
            dynamic range = areaObject;
            font = (object)range.Font;
            interior = (object)range.Interior;
            dynamic dynamicFont = font;
            dynamic dynamicInterior = interior;
            return key switch
            {
                ExcelStyleContract.Italic => ScalarBoolEquals((object?)dynamicFont.Italic, node.GetValue<bool>()),
                ExcelStyleContract.FontSize => ScalarDoubleEquals((object?)dynamicFont.Size, node.GetValue<double>()),
                ExcelStyleContract.NumberFormat => ExcelNumberFormatContract.ReadbackMatches(
                    node.GetValue<string>(),
                    Convert.ToString((object?)range.NumberFormat, CultureInfo.InvariantCulture)),
                ExcelStyleContract.FontColor => AppliedColorMatches(
                    areaObject, node.GetValue<double>(), fill: false),
                ExcelStyleContract.FillColor => AppliedColorMatches(
                    areaObject, node.GetValue<double>(), fill: true),
                ExcelStyleContract.FontName => ScalarStringEquals(
                    (object?)dynamicFont.Name, node.GetValue<string>()),
                ExcelStyleContract.HorizontalAlign => ScalarIntEquals(
                    (object?)range.HorizontalAlignment,
                    ExcelStyleContract.HorizontalAlignValue(node.GetValue<string>())),
                ExcelStyleContract.VerticalAlign => ScalarIntEquals(
                    (object?)range.VerticalAlignment,
                    ExcelStyleContract.VerticalAlignValue(node.GetValue<string>())),
                ExcelStyleContract.WrapText => ScalarBoolEquals(
                    (object?)range.WrapText, node.GetValue<bool>()),
                ExcelStyleContract.ShrinkToFit => ScalarBoolEquals(
                    (object?)range.ShrinkToFit, node.GetValue<bool>()),
                ExcelStyleContract.Underline => ScalarIntEquals(
                    (object?)dynamicFont.Underline,
                    ExcelStyleContract.UnderlineValue(node.GetValue<string>())),
                ExcelStyleContract.Strikethrough => ScalarBoolEquals(
                    (object?)dynamicFont.Strikethrough, node.GetValue<bool>()),
                ExcelStyleContract.Indent => ScalarIntEquals(
                    (object?)range.IndentLevel, node.GetValue<int>()),
                ExcelStyleContract.Orientation => ScalarIntEquals(
                    (object?)range.Orientation, ExcelStyleContract.OrientationValue(node)),
                ExcelStyleContract.Locked => ScalarBoolEquals(
                    (object?)range.Locked, node.GetValue<bool>()),
                ExcelStyleContract.FillPattern => ScalarIntEquals(
                    (object?)dynamicInterior.Pattern,
                    ExcelStyleContract.FillPatternValue(node.GetValue<string>())),
                ExcelStyleContract.NoFill => ScalarIntEquals(
                    (object?)dynamicInterior.Pattern, ExcelStyleContract.XlPatternNone),
                _ => false,
            };
        }
        catch
        {
            return false;
        }
        finally
        {
            RotHelper.ReleaseComReference(interior);
            RotHelper.ReleaseComReference(font);
        }
    }

    private static bool ScalarBoolEquals(object? raw, bool wanted)
    {
        if (IsMixed(raw)) return false;
        return Convert.ToBoolean(raw, CultureInfo.InvariantCulture) == wanted;
    }

    private static bool ScalarDoubleEquals(object? raw, double wanted)
    {
        if (IsMixed(raw)) return false;
        return NumbersEqual(Convert.ToDouble(raw, CultureInfo.InvariantCulture), wanted);
    }

    private static bool ScalarStringEquals(object? raw, string? wanted)
    {
        if (IsMixed(raw) || wanted is null) return false;
        return string.Equals(Convert.ToString(raw, CultureInfo.InvariantCulture), wanted, StringComparison.Ordinal);
    }

    private static bool ScalarIntEquals(object? raw, int wanted)
    {
        if (IsMixed(raw)) return false;
        return Convert.ToInt32(raw, CultureInfo.InvariantCulture) == wanted;
    }

    private static bool IsZeroOleColor(double ole) => Math.Abs(ole) <= 1e-9;

    private static bool AppliedColorMatches(object rangeObject, double wanted, bool fill)
    {
        object? font = null;
        object? interior = null;
        try
        {
            dynamic range = rangeObject;
            if (fill)
            {
                interior = (object)range.Interior;
                dynamic dynamicInterior = interior;
                var colorRaw = (object?)dynamicInterior.Color;
                var patternRaw = (object?)dynamicInterior.Pattern;
                if (IsMixed(colorRaw) || IsMixed(patternRaw))
                    return EveryCellAppliedColorMatches(rangeObject, wanted, fill: true);
                var aggregate = Convert.ToDouble(colorRaw, CultureInfo.InvariantCulture);
                if (!IsZeroOleColor(aggregate))
                {
                    return NumbersEqual(aggregate, wanted)
                        && ScalarIntEquals(patternRaw, ExcelStyleContract.XlPatternSolid);
                }
            }
            else
            {
                font = (object)range.Font;
                var colorRaw = (object?)((dynamic)font).Color;
                if (IsMixed(colorRaw))
                    return EveryCellAppliedColorMatches(rangeObject, wanted, fill: false);
                var aggregate = Convert.ToDouble(colorRaw, CultureInfo.InvariantCulture);
                if (!IsZeroOleColor(aggregate))
                    return NumbersEqual(aggregate, wanted);
            }
        }
        finally
        {
            RotHelper.ReleaseComReference(interior);
            RotHelper.ReleaseComReference(font);
        }

        return EveryCellAppliedColorMatches(rangeObject, wanted, fill);
    }

    private static bool EveryCellAppliedColorMatches(object rangeObject, double wanted, bool fill)
    {
        var allMatch = true;
        ExcelFormatAreas.ForEachArea(rangeObject, area =>
        {
            ExcelFormatAreas.ForEachCell(area, cell =>
            {
                object? font = null;
                object? interior = null;
                try
                {
                    dynamic dynamicCell = cell;
                    if (fill)
                    {
                        interior = (object)dynamicCell.Interior;
                        dynamic dynamicInterior = interior;
                        if (!ScalarDoubleEquals((object?)dynamicInterior.Color, wanted)
                            || !ScalarIntEquals((object?)dynamicInterior.Pattern, ExcelStyleContract.XlPatternSolid))
                            allMatch = false;
                    }
                    else
                    {
                        font = (object)dynamicCell.Font;
                        if (!ScalarDoubleEquals((object?)((dynamic)font).Color, wanted))
                            allMatch = false;
                    }
                }
                catch
                {
                    allMatch = false;
                }
                finally
                {
                    RotHelper.ReleaseComReference(interior);
                    RotHelper.ReleaseComReference(font);
                }
            });
        });
        return allMatch;
    }

    private static JsonObject CaptureRangeStyleSummary(object range)
    {
        object? font = null;
        object? interior = null;
        try
        {
            dynamic dynamicRange = range;
            font = (object)dynamicRange.Font;
            interior = (object)dynamicRange.Interior;
            dynamic dynamicFont = font;
            dynamic dynamicInterior = interior;
            return new JsonObject
            {
                ["bold"] = ComBoolNode(SafeGet(() => (object?)dynamicFont.Bold)),
                ["italic"] = ComBoolNode(SafeGet(() => (object?)dynamicFont.Italic)),
                ["fontSize"] = ComDoubleNode(SafeGet(() => (object?)dynamicFont.Size)),
                ["numberFormat"] = ComStringNode(SafeGet(() => (object?)dynamicRange.NumberFormat)),
                ["fontColor"] = ComDoubleNode(SafeGet(() => (object?)dynamicFont.Color)),
                ["fillColor"] = ComDoubleNode(SafeGet(() => (object?)dynamicInterior.Color)),
                ["fillColorIndex"] = ComIntNode(SafeGet(() => (object?)dynamicInterior.ColorIndex)),
                ["fillPattern"] = ComIntNode(SafeGet(() => (object?)dynamicInterior.Pattern)),
                ["fontName"] = ComStringNode(SafeGet(() => (object?)dynamicFont.Name)),
                ["horizontalAlign"] = ComAlignName(SafeGet(() => (object?)dynamicRange.HorizontalAlignment), vertical: false),
                ["verticalAlign"] = ComAlignName(SafeGet(() => (object?)dynamicRange.VerticalAlignment), vertical: true),
                ["wrapText"] = ComBoolNode(SafeGet(() => (object?)dynamicRange.WrapText)),
                ["shrinkToFit"] = ComBoolNode(SafeGet(() => (object?)dynamicRange.ShrinkToFit)),
                ["underline"] = ComUnderlineName(SafeGet(() => (object?)dynamicFont.Underline)),
                ["strikethrough"] = ComBoolNode(SafeGet(() => (object?)dynamicFont.Strikethrough)),
                ["indent"] = ComIntNode(SafeGet(() => (object?)dynamicRange.IndentLevel)),
                ["orientation"] = ComIntNode(SafeGet(() => (object?)dynamicRange.Orientation)),
                ["locked"] = ComBoolNode(SafeGet(() => (object?)dynamicRange.Locked)),
                ["noFill"] = ComBoolNode(SafeGet(() =>
                    (object?)(Convert.ToInt32(dynamicInterior.Pattern, CultureInfo.InvariantCulture) == ExcelStyleContract.XlPatternNone))),
                ["borders"] = CapturePublicBorderSummary(range),
                ["mixed"] = IsMixed(SafeGet(() => (object?)dynamicFont.Bold))
                    || IsMixed(SafeGet(() => (object?)dynamicInterior.Pattern)),
            };
        }
        finally
        {
            RotHelper.ReleaseComReference(interior);
            RotHelper.ReleaseComReference(font);
        }
    }

    private static JsonObject CaptureFormatOnlyCellStyle(object cell, string cellRef)
    {
        object? font = null;
        object? interior = null;
        try
        {
            dynamic dynamicCell = cell;
            font = (object)dynamicCell.Font;
            interior = (object)dynamicCell.Interior;
            dynamic dynamicFont = font;
            dynamic dynamicInterior = interior;
            var style = new JsonObject
            {
                ["bold"] = RequireBoolean(dynamicFont.Bold, "Font.Bold", cellRef),
                ["italic"] = RequireBoolean(dynamicFont.Italic, "Font.Italic", cellRef),
                ["fontSize"] = RequireDouble(dynamicFont.Size, "Font.Size", cellRef),
                ["numberFormat"] = RequireString(dynamicCell.NumberFormat, "NumberFormat", cellRef),
                ["fontColor"] = RequireDouble(dynamicFont.Color, "Font.Color", cellRef),
                ["fontColorIndex"] = RequireInt(dynamicFont.ColorIndex, "Font.ColorIndex", cellRef),
                ["fontTintAndShade"] = RequireDouble(dynamicFont.TintAndShade, "Font.TintAndShade", cellRef),
                ["fillColor"] = RequireDouble(dynamicInterior.Color, "Interior.Color", cellRef),
                ["fillColorIndex"] = RequireInt(dynamicInterior.ColorIndex, "Interior.ColorIndex", cellRef),
                ["fillTintAndShade"] = RequireDouble(dynamicInterior.TintAndShade, "Interior.TintAndShade", cellRef),
                ["fillPattern"] = RequireInt(dynamicInterior.Pattern, "Interior.Pattern", cellRef),
                ["fillPatternColor"] = RequireDouble(dynamicInterior.PatternColor, "Interior.PatternColor", cellRef),
                ["fillPatternColorIndex"] = RequireInt(dynamicInterior.PatternColorIndex, "Interior.PatternColorIndex", cellRef),
                ["fillPatternTintAndShade"] = RequireDouble(dynamicInterior.PatternTintAndShade, "Interior.PatternTintAndShade", cellRef),
            };
            if (TryReadThemeColor(dynamicFont, "Font.ThemeColor", cellRef, out int fontTheme))
                style["fontThemeColor"] = fontTheme;
            if (TryReadThemeColor(dynamicInterior, "Interior.ThemeColor", cellRef, out int fillTheme))
                style["fillThemeColor"] = fillTheme;
            if (TryReadThemeColor(dynamicInterior, "Interior.PatternThemeColor", cellRef, out int patternTheme))
                style["fillPatternThemeColor"] = patternTheme;
            RejectUnsupportedRgbTint(cellRef, "Font.TintAndShade", style, "fontTintAndShade", "fontThemeColor");
            RejectUnsupportedRgbTint(cellRef, "Interior.TintAndShade", style, "fillTintAndShade", "fillThemeColor");
            RejectUnsupportedRgbTint(cellRef, "Interior.PatternTintAndShade", style, "fillPatternTintAndShade", "fillPatternThemeColor");
            return style;
        }
        finally
        {
            RotHelper.ReleaseComReference(interior);
            RotHelper.ReleaseComReference(font);
        }
    }

    private static void RestoreFormatOnlyCellStyle(object cell, JsonObject style)
    {
        object? font = null;
        object? interior = null;
        try
        {
            dynamic dynamicCell = cell;
            font = (object)dynamicCell.Font;
            interior = (object)dynamicCell.Interior;
            dynamic dynamicFont = font;
            dynamic dynamicInterior = interior;
            dynamicFont.Bold = RequiredBool(style, "bold");
            dynamicFont.Italic = RequiredBool(style, "italic");
            dynamicFont.Size = RequiredNumber(style, "fontSize");
            AssignRangeNumberFormat(dynamicCell, RequiredText(style, "numberFormat"));
            RestoreLinkedColor(
                dynamicFont,
                RequiredNumber(style, "fontColor"),
                RequiredInt(style, "fontColorIndex"),
                RequiredNumber(style, "fontTintAndShade"),
                OptionalInt(style, "fontThemeColor"));
            RestoreLinkedColor(
                dynamicInterior,
                RequiredNumber(style, "fillColor"),
                RequiredInt(style, "fillColorIndex"),
                RequiredNumber(style, "fillTintAndShade"),
                OptionalInt(style, "fillThemeColor"));
            RestoreLinkedColor(
                dynamicInterior,
                RequiredNumber(style, "fillPatternColor"),
                RequiredInt(style, "fillPatternColorIndex"),
                RequiredNumber(style, "fillPatternTintAndShade"),
                OptionalInt(style, "fillPatternThemeColor"),
                pattern: true);
            dynamicInterior.Pattern = RequiredInt(style, "fillPattern");
        }
        finally
        {
            RotHelper.ReleaseComReference(interior);
            RotHelper.ReleaseComReference(font);
        }
    }

    private static bool FormatOnlyCellStyleMatches(object cell, JsonObject style)
    {
        object? font = null;
        object? interior = null;
        try
        {
            dynamic dynamicCell = cell;
            font = (object)dynamicCell.Font;
            interior = (object)dynamicCell.Interior;
            dynamic dynamicFont = font;
            dynamic dynamicInterior = interior;
            return RequireBoolean(dynamicFont.Bold, "Font.Bold", "verify") == RequiredBool(style, "bold")
                && RequireBoolean(dynamicFont.Italic, "Font.Italic", "verify") == RequiredBool(style, "italic")
                && NumbersEqual(RequireDouble(dynamicFont.Size, "Font.Size", "verify"), RequiredNumber(style, "fontSize"))
                && string.Equals(RequireString(dynamicCell.NumberFormat, "NumberFormat", "verify"), RequiredText(style, "numberFormat"), StringComparison.Ordinal)
                && LinkedColorMatches(
                    dynamicFont,
                    RequiredNumber(style, "fontColor"),
                    RequiredInt(style, "fontColorIndex"),
                    RequiredNumber(style, "fontTintAndShade"),
                    OptionalInt(style, "fontThemeColor"))
                && LinkedColorMatches(
                    dynamicInterior,
                    RequiredNumber(style, "fillColor"),
                    RequiredInt(style, "fillColorIndex"),
                    RequiredNumber(style, "fillTintAndShade"),
                    OptionalInt(style, "fillThemeColor"))
                && Convert.ToInt32(dynamicInterior.Pattern, CultureInfo.InvariantCulture) == RequiredInt(style, "fillPattern")
                && LinkedColorMatches(
                    dynamicInterior,
                    RequiredNumber(style, "fillPatternColor"),
                    RequiredInt(style, "fillPatternColorIndex"),
                    RequiredNumber(style, "fillPatternTintAndShade"),
                    OptionalInt(style, "fillPatternThemeColor"),
                    pattern: true);
        }
        catch (InvalidOperationException)
        {
            return false;
        }
        finally
        {
            RotHelper.ReleaseComReference(interior);
            RotHelper.ReleaseComReference(font);
        }
    }

    private static void RestoreLinkedColor(
        dynamic target, double rgb, int colorIndex, double tint, int? themeColor, bool pattern = false)
    {
        if (pattern)
        {
            if (themeColor is int theme)
            {
                target.PatternThemeColor = theme;
                target.PatternTintAndShade = tint;
                return;
            }
            if (IsAutomaticOrNoneColorIndex(colorIndex))
            {
                target.PatternColorIndex = colorIndex;
                return;
            }
            target.PatternColor = rgb;
            if (Math.Abs(tint) > 1e-9) target.PatternTintAndShade = tint;
            return;
        }

        if (themeColor is int fillTheme)
        {
            target.ThemeColor = fillTheme;
            target.TintAndShade = tint;
            return;
        }
        if (IsAutomaticOrNoneColorIndex(colorIndex))
        {
            target.ColorIndex = colorIndex;
            return;
        }
        target.Color = rgb;
        if (Math.Abs(tint) > 1e-9) target.TintAndShade = tint;
    }

    private static bool LinkedColorMatches(
        dynamic target, double rgb, int colorIndex, double tint, int? themeColor, bool pattern = false)
    {
        if (pattern)
        {
            if (themeColor is int theme)
            {
                object? themeRaw = target.PatternThemeColor;
                object? tintRaw = target.PatternTintAndShade;
                if (IsMixed(themeRaw) || IsMixed(tintRaw)) return false;
                return Convert.ToInt32(themeRaw, CultureInfo.InvariantCulture) == theme
                    && NumbersEqual(Convert.ToDouble(tintRaw, CultureInfo.InvariantCulture), tint);
            }
            if (IsAutomaticOrNoneColorIndex(colorIndex))
            {
                object? indexRaw = target.PatternColorIndex;
                if (IsMixed(indexRaw)) return false;
                return Convert.ToInt32(indexRaw, CultureInfo.InvariantCulture) == colorIndex;
            }

            object? patternColorRaw = target.PatternColor;
            object? patternTintRaw = target.PatternTintAndShade;
            if (IsMixed(patternColorRaw) || IsMixed(patternTintRaw)) return false;
            return NumbersEqual(Convert.ToDouble(patternColorRaw, CultureInfo.InvariantCulture), rgb)
                && NumbersEqual(Convert.ToDouble(patternTintRaw, CultureInfo.InvariantCulture), tint);
        }

        if (themeColor is int fillTheme)
        {
            object? themeRaw = target.ThemeColor;
            object? tintRaw = target.TintAndShade;
            if (IsMixed(themeRaw) || IsMixed(tintRaw)) return false;
            return Convert.ToInt32(themeRaw, CultureInfo.InvariantCulture) == fillTheme
                && NumbersEqual(Convert.ToDouble(tintRaw, CultureInfo.InvariantCulture), tint);
        }
        if (IsAutomaticOrNoneColorIndex(colorIndex))
        {
            object? indexRaw = target.ColorIndex;
            if (IsMixed(indexRaw)) return false;
            return Convert.ToInt32(indexRaw, CultureInfo.InvariantCulture) == colorIndex;
        }

        object? colorRaw = target.Color;
        object? shadeRaw = target.TintAndShade;
        if (IsMixed(colorRaw) || IsMixed(shadeRaw)) return false;
        return NumbersEqual(Convert.ToDouble(colorRaw, CultureInfo.InvariantCulture), rgb)
            && NumbersEqual(Convert.ToDouble(shadeRaw, CultureInfo.InvariantCulture), tint);
    }

    private static void RejectUnsupportedRgbTint(
        string cellRef, string property, JsonObject style, string tintKey, string themeKey)
    {
        if (style.ContainsKey(themeKey)) return;
        var tint = RequiredNumber(style, tintKey);
        if (Math.Abs(tint) <= 1e-9) return;
        throw new InvalidOperationException(
            $"[EXCEL_FORMAT_UNSUPPORTED_RGB_TINT] {cellRef} {property} is a non-theme RGB color with nonzero tint; " +
            "exact rollback cannot be proved because assigning Color then Tint reapplies the shade");
    }

    private static bool IsAutomaticOrNoneColorIndex(int colorIndex) =>
        colorIndex is ExcelStyleContract.XlColorIndexAutomatic or ExcelStyleContract.XlColorIndexNone;

    private static bool TryReadThemeColor(dynamic target, string property, string cellRef, out int themeColor)
    {
        themeColor = 0;
        try
        {
            object? raw = property.Contains("PatternThemeColor", StringComparison.Ordinal)
                ? target.PatternThemeColor
                : target.ThemeColor;
            if (IsMixed(raw))
                throw new InvalidOperationException(
                    $"[EXCEL_FORMAT_UNSUPPORTED_COLOR] {cellRef} {property} is mixed; theme/automatic colors are not coerced");
            var value = Convert.ToInt32(raw, CultureInfo.InvariantCulture);
            if (value is >= ExcelStyleContract.XlThemeColorMin and <= ExcelStyleContract.XlThemeColorMax)
            {
                themeColor = value;
                return true;
            }
            return false;
        }
        catch (InvalidOperationException ex) when (ex.Message.StartsWith("[EXCEL_FORMAT_", StringComparison.Ordinal))
        {
            throw;
        }
        catch (COMException ex) when (ex.HResult == unchecked((int)0x800A03EC))
        {
            // Excel 1004: ThemeColor is unset (direct RGB / no-fill). Not a dead or busy COM server.
            return false;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException(
                $"[EXCEL_FORMAT_THEME_READ_FAILED] {cellRef} {property} could not be read " +
                $"({ex.GetType().Name} {ComFailure.FormatHResult(ex.HResult)}): {ex.Message}",
                ex);
        }
    }

    private static JsonObject BuildOperationResult(
        int index, string name, bool ok, long elapsedMs, string stage,
        Exception? error, IEnumerable<string>? mismatches)
    {
        var result = new JsonObject
        {
            ["index"] = index,
            ["op"] = name,
            ["ok"] = ok,
            ["elapsedMs"] = elapsedMs,
            ["stage"] = stage,
            ["timingScope"] = "per-op",
        };
        if (error is not null)
        {
            var diag = ComFailure.Describe(error, stage);
            result["error"] = diag["message"]?.DeepClone();
            result["exceptionType"] = diag["exceptionType"]?.DeepClone();
            result["hresult"] = diag["hresult"]?.DeepClone();
            result["hresultValue"] = diag["hresultValue"]?.DeepClone();
            result["errorCode"] = diag["errorCode"]?.DeepClone();
        }
        else if (mismatches is not null)
        {
            var list = mismatches.ToList();
            if (list.Count > 0)
            {
                result["error"] = list[0];
                result["mismatches"] = Json.ToArray(list);
            }
        }
        return result;
    }

    private static JsonObject BuildSkippedOperation(int index, string name) => new()
    {
        ["index"] = index,
        ["op"] = name,
        ["ok"] = false,
        ["elapsedMs"] = 0,
        ["stage"] = "skipped",
        ["timingScope"] = "skipped",
        ["error"] = "not executed because a previous operation failed",
    };

    private static void ApplyCanonicalBorders(object rangeObject, JsonObject borders) =>
        ExcelBorderApply.ApplyCanonical(rangeObject, borders, ApplyApplicableBorder);

    /// <summary>
    /// Cross-batch counterpart of the EXCEL_BORDER_ORDER validation: a clear
    /// cannot know earlier batches, so the preview warns when a cleared edge
    /// is currently drawn on the neighboring strip (shared edge object).
    /// </summary>
    private static void WarnBorderClearNeighbors(object sheet, string sheetName, string address,
        JsonObject? borders, ApplyPreview preview)
    {
        var cleared = new List<string>();
        if (borders is not null)
        {
            foreach (var (key, node) in borders)
            {
                if (node is not JsonValue value || !value.TryGetValue<string>(out var text) ||
                    !string.Equals(text, "none", StringComparison.OrdinalIgnoreCase))
                    continue;
                if (key is "all" or "outline")
                    cleared.AddRange(new[] { "left", "right", "top", "bottom" });
                else if (key is "left" or "right" or "top" or "bottom")
                    cleared.Add(key.ToLowerInvariant());
            }
        }
        if (cleared.Count == 0) return;
        if (!ExcelDataOperationsContract.TryParseA1(address, out var box, allowUnion: false)) return;
        var neighbors = new List<(string Edge, string Address, int Opposite)>();
        if (box.Column > 1)
            neighbors.Add(("left", $"{ColName(box.Column - 1)}{box.Row}:{ColName(box.Column - 1)}{box.Row + box.Rows - 1}", 10));
        if (box.Column + box.Columns - 1 < 16_384)
            neighbors.Add(("right", $"{ColName(box.Column + box.Columns)}{box.Row}:{ColName(box.Column + box.Columns)}{box.Row + box.Rows - 1}", 7));
        if (box.Row > 1)
            neighbors.Add(("top", $"{ColName(box.Column)}{box.Row - 1}:{ColName(box.Column + box.Columns - 1)}{box.Row - 1}", 9));
        if (box.Row + box.Rows - 1 < 1_048_576)
            neighbors.Add(("bottom", $"{ColName(box.Column)}{box.Row + box.Rows}:{ColName(box.Column + box.Columns - 1)}{box.Row + box.Rows}", 8));
        foreach (var (edge, neighborAddress, opposite) in neighbors)
        {
            if (!cleared.Contains(edge)) continue;
            var state = ReadNeighborEdge(sheet, neighborAddress, opposite);
            if (state is null) continue;
            preview.Warnings.Add(
                $"[EXCEL_BORDER_SHARED_EDGE] clearing {sheetName}!{address} also removes the {edge} edge " +
                $"shared with {sheetName}!{neighborAddress} (currently {state}); redraw that outline afterwards");
        }
    }

    private static string? ReadNeighborEdge(object sheet, string address, int edgeIndex)
    {
        object? range = null;
        object? borders = null;
        object? border = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            borders = (object)((dynamic)range).Borders;
            border = (object)((dynamic)borders).Item(edgeIndex);
            object? line = null;
            try { line = ((dynamic)border).LineStyle; }
            catch { return null; }
            if (line is null) return null;
            var style = Convert.ToInt32(line, CultureInfo.InvariantCulture);
            if (style == ExcelBorderContract.XlLineStyleNone) return null;
            return ExcelBorderContract.LineStyleName(style);
        }
        catch { return null; }
        finally
        {
            RotHelper.ReleaseComReference(border);
            RotHelper.ReleaseComReference(borders);
            RotHelper.ReleaseComReference(range);
        }
    }

    private static void ApplyApplicableBorder(object target, ExcelBorderContract.EdgeSpec spec)
    {
        if (ExcelFormatAreas.TryGetDimensions(target, out var rows, out var columns) &&
            !ExcelBorderApplicability.InsideEdgeApplies(spec.Name, rows, columns))
            return;
        ApplyOneBorder(target, spec);
    }

    private static void ApplyOneBorder(object target, ExcelBorderContract.EdgeSpec spec)
    {
        var index = ExcelBorderContract.EdgeIndex(spec.Name);
        if (index == 0) return;
        object? borders = null;
        object? border = null;
        try
        {
            borders = (object)((dynamic)target).Borders;
            border = (object)((dynamic)borders).Item(index);
            if (string.Equals(spec.LineStyle, ExcelBorderContract.LineNone, StringComparison.OrdinalIgnoreCase))
            {
                ((dynamic)border).LineStyle = ExcelBorderContract.XlLineStyleNone;
                return;
            }

            ((dynamic)border).LineStyle = ExcelBorderContract.LineStyleValues[spec.LineStyle];
            ((dynamic)border).Weight = ExcelBorderContract.WeightValues[spec.Weight];
            ((dynamic)border).Color = spec.Color;
        }
        finally
        {
            RotHelper.ReleaseComReference(border);
            RotHelper.ReleaseComReference(borders);
        }
    }

    private static bool AppliedBordersMatch(object rangeObject, JsonObject? borders)
    {
        if (borders is null) return false;
        return ExcelFormatAreas.AllAreas(rangeObject, area =>
        {
            if (!ExcelFormatAreas.TryGetDimensions(area, out var rows, out var columns))
                return ExcelBorderReadback.AppliedBordersMatch(area, borders);
            return ExcelBorderReadback.AppliedBordersMatch(
                area, ExcelBorderApplicability.ForRange(borders, rows, columns));
        });
    }

    private static bool AppliedCellBorderMatches(object rangeObject, ExcelBorderContract.EdgeSpec spec) =>
        ExcelBorderReadback.AppliedCellBorderMatches(rangeObject, spec);

    private static bool OneBorderMatches(object target, ExcelBorderContract.EdgeSpec spec) =>
        ExcelBorderReadback.OneBorderMatches(target, spec);

    private static JsonNode? ComUnderlineName(object? raw)
    {
        if (IsMixed(raw)) return null;
        try
        {
            return JsonValue.Create(ExcelStyleContract.UnderlineName(Convert.ToInt32(raw, CultureInfo.InvariantCulture)));
        }
        catch { return null; }
    }

    private static JsonNode? ComAlignName(object? raw, bool vertical)
    {
        if (IsMixed(raw)) return null;
        try
        {
            var value = Convert.ToInt32(raw, CultureInfo.InvariantCulture);
            return vertical
                ? JsonValue.Create(ExcelStyleContract.VerticalAlignName(value))
                : JsonValue.Create(ExcelStyleContract.HorizontalAlignName(value));
        }
        catch { return null; }
    }

    private static JsonObject CapturePublicBorderSummary(object range)
    {
        var edges = new JsonObject();
        foreach (var (name, index) in new[]
                 {
                     ("left", ExcelBorderContract.XlEdgeLeft),
                     ("right", ExcelBorderContract.XlEdgeRight),
                     ("top", ExcelBorderContract.XlEdgeTop),
                     ("bottom", ExcelBorderContract.XlEdgeBottom),
                     ("insideVertical", ExcelBorderContract.XlInsideVertical),
                     ("insideHorizontal", ExcelBorderContract.XlInsideHorizontal),
                 })
        {
            object? borders = null;
            object? border = null;
            try
            {
                borders = (object)((dynamic)range).Borders;
                border = (object)((dynamic)borders).Item(index);
                var line = Convert.ToInt32(((dynamic)border).LineStyle, CultureInfo.InvariantCulture);
                edges[name] = new JsonObject
                {
                    ["lineStyle"] = ExcelBorderContract.LineStyleName(line),
                    ["weight"] = ExcelBorderContract.WeightName(
                        Convert.ToInt32(((dynamic)border).Weight, CultureInfo.InvariantCulture)),
                    ["color"] = Convert.ToDouble(((dynamic)border).Color, CultureInfo.InvariantCulture),
                };
            }
            catch
            {
                edges[name] = null;
            }
            finally
            {
                RotHelper.ReleaseComReference(border);
                RotHelper.ReleaseComReference(borders);
            }
        }

        return edges;
    }

    private static object? SafeGet(Func<object?> read)
    {
        try { return read(); }
        catch { return null; }
    }

    private static bool IsMixed(object? value) => value is null or DBNull;

    private static JsonNode? ComBoolNode(object? value)
    {
        if (IsMixed(value)) return null;
        try { return JsonValue.Create(Convert.ToBoolean(value, CultureInfo.InvariantCulture)); }
        catch { return null; }
    }

    private static JsonNode? ComDoubleNode(object? value)
    {
        if (IsMixed(value)) return null;
        try { return JsonValue.Create(Convert.ToDouble(value, CultureInfo.InvariantCulture)); }
        catch { return null; }
    }

    private static JsonNode? ComIntNode(object? value)
    {
        if (IsMixed(value)) return null;
        try { return JsonValue.Create(Convert.ToInt32(value, CultureInfo.InvariantCulture)); }
        catch { return null; }
    }

    private static JsonNode? ComStringNode(object? value)
    {
        if (IsMixed(value)) return null;
        return JsonValue.Create(Convert.ToString(value, CultureInfo.InvariantCulture));
    }

    private static object RequireScalar(object? value, string property, string cellRef)
    {
        if (IsMixed(value))
            throw new InvalidOperationException(
                $"[EXCEL_FORMAT_MIXED_RICHTEXT] {cellRef} {property} is mixed or rich-text; null/DBNull is not coerced to a default because exact rollback cannot be proved");
        return value!;
    }

    private static bool RequireBoolean(object? value, string property, string cellRef) =>
        Convert.ToBoolean(RequireScalar(value, property, cellRef), CultureInfo.InvariantCulture);

    private static double RequireDouble(object? value, string property, string cellRef) =>
        Convert.ToDouble(RequireScalar(value, property, cellRef), CultureInfo.InvariantCulture);

    private static int RequireInt(object? value, string property, string cellRef) =>
        Convert.ToInt32(RequireScalar(value, property, cellRef), CultureInfo.InvariantCulture);

    private static string RequireString(object? value, string property, string cellRef) =>
        Convert.ToString(RequireScalar(value, property, cellRef), CultureInfo.InvariantCulture)
        ?? throw new InvalidOperationException($"[EXCEL_FORMAT_MIXED_RICHTEXT] {cellRef} {property} is empty");

    private static bool RequiredBool(JsonObject style, string key) =>
        style[key] is JsonValue value && value.TryGetValue<bool>(out var flag)
            ? flag
            : throw new InvalidOperationException($"format-only style is missing boolean '{key}'");

    private static double RequiredNumber(JsonObject style, string key)
    {
        if (style[key] is not JsonValue value)
            throw new InvalidOperationException($"format-only style is missing number '{key}'");
        if (value.TryGetValue<double>(out var number) && double.IsFinite(number)) return number;
        if (value.TryGetValue<int>(out var i)) return i;
        if (value.TryGetValue<long>(out var l)) return l;
        if (value.TryGetValue<decimal>(out var dec)) return Convert.ToDouble(dec, CultureInfo.InvariantCulture);
        throw new InvalidOperationException($"format-only style '{key}' must be a finite number");
    }

    private static int RequiredInt(JsonObject style, string key)
    {
        if (style[key] is not JsonValue value)
            throw new InvalidOperationException($"format-only style is missing integer '{key}'");
        if (value.TryGetValue<int>(out var i)) return i;
        if (value.TryGetValue<long>(out var l) && l is >= int.MinValue and <= int.MaxValue) return (int)l;
        if (value.TryGetValue<double>(out var d) && double.IsFinite(d) && Math.Abs(d - Math.Round(d)) < 1e-9)
            return Convert.ToInt32(d, CultureInfo.InvariantCulture);
        throw new InvalidOperationException($"format-only style '{key}' must be an integer");
    }

    private static string RequiredText(JsonObject style, string key) =>
        style[key] is JsonValue value && value.TryGetValue<string>(out var text) && text is not null
            ? text
            : throw new InvalidOperationException($"format-only style is missing string '{key}'");

    private static int? OptionalInt(JsonObject style, string key) =>
        style.ContainsKey(key) ? RequiredInt(style, key) : null;

    private static bool TryReadPositiveInt(JsonObject obj, string key, out int value)
    {
        value = 0;
        if (obj[key] is not JsonValue node) return false;
        if (node.TryGetValue<int>(out var i) && i > 0) { value = i; return true; }
        if (node.TryGetValue<long>(out var l) && l > 0 && l <= int.MaxValue) { value = (int)l; return true; }
        if (node.TryGetValue<double>(out var d) && double.IsFinite(d) && d > 0 && Math.Abs(d - Math.Round(d)) < 1e-9)
        {
            value = Convert.ToInt32(d, CultureInfo.InvariantCulture);
            return value > 0;
        }
        return false;
    }

    private static bool TryReadNonNegativeInt(JsonObject obj, string key, out int value)
    {
        value = 0;
        if (obj[key] is not JsonValue node) return false;
        if (node.TryGetValue<int>(out var i) && i >= 0) { value = i; return true; }
        if (node.TryGetValue<long>(out var l) && l >= 0 && l <= int.MaxValue) { value = (int)l; return true; }
        if (node.TryGetValue<double>(out var d) && double.IsFinite(d) && d >= 0 && Math.Abs(d - Math.Round(d)) < 1e-9)
        {
            value = Convert.ToInt32(d, CultureInfo.InvariantCulture);
            return value >= 0;
        }
        return false;
    }

    private static bool NumbersEqual(double left, double right) => Math.Abs(left - right) < 1e-9;

    private static void AssignRangeNumberFormat(object range, string format)
    {
        ExcelNumberFormatContract.Assign(new ComNumberFormatSurface(range), format);
    }

    private sealed class ComNumberFormatSurface : ExcelNumberFormatContract.INumberFormatSurface
    {
        private readonly object _range;

        public ComNumberFormatSurface(object range) => _range = range;

        public void SetNumberFormat(string format) => ((dynamic)_range).NumberFormat = format;

        public void SetNumberFormatLocal(string format) => ((dynamic)_range).NumberFormatLocal = format;

        public string? LocalGeneralName
        {
            get
            {
                object? application = null;
                try
                {
                    application = (object)((dynamic)_range).Application;
                    return Convert.ToString(
                        ((dynamic)application).International[ExcelNumberFormatContract.XlGeneralFormatName],
                        CultureInfo.InvariantCulture);
                }
                finally { RotHelper.ReleaseComReference(application); }
            }
        }
    }
}
