using System.Globalization;
using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class ExcelAdapter
{
    internal const int FormatOnlyScopedSnapshotVersion = 3;
    internal const string WrittenStyleScope = "written-properties";
    internal const string FormatStyleModeUniform = "uniform";
    internal const string FormatStyleModeGroups = "groups";
    internal const string FormatStyleModeCells = "cells";

    private static readonly string[] FastPathFormatKeys =
    {
        ExcelStyleContract.Bold,
        ExcelStyleContract.Italic,
        ExcelStyleContract.FontSize,
        ExcelStyleContract.NumberFormat,
    };

    private static readonly string[] FontColorCoupledKeys =
    {
        ExcelStyleContract.FontColor,
        "fontColorIndex",
        "fontTintAndShade",
    };

    private static readonly string[] FillColorCoupledKeys =
    {
        ExcelStyleContract.FillColor,
        "fillColorIndex",
        "fillTintAndShade",
        "fillPattern",
        "fillPatternColor",
        "fillPatternColorIndex",
        "fillPatternTintAndShade",
    };

    private static readonly HashSet<string> WritableFormatPropertyKeys = new(StringComparer.Ordinal)
    {
        ExcelStyleContract.Bold,
        ExcelStyleContract.Italic,
        ExcelStyleContract.FontSize,
        ExcelStyleContract.NumberFormat,
        ExcelStyleContract.FontColor,
        ExcelStyleContract.FillColor,
    };

    private const int ExcelMaxRow = 1_048_576;
    private const int ExcelMaxColumn = 16_384;

    private static HashSet<string> CollectWrittenFormatScope(JsonObject op)
    {
        var errors = new List<string>();
        if (!ExcelStyleContract.TryNormalize(Json.GetObj(op, "style"), out var canonical, errors))
            throw new InvalidOperationException(string.Join("; ", errors));
        if (canonical.Count == 0)
            throw new InvalidOperationException(
                "format_range style does not name any properties; format-only snapshot refuses an empty scope");

        var written = new HashSet<string>(StringComparer.Ordinal);
        foreach (var (key, _) in canonical)
            written.Add(key);

        return ExactScopedProperties(written);
    }

    private static HashSet<string> ExactScopedProperties(IReadOnlySet<string> written)
    {
        var scoped = new HashSet<string>(written, StringComparer.Ordinal);
        if (scoped.Contains(ExcelStyleContract.FontColor))
        {
            foreach (var key in FontColorCoupledKeys)
                scoped.Add(key);
        }

        if (scoped.Contains(ExcelStyleContract.FillColor))
        {
            foreach (var key in FillColorCoupledKeys)
                scoped.Add(key);
        }

        return scoped;
    }

    private static JsonArray ToScopedPropertyArray(IReadOnlyCollection<string> scoped)
    {
        var properties = new JsonArray();
        foreach (var key in RequiredFormatOnlyStyleKeys)
        {
            if (scoped.Contains(key))
                properties.Add(key);
        }

        return properties;
    }

    private static bool TryReadScopedProperties(JsonObject formatState, out HashSet<string> scoped, out string error)
    {
        scoped = new HashSet<string>(StringComparer.Ordinal);
        error = "";
        if (formatState["scopedProperties"] is not JsonArray listed || listed.Count == 0)
        {
            error = "scoped format state is missing scopedProperties; refusing restore";
            return false;
        }

        var written = new HashSet<string>(StringComparer.Ordinal);
        foreach (var node in listed)
        {
            if (node is not JsonValue value || !value.TryGetValue<string>(out var key) || string.IsNullOrWhiteSpace(key))
            {
                error = "scoped format state scopedProperties must be non-empty strings; refusing restore";
                return false;
            }

            if (!scoped.Add(key))
            {
                error = "scoped format state scopedProperties contains duplicates; refusing restore";
                return false;
            }

            if (WritableFormatPropertyKeys.Contains(key))
                written.Add(key);
            else if (!IsKnownCoupledFormatKey(key))
            {
                error = $"scoped format state scopedProperties contains unsupported '{key}'; refusing restore";
                return false;
            }
        }

        if (written.Count == 0)
        {
            error = "scoped format state scopedProperties does not name any writable format property; refusing restore";
            return false;
        }

        var required = ExactScopedProperties(written);
        if (!scoped.SetEquals(required))
        {
            error = "scoped format state scopedProperties must be exactly the written keys plus required color-coupled fields; refusing restore";
            return false;
        }

        return true;
    }

    private static bool IsKnownCoupledFormatKey(string key) =>
        FontColorCoupledKeys.Contains(key, StringComparer.Ordinal)
        || FillColorCoupledKeys.Contains(key, StringComparer.Ordinal);

    private static bool TryValidateScopedStyleObject(JsonObject style, IReadOnlySet<string> scoped, out string error)
    {
        if (!HasRequiredScopedStyleKeys(style, scoped, out var missing))
        {
            error = $"is missing '{missing}'";
            return false;
        }

        foreach (var (key, node) in style)
        {
            if (node is null)
            {
                error = $"has null '{key}'";
                return false;
            }

            if (!IsAllowedScopedStyleKey(key, scoped))
            {
                error = $"has unexpected '{key}'";
                return false;
            }

            if (!ScopedStyleValueHasExpectedType(key, node))
            {
                error = $"has invalid '{key}'";
                return false;
            }
        }

        error = "";
        return true;
    }

    private static bool IsAllowedScopedStyleKey(string key, IReadOnlySet<string> scoped)
    {
        if (scoped.Contains(key)) return true;
        if (string.Equals(key, "fontThemeColor", StringComparison.Ordinal))
            return scoped.Contains(ExcelStyleContract.FontColor);
        if (string.Equals(key, "fillThemeColor", StringComparison.Ordinal)
            || string.Equals(key, "fillPatternThemeColor", StringComparison.Ordinal))
            return scoped.Contains(ExcelStyleContract.FillColor);
        return false;
    }

    private static bool ScopedStyleValueHasExpectedType(string key, JsonNode node)
    {
        if (node is not JsonValue value) return false;
        switch (key)
        {
            case ExcelStyleContract.Bold:
            case ExcelStyleContract.Italic:
                return value.TryGetValue<bool>(out _);
            case ExcelStyleContract.NumberFormat:
                return value.TryGetValue<string>(out var text) && text is not null;
            case "fontColorIndex":
            case "fillColorIndex":
            case "fillPattern":
            case "fillPatternColorIndex":
            case "fontThemeColor":
            case "fillThemeColor":
            case "fillPatternThemeColor":
                return TryReadIntegerValue(value);
            default:
                return TryReadFiniteNumberValue(value);
        }
    }

    private static bool TryReadIntegerValue(JsonValue value)
    {
        if (value.TryGetValue<int>(out _)) return true;
        if (value.TryGetValue<long>(out var l) && l is >= int.MinValue and <= int.MaxValue) return true;
        if (value.TryGetValue<double>(out var d) && double.IsFinite(d) && Math.Abs(d - Math.Round(d)) < 1e-9)
            return true;
        return false;
    }

    private static bool TryReadFiniteNumberValue(JsonValue value)
    {
        if (value.TryGetValue<double>(out var d)) return double.IsFinite(d);
        if (value.TryGetValue<int>(out _)) return true;
        if (value.TryGetValue<long>(out _)) return true;
        if (value.TryGetValue<decimal>(out _)) return true;
        if (value.TryGetValue<float>(out var f)) return float.IsFinite(f);
        return false;
    }

    private static bool HasRequiredScopedStyleKeys(JsonObject style, IReadOnlySet<string> scoped, out string missing)
    {
        foreach (var key in scoped)
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

    private static bool CanUseUniformRangeFastPath(IReadOnlySet<string> scoped)
    {
        if (scoped.Count == 0) return false;
        foreach (var key in scoped)
        {
            if (!FastPathFormatKeys.Contains(key, StringComparer.Ordinal))
                return false;
        }

        return true;
    }

    private static bool TryCaptureUniformFastPath(
        object area, IReadOnlySet<string> scoped, string areaRef, out JsonObject style)
    {
        style = new JsonObject();
        object? font = null;
        try
        {
            dynamic dynamicArea = area;
            if (scoped.Contains(ExcelStyleContract.Bold) ||
                scoped.Contains(ExcelStyleContract.Italic) ||
                scoped.Contains(ExcelStyleContract.FontSize))
            {
                font = (object)dynamicArea.Font;
            }

            if (scoped.Contains(ExcelStyleContract.Bold))
            {
                var raw = (object?)((dynamic)font!).Bold;
                if (IsMixed(raw)) return false;
                style[ExcelStyleContract.Bold] = RequireBoolean(raw, "Font.Bold", areaRef);
            }

            if (scoped.Contains(ExcelStyleContract.Italic))
            {
                var raw = (object?)((dynamic)font!).Italic;
                if (IsMixed(raw)) return false;
                style[ExcelStyleContract.Italic] = RequireBoolean(raw, "Font.Italic", areaRef);
            }

            if (scoped.Contains(ExcelStyleContract.FontSize))
            {
                var raw = (object?)((dynamic)font!).Size;
                if (IsMixed(raw)) return false;
                style[ExcelStyleContract.FontSize] = RequireDouble(raw, "Font.Size", areaRef);
            }

            if (scoped.Contains(ExcelStyleContract.NumberFormat))
            {
                var raw = (object?)dynamicArea.NumberFormat;
                if (IsMixed(raw)) return false;
                style[ExcelStyleContract.NumberFormat] = RequireString(raw, "NumberFormat", areaRef);
            }

            return true;
        }
        finally
        {
            RotHelper.ReleaseComReference(font);
        }
    }

    private static JsonObject CaptureScopedFormatStyle(object cell, string cellRef, IReadOnlySet<string> scoped)
    {
        object? font = null;
        object? interior = null;
        try
        {
            dynamic dynamicCell = cell;
            var needsFont = scoped.Contains(ExcelStyleContract.Bold)
                || scoped.Contains(ExcelStyleContract.Italic)
                || scoped.Contains(ExcelStyleContract.FontSize)
                || scoped.Contains(ExcelStyleContract.FontColor);
            var needsInterior = scoped.Contains(ExcelStyleContract.FillColor);
            if (needsFont) font = (object)dynamicCell.Font;
            if (needsInterior) interior = (object)dynamicCell.Interior;
            dynamic dynamicFont = font!;
            dynamic dynamicInterior = interior!;
            var style = new JsonObject();

            if (scoped.Contains(ExcelStyleContract.Bold))
                style[ExcelStyleContract.Bold] = RequireBoolean(dynamicFont.Bold, "Font.Bold", cellRef);
            if (scoped.Contains(ExcelStyleContract.Italic))
                style[ExcelStyleContract.Italic] = RequireBoolean(dynamicFont.Italic, "Font.Italic", cellRef);
            if (scoped.Contains(ExcelStyleContract.FontSize))
                style[ExcelStyleContract.FontSize] = RequireDouble(dynamicFont.Size, "Font.Size", cellRef);
            if (scoped.Contains(ExcelStyleContract.NumberFormat))
                style[ExcelStyleContract.NumberFormat] = RequireString(dynamicCell.NumberFormat, "NumberFormat", cellRef);

            if (scoped.Contains(ExcelStyleContract.FontColor))
            {
                style[ExcelStyleContract.FontColor] = RequireDouble(dynamicFont.Color, "Font.Color", cellRef);
                style["fontColorIndex"] = RequireInt(dynamicFont.ColorIndex, "Font.ColorIndex", cellRef);
                style["fontTintAndShade"] = RequireDouble(dynamicFont.TintAndShade, "Font.TintAndShade", cellRef);
                if (TryReadThemeColor(dynamicFont, "Font.ThemeColor", cellRef, out int fontTheme))
                    style["fontThemeColor"] = fontTheme;
                RejectUnsupportedRgbTint(cellRef, "Font.TintAndShade", style, "fontTintAndShade", "fontThemeColor");
            }

            if (scoped.Contains(ExcelStyleContract.FillColor))
            {
                style[ExcelStyleContract.FillColor] = RequireDouble(dynamicInterior.Color, "Interior.Color", cellRef);
                style["fillColorIndex"] = RequireInt(dynamicInterior.ColorIndex, "Interior.ColorIndex", cellRef);
                style["fillTintAndShade"] = RequireDouble(dynamicInterior.TintAndShade, "Interior.TintAndShade", cellRef);
                style["fillPattern"] = RequireInt(dynamicInterior.Pattern, "Interior.Pattern", cellRef);
                style["fillPatternColor"] = RequireDouble(dynamicInterior.PatternColor, "Interior.PatternColor", cellRef);
                style["fillPatternColorIndex"] = RequireInt(dynamicInterior.PatternColorIndex, "Interior.PatternColorIndex", cellRef);
                style["fillPatternTintAndShade"] = RequireDouble(dynamicInterior.PatternTintAndShade, "Interior.PatternTintAndShade", cellRef);
                if (TryReadThemeColor(dynamicInterior, "Interior.ThemeColor", cellRef, out int fillTheme))
                    style["fillThemeColor"] = fillTheme;
                if (TryReadThemeColor(dynamicInterior, "Interior.PatternThemeColor", cellRef, out int patternTheme))
                    style["fillPatternThemeColor"] = patternTheme;
                RejectUnsupportedRgbTint(cellRef, "Interior.TintAndShade", style, "fillTintAndShade", "fillThemeColor");
                RejectUnsupportedRgbTint(cellRef, "Interior.PatternTintAndShade", style, "fillPatternTintAndShade", "fillPatternThemeColor");
            }

            return style;
        }
        finally
        {
            RotHelper.ReleaseComReference(interior);
            RotHelper.ReleaseComReference(font);
        }
    }

    private static void RestoreScopedFormatStyle(object target, JsonObject style, IReadOnlySet<string> scoped)
    {
        object? font = null;
        object? interior = null;
        try
        {
            dynamic dynamicTarget = target;
            var needsFont = scoped.Contains(ExcelStyleContract.Bold)
                || scoped.Contains(ExcelStyleContract.Italic)
                || scoped.Contains(ExcelStyleContract.FontSize)
                || scoped.Contains(ExcelStyleContract.FontColor);
            var needsInterior = scoped.Contains(ExcelStyleContract.FillColor);
            if (needsFont) font = (object)dynamicTarget.Font;
            if (needsInterior) interior = (object)dynamicTarget.Interior;
            dynamic dynamicFont = font!;
            dynamic dynamicInterior = interior!;

            if (scoped.Contains(ExcelStyleContract.Bold) && style.ContainsKey(ExcelStyleContract.Bold))
                dynamicFont.Bold = RequiredBool(style, ExcelStyleContract.Bold);
            if (scoped.Contains(ExcelStyleContract.Italic) && style.ContainsKey(ExcelStyleContract.Italic))
                dynamicFont.Italic = RequiredBool(style, ExcelStyleContract.Italic);
            if (scoped.Contains(ExcelStyleContract.FontSize) && style.ContainsKey(ExcelStyleContract.FontSize))
                dynamicFont.Size = RequiredNumber(style, ExcelStyleContract.FontSize);
            if (scoped.Contains(ExcelStyleContract.NumberFormat) && style.ContainsKey(ExcelStyleContract.NumberFormat))
                dynamicTarget.NumberFormat = RequiredText(style, ExcelStyleContract.NumberFormat);

            if (scoped.Contains(ExcelStyleContract.FontColor) && style.ContainsKey(ExcelStyleContract.FontColor))
            {
                RestoreLinkedColor(
                    dynamicFont,
                    RequiredNumber(style, ExcelStyleContract.FontColor),
                    RequiredInt(style, "fontColorIndex"),
                    RequiredNumber(style, "fontTintAndShade"),
                    OptionalInt(style, "fontThemeColor"));
            }

            if (scoped.Contains(ExcelStyleContract.FillColor) && style.ContainsKey(ExcelStyleContract.FillColor))
            {
                RestoreLinkedColor(
                    dynamicInterior,
                    RequiredNumber(style, ExcelStyleContract.FillColor),
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
        }
        finally
        {
            RotHelper.ReleaseComReference(interior);
            RotHelper.ReleaseComReference(font);
        }
    }

    private static bool ScopedFormatStyleMatches(object target, JsonObject style)
    {
        if (NeedsPerCellColorProof(target, style))
            return EveryCellScopedFormatStyleMatches(target, style);
        return TargetScopedFormatStyleMatches(target, style);
    }

    private static bool NeedsPerCellColorProof(object target, JsonObject style)
    {
        if (!style.ContainsKey(ExcelStyleContract.FontColor) && !style.ContainsKey(ExcelStyleContract.FillColor))
            return false;
        // Unknown shape is not a single-cell proof; aggregate Color=0 stays ambiguous.
        if (!TryReadRangeShape(target, out var rows, out var columns))
            return true;
        return rows > 1 || columns > 1;
    }

    private static bool TryReadRangeShape(object target, out int rows, out int columns)
    {
        rows = 0;
        columns = 0;
        object? rowsObject = null;
        object? columnsObject = null;
        try
        {
            dynamic range = target;
            rowsObject = (object)range.Rows;
            columnsObject = (object)range.Columns;
            rows = Convert.ToInt32(((dynamic)rowsObject).Count, CultureInfo.InvariantCulture);
            columns = Convert.ToInt32(((dynamic)columnsObject).Count, CultureInfo.InvariantCulture);
            return rows >= 1 && columns >= 1;
        }
        catch
        {
            rows = 0;
            columns = 0;
            return false;
        }
        finally
        {
            RotHelper.ReleaseComReference(columnsObject);
            RotHelper.ReleaseComReference(rowsObject);
        }
    }

    private static bool EveryCellScopedFormatStyleMatches(object rangeObject, JsonObject style)
    {
        object? cells = null;
        object? rowsObject = null;
        object? columnsObject = null;
        try
        {
            dynamic range = rangeObject;
            rowsObject = (object)range.Rows;
            columnsObject = (object)range.Columns;
            var rows = Convert.ToInt32(((dynamic)rowsObject).Count, CultureInfo.InvariantCulture);
            var columns = Convert.ToInt32(((dynamic)columnsObject).Count, CultureInfo.InvariantCulture);
            if (rows < 1 || columns < 1) return false;
            cells = (object)range.Cells;
            for (var row = 1; row <= rows; row++)
            {
                for (var col = 1; col <= columns; col++)
                {
                    object? cell = null;
                    try
                    {
                        cell = (object)((dynamic)cells).Item(row, col);
                        if (!TargetScopedFormatStyleMatches(cell, style))
                            return false;
                    }
                    finally { RotHelper.ReleaseComReference(cell); }
                }
            }

            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            RotHelper.ReleaseComReference(cells);
            RotHelper.ReleaseComReference(columnsObject);
            RotHelper.ReleaseComReference(rowsObject);
        }
    }

    private static bool TargetScopedFormatStyleMatches(object target, JsonObject style)
    {
        object? font = null;
        object? interior = null;
        try
        {
            dynamic dynamicTarget = target;
            var needsFont = style.ContainsKey(ExcelStyleContract.Bold)
                || style.ContainsKey(ExcelStyleContract.Italic)
                || style.ContainsKey(ExcelStyleContract.FontSize)
                || style.ContainsKey(ExcelStyleContract.FontColor);
            var needsInterior = style.ContainsKey(ExcelStyleContract.FillColor)
                || style.ContainsKey("fillPattern");
            if (needsFont) font = (object)dynamicTarget.Font;
            if (needsInterior) interior = (object)dynamicTarget.Interior;
            dynamic dynamicFont = font!;
            dynamic dynamicInterior = interior!;

            if (style.ContainsKey(ExcelStyleContract.Bold)
                && RequireBoolean(dynamicFont.Bold, "Font.Bold", "verify") != RequiredBool(style, ExcelStyleContract.Bold))
                return false;
            if (style.ContainsKey(ExcelStyleContract.Italic)
                && RequireBoolean(dynamicFont.Italic, "Font.Italic", "verify") != RequiredBool(style, ExcelStyleContract.Italic))
                return false;
            if (style.ContainsKey(ExcelStyleContract.FontSize)
                && !NumbersEqual(RequireDouble(dynamicFont.Size, "Font.Size", "verify"), RequiredNumber(style, ExcelStyleContract.FontSize)))
                return false;
            if (style.ContainsKey(ExcelStyleContract.NumberFormat)
                && !string.Equals(
                    RequireString(dynamicTarget.NumberFormat, "NumberFormat", "verify"),
                    RequiredText(style, ExcelStyleContract.NumberFormat),
                    StringComparison.Ordinal))
                return false;
            if (style.ContainsKey(ExcelStyleContract.FontColor)
                && !LinkedColorMatches(
                    dynamicFont,
                    RequiredNumber(style, ExcelStyleContract.FontColor),
                    RequiredInt(style, "fontColorIndex"),
                    RequiredNumber(style, "fontTintAndShade"),
                    OptionalInt(style, "fontThemeColor")))
                return false;
            if (style.ContainsKey(ExcelStyleContract.FillColor))
            {
                if (!LinkedColorMatches(
                        dynamicInterior,
                        RequiredNumber(style, ExcelStyleContract.FillColor),
                        RequiredInt(style, "fillColorIndex"),
                        RequiredNumber(style, "fillTintAndShade"),
                        OptionalInt(style, "fillThemeColor")))
                    return false;
                if (!ScalarIntEquals((object?)dynamicInterior.Pattern, RequiredInt(style, "fillPattern")))
                    return false;
                if (!LinkedColorMatches(
                        dynamicInterior,
                        RequiredNumber(style, "fillPatternColor"),
                        RequiredInt(style, "fillPatternColorIndex"),
                        RequiredNumber(style, "fillPatternTintAndShade"),
                        OptionalInt(style, "fillPatternThemeColor"),
                        pattern: true))
                    return false;
            }

            return true;
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

    private static JsonArray GroupUniformStyleRectangles(
        JsonObject[,] styles, int rows, int cols, int startRow, int startCol)
    {
        var used = new bool[rows, cols];
        var groups = new JsonArray();
        var fingerprints = new string[rows, cols];
        for (var row = 0; row < rows; row++)
        {
            for (var col = 0; col < cols; col++)
                fingerprints[row, col] = Json.Canonical(styles[row, col]);
        }

        for (var row = 0; row < rows; row++)
        {
            for (var col = 0; col < cols; col++)
            {
                if (used[row, col]) continue;
                var fingerprint = fingerprints[row, col];
                var width = 1;
                while (col + width < cols
                       && !used[row, col + width]
                       && string.Equals(fingerprints[row, col + width], fingerprint, StringComparison.Ordinal))
                {
                    width++;
                }

                var height = 1;
                while (row + height < rows)
                {
                    var extends = true;
                    for (var offset = 0; offset < width; offset++)
                    {
                        if (used[row + height, col + offset]
                            || !string.Equals(fingerprints[row + height, col + offset], fingerprint, StringComparison.Ordinal))
                        {
                            extends = false;
                            break;
                        }
                    }

                    if (!extends) break;
                    height++;
                }

                for (var y = 0; y < height; y++)
                for (var x = 0; x < width; x++)
                    used[row + y, col + x] = true;

                var groupRow = startRow + row;
                var groupCol = startCol + col;
                groups.Add(new JsonObject
                {
                    ["range"] = FormatRangeAddress(groupRow, groupCol, height, width),
                    ["row"] = groupRow,
                    ["column"] = groupCol,
                    ["rows"] = height,
                    ["columns"] = width,
                    ["style"] = styles[row, col].DeepClone(),
                });
            }
        }

        return groups;
    }

    private static string FormatRangeAddress(int row, int column, int rows, int columns)
    {
        if (!TryGetBoundedExcelRange(row, column, rows, columns, out var endRow, out var endColumn, out var error))
            throw new InvalidOperationException($"format range {error}");
        var start = CellName(column, row);
        return rows == 1 && columns == 1
            ? start
            : $"{start}:{CellName(endColumn, endRow)}";
    }

    private static bool TryGetBoundedExcelRange(
        int row, int column, int rows, int columns, out int endRow, out int endColumn, out string error)
    {
        endRow = 0;
        endColumn = 0;
        error = "";
        if (row is < 1 or > ExcelMaxRow || column is < 1 or > ExcelMaxColumn || rows < 1 || columns < 1)
        {
            error = "is outside Excel bounds";
            return false;
        }

        try
        {
            checked
            {
                endRow = row + rows - 1;
                endColumn = column + columns - 1;
            }
        }
        catch (OverflowException)
        {
            error = "overflows Excel coordinates";
            return false;
        }

        if (endRow > ExcelMaxRow || endColumn > ExcelMaxColumn)
        {
            error = "is outside Excel bounds";
            return false;
        }

        return true;
    }

    private static bool TryParseA1Cell(string token, out int row, out int column)
    {
        row = 0;
        column = 0;
        if (string.IsNullOrWhiteSpace(token)) return false;

        var letters = new System.Text.StringBuilder();
        var digits = new System.Text.StringBuilder();
        var seenDigit = false;
        foreach (var ch in token.Trim())
        {
            if (ch == '$') continue;
            if (!seenDigit && char.IsAsciiLetter(ch))
            {
                letters.Append(ch);
                continue;
            }

            if (char.IsAsciiDigit(ch))
            {
                seenDigit = true;
                digits.Append(ch);
                continue;
            }

            return false;
        }

        if (letters.Length is < 1 or > 3 || digits.Length == 0) return false;
        if (!int.TryParse(digits.ToString(), NumberStyles.None, CultureInfo.InvariantCulture, out row))
            return false;
        if (row is < 1 or > ExcelMaxRow) return false;
        column = ColIndex(letters.ToString());
        if (column is < 1 or > ExcelMaxColumn) return false;
        return string.Equals(ColName(column), letters.ToString(), StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryParseA1Range(string? address, out int row, out int column, out int rows, out int columns)
    {
        row = 0;
        column = 0;
        rows = 0;
        columns = 0;
        if (string.IsNullOrWhiteSpace(address)) return false;

        var text = address.Trim();
        if (text.Contains('!', StringComparison.Ordinal)
            || text.Contains(',', StringComparison.Ordinal)
            || text.Contains(';', StringComparison.Ordinal)
            || text.Contains(' ', StringComparison.Ordinal))
        {
            return false;
        }

        var parts = text.Split(':');
        if (parts.Length is < 1 or > 2) return false;
        if (!TryParseA1Cell(parts[0], out var startRow, out var startCol)) return false;
        var endRow = startRow;
        var endCol = startCol;
        if (parts.Length == 2 && !TryParseA1Cell(parts[1], out endRow, out endCol)) return false;
        if (endRow < startRow || endCol < startCol) return false;

        try
        {
            checked
            {
                rows = endRow - startRow + 1;
                columns = endCol - startCol + 1;
            }
        }
        catch (OverflowException)
        {
            return false;
        }

        row = startRow;
        column = startCol;
        return rows >= 1 && columns >= 1;
    }

    private static bool TryValidateDeclaredA1Range(
        string address, int row, int column, int rows, int columns, out string error)
    {
        if (!TryGetBoundedExcelRange(row, column, rows, columns, out _, out _, out error))
            return false;

        if (!TryParseA1Range(address, out var parsedRow, out var parsedCol, out var parsedRows, out var parsedCols)
            || parsedRow != row
            || parsedCol != column
            || parsedRows != rows
            || parsedCols != columns)
        {
            error = $"parsed range '{address}' does not match row/column/rows/columns {row},{column},{rows},{columns}";
            return false;
        }

        error = "";
        return true;
    }

    private static JsonObject CreateScopedFormatState(
        string sheetName,
        string areaAddress,
        int startRow,
        int startCol,
        int rows,
        int cols,
        IReadOnlyCollection<string> scoped,
        string styleMode,
        JsonObject? uniformStyle,
        JsonArray? groups,
        JsonArray? cellStyles)
    {
        var state = new JsonObject
        {
            ["sheet"] = sheetName,
            ["range"] = areaAddress,
            ["row"] = startRow,
            ["column"] = startCol,
            ["rows"] = rows,
            ["columns"] = cols,
            ["styleMode"] = styleMode,
            ["styleScope"] = WrittenStyleScope,
            ["scopedProperties"] = ToScopedPropertyArray(scoped),
        };
        if (uniformStyle is not null) state["style"] = uniformStyle;
        if (groups is not null) state["groups"] = groups;
        if (cellStyles is not null) state["styles"] = cellStyles;
        return state;
    }

    private static void MarkFormatCoverage(
        string sheetName,
        int startRow,
        int startCol,
        int rows,
        int cols,
        Dictionary<string, int> occupied,
        ref int overlapping)
    {
        for (var row = 0; row < rows; row++)
        {
            for (var col = 0; col < cols; col++)
            {
                var cellRef = $"{sheetName}!{CellName(startCol + col, startRow + row)}";
                if (occupied.TryGetValue(cellRef, out var seen))
                {
                    overlapping++;
                    occupied[cellRef] = seen + 1;
                }
                else occupied[cellRef] = 1;
            }
        }
    }

    private static void RejectPartialMergesInArea(
        object areaObject,
        string sheetName,
        string areaAddress,
        int startRow,
        int startCol,
        int rows,
        int cols)
    {
        var mergeFlag = SafeGet(() => (object?)((dynamic)areaObject).MergeCells);
        if (!IsMixed(mergeFlag) && !Convert.ToBoolean(mergeFlag, CultureInfo.InvariantCulture))
            return;

        // Range.MergeArea is documented for a single-cell range only.
        // https://learn.microsoft.com/en-us/office/vba/api/excel.range.mergearea
        // MergeCells=true does not mean the target is one MergeArea.
        if (rows == 1 && cols == 1)
        {
            RejectPartialMergeCoverage(areaObject, $"{sheetName}!{areaAddress}", startRow, startCol, rows, cols);
            return;
        }

        if (TrySingleCellMergeCoversWholeArea(areaObject, sheetName, startRow, startCol, rows, cols))
            return;

        object? cells = null;
        try
        {
            cells = (object)((dynamic)areaObject).Cells;
            for (var row = 1; row <= rows; row++)
            {
                for (var col = 1; col <= cols; col++)
                {
                    object? cell = null;
                    try
                    {
                        cell = (object)((dynamic)cells).Item(row, col);
                        var cellRef = $"{sheetName}!{CellName(startCol + col - 1, startRow + row - 1)}";
                        RejectPartialMergeCoverage(cell, cellRef, startRow, startCol, rows, cols);
                    }
                    finally { RotHelper.ReleaseComReference(cell); }
                }
            }
        }
        finally { RotHelper.ReleaseComReference(cells); }
    }

    private static bool TrySingleCellMergeCoversWholeArea(
        object areaObject, string sheetName, int startRow, int startCol, int rows, int cols)
    {
        object? cells = null;
        object? cell = null;
        object? mergeArea = null;
        object? mergeRows = null;
        object? mergeColumns = null;
        try
        {
            cells = (object)((dynamic)areaObject).Cells;
            cell = (object)((dynamic)cells).Item(1, 1);
            var cellRef = $"{sheetName}!{CellName(startCol, startRow)}";
            var mergeCells = RequireScalar(((dynamic)cell).MergeCells, "MergeCells", cellRef);
            if (!Convert.ToBoolean(mergeCells, CultureInfo.InvariantCulture))
                return false;

            mergeArea = (object)((dynamic)cell).MergeArea;
            dynamic area = mergeArea;
            mergeRows = (object)area.Rows;
            mergeColumns = (object)area.Columns;
            var mergeRow = Convert.ToInt32(area.Row, CultureInfo.InvariantCulture);
            var mergeCol = Convert.ToInt32(area.Column, CultureInfo.InvariantCulture);
            var mergeRowCount = Convert.ToInt32(((dynamic)mergeRows).Count, CultureInfo.InvariantCulture);
            var mergeColCount = Convert.ToInt32(((dynamic)mergeColumns).Count, CultureInfo.InvariantCulture);
            return mergeRow == startRow
                && mergeCol == startCol
                && mergeRowCount == rows
                && mergeColCount == cols;
        }
        catch (InvalidOperationException ex) when (ex.Message.StartsWith("[EXCEL_FORMAT_", StringComparison.Ordinal))
        {
            throw;
        }
        catch
        {
            return false;
        }
        finally
        {
            RotHelper.ReleaseComReference(mergeColumns);
            RotHelper.ReleaseComReference(mergeRows);
            RotHelper.ReleaseComReference(mergeArea);
            RotHelper.ReleaseComReference(cell);
            RotHelper.ReleaseComReference(cells);
        }
    }

    private static bool TryValidateScopedFormatStateEntry(
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

        string sheetName = sheet;
        string rangeAddress = address;

        if (!TryReadPositiveInt(formatState, "row", out var startRow) ||
            !TryReadPositiveInt(formatState, "column", out var startCol) ||
            !TryReadPositiveInt(formatState, "rows", out var rows) ||
            !TryReadPositiveInt(formatState, "columns", out var columns))
        {
            error = $"format-only snapshot formatStates[{index}] ({sheetName}!{rangeAddress}) is missing dimensions; refusing restore";
            return false;
        }

        if (!TryReadScopedProperties(formatState, out var scoped, out error))
            return false;

        var areaCells = (long)rows * columns;
        if (areaCells > MaxFormatSnapshotCells)
        {
            error = $"format-only snapshot formatStates[{index}] ({sheetName}!{rangeAddress}) exceeds {MaxFormatSnapshotCells} cells; refusing restore";
            return false;
        }

        if (!TryValidateDeclaredA1Range(rangeAddress, startRow, startCol, rows, columns, out var rangeError))
        {
            error = $"format-only snapshot formatStates[{index}] ({sheetName}!{rangeAddress}) {rangeError}; refusing restore";
            return false;
        }

        if (!TryGetBoundedExcelRange(startRow, startCol, rows, columns, out var areaEndRow, out var areaEndCol, out var boundError))
        {
            error = $"format-only snapshot formatStates[{index}] ({sheetName}!{rangeAddress}) {boundError}; refusing restore";
            return false;
        }

        var mode = Json.GetString(formatState, "styleMode");
        if (string.Equals(mode, FormatStyleModeUniform, StringComparison.Ordinal))
        {
            if (formatState["style"] is not JsonObject style)
            {
                error = $"format-only snapshot formatStates[{index}] ({sheetName}!{rangeAddress}) uniform style is missing; refusing restore";
                return false;
            }

            if (!TryValidateScopedStyleObject(style, scoped, out var styleError))
            {
                error = $"format-only snapshot formatStates[{index}] ({sheetName}!{rangeAddress}) uniform style {styleError}; refusing restore";
                return false;
            }

            cells = rows * columns;
            return true;
        }

        if (string.Equals(mode, FormatStyleModeGroups, StringComparison.Ordinal))
        {
            if (formatState["groups"] is not JsonArray groups || groups.Count == 0)
            {
                error = $"format-only snapshot formatStates[{index}] ({sheetName}!{rangeAddress}) groups are missing; refusing restore";
                return false;
            }

            var covered = new HashSet<(int Row, int Column)>();
            for (var groupIndex = 0; groupIndex < groups.Count; groupIndex++)
            {
                if (groups[groupIndex] is not JsonObject group)
                {
                    error = $"format-only snapshot formatStates[{index}] ({sheetName}!{rangeAddress}) groups[{groupIndex}] is not an object; refusing restore";
                    return false;
                }

                var groupRange = Json.GetString(group, "range");
                if (!TryReadPositiveInt(group, "row", out var groupRow) ||
                    !TryReadPositiveInt(group, "column", out var groupCol) ||
                    !TryReadPositiveInt(group, "rows", out var groupRows) ||
                    !TryReadPositiveInt(group, "columns", out var groupCols) ||
                    string.IsNullOrWhiteSpace(groupRange))
                {
                    error = $"format-only snapshot formatStates[{index}] ({sheetName}!{rangeAddress}) groups[{groupIndex}] is missing dimensions; refusing restore";
                    return false;
                }

                string groupAddress = groupRange;
                if (!TryValidateDeclaredA1Range(groupAddress, groupRow, groupCol, groupRows, groupCols, out var groupRangeError))
                {
                    error = $"format-only snapshot formatStates[{index}] ({sheetName}!{rangeAddress}) groups[{groupIndex}] {groupRangeError}; refusing restore";
                    return false;
                }

                if (!TryGetBoundedExcelRange(groupRow, groupCol, groupRows, groupCols, out var groupEndRow, out var groupEndCol, out var groupBoundError))
                {
                    error = $"format-only snapshot formatStates[{index}] ({sheetName}!{rangeAddress}) groups[{groupIndex}] {groupBoundError}; refusing restore";
                    return false;
                }

                if (groupRow < startRow || groupCol < startCol || groupEndRow > areaEndRow || groupEndCol > areaEndCol)
                {
                    error = $"format-only snapshot formatStates[{index}] ({sheetName}!{rangeAddress}) groups[{groupIndex}] extends outside the captured range; refusing restore";
                    return false;
                }

                if (group["style"] is not JsonObject style)
                {
                    error = $"format-only snapshot formatStates[{index}] ({sheetName}!{rangeAddress}) groups[{groupIndex}] is missing style; refusing restore";
                    return false;
                }

                if (!TryValidateScopedStyleObject(style, scoped, out var styleError))
                {
                    error = $"format-only snapshot formatStates[{index}] ({sheetName}!{rangeAddress}) groups[{groupIndex}] {styleError}; refusing restore";
                    return false;
                }

                for (var row = 0; row < groupRows; row++)
                {
                    for (var col = 0; col < groupCols; col++)
                    {
                        var absoluteRow = groupRow + row;
                        var absoluteCol = groupCol + col;
                        if (absoluteRow < startRow || absoluteCol < startCol
                            || absoluteRow > areaEndRow || absoluteCol > areaEndCol)
                        {
                            error = $"format-only snapshot formatStates[{index}] ({sheetName}!{rangeAddress}) groups[{groupIndex}] extends outside the captured range; refusing restore";
                            return false;
                        }

                        if (!covered.Add((absoluteRow, absoluteCol)))
                        {
                            error = $"format-only snapshot formatStates[{index}] ({sheetName}!{rangeAddress}) groups overlap; refusing restore";
                            return false;
                        }
                    }
                }
            }

            if (covered.Count != rows * columns)
            {
                error = $"format-only snapshot formatStates[{index}] ({sheetName}!{rangeAddress}) groups cover {covered.Count} cells != {rows * columns}; refusing restore";
                return false;
            }

            cells = rows * columns;
            return true;
        }

        if (string.Equals(mode, FormatStyleModeCells, StringComparison.Ordinal)
            || formatState["styles"] is JsonArray)
        {
            if (formatState["styles"] is not JsonArray styleRows)
            {
                error = $"format-only snapshot formatStates[{index}] ({sheetName}!{rangeAddress}) is missing styles; refusing restore";
                return false;
            }

            if (styleRows.Count != rows)
            {
                error = $"format-only snapshot formatStates[{index}] ({sheetName}!{rangeAddress}) styles rows {styleRows.Count} != {rows}; refusing restore";
                return false;
            }

            for (var row = 0; row < styleRows.Count; row++)
            {
                if (styleRows[row] is not JsonArray styleCols || styleCols.Count != columns)
                {
                    error = $"format-only snapshot formatStates[{index}] ({sheetName}!{rangeAddress}) styles[{row}] does not match columns {columns}; refusing restore";
                    return false;
                }

                for (var col = 0; col < styleCols.Count; col++)
                {
                    if (styleCols[col] is not JsonObject style)
                    {
                        error = $"format-only snapshot formatStates[{index}] ({sheetName}!{rangeAddress}) styles[{row}][{col}] is not an object; refusing restore";
                        return false;
                    }

                    if (!TryValidateScopedStyleObject(style, scoped, out var styleError))
                    {
                        error = $"format-only snapshot formatStates[{index}] ({sheetName}!{rangeAddress}) styles[{row}][{col}] {styleError}; refusing restore";
                        return false;
                    }
                }
            }

            cells = rows * columns;
            return true;
        }

        error = $"format-only snapshot formatStates[{index}] ({sheetName}!{rangeAddress}) has an unsupported styleMode; refusing restore";
        return false;
    }

    private static void ApplyScopedFormatStates(
        object workbook, JsonArray formatStates, RestoreMismatchCollector mismatches, ref int checkedCells)
    {
        foreach (var formatNode in formatStates)
        {
            var formatState = (JsonObject)formatNode!;
            if (!TryReadScopedProperties(formatState, out var scoped, out var scopedError))
                throw new InvalidOperationException(scopedError);

            var sheetName = Json.GetString(formatState, "sheet")!;
            var startRow = RequiredInt(formatState, "row");
            var startCol = RequiredInt(formatState, "column");
            var rows = RequiredInt(formatState, "rows");
            var columns = RequiredInt(formatState, "columns");
            var mode = Json.GetString(formatState, "styleMode");
            dynamic sheet = GetSheet(workbook, sheetName);

            if (string.Equals(mode, FormatStyleModeUniform, StringComparison.Ordinal))
            {
                ApplyScopedRangeStyle(
                    sheet, sheetName, startRow, startCol, rows, columns,
                    (JsonObject)formatState["style"]!, scoped, mismatches, ref checkedCells);
                continue;
            }

            if (string.Equals(mode, FormatStyleModeGroups, StringComparison.Ordinal))
            {
                foreach (var groupNode in (JsonArray)formatState["groups"]!)
                {
                    var group = (JsonObject)groupNode!;
                    ApplyScopedRangeStyle(
                        sheet,
                        sheetName,
                        RequiredInt(group, "row"),
                        RequiredInt(group, "column"),
                        RequiredInt(group, "rows"),
                        RequiredInt(group, "columns"),
                        (JsonObject)group["style"]!,
                        scoped,
                        mismatches,
                        ref checkedCells);
                }

                continue;
            }

            ApplyScopedCellStyles(
                sheet, sheetName, startRow, startCol, rows, columns,
                (JsonArray)formatState["styles"]!, scoped, mismatches, ref checkedCells);
        }
    }

    private static void ApplyScopedRangeStyle(
        dynamic sheet,
        string sheetName,
        int row,
        int column,
        int rows,
        int columns,
        JsonObject style,
        IReadOnlySet<string> scoped,
        RestoreMismatchCollector mismatches,
        ref int checkedCells)
    {
        var address = FormatRangeAddress(row, column, rows, columns);
        object? rangeObject = null;
        try
        {
            rangeObject = (object)sheet.Range(address);
            RestoreScopedFormatStyle(rangeObject, style, scoped);
            checkedCells += rows * columns;
            if (!ScopedFormatStyleMatches(rangeObject, style))
                mismatches.Add($"{sheetName}!{address}: scoped style restore mismatch");
        }
        finally
        {
            RotHelper.ReleaseComReference(rangeObject);
        }
    }

    private static void ApplyScopedCellStyles(
        dynamic sheet,
        string sheetName,
        int startRow,
        int startCol,
        int rows,
        int columns,
        JsonArray styleRows,
        IReadOnlySet<string> scoped,
        RestoreMismatchCollector mismatches,
        ref int checkedCells)
    {
        var address = FormatRangeAddress(startRow, startCol, rows, columns);
        object? rangeObject = null;
        object? cells = null;
        try
        {
            rangeObject = (object)sheet.Range(address);
            dynamic range = rangeObject;
            cells = (object)range.Cells;
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
                        RestoreScopedFormatStyle(cell, style, scoped);
                        checkedCells++;
                        if (!ScopedFormatStyleMatches(cell, style))
                            mismatches.Add($"{sheetName}!{CellName(startCol + col, startRow + row)}: scoped style restore mismatch");
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
