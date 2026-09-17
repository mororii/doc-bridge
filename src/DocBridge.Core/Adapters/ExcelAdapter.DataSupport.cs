using System.Globalization;
using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class ExcelAdapter
{
    private static JsonArray ReadAutoFilterCriteria(object? filter, int count)
    {
        var items = new JsonArray();
        if (filter is null || count <= 0) return items;
        object? filters = null;
        try
        {
            filters = (object)((dynamic)filter).Filters;
            for (var index = 1; index <= count && index <= 64; index++)
            {
                object? item = null;
                try
                {
                    item = (object)((dynamic)filters).Item(index);
                    items.Add(new JsonObject
                    {
                        ["field"] = index,
                        ["on"] = Js(TryComBool(item, "On")),
                        ["operator"] = Js(TryComInt(item, "Operator")),
                        ["criteria1"] = ReadFilterCriteriaNode(item, "Criteria1"),
                        ["criteria2"] = ReadFilterCriteriaNode(item, "Criteria2"),
                    });
                }
                catch
                {
                    items.Add(new JsonObject { ["field"] = index, ["on"] = false });
                }
                finally { RotHelper.ReleaseComReference(item); }
            }
        }
        catch { /* Filters collection unavailable */ }
        finally { RotHelper.ReleaseComReference(filters); }
        return items;
    }

    private static JsonNode? ReadFilterCriteriaNode(object filter, string property)
    {
        try
        {
            object? raw = property == "Criteria2" ? ((dynamic)filter).Criteria2 : ((dynamic)filter).Criteria1;
            if (raw is null) return null;
            if (raw is Array array)
            {
                var values = new JsonArray();
                foreach (var item in array)
                    values.Add(Convert.ToString(item, CultureInfo.InvariantCulture));
                return values;
            }
            return JsonValue.Create(Convert.ToString(raw, CultureInfo.InvariantCulture));
        }
        catch { return null; }
    }

    private static JsonObject ReadOneConditionalRule(object item, int index)
    {
        var type = TryComInt(item, "Type");
        var rule = new JsonObject
        {
            ["index"] = index,
            ["conditionType"] = Js(type),
            ["operator"] = Js(TryComInt(item, "Operator")),
            ["formula1"] = TryComString(item, "Formula1"),
            ["formula2"] = TryComString(item, "Formula2"),
            ["style"] = ReadConditionalStyle(item),
        };
        if (type == ExcelDataObjectCatalog.XlUniqueValues)
            rule["dupeUnique"] = Js(TryComInt(item, "DupeUnique"));
        if (type == ExcelDataObjectCatalog.XlIconSet)
        {
            object? icon = null;
            try
            {
                icon = (object)((dynamic)item).IconSet;
                rule["iconSet"] = ExcelDataObjectCatalog.TokenFor(
                    ExcelDataObjectCatalog.IconSets, TryComInt(icon, "ID") ?? 0, "arrows3");
            }
            catch { /* icon set details unavailable */ }
            finally { RotHelper.ReleaseComReference(icon); }
        }
        return rule;
    }

    private static string ConditionalFingerprint(JsonObject rule) =>
        ExcelConditionalFormatContract.Fingerprint(rule);

    private static JsonObject ReadConditionalStyle(object condition)
    {
        object? font = null;
        object? interior = null;
        try
        {
            try { font = (object)((dynamic)condition).Font; } catch { font = null; }
            try { interior = (object)((dynamic)condition).Interior; } catch { interior = null; }
            return new JsonObject
            {
                ["bold"] = Js(font is null ? null : TryComBool(font, "Bold")),
                ["italic"] = Js(font is null ? null : TryComBool(font, "Italic")),
                ["fontColor"] = Js(font is null ? null : TryComDouble(font, "Color")),
                ["fillColor"] = Js(interior is null ? null : TryComDouble(interior, "Color")),
            };
        }
        finally
        {
            RotHelper.ReleaseComReference(interior);
            RotHelper.ReleaseComReference(font);
        }
    }

    private static JsonArray ReadChartSeries(object chart, out int seriesCount, out string? firstFormula)
    {
        seriesCount = 0;
        firstFormula = null;
        var list = new JsonArray();
        object? series = null;
        try
        {
            series = (object)((dynamic)chart).SeriesCollection();
            seriesCount = Convert.ToInt32(((dynamic)series).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= seriesCount && index <= 64; index++)
            {
                object? item = null;
                try
                {
                    item = (object)((dynamic)series).Item(index);
                    var formula = ReadComString(item, "Formula");
                    if (index == 1) firstFormula = formula;
                    var seriesType = TryComInt(item, "ChartType");
                    var axisGroup = TryComInt(item, "AxisGroup");
                    list.Add(new JsonObject
                    {
                        ["index"] = index,
                        ["formula"] = formula,
                        ["name"] = TryComString(item, "Name"),
                        ["values"] = TryComString(item, "Values") ?? formula,
                        ["categories"] = TryComString(item, "XValues"),
                        ["chartType"] = seriesType is null
                            ? null
                            : ExcelDataObjectCatalog.TokenFor(
                                ExcelDataObjectCatalog.ChartTypes, seriesType.Value, $"raw:{seriesType}"),
                        ["axisGroup"] = ExcelDataObjectCatalog.TokenFor(
                            ExcelDataObjectCatalog.AxisGroups,
                            axisGroup ?? ExcelDataObjectCatalog.XlPrimary,
                            axisGroup is null ? "primary" : $"raw:{axisGroup}"),
                        ["marker"] = ExcelDataObjectCatalog.TokenFor(
                            ExcelDataObjectCatalog.MarkerStyles,
                            TryComInt(item, "MarkerStyle") ?? ExcelDataObjectCatalog.XlMarkerStyleAutomatic,
                            "automatic"),
                        ["lineWeight"] = Js(ReadSeriesLineWeight(item)),
                        ["lineColor"] = Js(ReadSeriesLineColor(item)),
                        ["dataLabels"] = ReadChartDataLabels(item),
                        ["trendlines"] = ReadTrendlines(item),
                        ["points"] = ReadChartPoints(item),
                    });
                }
                finally { RotHelper.ReleaseComReference(item); }
            }
        }
        catch { /* chart has no series yet */ }
        finally { RotHelper.ReleaseComReference(series); }
        return list;
    }

    private static double? ReadSeriesLineWeight(object series)
    {
        object? format = null;
        object? line = null;
        try
        {
            format = (object)((dynamic)series).Format;
            line = (object)((dynamic)format).Line;
            return TryComDouble(line, "Weight");
        }
        catch { return null; }
        finally
        {
            RotHelper.ReleaseComReference(line);
            RotHelper.ReleaseComReference(format);
        }
    }

    private static double? ReadSeriesLineColor(object series)
    {
        object? format = null;
        object? line = null;
        object? color = null;
        try
        {
            format = (object)((dynamic)series).Format;
            line = (object)((dynamic)format).Line;
            color = (object)((dynamic)line).ForeColor;
            return TryComDouble(color, "RGB");
        }
        catch { return null; }
        finally
        {
            RotHelper.ReleaseComReference(color);
            RotHelper.ReleaseComReference(line);
            RotHelper.ReleaseComReference(format);
        }
    }

    private static JsonObject ReadCellContent(object range)
    {
        object? font = null;
        try
        {
            font = (object)((dynamic)range).Font;
            return new JsonObject
            {
                ["cellFormula"] = TryComString(range, "Formula"),
                ["cellValue"] = ToJsonValue(((dynamic)range).Value2),
                ["font"] = new JsonObject
                {
                    ["bold"] = Js(TryComBool(font, "Bold")),
                    ["color"] = Js(TryComDouble(font, "Color")),
                },
            };
        }
        catch
        {
            return new JsonObject { ["cellFormula"] = TryComString(range, "Formula") };
        }
        finally { RotHelper.ReleaseComReference(font); }
    }

    private static JsonObject CaptureRangeSurface(object sheet, string address)
    {
        object? range = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            var formulas = RangeToJson(range, out var cells, formulas: true, maxCells: MaxDataSnapshotCells);
            var values = RangeToJson(range, out _, formulas: false, maxCells: MaxDataSnapshotCells);
            var recovery = ExcelDataRecoveryContract.SurfaceRecoveryFields(DataSnapshotBackupMetadata);
            var surface = new JsonObject
            {
                ["formulas"] = formulas,
                ["values"] = values,
                ["numberFormats"] = CaptureNumberFormats(range),
                ["notes"] = CaptureNotesInRange(sheet, address),
                ["hyperlinks"] = CaptureHyperlinksInRange(range, address),
                ["cellCount"] = cells,
                ["wideSideEffect"] = true,
                ["uncaptured"] = new JsonArray("borders", "alignment", "fontName", "fontSize"),
                ["recoveryAvailable"] = recovery["recoveryAvailable"]?.DeepClone(),
                ["recoveryArtifact"] = recovery["recoveryArtifact"]?.DeepClone(),
                ["recoverySource"] = recovery["recoverySource"]?.DeepClone(),
            };
            if (recovery["recoveryReason"] is JsonNode reason)
                surface["recoveryReason"] = reason.DeepClone();
            return surface;
        }
        finally { RotHelper.ReleaseComReference(range); }
    }

    private static void MergeSurface(JsonObject entry, JsonObject surface)
    {
        foreach (var (key, value) in surface)
            entry[key] = value?.DeepClone();
    }

    private static JsonNode CaptureNumberFormats(object range)
    {
        try
        {
            object? raw = ((dynamic)range).NumberFormat;
            if (raw is null)
                return CaptureNumberFormatsPerCell(range);
            if (raw is object[,] arr)
            {
                var rows = new JsonArray();
                var r1 = arr.GetLowerBound(0); var r2 = arr.GetUpperBound(0);
                var c1 = arr.GetLowerBound(1); var c2 = arr.GetUpperBound(1);
                for (var r = r1; r <= r2; r++)
                {
                    var row = new JsonArray();
                    for (var c = c1; c <= c2; c++)
                        row.Add(Convert.ToString(arr[r, c], CultureInfo.InvariantCulture) ?? "General");
                    rows.Add(row);
                }
                return rows;
            }
            var scalar = Convert.ToString(raw, CultureInfo.InvariantCulture);
            if (string.IsNullOrEmpty(scalar))
                return CaptureNumberFormatsPerCell(range);
            return new JsonArray { new JsonArray { scalar } };
        }
        catch
        {
            return CaptureNumberFormatsPerCell(range);
        }
    }

    private static JsonArray CaptureNumberFormatsPerCell(object range)
    {
        var rows = new JsonArray();
        object? rowsObj = null;
        object? colsObj = null;
        try
        {
            rowsObj = (object)((dynamic)range).Rows;
            colsObj = (object)((dynamic)range).Columns;
            var rowCount = Math.Max(1, Convert.ToInt32(((dynamic)rowsObj).Count, CultureInfo.InvariantCulture));
            var colCount = Math.Max(1, Convert.ToInt32(((dynamic)colsObj).Count, CultureInfo.InvariantCulture));
            if (rowCount * colCount > MaxDataSnapshotCells)
            {
                rowCount = Math.Max(1, MaxDataSnapshotCells / Math.Max(1, colCount));
            }
            for (var r = 1; r <= rowCount; r++)
            {
                var row = new JsonArray();
                for (var c = 1; c <= colCount; c++)
                {
                    object? cell = null;
                    try
                    {
                        cell = (object)((dynamic)range).Cells[r, c];
                        row.Add(TryComString(cell, "NumberFormat") ?? "General");
                    }
                    finally { RotHelper.ReleaseComReference(cell); }
                }
                rows.Add(row);
            }
            return rows;
        }
        finally
        {
            RotHelper.ReleaseComReference(colsObj);
            RotHelper.ReleaseComReference(rowsObj);
        }
    }

    private static JsonArray CaptureNotesInRange(object sheet, string address)
    {
        var notes = new JsonArray();
        if (!ExcelDataOperationsContract.TryParseA1(address, out var box)) return notes;
        object? comments = null;
        try
        {
            comments = (object)((dynamic)sheet).Comments;
            var count = Convert.ToInt32(((dynamic)comments).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count && notes.Count < 256; index++)
            {
                object? comment = null;
                object? parent = null;
                try
                {
                    comment = (object)((dynamic)comments).Item(index);
                    parent = (object)((dynamic)comment).Parent;
                    var cell = TryAddress(parent);
                    if (string.IsNullOrWhiteSpace(cell) || !ExcelDataOperationsContract.TryParseA1(cell, out var cellBox))
                        continue;
                    if (cellBox.Row < box.Row || cellBox.EndRow > box.EndRow ||
                        cellBox.Column < box.Column || cellBox.EndColumn > box.EndColumn)
                        continue;
                    notes.Add(new JsonObject
                    {
                        ["range"] = cell,
                        ["text"] = TryInvokeText(comment),
                        ["visible"] = Js(TryComBool(comment, "Visible")),
                    });
                }
                catch { /* skip one comment */ }
                finally
                {
                    RotHelper.ReleaseComReference(parent);
                    RotHelper.ReleaseComReference(comment);
                }
            }
        }
        catch { /* sheet has no Comments collection */ }
        finally { RotHelper.ReleaseComReference(comments); }
        return notes;
    }

    private static JsonArray CaptureHyperlinksInRange(object range, string address)
    {
        var links = new JsonArray();
        object? hyperlinks = null;
        try
        {
            hyperlinks = (object)((dynamic)range).Hyperlinks;
            var count = Convert.ToInt32(((dynamic)hyperlinks).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count && links.Count < 256; index++)
            {
                object? item = null;
                try
                {
                    item = (object)((dynamic)hyperlinks).Item(index);
                    links.Add(new JsonObject
                    {
                        ["range"] = TryAddress((object?)((dynamic)item).Range) ?? address,
                        ["address"] = TryComString(item, "Address"),
                        ["subAddress"] = TryComString(item, "SubAddress"),
                        ["textToDisplay"] = TryComString(item, "TextToDisplay"),
                        ["screenTip"] = TryComString(item, "ScreenTip"),
                    });
                }
                catch { /* skip one link */ }
                finally { RotHelper.ReleaseComReference(item); }
            }
        }
        catch { /* range has no Hyperlinks */ }
        finally { RotHelper.ReleaseComReference(hyperlinks); }
        return links;
    }

    private static void RestoreRangeSurface(object sheet, JsonObject entry)
    {
        var address = Json.GetString(entry, "range") ?? Json.GetString(entry, "formulaRange");
        if (string.IsNullOrWhiteSpace(address)) return;
        if (entry["formulas"] is JsonArray formulas)
            RestoreRangeFormulas(sheet, address, formulas);
        else if (entry["values"] is JsonArray values)
            RestoreRangeValues(sheet, address, values);
        if (entry["numberFormats"] is JsonArray formats)
            RestoreNumberFormats(sheet, address, formats);
        ClearNotesInRange(sheet, address);
        ClearHyperlinksInRange(sheet, address);
        RestoreNotesList(sheet, Json.GetArr(entry, "notes"));
        RestoreHyperlinkList(sheet, Json.GetArr(entry, "hyperlinks"));
    }

    private static JsonNode CaptureRangeValues(object sheet, string address)
    {
        object? range = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            return RangeToJson(range, out _, formulas: false, maxCells: MaxDataSnapshotCells);
        }
        finally { RotHelper.ReleaseComReference(range); }
    }

    private static void RestoreRangeValues(object sheet, string address, JsonArray values)
    {
        object? range = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            var rows = values.Count;
            var cols = values[0] is JsonArray first ? first.Count : 0;
            var grid = new object[rows, cols];
            for (var r = 0; r < rows; r++)
            {
                if (values[r] is not JsonArray row) continue;
                for (var c = 0; c < cols && c < row.Count; c++)
                    grid[r, c] = NodeToComValue(row[c]) ?? "";
            }
            ((dynamic)range).Value2 = grid;
        }
        finally { RotHelper.ReleaseComReference(range); }
    }

    private static void RestoreNumberFormats(object sheet, string address, JsonArray formats)
    {
        if (formats.Count == 0)
            throw new InvalidOperationException(
                $"{address}: numberFormats snapshot is empty; mixed NumberFormat must be captured per-cell");
        object? range = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            var rows = formats.Count;
            var cols = formats[0] is JsonArray first ? first.Count : 0;
            if (rows == 1 && cols == 1 && formats[0] is JsonArray single)
            {
                ((dynamic)range).NumberFormat = single[0]?.ToString() ?? "General";
                return;
            }
            var grid = new object[rows, cols];
            for (var r = 0; r < rows; r++)
            {
                if (formats[r] is not JsonArray row) continue;
                for (var c = 0; c < cols && c < row.Count; c++)
                    grid[r, c] = row[c]?.ToString() ?? "General";
            }
            ((dynamic)range).NumberFormat = grid;
        }
        finally { RotHelper.ReleaseComReference(range); }
    }

    private static void RestoreNotesList(object sheet, JsonArray? notes)
    {
        if (notes is null) return;
        foreach (var node in notes.OfType<JsonObject>())
        {
            var address = Json.GetString(node, "range");
            if (string.IsNullOrWhiteSpace(address)) continue;
            object? range = null;
            object? comment = null;
            try
            {
                range = (object)((dynamic)sheet).Range(address);
                ((dynamic)range).ClearComments();
                comment = (object)((dynamic)range).AddComment(Json.GetString(node, "text") ?? "");
                if (node.ContainsKey("visible"))
                    ((dynamic)comment).Visible = Json.GetBool(node, "visible");
            }
            catch { /* skip one note */ }
            finally
            {
                RotHelper.ReleaseComReference(comment);
                RotHelper.ReleaseComReference(range);
            }
        }
    }

    private static void RestoreHyperlinkList(object sheet, JsonArray? links)
    {
        if (links is null) return;
        object? sheetLinks = null;
        try
        {
            sheetLinks = (object)((dynamic)sheet).Hyperlinks;
            foreach (var node in links.OfType<JsonObject>())
            {
                var address = Json.GetString(node, "range");
                if (string.IsNullOrWhiteSpace(address)) continue;
                object? range = null;
                object? existing = null;
                try
                {
                    range = (object)((dynamic)sheet).Range(address);
                    existing = (object)((dynamic)range).Hyperlinks;
                    try { ((dynamic)existing).Delete(); } catch { /* none */ }
                    ((dynamic)sheetLinks).Add(
                        range,
                        Json.GetString(node, "address") ?? "",
                        Json.GetString(node, "subAddress") ?? Type.Missing,
                        Json.GetString(node, "screenTip") ?? Type.Missing,
                        Json.GetString(node, "textToDisplay") ?? Type.Missing);
                }
                catch { /* skip one link */ }
                finally
                {
                    RotHelper.ReleaseComReference(existing);
                    RotHelper.ReleaseComReference(range);
                }
            }
        }
        finally { RotHelper.ReleaseComReference(sheetLinks); }
    }

    private static object RecreateDefinedName(object workbook, string name, string scope, string? sheetName,
        string refersTo, string? comment)
    {
        object? created;
        if (scope == "sheet")
        {
            object? sheet = null;
            object? names = null;
            try
            {
                sheet = GetExplicitTargetSheetReference(workbook, SheetOp(sheetName ?? ""));
                names = (object)((dynamic)sheet).Names;
                created = (object)((dynamic)names).Add(name, refersTo);
            }
            finally
            {
                RotHelper.ReleaseComReference(names);
                RotHelper.ReleaseComReference(sheet);
            }
        }
        else
        {
            object? names = null;
            try
            {
                names = (object)((dynamic)workbook).Names;
                created = (object)((dynamic)names).Add(name, refersTo);
            }
            finally { RotHelper.ReleaseComReference(names); }
        }
        if (!string.IsNullOrEmpty(comment))
            ((dynamic)created).Comment = comment;
        return created;
    }

    private static void ApplyValidationState(object validation, JsonObject state)
    {
        var type = Json.GetInt(state, "validationType") ?? Json.GetInt(state, "type");
        if (type is null) return;
        ((dynamic)validation).Add(
            type.Value,
            Json.GetInt(state, "alertStyle") ?? ExcelDataObjectCatalog.XlValidAlertStop,
            Json.GetInt(state, "operator") ?? ExcelDataObjectCatalog.XlBetween,
            Json.GetString(state, "formula1") ?? Type.Missing,
            Json.GetString(state, "formula2") ?? Type.Missing);
        if (state.ContainsKey("ignoreBlank"))
            ((dynamic)validation).IgnoreBlank = Json.GetBool(state, "ignoreBlank");
        if (state.ContainsKey("inCellDropdown"))
            ((dynamic)validation).InCellDropdown = Json.GetBool(state, "inCellDropdown");
        if (state.ContainsKey("showInput"))
            ((dynamic)validation).ShowInput = Json.GetBool(state, "showInput");
        if (state.ContainsKey("showError"))
            ((dynamic)validation).ShowError = Json.GetBool(state, "showError");
        if (state.ContainsKey("inputTitle")) ((dynamic)validation).InputTitle = Json.GetString(state, "inputTitle") ?? "";
        if (state.ContainsKey("inputMessage")) ((dynamic)validation).InputMessage = Json.GetString(state, "inputMessage") ?? "";
        if (state.ContainsKey("errorTitle")) ((dynamic)validation).ErrorTitle = Json.GetString(state, "errorTitle") ?? "";
        if (state.ContainsKey("errorMessage")) ((dynamic)validation).ErrorMessage = Json.GetString(state, "errorMessage") ?? "";
    }

    private static void RestoreChartSeries(object workbook, object chart, JsonObject state)
    {
        var source = Json.GetString(state, "sourceRange") ?? Json.GetString(state, "source");
        var defaultSheet = TryChartSheetName(chart);
        if (!string.IsNullOrWhiteSpace(source) &&
            ExcelDataRangeBinding.TrySourceSheet(defaultSheet, source, out var sheetName, out var address, out _))
        {
            object? sheet = null;
            object? range = null;
            try
            {
                sheet = (object)GetSheet(workbook, sheetName);
                range = (object)((dynamic)sheet).Range(address);
                ((dynamic)chart).SetSourceData(range);
            }
            catch { /* series values/categories below remain the restore path */ }
            finally
            {
                RotHelper.ReleaseComReference(range);
                RotHelper.ReleaseComReference(sheet);
            }
        }

        ApplyAdvertisedChartSeries(workbook, chart, Json.GetArr(state, "series"), defaultSheet ?? "");
        ExcelChartAxesApply.Apply(chart, Json.GetObj(state, "axes"));

        var series = Json.GetArr(state, "series");
        if (series is null || series.Count == 0)
        {
            var formula = Json.GetString(state, "series1Formula");
            if (string.IsNullOrWhiteSpace(formula)) return;
            series = new JsonArray { new JsonObject { ["index"] = 1, ["formula"] = formula } };
        }

        object? collection = null;
        try
        {
            collection = (object)((dynamic)chart).SeriesCollection();
            foreach (var node in series.OfType<JsonObject>())
            {
                var index = Json.GetInt(node, "index") ?? 1;
                var formula = Json.GetString(node, "formula");
                if (string.IsNullOrWhiteSpace(formula)) continue;
                if (node.ContainsKey("values") || node.ContainsKey("categories") || node.ContainsKey("name"))
                    continue;
                object? item = null;
                try
                {
                    item = (object)((dynamic)collection).Item(index);
                    ((dynamic)item).Formula = formula;
                }
                catch { /* series index may not exist after type change */ }
                finally { RotHelper.ReleaseComReference(item); }
            }
        }
        finally { RotHelper.ReleaseComReference(collection); }
    }

    private static string? TryChartSheetName(object chart)
    {
        object? chartObject = null;
        object? sheet = null;
        try
        {
            chartObject = (object)((dynamic)chart).Parent;
            sheet = (object)((dynamic)chartObject).Parent;
            return ReadComString(sheet, "Name");
        }
        catch
        {
            return null;
        }
        finally
        {
            RotHelper.ReleaseComReference(sheet);
            RotHelper.ReleaseComReference(chartObject);
        }
    }

    private static void RestoreCellContent(object range, JsonObject? state)
    {
        if (state is null) return;
        var formula = Json.GetString(state, "cellFormula");
        if (!string.IsNullOrWhiteSpace(formula))
            ((dynamic)range).Formula = formula;
        else if (state.ContainsKey("cellValue"))
            ((dynamic)range).Value2 = NodeToComValue(state["cellValue"]) ?? "";

        var fontState = Json.GetObj(state, "font");
        if (fontState is null) return;
        object? font = null;
        try
        {
            font = (object)((dynamic)range).Font;
            if (fontState.ContainsKey("bold"))
                ((dynamic)font).Bold = Json.GetBool(fontState, "bold");
            if (ExcelDataOperationsContract.TryGetFiniteNumber(fontState["color"], out var color))
                ((dynamic)font).Color = color;
        }
        finally { RotHelper.ReleaseComReference(font); }
    }

    private static IEnumerable<object> EnumerateDrawingShapes(object sheet)
    {
        object? shapes = null;
        try
        {
            shapes = (object)((dynamic)sheet).Shapes;
            var count = Convert.ToInt32(((dynamic)shapes).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                object? shape = null;
                try
                {
                    shape = (object)((dynamic)shapes).Item(index);
                    var type = TryComInt(shape, "Type");
                    if (type == ExcelDataObjectCatalog.MsoPicture) continue;
                    yield return shape;
                    shape = null;
                }
                finally { RotHelper.ReleaseComReference(shape); }
            }
        }
        finally { RotHelper.ReleaseComReference(shapes); }
    }

    private static IEnumerable<object> EnumeratePivotTables(object sheet)
    {
        object? pivots = null;
        try
        {
            pivots = (object)((dynamic)sheet).PivotTables();
            var count = Convert.ToInt32(((dynamic)pivots).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
                yield return (object)((dynamic)pivots).Item(index);
        }
        finally { RotHelper.ReleaseComReference(pivots); }
    }

    private static JsonObject ReadShapeState(object shape, string sheetName)
    {
        var type = TryComInt(shape, "Type");
        var text = TryShapeText(shape);
        var isConnector = TryComInt(shape, "Connector") == ExcelDataObjectCatalog.MsoTrue;
        var state = new JsonObject
        {
            ["type"] = isConnector ? "connector" : type == ExcelDataObjectCatalog.MsoTextBox ? "textbox" : "shape",
            ["sheet"] = sheetName,
            ["name"] = ReadComString(shape, "Name"),
            ["shapeType"] = Js(type),
            ["text"] = text,
            ["left"] = Js(TryComDouble(shape, "Left")),
            ["top"] = Js(TryComDouble(shape, "Top")),
            ["width"] = Js(TryComDouble(shape, "Width")),
            ["height"] = Js(TryComDouble(shape, "Height")),
        };
        ReadShapeRotation(shape, state);
        ReadShapeLine(shape, state);
        ReadShapeFont(shape, state);
        if (isConnector) ReadConnectorState(shape, state);
        return state;
    }

    private static void ReadShapeRotation(object shape, JsonObject state)
        => ReadShapeNumericProperty(state, "rotation", () => ((dynamic)shape).Rotation);

    private static void ReadShapeNumericProperty(JsonObject state, string field, Func<object?> read)
    {
        try
        {
            if (ExcelShapeFormatContract.TryCreateNumericReadbackValue(field, read(), out var value))
                state[field] = value;
            else
                state[field + "Unreadable"] = true;
        }
        catch { state[field + "Unreadable"] = true; }
    }

    private static void ReadShapeLine(object shape, JsonObject state)
    {
        object? line = null;
        object? foreColor = null;
        try
        {
            line = (object)((dynamic)shape).Line;
            try
            {
                object? rawVisible = ((dynamic)line).Visible;
                if (ExcelShapeFormatContract.TryNormalizeLineVisible(rawVisible, out bool visible))
                    state["lineVisible"] = visible;
                else
                    state["lineVisibleUnreadable"] = true;
            }
            catch { state["lineVisibleUnreadable"] = true; }
            ReadShapeNumericProperty(state, "lineWeight", () => ((dynamic)line).Weight);
            try
            {
                foreColor = (object)((dynamic)line).ForeColor;
                ReadShapeNumericProperty(state, "lineColor", () => ((dynamic)foreColor).RGB);
            }
            catch { state["lineColorUnreadable"] = true; }
        }
        catch { state["lineUnreadable"] = true; }
        finally
        {
            RotHelper.ReleaseComReference(foreColor);
            RotHelper.ReleaseComReference(line);
        }
    }

    private static void ReadShapeFont(object shape, JsonObject state)
    {
        object? frame = null;
        object? characters = null;
        object? font = null;
        try
        {
            frame = (object)((dynamic)shape).TextFrame;
            characters = (object)((dynamic)frame).Characters();
            font = (object)((dynamic)characters).Font;
            var values = new JsonObject();
            var unreadable = new JsonArray();
            TryReadShapeFontValue(values, unreadable, "name", () => ((dynamic)font).Name);
            TryReadShapeFontValue(values, unreadable, "size", () => ((dynamic)font).Size);
            TryReadShapeFontValue(values, unreadable, "bold", () => ((dynamic)font).Bold);
            TryReadShapeFontValue(values, unreadable, "italic", () => ((dynamic)font).Italic);
            TryReadShapeFontValue(values, unreadable, "color", () => ((dynamic)font).Color);
            if (values.Count > 0) state["font"] = values;
            if (unreadable.Count > 0) state["fontUnreadableFields"] = unreadable;
        }
        catch { state["fontUnreadable"] = true; }
        finally
        {
            RotHelper.ReleaseComReference(font);
            RotHelper.ReleaseComReference(characters);
            RotHelper.ReleaseComReference(frame);
        }
    }

    private static void TryReadShapeFontValue(JsonObject values, JsonArray unreadable, string name, Func<object?> read)
    {
        try
        {
            if (ExcelShapeFormatContract.TryCreateFontReadbackValue(name, read(), out var value))
                values[name] = value;
            else
                unreadable.Add(name);
        }
        catch { unreadable.Add(name); }
    }

    private static string? TryShapeText(object shape)
    {
        object? frame = null;
        object? characters = null;
        try
        {
            frame = (object)((dynamic)shape).TextFrame;
            characters = (object)((dynamic)frame).Characters();
            return TryComString(characters, "Text");
        }
        catch { return null; }
        finally
        {
            RotHelper.ReleaseComReference(characters);
            RotHelper.ReleaseComReference(frame);
        }
    }

    private static JsonObject ReadPivotState(object pivot, string sheetName)
    {
        object? cache = null;
        object? tableRange = null;
        try
        {
            try { cache = (object)((dynamic)pivot).PivotCache; } catch { cache = null; }
            try { tableRange = (object)((dynamic)pivot).TableRange2; } catch { tableRange = null; }
            return new JsonObject
            {
                ["type"] = "pivot",
                ["sheet"] = sheetName,
                ["name"] = ReadComString(pivot, "Name"),
                ["ownerWorkbook"] = ReadOwnerWorkbookName(pivot),
                ["source"] = cache is null ? null : TryComString(cache, "SourceData"),
                ["range"] = TryAddress(tableRange),
                ["rowFields"] = ReadPivotFieldNames(pivot, "RowFields"),
                ["columnFields"] = ReadPivotFieldNames(pivot, "ColumnFields"),
                ["pageFields"] = ReadPivotFieldNames(pivot, "PageFields"),
                ["dataFields"] = ReadPivotDataFields(pivot),
                ["aggregates"] = ReadPivotAggregates(pivot),
            };
        }
        finally
        {
            RotHelper.ReleaseComReference(tableRange);
            RotHelper.ReleaseComReference(cache);
        }
    }

    private static string? ReadOwnerWorkbookName(object excelObject)
    {
        object? current = excelObject;
        var acquired = new List<object>();
        try
        {
            for (var hop = 0; hop < 4; hop++)
            {
                var full = TryComString(current, "FullName");
                if (!string.IsNullOrWhiteSpace(full))
                    return PreferOwnerWorkbookName(TryComString(current, "Name"), full);
                object? next;
                try { next = (object)((dynamic)current).Parent; }
                catch { return null; }
                if (next is null) return null;
                acquired.Add(next);
                current = next;
            }
            return null;
        }
        finally
        {
            for (var i = acquired.Count - 1; i >= 0; i--)
                RotHelper.ReleaseComReference(acquired[i]);
        }
    }

    private static string? PreferOwnerWorkbookName(string? name, string? fullName)
    {
        if (!string.IsNullOrWhiteSpace(name)) return name.Trim();
        return string.IsNullOrWhiteSpace(fullName) ? null : fullName.Trim();
    }

    private static JsonArray ReadPivotFieldNames(object pivot, string collection)
    {
        var names = new JsonArray();
        object? fields = null;
        try
        {
            fields = collection switch
            {
                "RowFields" => (object)((dynamic)pivot).RowFields,
                "ColumnFields" => (object)((dynamic)pivot).ColumnFields,
                "PageFields" => (object)((dynamic)pivot).PageFields,
                _ => null,
            };
            if (fields is null) return names;
            var count = Convert.ToInt32(((dynamic)fields).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count && index <= 64; index++)
            {
                object? field = null;
                try
                {
                    field = (object)((dynamic)fields).Item(index);
                    names.Add(ReadComString(field, "Name"));
                }
                finally { RotHelper.ReleaseComReference(field); }
            }
        }
        catch { /* collection empty */ }
        finally { RotHelper.ReleaseComReference(fields); }
        return names;
    }

    private static JsonArray ReadPivotDataFields(object pivot)
    {
        var items = new JsonArray();
        object? fields = null;
        try
        {
            fields = (object)((dynamic)pivot).DataFields;
            var count = Convert.ToInt32(((dynamic)fields).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count && index <= 64; index++)
            {
                object? field = null;
                try
                {
                    field = (object)((dynamic)fields).Item(index);
                    items.Add(new JsonObject
                    {
                        ["name"] = ReadComString(field, "Name"),
                        ["caption"] = ReadComString(field, "Caption") ?? ReadComString(field, "Name"),
                        ["sourceName"] = TryComString(field, "SourceName"),
                        ["numberFormat"] = TryComString(field, "NumberFormat"),
                        ["function"] = ExcelDataObjectCatalog.TokenFor(
                            ExcelDataObjectCatalog.PivotFunctions,
                            TryComInt(field, "Function") ?? ExcelDataObjectCatalog.XlSum,
                            "sum"),
                    });
                }
                finally { RotHelper.ReleaseComReference(field); }
            }
        }
        catch { /* no data fields */ }
        finally { RotHelper.ReleaseComReference(fields); }
        return items;
    }

    private static JsonArray ReadPivotAggregates(object pivot)
    {
        var items = new JsonArray();
        object? fields = null;
        try
        {
            fields = (object)((dynamic)pivot).DataFields;
            var count = Convert.ToInt32(((dynamic)fields).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count && index <= 64; index++)
            {
                object? field = null;
                object? dataRange = null;
                try
                {
                    field = (object)((dynamic)fields).Item(index);
                    var name = ReadComString(field, "Name");
                    var function = ExcelDataObjectCatalog.TokenFor(
                        ExcelDataObjectCatalog.PivotFunctions,
                        TryComInt(field, "Function") ?? ExcelDataObjectCatalog.XlSum,
                        "sum");
                    var numbers = new JsonArray();
                    try
                    {
                        dataRange = (object)((dynamic)field).DataRange;
                        var grid = RangeToJson(dataRange, out _, formulas: false, maxCells: 64);
                        CollectFiniteNumbers(grid, numbers, 64);
                    }
                    catch { /* empty DataRange */ }
                    var hasAggregate = numbers.Count > 0;
                    var entry = new JsonObject
                    {
                        ["name"] = name,
                        ["field"] = name,
                        ["function"] = function,
                        ["values"] = numbers,
                        ["hasAggregate"] = hasAggregate,
                    };
                    if (hasAggregate &&
                        ExcelDataOperationsContract.TryGetFiniteNumber(numbers[0], out var first))
                        entry["value"] = first;
                    items.Add(entry);
                }
                finally
                {
                    RotHelper.ReleaseComReference(dataRange);
                    RotHelper.ReleaseComReference(field);
                }
            }
        }
        catch { /* no data fields */ }
        finally { RotHelper.ReleaseComReference(fields); }
        return items;
    }

    private static void CollectFiniteNumbers(JsonNode? node, JsonArray numbers, int max)
    {
        if (node is null || numbers.Count >= max) return;
        if (ExcelDataOperationsContract.TryGetFiniteNumber(node, out var number))
        {
            numbers.Add(number);
            return;
        }
        if (node is JsonArray array)
        {
            foreach (var child in array)
                CollectFiniteNumbers(child, numbers, max);
        }
    }

    private static bool ReadBoundedCellObjects(object sheet, string sheetName, string scope, bool includeAll,
        string? rangeFilter, Action<JsonObject> add)
    {
        var address = rangeFilter;
        if (string.IsNullOrWhiteSpace(address))
            address = TryUsedRangeAddress(sheet);
        if (string.IsNullOrWhiteSpace(address) ||
            !ExcelDataOperationsContract.TryParseA1(address, out var box))
            return false;

        var truncated = box.CellCount > MaxValidationSnapshotCells;
        if (truncated && ExcelDataOperationsContract.TryParseA1(address, out var limited))
        {
            var rows = Math.Max(1, MaxValidationSnapshotCells / Math.Max(1, limited.Columns));
            address = $"{ColName(limited.Column)}{limited.Row}:{ColName(limited.EndColumn)}{limited.Row + rows - 1}";
        }

        if (includeAll || scope == "validations")
            add(ReadValidationState(sheet, address));
        if (includeAll || scope == "conditionalFormats")
            add(ReadConditionalFormats(sheet, address));
        if (includeAll || scope == "notes")
        {
            foreach (var note in CaptureNotesInRange(sheet, address).OfType<JsonObject>())
            {
                var item = ExcelJsonOwnership.DetachObject(note);
                item["type"] = "note";
                item["sheet"] = sheetName;
                item["present"] = true;
                add(item);
            }
        }
        if (includeAll || scope == "hyperlinks")
        {
            object? range = null;
            try
            {
                range = (object)((dynamic)sheet).Range(address);
                foreach (var link in CaptureHyperlinksInRange(range, address).OfType<JsonObject>())
                {
                    var item = ExcelJsonOwnership.DetachObject(link);
                    item["type"] = "hyperlink";
                    item["sheet"] = sheetName;
                    item["present"] = true;
                    add(item);
                }
            }
            finally { RotHelper.ReleaseComReference(range); }
        }
        return truncated;
    }

    private static string? TryUsedRangeAddress(object sheet)
    {
        object? used = null;
        try
        {
            used = (object)((dynamic)sheet).UsedRange;
            return TryAddress(used);
        }
        catch { return null; }
        finally { RotHelper.ReleaseComReference(used); }
    }

    private static bool RangeIntersects(string? actual, string? filter)
    {
        if (string.IsNullOrWhiteSpace(filter) || string.IsNullOrWhiteSpace(actual)) return true;
        if (!ExcelDataOperationsContract.TryParseA1(actual, out var a) ||
            !ExcelDataOperationsContract.TryParseA1(filter, out var b))
            return true;
        return a.Row <= b.EndRow && b.Row <= a.EndRow && a.Column <= b.EndColumn && b.Column <= a.EndColumn;
    }

    private static object? FindPivotTable(object sheet, string name)
    {
        foreach (var pivot in EnumeratePivotTables(sheet))
        {
            if (string.Equals(ReadComString(pivot, "Name"), name, StringComparison.OrdinalIgnoreCase))
                return pivot;
            RotHelper.ReleaseComReference(pivot);
        }
        return null;
    }

    private static void ApplyPivotLayout(object pivot, JsonObject op) =>
        ApplyPivotLayout(pivot, op, replaceOmitted: true);

    private static void ApplyPivotLayout(object pivot, JsonObject op, bool replaceOmitted) =>
        ExcelPivotApply.ApplyLayout(pivot, op, replaceOmitted);

    private static void ApplyPivotCacheSource(object workbook, object pivot, string cacheSource)
    {
        object? cache = null;
        object? caches = null;
        object? created = null;
        try
        {
            cache = (object)((dynamic)pivot).PivotCache;
            var users = CountPivotsUsingCache(workbook, cache);
            if (ExcelDataRangeBinding.MustCreatePrivateCache(users))
            {
                caches = (object)((dynamic)workbook).PivotCaches();
                created = (object)((dynamic)caches).Create(ExcelDataObjectCatalog.XlDatabase, cacheSource);
                ((dynamic)pivot).ChangePivotCache(created);
            }
            else
            {
                try
                {
                    ((dynamic)cache).SourceData = cacheSource;
                }
                catch
                {
                    caches = (object)((dynamic)workbook).PivotCaches();
                    created = (object)((dynamic)caches).Create(ExcelDataObjectCatalog.XlDatabase, cacheSource);
                    ((dynamic)pivot).ChangePivotCache(created);
                }
            }
            ((dynamic)pivot).RefreshTable();
        }
        finally
        {
            RotHelper.ReleaseComReference(created);
            RotHelper.ReleaseComReference(caches);
            RotHelper.ReleaseComReference(cache);
        }
    }

    private static int CountPivotsUsingCache(object workbook, object cache)
    {
        var users = 0;
        object? worksheets = null;
        try
        {
            worksheets = (object)((dynamic)workbook).Worksheets;
            var count = Convert.ToInt32(((dynamic)worksheets).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                object? sheet = null;
                try
                {
                    sheet = (object)((dynamic)worksheets).Item(index);
                    object? pivots = null;
                    try
                    {
                        pivots = (object)((dynamic)sheet).PivotTables();
                        var pivotCount = Convert.ToInt32(((dynamic)pivots).Count, CultureInfo.InvariantCulture);
                        for (var pivotIndex = 1; pivotIndex <= pivotCount; pivotIndex++)
                        {
                            object? other = null;
                            object? otherCache = null;
                            try
                            {
                                other = (object)((dynamic)pivots).Item(pivotIndex);
                                otherCache = (object)((dynamic)other).PivotCache;
                                if (SamePivotCache(cache, otherCache))
                                    users++;
                            }
                            finally
                            {
                                RotHelper.ReleaseComReference(otherCache);
                                RotHelper.ReleaseComReference(other);
                            }
                        }
                    }
                    catch { /* sheet has no pivots */ }
                    finally { RotHelper.ReleaseComReference(pivots); }
                }
                finally { RotHelper.ReleaseComReference(sheet); }
            }
        }
        finally { RotHelper.ReleaseComReference(worksheets); }
        return users;
    }

    private static bool SamePivotCache(object left, object? right)
    {
        if (right is null) return false;
        if (ReferenceEquals(left, right)) return true;
        try
        {
            var a = Convert.ToInt32(((dynamic)left).Index, CultureInfo.InvariantCulture);
            var b = Convert.ToInt32(((dynamic)right).Index, CultureInfo.InvariantCulture);
            return a == b && a > 0;
        }
        catch
        {
            return false;
        }
    }

    private static void HideAllPivotFields(object pivot)
    {
        object? fields = null;
        try
        {
            fields = (object)((dynamic)pivot).PivotFields();
            var count = Convert.ToInt32(((dynamic)fields).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                object? field = null;
                try
                {
                    field = (object)((dynamic)fields).Item(index);
                    var name = ReadComString(field, "Name") ?? "";
                    if (name.Contains("값", StringComparison.Ordinal) ||
                        name.Contains("Values", StringComparison.OrdinalIgnoreCase) ||
                        name.Contains("Data", StringComparison.OrdinalIgnoreCase))
                        continue;
                    ((dynamic)field).Orientation = ExcelDataObjectCatalog.XlHidden;
                }
                catch { /* calculated / data field */ }
                finally { RotHelper.ReleaseComReference(field); }
            }
        }
        finally { RotHelper.ReleaseComReference(fields); }
    }

    private static void SetPivotOrientation(object pivot, JsonArray? names, int orientation)
    {
        if (names is null) return;
        foreach (var node in names)
        {
            var fieldName = node?.ToString();
            if (string.IsNullOrWhiteSpace(fieldName)) continue;
            object? field = null;
            try
            {
                field = (object)((dynamic)pivot).PivotFields(fieldName);
                ((dynamic)field).Orientation = orientation;
            }
            finally { RotHelper.ReleaseComReference(field); }
        }
    }

    private static void SetShapeText(object shape, string? text)
    {
        if (text is null) return;
        object? frame = null;
        object? characters = null;
        try
        {
            frame = (object)((dynamic)shape).TextFrame;
            characters = (object)((dynamic)frame).Characters();
            ((dynamic)characters).Text = text;
        }
        finally
        {
            RotHelper.ReleaseComReference(characters);
            RotHelper.ReleaseComReference(frame);
        }
    }

    private static void ApplyShapeFill(object shape, JsonObject op)
    {
        if (!op.ContainsKey("fillColor") || !ExcelStyleContract.TryParseColor(op["fillColor"], out var color))
            return;
        object? fill = null;
        try
        {
            fill = (object)((dynamic)shape).Fill;
            ((dynamic)fill).ForeColor.RGB = color;
        }
        catch { /* some shapes reject fill */ }
        finally { RotHelper.ReleaseComReference(fill); }
    }

    private static void ApplyShapeFormatting(object shape, JsonObject op)
    {
        object? line = null;
        object? foreColor = null;
        object? textFrame = null;
        object? characters = null;
        object? font = null;
        try
        {
            if (op.ContainsKey("lineColor") || op.ContainsKey("lineWeight") || op.ContainsKey("lineVisible"))
            {
                line = (object)((dynamic)shape).Line;
                if (op.ContainsKey("lineColor"))
                {
                    foreColor = (object)((dynamic)line).ForeColor;
                    ((dynamic)foreColor).RGB = ExcelStyleContract.TryParseColor(op["lineColor"], out var color)
                        ? color
                        : throw new InvalidOperationException("[EXCEL_SHAPE_LINE_COLOR] lineColor was not validated");
                }
                if (op.ContainsKey("lineWeight"))
                    ((dynamic)line).Weight = (float)ExcelShapeFormatContract.ReadFiniteNumber(op["lineWeight"]!);
                if (op.ContainsKey("lineVisible"))
                    ((dynamic)line).Visible = Json.GetBool(op, "lineVisible") ? -1 : 0;
            }

            if (op.ContainsKey("rotation"))
                ((dynamic)shape).Rotation = (float)ExcelShapeFormatContract.NormalizeRequestedRotation(op["rotation"]!);

            if (Json.GetObj(op, "font") is JsonObject requestedFont)
            {
                textFrame = (object)((dynamic)shape).TextFrame;
                characters = (object)((dynamic)textFrame).Characters();
                font = (object)((dynamic)characters).Font;
                if (requestedFont.ContainsKey("name")) ((dynamic)font).Name = Json.GetString(requestedFont, "name");
                if (requestedFont.ContainsKey("size"))
                    ((dynamic)font).Size = (float)ExcelShapeFormatContract.ReadFiniteNumber(requestedFont["size"]!);
                if (requestedFont.ContainsKey("bold")) ((dynamic)font).Bold = Json.GetBool(requestedFont, "bold");
                if (requestedFont.ContainsKey("italic")) ((dynamic)font).Italic = Json.GetBool(requestedFont, "italic");
                if (requestedFont.ContainsKey("color"))
                {
                    ((dynamic)font).Color = ExcelStyleContract.TryParseColor(requestedFont["color"], out var color)
                        ? color
                        : throw new InvalidOperationException("[EXCEL_SHAPE_FONT_COLOR] font.color was not validated");
                }
            }
        }
        finally
        {
            RotHelper.ReleaseComReference(font);
            RotHelper.ReleaseComReference(characters);
            RotHelper.ReleaseComReference(textFrame);
            RotHelper.ReleaseComReference(foreColor);
            RotHelper.ReleaseComReference(line);
        }
    }

    private static string ComputeTextToColumnsCaptureAddress(DataRangeLease source, string destOrigin, JsonObject op)
    {
        if (!ExcelDataOperationsContract.TryParseA1(source.Address, out var sourceBox))
            return destOrigin;
        var origin = sourceBox;
        if (ExcelDataOperationsContract.TryParseA1(DataLocalAddress(destOrigin), out var destBox))
            origin = destBox;
        var width = EstimateTextToColumnsWidth(source.Sheet, source.Address, sourceBox, op);
        var height = sourceBox.Rows;
        if (ExcelDataOperationsContract.TryParseA1(DataLocalAddress(destOrigin), out destBox))
        {
            width = Math.Max(width, destBox.Columns);
            height = Math.Max(height, destBox.Rows);
        }
        return new ExcelA1Box(origin.Row, origin.Column, height, width).Address;
    }

    private static int EstimateTextToColumnsWidth(object sheet, string address, ExcelA1Box sourceBox, JsonObject op)
    {
        var delimiters = ExcelDataReadbackContract.TextToColumnsDelimiters(op);
        var width = Math.Max(1, sourceBox.Columns);
        if (delimiters.Length == 0) return width;
        if (CaptureRangeValues(sheet, address) is not JsonArray rows) return width;
        foreach (var rowNode in rows)
        {
            if (rowNode is not JsonArray row || row.Count == 0) continue;
            var text = row[0]?.ToString() ?? "";
            if (text.Length == 0) continue;
            var fields = ExcelTextToColumnsParse.ExpectedFields(text, op);
            width = Math.Max(width, fields.Length);
        }
        return Math.Clamp(width, 1, 256);
    }

    private static void CollectTextToColumnsOverwriteConflicts(object destSheet, string destSheetName,
        string sourceSheetName, string sourceAddress, string captureAddress, ApplyPreview preview)
    {
        if (!ExcelDataOperationsContract.TryParseA1(captureAddress, out var capture)) return;
        ExcelA1Box? sourceBox = ExcelDataOperationsContract.TryParseA1(sourceAddress, out var parsedSource) &&
                                string.Equals(destSheetName, sourceSheetName, StringComparison.OrdinalIgnoreCase)
            ? parsedSource
            : null;
        if (CaptureRangeValues(destSheet, captureAddress) is not JsonArray rows) return;
        for (var r = 0; r < rows.Count; r++)
        {
            if (rows[r] is not JsonArray row) continue;
            for (var c = 0; c < row.Count; c++)
            {
                if (IsEmptySortCell(row[c])) continue;
                var cell = new ExcelA1Box(capture.Row + r, capture.Column + c, 1, 1);
                if (sourceBox is ExcelA1Box source && source.Intersects(cell))
                    continue;
                preview.Errors.Add(
                    $"[EXCEL_T2C_OVERWRITE] {destSheetName}!{cell.Address} is non-empty and outside source {sourceSheetName}!{sourceAddress}");
            }
        }
    }

    private static void ClearNotesInRange(object sheet, string address)
    {
        object? range = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            ((dynamic)range).ClearComments();
        }
        catch { /* no comments in range */ }
        finally { RotHelper.ReleaseComReference(range); }
    }

    private static void ClearHyperlinksInRange(object sheet, string address)
    {
        object? range = null;
        object? links = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            links = (object)((dynamic)range).Hyperlinks;
            ((dynamic)links).Delete();
        }
        catch { /* no hyperlinks in range */ }
        finally
        {
            RotHelper.ReleaseComReference(links);
            RotHelper.ReleaseComReference(range);
        }
    }
}
