using System.Globalization;
using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class ExcelAdapter
{
    private sealed class DataSheetLease : IDisposable
    {
        public DataSheetLease(object sheet, string sheetName)
        {
            Sheet = sheet;
            SheetName = sheetName;
        }

        public object Sheet { get; }
        public string SheetName { get; }
        public void Dispose() => RotHelper.ReleaseComReference(Sheet);
    }

    private sealed class DataRangeLease : IDisposable
    {
        public DataRangeLease(object sheet, string sheetName, string address, object? listObject = null,
            string? cacheSource = null)
        {
            Sheet = sheet;
            SheetName = sheetName;
            Address = address;
            ListObject = listObject;
            CacheSource = cacheSource ?? address;
        }

        public object Sheet { get; }
        public string SheetName { get; }
        public string Address { get; }
        public object? ListObject { get; }
        public string CacheSource { get; }
        public void Dispose()
        {
            RotHelper.ReleaseComReference(ListObject);
            RotHelper.ReleaseComReference(Sheet);
        }
    }

    private sealed class DataTableLease : IDisposable
    {
        public DataTableLease(object sheet, string sheetName, object listObject, string name, JsonObject state)
        {
            Sheet = sheet;
            SheetName = sheetName;
            ListObject = listObject;
            Name = name;
            State = state;
        }

        public object Sheet { get; }
        public string SheetName { get; }
        public object ListObject { get; }
        public string Name { get; }
        public JsonObject State { get; }
        public void Dispose()
        {
            RotHelper.ReleaseComReference(ListObject);
            RotHelper.ReleaseComReference(Sheet);
        }
    }

    private sealed class DataChartLease : IDisposable
    {
        public DataChartLease(object sheet, string sheetName, object chartObject, string name, JsonObject state)
        {
            Sheet = sheet;
            SheetName = sheetName;
            ChartObject = chartObject;
            Name = name;
            State = state;
        }

        public object Sheet { get; }
        public string SheetName { get; }
        public object ChartObject { get; }
        public string Name { get; }
        public JsonObject State { get; }
        public void Dispose()
        {
            RotHelper.ReleaseComReference(ChartObject);
            RotHelper.ReleaseComReference(Sheet);
        }
    }

    private sealed class DataPictureLease : IDisposable
    {
        public DataPictureLease(object sheet, string sheetName, object shape, string name, JsonObject state)
        {
            Sheet = sheet;
            SheetName = sheetName;
            Shape = shape;
            Name = name;
            State = state;
        }

        public object Sheet { get; }
        public string SheetName { get; }
        public object Shape { get; }
        public string Name { get; }
        public JsonObject State { get; }
        public void Dispose()
        {
            RotHelper.ReleaseComReference(Shape);
            RotHelper.ReleaseComReference(Sheet);
        }
    }

    private readonly record struct DataPosition(double Left, double Top, double Width, double Height);
    private readonly record struct DataSortKey(string Address, int Order);

    private static DataSheetLease BindSheet(object workbook, JsonObject op)
    {
        var sheet = GetExplicitTargetSheetReference(workbook, op);
        return new DataSheetLease(sheet, ReadComString(sheet, "Name") ?? "");
    }

    private static DataRangeLease BindRange(object workbook, JsonObject op, string field,
        bool independentSource = false)
    {
        var text = Json.GetString(op, field);
        if (string.IsNullOrWhiteSpace(text))
            throw new InvalidOperationException($"Excel data op requires '{field}'");
        var targetSheet = Json.GetString(Json.GetObj(op, "target"), "sheet");
        string? bindSheet = targetSheet;
        if (independentSource)
        {
            if (!ExcelDataRangeBinding.TrySourceSheet(targetSheet, text, out var sourceSheet, out _, out var error))
                throw new InvalidOperationException(error ?? $"Excel data op '{field}' is not a usable source");
            bindSheet = sourceSheet;
        }
        var resolved = ResolveRangeTarget(
            workbook,
            bindSheet,
            text,
            requireExplicitSheet: true);
        object? range = null;
        try
        {
            range = (object)((dynamic)resolved.Sheet).Range(resolved.Address);
            var sheetName = ReadComString(resolved.Sheet, "Name") ?? "";
            var address = Convert.ToString(((dynamic)range).Address(false, false), CultureInfo.InvariantCulture)
                          ?? resolved.Address;
            return new DataRangeLease(
                resolved.Sheet,
                sheetName,
                address,
                cacheSource: ExcelDataRangeBinding.QualifiedA1(sheetName, address));
        }
        catch
        {
            RotHelper.ReleaseComReference(resolved.Sheet);
            throw;
        }
        finally { RotHelper.ReleaseComReference(range); }
    }

    private static DataRangeLease BindPivotSource(object workbook, JsonObject op)
    {
        var text = Json.GetString(op, "sourceRange");
        if (string.IsNullOrWhiteSpace(text))
            throw new InvalidOperationException("Excel data op requires 'sourceRange'");
        if (ExcelDataRangeBinding.LooksLikeTableSource(text))
        {
            var table = FindListObjectInWorkbook(workbook, text, out var sheet);
            if (table is null || sheet is null)
            {
                RotHelper.ReleaseComReference(table);
                RotHelper.ReleaseComReference(sheet);
                throw new InvalidOperationException(
                    $"[EXCEL_TABLE_NOT_FOUND] ListObject '{text}' was not found in the target workbook");
            }
            object? range = null;
            try
            {
                range = (object)((dynamic)table).Range;
                var sheetName = ReadComString(sheet, "Name") ?? "";
                var address = Convert.ToString(((dynamic)range).Address(false, false), CultureInfo.InvariantCulture)
                              ?? text;
                return new DataRangeLease(sheet, sheetName, address, table, cacheSource: text);
            }
            catch
            {
                RotHelper.ReleaseComReference(table);
                RotHelper.ReleaseComReference(sheet);
                throw;
            }
            finally { RotHelper.ReleaseComReference(range); }
        }

        return BindRange(workbook, op, "sourceRange", independentSource: true);
    }

    private static DataRangeLease BindCell(object workbook, JsonObject op)
    {
        var field = !string.IsNullOrWhiteSpace(Json.GetString(op, "range")) ? "range" : "cell";
        return BindRange(workbook, op, field);
    }

    private static DataTableLease BindTable(object workbook, JsonObject op)
    {
        var sheet = GetExplicitTargetSheetReference(workbook, op);
        var sheetName = ReadComString(sheet, "Name") ?? "";
        var name = Json.GetString(op, "name")!;
        var table = FindListObject(sheet, name);
        if (table is null)
        {
            RotHelper.ReleaseComReference(sheet);
            throw new InvalidOperationException(
                $"[EXCEL_TABLE_NOT_FOUND] ListObject '{name}' was not found on '{sheetName}'");
        }
        var state = ReadListObjectState(table);
        return new DataTableLease(sheet, sheetName, table, Json.GetString(state, "name") ?? name, state);
    }

    private static DataRangeLease BindFilterTarget(object workbook, JsonObject op)
    {
        if (!string.IsNullOrWhiteSpace(Json.GetString(op, "name")))
        {
            var table = BindTable(workbook, op);
            var address = Json.GetString(table.State, "range") ?? "";
            var lease = new DataRangeLease(table.Sheet, table.SheetName, address, table.ListObject);
            return lease;
        }
        return BindRange(workbook, op, "range");
    }

    private static DataChartLease BindChart(object workbook, JsonObject op)
    {
        var sheet = GetExplicitTargetSheetReference(workbook, op);
        var name = Json.GetString(op, "name")!;
        var chart = FindChartObject(sheet, name);
        if (chart is null)
        {
            RotHelper.ReleaseComReference(sheet);
            throw new InvalidOperationException($"[EXCEL_CHART_NOT_FOUND] ChartObject '{name}' was not found");
        }
        var state = ReadChartState(chart, ReadComString(sheet, "Name") ?? "");
        return new DataChartLease(sheet, ReadComString(sheet, "Name") ?? "", chart, Json.GetString(state, "name") ?? name, state);
    }

    private static DataPictureLease BindPicture(object workbook, JsonObject op)
    {
        var sheet = GetExplicitTargetSheetReference(workbook, op);
        var name = Json.GetString(op, "name")!;
        var shape = FindShape(sheet, name);
        if (shape is null)
        {
            RotHelper.ReleaseComReference(sheet);
            throw new InvalidOperationException($"[EXCEL_PICTURE_NOT_FOUND] Shape '{name}' was not found");
        }
        var state = ReadPictureState(shape, ReadComString(sheet, "Name") ?? "");
        return new DataPictureLease(sheet, ReadComString(sheet, "Name") ?? "", shape, Json.GetString(state, "name") ?? name, state);
    }

    private static void EnsureSheetWritable(object sheet)
    {
        if (Convert.ToBoolean(((dynamic)sheet).ProtectContents, CultureInfo.InvariantCulture))
            throw new InvalidOperationException(
                "[EXCEL_SHEET_PROTECTED] the target worksheet is protected; unprotect it before data/reporting writes");
    }

    private static object? FindListObject(object sheet, string name)
    {
        foreach (var table in EnumerateListObjects(sheet))
        {
            var actual = ReadComString(table, "Name");
            if (string.Equals(actual, name, StringComparison.OrdinalIgnoreCase))
                return table;
            RotHelper.ReleaseComReference(table);
        }
        return null;
    }

    private static object? FindListObjectInWorkbook(object workbook, string name, out object? owningSheet)
    {
        owningSheet = null;
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
                    var table = FindListObject(sheet, name);
                    if (table is null) continue;
                    owningSheet = sheet;
                    sheet = null;
                    return table;
                }
                finally { RotHelper.ReleaseComReference(sheet); }
            }
            return null;
        }
        finally { RotHelper.ReleaseComReference(worksheets); }
    }

    private static string? FindOverlappingTable(object sheet, string address)
    {
        object? range = null;
        object? list = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            list = (object?)((dynamic)range).ListObject;
            return list is null ? null : ReadComString(list, "Name");
        }
        catch { return null; }
        finally
        {
            RotHelper.ReleaseComReference(list);
            RotHelper.ReleaseComReference(range);
        }
    }

    private static object? FindChartObject(object sheet, string name)
    {
        foreach (var chart in EnumerateChartObjects(sheet))
        {
            if (string.Equals(ReadComString(chart, "Name"), name, StringComparison.OrdinalIgnoreCase))
                return chart;
            RotHelper.ReleaseComReference(chart);
        }
        return null;
    }

    private static object? FindShape(object sheet, string name)
    {
        object? shapes = null;
        try
        {
            shapes = (object)((dynamic)sheet).Shapes;
            try { return (object)((dynamic)shapes).Item(name); }
            catch { return null; }
        }
        finally { RotHelper.ReleaseComReference(shapes); }
    }

    private static object? FindDefinedName(object workbook, string name, string scope, string? sheetName)
    {
        var candidates = new List<string> { name };
        if (scope == "sheet" && !string.IsNullOrWhiteSpace(sheetName))
        {
            candidates.Insert(0, $"{QuoteSheetName(sheetName)}!{name}");
            var quoted = "'" + sheetName.Replace("'", "''", StringComparison.Ordinal) + "'!" + name;
            if (!candidates.Exists(item => string.Equals(item, quoted, StringComparison.OrdinalIgnoreCase)))
                candidates.Insert(1, quoted);
        }
        object? names = null;
        try
        {
            names = (object)((dynamic)workbook).Names;
            foreach (var candidate in candidates)
            {
                try { return (object)((dynamic)names).Item(candidate); }
                catch { /* try next */ }
            }
            var count = Convert.ToInt32(((dynamic)names).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                object? item = null;
                try
                {
                    item = (object)((dynamic)names).Item(index);
                    var actual = ReadComString(item, "Name") ?? "";
                    if (NameEqualsDefined(actual, name, scope, sheetName))
                    {
                        var keep = item;
                        item = null;
                        return keep;
                    }
                }
                finally { RotHelper.ReleaseComReference(item); }
            }
            return null;
        }
        finally { RotHelper.ReleaseComReference(names); }
    }

    private static bool NameEqualsDefined(string actual, string name, string scope, string? sheetName)
    {
        if (!ExcelFormulaReference.TryParseDefinedName(actual, out var actualScope, out var actualSheet, out var local))
            return false;
        if (!string.Equals(local, name, StringComparison.OrdinalIgnoreCase))
            return false;
        if (scope == "workbook")
            return string.Equals(actualScope, "workbook", StringComparison.OrdinalIgnoreCase);
        return string.Equals(actualScope, "sheet", StringComparison.OrdinalIgnoreCase) &&
               (string.IsNullOrWhiteSpace(sheetName) ||
                string.Equals(actualSheet, sheetName, StringComparison.OrdinalIgnoreCase));
    }

    private static IEnumerable<object> EnumerateWorksheets(object workbook, string? sheetName)
    {
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
                    var name = ReadComString(sheet, "Name") ?? "";
                    if (!string.IsNullOrWhiteSpace(sheetName) &&
                        !string.Equals(name, sheetName, StringComparison.OrdinalIgnoreCase))
                        continue;
                    yield return sheet;
                    sheet = null;
                }
                finally { RotHelper.ReleaseComReference(sheet); }
            }
        }
        finally { RotHelper.ReleaseComReference(worksheets); }
    }

    private static IEnumerable<object> EnumerateListObjects(object sheet)
    {
        object? listObjects = null;
        try
        {
            listObjects = (object)((dynamic)sheet).ListObjects;
            var count = Convert.ToInt32(((dynamic)listObjects).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
                yield return (object)((dynamic)listObjects).Item(index);
        }
        finally { RotHelper.ReleaseComReference(listObjects); }
    }

    private static IEnumerable<object> EnumerateChartObjects(object sheet)
    {
        object? charts = null;
        try
        {
            charts = (object)((dynamic)sheet).ChartObjects();
            var count = Convert.ToInt32(((dynamic)charts).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
                yield return (object)((dynamic)charts).Item(index);
        }
        finally { RotHelper.ReleaseComReference(charts); }
    }

    private static IEnumerable<object> EnumeratePictures(object sheet)
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
                    if (type == ExcelDataObjectCatalog.MsoPicture)
                    {
                        yield return shape;
                        shape = null;
                    }
                }
                finally { RotHelper.ReleaseComReference(shape); }
            }
        }
        finally { RotHelper.ReleaseComReference(shapes); }
    }

    private static IEnumerable<object> EnumerateNames(object workbook)
    {
        object? names = null;
        try
        {
            names = (object)((dynamic)workbook).Names;
            var count = Convert.ToInt32(((dynamic)names).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
                yield return (object)((dynamic)names).Item(index);
        }
        finally { RotHelper.ReleaseComReference(names); }
    }

    private static JsonObject ReadListObjectState(object table)
    {
        object? range = null;
        object? data = null;
        object? header = null;
        object? columns = null;
        try
        {
            range = (object)((dynamic)table).Range;
            try { data = (object?)((dynamic)table).DataBodyRange; } catch { data = null; }
            try { header = (object?)((dynamic)table).HeaderRowRange; } catch { header = null; }
            columns = (object)((dynamic)table).ListColumns;
            var columnNames = new JsonArray();
            var totals = new JsonArray();
            var count = Convert.ToInt32(((dynamic)columns).Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                object? column = null;
                try
                {
                    column = (object)((dynamic)columns).Item(index);
                    var columnName = ReadComString(column, "Name") ?? $"Column{index}";
                    columnNames.Add(columnName);
                    var total = new JsonObject
                    {
                        ["column"] = columnName,
                        ["function"] = ExcelDataObjectCatalog.TokenFor(
                            ExcelDataObjectCatalog.TotalsFunctions,
                            TryComInt(column, "TotalsCalculation") ?? ExcelDataObjectCatalog.XlTotalsCalculationNone,
                            "none"),
                    };
                    object? totalsCell = null;
                    try
                    {
                        totalsCell = (object?)((dynamic)column).TotalsRowRange;
                        if (totalsCell is not null && TryComDouble(totalsCell, "Value") is double cellValue &&
                            double.IsFinite(cellValue))
                            total["value"] = cellValue;
                    }
                    catch { /* totals row hidden or empty */ }
                    finally { RotHelper.ReleaseComReference(totalsCell); }
                    totals.Add(total);
                }
                finally { RotHelper.ReleaseComReference(column); }
            }
            return new JsonObject
            {
                ["type"] = "table",
                ["name"] = ReadComString(table, "Name"),
                ["sheet"] = TryParentSheetName(table),
                ["range"] = TryAddress(range),
                ["dataRange"] = TryAddress(data),
                ["headerRange"] = TryAddress(header),
                ["styleName"] = TryTableStyleName(table),
                ["showTotals"] = Js(TryComBool(table, "ShowTotals")),
                ["showHeaders"] = Js(TryComBool(table, "ShowHeaders")),
                ["showAutoFilter"] = Js(TryComBool(table, "ShowAutoFilter")),
                ["showRowStripes"] = Js(TryComBool(table, "ShowTableStyleRowStripes")),
                ["showColumnStripes"] = Js(TryComBool(table, "ShowTableStyleColumnStripes")),
                ["columns"] = columnNames,
                ["totals"] = totals,
            };
        }
        finally
        {
            RotHelper.ReleaseComReference(columns);
            RotHelper.ReleaseComReference(header);
            RotHelper.ReleaseComReference(data);
            RotHelper.ReleaseComReference(range);
        }
    }

    private static JsonObject ReadChartState(object chartObject, string sheetName)
    {
        object? chart = null;
        object? title = null;
        object? legend = null;
        object? series = null;
        try
        {
            chart = (object)((dynamic)chartObject).Chart;
            var chartType = TryComInt(chart, "ChartType");
            var hasTitle = TryComBool(chart, "HasTitle");
            var hasLegend = TryComBool(chart, "HasLegend");
            string? titleText = null;
            if (hasTitle == true)
            {
                title = (object)((dynamic)chart).ChartTitle;
                titleText = ReadComString(title, "Text");
            }
            string? legendPosition = null;
            if (hasLegend == true)
            {
                legend = (object)((dynamic)chart).Legend;
                legendPosition = ExcelDataObjectCatalog.TokenFor(
                    ExcelDataObjectCatalog.LegendPositions,
                    TryComInt(legend, "Position") ?? 0,
                    "right");
            }
            var seriesList = ReadChartSeries(chart, out var seriesCount, out var seriesFormula);
            return new JsonObject
            {
                ["type"] = "chart",
                ["sheet"] = sheetName,
                ["name"] = ReadComString(chartObject, "Name"),
                ["ownerWorkbook"] = ReadOwnerWorkbookName(chartObject),
                ["chartType"] = chartType is null
                    ? null
                    : ExcelDataObjectCatalog.TokenFor(ExcelDataObjectCatalog.ChartTypes, chartType.Value, $"raw:{chartType}"),
                ["chartTypeValue"] = Js(chartType),
                ["combination"] = SeriesLooksLikeCombination(seriesList),
                ["title"] = titleText,
                ["hasTitle"] = Js(hasTitle),
                ["hasLegend"] = Js(hasLegend),
                ["legendPosition"] = legendPosition,
                ["left"] = Js(TryComDouble(chartObject, "Left")),
                ["top"] = Js(TryComDouble(chartObject, "Top")),
                ["width"] = Js(TryComDouble(chartObject, "Width")),
                ["height"] = Js(TryComDouble(chartObject, "Height")),
                ["seriesCount"] = seriesCount,
                ["series1Formula"] = seriesFormula,
                ["series"] = seriesList,
                ["chartFillVisible"] = Js(ReadFormatVisible(chart, "ChartArea", fill: true)),
                ["plotFillVisible"] = Js(ReadFormatVisible(chart, "PlotArea", fill: true)),
                ["chartLineVisible"] = Js(ReadFormatVisible(chart, "ChartArea", fill: false)),
                ["plotLineVisible"] = Js(ReadFormatVisible(chart, "PlotArea", fill: false)),
                ["axes"] = ReadChartAxes(chart),
                ["dataLabels"] = ReadChartDataLabels(chart),
            };
        }
        finally
        {
            RotHelper.ReleaseComReference(series);
            RotHelper.ReleaseComReference(legend);
            RotHelper.ReleaseComReference(title);
            RotHelper.ReleaseComReference(chart);
        }
    }

    private static bool SeriesLooksLikeCombination(JsonArray series)
    {
        var types = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var secondary = false;
        foreach (var node in series.OfType<JsonObject>())
        {
            var type = Json.GetString(node, "chartType");
            if (!string.IsNullOrWhiteSpace(type) &&
                !type.StartsWith("raw:", StringComparison.OrdinalIgnoreCase))
                types.Add(type);
            if (string.Equals(Json.GetString(node, "axisGroup"), "secondary", StringComparison.OrdinalIgnoreCase))
                secondary = true;
        }
        return secondary || types.Count > 1;
    }

    private static JsonObject ReadPictureState(object shape, string sheetName)
    {
        var state = new JsonObject
        {
            ["type"] = "picture", ["sheet"] = sheetName, ["name"] = ReadComString(shape, "Name"),
            ["shapeType"] = Js(TryComInt(shape, "Type")), ["left"] = Js(TryComDouble(shape, "Left")), ["top"] = Js(TryComDouble(shape, "Top")),
            ["width"] = Js(TryComDouble(shape, "Width")), ["height"] = Js(TryComDouble(shape, "Height")), ["rotation"] = Js(TryComDouble(shape, "Rotation")),
            ["lockAspectRatio"] = TryComInt(shape, "LockAspectRatio") switch { 0 => JsonValue.Create(false), -1 => JsonValue.Create(true), _ => null },
        };
        object? format = null;
        try { format = (object)((dynamic)shape).PictureFormat; state["crop"] = new JsonObject { ["left"] = Js(TryComDouble(format,"CropLeft")), ["top"] = Js(TryComDouble(format,"CropTop")), ["right"] = Js(TryComDouble(format,"CropRight")), ["bottom"] = Js(TryComDouble(format,"CropBottom")) }; }
        catch { state["cropUnreadable"] = true; }
        finally { RotHelper.ReleaseComReference(format); }
        return state;
    }

    private static JsonObject ReadDefinedNameState(object name)
    {
        var actualName = ReadComString(name, "Name") ?? "";
        ExcelFormulaReference.TryParseDefinedName(actualName, out var scope, out var sheet, out var local);
        return new JsonObject
        {
            ["type"] = "definedName",
            ["name"] = actualName,
            ["localName"] = local,
            ["scope"] = scope,
            ["sheet"] = sheet,
            ["refersTo"] = ReadComString(name, "RefersTo"),
            ["comment"] = TryComString(name, "Comment"),
            ["visible"] = Js(TryComBool(name, "Visible")),
        };
    }

    private static JsonObject ReadAutoFilterState(object sheet, string sheetName)
    {
        object? filter = null;
        object? range = null;
        try
        {
            var mode = TryComBool(sheet, "AutoFilterMode") == true;
            var filtered = TryComBool(sheet, "FilterMode") == true;
            string? address = null;
            var criteriaCount = 0;
            if (mode)
            {
                try
                {
                    filter = (object)((dynamic)sheet).AutoFilter;
                    range = (object)((dynamic)filter).Range;
                    address = TryAddress(range);
                    object? filters = null;
                    try
                    {
                        filters = (object)((dynamic)filter).Filters;
                        criteriaCount = Convert.ToInt32(((dynamic)filters).Count, CultureInfo.InvariantCulture);
                    }
                    finally { RotHelper.ReleaseComReference(filters); }
                }
                catch { /* AutoFilterMode can race with a closing sheet */ }
            }
            return new JsonObject
            {
                ["type"] = "autoFilter",
                ["sheet"] = sheetName,
                ["enabled"] = mode,
                ["filtered"] = filtered,
                ["range"] = address,
                ["filterCount"] = criteriaCount,
                ["filters"] = ReadAutoFilterCriteria(filter, criteriaCount),
            };
        }
        finally
        {
            RotHelper.ReleaseComReference(range);
            RotHelper.ReleaseComReference(filter);
        }
    }

    private static JsonObject ReadValidationState(object sheet, string address)
    {
        object? range = null;
        object? validation = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            validation = (object)((dynamic)range).Validation;
            var type = TryComInt(validation, "Type");
            if (type is null)
                return new JsonObject { ["type"] = "validation", ["present"] = false, ["range"] = address };
            return new JsonObject
            {
                ["type"] = "validation",
                ["present"] = true,
                ["range"] = address,
                ["validationType"] = Js(type),
                ["operator"] = Js(TryComInt(validation, "Operator")),
                ["alertStyle"] = Js(TryComInt(validation, "AlertStyle")),
                ["formula1"] = TryComString(validation, "Formula1"),
                ["formula2"] = TryComString(validation, "Formula2"),
                ["ignoreBlank"] = Js(TryComBool(validation, "IgnoreBlank")),
                ["inCellDropdown"] = Js(TryComBool(validation, "InCellDropdown")),
                ["showInput"] = Js(TryComBool(validation, "ShowInput")),
                ["showError"] = Js(TryComBool(validation, "ShowError")),
                ["inputTitle"] = TryComString(validation, "InputTitle"),
                ["inputMessage"] = TryComString(validation, "InputMessage"),
                ["errorTitle"] = TryComString(validation, "ErrorTitle"),
                ["errorMessage"] = TryComString(validation, "ErrorMessage"),
            };
        }
        catch
        {
            return new JsonObject { ["type"] = "validation", ["present"] = false, ["range"] = address };
        }
        finally
        {
            RotHelper.ReleaseComReference(validation);
            RotHelper.ReleaseComReference(range);
        }
    }

    private static JsonObject ReadConditionalFormats(object sheet, string address)
    {
        object? range = null;
        object? conditions = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            conditions = (object)((dynamic)range).FormatConditions;
            var count = Convert.ToInt32(((dynamic)conditions).Count, CultureInfo.InvariantCulture);
            var rules = new JsonArray();
            for (var index = 1; index <= count && index <= 64; index++)
            {
                object? item = null;
                try
                {
                    item = (object)((dynamic)conditions).Item(index);
                    var rule = ReadOneConditionalRule(item, index);
                    rule["fingerprint"] = ConditionalFingerprint(rule);
                    rules.Add(rule);
                }
                catch
                {
                    rules.Add(new JsonObject { ["index"] = index, ["unreadable"] = true });
                }
                finally { RotHelper.ReleaseComReference(item); }
            }
            return new JsonObject
            {
                ["type"] = "conditionalFormat",
                ["range"] = address,
                ["count"] = count,
                ["rules"] = rules,
            };
        }
        catch
        {
            return new JsonObject { ["type"] = "conditionalFormat", ["range"] = address, ["count"] = 0, ["rules"] = new JsonArray() };
        }
        finally
        {
            RotHelper.ReleaseComReference(conditions);
            RotHelper.ReleaseComReference(range);
        }
    }

    private static JsonObject ReadNoteState(object sheet, string address)
    {
        object? range = null;
        object? comment = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            comment = (object?)((dynamic)range).Comment;
            if (comment is null)
                return new JsonObject { ["type"] = "note", ["present"] = false, ["range"] = address };
            return new JsonObject
            {
                ["type"] = "note",
                ["present"] = true,
                ["range"] = address,
                ["text"] = TryInvokeText(comment),
                ["visible"] = Js(TryComBool(comment, "Visible")),
            };
        }
        catch
        {
            return new JsonObject { ["type"] = "note", ["present"] = false, ["range"] = address };
        }
        finally
        {
            RotHelper.ReleaseComReference(comment);
            RotHelper.ReleaseComReference(range);
        }
    }

    private static JsonObject ReadHyperlinkState(object sheet, string address)
    {
        object? range = null;
        object? links = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            links = (object)((dynamic)range).Hyperlinks;
            var count = Convert.ToInt32(((dynamic)links).Count, CultureInfo.InvariantCulture);
            if (count <= 0)
            {
                var empty = ReadCellContent(range);
                return new JsonObject
                {
                    ["type"] = "hyperlink",
                    ["present"] = false,
                    ["range"] = address,
                    ["cellFormula"] = Json.GetString(empty, "cellFormula"),
                    ["cellValue"] = empty["cellValue"]?.DeepClone(),
                    ["font"] = Json.GetObj(empty, "font")?.DeepClone(),
                };
            }
            object? first = null;
            try
            {
                first = (object)((dynamic)links).Item(1);
                var content = ReadCellContent(range);
                return new JsonObject
                {
                    ["type"] = "hyperlink",
                    ["present"] = true,
                    ["range"] = address,
                    ["address"] = TryComString(first, "Address"),
                    ["subAddress"] = TryComString(first, "SubAddress"),
                    ["textToDisplay"] = TryComString(first, "TextToDisplay"),
                    ["screenTip"] = TryComString(first, "ScreenTip"),
                    ["cellFormula"] = Json.GetString(content, "cellFormula"),
                    ["cellValue"] = content["cellValue"]?.DeepClone(),
                    ["font"] = Json.GetObj(content, "font")?.DeepClone(),
                };
            }
            finally { RotHelper.ReleaseComReference(first); }
        }
        catch
        {
            return new JsonObject { ["type"] = "hyperlink", ["present"] = false, ["range"] = address };
        }
        finally
        {
            RotHelper.ReleaseComReference(links);
            RotHelper.ReleaseComReference(range);
        }
    }

    private static JsonNode CaptureRangeFormulas(object sheet, string address)
    {
        object? range = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            return RangeToJson(range, out _, formulas: true, maxCells: MaxDataSnapshotCells);
        }
        finally { RotHelper.ReleaseComReference(range); }
    }

    private static void RestoreRangeFormulas(object sheet, string address, JsonArray formulas)
    {
        object? range = null;
        try
        {
            range = (object)((dynamic)sheet).Range(address);
            var rows = formulas.Count;
            var cols = formulas[0] is JsonArray first ? first.Count : 0;
            var values = new object[rows, cols];
            for (var r = 0; r < rows; r++)
            {
                if (formulas[r] is not JsonArray row) continue;
                for (var c = 0; c < cols && c < row.Count; c++)
                    values[r, c] = NodeToComValue(row[c]) ?? "";
            }
            ((dynamic)range).Formula = values;
        }
        finally { RotHelper.ReleaseComReference(range); }
    }

    private static void ApplyTotalsColumns(object table, JsonObject op)
    {
        var columns = Json.GetArr(op, "columns") ?? Json.GetArr(op, "totals");
        if (columns is null) return;
        object? listColumns = null;
        try
        {
            listColumns = (object)((dynamic)table).ListColumns;
            foreach (var node in columns.OfType<JsonObject>())
            {
                object? column = null;
                try
                {
                    column = ResolveListColumn(listColumns, node["column"]);
                    if (!ExcelDataObjectCatalog.TryTotalsFunction(Json.GetString(node, "function"), out var function))
                        throw new InvalidOperationException("[EXCEL_DATA_UNSUPPORTED] totals function is not mapped");
                    ((dynamic)column).TotalsCalculation = function;
                    if (function == ExcelDataObjectCatalog.XlTotalsCalculationCustom)
                    {
                        object? totalsCell = null;
                        try
                        {
                            totalsCell = (object)((dynamic)column).TotalsRowRange;
                            ((dynamic)totalsCell).Formula = Json.GetString(node, "formula");
                        }
                        finally { RotHelper.ReleaseComReference(totalsCell); }
                    }
                }
                finally { RotHelper.ReleaseComReference(column); }
            }
        }
        finally { RotHelper.ReleaseComReference(listColumns); }
    }

    private static object ResolveListColumn(object listColumns, JsonNode? key)
    {
        if (key is JsonValue number && number.TryGetValue<int>(out var index))
            return (object)((dynamic)listColumns).Item(index);
        if (key is JsonValue text && text.TryGetValue<string>(out var name) && !string.IsNullOrWhiteSpace(name))
            return (object)((dynamic)listColumns).Item(name);
        throw new InvalidOperationException("table column key must be a 1-based index or header name");
    }

    private static List<DataSortKey> ResolveSortKeys(object sheet, string rangeAddress, JsonArray keys, bool table,
        object? tableObject = null)
    {
        if (!ExcelDataOperationsContract.TryParseA1(rangeAddress, out var box))
            throw new InvalidOperationException($"cannot resolve sort keys for '{rangeAddress}'");
        var result = new List<DataSortKey>();
        foreach (var node in keys.OfType<JsonObject>())
        {
            var column = ResolveColumnNumber(sheet, box, node["column"], table, tableObject);
            var address = $"{ColName(column)}{box.Row}:{ColName(column)}{box.EndRow}";
            result.Add(new DataSortKey(address, ExcelDataObjectCatalog.SortOrder(Json.GetString(node, "order"))));
        }
        return result;
    }

    private static int ResolveColumnNumber(object sheet, ExcelA1Box box, JsonNode? key, bool table, object? tableObject)
    {
        if (key is JsonValue number && number.TryGetValue<int>(out var index))
            return box.Column + index - 1;
        if (key is not JsonValue text || !text.TryGetValue<string>(out var token) || string.IsNullOrWhiteSpace(token))
            throw new InvalidOperationException("sort/filter column is missing");
        if (ExcelDataOperationsContract.TryParseCell(token + "1", out _, out var colFromLetter) &&
            token.All(ch => char.IsLetter(ch) || ch == '$'))
            return colFromLetter;
        if (table && tableObject is not null)
        {
            object? columns = null;
            object? column = null;
            object? range = null;
            try
            {
                columns = (object)((dynamic)tableObject).ListColumns;
                column = (object)((dynamic)columns).Item(token);
                range = (object)((dynamic)column).Range;
                return Convert.ToInt32(((dynamic)range).Column, CultureInfo.InvariantCulture);
            }
            finally
            {
                RotHelper.ReleaseComReference(range);
                RotHelper.ReleaseComReference(column);
                RotHelper.ReleaseComReference(columns);
            }
        }
        object? header = null;
        try
        {
            header = (object)((dynamic)sheet).Range($"{ColName(box.Column)}{box.Row}:{ColName(box.EndColumn)}{box.Row}");
            var values = RangeToJson(header, out _, formulas: false, maxCells: 256);
            if (values.Count > 0 && values[0] is JsonArray row)
            {
                for (var i = 0; i < row.Count; i++)
                {
                    if (string.Equals(row[i]?.ToString(), token, StringComparison.OrdinalIgnoreCase))
                        return box.Column + i;
                }
            }
        }
        finally { RotHelper.ReleaseComReference(header); }
        throw new InvalidOperationException($"[EXCEL_COLUMN_NOT_FOUND] column '{token}' was not found in {ColName(box.Column)}{box.Row}:{ColName(box.EndColumn)}{box.Row}");
    }

    private static void InvokeRangeSort(object comRange, List<DataSortKey> keys, bool hasHeaders)
    {
        object? key1 = null;
        object? key2 = null;
        object? key3 = null;
        try
        {
            dynamic range = comRange;
            dynamic sheet = range.Worksheet;
            key1 = (object)sheet.Range(keys[0].Address);
            if (keys.Count > 1) key2 = (object)sheet.Range(keys[1].Address);
            if (keys.Count > 2) key3 = (object)sheet.Range(keys[2].Address);
            range.Sort(
                Key1: key1,
                Order1: keys[0].Order,
                Key2: key2 ?? Type.Missing,
                Type: Type.Missing,
                Order2: keys.Count > 1 ? keys[1].Order : Type.Missing,
                Key3: key3 ?? Type.Missing,
                Order3: keys.Count > 2 ? keys[2].Order : Type.Missing,
                Header: ExcelDataObjectCatalog.HasHeaders(hasHeaders),
                OrderCustom: Type.Missing,
                MatchCase: false,
                Orientation: ExcelDataObjectCatalog.XlSortRows);
        }
        finally
        {
            RotHelper.ReleaseComReference(key3);
            RotHelper.ReleaseComReference(key2);
            RotHelper.ReleaseComReference(key1);
        }
    }

    private static bool SortApplied(object comRange)
    {
        try
        {
            _ = ((dynamic)comRange).Address(false, false);
            return true;
        }
        catch { return false; }
    }

    private static void ApplyOneFilter(object comRange, DataRangeLease target, JsonObject criterion)
    {
        if (!ExcelDataOperationsContract.TryParseA1(target.Address, out var box))
            throw new InvalidOperationException("filter target is not a rectangular A1 range");
        var column = ResolveColumnNumber(target.Sheet, box, criterion["column"], target.ListObject is not null, target.ListObject);
        var field = column - box.Column + 1;
        var filterOp = Json.GetString(criterion, "operator") ?? "equals";
        var value = CriterionText(criterion["value"]);
        dynamic range = comRange;
        switch (filterOp.ToLowerInvariant())
        {
            case "equals":
                range.AutoFilter(field, value);
                break;
            case "notequals":
                range.AutoFilter(field, "<>" + value);
                break;
            case "greaterthan":
                range.AutoFilter(field, ">" + value);
                break;
            case "lessthan":
                range.AutoFilter(field, "<" + value);
                break;
            case "greaterorequal":
                range.AutoFilter(field, ">=" + value);
                break;
            case "lessorequal":
                range.AutoFilter(field, "<=" + value);
                break;
            case "between":
                range.AutoFilter(field, ">=" + value, ExcelDataObjectCatalog.XlAnd, "<=" + CriterionText(criterion["value2"]));
                break;
            case "contains":
                range.AutoFilter(field, "*" + value + "*");
                break;
            case "beginswith":
                range.AutoFilter(field, value + "*");
                break;
            case "endswith":
                range.AutoFilter(field, "*" + value);
                break;
            case "values":
                var values = (Json.GetArr(criterion, "values") ?? new JsonArray())
                    .Select(CriterionText)
                    .ToArray();
                range.AutoFilter(field, values, ExcelDataObjectCatalog.XlFilterValues);
                break;
            default:
                throw new InvalidOperationException($"[EXCEL_DATA_UNSUPPORTED] filter operator '{filterOp}' is not mapped");
        }
    }

    private static string CriterionText(JsonNode? node)
    {
        if (node is null) return "";
        if (node is JsonValue value)
        {
            if (value.TryGetValue<string>(out var text)) return text;
            if (ExcelDataOperationsContract.TryGetFiniteNumber(value, out var number))
                return ExcelDataObjectCatalog.FormatInvariant(number);
            if (value.TryGetValue<bool>(out var flag)) return flag ? "TRUE" : "FALSE";
        }
        return node.ToJsonString();
    }

    private static object AddConditionalRule(object conditions, JsonObject rule)
    {
        var type = Json.GetString(rule, "type")!;
        switch (type)
        {
            case "cellValue":
                if (!ExcelDataObjectCatalog.TryComparisonOperator(Json.GetString(rule, "operator"), out var comparison))
                    throw new InvalidOperationException("[EXCEL_DATA_UNSUPPORTED] conditional operator is not mapped");
                return (object)((dynamic)conditions).Add(
                    ExcelDataObjectCatalog.XlCellValue,
                    comparison,
                    Json.GetString(rule, "formula1"),
                    Json.GetString(rule, "formula2") ?? Type.Missing);
            case "expression":
                return (object)((dynamic)conditions).Add(
                    ExcelDataObjectCatalog.XlExpression,
                    Type.Missing,
                    Json.GetString(rule, "formula1"));
            case "uniqueValues":
            case "duplicateValues":
            {
                object added = ((dynamic)conditions).AddUniqueValues();
                ((dynamic)added).DupeUnique = type == "uniqueValues"
                    ? ExcelDataObjectCatalog.XlUnique
                    : ExcelDataObjectCatalog.XlDuplicate;
                return added;
            }
            case "colorScale":
            {
                var points = Json.GetInt(rule, "points") ?? 2;
                object scale = ((dynamic)conditions).AddColorScale(points);
                ApplyColorScaleColors(scale, rule, points);
                return scale;
            }
            case "dataBar":
                return (object)((dynamic)conditions).AddDatabar();
            case "iconSet":
            {
                object added = ((dynamic)conditions).AddIconSetCondition();
                if (ExcelDataObjectCatalog.TryIconSet(Json.GetString(rule, "iconSet"), out var iconSet))
                {
                    object? icon = null;
                    try
                    {
                        icon = (object)((dynamic)added).IconSet;
                        ((dynamic)icon).ID = iconSet;
                    }
                    finally { RotHelper.ReleaseComReference(icon); }
                }
                return added;
            }
            default:
                throw new InvalidOperationException($"[EXCEL_DATA_UNSUPPORTED] conditional rule '{type}' is not implemented");
        }
    }

    private static void ApplyColorScaleColors(object scale, JsonObject rule, int points)
    {
        var colors = new[] { "minColor", "midColor", "maxColor" };
        for (var index = 0; index < points; index++)
        {
            if (!rule.ContainsKey(colors[index])) continue;
            if (!ExcelStyleContract.TryParseColor(rule[colors[index]], out var ole)) continue;
            object? criteria = null;
            object? format = null;
            try
            {
                criteria = (object)((dynamic)scale).ColorScaleCriteria(index + 1);
                format = (object)((dynamic)criteria).FormatColor;
                ((dynamic)format).Color = ole;
            }
            finally
            {
                RotHelper.ReleaseComReference(format);
                RotHelper.ReleaseComReference(criteria);
            }
        }
    }

    private static void ApplyConditionalStyle(object condition, JsonObject? style)
    {
        if (style is null) return;
        object? font = null;
        object? interior = null;
        try
        {
            try { font = (object)((dynamic)condition).Font; } catch { font = null; }
            try { interior = (object)((dynamic)condition).Interior; } catch { interior = null; }
            if (font is not null)
            {
                if (style.ContainsKey("bold") || style.ContainsKey("fontBold"))
                    ((dynamic)font).Bold = Json.GetBool(style, "bold") || Json.GetBool(style, "fontBold");
                if (style.ContainsKey("italic") || style.ContainsKey("fontItalic"))
                    ((dynamic)font).Italic = Json.GetBool(style, "italic") || Json.GetBool(style, "fontItalic");
                if ((style.ContainsKey("fontColor") || style.ContainsKey("fill")) &&
                    ExcelStyleContract.TryParseColor(style["fontColor"] ?? style["fill"], out var fontColor))
                    ((dynamic)font).Color = fontColor;
            }
            if (interior is not null &&
                (style.ContainsKey("fillColor") || style.ContainsKey("interiorColor") || style.ContainsKey("fill")) &&
                ExcelStyleContract.TryParseColor(style["fillColor"] ?? style["interiorColor"] ?? style["fill"], out var fill))
            {
                ((dynamic)interior).Pattern = ExcelStyleContract.XlPatternSolid;
                ((dynamic)interior).Color = fill;
            }
        }
        finally
        {
            RotHelper.ReleaseComReference(interior);
            RotHelper.ReleaseComReference(font);
        }
    }

    private static void ApplyAdvertisedChartSeries(object workbook, object chart, JsonArray? series, string defaultSheet)
    {
        ExcelChartSeriesApply.ApplyData(chart, series, text =>
            ResolveChartSeriesRange(workbook, defaultSheet, text));
    }

    private static object? ResolveChartSeriesRange(object workbook, string defaultSheet, string text)
    {
        if (!ExcelDataRangeBinding.TrySourceSheet(defaultSheet, text, out var sheetName, out var address, out _))
            return null;
        object? sheet = null;
        try
        {
            sheet = (object)GetSheet(workbook, sheetName);
            return (object)((dynamic)sheet).Range(address);
        }
        finally { RotHelper.ReleaseComReference(sheet); }
    }

    private static bool TrySetChartSourceData(object chart, DataRangeLease source, JsonObject op)
    {
        object? comRange = null;
        try
        {
            comRange = (object)((dynamic)source.Sheet).Range(source.Address);
            ((dynamic)chart).SetSourceData(comRange, ExcelDataObjectCatalog.PlotBy(Json.GetString(op, "plotBy")));
            return true;
        }
        catch (Exception) when (ExcelChartSeriesContract.HasSeriesData(Json.GetArr(op, "series")))
        {
            return false;
        }
        finally { RotHelper.ReleaseComReference(comRange); }
    }

    private static void ApplyChartTitleLegend(object chart, JsonObject op)
    {
        if (op.ContainsKey("title"))
        {
            var title = Json.GetString(op, "title");
            if (string.IsNullOrEmpty(title))
            {
                ((dynamic)chart).HasTitle = false;
            }
            else
            {
                ((dynamic)chart).HasTitle = true;
                object? chartTitle = null;
                try
                {
                    chartTitle = (object)((dynamic)chart).ChartTitle;
                    ((dynamic)chartTitle).Text = title;
                }
                finally { RotHelper.ReleaseComReference(chartTitle); }
            }
        }
        if (op.ContainsKey("hasLegend"))
            ((dynamic)chart).HasLegend = Json.GetBool(op, "hasLegend");
        if (op.ContainsKey("legendPosition") && TryComBool(chart, "HasLegend") == true &&
            ExcelDataObjectCatalog.TryLegendPosition(Json.GetString(op, "legendPosition"), out var position))
        {
            object? legend = null;
            try
            {
                legend = (object)((dynamic)chart).Legend;
                ((dynamic)legend).Position = position;
            }
            finally { RotHelper.ReleaseComReference(legend); }
        }
    }

    private static void ApplyChartPresentation(object chart, JsonObject op)
    {
        ApplyChartAreaFormat(chart, "ChartArea", op, "chartFill", "chartBorder");
        ApplyChartAreaFormat(chart, "PlotArea", op, "plotFill", "plotBorder");
        ApplyChartSeriesPresentation(chart, Json.GetArr(op, "series"));
        ApplyChartAxes(chart, Json.GetObj(op, "axes"));
        ApplyChartPlotArea(chart, Json.GetObj(op, "plotArea"));
    }

    private static void ApplyChartAreaFormat(object chart, string areaName, JsonObject op, string fillField,
        string lineField)
    {
        if (!op.ContainsKey(fillField) && !op.ContainsKey(lineField)) return;
        object? area = null;
        object? format = null;
        object? fill = null;
        object? line = null;
        try
        {
            area = areaName == "PlotArea" ? (object)((dynamic)chart).PlotArea : (object)((dynamic)chart).ChartArea;
            format = (object)((dynamic)area).Format;
            if (op.ContainsKey(fillField))
            {
                fill = (object)((dynamic)format).Fill;
                ApplyChartFillOrNone(fill, op[fillField]);
            }
            if (op.ContainsKey(lineField))
            {
                line = (object)((dynamic)format).Line;
                ApplyChartLineOrNone(line, op[lineField]);
            }
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"[EXCEL_CHART_FORMAT] {areaName} fill/line failed: {ex.Message}", ex);
        }
        finally
        {
            RotHelper.ReleaseComReference(line);
            RotHelper.ReleaseComReference(fill);
            RotHelper.ReleaseComReference(format);
            RotHelper.ReleaseComReference(area);
        }
    }

    private static void ApplyChartFillOrNone(object fill, JsonNode? node)
    {
        var text = node is JsonValue value && value.TryGetValue<string>(out var s) ? s : null;
        if (string.Equals(text, "none", StringComparison.OrdinalIgnoreCase))
        {
            ((dynamic)fill).Visible = ExcelDataObjectCatalog.MsoFalse;
            return;
        }
        if (!ExcelStyleContract.TryParseColor(node, out var color)) return;
        ((dynamic)fill).Visible = ExcelDataObjectCatalog.MsoTrue;
        ((dynamic)fill).ForeColor.RGB = color;
    }

    private static void ApplyChartLineOrNone(object line, JsonNode? node)
    {
        var text = node is JsonValue value && value.TryGetValue<string>(out var s) ? s : null;
        if (string.Equals(text, "none", StringComparison.OrdinalIgnoreCase))
        {
            ((dynamic)line).Visible = ExcelDataObjectCatalog.MsoFalse;
            return;
        }
        if (!ExcelStyleContract.TryParseColor(node, out var color)) return;
        ((dynamic)line).Visible = ExcelDataObjectCatalog.MsoTrue;
        ((dynamic)line).ForeColor.RGB = color;
    }

    private static void ApplyChartSeriesPresentation(object chart, JsonArray? series)
    {
        if (series is null) return;
        object? collection = null;
        try
        {
            collection = (object)((dynamic)chart).SeriesCollection();
            var count = Convert.ToInt32(((dynamic)collection).Count, CultureInfo.InvariantCulture);
            for (var i = 0; i < series.Count && i < count; i++)
            {
                if (series[i] is not JsonObject requested) continue;
                object? item = null;
                object? format = null;
                object? line = null;
                try
                {
                    item = (object)((dynamic)collection).Item(i + 1);
                    if (requested.ContainsKey("lineColor") || requested.ContainsKey("lineWeight"))
                    {
                        format = (object)((dynamic)item).Format;
                        line = (object)((dynamic)format).Line;
                        ((dynamic)line).Visible = ExcelDataObjectCatalog.MsoTrue;
                        if (ExcelStyleContract.TryParseColor(requested["lineColor"], out var color))
                            ((dynamic)line).ForeColor.RGB = color;
                        if (ExcelDataOperationsContract.TryGetFiniteNumber(requested["lineWeight"], out var weight))
                            ((dynamic)line).Weight = weight;
                    }
                    if (ExcelDataObjectCatalog.TryMarkerStyle(Json.GetString(requested, "marker"), out var marker))
                        ((dynamic)item).MarkerStyle = marker;
                }
                finally
                {
                    RotHelper.ReleaseComReference(line);
                    RotHelper.ReleaseComReference(format);
                    RotHelper.ReleaseComReference(item);
                }
            }
        }
        finally { RotHelper.ReleaseComReference(collection); }
    }

    private static void ApplyChartAxes(object chart, JsonObject? axes) =>
        ExcelChartAxesApply.Apply(chart, axes);

    private static void ApplyChartPlotArea(object chart, JsonObject? plot)
    {
        if (plot is null) return;
        object? plotArea = null;
        object? chartArea = null;
        try
        {
            plotArea = (object)((dynamic)chart).PlotArea;
            if (Json.GetBool(plot, "spanChart"))
            {
                chartArea = (object)((dynamic)chart).ChartArea;
                ((dynamic)plotArea).Left = 0;
                ((dynamic)plotArea).Top = 0;
                ((dynamic)plotArea).Width = ((dynamic)chartArea).Width;
                ((dynamic)plotArea).Height = ((dynamic)chartArea).Height;
            }
            if (plot.ContainsKey("left") && ExcelDataOperationsContract.TryGetFiniteNumber(plot["left"], out var left))
                ((dynamic)plotArea).Left = left;
            if (plot.ContainsKey("top") && ExcelDataOperationsContract.TryGetFiniteNumber(plot["top"], out var top))
                ((dynamic)plotArea).Top = top;
            if (plot.ContainsKey("width") && ExcelDataOperationsContract.TryGetFiniteNumber(plot["width"], out var width))
                ((dynamic)plotArea).Width = width;
            if (plot.ContainsKey("height") && ExcelDataOperationsContract.TryGetFiniteNumber(plot["height"], out var height))
                ((dynamic)plotArea).Height = height;
        }
        finally
        {
            RotHelper.ReleaseComReference(chartArea);
            RotHelper.ReleaseComReference(plotArea);
        }
    }

    private static DataPosition ReadPosition(JsonObject op, double left, double top, double width, double height,
        bool optionalSize = false)
    {
        var position = Json.GetObj(op, "position");
        if (position is null) return new DataPosition(left, top, width, height);
        return new DataPosition(
            ReadPoint(position, "left", left),
            ReadPoint(position, "top", top),
            ReadPoint(position, "width", optionalSize ? width : width == 0 ? DefaultChartWidth : width),
            ReadPoint(position, "height", optionalSize ? height : height == 0 ? DefaultChartHeight : height));
    }

    private static double ReadPoint(JsonObject position, string field, double fallback)
    {
        if (!position.ContainsKey(field) ||
            !ExcelDataOperationsContract.TryGetFiniteNumber(position[field], out var value))
            return fallback;
        return value;
    }

    private static void ApplyPosition(object target, DataPosition position)
    {
        if (position.Left > 0 || position.Top > 0 || position.Width > 0 || position.Height > 0)
        {
            if (position.Width > 0) ((dynamic)target).Width = position.Width;
            if (position.Height > 0) ((dynamic)target).Height = position.Height;
            ((dynamic)target).Left = position.Left;
            ((dynamic)target).Top = position.Top;
        }
    }

    private static void SetOptionalBool(object listObject, JsonObject op, string field, string comProperty,
        string? alias = null)
    {
        if (!op.ContainsKey(field)) return;
        var value = Json.GetBool(op, field);
        switch (comProperty)
        {
            case "ShowHeaders": ((dynamic)listObject).ShowHeaders = value; break;
            case "ShowTotals": ((dynamic)listObject).ShowTotals = value; break;
            case "ShowAutoFilter": ((dynamic)listObject).ShowAutoFilter = value; break;
            case "ShowTableStyleRowStripes": ((dynamic)listObject).ShowTableStyleRowStripes = value; break;
            case "ShowTableStyleColumnStripes": ((dynamic)listObject).ShowTableStyleColumnStripes = value; break;
            case "ShowRowStripes": ((dynamic)listObject).ShowTableStyleRowStripes = value; break;
            default:
                if (alias == "ShowTableStyleRowStripes")
                    ((dynamic)listObject).ShowTableStyleRowStripes = value;
                else
                    throw new InvalidOperationException($"[EXCEL_DATA_UNSUPPORTED] table flag '{comProperty}' is not mapped");
                break;
        }
    }

    private static void SetOptionalString(object target, JsonObject op, string field, string comProperty)
    {
        if (!op.ContainsKey(field)) return;
        var value = Json.GetString(op, field) ?? "";
        switch (comProperty)
        {
            case "InputTitle": ((dynamic)target).InputTitle = value; break;
            case "InputMessage": ((dynamic)target).InputMessage = value; break;
            case "ErrorTitle": ((dynamic)target).ErrorTitle = value; break;
            case "ErrorMessage": ((dynamic)target).ErrorMessage = value; break;
            default:
                throw new InvalidOperationException($"[EXCEL_DATA_UNSUPPORTED] validation text '{comProperty}' is not mapped");
        }
    }

    private static int VerifyTable(JsonObject actual, JsonObject op, DataRangeLease range, List<string> mismatches)
    {
        var checkedCells = 1;
        var requested = Json.GetString(op, "name");
        if (!string.IsNullOrWhiteSpace(requested) &&
            !string.Equals(Json.GetString(actual, "name"), requested, StringComparison.OrdinalIgnoreCase))
            mismatches.Add($"{range.SheetName}: created ListObject name '{Json.GetString(actual, "name")}' != '{requested}'");
        var actualRange = Json.GetString(actual, "range");
        if (string.IsNullOrWhiteSpace(actualRange))
            mismatches.Add($"{range.SheetName}: created ListObject has no Range");
        else if (!AddressesEqual(actualRange, range.Address))
            mismatches.Add($"{range.SheetName}: created ListObject range '{actualRange}' != '{range.Address}'");
        if (op.ContainsKey("hasHeaders") && actual.ContainsKey("showHeaders") &&
            Json.GetBool(actual, "showHeaders") != Json.GetBool(op, "hasHeaders"))
            mismatches.Add($"{range.SheetName}: ShowHeaders readback {Json.GetBool(actual, "showHeaders")} != hasHeaders {Json.GetBool(op, "hasHeaders")}");
        if (op.ContainsKey("showTotals") && Json.GetBool(actual, "showTotals") != Json.GetBool(op, "showTotals"))
            mismatches.Add($"{range.SheetName}: ShowTotals readback mismatch");
        var styleName = Json.GetString(op, "styleName");
        if (!string.IsNullOrWhiteSpace(styleName))
        {
            var actualStyle = Json.GetString(actual, "styleName") ?? "";
            if (!actualStyle.Contains(styleName, StringComparison.OrdinalIgnoreCase))
                mismatches.Add($"{range.SheetName}: TableStyle readback '{actualStyle}' does not contain '{styleName}'");
        }
        ExcelDataReadbackContract.CompareRequestedTotals(actual, op, range.SheetName, mismatches);
        return checkedCells;
    }

    private static int VerifyTableStyle(JsonObject actual, JsonObject op, DataTableLease table, List<string> mismatches)
    {
        var checkedCells = 1;
        var styleName = Json.GetString(op, "styleName");
        if (!string.IsNullOrWhiteSpace(styleName))
        {
            var actualStyle = Json.GetString(actual, "styleName") ?? "";
            if (!actualStyle.Contains(styleName, StringComparison.OrdinalIgnoreCase))
                mismatches.Add($"{table.SheetName}!{table.Name}: TableStyle readback '{actualStyle}' does not contain '{styleName}'");
        }
        foreach (var (field, key) in new[]
                 {
                     ("showHeaders", "showHeaders"), ("showTotals", "showTotals"),
                     ("showAutoFilter", "showAutoFilter"), ("showRowStripes", "showRowStripes"),
                     ("showColumnStripes", "showColumnStripes"),
                 })
        {
            if (op.ContainsKey(field) && Json.GetBool(actual, key) != Json.GetBool(op, field))
                mismatches.Add($"{table.SheetName}!{table.Name}: {field} readback mismatch");
        }
        return checkedCells;
    }

    private static int VerifyChart(JsonObject actual, JsonObject op, DataRangeLease? source, List<string> mismatches)
    {
        var checkedCells = 1;
        var label = $"{Json.GetString(actual, "sheet")}!{Json.GetString(actual, "name")}";
        if (op.ContainsKey("chartType") &&
            !ExcelChartSeriesContract.ChartLevelTypeMatches(
                Json.GetString(actual, "chartType"),
                Json.GetString(op, "chartType"),
                Json.GetArr(op, "series")))
            mismatches.Add($"{label}: ChartType readback '{Json.GetString(actual, "chartType")}' != '{Json.GetString(op, "chartType")}'");
        if (op.ContainsKey("title") &&
            !string.Equals(Json.GetString(actual, "title") ?? "", Json.GetString(op, "title") ?? "", StringComparison.Ordinal))
            mismatches.Add($"{label}: ChartTitle readback mismatch");
        if (op.ContainsKey("hasLegend") && Json.GetBool(actual, "hasLegend") != Json.GetBool(op, "hasLegend"))
            mismatches.Add($"{label}: HasLegend readback mismatch");
        ExcelDataReadbackContract.CompareRequestedChartSeries(actual, op, label, mismatches);
        if (source is not null)
        {
            var hasSeriesData = ExcelChartSeriesContract.HasSeriesData(Json.GetArr(op, "series"));
            if ((Json.GetInt(actual, "seriesCount") ?? 0) <= 0 && !hasSeriesData)
                mismatches.Add($"{label}: chart has no SeriesCollection after SetSourceData");
            var seriesFormula = Json.GetString(actual, "series1Formula") ?? "";
            if (!hasSeriesData && !ChartFormulaMentionsSource(
                    seriesFormula, source, Json.GetString(actual, "ownerWorkbook")))
                mismatches.Add($"{label}: series formula '{seriesFormula}' does not reference source {source.SheetName}!{source.Address}");
        }
        VerifyChartPresentation(actual, op, label, mismatches);
        VerifyChartDetails(actual, op, label, mismatches);
        return checkedCells;
    }

    private static bool ChartFormulaMentionsSource(string? formula, DataRangeLease source, string? ownerWorkbook)
    {
        if (string.IsNullOrWhiteSpace(formula)) return false;
        if (ExcelChartSeriesContract.FormulaMentionsSource(
                formula, source.CacheSource, source.SheetName, ownerWorkbook) ||
            ExcelChartSeriesContract.FormulaMentionsSource(
                formula,
                ExcelDataRangeBinding.QualifiedA1(source.SheetName, source.Address),
                source.SheetName,
                ownerWorkbook))
            return true;
        foreach (var area in ExcelDataRangeBinding.AddressAreas(source.Address))
        {
            if (ExcelChartSeriesContract.FormulaMentionsSource(
                    formula,
                    ExcelDataRangeBinding.QualifiedA1(source.SheetName, area),
                    source.SheetName,
                    ownerWorkbook))
                return true;
        }
        return false;
    }

    private static void VerifyChartPresentation(JsonObject actual, JsonObject op, string label,
        List<string> mismatches)
    {
        if (string.Equals(Json.GetString(op, "chartFill"), "none", StringComparison.OrdinalIgnoreCase) &&
            Json.GetBool(actual, "chartFillVisible"))
            mismatches.Add($"{label}: ChartArea fill is still visible");
        if (string.Equals(Json.GetString(op, "plotFill"), "none", StringComparison.OrdinalIgnoreCase) &&
            Json.GetBool(actual, "plotFillVisible"))
            mismatches.Add($"{label}: PlotArea fill is still visible");
        if (string.Equals(Json.GetString(op, "chartBorder"), "none", StringComparison.OrdinalIgnoreCase) &&
            Json.GetBool(actual, "chartLineVisible"))
            mismatches.Add($"{label}: ChartArea border is still visible");
        if (string.Equals(Json.GetString(op, "plotBorder"), "none", StringComparison.OrdinalIgnoreCase) &&
            Json.GetBool(actual, "plotLineVisible"))
            mismatches.Add($"{label}: PlotArea border is still visible");

        var requestedSeries = Json.GetArr(op, "series");
        var actualSeries = Json.GetArr(actual, "series") ?? new JsonArray();
        if (requestedSeries is not null)
        {
            for (var i = 0; i < requestedSeries.Count; i++)
            {
                if (requestedSeries[i] is not JsonObject req) continue;
                var seriesIndex = Json.GetInt(req, "index") ?? (i + 1);
                var got = seriesIndex <= actualSeries.Count ? actualSeries[seriesIndex - 1] as JsonObject : null;
                if (got is null) continue;
                var marker = Json.GetString(req, "marker");
                if (!string.IsNullOrWhiteSpace(marker) &&
                    !string.Equals(Json.GetString(got, "marker"), marker, StringComparison.OrdinalIgnoreCase))
                    mismatches.Add($"{label}: series[{i}].marker '{Json.GetString(got, "marker")}' != '{marker}'");
                if (ExcelDataOperationsContract.TryGetFiniteNumber(req["lineWeight"], out var weight))
                {
                    var actualWeight = ExcelDataOperationsContract.TryGetFiniteNumber(got["lineWeight"], out var aw)
                        ? aw
                        : double.NaN;
                    if (double.IsNaN(actualWeight) || Math.Abs(actualWeight - weight) > 0.15)
                        mismatches.Add($"{label}: series[{i}].lineWeight {actualWeight} != {weight}");
                }
                if (ExcelStyleContract.TryParseColor(req["lineColor"], out var color))
                {
                    var actualColor = ExcelDataOperationsContract.TryGetFiniteNumber(got["lineColor"], out var ac)
                        ? ac
                        : double.NaN;
                    if (double.IsNaN(actualColor) || Math.Abs(actualColor - color) > 0.5)
                        mismatches.Add($"{label}: series[{i}].lineColor {actualColor} != {color}");
                }
            }
        }

        ExcelChartAxesContract.CompareRequested(actual, op, label, mismatches);
    }

    private static bool? ReadFormatVisible(object chart, string areaName, bool fill)
    {
        object? area = null;
        object? format = null;
        object? target = null;
        try
        {
            area = areaName == "PlotArea" ? (object)((dynamic)chart).PlotArea : (object)((dynamic)chart).ChartArea;
            format = (object)((dynamic)area).Format;
            target = fill ? (object)((dynamic)format).Fill : (object)((dynamic)format).Line;
            var visible = TryComInt(target, "Visible");
            return visible is null ? null : visible != ExcelDataObjectCatalog.MsoFalse;
        }
        catch { return null; }
        finally
        {
            RotHelper.ReleaseComReference(target);
            RotHelper.ReleaseComReference(format);
            RotHelper.ReleaseComReference(area);
        }
    }

    private static JsonObject ReadChartAxes(object chart)
    {
        var axes = new JsonObject
        {
            ["category"] = ReadOneChartAxis(chart, ExcelDataObjectCatalog.XlCategory, ExcelDataObjectCatalog.XlPrimary, required: true),
            ["value"] = ReadOneChartAxis(chart, ExcelDataObjectCatalog.XlValue, ExcelDataObjectCatalog.XlPrimary, required: true),
        };
        var secondary = ReadOneChartAxis(chart, ExcelDataObjectCatalog.XlValue, ExcelDataObjectCatalog.XlSecondary, required: false);
        if (secondary is not null)
            axes["valueSecondary"] = secondary;
        return axes;
    }

    private static JsonObject? ReadOneChartAxis(object chart, int axisType, int axisGroup, bool required)
    {
        object? axis = null;
        object? labels = null;
        object? title = null;
        try
        {
            try
            {
                axis = (object)((dynamic)chart).Axes(axisType, axisGroup);
            }
            catch
            {
                return required
                    ? UnreadableAxis(axisGroup)
                    : null;
            }

            var tick = TryComInt(axis, "TickLabelPosition");
            var visible = tick != ExcelDataObjectCatalog.XlTickLabelPositionNone;
            var hasTitle = TryComBool(axis, "HasTitle");
            string? titleText = null;
            if (hasTitle == true)
            {
                title = (object)((dynamic)axis).AxisTitle;
                titleText = ReadComString(title, "Text");
            }
            try { labels = (object)((dynamic)axis).TickLabels; }
            catch { labels = null; }
            return new JsonObject
            {
                ["group"] = ExcelDataObjectCatalog.TokenFor(
                    ExcelDataObjectCatalog.AxisGroups, axisGroup, $"raw:{axisGroup}"),
                ["visible"] = visible,
                ["minimum"] = Js(TryComDouble(axis, "MinimumScale")),
                ["maximum"] = Js(TryComDouble(axis, "MaximumScale")),
                ["numberFormat"] = labels is null ? null : TryComString(labels, "NumberFormat"),
                ["title"] = titleText,
            };
        }
        catch
        {
            return UnreadableAxis(axisGroup);
        }
        finally
        {
            RotHelper.ReleaseComReference(title);
            RotHelper.ReleaseComReference(labels);
            RotHelper.ReleaseComReference(axis);
        }
    }

    private static JsonObject UnreadableAxis(int axisGroup) => new()
    {
        ["group"] = ExcelDataObjectCatalog.TokenFor(
            ExcelDataObjectCatalog.AxisGroups, axisGroup, $"raw:{axisGroup}"),
        ["visible"] = null,
        ["unreadable"] = true,
    };

    private static int VerifyPicture(JsonObject actual, JsonObject op, string sheetName, List<string> mismatches)
    {
        var checkedCells = 1;
        if (Json.GetInt(actual, "shapeType") != ExcelDataObjectCatalog.MsoPicture)
            mismatches.Add($"{sheetName}!{Json.GetString(actual, "name")}: Shape.Type is not msoPicture");
        var requested = Json.GetString(op, "name");
        if (!string.IsNullOrWhiteSpace(requested) &&
            !string.Equals(Json.GetString(actual, "name"), requested, StringComparison.OrdinalIgnoreCase))
            mismatches.Add($"{sheetName}: picture name '{Json.GetString(actual, "name")}' != '{requested}'");
        if (Json.GetObj(op, "position") is JsonObject position)
        {
            foreach (var field in new[] { "left", "top", "width", "height" })
            {
                if (!position.ContainsKey(field) ||
                    !ExcelDataOperationsContract.TryGetFiniteNumber(position[field], out var expected))
                    continue;
                var actualValue = actual[field] is JsonValue value &&
                                  ExcelDataOperationsContract.TryGetFiniteNumber(value, out var number)
                    ? number
                    : double.NaN;
                if (double.IsNaN(actualValue) || Math.Abs(actualValue - expected) > 1.5)
                    mismatches.Add($"{sheetName}!{Json.GetString(actual, "name")}: {field} readback {actualValue} != {expected}");
            }
        }
        if (op.ContainsKey("lockAspectRatio") && (actual["lockAspectRatio"] is not JsonValue flag || !flag.TryGetValue<bool>(out var observedAspect) || observedAspect != Json.GetBool(op, "lockAspectRatio")))
            mismatches.Add($"{sheetName}: picture aspect setting readback mismatch");
        if (op.ContainsKey("rotation") && (!ExcelDataOperationsContract.TryGetFiniteNumber(actual["rotation"], out var observedRotation) || Math.Abs(observedRotation - ExcelShapeFormatContract.ReadFiniteNumber(op["rotation"]!)) > 0.01))
            mismatches.Add($"{sheetName}: picture rotation readback mismatch");
        if (Json.GetObj(op, "crop") is JsonObject crop)
            foreach (var side in crop)
                if (!ExcelDataOperationsContract.TryGetFiniteNumber(Json.GetObj(actual, "crop")?[side.Key], out var observedCrop) || Math.Abs(observedCrop - ExcelShapeFormatContract.ReadFiniteNumber(side.Value!)) > 0.01)
                    mismatches.Add($"{sheetName}: picture crop.{side.Key} readback mismatch");
        return checkedCells;
    }

    private static void VerifySortOrder(object sheet, string rangeAddress, JsonArray keys, bool hasHeaders,
        List<string> mismatches, JsonNode? before, object? tableObject = null)
    {
        if (keys.Count == 0) return;
        if (!ExcelDataOperationsContract.TryParseA1(DataLocalAddress(rangeAddress), out var box))
        {
            mismatches.Add($"{rangeAddress}: sort readback range is not rectangular");
            return;
        }

        var specs = new List<(int GridColumn, bool Descending)>();
        foreach (var key in keys.OfType<JsonObject>())
        {
            try
            {
                var column = ResolveColumnNumber(sheet, box, key["column"], tableObject is not null, tableObject);
                specs.Add((
                    GridColumn: column - box.Column,
                    Descending: ExcelDataObjectCatalog.SortOrder(Json.GetString(key, "order")) ==
                                ExcelDataObjectCatalog.XlDescending));
            }
            catch (Exception ex)
            {
                mismatches.Add($"{rangeAddress}: sort key column could not be resolved: {ex.Message}");
                return;
            }
        }

        var after = CaptureRangeValues(sheet, rangeAddress) as JsonArray;
        ExcelDataReadbackContract.CompareSortReadback(
            before as JsonArray, after, specs, hasHeaders, rangeAddress, mismatches);
    }

    private static void VerifyDuplicatesRemoved(object sheet, string rangeAddress, JsonArray? columns, bool hasHeaders,
        List<string> mismatches)
    {
        if (!ExcelDataOperationsContract.TryParseA1(DataLocalAddress(rangeAddress), out var box))
        {
            mismatches.Add($"{rangeAddress}: remove_duplicates readback range is not rectangular");
            return;
        }
        object? comRange = null;
        try
        {
            comRange = (object)((dynamic)sheet).Range(box.Address);
            var grid = RangeToJson(comRange, out _, formulas: false, maxCells: MaxDataSnapshotCells);
            var start = hasHeaders ? 1 : 0;
            var keyIndexes = new List<int>();
            if (columns is { Count: > 0 })
            {
                foreach (var node in columns)
                {
                    if (node is JsonValue value && value.TryGetValue<int>(out var index) && index >= 1)
                        keyIndexes.Add(index - 1);
                }
            }
            else
            {
                var width = grid.Count > 0 && grid[0] is JsonArray first ? first.Count : box.Columns;
                for (var i = 0; i < width; i++) keyIndexes.Add(i);
            }

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (var r = start; r < grid.Count; r++)
            {
                if (grid[r] is not JsonArray row) continue;
                var key = string.Join('\u001f', keyIndexes.Select(i => i < row.Count ? CellKey(row[i]) : ""));
                if (!seen.Add(key))
                {
                    mismatches.Add($"{rangeAddress}: duplicate key still present after RemoveDuplicates");
                    return;
                }
            }
        }
        finally { RotHelper.ReleaseComReference(comRange); }
    }

    private static void VerifyTextToColumns(object sheet, string sourceAddress, string afterAddress,
        JsonNode? before, JsonObject op, List<string> mismatches)
    {
        if (!ExcelDataOperationsContract.TryParseA1(DataLocalAddress(afterAddress), out _))
        {
            mismatches.Add($"{afterAddress}: text_to_columns destination is not rectangular");
            return;
        }
        var after = CaptureRangeValues(sheet, afterAddress) as JsonArray;
        ExcelDataReadbackContract.CompareTextToColumnsReadback(
            before as JsonArray, after, op, $"{sourceAddress}->{afterAddress}", mismatches);
    }

    private static void VerifyPivotRequested(JsonObject actual, JsonObject op, string sheetName,
        List<string> mismatches, bool requireAggregates = false)
    {
        var requestedName = Json.GetString(op, "name");
        var actualName = Json.GetString(actual, "name");
        var label = $"{sheetName}!{actualName ?? requestedName}";
        ExcelDataReadbackContract.CompareRequestedPivot(
            actual, op, label, mismatches,
            requireAggregates: requireAggregates || Json.GetArr(op, "values") is { Count: > 0 });
    }

    private static void VerifyRequestedShape(JsonObject actual, JsonObject op, string sheetName, bool textbox,
        List<string> mismatches)
    {
        var requested = Json.GetString(op, "name");
        var actualName = Json.GetString(actual, "name");
        var label = $"{sheetName}!{actualName ?? requested}";
        if (!string.IsNullOrWhiteSpace(requested) &&
            !string.Equals(actualName, requested, StringComparison.OrdinalIgnoreCase))
            mismatches.Add($"{sheetName}: {(textbox ? "textbox" : "shape")} name '{actualName}' != '{requested}'");

        if (textbox)
        {
            if (Json.GetInt(actual, "shapeType") != ExcelDataObjectCatalog.MsoTextBox)
                mismatches.Add($"{label}: Type is not msoTextBox");
        }
        else if (ExcelDataObjectCatalog.TryShapeType(Json.GetString(op, "shapeType"), out var expectedType) &&
                 Json.GetInt(actual, "shapeType") != expectedType)
        {
            mismatches.Add($"{label}: AutoShapeType {Json.GetInt(actual, "shapeType")} != {Json.GetString(op, "shapeType")}");
        }

        if (op.ContainsKey("text") &&
            !string.Equals(Json.GetString(actual, "text") ?? "", Json.GetString(op, "text") ?? "", StringComparison.Ordinal))
            mismatches.Add($"{label}: shape text readback mismatch");

        if (Json.GetObj(op, "position") is JsonObject position)
        {
            foreach (var field in new[] { "left", "top", "width", "height" })
            {
                if (!position.ContainsKey(field) ||
                    !ExcelDataOperationsContract.TryGetFiniteNumber(position[field], out var expected))
                    continue;
                var actualValue = actual[field] is JsonValue value &&
                                  ExcelDataOperationsContract.TryGetFiniteNumber(value, out var number)
                    ? number
                    : double.NaN;
                if (double.IsNaN(actualValue) || Math.Abs(actualValue - expected) > 1.5)
                    mismatches.Add($"{label}: {field} readback {actualValue} != {expected}");
            }
        }
        ExcelShapeFormatContract.CompareRequestedReadback(actual, op, label, mismatches);
    }

    private static void VerifyCalculatedColumn(object column, string requestedFormula, string label,
        List<string> mismatches)
    {
        object? body = null;
        try
        {
            body = (object)((dynamic)column).DataBodyRange;
            var actual = TryComString(body, "Formula") ?? "";
            if (string.IsNullOrWhiteSpace(actual) || !actual.StartsWith('='))
            {
                mismatches.Add($"{label}: calculated column Formula is empty after Add");
                return;
            }
            var requested = requestedFormula.Trim();
            if (string.Equals(actual, requested, StringComparison.OrdinalIgnoreCase))
                return;
            var requestedBody = requested.TrimStart('=');
            if (!string.IsNullOrWhiteSpace(requestedBody) &&
                actual.Contains(requestedBody, StringComparison.OrdinalIgnoreCase))
                return;
            mismatches.Add($"{label}: calculated column Formula '{actual}' does not match '{requested}'");
        }
        finally { RotHelper.ReleaseComReference(body); }
    }

    private static void CompareRequestedFieldList(JsonObject actual, string actualField, JsonObject op,
        string requestedField, string label, List<string> mismatches)
    {
        if (!op.ContainsKey(requestedField)) return;
        var requested = Json.GetArr(op, requestedField) ?? new JsonArray();
        var actualList = Json.GetArr(actual, actualField) ?? new JsonArray();
        foreach (var node in requested)
        {
            var name = node?.ToString();
            if (string.IsNullOrWhiteSpace(name)) continue;
            if (!actualList.Any(item => string.Equals(item?.ToString(), name, StringComparison.OrdinalIgnoreCase)))
                mismatches.Add($"{label}: missing {requestedField} '{name}'");
        }
    }

    private static string DataLocalAddress(string address)
    {
        try { return ExcelRangeReference.Parse(address).Address; }
        catch (FormatException) { return address; }
    }

    private static bool IsEmptySortCell(JsonNode? cell) =>
        cell is null || string.IsNullOrWhiteSpace(cell.ToString());

    private static string CellKey(JsonNode? cell) =>
        cell is null ? "" : cell.ToString();

    private static int CompareSortValues(JsonNode? left, JsonNode? right)
    {
        if (TryNodeNumber(left, out var ln) && TryNodeNumber(right, out var rn))
            return ln.CompareTo(rn);
        return string.Compare(left?.ToString() ?? "", right?.ToString() ?? "", StringComparison.OrdinalIgnoreCase);
    }

    private static bool TryNodeNumber(JsonNode? node, out double number)
    {
        if (ExcelDataOperationsContract.TryGetFiniteNumber(node, out number)) return true;
        if (node is JsonValue value && value.TryGetValue<string>(out var text) &&
            double.TryParse(text, NumberStyles.Float, CultureInfo.InvariantCulture, out number) &&
            double.IsFinite(number))
            return true;
        number = 0;
        return false;
    }

    private static string? ResolveValidationFormula(JsonObject op, string field)
    {
        if (op[field] is JsonValue value)
        {
            if (value.TryGetValue<string>(out var text)) return text;
            if (ExcelDataOperationsContract.TryGetFiniteNumber(value, out var number))
                return ExcelDataObjectCatalog.FormatInvariant(number);
        }
        return null;
    }

    private static string? ResolveListSource(string sheetName, JsonObject op)
    {
        var source = Json.GetString(op, "source");
        if (string.IsNullOrWhiteSpace(source)) return ResolveValidationFormula(op, "formula1");
        if (ExcelDataOperationsContract.TryParseA1(source, out _))
        {
            var parsed = ExcelRangeReference.Parse(source);
            var sheet = parsed.SheetName ?? sheetName;
            return $"={QuoteSheetName(sheet)}!{parsed.Address}";
        }
        return source;
    }

    private static bool AddressesEqual(string? left, string? right)
    {
        if (string.IsNullOrWhiteSpace(left) || string.IsNullOrWhiteSpace(right)) return false;
        if (string.Equals(left, right, StringComparison.OrdinalIgnoreCase)) return true;
        return ExcelDataOperationsContract.TryParseA1(left, out var a) &&
               ExcelDataOperationsContract.TryParseA1(right, out var b) &&
               a == b;
    }

    private static bool RefersToMatches(string? actual, string expected) =>
        ExcelFormulaReference.SemanticEquals(actual, expected);

    private static bool HyperlinkMatches(JsonObject actual, JsonObject op)
    {
        var address = Json.GetString(op, "address");
        var sub = Json.GetString(op, "subAddress");
        if (!string.IsNullOrWhiteSpace(address) &&
            !string.Equals(Json.GetString(actual, "address") ?? "", address, StringComparison.OrdinalIgnoreCase))
            return false;
        if (!string.IsNullOrWhiteSpace(sub) &&
            !string.Equals(Json.GetString(actual, "subAddress") ?? "", sub, StringComparison.OrdinalIgnoreCase))
            return false;
        return true;
    }

    private static string NameRef(string scope, string name, JsonObject op) =>
        scope == "sheet"
            ? $"{Json.GetString(Json.GetObj(op, "target"), "sheet")}!{name}"
            : name;

    private static bool NameMatches(string? filter, string? name) =>
        string.IsNullOrWhiteSpace(filter) ||
        string.Equals(filter, name, StringComparison.OrdinalIgnoreCase);

    private static string QuoteSheetName(string sheetName) =>
        ExcelFormulaReference.QuoteSheetName(sheetName);

    private static string DataComMessage(Exception ex, string stage, string? op) =>
        ComFailure.FormatMessage(ex, stage, op ?? "data", 1);

    private static string? ReadComString(object target, string property) => TryComString(target, property);

    private static string? TryTableStyleName(object table)
    {
        object? style = null;
        try
        {
            try { style = (object)((dynamic)table).TableStyle; }
            catch { return null; }
            return ReadComString(style, "Name");
        }
        finally { RotHelper.ReleaseComReference(style); }
    }

    private static string? TryComString(object target, string property)
    {
        try
        {
            object? value = property switch
            {
                "Name" => ((dynamic)target).Name,
                "Text" => ((dynamic)target).Text,
                "RefersTo" => ((dynamic)target).RefersTo,
                "Comment" => ((dynamic)target).Comment,
                "TableStyle" => ((dynamic)target).TableStyle,
                "Formula" => ((dynamic)target).Formula,
                "Formula1" => ((dynamic)target).Formula1,
                "Formula2" => ((dynamic)target).Formula2,
                "Address" => ((dynamic)target).Address,
                "SubAddress" => ((dynamic)target).SubAddress,
                "TextToDisplay" => ((dynamic)target).TextToDisplay,
                "ScreenTip" => ((dynamic)target).ScreenTip,
                "InputTitle" => ((dynamic)target).InputTitle,
                "InputMessage" => ((dynamic)target).InputMessage,
                "ErrorTitle" => ((dynamic)target).ErrorTitle,
                "ErrorMessage" => ((dynamic)target).ErrorMessage,
                "NumberFormat" => ((dynamic)target).NumberFormat,
                "SourceData" => ((dynamic)target).SourceData,
                "FullName" => ((dynamic)target).FullName,
                _ => null,
            };
            return Convert.ToString(value, CultureInfo.InvariantCulture);
        }
        catch { return null; }
    }

    private static JsonNode? Js(bool? value) => value is null ? null : JsonValue.Create(value.Value);
    private static JsonNode? Js(int? value) => value is null ? null : JsonValue.Create(value.Value);
    private static JsonNode? Js(double? value) => value is null ? null : JsonValue.Create(value.Value);

    private static int? TryComInt(object target, string property)
    {
        try
        {
            object? value = property switch
            {
                "Type" => ((dynamic)target).Type,
                "ChartType" => ((dynamic)target).ChartType,
                "Position" => ((dynamic)target).Position,
                "TotalsCalculation" => ((dynamic)target).TotalsCalculation,
                "Operator" => ((dynamic)target).Operator,
                "LockAspectRatio" => ((dynamic)target).LockAspectRatio,
                "AlertStyle" => ((dynamic)target).AlertStyle,
                "Count" => ((dynamic)target).Count,
                "DupeUnique" => ((dynamic)target).DupeUnique,
                "Function" => ((dynamic)target).Function,
                "Orientation" => ((dynamic)target).Orientation,
                "ID" => ((dynamic)target).ID,
                "MarkerStyle" => ((dynamic)target).MarkerStyle,
                "AxisGroup" => ((dynamic)target).AxisGroup,
                "Visible" => ((dynamic)target).Visible,
                "TickLabelPosition" => ((dynamic)target).TickLabelPosition,
                "Underline" => ((dynamic)target).Underline,
                "ConnectionSiteCount" => ((dynamic)target).ConnectionSiteCount,
                "BeginConnected" => ((dynamic)target).BeginConnected,
                "EndConnected" => ((dynamic)target).EndConnected,
                "BeginConnectionSite" => ((dynamic)target).BeginConnectionSite,
                "EndConnectionSite" => ((dynamic)target).EndConnectionSite,
                "HorizontalFlip" => ((dynamic)target).HorizontalFlip,
                "VerticalFlip" => ((dynamic)target).VerticalFlip,
                "Placement" => ((dynamic)target).Placement,
                _ => null,
            };
            return value is null ? null : Convert.ToInt32(value, CultureInfo.InvariantCulture);
        }
        catch { return null; }
    }

    private static bool? TryComBool(object target, string property)
    {
        try
        {
            object? value = property switch
            {
                "ShowTotals" => ((dynamic)target).ShowTotals,
                "ShowHeaders" => ((dynamic)target).ShowHeaders,
                "ShowAutoFilter" => ((dynamic)target).ShowAutoFilter,
                "ShowTableStyleRowStripes" => ((dynamic)target).ShowTableStyleRowStripes,
                "ShowTableStyleColumnStripes" => ((dynamic)target).ShowTableStyleColumnStripes,
                "HasTitle" => ((dynamic)target).HasTitle,
                "HasLegend" => ((dynamic)target).HasLegend,
                "AutoFilterMode" => ((dynamic)target).AutoFilterMode,
                "FilterMode" => ((dynamic)target).FilterMode,
                "IgnoreBlank" => ((dynamic)target).IgnoreBlank,
                "InCellDropdown" => ((dynamic)target).InCellDropdown,
                "Visible" => ((dynamic)target).Visible,
                "ShowInput" => ((dynamic)target).ShowInput,
                "ShowError" => ((dynamic)target).ShowError,
                "On" => ((dynamic)target).On,
                "Bold" => ((dynamic)target).Bold,
                "Italic" => ((dynamic)target).Italic,
                "BuiltIn" => ((dynamic)target).BuiltIn,
                "ProtectStructure" => ((dynamic)target).ProtectStructure,
                "ProtectWindows" => ((dynamic)target).ProtectWindows,
                _ => null,
            };
            if (value is null or DBNull) return null;
            if (value is bool flag) return flag;
            return Convert.ToInt32(value, CultureInfo.InvariantCulture) switch { 0 => false, -1 or 1 => true, _ => null };
        }
        catch { return null; }
    }

    private static double? TryComDouble(object target, string property)
    {
        try
        {
            object? value = property switch
            {
                "Left" => ((dynamic)target).Left,
                "Top" => ((dynamic)target).Top,
                "Width" => ((dynamic)target).Width,
                "Height" => ((dynamic)target).Height,
                "Color" => ((dynamic)target).Color,
                "Weight" => ((dynamic)target).Weight,
                "RGB" => ((dynamic)target).RGB,
                "MinimumScale" => ((dynamic)target).MinimumScale,
                "MaximumScale" => ((dynamic)target).MaximumScale,
                "Size" => ((dynamic)target).Size,
                "Rotation" => ((dynamic)target).Rotation,
                "CropLeft" => ((dynamic)target).CropLeft,
                "CropTop" => ((dynamic)target).CropTop,
                "CropRight" => ((dynamic)target).CropRight,
                "CropBottom" => ((dynamic)target).CropBottom,
                _ => null,
            };
            if (value is null or DBNull) return null;
            var number = Convert.ToDouble(value, CultureInfo.InvariantCulture);
            return double.IsFinite(number) ? number : null;
        }
        catch { return null; }
    }

    private static string? TryAddress(object? range)
    {
        if (range is null) return null;
        try { return Convert.ToString(((dynamic)range).Address(false, false), CultureInfo.InvariantCulture); }
        catch { return null; }
    }

    private static string? TryParentSheetName(object table)
    {
        object? parent = null;
        try
        {
            parent = (object)((dynamic)table).Parent;
            return ReadComString(parent, "Name");
        }
        catch { return null; }
        finally { RotHelper.ReleaseComReference(parent); }
    }

    private static string? TryInvokeText(object comment)
    {
        try { return Convert.ToString(((dynamic)comment).Text(), CultureInfo.InvariantCulture); }
        catch
        {
            try { return Convert.ToString(((dynamic)comment).Text, CultureInfo.InvariantCulture); }
            catch { return null; }
        }
    }
}
