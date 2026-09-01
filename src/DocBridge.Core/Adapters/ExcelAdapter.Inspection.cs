using System.Globalization;
using System.Runtime.InteropServices;
using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class ExcelAdapter
{
    private const int XlCellTypeFormulas = -4123;
    private const int XlErrors = 16;

    private static JsonObject InspectExcelDiagnostics(dynamic app)
    {
        var protectedViews = new JsonArray();
        object? protectedViewWindows = null;
        try
        {
            protectedViewWindows = (object)app.ProtectedViewWindows;
            dynamic windows = protectedViewWindows;
            var count = Convert.ToInt32(windows.Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                object? protectedViewWindow = null;
                try
                {
                    protectedViewWindow = (object)windows.Item(index);
                    dynamic window = protectedViewWindow;
                    protectedViews.Add(new JsonObject
                    {
                        ["caption"] = Convert.ToString(window.Caption, CultureInfo.InvariantCulture),
                        ["sourcePath"] = TryString(() => window.SourcePath),
                        ["sourceName"] = TryString(() => window.SourceName),
                    });
                }
                finally { RotHelper.ReleaseComObject(protectedViewWindow); }
            }
        }
        catch { }
        finally { RotHelper.ReleaseComObject(protectedViewWindows); }

        var ready = TryBool(() => app.Ready);
        var interactive = TryBool(() => app.Interactive);
        var workbookCount = 0;
        object? workbooks = null;
        try
        {
            workbooks = (object)app.Workbooks;
            workbookCount = Convert.ToInt32(((dynamic)workbooks).Count, CultureInfo.InvariantCulture);
        }
        catch { }
        finally { RotHelper.ReleaseComObject(workbooks); }
        var state = protectedViews.Count > 0 && workbookCount == 0
            ? "EXCEL_PROTECTED_VIEW"
            : interactive == false || ready == false
                ? "EXCEL_MODAL_OR_BUSY"
                : "CHECK_PASSED";
        return new JsonObject
        {
            ["ok"] = state == "CHECK_PASSED",
            ["app"] = "excel",
            ["scope"] = "diagnostics",
            ["state"] = state,
            ["errorCode"] = state == "CHECK_PASSED" ? null : state,
            ["retryable"] = state == "EXCEL_MODAL_OR_BUSY",
            ["ready"] = ready,
            ["interactive"] = interactive,
            ["openWorkbookCount"] = workbookCount,
            ["protectedViewWindows"] = protectedViews,
            ["userAction"] = state switch
            {
                "EXCEL_PROTECTED_VIEW" => "Excel에서 편집 사용을 눌러 제한된 보기를 해제한 뒤 다시 시도하세요.",
                "EXCEL_MODAL_OR_BUSY" => "Excel의 대화상자 또는 수식 편집을 끝낸 뒤 다시 시도하세요.",
                _ => "조치가 필요하지 않습니다.",
            },
        };
    }

    private static JsonObject InspectWorkbook(dynamic app, dynamic workbook, JsonObject args, string scope)
    {
        return scope switch
        {
            "scan" => ScanWorkbook(workbook, args),
            "objects" => InspectObjects(workbook, args),
            "errors" => InspectFormulaErrors(workbook, args),
            _ => throw new ArgumentException($"unknown Excel inspect scope '{scope}' (scan|objects|errors|diagnostics)"),
        };
    }

    private static JsonObject ScanWorkbook(dynamic workbook, JsonObject args)
    {
        var requestedSheet = Json.GetString(args, "sheet");
        var sheets = new JsonArray();
        var count = Convert.ToInt32(workbook.Worksheets.Count, CultureInfo.InvariantCulture);
        for (var index = 1; index <= count; index++)
        {
            dynamic sheet = workbook.Worksheets.Item(index);
            var name = Convert.ToString(sheet.Name, CultureInfo.InvariantCulture) ?? "";
            if (!string.IsNullOrWhiteSpace(requestedSheet) &&
                !string.Equals(name, requestedSheet, StringComparison.OrdinalIgnoreCase)) continue;
            dynamic used = sheet.UsedRange;
            sheets.Add(new JsonObject
            {
                ["name"] = name,
                ["visible"] = TryInt(() => sheet.Visible),
                ["usedRange"] = TryString(() => used.Address(false, false)),
                ["usedRows"] = TryInt(() => used.Rows.Count),
                ["usedColumns"] = TryInt(() => used.Columns.Count),
                ["listObjectCount"] = TryInt(() => sheet.ListObjects.Count),
                ["chartCount"] = TryInt(() => sheet.ChartObjects().Count),
                ["shapeCount"] = TryInt(() => sheet.Shapes.Count),
                ["pivotTableCount"] = TryInt(() => sheet.PivotTables().Count),
                ["formulaCellCount"] = CountSpecialCells(used, XlCellTypeFormulas, null),
                ["formulaErrorCount"] = CountSpecialCells(used, XlCellTypeFormulas, XlErrors),
            });
        }
        return new JsonObject
        {
            ["ok"] = true,
            ["app"] = "excel",
            ["scope"] = "scan",
            ["workbook"] = TryString(() => workbook.FullName),
            ["saved"] = TryBool(() => workbook.Saved),
            ["sheetCount"] = count,
            ["sheets"] = sheets,
            ["coverage"] = new JsonObject { ["complete"] = string.IsNullOrWhiteSpace(requestedSheet), ["returnedSheets"] = sheets.Count },
        };
    }

    private static JsonObject InspectObjects(dynamic workbook, JsonObject args)
    {
        var requestedSheet = Json.GetString(args, "sheet");
        var limit = Math.Clamp(Json.GetInt(args, "limit") ?? 500, 1, 2000);
        var objects = new JsonArray();
        var truncated = false;
        void Add(JsonObject item)
        {
            if (objects.Count >= limit) { truncated = true; return; }
            objects.Add(item);
        }

        var sheetCount = Convert.ToInt32(workbook.Worksheets.Count, CultureInfo.InvariantCulture);
        for (var sheetIndex = 1; sheetIndex <= sheetCount; sheetIndex++)
        {
            dynamic sheet = workbook.Worksheets.Item(sheetIndex);
            var sheetName = Convert.ToString(sheet.Name, CultureInfo.InvariantCulture) ?? "";
            if (!string.IsNullOrWhiteSpace(requestedSheet) &&
                !string.Equals(sheetName, requestedSheet, StringComparison.OrdinalIgnoreCase)) continue;

            try
            {
                var count = Convert.ToInt32(sheet.ListObjects.Count, CultureInfo.InvariantCulture);
                for (var index = 1; index <= count; index++)
                {
                    dynamic item = sheet.ListObjects.Item(index);
                    Add(new JsonObject { ["type"] = "table", ["sheet"] = sheetName, ["name"] = TryString(() => item.Name), ["range"] = TryString(() => item.Range.Address(false, false)) });
                }
            }
            catch { }
            try
            {
                dynamic charts = sheet.ChartObjects();
                var count = Convert.ToInt32(charts.Count, CultureInfo.InvariantCulture);
                for (var index = 1; index <= count; index++)
                {
                    dynamic item = charts.Item(index);
                    Add(new JsonObject { ["type"] = "chart", ["sheet"] = sheetName, ["name"] = TryString(() => item.Name), ["chartType"] = TryInt(() => item.Chart.ChartType) });
                }
            }
            catch { }
            try
            {
                var count = Convert.ToInt32(sheet.PivotTables().Count, CultureInfo.InvariantCulture);
                for (var index = 1; index <= count; index++)
                {
                    dynamic item = sheet.PivotTables().Item(index);
                    Add(new JsonObject { ["type"] = "pivotTable", ["sheet"] = sheetName, ["name"] = TryString(() => item.Name), ["range"] = TryString(() => item.TableRange2.Address(false, false)) });
                }
            }
            catch { }
            try
            {
                var count = Convert.ToInt32(sheet.Shapes.Count, CultureInfo.InvariantCulture);
                for (var index = 1; index <= count; index++)
                {
                    dynamic item = sheet.Shapes.Item(index);
                    Add(new JsonObject { ["type"] = "shape", ["sheet"] = sheetName, ["name"] = TryString(() => item.Name), ["shapeType"] = TryInt(() => item.Type) });
                }
            }
            catch { }
        }

        try
        {
            var count = Convert.ToInt32(workbook.Names.Count, CultureInfo.InvariantCulture);
            for (var index = 1; index <= count; index++)
            {
                dynamic item = workbook.Names.Item(index);
                Add(new JsonObject { ["type"] = "definedName", ["name"] = TryString(() => item.Name), ["refersTo"] = TryString(() => item.RefersTo) });
            }
        }
        catch { }

        return new JsonObject
        {
            ["ok"] = true,
            ["app"] = "excel",
            ["scope"] = "objects",
            ["workbook"] = TryString(() => workbook.FullName),
            ["objects"] = objects,
            ["coverage"] = new JsonObject { ["limit"] = limit, ["returned"] = objects.Count, ["truncated"] = truncated, ["complete"] = !truncated },
        };
    }

    private static JsonObject InspectFormulaErrors(dynamic workbook, JsonObject args)
    {
        var requestedSheet = Json.GetString(args, "sheet");
        var limit = Math.Clamp(Json.GetInt(args, "limit") ?? 500, 1, 2000);
        var errors = new JsonArray();
        var total = 0L;
        var sheetCount = Convert.ToInt32(workbook.Worksheets.Count, CultureInfo.InvariantCulture);
        for (var sheetIndex = 1; sheetIndex <= sheetCount; sheetIndex++)
        {
            dynamic sheet = workbook.Worksheets.Item(sheetIndex);
            var sheetName = Convert.ToString(sheet.Name, CultureInfo.InvariantCulture) ?? "";
            if (!string.IsNullOrWhiteSpace(requestedSheet) &&
                !string.Equals(sheetName, requestedSheet, StringComparison.OrdinalIgnoreCase)) continue;
            try
            {
                dynamic cells = sheet.UsedRange.SpecialCells(XlCellTypeFormulas, XlErrors);
                total += Convert.ToInt64(cells.CountLarge, CultureInfo.InvariantCulture);
                foreach (dynamic cell in cells.Cells)
                {
                    if (errors.Count >= limit) break;
                    errors.Add(new JsonObject
                    {
                        ["sheet"] = sheetName,
                        ["address"] = TryString(() => cell.Address(false, false)),
                        ["formula"] = TryString(() => cell.Formula),
                        ["displayText"] = TryString(() => cell.Text),
                    });
                }
            }
            catch (COMException ex) when (ex.HResult == unchecked((int)0x800A03EC))
            {
                // SpecialCells는 일치 셀이 없을 때 1004를 던진다. 정상적인 0건이다.
            }
        }
        return new JsonObject
        {
            ["ok"] = true,
            ["app"] = "excel",
            ["scope"] = "errors",
            ["workbook"] = TryString(() => workbook.FullName),
            ["errorCount"] = total,
            ["errorsFound"] = errors,
            ["coverage"] = new JsonObject { ["limit"] = limit, ["returned"] = errors.Count, ["truncated"] = total > errors.Count, ["complete"] = total <= errors.Count },
        };
    }

    private static long CountSpecialCells(dynamic range, int cellType, int? valueType)
    {
        object? cells = null;
        try
        {
            cells = (object)(valueType is null ? range.SpecialCells(cellType) : range.SpecialCells(cellType, valueType.Value));
            return Convert.ToInt64(((dynamic)cells).CountLarge, CultureInfo.InvariantCulture);
        }
        catch (COMException ex) when (ex.HResult == unchecked((int)0x800A03EC)) { return 0; }
        catch { return 0; }
        finally { RotHelper.ReleaseComObject(cells); }
    }

    private JsonObject ExcelErrorResult(Exception ex)
    {
        var disconnected = IsComDisconnected(ex);
        var busy = IsCallRejected(ex);
        var diagnostics = new JsonObject();
        try
        {
            var app = AttachExcel();
            if (app is not null) diagnostics = InspectExcelDiagnostics((dynamic)app);
        }
        catch { }
        var protectedView = Json.GetString(diagnostics, "state") == "EXCEL_PROTECTED_VIEW";
        var code = protectedView ? "EXCEL_PROTECTED_VIEW" : disconnected ? "EXCEL_COM_DISCONNECTED" : busy ? "EXCEL_MODAL_OR_BUSY" : "EXCEL_READ_FAILED";
        return new JsonObject
        {
            ["ok"] = false,
            ["app"] = App,
            ["errorCode"] = code,
            ["retryable"] = busy || disconnected,
            ["errors"] = new JsonArray(ex.Message),
            ["diagnostics"] = diagnostics,
            ["userAction"] = code switch
            {
                "EXCEL_PROTECTED_VIEW" => "Excel에서 편집 사용을 눌러 제한된 보기를 해제하세요.",
                "EXCEL_MODAL_OR_BUSY" => "Excel 대화상자나 수식 편집을 끝낸 뒤 다시 시도하세요.",
                "EXCEL_COM_DISCONNECTED" => "Excel이 재시작되었습니다. 다음 호출에서 자동 재연결됩니다.",
                _ => "workbook·sheet·range 인수를 확인하세요.",
            },
        };
    }

    private static string? TryString(Func<object?> getter) { try { return Convert.ToString(getter(), CultureInfo.InvariantCulture); } catch { return null; } }
    private static int? TryInt(Func<object?> getter) { try { return Convert.ToInt32(getter(), CultureInfo.InvariantCulture); } catch { return null; } }
    private static bool? TryBool(Func<object?> getter) { try { return Convert.ToBoolean(getter(), CultureInfo.InvariantCulture); } catch { return null; } }
}
