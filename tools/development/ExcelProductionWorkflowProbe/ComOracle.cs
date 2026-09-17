using System.Runtime.InteropServices;

namespace DocBridge.Development.ExcelProductionWorkflowProbe;

/// <summary>
/// Isolated native Excel oracle. Never product authoring.
/// New Excel.Application instance only (never GetActiveObject).
/// Opens the SOURCE read-only, UpdateLinks=0, macros forced off.
/// Closes only that book and Quits only this instance.
/// </summary>
internal static class ComOracle
{
    private const int MsoAutomationSecurityForceDisable = 3;
    private const string InstanceCaption = "DocBridge-Oracle-ReadOnly-Source";

    public static JsonObject ReadWorkbook(ProbeOptions options, string path)
    {
        if (!options.OracleCom)
            return new JsonObject { ["ok"] = false, ["skipped"] = "oracle-com-disabled" };
        if (!options.LeaseOk)
            return new JsonObject { ["ok"] = false, ["skipped"] = "no-coordinator-lease" };
        if (!string.Equals(Environment.GetEnvironmentVariable("DOCBRIDGE_E2E"), "1", StringComparison.Ordinal))
            return new JsonObject { ["ok"] = false, ["skipped"] = "DOCBRIDGE_E2E!=1" };

        var source = Path.GetFullPath(path);
        SourceIntegrity.RefuseIfProtectedPath(source + ".oracle-touch", options.SourceXlsx);
        if (!string.Equals(source, Path.GetFullPath(options.SourceXlsx), StringComparison.OrdinalIgnoreCase))
            return new JsonObject { ["ok"] = false, ["error"] = "COM oracle may open only the configured source xlsx" };
        if (SourceIntegrity.IsForbiddenUserBook(source))
            return new JsonObject { ["ok"] = false, ["error"] = "refusing user untitled book" };

        var before = SourceIntegrity.Capture(source, options.ExpectedSha256);
        if (JsonUtil.Bool(before, "match") != true)
            return new JsonObject { ["ok"] = false, ["error"] = "source hash mismatch before COM", ["hash"] = before };

        var type = Type.GetTypeFromProgID("Excel.Application");
        if (type is null)
            return new JsonObject { ["ok"] = false, ["error"] = "Excel.Application ProgID not registered" };

        var pdf = Path.Combine(options.OutputDir, "golden-source.pdf");
        if (File.Exists(pdf))
            return new JsonObject { ["ok"] = false, ["error"] = "refusing to overwrite " + pdf };

        object? app = null;
        object? books = null;
        object? wb = null;
        object? sheet = null;
        JsonObject? payload = null;
        string? openedFull = null;
        try
        {
            ExcelMessageFilter.Register();
            app = CreateExcelWithRetry(type);
            dynamic dApp = app;
            books = dApp.Workbooks;
            if (Convert.ToInt32(((dynamic)books).Count, CultureInfo.InvariantCulture) != 0)
            {
                books = null;
                app = null;
                throw new InvalidOperationException("Excel instance already has workbooks; refusing to reuse a user session. Not quitting or releasing that Application.");
            }
            dApp.Visible = false;
            dApp.DisplayAlerts = false;
            dApp.ScreenUpdating = false;
            dApp.EnableEvents = false;
            dApp.AskToUpdateLinks = false;
            dApp.Caption = InstanceCaption;
            try { dApp.AutomationSecurity = MsoAutomationSecurityForceDisable; } catch { /* older Excel */ }
            dynamic dBooks = books;
            var missing = Type.Missing;
            wb = dBooks.Open(
                source,
                0,
                true,
                missing, missing, missing,
                true,
                missing, missing, missing, missing, missing,
                false);
            dynamic dWb = wb;
            openedFull = Convert.ToString(dWb.FullName, CultureInfo.InvariantCulture);
            if (!string.Equals(Path.GetFullPath(openedFull ?? ""), source, StringComparison.OrdinalIgnoreCase))
                throw new InvalidOperationException($"opened FullName '{openedFull}' is not the source");
            if (Convert.ToInt32(dBooks.Count, CultureInfo.InvariantCulture) != 1)
                throw new InvalidOperationException("owned Excel instance has unexpected extra workbooks; aborting without closing them");

            sheet = dWb.Worksheets.Item(options.ScheduleSheet);
            dynamic dSheet = sheet;
            payload = ReadSheet(dSheet, options.ScheduleSheet);
            payload["openedFullName"] = openedFull;
            payload["readOnly"] = true;
            payload["updateLinks"] = 0;
            payload["macros"] = "AutomationSecurityForceDisable";
            payload["authoring"] = false;
            payload["instanceCaption"] = InstanceCaption;

            try
            {
                Directory.CreateDirectory(options.OutputDir);
                dSheet.ExportAsFixedFormat(0, pdf);
                payload["goldenPdf"] = File.Exists(pdf) ? pdf : null;
                payload["goldenPdfBytes"] = File.Exists(pdf) ? new FileInfo(pdf).Length : 0;
            }
            catch (Exception ex)
            {
                payload["goldenPdf"] = null;
                payload["goldenPdfError"] = ex.Message;
            }
        }
        catch (Exception ex)
        {
            payload = new JsonObject { ["ok"] = false, ["error"] = ex.ToString(), ["authoring"] = false };
        }
        finally
        {
            TryCloseOwned(wb, source);
            TryQuitOwned(app);
            Release(sheet);
            Release(wb);
            Release(books);
            Release(app);
            sheet = wb = books = app = null;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            ExcelMessageFilter.Revoke();
        }

        var after = SourceIntegrity.Capture(source, options.ExpectedSha256);
        payload ??= new JsonObject { ["ok"] = false };
        payload["ok"] = JsonUtil.Bool(payload, "ok") == true
                        && JsonUtil.Bool(after, "match") == true
                        && string.Equals(JsonUtil.Str(before, "sha256"), JsonUtil.Str(after, "sha256"), StringComparison.OrdinalIgnoreCase);
        payload["hashBefore"] = before;
        payload["hashAfter"] = after;
        payload["sourceUnchanged"] = string.Equals(JsonUtil.Str(before, "sha256"), JsonUtil.Str(after, "sha256"), StringComparison.OrdinalIgnoreCase);
        payload["closedOwnedBook"] = openedFull;
        return payload;
    }

    private static object CreateExcelWithRetry(Type type)
    {
        Exception? last = null;
        for (var i = 0; i < 20; i++)
        {
            try
            {
                return Activator.CreateInstance(type)
                       ?? throw new InvalidOperationException("failed to create Excel.Application");
            }
            catch (COMException ex) when (unchecked((uint)ex.HResult) is 0x8001010A or 0x80010001)
            {
                last = ex;
                Thread.Sleep(500 * (i + 1));
            }
        }
        throw last ?? new InvalidOperationException("Excel.Application create failed");
    }

    private static JsonObject ReadSheet(dynamic sheet, string name)
    {
        var columns = new JsonArray();
        for (var c = 1; c <= 69; c++)
        {
            dynamic col = sheet.Columns[c];
            try
            {
                columns.Add(new JsonObject
                {
                    ["col"] = A1.Col(c),
                    ["columnWidth"] = AsDouble(col.ColumnWidth),
                    ["widthPoints"] = AsDouble(col.Width),
                });
            }
            finally { Release(col); }
        }

        var rows = new JsonArray();
        for (var r = 1; r <= 96; r++)
        {
            dynamic row = sheet.Rows[r];
            try
            {
                rows.Add(new JsonObject
                {
                    ["row"] = r,
                    ["heightPoints"] = AsDouble(row.Height),
                });
            }
            finally { Release(row); }
        }

        var cells = new JsonArray();
        var signatures = new Dictionary<string, int>(StringComparer.Ordinal);
        for (var r = 1; r <= 96; r++)
        {
            for (var c = 1; c <= 69; c++)
            {
                dynamic cell = sheet.Cells[r, c];
                try
                {
                    var font = cell.Font;
                    var alignH = Convert.ToInt32(cell.HorizontalAlignment, CultureInfo.InvariantCulture);
                    var alignV = Convert.ToInt32(cell.VerticalAlignment, CultureInfo.InvariantCulture);
                    var shrink = AsBool(cell.ShrinkToFit);
                    var wrap = AsBool(cell.WrapText);
                    var borders = ReadBorders(cell);
                    var sig = new JsonObject
                    {
                        ["fontName"] = Convert.ToString(font.Name, CultureInfo.InvariantCulture),
                        ["fontSize"] = AsDouble(font.Size),
                        ["bold"] = AsBool(font.Bold),
                        ["horizontalAlign"] = MapH(alignH),
                        ["verticalAlign"] = MapV(alignV),
                        ["shrinkToFit"] = shrink,
                        ["wrapText"] = wrap,
                        ["borders"] = borders,
                    }.ToJsonString(JsonUtil.WriteCompact);
                    signatures[sig] = signatures.TryGetValue(sig, out var n) ? n + 1 : 1;

                    var hasValue = cell.Value2 is not null || !string.IsNullOrEmpty(Convert.ToString(cell.Formula, CultureInfo.InvariantCulture));
                    if (hasValue || signatures[sig] == 1)
                    {
                        cells.Add(new JsonObject
                        {
                            ["a1"] = A1.Cell(r, c),
                            ["fontName"] = Convert.ToString(font.Name, CultureInfo.InvariantCulture),
                            ["fontSize"] = AsDouble(font.Size),
                            ["bold"] = AsBool(font.Bold),
                            ["horizontalAlign"] = MapH(alignH),
                            ["verticalAlign"] = MapV(alignV),
                            ["shrinkToFit"] = shrink,
                            ["wrapText"] = wrap,
                            ["borders"] = borders.DeepClone(),
                        });
                    }
                    Release(font);
                }
                finally { Release(cell); }
            }
        }

        dynamic page = sheet.PageSetup;
        JsonObject print;
        try
        {
            print = new JsonObject
            {
                ["paperSize"] = Convert.ToInt32(page.PaperSize, CultureInfo.InvariantCulture),
                ["orientation"] = Convert.ToInt32(page.Orientation, CultureInfo.InvariantCulture),
                ["zoom"] = TryDouble(() => AsDouble(page.Zoom)),
                ["fitToPagesWide"] = TryDouble(() => AsDouble(page.FitToPagesWide)),
                ["fitToPagesTall"] = TryDouble(() => AsDouble(page.FitToPagesTall)),
                ["printArea"] = Convert.ToString(page.PrintArea, CultureInfo.InvariantCulture),
                ["printTitleRows"] = Convert.ToString(page.PrintTitleRows, CultureInfo.InvariantCulture),
                ["leftMarginPoints"] = AsDouble(page.LeftMargin),
                ["rightMarginPoints"] = AsDouble(page.RightMargin),
                ["topMarginPoints"] = AsDouble(page.TopMargin),
                ["bottomMarginPoints"] = AsDouble(page.BottomMargin),
                ["headerMarginPoints"] = AsDouble(page.HeaderMargin),
                ["footerMarginPoints"] = AsDouble(page.FooterMargin),
                ["centerHeader"] = Convert.ToString(page.CenterHeader, CultureInfo.InvariantCulture),
                ["centerFooter"] = Convert.ToString(page.CenterFooter, CultureInfo.InvariantCulture),
                ["leftFooter"] = Convert.ToString(page.LeftFooter, CultureInfo.InvariantCulture),
                ["rightFooter"] = Convert.ToString(page.RightFooter, CultureInfo.InvariantCulture),
                ["oddFooter"] = Convert.ToString(page.CenterFooter, CultureInfo.InvariantCulture),
            };
        }
        finally { Release(page); }

        dynamic window = sheet.Parent.Windows.Item(1);
        JsonObject freeze;
        try
        {
            freeze = new JsonObject
            {
                ["freezePanes"] = AsBool(window.FreezePanes),
                ["splitColumn"] = Convert.ToInt32(window.SplitColumn, CultureInfo.InvariantCulture),
                ["splitRow"] = Convert.ToInt32(window.SplitRow, CultureInfo.InvariantCulture),
            };
        }
        finally { Release(window); }

        return new JsonObject
        {
            ["ok"] = true,
            ["sheet"] = name,
            ["columnsABQ"] = columns,
            ["rows1to96"] = rows,
            ["uniqueNativeStyleSignatures"] = signatures.Count,
            ["cellSamples"] = cells,
            ["cellSampleCount"] = cells.Count,
            ["print"] = print,
            ["freeze"] = freeze,
            ["note"] = "COM ColumnWidth/Width/Height are native units. They are not claimed equal to OOXML width without tolerance.",
        };
    }

    private static JsonObject ReadBorders(dynamic cell)
    {
        var o = new JsonObject();
        foreach (var (key, idx) in new (string, int)[] { ("left", 7), ("right", 10), ("top", 8), ("bottom", 9) })
        {
            dynamic edge = cell.Borders[idx];
            try
            {
                var weight = Convert.ToInt32(edge.Weight, CultureInfo.InvariantCulture);
                var line = Convert.ToInt32(edge.LineStyle, CultureInfo.InvariantCulture);
                if (line == -4142) continue;
                o[key] = new JsonObject { ["weight"] = MapWeight(weight), ["lineStyle"] = MapLine(line), ["weightRaw"] = weight, ["lineRaw"] = line };
            }
            catch { /* xlLineStyleNone */ }
            finally { Release(edge); }
        }
        return o;
    }

    private static void TryCloseOwned(object? wb, string source)
    {
        if (wb is null) return;
        try
        {
            dynamic d = wb;
            var full = Convert.ToString(d.FullName, CultureInfo.InvariantCulture);
            if (!string.Equals(Path.GetFullPath(full ?? ""), Path.GetFullPath(source), StringComparison.OrdinalIgnoreCase))
                return;
            d.Close(false);
        }
        catch { /* still quit the instance */ }
    }

    private static void TryQuitOwned(object? app)
    {
        if (app is null) return;
        object? books = null;
        try
        {
            dynamic d = app;
            var caption = Convert.ToString(d.Caption, CultureInfo.InvariantCulture) ?? "";
            if (!caption.Contains(InstanceCaption, StringComparison.Ordinal))
                return;
            books = d.Workbooks;
            var remaining = Convert.ToInt32(((dynamic)books).Count, CultureInfo.InvariantCulture);
            if (remaining != 0)
                return;
            d.DisplayAlerts = false;
            d.Quit();
        }
        catch { /* process may already be gone */ }
        finally { Release(books); }
    }

    private static void Release(object? com)
    {
        if (com is null || !Marshal.IsComObject(com)) return;
        try { Marshal.ReleaseComObject(com); } catch { /* ignore */ }
    }

    private static double? AsDouble(object? v)
    {
        if (v is null || v is DBNull) return null;
        try { return Convert.ToDouble(v, CultureInfo.InvariantCulture); }
        catch { return null; }
    }

    private static bool? AsBool(object? v)
    {
        if (v is null || v is DBNull) return null;
        if (v is bool b) return b;
        try { return Convert.ToBoolean(v, CultureInfo.InvariantCulture); }
        catch { return null; }
    }

    private static double? TryDouble(Func<double?> read)
    {
        try { return read(); }
        catch { return null; }
    }

    private static string MapH(int v) => v switch
    {
        -4131 => "left",
        -4108 => "center",
        -4152 => "right",
        1 => "general",
        5 => "distributed",
        7 => "centerAcross",
        _ => v.ToString(CultureInfo.InvariantCulture),
    };

    private static string MapV(int v) => v switch
    {
        -4160 => "top",
        -4108 => "center",
        -4107 => "bottom",
        _ => v.ToString(CultureInfo.InvariantCulture),
    };

    private static string MapWeight(int v) => v switch
    {
        1 => "hairline",
        2 => "thin",
        -4138 => "medium",
        4 => "thick",
        _ => v.ToString(CultureInfo.InvariantCulture),
    };

    private static string MapLine(int v) => v switch
    {
        1 => "continuous",
        -4118 => "dotted",
        -4115 => "dashed",
        -4119 => "double",
        _ => v.ToString(CultureInfo.InvariantCulture),
    };
}
