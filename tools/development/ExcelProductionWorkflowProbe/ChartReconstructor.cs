using System.IO.Compression;
using System.Xml.Linq;

namespace DocBridge.Development.ExcelProductionWorkflowProbe;

/// <summary>
/// Read-only Chart1 reconstruction from the source package + native column/row sizes.
/// Not a raster and not a copy of the formatted workbook.
/// </summary>
internal static class ChartReconstructor
{
    private static readonly XNamespace Xdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
    private static readonly XNamespace C = "http://schemas.openxmlformats.org/drawingml/2006/chart";
    private static readonly XNamespace A = "http://schemas.openxmlformats.org/drawingml/2006/main";
    private const double EmuPerPoint = 12700;

    public static JsonObject Reconstruct(string sourceXlsx, string? comOraclePath)
    {
        using var fs = new FileStream(sourceXlsx, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Read, leaveOpen: false);
        var drawing = Read(zip, "xl/drawings/drawing1.xml");
        var chart = Read(zip, "xl/charts/chart1.xml");

        var from = drawing.Descendants(Xdr + "from").FirstOrDefault()
                   ?? throw new InvalidOperationException("drawing1.xml missing twoCellAnchor from");
        var to = drawing.Descendants(Xdr + "to").FirstOrDefault()
                 ?? throw new InvalidOperationException("drawing1.xml missing twoCellAnchor to");
        var name = (string?)drawing.Descendants(Xdr + "cNvPr").FirstOrDefault()?.Attribute("name") ?? "Chart 1";

        var formula = chart.Descendants(C + "f").Select(e => e.Value).FirstOrDefault(s => s.Contains('$'))
                      ?? "예정공정표!$G$95:$BP$95";
        var rgb = (string?)chart.Descendants(A + "srgbClr").FirstOrDefault()?.Attribute("val") ?? "FF0000";
        var lineEmu = (string?)chart.Descendants(A + "ln").FirstOrDefault()?.Attribute("w");
        var linePt = double.TryParse(lineEmu, NumberStyles.Float, CultureInfo.InvariantCulture, out var emu)
            ? Math.Round(emu / EmuPerPoint, 2)
            : 2.25;
        var marker = (string?)chart.Descendants(C + "marker").FirstOrDefault()?.Element(C + "symbol")?.Attribute("val") ?? "none";
        var valMax = (string?)chart.Descendants(C + "valAx").FirstOrDefault()?.Descendants(C + "max").FirstOrDefault()?.Attribute("val");
        var catDeleted = string.Equals((string?)chart.Descendants(C + "catAx").FirstOrDefault()?.Element(C + "delete")?.Attribute("val"), "1", StringComparison.Ordinal);
        var valDeleted = string.Equals((string?)chart.Descendants(C + "valAx").FirstOrDefault()?.Element(C + "delete")?.Attribute("val"), "1", StringComparison.Ordinal);
        var sourceRange = formula.Contains('!', StringComparison.Ordinal)
            ? formula[(formula.IndexOf('!') + 1)..].Replace("$", "", StringComparison.Ordinal)
            : formula.Replace("$", "", StringComparison.Ordinal);

        var position = PositionFromAnchor(from, to, comOraclePath);
        var create = new JsonObject
        {
            ["op"] = "create_chart",
            ["target"] = JsonUtil.Target(ScenarioE1.Sheet),
            ["name"] = name.Replace(" ", "", StringComparison.Ordinal),
            ["chartType"] = "line",
            ["sourceRange"] = "'예정공정표'!$G$95:$BP$95",
            ["title"] = "",
            ["hasLegend"] = false,
            ["plotBy"] = "rows",
            ["position"] = position,
            ["series"] = JsonUtil.Arr(new JsonObject
            {
                ["lineColor"] = "#" + rgb.TrimStart('#'),
                ["lineWeight"] = linePt,
                ["marker"] = marker,
            }),
            ["chartFill"] = "none",
            ["plotFill"] = "none",
            ["chartBorder"] = "none",
            ["plotBorder"] = "none",
            ["axes"] = new JsonObject
            {
                ["category"] = new JsonObject { ["visible"] = !catDeleted ? true : false },
                ["value"] = new JsonObject
                {
                    ["visible"] = !valDeleted ? true : false,
                    ["maximum"] = double.TryParse(valMax, NumberStyles.Float, CultureInfo.InvariantCulture, out var mx) ? mx : 100,
                },
            },
            ["plotArea"] = new JsonObject { ["spanChart"] = true },
        };

        return new JsonObject
        {
            ["sourceFormula"] = formula,
            ["sourceRange"] = sourceRange,
            ["drawingName"] = name,
            ["anchor"] = new JsonObject
            {
                ["fromCol0"] = Int(from, "col"),
                ["fromColOffEmu"] = Int(from, "colOff"),
                ["fromRow0"] = Int(from, "row"),
                ["fromRowOffEmu"] = Int(from, "rowOff"),
                ["toCol0"] = Int(to, "col"),
                ["toColOffEmu"] = Int(to, "colOff"),
                ["toRow0"] = Int(to, "row"),
                ["toRowOffEmu"] = Int(to, "rowOff"),
            },
            ["createOp"] = create,
            ["notRaster"] = true,
        };
    }

    private static JsonObject PositionFromAnchor(XElement from, XElement to, string? comOraclePath)
    {
        var widths = new double[70];
        var heights = new double[97];
        Array.Fill(widths, 24.6);
        Array.Fill(heights, 15);
        if (!string.IsNullOrWhiteSpace(comOraclePath) && File.Exists(comOraclePath))
        {
            var com = JsonUtil.Load(comOraclePath).AsObject();
            if (com["columnsABQ"] is JsonArray cols)
            {
                foreach (var col in cols.OfType<JsonObject>())
                {
                    var letter = JsonUtil.Str(col, "col");
                    if (letter is null) continue;
                    var idx = A1.ColIndex(letter);
                    if (idx is >= 1 and <= 69)
                        widths[idx] = JsonUtil.Num(col, "widthPoints") ?? widths[idx];
                }
            }
            if (com["rows1to96"] is JsonArray rows)
            {
                foreach (var row in rows.OfType<JsonObject>())
                {
                    var r = (int)(JsonUtil.Num(row, "row") ?? 0);
                    if (r is >= 1 and <= 96)
                        heights[r] = JsonUtil.Num(row, "heightPoints") ?? heights[r];
                }
            }
        }

        var left = Sum(widths, 1, Int(from, "col")) + Int(from, "colOff") / EmuPerPoint;
        var top = Sum(heights, 1, Int(from, "row")) + Int(from, "rowOff") / EmuPerPoint;
        var right = Sum(widths, 1, Int(to, "col")) + Int(to, "colOff") / EmuPerPoint;
        var bottom = Sum(heights, 1, Int(to, "row")) + Int(to, "rowOff") / EmuPerPoint;
        return new JsonObject
        {
            ["left"] = Math.Round(left, 2),
            ["top"] = Math.Round(top, 2),
            ["width"] = Math.Round(Math.Max(1, right - left), 2),
            ["height"] = Math.Round(Math.Max(1, bottom - top), 2),
        };
    }

    private static double Sum(double[] values, int from1, int countExclusive)
    {
        var t = 0d;
        for (var i = from1; i <= countExclusive && i < values.Length; i++)
            t += values[i];
        return t;
    }

    private static int Int(XElement parent, string local)
    {
        var text = parent.Elements().FirstOrDefault(e => e.Name.LocalName == local)?.Value;
        return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : 0;
    }

    private static XDocument Read(ZipArchive zip, string path)
    {
        var entry = zip.GetEntry(path) ?? throw new InvalidOperationException("package part missing: " + path);
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }
}
