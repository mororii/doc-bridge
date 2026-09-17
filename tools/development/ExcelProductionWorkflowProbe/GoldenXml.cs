using System.IO.Compression;
using System.Xml.Linq;

namespace DocBridge.Development.ExcelProductionWorkflowProbe;

/// <summary>
/// Read-only OOXML inspection of an existing .xlsx. Never writes the package.
/// Column widths are the OOXML values; they are not claimed equal to COM ColumnWidth.
/// </summary>
internal static class GoldenXml
{
    private static readonly XNamespace Ss = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static readonly XNamespace Rel = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace PkgRel = "http://schemas.openxmlformats.org/package/2006/relationships";

    public static JsonObject Inspect(string xlsxPath, string requestedSheet)
    {
        using var fs = new FileStream(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Read, leaveOpen: false);

        var workbook = ReadXml(zip, "xl/workbook.xml");
        var rels = ReadXml(zip, "xl/_rels/workbook.xml.rels");
        var styles = ReadXml(zip, "xl/styles.xml");
        var sheetPath = ResolveSheetPath(workbook, rels, requestedSheet);
        var sheet = ReadXml(zip, sheetPath);

        var merges = sheet.Descendants(Ss + "mergeCell")
            .Select(e => (string?)e.Attribute("ref"))
            .Where(r => !string.IsNullOrWhiteSpace(r))
            .Select(r => r!)
            .ToList();

        var month = merges.Where(A1.IsMonthHeaderPair).OrderBy(m => A1.ParseRange(m).C1).ToList();
        var remaining = merges.Where(m => !A1.IsMonthHeaderPair(m)).OrderBy(m => m, StringComparer.Ordinal).ToList();

        var pane = sheet.Descendants(Ss + "pane").FirstOrDefault();
        var page = sheet.Element(Ss + "pageSetup") ?? sheet.Descendants(Ss + "pageSetup").FirstOrDefault();
        var margins = sheet.Element(Ss + "pageMargins") ?? sheet.Descendants(Ss + "pageMargins").FirstOrDefault();
        var footer = sheet.Descendants(Ss + "oddFooter").FirstOrDefault()?.Value
                     ?? sheet.Descendants(Ss + "evenFooter").FirstOrDefault()?.Value;
        var format = sheet.Element(Ss + "sheetFormatPr");
        var dimension = (string?)sheet.Element(Ss + "dimension")?.Attribute("ref");

        var fonts = RootList(styles, "fonts", "font").Select(DescribeFont).ToList();
        var borders = RootList(styles, "borders", "border").Select(DescribeBorder).ToList();
        var cellXfs = RootList(styles, "cellXfs", "xf");

        var printNames = workbook.Descendants(Ss + "definedName")
            .Select(n => new JsonObject
            {
                ["name"] = (string?)n.Attribute("name"),
                ["localSheetId"] = (string?)n.Attribute("localSheetId"),
                ["refersTo"] = n.Value,
            })
            .Cast<JsonNode>()
            .ToList();

        return new JsonObject
        {
            ["sheetPath"] = sheetPath,
            ["requestedSheet"] = requestedSheet,
            ["resolvedSheetName"] = requestedSheet,
            ["dimension"] = dimension,
            ["mergeCount"] = merges.Count,
            ["merges"] = new JsonArray(merges.Select(m => JsonValue.Create(m)).ToArray()),
            ["monthHeaderMerges"] = new JsonArray(month.Select(m => JsonValue.Create(m)).ToArray()),
            ["monthHeaderCount"] = month.Count,
            ["remainingMergeCount"] = remaining.Count,
            ["remainingMerges"] = new JsonArray(remaining.Select(m => JsonValue.Create(m)).ToArray()),
            ["freezePanes"] = pane is null ? null : new JsonObject
            {
                ["xSplit"] = (string?)pane.Attribute("xSplit"),
                ["ySplit"] = (string?)pane.Attribute("ySplit"),
                ["topLeftCell"] = (string?)pane.Attribute("topLeftCell"),
                ["state"] = (string?)pane.Attribute("state"),
            },
            ["pageSetup"] = page is null ? null : new JsonObject
            {
                ["paperSize"] = (string?)page.Attribute("paperSize"),
                ["orientation"] = (string?)page.Attribute("orientation"),
                ["scale"] = (string?)page.Attribute("scale"),
                ["fitToWidth"] = (string?)page.Attribute("fitToWidth"),
                ["fitToHeight"] = (string?)page.Attribute("fitToHeight"),
            },
            ["pageMarginsInches"] = pageMargins(margins),
            ["pageMarginsMmApprox"] = pageMarginsMm(margins),
            ["centerFooter"] = footer,
            ["defaultRowHeightPoints"] = (string?)format?.Attribute("defaultRowHeight"),
            ["columns"] = new JsonArray(sheet.Descendants(Ss + "col").Select(c => new JsonObject
            {
                ["min"] = (string?)c.Attribute("min"),
                ["max"] = (string?)c.Attribute("max"),
                ["widthXml"] = (string?)c.Attribute("width"),
                ["customWidth"] = (string?)c.Attribute("customWidth"),
                ["note"] = "OOXML width; do not treat as COM ColumnWidth",
            }).ToArray()),
            ["fonts"] = new JsonArray(fonts.ToArray()),
            ["borderStyles"] = new JsonArray(borders.ToArray()),
            ["cellXfCount"] = cellXfs.Count,
            ["usedStyleHint"] = "cellXfs count is the style table; used-in-cells is counted separately",
            ["definedNames"] = new JsonArray(printNames.ToArray()),
            ["row5Unmerged"] = !merges.Any(m => A1.Overlaps(m, "G5:BP5")),
            ["row5IntersectingMerges"] = new JsonArray(merges.Where(m => A1.Overlaps(m, "G5:BP5")).Select(m => JsonValue.Create(m)).ToArray()),
        };
    }

    public static JsonObject CountUsedStylesAndMediumBottoms(string xlsxPath, string requestedSheet, IReadOnlyCollection<(int R, int C)>? barCells)
    {
        using var fs = new FileStream(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Read, leaveOpen: false);
        var workbook = ReadXml(zip, "xl/workbook.xml");
        var rels = ReadXml(zip, "xl/_rels/workbook.xml.rels");
        var styles = ReadXml(zip, "xl/styles.xml");
        var sheet = ReadXml(zip, ResolveSheetPath(workbook, rels, requestedSheet));

        var borders = RootList(styles, "borders", "border");
        var mediumBottomIds = new HashSet<int>();
        for (var i = 0; i < borders.Count; i++)
        {
            var bottom = borders[i].Element(Ss + "bottom");
            if (string.Equals((string?)bottom?.Attribute("style"), "medium", StringComparison.OrdinalIgnoreCase))
                mediumBottomIds.Add(i);
        }

        var xfs = RootList(styles, "cellXfs", "xf");
        var used = new HashSet<int>();
        var scheduleBars = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var otherMedium = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var cell in sheet.Descendants(Ss + "c"))
        {
            var addr = (string?)cell.Attribute("r");
            var s = (string?)cell.Attribute("s");
            if (!int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var styleId))
                continue;
            used.Add(styleId);
            if (addr is null || styleId >= xfs.Count) continue;
            var borderId = (string?)xfs[styleId].Attribute("borderId");
            if (!int.TryParse(borderId, NumberStyles.Integer, CultureInfo.InvariantCulture, out var bid) ||
                !mediumBottomIds.Contains(bid))
                continue;
            var parsed = A1.ParseCell(addr);
            var inScheduleGrid = parsed.Row is >= 1 and <= 96 && parsed.Col is >= 7 and <= 68;
            if (inScheduleGrid) scheduleBars.Add(addr);
            else otherMedium.Add(addr);
        }

        JsonObject? barCompare = null;
        if (barCells is not null)
        {
            var rebuild = barCells
                .Where(b => b.R is >= 1 and <= 96 && b.C is >= 7 and <= 68)
                .Select(b => A1.Cell(b.R, b.C))
                .ToHashSet(StringComparer.OrdinalIgnoreCase);
            barCompare = new JsonObject
            {
                ["rebuildCount"] = rebuild.Count,
                ["xmlScheduleMediumBottomCount"] = scheduleBars.Count,
                ["onlyInRebuild"] = new JsonArray(rebuild.Except(scheduleBars, StringComparer.OrdinalIgnoreCase).Select(s => JsonValue.Create(s)).ToArray()),
                ["onlyInXmlScheduleGrid"] = new JsonArray(scheduleBars.Except(rebuild, StringComparer.OrdinalIgnoreCase).Select(s => JsonValue.Create(s)).ToArray()),
                ["otherMediumBottomsOutsideG1BP96"] = new JsonArray(otherMedium.Select(s => JsonValue.Create(s)).ToArray()),
                ["note"] = "Schedule bars are G1:BP96 cell-bottom medium. Left-of-G or row>96 mediums are table/boundary, not bars.",
            };
        }

        return new JsonObject
        {
            ["usedCellStyleCount"] = used.Count,
            ["mediumBottomBorderStyleIds"] = new JsonArray(mediumBottomIds.Select(i => JsonValue.Create(i)).ToArray()),
            ["xmlScheduleMediumBottomCellCount"] = scheduleBars.Count,
            ["barCompare"] = barCompare,
        };
    }

    private static JsonObject? pageMargins(XElement? m)
    {
        if (m is null) return null;
        return new JsonObject
        {
            ["left"] = (string?)m.Attribute("left"),
            ["right"] = (string?)m.Attribute("right"),
            ["top"] = (string?)m.Attribute("top"),
            ["bottom"] = (string?)m.Attribute("bottom"),
            ["header"] = (string?)m.Attribute("header"),
            ["footer"] = (string?)m.Attribute("footer"),
        };
    }

    private static JsonObject? pageMarginsMm(XElement? m)
    {
        if (m is null) return null;
        double Mm(string? inches) =>
            double.TryParse(inches, NumberStyles.Float, CultureInfo.InvariantCulture, out var v)
                ? Math.Round(v * 25.4, 3)
                : double.NaN;
        return new JsonObject
        {
            ["left"] = Mm((string?)m.Attribute("left")),
            ["right"] = Mm((string?)m.Attribute("right")),
            ["top"] = Mm((string?)m.Attribute("top")),
            ["bottom"] = Mm((string?)m.Attribute("bottom")),
            ["header"] = Mm((string?)m.Attribute("header")),
            ["footer"] = Mm((string?)m.Attribute("footer")),
            ["note"] = "inches * 25.4; compare with tolerance, not bit-exact COM",
        };
    }

    private static JsonObject DescribeFont(XElement font) => new()
    {
        ["name"] = (string?)font.Element(Ss + "name")?.Attribute("val"),
        ["sz"] = (string?)font.Element(Ss + "sz")?.Attribute("val"),
        ["bold"] = font.Element(Ss + "b") is not null,
        ["italic"] = font.Element(Ss + "i") is not null,
    };

    private static JsonObject DescribeBorder(XElement border)
    {
        JsonObject Edge(string name)
        {
            var e = border.Element(Ss + name);
            return new JsonObject
            {
                ["style"] = (string?)e?.Attribute("style"),
            };
        }
        return new JsonObject
        {
            ["left"] = Edge("left"),
            ["right"] = Edge("right"),
            ["top"] = Edge("top"),
            ["bottom"] = Edge("bottom"),
        };
    }

    private static IReadOnlyList<XElement> RootList(XDocument styles, string parent, string child)
    {
        var root = styles.Root ?? throw new InvalidOperationException("styles.xml has no root");
        return root.Element(Ss + parent)?.Elements(Ss + child).ToList() ?? [];
    }

    internal static string ResolveSheetPathPublic(XDocument workbook, XDocument rels, string requestedSheet) =>
        ResolveSheetPath(workbook, rels, requestedSheet);

    private static string ResolveSheetPath(XDocument workbook, XDocument rels, string requestedSheet)
    {
        var sheets = workbook.Descendants(Ss + "sheet").ToList();
        var sheet = sheets.FirstOrDefault(s =>
                       string.Equals((string?)s.Attribute("name"), requestedSheet, StringComparison.OrdinalIgnoreCase))
                   ?? throw new InvalidOperationException($"workbook has no sheet named {requestedSheet}");
        var rid = (string?)sheet.Attribute(Rel + "id") ?? throw new InvalidOperationException("sheet r:id missing");
        var target = rels.Descendants(PkgRel + "Relationship")
            .FirstOrDefault(r => (string?)r.Attribute("Id") == rid)
            ?.Attribute("Target")?.Value
            ?? throw new InvalidOperationException($"missing relationship {rid}");
        return NormalizePartPath(target);
    }

    /// <summary>
    /// workbook.xml.rels targets are relative to xl/ or package-absolute (/xl/...).
    /// Never prefix xl/ onto an already-absolute /xl/... path.
    /// </summary>
    internal static string NormalizePartPath(string target)
    {
        var t = target.Replace('\\', '/');
        if (t.StartsWith('/'))
            return t.TrimStart('/');
        if (t.StartsWith("xl/", StringComparison.OrdinalIgnoreCase))
            return t;
        return "xl/" + t.TrimStart('/');
    }

    private static XDocument ReadXml(ZipArchive zip, string path)
    {
        var entry = zip.GetEntry(path) ?? throw new InvalidOperationException($"package part missing: {path}");
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }
}
