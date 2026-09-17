using System.IO.Compression;
using System.Xml.Linq;

namespace DocBridge.Development.ExcelProductionWorkflowProbe;

/// <summary>
/// Read-only OOXML proof that a freshly saved owned E1 workbook is actually
/// blank: no values, no merges, and no effective cell/row/column fill on A1:BQ96.
/// XF0 applies to omitted cells and cells without <c>s</c>. applyFill=0 inherits
/// <c>cellStyleXfs[xfId]</c>. Unused gray125 in the palette is allowed.
/// Planner <c>blankSheetNoFillAssumed</c> is not this proof.
/// </summary>
internal static class BlankCheckpoint
{
    public const int TargetRows = 96;
    public const int TargetCols = 69;
    public const string BatchId = "e1-blank-checkpoint";

    private static readonly XNamespace Ss = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    public static JsonObject VerifyOwned(string xlsxPath, string sheetName)
    {
        if (string.IsNullOrWhiteSpace(xlsxPath) || !File.Exists(xlsxPath))
        {
            return Fail("owned checkpoint file is missing; cannot omit noFill",
                new JsonObject { ["path"] = xlsxPath });
        }

        using var fs = new FileStream(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        return VerifyStream(fs, sheetName, xlsxPath);
    }

    public static JsonObject VerifyStream(Stream stream, string sheetName, string? path = null)
    {
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read, leaveOpen: true);
        var workbook = ReadXml(zip, "xl/workbook.xml");
        var rels = ReadXml(zip, "xl/_rels/workbook.xml.rels");
        var styles = ReadXml(zip, "xl/styles.xml");
        var sheetPath = GoldenXml.ResolveSheetPathPublic(workbook, rels, sheetName);
        var sheet = ReadXml(zip, sheetPath);
        XDocument? sst = null;
        var sstEntry = zip.GetEntry("xl/sharedStrings.xml");
        if (sstEntry is not null)
        {
            using var sstStream = sstEntry.Open();
            sst = XDocument.Load(sstStream);
        }

        return Evaluate(sheet, styles, sst, path, sheetPath);
    }

    public static JsonObject SelfCheck()
    {
        var errors = new JsonArray();
        var blank = VerifyBytes(MinimalPackage(BlankSheet()), ScenarioE1.Sheet);
        if (JsonUtil.Bool(blank, "ok") != true)
            errors.Add("blank package must pass: " + blank);
        if (JsonUtil.Bool(VerifyBytes(MinimalPackage(ValueSheet()), ScenarioE1.Sheet), "ok") == true)
            errors.Add("package with A1 value must fail");
        if (JsonUtil.Bool(VerifyBytes(MinimalPackage(MergeSheet()), ScenarioE1.Sheet), "ok") == true)
            errors.Add("package with a merge must fail");
        if (JsonUtil.Bool(VerifyBytes(MinimalPackage(FilledCellSheet()), ScenarioE1.Sheet), "ok") == true)
            errors.Add("package with cell fill on A1 must fail");
        if (JsonUtil.Bool(VerifyBytes(MinimalPackage(FilledRowSheet()), ScenarioE1.Sheet), "ok") == true)
            errors.Add("package with row fill on row 1 must fail");
        if (JsonUtil.Bool(VerifyBytes(MinimalPackage(FilledColSheet()), ScenarioE1.Sheet), "ok") == true)
            errors.Add("package with column fill on A must fail");
        var outside = VerifyBytes(MinimalPackage(FilledColOutsideSheet()), ScenarioE1.Sheet);
        if (JsonUtil.Bool(outside, "ok") != true)
            errors.Add("fill on column BR (outside A:BQ) must not fail the target grid: " + outside);

        var xf0Empty = VerifyBytes(MinimalPackage(BlankSheet(), Xf0FilledStyles()), ScenarioE1.Sheet);
        if (JsonUtil.Bool(xf0Empty, "ok") == true)
            errors.Add("blank-sheet default-fill: empty sheet with filled default XF0 must fail");
        var parentBlank = VerifyBytes(MinimalPackage(BlankSheet(), Xf0InheritsFilledParentStyles()), ScenarioE1.Sheet);
        if (JsonUtil.Bool(parentBlank, "ok") == true)
            errors.Add("blank-sheet parent-fill: XF0 applyFill=0 inheriting a filled cellStyleXfs parent must fail");
        var noStyleXf0 = VerifyBytes(MinimalPackage(CellWithoutStyleSheet(), Xf0FilledStyles()), ScenarioE1.Sheet);
        if (JsonUtil.Bool(noStyleXf0, "ok") == true)
            errors.Add("cell without s still uses filled XF0 and must fail");
        var inherited = VerifyBytes(MinimalPackage(FilledCellSheet(), InheritedFillStyles()), ScenarioE1.Sheet);
        if (JsonUtil.Bool(inherited, "ok") == true)
            errors.Add("parent-fill: applyFill=0 must inherit a filled cellStyleXfs[xfId] and fail");
        var unusedGray = VerifyBytes(MinimalPackage(BlankSheet(), UnusedGray125Styles()), ScenarioE1.Sheet);
        if (JsonUtil.Bool(unusedGray, "ok") != true)
            errors.Add("unused gray125 in the palette must pass when no effective target xf uses it: " + unusedGray);
        var inheritNone = VerifyBytes(MinimalPackage(BlankSheet(), ApplyFillZeroInheritsNoneStyles()), ScenarioE1.Sheet);
        if (JsonUtil.Bool(inheritNone, "ok") != true)
            errors.Add("applyFill=0 inheriting none from cellStyleXfs must pass: " + inheritNone);
        var implicitNone = VerifyBytes(MinimalPackage(CellWithoutStyleSheet()), ScenarioE1.Sheet);
        if (JsonUtil.Bool(implicitNone, "ok") != true)
            errors.Add("cell without s and XF0 none must pass: " + implicitNone);

        return new JsonObject
        {
            ["ok"] = errors.Count == 0,
            ["errors"] = errors,
        };
    }

    private static JsonObject Evaluate(XDocument sheet, XDocument styles, XDocument? sst, string? path, string sheetPath)
    {
        var errors = new JsonArray();
        var fillIds = EffectiveFilledXfIds(styles);
        var xf0Filled = fillIds.Contains(0);
        var nonempty = new JsonArray();
        var filledCells = new JsonArray();
        var filledRows = new JsonArray();
        var filledCols = new JsonArray();

        if (xf0Filled)
            errors.Add("default XF0 effective fill applies to omitted cells and cells without s");

        foreach (var cell in sheet.Descendants(Ss + "c"))
        {
            var addr = (string?)cell.Attribute("r") ?? "";
            if (HasCellValue(cell, sst))
            {
                nonempty.Add(addr);
                if (nonempty.Count <= 12)
                    errors.Add($"nonempty cell {addr}");
            }

            var styleId = ParseInt(cell.Attribute("s")) ?? 0;
            if (fillIds.Contains(styleId) && InTarget(addr))
            {
                if (styleId == 0 && xf0Filled)
                    continue;
                filledCells.Add(addr);
                if (filledCells.Count <= 12)
                    errors.Add($"cell fill on target {addr} xf={styleId}");
            }
        }

        if (nonempty.Count > 12)
            errors.Add($"and {nonempty.Count - 12} more nonempty cells");

        var merges = sheet.Descendants(Ss + "mergeCell")
            .Select(e => (string?)e.Attribute("ref"))
            .Where(r => !string.IsNullOrWhiteSpace(r))
            .Select(r => r!)
            .ToList();
        foreach (var merge in merges.Take(8))
            errors.Add("merge " + merge);
        if (merges.Count > 8)
            errors.Add($"and {merges.Count - 8} more merges");

        foreach (var row in sheet.Descendants(Ss + "row"))
        {
            var r = ParseInt(row.Attribute("r"));
            var styleId = ParseInt(row.Attribute("s"));
            var custom = (string?)row.Attribute("customFormat") == "1";
            if (r is int rowNum && rowNum is >= 1 and <= TargetRows && styleId is int sid && fillIds.Contains(sid) && custom)
            {
                filledRows.Add(rowNum);
                errors.Add($"row fill on target row {rowNum} xf={sid}");
            }
        }

        foreach (var col in sheet.Descendants(Ss + "col"))
        {
            var min = ParseInt(col.Attribute("min")) ?? 0;
            var max = ParseInt(col.Attribute("max")) ?? min;
            var styleId = ParseInt(col.Attribute("style"));
            if (styleId is int sid && fillIds.Contains(sid) && RangesOverlap(min, max, 1, TargetCols))
            {
                filledCols.Add($"{min}:{max}");
                errors.Add($"column fill on target cols {min}:{max} xf={sid}");
            }
        }

        var ok = errors.Count == 0;
        return new JsonObject
        {
            ["ok"] = ok,
            ["path"] = path,
            ["sheetPath"] = sheetPath,
            ["target"] = "A1:BQ96",
            ["nonemptyCellCount"] = nonempty.Count,
            ["mergeCount"] = merges.Count,
            ["filledCellCount"] = filledCells.Count,
            ["filledRowCount"] = filledRows.Count,
            ["filledColCount"] = filledCols.Count,
            ["filledStyleXfIds"] = new JsonArray(fillIds.OrderBy(i => i).Select(i => JsonValue.Create(i)).ToArray()),
            ["xf0EffectiveFill"] = xf0Filled,
            ["applyFillZeroInheritsCellStyleXfs"] = true,
            ["unusedGray125Allowed"] = true,
            ["errors"] = errors,
            ["plannerFlagIsNotProof"] = "blankSheetNoFillAssumed is not a checkpoint",
            ["note"] = ok
                ? "owned blank OOXML proved; noFill may stay omitted on this checkpoint only"
                : "blank/noFill baseline failed; refuse values/styles; do not omit noFill",
        };
    }

    private static bool HasCellValue(XElement cell, XDocument? sst)
    {
        if (cell.Element(Ss + "f") is { } f && !string.IsNullOrWhiteSpace(f.Value))
            return true;
        if (cell.Element(Ss + "is") is { } inline)
            return inline.Descendants(Ss + "t").Any(t => t.Value.Length > 0);
        var v = cell.Element(Ss + "v");
        if (v is null || string.IsNullOrEmpty(v.Value))
            return false;
        var type = (string?)cell.Attribute("t");
        if (string.Equals(type, "s", StringComparison.OrdinalIgnoreCase) && sst is not null
            && int.TryParse(v.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var idx))
        {
            var item = sst.Descendants(Ss + "si").Skip(idx).FirstOrDefault();
            return item is not null && item.Descendants(Ss + "t").Any(t => t.Value.Length > 0);
        }
        return true;
    }

    private static HashSet<int> EffectiveFilledXfIds(XDocument styles)
    {
        var fills = RootList(styles, "fills", "fill");
        var filledFill = new HashSet<int>();
        for (var i = 0; i < fills.Count; i++)
        {
            if (IsFilled(fills[i]))
                filledFill.Add(i);
        }

        var styleXfs = RootList(styles, "cellStyleXfs", "xf");
        var xfs = RootList(styles, "cellXfs", "xf");
        var ids = new HashSet<int>();
        for (var i = 0; i < xfs.Count; i++)
        {
            var fillXf = AppliedFillXf(xfs[i], styleXfs);
            var fillId = ParseInt(fillXf.Attribute("fillId")) ?? 0;
            if (filledFill.Contains(fillId))
                ids.Add(i);
        }
        return ids;
    }

    private static XElement AppliedFillXf(XElement xf, IReadOnlyList<XElement> styleXfs)
    {
        var apply = (string?)xf.Attribute("applyFill");
        if (apply is "1" or "true") return xf;
        if (apply is "0" or "false")
        {
            var id = ParseInt(xf.Attribute("xfId")) ?? 0;
            if (id >= 0 && id < styleXfs.Count) return styleXfs[id];
        }
        return xf;
    }

    private static bool IsFilled(XElement fill)
    {
        var pattern = fill.Element(Ss + "patternFill") ?? fill.Descendants(Ss + "patternFill").FirstOrDefault();
        if (pattern is not null)
        {
            var type = (string?)pattern.Attribute("patternType");
            if (string.IsNullOrWhiteSpace(type) || string.Equals(type, "none", StringComparison.OrdinalIgnoreCase))
                return false;
            return true;
        }
        return fill.Element(Ss + "gradientFill") is not null
               || fill.Descendants(Ss + "gradientFill").Any();
    }

    private static bool InTarget(string address)
    {
        if (string.IsNullOrWhiteSpace(address)) return false;
        try
        {
            var parsed = A1.ParseCell(address);
            return parsed.Row is >= 1 and <= TargetRows && parsed.Col is >= 1 and <= TargetCols;
        }
        catch
        {
            return false;
        }
    }

    private static bool RangesOverlap(int a1, int a2, int b1, int b2) => a1 <= b2 && b1 <= a2;

    private static int? ParseInt(XAttribute? attr)
    {
        if (attr is null) return null;
        return int.TryParse(attr.Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out var n) ? n : null;
    }

    private static IReadOnlyList<XElement> RootList(XDocument styles, string parent, string child) =>
        styles.Root?.Element(Ss + parent)?.Elements(Ss + child).ToList() ?? [];

    private static XDocument ReadXml(ZipArchive zip, string path)
    {
        var entry = zip.GetEntry(path) ?? throw new InvalidOperationException("package part missing: " + path);
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    private static JsonObject Fail(string error, JsonObject extra)
    {
        extra["ok"] = false;
        extra["errors"] = JsonUtil.Arr(JsonValue.Create(error));
        return extra;
    }

    private static JsonObject VerifyBytes(byte[] bytes, string sheet)
    {
        using var ms = new MemoryStream(bytes);
        return VerifyStream(ms, sheet, "self-check");
    }

    private static byte[] MinimalPackage(string sheetXml, string? stylesXml = null)
    {
        using var ms = new MemoryStream();
        using (var zip = new ZipArchive(ms, ZipArchiveMode.Create, leaveOpen: true))
        {
            Write(zip, "[Content_Types].xml",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
                  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
                  <Default Extension="xml" ContentType="application/xml"/>
                  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
                  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
                  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
                </Types>
                """);
            Write(zip, "_rels/.rels",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
                </Relationships>
                """);
            Write(zip, "xl/workbook.xml",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"
                          xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
                  <sheets><sheet name="예정공정표" sheetId="1" r:id="rId1"/></sheets>
                </workbook>
                """);
            Write(zip, "xl/_rels/workbook.xml.rels",
                """
                <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
                <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
                  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
                  <Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
                </Relationships>
                """);
            Write(zip, "xl/styles.xml", stylesXml ?? DefaultStyles());
            Write(zip, "xl/worksheets/sheet1.xml", sheetXml);
        }
        return ms.ToArray();
    }

    private static string DefaultStyles() =>
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fills count="2">
            <fill><patternFill patternType="none"/></fill>
            <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
          </fills>
          <cellXfs count="2">
            <xf fillId="0"/>
            <xf fillId="1" applyFill="1"/>
          </cellXfs>
        </styleSheet>
        """;

    private static string Xf0FilledStyles() =>
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fills count="2">
            <fill><patternFill patternType="none"/></fill>
            <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
          </fills>
          <cellStyleXfs count="1"><xf fillId="0"/></cellStyleXfs>
          <cellXfs count="1"><xf fillId="1" applyFill="1" xfId="0"/></cellXfs>
        </styleSheet>
        """;

    private static string Xf0InheritsFilledParentStyles() =>
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fills count="2">
            <fill><patternFill patternType="none"/></fill>
            <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
          </fills>
          <cellStyleXfs count="1"><xf fillId="1" applyFill="1"/></cellStyleXfs>
          <cellXfs count="1"><xf fillId="0" applyFill="0" xfId="0"/></cellXfs>
        </styleSheet>
        """;

    private static string InheritedFillStyles() =>
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fills count="2">
            <fill><patternFill patternType="none"/></fill>
            <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
          </fills>
          <cellStyleXfs count="1"><xf fillId="1" applyFill="1"/></cellStyleXfs>
          <cellXfs count="2">
            <xf fillId="0" xfId="0"/>
            <xf fillId="0" applyFill="0" xfId="0"/>
          </cellXfs>
        </styleSheet>
        """;

    private static string UnusedGray125Styles() =>
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fills count="3">
            <fill><patternFill patternType="none"/></fill>
            <fill><patternFill patternType="gray125"/></fill>
            <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
          </fills>
          <cellStyleXfs count="1"><xf fillId="0"/></cellStyleXfs>
          <cellXfs count="1"><xf fillId="0" xfId="0"/></cellXfs>
        </styleSheet>
        """;

    private static string ApplyFillZeroInheritsNoneStyles() =>
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <fills count="2">
            <fill><patternFill patternType="none"/></fill>
            <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
          </fills>
          <cellStyleXfs count="1"><xf fillId="0"/></cellStyleXfs>
          <cellXfs count="1"><xf fillId="1" applyFill="0" xfId="0"/></cellXfs>
        </styleSheet>
        """;

    private static void Write(ZipArchive zip, string name, string xml)
    {
        var entry = zip.CreateEntry(name, CompressionLevel.Fastest);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(xml.Trim());
    }

    private static string BlankSheet() =>
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData/>
        </worksheet>
        """;

    private static string ValueSheet() =>
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>x</t></is></c></row></sheetData>
        </worksheet>
        """;

    private static string MergeSheet() =>
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData/>
          <mergeCells count="1"><mergeCell ref="A1:B1"/></mergeCells>
        </worksheet>
        """;

    private static string CellWithoutStyleSheet() =>
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData><row r="1"><c r="A1"/></row></sheetData>
        </worksheet>
        """;

    private static string FilledCellSheet() =>
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData><row r="1"><c r="A1" s="1"/></row></sheetData>
        </worksheet>
        """;

    private static string FilledRowSheet() =>
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <sheetData><row r="1" s="1" customFormat="1"/></sheetData>
        </worksheet>
        """;

    private static string FilledColSheet() =>
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <cols><col min="1" max="1" style="1"/></cols>
          <sheetData/>
        </worksheet>
        """;

    private static string FilledColOutsideSheet() =>
        """
        <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
        <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          <cols><col min="70" max="70" style="1"/></cols>
          <sheetData/>
        </worksheet>
        """;
}
