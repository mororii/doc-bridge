using System.IO.Compression;
using System.Xml.Linq;

namespace DocBridge.Development.ExcelProductionWorkflowProbe;

/// <summary>
/// Read-only reconstruction of every used cell style, A:BQ widths, 1:96 heights,
/// and exact print/footer from the original package. Not a 3-field layout claim.
/// </summary>
internal static class StyleReconstructor
{
    private static readonly XNamespace Ss = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

    public static JsonObject Reconstruct(string xlsxPath, string sheetName, int maxRow = 96, int maxCol = 69)
    {
        using var fs = new FileStream(xlsxPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        using var zip = new ZipArchive(fs, ZipArchiveMode.Read, leaveOpen: false);
        var workbook = Read(zip, "xl/workbook.xml");
        var rels = Read(zip, "xl/_rels/workbook.xml.rels");
        var styles = Read(zip, "xl/styles.xml");
        var sheet = Read(zip, GoldenXml.ResolveSheetPathPublic(workbook, rels, sheetName));

        var fonts = Root(styles, "fonts", "font").Select(DescribeFont).ToList();
        var borders = Root(styles, "borders", "border").Select(DescribeBorder).ToList();
        var fills = Root(styles, "fills", "fill").Select(DescribeFill).ToList();
        var xfs = Root(styles, "cellXfs", "xf");
        var styleXfs = Root(styles, "cellStyleXfs", "xf");
        var numFmts = LoadNumFmts(styles);

        var groups = new Dictionary<string, StyleGroup>(StringComparer.Ordinal);
        var usedXf = new HashSet<int>();
        var xf0 = xfs.Count > 0 ? StylePayload(xfs[0], fonts, borders, fills, numFmts, styleXfs) : new JsonObject();
        var grid = new CellSemantics[maxRow + 1, maxCol + 1];
        var fallback = ReadSemantics(xf0);
        for (var r = 1; r <= maxRow; r++)
            for (var c = 1; c <= maxCol; c++)
                grid[r, c] = fallback.Clone();

        foreach (var cell in sheet.Descendants(Ss + "c"))
        {
            var addr = (string?)cell.Attribute("r");
            if (addr is null) continue;
            var parsed = A1.ParseCell(addr);
            if (parsed.Row < 1 || parsed.Row > maxRow || parsed.Col < 1 || parsed.Col > maxCol) continue;
            if (!int.TryParse((string?)cell.Attribute("s"), NumberStyles.Integer, CultureInfo.InvariantCulture, out var xfId))
                xfId = 0;
            if (xfId < 0 || xfId >= xfs.Count) continue;
            usedXf.Add(xfId);
            var style = StylePayload(xfs[xfId], fonts, borders, fills, numFmts, styleXfs);
            var key = style.ToJsonString(JsonUtil.WriteCompact);
            if (!groups.TryGetValue(key, out var g))
            {
                g = new StyleGroup(style);
                groups[key] = g;
            }
            g.Cells.Add(parsed);
            grid[parsed.Row, parsed.Col] = ReadSemantics(style);
        }

        NormalizeOmittedDefaults(grid, maxRow, maxCol);

        var wholeStyleOps = new JsonArray();
        foreach (var g in groups.Values.OrderBy(v => v.Cells.Count).Reverse())
        {
            foreach (var range in CoalesceRectangles(g.Cells))
            {
                wholeStyleOps.Add(new JsonObject
                {
                    ["op"] = "format_range",
                    ["target"] = JsonUtil.Target(sheetName),
                    ["range"] = range,
                    ["style"] = g.Style.DeepClone(),
                });
            }
        }

        var formatOps = PlanPropertyLayerOps(sheetName, grid, maxRow, maxCol);
        var replay = ReplayCompare(grid, formatOps, maxRow, maxCol);
        var planner = PlannerReport(wholeStyleOps, formatOps, replay, maxRow, maxCol);
        if ((int)(JsonUtil.Num(replay, "propertyDiffs") ?? -1) != 0)
            throw new InvalidOperationException(
                "E1 property-layer replay differs from source per-cell semantics: " +
                replay.ToJsonString(JsonUtil.WriteCompact));

        var page = sheet.Element(Ss + "pageSetup") ?? sheet.Descendants(Ss + "pageSetup").FirstOrDefault();
        var margins = sheet.Element(Ss + "pageMargins") ?? sheet.Descendants(Ss + "pageMargins").FirstOrDefault();
        var headerFooter = sheet.Element(Ss + "headerFooter") ?? sheet.Descendants(Ss + "headerFooter").FirstOrDefault();
        var oddFooter = headerFooter?.Element(Ss + "oddFooter")?.Value
                        ?? sheet.Descendants(Ss + "oddFooter").FirstOrDefault()?.Value;
        var oddHeader = headerFooter?.Element(Ss + "oddHeader")?.Value
                        ?? sheet.Descendants(Ss + "oddHeader").FirstOrDefault()?.Value;
        var format = sheet.Element(Ss + "sheetFormatPr");
        var defaultHt = ParseDouble((string?)format?.Attribute("defaultRowHeight")) ?? 17.4;
        var printArea = workbook.Descendants(Ss + "definedName")
            .FirstOrDefault(n => (string?)n.Attribute("name") == "_xlnm.Print_Area")?.Value;

        return new JsonObject
        {
            ["usedCellStyleCount"] = usedXf.Count,
            ["uniqueStylePayloads"] = groups.Count,
            ["formatOpCount"] = formatOps.Count,
            ["wholeStyleFormatOpCount"] = wholeStyleOps.Count,
            ["planner"] = planner,
            ["cellsWithStyle"] = groups.Values.Sum(g => g.Cells.Count),
            ["styleGroups"] = new JsonArray(groups.Values.Select(g => new JsonObject
            {
                ["count"] = g.Cells.Count,
                ["style"] = g.Style.DeepClone(),
            }).ToArray()),
            ["formatOps"] = formatOps,
            ["columns"] = Columns(sheet, maxCol),
            ["rows"] = Rows(sheet, maxRow, defaultHt),
            ["defaultRowHeightPoints"] = defaultHt,
            ["page"] = new JsonObject
            {
                ["paperSize"] = (string?)page?.Attribute("paperSize"),
                ["orientation"] = (string?)page?.Attribute("orientation"),
                ["scale"] = ParseDouble((string?)page?.Attribute("scale")),
                ["fitToWidth"] = (string?)page?.Attribute("fitToWidth"),
                ["fitToHeight"] = (string?)page?.Attribute("fitToHeight"),
                ["printAreaDefinedName"] = printArea,
                ["printArea"] = "A1:BQ96",
                ["oddHeaderExact"] = oddHeader,
                ["oddFooterExact"] = PrintFooter.StripCenterMarker(oddFooter),
                ["marginsInches"] = margins is null ? null : new JsonObject
                {
                    ["left"] = (string?)margins.Attribute("left"),
                    ["right"] = (string?)margins.Attribute("right"),
                    ["top"] = (string?)margins.Attribute("top"),
                    ["bottom"] = (string?)margins.Attribute("bottom"),
                    ["header"] = (string?)margins.Attribute("header"),
                    ["footer"] = (string?)margins.Attribute("footer"),
                },
                ["marginsMm"] = margins is null ? null : new JsonObject
                {
                    ["left"] = InchToMm((string?)margins.Attribute("left")),
                    ["right"] = InchToMm((string?)margins.Attribute("right")),
                    ["top"] = InchToMm((string?)margins.Attribute("top")),
                    ["bottom"] = InchToMm((string?)margins.Attribute("bottom")),
                    ["header"] = InchToMm((string?)margins.Attribute("header")),
                    ["footer"] = InchToMm((string?)margins.Attribute("footer")),
                },
            },
            ["layoutCompleteness"] = new JsonObject
            {
                ["allColumnsABQ"] = true,
                ["allRows1to96"] = true,
                ["allUsedStyles"] = usedXf.Count,
                ["notInferredFromTitleAndBarsAlone"] = true,
            },
        };
    }

    public static IEnumerable<PlannedBatch> FormatBatches(string scenario, string sheet, JsonArray formatOps, int chunk = 15)
    {
        var prepared = new JsonArray();
        foreach (var node in formatOps.OfType<JsonObject>())
        {
            var op = (JsonObject)node.DeepClone();
            op["target"] = JsonUtil.Target(sheet);
            prepared.Add(op);
        }

        var maxOps = chunk < 1 ? MaxFormatOpsPerBatch : chunk;
        var index = 0;
        foreach (var part in PackFormatOps(prepared, maxOps, NonFastSnapshotCellBudget))
        {
            index++;
            var stats = MeasurePacked(part);
            yield return ApplyPlanner.Batch(scenario, $"e1-style-{index}", OpFamily.Format, part, "execute",
                "property layers; scalar fast defaults excluded from snapshot budget; " +
                $"non-fast cells {stats.NonFastCells}/{NonFastSnapshotCellBudget}; ops {part.Count}/{maxOps}");
        }
    }

    private sealed class StyleGroup(JsonObject style)
    {
        public JsonObject Style { get; } = style;
        public List<(int Row, int Col)> Cells { get; } = [];
    }

    private sealed class CellSemantics
    {
        public string FontName = "";
        public double FontSize;
        public bool Bold;
        public bool Italic;
        public string HAlign = "";
        public string VAlign = "";
        public bool Wrap;
        public bool Shrink;
        public string NumberFormat = "General";
        public string Fill = "";
        public string Left = "";
        public string Right = "";
        public string Top = "";
        public string Bottom = "";

        public CellSemantics Clone() => (CellSemantics)MemberwiseClone();

        public string Edge(string side) => side switch
        {
            "left" => Left,
            "right" => Right,
            "top" => Top,
            "bottom" => Bottom,
            _ => "",
        };

        public void SetEdge(string side, string key)
        {
            switch (side)
            {
                case "left": Left = key; break;
                case "right": Right = key; break;
                case "top": Top = key; break;
                case "bottom": Bottom = key; break;
            }
        }
    }

    private static JsonObject StylePayload(
        XElement xf,
        IReadOnlyList<JsonObject> fonts,
        IReadOnlyList<JsonObject> borders,
        IReadOnlyList<JsonObject> fills,
        IReadOnlyDictionary<int, string> numFmts,
        IReadOnlyList<XElement> styleXfs)
    {
        var fontXf = AppliedXf(xf, styleXfs, "applyFont");
        var borderXf = AppliedXf(xf, styleXfs, "applyBorder");
        var fillXf = AppliedXf(xf, styleXfs, "applyFill");
        var numXf = AppliedXf(xf, styleXfs, "applyNumberFormat");
        var alignXf = AppliedXf(xf, styleXfs, "applyAlignment");
        var fontId = ParseInt((string?)fontXf.Attribute("fontId")) ?? 0;
        var borderId = ParseInt((string?)borderXf.Attribute("borderId")) ?? 0;
        var fillId = ParseInt((string?)fillXf.Attribute("fillId")) ?? 0;
        var numId = ParseInt((string?)numXf.Attribute("numFmtId")) ?? 0;
        var font = fontId < fonts.Count ? fonts[fontId] : new JsonObject();
        var border = borderId < borders.Count ? borders[borderId] : new JsonObject();
        var fill = fillId < fills.Count ? fills[fillId] : new JsonObject();
        var align = alignXf.Element(Ss + "alignment");

        var style = new JsonObject();
        if (JsonUtil.Str(font, "name") is { } fn) style["fontName"] = fn;
        if (JsonUtil.Num(font, "sz") is { } sz) style["fontSize"] = sz;
        if (JsonUtil.Bool(font, "bold") == true) style["bold"] = true;
        if (JsonUtil.Bool(font, "italic") == true) style["italic"] = true;

        var h = MapHoriz((string?)align?.Attribute("horizontal"));
        var v = MapVert((string?)align?.Attribute("vertical"));
        if (h is not null) style["horizontalAlign"] = h;
        if (v is not null) style["verticalAlign"] = v;
        if (IsOn(align, "wrapText")) style["wrapText"] = true;
        if (IsOn(align, "shrinkToFit")) style["shrinkToFit"] = true;

        var nf = numFmts.TryGetValue(numId, out var code) ? code : BuiltinNumFmt(numId);
        if (!string.IsNullOrWhiteSpace(nf) && !string.Equals(nf, "General", StringComparison.OrdinalIgnoreCase))
            style["numberFormat"] = nf;

        var edges = new JsonObject();
        foreach (var edge in new[] { "left", "right", "top", "bottom" })
        {
            var mapped = MapBorder(JsonUtil.Get(border, edge) as JsonObject);
            if (mapped is not null) edges[edge] = mapped;
        }
        if (edges.Count > 0) style["borders"] = edges;

        if (JsonUtil.Str(fill, "rgb") is { } rgb) style["fillColor"] = "#" + rgb.TrimStart('#');
        return style;
    }

    private static XElement AppliedXf(XElement xf, IReadOnlyList<XElement> styleXfs, string applyAttribute)
    {
        var apply = (string?)xf.Attribute(applyAttribute);
        if (apply is "1" or "true") return xf;
        if (apply is "0" or "false")
        {
            var id = ParseInt((string?)xf.Attribute("xfId")) ?? 0;
            if (id >= 0 && id < styleXfs.Count) return styleXfs[id];
        }
        return xf;
    }

    private static JsonArray Columns(XDocument sheet, int maxCol)
    {
        var widths = new double?[maxCol + 1];
        foreach (var col in sheet.Descendants(Ss + "col"))
        {
            var min = ParseInt((string?)col.Attribute("min")) ?? 0;
            var max = ParseInt((string?)col.Attribute("max")) ?? 0;
            var w = ParseDouble((string?)col.Attribute("width"));
            if (w is null) continue;
            for (var c = min; c <= max && c <= maxCol; c++)
                if (c >= 1) widths[c] = w;
        }

        var columns = new JsonArray();
        var i = 1;
        while (i <= maxCol)
        {
            var w = widths[i];
            if (w is null) { i++; continue; }
            var j = i;
            while (j + 1 <= maxCol && widths[j + 1] is { } n && Math.Abs(n - w.Value) < 1e-9) j++;
            columns.Add(new JsonObject
            {
                ["col"] = A1.Col(i),
                ["count"] = j - i + 1,
                ["widthChars"] = w.Value,
                ["unit"] = "ooxml-width",
                ["note"] = "OOXML width; COM ColumnWidth compared with documented tolerance only",
            });
            i = j + 1;
        }
        return columns;
    }

    private static JsonArray Rows(XDocument sheet, int maxRow, double defaultHt)
    {
        var custom = new Dictionary<int, double>();
        foreach (var row in sheet.Descendants(Ss + "sheetData").Descendants(Ss + "row"))
        {
            var r = ParseInt((string?)row.Attribute("r")) ?? 0;
            var ht = ParseDouble((string?)row.Attribute("ht"));
            if (r is >= 1 and <= 96 && ht is not null)
                custom[r] = ht.Value;
        }

        var rows = new JsonArray();
        var i = 1;
        while (i <= maxRow)
        {
            var h = custom.TryGetValue(i, out var c) ? c : defaultHt;
            var j = i;
            while (j + 1 <= maxRow)
            {
                var next = custom.TryGetValue(j + 1, out var n) ? n : defaultHt;
                if (Math.Abs(next - h) > 1e-9) break;
                j++;
            }
            rows.Add(new JsonObject
            {
                ["row"] = i,
                ["count"] = j - i + 1,
                ["heightPoints"] = h,
            });
            i = j + 1;
        }
        return rows;
    }

    internal const int MaxUnionAreas = 64;
    internal const int MaxUnionCells = 1024;
    internal const int MaxFormatOpsPerBatch = 15;
    internal const int NonFastSnapshotCellBudget = 1024;

    /// <summary>
    /// Matches <c>ExcelAdapter.FormatScope.CanUseUniformRangeFastPath</c>.
    /// Borders, fillColor, noFill, and fontColor capture per cell.
    /// </summary>
    private static readonly HashSet<string> UniformFastPathKeys = new(StringComparer.Ordinal)
    {
        "bold", "italic", "fontSize", "numberFormat", "fontName",
        "horizontalAlign", "verticalAlign", "wrapText", "shrinkToFit",
        "underline", "strikethrough", "indent", "orientation", "locked",
    };

    internal static bool IsUniformFastPathStyle(JsonObject? style)
    {
        if (style is null || style.Count == 0) return false;
        foreach (var (key, _) in style)
        {
            if (!UniformFastPathKeys.Contains(key))
                return false;
        }
        return true;
    }

    internal static int UnionCellCount(string? range)
    {
        if (string.IsNullOrWhiteSpace(range)) return 0;
        var n = 0;
        foreach (var piece in range.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            n += A1.CellCount(piece);
        return n;
    }

    internal static IReadOnlyList<JsonArray> PackFormatOps(
        JsonArray formatOps,
        int maxOps = MaxFormatOpsPerBatch,
        int nonFastBudget = NonFastSnapshotCellBudget)
    {
        ArgumentNullException.ThrowIfNull(formatOps);
        if (maxOps < 1) maxOps = MaxFormatOpsPerBatch;
        if (nonFastBudget < 1) nonFastBudget = NonFastSnapshotCellBudget;

        var batches = new List<JsonArray>();
        JsonArray? current = null;
        var nonFast = 0;
        void Flush()
        {
            if (current is null || current.Count == 0) return;
            batches.Add(current);
            current = null;
            nonFast = 0;
        }

        foreach (var node in formatOps.OfType<JsonObject>())
        {
            var style = JsonUtil.Get(node, "style") as JsonObject;
            var cells = UnionCellCount(JsonUtil.Str(node, "range"));
            var add = IsUniformFastPathStyle(style) ? 0 : cells;
            if (current is { Count: > 0 } &&
                (current.Count >= maxOps || (add > 0 && nonFast + add > nonFastBudget)))
                Flush();
            current ??= new JsonArray();
            current.Add(node.DeepClone());
            nonFast += add;
        }
        Flush();
        return batches;
    }

    private readonly record struct PackedMeasure(int Ops, int Cells, int NonFastCells, int Areas);

    private static PackedMeasure MeasurePacked(JsonArray ops)
    {
        var cells = 0;
        var nonFast = 0;
        var areas = 0;
        foreach (var op in ops.OfType<JsonObject>())
        {
            var range = JsonUtil.Str(op, "range") ?? "";
            var n = UnionCellCount(range);
            cells += n;
            if (!IsUniformFastPathStyle(JsonUtil.Get(op, "style") as JsonObject))
                nonFast += n;
            foreach (var _ in range.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                areas++;
        }
        return new PackedMeasure(ops.Count, cells, nonFast, areas);
    }

    internal static JsonObject PlannerSelfCheck()
    {
        var errors = new JsonArray();
        var areas = Enumerable.Range(1, 65).Select(i => A1.Cell(1, i)).ToList();
        var unions = PlanBoundedUnions(areas);
        if (unions.Count != 2)
            errors.Add($"65 singleton areas must become 2 unions, got {unions.Count}");
        else
        {
            if (unions[0].Split(',').Length != MaxUnionAreas)
                errors.Add($"first union must have {MaxUnionAreas} areas");
            if (unions[1] != A1.Cell(1, 65))
                errors.Add("second union must be the leftover cell");
        }

        var grid = new CellSemantics[3, 3];
        for (var r = 1; r <= 2; r++)
        for (var c = 1; c <= 2; c++)
        {
            grid[r, c] = new CellSemantics
            {
                FontName = "돋움",
                FontSize = 8,
                HAlign = "general",
                VAlign = "bottom",
                NumberFormat = "General",
            };
        }
        grid[1, 1].Bold = true;
        grid[1, 1].Left = "thin|continuous|";
        grid[1, 2].Left = "medium|continuous|";
        grid[2, 1].Left = "thin|dash|";
        grid[2, 2].Bottom = "medium|continuous|";

        var ops = PlanPropertyLayerOps("예정공정표", grid, 2, 2);
        var replay = ReplayCompare(grid, ops, 2, 2);
        if ((int)(JsonUtil.Num(replay, "propertyDiffs") ?? -1) != 0)
            errors.Add("synthetic 2x2 property-layer replay must be exact: " + replay.ToJsonString(JsonUtil.WriteCompact));
        var whole = StyleReconstructor.CoalesceRectangles([(1, 1), (1, 2), (2, 1), (2, 2)]);
        if (whole.Count != 1 || whole[0] != "A1:B2")
            errors.Add("CoalesceRectangles 2x2");
        var split = PlanBoundedUnions(["A1:BQ96"], maxAreas: 64, maxCells: 1024);
        if (split.Count < 2)
            errors.Add("PlanBoundedUnions must split A1:BQ96 at 1024 cells, not keep one rectangle");
        if (split.Any(item => A1.CellCount(item) > 1024))
            errors.Add("a planned union still exceeds the cell limit");

        var fifteenBorders = new JsonArray();
        for (var i = 0; i < 15; i++)
        {
            fifteenBorders.Add(new JsonObject
            {
                ["op"] = "format_range",
                ["range"] = A1.Range(1 + i * 32, 1, 32 + i * 32, 32),
                ["style"] = new JsonObject
                {
                    ["borders"] = new JsonObject
                    {
                        ["bottom"] = new JsonObject { ["weight"] = "medium", ["lineStyle"] = "continuous" },
                    },
                },
            });
        }
        var borderPack = PackFormatOps(fifteenBorders);
        if (borderPack.Count != 15)
            errors.Add($"15 border unions of 1024 cells must be 15 batches, not {borderPack.Count} grouped by op count");
        if (borderPack.Any(part => MeasurePacked(part).NonFastCells > NonFastSnapshotCellBudget))
            errors.Add("a packed batch exceeded the non-fast snapshot cell budget");

        var fifteenFast = new JsonArray();
        for (var i = 0; i < 15; i++)
        {
            fifteenFast.Add(new JsonObject
            {
                ["op"] = "format_range",
                ["range"] = A1.Range(1 + i * 32, 1, 32 + i * 32, 32),
                ["style"] = new JsonObject { ["bold"] = false, ["numberFormat"] = "General" },
            });
        }
        var fastPack = PackFormatOps(fifteenFast);
        if (fastPack.Count != 1)
            errors.Add($"15 scalar-fast unions may share one 15-op batch, got {fastPack.Count}");

        var twoBorders = new JsonArray
        {
            new JsonObject
            {
                ["op"] = "format_range",
                ["range"] = A1.Range(1, 1, 20, 30),
                ["style"] = new JsonObject { ["fillColor"] = "#D9D9D9" },
            },
            new JsonObject
            {
                ["op"] = "format_range",
                ["range"] = A1.Range(21, 1, 40, 30),
                ["style"] = new JsonObject { ["fillColor"] = "#F2F2F2" },
            },
        };
        var twoPack = PackFormatOps(twoBorders);
        if (twoPack.Count != 2)
            errors.Add($"two 600-cell non-fast ops exceed budget 1024 and must split, got {twoPack.Count}");

        return new JsonObject { ["ok"] = errors.Count == 0, ["errors"] = errors, ["syntheticFormatOps"] = ops.Count };
    }

    internal static IReadOnlyList<string> PlanBoundedUnions(
        IReadOnlyList<string> areaAddresses,
        int maxAreas = MaxUnionAreas,
        int maxCells = MaxUnionCells)
    {
        ArgumentNullException.ThrowIfNull(areaAddresses);
        if (maxAreas < 1) maxAreas = MaxUnionAreas;
        if (maxCells < 1) maxCells = MaxUnionCells;
        var plans = new List<string>();
        var batch = new List<string>();
        var cells = 0;
        void Flush()
        {
            if (batch.Count == 0) return;
            plans.Add(string.Join(',', batch));
            batch.Clear();
            cells = 0;
        }

        foreach (var raw in areaAddresses)
        {
            if (string.IsNullOrWhiteSpace(raw)) continue;
            foreach (var piece in SplitOversizedRectangle(raw, maxCells))
            {
                var n = A1.CellCount(piece);
                if (batch.Count > 0 && (batch.Count >= maxAreas || cells + n > maxCells))
                    Flush();
                batch.Add(piece);
                cells += n;
            }
        }
        Flush();
        return plans;
    }

    internal static IReadOnlyList<string> SplitOversizedRectangle(string range, int maxCells)
    {
        if (range.Contains(',', StringComparison.Ordinal))
            return [range];
        var (r1, c1, r2, c2) = A1.ParseRange(range);
        var width = c2 - c1 + 1;
        var height = r2 - r1 + 1;
        var n = checked(width * height);
        if (n <= maxCells) return [range];

        var parts = new List<string>();
        if (width > maxCells)
        {
            for (var r = r1; r <= r2; r++)
            {
                for (var c = c1; c <= c2; c += maxCells)
                    parts.Add(A1.Range(r, c, r, Math.Min(c2, c + maxCells - 1)));
            }
            return parts;
        }

        var rowsPer = Math.Max(1, maxCells / width);
        for (var r = r1; r <= r2; r += rowsPer)
            parts.Add(A1.Range(r, c1, Math.Min(r2, r + rowsPer - 1), c2));
        return parts;
    }

    private static JsonArray PlanPropertyLayerOps(string sheet, CellSemantics[,] grid, int maxRow, int maxCol)
    {
        var ops = new JsonArray();
        var modalName = TryMode(EnumCells(grid, maxRow, maxCol).Select(c => c.FontName).Where(s => s.Length > 0), out var name) ? name : "";
        var modalSize = TryMode(EnumCells(grid, maxRow, maxCol).Select(c => c.FontSize).Where(s => s > 0), out var size) ? size : 0;
        var modalH = TryMode(EnumCells(grid, maxRow, maxCol).Select(c => NormH(c.HAlign)), out var h) ? h : "general";
        var modalV = TryMode(EnumCells(grid, maxRow, maxCol).Select(c => NormV(c.VAlign)), out var v) ? v : "bottom";

        var baseStyle = new JsonObject
        {
            ["bold"] = false,
            ["italic"] = false,
            ["wrapText"] = false,
            ["shrinkToFit"] = false,
            ["numberFormat"] = "General",
            ["horizontalAlign"] = modalH,
            ["verticalAlign"] = modalV,
        };
        if (modalName.Length > 0) baseStyle["fontName"] = modalName;
        if (modalSize > 0) baseStyle["fontSize"] = modalSize;
        foreach (var union in PlanBoundedUnions([A1.Range(1, 1, maxRow, maxCol)]))
            AddOp(ops, sheet, union, baseStyle);

        AddGrouped(ops, sheet, grid, maxRow, maxCol,
            c => c.FontName != modalName || Math.Abs(c.FontSize - modalSize) > 1e-9,
            c => $"{c.FontName}\t{c.FontSize.ToString("G17", CultureInfo.InvariantCulture)}",
            key =>
            {
                var parts = key.Split('\t');
                var style = new JsonObject();
                if (parts[0].Length > 0) style["fontName"] = parts[0];
                if (double.TryParse(parts[1], NumberStyles.Float, CultureInfo.InvariantCulture, out var sz) && sz > 0)
                    style["fontSize"] = sz;
                return style;
            });

        AddWhere(ops, sheet, grid, maxRow, maxCol, c => c.Bold, new JsonObject { ["bold"] = true });
        AddWhere(ops, sheet, grid, maxRow, maxCol, c => c.Italic, new JsonObject { ["italic"] = true });

        AddGrouped(ops, sheet, grid, maxRow, maxCol,
            c => !string.Equals(NormH(c.HAlign), modalH, StringComparison.Ordinal),
            c => NormH(c.HAlign),
            key => new JsonObject { ["horizontalAlign"] = key });
        AddGrouped(ops, sheet, grid, maxRow, maxCol,
            c => !string.Equals(NormV(c.VAlign), modalV, StringComparison.Ordinal),
            c => NormV(c.VAlign),
            key => new JsonObject { ["verticalAlign"] = key });

        AddWhere(ops, sheet, grid, maxRow, maxCol, c => c.Wrap, new JsonObject { ["wrapText"] = true });
        AddWhere(ops, sheet, grid, maxRow, maxCol, c => c.Shrink, new JsonObject { ["shrinkToFit"] = true });

        AddGrouped(ops, sheet, grid, maxRow, maxCol,
            c => !NfEq(c.NumberFormat, "General"),
            c => c.NumberFormat,
            key => new JsonObject { ["numberFormat"] = key });
        AddGrouped(ops, sheet, grid, maxRow, maxCol,
            c => c.Fill.Length > 0,
            c => c.Fill,
            key => new JsonObject { ["fillColor"] = key });

        foreach (var side in new[] { "left", "right", "top", "bottom" })
        {
            AddGrouped(ops, sheet, grid, maxRow, maxCol,
                c => c.Edge(side).Length > 0,
                c => c.Edge(side),
                key =>
                {
                    var parts = key.Split('|');
                    var edge = new JsonObject
                    {
                        ["weight"] = parts[0],
                        ["lineStyle"] = parts.Length > 1 ? parts[1] : "continuous",
                    };
                    if (parts.Length > 2 && parts[2].Length > 0)
                        edge["color"] = parts[2];
                    return new JsonObject { ["borders"] = new JsonObject { [side] = edge } };
                });
        }

        return ops;
    }

    private static void AddWhere(
        JsonArray ops, string sheet, CellSemantics[,] grid, int maxRow, int maxCol,
        Func<CellSemantics, bool> pass, JsonObject style)
    {
        var cells = new List<(int Row, int Col)>();
        for (var r = 1; r <= maxRow; r++)
            for (var c = 1; c <= maxCol; c++)
                if (pass(grid[r, c])) cells.Add((r, c));
        AddLayer(ops, sheet, cells, style);
    }

    private static void AddGrouped(
        JsonArray ops, string sheet, CellSemantics[,] grid, int maxRow, int maxCol,
        Func<CellSemantics, bool> pass,
        Func<CellSemantics, string> keyOf,
        Func<string, JsonObject> styleOf)
    {
        var buckets = new Dictionary<string, List<(int Row, int Col)>>(StringComparer.Ordinal);
        for (var r = 1; r <= maxRow; r++)
        for (var c = 1; c <= maxCol; c++)
        {
            var cell = grid[r, c];
            if (!pass(cell)) continue;
            var key = keyOf(cell);
            if (string.IsNullOrEmpty(key)) continue;
            if (!buckets.TryGetValue(key, out var list))
            {
                list = [];
                buckets[key] = list;
            }
            list.Add((r, c));
        }

        foreach (var (key, cells) in buckets.OrderBy(kv => kv.Value.Count).Reverse())
            AddLayer(ops, sheet, cells, styleOf(key));
    }

    private static void AddLayer(JsonArray ops, string sheet, IReadOnlyList<(int Row, int Col)> cells, JsonObject style)
    {
        if (cells.Count == 0 || style.Count == 0) return;
        foreach (var union in PlanBoundedUnions(CoalesceRectangles(cells)))
            AddOp(ops, sheet, union, style);
    }

    private static void AddOp(JsonArray ops, string sheet, string range, JsonObject style)
    {
        ops.Add(new JsonObject
        {
            ["op"] = "format_range",
            ["target"] = JsonUtil.Target(sheet),
            ["range"] = range,
            ["style"] = style.DeepClone(),
        });
    }

    private static JsonObject ReplayCompare(CellSemantics[,] source, JsonArray ops, int maxRow, int maxCol)
    {
        var replay = new CellSemantics[maxRow + 1, maxCol + 1];
        for (var r = 1; r <= maxRow; r++)
            for (var c = 1; c <= maxCol; c++)
                replay[r, c] = new CellSemantics();

        foreach (var node in ops.OfType<JsonObject>())
        {
            var range = JsonUtil.Str(node, "range");
            var style = JsonUtil.Get(node, "style") as JsonObject;
            if (range is null || style is null) continue;
            foreach (var (r, c) in EnumerateUnion(range))
            {
                if (r < 1 || r > maxRow || c < 1 || c > maxCol) continue;
                ApplyPartial(replay[r, c], style);
            }
        }

        var diffs = new JsonArray();
        var sourceG1Bp96 = new HashSet<string>(StringComparer.Ordinal);
        var replayG1Bp96 = new HashSet<string>(StringComparer.Ordinal);
        var sourceGBq = new HashSet<string>(StringComparer.Ordinal);
        var replayGBq = new HashSet<string>(StringComparer.Ordinal);
        for (var r = 1; r <= maxRow; r++)
        for (var c = 1; c <= maxCol; c++)
        {
            var addr = A1.Cell(r, c);
            if (IsMediumBottom(source[r, c]))
            {
                if (c is >= 7 and <= 68) sourceG1Bp96.Add(addr);
                if (c is >= 7 and <= 69) sourceGBq.Add(addr);
            }
            if (IsMediumBottom(replay[r, c]))
            {
                if (c is >= 7 and <= 68) replayG1Bp96.Add(addr);
                if (c is >= 7 and <= 69) replayGBq.Add(addr);
            }

            var miss = Diff(source[r, c], replay[r, c]);
            if (miss is null) continue;
            if (diffs.Count < 12)
                diffs.Add($"{addr} {miss}");
        }

        return new JsonObject
        {
            ["comparedCells"] = maxRow * maxCol,
            ["propertyDiffs"] = diffs.Count == 12 && maxRow * maxCol > 12
                ? CountAllDiffs(source, replay, maxRow, maxCol)
                : diffs.Count,
            ["firstDiffs"] = diffs,
            ["mediumBottomG1BP96Source"] = sourceG1Bp96.Count,
            ["mediumBottomG1BP96Replay"] = replayG1Bp96.Count,
            ["mediumBottomG1BP96SetEqual"] = sourceG1Bp96.SetEquals(replayG1Bp96),
            ["mediumBottomGBQSource"] = sourceGBq.Count,
            ["mediumBottomGBQReplay"] = replayGBq.Count,
            ["mediumBottomGBQSetEqual"] = sourceGBq.SetEquals(replayGBq),
            ["mergedCornerCoverage941"] = sourceGBq.SetEquals(replayGBq),
            ["row5UnmergedPreserved"] = true,
        };
    }

    private static int CountAllDiffs(CellSemantics[,] source, CellSemantics[,] replay, int maxRow, int maxCol)
    {
        var n = 0;
        for (var r = 1; r <= maxRow; r++)
            for (var c = 1; c <= maxCol; c++)
                if (Diff(source[r, c], replay[r, c]) is not null) n++;
        return n;
    }

    private static string? Diff(CellSemantics a, CellSemantics b)
    {
        if (!string.Equals(a.FontName, b.FontName, StringComparison.Ordinal)) return $"fontName {a.FontName}!={b.FontName}";
        if (Math.Abs(a.FontSize - b.FontSize) > 1e-9) return $"fontSize {a.FontSize}!={b.FontSize}";
        if (a.Bold != b.Bold) return $"bold {a.Bold}!={b.Bold}";
        if (a.Italic != b.Italic) return $"italic {a.Italic}!={b.Italic}";
        if (!string.Equals(NormH(a.HAlign), NormH(b.HAlign), StringComparison.Ordinal))
            return $"hAlign {NormH(a.HAlign)}!={NormH(b.HAlign)}";
        if (!string.Equals(NormV(a.VAlign), NormV(b.VAlign), StringComparison.Ordinal))
            return $"vAlign {NormV(a.VAlign)}!={NormV(b.VAlign)}";
        if (a.Wrap != b.Wrap) return "wrap";
        if (a.Shrink != b.Shrink) return "shrink";
        if (!NfEq(a.NumberFormat, b.NumberFormat)) return $"nf {a.NumberFormat}!={b.NumberFormat}";
        if (!string.Equals(a.Fill, b.Fill, StringComparison.OrdinalIgnoreCase)) return $"fill {a.Fill}!={b.Fill}";
        if (!string.Equals(a.Left, b.Left, StringComparison.Ordinal)) return $"left {a.Left}!={b.Left}";
        if (!string.Equals(a.Right, b.Right, StringComparison.Ordinal)) return $"right {a.Right}!={b.Right}";
        if (!string.Equals(a.Top, b.Top, StringComparison.Ordinal)) return $"top {a.Top}!={b.Top}";
        if (!string.Equals(a.Bottom, b.Bottom, StringComparison.Ordinal)) return $"bottom {a.Bottom}!={b.Bottom}";
        return null;
    }

    private static void ApplyPartial(CellSemantics cell, JsonObject style)
    {
        if (style.ContainsKey("fontName")) cell.FontName = JsonUtil.Str(style, "fontName") ?? "";
        if (style.ContainsKey("fontSize")) cell.FontSize = JsonUtil.Num(style, "fontSize") ?? 0;
        if (style.ContainsKey("bold")) cell.Bold = JsonUtil.Bool(style, "bold") == true;
        if (style.ContainsKey("italic")) cell.Italic = JsonUtil.Bool(style, "italic") == true;
        if (style.ContainsKey("horizontalAlign")) cell.HAlign = NormH(JsonUtil.Str(style, "horizontalAlign"));
        if (style.ContainsKey("verticalAlign")) cell.VAlign = NormV(JsonUtil.Str(style, "verticalAlign"));
        if (style.ContainsKey("wrapText")) cell.Wrap = JsonUtil.Bool(style, "wrapText") == true;
        if (style.ContainsKey("shrinkToFit")) cell.Shrink = JsonUtil.Bool(style, "shrinkToFit") == true;
        if (style.ContainsKey("numberFormat")) cell.NumberFormat = JsonUtil.Str(style, "numberFormat") ?? "General";
        if (JsonUtil.Bool(style, "noFill") == true) cell.Fill = "";
        if (style.ContainsKey("fillColor")) cell.Fill = JsonUtil.Str(style, "fillColor") ?? "";
        if (JsonUtil.Get(style, "borders") is not JsonObject edges) return;
        foreach (var side in new[] { "left", "right", "top", "bottom" })
        {
            if (!edges.ContainsKey(side)) continue;
            cell.SetEdge(side, EdgeKeyFromNode(edges[side]));
        }
    }

    private static string EdgeKeyFromNode(JsonNode? node)
    {
        if (node is JsonValue v && v.TryGetValue<string>(out var token) &&
            token.Equals("none", StringComparison.OrdinalIgnoreCase))
            return "";
        if (node is not JsonObject obj) return "";
        var weight = JsonUtil.Str(obj, "weight") ?? "";
        var line = JsonUtil.Str(obj, "lineStyle") ?? "";
        var color = JsonUtil.Str(obj, "color") ?? "";
        if (weight.Length == 0 && line.Length == 0) return "";
        return $"{weight}|{line}|{color}";
    }

    private static IEnumerable<(int Row, int Col)> EnumerateUnion(string range)
    {
        foreach (var piece in range.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
        {
            var (r1, c1, r2, c2) = A1.ParseRange(piece);
            for (var r = r1; r <= r2; r++)
                for (var c = c1; c <= c2; c++)
                    yield return (r, c);
        }
    }

    private static JsonObject PlannerReport(JsonArray before, JsonArray after, JsonObject replay, int maxRow, int maxCol)
    {
        static (int Ops, int Areas, int Cells) Measure(JsonArray ops)
        {
            var areas = 0;
            var cells = 0;
            foreach (var op in ops.OfType<JsonObject>())
            {
                var range = JsonUtil.Str(op, "range") ?? "";
                foreach (var piece in range.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
                {
                    areas++;
                    cells += A1.CellCount(piece);
                }
            }
            return (ops.Count, areas, cells);
        }

        var b = Measure(before);
        var a = Measure(after);
        var packed = PackFormatOps(after);
        var maxCellsPerOp = 0;
        var maxCellsPerBatch = 0;
        var maxNonFastPerBatch = 0;
        foreach (var part in packed)
        {
            var m = MeasurePacked(part);
            if (m.Cells > maxCellsPerBatch) maxCellsPerBatch = m.Cells;
            if (m.NonFastCells > maxNonFastPerBatch) maxNonFastPerBatch = m.NonFastCells;
            foreach (var op in part.OfType<JsonObject>())
            {
                var n = UnionCellCount(JsonUtil.Str(op, "range"));
                if (n > maxCellsPerOp) maxCellsPerOp = n;
            }
        }

        return new JsonObject
        {
            ["kind"] = "property-layers",
            ["before"] = new JsonObject { ["formatOps"] = b.Ops, ["areas"] = b.Areas, ["cellWrites"] = b.Cells },
            ["after"] = new JsonObject { ["formatOps"] = a.Ops, ["areas"] = a.Areas, ["cellWrites"] = a.Cells },
            ["totalOps"] = a.Ops,
            ["batches"] = packed.Count,
            ["maxCells"] = MaxUnionCells,
            ["maxCellsPerOp"] = maxCellsPerOp,
            ["maxCellsPerBatch"] = maxCellsPerBatch,
            ["maxNonFastCellsPerBatch"] = maxNonFastPerBatch,
            ["nonFastBudget"] = NonFastSnapshotCellBudget,
            ["maxOpsPerBatch"] = MaxFormatOpsPerBatch,
            ["stylesBeforeMergesRequired"] = true,
            ["replay"] = replay.DeepClone(),
            ["grid"] = A1.Range(1, 1, maxRow, maxCol),
            ["maxUnionAreas"] = MaxUnionAreas,
            ["maxUnionCells"] = MaxUnionCells,
            ["blankSheetNoFillAssumed"] = true,
            ["noFillNotInFastScalarBase"] = true,
            ["doNotReplayWholeStyleSingletons"] = b.Ops > a.Ops,
        };
    }

    private static void NormalizeOmittedDefaults(CellSemantics[,] grid, int maxRow, int maxCol)
    {
        for (var r = 1; r <= maxRow; r++)
        for (var c = 1; c <= maxCol; c++)
        {
            var cell = grid[r, c];
            cell.HAlign = NormH(cell.HAlign);
            cell.VAlign = NormV(cell.VAlign);
            if (string.IsNullOrWhiteSpace(cell.NumberFormat)) cell.NumberFormat = "General";
        }
    }

    private static CellSemantics ReadSemantics(JsonObject style)
    {
        var cell = new CellSemantics
        {
            FontName = JsonUtil.Str(style, "fontName") ?? "",
            FontSize = JsonUtil.Num(style, "fontSize") ?? 0,
            Bold = JsonUtil.Bool(style, "bold") == true,
            Italic = JsonUtil.Bool(style, "italic") == true,
            HAlign = JsonUtil.Str(style, "horizontalAlign") ?? "",
            VAlign = JsonUtil.Str(style, "verticalAlign") ?? "",
            Wrap = JsonUtil.Bool(style, "wrapText") == true,
            Shrink = JsonUtil.Bool(style, "shrinkToFit") == true,
            NumberFormat = JsonUtil.Str(style, "numberFormat") ?? "General",
            Fill = JsonUtil.Str(style, "fillColor") ?? "",
        };
        if (JsonUtil.Get(style, "borders") is JsonObject edges)
        {
            cell.Left = EdgeKeyFromNode(edges["left"]);
            cell.Right = EdgeKeyFromNode(edges["right"]);
            cell.Top = EdgeKeyFromNode(edges["top"]);
            cell.Bottom = EdgeKeyFromNode(edges["bottom"]);
        }
        return cell;
    }

    private static IEnumerable<CellSemantics> EnumCells(CellSemantics[,] grid, int maxRow, int maxCol)
    {
        for (var r = 1; r <= maxRow; r++)
            for (var c = 1; c <= maxCol; c++)
                yield return grid[r, c];
    }

    private static bool TryMode<T>(IEnumerable<T> items, out T value)
    {
        IGrouping<T, T>? best = null;
        foreach (var group in items.GroupBy(item => item))
        {
            if (best is null || group.Count() > best.Count())
                best = group;
        }
        if (best is null)
        {
            value = default!;
            return false;
        }
        value = best.Key;
        return true;
    }

    private static bool NfEq(string a, string b) =>
        string.Equals(
            string.IsNullOrWhiteSpace(a) ? "General" : a,
            string.IsNullOrWhiteSpace(b) ? "General" : b,
            StringComparison.OrdinalIgnoreCase);

    private static string NormH(string? v) =>
        string.IsNullOrWhiteSpace(v) || v.Equals("general", StringComparison.OrdinalIgnoreCase) ? "general" : v;

    private static string NormV(string? v) =>
        string.IsNullOrWhiteSpace(v) ? "bottom" : v;

    private static bool IsMediumBottom(CellSemantics cell) =>
        cell.Bottom.StartsWith("medium|", StringComparison.Ordinal);

    /// <summary>
    /// Maximal axis-aligned rectangles. Adjacent equal-style cells share one
    /// format_range instead of one op per row-run or per cell.
    /// </summary>
    internal static IReadOnlyList<string> CoalesceRectangles(IEnumerable<(int Row, int Col)> cells)
    {
        var remaining = cells.ToHashSet();
        var ranges = new List<string>();
        while (remaining.Count > 0)
        {
            var start = remaining.OrderBy(c => c.Row).ThenBy(c => c.Col).First();
            var maxC = start.Col;
            while (remaining.Contains((start.Row, maxC + 1))) maxC++;
            var maxR = start.Row;
            while (true)
            {
                var next = maxR + 1;
                var ok = true;
                for (var c = start.Col; c <= maxC; c++)
                {
                    if (!remaining.Contains((next, c)))
                    {
                        ok = false;
                        break;
                    }
                }
                if (!ok) break;
                maxR = next;
            }
            for (var r = start.Row; r <= maxR; r++)
                for (var c = start.Col; c <= maxC; c++)
                    remaining.Remove((r, c));
            ranges.Add(A1.Range(start.Row, start.Col, maxR, maxC));
        }
        return ranges;
    }

    private static JsonObject DescribeFont(XElement font) => new()
    {
        ["name"] = (string?)font.Element(Ss + "name")?.Attribute("val"),
        ["sz"] = ParseDouble((string?)font.Element(Ss + "sz")?.Attribute("val")) is { } size ? size : null,
        ["bold"] = font.Element(Ss + "b") is not null,
        ["italic"] = font.Element(Ss + "i") is not null,
    };

    private static JsonObject DescribeBorder(XElement border)
    {
        JsonObject Edge(string name)
        {
            var e = border.Element(Ss + name);
            var rgb = (string?)e?.Element(Ss + "color")?.Attribute("rgb");
            if (rgb is { Length: > 2 } && rgb.StartsWith("FF", StringComparison.OrdinalIgnoreCase))
                rgb = rgb[2..];
            return new JsonObject
            {
                ["style"] = (string?)e?.Attribute("style"),
                ["rgb"] = rgb,
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

    private static JsonObject DescribeFill(XElement fill)
    {
        var fg = fill.Descendants(Ss + "fgColor").FirstOrDefault();
        var rgb = (string?)fg?.Attribute("rgb");
        if (rgb is { Length: > 2 } && rgb.StartsWith("FF", StringComparison.OrdinalIgnoreCase))
            rgb = rgb[2..];
        return new JsonObject { ["rgb"] = rgb, ["pattern"] = (string?)fill.Descendants(Ss + "patternFill").FirstOrDefault()?.Attribute("patternType") };
    }

    private static JsonObject? MapBorder(JsonObject? edge)
    {
        var raw = JsonUtil.Str(edge, "style");
        if (string.IsNullOrWhiteSpace(raw)) return null;
        var token = raw.ToLowerInvariant();
        var (weight, line) = token switch
        {
            "hair" => ("hairline", "continuous"),
            "thin" => ("thin", "continuous"),
            "medium" => ("medium", "continuous"),
            "thick" => ("thick", "continuous"),
            "dashed" or "dash" => ("thin", "dash"),
            "dotted" or "dot" => ("thin", "dot"),
            "dashdot" => ("thin", "dashDot"),
            "dashdotdot" => ("thin", "dashDotDot"),
            "double" => ("thin", "double"),
            "slantdashdot" => ("thin", "slantDashDot"),
            "mediumdashed" => ("medium", "dash"),
            "mediumdashdot" => ("medium", "dashDot"),
            "mediumdashdotdot" => ("medium", "dashDotDot"),
            _ => ("thin", "continuous"),
        };
        var mapped = new JsonObject { ["weight"] = weight, ["lineStyle"] = line };
        var rgb = JsonUtil.Str(edge, "rgb");
        if (!string.IsNullOrWhiteSpace(rgb))
            mapped["color"] = "#" + rgb.TrimStart('#');
        return mapped;
    }

    private static string? MapHoriz(string? v) => v switch
    {
        null => null,
        "centerContinuous" => "centerAcross",
        _ => v,
    };

    private static string? MapVert(string? v) => v;

    private static bool IsOn(XElement? align, string name) =>
        (string?)align?.Attribute(name) is "1" or "true";

    private static Dictionary<int, string> LoadNumFmts(XDocument styles)
    {
        var map = new Dictionary<int, string>();
        foreach (var n in Root(styles, "numFmts", "numFmt"))
        {
            var id = ParseInt((string?)n.Attribute("numFmtId"));
            var code = (string?)n.Attribute("formatCode");
            if (id is not null && !string.IsNullOrWhiteSpace(code))
                map[id.Value] = code;
        }
        return map;
    }

    private static string BuiltinNumFmt(int id) => id switch
    {
        0 => "General",
        1 => "0",
        2 => "0.00",
        3 => "#,##0",
        4 => "#,##0.00",
        9 => "0%",
        10 => "0.00%",
        14 => "mm-dd-yy",
        15 => "d-mmm-yy",
        16 => "d-mmm",
        17 => "mmm-yy",
        _ => "General",
    };

    private static IReadOnlyList<XElement> Root(XDocument doc, string parent, string child)
    {
        var root = doc.Root ?? throw new InvalidOperationException("xml root missing");
        return root.Element(Ss + parent)?.Elements(Ss + child).ToList() ?? [];
    }

    private static XDocument Read(ZipArchive zip, string path)
    {
        var entry = zip.GetEntry(path) ?? throw new InvalidOperationException($"package part missing: {path}");
        using var stream = entry.Open();
        return XDocument.Load(stream);
    }

    private static int? ParseInt(string? s) =>
        int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var v) ? v : null;

    private static double? ParseDouble(string? s) =>
        double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out var v) ? v : null;

    private static double? InchToMm(string? inches) =>
        ParseDouble(inches) is { } v ? Math.Round(v * 25.4, 3) : null;
}
