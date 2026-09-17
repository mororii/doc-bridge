namespace DocBridge.Development.ExcelProductionWorkflowProbe;

internal readonly record struct LayoutBox(string Name, double Left, double Top, double Width, double Height)
{
    public double Right => Left + Width;
    public double Bottom => Top + Height;

    public bool Overlaps(LayoutBox other, double eps = 0.01) =>
        Left < other.Right - eps && Right > other.Left + eps &&
        Top < other.Bottom - eps && Bottom > other.Top + eps;

    public JsonObject ToJson() => new()
    {
        ["name"] = Name,
        ["left"] = Left,
        ["top"] = Top,
        ["width"] = Width,
        ["height"] = Height,
        ["right"] = Right,
        ["bottom"] = Bottom,
    };
}

/// <summary>
/// Point/character layout for practical E3/E6 fixtures.
/// Object boxes are checked against populated cell/text boxes at MDW 6/7/8.
/// Photo/chart/signature tops are derived from cumulative row heights.
/// </summary>
internal static class PracticalLayoutMetrics
{
    public const double Dpi = 96;
    public const double MmPerInch = 25.4;
    public const double A4WidthMm = 210;
    public const double A4HeightMm = 297;
    public const double DefaultMarginMm = 12;
    public const double ChartHeightPt = 180;
    public const double ChartInsetLeftPt = 12;
    public const double ChartWidthPt = 340;
    public const double ChartBandPadPt = 6;
    public const double PhotoWidthPt = 160;
    public const double PhotoHeightPt = 120;
    public const double Photo1LeftPt = 16;
    public const double PhotoGapPt = 12;
    public const double PhotoBandPadPt = 8;
    public const double SampleMarkWidthPt = 56;
    public const double SampleMarkHeightPt = 28;
    public const int PageScale = 100;

    public static readonly double[] E6ColumnChars = [22, 20, 20, 20];
    public static readonly double[] E3ColumnChars = [14, 16, 14, 12, 14, 16];
    public static readonly double[] DigitWidthsPx = [6, 7, 8];

    public static double A4WidthPt => MmToPoints(A4WidthMm);
    public static double A4HeightPt => MmToPoints(A4HeightMm);
    public static double Photo2LeftPt => Photo1LeftPt + PhotoWidthPt + PhotoGapPt;

    public static double MmToPoints(double mm) => mm * 72.0 / MmPerInch;

    public static double PrintableWidthPt(double marginMm = DefaultMarginMm) =>
        A4WidthPt - 2 * MmToPoints(marginMm);

    public static double PrintableHeightPt(double marginMm = DefaultMarginMm) =>
        A4HeightPt - 2 * MmToPoints(marginMm);

    public static double WidthCharsToPoints(double widthChars, double maxDigitWidthPx = 7)
    {
        var pixels = Math.Truncate(((256.0 * widthChars + Math.Truncate(128.0 / maxDigitWidthPx)) / 256.0) * maxDigitWidthPx);
        return pixels * 72.0 / Dpi;
    }

    public static double[] ColumnLefts(IReadOnlyList<double> widthChars, double maxDigitWidthPx)
    {
        var lefts = new double[widthChars.Count + 1];
        for (var i = 0; i < widthChars.Count; i++)
            lefts[i + 1] = lefts[i] + WidthCharsToPoints(widthChars[i], maxDigitWidthPx);
        return lefts;
    }

    public static double[] CumulativeTops(IReadOnlyList<double> heights)
    {
        var tops = new double[heights.Count + 1];
        for (var i = 0; i < heights.Count; i++)
            tops[i + 1] = tops[i] + heights[i];
        return tops;
    }

    public static bool IsPointStep(double points) =>
        Math.Abs(points / 0.6 - Math.Round(points / 0.6)) < 1e-9;

    public static LayoutBox BoxFromRange(string name, string range, IReadOnlyList<double> heights, IReadOnlyList<double> widthChars, double mdw)
    {
        var (r1, c1, r2, c2) = A1.ParseRange(range);
        var tops = CumulativeTops(heights);
        var lefts = ColumnLefts(widthChars, mdw);
        if (r2 > heights.Count || c2 > widthChars.Count)
            throw new InvalidOperationException($"{name} {range} exceeds modeled grid");
        return new LayoutBox(name, lefts[c1 - 1], tops[r1 - 1], lefts[c2] - lefts[c1 - 1], tops[r2] - tops[r1 - 1]);
    }

    public static JsonObject ColumnSandwich(IReadOnlyList<double> widthChars)
    {
        var widths = new JsonObject();
        var min = double.MaxValue;
        var max = 0d;
        foreach (var mdw in DigitWidthsPx)
        {
            var pts = ColumnLefts(widthChars, mdw)[^1];
            widths[mdw.ToString(CultureInfo.InvariantCulture)] = pts;
            min = Math.Min(min, pts);
            max = Math.Max(max, pts);
        }
        return new JsonObject
        {
            ["byMaxDigitWidthPx"] = widths,
            ["narrowestPoints"] = min,
            ["widestPoints"] = max,
            ["printableWidthPoints"] = PrintableWidthPt(),
            ["fitsPrintableAtScale100"] = max <= PrintableWidthPt() + 0.01,
        };
    }

    public static double[] E6RowHeights()
    {
        // 45 rows. Image lives on empty row 24. Analysis starts at 36.
        var h = new double[45];
        Array.Fill(h, 18);
        h[0] = 30;
        h[23] = 30; // empty reserved image row 24
        h[35] = 30; // 분석
        h[36] = 42; // 요지 — longest wrap
        h[37] = 36; // 원인
        h[38] = 36; // 리스크
        h[39] = 36; // 조치
        h[40] = 30; // 담당
        for (var i = 42; i <= 44; i++) h[i] = 24; // signatures 43-45
        return h;
    }

    public static double[] E3RowHeights()
    {
        var h = new double[23];
        Array.Fill(h, 18);
        h[0] = 30;
        h[1] = 30; // wrapped metadata
        h[2] = 42; // D3 station wrap
        h[8] = 168; // long note stress
        for (var i = 20; i <= 22; i++) h[i] = 24;
        return h;
    }

    public static readonly string[] E6TextRanges =
    [
        "A1:D1", "A2:D2", "A3:D8", "A9:D9",
        "A22:C22", "A23", "B23:C23",
        "A36:D42", "A43:B45", "C43:D45",
    ];

    public static readonly string[] E6MergeRanges =
    [
        "A1:D1", "B9:D9", "A22:C22", "B23:C23",
        "B36:D36", "B37:D37", "B38:D38", "B39:D39", "B40:D40", "B41:D41",
        "A43:B45", "C43:D45",
    ];

    public static readonly string[] E3TextRanges =
    [
        "A1:F1", "A2:F3", "A4:F8", "A9:F9", "A10:F10", "A21:F23",
    ];

    public static LayoutBox[] E6Objects()
    {
        var tops = CumulativeTops(E6RowHeights());
        return
        [
            new("PlanActualLine", ChartInsetLeftPt, tops[9] + ChartBandPadPt, ChartWidthPt, ChartHeightPt),
            new("DiffColumn", ChartInsetLeftPt, tops[24] + ChartBandPadPt, ChartWidthPt, ChartHeightPt),
            new("Logo", ChartInsetLeftPt, tops[23] + 1, SampleMarkWidthPt, SampleMarkHeightPt),
        ];
    }

    public static LayoutBox[] E3Objects()
    {
        var tops = CumulativeTops(E3RowHeights());
        var photoTop = tops[10] + PhotoBandPadPt;
        return
        [
            new("Photo1", Photo1LeftPt, photoTop, PhotoWidthPt, PhotoHeightPt),
            new("Photo2", Photo2LeftPt, photoTop, PhotoWidthPt, PhotoHeightPt),
        ];
    }

    public static JsonObject Position(LayoutBox box) => new()
    {
        ["left"] = box.Left,
        ["top"] = box.Top,
        ["width"] = box.Width,
        ["height"] = box.Height,
    };

    public static JsonObject E6Bounds()
    {
        var heights = E6RowHeights();
        var tops = CumulativeTops(heights);
        var sandwich = ColumnSandwich(E6ColumnChars);
        var objects = E6Objects();
        return new JsonObject
        {
            ["columnChars"] = CharsJson(E6ColumnChars),
            ["columnSandwich"] = sandwich,
            ["scale"] = PageScale,
            ["chart1"] = Position(objects[0]),
            ["chart2"] = Position(objects[1]),
            ["sampleMark"] = Position(objects[2]),
            ["row22Top"] = tops[21],
            ["row23Top"] = tops[22],
            ["row24Top"] = tops[23],
            ["analysisRow36Top"] = tops[35],
            ["page1ContentBottom"] = tops[21],
            ["page2ContentBottom"] = tops[45],
            ["page2WithRepeatedTitles"] = tops[45] - tops[21] + heights[0] + heights[1],
            ["printArea"] = "A1:D45",
            ["pageBreak"] = "A22",
            ["printTitleRows"] = "$1:$2",
            ["imageRow"] = 24,
            ["reservedImageCells"] = "D22:D23 empty; image on empty row 24",
            ["proof"] = "object boxes vs populated text boxes at MDW 6/7/8; tops from cumulative row heights",
        };
    }

    public static JsonObject E3Bounds()
    {
        var heights = E3RowHeights();
        var tops = CumulativeTops(heights);
        var objects = E3Objects();
        return new JsonObject
        {
            ["columnChars"] = CharsJson(E3ColumnChars),
            ["columnSandwich"] = ColumnSandwich(E3ColumnChars),
            ["scale"] = PageScale,
            ["photo1"] = Position(objects[0]),
            ["photo2"] = Position(objects[1]),
            ["noteRowTop"] = tops[8],
            ["noteRowBottom"] = tops[9],
            ["photoBandTop"] = tops[10],
            ["photoTop"] = objects[0].Top,
            ["photoBandBottom"] = tops[20],
            ["signatureTop"] = tops[20],
            ["contentBottom"] = tops[23],
            ["printArea"] = "A1:F23",
            ["row2Height"] = heights[1],
            ["row3Height"] = heights[2],
            ["proof"] = "photo/signature tops from cumulative row heights; D3 wrap uses row3=42",
        };
    }

    public static JsonObject SelfCheck()
    {
        var errors = new JsonArray();
        var overlaps = new JsonArray();
        CheckGrid(errors, "e6", E6RowHeights(), E6ColumnChars, E6TextRanges, E6Objects(), overlaps, pageBreakBeforeRow: 22);
        CheckGrid(errors, "e3", E3RowHeights(), E3ColumnChars, E3TextRanges, E3Objects(), overlaps, pageBreakBeforeRow: null);

        var e3Tops = CumulativeTops(E3RowHeights());
        var photos = E3Objects();
        if (Math.Abs(photos[0].Top - (e3Tops[10] + PhotoBandPadPt)) > 0.01)
            errors.Add("E3 photo top must be derived from row 11, not a hardcoded 350");
        if (photos[0].Top < e3Tops[9] - 0.01)
            errors.Add("E3 photos overlap the long note");
        if (photos[0].Bottom > e3Tops[20] + 0.01)
            errors.Add("E3 photos leave the reserved photo band");
        if (E3RowHeights()[1] < 30 || E3RowHeights()[2] < 42)
            errors.Add("E3 rows 2/3 must be at least 30/42 for wrapped metadata");
        if (E6RowHeights()[35] < 30 || E6RowHeights()[39] < 36 || E6RowHeights()[40] < 30)
            errors.Add("E6 analysis rows 36/40/41 must have 30+/36+/30 heights");

        foreach (var h in E6RowHeights().Concat(E3RowHeights()))
        {
            if (!IsPointStep(h))
                errors.Add($"row height {h} is not a 0.6pt step");
        }

        if (Math.Abs(AcceptanceNumbers.ActualRates[2] - AcceptanceNumbers.PlanRates[2] - 3) > 1e-9)
            errors.Add("D5 initial must remain 3");
        if (Math.Abs(AcceptanceNumbers.MarDiffChanged - 5) > 1e-9)
            errors.Add("D5 after Mar change must remain 5");

        return new JsonObject
        {
            ["ok"] = errors.Count == 0,
            ["errors"] = errors,
            ["overlaps"] = overlaps,
            ["e3"] = E3Bounds(),
            ["e6"] = E6Bounds(),
        };
    }

    private static void CheckGrid(
        JsonArray errors,
        string id,
        double[] heights,
        double[] cols,
        IReadOnlyList<string> textRanges,
        LayoutBox[] objects,
        JsonArray overlaps,
        int? pageBreakBeforeRow)
    {
        var sandwich = ColumnSandwich(cols);
        var narrowest = (double)sandwich["narrowestPoints"]!;
        var widest = (double)sandwich["widestPoints"]!;
        var tops = CumulativeTops(heights);
        if (widest > PrintableWidthPt() + 0.01)
            errors.Add($"{id} widest columns {widest} exceed printable {PrintableWidthPt()} at scale 100");
        if (pageBreakBeforeRow is { } br)
        {
            var breakY = tops[br - 1];
            if (breakY > PrintableHeightPt())
                errors.Add($"{id} page 1 through row {br - 1} exceeds printable height");
            var page2 = tops[^1] - breakY + heights[0] + heights[1];
            if (page2 > PrintableHeightPt())
                errors.Add($"{id} page 2 + repeated titles {page2} exceeds printable height");
            foreach (var obj in objects)
            {
                if (obj.Top < breakY - 0.01 && obj.Bottom > breakY + 0.01)
                    errors.Add($"{id} {obj.Name} crosses page break at {breakY}");
                else if (obj.Bottom <= breakY + 0.01)
                {
                    if (obj.Bottom > PrintableHeightPt() + 0.01)
                        errors.Add($"{id} {obj.Name} leaves page 1 printable height");
                }
                else
                {
                    var page2Bottom = obj.Bottom - breakY + heights[0] + heights[1];
                    if (page2Bottom > PrintableHeightPt() + 0.01)
                        errors.Add($"{id} {obj.Name} leaves page 2 printable height with repeated titles");
                }
            }
        }
        else if (tops[^1] > PrintableHeightPt())
        {
            errors.Add($"{id} content {tops[^1]} exceeds one A4 page");
        }

        foreach (var obj in objects)
        {
            if (obj.Right > narrowest + 0.01)
                errors.Add($"{id} {obj.Name} right {obj.Right} exceeds narrowest columns {narrowest}");
            if (obj.Left < -0.01 || obj.Top < -0.01)
                errors.Add($"{id} {obj.Name} has negative origin");
        }

        for (var i = 0; i < objects.Length; i++)
        {
            for (var j = i + 1; j < objects.Length; j++)
            {
                if (!objects[i].Overlaps(objects[j])) continue;
                errors.Add($"{id} {objects[i].Name} overlaps {objects[j].Name}");
                overlaps.Add(new JsonObject { ["a"] = objects[i].ToJson(), ["b"] = objects[j].ToJson() });
            }
        }

        foreach (var mdw in DigitWidthsPx)
        {
            var texts = textRanges.Select(r => BoxFromRange(r, r, heights, cols, mdw)).ToArray();
            foreach (var obj in objects)
            {
                foreach (var text in texts)
                {
                    if (!obj.Overlaps(text)) continue;
                    errors.Add($"{id} {obj.Name} overlaps text {text.Name} at MDW={mdw}");
                    overlaps.Add(new JsonObject
                    {
                        ["mdw"] = mdw,
                        ["object"] = obj.ToJson(),
                        ["text"] = text.ToJson(),
                    });
                }
            }
        }
    }

    private static JsonArray CharsJson(IReadOnlyList<double> chars) =>
        new(chars.Select(v => JsonValue.Create(v)).ToArray());
}
