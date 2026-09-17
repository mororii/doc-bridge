using System.Globalization;
using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// Excel set_page_setup contract. <c>scale</c> wins over fit-to-page.
/// Margins are authored in millimetres and stored as Excel points (72/inch).
/// </summary>
public static class ExcelPageSetupContract
{
    public const int XlPaperA3 = 8;
    public const int XlPaperA4 = 9;
    public const int XlPaperLetter = 1;
    public const int XlPortrait = 1;
    public const int XlLandscape = 2;
    public const string ModeScale = "scale";
    public const string ModeFit = "fit";

    public static readonly IReadOnlyDictionary<string, int> PaperSizes =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["letter"] = XlPaperLetter,
            ["a3"] = XlPaperA3,
            ["a4"] = XlPaperA4,
            ["tabloid"] = 3,
            ["legal"] = 5,
            ["b4"] = 12,
            ["b5"] = 13,
        };

    public static bool TryNormalize(JsonObject? page, int opIndex, ICollection<string> errors, out JsonObject canonical)
    {
        canonical = new JsonObject();
        if (page is null)
        {
            errors.Add($"ops[{opIndex}] 'set_page_setup' field 'page' must be an object");
            return false;
        }

        var ok = true;
        if (page.ContainsKey("paperSize"))
        {
            int paper;
            if (!TryReadPaper(page["paperSize"], opIndex, errors, out paper))
                ok = false;
            else
                canonical["paperSize"] = paper;
        }

        if (page.ContainsKey("orientation"))
        {
            var orientation = Json.GetString(page, "orientation");
            if (!string.Equals(orientation, "portrait", StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(orientation, "landscape", StringComparison.OrdinalIgnoreCase))
            {
                errors.Add($"ops[{opIndex}] 'set_page_setup' page.orientation must be portrait or landscape");
                ok = false;
            }
            else
            {
                canonical["orientation"] = orientation!.ToLowerInvariant();
                canonical["orientationValue"] = orientation.Equals("landscape", StringComparison.OrdinalIgnoreCase)
                    ? XlLandscape
                    : XlPortrait;
            }
        }

        var hasScale = page.ContainsKey("scale");
        var hasFit = page.ContainsKey("fitToWidth") || page.ContainsKey("fitToHeight");
        if (hasScale)
        {
            if (!TryReadInt(page["scale"], 10, 400, out var scale))
            {
                errors.Add($"ops[{opIndex}] 'set_page_setup' page.scale must be an integer 10..400");
                ok = false;
            }
            else
            {
                canonical["scale"] = scale;
                canonical["zoomMode"] = ModeScale;
                if (hasFit)
                    canonical["fitIgnored"] = true;
            }
        }
        else if (hasFit)
        {
            if (page.ContainsKey("fitToWidth") && !TryReadInt(page["fitToWidth"], 0, 32767, out var fitW))
            {
                errors.Add($"ops[{opIndex}] 'set_page_setup' page.fitToWidth must be an integer 0..32767");
                ok = false;
            }
            else if (page.ContainsKey("fitToWidth"))
                canonical["fitToWidth"] = Json.GetInt(page, "fitToWidth");

            if (page.ContainsKey("fitToHeight") && !TryReadInt(page["fitToHeight"], 0, 32767, out var fitH))
            {
                errors.Add($"ops[{opIndex}] 'set_page_setup' page.fitToHeight must be an integer 0..32767");
                ok = false;
            }
            else if (page.ContainsKey("fitToHeight"))
                canonical["fitToHeight"] = Json.GetInt(page, "fitToHeight");

            canonical["zoomMode"] = ModeFit;
        }

        CopyString(page, canonical, "printArea");
        CopyTitleRange(page, canonical, "printTitleRows", columns: false);
        CopyTitleRange(page, canonical, "printTitleColumns", columns: true);
        CopyString(page, canonical, "leftHeader");
        CopyString(page, canonical, "centerHeader");
        CopyString(page, canonical, "rightHeader");
        CopyString(page, canonical, "leftFooter");
        CopyString(page, canonical, "centerFooter");
        CopyString(page, canonical, "rightFooter");

        foreach (var key in new[]
                 {
                     "leftMarginMm", "rightMarginMm", "topMarginMm", "bottomMarginMm",
                     "headerMarginMm", "footerMarginMm",
                 })
        {
            if (!page.ContainsKey(key)) continue;
            if (!TryReadNumber(page[key], 0, 200, out var mm))
            {
                errors.Add($"ops[{opIndex}] 'set_page_setup' page.{key} must be a finite number 0..200");
                ok = false;
                continue;
            }

            canonical[key] = mm;
            canonical[key.Replace("Mm", "Points", StringComparison.Ordinal)] = MillimetresToPoints(mm);
        }

        if (canonical.Count == 0)
        {
            errors.Add($"ops[{opIndex}] 'set_page_setup' page does not name any supported properties");
            return false;
        }

        EnsurePointKeys(canonical);
        return ok;
    }

    /// <summary>
    /// Captured page setups historically stored millimetres only. Writers must
    /// accept those keys and materialize Excel point values before COM assign.
    /// </summary>
    public static void EnsurePointKeys(JsonObject page)
    {
        foreach (var key in new[]
                 {
                     "leftMarginMm", "rightMarginMm", "topMarginMm", "bottomMarginMm",
                     "headerMarginMm", "footerMarginMm",
                 })
        {
            var pointsKey = key.Replace("Mm", "Points", StringComparison.Ordinal);
            if (page.ContainsKey(pointsKey) || !page.ContainsKey(key)) continue;
            if (TryReadNumber(page[key], 0, 200, out var mm))
                page[pointsKey] = MillimetresToPoints(mm);
        }
    }

    public static string NormalizePrintRange(string? value) =>
        (value ?? "").Replace("$", "", StringComparison.Ordinal).Trim();

    /// <summary>
    /// Excel resolves a relative title range ("7:7", "A:B") against the active
    /// cell, which silently retargets the titles (observed: "7:7" became
    /// $22:$22). Canonicalize to absolute form ("$7:$7") before COM assign so
    /// the written titles never depend on selection state.
    /// </summary>
    public static string AbsolutizeTitleRange(string? value, bool columns)
    {
        var text = (value ?? "").Replace("$", "", StringComparison.Ordinal).Trim();
        var parts = text.Split(':');
        if (parts.Length != 2) return text;
        var left = parts[0].Trim();
        var right = parts[1].Trim();
        if (columns)
        {
            if (ExcelA1Box.ColumnIndex(left) is < 1 or > 16_384) return text;
            if (ExcelA1Box.ColumnIndex(right) is < 1 or > 16_384) return text;
            return $"${left.ToUpperInvariant()}:${right.ToUpperInvariant()}";
        }
        if (!int.TryParse(left, NumberStyles.None, CultureInfo.InvariantCulture, out var first) ||
            !int.TryParse(right, NumberStyles.None, CultureInfo.InvariantCulture, out var last))
            return text;
        if (first is < 1 or > 1_048_576 || last is < 1 or > 1_048_576 || last < first) return text;
        return $"${first}:${last}";
    }

    private static void CopyTitleRange(JsonObject page, JsonObject canonical, string key, bool columns)
    {
        var text = Json.GetString(page, key);
        if (string.IsNullOrWhiteSpace(text)) return;
        canonical[key] = AbsolutizeTitleRange(text, columns);
    }

    public static bool PrintRangesMatch(string? expected, string? actual) =>
        string.Equals(NormalizePrintRange(expected), NormalizePrintRange(actual), StringComparison.OrdinalIgnoreCase);

    public static bool HeaderFooterMatch(string? expected, string? actual) =>
        string.Equals(expected ?? "", actual ?? "", StringComparison.Ordinal);

    public static bool MatchesRequested(JsonObject actual, JsonObject expected) =>
        DescribeMismatches(actual, expected).Count == 0;

    public static IReadOnlyList<string> DescribeMismatches(JsonObject actual, JsonObject expected)
    {
        var mismatches = new List<string>();
        if (actual.ContainsKey("error"))
        {
            mismatches.Add($"pageSetup error: {Json.GetString(actual, "error")}");
            return mismatches;
        }

        AddIfDifferent(mismatches, "paperSize", Json.GetInt(expected, "paperSize"), Json.GetInt(actual, "paperSize"),
            expected.ContainsKey("paperSize"));
        if (expected.ContainsKey("orientation") &&
            !string.Equals(Json.GetString(actual, "orientation"), Json.GetString(expected, "orientation"), StringComparison.OrdinalIgnoreCase))
            mismatches.Add(FieldMismatch("orientation", Json.GetString(expected, "orientation"), Json.GetString(actual, "orientation")));
        if (string.Equals(Json.GetString(expected, "zoomMode"), ModeScale, StringComparison.Ordinal))
            AddIfDifferent(mismatches, "scale", Json.GetInt(expected, "scale"), Json.GetInt(actual, "scale"), true);
        if (string.Equals(Json.GetString(expected, "zoomMode"), ModeFit, StringComparison.Ordinal))
        {
            AddIfDifferent(mismatches, "fitToWidth", Json.GetInt(expected, "fitToWidth"), Json.GetInt(actual, "fitToWidth"),
                expected.ContainsKey("fitToWidth"));
            AddIfDifferent(mismatches, "fitToHeight", Json.GetInt(expected, "fitToHeight"), Json.GetInt(actual, "fitToHeight"),
                expected.ContainsKey("fitToHeight"));
        }

        if (expected.ContainsKey("printArea") &&
            !PrintRangesMatch(Json.GetString(expected, "printArea"), Json.GetString(actual, "printArea")))
            mismatches.Add(FieldMismatch("printArea", Json.GetString(expected, "printArea"), Json.GetString(actual, "printArea")));

        foreach (var key in new[] { "printTitleRows", "printTitleColumns" })
        {
            if (!expected.ContainsKey(key)) continue;
            if (!PrintRangesMatch(Json.GetString(expected, key), Json.GetString(actual, key)))
                mismatches.Add(FieldMismatch(key, Json.GetString(expected, key), Json.GetString(actual, key)));
        }

        foreach (var key in new[]
                 {
                     "leftHeader", "centerHeader", "rightHeader",
                     "leftFooter", "centerFooter", "rightFooter",
                 })
        {
            if (!expected.ContainsKey(key)) continue;
            if (!HeaderFooterMatch(Json.GetString(expected, key), Json.GetString(actual, key)))
                mismatches.Add(FieldMismatch(key, Json.GetString(expected, key), Json.GetString(actual, key)));
        }

        foreach (var key in new[]
                 {
                     "leftMarginMm", "rightMarginMm", "topMarginMm", "bottomMarginMm",
                     "headerMarginMm", "footerMarginMm",
                 })
        {
            if (!expected.ContainsKey(key) && !expected.ContainsKey(key.Replace("Mm", "Points", StringComparison.Ordinal)))
                continue;
            if (!TryReadNumber(expected.ContainsKey(key) ? expected[key] : null, 0, 200, out var expectedMm))
            {
                var pointsKey = key.Replace("Mm", "Points", StringComparison.Ordinal);
                if (expected.ContainsKey(pointsKey) && TryReadNumber(expected[pointsKey], 0, 2000, out var expectedPoints))
                    expectedMm = PointsToMillimetres(expectedPoints);
                else
                    continue;
            }

            if (!TryReadNumber(actual[key], 0, 200, out var actualMm))
            {
                mismatches.Add(FieldMismatch(key, expectedMm.ToString(CultureInfo.InvariantCulture), "unreadable"));
                continue;
            }

            if (Math.Abs(actualMm - expectedMm) > 0.15)
                mismatches.Add(FieldMismatch(
                    key,
                    expectedMm.ToString("0.###", CultureInfo.InvariantCulture),
                    actualMm.ToString("0.###", CultureInfo.InvariantCulture)));
        }

        return mismatches;
    }

    private static void AddIfDifferent(List<string> mismatches, string field, int? expected, int? actual, bool present)
    {
        if (!present) return;
        if (expected != actual)
            mismatches.Add(FieldMismatch(field, expected?.ToString(CultureInfo.InvariantCulture), actual?.ToString(CultureInfo.InvariantCulture)));
    }

    private static string FieldMismatch(string field, string? expected, string? actual) =>
        $"{field} expected '{expected ?? ""}' actual '{actual ?? ""}'";

    public static string PaperName(int value) => value switch
    {
        XlPaperLetter => "letter",
        XlPaperA3 => "A3",
        XlPaperA4 => "A4",
        3 => "tabloid",
        5 => "legal",
        12 => "B4",
        13 => "B5",
        _ => $"paper({value})",
    };

    public static double MillimetresToPoints(double mm) => mm * 72d / 25.4d;
    public static double PointsToMillimetres(double points) => points * 25.4d / 72d;

    private static bool TryReadPaper(JsonNode? node, int opIndex, ICollection<string> errors, out int paper)
    {
        paper = 0;
        if (node is JsonValue value && value.TryGetValue<int>(out paper) && paper is >= 1 and <= 256)
            return true;
        if (node is JsonValue text && text.TryGetValue<string>(out var name) &&
            PaperSizes.TryGetValue(name, out paper))
            return true;
        errors.Add($"ops[{opIndex}] 'set_page_setup' page.paperSize must be A3|A4|letter|legal|tabloid|B4|B5 or Excel paper integer");
        return false;
    }

    private static bool TryReadInt(JsonNode? node, int min, int max, out int number)
    {
        number = 0;
        if (node is not JsonValue value) return false;
        if (value.TryGetValue<int>(out number)) return number >= min && number <= max;
        if (value.TryGetValue<long>(out var l) && l >= min && l <= max)
        {
            number = (int)l;
            return true;
        }

        return false;
    }

    private static bool TryReadNumber(JsonNode? node, double min, double max, out double number)
    {
        number = 0;
        if (node is not JsonValue value) return false;
        if (value.TryGetValue<double>(out number) && double.IsFinite(number))
            return number >= min && number <= max;
        if (value.TryGetValue<int>(out var i))
        {
            number = i;
            return number >= min && number <= max;
        }

        return false;
    }

    private static void CopyString(JsonObject source, JsonObject dest, string key)
    {
        var text = Json.GetString(source, key);
        if (text is not null) dest[key] = text;
    }
}
