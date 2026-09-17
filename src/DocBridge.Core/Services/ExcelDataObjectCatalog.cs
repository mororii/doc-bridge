using System.Globalization;

namespace DocBridge.Core.Services;

/// <summary>
/// Excel COM enumerations and public token maps for data/reporting ops.
/// Values are the official Xl* / Mso* constants; unknown tokens are rejected.
/// </summary>
public static class ExcelDataObjectCatalog
{
    public const int XlSrcRange = 1;
    public const int XlYes = 1;
    public const int XlNo = 2;
    public const int XlGuess = 0;
    public const int XlAscending = 1;
    public const int XlDescending = 2;
    public const int XlSortColumns = 1;
    public const int XlSortRows = 2;
    public const int XlAnd = 1;
    public const int XlOr = 2;
    public const int XlFilterValues = 7;
    public const int XlValidateWholeNumber = 1;
    public const int XlValidateDecimal = 2;
    public const int XlValidateList = 3;
    public const int XlValidateDate = 4;
    public const int XlValidateTime = 5;
    public const int XlValidateTextLength = 6;
    public const int XlValidateCustom = 7;
    public const int XlBetween = 1;
    public const int XlNotBetween = 2;
    public const int XlEqual = 3;
    public const int XlNotEqual = 4;
    public const int XlGreater = 5;
    public const int XlLess = 6;
    public const int XlGreaterEqual = 7;
    public const int XlLessEqual = 8;
    public const int XlValidAlertStop = 1;
    public const int XlValidAlertWarning = 2;
    public const int XlValidAlertInformation = 3;
    public const int XlCellValue = 1;
    public const int XlExpression = 2;
    public const int XlColorScale = 3;
    public const int XlDatabar = 4;
    public const int XlUniqueValues = 8;
    public const int XlDuplicate = 1;
    public const int XlUnique = 0;
    public const int XlConditionValueLowestValue = 1;
    public const int XlConditionValueHighestValue = 2;
    public const int XlColumnClustered = 51;
    public const int XlColumnStacked = 52;
    public const int XlBarClustered = 57;
    public const int XlBarStacked = 58;
    public const int XlLine = 4;
    public const int XlLineMarkers = 65;
    public const int XlPie = 5;
    public const int XlArea = 1;
    public const int XlXYScatter = -4169;
    public const int XlXYScatterLines = 74;
    public const int XlRows = 1;
    public const int XlColumns = 2;
    public const int XlLegendPositionBottom = -4107;
    public const int XlLegendPositionCorner = 2;
    public const int XlLegendPositionTop = -4160;
    public const int XlLegendPositionRight = -4152;
    public const int XlLegendPositionLeft = -4131;
    public const int XlTotalsCalculationNone = 0;
    public const int XlTotalsCalculationAverage = 1;
    public const int XlTotalsCalculationCount = 2;
    public const int XlTotalsCalculationCountNums = 3;
    public const int XlTotalsCalculationMax = 4;
    public const int XlTotalsCalculationMin = 5;
    public const int XlTotalsCalculationSum = 6;
    public const int XlTotalsCalculationStdDev = 7;
    public const int XlTotalsCalculationVar = 8;
    public const int XlTotalsCalculationCustom = 9;
    public const int MsoFalse = 0;
    public const int MsoTrue = -1;
    public const int MsoPicture = 13;
    public const int XlErrRef = 2023;
    public const int XlDatabase = 1;
    public const int XlRowField = 1;
    public const int XlColumnField = 2;
    public const int XlPageField = 3;
    public const int XlDataField = 4;
    public const int XlHidden = 0;
    public const int XlSum = -4157;
    public const int XlAverage = -4106;
    public const int XlCount = -4112;
    public const int XlCountNums = -4113;
    public const int XlMax = -4136;
    public const int XlMin = -4139;
    public const int XlProduct = -4149;
    public const int XlStDev = -4155;
    public const int XlVar = -4164;
    public const int XlDelimited = 1;
    public const int XlFixedWidth = 2;
    public const int XlTextQualifierDoubleQuote = 1;
    public const int XlTextQualifierSingleQuote = 2;
    public const int XlTextQualifierNone = -4142;
    public const int XlIconSet = 6;
    public const int Xl3Arrows = 1;
    public const int Xl3TrafficLights1 = 4;
    public const int Xl3Symbols = 6;
    public const int Xl3Symbols2 = 7;
    public const int Xl4Arrows = 8;
    public const int Xl5Arrows = 11;
    public const int MsoShapeRectangle = 1;
    public const int MsoShapeRoundedRectangle = 5;
    public const int MsoShapeOval = 9;
    public const int MsoShapeRightArrow = 33;
    public const int MsoShapeDownArrow = 36;
    public const int MsoTextBox = 17;
    public const int MsoTextOrientationHorizontal = 1;
    public const int XlCategory = 1;
    public const int XlValue = 2;
    public const int XlPrimary = 1;
    public const int XlSecondary = 2;
    public const int XlTickLabelPositionNone = -4142;
    public const int XlTickMarkNone = -4142;
    public const int XlMarkerStyleNone = -4142;
    public const int XlMarkerStyleAutomatic = -4105;
    public const int XlMarkerStyleCircle = 8;
    public const int XlMarkerStyleDash = -4115;
    public const int XlMarkerStyleDiamond = 2;
    public const int XlMarkerStyleDot = -4118;
    public const int XlMarkerStyleSquare = 1;
    public const int XlMarkerStyleTriangle = 3;

    public static readonly IReadOnlyDictionary<string, int> ChartTypes =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["columnClustered"] = XlColumnClustered,
            ["columnStacked"] = XlColumnStacked,
            ["barClustered"] = XlBarClustered,
            ["barStacked"] = XlBarStacked,
            ["line"] = XlLine,
            ["lineMarkers"] = XlLineMarkers,
            ["pie"] = XlPie,
            ["area"] = XlArea,
            ["scatter"] = XlXYScatter,
            ["scatterLines"] = XlXYScatterLines,
        };

    public static readonly IReadOnlyDictionary<string, int> AxisGroups =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["primary"] = XlPrimary,
            ["secondary"] = XlSecondary,
        };

    public static readonly IReadOnlyDictionary<string, int> MarkerStyles =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["none"] = XlMarkerStyleNone,
            ["automatic"] = XlMarkerStyleAutomatic,
            ["circle"] = XlMarkerStyleCircle,
            ["dash"] = XlMarkerStyleDash,
            ["diamond"] = XlMarkerStyleDiamond,
            ["dot"] = XlMarkerStyleDot,
            ["square"] = XlMarkerStyleSquare,
            ["triangle"] = XlMarkerStyleTriangle,
        };

    public static readonly IReadOnlyDictionary<string, int> LegendPositions =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["bottom"] = XlLegendPositionBottom,
            ["corner"] = XlLegendPositionCorner,
            ["top"] = XlLegendPositionTop,
            ["right"] = XlLegendPositionRight,
            ["left"] = XlLegendPositionLeft,
        };

    public static readonly IReadOnlyDictionary<string, int> ValidationTypes =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["list"] = XlValidateList,
            ["whole"] = XlValidateWholeNumber,
            ["decimal"] = XlValidateDecimal,
            ["date"] = XlValidateDate,
            ["time"] = XlValidateTime,
            ["textLength"] = XlValidateTextLength,
            ["custom"] = XlValidateCustom,
        };

    public static readonly IReadOnlyDictionary<string, int> ComparisonOperators =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["between"] = XlBetween,
            ["notBetween"] = XlNotBetween,
            ["equal"] = XlEqual,
            ["notEqual"] = XlNotEqual,
            ["greater"] = XlGreater,
            ["less"] = XlLess,
            ["greaterEqual"] = XlGreaterEqual,
            ["lessEqual"] = XlLessEqual,
        };

    public static readonly IReadOnlyDictionary<string, int> ErrorStyles =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["stop"] = XlValidAlertStop,
            ["warning"] = XlValidAlertWarning,
            ["information"] = XlValidAlertInformation,
        };

    public static readonly IReadOnlyDictionary<string, int> TotalsFunctions =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["none"] = XlTotalsCalculationNone,
            ["average"] = XlTotalsCalculationAverage,
            ["count"] = XlTotalsCalculationCount,
            ["countNums"] = XlTotalsCalculationCountNums,
            ["max"] = XlTotalsCalculationMax,
            ["min"] = XlTotalsCalculationMin,
            ["sum"] = XlTotalsCalculationSum,
            ["stdDev"] = XlTotalsCalculationStdDev,
            ["var"] = XlTotalsCalculationVar,
            ["custom"] = XlTotalsCalculationCustom,
        };

    public static readonly IReadOnlySet<string> ConditionalRuleTypes =
        new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "cellValue", "expression", "uniqueValues", "duplicateValues", "colorScale", "dataBar", "iconSet",
        };

    public static readonly IReadOnlyDictionary<string, int> IconSets =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["arrows3"] = Xl3Arrows,
            ["trafficLights3"] = Xl3TrafficLights1,
            ["symbols3"] = Xl3Symbols,
            ["symbols3Alt"] = Xl3Symbols2,
            ["arrows4"] = Xl4Arrows,
            ["arrows5"] = Xl5Arrows,
        };

    public static readonly IReadOnlyDictionary<string, int> PivotFunctions =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["sum"] = XlSum,
            ["count"] = XlCount,
            ["average"] = XlAverage,
            ["max"] = XlMax,
            ["min"] = XlMin,
            ["countNums"] = XlCountNums,
            ["product"] = XlProduct,
            ["stdDev"] = XlStDev,
            ["var"] = XlVar,
        };

    public static readonly IReadOnlyDictionary<string, int> ShapeTypes =
        new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase)
        {
            ["rectangle"] = MsoShapeRectangle,
            ["roundedRectangle"] = MsoShapeRoundedRectangle,
            ["oval"] = MsoShapeOval,
            ["rightArrow"] = MsoShapeRightArrow,
            ["downArrow"] = MsoShapeDownArrow,
        };

    public static readonly IReadOnlySet<string> FilterOperators =
        new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            "equals", "notEquals", "greaterThan", "lessThan", "greaterOrEqual", "lessOrEqual",
            "between", "contains", "beginsWith", "endsWith", "values",
        };

    public static readonly IReadOnlySet<string> PictureExtensions =
        new HashSet<string>(StringComparer.OrdinalIgnoreCase)
        {
            ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".emf", ".wmf", ".tif", ".tiff",
        };

    public static string TokenFor(IReadOnlyDictionary<string, int> map, int value, string fallback)
    {
        foreach (var pair in map)
        {
            if (pair.Value == value) return pair.Key;
        }
        return fallback;
    }

    public static bool TryChartType(string? token, out int value) =>
        TryMap(ChartTypes, token, out value);

    public static bool TryAxisGroup(string? token, out int value) =>
        TryMap(AxisGroups, token, out value);

    public static bool TryLegendPosition(string? token, out int value) =>
        TryMap(LegendPositions, token, out value);

    public static bool TryValidationType(string? token, out int value) =>
        TryMap(ValidationTypes, token, out value);

    public static bool TryComparisonOperator(string? token, out int value) =>
        TryMap(ComparisonOperators, token, out value);

    public static bool TryErrorStyle(string? token, out int value) =>
        TryMap(ErrorStyles, token, out value);

    public static bool TryTotalsFunction(string? token, out int value) =>
        TryMap(TotalsFunctions, token, out value);

    public static bool TryIconSet(string? token, out int value) =>
        TryMap(IconSets, token, out value);

    public static bool TryPivotFunction(string? token, out int value) =>
        TryMap(PivotFunctions, token, out value);

    public static bool TryShapeType(string? token, out int value) =>
        TryMap(ShapeTypes, token, out value);

    public static bool TryMarkerStyle(string? token, out int value) =>
        TryMap(MarkerStyles, token, out value);

    public static int SortOrder(string? token) =>
        string.Equals(token, "desc", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(token, "descending", StringComparison.OrdinalIgnoreCase)
            ? XlDescending
            : XlAscending;

    public static int HasHeaders(bool hasHeaders) => hasHeaders ? XlYes : XlNo;

    public static int PlotBy(string? token) =>
        string.Equals(token, "rows", StringComparison.OrdinalIgnoreCase) ? XlRows : XlColumns;

    public static string FormatInvariant(double value) =>
        value.ToString("G15", CultureInfo.InvariantCulture);

    private static bool TryMap(IReadOnlyDictionary<string, int> map, string? token, out int value)
    {
        if (!string.IsNullOrWhiteSpace(token) && map.TryGetValue(token.Trim(), out value))
            return true;
        value = 0;
        return false;
    }
}
