using System.Linq;
using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// Advertised chart series data: values, categories, name, optional range.
/// Union sourceRange is legal; SetSourceData or explicit series both apply.
/// </summary>
public static class ExcelChartSeriesContract
{
    public readonly record struct SeriesSpec(
        int Index,
        string? Values,
        string? Categories,
        string? Name,
        bool NameIsRange,
        string? Range,
        string? ChartType = null,
        string? AxisGroup = null);

    public static IReadOnlyList<SeriesSpec> ReadSeries(JsonArray? series)
    {
        if (series is null || series.Count == 0) return [];
        var list = new List<SeriesSpec>(series.Count);
        for (var i = 0; i < series.Count; i++)
        {
            if (series[i] is not JsonObject item) continue;
            var name = Json.GetString(item, "name");
            var nameIsRange = LooksLikeNameRange(name);
            var index = Json.GetInt(item, "index");
            list.Add(new SeriesSpec(
                index is > 0 ? index.Value : i + 1,
                EmptyToNull(Json.GetString(item, "values")),
                EmptyToNull(Json.GetString(item, "categories")),
                EmptyToNull(name),
                nameIsRange,
                EmptyToNull(Json.GetString(item, "range")),
                EmptyToNull(Json.GetString(item, "chartType")),
                EmptyToNull(Json.GetString(item, "axisGroup"))));
        }
        return list;
    }

    public static bool HasSeriesData(JsonArray? series) =>
        ReadSeries(series).Any(item =>
            item.Values is not null || item.Categories is not null ||
            item.Name is not null || item.Range is not null);

    /// <summary>
    /// Per-series ChartType / AxisGroup override. Omitted tokens are preserved.
    /// </summary>
    public static bool HasComboOverride(JsonArray? series) =>
        ReadSeries(series).Any(item => item.ChartType is not null || item.AxisGroup is not null);

    public static string? EffectiveChartType(SeriesSpec spec, string? chartLevelType) =>
        spec.ChartType ?? chartLevelType;

    /// <summary>
    /// Chart.ChartType may become an unmapped combination enum only when
    /// requested series types actually differ. AxisGroup-only does not waive
    /// a requested base type.
    /// </summary>
    public static bool ChartLevelTypeMatches(string? actualType, string? requestedType, JsonArray? series)
    {
        if (string.IsNullOrWhiteSpace(requestedType)) return true;
        if (string.Equals(actualType, requestedType, StringComparison.OrdinalIgnoreCase))
            return true;
        if (string.IsNullOrWhiteSpace(actualType) ||
            !actualType.StartsWith("raw:", StringComparison.OrdinalIgnoreCase))
            return false;
        return DistinctRequestedChartTypes(series, requestedType).Count > 1;
    }

    public static IReadOnlyList<string> DistinctRequestedChartTypes(JsonArray? series, string? chartLevelType)
    {
        var types = new List<string>();
        foreach (var spec in ReadSeries(series))
        {
            var type = EffectiveChartType(spec, chartLevelType);
            if (type is null) continue;
            if (!types.Exists(existing => string.Equals(existing, type, StringComparison.OrdinalIgnoreCase)))
                types.Add(type);
        }
        return types;
    }

    public static bool IndexesAreDenseFromOne(IReadOnlyList<SeriesSpec> specs)
    {
        if (specs.Count == 0) return true;
        var indexes = specs.Select(item => item.Index).OrderBy(index => index).ToArray();
        if (indexes[0] != 1) return false;
        for (var i = 1; i < indexes.Length; i++)
        {
            if (indexes[i] == indexes[i - 1] || indexes[i] != indexes[i - 1] + 1)
                return false;
        }
        return true;
    }

    public static bool LooksLikeRange(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return false;
        return ExcelDataOperationsContract.TryParseA1(text, out _, allowUnion: true);
    }

    /// <summary>
    /// Microsoft Series.Name is a literal String unless the caller writes an
    /// explicit formula such as =Sheet1!R1C1. Q1 and 검수완료! are names.
    /// </summary>
    public static bool LooksLikeNameRange(string? text) =>
        ExcelReferenceIdentity.IsExplicitNameReference(text);

    public static bool FormulaMentions(
        string? formula, string? range, string? defaultSheet = null, string? ownerWorkbook = null) =>
        FormulaMentionsValues(formula, range, defaultSheet, ownerWorkbook);

    public static bool FormulaMentionsValues(
        string? formula, string? range, string? defaultSheet = null, string? ownerWorkbook = null) =>
        ExcelReferenceIdentity.FormulaContainsReference(formula, range, defaultSheet, slot: 2, ownerWorkbook);

    public static bool FormulaMentionsCategories(
        string? formula, string? range, string? defaultSheet = null, string? ownerWorkbook = null) =>
        ExcelReferenceIdentity.FormulaContainsReference(formula, range, defaultSheet, slot: 1, ownerWorkbook);

    public static bool FormulaMentionsName(
        string? formula, string? range, string? defaultSheet = null, string? ownerWorkbook = null) =>
        ExcelReferenceIdentity.FormulaContainsReference(formula, range, defaultSheet, slot: 0, ownerWorkbook);

    public static bool FormulaMentionsSource(
        string? formula, string? range, string? defaultSheet = null, string? ownerWorkbook = null) =>
        ExcelReferenceIdentity.FormulaContainsSourceArea(formula, range, defaultSheet, ownerWorkbook);

    private static string? EmptyToNull(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value;
}
