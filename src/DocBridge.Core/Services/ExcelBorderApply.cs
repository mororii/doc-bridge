using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// Applies format_range borders on every <c>Range.Areas</c> member.
/// ScopeCell uses one contiguous-area edge plus its inside companion
/// (xlEdgeBottom + xlInsideHorizontal on A1:A3), not per-cell COM.
/// ScopeRange writes outline/inside on each area only — never the union
/// bounding box. Multi-area unions still walk Areas; merged proof stays
/// on ExcelBorderReadback (outer+inside or cell scan).
/// </summary>
public static class ExcelBorderApply
{
    public static void ApplyCanonical(
        object rangeObject,
        JsonObject borders,
        Action<object, ExcelBorderContract.EdgeSpec> applyOne)
    {
        ArgumentNullException.ThrowIfNull(rangeObject);
        ArgumentNullException.ThrowIfNull(borders);
        ArgumentNullException.ThrowIfNull(applyOne);

        var specs = ExcelBorderContract.ReadEdges(borders);
        ExcelFormatAreas.ForEachArea(rangeObject, area =>
        {
            var geometry = ExcelBorderApplicability.ReadGeometry(area);
            foreach (var spec in specs)
            {
                if (!ExcelBorderApplicability.EdgeApplies(spec.Name, geometry))
                    continue;
                if (spec.Scope == ExcelBorderContract.ScopeCell)
                    ApplyCellEdgesOnContiguousArea(area, spec, applyOne, geometry);
                else
                    applyOne(area, spec);
            }
        });
    }

    public static void ApplyCellEdgesOnContiguousArea(
        object areaObject,
        ExcelBorderContract.EdgeSpec spec,
        Action<object, ExcelBorderContract.EdgeSpec> applyOne)
    {
        ApplyCellEdgesOnContiguousArea(
            areaObject, spec, applyOne, ExcelBorderApplicability.ReadGeometry(areaObject));
    }

    public static void ApplyCellEdgesOnContiguousArea(
        object areaObject,
        ExcelBorderContract.EdgeSpec spec,
        Action<object, ExcelBorderContract.EdgeSpec> applyOne,
        ExcelBorderApplicability.AreaGeometry geometry)
    {
        ArgumentNullException.ThrowIfNull(areaObject);
        ArgumentNullException.ThrowIfNull(applyOne);
        applyOne(areaObject, spec);
        if (!ExcelBorderApplicability.ScopeCellNeedsInsideCompanion(spec.Name, geometry))
            return;
        var inside = ExcelBorderReadback.InsideCompanionName(spec.Name);
        if (inside is null) return;
        applyOne(areaObject, new ExcelBorderContract.EdgeSpec(
            inside, ExcelBorderContract.ScopeRange, spec.Weight, spec.LineStyle, spec.Color));
    }
}
