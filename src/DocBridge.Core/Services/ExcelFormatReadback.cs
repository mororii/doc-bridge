using System.Globalization;

namespace DocBridge.Core.Services;

/// <summary>
/// format_range non-border readback. Union Font.Bold can report the first
/// area only, so every <c>Range.Areas</c> member is checked.
/// </summary>
public static class ExcelFormatReadback
{
    public static bool AppliedBoldMatches(object rangeObject, bool wanted) =>
        ExcelFormatAreas.AllAreas(rangeObject, area => BoldOnArea(area, wanted));

    public static bool BoldOnArea(object areaObject, bool wanted)
    {
        object? font = null;
        try
        {
            font = (object)((dynamic)areaObject).Font;
            var raw = (object?)((dynamic)font).Bold;
            if (raw is null or DBNull) return false;
            return Convert.ToBoolean(raw, CultureInfo.InvariantCulture) == wanted;
        }
        catch
        {
            return false;
        }
        finally
        {
            RotHelper.ReleaseComReference(font);
        }
    }
}
