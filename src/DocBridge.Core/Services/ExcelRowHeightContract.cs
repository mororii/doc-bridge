using System.Globalization;

namespace DocBridge.Core.Services;

/// <summary>
/// set_row_heights writes the requested points once. Excel may snap the
/// stored height to the owner window's pixel grid. Measure that grid with
/// PointsToScreenPixelsY(72)-PointsToScreenPixelsY(0). Unmeasured windows
/// use an explicit bounded fallback — not an invented 96/120 DPI.
/// Observed height must sit on the nearest pixel steps of the request
/// within floating tolerance. Return the observed height.
/// </summary>
public static class ExcelRowHeightContract
{
    public const double MinPoints = 0.1;
    public const double MaxPoints = 409.5;
    public const double FloatTolerance = 1e-3;
    public const int MeasureSpanPoints = 72;

    public readonly record struct PixelMapping(bool Measured, double? PointsPerPixel, int PixelDeltaAt72, string Source)
    {
        public static PixelMapping Unmeasured { get; } = new(
            false, null, 0, "unmeasured-owner-window; bounded fallback");
    }

    public static bool MustMeasureTargetWorkbookWindow(string targetWorkbook, string? activeWorkbook)
    {
        if (string.IsNullOrWhiteSpace(targetWorkbook) || string.IsNullOrWhiteSpace(activeWorkbook))
            return true;
        return !string.Equals(
            System.IO.Path.GetFileName(targetWorkbook),
            System.IO.Path.GetFileName(activeWorkbook),
            StringComparison.OrdinalIgnoreCase);
    }

    public static bool ZoomAffectsPixelMapping(int zoom) => zoom != 100;

    public static PixelMapping FromScreenPixelDelta(int pixelsAt0, int pixelsAt72)
    {
        var delta = pixelsAt72 - pixelsAt0;
        if (delta == 0)
            return new PixelMapping(false, null, 0, "PointsToScreenPixelsY delta 0; bounded fallback");
        var pointsPerPixel = (double)MeasureSpanPoints / delta;
        if (pointsPerPixel <= 0 || double.IsNaN(pointsPerPixel) || double.IsInfinity(pointsPerPixel))
            return new PixelMapping(false, null, delta, "PointsToScreenPixelsY unusable delta; bounded fallback");
        return new PixelMapping(true, pointsPerPixel, delta, "PointsToScreenPixelsY over 72pt");
    }

    public static double NearestPixelStep(double points, double pointsPerPixel)
    {
        if (pointsPerPixel <= 0)
            return points;
        var steps = Math.Round(points / pointsPerPixel, MidpointRounding.AwayFromZero);
        return steps * pointsPerPixel;
    }

    public static bool AlmostEqual(double left, double right, double epsilon = FloatTolerance) =>
        Math.Abs(left - right) <= epsilon + 1e-12;

    public static bool ObservedMatchesNearestPixelStep(
        double requested, double actual, PixelMapping mapping,
        out double normalized, out string? diagnostic)
    {
        normalized = actual;
        diagnostic = null;
        if (AlmostEqual(requested, actual))
            return true;

        if (!mapping.Measured || mapping.PointsPerPixel is not double pointsPerPixel)
        {
            diagnostic = FormatDiagnostic(requested, actual, mapping, nearest: null);
            return false;
        }

        var nearest = NearestPixelStep(requested, pointsPerPixel);
        var lower = Math.Floor(requested / pointsPerPixel) * pointsPerPixel;
        var upper = Math.Ceiling(requested / pointsPerPixel) * pointsPerPixel;
        var onNearestStep = AlmostEqual(actual, nearest) ||
                            AlmostEqual(actual, lower) ||
                            AlmostEqual(actual, upper);
        var withinOneMeasuredPixel = Math.Abs(actual - requested) <= pointsPerPixel + FloatTolerance;
        if (onNearestStep && withinOneMeasuredPixel)
            return true;

        diagnostic = FormatDiagnostic(requested, actual, mapping, nearest);
        return false;
    }

    public static string MismatchMessage(
        string sheetName, int row, double requested, double actual, PixelMapping mapping)
    {
        ObservedMatchesNearestPixelStep(requested, actual, mapping, out _, out var diagnostic);
        return $"{sheetName}!row {row}: heightPoints readback mismatch: {diagnostic}";
    }

    private static string FormatDiagnostic(
        double requested, double actual, PixelMapping mapping, double? nearest)
    {
        var text =
            $"expected {requested.ToString("0.###", CultureInfo.InvariantCulture)}, " +
            $"actual {actual.ToString("0.###", CultureInfo.InvariantCulture)}, " +
            mapping.Source;
        if (mapping.Measured && mapping.PointsPerPixel is double pointsPerPixel)
        {
            text +=
                $", pointsPerPixel {pointsPerPixel.ToString("0.###", CultureInfo.InvariantCulture)}" +
                $", delta72={mapping.PixelDeltaAt72.ToString(CultureInfo.InvariantCulture)}";
            if (nearest is double step)
                text += $", nearest-pixel-step {step.ToString("0.###", CultureInfo.InvariantCulture)}";
        }

        return text;
    }
}
