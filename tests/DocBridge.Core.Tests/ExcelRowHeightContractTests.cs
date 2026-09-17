using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelRowHeightContractTests
{
    [Fact]
    public void Screen_pixel_delta_over_72pt_is_the_owner_mapping()
    {
        var dpi96 = ExcelRowHeightContract.FromScreenPixelDelta(0, 96);
        Assert.True(dpi96.Measured);
        Assert.Equal(0.75, dpi96.PointsPerPixel);
        Assert.Equal(96, dpi96.PixelDeltaAt72);
        Assert.Contains("PointsToScreenPixelsY", dpi96.Source, StringComparison.Ordinal);

        var dpi120 = ExcelRowHeightContract.FromScreenPixelDelta(0, 120);
        Assert.True(dpi120.Measured);
        Assert.Equal(0.6, dpi120.PointsPerPixel);

        var zero = ExcelRowHeightContract.FromScreenPixelDelta(10, 10);
        Assert.False(zero.Measured);
        Assert.Null(zero.PointsPerPixel);
        Assert.Contains("bounded fallback", zero.Source, StringComparison.Ordinal);
    }

    [Fact]
    public void Observed_must_match_nearest_measured_pixel_steps()
    {
        var dpi96 = ExcelRowHeightContract.FromScreenPixelDelta(0, 96);
        Assert.Equal(27.75, ExcelRowHeightContract.NearestPixelStep(28, 0.75));
        Assert.True(ExcelRowHeightContract.ObservedMatchesNearestPixelStep(
            28, 27.75, dpi96, out var from28, out var okDiag));
        Assert.Equal(27.75, from28);
        Assert.Null(okDiag);

        var dpi120 = ExcelRowHeightContract.FromScreenPixelDelta(0, 120);
        Assert.True(ExcelRowHeightContract.ObservedMatchesNearestPixelStep(22, 21.6, dpi120, out var from22, out _));
        Assert.Equal(21.6, from22);
        Assert.True(ExcelRowHeightContract.ObservedMatchesNearestPixelStep(18, 18, dpi120, out _, out _));

        Assert.False(ExcelRowHeightContract.ObservedMatchesNearestPixelStep(
            22, 15, dpi96, out _, out var miss));
        Assert.Contains("expected 22", miss, StringComparison.Ordinal);
        Assert.Contains("actual 15", miss, StringComparison.Ordinal);
        Assert.Contains("pointsPerPixel", miss, StringComparison.Ordinal);
        Assert.DoesNotContain("owner-window one-pixel", miss, StringComparison.Ordinal);
    }

    [Fact]
    public void Unmeasured_mapping_is_explicit_fallback_not_invented_dpi()
    {
        Assert.False(ExcelRowHeightContract.PixelMapping.Unmeasured.Measured);
        Assert.True(ExcelRowHeightContract.ObservedMatchesNearestPixelStep(
            18, 18, ExcelRowHeightContract.PixelMapping.Unmeasured, out var exact, out _));
        Assert.Equal(18, exact);
        Assert.False(ExcelRowHeightContract.ObservedMatchesNearestPixelStep(
            28, 27.75, ExcelRowHeightContract.PixelMapping.Unmeasured, out _, out var fallback));
        Assert.Contains("bounded fallback", fallback, StringComparison.Ordinal);
        Assert.DoesNotContain("0.75", fallback, StringComparison.Ordinal);
        Assert.Contains("bounded fallback",
            ExcelRowHeightContract.MismatchMessage("내역서", 1, 28, 27.75, ExcelRowHeightContract.PixelMapping.Unmeasured),
            StringComparison.Ordinal);
        Assert.True(ExcelRowHeightContract.MustMeasureTargetWorkbookWindow("owned.xlsx", "other.xlsx"));
        Assert.False(ExcelRowHeightContract.MustMeasureTargetWorkbookWindow(@"C:\out\owned.xlsx", "owned.xlsx"));
        Assert.True(ExcelRowHeightContract.ZoomAffectsPixelMapping(75));
        Assert.False(ExcelRowHeightContract.ZoomAffectsPixelMapping(100));
    }
}
