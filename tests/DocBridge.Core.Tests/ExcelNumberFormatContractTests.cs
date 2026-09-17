using System.Runtime.InteropServices;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelNumberFormatContractTests
{
    [Fact]
    public void Public_mixed_fixture_keeps_general_and_quoted_lit()
    {
        Assert.Equal("General", ExcelNumberFormatContract.General);
        Assert.Equal("General", ExcelNumberFormatPreserve.MixedRunFixtureFormats[0]);
        Assert.Equal("0.00", ExcelNumberFormatPreserve.MixedRunFixtureFormats[1]);
        Assert.Equal("yyyy-mm-dd", ExcelNumberFormatPreserve.MixedRunFixtureFormats[2]);
        Assert.True(ExcelNumberFormatContract.ReadbackMatches("General", "G/표준"));
        Assert.True(ExcelNumberFormatContract.ReadbackMatches("General", "G/일반"));
        Assert.False(ExcelNumberFormatContract.ReadbackMatches("0.00", "General"));
    }

    [Fact]
    public void General_cannot_set_falls_back_to_number_format_local()
    {
        var surface = new FakeNumberFormatSurface
        {
            LocalGeneralName = "G/표준",
            ThrowOnNumberFormat = ExcelNumberFormatContract.CannotSetProperty,
        };

        ExcelNumberFormatContract.Assign(surface, ExcelNumberFormatContract.General);

        Assert.Equal(new[] { "General" }, surface.NumberFormatWrites);
        Assert.Equal(new[] { "G/표준" }, surface.NumberFormatLocalWrites);
    }

    [Fact]
    public void Non_general_cannot_set_does_not_rewrite_format()
    {
        var surface = new FakeNumberFormatSurface
        {
            LocalGeneralName = "G/표준",
            ThrowOnNumberFormat = ExcelNumberFormatContract.CannotSetProperty,
        };

        var ex = Assert.Throws<COMException>(() => ExcelNumberFormatContract.Assign(surface, "0.00"));
        Assert.Equal(ExcelNumberFormatContract.CannotSetProperty, ex.HResult);
        Assert.Equal(new[] { "0.00" }, surface.NumberFormatWrites);
        Assert.Empty(surface.NumberFormatLocalWrites);
    }

    [Fact]
    public void Missing_local_general_name_rethrows()
    {
        var surface = new FakeNumberFormatSurface
        {
            LocalGeneralName = null,
            ThrowOnNumberFormat = ExcelNumberFormatContract.CannotSetProperty,
        };

        Assert.Throws<COMException>(() => ExcelNumberFormatContract.Assign(surface, "General"));
        Assert.Empty(surface.NumberFormatLocalWrites);
    }

    [Fact]
    public void Quoted_lit_and_plain_general_write_requested_format()
    {
        var surface = new FakeNumberFormatSurface();
        ExcelNumberFormatContract.Assign(surface, "\"LIT\"");
        ExcelNumberFormatContract.Assign(surface, "General");
        Assert.Equal(new[] { "\"LIT\"", "General" }, surface.NumberFormatWrites);
        Assert.Empty(surface.NumberFormatLocalWrites);
        Assert.False(ExcelNumberFormatContract.IsGeneral("\"LIT\""));
    }

    private sealed class FakeNumberFormatSurface : ExcelNumberFormatContract.INumberFormatSurface
    {
        public List<string> NumberFormatWrites { get; } = [];
        public List<string> NumberFormatLocalWrites { get; } = [];
        public int? ThrowOnNumberFormat { get; set; }
        public string? LocalGeneralName { get; set; }

        public void SetNumberFormat(string format)
        {
            NumberFormatWrites.Add(format);
            if (ThrowOnNumberFormat is int hr)
                throw new COMException("Range 클래스 중 NumberFormat 속성을 설정할 수 없습니다.", hr);
        }

        public void SetNumberFormatLocal(string format) => NumberFormatLocalWrites.Add(format);
    }
}
