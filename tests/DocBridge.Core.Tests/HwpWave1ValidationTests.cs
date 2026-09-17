using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;

namespace DocBridge.Core.Tests;

/// <summary>Wave 1 신규/확장 HWP op의 COM-free 검증 규칙 테스트.</summary>
public sealed class HwpWave1ValidationTests
{
    private static JsonObject Obj(string json) => JsonNode.Parse(json)!.AsObject();

    [Fact]
    public void Character_effects_accept_valid_values()
    {
        var style = Obj("""{"outline":true,"shadow":true,"shadowColor":"#FF0000","emboss":true,"smallCaps":true,"kerning":true}""");
        HwpAdapter.ValidateCharacterStyle(style);
        var engrave = Obj("""{"engrave":true}""");
        HwpAdapter.ValidateCharacterStyle(engrave);
    }

    [Fact]
    public void Character_effects_reject_bad_shadow_color()
    {
        Assert.Throws<ArgumentException>(() =>
            HwpAdapter.ValidateCharacterStyle(Obj("""{"shadowColor":"red"}""")));
    }

    [Fact]
    public void Character_effects_reject_emboss_with_engrave()
    {
        Assert.Throws<ArgumentException>(() =>
            HwpAdapter.ValidateCharacterStyle(Obj("""{"emboss":true,"engrave":true}""")));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(9)]
    public void Paragraph_level_accepts_outline_range(int level)
    {
        HwpAdapter.ValidateParagraphStyle(Obj($"{{\"level\":{level}}}"));
    }

    [Theory]
    [InlineData(-1)]
    [InlineData(10)]
    public void Paragraph_level_rejects_out_of_range(int level)
    {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            HwpAdapter.ValidateParagraphStyle(Obj($"{{\"level\":{level}}}")));
    }

    [Fact]
    public void Page_setup_accepts_line_numbers()
    {
        HwpAdapter.ValidatePageSetup(Obj("""{"page":{"lineNumbers":true,"lineNumberStart":3}}"""));
        HwpAdapter.ValidatePageSetup(Obj("""{"page":{}}"""));
    }

    [Fact]
    public void Page_setup_rejects_bad_line_number_start()
    {
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            HwpAdapter.ValidatePageSetup(Obj("""{"page":{"lineNumberStart":0}}""")));
    }

    [Fact]
    public void Footnote_requires_text()
    {
        HwpAdapter.ValidateFootnote(Obj("""{"text":"주석"}"""), endnote: false);
        HwpAdapter.ValidateFootnote(Obj("""{"text":"주석"}"""), endnote: true);
        Assert.Throws<ArgumentException>(() =>
            HwpAdapter.ValidateFootnote(Obj("""{"text":""}"""), endnote: false));
        Assert.Throws<ArgumentException>(() =>
            HwpAdapter.ValidateFootnote(Obj("""{}"""), endnote: true));
    }

    [Fact]
    public void Repeat_header_rejects_negative_table_index()
    {
        HwpAdapter.ValidateTableSetRepeatHeader(Obj("""{}"""));
        HwpAdapter.ValidateTableSetRepeatHeader(Obj("""{"tableIndex":2,"repeat":false}"""));
        Assert.Throws<ArgumentException>(() =>
            HwpAdapter.ValidateTableSetRepeatHeader(Obj("""{"tableIndex":-1}""")));
    }

    [Fact]
    public void Borders_require_width_in_range()
    {
        HwpAdapter.ValidateTableSetBorders(Obj("""{"widthMm":0.5}"""));
        HwpAdapter.ValidateTableSetBorders(Obj("""{"tableIndex":1,"widthMm":1.0,"color":"#000000"}"""));
        HwpAdapter.ValidateTableSetBorders(Obj("""{"widthMm":0.5,"edges":["right","bottom"]}"""));
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            HwpAdapter.ValidateTableSetBorders(Obj("""{"widthMm":0.05}""")));
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            HwpAdapter.ValidateTableSetBorders(Obj("""{"widthMm":6}""")));
        Assert.Throws<ArgumentException>(() =>
            HwpAdapter.ValidateTableSetBorders(Obj("""{"widthMm":0.5,"color":"black"}""")));
        Assert.Throws<ArgumentException>(() =>
            HwpAdapter.ValidateTableSetBorders(Obj("""{"tableIndex":-1,"widthMm":0.5}""")));
        Assert.Throws<ArgumentException>(() =>
            HwpAdapter.ValidateTableSetBorders(Obj("""{"widthMm":0.5,"edges":[]}""")));
        Assert.Throws<ArgumentException>(() =>
            HwpAdapter.ValidateTableSetBorders(Obj("""{"widthMm":0.5,"edges":["diagonal"]}""")));
    }

    [Fact]
    public void Table_delete_rejects_negative_table_index()
    {
        HwpAdapter.ValidateTableDelete(Obj("""{}"""));
        Assert.Throws<ArgumentException>(() =>
            HwpAdapter.ValidateTableDelete(Obj("""{"tableIndex":-2}""")));
    }
}
