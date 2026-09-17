using System.Reflection;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public sealed class ExcelWorkbookObjectsFakeTests
{
    private static List<JsonObject> ReadStyles(object workbook)
    {
        var method = typeof(ExcelAdapter).GetMethod("ReadCellStyleStates", BindingFlags.Static | BindingFlags.NonPublic)!;
        return (List<JsonObject>)method.Invoke(null, new[] { workbook })!;
    }

    [Fact]
    public void Cell_style_inventory_lists_fake_styles()
    {
        var workbook = new StyleWorkbookFake();
        var items = ReadStyles(workbook);
        Assert.Equal(2, items.Count);
        Assert.Equal("Title", Json.GetString(items[0], "name"));
        Assert.True(Json.GetBool(items[0], "builtIn"));
    }

    [Fact]
    public void Link_inventory_lists_fake_sources_with_status()
    {
        var method = typeof(ExcelAdapter).GetMethod("ReadExcelLinks", BindingFlags.Static | BindingFlags.NonPublic)!;
        var items = (List<JsonObject>)method.Invoke(null, new object[] { new LinkWorkbookFake() })!;
        Assert.Single(items);
        Assert.Equal("C:\\bridge\\적용수량집계_r2.xlsx", Json.GetString(items[0], "source"));
        Assert.Equal("적용수량집계_r2.xlsx", Json.GetString(items[0], "leaf"));
        Assert.Equal(0, Json.GetInt(items[0], "status"));
        Assert.Equal("ok", Json.GetString(items[0], "statusName"));
    }

    [Fact]
    public void Empty_link_sources_yield_no_items()
    {
        var method = typeof(ExcelAdapter).GetMethod("ReadExcelLinks", BindingFlags.Static | BindingFlags.NonPublic)!;
        var items = (List<JsonObject>)method.Invoke(null, new object[] { new EmptyLinkWorkbookFake() })!;
        Assert.Empty(items);
    }

    [Fact]
    public void Border_clear_warns_about_drawn_shared_neighbor()
    {
        var method = typeof(ExcelAdapter).GetMethod("WarnBorderClearNeighbors", BindingFlags.Static | BindingFlags.NonPublic)!;
        var preview = new ApplyPreview();
        var borders = Json.ParseObject("""{ "all": "none" }""")!;
        method.Invoke(null, new object[] { new BorderSheetFake(), "S", "A16:E18", borders, preview });
        Assert.Equal(3, preview.Warnings.Count);
        Assert.Contains(preview.Warnings, warning =>
            warning.Contains("shared with S!F16:F18", StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void Border_clear_is_quiet_when_neighbors_are_bare()
    {
        var method = typeof(ExcelAdapter).GetMethod("WarnBorderClearNeighbors", BindingFlags.Static | BindingFlags.NonPublic)!;
        var preview = new ApplyPreview();
        var borders = Json.ParseObject("""{ "left": "none" }""")!;
        method.Invoke(null, new object[] { new BareBorderSheetFake(), "S", "B2:B4", borders, preview });
        Assert.Empty(preview.Warnings);
    }
}

public sealed class StyleFake
{
    private readonly string _name;
    public StyleFake(string name) => _name = name;
    public string Name => _name;
    public bool BuiltIn => true;
}

public sealed class StylesFake
{
    private readonly StyleFake[] _items = new[] { new StyleFake("Title"), new StyleFake("Normal") };
    public int Count => _items.Length;
    public StyleFake Item(int index) => _items[index - 1];
}

public sealed class StyleWorkbookFake
{
    private readonly StylesFake _styles = new();
    public StylesFake Styles => _styles;
}

public sealed class LinkWorkbookFake
{
    public string[] LinkSources(int type) => new[] { "C:\\bridge\\적용수량집계_r2.xlsx" };
    public int LinkInfo(string name, int info) => 0;
}

public sealed class EmptyLinkWorkbookFake
{
    public object? LinkSources(int type) => null;
}

public sealed class BorderFake
{
    private readonly int _style;
    public BorderFake(int style) => _style = style;
    public int LineStyle => _style;
}

public sealed class BorderBordersFake
{
    private readonly int _style;
    public BorderBordersFake(int style) => _style = style;
    public BorderFake Item(int index) => new(_style);
}

public sealed class BorderRangeFake
{
    private readonly int _style;
    public BorderRangeFake(int style) => _style = style;
    public BorderBordersFake Borders => new(_style);
}

public sealed class BorderSheetFake
{
    public BorderRangeFake Range(string address) => new(1);
}

public sealed class BareBorderSheetFake
{
    public BorderRangeFake Range(string address) => new(-4142);
}
