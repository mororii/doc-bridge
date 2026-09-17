namespace DocBridge.Development.ExcelProductionWorkflowProbe;

/// <summary>
/// Per-cell style facts for the E1 property-layer planner.
/// Lives beside StyleReconstructor so that file can keep evolving without
/// blocking other scenario compiles.
/// </summary>
internal sealed class CellSemantics
{
    public string FontName { get; set; } = "";
    public double FontSize { get; set; }
    public bool Bold { get; set; }
    public bool Italic { get; set; }
    public string HAlign { get; set; } = "";
    public string VAlign { get; set; } = "";
    public bool Wrap { get; set; }
    public bool Shrink { get; set; }
    public string NumberFormat { get; set; } = "General";
    public string Fill { get; set; } = "";
    public string Left { get; set; } = "";
    public string Right { get; set; } = "";
    public string Top { get; set; } = "";
    public string Bottom { get; set; } = "";

    public CellSemantics Clone() => new()
    {
        FontName = FontName,
        FontSize = FontSize,
        Bold = Bold,
        Italic = Italic,
        HAlign = HAlign,
        VAlign = VAlign,
        Wrap = Wrap,
        Shrink = Shrink,
        NumberFormat = NumberFormat,
        Fill = Fill,
        Left = Left,
        Right = Right,
        Top = Top,
        Bottom = Bottom,
    };
}
