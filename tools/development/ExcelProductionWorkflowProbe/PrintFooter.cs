namespace DocBridge.Development.ExcelProductionWorkflowProbe;

internal static class PrintFooter
{
    public static string StripCenterMarker(string? footer)
    {
        if (string.IsNullOrEmpty(footer)) return "";
        return footer.StartsWith("&C", StringComparison.Ordinal) ? footer[2..] : footer;
    }
}
