namespace DocBridge.Core.Services;

/// <summary>
/// Absolute paths that product save/export must never overwrite.
/// Configured by <c>DOCBRIDGE_EXCEL_PROTECTED_SOURCES</c> (semicolon or pipe separated).
/// There is no compiled-in user workbook path.
/// </summary>
public static class ExcelAuthoringPaths
{
    public const string ProtectedSourcesVariable = "DOCBRIDGE_EXCEL_PROTECTED_SOURCES";

    public static IReadOnlyList<string> ProtectedSourceWorkbooks => ResolveProtectedSources();

    public static bool IsProtectedSource(string? path)
    {
        if (string.IsNullOrWhiteSpace(path)) return false;
        string full;
        try { full = Path.GetFullPath(path); }
        catch (Exception) { return false; }

        return ResolveProtectedSources().Any(item =>
        {
            try { return string.Equals(Path.GetFullPath(item), full, StringComparison.OrdinalIgnoreCase); }
            catch { return false; }
        });
    }

    public static bool IsLikelyUnsavedWorkbookName(string? pathOrName)
    {
        if (string.IsNullOrWhiteSpace(pathOrName)) return true;
        var name = Path.GetFileName(pathOrName);
        if (string.IsNullOrWhiteSpace(name)) return true;
        if (name.Contains(Path.DirectorySeparatorChar) || name.Contains(Path.AltDirectorySeparatorChar))
            return false;
        if (name.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase) ||
            name.EndsWith(".xlsm", StringComparison.OrdinalIgnoreCase) ||
            name.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
            return false;
        return name.StartsWith("Book", StringComparison.OrdinalIgnoreCase) ||
               name.StartsWith("통합 문서", StringComparison.OrdinalIgnoreCase) ||
               name.StartsWith("Workbook", StringComparison.OrdinalIgnoreCase);
    }

    private static IReadOnlyList<string> ResolveProtectedSources()
    {
        var raw = Environment.GetEnvironmentVariable(ProtectedSourcesVariable);
        if (string.IsNullOrWhiteSpace(raw))
            return Array.Empty<string>();
        return raw
            .Split(new[] { ';', '|' }, StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Where(item => !string.IsNullOrWhiteSpace(item))
            .ToArray();
    }
}
