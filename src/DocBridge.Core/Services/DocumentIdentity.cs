namespace DocBridge.Core.Services;

/// <summary>
/// Execute-only document identity. Bare unsaved names (Book1, Drawing1) are
/// ambiguous across instances. Excel/CAD/Gstar require an absolute path or an
/// adapter-supplied instance-bound ref. HWP keeps hwp:PID:id / untitled-PID-id.
/// Documents are never saved to satisfy this check.
/// </summary>
public static class DocumentIdentity
{
    public const string LegacyPreviewGuidance =
        "Unsaved or bare names such as Book1/Drawing1 are ambiguous across application instances. " +
        "executionMode=execute will not save the document to create a path. " +
        "Use dryRun=true to receive a confirmToken, then apply the same ops with that token.";

    public static bool RequiresAbsoluteOrInstanceBound(string app) =>
        app.Equals("excel", StringComparison.OrdinalIgnoreCase) ||
        app.Equals("cad", StringComparison.OrdinalIgnoreCase) ||
        app.Equals("gstarcad", StringComparison.OrdinalIgnoreCase);

    public static bool TryParseHwpStableRef(string value, out string processId, out string documentId)
    {
        processId = "";
        documentId = "";
        string[] parts;
        if (value.StartsWith("hwp:", StringComparison.OrdinalIgnoreCase))
            parts = value.Split(':', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        else if (value.StartsWith("untitled-", StringComparison.OrdinalIgnoreCase))
            parts = value.Split('-', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        else
            return false;

        if (parts.Length < 3) return false;
        processId = parts[1];
        documentId = parts[^1];
        return int.TryParse(processId, out _) && int.TryParse(documentId, out _);
    }

    public static bool TryNormalizeExecuteRef(string app, string? raw, out string normalized, out string? error)
    {
        normalized = "";
        error = null;
        if (string.IsNullOrWhiteSpace(raw))
        {
            error = "expectedDocumentRef is empty. " + LegacyPreviewGuidance;
            return false;
        }

        var trimmed = raw.Trim();
        if (TryAbsolutePath(trimmed, out var full))
        {
            normalized = full;
            return true;
        }

        if (app.Equals("hwp", StringComparison.OrdinalIgnoreCase) &&
            TryParseHwpStableRef(trimmed, out var processId, out var documentId))
        {
            normalized = $"hwp:{processId}:{documentId}";
            return true;
        }

        if (IsAdapterInstanceBoundRef(app, trimmed))
        {
            normalized = trimmed;
            return true;
        }

        if (RequiresAbsoluteOrInstanceBound(app) || app.Equals("hwp", StringComparison.OrdinalIgnoreCase))
        {
            error = $"'{trimmed}' is not a stable document identity. " + LegacyPreviewGuidance;
            return false;
        }

        normalized = trimmed;
        return true;
    }

    public static bool IsAdapterInstanceBoundRef(string app, string value)
    {
        if (!RequiresAbsoluteOrInstanceBound(app) || string.IsNullOrWhiteSpace(value))
            return false;
        var prefix = app.Trim() + "-instance:";
        return value.StartsWith(prefix, StringComparison.OrdinalIgnoreCase) &&
               value.Length > prefix.Length;
    }

    public static bool TryAbsolutePath(string value, out string fullPath)
    {
        fullPath = "";
        if (!Path.IsPathFullyQualified(value)) return false;
        try
        {
            fullPath = Path.GetFullPath(value);
            return true;
        }
        catch
        {
            return false;
        }
    }
}
