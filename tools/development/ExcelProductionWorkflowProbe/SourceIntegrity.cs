using System.Security.Cryptography;

namespace DocBridge.Development.ExcelProductionWorkflowProbe;

internal static class SourceIntegrity
{
    public const string ForbiddenUserBook = "통합 문서1";

    public static bool IsForbiddenUserBook(string? candidate)
    {
        if (string.IsNullOrWhiteSpace(candidate)) return false;
        var name = Path.GetFileNameWithoutExtension(candidate);
        return name.Equals(ForbiddenUserBook, StringComparison.OrdinalIgnoreCase) ||
               candidate.Contains(ForbiddenUserBook, StringComparison.OrdinalIgnoreCase);
    }

    public static string Sha256File(string path)
    {
        using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
        var hash = SHA256.HashData(fs);
        return Convert.ToHexString(hash);
    }

    public static JsonObject Capture(string sourcePath, string expected)
    {
        if (!File.Exists(sourcePath))
            throw new FileNotFoundException("original schedule xlsx not found", sourcePath);

        var sha = Sha256File(sourcePath);
        var ok = string.Equals(sha, expected, StringComparison.OrdinalIgnoreCase);
        return new JsonObject
        {
            ["path"] = Path.GetFullPath(sourcePath),
            ["sha256"] = sha,
            ["expectedSha256"] = expected,
            ["match"] = ok,
            ["length"] = new FileInfo(sourcePath).Length,
            ["utc"] = DateTime.UtcNow.ToString("o", CultureInfo.InvariantCulture),
            ["openMode"] = "read-only FileShare.ReadWrite",
        };
    }

    public static void RefuseIfProtectedPath(string candidate, string sourceXlsx)
    {
        var full = Path.GetFullPath(candidate);
        if (string.Equals(full, Path.GetFullPath(sourceXlsx), StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("refusing to write the user original schedule xlsx");
        if (string.Equals(Path.GetFileNameWithoutExtension(full), ForbiddenUserBook, StringComparison.OrdinalIgnoreCase) ||
            full.Contains(ForbiddenUserBook, StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("refusing to write the user's 통합 문서1");
    }
}
