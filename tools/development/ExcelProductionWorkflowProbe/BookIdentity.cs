namespace DocBridge.Development.ExcelProductionWorkflowProbe;

/// <summary>
/// Open-workbook inventory from public <c>excel_get_active_context.summary.openWorkbooks</c>.
/// <c>core_get_status.apps.excel.document</c> is a string and is not an inventory.
/// </summary>
internal sealed class BookInventory
{
    public required IReadOnlyList<BookRef> Books { get; init; }
    public string? ActiveRef { get; init; }

    public IEnumerable<string> Keys =>
        Books.SelectMany(b => new[] { b.DocumentRef, b.FullName, b.Name })
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .Select(s => s!)
            .Distinct(StringComparer.OrdinalIgnoreCase);

    public bool Contains(string? id) =>
        !string.IsNullOrWhiteSpace(id) &&
        Books.Any(b => b.Matches(id));

    public JsonObject ToJson() => new()
    {
        ["activeRef"] = ActiveRef,
        ["books"] = new JsonArray(Books.Select(b => b.ToJson()).ToArray()),
    };

    public static BookInventory Parse(params JsonNode?[] nodes)
    {
        var books = new Dictionary<string, BookRef>(StringComparer.OrdinalIgnoreCase);
        string? active = null;
        var sawOpenWorkbooks = false;
        foreach (var node in nodes)
        {
            if (node is not JsonObject o) continue;
            var summary = JsonUtil.Get(o, "summary") as JsonObject;
            var open = JsonUtil.Get(summary, "openWorkbooks") as JsonArray
                       ?? JsonUtil.Get(o, "openWorkbooks") as JsonArray;
            if (open is null) continue;
            sawOpenWorkbooks = true;
            foreach (var item in open.OfType<JsonObject>())
            {
                var book = BookRef.FromOpenWorkbook(item);
                if (string.IsNullOrWhiteSpace(book.Primary)) continue;
                books[book.Primary] = book;
                if (JsonUtil.Bool(item, "activeInInstance") == true)
                    active = book.Primary;
            }
        }

        if (!sawOpenWorkbooks)
            return new BookInventory { Books = [], ActiveRef = null };

        return new BookInventory { Books = books.Values.ToList(), ActiveRef = active };
    }

    public IReadOnlyList<BookRef> Except(BookInventory later) =>
        later.Books.Where(b => !Contains(b.Primary)).ToList();

    public BookInventory Without(string? path)
    {
        if (string.IsNullOrWhiteSpace(path)) return this;
        var kept = Books.Where(b => !b.Matches(path)).ToList();
        var active = !string.IsNullOrWhiteSpace(ActiveRef) && kept.Any(b => b.Matches(ActiveRef))
            ? ActiveRef
            : kept.FirstOrDefault().Primary is { Length: > 0 } p ? p : null;
        return new BookInventory { Books = kept, ActiveRef = active };
    }
}

internal readonly record struct BookRef(
    string? Name,
    string? DocumentRef,
    string? FullName,
    long? ExcelHwnd = null,
    IReadOnlyList<string>? Sheets = null,
    bool? Saved = null)
{
    public string Primary =>
        !string.IsNullOrWhiteSpace(FullName) ? FullName! :
        !string.IsNullOrWhiteSpace(DocumentRef) ? DocumentRef! :
        Name ?? "";

    public bool Matches(string id) =>
        string.Equals(DocumentRef, id, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(FullName, id, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(Name, id, StringComparison.OrdinalIgnoreCase) ||
        string.Equals(Primary, id, StringComparison.OrdinalIgnoreCase);

    public static BookRef FromOpenWorkbook(JsonObject item)
    {
        var sheets = (JsonUtil.Get(item, "sheets") as JsonArray)?
            .Select(n => n?.GetValue<string>())
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .Select(s => s!)
            .ToList();
        return new BookRef(
            JsonUtil.Str(item, "name"),
            JsonUtil.Str(item, "documentRef") ?? JsonUtil.Str(item, "fullName") ?? JsonUtil.Str(item, "name"),
            JsonUtil.Str(item, "fullName") ?? JsonUtil.Str(item, "name"),
            (long?)JsonUtil.Num(item, "excelHwnd"),
            sheets,
            JsonUtil.Bool(item, "saved"));
    }

    public JsonObject ToJson() => new()
    {
        ["name"] = Name,
        ["documentRef"] = DocumentRef,
        ["fullName"] = FullName,
        ["excelHwnd"] = ExcelHwnd,
        ["saved"] = Saved,
        ["sheets"] = Sheets is null ? null : new JsonArray(Sheets.Select(s => JsonValue.Create(s)).ToArray()),
    };
}

internal static class IdentityGuard
{
    public const string OwnedPlaceholder = "$owned";

    public static void RefuseIfPlaceholder(string? candidate)
    {
        if (string.Equals(candidate, OwnedPlaceholder, StringComparison.Ordinal))
            throw new InvalidOperationException("public read/inspect cannot use the $owned placeholder; pass the actual saved xlsx identity");
    }

    public static void RefuseProtected(string? candidate, string sourceXlsx, BookInventory? bystanders, string? allowCreatedUnsaved = null)
    {
        if (string.IsNullOrWhiteSpace(candidate) || candidate == OwnedPlaceholder) return;
        var allowCreated = IsAllowedCreatedUnsaved(candidate, allowCreatedUnsaved, bystanders);
        if (!allowCreated)
            SourceIntegrity.RefuseIfProtectedPath(candidate, sourceXlsx);
        if (SourceIntegrity.IsForbiddenUserBook(candidate) && !allowCreated)
            throw new InvalidOperationException($"refusing user workbook identity: {candidate}");
        if (string.Equals(Path.GetFullPath(SafeFull(candidate) ?? candidate), Path.GetFullPath(sourceXlsx), StringComparison.OrdinalIgnoreCase))
            throw new InvalidOperationException("refusing the original schedule workbook as a target");
        if (bystanders is not null && bystanders.Contains(candidate))
            throw new InvalidOperationException($"refusing bystander workbook: {candidate}");
    }

    /// <summary>
    /// Excel names the first unsaved workbook 통합 문서1 / Book1. That is not the
    /// lost user bystander when the name was absent from the pre-create inventory
    /// and public create_workbook just added it. A name already present before
    /// create stays forbidden.
    /// </summary>
    public static bool IsAllowedCreatedUnsaved(string? candidate, string? createdName, BookInventory? before)
    {
        if (string.IsNullOrWhiteSpace(candidate) || string.IsNullOrWhiteSpace(createdName))
            return false;
        if (!SourceIntegrity.IsForbiddenUserBook(candidate) && !SourceIntegrity.IsForbiddenUserBook(createdName))
            return false;
        if (before is not null && (before.Contains(candidate) || before.Contains(createdName)))
            return false;
        return string.Equals(candidate, createdName, StringComparison.OrdinalIgnoreCase)
               || string.Equals(
                   Path.GetFileNameWithoutExtension(candidate),
                   Path.GetFileNameWithoutExtension(createdName),
                   StringComparison.OrdinalIgnoreCase);
    }

    public static string? FirstWorkbookIdentity(JsonNode? response)
    {
        var direct = JsonUtil.Str(response, "documentRef")
                     ?? JsonUtil.Str(JsonUtil.Get(response, "document"), "documentRef")
                     ?? JsonUtil.Str(JsonUtil.Get(response, "workbook"), "documentRef");
        if (!string.IsNullOrWhiteSpace(direct)) return direct;

        var fromOpen = BookInventory.Parse(response);
        if (fromOpen.Books.Count == 1) return fromOpen.Books[0].Primary;
        if (!string.IsNullOrWhiteSpace(fromOpen.ActiveRef)) return fromOpen.ActiveRef;
        return null;
    }

    public static BookRef RequireNewOwned(BookInventory before, BookInventory after, string sourceXlsx)
    {
        var added = before.Except(after);
        if (added.Count != 1)
            throw new InvalidOperationException($"create_workbook must add exactly one workbook; added {added.Count}: {string.Join(",", added.Select(b => b.Primary))}");
        var owned = added[0];
        RefuseProtected(owned.Primary, sourceXlsx, before, allowCreatedUnsaved: owned.Primary);
        if (string.IsNullOrWhiteSpace(owned.Primary))
            throw new InvalidOperationException("create_workbook returned an empty identity");
        return owned;
    }

    public static JsonObject CompareBystanders(BookInventory before, BookInventory after, IEnumerable<string?> owned)
    {
        var ownedList = owned.Where(s => !string.IsNullOrWhiteSpace(s)).Select(s => s!).ToList();
        bool IsOwned(BookRef b) => ownedList.Any(b.Matches);
        var lost = before.Books.Where(b => !after.Contains(b.Primary) && !IsOwned(b)).Select(b => b.Primary).ToList();
        var unexpected = after.Books.Where(b => !before.Contains(b.Primary) && !IsOwned(b)).Select(b => b.Primary).ToList();
        return new JsonObject
        {
            ["ok"] = lost.Count == 0 && unexpected.Count == 0,
            ["lostBystanders"] = new JsonArray(lost.Select(s => JsonValue.Create(s)).ToArray()),
            ["unexpectedBooks"] = new JsonArray(unexpected.Select(s => JsonValue.Create(s)).ToArray()),
            ["ownedTracked"] = new JsonArray(ownedList.Select(s => JsonValue.Create(s)).ToArray()),
        };
    }

    public static JsonObject VerifyPublicFixtures(string fixtureDir)
    {
        var errors = new JsonArray();
        var contextPath = Path.Combine(fixtureDir, "existing-user-context-after-oracle.json");
        var statusPath = Path.Combine(fixtureDir, "existing-user-status-after-oracle.json");
        if (!File.Exists(contextPath) || !File.Exists(statusPath))
            return new JsonObject { ["ok"] = false, ["error"] = "inventory fixtures missing" };

        var context = JsonUtil.Load(contextPath);
        var status = JsonUtil.Load(statusPath);
        var fromStatus = ParseSafe(status);
        if (fromStatus.Books.Count != 0)
            errors.Add("status-only parse must not invent workbook rows from apps.excel.document");

        var fromContext = ParseSafe(context);
        if (fromContext.Books.Count != 1)
            errors.Add($"context openWorkbooks must yield exactly 1 book, got {fromContext.Books.Count}");
        else
        {
            var book = fromContext.Books[0];
            if (!book.Matches("통합 문서1")) errors.Add("expected user book 통합 문서1");
            if (book.ExcelHwnd != 8654646) errors.Add($"excelHwnd {book.ExcelHwnd} != 8654646");
            if (book.Sheets is null || book.Sheets.Count != 1 || book.Sheets[0] != "Sheet1")
                errors.Add("expected sheets [Sheet1]");
        }

        var both = ParseSafe(status, context);
        if (both.Books.Count != 1)
            errors.Add("status+context must still be one user book, not duplicated path/name rows");

        return new JsonObject { ["ok"] = errors.Count == 0, ["errors"] = errors };
    }

    private static BookInventory ParseSafe(params JsonNode?[] nodes) => BookInventory.Parse(nodes);

    private static string? SafeFull(string candidate)
    {
        try { return Path.IsPathRooted(candidate) ? Path.GetFullPath(candidate) : null; }
        catch { return null; }
    }
}
