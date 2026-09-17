namespace DocBridge.Development.ExcelProductionWorkflowProbe;

/// <summary>
/// Live public apply. DocumentRef is set only after a successful save whose
/// public identity matches the owned path. Every op except create/open is
/// stamped with target.workbook. Success requires ok==true.
/// </summary>
internal sealed class ApplySession
{
    private readonly PublicClient _client;
    private readonly string _sourceXlsx;
    private readonly BookInventory _bystanders;
    private readonly IReadOnlySet<string> _highRiskOps;
    private BookInventory? _createBaseline;
    private string? _createdName;
    private string? _documentRef;

    public ApplySession(PublicClient client, string sourceXlsx, BookInventory bystanders, IReadOnlySet<string> highRiskOps)
    {
        _client = client;
        _sourceXlsx = sourceXlsx;
        _bystanders = bystanders;
        _highRiskOps = highRiskOps;
    }

    public static HashSet<string> DiscoverHighRiskOps(JsonNode? capabilities, string? mcpPath)
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        AddHighRiskNames(set, capabilities);
        foreach (var path in PolicyPathsNear(mcpPath))
        {
            if (!File.Exists(path)) continue;
            try { AddHighRiskNames(set, JsonUtil.Load(path)); }
            catch { /* keep other sources */ }
        }
        return set;
    }

    public string? DocumentRef => _documentRef;
    public string? CreatedName => _createdName;
    public string? Binding => _documentRef ?? _createdName;

    public void ResetOwned()
    {
        _createdName = null;
        _documentRef = null;
        _createBaseline = null;
    }

    public void BindExistingOwned(string path)
    {
        if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            throw new InvalidOperationException("resume-owned path missing: " + path);
        SourceIntegrity.RefuseIfProtectedPath(path, _sourceXlsx);
        RefuseUnlessOwned(path);
        var inventory = CaptureContextInventory();
        var match = inventory.Books.FirstOrDefault(b => b.Matches(path));
        if (string.IsNullOrWhiteSpace(match.Primary))
            throw new InvalidOperationException("resume-owned workbook is not in excel_get_active_context.openWorkbooks: " + path);
        RefuseUnlessOwned(match.Primary);
        _documentRef = Path.GetFullPath(path);
        _createdName = match.Name ?? Path.GetFileName(path);
    }

    public JsonNode Apply(PlannedBatch batch)
    {
        foreach (var op in batch.Ops.OfType<JsonObject>())
            GuardPaths(op);

        var names = batch.Ops.OfType<JsonObject>().Select(o => JsonUtil.Str(o, "op") ?? "").ToList();
        if (names.Contains("create_workbook"))
            _createBaseline = CaptureContextInventory();

        var ops = Stamp(batch.Ops);
        JsonNode response;
        if (batch.Path == "execute")
        {
            var path = _documentRef
                ?? throw new InvalidOperationException($"{batch.Id}: execute requires a verified saved documentRef");
            RefuseUnlessOwned(path);
            response = _client.Call("excel_apply_ops", new JsonObject
            {
                ["ops"] = ops.DeepClone(),
                ["executionMode"] = "execute",
                ["requestId"] = Guid.NewGuid().ToString(),
                ["expectedDocumentRef"] = path,
            });
        }
        else
        {
            var preview = _client.Call("excel_apply_ops", new JsonObject
            {
                ["ops"] = ops.DeepClone(),
                ["dryRun"] = true,
            });
            if (batch.InjectedFailure)
            {
                if (JsonUtil.Bool(preview, "ok") == true && !string.IsNullOrWhiteSpace(JsonUtil.Str(preview, "confirmToken")))
                    throw new InvalidOperationException($"{batch.Id}: injected failure preview unexpectedly issued a confirmToken");
                return preview;
            }

            RequireOk(preview, batch.Id, "dry-run");
            var token = JsonUtil.Str(preview, "confirmToken");
            if (string.IsNullOrWhiteSpace(token))
                throw new InvalidOperationException($"{batch.Id}: dry-run returned no confirmToken");

            var confirm = new JsonObject
            {
                ["ops"] = ops.DeepClone(),
                ["dryRun"] = false,
                ["confirmToken"] = token,
            };
            if (IsHighRisk(ops))
                confirm["highRiskConfirm"] = true;
            response = _client.Call("excel_apply_ops", confirm);
        }

        if (batch.InjectedFailure)
        {
            if (JsonUtil.Bool(response, "ok") == true)
                throw new InvalidOperationException($"{batch.Id}: injected failure unexpectedly succeeded");
            return response;
        }

        RequireOk(response, batch.Id, "apply");
        NoteIdentities(batch, response);
        return response;
    }

    public bool ShouldCheckpoint(PlannedBatch batch)
    {
        if (string.IsNullOrWhiteSpace(_documentRef)) return false;
        if (batch.InjectedFailure) return false;
        var names = batch.Ops.OfType<JsonObject>().Select(o => JsonUtil.Str(o, "op") ?? "").ToList();
        if (names.Count == 0) return false;
        if (names.All(n => n is "create_workbook" or "open_workbook" or "close_workbook"
            or "save_workbook" or "export_pdf"))
            return false;
        return true;
    }

    public JsonNode CheckpointSave(string reason)
    {
        var path = _documentRef
            ?? throw new InvalidOperationException("checkpoint save requires a verified saved documentRef");
        RefuseUnlessOwned(path);
        if (!IsOwnedUnsavedDefault(path))
            SourceIntegrity.RefuseIfProtectedPath(path, _sourceXlsx);
        var batch = ApplyPlanner.Batch("checkpoint", "checkpoint-save", OpFamily.Lifecycle, JsonUtil.Arr(
            new JsonObject
            {
                ["op"] = "save_workbook",
                ["output"] = path,
                ["overwrite"] = true,
            }), "token", reason);
        return Apply(batch);
    }

    public JsonObject TryCheckpointSave(string reason)
    {
        if (string.IsNullOrWhiteSpace(_documentRef))
            return new JsonObject { ["ok"] = false, ["skipped"] = true, ["reason"] = "no documentRef" };
        try
        {
            var response = CheckpointSave(reason);
            return new JsonObject
            {
                ["ok"] = JsonUtil.Bool(response, "ok") == true,
                ["reason"] = reason,
                ["documentRef"] = _documentRef,
                ["responseOk"] = JsonUtil.Bool(response, "ok"),
            };
        }
        catch (Exception ex)
        {
            return new JsonObject { ["ok"] = false, ["reason"] = reason, ["error"] = ex.Message };
        }
    }

    public JsonNode TryExecuteValues(string sheet, string range, JsonArray values)
    {
        var path = _documentRef
            ?? throw new InvalidOperationException("protect probe requires a verified saved documentRef");
        RefuseUnlessOwned(path);
        return _client.Call("excel_apply_ops", new JsonObject
        {
            ["ops"] = JsonUtil.Arr(new JsonObject
            {
                ["op"] = "set_values",
                ["target"] = JsonUtil.Target(sheet, path),
                ["range"] = range,
                ["values"] = values,
            }),
            ["executionMode"] = "execute",
            ["requestId"] = Guid.NewGuid().ToString(),
            ["expectedDocumentRef"] = path,
        });
    }

    private void NoteIdentities(PlannedBatch batch, JsonNode response)
    {
        var names = batch.Ops.OfType<JsonObject>().Select(o => JsonUtil.Str(o, "op") ?? "").ToList();
        if (names.Contains("create_workbook"))
        {
            var fromResponse = IdentityGuard.FirstWorkbookIdentity(response);
            var after = CaptureContextInventory();
            var baseline = _createBaseline ?? throw new InvalidOperationException("create_workbook missing immediate pre-create inventory");
            var owned = IdentityGuard.RequireNewOwned(baseline, after, _sourceXlsx);
            if (!string.IsNullOrWhiteSpace(fromResponse) &&
                !owned.Matches(fromResponse) &&
                !owned.Matches(Path.GetFileName(fromResponse)))
            {
                throw new InvalidOperationException($"create identity mismatch: response '{fromResponse}' vs inventory '{owned.Primary}'");
            }
            _createdName = owned.Name ?? owned.Primary;
            RefuseUnlessOwned(owned.Primary);
            _documentRef = null;
            return;
        }

        if (names.Contains("open_workbook"))
        {
            var plannedOpen = batch.Ops.OfType<JsonObject>()
                .Where(o => JsonUtil.Str(o, "op") == "open_workbook")
                .Select(o => JsonUtil.Str(o, "path"))
                .FirstOrDefault(s => !string.IsNullOrWhiteSpace(s));
            if (string.IsNullOrWhiteSpace(plannedOpen))
                throw new InvalidOperationException($"{batch.Id}: open_workbook has no path");
            RefuseUnlessOwned(plannedOpen);
            if (!IsOwnedUnsavedDefault(plannedOpen))
                SourceIntegrity.RefuseIfProtectedPath(plannedOpen, _sourceXlsx);

            var opened = ResolveOwnedIdentity(response, plannedOpen);
            if (string.IsNullOrWhiteSpace(opened))
                throw new InvalidOperationException($"{batch.Id}: open returned no public identity matching planned '{plannedOpen}' (active bystander is not used)");
            if (JsonUtil.Bool(JsonUtil.Get(response, "readback"), "verified") != true)
                throw new InvalidOperationException($"{batch.Id}: open requires readback.verified");
            if (!AffectedWorkbookMatches(response, plannedOpen))
                throw new InvalidOperationException($"{batch.Id}: open requires affected workbook ref matching '{plannedOpen}'");

            RefuseUnlessOwned(opened);
            if (!File.Exists(plannedOpen) || !File.Exists(opened))
                throw new InvalidOperationException($"{batch.Id}: opened path missing: {opened}");
            if (!SamePath(plannedOpen, opened))
                throw new InvalidOperationException($"{batch.Id}: open identity mismatch: public '{opened}' vs planned '{plannedOpen}'");

            var inventory = CaptureContextInventory();
            if (!inventory.Contains(opened) && !inventory.Contains(plannedOpen))
                throw new InvalidOperationException($"{batch.Id}: open inventory does not list owned path '{opened}'");
            var sheet = inventory.Books.FirstOrDefault(b => b.Matches(opened) || b.Matches(plannedOpen)).Sheets?.FirstOrDefault()
                        ?? "Sheet1";
            var read = PublicApi.ReadRange(_client, opened, sheet, "A1");
            if (JsonUtil.Bool(read, "ok") != true || !PublicApi.HasReadableValues(read))
                throw new InvalidOperationException($"{batch.Id}: explicit owned-path read failed: {read}");

            _documentRef = Path.GetFullPath(opened);
            return;
        }

        if (!names.Contains("save_workbook")) return;

        var planned = batch.Ops.OfType<JsonObject>()
            .Where(o => JsonUtil.Str(o, "op") == "save_workbook")
            .Select(o => JsonUtil.Str(o, "output"))
            .FirstOrDefault(s => !string.IsNullOrWhiteSpace(s));
        if (string.IsNullOrWhiteSpace(planned))
            throw new InvalidOperationException($"{batch.Id}: save_workbook has no output path");

        var actual = ResolveOwnedIdentity(response, planned);
        if (string.IsNullOrWhiteSpace(actual))
            throw new InvalidOperationException($"{batch.Id}: save returned no public identity matching planned '{planned}' (active bystander is not used)");

        RefuseUnlessOwned(actual);
        if (!File.Exists(planned) || !File.Exists(actual))
            throw new InvalidOperationException($"save_workbook claimed success but file missing: {actual}");
        if (!SamePath(planned, actual))
            throw new InvalidOperationException($"save identity mismatch: public '{actual}' vs planned '{planned}'");
        _documentRef = Path.GetFullPath(actual);
    }

    private string? ResolveOwnedIdentity(JsonNode response, string planned)
    {
        foreach (var candidate in IdentityCandidates(response))
        {
            if (SamePath(planned, candidate))
                return candidate;
        }

        var inventory = CaptureContextInventory();
        var match = inventory.Books.FirstOrDefault(b =>
            b.Matches(planned) || SamePath(planned, b.Primary) || SamePath(planned, b.FullName ?? "") || SamePath(planned, b.DocumentRef ?? ""));
        if (!string.IsNullOrWhiteSpace(match.Primary))
            return match.Primary;
        return null;
    }

    private static bool AffectedWorkbookMatches(JsonNode response, string planned)
    {
        var affected = JsonUtil.Get(response, "affected") as JsonArray
                       ?? JsonUtil.Get(JsonUtil.Get(response, "result"), "affected") as JsonArray;
        if (affected is null) return false;
        foreach (var item in affected.OfType<JsonObject>())
        {
            var r = JsonUtil.Str(item, "ref") ?? JsonUtil.Str(item, "workbook") ?? JsonUtil.Str(item, "fullName");
            if (!string.IsNullOrWhiteSpace(r) && SamePath(planned, r))
                return true;
        }
        return false;
    }

    private static IEnumerable<string> IdentityCandidates(JsonNode response)
    {
        var direct = IdentityGuard.FirstWorkbookIdentity(response);
        if (!string.IsNullOrWhiteSpace(direct))
            yield return direct;

        var affected = JsonUtil.Get(response, "affected") as JsonArray
                       ?? JsonUtil.Get(JsonUtil.Get(response, "result"), "affected") as JsonArray;
        if (affected is null) yield break;
        foreach (var item in affected.OfType<JsonObject>())
        {
            var r = JsonUtil.Str(item, "ref") ?? JsonUtil.Str(item, "workbook") ?? JsonUtil.Str(item, "fullName");
            if (!string.IsNullOrWhiteSpace(r))
                yield return r;
        }
    }

    private JsonArray Stamp(JsonArray ops)
    {
        var copy = (JsonArray)ops.DeepClone();
        foreach (var op in copy.OfType<JsonObject>())
        {
            var name = JsonUtil.Str(op, "op") ?? "";
            if (name is "create_workbook" or "open_workbook") continue;

            var binding = Binding
                ?? throw new InvalidOperationException($"{name} requires an owned workbook identity");
            if (string.IsNullOrWhiteSpace(binding))
                throw new InvalidOperationException($"{name} refuses a null/empty workbook identity (product close would bind the active user book)");
            RefuseUnlessOwned(binding);
            var target = JsonUtil.Get(op, "target") as JsonObject;
            if (target is null)
            {
                target = new JsonObject();
                op["target"] = target;
            }
            target["workbook"] = binding;
            if (name is "close_workbook" or "save_workbook" or "export_pdf")
                op["targetWorkbook"] = binding;
        }
        return copy;
    }

    private BookInventory CaptureContextInventory()
    {
        var context = _client.Call("excel_get_active_context", new JsonObject());
        var inventory = BookInventory.Parse(context);
        if (inventory.Books.Count == 0)
            throw new InvalidOperationException("refusing empty owned-inventory: excel_get_active_context.summary.openWorkbooks is missing or empty. An empty list is not 'no bystanders' when Excel may already have 통합 문서1.");
        return inventory;
    }

    private static void RequireOk(JsonNode response, string batchId, string phase)
    {
        if (JsonUtil.Bool(response, "ok") != true)
            throw new InvalidOperationException($"{batchId}: {phase} requires ok==true: {Redact.Tokens(response)}");
    }

    private static bool SamePath(string planned, string actual)
    {
        if (string.Equals(planned, actual, StringComparison.OrdinalIgnoreCase)) return true;
        if (!Path.IsPathRooted(actual)) return false;
        try
        {
            return string.Equals(Path.GetFullPath(planned), Path.GetFullPath(actual), StringComparison.OrdinalIgnoreCase);
        }
        catch { return false; }
    }

    private void GuardPaths(JsonObject op)
    {
        var name = JsonUtil.Str(op, "op") ?? "";
        if (name is "save_workbook" or "open_workbook" or "export_pdf" or "import_csv" or "export_csv")
        {
            var output = JsonUtil.Str(op, "output") ?? JsonUtil.Str(op, "path");
            if (!string.IsNullOrWhiteSpace(output))
            {
                if (!IsOwnedUnsavedDefault(output))
                    SourceIntegrity.RefuseIfProtectedPath(output, _sourceXlsx);
                RefuseUnlessOwned(output);
            }
        }
        var wb = JsonUtil.Str(JsonUtil.Get(op, "target"), "workbook");
        if (!string.IsNullOrWhiteSpace(wb) && wb != IdentityGuard.OwnedPlaceholder)
            RefuseUnlessOwned(wb);
    }

    private void RefuseUnlessOwned(string? candidate) =>
        IdentityGuard.RefuseProtected(candidate, _sourceXlsx, _bystanders,
            allowCreatedUnsaved: _documentRef is null ? _createdName : null);

    private bool IsOwnedUnsavedDefault(string? candidate) =>
        _documentRef is null &&
        IdentityGuard.IsAllowedCreatedUnsaved(candidate, _createdName, _bystanders);

    private bool IsHighRisk(JsonArray ops) =>
        ops.OfType<JsonObject>().Any(o => _highRiskOps.Contains(JsonUtil.Str(o, "op") ?? ""));

    private static void AddHighRiskNames(HashSet<string> set, JsonNode? node)
    {
        if (node is null) return;
        var arrays = new List<JsonArray?>
        {
            JsonUtil.Get(node, "highRiskOps") as JsonArray,
            JsonUtil.Get(JsonUtil.Get(node, "excel"), "highRiskOps") as JsonArray,
            JsonUtil.Get(JsonUtil.Get(JsonUtil.Get(node, "apps"), "excel"), "highRiskOps") as JsonArray,
            JsonUtil.Get(JsonUtil.Get(JsonUtil.Get(JsonUtil.Get(node, "result"), "apps"), "excel"), "highRiskOps") as JsonArray,
        };
        foreach (var arr in arrays)
        {
            if (arr is null) continue;
            foreach (var item in arr)
            {
                if (item is JsonValue v && v.TryGetValue<string>(out var name) && !string.IsNullOrWhiteSpace(name))
                    set.Add(name);
            }
        }
    }

    private static IEnumerable<string> PolicyPathsNear(string? mcpPath)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        void Yield(string? path)
        {
            if (string.IsNullOrWhiteSpace(path)) return;
            var full = Path.GetFullPath(path);
            if (seen.Add(full)) { /* unique */ }
        }

        if (!string.IsNullOrWhiteSpace(mcpPath))
        {
            var dir = Path.GetDirectoryName(Path.GetFullPath(mcpPath));
            for (var i = 0; i < 8 && !string.IsNullOrWhiteSpace(dir); i++)
            {
                Yield(Path.Combine(dir, "ops", "policies", "default.policy.json"));
                dir = Directory.GetParent(dir)?.FullName;
            }
        }

        Yield(Path.Combine(AppContext.BaseDirectory, "ops", "policies", "default.policy.json"));
        return seen;
    }
}
