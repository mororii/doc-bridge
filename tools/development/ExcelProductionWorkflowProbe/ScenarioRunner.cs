namespace DocBridge.Development.ExcelProductionWorkflowProbe;

internal sealed class FailedRequirementException : InvalidOperationException
{
    public FailedRequirementException(string message) : base(message) { }
}

internal static class ScenarioRunner
{
    public static IReadOnlyList<string> Expand(IReadOnlyList<string> requested)
    {
        if (requested.Count == 1 && requested[0] == "all")
            return ["e1", "e2", "e3", "e4", "e5", "e6", "e7", "e8", "e9"];
        return requested;
    }

    public static List<PlannedBatch> Plan(string id, ProbeOptions options, string artifactDir, string pictureDir)
    {
        var fixtureDir = id == "e9" ? Path.Combine(options.OutputDir, "fixtures") : pictureDir;
        var fixture = PracticalReportFixtures.TryPlan(id, artifactDir, options.SourceXlsx, fixtureDir);
        if (fixture is not null)
            return InsertBeforeSave(fixture, RemainingAfterFixture(id, pictureDir));

        return id switch
        {
            "e1" => ScenarioE1.Plan(options, artifactDir),
            "e2" => ScenarioCore.E2(artifactDir, options.SourceXlsx),
            "e4" => ScenarioCore.E4(artifactDir, options.SourceXlsx),
            "e5" => ScenarioCore.E5(artifactDir, options.SourceXlsx),
            "e7" => ScenarioExtended.E7(artifactDir, options.SourceXlsx),
            _ => throw new ArgumentException($"unknown scenario {id}"),
        };
    }

    public static JsonObject Expected(string id)
    {
        var fixture = PracticalReportFixtures.TryExpected(id);
        if (fixture is not null)
        {
            if (id == "e8")
            {
                fixture["amountInitial"] = 10_000;
                fixture["amountAfterB1"] = 11_000;
                fixture["protect"] = "locked B3 write refused; unlocked B1=11 recalculates B3 to 11000; no password field";
                fixture["remainingRecalc"] = "e8-b1-unprotect + e8-unlock-b1 + e8-calculate + e8-b1-reprotect; fixture b3-still-10000 is not acceptance";
            }
            return fixture;
        }

        return id switch
        {
            "e1" => ScenarioE1.Expected(),
            "e2" => ScenarioCore.ExpectedE2(),
            "e4" => ScenarioCore.ExpectedE4(),
            "e5" => ScenarioCore.ExpectedE5(),
            "e7" => ScenarioExtended.ExpectedE7(),
            _ => new JsonObject(),
        };
    }

    private static IEnumerable<PlannedBatch> RemainingAfterFixture(string id, string pictureDir) => id switch
    {
        "e3" => RemainingOpWorkflows.E3("작업일보", pictureDir),
        "e6" => RemainingOpWorkflows.E6("월간보고"),
        "e8" => RemainingOpWorkflows.E8("입력양식"),
        _ => [],
    };

    private static List<PlannedBatch> InsertBeforeSave(List<PlannedBatch> plan, IEnumerable<PlannedBatch> extra)
    {
        var add = extra.ToList();
        if (add.Count == 0) return plan;
        var scenario = plan[0].Scenario;
        var idx = plan.FindIndex(b => b.Id == scenario + "-save");
        if (idx < 0)
        {
            plan.AddRange(add);
            return plan;
        }
        plan.InsertRange(idx, add);
        return plan;
    }

    public static JsonObject RunLive(ProbeOptions options, IReadOnlyList<string> ids)
    {
        if (!options.LeaseOk)
            throw new InvalidOperationException("live mode refuses without --lease-ok (root Excel lease after safety review)");
        if (!string.Equals(Environment.GetEnvironmentVariable("DOCBRIDGE_E2E"), "1", StringComparison.Ordinal))
            throw new InvalidOperationException("live mode refuses unless DOCBRIDGE_E2E=1");
        if (options.Transport != "mcp")
            throw new InvalidOperationException("live mode requires --transport mcp; CLI exits cannot prove workbook ownership and cannot carry confirmToken");
        if (string.IsNullOrWhiteSpace(options.McpPath) || !File.Exists(options.McpPath))
            throw new InvalidOperationException("live mode requires --mcp pointing at an already-built doc-bridge-mcp");

        Directory.CreateDirectory(options.ArtifactDir);
        var pictureDir = Path.Combine(options.OutputDir, "fixtures");
        Directory.CreateDirectory(pictureDir);
        Program.WriteFixtureManifest(options, pictureDir);
        using var client = new PublicClient(options);
        var protocol = client.ProtocolSelfTest();
        if (JsonUtil.Bool(protocol, "ok") != true)
            throw new InvalidOperationException("MCP protocol self-test failed; no Excel writes: " + protocol);

        var statusBefore = client.Call("core_get_status", new JsonObject { ["app"] = "excel" });
        var ctxBefore = SafeCall(client, "excel_get_active_context", new JsonObject());
        var startInventory = BookInventory.Parse(ctxBefore);
        var excel = JsonUtil.Get(JsonUtil.Get(statusBefore, "apps"), "excel") as JsonObject;
        var connected = JsonUtil.Bool(excel, "connected") == true;
        if (connected && startInventory.Books.Count == 0)
            throw new InvalidOperationException("false empty inventory: Excel is connected (" +
                (JsonUtil.Str(excel, "document") ?? "unknown") +
                ") but summary.openWorkbooks is empty. Refusing to treat that as no bystanders.");
        if (!string.IsNullOrWhiteSpace(options.ResumeOwnedPath))
            startInventory = startInventory.Without(options.ResumeOwnedPath);
        var bystanderSamples = SampleBystanders(client, startInventory, options.SourceXlsx);
        if (connected && JsonUtil.Num(bystanderSamples, "capturedCount") != startInventory.Books.Count)
            throw new InvalidOperationException("StartBook sampling did not capture every open workbook with a successful public read: " + bystanderSamples);

        var caps = client.Call("core_get_capabilities", new JsonObject());
        var advertised = ContractCatalog.AdvertisedWriteOps(caps);
        var highRiskOps = ApplySession.DiscoverHighRiskOps(caps, options.McpPath);
        if (highRiskOps.Count == 0)
            throw new InvalidOperationException("high-risk op discovery found no names in capabilities or ops/policies/default.policy.json next to the candidate MCP");
        var session = new ApplySession(client, options.SourceXlsx, startInventory, highRiskOps);

        var results = new JsonArray();
        var ownedBooks = new List<string>();
        Exception? error = null;
        JsonObject? success = null;
        var sourceHashBefore = SourceIntegrity.Capture(options.SourceXlsx, options.ExpectedSha256);
        try
        {
            foreach (var id in ids)
            {
                session.ResetOwned();
                var scenarioDir = Path.Combine(options.ArtifactDir, id);
                Directory.CreateDirectory(scenarioDir);
                var batches = Plan(id, options, scenarioDir, pictureDir);
                if (!string.IsNullOrWhiteSpace(options.ResumeOwnedPath))
                {
                    session.BindExistingOwned(options.ResumeOwnedPath);
                    TrackOwned(ownedBooks, session.DocumentRef);
                }
                var scenario = new JsonObject { ["id"] = id, ["batches"] = new JsonArray() };
                JsonNode? last = null;
                JsonNode? injected = null;
                JsonObject? initialRead = null;
                var captureBatchId = batches.FirstOrDefault(b => b.Capture == "initial")?.Id;
                var initialPath = Path.Combine(scenarioDir, "initial-capture.json");
                if (!string.IsNullOrWhiteSpace(options.ResumeFromBatchId) && captureBatchId is not null)
                {
                    var skipped = batches.Select(b => b.Id)
                        .TakeWhile(id0 => !string.Equals(id0, options.ResumeFromBatchId, StringComparison.OrdinalIgnoreCase))
                        .ToHashSet(StringComparer.OrdinalIgnoreCase);
                    if (skipped.Contains(captureBatchId))
                    {
                        if (File.Exists(initialPath) && (id != "e2" || skipped.Contains("e2-recalc-qty")))
                            initialRead = JsonUtil.Load(initialPath).AsObject();
                        else if (!string.IsNullOrWhiteSpace(options.ResumeOwnedPath) && session.DocumentRef is { } resumeCap)
                            initialRead = LiveVerifier.CaptureInitial(client, resumeCap, id);
                    }
                }
                if (initialRead is not null)
                    JsonUtil.Write(initialPath, initialRead);
                var skipping = !string.IsNullOrWhiteSpace(options.ResumeFromBatchId);
                var stopBeforeHit = false;
                foreach (var batch in batches)
                {
                    if (!string.IsNullOrWhiteSpace(options.StopBeforeBatchId)
                        && string.Equals(batch.Id, options.StopBeforeBatchId, StringComparison.OrdinalIgnoreCase))
                    {
                        stopBeforeHit = true;
                    }
                    if (stopBeforeHit)
                    {
                        scenario["batches"]!.AsArray().Add(new JsonObject
                        {
                            ["id"] = batch.Id,
                            ["ok"] = true,
                            ["skippedStopBefore"] = true,
                            ["stopBefore"] = options.StopBeforeBatchId,
                        });
                        continue;
                    }
                    if (skipping)
                    {
                        if (!string.Equals(batch.Id, options.ResumeFromBatchId, StringComparison.OrdinalIgnoreCase))
                        {
                            scenario["batches"]!.AsArray().Add(new JsonObject
                            {
                                ["id"] = batch.Id,
                                ["ok"] = true,
                                ["skippedResume"] = true,
                            });
                            continue;
                        }
                        skipping = false;
                    }
                    foreach (var op in batch.Ops.OfType<JsonObject>())
                    {
                        var name = JsonUtil.Str(op, "op") ?? "";
                        if (advertised.Count > 0 && !advertised.Contains(name) && ContractCatalog.FamilyOf(name) != OpFamily.Unknown)
                            throw new FailedRequirementException($"{id}/{batch.Id}: missing advertised op '{name}' is a failed requirement, not a skip");
                    }

                    if (string.Equals(batch.Path, "checkpoint", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(batch.Capture, "blank-nofill", StringComparison.OrdinalIgnoreCase))
                    {
                        var ownedBlank = session.DocumentRef
                            ?? throw new FailedRequirementException($"{id}/{batch.Id}: blank checkpoint requires the owned save path");
                        var check = BlankCheckpoint.VerifyOwned(ownedBlank, ScenarioE1.Sheet);
                        var checkPath = Path.Combine(scenarioDir, batch.Id + ".json");
                        JsonUtil.Write(checkPath, check);
                        var blankOk = JsonUtil.Bool(check, "ok") == true;
                        scenario["batches"]!.AsArray().Add(new JsonObject
                        {
                            ["id"] = batch.Id,
                            ["ok"] = blankOk,
                            ["path"] = "checkpoint",
                            ["blankCheckpoint"] = check.DeepClone(),
                            ["file"] = checkPath,
                        });
                        if (!blankOk)
                            throw new FailedRequirementException(
                                $"{id}/{batch.Id}: owned OOXML is not a blank/noFill baseline; refusing values and style noFill omission: " + check);
                        last = check;
                        continue;
                    }

                    JsonNode response;
                    try
                    {
                        response = session.Apply(batch);
                    }
                    catch (Exception ex)
                    {
                        scenario["batches"]!.AsArray().Add(new JsonObject
                        {
                            ["id"] = batch.Id,
                            ["ok"] = false,
                            ["error"] = ex.Message,
                            ["injectedFailure"] = batch.InjectedFailure,
                        });
                        results.Add(scenario);
                        throw;
                    }
                    last = response;
                    if (batch.InjectedFailure) injected = response;
                    TrackOwned(ownedBooks, session.DocumentRef);
                    if (batch.Capture == "initial" && session.DocumentRef is { } ownedForCapture)
                    {
                        initialRead = LiveVerifier.CaptureInitial(client, ownedForCapture, id);
                        JsonUtil.Write(initialPath, initialRead);
                    }
                    if (!string.IsNullOrWhiteSpace(batch.Capture)
                        && batch.Capture is not "initial" and not "blank-nofill"
                        && session.DocumentRef is { } evidenceOwned)
                    {
                        var evidence = LiveVerifier.CaptureEvidence(client, evidenceOwned, id, batch);
                        JsonUtil.Write(Path.Combine(scenarioDir, batch.Id + "-evidence.json"), evidence);
                    }
                    JsonObject? checkpoint = null;
                    if (session.ShouldCheckpoint(batch))
                        checkpoint = session.TryCheckpointSave("after " + batch.Id);
                    scenario["batches"]!.AsArray().Add(new JsonObject
                    {
                        ["id"] = batch.Id,
                        ["ok"] = batch.InjectedFailure
                            ? JsonUtil.Bool(response, "ok") != true
                            : JsonUtil.Bool(response, "ok") == true,
                        ["productReadbackVerified"] = JsonUtil.Bool(JsonUtil.Get(response, "readback"), "verified"),
                        ["rollbackVerified"] = JsonUtil.Bool(JsonUtil.Get(response, "rollback"), "verified"),
                        ["injectedFailure"] = batch.InjectedFailure,
                        ["injectedPhase"] = batch.InjectedFailure ? InjectedPhase(response) : null,
                        ["checkpointSave"] = checkpoint,
                    });
                }

                var ownedPath = session.DocumentRef
                    ?? throw new InvalidOperationException($"{id}: scenario ended without a verified saved documentRef");
                var independent = LiveVerifier.VerifyScenario(new LiveVerifyRequest
                {
                    Client = client,
                    Session = session,
                    Id = id,
                    ArtifactDir = scenarioDir,
                    SourceXlsx = options.SourceXlsx,
                    ExpectedSha = options.ExpectedSha256,
                    RebuildDir = options.RebuildDir,
                    ComOraclePath = options.ComOraclePath,
                    OwnedWorkbook = ownedPath,
                    LastApply = last,
                    InjectedApply = injected,
                    InitialRead = initialRead,
                    StopBeforeBatchId = options.StopBeforeBatchId,
                });
                scenario["independent"] = independent;
                if (initialRead is not null)
                    scenario["initialRead"] = initialRead;
                scenario["ownedWorkbook"] = ownedPath;
                ownedBooks.Add(ownedPath);
                results.Add(scenario);
                if (JsonUtil.Bool(independent, "ok") != true)
                    throw new InvalidOperationException($"{id}: independent verification failed");
            }

            var afterState = CollectEvidence(client, options, startInventory, bystanderSamples, session, ownedBooks, sourceHashBefore);
            if (JsonUtil.Bool(JsonUtil.Get(afterState, "inventoryCompare"), "ok") != true)
                throw new InvalidOperationException("bystander workbook inventory changed: " + afterState["inventoryCompare"]);
            if (JsonUtil.Bool(afterState, "bystanderContentPreserved") != true)
                throw new InvalidOperationException("bystander workbook content changed or a bystander sample failed: " + afterState["bystanderCompare"]);
            if (JsonUtil.Bool(afterState, "sourceUnchangedDuringRun") != true)
                throw new InvalidOperationException("source hash changed during live run");

            success = new JsonObject
            {
                ["ok"] = true,
                ["protocolSelfTest"] = protocol,
                ["statusBefore"] = Redact.Tokens(statusBefore),
                ["inventoryBefore"] = startInventory.ToJson(),
                ["bystanderSamplesBefore"] = bystanderSamples,
                ["finallyEvidence"] = afterState,
                ["scenarios"] = results,
                ["timings"] = client.TimingLedger(),
                ["documentRef"] = session.DocumentRef,
                ["createdName"] = session.CreatedName,
            };
        }
        catch (Exception ex)
        {
            error = ex;
        }
        finally
        {
            TrackOwned(ownedBooks, session.DocumentRef);
            JsonObject? failureCheckpoint = null;
            if (error is not null && !string.IsNullOrWhiteSpace(session.DocumentRef))
                failureCheckpoint = session.TryCheckpointSave("preserve owned workbook after failure; never auto-close");
            var evidence = CollectEvidence(client, options, startInventory, bystanderSamples, session, ownedBooks, sourceHashBefore);
            if (error is not null)
            {
                success = new JsonObject
                {
                    ["ok"] = false,
                    ["error"] = error.ToString(),
                    ["failureCleanup"] = new JsonObject
                    {
                        ["autoClosed"] = false,
                        ["ownedPreserved"] = ownedBooks.Count > 0,
                        ["checkpointSave"] = failureCheckpoint,
                        ["note"] = "Failure leaves the owned workbook open. Public save is attempted; close_workbook is never emitted here.",
                    },
                    ["protocolSelfTest"] = protocol,
                    ["statusBefore"] = Redact.Tokens(statusBefore),
                    ["inventoryBefore"] = startInventory.ToJson(),
                    ["bystanderSamplesBefore"] = bystanderSamples,
                    ["finallyEvidence"] = evidence,
                    ["scenarios"] = results,
                    ["timings"] = client.TimingLedger(),
                    ["documentRef"] = session.DocumentRef,
                    ["createdName"] = session.CreatedName,
                };
                try
                {
                    JsonUtil.Write(Path.Combine(options.OutputDir, "live-failure.json"), success);
                }
                catch { /* still return the in-memory report */ }
            }
            else if (success is not null)
            {
                success["finallyEvidence"] = evidence;
                if (!EvidenceOk(evidence, out var evidenceError))
                {
                    success["ok"] = false;
                    success["error"] = "CollectEvidence failed after a previously successful report: " + evidenceError;
                }
            }
        }

        return success ?? new JsonObject { ["ok"] = false, ["error"] = "live run produced no report" };
    }

    private static void TrackOwned(List<string> ownedBooks, string? documentRef)
    {
        if (string.IsNullOrWhiteSpace(documentRef)) return;
        if (ownedBooks.Any(p => string.Equals(p, documentRef, StringComparison.OrdinalIgnoreCase))) return;
        ownedBooks.Add(documentRef);
    }

    private static string InjectedPhase(JsonNode response)
    {
        if (JsonUtil.Bool(response, "ok") == true)
            return "unexpected-ok";
        if (!string.IsNullOrWhiteSpace(JsonUtil.Str(response, "confirmToken")))
            return "apply-refused";
        return "preview-refused";
    }

    private static JsonNode SafeCall(PublicClient client, string tool, JsonObject args)
    {
        try { return client.Call(tool, args); }
        catch (Exception ex) { return new JsonObject { ["ok"] = false, ["error"] = ex.Message }; }
    }

    internal static JsonObject SampleBystanders(PublicClient client, BookInventory inventory, string sourceXlsx)
    {
        var samples = new JsonArray();
        foreach (var book in inventory.Books)
        {
            var workbook = book.DocumentRef ?? book.FullName ?? book.Name;
            if (string.IsNullOrWhiteSpace(workbook))
            {
                samples.Add(FailedSample(book.Primary, "", "", "empty workbook identity"));
                continue;
            }

            var isSource = book.Matches(sourceXlsx)
                           || string.Equals(book.FullName, sourceXlsx, StringComparison.OrdinalIgnoreCase)
                           || string.Equals(book.DocumentRef, sourceXlsx, StringComparison.OrdinalIgnoreCase);
            var sheet = isSource && book.Sheets?.Contains("예정공정표", StringComparer.OrdinalIgnoreCase) != false
                ? "예정공정표"
                : book.Sheets?.FirstOrDefault() ?? ResolveSheet(client, workbook);
            var range = isSource ? "BP95:BP96" : "A1:C3";
            if (string.IsNullOrWhiteSpace(sheet))
            {
                samples.Add(FailedSample(book.Primary, workbook, range, "could not resolve a sheet via excel_inspect scope=scan"));
                continue;
            }

            JsonObject read;
            try
            {
                read = PublicApi.ReadRange(client, workbook, sheet, range, formulas: true);
            }
            catch (Exception ex)
            {
                samples.Add(FailedSample(book.Primary, workbook, range, ex.Message, sheet));
                continue;
            }

            var values = PublicApi.ValuesOf(read);
            var captured = PublicApi.HasReadableValues(read) && values is not null;
            samples.Add(new JsonObject
            {
                ["book"] = book.Primary,
                ["workbook"] = workbook,
                ["sheet"] = sheet,
                ["range"] = range,
                ["ok"] = captured,
                ["captured"] = captured,
                ["values"] = captured ? values!.DeepClone() : null,
                ["error"] = captured ? null : (JsonUtil.Str(read, "error") ?? "read returned no values"),
            });
        }
        return new JsonObject
        {
            ["samples"] = samples,
            ["expectedCount"] = inventory.Books.Count,
            ["capturedCount"] = samples.OfType<JsonObject>().Count(s => JsonUtil.Bool(s, "captured") == true),
        };
    }

    private static string? ResolveSheet(PublicClient client, string workbook)
    {
        try
        {
            var scan = PublicApi.Inspect(client, workbook, "scan");
            if (JsonUtil.Bool(scan, "ok") != true) return null;
            return PublicApi.FirstSheetName(scan);
        }
        catch { return null; }
    }

    private static JsonObject FailedSample(string book, string workbook, string range, string error, string? sheet = null) =>
        new()
        {
            ["book"] = book,
            ["workbook"] = workbook,
            ["sheet"] = sheet,
            ["range"] = range,
            ["ok"] = false,
            ["captured"] = false,
            ["values"] = null,
            ["error"] = error,
        };

    private static JsonObject CollectEvidence(
        PublicClient client,
        ProbeOptions options,
        BookInventory startInventory,
        JsonObject samplesBefore,
        ApplySession session,
        IReadOnlyList<string> ownedBooks,
        JsonObject sourceHashBefore)
    {
        JsonNode statusAfter;
        JsonNode ctxAfter;
        try { statusAfter = client.Call("core_get_status", new JsonObject { ["app"] = "excel" }); }
        catch (Exception ex) { statusAfter = new JsonObject { ["ok"] = false, ["error"] = ex.Message }; }
        try { ctxAfter = client.Call("excel_get_active_context", new JsonObject()); }
        catch (Exception ex) { ctxAfter = new JsonObject { ["ok"] = false, ["error"] = ex.Message }; }

        var after = BookInventory.Parse(ctxAfter);
        JsonObject samplesAfter;
        try { samplesAfter = SampleBystanders(client, startInventory, options.SourceXlsx); }
        catch (Exception ex) { samplesAfter = new JsonObject { ["ok"] = false, ["error"] = ex.Message, ["samples"] = new JsonArray() }; }

        JsonObject hash;
        try { hash = SourceIntegrity.Capture(options.SourceXlsx, options.ExpectedSha256); }
        catch (Exception ex) { hash = new JsonObject { ["match"] = false, ["error"] = ex.Message }; }

        var owned = ownedBooks.Count > 0
            ? ownedBooks
            : new[] { session.DocumentRef ?? session.CreatedName ?? "" };
        var inventory = IdentityGuard.CompareBystanders(startInventory, after, owned);
        var content = CompareSamples(samplesBefore, samplesAfter);
        var beforeSha = JsonUtil.Str(sourceHashBefore, "sha256");
        var afterSha = JsonUtil.Str(hash, "sha256");
        var unchangedDuringRun = !string.IsNullOrWhiteSpace(beforeSha)
            && string.Equals(beforeSha, afterSha, StringComparison.OrdinalIgnoreCase);
        return new JsonObject
        {
            ["statusAfter"] = Redact.Tokens(statusAfter),
            ["contextAfter"] = Redact.Tokens(ctxAfter),
            ["inventoryAfter"] = after.ToJson(),
            ["inventoryCompare"] = inventory,
            ["bystanderSamplesAfter"] = samplesAfter,
            ["bystanderCompare"] = content,
            ["bystanderContentPreserved"] = JsonUtil.Bool(content, "ok") == true,
            ["sourceHash"] = hash,
            ["sourceHashBefore"] = sourceHashBefore.DeepClone(),
            ["sourceHashMatchFrozen"] = JsonUtil.Bool(hash, "match") == true,
            ["sourceUnchangedDuringRun"] = unchangedDuringRun,
            ["sourceHashMatch"] = unchangedDuringRun,
        };
    }

    private static JsonObject CompareSamples(JsonObject before, JsonObject after)
    {
        var a = JsonUtil.Get(before, "samples") as JsonArray ?? [];
        var b = JsonUtil.Get(after, "samples") as JsonArray ?? [];
        if (a.Count != b.Count)
            return new JsonObject { ["ok"] = false, ["error"] = $"sample count {a.Count} vs {b.Count}" };

        for (var i = 0; i < a.Count; i++)
        {
            if (a[i] is not JsonObject left || b[i] is not JsonObject right)
                return new JsonObject { ["ok"] = false, ["error"] = $"sample {i} is not an object" };
            if (JsonUtil.Bool(left, "captured") != true || JsonUtil.Bool(right, "captured") != true)
                return new JsonObject
                {
                    ["ok"] = false,
                    ["error"] = $"sample {i} was not captured on both sides; failed reads are not treated as preserved",
                    ["before"] = left.DeepClone(),
                    ["after"] = right.DeepClone(),
                };
            var lv = JsonUtil.Get(left, "values");
            var rv = JsonUtil.Get(right, "values");
            if (lv is null || rv is null)
                return new JsonObject { ["ok"] = false, ["error"] = $"sample {i} has null values; null==null is not preservation" };
            if (lv.ToJsonString(JsonUtil.WriteCompact) != rv.ToJsonString(JsonUtil.WriteCompact))
                return new JsonObject { ["ok"] = false, ["error"] = $"sample {i} values changed" };
        }
        return new JsonObject { ["ok"] = true, ["compared"] = a.Count };
    }

    private static bool EvidenceOk(JsonObject evidence, out string error)
    {
        if (JsonUtil.Bool(JsonUtil.Get(evidence, "inventoryCompare"), "ok") != true)
        {
            error = "bystander workbook inventory changed: " + evidence["inventoryCompare"];
            return false;
        }
        if (JsonUtil.Bool(evidence, "bystanderContentPreserved") != true)
        {
            error = "bystander workbook content changed or a bystander sample failed: " + evidence["bystanderCompare"];
            return false;
        }
        if (JsonUtil.Bool(evidence, "sourceUnchangedDuringRun") != true
            && JsonUtil.Bool(evidence, "sourceHashMatch") != true)
        {
            error = "source hash changed during live run";
            return false;
        }
        error = "";
        return true;
    }
}
