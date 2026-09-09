using System.Diagnostics.CodeAnalysis;
using System.Globalization;
using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

public sealed partial class CadAdapter
{
    private const int PartialTextBackupLimit = 500;
    private static readonly HashSet<string> ExactScopedSnapshotOps = new(StringComparer.OrdinalIgnoreCase)
    {
        "set_text_value",
        "set_layer_visibility",
        "set_layer_color",
        "regen_document",
    };
    private static readonly HashSet<string> LayerSnapshotFields = new(StringComparer.OrdinalIgnoreCase)
    {
        "fields", "on", "color",
    };
    private static readonly HashSet<string> TextSnapshotFields = new(StringComparer.OrdinalIgnoreCase)
    {
        "fields", "text", "entityName",
    };
    private static readonly HashSet<string> AllowedLayerRestoreFields = new(StringComparer.OrdinalIgnoreCase)
    {
        "on", "color",
    };
    private static readonly HashSet<string> AllowedTextRestoreFields = new(StringComparer.OrdinalIgnoreCase)
    {
        "text",
    };

    private void CaptureOperationSnapshot(string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject>? ops)
    {
        ComInvokeWithRetry(() =>
        {
            var app = AttachCad();
            if (app is null) { metadata["payload"] = "none (autocad not running)"; return; }
            dynamic d = app;
            var doc = ActiveDocWait(d);
            if (doc is null) { metadata["payload"] = "none (no drawing)"; return; }

            var identity = ReadDocumentIdentity(doc);
            TryCopyDrawingBackup(snapshotDir, identity.FullName, metadata);
            if (ops is { Count: > 0 })
            {
                string? documentError = DenyMismatchedScopedDocuments((object)doc, ops);
                if (documentError is not null)
                    throw new InvalidOperationException(documentError);
            }

            var scoped = ops is { Count: > 0 } && IsExactScopedSnapshot(ops);
            JsonObject state;
            if (scoped)
            {
                state = CaptureExactScopedState(doc, ops!, identity);
                metadata["payload"] = "drawing-backup + operation-scoped state.json";
                metadata["snapshotCoverage"] = "complete";
                metadata["snapshotKind"] = "operation-scoped";
            }
            else
            {
                state = CapturePartialLegacyBackup(doc, identity, ops,
                    "geometry, manual, or unsupported ops cannot be fully rolled back from COM state; drawing-backup and partial text/layer backup retained");
                metadata["payload"] = "drawing-backup + incomplete state.json with partial text/layer backup";
                metadata["snapshotCoverage"] = "incomplete";
                metadata["snapshotKind"] = "legacy-geometry-or-unsupported";
            }

            File.WriteAllText(Path.Combine(snapshotDir, "state.json"), state.ToJsonString(Json.Pretty));
            metadata["documentRef"] = string.IsNullOrEmpty(identity.FullName) ? metadata["documentRef"] : identity.FullName;
            metadata["savedAtSnapshot"] = identity.Saved;
            metadata["snapshotStateVersion"] = 3;
            if (ops is { Count: > 0 })
            {
                metadata["operationStateSha256"] = CadOperationStateHash(doc, ops);
                metadata["operationStateVersion"] = 2;
            }
        });
    }

    private JsonObject RestoreOperationSnapshot(string snapshotDir, JsonObject metadata)
    {
        return ComInvokeWithRetry(() =>
        {
            var statePath = Path.Combine(snapshotDir, "state.json");
            if (!File.Exists(statePath)) return Json.ErrorResult("state.json not found in snapshot", App);

            var app = AttachCad();
            if (app is null) return Json.ErrorResult($"{ProductName} not running", App);
            dynamic d = app;
            var doc = ActiveDocWait(d);
            if (doc is null) return Json.ErrorResult("열린 도면이 없습니다", App);

            var state = JsonNode.Parse(File.ReadAllText(statePath)) as JsonObject ?? new JsonObject();
            string? identityError = CheckDocumentIdentity((object)doc, state, requireStableIdentity: ClaimsCompleteScopedState(state));
            if (identityError is not null)
            {
                string identityMessage = identityError;
                return IncompleteRestore(state, metadata, identityMessage, checkedCount: 0,
                    mismatches: new string[] { identityMessage }, regenerated: false, coverage: "incomplete");
            }

            if (ClaimsCompleteScopedState(state))
            {
                if (!TryValidateCompleteScopedState(state, out string? shapeError))
                {
                    string shapeMessage = shapeError ?? "complete snapshot failed shape validation";
                    return IncompleteRestore(state, metadata, shapeMessage, checkedCount: 0,
                        mismatches: new string[] { shapeMessage }, regenerated: false, coverage: "incomplete");
                }

                string? bindError = BindRequestedRestoreTargets((object)doc, state);
                if (bindError is not null)
                {
                    string bindMessage = bindError;
                    return IncompleteRestore(state, metadata, bindMessage, checkedCount: 0,
                        mismatches: new string[] { bindMessage }, regenerated: false, coverage: "incomplete");
                }

                var mismatches = new List<string>();
                var checkedCount = 0;
                RestoreScopedLayers(doc, state, mismatches, ref checkedCount);
                RestoreScopedTexts(doc, state, mismatches, ref checkedCount);
                var regenerated = TryRegen(doc, mismatches);
                var verified = mismatches.Count == 0;
                return new JsonObject
                {
                    ["ok"] = verified,
                    ["restored"] = verified,
                    ["coverage"] = "complete",
                    ["kind"] = "operation-scoped",
                    ["drawingBackup"] = Json.GetString(metadata, "drawingBackup"),
                    ["regenerated"] = regenerated,
                    ["readback"] = new JsonObject
                    {
                        ["verified"] = verified,
                        ["checked"] = checkedCount,
                        ["mismatches"] = Json.ToArray(mismatches),
                    },
                    ["warnings"] = Json.ToArray(verified
                        ? Array.Empty<string>()
                        : new[] { "operation-scoped restore did not verify every requested field" }),
                    ["errors"] = Json.ToArray(mismatches),
                };
            }

            var partialMismatches = new List<string>();
            var partialChecked = 0;
            RestorePartialLegacyBackup(doc, state, partialMismatches, ref partialChecked);
            var partialRegen = TryRegen(doc, partialMismatches);
            return IncompleteRestore(state, metadata,
                Json.GetString(state, "reason") ??
                "snapshot coverage is incomplete; partial text/layer backup may have been applied and full batch rollback is not claimed",
                partialChecked, partialMismatches, partialRegen);
        });
    }

    private static bool IsExactScopedSnapshot(IReadOnlyList<JsonObject> ops) =>
        ops.All(op => ExactScopedSnapshotOps.Contains(Json.GetString(op, "op") ?? ""));

    private JsonObject CaptureExactScopedState(dynamic doc, IReadOnlyList<JsonObject> ops, DocumentIdentity identity)
    {
        var layerFields = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
        var textHandles = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var op in ops)
        {
            var name = Json.GetString(op, "op");
            if (name == "set_layer_visibility")
                AddLayerField(layerFields, Json.GetString(op, "layer"), "on");
            else if (name == "set_layer_color")
                AddLayerField(layerFields, Json.GetString(op, "layer"), "color");
            else if (name == "set_text_value")
            {
                var handle = Json.GetString(op, "handle");
                if (string.IsNullOrWhiteSpace(handle))
                    throw new InvalidOperationException("set_text_value snapshot requires handle");
                textHandles.Add(handle);
            }
        }

        var layers = new JsonObject();
        foreach (var (layerName, fields) in layerFields.OrderBy(kv => kv.Key, StringComparer.OrdinalIgnoreCase))
        {
            dynamic layer;
            try { layer = doc.Layers.Item(layerName); }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"CAD snapshot missing requested layer '{layerName}': {ex.Message}");
            }

            var item = new JsonObject { ["fields"] = Json.ToArray(fields.OrderBy(f => f)) };
            if (fields.Contains("on"))
                item["on"] = (bool)layer.LayerOn;
            if (fields.Contains("color"))
                item["color"] = (int)layer.Color;
            layers[layerName] = item;
        }

        var texts = new JsonObject();
        foreach (var handle in textHandles.OrderBy(h => h, StringComparer.OrdinalIgnoreCase))
        {
            dynamic ent;
            try { ent = doc.HandleToObject(handle); }
            catch (Exception ex)
            {
                throw new InvalidOperationException(
                    $"CAD snapshot missing requested handle '{handle}': {ex.Message}");
            }

            var type = (string)ent.EntityName;
            if (!IsTextLike(type))
                throw new InvalidOperationException($"CAD snapshot handle '{handle}' is {type} (not text-like)");
            texts[handle] = new JsonObject
            {
                ["text"] = TextOf(ent),
                ["fields"] = new JsonArray("text"),
                ["entityName"] = type,
            };
        }

        if (layerFields.Count != layers.Count || textHandles.Count != texts.Count)
            throw new InvalidOperationException("CAD snapshot shape is incomplete for the requested text/layer ops");

        return new JsonObject
        {
            ["version"] = 3,
            ["coverage"] = "complete",
            ["kind"] = "operation-scoped",
            ["complete"] = true,
            ["fullName"] = identity.FullName,
            ["documentName"] = identity.Name,
            ["modelSpaceCount"] = identity.ModelSpaceCount,
            ["ops"] = Json.ToArray(ops.Select(op => Json.GetString(op, "op") ?? "")),
            ["requested"] = new JsonObject
            {
                ["layers"] = Json.ToArray(layerFields.Keys.OrderBy(k => k, StringComparer.OrdinalIgnoreCase)),
                ["texts"] = Json.ToArray(textHandles.OrderBy(h => h, StringComparer.OrdinalIgnoreCase)),
            },
            ["layers"] = layers,
            ["texts"] = texts,
        };
    }

    private JsonObject CapturePartialLegacyBackup(
        dynamic doc, DocumentIdentity identity, IReadOnlyList<JsonObject>? ops, string reason)
    {
        var layers = new JsonObject();
        string? layerError = null;
        try
        {
            foreach (dynamic layer in doc.Layers)
            {
                var name = (string)layer.Name;
                layers[name] = new JsonObject
                {
                    ["on"] = (bool)layer.LayerOn,
                    ["color"] = (int)layer.Color,
                };
            }
        }
        catch (Exception ex) { layerError = ex.Message; }

        var texts = new JsonObject();
        string? textError = null;
        var tcount = 0;
        try
        {
            foreach (dynamic ent in doc.ModelSpace)
            {
                var type = (string)ent.EntityName;
                if (!IsTextLike(type)) continue;
                try { texts[(string)ent.Handle] = TextOf(ent); } catch { }
                if (++tcount >= PartialTextBackupLimit) break;
            }
        }
        catch (Exception ex) { textError = ex.Message; }

        var state = new JsonObject
        {
            ["version"] = 3,
            ["coverage"] = "incomplete",
            ["kind"] = "legacy-geometry-or-unsupported",
            ["complete"] = false,
            ["reason"] = reason,
            ["fullName"] = identity.FullName,
            ["documentName"] = identity.Name,
            ["modelSpaceCount"] = identity.ModelSpaceCount,
            ["ops"] = ops is null ? new JsonArray() : Json.ToArray(ops.Select(op => Json.GetString(op, "op") ?? "")),
            ["layers"] = layers,
            ["texts"] = texts,
        };
        if (layerError is not null) state["layerBackupError"] = layerError;
        if (textError is not null) state["textBackupError"] = textError;
        return state;
    }

    private static void RestoreScopedLayers(dynamic doc, JsonObject state, List<string> mismatches, ref int checkedCount)
    {
        if (Json.GetObj(state, "layers") is not { } layers) return;
        foreach (var (layerName, node) in layers)
        {
            if (node is not JsonObject layerState) continue;
            var fields = RequestedFields(layerState);
            if (fields.Count == 0)
            {
                if (layerState.ContainsKey("on")) fields.Add("on");
                if (layerState.ContainsKey("color")) fields.Add("color");
            }
            try
            {
                dynamic layer = doc.Layers.Item(layerName);
                if (fields.Contains("on"))
                {
                    var wantOn = Json.GetBool(layerState, "on");
                    layer.LayerOn = wantOn;
                    checkedCount++;
                    if ((bool)layer.LayerOn != wantOn)
                        mismatches.Add($"layer {layerName}: LayerOn restore readback mismatch");
                }
                if (fields.Contains("color"))
                {
                    var wantColor = Json.GetInt(layerState, "color");
                    if (wantColor is null)
                    {
                        mismatches.Add($"layer {layerName}: snapshot color field missing");
                        continue;
                    }
                    layer.Color = wantColor.Value;
                    checkedCount++;
                    if ((int)layer.Color != wantColor.Value)
                        mismatches.Add($"layer {layerName}: Color restore readback mismatch");
                }
            }
            catch (Exception ex) { mismatches.Add($"layer {layerName}: {ex.Message}"); }
        }
    }

    private void RestoreScopedTexts(dynamic doc, JsonObject state, List<string> mismatches, ref int checkedCount)
    {
        if (Json.GetObj(state, "texts") is not { } texts) return;
        foreach (var (handle, node) in texts)
        {
            try
            {
                dynamic ent = doc.HandleToObject(handle);
                var want = node is JsonObject obj
                    ? Json.GetString(obj, "text")
                    : node?.GetValue<string>();
                if (want is null)
                {
                    mismatches.Add($"entity {handle}: snapshot text field missing");
                    continue;
                }
                ent.TextString = want;
                checkedCount++;
                if (TextOf(ent) != want)
                    mismatches.Add($"entity {handle}: restore readback mismatch");
            }
            catch (Exception ex) { mismatches.Add($"entity {handle}: {ex.Message}"); }
        }
    }

    private JsonObject IncompleteRestore(
        JsonObject state, JsonObject metadata, string reason, int checkedCount, IEnumerable<string> mismatches, bool regenerated,
        string? coverage = null)
    {
        var errors = mismatches.ToArray();
        var reportedCoverage = coverage
            ?? (string.Equals(Json.GetString(state, "coverage"), "complete", StringComparison.OrdinalIgnoreCase)
                ? "incomplete"
                : Json.GetString(state, "coverage") ?? "incomplete");
        return new JsonObject
        {
            ["ok"] = false,
            ["restored"] = false,
            ["partialApplied"] = checkedCount > 0,
            ["coverage"] = reportedCoverage,
            ["kind"] = Json.GetString(state, "kind") ?? "legacy-geometry-or-unsupported",
            ["complete"] = false,
            ["drawingBackup"] = Json.GetString(metadata, "drawingBackup"),
            ["regenerated"] = regenerated,
            ["readback"] = new JsonObject
            {
                ["verified"] = false,
                ["checked"] = checkedCount,
                ["mismatches"] = Json.ToArray(errors),
            },
            ["warnings"] = Json.ToArray(new[]
            {
                "CAD snapshot restore is partial/incomplete; drawing-backup was retained and full batch rollback is not claimed.",
                reason,
            }),
            ["errors"] = Json.ToArray(errors.Length == 0 ? new[] { reason } : errors),
        };
    }

    private static string? DenyMismatchedScopedDocuments(object usedDoc, IReadOnlyList<JsonObject> ops)
    {
        foreach (var op in ops)
        {
            string? error = DenyMismatchedScopedDocument(usedDoc, op);
            if (error is not null)
                return error;
        }
        return null;
    }

    private static string? DenyMismatchedScopedDocument(object usedDoc, JsonObject op)
    {
        var name = Json.GetString(op, "op");
        if (!ExactScopedSnapshotOps.Contains(name ?? "")) return null;
        var selector = Json.GetString(op, "document");
        if (string.IsNullOrWhiteSpace(selector)) return null;
        if (UsedDocumentMatchesSelector(usedDoc, selector)) return null;
        return $"op.document '{selector}' does not match the CAD document actually used; write denied before mutation";
    }

    private static bool UsedDocumentMatchesSelector(object usedDocObj, string selector)
    {
        dynamic usedDoc = usedDocObj;
        string name = "";
        string fullName = "";
        try { name = (string)(usedDoc.Name ?? ""); } catch { }
        try { fullName = (string)(usedDoc.FullName ?? ""); } catch { }
        return ScopedDocumentIdentityEquals(fullName, name, selector);
    }

    /// <summary>
    /// Scoped op.document / expected identity: exact normalized absolute path, or exact
    /// unsaved Name/FullName. Substring and wildcard matching stay on discovery helpers.
    /// </summary>
    private static bool ScopedDocumentIdentityEquals(string fullName, string name, string selector)
    {
        if (string.IsNullOrWhiteSpace(selector)) return false;
        if (selector.IndexOfAny(Path.GetInvalidPathChars()) >= 0) return false;

        if (IsStableCadDocumentIdentity(selector))
        {
            if (!IsStableCadDocumentIdentity(fullName)) return false;
            string provenSelector = selector;
            string provenFullName = fullName;
            try
            {
                return string.Equals(
                    Path.GetFullPath(provenFullName),
                    Path.GetFullPath(provenSelector),
                    StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        if (selector.Contains('*') || selector.Contains('?')) return false;
        return string.Equals(name, selector, StringComparison.OrdinalIgnoreCase)
            || string.Equals(fullName, selector, StringComparison.OrdinalIgnoreCase);
    }

    private static bool ClaimsCompleteScopedState(JsonObject state) =>
        string.Equals(Json.GetString(state, "coverage"), "complete", StringComparison.OrdinalIgnoreCase) ||
        string.Equals(Json.GetString(state, "kind"), "operation-scoped", StringComparison.OrdinalIgnoreCase) ||
        Json.GetBool(state, "complete");

    private static bool TryValidateCompleteScopedState(JsonObject state, out string? error)
    {
        error = null;
        if (Json.GetInt(state, "version") != 3 ||
            !string.Equals(Json.GetString(state, "kind"), "operation-scoped", StringComparison.Ordinal) ||
            !string.Equals(Json.GetString(state, "coverage"), "complete", StringComparison.Ordinal) ||
            !Json.GetBool(state, "complete"))
        {
            error = "snapshot is not a validated operation-scoped complete state";
            return false;
        }

        if (!IsStableCadDocumentIdentity(Json.GetString(state, "fullName")))
        {
            error = "complete CAD snapshot is missing a stable document identity";
            return false;
        }

        if (Json.GetArr(state, "ops") is not { } ops || ops.Count == 0)
        {
            error = "complete CAD snapshot is missing ops coverage";
            return false;
        }

        var opNames = new List<string>();
        foreach (var node in ops)
        {
            if (node is not JsonValue value || !value.TryGetValue<string>(out var opName) ||
                !ExactScopedSnapshotOps.Contains(opName))
            {
                error = "complete CAD snapshot ops must be the text/layer/regen allowlist only";
                return false;
            }
            opNames.Add(opName);
        }

        if (Json.GetObj(state, "layers") is not { } layers || Json.GetObj(state, "texts") is not { } texts)
        {
            error = "complete CAD snapshot is missing layers or texts objects";
            return false;
        }

        var requested = Json.GetObj(state, "requested") ?? new JsonObject();
        var requestedLayers = StringSet(Json.GetArr(requested, "layers"));
        var requestedTexts = StringSet(Json.GetArr(requested, "texts"));
        if (opNames.Any(op => op is "set_layer_visibility" or "set_layer_color") && requestedLayers.Count == 0)
        {
            error = "complete CAD snapshot requested no layers for layer ops";
            return false;
        }
        if (opNames.Contains("set_text_value", StringComparer.OrdinalIgnoreCase) && requestedTexts.Count == 0)
        {
            error = "complete CAD snapshot requested no text handles for set_text_value";
            return false;
        }

        if (!SameKeys(layers, requestedLayers, "layers", out error)) return false;
        if (!SameKeys(texts, requestedTexts, "texts", out error)) return false;
        if (!ValidateLayerEntries(layers, requestedLayers, out error)) return false;
        if (!ValidateTextEntries(texts, requestedTexts, out error)) return false;
        return true;
    }

    private static string? BindRequestedRestoreTargets(object docObj, JsonObject state)
    {
        dynamic doc = docObj;
        var requested = Json.GetObj(state, "requested") ?? new JsonObject();
        foreach (var layerName in StringSet(Json.GetArr(requested, "layers")))
        {
            try { _ = doc.Layers.Item(layerName); }
            catch (Exception ex)
            {
                return $"CAD restore missing requested layer '{layerName}' before writes: {ex.Message}";
            }
        }

        foreach (var handle in StringSet(Json.GetArr(requested, "texts")))
        {
            try { _ = doc.HandleToObject(handle); }
            catch (Exception ex)
            {
                return $"CAD restore missing requested handle '{handle}' before writes: {ex.Message}";
            }
        }

        return null;
    }

    private void RestorePartialLegacyBackup(dynamic doc, JsonObject state, List<string> mismatches, ref int checkedCount)
    {
        RestoreScopedLayers(doc, state, mismatches, ref checkedCount);
        RestoreScopedTexts(doc, state, mismatches, ref checkedCount);
    }

    private static bool TryRegen(dynamic doc, List<string> mismatches)
    {
        try
        {
            doc.Regen(1);
            return true;
        }
        catch (Exception ex)
        {
            mismatches.Add($"regen after restore failed: {ex.Message}");
            return false;
        }
    }

    private static bool ValidateLayerEntries(JsonObject layers, HashSet<string> requested, out string? error)
    {
        error = null;
        foreach (var (name, node) in layers)
        {
            if (!requested.Contains(name))
            {
                error = $"complete CAD snapshot has unexpected layer '{name}'";
                return false;
            }
            if (node is not JsonObject item)
            {
                error = $"complete CAD snapshot layer '{name}' must be an object";
                return false;
            }
            if (HasUnknownKeys(item, LayerSnapshotFields, out var extra))
            {
                error = $"complete CAD snapshot layer '{name}' has unsupported field '{extra}'";
                return false;
            }
            var fields = RequestedFields(item);
            if (fields.Count == 0 || fields.Any(field => !AllowedLayerRestoreFields.Contains(field)))
            {
                error = $"complete CAD snapshot layer '{name}' fields must be on/color only";
                return false;
            }
            if (fields.Contains("on") &&
                (item["on"] is not JsonValue onValue || !onValue.TryGetValue<bool>(out _)))
            {
                error = $"complete CAD snapshot layer '{name}' on must be boolean";
                return false;
            }
            if (fields.Contains("color") && Json.GetInt(item, "color") is null)
            {
                error = $"complete CAD snapshot layer '{name}' color must be an integer";
                return false;
            }
        }
        return true;
    }

    private static bool ValidateTextEntries(JsonObject texts, HashSet<string> requested, out string? error)
    {
        error = null;
        foreach (var (handle, node) in texts)
        {
            if (!requested.Contains(handle))
            {
                error = $"complete CAD snapshot has unexpected handle '{handle}'";
                return false;
            }
            if (node is not JsonObject item)
            {
                error = $"complete CAD snapshot text '{handle}' must be an object";
                return false;
            }
            if (HasUnknownKeys(item, TextSnapshotFields, out var extra))
            {
                error = $"complete CAD snapshot text '{handle}' has unsupported field '{extra}'";
                return false;
            }
            var fields = RequestedFields(item);
            if (fields.Count == 0) fields.Add("text");
            if (fields.Any(field => !AllowedTextRestoreFields.Contains(field)))
            {
                error = $"complete CAD snapshot text '{handle}' fields must be text only";
                return false;
            }
            if (item["text"] is not JsonValue textValue || !textValue.TryGetValue<string>(out _))
            {
                error = $"complete CAD snapshot text '{handle}' text must be a string";
                return false;
            }
        }
        return true;
    }

    private static bool SameKeys(JsonObject map, HashSet<string> requested, string label, out string? error)
    {
        error = null;
        if (map.Count != requested.Count || requested.Any(key => !map.ContainsKey(key)))
        {
            error = $"complete CAD snapshot {label} keys do not match requested coverage";
            return false;
        }
        return true;
    }

    private static bool HasUnknownKeys(JsonObject item, HashSet<string> allow, out string? extra)
    {
        extra = item.Select(kv => kv.Key).FirstOrDefault(key => !allow.Contains(key));
        return extra is not null;
    }

    private static HashSet<string> StringSet(JsonArray? values)
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (values is null) return set;
        foreach (var node in values)
            if (node is JsonValue value && value.TryGetValue<string>(out var text) && !string.IsNullOrWhiteSpace(text))
                set.Add(text);
        return set;
    }

    private static string? CheckDocumentIdentity(object docObj, JsonObject state, bool requireStableIdentity)
    {
        dynamic doc = docObj;
        var expected = Json.GetString(state, "fullName");
        if (!IsStableCadDocumentIdentity(expected))
        {
            if (string.IsNullOrWhiteSpace(expected) && !requireStableIdentity)
                return null;
            return string.IsNullOrWhiteSpace(expected)
                ? "CAD snapshot is missing a stable document identity; restore refused"
                : "CAD snapshot document identity is ambiguous; restore refused (no autosave)";
        }

        string provenExpected = expected;
        string current = "";
        try { current = (string)(doc.FullName ?? ""); } catch { }
        if (!IsStableCadDocumentIdentity(current))
            return "active CAD document identity is missing or ambiguous; restore refused (no autosave)";
        string provenCurrent = current;
        if (!string.Equals(Path.GetFullPath(provenCurrent), Path.GetFullPath(provenExpected), StringComparison.OrdinalIgnoreCase))
            return $"현재 도면 '{provenCurrent}'가 스냅샷 도면 '{provenExpected}'와 다릅니다.";
        return null;
    }

    private static bool IsStableCadDocumentIdentity([NotNullWhen(true)] string? value)
    {
        if (string.IsNullOrWhiteSpace(value)) return false;
        if (value.IndexOfAny(Path.GetInvalidPathChars()) >= 0) return false;
        return Path.IsPathFullyQualified(value);
    }

    private static void TryCopyDrawingBackup(string snapshotDir, string fullName, JsonObject metadata)
    {
        if (string.IsNullOrEmpty(fullName) || !File.Exists(fullName)) return;
        var dest = Path.Combine(snapshotDir, "drawing-backup" + Path.GetExtension(fullName));
        try
        {
            using var src = new FileStream(fullName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var dst = new FileStream(dest, FileMode.Create, FileAccess.Write);
            src.CopyTo(dst);
            metadata["drawingBackup"] = Path.GetFileName(dest);
            metadata["fileSha256"] = CadFileHash(dest);
        }
        catch (Exception ex) { metadata["drawingBackupError"] = ex.Message; }
    }

    private static DocumentIdentity ReadDocumentIdentity(dynamic doc)
    {
        string fullName = "";
        string name = "";
        var saved = false;
        int? count = null;
        try { fullName = (string)(doc.FullName ?? ""); } catch { }
        try { name = (string)(doc.Name ?? ""); } catch { }
        try { saved = (bool)doc.Saved; } catch { }
        try { count = Convert.ToInt32(doc.ModelSpace.Count, CultureInfo.InvariantCulture); } catch { }
        return new DocumentIdentity(fullName, name, saved, count);
    }

    private static void AddLayerField(Dictionary<string, HashSet<string>> layerFields, string? layer, string field)
    {
        if (string.IsNullOrWhiteSpace(layer))
            throw new InvalidOperationException($"CAD snapshot requires layer for {field}");
        if (!layerFields.TryGetValue(layer, out var fields))
        {
            fields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            layerFields[layer] = fields;
        }
        fields.Add(field);
    }

    private static HashSet<string> RequestedFields(JsonObject layerState)
    {
        var fields = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (Json.GetArr(layerState, "fields") is { } listed)
        {
            foreach (var node in listed)
                if (node is JsonValue value && value.TryGetValue<string>(out var field))
                    fields.Add(field);
        }
        return fields;
    }

    private readonly record struct DocumentIdentity(string FullName, string Name, bool Saved, int? ModelSpaceCount);
}
