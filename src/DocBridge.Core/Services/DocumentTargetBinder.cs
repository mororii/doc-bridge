using System.Text.Json.Nodes;

namespace DocBridge.Core.Services;

/// <summary>
/// Binds expectedDocumentRef to every explicit op target before execute mutation.
/// Mixed targets and unqualified/missing explicit workbooks are rejected; the
/// active status document is never used to resolve an explicit workbook path.
/// Bare unsaved names are not saved to invent a path.
/// </summary>
public static class DocumentTargetBinder
{
    public sealed record Result(
        bool Ok, string? ExplicitDocumentRef, IReadOnlyList<string> Errors, string? NormalizedExpected = null);

    public static Result BindForExecute(string app, IReadOnlyList<JsonObject> ops, string expectedDocumentRef)
    {
        var errors = new List<string>();
        if (!DocumentIdentity.TryNormalizeExecuteRef(app, expectedDocumentRef, out var expectedNorm, out var identityError))
        {
            errors.Add(identityError ?? DocumentIdentity.LegacyPreviewGuidance);
            return new Result(false, null, errors);
        }

        var explicitRefs = new List<string>();
        var missingExplicit = new List<int>();
        for (var i = 0; i < ops.Count; i++)
        {
            var extracted = ExtractExplicit(app, ops[i], i + 1, errors);
            if (extracted.Ambiguous) return new Result(false, null, errors);
            if (extracted.Absent) missingExplicit.Add(i + 1);
            else if (extracted.Resolved is not null)
            {
                if (!DocumentIdentity.TryNormalizeExecuteRef(app, extracted.Resolved, out var explicitNorm, out var explicitError))
                {
                    errors.Add($"ops[{i + 1}] {explicitError ?? DocumentIdentity.LegacyPreviewGuidance}");
                    return new Result(false, null, errors);
                }
                explicitRefs.Add(explicitNorm);
            }
        }

        if (errors.Count > 0) return new Result(false, null, errors);

        if (explicitRefs.Count > 0 && missingExplicit.Count > 0)
        {
            errors.Add(
                "executionMode=execute rejects mixed document targets: some ops omit an explicit document " +
                $"while others name one (ops without a target: {string.Join(",", missingExplicit)})");
            return new Result(false, null, errors);
        }

        if (explicitRefs.Count == 0)
            return new Result(true, null, errors, expectedNorm);

        var first = explicitRefs[0];
        for (var i = 1; i < explicitRefs.Count; i++)
        {
            if (!DocBridgeHost.SameDocumentRef(app, first, explicitRefs[i]))
            {
                errors.Add(
                    $"executionMode=execute rejects mixed explicit document targets: '{first}' vs '{explicitRefs[i]}'");
                return new Result(false, null, errors);
            }
        }

        if (!DocBridgeHost.SameDocumentRef(app, expectedNorm, first))
        {
            errors.Add(
                $"expectedDocumentRef '{expectedNorm}' does not match explicit op target '{first}'");
            return new Result(false, null, errors);
        }

        return new Result(true, first, errors, expectedNorm);
    }

    /// <summary>
    /// Deep-clone ops and stamp expectedDocumentRef so later active-document
    /// switches cannot retarget omitted ops. Caller JSON is not mutated.
    /// </summary>
    public static List<JsonObject> CloneBoundOps(string app, IReadOnlyList<JsonObject> ops, string expectedNorm)
    {
        var bound = new List<JsonObject>(ops.Count);
        foreach (var op in ops)
        {
            var clone = (JsonObject)op.DeepClone();
            StampExpected(app, clone, expectedNorm);
            bound.Add(clone);
        }
        return bound;
    }

    private static void StampExpected(string app, JsonObject op, string expected)
    {
        if (app.Equals("excel", StringComparison.OrdinalIgnoreCase))
        {
            op["targetWorkbook"] = expected;
            if (Json.GetObj(op, "target") is { } target && target.ContainsKey("workbook"))
                target["workbook"] = expected;
            return;
        }

        if (app.Equals("cad", StringComparison.OrdinalIgnoreCase) ||
            app.Equals("gstarcad", StringComparison.OrdinalIgnoreCase))
        {
            op["document"] = expected;
            return;
        }

        if (app.Equals("hwp", StringComparison.OrdinalIgnoreCase) && !op.ContainsKey("file"))
            op["documentRef"] = expected;
    }

    private readonly record struct Extraction(bool Absent, bool Ambiguous, string? Resolved);

    private static Extraction ExtractExplicit(string app, JsonObject op, int index, List<string> errors)
    {
        if (app.Equals("excel", StringComparison.OrdinalIgnoreCase))
            return ExtractExcel(op, index, errors);
        if (app.Equals("hwp", StringComparison.OrdinalIgnoreCase))
            return ExtractHwp(op, index, errors);
        if (app.Equals("cad", StringComparison.OrdinalIgnoreCase) ||
            app.Equals("gstarcad", StringComparison.OrdinalIgnoreCase))
            return ExtractNamed(op, index, "document", errors, requireFullyQualifiedPath: false);
        return new Extraction(true, false, null);
    }

    private static Extraction ExtractExcel(JsonObject op, int index, List<string> errors)
    {
        var fromRoot = ReadPresentString(op, "targetWorkbook");
        var fromTarget = ReadPresentString(Json.GetObj(op, "target"), "workbook");
        if (!fromRoot.Present && !fromTarget.Present)
            return new Extraction(true, false, null);

        if (fromRoot.Present && fromTarget.Present &&
            !string.Equals(fromRoot.Value, fromTarget.Value, StringComparison.OrdinalIgnoreCase))
        {
            errors.Add($"ops[{index}] has conflicting targetWorkbook and target.workbook");
            return new Extraction(false, true, null);
        }

        var raw = fromRoot.Present ? fromRoot.Value : fromTarget.Value;
        if (string.IsNullOrWhiteSpace(raw))
        {
            errors.Add(
                $"ops[{index}] names an explicit Excel workbook but the value is empty; " +
                "the active workbook is not used to fill it");
            return new Extraction(false, true, null);
        }

        if (!Path.IsPathFullyQualified(raw))
        {
            errors.Add(
                $"ops[{index}] explicit Excel workbook '{raw}' is not a fully qualified path; " +
                "the active workbook is not trusted to resolve it. " +
                DocumentIdentity.LegacyPreviewGuidance);
            return new Extraction(false, true, null);
        }

        try
        {
            return new Extraction(false, false, Path.GetFullPath(raw));
        }
        catch (Exception ex)
        {
            errors.Add($"ops[{index}] explicit Excel workbook is not a usable path: {ex.Message}");
            return new Extraction(false, true, null);
        }
    }

    private static Extraction ExtractHwp(JsonObject op, int index, List<string> errors)
    {
        var documentRef = ReadPresentString(op, "documentRef");
        var file = ReadPresentString(op, "file");
        if (!documentRef.Present && !file.Present)
            return new Extraction(true, false, null);

        if (documentRef.Present && file.Present)
        {
            errors.Add($"ops[{index}] cannot combine file and documentRef");
            return new Extraction(false, true, null);
        }

        if (file.Present)
        {
            if (string.IsNullOrWhiteSpace(file.Value))
            {
                errors.Add($"ops[{index}] file is empty; the active HWP document is not used to fill it");
                return new Extraction(false, true, null);
            }
            if (!Path.IsPathFullyQualified(file.Value))
            {
                errors.Add(
                    $"ops[{index}] file '{file.Value}' is not a fully qualified path; " +
                    "the active document is not trusted to resolve it");
                return new Extraction(false, true, null);
            }
            try { return new Extraction(false, false, Path.GetFullPath(file.Value)); }
            catch (Exception ex)
            {
                errors.Add($"ops[{index}] file is not a usable path: {ex.Message}");
                return new Extraction(false, true, null);
            }
        }

        if (string.IsNullOrWhiteSpace(documentRef.Value))
        {
            errors.Add($"ops[{index}] documentRef is empty; the active HWP document is not used to fill it");
            return new Extraction(false, true, null);
        }
        return new Extraction(false, false, documentRef.Value);
    }

    private static Extraction ExtractNamed(
        JsonObject op, int index, string field, List<string> errors, bool requireFullyQualifiedPath)
    {
        var present = ReadPresentString(op, field);
        if (!present.Present) return new Extraction(true, false, null);
        if (string.IsNullOrWhiteSpace(present.Value))
        {
            errors.Add($"ops[{index}] {field} is empty; the active document is not used to fill it");
            return new Extraction(false, true, null);
        }
        if (requireFullyQualifiedPath && !Path.IsPathFullyQualified(present.Value))
        {
            errors.Add($"ops[{index}] {field} '{present.Value}' is not a fully qualified path");
            return new Extraction(false, true, null);
        }
        if (Path.IsPathFullyQualified(present.Value))
        {
            try { return new Extraction(false, false, Path.GetFullPath(present.Value)); }
            catch (Exception ex)
            {
                errors.Add($"ops[{index}] {field} is not a usable path: {ex.Message}");
                return new Extraction(false, true, null);
            }
        }
        return new Extraction(false, false, present.Value);
    }

    private readonly record struct PresentString(bool Present, string? Value);

    private static PresentString ReadPresentString(JsonObject? obj, string key)
    {
        if (obj is null || !obj.TryGetPropertyValue(key, out var node))
            return new PresentString(false, null);
        if (node is null) return new PresentString(true, null);
        if (node is JsonValue value && value.TryGetValue<string>(out var text))
            return new PresentString(true, text);
        return new PresentString(true, node.ToJsonString());
    }
}
