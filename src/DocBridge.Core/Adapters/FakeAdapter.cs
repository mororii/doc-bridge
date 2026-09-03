using System.Text;
using System.Text.Json.Nodes;
using DocBridge.Core.Models;
using DocBridge.Core.Services;

namespace DocBridge.Core.Adapters;

/// <summary>
/// M0 검증용 인메모리 가짜 문서 어댑터.
/// 시트 기반 셀 표(dict) + 텍스트 본문을 가진 간이 문서 모델로
/// read/dry-run/confirm/apply/snapshot/restore 전체 흐름을 테스트한다.
/// </summary>
public sealed class FakeAdapter : IAppAdapter, IPreviewReuseAdapter
{
    public string App => "fake";

    /// <summary>시트명 → (셀주소 → 값)</summary>
    public Dictionary<string, Dictionary<string, string>> Sheets = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Sheet1"] = new(StringComparer.OrdinalIgnoreCase)
        {
            ["A1"] = "이름", ["B1"] = "금액",
            ["A2"] = "사과", ["B2"] = "1000",
            ["A3"] = "배",   ["B3"] = "2000",
        },
    };

    public string BodyText = "안녕하세요. 이것은 테스트 문서입니다.\n두 번째 문단입니다.";
    public string Selection { get; set; } = "Sheet1!A1:B3";
    public string DocumentRef { get; set; } = "fake-document";
    public int PreviewCallCount { get; private set; }
    public int StatusCallCount { get; private set; }

    public JsonObject CaptureState()
    {
        var sheets = new JsonObject();
        foreach (var (sheet, cells) in Sheets)
        {
            var c = new JsonObject();
            foreach (var (addr, val) in cells) c[addr] = val;
            sheets[sheet] = c;
        }
        return new JsonObject { ["sheets"] = sheets, ["bodyText"] = BodyText };
    }

    public void LoadState(JsonObject state)
    {
        Sheets = new(StringComparer.OrdinalIgnoreCase);
        if (Json.GetObj(state, "sheets") is { } sheets)
            foreach (var (sheet, node) in sheets)
            {
                var cells = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                if (node is JsonObject co)
                    foreach (var (addr, val) in co)
                        cells[addr] = val?.GetValue<string>() ?? "";
                Sheets[sheet] = cells;
            }
        BodyText = Json.GetString(state, "bodyText") ?? "";
    }

    public AdapterStatus GetStatus()
    {
        StatusCallCount++;
        return new(true, true, "fake", "1.0", DocumentRef, "in-memory adapter");
    }

    public JsonObject GetCapabilities() => new()
    {
        ["app"] = App,
        ["automation"] = "in-memory",
        ["directAppControl"] = false,
        ["readOps"] = new JsonArray("context", "range"),
        ["writeOps"] = new JsonArray("set_values", "find_replace", "insert_text", "append_text", "insert_before_text", "insert_after_text", "delete_entities"),
        ["limits"] = new JsonObject(),
    };

    public ContextResult GetActiveContext()
    {
        var r = new ContextResult { Ok = true, App = App, DocumentRef = "fake://document" };
        var sheets = new JsonArray();
        foreach (var s in Sheets.Keys) sheets.Add(s);
        r.Summary["sheets"] = sheets;
        r.Summary["usedRange"] = "Sheet1!A1:B3";
        r.Summary["paragraphChars"] = BodyText.Length;
        r.Selection = new JsonObject { ["ref"] = Selection, ["text"] = BodyText[..Math.Min(40, BodyText.Length)] };
        return r;
    }

    public JsonObject Read(JsonObject args)
    {
        var range = Json.GetString(args, "range") ?? "Sheet1!A1:B3";
        var (sheet, cells) = ResolveRange(range);
        var values = new JsonArray();
        foreach (var row in cells)
        {
            var jr = new JsonArray();
            foreach (var v in row) jr.Add(v);
            values.Add(jr);
        }
        return new JsonObject
        {
            ["ok"] = true, ["app"] = App,
            ["sheet"] = sheet, ["range"] = range, ["values"] = values,
        };
    }

    private (string Sheet, List<List<string>> Rows) ResolveRange(string range)
    {
        // "Sheet1!A1:B3" 형태만 지원 (fake용 최소 파서)
        var sheet = "Sheet1";
        var span = range;
        var bang = range.IndexOf('!');
        if (bang >= 0) { sheet = range[..bang]; span = range[(bang + 1)..]; }
        var parts = span.Split(':');
        var (c1, r1) = ParseCell(parts[0]);
        var (c2, r2) = parts.Length > 1 ? ParseCell(parts[1]) : (c1, r1);
        var rows = new List<List<string>>();
        Sheets.TryGetValue(sheet, out var table);
        for (var r = Math.Min(r1, r2); r <= Math.Max(r1, r2); r++)
        {
            var row = new List<string>();
            for (var c = Math.Min(c1, c2); c <= Math.Max(c1, c2); c++)
            {
                var addr = $"{ColName(c)}{r}";
                row.Add(table is not null && table.TryGetValue(addr, out var v) ? v : "");
            }
            rows.Add(row);
        }
        return (sheet, rows);
    }

    private static (int Col, int Row) ParseCell(string a)
    {
        var i = 0;
        while (i < a.Length && char.IsLetter(a[i])) i++;
        var col = 0;
        foreach (var ch in a[..i].ToUpperInvariant()) col = col * 26 + (ch - 'A' + 1);
        return (col, int.Parse(a[i..]));
    }

    private static string ColName(int c)
    {
        var sb = new StringBuilder();
        while (c > 0) { var m = (c - 1) % 26; sb.Insert(0, (char)('A' + m)); c = (c - 1) / 26; }
        return sb.ToString();
    }

    public ApplyPreview Preview(IReadOnlyList<JsonObject> ops)
    {
        PreviewCallCount++;
        var p = new ApplyPreview();
        var maxDiff = 100;
        foreach (var op in ops)
        {
            var name = Json.GetString(op, "op")!;
            switch (name)
            {
                case "set_values":
                {
                    var range = Json.GetString(op, "range")!;
                    var values = Json.GetArr(op, "values")!;
                    var (sheet, rows) = ResolveRange(range);
                    var bang = range.IndexOf('!');
                    var span = bang >= 0 ? range[(bang + 1)..] : range;
                    var (c1, r1) = ParseCell(span.Split(':')[0]);
                    p.Affected.Add(new AffectedRef("range", $"{sheet}!{span}"));
                    var i = 0;
                    foreach (var rowNode in values)
                    {
                        if (rowNode is not JsonArray rowArr) continue;
                        var j = 0;
                        foreach (var cellNode in rowArr)
                        {
                            var addr = $"{ColName(c1 + j)}{r1 + i}";
                            Sheets.TryGetValue(sheet, out var table);
                            var before = table is not null && table.TryGetValue(addr, out var b) ? b : "";
                            var after = cellNode?.GetValue<string>() ?? "";
                            if (p.Diff.Count < maxDiff)
                                p.Diff.Add(new DiffEntry { Ref = $"{sheet}!{addr}", Before = before, After = after });
                            else p.DiffTruncated = true;
                            j++;
                        }
                        i++;
                    }
                    break;
                }
                case "find_replace":
                {
                    var find = Json.GetString(op, "find")!;
                    var replace = Json.GetString(op, "replace")!;
                    var matchCase = Json.GetBool(Json.GetObj(op, "options"), "matchCase");
                    var cmp = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                    var count = 0;
                    foreach (var (sheet, table) in Sheets)
                        foreach (var (addr, val) in table)
                            if (val.Contains(find, cmp))
                            {
                                count++;
                                if (p.Diff.Count < maxDiff)
                                    p.Diff.Add(new DiffEntry { Ref = $"{sheet}!{addr}", Before = val, After = val.Replace(find, replace, cmp) });
                            }
                    var bodyCount = CountOccurrences(BodyText, find, cmp);
                    if (bodyCount > 0 && p.Diff.Count < maxDiff)
                        p.Diff.Add(new DiffEntry { Ref = "body", Before = find, After = replace });
                    p.Affected.Add(new AffectedRef("matches", $"{count + bodyCount} cell/body matches"));
                    break;
                }
                case "insert_text":
                {
                    var text = Json.GetString(op, "text")!;
                    p.Affected.Add(new AffectedRef("body", "append"));
                    p.Diff.Add(new DiffEntry { Ref = "body", Before = BodyText[^Math.Min(20, BodyText.Length)..], After = text });
                    break;
                }
                case "insert_before_text":
                case "insert_after_text":
                {
                    var anchor = Json.GetString(op, "anchor")!;
                    var text = Json.GetString(op, "text")!;
                    var cmp = Json.GetBool(op, "matchCase", true)
                        ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                    var count = CountOccurrences(BodyText, anchor, cmp);
                    var occurrence = Json.GetInt(op, "occurrence") ?? 1;
                    if (count == 0) p.Errors.Add($"{name}: anchor not found");
                    else if (!op.ContainsKey("occurrence") && count != 1)
                        p.Errors.Add($"{name}: anchor has {count} matches; occurrence is required");
                    else if (occurrence < 1 || occurrence > count)
                        p.Errors.Add($"{name}: occurrence outside 1..{count}");
                    else
                    {
                        p.Affected.Add(new AffectedRef("body-anchor", $"{name} occurrence {occurrence}/{count}"));
                        p.Diff.Add(new DiffEntry { Ref = "body", Before = anchor, After = name == "insert_before_text" ? text + anchor : anchor + text });
                    }
                    break;
                }
                case "delete_entities":
                {
                    p.RequiresHighRiskApproval = true;
                    var handles = Json.GetArr(op, "handles")!;
                    p.Affected.Add(new AffectedRef("entities", $"{handles.Count} entities"));
                    break;
                }
            }
        }
        return p;
    }

    private static int CountOccurrences(string hay, string needle, StringComparison cmp)
    {
        var n = 0; var idx = 0;
        while ((idx = hay.IndexOf(needle, idx, cmp)) >= 0) { n++; idx += needle.Length; }
        return n;
    }

    public ApplyExecution Apply(IReadOnlyList<JsonObject> ops, string snapshotId)
    {
        var exec = new ApplyExecution { Ok = true };
        var checkedCount = 0;
        var mismatches = new List<string>();

        foreach (var op in ops)
        {
            var name = Json.GetString(op, "op")!;
            switch (name)
            {
                case "set_values":
                {
                    var range = Json.GetString(op, "range")!;
                    var values = Json.GetArr(op, "values")!;
                    var bang = range.IndexOf('!');
                    var sheet = bang >= 0 ? range[..bang] : "Sheet1";
                    var span = bang >= 0 ? range[(bang + 1)..] : range;
                    var (c1, r1) = ParseCell(span.Split(':')[0]);
                    if (!Sheets.TryGetValue(sheet, out var table))
                        Sheets[sheet] = table = new(StringComparer.OrdinalIgnoreCase);
                    var expected = new Dictionary<string, string>();
                    var i = 0;
                    foreach (var rowNode in values)
                    {
                        if (rowNode is not JsonArray rowArr) continue;
                        var j = 0;
                        foreach (var cellNode in rowArr)
                        {
                            var addr = $"{ColName(c1 + j)}{r1 + i}";
                            var v = cellNode?.GetValue<string>() ?? "";
                            table[addr] = v;
                            expected[$"{sheet}!{addr}"] = v;
                            exec.Affected.Add(new AffectedRef("range", $"{sheet}!{addr}"));
                            j++;
                        }
                        i++;
                    }
                    // readback 검증
                    foreach (var (refAddr, want) in expected)
                    {
                        checkedCount++;
                        var addrOnly = refAddr[(refAddr.IndexOf('!') + 1)..];
                        if (!table.TryGetValue(addrOnly, out var got) || got != want)
                            mismatches.Add($"{refAddr}: expected '{want}', got '{got}'");
                    }
                    break;
                }
                case "find_replace":
                {
                    var find = Json.GetString(op, "find")!;
                    var replace = Json.GetString(op, "replace")!;
                    var matchCase = Json.GetBool(Json.GetObj(op, "options"), "matchCase");
                    var cmp = matchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                    var replaced = 0;
                    foreach (var (sheet, table) in Sheets)
                    {
                        foreach (var key in table.Keys.ToList())
                        {
                            var val = table[key];
                            if (val.Contains(find, cmp))
                            {
                                var nv = val.Replace(find, replace, cmp);
                                table[key] = nv;
                                replaced++;
                                exec.Diff.Add(new DiffEntry { Ref = $"{sheet}!{key}", Before = val, After = nv });
                            }
                        }
                    }
                    BodyText = BodyText.Replace(find, replace, cmp);
                    exec.Affected.Add(new AffectedRef("matches", $"{replaced} cells replaced"));
                    checkedCount++;
                    break;
                }
                case "insert_text":
                case "append_text":
                {
                    var text = Json.GetString(op, "text")!;
                    if (name == "append_text" && Json.GetBool(op, "startNewParagraph", true) &&
                        BodyText.Length > 0 && !BodyText.EndsWith('\n') && !text.StartsWith('\n'))
                        BodyText += "\n";
                    BodyText += text;
                    exec.Affected.Add(new AffectedRef("body", name));
                    checkedCount++;
                    if (!BodyText.EndsWith(text, StringComparison.Ordinal))
                        mismatches.Add($"body: {name} readback failed");
                    break;
                }
                case "insert_before_text":
                case "insert_after_text":
                {
                    var anchor = Json.GetString(op, "anchor")!;
                    var text = Json.GetString(op, "text")!;
                    var cmp = Json.GetBool(op, "matchCase", true)
                        ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;
                    var count = CountOccurrences(BodyText, anchor, cmp);
                    var occurrence = Json.GetInt(op, "occurrence") ?? 1;
                    if (count == 0 || (!op.ContainsKey("occurrence") && count != 1) || occurrence < 1 || occurrence > count)
                    {
                        mismatches.Add($"{name}: invalid anchor/occurrence");
                        break;
                    }
                    var index = 0;
                    for (var i = 0; i < occurrence; i++)
                    {
                        index = BodyText.IndexOf(anchor, index, cmp);
                        if (i < occurrence - 1) index += anchor.Length;
                    }
                    var before = name == "insert_before_text";
                    var mode = (Json.GetString(op, "mode") ?? "paragraph").ToLowerInvariant();
                    if (mode == "inline")
                    {
                        var insertAt = before ? index : index + anchor.Length;
                        BodyText = BodyText.Insert(insertAt, text);
                    }
                    else if (mode == "paragraph")
                    {
                        if (before)
                        {
                            var lineStart = BodyText.LastIndexOf('\n', Math.Max(0, index - 1));
                            lineStart = lineStart < 0 ? 0 : lineStart + 1;
                            var separator = text.EndsWith('\n') ? "" : "\n";
                            BodyText = BodyText.Insert(lineStart, text + separator);
                        }
                        else
                        {
                            var lineEnd = BodyText.IndexOf('\n', index + anchor.Length);
                            if (lineEnd < 0) lineEnd = BodyText.Length;
                            var separator = text.StartsWith('\n') ? "" : "\n";
                            BodyText = BodyText.Insert(lineEnd, separator + text);
                        }
                    }
                    else
                    {
                        mismatches.Add($"{name}: invalid mode");
                        break;
                    }
                    checkedCount++;
                    exec.Affected.Add(new AffectedRef("body-anchor", name));
                    if (!BodyText.Contains(text, StringComparison.Ordinal)) mismatches.Add($"{name}: readback failed");
                    break;
                }
                case "delete_entities":
                {
                    // fake에는 엔티티가 없으므로 no-op + 경고
                    exec.Warnings.Add("fake adapter: delete_entities is a no-op (high-risk flow exercised)");
                    break;
                }
            }
        }

        exec.Readback = new JsonObject
        {
            ["verified"] = mismatches.Count == 0,
            ["checked"] = checkedCount,
            ["mismatches"] = Json.ToArray(mismatches),
        };
        exec.Ok = mismatches.Count == 0;
        if (mismatches.Count > 0) exec.Errors.AddRange(mismatches);
        return exec;
    }

    public void CaptureSnapshot(string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject>? ops = null)
    {
        var state = CaptureState();
        File.WriteAllText(Path.Combine(snapshotDir, "state.json"), state.ToJsonString(Json.Pretty));
        metadata["payload"] = "state.json";
    }

    public JsonObject ValidatePreviewReuse(
        string snapshotDir, JsonObject metadata, IReadOnlyList<JsonObject> ops)
    {
        var statePath = Path.Combine(snapshotDir, "state.json");
        if (!File.Exists(statePath))
            return new JsonObject { ["ok"] = true, ["reusable"] = false, ["reason"] = "snapshot state missing" };
        var snapshot = JsonNode.Parse(File.ReadAllText(statePath)) as JsonObject;
        if (snapshot is null)
            return new JsonObject { ["ok"] = true, ["reusable"] = false, ["reason"] = "snapshot state invalid" };
        var reusable = string.Equals(Json.Canonical(snapshot), Json.Canonical(CaptureState()), StringComparison.Ordinal);
        return new JsonObject
        {
            ["ok"] = true,
            ["reusable"] = reusable,
            ["fingerprintMethod"] = "fake-full-state",
            ["reason"] = reusable ? "snapshot fingerprint matched" : "document state changed after dry-run",
        };
    }

    public JsonObject RestoreSnapshot(string snapshotDir, JsonObject metadata)
    {
        var statePath = Path.Combine(snapshotDir, "state.json");
        if (!File.Exists(statePath))
            return new JsonObject { ["ok"] = false, ["errors"] = Json.ToArray(new[] { "state.json not found in snapshot" }) };

        var state = JsonNode.Parse(File.ReadAllText(statePath)) as JsonObject ?? new JsonObject();
        var beforeHash = Convert.ToHexString(
            System.Security.Cryptography.SHA256.HashData(Encoding.UTF8.GetBytes(CaptureState().ToJsonString(Json.Compact))));
        LoadState(state);
        var afterHash = Convert.ToHexString(
            System.Security.Cryptography.SHA256.HashData(Encoding.UTF8.GetBytes(CaptureState().ToJsonString(Json.Compact))));
        var snapshotHash = Convert.ToHexString(
            System.Security.Cryptography.SHA256.HashData(Encoding.UTF8.GetBytes(state.ToJsonString(Json.Compact))));

        return new JsonObject
        {
            ["ok"] = true,
            ["restored"] = true,
            ["readback"] = new JsonObject
            {
                ["verified"] = afterHash == snapshotHash,
                ["beforeHash"] = beforeHash,
                ["afterHash"] = afterHash,
            },
        };
    }

    public void Dispose() { }
}
