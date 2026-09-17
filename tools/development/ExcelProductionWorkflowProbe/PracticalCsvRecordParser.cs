namespace DocBridge.Development.ExcelProductionWorkflowProbe;

/// <summary>
/// Independent RFC4180 CSV oracle for E9 checks. Read-only: never feed these
/// records into set_values on the acceptance sheet. Product import_csv uses
/// ExcelCsvContract.Parse; this parser is expected-output comparison only.
/// </summary>
internal static class PracticalCsvRecordParser
{
    public const string HeaderLine = "id,name,qty,note,sku,literal,packed";

    public static string CanonicalSample()
    {
        // Duplicate first data record so remove_duplicates yields 2 unique rows.
        // note contains quoted CRLF; name contains quoted comma; sku keeps leading zeros;
        // literal is a formula-prefixed string, not a formula.
        return HeaderLine + "\r\n" +
               "1,\"PVC, pipe\",10,\"누수\r\n점검\",\"0012\",\"=1+1\",\"1|PVC, pipe\"\r\n" +
               "1,\"PVC, pipe\",10,\"누수\r\n점검\",\"0012\",\"=1+1\",\"1|PVC, pipe\"\r\n" +
               "2,\"맨홀\",3,\"양호\",\"0003\",\"=2+2\",\"2|맨홀\"\r\n";
    }

    public static List<string[]> Parse(string text)
    {
        if (text is null) throw new ArgumentNullException(nameof(text));
        var rows = new List<string[]>();
        var fields = new List<string>();
        var sb = new StringBuilder();
        var inQuotes = false;
        var i = 0;
        while (i < text.Length)
        {
            var c = text[i];
            if (inQuotes)
            {
                if (c == '"')
                {
                    if (i + 1 < text.Length && text[i + 1] == '"')
                    {
                        sb.Append('"');
                        i += 2;
                        continue;
                    }
                    inQuotes = false;
                    i++;
                    continue;
                }
                sb.Append(c);
                i++;
                continue;
            }

            switch (c)
            {
                case '"':
                    inQuotes = true;
                    i++;
                    break;
                case ',':
                    fields.Add(sb.ToString());
                    sb.Clear();
                    i++;
                    break;
                case '\r':
                    if (i + 1 < text.Length && text[i + 1] == '\n') i++;
                    FlushRow(rows, fields, sb);
                    i++;
                    break;
                case '\n':
                    FlushRow(rows, fields, sb);
                    i++;
                    break;
                default:
                    sb.Append(c);
                    i++;
                    break;
            }
        }

        if (inQuotes)
            throw new FormatException("CSV ended inside a quoted field");
        if (sb.Length > 0 || fields.Count > 0)
            FlushRow(rows, fields, sb);
        return rows;
    }

    public static string Write(IEnumerable<string[]> rows)
    {
        var sb = new StringBuilder();
        foreach (var row in rows)
        {
            for (var i = 0; i < row.Length; i++)
            {
                if (i > 0) sb.Append(',');
                sb.Append(Escape(row[i] ?? ""));
            }
            sb.Append("\r\n");
        }
        return sb.ToString();
    }

    public static string Escape(string value)
    {
        var mustQuote = value.Contains(',') || value.Contains('"') ||
                        value.Contains('\n') || value.Contains('\r');
        if (!mustQuote) return value;
        return "\"" + value.Replace("\"", "\"\"") + "\"";
    }

    public static List<string[]> UniqueDataRows(IReadOnlyList<string[]> rows, StringComparer? comparer = null)
    {
        comparer ??= StringComparer.Ordinal;
        var seen = new HashSet<string>(comparer);
        var unique = new List<string[]>();
        for (var i = 1; i < rows.Count; i++)
        {
            var key = string.Join('\u0001', rows[i]);
            if (seen.Add(key))
                unique.Add(rows[i]);
        }
        return unique;
    }

    public static JsonObject Describe(IReadOnlyList<string[]> rows)
    {
        var data = UniqueDataRows(rows);
        return new JsonObject
        {
            ["recordCount"] = rows.Count,
            ["headerCount"] = rows.Count == 0 ? 0 : 1,
            ["dataRowsIncludingDuplicates"] = Math.Max(0, rows.Count - 1),
            ["uniqueDataRows"] = data.Count,
            ["hasQuotedComma"] = rows.Any(r => r.Any(f => f.Contains(','))),
            ["hasQuotedCrlf"] = rows.Any(r => r.Any(f => f.Contains('\r') || f.Contains('\n'))),
            ["hasKorean"] = rows.Any(r => r.Any(ContainsHangul)),
            ["leadingZeroStrings"] = new JsonArray(data.Select(r => Field(r, 4)).Where(s => s.StartsWith('0')).Select(s => JsonValue.Create(s)).ToArray()),
            ["formulaPrefixedLiterals"] = new JsonArray(data.Select(r => Field(r, 5)).Where(s => s.StartsWith('=')).Select(s => JsonValue.Create(s)).ToArray()),
        };
    }

    public static JsonObject SelfCheck()
    {
        var errors = new JsonArray();
        var rows = Parse(CanonicalSample());
        if (rows.Count != 4) errors.Add($"canonical records {rows.Count} != 4");
        if (rows[0][0] != "id" || rows[0][6] != "packed") errors.Add("header");
        if (rows[1][1] != "PVC, pipe") errors.Add("quoted comma not preserved");
        if (rows[1][3] != "누수\r\n점검") errors.Add("quoted CRLF not preserved");
        if (rows[3][1] != "맨홀") errors.Add("Korean field");
        if (rows[1][4] != "0012") errors.Add("leading-zero sku must stay 0012");
        if (rows[1][5] != "=1+1") errors.Add("formula-prefixed literal must stay =1+1");
        if (rows[3][4] != "0003") errors.Add("second leading-zero sku");
        if (rows[3][5] != "=2+2") errors.Add("second formula-prefixed literal");
        var unique = UniqueDataRows(rows);
        if (unique.Count != 2) errors.Add($"unique data rows {unique.Count} != 2");

        var oracle = OracleFromRows(rows);
        var assert = JsonUtil.Get(oracle, "assertImportedThenCleaned") as JsonObject;
        if (JsonUtil.Str(assert, "B2") != "PVC, pipe") errors.Add("oracle B2 quoted comma");
        if (JsonUtil.Str(assert, "D2") != "누수\r\n점검") errors.Add("oracle D2 quoted CRLF");
        if (JsonUtil.Str(assert, "E2") != "0012") errors.Add("oracle E2 leading-zero");
        if (JsonUtil.Str(assert, "F2") != "=1+1") errors.Add("oracle F2 formula-prefixed literal");
        if (JsonUtil.Str(assert, "B3") != "맨홀") errors.Add("oracle B3 Korean");
        if (JsonUtil.Str(assert, "E3") != "0003") errors.Add("oracle E3 second leading-zero");
        if (JsonUtil.Str(assert, "F3") != "=2+2") errors.Add("oracle F3 second literal");

        var escaped = Parse("a,\"b\"\"c\",d\r\n");
        if (escaped.Count != 1 || escaped[0][1] != "b\"c") errors.Add("escaped quote");

        var roundTrip = Parse(Write(rows));
        if (roundTrip.Count != rows.Count) errors.Add("round-trip count");
        else
        {
            for (var i = 0; i < rows.Count; i++)
            {
                if (!rows[i].SequenceEqual(roundTrip[i]))
                    errors.Add("round-trip row " + i);
            }
        }

        return new JsonObject
        {
            ["ok"] = errors.Count == 0,
            ["errors"] = errors,
            ["canonical"] = Describe(rows),
        };
    }

    public static string WriteSampleFile(string directory, string fileName = "e9-sample.csv")
    {
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, fileName);
        File.WriteAllText(path, CanonicalSample(), new UTF8Encoding(false));
        return path;
    }

    public static JsonObject Oracle(string csvPath) =>
        OracleFromRows(Parse(File.ReadAllText(csvPath, new UTF8Encoding(false))));

    public static JsonObject OracleFromCanonical() => OracleFromRows(Parse(CanonicalSample()));

    public static JsonObject OracleFromRows(IReadOnlyList<string[]> rows)
    {
        var desc = Describe(rows);
        var unique = UniqueDataRows(rows);
        unique.Sort((a, b) => string.CompareOrdinal(Field(a, 0), Field(b, 0)));
        var cleaned = new List<string[]>();
        if (rows.Count > 0)
            cleaned.Add(rows[0]);
        cleaned.AddRange(unique);

        desc["role"] = "read-only expected-output oracle; never set_values this matrix onto 정리";
        desc["importedRecords"] = TableJson(rows);
        desc["expectedAfterDedupeSort"] = TableJson(cleaned);
        desc["assertImportedThenCleaned"] = new JsonObject
        {
            ["B2"] = Field(unique.Count > 0 ? unique[0] : [], 1),
            ["D2"] = Field(unique.Count > 0 ? unique[0] : [], 3),
            ["E2"] = Field(unique.Count > 0 ? unique[0] : [], 4),
            ["F2"] = Field(unique.Count > 0 ? unique[0] : [], 5),
            ["B3"] = Field(unique.Count > 1 ? unique[1] : [], 1),
            ["D3"] = Field(unique.Count > 1 ? unique[1] : [], 3),
            ["E3"] = Field(unique.Count > 1 ? unique[1] : [], 4),
            ["F3"] = Field(unique.Count > 1 ? unique[1] : [], 5),
            ["uniqueDataRows"] = unique.Count,
        };
        return desc;
    }

    private static void FlushRow(List<string[]> rows, List<string> fields, StringBuilder sb)
    {
        fields.Add(sb.ToString());
        sb.Clear();
        if (fields.Count == 1 && fields[0].Length == 0 && rows.Count > 0)
        {
            fields.Clear();
            return;
        }
        rows.Add(fields.ToArray());
        fields.Clear();
    }

    private static JsonArray TableJson(IEnumerable<string[]> rows)
    {
        var table = new JsonArray();
        foreach (var row in rows)
        {
            var cells = new JsonArray();
            foreach (var cell in row)
                cells.Add(cell);
            table.Add(cells);
        }
        return table;
    }

    private static string Field(string[] row, int index) =>
        index >= 0 && index < row.Length ? row[index] : "";

    private static bool ContainsHangul(string value)
    {
        foreach (var ch in value)
        {
            if (ch is >= '\uAC00' and <= '\uD7A3') return true;
        }
        return false;
    }
}
