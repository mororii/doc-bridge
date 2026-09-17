using System.Text;
using System.Text.Json.Nodes;
using DocBridge.Core.Services;

namespace DocBridge.Core.Tests;

public class ExcelCsvContractTests
{
    [Theory]
    [InlineData("id,note\r\n0012,\"PVC, \"\"200\"\"\"", "0012", "PVC, \"200\"")]
    [InlineData("name,note\r\n맨홀,\"line1\r\n=1+1\"", "맨홀", "line1\r\n=1+1")]
    public void Parse_preserves_quoted_comma_escape_crlf_korean_and_literals(string csv, string first, string second)
    {
        var rows = ExcelCsvContract.Parse(csv);
        Assert.Equal(2, rows.Count);
        Assert.Equal(first, rows[1][0]);
        Assert.Equal(second, rows[1][1]);
        Assert.True(ExcelCsvContract.LooksLikeIdentifierOrFormulaText("0012"));
        Assert.True(ExcelCsvContract.LooksLikeIdentifierOrFormulaText("=1+1"));
    }

    [Fact]
    public void Parse_keeps_trailing_empty_field_and_empty_quoted_first_record()
    {
        var trailing = ExcelCsvContract.Parse("a,b,");
        Assert.Equal(new[] { "a", "b", "" }, trailing[0]);
        var firstEmpty = ExcelCsvContract.Parse("\"\"");
        Assert.Equal(new[] { "" }, firstEmpty[0]);
    }

    [Fact]
    public void Parse_keeps_explicit_final_quoted_empty_record()
    {
        var rows = ExcelCsvContract.Parse("a\r\n\"\"");
        Assert.Equal(2, rows.Count);
        Assert.Equal(new[] { "a" }, rows[0]);
        Assert.Equal(new[] { "" }, rows[1]);
    }

    [Theory]
    [InlineData("\"a\"x,b")]
    [InlineData("ab\"cd\"e,f")]
    public void Parse_refuses_malformed_quote_sequences(string csv)
    {
        var ex = Assert.Throws<InvalidOperationException>(() => ExcelCsvContract.Parse(csv));
        Assert.Contains("malformed quote sequence", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Parse_still_refuses_unclosed_quote()
    {
        var ex = Assert.Throws<InvalidOperationException>(() => ExcelCsvContract.Parse("\"unclosed"));
        Assert.Contains("unclosed quoted field", ex.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Write_then_parse_round_trips_terminal_empty_record()
    {
        IReadOnlyList<IReadOnlyList<string>> table = [["a"], [""]];
        var written = ExcelCsvContract.Write(table);
        Assert.Equal("a\r\n\"\"", written);
        var again = ExcelCsvContract.Parse(written);
        Assert.True(ExcelCsvContract.TablesEqual(table, again));
    }

    [Fact]
    public void Write_quotes_empty_fields_and_escaped_quotes()
    {
        Assert.Equal("\"\"", ExcelCsvContract.Escape(""));
        Assert.Equal("\"PVC, \"\"200\"\"\"", ExcelCsvContract.Escape("PVC, \"200\""));
        var written = ExcelCsvContract.Write([["0012", "PVC, \"200\""], ["=1+1", ""]]);
        var again = ExcelCsvContract.Parse(written);
        Assert.Equal("0012", again[0][0]);
        Assert.Equal("PVC, \"200\"", again[0][1]);
        Assert.Equal("=1+1", again[1][0]);
        Assert.Equal("", again[1][1]);
    }

    [Fact]
    public void Custom_delimiter_round_trips_without_comma_split()
    {
        var table = ExcelCsvContract.NormalizeRectangle(new[]
        {
            new[] { "0012", "a;b", "" },
            new[] { "=1+1", "맨홀", "x" },
        });
        var written = ExcelCsvContract.Write(table, ';');
        Assert.Contains(';', written);
        Assert.DoesNotContain(",", written.Split('"')[0], StringComparison.Ordinal);
        var again = ExcelCsvContract.Parse(written, ';');
        Assert.True(ExcelCsvContract.TablesEqual(table, ExcelCsvContract.NormalizeRectangle(again)));
        Assert.Equal(',', ExcelCsvContract.ResolveDelimiter(new JsonObject()));
        Assert.Equal(';', ExcelCsvContract.ResolveDelimiter(new JsonObject { ["delimiter"] = ";" }));
    }

    [Fact]
    public void Utf8_file_bytes_preserve_korean_and_literals()
    {
        var csv = "id,name,sku,literal\r\n1,\"맨홀\",\"0012\",\"=1+1\"\r\n";
        var path = Path.Combine(Path.GetTempPath(), "docbridge-csv-" + Guid.NewGuid().ToString("N")[..8] + ".csv");
        File.WriteAllText(path, csv, new UTF8Encoding(false));
        try
        {
            var text = File.ReadAllText(path, new UTF8Encoding(false));
            var rows = ExcelCsvContract.Parse(text);
            Assert.Equal("맨홀", rows[1][1]);
            Assert.Equal("0012", rows[1][2]);
            Assert.Equal("=1+1", rows[1][3]);
        }
        finally
        {
            try { File.Delete(path); } catch { /* temp */ }
        }
    }

    [Fact]
    public void NormalizeRectangle_refuses_ragged_rows_times_width_before_padding()
    {
        var parsed = ExcelCsvContract.Parse("a\r\nb\r\nc,d,e,f,g", ',', maxCells: 10);
        Assert.Equal(3, parsed.Count);
        Assert.Equal(7, parsed.Sum(row => row.Count));
        var ex = Assert.Throws<InvalidOperationException>(() =>
            ExcelCsvContract.NormalizeRectangle(parsed, maxCells: 10));
        Assert.Contains("exceeds", ex.Message, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("3x5", ex.Message, StringComparison.Ordinal);
        Assert.Null(ExcelCsvContract.ImportLimitError(2, 5, 10));
        Assert.Contains("exceeds", ExcelCsvContract.ImportLimitError(3, 5, 10), StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void NormalizeRectangle_keeps_empty_rows_and_pads_ragged_columns_when_under_limit()
    {
        var parsed = ExcelCsvContract.NormalizeRectangle(ExcelCsvContract.Parse("a,b\r\n\r\nc"));
        Assert.Equal(3, parsed.Count);
        Assert.Equal(2, parsed[0].Count);
        Assert.Equal("", parsed[1][0]);
        Assert.Equal("", parsed[1][1]);
        Assert.Equal("c", parsed[2][0]);
        Assert.Equal("", parsed[2][1]);
    }
}
