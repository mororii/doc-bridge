using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;
using DocBridge.Core.Adapters;
using DocBridge.Core.Services;
using Xunit.Abstractions;

namespace DocBridge.Core.Tests;

public sealed class HwpDocumentMapPerformanceTests
{
    private readonly ITestOutputHelper _output;

    public HwpDocumentMapPerformanceTests(ITestOutputHelper output) => _output = output;

    [Fact]
    public void Later_page_duplicate_line_ids_match_full_scan_occurrence_counts()
    {
        var lines = new List<string>(240);
        for (var i = 0; i < 240; i++)
        {
            lines.Add(i % 3 == 0 ? "REPEAT-A" : i % 3 == 1 ? "REPEAT-B" : $"unique-{i}");
        }

        var text = string.Join('\n', lines);
        foreach (var start in new[] { 0, 80, 160 })
            AssertBodyEqual(text, start, 80);
    }

    [Fact]
    public void Empty_and_trailing_lines_keep_blank_ids_and_coverage()
    {
        AssertBodyEqual("\nhello\n\nworld\n", startParagraph: 0, maxParagraphs: 80);
        AssertBodyEqual("only\n", startParagraph: 0, maxParagraphs: 80);
        AssertBodyEqual("\n\n", startParagraph: 1, maxParagraphs: 80);
    }

    [Fact]
    public void Start_at_or_beyond_end_returns_empty_window_with_full_total()
    {
        const string text = "a\nb\nc";
        AssertBodyEqual(text, startParagraph: 3, maxParagraphs: 80);
        AssertBodyEqual(text, startParagraph: 99, maxParagraphs: 80);

        var beyond = HwpAdapter.BuildDocumentMapBodyFromText(text, 99, 80);
        var coverage = Json.GetObj(beyond, "coverage")!;
        Assert.Equal(3, Json.GetInt(coverage, "totalParagraphs"));
        Assert.Equal(3, Json.GetInt(coverage, "startParagraph"));
        Assert.Equal(3, Json.GetInt(coverage, "endParagraphExclusive"));
        Assert.Equal(0, Json.GetInt(coverage, "returnedParagraphs"));
        Assert.Empty(Json.GetArr(beyond, "paragraphs")!);
    }

    [Fact]
    public void Fifty_thousand_paragraphs_max80_matches_old_loop_and_reports_timing()
    {
        var lines = new string[50_000];
        for (var i = 0; i < lines.Length; i++)
            lines[i] = i % 17 == 0 ? "REPEAT-PAGE" : $"para-{i}";
        var text = string.Join('\n', lines);

        AssertBodyEqual(text, startParagraph: 0, maxParagraphs: 80);
        AssertBodyEqual(text, startParagraph: 160, maxParagraphs: 80);

        BuildDocumentMapBodyOld(text, 0, 80);
        HwpAdapter.BuildDocumentMapBodyFromText(text, 0, 80);

        var oldMs = TimeMs(() => BuildDocumentMapBodyOld(text, 0, 80));
        var newMs = TimeMs(() => HwpAdapter.BuildDocumentMapBodyFromText(text, 0, 80));
        _output.WriteLine(
            "hwp-document-map-a50000-max80 oldLoopMs={0} endExclusiveLoopMs={1} timingIsReportNotAssert=true",
            oldMs, newMs);

        var optionalPath = Environment.GetEnvironmentVariable("DOCBRIDGE_HWP_MAP_TIMING_PATH");
        if (string.IsNullOrWhiteSpace(optionalPath)) return;

        var report = new JsonObject
        {
            ["name"] = "hwp-document-map-a50000-max80",
            ["paragraphs"] = 50_000,
            ["maxParagraphs"] = 80,
            ["oldLoopMs"] = oldMs,
            ["endExclusiveLoopMs"] = newMs,
            ["timingIsReportNotAssert"] = true,
            ["outputsEqual"] = true,
        };
        Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(optionalPath))!);
        File.WriteAllText(optionalPath, report.ToJsonString(Json.Pretty));
    }

    private static void AssertBodyEqual(string text, int startParagraph, int maxParagraphs)
    {
        var expected = BuildDocumentMapBodyOld(text, startParagraph, maxParagraphs);
        var actual = HwpAdapter.BuildDocumentMapBodyFromText(text, startParagraph, maxParagraphs);
        Assert.Equal(Json.Canonical(expected["paragraphs"]), Json.Canonical(actual["paragraphs"]));
        Assert.Equal(Json.Canonical(expected["coverage"]), Json.Canonical(actual["coverage"]));
    }

    private static JsonObject BuildDocumentMapBodyOld(string normalizedText, int startParagraph, int maxParagraphs)
    {
        string text = normalizedText.Replace("\r\n", "\n").Replace('\r', '\n');
        string[] paragraphs = text.Split('\n');
        startParagraph = Math.Clamp(startParagraph, 0, paragraphs.Length);
        maxParagraphs = Math.Clamp(maxParagraphs, 1, 2000);
        var endExclusive = Math.Min(paragraphs.Length, startParagraph + maxParagraphs);
        var occurrences = new Dictionary<string, int>(StringComparer.Ordinal);
        var items = new JsonArray();

        for (var index = 0; index < paragraphs.Length; index++)
        {
            var paragraph = paragraphs[index];
            var digest = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(paragraph)))
                .ToLowerInvariant()[..12];
            occurrences.TryGetValue(digest, out int occurrence);
            occurrence++;
            occurrences[digest] = occurrence;
            if (index < startParagraph || index >= endExclusive) continue;

            items.Add(new JsonObject
            {
                ["lineId"] = $"p-{digest}-{occurrence}",
                ["paragraphIndex"] = index,
                ["text"] = paragraph,
                ["textLength"] = paragraph.Length,
                ["blank"] = paragraph.Length == 0,
            });
        }

        return new JsonObject
        {
            ["paragraphs"] = items,
            ["coverage"] = new JsonObject
            {
                ["totalParagraphs"] = paragraphs.Length,
                ["startParagraph"] = startParagraph,
                ["endParagraphExclusive"] = endExclusive,
                ["returnedParagraphs"] = items.Count,
                ["complete"] = startParagraph == 0 && endExclusive == paragraphs.Length,
                ["truncated"] = endExclusive < paragraphs.Length,
                ["nextStartParagraph"] = endExclusive < paragraphs.Length ? endExclusive : null,
            },
        };
    }

    private static long TimeMs(Action work)
    {
        var sw = Stopwatch.StartNew();
        work();
        sw.Stop();
        return sw.ElapsedMilliseconds;
    }
}
