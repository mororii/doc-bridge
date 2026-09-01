using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;

namespace DocBridge.Core.Adapters;

public sealed partial class HwpAdapter
{
    /// <summary>
    /// 전체 TEXT를 안정적인 lineId와 coverage가 있는 문단 지도(CVD-lite)로 변환한다.
    /// 동일 문단은 내용 hash+등장 순번으로 구분되어 단순 절대 문단 번호보다 앞쪽 삽입에 강하다.
    /// </summary>
    private static JsonObject BuildDocumentMap(dynamic hwp, int startParagraph, int maxParagraphs)
    {
        string text = NormalizeNewlines((string)GetDocText(hwp));
        return BuildDocumentMapFromText(hwp, text, startParagraph, maxParagraphs);
    }

    private static JsonObject BuildDocumentMapFromText(
        dynamic hwp, string normalizedText, int startParagraph, int maxParagraphs)
    {
        string text = NormalizeNewlines(normalizedText);
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
            ["documentIdentity"] = CaptureDocumentIdentity(hwp),
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

    private static JsonObject CapturePostEditReread(dynamic hwp)
    {
        string text = NormalizeNewlines((string)GetDocText(hwp));
        var digest = Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text))).ToLowerInvariant();
        return new JsonObject
        {
            ["documentIdentity"] = CaptureDocumentIdentity(hwp),
            ["textSha256"] = digest,
            ["textLength"] = text.Length,
            ["textPreview"] = text[..Math.Min(500, text.Length)],
            ["textTail"] = text[Math.Max(0, text.Length - 500)..],
            ["documentMap"] = BuildDocumentMapFromText(hwp, text, 0, 80),
            ["instruction"] = "다음 편집은 이 갱신된 documentMap/lineId와 텍스트를 기준으로 다시 계획하세요. 실패 전의 오래된 위치 정보는 재사용하지 마세요.",
        };
    }
}
