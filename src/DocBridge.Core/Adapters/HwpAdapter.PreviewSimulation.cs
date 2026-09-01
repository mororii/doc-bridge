using System.Text;

namespace DocBridge.Core.Adapters;

public sealed partial class HwpAdapter
{
    /// <summary>
    /// GetTextFile("TEXT") serializes a subset of Unicode as HTML numeric entities even
    /// though JSON can carry the original Unicode safely. Public read results and all
    /// preview/apply decisions must use the same decoded text.
    /// </summary>
    internal static string DecodeHwpSerializedText(string value) =>
        System.Net.WebUtility.HtmlDecode(value ?? string.Empty);

    /// <summary>
    /// HWP TEXT readback can return compatibility units and mathematical minus signs in a
    /// different but visually equivalent Unicode representation. Verification must compare
    /// canonical text without changing the text that is actually written to the document.
    /// </summary>
    internal static string NormalizeHwpReadbackComparable(string value)
    {
        // GetTextFile("TEXT") can serialize unsupported Unicode as a literal numeric
        // entity (for example U+2212 becomes "&#8722;"). Decode only for comparison.
        var decoded = DecodeHwpSerializedText(NormalizeNewlines(value));
        var normalized = decoded.Normalize(NormalizationForm.FormKC);
        return normalized
            .Replace('\u2010', '-')
            .Replace('\u2011', '-')
            .Replace('\u2012', '-')
            .Replace('\u2013', '-')
            .Replace('\u2043', '-')
            .Replace('\u2212', '-')
            .Replace('\uFE63', '-')
            .Replace('\uFF0D', '-');
    }

    internal static bool HwpReadbackContainsEquivalent(string actual, string expected) =>
        NormalizeHwpReadbackComparable(actual).Contains(
            NormalizeHwpReadbackComparable(expected), StringComparison.Ordinal);

    internal static string SimulatePreviewAppend(
        string document, string text, bool startNewParagraph)
    {
        var before = NormalizeNewlines(document);
        var inserted = NormalizeNewlines(text);
        if (startNewParagraph && before.TrimEnd('\n').Length > 0 &&
            !before.EndsWith('\n') && !inserted.StartsWith('\n'))
            return before + "\n" + inserted;
        return before + inserted;
    }

    internal static string SimulatePreviewRelativeInsert(
        string document, string anchor, string text, int occurrence,
        bool matchCase, bool before, string mode)
    {
        var source = NormalizeNewlines(document);
        var inserted = NormalizeNewlines(text);
        var anchorIndex = IndexOfTextOccurrence(source, anchor, occurrence, matchCase);
        if (anchorIndex < 0) return source;

        if (string.Equals(mode, "inline", StringComparison.OrdinalIgnoreCase))
        {
            var offset = before ? anchorIndex : anchorIndex + anchor.Length;
            return source.Insert(offset, inserted);
        }

        if (before)
        {
            var paragraphStart = source.LastIndexOf('\n', Math.Max(0, anchorIndex - 1));
            paragraphStart = paragraphStart < 0 ? 0 : paragraphStart + 1;
            var value = inserted.EndsWith('\n') ? inserted : inserted + "\n";
            return source.Insert(paragraphStart, value);
        }

        var anchorEnd = anchorIndex + anchor.Length;
        var paragraphEnd = source.IndexOf('\n', anchorEnd);
        if (paragraphEnd < 0) paragraphEnd = source.Length;
        var prefix = inserted.StartsWith('\n') ? "" : "\n";
        return source.Insert(paragraphEnd, prefix + inserted);
    }
}
