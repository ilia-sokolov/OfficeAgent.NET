namespace OfficeAgent.Word;

/// <summary>
/// Narrows a replacement to the words that differ, for <c>changeText</c> with granularity
/// <c>Minimal</c>.
/// </summary>
/// <remarks>
/// Text is read as alternating segments: maximal runs of whitespace and maximal runs of
/// anything else, which are the words. Punctuation is part of the word it touches, so
/// "45" against "forty-five" is one word against one word, and "days." differs from "days".
/// Segments compare ordinally, so a change of case is a change. The segments both texts share
/// at the start, and then at the end, are kept; whatever lies between them is the edit. Either
/// side of it may be empty, which makes the edit a pure insertion or a pure deletion; both
/// empty means nothing differs.
/// </remarks>
internal static class MinimalTextEdit
{
    /// <summary>
    /// Returns the part of <paramref name="before"/> to delete, as an offset and length, and
    /// the text from <paramref name="after"/> to insert in its place.
    /// </summary>
    public static (int Offset, int DeleteLength, string Insert) Narrow(string before, string after)
    {
        var a = Segments(before);
        var b = Segments(after);

        int prefix = 0;
        while (prefix < a.Count && prefix < b.Count && string.Equals(a[prefix], b[prefix], StringComparison.Ordinal))
            prefix++;

        // The suffix stops where the prefix ended on the shorter side, so the two never share
        // a segment.
        int suffix = 0;
        while (suffix < a.Count - prefix && suffix < b.Count - prefix &&
               string.Equals(a[a.Count - 1 - suffix], b[b.Count - 1 - suffix], StringComparison.Ordinal))
            suffix++;

        int head = Length(a, 0, prefix);
        int tailBefore = Length(a, a.Count - suffix, suffix);
        int tailAfter = Length(b, b.Count - suffix, suffix);

        return (head, before.Length - head - tailBefore, after.Substring(head, after.Length - head - tailAfter));
    }

    /// <summary>Splits text into maximal runs of whitespace and of non-whitespace, in order.</summary>
    internal static List<string> Segments(string text)
    {
        var segments = new List<string>();
        int i = 0;
        while (i < text.Length)
        {
            bool space = char.IsWhiteSpace(text[i]);
            int j = i + 1;
            while (j < text.Length && char.IsWhiteSpace(text[j]) == space) j++;
            segments.Add(text.Substring(i, j - i));
            i = j;
        }
        return segments;
    }

    private static int Length(List<string> segments, int from, int count)
    {
        int length = 0;
        for (int i = from; i < from + count; i++) length += segments[i].Length;
        return length;
    }
}
