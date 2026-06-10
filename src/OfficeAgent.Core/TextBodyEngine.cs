using System.Text;
using DocumentFormat.OpenXml;

namespace OfficeAgent.Core;

/// <summary>
/// Provides run-spanning text operations for Open XML paragraphs whose logical text is split across multiple runs.
/// </summary>
internal sealed class TextBodyEngine
{
    private readonly ITextDialect _dialect;

    public TextBodyEngine(ITextDialect dialect) => _dialect = dialect;


    public string GetLogicalText(OpenXmlElement paragraph)
    {
        var sb = new StringBuilder();
        foreach (var run in _dialect.GetRuns(paragraph))
            if (_dialect.IsTextRun(run))
                sb.Append(_dialect.GetRunText(run));
        return sb.ToString();
    }

    public int CountOccurrences(string text, string needle, StringComparison comparison)
    {
        if (string.IsNullOrEmpty(needle)) return 0;

        int count = 0, from = 0;
        while (true)
        {
            int idx = text.IndexOf(needle, from, comparison);
            if (idx < 0) break;
            count++;
            from = idx + needle.Length;
        }
        return count;
    }


    public int IndexOfOccurrence(string text, string needle, int occurrence, StringComparison comparison)
    {
        if (string.IsNullOrEmpty(needle) || occurrence < 0) return -1;

        int idx = -1, from = 0;
        for (int i = 0; i <= occurrence; i++)
        {
            idx = text.IndexOf(needle, from, comparison);
            if (idx < 0) return -1;
            from = idx + needle.Length;
        }
        return idx;
    }







    public IReadOnlyList<OpenXmlElement> IsolateSpan(OpenXmlElement paragraph, int start, int length)
    {
        int end = start + length;
        var covered = new List<OpenXmlElement>();
        int cursor = 0;

        foreach (var run in _dialect.GetRuns(paragraph))
        {
            if (!_dialect.IsTextRun(run))
                continue;

            string text = _dialect.GetRunText(run);
            int runStart = cursor;
            int runEnd = cursor + text.Length;
            cursor = runEnd;

            if (runEnd <= start || runStart >= end)
                continue;

            int localStart = Math.Max(start, runStart) - runStart;
            int localEnd = Math.Min(end, runEnd) - runStart;

            string pre = text.Substring(0, localStart);
            string mid = text.Substring(localStart, localEnd - localStart);
            string post = text.Substring(localEnd);

            var parent = run.Parent
                ?? throw new InvalidOperationException("Run has no parent paragraph.");

            if (pre.Length > 0)
                parent.InsertBefore(_dialect.CloneRunShell(run, pre), run);

            _dialect.SetRunText(run, mid);
            covered.Add(run);

            if (post.Length > 0)
                parent.InsertAfter(_dialect.CloneRunShell(run, post), run);
        }

        return covered;
    }
}
