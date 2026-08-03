using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using OfficeAgent.Core;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// DrawingML dialect: <c>a:r</c> runs carrying <c>a:t</c> text, as used by every text
/// body in a presentation - shapes, table cells, and notes alike.
/// </summary>
/// <remarks>
/// Supplying this to the shared <see cref="TextBodyEngine"/> is what lets PowerPoint
/// reuse Word's run-spanning replacement: the engine finds text across run boundaries
/// and rewrites the minimum number of runs, so character formatting on the runs it does
/// not touch survives untouched. The two dialects differ only in element vocabulary.
/// </remarks>
public sealed class PresentationmlDialect : ITextDialect
{
    /// <inheritdoc />
    public DocFormat Format => DocFormat.PowerPoint;

    /// <inheritdoc />
    public IReadOnlyList<OpenXmlElement> GetRuns(OpenXmlElement paragraph) =>
        paragraph.Elements<Run>().Cast<OpenXmlElement>().ToList();

    /// <inheritdoc />
    public bool IsTextRun(OpenXmlElement run) =>
        run is Run r && r.Elements<Text>().Any();

    /// <inheritdoc />
    public string GetRunText(OpenXmlElement run) =>
        run is Run r ? string.Concat(r.Elements<Text>().Select(t => t.Text)) : string.Empty;

    /// <inheritdoc />
    public void SetRunText(OpenXmlElement run, string text)
    {
        if (run is not Run r) return;

        // a:r allows exactly one a:t, but be tolerant of what other tools wrote.
        foreach (var extra in r.Elements<Text>().Skip(1).ToList())
            extra.Remove();

        var first = r.Elements<Text>().FirstOrDefault();
        if (first is null)
        {
            first = new Text();
            r.AppendChild(first);
        }

        first.Text = text;
    }

    /// <inheritdoc />
    public OpenXmlElement CloneRunShell(OpenXmlElement run, string text)
    {
        var clone = (Run)run.CloneNode(deep: true);

        foreach (var t in clone.Elements<Text>().ToList())
            t.Remove();

        // a:rPr must stay first in a:r, so the text is appended after the cloned
        // properties rather than inserted ahead of them.
        clone.AppendChild(new Text(text));
        return clone;
    }
}
