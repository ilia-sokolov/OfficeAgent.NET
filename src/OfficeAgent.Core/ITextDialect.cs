using System.Diagnostics.CodeAnalysis;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Core;

/// <summary>
/// Strategy that lets one <see cref="TextBodyEngine"/> serve different run/text
/// vocabularies. Today only the Word dialect is registered; the seam is kept so
/// future modules (e.g. PowerPoint DrawingML) can plug in without changing the engine.
/// </summary>
[Experimental(EngineExtensibility.DiagnosticId, UrlFormat = EngineExtensibility.UrlFormat)]
public interface ITextDialect
{
    /// <summary>Gets the document format whose run vocabulary this dialect understands.</summary>
    DocFormat Format { get; }

    /// <summary>Gets a paragraph's runs in document order, including runs nested in wrappers.</summary>
    IReadOnlyList<OpenXmlElement> GetRuns(OpenXmlElement paragraph);

    /// <summary>Gets whether a run carries text.</summary>
    bool IsTextRun(OpenXmlElement run);

    /// <summary>Gets the text a run carries, or an empty string.</summary>
    string GetRunText(OpenXmlElement run);

    /// <summary>Replaces a run's text, keeping its formatting.</summary>
    void SetRunText(OpenXmlElement run, string text);

    /// <summary>Creates a copy of a run's formatting that carries new text.</summary>
    OpenXmlElement CloneRunShell(OpenXmlElement run, string text);
}

/// <summary>WordprocessingML dialect: <c>w:r</c> runs carrying <c>w:t</c> text.</summary>
[Experimental(EngineExtensibility.DiagnosticId, UrlFormat = EngineExtensibility.UrlFormat)]
public sealed class WordmlDialect : ITextDialect
{
    /// <inheritdoc />
    public DocFormat Format => DocFormat.Word;

    /// <inheritdoc />
    public IReadOnlyList<OpenXmlElement> GetRuns(OpenXmlElement paragraph)
    {
        var runs = new List<OpenXmlElement>();
        Collect(paragraph, runs);
        return runs;
    }

    /// <summary>
    /// Walks the paragraph's inline content in document order. A run is not always a
    /// direct child of <c>w:p</c>: tracked insertions (<c>w:ins</c>), hyperlinks,
    /// content controls and fields all wrap their runs, and text inside them is live
    /// text the reader sees. Enumerating only direct children makes that text
    /// invisible to find, changeText and format - so a tracked edit or a content
    /// control could not be targeted afterwards.
    /// </summary>
    private static void Collect(OpenXmlElement element, List<OpenXmlElement> runs)
    {
        foreach (var child in element.ChildElements)
        {
            // Stop at the run itself: a run may carry a text box, whose paragraphs
            // belong to their own body and must not fold into this one's text.
            if (child is Run)
            {
                runs.Add(child);
                continue;
            }

            // Deleted and moved-from content is struck through, not part of the text.
            if (child is DeletedRun or MoveFromRun)
                continue;

            Collect(child, runs);
        }
    }

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

        foreach (var extra in r.Elements<Text>().Skip(1).ToList())
            extra.Remove();

        var first = r.Elements<Text>().FirstOrDefault();
        if (first is null)
        {
            first = new Text();
            r.AppendChild(first);
        }

        first.Text = text;
        first.Space = SpaceProcessingModeValues.Preserve;
    }

    /// <inheritdoc />
    public OpenXmlElement CloneRunShell(OpenXmlElement run, string text)
    {
        var clone = (Run)run.CloneNode(deep: true);

        foreach (var t in clone.Elements<Text>().ToList())
            t.Remove();

        clone.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return clone;
    }
}
