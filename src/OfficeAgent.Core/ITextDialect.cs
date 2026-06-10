using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Core;

/// <summary>
/// Strategy that lets one <see cref="TextBodyEngine"/> serve different run/text
/// vocabularies. Today only the Word dialect is registered; the seam is kept so
/// future modules (e.g. PowerPoint DrawingML) can plug in without changing the engine.
/// </summary>
public interface ITextDialect
{
    DocFormat Format { get; }

    IReadOnlyList<OpenXmlElement> GetRuns(OpenXmlElement paragraph);

    bool IsTextRun(OpenXmlElement run);

    string GetRunText(OpenXmlElement run);

    void SetRunText(OpenXmlElement run, string text);

    OpenXmlElement CloneRunShell(OpenXmlElement run, string text);
}

/// <summary>WordprocessingML dialect: <c>w:r</c> runs carrying <c>w:t</c> text.</summary>
public sealed class WordmlDialect : ITextDialect
{
    public DocFormat Format => DocFormat.Word;

    public IReadOnlyList<OpenXmlElement> GetRuns(OpenXmlElement paragraph) =>
        paragraph.Elements<Run>().Cast<OpenXmlElement>().ToList();

    public bool IsTextRun(OpenXmlElement run) =>
        run is Run r && r.Elements<Text>().Any();

    public string GetRunText(OpenXmlElement run) =>
        run is Run r ? string.Concat(r.Elements<Text>().Select(t => t.Text)) : string.Empty;

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

    public OpenXmlElement CloneRunShell(OpenXmlElement run, string text)
    {
        var clone = (Run)run.CloneNode(deep: true);

        foreach (var t in clone.Elements<Text>().ToList())
            t.Remove();

        clone.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
        return clone;
    }
}
