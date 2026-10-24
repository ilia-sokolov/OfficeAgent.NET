using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>Writes text into table cells and prepares copied rows, for the verbs that fill tables.</summary>
internal static class WordTableCells
{
    /// <summary>
    /// Replaces a cell's text with one run carrying the formatting of the first run it replaces.
    /// The first paragraph and its properties stay. Anchors stay too - bookmarks, comment ranges
    /// and comment references - so a comment on a figure is still on the figure after the figure
    /// changes; anchors in later paragraphs move into the first before those paragraphs go.
    /// </summary>
    public static void SetText(TableCell cell, string text)
    {
        var paragraphs = cell.Elements<Paragraph>().ToList();
        var first = paragraphs.FirstOrDefault() ?? cell.AppendChild(new Paragraph());
        foreach (var extra in paragraphs.Skip(1))
        {
            foreach (var anchor in extra.ChildElements.Where(IsAnchor).ToList())
            {
                anchor.Remove();
                first.AppendChild(anchor);
            }

            extra.Remove();
        }

        var content = first.ChildElements.Where(child => child is not ParagraphProperties && !IsAnchor(child)).ToList();
        var pattern = content.SelectMany(child => child is Run run ? new[] { run } : child.Descendants<Run>())
            .FirstOrDefault(run => run.Elements<Text>().Any());
        var properties = pattern?.RunProperties?.CloneNode(deep: true);

        if (text.Length > 0)
        {
            var run = new Run();
            if (properties is not null) run.AppendChild(properties);
            run.AppendChild(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            if (content.FirstOrDefault() is { } place) place.InsertBeforeSelf(run);
            else first.AppendChild(run);
        }

        foreach (var child in content) child.Remove();
    }

    /// <summary>
    /// Drops the anchors a copied row inherited. Bookmark names and comment ids must be unique,
    /// so a copy carrying its source's would give one comment two anchors and one name two places.
    /// </summary>
    public static void StripAnchors(TableRow row)
    {
        foreach (var anchor in row.Descendants().Where(IsAnchor).ToList())
            anchor.Remove();
    }

    private static bool IsAnchor(OpenXmlElement element) =>
        element is BookmarkStart or BookmarkEnd or CommentRangeStart or CommentRangeEnd or PermStart or PermEnd
        || element is Run run && run.Elements<CommentReference>().Any() && !run.Elements<Text>().Any();
}
