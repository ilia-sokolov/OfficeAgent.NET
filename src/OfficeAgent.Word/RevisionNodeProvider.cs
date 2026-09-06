using DocumentFormat.OpenXml;
using OfficeAgent.Core;
using OfficeAgent.Abstractions;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Surfaces and resolves tracked revisions across every text host (body, headers, footers,
/// footnotes, endnotes).
/// </summary>
/// <remarks>
/// <para>
/// A Word redline is not only <c>w:ins</c> and <c>w:del</c> around runs. Splitting a
/// paragraph marks its <em>mark</em>; adding a table row marks the row; restyling a
/// heading records a <c>w:pPrChange</c>; moving a clause writes <c>w:moveFrom</c> and
/// <c>w:moveTo</c>. Surfacing only the two run wrappers made "accept everything" report
/// success on a document that still had revisions in it - so every marker
/// WordprocessingML defines is enumerated here, each under a tag naming what it is.
/// </para>
/// <para>
/// Each node's summary carries the author and the date, because the first question asked
/// of a redline is whose it is.
/// </para>
/// </remarks>
internal sealed class RevisionNodeProvider : IWordNodeProvider
{
    public string Kind => "revision";

    /// <summary>The path prefix that selects by author rather than by id.</summary>
    private const string AuthorPrefix = "author:";

    public IEnumerable<NodeInfo> Enumerate(WordObjectMap map)
    {
        foreach (var revision in Revisions(map.Package))
            yield return Node(revision);
    }

    public ResolvedNode? Resolve(NodeAnchor anchor, WordObjectMap map)
    {
        var path = anchor.Path ?? string.Empty;
        var all = Revisions(map.Package).ToList();

        List<Revision> selected;
        if (string.Equals(path, "all", StringComparison.OrdinalIgnoreCase))
        {
            selected = all;
        }
        else if (path.StartsWith(AuthorPrefix, StringComparison.OrdinalIgnoreCase))
        {
            var author = path.Substring(AuthorPrefix.Length);
            selected = all
                .Where(r => string.Equals(r.Author, author, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }
        else
        {
            selected = all.Where(r => string.Equals(r.Path, path, StringComparison.Ordinal)).ToList();
        }

        // "all" and an author filter are set selections: matching nothing means there is
        // nothing to do, not that the caller named something that does not exist.
        var isSet = string.Equals(path, "all", StringComparison.OrdinalIgnoreCase)
                 || path.StartsWith(AuthorPrefix, StringComparison.OrdinalIgnoreCase);
        if (selected.Count == 0 && !isSet) return null;

        return new ResolvedNode
        {
            Kind = Kind,
            Elements = selected.Select(r => r.Element).ToList()
        };
    }

    // ── Enumeration ──────────────────────────────────────────────────────

    /// <summary>One tracked revision: the marker element, what kind it is, and who left it.</summary>
    internal readonly record struct Revision(
        OpenXmlElement Element, string Tag, string Id, string Author, string Date, string Text)
    {
        public string Path => $"{Tag}#{Id}";
    }

    /// <summary>Every revision marker in the document, in document order per host.</summary>
    internal static IEnumerable<Revision> Revisions(IOpenXmlPackage package)
    {
        foreach (var (root, _) in WordModel.TextHosts(package))
            foreach (var element in root.Descendants<OpenXmlElement>())
                if (WordRevisions.TagOf(element) is { } tag)
                    yield return Describe(element, tag);
    }

    private static Revision Describe(OpenXmlElement element, string tag) => new(
        element,
        tag,
        WordRevisions.IdOf(element) ?? string.Empty,
        WordRevisions.AuthorOf(element) ?? string.Empty,
        WordRevisions.DateOf(element) ?? string.Empty,
        TextOf(element));

    /// <summary>
    /// What the revision is about, in words: the text for a run wrapper, and the affected
    /// content for a marker that carries none of its own.
    /// </summary>
    private static string TextOf(OpenXmlElement element) => element switch
    {
        DeletedRun or MoveFromRun => string.Concat(element.Descendants<DeletedText>().Select(t => t.Text)),
        InsertedRun or MoveToRun => string.Concat(element.Descendants<Text>().Select(t => t.Text)),
        _ => Owner(element) is { } owner ? Snippet(owner) : string.Empty
    };

    /// <summary>The paragraph, row, cell, or table a property-level marker belongs to.</summary>
    private static OpenXmlElement? Owner(OpenXmlElement element) =>
        element.Ancestors<Paragraph>().FirstOrDefault()
        ?? (OpenXmlElement?)element.Ancestors<TableRow>().FirstOrDefault()
        ?? element.Ancestors<Table>().FirstOrDefault();

    private static string Snippet(OpenXmlElement owner, int max = 60)
    {
        var text = string.Concat(owner.Descendants<Text>().Select(t => t.Text));
        return text.Length <= max ? text : text.Substring(0, max) + "…";
    }

    private NodeInfo Node(Revision revision)
    {
        var who = string.IsNullOrEmpty(revision.Author) ? "unknown" : revision.Author;
        var when = string.IsNullOrEmpty(revision.Date) ? string.Empty : $", {revision.Date}";

        return new NodeInfo
        {
            Kind = Kind,
            Path = revision.Path,
            Summary = $"{revision.Tag} by {who}{when}: {revision.Text}",
            Anchor = new NodeAnchor
            {
                Id = $"rev:{revision.Tag}:{revision.Id}",
                Kind = Kind,
                Path = revision.Path,
                Expect = revision.Text
            }
        };
    }
}
