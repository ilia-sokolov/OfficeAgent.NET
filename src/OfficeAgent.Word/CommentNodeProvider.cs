using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;
using W15 = DocumentFormat.OpenXml.Office2013.Word;

namespace OfficeAgent.Word;

/// <summary>
/// Surfaces each Word comment as an addressable node - its author, its text, whether it
/// has been resolved, and which thread it belongs to - so an agent can read a document's
/// open feedback and act on it.
/// </summary>
/// <remarks>
/// A comment is two things in two parts: the body in <c>comments.xml</c>, and everything
/// about the conversation - the reply parentage and the resolved flag - in
/// <c>commentsExtended.xml</c>, keyed by the <c>w14:paraId</c> of the comment's last
/// paragraph. Reading only the first part gives an agent a list of remarks with no idea
/// which are answers to which, or which have already been dealt with.
/// </remarks>
internal sealed class CommentNodeProvider : IWordNodeProvider
{
    /// <inheritdoc />
    public string Kind => "comment";

    /// <inheritdoc />
    public IEnumerable<NodeInfo> Enumerate(WordObjectMap map)
    {
        var extended = ExtendedIndex(map.Main);

        foreach (var comment in Comments(map.Main))
        {
            var id = comment.Id?.Value;
            if (string.IsNullOrEmpty(id)) continue;

            var text = TextOf(comment);
            var author = comment.Author?.Value ?? "unknown";
            var info = LastParaId(comment) is { Length: > 0 } paraId && extended.TryGetValue(paraId, out var ex)
                ? ex
                : null;

            var resolved = info?.Done?.Value == true;
            var parent = info?.ParaIdParent?.Value;
            var replyTo = parent is { Length: > 0 } && ParentOf(map.Main, parent) is { } target
                ? $" (reply to comment#{target})"
                : string.Empty;

            yield return new NodeInfo
            {
                Kind = Kind,
                Path = PathOf(id!),
                Summary = $"{author}: \"{Truncate(text)}\"" +
                          (resolved ? " (resolved)" : string.Empty) + replyTo,
                Anchor = new NodeAnchor
                {
                    Id = $"comment:{id}",
                    Kind = Kind,
                    Path = PathOf(id!),
                    Expect = text
                }
            };
        }
    }

    /// <inheritdoc />
    public ResolvedNode? Resolve(NodeAnchor anchor, WordObjectMap map)
    {
        var comment = Locate(map.Main, anchor.Path);
        return comment is null
            ? null
            : new ResolvedNode
            {
                Kind = Kind,
                Elements = new OpenXmlElement[] { comment },
                Value = TextOf(comment)
            };
    }

    // ── Shared with the handler ──────────────────────────────────────────

    public static string PathOf(string id) => $"comment#{id}";

    /// <summary>Finds the comment a <c>comment#N</c> path names.</summary>
    public static Comment? Locate(MainDocumentPart main, string path)
    {
        if (!path.StartsWith("comment#", StringComparison.Ordinal)) return null;
        var id = path.Substring("comment#".Length);
        return Comments(main).FirstOrDefault(c => string.Equals(c.Id?.Value, id, StringComparison.Ordinal));
    }

    public static IEnumerable<Comment> Comments(MainDocumentPart main) =>
        main.WordprocessingCommentsPart?.Comments?.Elements<Comment>() ?? Enumerable.Empty<Comment>();

    public static string TextOf(Comment comment) =>
        string.Concat(comment.Descendants<Text>().Select(t => t.Text));

    /// <summary>
    /// The <c>w14:paraId</c> of a comment's last paragraph, which is the handle
    /// <c>commentsExtended</c> uses to attach parentage and the resolved flag.
    /// </summary>
    public static string? LastParaId(Comment comment) =>
        comment.Elements<Paragraph>().LastOrDefault()?.ParagraphId?.Value;

    /// <summary>The <c>w15:commentEx</c> entries of the document, keyed by paragraph id.</summary>
    public static Dictionary<string, W15.CommentEx> ExtendedIndex(MainDocumentPart main)
    {
        var index = new Dictionary<string, W15.CommentEx>(StringComparer.OrdinalIgnoreCase);
        var extended = main.WordprocessingCommentsExPart?.CommentsEx;
        if (extended is null) return index;

        foreach (var entry in extended.Elements<W15.CommentEx>())
            if (entry.ParaId?.Value is { Length: > 0 } paraId)
                index[paraId] = entry;
        return index;
    }

    /// <summary>The id of the comment whose last paragraph carries <paramref name="paraId"/>.</summary>
    public static string? ParentOf(MainDocumentPart main, string paraId) =>
        Comments(main)
            .FirstOrDefault(c => string.Equals(LastParaId(c), paraId, StringComparison.OrdinalIgnoreCase))
            ?.Id?.Value;

    private static string Truncate(string text, int max = 80) =>
        text.Length <= max ? text : text.Substring(0, max) + "…";
}
