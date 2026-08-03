using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using A = DocumentFormat.OpenXml.Drawing;
using PC = DocumentFormat.OpenXml.Office2021.PowerPoint.Comment;
using PCMD = DocumentFormat.OpenXml.Office2016.Presentation.Command;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// Surfaces each modern PowerPoint comment as an addressable node, including whether it
/// has already been resolved, so an agent can review a deck's open feedback.
/// </summary>
/// <remarks>
/// This is the Office 2021 comment model (<c>p188:cm</c> in a
/// <see cref="PowerPointCommentPart"/>), which is what current PowerPoint writes and the
/// only one carrying a status an agent can resolve. Legacy <c>p:cm</c> comments have no
/// resolved state at all, so they are deliberately not surfaced as resolvable.
/// Word records resolution differently, in a <c>commentsExtended</c> part; the Word
/// module does not implement it yet, so <c>action: "Resolve"</c> is refused there.
/// </remarks>
internal sealed class SlideCommentNodeProvider : IPowerPointNodeProvider
{
    /// <inheritdoc />
    public string Kind => "comment";

    /// <inheritdoc />
    public IEnumerable<NodeInfo> Enumerate(PowerPointObjectMap map)
    {
        var authors = SlideCommentWriter.Authors(map);

        foreach (var comment in Comments(map))
        {
            var author = comment.Comment.AuthorId?.Value is { } id && authors.TryGetValue(id, out var name)
                ? name
                : "unknown";
            var status = comment.Comment.Status?.Value;
            var resolved = status is not null && status == PC.CommentStatus.Resolved;

            yield return new NodeInfo
            {
                Kind = Kind,
                Path = comment.Path,
                Summary = $"slide {comment.Slide.Number}: {author}" +
                          $" - \"{Truncate(SlideCommentWriter.TextOf(comment.Comment))}\"" +
                          (resolved ? " (resolved)" : string.Empty),
                Anchor = new NodeAnchor { Id = comment.Path, Kind = Kind, Path = comment.Path }
            };
        }
    }

    /// <inheritdoc />
    public ResolvedNode? Resolve(NodeAnchor anchor, PowerPointObjectMap map)
    {
        var located = Locate(anchor.Path, map);
        return located is null
            ? null
            : new ResolvedNode
            {
                Kind = Kind,
                Elements = new OpenXmlElement[] { located.Comment },
                Value = SlideCommentWriter.TextOf(located.Comment)
            };
    }

    /// <summary>Finds one comment by its path, or null when it is gone.</summary>
    internal static CommentRef? Locate(string path, PowerPointObjectMap map)
    {
        foreach (var comment in Comments(map))
            if (string.Equals(comment.Path, path, StringComparison.OrdinalIgnoreCase))
                return comment;
        return null;
    }

    /// <summary>Enumerates every modern comment in the deck, in slide order.</summary>
    internal static IEnumerable<CommentRef> Comments(PowerPointObjectMap map)
    {
        foreach (var slide in PowerPointModel.Slides(map.Package))
        {
            var part = slide.Part.GetPartsOfType<PowerPointCommentPart>().FirstOrDefault();
            if (part?.CommentList is null) continue;

            foreach (var comment in part.CommentList.Elements<PC.Comment>())
            {
                var id = comment.Id?.Value;
                if (string.IsNullOrEmpty(id)) continue;
                yield return new CommentRef($"comment#{slide.SlideId}/{id}", comment, slide, part);
            }
        }
    }

    private static string Truncate(string text) =>
        text.Length <= 40 ? text : text.Substring(0, 40) + "…";
}

/// <summary>One comment, the part that holds it, and the slide it belongs to.</summary>
internal sealed class CommentRef
{
    public CommentRef(string path, PC.Comment comment, SlideRef slide, PowerPointCommentPart part)
    {
        Path = path;
        Comment = comment;
        Slide = slide;
        Part = part;
    }

    public string Path { get; }
    public PC.Comment Comment { get; }
    public SlideRef Slide { get; }
    public PowerPointCommentPart Part { get; }
}

/// <summary>
/// Adds a comment to a slide, or marks an existing one resolved.
/// </summary>
/// <remarks>
/// Adding targets the slide (<c>{ "kind": "slide", "path": "slide#256" }</c>); resolving
/// targets the comment (<c>{ "kind": "comment", "path": "comment#256/{id}" }</c>), whose
/// path comes from inspection. Resolving keeps the comment and any replies and changes
/// only its status, so the review trail is preserved rather than deleted.
/// </remarks>
internal sealed class SlideCommentHandler : IOperationHandler
{
    private readonly TimeProvider _clock;

    public SlideCommentHandler(TimeProvider clock) => _clock = clock;

    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is CommentOp { Target: NodeAnchor { Kind: "slide" or "comment" } };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (CommentOp)operation;
        var anchor = (NodeAnchor)op.Target!;
        var map = new PowerPointObjectMap(context.Package);

        return op.Action == CommentAction.Resolve
            ? PreviewResolve(op, anchor, map)
            : PreviewAdd(op, anchor, context);
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (CommentOp)operation;
        var anchor = (NodeAnchor)op.Target!;
        var map = new PowerPointObjectMap(context.Package);

        if (op.Action == CommentAction.Resolve)
        {
            var comment = SlideCommentNodeProvider.Locate(anchor.Path, map)
                ?? throw new InvalidOperationException($"Comment '{anchor.Path}' vanished before apply.");
            comment.Comment.Status = PC.CommentStatus.Resolved;
            return;
        }

        var slide = ResolveSlide(context, anchor)
            ?? throw new InvalidOperationException($"Slide '{anchor.Path}' vanished before apply.");
        SlideCommentWriter.Add(map, slide, op, _clock.GetUtcNow());
    }

    private static OperationPreview PreviewAdd(CommentOp op, NodeAnchor anchor, ApplyContext context)
    {
        var slide = ResolveSlide(context, anchor);
        if (slide is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No slide with path '{anchor.Path}'. To add a comment, target a slide; " +
                "to resolve one, target a comment path from inspect_document.nodes.",
                anchor));

        if (string.IsNullOrWhiteSpace(op.Text))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "comment requires non-empty text.", anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "comment",
            Before = string.Empty,
            After = op.Text,
            Context = $"slide {slide.Number}",
            BlastRadius = 1
        });
    }

    private static OperationPreview PreviewResolve(CommentOp op, NodeAnchor anchor, PowerPointObjectMap map)
    {
        if (anchor.Kind != "comment")
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "Resolving a comment targets a comment path from inspect_document.nodes, not a slide.",
                anchor));

        var comment = SlideCommentNodeProvider.Locate(anchor.Path, map);
        if (comment is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No comment with path '{anchor.Path}'.", anchor));

        var already = comment.Comment.Status?.Value is { } status && status == PC.CommentStatus.Resolved;
        if (already)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"Comment '{anchor.Path}' is already resolved.", anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "comment",
            Before = SlideCommentWriter.TextOf(comment.Comment),
            After = "(resolved)",
            Context = $"slide {comment.Slide.Number}",
            BlastRadius = 1
        });
    }

    private static SlideRef? ResolveSlide(ApplyContext context, NodeAnchor anchor) =>
        anchor.Kind == "slide" && SlideNodeProvider.TryParseSlideId(anchor.Path, out var slideId)
            ? PowerPointModel.Slide(context.Package, slideId)
            : null;
}

/// <summary>
/// Writes the parts a modern comment needs: an author entry on the presentation and a
/// comment list on the slide, both created only when they are missing.
/// </summary>
internal static class SlideCommentWriter
{
    /// <summary>Author id → display name, for rendering inspection summaries.</summary>
    public static Dictionary<string, string> Authors(PowerPointObjectMap map)
    {
        var authors = new Dictionary<string, string>(StringComparer.Ordinal);
        var part = map.Main.GetPartsOfType<PowerPointAuthorsPart>().FirstOrDefault();
        if (part?.AuthorList is null) return authors;

        foreach (var author in part.AuthorList.Elements<PC.Author>())
            if (author.Id?.Value is { } id)
                authors[id] = author.Name?.Value ?? "unknown";
        return authors;
    }

    /// <summary>The comment's plain text, joined across the paragraphs of its text body.</summary>
    public static string TextOf(PC.Comment comment)
    {
        var body = comment.GetFirstChild<PC.TextBodyType>();
        if (body is null) return string.Empty;
        return string.Join(" ", body.Elements<A.Paragraph>()
            .Select(PowerPointModel.TextOf)
            .Where(t => t.Length > 0));
    }

    /// <summary>Adds one comment to a slide, creating the author and comment parts as needed.</summary>
    public static void Add(PowerPointObjectMap map, SlideRef slide, CommentOp op, DateTimeOffset now)
    {
        var authorId = EnsureAuthor(map, op.Author, op.Initials);

        var part = slide.Part.GetPartsOfType<PowerPointCommentPart>().FirstOrDefault();
        if (part is null)
        {
            part = slide.Part.AddNewPart<PowerPointCommentPart>();
            part.CommentList = new PC.CommentList();
        }
        part.CommentList ??= new PC.CommentList();

        part.CommentList.Append(new PC.Comment(
            // The moniker is what ties the comment to a slide. PowerPoint writes it on
            // every comment it creates and refuses the package as corrupt without it -
            // schema validation does not require it, so its absence is invisible until a
            // real client opens the deck.
            new PCMD.SlideMonikerList(
                new PCMD.DocumentMoniker(),
                new PCMD.SlideMoniker
                {
                    CId = CreationIdOf(slide),
                    SldId = slide.SlideId
                }),
            // A position places the marker; the plan does not carry one, so it goes a
            // little in from the top-left rather than exactly on the corner.
            new PC.Point2DType { X = 127000L, Y = 127000L },
            new PC.TextBodyType(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(new A.Run(
                    new A.RunProperties { Language = "en-US" },
                    new A.Text(op.Text)))))
        {
            Id = "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}",
            AuthorId = authorId,
            // Local time without a zone suffix, matching what PowerPoint itself writes.
            Created = now.UtcDateTime
            // No status: PowerPoint omits it on a new comment, and an unresolved comment
            // is the default state. Resolving sets it explicitly.
        });
    }

    /// <summary>
    /// The slide's creation id, which the comment moniker pairs with the slide id. Slides
    /// authored elsewhere may not carry one; zero is what PowerPoint accepts in that case.
    /// </summary>
    private static uint CreationIdOf(SlideRef slide)
    {
        foreach (var creationId in slide.Part.Slide
                     .Descendants<DocumentFormat.OpenXml.Office2010.PowerPoint.CreationId>())
            if (creationId.Val?.Value is { } value) return value;
        return 0U;
    }

    /// <summary>
    /// Returns the id of the author entry for a display name, adding one when the deck
    /// has never seen that author. Reusing the entry keeps PowerPoint from showing the
    /// same person once per comment.
    /// </summary>
    private static string EnsureAuthor(PowerPointObjectMap map, string name, string initials)
    {
        var part = map.Main.GetPartsOfType<PowerPointAuthorsPart>().FirstOrDefault();
        if (part is null)
        {
            part = map.Main.AddNewPart<PowerPointAuthorsPart>();
            part.AuthorList = new PC.AuthorList();
        }
        part.AuthorList ??= new PC.AuthorList();

        foreach (var existing in part.AuthorList.Elements<PC.Author>())
            if (string.Equals(existing.Name?.Value, name, StringComparison.Ordinal) &&
                existing.Id?.Value is { } found)
                return found;

        var id = "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}";
        part.AuthorList.Append(new PC.Author
        {
            Id = id,
            Name = name,
            Initials = initials,
            UserId = name,
            ProviderId = "None"
        });
        return id;
    }
}
