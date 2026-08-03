using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.PowerPoint;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;
using PC = DocumentFormat.OpenXml.Office2021.PowerPoint.Comment;

namespace OfficeAgent.Tests;

/// <summary>
/// Comment verbs on a slide, over the Office 2021 comment model - the one current
/// PowerPoint writes and the only one carrying a status that can be resolved.
/// </summary>
public class PowerPointCommentTests
{
    private static OfficeAgentClient Client() => new(new PowerPointModule());

    [Fact]
    public void Comments_are_added_to_a_slide_and_surface_as_nodes()
    {
        var client = Client();

        var applied = Apply(client, PptxFactory.Deck(), new CommentOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Text = "Confirm the EMEA figure.",
            Author = "Reviewer",
            Initials = "RV"
        });

        var comment = Assert.Single(client.Inspect(applied).Nodes, n => n.Kind == "comment");
        Assert.StartsWith("comment#256/", comment.Path);
        Assert.Contains("Reviewer", comment.Summary);
        Assert.Contains("Confirm the EMEA figure.", comment.Summary);
        // A freshly added comment is open, not resolved. PowerPoint writes no status
        // attribute at all on a new comment - writing "active" instead is one of the
        // things that made it reject the package - so absence is the correct state, and
        // resolving is what puts a status there.
        Assert.DoesNotContain("(resolved)", comment.Summary);
        Assert.Null(StatusOf(applied));

        // Initials are what PowerPoint shows on the comment marker; the summary never
        // mentions them, so dropping them would go unnoticed.
        Assert.Equal("RV", AuthorInitialsOf(applied));

        AssertValid(applied);
    }

    /// <summary>The status recorded on the deck's single comment.</summary>
    private static PC.CommentStatus? StatusOf(byte[] deck)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);

        return document.PresentationPart!.SlideParts
            .SelectMany(s => s.GetPartsOfType<PowerPointCommentPart>())
            .SelectMany(p => p.CommentList!.Elements<PC.Comment>())
            .Select(c => c.Status?.Value)
            .Single();
    }

    /// <summary>The initials recorded on the deck's single comment author.</summary>
    private static string? AuthorInitialsOf(byte[] deck)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);

        return document.PresentationPart!
            .GetPartsOfType<PowerPointAuthorsPart>()
            .SelectMany(p => p.AuthorList!.Elements<PC.Author>())
            .Select(a => a.Initials?.Value)
            .Single();
    }

    [Fact]
    public void A_comment_can_be_resolved_and_keeps_its_text()
    {
        var client = Client();
        var withComment = Apply(client, PptxFactory.Deck(), new CommentOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Text = "Confirm the EMEA figure."
        });
        var path = Assert.Single(client.Inspect(withComment).Nodes, n => n.Kind == "comment").Path;

        var resolved = Apply(client, withComment, new CommentOp
        {
            Target = new NodeAnchor { Kind = "comment", Path = path },
            Action = CommentAction.Resolve
        });

        var comment = Assert.Single(client.Inspect(resolved).Nodes, n => n.Kind == "comment");
        Assert.Contains("(resolved)", comment.Summary);
        // Resolving keeps the review trail rather than deleting it.
        Assert.Contains("Confirm the EMEA figure.", comment.Summary);

        AssertStatus(resolved, PC.CommentStatus.Resolved);
        AssertValid(resolved);
    }

    [Fact]
    public void Two_comments_by_one_author_share_a_single_author_entry()
    {
        var client = Client();
        var once = Apply(client, PptxFactory.Deck(), new CommentOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Text = "First",
            Author = "Reviewer"
        });
        var twice = Apply(client, once, new CommentOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#257" },
            Text = "Second",
            Author = "Reviewer"
        });

        Assert.Equal(2, client.Inspect(twice).Nodes.Count(n => n.Kind == "comment"));

        // Duplicating the author entry makes PowerPoint list the same person repeatedly.
        using var stream = new MemoryStream(twice);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var authors = document.PresentationPart!
            .GetPartsOfType<PowerPointAuthorsPart>().Single()
            .AuthorList!.Elements<PC.Author>().ToList();
        Assert.Single(authors);
        Assert.Equal("Reviewer", authors[0].Name!.Value);
    }

    [Fact]
    public void Resolving_a_comment_twice_is_refused_rather_than_silently_repeated()
    {
        var client = Client();
        var withComment = Apply(client, PptxFactory.Deck(), new CommentOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Text = "Check this."
        });
        var path = Assert.Single(client.Inspect(withComment).Nodes, n => n.Kind == "comment").Path;
        var resolved = Apply(client, withComment, new CommentOp
        {
            Target = new NodeAnchor { Kind = "comment", Path = path },
            Action = CommentAction.Resolve
        });

        var again = Preview(client, resolved, new CommentOp
        {
            Target = new NodeAnchor { Kind = "comment", Path = path },
            Action = CommentAction.Resolve
        });

        Assert.False(again.IsValid);
        Assert.Contains("already resolved", Assert.Single(again.Errors).Message);
    }

    [Fact]
    public void Empty_text_missing_slides_and_missing_comments_are_reported()
    {
        var client = Client();
        var deck = PptxFactory.Deck();

        var noText = Preview(client, deck, new CommentOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Text = "   "
        });
        var noSlide = Preview(client, deck, new CommentOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#999" },
            Text = "Hello"
        });
        var noComment = Preview(client, deck, new CommentOp
        {
            Target = new NodeAnchor { Kind = "comment", Path = "comment#256/{missing}" },
            Action = CommentAction.Resolve
        });

        Assert.Equal(ValidationErrorCodes.InvalidOperation, Assert.Single(noText.Errors).Code);
        Assert.Equal(ValidationErrorCodes.AnchorNotFound, Assert.Single(noSlide.Errors).Code);
        Assert.Equal(ValidationErrorCodes.AnchorNotFound, Assert.Single(noComment.Errors).Code);
    }

    [Fact]
    public void Resolving_against_a_slide_target_explains_what_to_target_instead()
    {
        var client = Client();

        var report = Preview(client, PptxFactory.Deck(), new CommentOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
            Action = CommentAction.Resolve
        });

        Assert.False(report.IsValid);
        Assert.Contains("comment path", Assert.Single(report.Errors).Message);
    }

    // ── helpers ───────────────────────────────────────────────────────────────

    private static byte[] Apply(OfficeAgentClient client, byte[] deck, PlanOperation operation)
    {
        using var applied = client.Commit(new StreamHandle(new MemoryStream(deck)), Plan(operation));
        Assert.True(applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return applied.ToBytes();
    }

    private static ChangeReport Preview(OfficeAgentClient client, byte[] deck, PlanOperation operation) =>
        client.Preview(new StreamHandle(new MemoryStream(deck)), Plan(operation));

    private static DocumentPlan Plan(PlanOperation operation) => new()
    {
        Format = DocFormat.PowerPoint,
        Operations = new[] { operation }
    };

    private static void AssertStatus(byte[] deck, PC.CommentStatus expected)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var comment = document.PresentationPart!.SlideParts
            .SelectMany(p => p.GetPartsOfType<PowerPointCommentPart>())
            .SelectMany(p => p.CommentList!.Elements<PC.Comment>())
            .Single();
        Assert.Equal(expected, comment.Status!.Value);
    }

    private static void AssertValid(byte[] deck)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2021).Validate(document).ToList();
        Assert.True(errors.Count == 0,
            string.Join("; ", errors.Select(e => $"{e.Path?.XPath}: {e.Description}")));
    }
}
