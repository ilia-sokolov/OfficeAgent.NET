using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.PowerPoint;

namespace OfficeAgent.Tests;

/// <summary>
/// Relationship-graph invariants for presentations. These sit outside what
/// <c>OpenXmlValidator</c> checks: it validates each part's XML against the schema, not
/// whether the parts reference one another the way PowerPoint requires. A deck can be
/// schema-clean and still be refused with "PowerPoint found a problem with content".
/// </summary>
public class PowerPointPackageIntegrityTests
{
    [Fact]
    public void Every_slide_layout_references_its_slide_master()
    {
        // The bug this pins down: creating a layout under a master writes only the
        // master→layout relationship. Without the reverse one, real PowerPoint declares
        // the package corrupt and offers to repair it, while every schema check passes.
        AssertLayoutsReferenceMasters(new PowerPointModule().CreateBlank(), "blank deck");
        AssertLayoutsReferenceMasters(PptxFactory.Deck(), "fixture deck");
        AssertLayoutsReferenceMasters(PptxFactory.DeckWithTable(), "fixture table deck");
    }

    [Fact]
    public void Editing_a_deck_preserves_the_layout_to_master_relationship()
    {
        var client = new OfficeAgentClient(new PowerPointModule());

        using var applied = client.Commit(
            new StreamHandle(new MemoryStream(PptxFactory.Deck())),
            new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new InsertTableOp
                    {
                        Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
                        Table = new TableData { Headers = new[] { "A" }, Rows = new[] { new[] { "1" } } }
                    }
                }
            });

        Assert.True(applied.Committed);
        AssertLayoutsReferenceMasters(applied.ToBytes(), "edited deck");
    }

    [Fact]
    public void Every_slide_references_a_layout_and_every_slide_id_resolves()
    {
        using var stream = new MemoryStream(new PowerPointModule().CreateBlank());
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var presentation = document.PresentationPart!;

        foreach (var slide in presentation.SlideParts)
            Assert.True(slide.SlideLayoutPart is not null,
                "a slide with no layout relationship cannot be rendered");

        // Every id in p:sldIdLst must resolve to a slide part that exists.
        foreach (var slideId in presentation.Presentation.SlideIdList!
                     .Elements<DocumentFormat.OpenXml.Presentation.SlideId>())
        {
            var relationshipId = slideId.RelationshipId?.Value;
            Assert.False(string.IsNullOrEmpty(relationshipId));
            Assert.IsType<SlidePart>(presentation.GetPartById(relationshipId!));
        }

        // …and the master a slide's layout belongs to must be one the presentation lists.
        Assert.NotEmpty(presentation.SlideMasterParts);
        Assert.NotNull(presentation.SlideMasterParts.First().ThemePart);
    }

    [Fact]
    public void An_added_comment_carries_the_slide_moniker_powerpoint_requires()
    {
        var client = new OfficeAgentClient(new PowerPointModule());

        using var applied = client.Commit(
            new StreamHandle(new MemoryStream(PptxFactory.Deck())),
            new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new CommentOp
                    {
                        Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
                        Text = "Check this figure.",
                        Author = "Reviewer",
                        Initials = "RV"
                    }
                }
            });

        Assert.True(applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));

        using var stream = new MemoryStream(applied.ToBytes());
        using var document = PresentationDocument.Open(stream, isEditable: false);

        var comment = document.PresentationPart!.SlideParts
            .SelectMany(s => s.GetPartsOfType<PowerPointCommentPart>())
            .SelectMany(p => p.CommentList!
                .Elements<DocumentFormat.OpenXml.Office2021.PowerPoint.Comment.Comment>())
            .Single();

        // Without the moniker PowerPoint cannot tell which slide the comment belongs to
        // and rejects the whole package - a failure schema validation never sees.
        var moniker = comment
            .GetFirstChild<DocumentFormat.OpenXml.Office2016.Presentation.Command.SlideMonikerList>();
        Assert.NotNull(moniker);
        var slideMoniker = moniker!
            .GetFirstChild<DocumentFormat.OpenXml.Office2016.Presentation.Command.SlideMoniker>();
        Assert.NotNull(slideMoniker);
        Assert.Equal(256U, slideMoniker!.SldId?.Value);

        // PowerPoint writes no status on a new comment; resolving is what sets one.
        Assert.Null(comment.Status?.Value);
    }

    private static void AssertLayoutsReferenceMasters(byte[] deck, string what)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);

        var layouts = 0;
        foreach (var master in document.PresentationPart!.SlideMasterParts)
            foreach (var layout in master.SlideLayoutParts)
            {
                layouts++;
                Assert.True(layout.SlideMasterPart is not null,
                    $"{what}: slide layout has no relationship back to its master, " +
                    "which PowerPoint treats as a corrupt package");
            }

        Assert.True(layouts > 0, $"{what}: expected at least one slide layout");
    }
}
