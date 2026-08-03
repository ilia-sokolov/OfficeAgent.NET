using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Tests;

/// <summary>
/// Degenerate decks and malformed operations. The contract these lock in: <c>Preview</c>
/// never throws - every reachable bad state comes back as a <see cref="ValidationError"/>
/// the agent can read - and a plan that fails validation writes nothing.
/// </summary>
/// <remarks>
/// A deck built by another tool can carry shapes this module never creates: a table with
/// no <c>a:tblGrid</c>, a table with no rows, a presentation with no slides. Those are
/// possible states rather than impossible ones, so they belong in validation rather than
/// in an exception thrown half-way through an apply.
/// </remarks>
public class PowerPointEdgeCaseTests
{
    private static OfficeAgentClient Client() => new(new PowerPointModule());

    private static readonly NodeAnchor Table = new() { Kind = "table", Path = "table#256/3" };
    private static readonly NodeAnchor Slide = new() { Kind = "slide", Path = "slide#256" };

    [Fact]
    public void A_deck_with_no_slides_inspects_and_refuses_edits_cleanly()
    {
        var client = Client();
        var empty = EmptyDeck();

        var inspection = client.Inspect(empty);
        Assert.Empty(inspection.Paragraphs);
        Assert.Empty(inspection.Outline);

        var report = Preview(empty, new InsertTableOp
        {
            Target = Slide,
            Table = new TableData { Headers = new[] { "A" } }
        });
        Assert.Equal(ValidationErrorCodes.AnchorNotFound, Assert.Single(report.Errors).Code);
    }

    [Fact]
    public void A_table_with_no_column_grid_is_refused_by_validation_not_by_a_crash()
    {
        var gridless = GridlessTableDeck();

        // Preview used to pass these - ColumnCount reports 0 for a missing grid, and only
        // Before/After range-checked it - and Apply then hit a `?? throw` mid-plan.
        var insert = Preview(gridless, new InsertTableColumnsOp
        {
            Target = Table,
            Columns = new[] { new[] { "x" } }
        });
        var remove = Preview(gridless, new RemoveTableColumnsOp { Target = Table });

        Assert.Equal(ValidationErrorCodes.InvalidOperation, Assert.Single(insert.Errors).Code);
        Assert.Equal(ValidationErrorCodes.InvalidOperation, Assert.Single(remove.Errors).Code);

        // And committing the same plan writes nothing rather than throwing.
        using var applied = Client().Commit(
            new StreamHandle(new MemoryStream(gridless)),
            Plan(new InsertTableColumnsOp { Target = Table, Columns = new[] { new[] { "x" } } }));
        Assert.False(applied.Committed);
    }

    [Fact]
    public void A_table_with_no_rows_accepts_an_append_and_refuses_a_positional_insert()
    {
        var rowless = RowlessTableDeck();

        // Nothing to position against…
        var positional = Preview(rowless, new InsertTableRowsOp
        {
            Target = Table,
            Rows = new[] { new[] { "a", "b" } },
            Position = TablePosition.After,
            RowIndex = 0
        });
        Assert.Equal(ValidationErrorCodes.InvalidOperation, Assert.Single(positional.Errors).Code);

        // …but appending into an empty table is well defined.
        var appended = Apply(rowless, new InsertTableRowsOp
        {
            Target = Table,
            Rows = new[] { new[] { "a", "b" } }
        });
        Assert.Contains("1×2", Assert.Single(Client().Inspect(appended).Nodes, n => n.Kind == "table").Summary);
    }

    [Fact]
    public void Duplicate_indices_remove_each_row_or_column_once()
    {
        // [0,0] must not remove two rows, and must not fail: the second mention is a
        // no-op, not a second victim.
        var rows = Apply(PptxFactory.DeckWithTable(), new RemoveTableRowsOp
        {
            Target = Table,
            RowIndices = new[] { 0, 0 }
        });
        Assert.Contains("1×2", Assert.Single(Client().Inspect(rows).Nodes, n => n.Kind == "table").Summary);

        var columns = Apply(PptxFactory.DeckWithTable(), new RemoveTableColumnsOp
        {
            Target = Table,
            ColumnIndices = new[] { 0, 0 }
        });
        Assert.Contains("2×1", Assert.Single(Client().Inspect(columns).Nodes, n => n.Kind == "table").Summary);
    }

    [Fact]
    public void A_negative_row_index_counts_from_the_end()
    {
        var applied = Apply(PptxFactory.DeckWithTable(), new InsertTableRowsOp
        {
            Target = Table,
            Rows = new[] { new[] { "LAST", "9" } },
            Position = TablePosition.After,
            RowIndex = -1
        });

        // -1 is the last row, so the new row lands after it.
        var texts = Client().Inspect(applied).Paragraphs
            .Where(p => p.ParaId.Contains("r2c0")).Select(p => p.Text);
        Assert.Contains("LAST", texts);
    }

    [Fact]
    public void Column_data_longer_than_the_table_is_truncated_not_crashed()
    {
        var applied = Apply(PptxFactory.DeckWithTable(), new InsertTableColumnsOp
        {
            Target = Table,
            // Five entries for a two-row table: the surplus has nowhere to go.
            Columns = new[] { new[] { "a", "b", "c", "d", "e" } }
        });

        Assert.Contains("2×3", Assert.Single(Client().Inspect(applied).Nodes, n => n.Kind == "table").Summary);
    }

    [Theory]
    [InlineData("not-a-path")]
    [InlineData("")]
    [InlineData("slide#")]
    [InlineData("slide#-1")]
    [InlineData("slide#99999999999999999999")]
    public void A_malformed_slide_path_is_reported_rather_than_parsed_loosely(string path)
    {
        var report = Preview(PptxFactory.Deck(), new InsertTableOp
        {
            Target = new NodeAnchor { Kind = "slide", Path = path },
            Table = new TableData { Headers = new[] { "A" } }
        });

        Assert.Equal(ValidationErrorCodes.AnchorNotFound, Assert.Single(report.Errors).Code);
    }

    [Fact]
    public void Resolving_a_comment_that_does_not_exist_is_reported()
    {
        var report = Preview(PptxFactory.Deck(), new CommentOp
        {
            Target = new NodeAnchor { Kind = "comment", Path = "comment#256/{MISSING}" },
            Action = CommentAction.Resolve
        });

        Assert.Equal(ValidationErrorCodes.AnchorNotFound, Assert.Single(report.Errors).Code);
    }

    // ── helpers ───────────────────────────────────────────────────────────────

    private static DocumentPlan Plan(PlanOperation operation) => new()
    {
        Format = DocFormat.PowerPoint,
        Operations = new[] { operation }
    };

    private static ChangeReport Preview(byte[] deck, PlanOperation operation) =>
        Client().Preview(new StreamHandle(new MemoryStream(deck)), Plan(operation));

    private static byte[] Apply(byte[] deck, PlanOperation operation)
    {
        using var applied = Client().Commit(new StreamHandle(new MemoryStream(deck)), Plan(operation));
        Assert.True(applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return applied.ToBytes();
    }

    /// <summary>A valid presentation with a master and layout but no slides at all.</summary>
    private static byte[] EmptyDeck()
    {
        var bytes = new PowerPointModule().CreateBlank();
        using var stream = new MemoryStream();
        stream.Write(bytes, 0, bytes.Length);
        stream.Position = 0;
        using (var document = PresentationDocument.Open(stream, isEditable: true))
        {
            var presentation = document.PresentationPart!;
            foreach (var id in presentation.Presentation.SlideIdList!.Elements<SlideId>().ToList())
                id.Remove();
            presentation.Presentation.Save();
        }
        return stream.ToArray();
    }

    private static byte[] MutateTable(Action<A.Table> mutate)
    {
        var bytes = PptxFactory.DeckWithTable();
        using var stream = new MemoryStream();
        stream.Write(bytes, 0, bytes.Length);
        stream.Position = 0;
        using (var document = PresentationDocument.Open(stream, isEditable: true))
        {
            var part = document.PresentationPart!.SlideParts.First();
            mutate(part.Slide.Descendants<A.Table>().First());
            part.Slide.Save();
        }
        return stream.ToArray();
    }

    /// <summary>A deck whose table lost its <c>a:tblGrid</c>, as a foreign tool might leave it.</summary>
    private static byte[] GridlessTableDeck() => MutateTable(t => t.TableGrid!.Remove());

    /// <summary>A deck whose table has a grid but no rows.</summary>
    private static byte[] RowlessTableDeck() => MutateTable(t =>
    {
        foreach (var row in t.Elements<A.TableRow>().ToList()) row.Remove();
    });
}
