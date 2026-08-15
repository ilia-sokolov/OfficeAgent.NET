using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// Writing into a paragraph that has no text. A blank document is one empty paragraph, so
/// without this there is no way to put the first words into a document the caller just
/// created - <c>insert</c> would leave the empty paragraph above them.
/// </summary>
/// <remarks>
/// An empty <c>expect</c> stays refused against a paragraph that has text: there it cannot
/// be told apart from a caller who left the field out, and acting on the guess would
/// rewrite a paragraph nobody named.
/// </remarks>
public class WordEmptyParagraphTests
{
    private static OfficeAgentClient Client() => new(new WordModule());

    [Fact]
    public void An_empty_paragraph_can_be_filled_in()
    {
        var client = Client();
        var blank = new WordModule().CreateBlank();

        var paraId = client.Inspect(blank).Paragraphs.Single().ParaId;
        var filled = Apply(client, blank, new ChangeTextOp
        {
            Target = new TextSpanAnchor { ParaId = paraId, Expect = string.Empty },
            With = "Northwind Traders — Q2 Board Review",
            Mode = ChangeMode.Direct
        });

        var paragraph = Assert.Single(client.Inspect(filled).Paragraphs);
        Assert.Equal("Northwind Traders — Q2 Board Review", paragraph.Text);

        // A blank document's paragraph carries no w14:paraId, so inspect names it with a
        // synthetic one. Writing to it mints a real id - which means a caller that wants to
        // style what it just wrote has to inspect again rather than reuse the id it passed.
        Assert.StartsWith("auto-", paraId);
        Assert.StartsWith("w14:", paragraph.ParaId);
        AssertValid(filled);
    }

    [Fact]
    public void Filling_an_empty_paragraph_with_tracking_on_is_recorded_as_an_insertion()
    {
        var client = Client();
        var blank = new WordModule().CreateBlank();
        var paraId = client.Inspect(blank).Paragraphs.Single().ParaId;

        var filled = Apply(client, blank, new ChangeTextOp
        {
            Target = new TextSpanAnchor { ParaId = paraId, Expect = string.Empty },
            With = "Draft for review",
            Mode = ChangeMode.Tracked
        });

        using var stream = new MemoryStream(filled);
        using var document = WordprocessingDocument.Open(stream, isEditable: false);
        var inserted = Assert.Single(document.MainDocumentPart!.Document.Body!.Descendants<InsertedRun>());

        Assert.Equal("Draft for review", inserted.InnerText);
        Assert.Equal("OfficeAgent", inserted.Author!.Value);

        // Still one paragraph of readable text, and Word can still open it.
        Assert.Equal("Draft for review", Assert.Single(client.Inspect(filled).Paragraphs).Text);
        AssertValid(filled);
    }

    [Fact]
    public void An_empty_expect_is_still_refused_where_there_is_text_to_lose()
    {
        var client = Client();

        var report = Preview(client, DocxFactory.Contract(), new ChangeTextOp
        {
            Target = new TextSpanAnchor { ParaId = "w14:00000002", Expect = string.Empty },
            With = "Something else"
        });

        var error = Assert.Single(report.Errors);
        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        Assert.Contains("non-empty 'expect'", error.Message);
    }

    [Fact]
    public void A_paragraph_an_insert_created_comes_back_with_a_stable_id()
    {
        var client = Client();
        var blank = new WordModule().CreateBlank();

        var start = client.Inspect(blank).Paragraphs.Single().ParaId;
        var written = Apply(client, blank,
            new ChangeTextOp
            {
                Target = new TextSpanAnchor { ParaId = start, Expect = string.Empty },
                With = "Title",
                Mode = ChangeMode.Direct
            });

        var titleId = client.Inspect(written).Paragraphs.Single().ParaId;
        var grown = Apply(client, written, new InsertOp
        {
            Target = new TextSpanAnchor { ParaId = titleId, Expect = string.Empty },
            Position = InsertPosition.After,
            Text = "Body"
        });

        var paragraphs = client.Inspect(grown).Paragraphs.ToList();
        Assert.Equal(new[] { "Title", "Body" }, paragraphs.Select(p => p.Text));

        // Both are addressable by a real id. Without this the new paragraph is named by its
        // position, and the name moves the moment anything is inserted above it - so a
        // caller composing a document over several plans styles the wrong line.
        Assert.All(paragraphs, p => Assert.StartsWith("w14:", p.ParaId));
        Assert.Equal(titleId, paragraphs[0].ParaId);

        // And the new id keeps working against the document it came from.
        var styled = Apply(client, grown, new FormatOp
        {
            Target = new TextSpanAnchor { ParaId = paragraphs[1].ParaId, Expect = string.Empty },
            Bold = true
        });

        Assert.Equal("Body", client.Inspect(styled).Paragraphs.Last().Text);
        AssertValid(styled);
    }

    [Fact]
    public void Filling_an_empty_paragraph_with_nothing_changes_nothing()
    {
        var client = Client();
        var blank = new WordModule().CreateBlank();
        var paraId = client.Inspect(blank).Paragraphs.Single().ParaId;

        var applied = Apply(client, blank, new ChangeTextOp
        {
            Target = new TextSpanAnchor { ParaId = paraId, Expect = string.Empty },
            With = string.Empty,
            Mode = ChangeMode.Direct
        });

        Assert.Equal(string.Empty, Assert.Single(client.Inspect(applied).Paragraphs).Text);
        AssertValid(applied);
    }

    // ── helpers ───────────────────────────────────────────────────────────────

    private static byte[] Apply(OfficeAgentClient client, byte[] document, params PlanOperation[] operations)
    {
        using var applied = client.Commit(
            new StreamHandle(new MemoryStream(document)),
            new DocumentPlan { Operations = operations });

        Assert.True(applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));

        return applied.ToBytes();
    }

    private static ChangeReport Preview(
        OfficeAgentClient client, byte[] document, params PlanOperation[] operations) =>
        client.Preview(
            new StreamHandle(new MemoryStream(document)),
            new DocumentPlan { Operations = operations });

    private static void AssertValid(byte[] document)
    {
        using var stream = new MemoryStream(document);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);

        var errors = new OpenXmlValidator(DocumentFormat.OpenXml.FileFormatVersions.Office2019)
            .Validate(opened).ToList();

        Assert.True(errors.Count == 0,
            string.Join("; ", errors.Select(e => $"{e.Description} @ {e.Path?.XPath}")));
    }
}
