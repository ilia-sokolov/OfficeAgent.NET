using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// Running heads, page numbers, and the distinct first page a cover needs.
/// </summary>
/// <remarks>
/// <c>w:sectPr</c> is a strict sequence: the header and footer references open it and
/// <c>w:titlePg</c> comes near the end, after the column settings. Appending any of them -
/// the obvious thing to do - produces a document Word offers to repair, so every test here
/// checks placement as well as content.
/// </remarks>
public class WordHeaderFooterTests
{
    private static OfficeAgentClient Client() => new(new WordModule());

    [Fact]
    public void A_header_is_created_and_referenced_from_the_section()
    {
        var client = Client();

        var document = Apply(client, Blank(), new HeaderFooterOp
        {
            Header = "Northwind Traders — Q2 Board Review"
        });

        using var stream = new MemoryStream(document);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var main = opened.MainDocumentPart!;

        var header = Assert.Single(main.HeaderParts);
        Assert.Equal("Northwind Traders — Q2 Board Review", header.Header!.InnerText);

        var section = main.Document.Body!.GetFirstChild<SectionProperties>()!;
        var reference = Assert.Single(section.Elements<HeaderReference>());
        Assert.Equal(HeaderFooterValues.Default, reference.Type!.Value);
        Assert.Equal(main.GetIdOfPart(header), reference.Id!.Value);

        // The reference opens w:sectPr, ahead of the page size the blank document sets.
        Assert.Equal(0, section.ChildElements.ToList().IndexOf(reference));
        AssertValid(document);
    }

    [Fact]
    public void A_page_number_is_a_field_rather_than_a_number()
    {
        var client = Client();

        var document = Apply(client, Blank(), new HeaderFooterOp { ShowPageNumber = true });

        using var stream = new MemoryStream(document);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var footer = Assert.Single(opened.MainDocumentPart!.FooterParts);

        // A literal number would be wrong on page two. The field keeps up.
        var code = Assert.Single(footer.Footer!.Descendants<FieldCode>());
        Assert.Contains("PAGE", code.Text);

        var chars = footer.Footer.Descendants<FieldChar>().Select(c => c.FieldCharType!.Value).ToList();
        Assert.Equal(
            new[] { FieldCharValues.Begin, FieldCharValues.Separate, FieldCharValues.End },
            chars);
        AssertValid(document);
    }

    [Fact]
    public void A_running_head_can_put_the_text_and_the_number_at_opposite_edges()
    {
        var client = Client();

        var document = Apply(client, Blank(), new HeaderFooterOp
        {
            Footer = "Confidential",
            ShowPageNumber = true,
            Alignment = "edges"
        });

        using var stream = new MemoryStream(document);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var footer = Assert.Single(opened.MainDocumentPart!.FooterParts);

        // A right tab at the text edge is what separates them; spaces would not hold.
        var tab = Assert.Single(footer.Footer!.Descendants<TabStop>());
        Assert.Equal(TabStopValues.Right, tab.Val!.Value);
        Assert.Equal(12240 - 1440 - 1440, tab.Position!.Value);
        Assert.Single(footer.Footer.Descendants<TabChar>());
        AssertValid(document);
    }

    [Fact]
    public void A_distinct_first_page_lets_a_cover_keep_the_running_head_off()
    {
        var client = Client();

        var document = Apply(client, Blank(),
            new HeaderFooterOp { DifferentFirstPage = true, Header = "Running head" },
            new HeaderFooterOp { Scope = "firstPage", Header = string.Empty });

        using var stream = new MemoryStream(document);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var main = opened.MainDocumentPart!;
        var section = main.Document.Body!.GetFirstChild<SectionProperties>()!;

        Assert.NotNull(section.GetFirstChild<TitlePage>());

        var references = section.Elements<HeaderReference>().ToList();
        Assert.Equal(2, references.Count);
        Assert.Contains(references, r => r.Type!.Value == HeaderFooterValues.First);

        var first = references.Single(r => r.Type!.Value == HeaderFooterValues.First);
        Assert.Equal(string.Empty, ((HeaderPart)main.GetPartById(first.Id!.Value!)).Header!.InnerText);

        var running = references.Single(r => r.Type!.Value == HeaderFooterValues.Default);
        Assert.Equal("Running head", ((HeaderPart)main.GetPartById(running.Id!.Value!)).Header!.InnerText);
        AssertValid(document);
    }

    [Fact]
    public void Title_page_lands_after_the_page_size_not_before_it()
    {
        var client = Client();

        var document = Apply(client, Blank(), new HeaderFooterOp { DifferentFirstPage = true });

        using var stream = new MemoryStream(document);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var section = opened.MainDocumentPart!.Document.Body!.GetFirstChild<SectionProperties>()!;

        var names = section.ChildElements.Select(c => c.LocalName).ToList();
        if (names.Contains("pgSz"))
            Assert.True(names.IndexOf("pgSz") < names.IndexOf("titlePg"), string.Join(", ", names));
        AssertValid(document);
    }

    [Fact]
    public void Writing_the_header_twice_replaces_the_text_rather_than_adding_to_it()
    {
        var client = Client();

        var once = Apply(client, Blank(), new HeaderFooterOp { Header = "Draft" });
        var twice = Apply(client, once, new HeaderFooterOp { Header = "Final" });

        using var stream = new MemoryStream(twice);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);

        var header = Assert.Single(opened.MainDocumentPart!.HeaderParts);
        Assert.Equal("Final", header.Header!.InnerText);
        AssertValid(twice);
    }

    [Fact]
    public void A_header_can_be_cleared()
    {
        var client = Client();

        var written = Apply(client, Blank(), new HeaderFooterOp { Header = "Draft" });
        var cleared = Apply(client, written, new HeaderFooterOp { Header = string.Empty });

        using var stream = new MemoryStream(cleared);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        Assert.Equal(string.Empty, Assert.Single(opened.MainDocumentPart!.HeaderParts).Header!.InnerText);
        AssertValid(cleared);
    }

    [Fact]
    public void Deck_only_settings_are_refused_rather_than_dropped()
    {
        var client = Client();

        var report = Preview(client, new HeaderFooterOp { Header = "Head", ShowSlideNumber = true });

        var error = Assert.Single(report.Errors);
        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        Assert.Contains("showPageNumber", error.Message);
    }

    [Fact]
    public void A_scope_that_is_not_a_page_kind_is_refused()
    {
        var client = Client();

        var report = Preview(client, new HeaderFooterOp { Header = "Head", Scope = "oddPage" });

        Assert.Contains("default, firstPage, or evenPage", Assert.Single(report.Errors).Message);
    }

    [Fact]
    public void An_operation_that_sets_nothing_says_what_it_wanted()
    {
        var client = Client();

        var report = Preview(client, new HeaderFooterOp());

        Assert.Contains("at least one of", Assert.Single(report.Errors).Message);
    }

    [Fact]
    public void A_logo_placed_in_the_header_keeps_its_bytes_in_the_header()
    {
        var client = Client();

        var withHeader = Apply(client, Blank(), new HeaderFooterOp { Header = "Acme Corp" });
        var paraId = client.Inspect(withHeader).Paragraphs.Single(p => p.Location == "header").ParaId;

        var withLogo = Apply(client, withHeader, new InsertImageOp
        {
            Target = new TextSpanAnchor { ParaId = paraId, Expect = string.Empty },
            Base64Bytes = Png,
            ImageType = "png",
            WidthPx = 60,
            HeightPx = 60,
            Position = InsertPosition.Before
        });

        using var stream = new MemoryStream(withLogo);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var main = opened.MainDocumentPart!;
        var header = Assert.Single(main.HeaderParts);

        // A relationship id means nothing outside its own part. Bytes added to the document
        // part for a drawing that lives in the header leave an id the header cannot resolve,
        // and Word calls that file corrupt - which is the whole point of this test.
        Assert.Single(header.ImageParts);
        Assert.Empty(main.ImageParts);

        var blip = Assert.Single(header.Header!.Descendants<DocumentFormat.OpenXml.Drawing.Blip>());
        Assert.NotNull(header.GetPartById(blip.Embed!.Value!));
        AssertValid(withLogo);
    }

    [Fact]
    public void Find_reports_the_host_a_match_lives_in()
    {
        var client = Client();

        var withHeader = Apply(client, Blank(), new HeaderFooterOp { Header = "Acme Corp" });
        var inspect = client.Inspect(withHeader);
        var document = Apply(client, withHeader,
            new ChangeTextOp
            {
                Target = new TextSpanAnchor
                {
                    ParaId = inspect.Paragraphs.First(p => p.Location != "header").ParaId,
                    Expect = string.Empty
                },
                With = "Acme Corp supplies the goods.",
                Mode = ChangeMode.Direct
            });

        var hits = client.Find(
            new StreamHandle(new MemoryStream(document)),
            new FindQuery("Acme Corp"));

        Assert.Equal(
            new[] { "body", "header" },
            hits.Select(h => h.Location).OrderBy(location => location).ToArray());
        AssertValid(document);
    }

    [Fact]
    public void The_same_logo_in_the_body_and_the_header_is_stored_in_both()
    {
        var client = Client();

        var withHeader = Apply(client, Blank(), new HeaderFooterOp { Header = "Acme Corp" });
        var inspect = client.Inspect(withHeader);

        var document = Apply(client, withHeader,
            new InsertImageOp
            {
                Target = new TextSpanAnchor
                {
                    ParaId = inspect.Paragraphs.Single(p => p.Location == "header").ParaId,
                    Expect = string.Empty
                },
                Base64Bytes = Png, ImageType = "png", WidthPx = 40, HeightPx = 40
            },
            new InsertImageOp
            {
                Target = new TextSpanAnchor
                {
                    ParaId = inspect.Paragraphs.First(p => p.Location != "header").ParaId,
                    Expect = string.Empty
                },
                Base64Bytes = Png, ImageType = "png", WidthPx = 40, HeightPx = 40
            });

        using var stream = new MemoryStream(document);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var main = opened.MainDocumentPart!;

        // Deduplication is per part, not per package: the two drawings live in different
        // parts and each needs a relationship its own part can resolve.
        Assert.Single(main.ImageParts);
        Assert.Single(Assert.Single(main.HeaderParts).ImageParts);
        AssertValid(document);
    }

    // ── helpers ───────────────────────────────────────────────────────────────

    /// <summary>A 1x1 PNG. The bytes only have to be a real image; nothing here renders one.</summary>
    private const string Png =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==";

    private static byte[] Blank() => new WordModule().CreateBlank();

    private static byte[] Apply(OfficeAgentClient client, byte[] document, params PlanOperation[] operations)
    {
        using var applied = client.Commit(
            new StreamHandle(new MemoryStream(document)),
            new DocumentPlan { Operations = operations });
        Assert.True(applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return applied.ToBytes();
    }

    private static ChangeReport Preview(OfficeAgentClient client, params PlanOperation[] operations) =>
        client.Preview(
            new StreamHandle(new MemoryStream(Blank())),
            new DocumentPlan { Operations = operations });

    private static void AssertValid(byte[] document)
    {
        using var stream = new MemoryStream(document);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(opened).ToList();
        Assert.True(errors.Count == 0,
            string.Join("; ", errors.Select(e => $"{e.Description} @ {e.Path?.XPath}")));
    }
}
