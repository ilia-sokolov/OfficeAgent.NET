using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.PowerPoint;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;
using PC = DocumentFormat.OpenXml.Office2021.PowerPoint.Comment;
using PCMD = DocumentFormat.OpenXml.Office2016.Presentation.Command;

namespace OfficeAgent.Tests;

/// <summary>
/// What a duplicated slide owns and what it shares. A copy that shares a chart part with its
/// source is a deck PowerPoint refuses to open, and a chart write on either slide lands on both;
/// the same holds for every other part a slide edits in place.
/// </summary>
public class SlideDuplicateTests
{
    private static OfficeAgentClient Client() => new(new PowerPointModule());

    [Fact]
    public void Each_duplicate_of_a_chart_slide_gets_its_own_chart_and_workbook()
    {
        var client = Client();
        var deck = ChartDeck(client);

        // Two plans, as an agent sends them: the same slide may not be duplicated twice in one.
        var once = Apply(client, deck, new DuplicateSlideOp { Target = SlideAnchor(256) });
        var twice = Apply(client, once, new DuplicateSlideOp { Target = SlideAnchor(256), Position = SlidePosition.End });

        using var stream = new MemoryStream(twice);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var slides = OrderedParts(document).ToList();
        Assert.Equal(3, slides.Count);

        // Every slide's c:chart r:id resolves, on that slide, to a chart part no other slide has.
        var charts = slides.Select(ChartOf).ToList();
        Assert.Equal(3, charts.Select(c => c.Uri).Distinct().Count());
        Assert.Equal(3, document.PresentationPart!.SlideParts.SelectMany(s => s.ChartParts).Distinct().Count());

        // Each chart's c:externalData r:id resolves to a workbook of its own.
        var workbooks = charts.Select(WorkbookOf).ToList();
        Assert.Equal(3, workbooks.Select(w => w.Uri).Distinct().Count());
        Assert.All(workbooks, w => Assert.Equal(WorkbookBytes(workbooks[0]), WorkbookBytes(w)));

        Assert.All(charts, c => Assert.Equal(new[] { 10.0, 12.0, 15.0 }, ValuesOf(c)));
        AssertValid(twice);
    }

    [Fact]
    public void Chart_data_set_on_a_copy_leaves_the_original_and_the_other_copy_alone()
    {
        var client = Client();
        var once = Apply(client, ChartDeck(client), new DuplicateSlideOp { Target = SlideAnchor(256) });
        var twice = Apply(client, once, new DuplicateSlideOp { Target = SlideAnchor(256), Position = SlidePosition.End });
        var shapeId = ChartShapeId(client, twice, 256);
        var firstCopy = SlideIdAt(twice, 1);
        var secondCopy = SlideIdAt(twice, 2);
        var originalWorkbook = WorkbookBytesAt(twice, 0);

        var edited = Apply(client, twice,
            SetData(firstCopy, shapeId, 1, 2, 3),
            SetData(secondCopy, shapeId, 4, 5, 6));

        Assert.Equal(new[] { 10.0, 12.0, 15.0 }, ValuesAt(edited, 0));
        Assert.Equal(new[] { 1.0, 2.0, 3.0 }, ValuesAt(edited, 1));
        Assert.Equal(new[] { 4.0, 5.0, 6.0 }, ValuesAt(edited, 2));
        Assert.Equal(originalWorkbook, WorkbookBytesAt(edited, 0));
        Assert.NotEqual(originalWorkbook, WorkbookBytesAt(edited, 1));
        AssertValid(edited);
    }

    [Fact]
    public void A_chart_written_earlier_in_the_same_plan_is_copied_as_written()
    {
        var client = Client();
        var deck = ChartDeck(client);
        var shapeId = ChartShapeId(client, deck, 256);

        var edited = Apply(client, deck,
            SetData(256, shapeId, 7, 8, 9),
            new DuplicateSlideOp { Target = SlideAnchor(256) });

        Assert.Equal(new[] { 7.0, 8.0, 9.0 }, ValuesAt(edited, 0));
        Assert.Equal(new[] { 7.0, 8.0, 9.0 }, ValuesAt(edited, 1));
        Assert.Equal(WorkbookBytesAt(edited, 0), WorkbookBytesAt(edited, 1));
        AssertValid(edited);
    }

    [Fact]
    public void Notes_edited_earlier_in_the_same_plan_are_copied_as_edited()
    {
        var client = Client();
        var deck = Apply(client, new PowerPointModule().CreateBlank(), new InsertSlideOp
        {
            Slide = new SlideData { Title = "Delivery risks", Notes = "Do not commit to a recovery date." }
        });
        var source = SlideIdAt(deck, 1);

        // The edited notes are still only in memory when the duplicate runs: the copy must be
        // made from them, not from the part's last saved bytes.
        var edited = Apply(client, deck,
            new ChangeTextOp
            {
                Target = new TextSpanAnchor { ParaId = $"slide{source}/notes/shape2/p0", Expect = "Do not commit to a recovery date." },
                With = "Recovery plan due 17 August.",
                Mode = ChangeMode.Direct
            },
            new DuplicateSlideOp { Target = SlideAnchor(source) });

        var notes = client.Inspect(edited).Paragraphs.Where(p => p.Location == "notes").Select(p => p.Text).ToList();
        Assert.Equal(new[] { "Recovery plan due 17 August.", "Recovery plan due 17 August." }, notes);

        // The copy has notes of its own.
        using var stream = new MemoryStream(edited);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var slides = OrderedParts(document).ToList();
        Assert.NotEqual(slides[1].NotesSlidePart!.Uri, slides[2].NotesSlidePart!.Uri);
        AssertValid(edited);
    }

    [Fact]
    public void Chart_styles_smartart_and_tags_are_copied_and_links_kept()
    {
        var client = Client();
        var deck = RichSlideDeck(ChartDeck(client));

        var copied = Apply(client, deck, new DuplicateSlideOp { Target = SlideAnchor(256) });

        using var stream = new MemoryStream(copied);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var slides = OrderedParts(document).ToList();
        var (source, copy) = (slides[0], slides[1]);

        // The chart's own children, under the same ids so the chart XML resolves them.
        var (sourceChart, copyChart) = (ChartOf(source), ChartOf(copy));
        foreach (var (id, part) in sourceChart.Parts.Select(p => (p.RelationshipId, p.OpenXmlPart)))
        {
            var twin = copyChart.GetPartById(id);
            Assert.Equal(part.GetType(), twin.GetType());
            Assert.NotEqual(part.Uri, twin.Uri);
            Assert.Equal(Bytes(part), Bytes(twin));
        }
        Assert.Single(copyChart.GetPartsOfType<ChartStylePart>());
        Assert.Single(copyChart.GetPartsOfType<ChartColorStylePart>());

        // SmartArt parts and tags are the slide's own content.
        foreach (var id in new[] { "rIdDm", "rIdLo", "rIdQs", "rIdCs", "rIdDr", "rIdTags" })
        {
            var (from, to) = (source.GetPartById(id), copy.GetPartById(id));
            Assert.Equal(from.GetType(), to.GetType());
            Assert.NotEqual(from.Uri, to.Uri);
            Assert.Equal(Bytes(from), Bytes(to));
        }

        // The layout is shared, and the hyperlink survives under its id.
        Assert.Equal(source.SlideLayoutPart!.Uri, copy.SlideLayoutPart!.Uri);
        var link = Assert.Single(copy.HyperlinkRelationships);
        Assert.Equal("rIdLink", link.Id);
        Assert.Equal(new Uri("https://example.com/q3"), link.Uri);

        AssertNoNewSchemaErrors(deck, copied);
    }

    [Fact]
    public void A_duplicated_video_keeps_its_media_relationships()
    {
        var client = Client();
        var deck = Apply(client, new PowerPointModule().CreateBlank(), new InsertMediaOp
        {
            Target = SlideAnchor(256),
            Kind = MediaKind.Video,
            Base64Bytes = "AAAAIGZ0eXBpc29tAAACAGlzb21pc28yYXZjMW1wNDE=",
            MediaType = "mp4",
            PosterBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==",
            XPx = 100, YPx = 120, WidthPx = 480, HeightPx = 270
        });

        var copied = Apply(client, deck, new DuplicateSlideOp { Target = SlideAnchor(256) });

        using var stream = new MemoryStream(copied);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var slides = OrderedParts(document).ToList();
        foreach (var slide in slides)
        {
            // Without these the copy's a:videoFile r:link and p14:media r:embed dangle.
            Assert.Single(slide.DataPartReferenceRelationships.OfType<VideoReferenceRelationship>());
            Assert.Single(slide.DataPartReferenceRelationships.OfType<MediaReferenceRelationship>());
        }
        // The video bytes themselves are shared, like an image.
        Assert.Equal(
            slides[0].DataPartReferenceRelationships.First().DataPart.Uri,
            slides[1].DataPartReferenceRelationships.First().DataPart.Uri);
        AssertValid(copied);
    }

    [Fact]
    public void A_duplicated_slide_carries_its_comments_addressed_to_the_copy()
    {
        var client = Client();
        var commented = Apply(client, CreationIdDeck(), new CommentOp
        {
            Target = SlideAnchor(256),
            Text = "Confirm the EMEA figure.",
            Author = "Reviewer",
            Initials = "RV"
        });

        var copied = Apply(client, commented, new DuplicateSlideOp { Target = SlideAnchor(256) });
        var copyId = SlideIdAt(copied, 1);

        using (var stream = new MemoryStream(copied))
        using (var document = PresentationDocument.Open(stream, isEditable: false))
        {
            var slides = OrderedParts(document).ToList();
            var (sourcePart, copyPart) = (CommentPartOf(slides[0]), CommentPartOf(slides[1]));
            Assert.NotEqual(sourcePart.Uri, copyPart.Uri);

            // The copy has a creation id of its own, and its comment's moniker names the copy.
            var sourceCreation = CreationIdOf(slides[0]);
            var copyCreation = CreationIdOf(slides[1]);
            Assert.NotEqual(sourceCreation, copyCreation);
            var moniker = copyPart.CommentList!.Descendants<PCMD.SlideMoniker>().Single();
            Assert.Equal(copyId, moniker.SldId!.Value);
            Assert.Equal(copyCreation, moniker.CId!.Value);
            Assert.Equal(256U, sourcePart.CommentList!.Descendants<PCMD.SlideMoniker>().Single().SldId!.Value);

            Assert.NotEqual(
                sourcePart.CommentList.Elements<PC.Comment>().Single().Id!.Value,
                copyPart.CommentList.Elements<PC.Comment>().Single().Id!.Value);
        }

        // Both comments surface, one on each slide, and resolving the copy's leaves the original open.
        var comments = client.Inspect(copied).Nodes.Where(n => n.Kind == "comment").ToList();
        Assert.Equal(2, comments.Count);
        Assert.Contains(comments, c => c.Path.StartsWith("comment#256/", StringComparison.Ordinal));
        Assert.Contains(comments, c => c.Path.StartsWith($"comment#{copyId}/", StringComparison.Ordinal));
        AssertValid(copied);
    }

    // ── fixtures ──────────────────────────────────────────────────────────────────────────

    private static NodeAnchor SlideAnchor(uint slideId) => new() { Kind = "slide", Path = $"slide#{slideId}" };

    /// <summary>A one-slide deck holding a native column chart with an embedded workbook.</summary>
    internal static byte[] ChartDeck(OfficeAgentClient client) =>
        Apply(client, new PowerPointModule().CreateBlank(), new InsertChartOp
        {
            Target = SlideAnchor(256),
            Kind = ChartKind.ClusteredColumn,
            Title = "Revenue",
            Categories = new[] { "Q1", "Q2", "Q3" },
            Series = new[] { new ChartSeries { Name = "Revenue", Values = new double?[] { 10, 12, 15 } } },
            Description = "Revenue by quarter"
        });

    private static SetChartDataOp SetData(uint slideId, uint shapeId, params double[] values) => new()
    {
        Target = new NodeAnchor { Kind = "chart", Path = $"chart#{slideId}/{shapeId}" },
        Categories = new[] { "Q1", "Q2", "Q3" },
        Series = new[] { new ChartSeries { Name = "Revenue", Values = values.Select(v => (double?)v).ToArray() } }
    };

    private static uint ChartShapeId(OfficeAgentClient client, byte[] deck, uint slideId)
    {
        var path = client.Inspect(deck).Nodes.Single(n => n.Kind == "chart" && n.Path.StartsWith($"chart#{slideId}/", StringComparison.Ordinal)).Path;
        return uint.Parse(path.Substring(path.IndexOf('/') + 1), CultureInfo.InvariantCulture);
    }

    /// <summary>
    /// The chart deck as a PowerPoint-authored template has it: chart style and colour parts under
    /// the chart, SmartArt and tags on the slide, an external hyperlink, and a creation id.
    /// </summary>
    private static byte[] RichSlideDeck(byte[] chartDeck)
    {
        var stream = new MemoryStream();
        stream.Write(chartDeck, 0, chartDeck.Length);
        using (var document = PresentationDocument.Open(stream, isEditable: true))
        {
            var slide = document.PresentationPart!.SlideParts.Single();
            var chart = slide.ChartParts.Single();
            const string cs = "http://schemas.microsoft.com/office/drawing/2012/chartStyle";
            const string a = "http://schemas.openxmlformats.org/drawingml/2006/main";
            const string dgm = "http://schemas.openxmlformats.org/drawingml/2006/diagram";
            Feed(chart.AddNewPart<ChartStylePart>("rIdStyle"), $"<cs:chartStyle xmlns:cs=\"{cs}\" id=\"201\"/>");
            Feed(chart.AddNewPart<ChartColorStylePart>("rIdColors"),
                $"<cs:colorStyle xmlns:cs=\"{cs}\" xmlns:a=\"{a}\" meth=\"cycle\" id=\"10\"><a:schemeClr val=\"accent1\"/></cs:colorStyle>");

            Feed(slide.AddNewPart<DiagramDataPart>("rIdDm"),
                $"<dgm:dataModel xmlns:dgm=\"{dgm}\"><dgm:ptLst><dgm:pt modelId=\"{{00000000-0000-0000-0000-000000000001}}\" type=\"doc\"/></dgm:ptLst></dgm:dataModel>");
            Feed(slide.AddNewPart<DiagramLayoutDefinitionPart>("rIdLo"),
                $"<dgm:layoutDef xmlns:dgm=\"{dgm}\" uniqueId=\"urn:test/layout\"><dgm:layoutNode/></dgm:layoutDef>");
            Feed(slide.AddNewPart<DiagramStylePart>("rIdQs"), $"<dgm:styleDef xmlns:dgm=\"{dgm}\" uniqueId=\"urn:test/style\"/>");
            Feed(slide.AddNewPart<DiagramColorsPart>("rIdCs"), $"<dgm:colorsDef xmlns:dgm=\"{dgm}\" uniqueId=\"urn:test/colors\"/>");
            Feed(slide.AddNewPart<DiagramPersistLayoutPart>("rIdDr"),
                "<dsp:drawing xmlns:dsp=\"http://schemas.microsoft.com/office/drawing/2008/diagram\"><dsp:spTree><dsp:nvGrpSpPr><dsp:cNvPr id=\"0\" name=\"\"/><dsp:cNvGrpSpPr/></dsp:nvGrpSpPr><dsp:grpSpPr/></dsp:spTree></dsp:drawing>");
            Feed(slide.AddNewPart<UserDefinedTagsPart>("rIdTags"),
                "<p:tagLst xmlns:p=\"http://schemas.openxmlformats.org/presentationml/2006/main\"><p:tag name=\"REGION\" val=\"EMEA\"/></p:tagLst>");
            slide.AddHyperlinkRelationship(new Uri("https://example.com/q3"), true, "rIdLink");
        }
        return stream.ToArray();
    }

    /// <summary>A blank deck whose slide carries a creation id, as PowerPoint writes it.</summary>
    private static byte[] CreationIdDeck()
    {
        var blank = new PowerPointModule().CreateBlank();
        var stream = new MemoryStream();
        stream.Write(blank, 0, blank.Length);
        using (var document = PresentationDocument.Open(stream, isEditable: true))
        {
            var slide = document.PresentationPart!.SlideParts.Single().Slide;
            slide.Append(new SlideExtensionList(new SlideExtension(
                new P14.CreationId { Val = 1234567890U })
            { Uri = "{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}" }));
        }
        return stream.ToArray();
    }

    private static void Feed(OpenXmlPart part, string xml)
    {
        using var data = new MemoryStream(Encoding.UTF8.GetBytes(xml));
        part.FeedData(data);
    }

    // ── reading ───────────────────────────────────────────────────────────────────────────

    private static byte[] Apply(OfficeAgentClient client, byte[] deck, params PlanOperation[] operations)
    {
        using var applied = client.Commit(new StreamHandle(new MemoryStream(deck)),
            new DocumentPlan { Format = DocFormat.PowerPoint, Operations = operations });
        Assert.True(applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return applied.ToBytes();
    }

    private static IEnumerable<SlidePart> OrderedParts(PresentationDocument document)
    {
        var main = document.PresentationPart!;
        foreach (var entry in main.Presentation.SlideIdList!.Elements<SlideId>())
            yield return (SlidePart)main.GetPartById(entry.RelationshipId!);
    }

    private static uint SlideIdAt(byte[] deck, int index)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        return document.PresentationPart!.Presentation.SlideIdList!.Elements<SlideId>().ElementAt(index).Id!.Value;
    }

    /// <summary>The chart part the slide's single c:chart reference resolves to on that slide.</summary>
    private static ChartPart ChartOf(SlidePart slide)
    {
        var id = slide.Slide.Descendants<C.ChartReference>().Single().Id!.Value!;
        return Assert.IsType<ChartPart>(slide.GetPartById(id));
    }

    private static EmbeddedPackagePart WorkbookOf(ChartPart chart)
    {
        var id = chart.ChartSpace.GetFirstChild<C.ExternalData>()!.Id!.Value!;
        return Assert.IsType<EmbeddedPackagePart>(chart.GetPartById(id));
    }

    private static double[] ValuesOf(ChartPart chart) =>
        chart.ChartSpace.Descendants<C.Values>().First().Descendants<C.NumericValue>()
            .Select(v => double.Parse(v.Text, CultureInfo.InvariantCulture)).ToArray();

    private static double[] ValuesAt(byte[] deck, int index)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        return ValuesOf(ChartOf(OrderedParts(document).ElementAt(index)));
    }

    private static byte[] WorkbookBytesAt(byte[] deck, int index)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        return WorkbookBytes(WorkbookOf(ChartOf(OrderedParts(document).ElementAt(index))));
    }

    private static byte[] WorkbookBytes(EmbeddedPackagePart part) => Bytes(part);

    private static byte[] Bytes(OpenXmlPart part)
    {
        using var input = part.GetStream(FileMode.Open, FileAccess.Read);
        using var copy = new MemoryStream();
        input.CopyTo(copy);
        return copy.ToArray();
    }

    private static PowerPointCommentPart CommentPartOf(SlidePart slide) =>
        slide.GetPartsOfType<PowerPointCommentPart>().Single();

    /// <summary>The slide's creation id, which a slide read from a file holds as an unknown element.</summary>
    private static uint CreationIdOf(SlidePart slide) =>
        uint.Parse(slide.Slide.Descendants()
            .Single(e => e.LocalName == "creationId" && e.NamespaceUri == "http://schemas.microsoft.com/office/powerpoint/2010/main")
            .GetAttribute("val", string.Empty).Value!, CultureInfo.InvariantCulture);

    private static IEnumerable<string> SchemaErrors(byte[] deck)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        return new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document)
            .Select(e => e.Description ?? string.Empty).ToList();
    }

    private static void AssertValid(byte[] deck)
    {
        var problems = SchemaErrors(deck).ToList();
        Assert.True(problems.Count == 0, string.Join("; ", problems.Take(3)));
    }

    /// <summary>
    /// The fixture's hand-written style and SmartArt parts are minimal rather than complete, so
    /// the copy is held to adding no kind of error the input did not already have.
    /// </summary>
    private static void AssertNoNewSchemaErrors(byte[] input, byte[] output)
    {
        var added = SchemaErrors(output).Except(SchemaErrors(input)).ToList();
        Assert.True(added.Count == 0, string.Join("; ", added.Take(3)));
    }
}
