using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// List numbering that Word owns, rather than digits typed into the text.
/// </summary>
/// <remarks>
/// The difference matters most in the documents that most need it. Insert a clause into the
/// middle of a contract whose numbers are literal text and everything below it is now wrong,
/// silently - which is how a signed agreement ends up with two clause 7s. A real
/// <c>w:numPr</c> renumbers on its own.
/// </remarks>
public class WordNumberingTests
{
    private static OfficeAgentClient Client() => new(new WordModule());

    [Fact]
    public void A_paragraph_can_be_made_a_numbered_clause()
    {
        var client = Client();
        var document = Apply(client, WithLines("Term", "Payment"),
            Number("Term", "clause", 0),
            Number("Payment", "clause", 0));

        using var stream = new MemoryStream(document);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var main = opened.MainDocumentPart!;

        // The numbering part is created on first use; a blank document has none.
        Assert.NotNull(main.NumberingDefinitionsPart);

        var numbering = main.NumberingDefinitionsPart!.Numbering!;
        var instance = Assert.Single(numbering.Elements<NumberingInstance>());
        var definition = Assert.Single(numbering.Elements<AbstractNum>());

        // Both clauses point at the same instance, so they are one running sequence rather
        // than two lists that both start at 1.
        var referenced = ParagraphsOf(opened)
            .Where(p => p.ParagraphProperties?.NumberingProperties is not null)
            .Select(p => p.ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value)
            .ToList();

        Assert.Equal(2, referenced.Count);
        Assert.All(referenced, id => Assert.Equal(instance.NumberID!.Value, id));
        Assert.Equal(definition.AbstractNumberId!.Value, instance.AbstractNumId!.Val!.Value);
        AssertValid(document);
    }

    [Fact]
    public void A_clause_level_prints_its_parents()
    {
        var client = Client();
        var document = Apply(client, WithLines("Top", "Sub"),
            Number("Top", "clause", 0),
            Number("Sub", "clause", 1));

        using var stream = new MemoryStream(document);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var levels = opened.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!
            .Elements<AbstractNum>().Single()
            .Elements<Level>().ToList();

        // 1.  then 1.1  then 1.1.1 - which is what makes a clause citable from another one.
        Assert.Equal("%1.", levels[0].LevelText!.Val!.Value);
        Assert.Equal("%1.%2", levels[1].LevelText!.Val!.Value);
        Assert.Equal("%1.%2.%3", levels[2].LevelText!.Val!.Value);

        // The paragraph asked for depth 1 and got it.
        var sub = ParagraphsOf(opened).Single(p => p.InnerText == "Sub");
        Assert.Equal(1, sub.ParagraphProperties!.NumberingProperties!.NumberingLevelReference!.Val!.Value);
        AssertValid(document);
    }

    [Fact]
    public void A_second_list_id_starts_a_separate_sequence()
    {
        var client = Client();

        // A manual whose second chapter restarts its steps at 1.
        var document = Apply(client, WithLines("Step one", "Step two"),
            Number("Step one", "decimal", 0, listId: 0),
            Number("Step two", "decimal", 0, listId: 1));

        using var stream = new MemoryStream(document);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var numbering = opened.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;

        Assert.Equal(2, numbering.Elements<NumberingInstance>().Count());

        var ids = ParagraphsOf(opened)
            .Where(p => p.ParagraphProperties?.NumberingProperties is not null)
            .Select(p => p.ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value)
            .ToList();
        Assert.Equal(2, ids.Distinct().Count());
        AssertValid(document);
    }

    [Fact]
    public void The_same_list_asked_for_twice_reuses_one_definition()
    {
        var client = Client();

        // Composed over two plans, the way a real caller builds a document.
        var once = Apply(client, WithLines("A", "B"), Number("A", "bullet", 0));
        var twice = Apply(client, once, Number("B", "bullet", 0));

        using var stream = new MemoryStream(twice);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var numbering = opened.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;

        Assert.Single(numbering.Elements<AbstractNum>());
        Assert.Single(numbering.Elements<NumberingInstance>());
        AssertValid(twice);
    }

    [Fact]
    public void Numbering_can_be_taken_off_again()
    {
        var client = Client();
        var numbered = Apply(client, WithLines("A", "B"), Number("A", "clause", 0));
        var plain = Apply(client, numbered, Number("A", "none", 0));

        using var stream = new MemoryStream(plain);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);

        var paragraph = ParagraphsOf(opened).Single(p => p.InnerText == "A");
        Assert.Null(paragraph.ParagraphProperties?.NumberingProperties);
        AssertValid(plain);
    }

    [Fact]
    public void Numbering_lands_ahead_of_the_indent_and_the_spacing()
    {
        var client = Client();
        var document = Apply(client, WithLines("A", "B"), new FormatOp
        {
            // Named by text; Resolve turns it into the real id against this same document.
            Target = new TextSpanAnchor { ParaId = "A", Expect = string.Empty },
            ListStyle = "clause",
            IndentLeftTwips = 480,
            SpacingBeforeTwips = 120
        });

        using var stream = new MemoryStream(document);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var properties = ParagraphsOf(opened).Single(p => p.InnerText == "A").ParagraphProperties!;

        // w:pPr is a sequence: w:numPr precedes w:spacing and w:ind however late it is asked
        // for, or Word offers to repair the file.
        var names = properties.ChildElements.Select(c => c.LocalName).ToList();
        Assert.True(names.IndexOf("numPr") < names.IndexOf("spacing"), string.Join(", ", names));
        Assert.True(names.IndexOf("numPr") < names.IndexOf("ind"), string.Join(", ", names));
        AssertValid(document);
    }

    [Fact]
    public void A_numbered_paragraph_can_also_start_a_page()
    {
        var client = Client();

        // The combination a manual needs on its first section, and the one that gets the
        // order wrong: w:pageBreakBefore precedes w:numPr in w:pPr, not the other way round.
        var document = Apply(client, WithLines("Chapter", "Body"), new FormatOp
        {
            Target = new TextSpanAnchor { ParaId = "Chapter", Expect = string.Empty },
            ListStyle = "clause",
            PageBreakBefore = true,
            SpacingBeforeTwips = 400
        });

        using var stream = new MemoryStream(document);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var names = ParagraphsOf(opened).Single(p => p.InnerText == "Chapter")
            .ParagraphProperties!.ChildElements.Select(c => c.LocalName).ToList();

        Assert.True(names.IndexOf("pageBreakBefore") < names.IndexOf("numPr"), string.Join(", ", names));
        AssertValid(document);
    }

    [Fact]
    public void A_look_that_is_not_a_list_style_is_refused()
    {
        var client = Client();
        var report = Preview(client, WithLines("A"), Number("A", "roman", 0));

        var error = Assert.Single(report.Errors);
        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        Assert.Contains("bullet, decimal, clause, none", error.Message);
    }

    [Fact]
    public void A_level_past_the_ninth_is_refused()
    {
        var client = Client();
        var report = Preview(client, WithLines("A"), Number("A", "clause", 9));

        Assert.Contains("between 0 and 8", Assert.Single(report.Errors).Message);
    }

    [Fact]
    public void A_level_with_no_style_to_belong_to_is_refused()
    {
        var client = Client();
        var report = Preview(client, WithLines("A"), new FormatOp
        {
            Target = new TextSpanAnchor { ParaId = "A", Expect = string.Empty },
            ListLevel = 2
        });

        Assert.Contains("need a listStyle", Assert.Single(report.Errors).Message);
    }

    [Fact]
    public void A_deck_refuses_list_numbering_and_says_what_to_use()
    {
        var client = new OfficeAgentClient(new OfficeAgent.PowerPoint.PowerPointModule());

        var report = client.Preview(
            new StreamHandle(new MemoryStream(new OfficeAgent.PowerPoint.PowerPointModule().CreateBlank())),
            new DocumentPlan
            {
                Format = OfficeAgent.Abstractions.DocumentFormat.PowerPoint,
                Operations = new PlanOperation[]
                {
                    new FormatOp
                    {
                        Target = new TextSpanAnchor { ParaId = "slide256/shape2/p0", Expect = string.Empty },
                        ListStyle = "clause"
                    }
                }
            });

        Assert.Contains("bullets come from its layout", Assert.Single(report.Errors).Message);
    }

    // ── helpers ───────────────────────────────────────────────────────────────

    /// <summary>A blank document carrying one paragraph per line of text.</summary>
    private static byte[] WithLines(params string[] lines)
    {
        var client = Client();
        var document = new WordModule().CreateBlank();

        var first = client.Inspect(document).Paragraphs.Single().ParaId;
        document = Apply(client, document, new ChangeTextOp
        {
            Target = new TextSpanAnchor { ParaId = first, Expect = string.Empty },
            With = lines[0],
            Mode = ChangeMode.Direct
        });

        foreach (var line in lines.Skip(1))
        {
            var last = client.Inspect(document).Paragraphs.Last().ParaId;
            document = Apply(client, document, new InsertOp
            {
                Target = new TextSpanAnchor { ParaId = last, Expect = string.Empty },
                Position = InsertPosition.After,
                Text = line
            });
        }

        return document;
    }

    private static FormatOp Number(string text, string style, int level, int? listId = null) => new()
    {
        Target = new TextSpanAnchor { ParaId = text, Expect = string.Empty },
        ListStyle = style,
        ListLevel = level,
        ListId = listId
    };

    /// <summary>
    /// Rewrites the placeholder ids in <see cref="Number"/> - which name paragraphs by their
    /// text so the tests stay readable - into the real ones.
    /// </summary>
    private static PlanOperation[] Resolve(OfficeAgentClient client, byte[] document, PlanOperation[] operations)
    {
        var byText = client.Inspect(document).Paragraphs
            .GroupBy(p => p.Text)
            .ToDictionary(g => g.Key, g => g.First().ParaId, StringComparer.Ordinal);

        return operations.Select(op =>
        {
            if (op is FormatOp format &&
                format.Target is TextSpanAnchor anchor &&
                byText.TryGetValue(anchor.ParaId, out var real))
                return new FormatOp
                {
                    Target = new TextSpanAnchor { ParaId = real, Expect = anchor.Expect },
                    ListStyle = format.ListStyle,
                    ListLevel = format.ListLevel,
                    ListId = format.ListId,
                    IndentLeftTwips = format.IndentLeftTwips,
                    SpacingBeforeTwips = format.SpacingBeforeTwips,
                    PageBreakBefore = format.PageBreakBefore
                };
            return op;
        }).ToArray();
    }

    private static byte[] Apply(OfficeAgentClient client, byte[] document, params PlanOperation[] operations)
    {
        using var applied = client.Commit(
            new StreamHandle(new MemoryStream(document)),
            new DocumentPlan { Operations = Resolve(client, document, operations) });

        Assert.True(applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return applied.ToBytes();
    }

    private static ChangeReport Preview(
        OfficeAgentClient client, byte[] document, params PlanOperation[] operations) =>
        client.Preview(
            new StreamHandle(new MemoryStream(document)),
            new DocumentPlan { Operations = Resolve(client, document, operations) });

    private static List<Paragraph> ParagraphsOf(WordprocessingDocument document) =>
        document.MainDocumentPart!.Document.Body!.Elements<Paragraph>().ToList();

    private static void AssertValid(byte[] document)
    {
        using var stream = new MemoryStream(document);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(opened).ToList();
        Assert.True(errors.Count == 0,
            string.Join("; ", errors.Select(e => $"{e.Description} @ {e.Path?.XPath}")));
    }
}
