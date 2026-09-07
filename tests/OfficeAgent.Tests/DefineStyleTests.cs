using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// A style states a look once. Without it a plan has to write the same direct formatting
/// onto every paragraph, which produces a document that cannot be maintained: changing the
/// heading colour means finding every heading, and an author who edits the style in Word
/// sees nothing happen.
/// </summary>
public class DefineStyleTests
{
    private static OfficeAgentClient Office() => new(new WordModule());
    private static StreamHandle Handle(byte[] bytes) => new(new MemoryStream(bytes), "doc.docx");

    [Fact]
    public void A_new_style_is_added_to_the_catalogue_with_its_properties()
    {
        var result = Office().Commit(Handle(new WordModule().CreateBlank()), Plan(new DefineStyleOp
        {
            StyleId = "Quote",
            Name = "Pull Quote",
            BasedOn = "Normal",
            Next = "Normal",
            FontFamily = "Georgia",
            SizeHalfPoints = 28,
            Italic = true,
            Color = "444444",
            Alignment = "center",
            IndentLeftTwips = 720,
            SpacingBeforeTwips = 240
        }));

        Assert.True(result.Committed);
        var style = StyleIn(result, "Quote");

        Assert.Equal(StyleValues.Paragraph, style.Type!.Value);
        Assert.Equal("Pull Quote", style.StyleName!.Val!.Value);
        Assert.Equal("Normal", style.GetFirstChild<BasedOn>()!.Val!.Value);
        Assert.Equal("Normal", style.GetFirstChild<NextParagraphStyle>()!.Val!.Value);

        var runProperties = style.GetFirstChild<StyleRunProperties>()!;
        Assert.Equal("Georgia", runProperties.GetFirstChild<RunFonts>()!.Ascii!.Value);
        Assert.Equal("28", runProperties.GetFirstChild<FontSize>()!.Val!.Value);
        Assert.NotNull(runProperties.GetFirstChild<Italic>());
        Assert.Equal("444444", runProperties.GetFirstChild<Color>()!.Val!.Value);

        var paragraphProperties = style.GetFirstChild<StyleParagraphProperties>()!;
        Assert.Equal(JustificationValues.Center, paragraphProperties.GetFirstChild<Justification>()!.Val!.Value);
        Assert.Equal("720", paragraphProperties.GetFirstChild<Indentation>()!.Left!.Value);
        Assert.Equal("240", paragraphProperties.GetFirstChild<SpacingBetweenLines>()!.Before!.Value);
    }

    [Fact]
    public void A_style_definition_produces_a_document_that_validates()
    {
        // The properties are deliberately written in an order the schema does not declare
        // them in, because the containers are sequences: a colour before a bold, or a
        // basedOn after the properties, is what turns this into a repair prompt.
        var result = Office().Commit(Handle(new WordModule().CreateBlank()), Plan(new DefineStyleOp
        {
            StyleId = "Callout",
            Color = "AA0000",
            Underline = true,
            Bold = true,
            SizeHalfPoints = 24,
            FontFamily = "Arial",
            BasedOn = "Normal",
            OutlineLevel = 2,
            BorderStyle = "single",
            BorderSizeEighths = 8,
            BorderColor = "AA0000",
            SpacingAfterTwips = 120,
            IndentLeftTwips = 360
        }));

        Assert.True(result.Committed);

        using var document = WordprocessingDocument.Open(
            new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document).ToList();
        Assert.True(errors.Count == 0, string.Join("; ", errors.Select(e => $"{e.Description} at {e.Path?.XPath}")));
    }

    [Fact]
    public void Defining_a_style_that_exists_updates_it_and_keeps_what_was_not_mentioned()
    {
        var first = Office().Commit(Handle(new WordModule().CreateBlank()), Plan(new DefineStyleOp
        {
            StyleId = "Quote",
            Name = "Pull Quote",
            FontFamily = "Georgia",
            SizeHalfPoints = 28
        }));

        var second = Office().Commit(
            Handle(OfficeAgentClient.ToBytes(first)),
            Plan(new DefineStyleOp { StyleId = "Quote", SizeHalfPoints = 22 }));

        Assert.True(second.Committed);
        var style = StyleIn(second, "Quote");
        var runProperties = style.GetFirstChild<StyleRunProperties>()!;

        Assert.Equal("22", runProperties.GetFirstChild<FontSize>()!.Val!.Value);

        // Untouched properties survive, so "make the quotes smaller" is one property rather
        // than a restatement of the whole style.
        Assert.Equal("Georgia", runProperties.GetFirstChild<RunFonts>()!.Ascii!.Value);
        Assert.Equal("Pull Quote", style.StyleName!.Val!.Value);
    }

    [Fact]
    public void A_style_can_be_defined_in_a_document_that_has_no_style_catalogue_yet()
    {
        var result = Office().Commit(
            Handle(DocxFactory.Contract()),
            Plan(new DefineStyleOp { StyleId = "Quote", Italic = true }));

        Assert.True(result.Committed);
        Assert.Equal("Quote", StyleIn(result, "Quote").StyleId!.Value);
    }

    [Fact]
    public void A_paragraph_carrying_the_style_reports_it_after_the_plan_runs()
    {
        var bytes = DocxFactory.Contract();
        var paraId = Office().Inspect(Handle(bytes)).Paragraphs
            .First(p => p.Text == DocxFactory.ClauseText).ParaId;

        // Definition and use in one plan: the style has to exist by the time the format
        // operation names it, which is the whole point of defining it first.
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new DefineStyleOp { StyleId = "Quote", Name = "Pull Quote", Italic = true },
                new FormatOp
                {
                    Target = new TextSpanAnchor { ParaId = paraId, Expect = string.Empty },
                    StyleId = "Quote"
                }
            }
        };

        var result = Office().Commit(Handle(bytes), plan);
        Assert.True(result.Committed, string.Join("; ", result.Report.Errors.Select(e => e.Message)));

        var inspect = Office().Inspect(Handle(OfficeAgentClient.ToBytes(result)));
        Assert.Contains(inspect.Paragraphs, p => p.StyleId == "Quote");
    }

    [Fact]
    public void A_style_based_on_one_defined_earlier_in_the_same_plan_is_accepted()
    {
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new DefineStyleOp { StyleId = "House", FontFamily = "Georgia" },
                new DefineStyleOp { StyleId = "HouseQuote", BasedOn = "House", Italic = true }
            }
        };

        var report = Office().Preview(Handle(new WordModule().CreateBlank()), plan);

        Assert.True(report.IsValid);
    }

    [Fact]
    public void A_style_based_on_one_nothing_defines_is_refused()
    {
        var report = Office().Preview(
            Handle(new WordModule().CreateBlank()),
            Plan(new DefineStyleOp { StyleId = "Quote", BasedOn = "NoSuchStyle" }));

        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, e => e.Message.Contains("NoSuchStyle", StringComparison.Ordinal));
    }

    [Fact]
    public void A_style_based_on_itself_is_refused()
    {
        var report = Office().Preview(
            Handle(new WordModule().CreateBlank()),
            Plan(new DefineStyleOp { StyleId = "Quote", BasedOn = "Quote" }));

        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, e => e.Message.Contains("itself", StringComparison.Ordinal));
    }

    [Fact]
    public void A_cycle_through_two_styles_is_refused()
    {
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new DefineStyleOp { StyleId = "A", BasedOn = "B" },
                new DefineStyleOp { StyleId = "B", BasedOn = "A" }
            }
        };

        var report = Office().Preview(Handle(new WordModule().CreateBlank()), plan);

        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, e => e.Message.Contains("inherit from itself", StringComparison.Ordinal));
    }

    [Theory]
    [InlineData("", "styleId")]
    [InlineData("Pull Quote", "whitespace")]
    public void An_unusable_style_id_is_refused(string styleId, string expected)
    {
        var report = Office().Preview(
            Handle(new WordModule().CreateBlank()),
            Plan(new DefineStyleOp { StyleId = styleId, Italic = true }));

        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, e => e.Message.Contains(expected, StringComparison.Ordinal));
    }

    [Fact]
    public void An_unknown_style_type_is_refused_by_name()
    {
        var report = Office().Preview(
            Handle(new WordModule().CreateBlank()),
            Plan(new DefineStyleOp { StyleId = "Quote", Type = "sidebar" }));

        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, e => e.Message.Contains("paragraph, character, table", StringComparison.Ordinal));
    }

    [Fact]
    public void A_highlight_is_refused_because_a_style_cannot_carry_one()
    {
        // WordprocessingML rejects w:highlight inside a style outright - not a question of
        // ordering. Accepting it and dropping it would return a style that silently differs
        // from what was asked for.
        var report = Office().Preview(
            Handle(new WordModule().CreateBlank()),
            Plan(new DefineStyleOp { StyleId = "Quote", Highlight = "yellow" }));

        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, e => e.Message.Contains("cannot carry a highlight", StringComparison.Ordinal));
    }

    [Fact]
    public void A_character_style_is_defined_as_a_character_style()
    {
        var result = Office().Commit(
            Handle(new WordModule().CreateBlank()),
            Plan(new DefineStyleOp { StyleId = "Term", Type = "character", Bold = true }));

        Assert.True(result.Committed);
        var style = StyleIn(result, "Term");

        Assert.Equal(StyleValues.Character, style.Type!.Value);

        // A character style carries no paragraph properties, whatever the vocabulary allows.
        Assert.Null(style.GetFirstChild<StyleParagraphProperties>());
    }

    private static DocumentPlan Plan(params PlanOperation[] operations) => new() { Operations = operations };

    private static Style StyleIn(ApplyResult result, string styleId)
    {
        using var document = WordprocessingDocument.Open(
            new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var styles = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
        return Assert.Single(styles.Elements<Style>().Where(s => s.StyleId?.Value == styleId));
    }
}
