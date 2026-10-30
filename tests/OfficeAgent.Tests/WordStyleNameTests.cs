using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// A company template's styles, named the way its authors see them. Word shows a paragraph
/// whose style id it does not know as Normal, without a word, so a plan that names a style
/// the document lacks has to be refused, and one that names it by the name Word shows has
/// to land on the right style.
/// </summary>
public class WordStyleNameTests
{
    private static OfficeAgentClient Office() => new(new WordModule());
    private static StreamHandle Handle(byte[] bytes) => new(new MemoryStream(bytes), "doc.docx");

    [Fact]
    public void A_paragraph_takes_a_style_by_the_name_Word_shows()
    {
        var bytes = new WordModule().CreateBlank();

        var result = Commit(bytes, new FormatOp { Target = FirstParagraph(bytes), StyleId = "Heading 2" });

        Assert.Equal("Heading2", StyleOfFirstParagraph(result));
    }

    [Fact]
    public void A_company_style_defined_in_the_same_plan_is_found_by_its_name()
    {
        var bytes = new WordModule().CreateBlank();

        var result = Commit(bytes,
            new DefineStyleOp { StyleId = "BodyTextBlue", Name = "Body Text Blue", BasedOn = "Normal", Color = "1F4E79" },
            new InsertOp { Target = FirstParagraph(bytes), Position = InsertPosition.After, Text = "Scope", StyleId = "body text blue" });

        using var document = Open(result);
        var inserted = document.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().ElementAt(1);
        Assert.Equal("BodyTextBlue", inserted.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
    }

    [Fact]
    public void An_unknown_style_is_refused_with_the_document_own_styles_and_nothing_is_written()
    {
        var bytes = new WordModule().CreateBlank();

        using var result = Office().Commit(Handle(bytes), Plan(
            new FormatOp { Target = FirstParagraph(bytes), StyleId = "Clause Heading" }));

        Assert.False(result.Committed);
        var error = Assert.Single(result.Report.Errors);
        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        Assert.Contains("no paragraph style named 'Clause Heading'", error.Message);
        Assert.Contains("'heading 1'", error.Message);
    }

    [Fact]
    public void A_style_of_the_wrong_type_is_refused()
    {
        var bytes = new WordModule().CreateBlank();

        var report = Office().Preview(Handle(bytes), Plan(
            new FormatOp { Target = FirstParagraph(bytes), StyleId = "Table Grid" }));

        Assert.Contains("is a table style; this needs a paragraph style", Assert.Single(report.Errors).Message);
    }

    [Fact]
    public void A_built_in_style_the_document_lacks_is_added_as_Word_adds_it()
    {
        // A document with no styles part at all, as some generators write them.
        var bytes = NoStyles();

        var result = Commit(bytes, new FormatOp { Target = FirstParagraph(bytes), StyleId = "Heading 1" });

        Assert.Equal("Heading1", StyleOfFirstParagraph(result));
        using var document = Open(result);
        var ids = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Elements<Style>()
            .Select(s => s.StyleId!.Value).ToList();
        // Heading 1 is based on Normal, so Normal arrives first.
        Assert.Equal(new[] { "Normal", "Heading1" }, ids);
    }

    [Fact]
    public void A_definition_may_name_its_parent_by_the_name_Word_shows()
    {
        var bytes = new WordModule().CreateBlank();

        var result = Commit(bytes,
            new DefineStyleOp { StyleId = "ClauseHeading", Name = "Clause Heading", BasedOn = "heading 1", Next = "Normal" });

        using var document = Open(result);
        var style = document.MainDocumentPart!.StyleDefinitionsPart!.Styles!.Elements<Style>()
            .Single(s => s.StyleId?.Value == "ClauseHeading");
        Assert.Equal("Heading1", style.BasedOn?.Val?.Value);
    }

    // ── helpers ───────────────────────────────────────────────────────────────

    private static TextSpanAnchor FirstParagraph(byte[] bytes) =>
        new() { ParaId = Office().Inspect(Handle(bytes)).Paragraphs.First().ParaId, Expect = "" };

    private static byte[] Commit(byte[] bytes, params PlanOperation[] operations)
    {
        using var result = Office().Commit(Handle(bytes), Plan(operations));
        Assert.True(result.Committed,
            string.Join("; ", result.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return OfficeAgentClient.ToBytes(result);
    }

    private static DocumentPlan Plan(params PlanOperation[] operations) => new() { Operations = operations };

    private static string? StyleOfFirstParagraph(byte[] bytes)
    {
        using var document = Open(bytes);
        return document.MainDocumentPart!.Document!.Body!.Elements<Paragraph>().First()
            .ParagraphProperties?.ParagraphStyleId?.Val?.Value;
    }

    private static WordprocessingDocument Open(byte[] bytes) =>
        WordprocessingDocument.Open(new MemoryStream(bytes), false);

    private static byte[] NoStyles()
    {
        var buffer = new MemoryStream();
        var blank = new WordModule().CreateBlank();
        buffer.Write(blank, 0, blank.Length);
        using (var document = WordprocessingDocument.Open(buffer, true))
        {
            var main = document.MainDocumentPart!;
            main.DeletePart(main.StyleDefinitionsPart!);
        }
        return buffer.ToArray();
    }
}
