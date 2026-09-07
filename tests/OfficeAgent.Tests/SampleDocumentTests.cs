using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// The sample contract in <c>samples/documents</c> and the result the README promises for it.
/// </summary>
/// <remarks>
/// A first-run experience that does not do what the page said is worse than no sample at
/// all: the reader cannot tell whether they misconfigured something or the project
/// oversold itself. These pin the two things the README claims - that the payment clause
/// becomes a tracked change, and that nothing else in the document moves.
/// </remarks>
public class SampleDocumentTests
{
    private static byte[] Sample()
    {
        // Walk up from the test binary to the repository root.
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "OfficeAgent.NET.sln")))
            directory = directory.Parent;

        Assert.NotNull(directory);
        var path = Path.Combine(directory!.FullName, "samples", "documents", "services-agreement.docx");
        Assert.True(File.Exists(path), $"The sample document is missing from {path}.");
        return File.ReadAllBytes(path);
    }

    [Fact]
    public void The_sample_carries_what_the_documentation_says_it_does()
    {
        using var stream = new MemoryStream(Sample());
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        var main = package.MainDocumentPart!;
        var body = main.Document.Body!;

        // The clause the README asks the reader to change.
        Assert.Contains("within thirty days of receipt", body.InnerText);

        // A table to edit, an open comment to answer, and a revision to review - the three
        // things that make it a document under review rather than a blank page.
        Assert.Single(body.Descendants<Table>());
        Assert.Single(main.WordprocessingCommentsPart!.Comments!.Elements<Comment>());
        Assert.Contains("Priya Raman",
            main.WordprocessingCommentsPart.Comments.Elements<Comment>().Single().Author?.Value);
        Assert.NotEmpty(body.Descendants<InsertedRun>());
        Assert.NotEmpty(body.Descendants<DeletedRun>());

        var problems = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(package).ToList();
        Assert.True(problems.Count == 0,
            string.Join("; ", problems.Take(3).Select(p => $"{p.Path?.XPath}: {p.Description}")));
    }

    [Fact]
    public void The_first_edit_lands_as_a_redline_and_disturbs_nothing_else()
    {
        var client = new OfficeAgentClient(new WordModule());
        var before = Sample();

        var clause = client.Inspect(before).Paragraphs
            .First(p => p.Text.Contains("thirty days", StringComparison.Ordinal));

        // Exactly the edit the README's prompt describes, with no mode stated - so it takes
        // the same Tracked default a reader's connection would.
        using var applied = client.Commit(
            new StreamHandle(new MemoryStream(before)),
            new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new ChangeTextOp
                    {
                        Target = new TextSpanAnchor { ParaId = clause.ParaId, Expect = "thirty days" },
                        With = "forty-five days"
                    }
                }
            });

        Assert.True(applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));

        using var stream = new MemoryStream(applied.ToBytes());
        using var package = WordprocessingDocument.Open(stream, isEditable: false);
        var main = package.MainDocumentPart!;
        var body = main.Document.Body!;

        // "reads forty-five days, as a tracked change"
        Assert.Contains("forty-five days", body.Descendants<InsertedRun>().Select(r => r.InnerText));
        Assert.Contains("thirty days", body.Descendants<DeletedText>().Select(t => t.Text));

        // "everything else is exactly as it was" - the claim worth pinning, because it is
        // what distinguishes this from extracting the text and writing a new file.
        Assert.Single(body.Descendants<Table>());
        Assert.Equal(4, body.Descendants<Table>().Single().Elements<TableRow>().Count());
        Assert.Single(main.WordprocessingCommentsPart!.Comments!.Elements<Comment>());
        Assert.Contains("3% above base rate", body.Descendants<InsertedRun>().Select(r => r.InnerText));
        Assert.Contains("2% above base rate", body.Descendants<DeletedText>().Select(t => t.Text));
        Assert.Contains("Master Services Agreement", body.InnerText);

        var problems = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(package).ToList();
        Assert.True(problems.Count == 0,
            string.Join("; ", problems.Take(3).Select(p => $"{p.Path?.XPath}: {p.Description}")));
    }
}
