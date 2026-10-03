using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// Deleting a span with <c>changeText</c> and an empty <c>with</c> is one revision: a
/// <c>w:del</c> and nothing else. An empty <c>w:ins</c> beside it is a second revision that
/// changes nothing, which Word's Review pane counts and steps through (#18).
/// </summary>
public class TrackedDeletionTests
{
    private const string Before = "The term lasts five years";
    private const string Deleted = " unless terminated earlier under clause 6";
    private const string Original = Before + Deleted + ".";

    private static OfficeAgentClient Client() => new(new WordModule());

    [Fact]
    public void A_tracked_deletion_writes_one_deletion_and_no_insertion()
    {
        var result = Delete(Contract(), ChangeMode.Tracked);

        AssertValid(result);
        using var body = Body(result);
        Assert.Single(body.Root.Descendants<DeletedRun>());
        Assert.Empty(body.Root.Descendants<InsertedRun>());
        Assert.Equal(Deleted, string.Concat(body.Root.Descendants<DeletedText>().Select(t => t.Text)));
        Assert.Equal(Before + ".", VisibleText(body.Root));
    }

    [Theory]
    [InlineData(RevisionAction.Accept, Before + ".")]
    [InlineData(RevisionAction.Reject, Original)]
    public void Resolving_a_tracked_deletion_removes_or_restores_the_text_exactly(RevisionAction action, string expected)
    {
        var resolved = Apply(Delete(Contract(), ChangeMode.Tracked), new RevisionOp
        {
            Target = new NodeAnchor { Kind = "revision", Path = "all" },
            Action = action
        });

        AssertValid(resolved);
        using var body = Body(resolved);
        Assert.Empty(body.Root.Descendants<DeletedRun>());
        Assert.Empty(body.Root.Descendants<InsertedRun>());
        Assert.Equal(expected, VisibleText(body.Root));
    }

    [Fact]
    public void A_tracked_deletion_inside_a_comment_range_keeps_the_comment_anchors()
    {
        var result = Delete(Contract(withComment: true), ChangeMode.Tracked);

        AssertValid(result);
        using var body = Body(result);
        Assert.Single(body.Root.Descendants<CommentRangeStart>());
        Assert.Single(body.Root.Descendants<CommentRangeEnd>());
        Assert.Single(body.Root.Descendants<CommentReference>());
        Assert.Single(body.Root.Descendants<DeletedRun>());
        Assert.Empty(body.Root.Descendants<InsertedRun>());
    }

    [Fact]
    public void A_direct_deletion_leaves_no_empty_run_behind()
    {
        var result = Delete(Contract(), ChangeMode.Direct);

        AssertValid(result);
        using var body = Body(result);
        Assert.Equal(Before + ".", VisibleText(body.Root));
        Assert.DoesNotContain(body.Root.Descendants<Run>(), run => run.Elements<Text>().All(t => t.Text.Length == 0));
        Assert.Empty(body.Root.Descendants<DeletedRun>());
        Assert.Empty(body.Root.Descendants<InsertedRun>());
    }

    [Fact]
    public void A_tracked_replacement_still_writes_one_deletion_and_one_insertion()
    {
        var result = Apply(Contract(), new ChangeTextOp
        {
            Target = Span(Contract(), Deleted),
            With = " unless renewed",
            Mode = ChangeMode.Tracked
        });

        AssertValid(result);
        using var body = Body(result);
        Assert.Single(body.Root.Descendants<DeletedRun>());
        var inserted = Assert.Single(body.Root.Descendants<InsertedRun>());
        Assert.Equal(" unless renewed", inserted.InnerText);
    }

    private static byte[] Delete(byte[] document, ChangeMode mode) =>
        Apply(document, new ChangeTextOp { Target = Span(document, Deleted), With = string.Empty, Mode = mode });

    private static TextSpanAnchor Span(byte[] document, string text) =>
        (TextSpanAnchor)Client().Find(new StreamHandle(new MemoryStream(document)), new FindQuery(text)).First().Anchor;

    /// <summary>
    /// One paragraph whose deleted span crosses a run boundary, so the deletion has to
    /// isolate and remove more than one run. The comment, when present, brackets the whole
    /// sentence.
    /// </summary>
    private static byte[] Contract(bool withComment = false)
    {
        var buffer = new MemoryStream();
        using (var package = WordprocessingDocument.Create(buffer, WordprocessingDocumentType.Document))
        {
            var main = package.AddMainDocumentPart();
            Run Plain(string text) => new(new Text(text) { Space = SpaceProcessingModeValues.Preserve });
            var runs = new OpenXmlElement[]
            {
                Plain(Before + " unless terminated"),
                Plain(" earlier under clause 6."),
            };

            var paragraph = new Paragraph();
            if (withComment)
            {
                paragraph.AppendChild(new CommentRangeStart { Id = "0" });
                paragraph.Append(runs);
                paragraph.AppendChild(new CommentRangeEnd { Id = "0" });
                paragraph.AppendChild(new Run(new CommentReference { Id = "0" }));

                var comments = main.AddNewPart<WordprocessingCommentsPart>();
                comments.Comments = new Comments(new Comment(
                    new Paragraph(new Run(new Text("Check the term."))))
                { Id = "0", Author = "Reviewer", Initials = "R" });
            }
            else
            {
                paragraph.Append(runs);
            }

            main.Document = new Document(new Body(paragraph));
        }
        return buffer.ToArray();
    }

    private static byte[] Apply(byte[] document, PlanOperation operation)
    {
        using var applied = Client().Commit(
            new StreamHandle(new MemoryStream(document)),
            new DocumentPlan { Revision = new RevisionMetadata { Author = "Legal Ops" }, Operations = new[] { operation } });

        Assert.True(applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return applied.ToBytes();
    }

    /// <summary>The text a reader sees: run text, without struck-through deletions.</summary>
    private static string VisibleText(OpenXmlElement root) =>
        string.Concat(root.Descendants<Text>().Select(t => t.Text));

    private sealed class OpenBody : IDisposable
    {
        private readonly MemoryStream _stream;
        private readonly WordprocessingDocument _document;

        public Body Root { get; }

        public OpenBody(byte[] bytes)
        {
            _stream = new MemoryStream(bytes);
            _document = WordprocessingDocument.Open(_stream, isEditable: false);
            Root = _document.MainDocumentPart!.Document.Body!;
        }

        public void Dispose()
        {
            _document.Dispose();
            _stream.Dispose();
        }
    }

    private static OpenBody Body(byte[] document) => new(document);

    private static void AssertValid(byte[] document)
    {
        using var stream = new MemoryStream(document);
        using var opened = WordprocessingDocument.Open(stream, isEditable: false);
        var problems = new OpenXmlValidator(FileFormatVersions.Office2019)
            .Validate(opened)
            .Select(e => $"{e.Path?.XPath}: {e.Description}")
            .ToList();
        Assert.True(problems.Count == 0, string.Join("; ", problems.Take(3)));
    }
}
