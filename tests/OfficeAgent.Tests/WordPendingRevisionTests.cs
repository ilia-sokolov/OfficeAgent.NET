using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// The contract for editing a document that already carries pending revisions: what the
/// text view shows, which successive edits are allowed, and which are refused before any
/// mutation. Accept and reject mechanics per revision kind live in
/// <see cref="WordRedlineTests"/>; these cases are about a second author arriving at a
/// document a first author has already redlined.
/// </summary>
public sealed class WordPendingRevisionTests
{
    private static readonly DateTimeOffset Stamp = DateTimeOffset.Parse("2026-01-01T00:00:00Z");
    private static readonly DateTimeOffset LaterStamp = DateTimeOffset.Parse("2026-02-02T00:00:00Z");

    // ── The text view ────────────────────────────────────────────────────

    /// <summary>
    /// Inspect and find show the document as it would read if every pending change were
    /// accepted: insertions are present, deletions are not.
    /// </summary>
    [Fact]
    public void The_text_view_shows_insertions_and_hides_deletions()
    {
        var client = Client();
        var original = Docx("Alpha beta gamma delta.");
        var edited = Tracked(client, original, "beta ", "", "Author A");

        var paragraph = Assert.Single(client.Inspect(edited).Paragraphs);
        Assert.Equal("Alpha gamma delta.", paragraph.Text);

        // The deleted words are still in the package, just not in the view.
        Assert.Contains("beta", DeletedText(edited), StringComparison.Ordinal);

        // Find agrees with inspect: it cannot match text that is pending deletion.
        Assert.Empty(client.Find(Handle(edited), new FindQuery("beta")));
        Assert.Single(client.Find(Handle(edited), new FindQuery("Alpha gamma")));
    }

    /// <summary>Every pending revision is discoverable with the author who left it.</summary>
    [Fact]
    public void Pending_revisions_are_discoverable_with_their_author_and_date()
    {
        var client = Client();
        var edited = Tracked(client, Docx("The agreement stands."), "agreement", "contract", "Dana Reyes");

        var revisions = client.Inspect(edited).Nodes.Where(node => node.Kind == "revision").ToList();

        Assert.Equal(2, revisions.Count);
        Assert.Contains(revisions, node => node.Path.StartsWith("del#", StringComparison.Ordinal));
        Assert.Contains(revisions, node => node.Path.StartsWith("ins#", StringComparison.Ordinal));
        Assert.All(revisions, node => Assert.Contains("Dana Reyes", node.Summary, StringComparison.Ordinal));
        Assert.All(revisions, node => Assert.Contains("2026-01-01", node.Summary, StringComparison.Ordinal));
    }

    // ── Allowed successive edits ─────────────────────────────────────────

    /// <summary>
    /// An edit that touches no pending revision is ordinary, and it leaves the earlier
    /// author's revision identity untouched.
    /// </summary>
    [Fact]
    public void An_edit_clear_of_pending_revisions_preserves_the_earlier_identity()
    {
        var client = Client();
        var afterA = Tracked(client, Docx("Clause one.", "Clause two."), "one", "ONE", "Author A");
        var afterB = Tracked(client, afterA, "two", "TWO", "Author B", LaterStamp);

        var revisions = client.Inspect(afterB).Nodes.Where(node => node.Kind == "revision").ToList();

        Assert.Equal(2, revisions.Count(node => node.Summary.Contains("Author A", StringComparison.Ordinal)));
        Assert.Equal(2, revisions.Count(node => node.Summary.Contains("Author B", StringComparison.Ordinal)));
        Assert.Contains(revisions, node =>
            node.Summary.Contains("Author A", StringComparison.Ordinal) &&
            node.Summary.Contains("2026-01-01", StringComparison.Ordinal));
        Assert.Contains(revisions, node =>
            node.Summary.Contains("Author B", StringComparison.Ordinal) &&
            node.Summary.Contains("2026-02-02", StringComparison.Ordinal));
        AssertSchemaValid(afterB);
    }

    /// <summary>
    /// Editing text that lives wholly inside another author's pending insertion is the
    /// nested w:ins/w:del case Word itself writes, and it still resolves both ways.
    /// </summary>
    [Fact]
    public void An_edit_inside_another_authors_pending_insertion_still_resolves_both_ways()
    {
        var client = Client();
        var original = Docx("The agreement stands.");
        var afterA = Tracked(client, original, "The agreement stands.", "The revised agreement stands.", "Author A");
        var afterB = Tracked(client, afterA, "revised", "amended", "Author B", LaterStamp);

        Assert.Equal("The agreement stands.", VisibleText(Resolve(client, afterB, "all", RevisionAction.Reject)));
        Assert.Equal("The amended agreement stands.", VisibleText(Resolve(client, afterB, "all", RevisionAction.Accept)));
        AssertSchemaValid(afterB);
    }

    // ── Refused successive edits ─────────────────────────────────────────

    /// <summary>
    /// The case that made this contract necessary. The text view hides a pending
    /// deletion, so two runs read as adjacent while another author's deleted words sit
    /// physically between them. Redlining across that gap produces a document whose
    /// reject-all no longer restores the original, so it is refused before any mutation.
    /// </summary>
    [Fact]
    public void An_edit_spanning_a_pending_deletion_is_refused_before_it_mutates()
    {
        var client = Client();
        var original = Docx("Alpha beta gamma delta.");
        var afterA = Tracked(client, original, "beta ", "", "Author A");
        var before = Sha256(afterA);

        var hit = Assert.Single(client.Find(Handle(afterA), new FindQuery("Alpha gamma")));
        var plan = Plan("Author B", LaterStamp, new ChangeTextOp
        {
            Target = hit.Anchor,
            With = "Alpha GAMMA",
            Mode = ChangeMode.Tracked
        });

        var report = client.Preview(Handle(afterA), plan);
        Assert.False(report.IsValid);
        var error = Assert.Single(report.Errors);
        Assert.Equal(ValidationErrorCodes.RevisionOverlap, error.Code);
        Assert.Equal(ValidationErrorCode.RevisionOverlap, ValidationErrorCodes.Parse(error.Code));

        // The recovery instruction names the way out rather than only the problem.
        Assert.Contains("accept or reject", error.Message, StringComparison.Ordinal);

        using var commit = client.Commit(Handle(afterA), plan);
        Assert.False(commit.Committed);
        Assert.Equal(before, Sha256(afterA));
    }

    /// <summary>
    /// The same refusal in direct mode. A direct edit across the gap would strand the
    /// earlier author's deletion mid-paragraph, which is resolving their revision by
    /// implication, and the contract never does that silently.
    /// </summary>
    [Fact]
    public void A_direct_edit_spanning_a_pending_deletion_is_refused_too()
    {
        var client = Client();
        var afterA = Tracked(client, Docx("Alpha beta gamma delta."), "beta ", "", "Author A");

        var hit = Assert.Single(client.Find(Handle(afterA), new FindQuery("Alpha gamma")));
        var report = client.Preview(Handle(afterA), Plan("Author B", LaterStamp, new ChangeTextOp
        {
            Target = hit.Anchor,
            With = "Alpha GAMMA",
            Mode = ChangeMode.Direct
        }));

        Assert.False(report.IsValid);
        Assert.Equal(ValidationErrorCodes.RevisionOverlap, Assert.Single(report.Errors).Code);
    }

    /// <summary>
    /// Reaching out of a pending insertion into unmarked text is refused for the same
    /// reason: no single insertion point represents the whole edit.
    /// </summary>
    [Fact]
    public void An_edit_reaching_out_of_a_pending_insertion_is_refused()
    {
        var client = Client();
        var afterA = Tracked(client, Docx("Alpha beta."), "Alpha", "Gamma", "Author A");

        // "Gamma beta." spans A's inserted run and the untouched run after it.
        var hit = Assert.Single(client.Find(Handle(afterA), new FindQuery("Gamma beta")));
        var report = client.Preview(Handle(afterA), Plan("Author B", LaterStamp, new ChangeTextOp
        {
            Target = hit.Anchor,
            With = "Delta epsilon",
            Mode = ChangeMode.Tracked
        }));

        Assert.False(report.IsValid);
        var error = Assert.Single(report.Errors);
        Assert.Equal(ValidationErrorCodes.RevisionOverlap, error.Code);
        Assert.Contains("container", error.Message, StringComparison.Ordinal);
    }

    /// <summary>
    /// The refusal is specific, not a blanket ban on editing a redlined document: the
    /// same document still takes an edit that stays clear of the pending revision.
    /// </summary>
    [Fact]
    public void A_refused_document_still_accepts_an_edit_clear_of_the_revision()
    {
        var client = Client();
        var afterA = Tracked(client, Docx("Alpha beta gamma delta."), "beta ", "", "Author A");

        var afterB = Tracked(client, afterA, "delta", "epsilon", "Author B", LaterStamp);

        Assert.Equal("Alpha gamma epsilon.", VisibleTextOfView(client, afterB));
        Assert.Equal("Alpha beta gamma delta.", VisibleText(Resolve(client, afterB, "all", RevisionAction.Reject)));
        Assert.Equal("Alpha gamma epsilon.", VisibleText(Resolve(client, afterB, "all", RevisionAction.Accept)));
    }

    // ── Resolution across two authors ────────────────────────────────────

    /// <summary>
    /// Resolving one author leaves the other author's revisions pending, in either order,
    /// and the final document is the same either way.
    /// </summary>
    [Fact]
    public void Resolving_one_author_leaves_the_other_pending_in_either_order()
    {
        var client = Client();
        var both = Tracked(
            client,
            Tracked(client, Docx("Clause one.", "Clause two."), "one", "ONE", "Author A"),
            "two", "TWO", "Author B", LaterStamp);

        var aFirst = Resolve(client, Resolve(client, both, "author:Author A", RevisionAction.Accept),
            "author:Author B", RevisionAction.Accept);
        var bFirst = Resolve(client, Resolve(client, both, "author:Author B", RevisionAction.Accept),
            "author:Author A", RevisionAction.Accept);

        Assert.Equal(VisibleText(aFirst), VisibleText(bFirst));
        Assert.Equal("Clause ONE.Clause TWO.", VisibleText(aFirst));

        // Half-resolved: A accepted, B still pending and still attributed to B.
        var halfway = Resolve(client, both, "author:Author A", RevisionAction.Accept);
        var pending = client.Inspect(halfway).Nodes.Where(node => node.Kind == "revision").ToList();
        Assert.NotEmpty(pending);
        Assert.All(pending, node => Assert.Contains("Author B", node.Summary, StringComparison.Ordinal));
    }

    /// <summary>
    /// Rejecting one author restores only that author's original text and leaves the
    /// other author's proposal standing.
    /// </summary>
    [Fact]
    public void Rejecting_one_author_restores_only_that_authors_text()
    {
        var client = Client();
        var both = Tracked(
            client,
            Tracked(client, Docx("Clause one.", "Clause two."), "one", "ONE", "Author A"),
            "two", "TWO", "Author B", LaterStamp);

        var rejectedA = Resolve(client, both, "author:Author A", RevisionAction.Reject);

        Assert.Equal("Clause one.", VisibleTextOfParagraph(client, rejectedA, 0));
        Assert.Equal("Clause TWO.", VisibleTextOfParagraph(client, rejectedA, 1));
        Assert.All(
            client.Inspect(rejectedA).Nodes.Where(node => node.Kind == "revision"),
            node => Assert.Contains("Author B", node.Summary, StringComparison.Ordinal));
    }

    /// <summary>
    /// Accepting or rejecting revisions is not allowed to disturb comments that have
    /// nothing to do with them.
    /// </summary>
    [Fact]
    public void Resolving_revisions_retains_unrelated_comments()
    {
        var client = Client();
        var withComment = Commented(client, Docx("Clause one.", "Clause two."), "Clause two.", "Please check this.");
        var edited = Tracked(client, withComment, "one", "ONE", "Author A");

        var accepted = Resolve(client, edited, "all", RevisionAction.Accept);
        var rejected = Resolve(client, edited, "all", RevisionAction.Reject);

        foreach (var document in new[] { accepted, rejected })
        {
            Assert.Contains("Please check this.", AllCommentText(document), StringComparison.Ordinal);
            AssertSchemaValid(document);
        }
    }


    /// <summary>
    /// Formatting text that already carries a pending formatting revision keeps the
    /// first revision rather than stacking a second one. WordprocessingML allows one
    /// <c>w:rPrChange</c> per run, and the one that matters is the oldest: it is what
    /// reject must restore. The later author's formatting is applied but is not
    /// separately attributed, which is a documented limit of formatting history.
    /// </summary>
    [Fact]
    public void Formatting_over_a_pending_format_revision_keeps_the_original_restore_point()
    {
        var client = Client();
        var original = Docx("Payment terms are net thirty days.");

        var afterA = Formatted(client, original, "net thirty days", "Author A", Stamp, bold: true);
        var afterB = Formatted(client, afterA, "net thirty days", "Author B", LaterStamp, italic: true);

        // One formatting revision, still the first author's: that is the restore point.
        var formatRevisions = client.Inspect(afterB).Nodes
            .Where(node => node.Kind == "revision" && node.Path.StartsWith("runFormat#", StringComparison.Ordinal))
            .ToList();
        var only = Assert.Single(formatRevisions);
        Assert.Contains("Author A", only.Summary, StringComparison.Ordinal);
        AssertSchemaValid(afterB);

        // Both authors' formatting is present in the document.
        using (var document = WordprocessingDocument.Open(new MemoryStream(afterB, writable: false), isEditable: false))
        {
            var run = document.MainDocumentPart!.Document!.Body!
                .Descendants<Run>().Single(r => r.InnerText == "net thirty days");
            Assert.NotNull(run.RunProperties?.GetFirstChild<Bold>());
            Assert.NotNull(run.RunProperties?.GetFirstChild<Italic>());
        }

        // Rejecting restores the genuine original, not an intermediate state that never
        // existed: the run had neither bold nor italic before either author touched it.
        var rejected = Resolve(client, afterB, "all", RevisionAction.Reject);
        using (var document = WordprocessingDocument.Open(new MemoryStream(rejected, writable: false), isEditable: false))
        {
            var run = document.MainDocumentPart!.Document!.Body!
                .Descendants<Run>().Single(r => r.InnerText == "net thirty days");
            Assert.Null(run.RunProperties?.GetFirstChild<Bold>());
            Assert.Null(run.RunProperties?.GetFirstChild<Italic>());
        }
    }

    /// <summary>
    /// A structural row verb applied to a table that already has a pending row revision
    /// marks its own row and leaves the earlier author's row marker intact.
    /// </summary>
    [Fact]
    public void A_row_verb_over_a_pending_row_revision_keeps_both_markers()
    {
        var client = Client();
        var withTable = Committed(client, Docx("Regional results."), new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertTableOp
                {
                    Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = "Regional results." },
                    Position = InsertPosition.After,
                    Table = new TableData { Headers = new[] { "Region", "Q1" }, Rows = new[] { new[] { "NL", "41850" } } },
                    Mode = ChangeMode.Direct
                }
            }
        });

        var afterA = Committed(client, withTable, Plan("Author A", Stamp, new InsertTableRowsOp
        {
            Target = new NodeAnchor { Kind = "table", Path = "table#0" },
            Rows = new[] { new[] { "BE", "12300" } },
            Position = TablePosition.End
        }));

        var afterB = Committed(client, afterA, Plan("Author B", LaterStamp, new InsertTableRowsOp
        {
            Target = new NodeAnchor { Kind = "table", Path = "table#0" },
            Rows = new[] { new[] { "DE", "58200" } },
            Position = TablePosition.End
        }));

        var rowRevisions = client.Inspect(afterB).Nodes
            .Where(node => node.Kind == "revision" && node.Path.StartsWith("rowIns#", StringComparison.Ordinal))
            .ToList();

        Assert.Equal(2, rowRevisions.Count);
        Assert.Contains(rowRevisions, node => node.Summary.Contains("Author A", StringComparison.Ordinal));
        Assert.Contains(rowRevisions, node => node.Summary.Contains("Author B", StringComparison.Ordinal));
        AssertSchemaValid(afterB);

        // Rejecting only B leaves A's proposed row standing.
        var rejectedB = Resolve(client, afterB, "author:Author B", RevisionAction.Reject);
        using var document = WordprocessingDocument.Open(new MemoryStream(rejectedB, writable: false), isEditable: false);
        var text = document.MainDocumentPart!.Document!.Body!.InnerText;
        Assert.Contains("BE", text, StringComparison.Ordinal);
        Assert.DoesNotContain("DE", text, StringComparison.Ordinal);
    }

    // ── Stale anchors ────────────────────────────────────────────────────

    /// <summary>
    /// An anchor captured before an intervening edit is refused rather than applied to
    /// whatever now sits at that address.
    /// </summary>
    [Fact]
    public void A_stale_anchor_is_refused_after_an_intervening_edit()
    {
        var client = Client();
        var original = Docx("Clause one.", "Clause two.");
        var inspection = client.Inspect(original);
        var hit = Assert.Single(client.Find(Handle(original), new FindQuery("Clause two.")));

        // Another author rewrites the first paragraph, which renumbers the automatic ids.
        var afterA = Tracked(client, original, "Clause one.", "Clause one revised.", "Author A");

        var withSnapshot = new DocumentPlan
        {
            Snapshot = inspection.Snapshot,
            Operations = new PlanOperation[]
            {
                new ChangeTextOp { Target = hit.Anchor, With = "Clause two revised.", Mode = ChangeMode.Tracked }
            }
        };

        var report = client.Preview(Handle(afterA), withSnapshot);
        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, error => error.Code == ValidationErrorCodes.StaleSnapshot);

        // Dropping the snapshot does not make the stale anchor safe: it is still refused,
        // now by the anchor itself.
        var withoutSnapshot = new DocumentPlan { Operations = withSnapshot.Operations };
        var second = client.Preview(Handle(afterA), withoutSnapshot);
        Assert.False(second.IsValid);
        Assert.Contains(second.Errors, error =>
            error.Code == ValidationErrorCodes.AnchorNotFound ||
            error.Code == ValidationErrorCodes.ExpectMismatch);
    }

    // ── Helpers ──────────────────────────────────────────────────────────

    private static OfficeAgentClient Client() => new(new WordModule());

    private static StreamHandle Handle(byte[] bytes) => new(new MemoryStream(bytes, writable: false));

    private static DocumentPlan Plan(string author, DateTimeOffset stamp, params PlanOperation[] operations) => new()
    {
        Revision = new RevisionMetadata { Author = author, TimestampUtc = stamp },
        Operations = operations
    };

    private static byte[] Tracked(
        OfficeAgentClient client, byte[] input, string find, string with, string author, DateTimeOffset? stamp = null)
    {
        var hit = client.Find(Handle(input), new FindQuery(find)).First();
        using var result = client.Commit(Handle(input), Plan(author, stamp ?? Stamp, new ChangeTextOp
        {
            Target = hit.Anchor,
            With = with,
            Mode = ChangeMode.Tracked
        }));
        Assert.True(result.Committed, Errors(result.Report));
        return result.ToBytes();
    }

    private static byte[] Commented(OfficeAgentClient client, byte[] input, string find, string text)
    {
        var hit = client.Find(Handle(input), new FindQuery(find)).First();
        using var result = client.Commit(Handle(input), new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new CommentOp { Target = hit.Anchor, Text = text, Author = "Reviewer" }
            }
        });
        Assert.True(result.Committed, Errors(result.Report));
        return result.ToBytes();
    }

    private static byte[] Resolve(OfficeAgentClient client, byte[] input, string path, RevisionAction action)
    {
        using var result = client.Commit(Handle(input), new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new RevisionOp { Target = new NodeAnchor { Kind = "revision", Path = path }, Action = action }
            }
        });
        Assert.True(result.Committed, Errors(result.Report));
        return result.ToBytes();
    }

    private static byte[] Committed(OfficeAgentClient client, byte[] input, DocumentPlan plan)
    {
        using var result = client.Commit(Handle(input), plan);
        Assert.True(result.Committed, Errors(result.Report));
        return result.ToBytes();
    }

    private static byte[] Formatted(
        OfficeAgentClient client, byte[] input, string expect, string author, DateTimeOffset stamp,
        bool? bold = null, bool? italic = null)
    {
        var hit = client.Find(Handle(input), new FindQuery(expect)).First();
        return Committed(client, input, Plan(author, stamp, new FormatOp
        {
            Target = hit.Anchor,
            Bold = bold,
            Italic = italic,
            Mode = ChangeMode.Tracked
        }));
    }

    private static string Errors(ChangeReport report) =>
        string.Join("; ", report.Errors.Select(error => $"{error.Code}: {error.Message}"));

    private static string VisibleText(byte[] bytes)
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(bytes, writable: false), isEditable: false);
        return string.Concat(document.MainDocumentPart!.Document!.Body!.Descendants<Text>().Select(t => t.Text));
    }

    private static string DeletedText(byte[] bytes)
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(bytes, writable: false), isEditable: false);
        return string.Concat(document.MainDocumentPart!.Document!.Body!
            .Descendants<DeletedText>().Select(t => t.Text));
    }

    private static string AllCommentText(byte[] bytes)
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(bytes, writable: false), isEditable: false);
        var comments = document.MainDocumentPart?.WordprocessingCommentsPart?.Comments;
        return comments is null ? string.Empty : string.Concat(comments.Descendants<Text>().Select(t => t.Text));
    }

    private static string VisibleTextOfView(OfficeAgentClient client, byte[] bytes) =>
        string.Concat(client.Inspect(bytes).Paragraphs.Select(paragraph => paragraph.Text));

    private static string VisibleTextOfParagraph(OfficeAgentClient client, byte[] bytes, int index) =>
        client.Inspect(bytes).Paragraphs[index].Text;

    private static string Sha256(byte[] bytes) =>
        Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(bytes));

    private static void AssertSchemaValid(byte[] bytes)
    {
        using var document = WordprocessingDocument.Open(new MemoryStream(bytes, writable: false), isEditable: false);
        var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document).ToList();
        Assert.Empty(errors.Select(error => error.Description));
    }

    private static byte[] Docx(params string[] paragraphs)
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                paragraphs.Select(text => new Paragraph(new Run(new Text(text))))));
            main.Document.Save();
        }

        return ms.ToArray();
    }
}
