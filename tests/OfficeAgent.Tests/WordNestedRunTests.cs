using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;
using V = DocumentFormat.OpenXml.Vml;

namespace OfficeAgent.Tests;

/// <summary>
/// A run is not always a direct child of <c>w:p</c>. Tracked insertions, hyperlinks and
/// content controls wrap theirs, and the text inside them is text the reader sees - so it
/// has to be findable and editable. These tests pin the second pass over a redlined
/// document: change something with tracking on, then come back and work on what you wrote.
/// </summary>
public class WordNestedRunTests
{
    private static OfficeAgentClient Client() => new(new WordModule());

    [Fact]
    public void Text_a_tracked_edit_inserted_is_part_of_the_paragraph_afterwards()
    {
        var client = Client();
        var redlined = TrackedRename(client);

        var paragraph = client.Inspect(redlined).Paragraphs
            .Single(p => p.Text.Contains("shall provide services"));

        // Before the fix this read "  shall provide services to Acme Corp." - the new
        // supplier name was in the file but invisible to every text-based verb.
        Assert.Contains("Globex Industries", paragraph.Text);
        // The struck-through original must not come back as live text.
        Assert.DoesNotContain("Acme Corp shall", paragraph.Text);
    }

    [Fact]
    public void A_second_plan_can_target_the_text_the_first_one_inserted()
    {
        var client = Client();
        var redlined = TrackedRename(client);

        // The workflow: counsel renames the party, then bolds the new name.
        var hit = Assert.Single(client.Find(
            new StreamHandle(new MemoryStream(redlined)),
            new FindQuery { Pattern = "Globex Industries" }));

        using var formatted = client.Commit(
            new StreamHandle(new MemoryStream(redlined)),
            new DocumentPlan
            {
                Operations = new PlanOperation[] { new FormatOp { Target = hit.Anchor!, Bold = true } }
            });

        Assert.True(formatted.Committed,
            string.Join("; ", formatted.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));

        using var stream = new MemoryStream(formatted.ToBytes());
        using var document = WordprocessingDocument.Open(stream, isEditable: false);
        var bold = document.MainDocumentPart!.Document.Descendants<Run>()
            .Where(r => r.RunProperties?.Bold is not null)
            .Select(r => string.Concat(r.Elements<Text>().Select(t => t.Text)));

        Assert.Equal("Globex Industries", string.Concat(bold));
    }

    [Fact]
    public void Content_control_text_is_findable_like_any_other_text()
    {
        var client = Client();

        // The placeholder lives inside w:sdt, two levels below the paragraph.
        var hit = Assert.Single(client.Find(
            new StreamHandle(new MemoryStream(DocxFactory.Contract())),
            new FindQuery { Pattern = DocxFactory.ClientPlaceholder }));

        Assert.Equal("Client: PLACEHOLDER", ParagraphText(client, "Client: "));
        Assert.NotNull(hit.Anchor);
    }

    [Fact]
    public void A_text_box_does_not_fold_into_the_paragraph_that_carries_it()
    {
        var client = Client();

        // A text box hangs off a run and owns its paragraphs. Walking blindly into every
        // descendant run would splice its words into the body text and corrupt offsets.
        var text = ParagraphText(client, "Body text", WithTextBox());

        Assert.Equal("Body text before the box.", text);
    }

    [Fact]
    public void Rows_cloned_from_a_template_row_get_their_own_paragraph_ids()
    {
        var client = Client();

        using var applied = client.Commit(
            new StreamHandle(new MemoryStream(WithTable())),
            new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new InsertTableRowsOp
                    {
                        Target = new NodeAnchor { Kind = "table", Path = "table#0" },
                        Position = TablePosition.End,
                        Rows = new[] { new[] { "UK", "68" } }
                    }
                }
            });

        Assert.True(applied.Committed);

        using var stream = new MemoryStream(applied.ToBytes());
        using var document = WordprocessingDocument.Open(stream, isEditable: false);
        var ids = document.MainDocumentPart!.Document.Descendants<Paragraph>()
            .Select(p => p.ParagraphId?.Value)
            .Where(id => id is not null)
            .ToList();

        // The row is cloned from the last one to keep its widths and formatting. Carrying
        // the template's ids across would put two paragraphs under one anchor, and every
        // later operation targeting that id would hit whichever came first.
        Assert.Equal(ids.Count, ids.Distinct().Count());
        Assert.NotEmpty(ids);
    }

    // ── helpers ───────────────────────────────────────────────────────────────

    private static byte[] TrackedRename(OfficeAgentClient client)
    {
        using var applied = client.Commit(
            new StreamHandle(new MemoryStream(DocxFactory.Contract())),
            new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new ChangeTextOp
                    {
                        Target = new TextSpanAnchor { ParaId = "w14:00000002", Expect = "Acme Corp" },
                        With = "Globex Industries",
                        Mode = ChangeMode.Tracked
                    }
                }
            });

        Assert.True(applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return applied.ToBytes();
    }

    private static string ParagraphText(OfficeAgentClient client, string startsWith, byte[]? document = null) =>
        client.Inspect(document ?? DocxFactory.Contract()).Paragraphs
            .Single(p => p.Text.StartsWith(startsWith, StringComparison.Ordinal))
            .Text;

    private static byte[] WithTextBox()
    {
        using var buffer = new MemoryStream();
        using (var document = WordprocessingDocument.Create(buffer, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();

            var boxed = new Paragraph(new Run(new Text("Caption inside the box")))
            {
                ParagraphId = "0000BB01"
            };
            var picture = new Picture(
                new V.Shape(
                    new V.TextBox(new TextBoxContent(boxed)))
                {
                    Id = "shape1",
                    Style = "width:100pt;height:40pt"
                });

            var paragraph = new Paragraph(
                new Run(new Text("Body text before the box.") { Space = SpaceProcessingModeValues.Preserve }),
                new Run(picture))
            {
                ParagraphId = "0000AA01"
            };

            main.Document = new Document(new Body(paragraph));
            main.Document.Save();
        }
        return buffer.ToArray();
    }

    private static byte[] WithTable()
    {
        using var buffer = new MemoryStream();
        using (var document = WordprocessingDocument.Create(buffer, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(new Table(
                new TableProperties(new TableStyle { Val = "TableGrid" }),
                Row("Country", "Population mil"),
                Row("RU", "146"),
                Row("US", "332"))));
            main.Document.Save();
        }
        return buffer.ToArray();
    }

    private static TableRow Row(params string[] cells)
    {
        var row = new TableRow();
        foreach (var cell in cells)
            row.AppendChild(new TableCell(new Paragraph(
                new Run(new Text(cell) { Space = SpaceProcessingModeValues.Preserve }))));
        return row;
    }
}
