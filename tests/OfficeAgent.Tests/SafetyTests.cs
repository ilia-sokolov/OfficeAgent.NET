using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

public class SafetyTests
{
    private static OfficeAgentClient Office() => new(new WordModule());
    private static StreamHandle Handle(byte[] bytes) => new(new MemoryStream(bytes), "contract.docx");

    [Fact]
    public void Stale_snapshot_in_plan_is_rejected()
    {
        var office = Office();
        var bytes = DocxFactory.Contract();
        var inspect = office.Inspect(Handle(bytes));
        var clauseParaId = inspect.Paragraphs.First(p => p.Text.Contains("shall provide")).ParaId;

        var plan = new DocumentPlan
        {
            Snapshot = new SnapshotToken("deadbeef-not-the-real-etag"),
            Operations = new PlanOperation[]
            {
                new ChangeTextOp
                {
                    Target = new TextSpanAnchor { ParaId = clauseParaId, Expect = "Acme Corp", Occurrence = 0 },
                    With = "Globex",
                    Mode = ChangeMode.Direct
                }
            }
        };

        var report = office.Preview(Handle(bytes), plan);

        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, e => e.Code == ValidationErrorCodes.StaleSnapshot);
    }

    [Fact]
    public void Matching_snapshot_passes_validation()
    {
        var office = Office();
        var bytes = DocxFactory.Contract();
        var inspect = office.Inspect(Handle(bytes));
        var clauseParaId = inspect.Paragraphs.First(p => p.Text.Contains("shall provide")).ParaId;

        var plan = new DocumentPlan
        {
            Snapshot = inspect.Snapshot,
            Operations = new PlanOperation[]
            {
                new ChangeTextOp
                {
                    Target = new TextSpanAnchor { ParaId = clauseParaId, Expect = "Acme Corp", Occurrence = 0 },
                    With = "Globex",
                    Mode = ChangeMode.Direct
                }
            }
        };

        var report = office.Preview(Handle(bytes), plan);
        Assert.True(report.IsValid);
    }

    [Fact]
    public void Insert_then_edit_a_later_paragraph_stays_anchored_to_the_right_paragraph()
    {
        // Document without w14:paraId - the engine must assign one and rewrite anchors
        // so that the InsertOp's offset shift doesn't redirect the ChangeTextOp.
        var bytes = BuildPositionalDoc();
        var office = Office();
        var inspect = office.Inspect(Handle(bytes));

        var firstId = inspect.Paragraphs[0].ParaId;
        var thirdId = inspect.Paragraphs[2].ParaId;

        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertOp
                {
                    Target = new TextSpanAnchor { ParaId = firstId, Expect = "Alpha" },
                    Position = InsertPosition.After,
                    Text = "Inserted line"
                },
                new ChangeTextOp
                {
                    Target = new TextSpanAnchor { ParaId = thirdId, Expect = "Charlie", Occurrence = 0 },
                    With = "Charlie!",
                    Mode = ChangeMode.Direct
                }
            }
        };

        var result = office.Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var texts = doc.MainDocumentPart!.Document.Body!.Descendants<Paragraph>()
            .Select(p => string.Concat(p.Descendants<Text>().Select(t => t.Text)))
            .ToList();

        Assert.Equal(new[] { "Alpha", "Inserted line", "Bravo", "Charlie!" }, texts);
    }

    [Fact]
    public void TrackedChange_revision_ids_do_not_collide_with_existing_ones()
    {
        var bytes = BuildDocWithExistingRevision(maxExistingId: 500);
        var office = Office();
        var inspect = office.Inspect(Handle(bytes));
        var paraId = inspect.Paragraphs.First(p => p.Text.Contains("Acme")).ParaId;

        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new ChangeTextOp
                {
                    Target = new TextSpanAnchor { ParaId = paraId, Expect = "Acme", Occurrence = 0 },
                    With = "Globex",
                    Mode = ChangeMode.Tracked
                }
            }
        };

        var result = office.Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var body = doc.MainDocumentPart!.Document.Body!;
        var allIds = body.Descendants<InsertedRun>().Select(i => int.Parse(i.Id!.Value!))
            .Concat(body.Descendants<DeletedRun>().Select(d => int.Parse(d.Id!.Value!)))
            .ToList();

        var newIds = allIds.Where(id => id != 500).ToList();
        Assert.Equal(2, newIds.Count);
        Assert.True(newIds.All(id => id > 500), $"New revision ids should be > 500, got [{string.Join(",", newIds)}]");
        Assert.Equal(allIds.Count, allIds.Distinct().Count());
    }

    [Fact]
    public void Tracked_edits_use_injected_clock_for_deterministic_timestamps()
    {
        var fixedInstant = new DateTimeOffset(2026, 1, 2, 3, 4, 5, TimeSpan.Zero);
        var clock = new FixedTimeProvider(fixedInstant);
        var client = new OfficeAgentClient(new WordModule(clock));

        var bytes = DocxFactory.Contract();
        var inspect = client.Inspect(Handle(bytes));
        var paraId = inspect.Paragraphs.First(p => p.Text.Contains("shall provide")).ParaId;

        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new ChangeTextOp
                {
                    Target = new TextSpanAnchor { ParaId = paraId, Expect = "Acme Corp", Occurrence = 0 },
                    With = "Globex",
                    Mode = ChangeMode.Tracked
                }
            }
        };

        var result = client.Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var ins = doc.MainDocumentPart!.Document.Body!.Descendants<InsertedRun>().Single();
        Assert.Equal(fixedInstant.UtcDateTime, ins.Date!.Value);
    }

    private static byte[] BuildPositionalDoc()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Alpha"))),
                new Paragraph(new Run(new Text("Bravo"))),
                new Paragraph(new Run(new Text("Charlie")))));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    private static byte[] BuildDocWithExistingRevision(int maxExistingId)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var paragraph = new Paragraph { ParagraphId = "ABCDEF01" };
            paragraph.AppendChild(new Run(new Text("Acme") { Space = SpaceProcessingModeValues.Preserve }));

            var existingIns = new InsertedRun { Id = maxExistingId.ToString(), Author = "Other", Date = DateTime.UtcNow };
            existingIns.AppendChild(new Run(new Text(" (added)") { Space = SpaceProcessingModeValues.Preserve }));
            paragraph.AppendChild(existingIns);

            main.Document = new Document(new Body(paragraph));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    [Fact]
    public void DI_resolves_client_with_registered_word_module()
    {
        var services = new ServiceCollection();
        services.AddSingleton<IFormatModule, WordModule>();
        services.AddOfficeAgent();

        using var provider = services.BuildServiceProvider();
        var client = provider.GetRequiredService<OfficeAgentClient>();
        var inspect = client.Inspect(Handle(DocxFactory.Contract()));

        Assert.Equal(OfficeAgent.Abstractions.DocumentFormat.Word, inspect.Format);
    }

    private sealed class FixedTimeProvider : TimeProvider
    {
        private readonly DateTimeOffset _now;
        public FixedTimeProvider(DateTimeOffset now) => _now = now;
        public override DateTimeOffset GetUtcNow() => _now;
    }
}
