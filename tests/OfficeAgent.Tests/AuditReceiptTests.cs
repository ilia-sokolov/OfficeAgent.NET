using System.Security.Cryptography;
using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

public class AuditReceiptTests
{
    [Fact]
    public void Revision_identity_round_trips_with_the_plan_contract()
    {
        var timestamp = new DateTimeOffset(2026, 9, 9, 10, 11, 12, TimeSpan.Zero);
        var plan = new DocumentPlan
        {
            Revision = new RevisionMetadata { Author = "Legal Review", TimestampUtc = timestamp },
            Operations = Array.Empty<PlanOperation>()
        };

        var roundTrip = JsonSerializer.Deserialize<DocumentPlan>(JsonSerializer.Serialize(plan));

        Assert.NotNull(roundTrip?.Revision);
        Assert.Equal("Legal Review", roundTrip.Revision.Author);
        Assert.Equal(timestamp, roundTrip.Revision.TimestampUtc);
    }

    [Fact]
    public void Plan_revision_identity_is_shared_by_every_tracked_operation()
    {
        var engineInstant = new DateTimeOffset(2026, 9, 9, 8, 0, 0, TimeSpan.Zero);
        var revisionInstant = new DateTimeOffset(2026, 9, 9, 12, 34, 56, TimeSpan.FromHours(2));
        var client = new OfficeAgentClient(new WordModule(new FixedTimeProvider(engineInstant)));
        var input = DocxFactory.Contract();
        var handle = Handle(input);
        var paragraph = client.Inspect(handle).Paragraphs.First(p => p.Text.Contains("Acme Corp"));
        var target = new TextSpanAnchor
        {
            ParaId = paragraph.ParaId,
            Expect = "Acme Corp",
            Occurrence = 0
        };
        var actor = new AuditActor
        {
            Subject = "user-42",
            Issuer = "https://identity.example",
            DisplayName = "Authenticated User"
        };
        var plan = new DocumentPlan
        {
            Revision = new RevisionMetadata
            {
                Author = "Contract Review Agent",
                TimestampUtc = revisionInstant
            },
            Operations = new PlanOperation[]
            {
                new FormatOp { Target = target, Bold = true, Mode = ChangeMode.Tracked },
                new ChangeTextOp
                {
                    Target = target,
                    With = "Globex",
                    Mode = ChangeMode.Tracked
                }
            }
        };

        using var result = client.Apply(
            Handle(input), plan, new ApplyOptions { DryRun = false, Actor = actor });
        var receipt = Assert.IsType<ApplyReceipt>(result.Receipt);

        Assert.True(result.Committed);
        Assert.Equal(ApplyOutcome.Committed, receipt.Outcome);
        Assert.Equal(engineInstant, receipt.TimestampUtc);
        Assert.Equal("Contract Review Agent", receipt.Revision.Author);
        Assert.Equal(revisionInstant.ToUniversalTime(), receipt.Revision.TimestampUtc);
        Assert.Same(actor, receipt.Actor);
        Assert.Null(receipt.OutputDocument);
        Assert.Equal(Hash(input), receipt.InputSha256);
        Assert.Equal(Hash(JsonSerializer.SerializeToUtf8Bytes(plan)), receipt.PlanSha256);
        Assert.Equal(Hash(result.ToBytes()), receipt.OutputSha256);

        using var document = WordprocessingDocument.Open(new MemoryStream(result.ToBytes()), false);
        var markers = document.MainDocumentPart!.Document
            .Descendants()
            .Where(element => RevisionAttribute(element, "author") is not null)
            .ToList();

        Assert.True(markers.Count >= 3);
        Assert.All(markers, marker =>
        {
            Assert.Equal("Contract Review Agent", RevisionAttribute(marker, "author"));
            Assert.Equal("2026-09-09T10:34:56Z", RevisionAttribute(marker, "date"));
        });
    }

    [Fact]
    public void Default_revision_identity_uses_module_clock_once()
    {
        var instant = new DateTimeOffset(2026, 9, 9, 8, 1, 2, TimeSpan.Zero);
        var client = new OfficeAgentClient(new WordModule(new FixedTimeProvider(instant)));
        var input = DocxFactory.Contract();
        var paragraph = client.Inspect(Handle(input)).Paragraphs.First(p => p.Text.Contains("Acme Corp"));
        var plan = ChangePlan(paragraph.ParaId);

        using var result = client.Commit(Handle(input), plan);
        var receipt = Assert.IsType<ApplyReceipt>(result.Receipt);

        Assert.Equal("OfficeAgent", receipt.Revision.Author);
        Assert.Equal(instant, receipt.Revision.TimestampUtc);
        Assert.Equal(instant, receipt.TimestampUtc);
    }

    [Fact]
    public void Invalid_revision_author_rejects_plan_and_records_receipt()
    {
        var client = new OfficeAgentClient(new WordModule());
        var input = DocxFactory.Contract();
        var paragraph = client.Inspect(Handle(input)).Paragraphs.First(p => p.Text.Contains("Acme Corp"));
        var plan = ChangePlan(paragraph.ParaId, new RevisionMetadata { Author = "bad\nauthor" });

        using var result = client.Commit(Handle(input), plan);
        var receipt = Assert.IsType<ApplyReceipt>(result.Receipt);

        Assert.False(result.Committed);
        Assert.Equal(ApplyOutcome.Rejected, receipt.Outcome);
        Assert.Null(receipt.OutputSha256);
        Assert.Contains(result.Report.Errors, e => e.Code == ValidationErrorCodes.InvalidOperation);
    }

    [Fact]
    public void Preview_receipt_has_no_output_hash()
    {
        var client = new OfficeAgentClient(new WordModule());
        var input = DocxFactory.Contract();
        var paragraph = client.Inspect(Handle(input)).Paragraphs.First(p => p.Text.Contains("Acme Corp"));

        using var result = client.Apply(Handle(input), ChangePlan(paragraph.ParaId), ApplyOptions.Preview);
        var receipt = Assert.IsType<ApplyReceipt>(result.Receipt);

        Assert.False(result.Committed);
        Assert.Equal(ApplyOutcome.Previewed, receipt.Outcome);
        Assert.Null(receipt.OutputSha256);
        Assert.Equal(Hash(input), receipt.InputSha256);
    }

    [Fact]
    public async Task Provider_receipt_uses_host_actor_and_names_saved_document()
    {
        var root = Path.Combine(Path.GetTempPath(), $"officeagent-audit-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        try
        {
            var actor = new AuditActor { Subject = "oidc-subject", Issuer = "tenant-a" };
            var services = new ServiceCollection();
            services.AddWordFormat();
            services.AddFileSystemDocumentProvider("workspace", root);
            services.AddSingleton<IAuditActorProvider>(new FixedActorProvider(actor));
            services.AddOfficeAgent();
            using var provider = services.BuildServiceProvider();
            var client = provider.GetRequiredService<OfficeAgentClient>();
            var reference = await client.RegisterBytesAsync(
                "workspace", root, DocxFactory.Contract(), "contract.docx");
            var paragraph = (await client.InspectAsync(reference)).Paragraphs
                .First(p => p.Text.Contains("Acme Corp"));

            var result = await client.CommitAsync(reference, ChangePlan(paragraph.ParaId));

            Assert.True(result.Committed);
            Assert.NotNull(result.Receipt);
            Assert.Equal("oidc-subject", result.Receipt!.Actor!.Subject);
            Assert.Equal(result.Document, result.Receipt.OutputDocument);
            Assert.Equal(ApplyOutcome.Committed, result.Receipt.Outcome);
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    private static DocumentPlan ChangePlan(string paragraphId, RevisionMetadata? revision = null) => new()
    {
        Revision = revision,
        Operations = new PlanOperation[]
        {
            new ChangeTextOp
            {
                Target = new TextSpanAnchor
                {
                    ParaId = paragraphId,
                    Expect = "Acme Corp",
                    Occurrence = 0
                },
                With = "Globex",
                Mode = ChangeMode.Tracked
            }
        }
    };

    private static StreamHandle Handle(byte[] bytes) =>
        new(new MemoryStream(bytes, writable: false), "contract.docx");

    private static string Hash(byte[] bytes) =>
        Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static string? RevisionAttribute(OpenXmlElement element, string name) =>
        element.GetAttributes().FirstOrDefault(attribute =>
            attribute.LocalName == name &&
            attribute.NamespaceUri == "http://schemas.openxmlformats.org/wordprocessingml/2006/main").Value;

    private sealed class FixedActorProvider : IAuditActorProvider
    {
        private readonly AuditActor _actor;

        public FixedActorProvider(AuditActor actor) => _actor = actor;

        public AuditActor GetCurrentActor() => _actor;
    }

    private sealed class FixedTimeProvider : TimeProvider
    {
        private readonly DateTimeOffset _now;

        public FixedTimeProvider(DateTimeOffset now) => _now = now;

        public override DateTimeOffset GetUtcNow() => _now;
    }
}
