using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Tests;

public class NodeModelTests
{
    private static OfficeAgentClient Office() => new(new WordModule());

    private static StreamHandle Handle(byte[] bytes) => new(new MemoryStream(bytes), "contract.docx");

    [Fact]
    public void Inspect_surfaces_document_property_nodes()
    {
        var inspect = Office().Inspect(Handle(DocxFactory.Contract()));

        Assert.Contains(inspect.Nodes, n => n.Kind == "docProperty" && n.Path == "core/title");
        Assert.All(inspect.Nodes.Where(n => n.Kind == "docProperty"), n => Assert.NotNull(n.Anchor));
    }

    [Fact]
    public void SetProperty_sets_a_core_document_property()
    {
        var bytes = DocxFactory.Contract();
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new SetPropertyOp
                {
                    Target = new NodeAnchor { Kind = "docProperty", Path = "core/title" },
                    Value = "Service Agreement (Globex)"
                }
            }
        };

        var result = Office().Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        Assert.Equal("Service Agreement (Globex)", doc.PackageProperties.Title);
    }

    [Fact]
    public void SetProperty_update_fields_on_open_is_deferred_but_commits()
    {
        var office = Office();
        var bytes = DocxFactory.Contract();
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new SetPropertyOp { Target = new NodeAnchor { Kind = "field" }, Name = "updateOnOpen" }
            }
        };

        var report = office.Preview(Handle(bytes), plan);
        Assert.True(report.IsValid);
        Assert.Contains(report.Changes, c => c.Capability == Capability.DeferredToWordOnOpen);

        var result = office.Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var settings = doc.MainDocumentPart!.DocumentSettingsPart!.Settings;
        Assert.NotNull(settings.GetFirstChild<UpdateFieldsOnOpen>());
    }

    [Fact]
    public void SetProperty_field_refresh_needs_renderer_and_is_rejected()
    {
        var office = Office();
        var bytes = DocxFactory.Contract();
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new SetPropertyOp { Target = new NodeAnchor { Kind = "field" }, Name = "refresh" }
            }
        };

        var report = office.Preview(Handle(bytes), plan);
        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, e => e.Code == ValidationErrorCodes.RequiresRenderer);
        Assert.Contains(report.Changes, c => c.Capability == Capability.NeedsRenderer);

        var result = office.Commit(Handle(bytes), plan);
        Assert.False(result.Committed);
    }

    [Fact]
    public void Revision_accept_resolves_tracked_change_to_plain_text()
    {
        var office = Office();
        var bytes = DocxFactory.Contract();
        var clauseParaId = office.Inspect(Handle(bytes)).Paragraphs
            .First(p => p.Text.Contains("shall provide", StringComparison.Ordinal)).ParaId;

        var trackedPlan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new ChangeTextOp
                {
                    Target = new TextSpanAnchor { ParaId = clauseParaId, Expect = "Acme Corp", Occurrence = 0 },
                    With = "Globex Inc.",
                    Mode = ChangeMode.Tracked
                }
            }
        };
        var tracked = office.Commit(Handle(bytes), trackedPlan);
        Assert.True(tracked.Committed);
        var trackedBytes = OfficeAgentClient.ToBytes(tracked);

        var acceptPlan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new RevisionOp
                {
                    Target = new NodeAnchor { Kind = "revision", Path = "all" },
                    Action = RevisionAction.Accept
                }
            }
        };
        var accepted = office.Commit(Handle(trackedBytes), acceptPlan);
        Assert.True(accepted.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(accepted)), false);
        var body = doc.MainDocumentPart!.Document.Body!;

        Assert.Empty(body.Descendants<InsertedRun>());
        Assert.Empty(body.Descendants<DeletedRun>());

        var clause = body.Descendants<Paragraph>().Single(p => p.ParagraphId?.Value == "00000002");
        var text = string.Concat(clause.Descendants<Text>().Select(t => t.Text));
        Assert.Equal("Globex Inc. shall provide services to Acme Corp.", text);
    }

    [Fact]
    public void Conflicting_operations_on_the_same_anchor_are_rejected()
    {
        var office = Office();
        var bytes = DocxFactory.Contract();
        var clauseParaId = office.Inspect(Handle(bytes)).Paragraphs
            .First(p => p.Text.Contains("shall provide", StringComparison.Ordinal)).ParaId;

        var anchor = new TextSpanAnchor { ParaId = clauseParaId, Expect = "Acme Corp", Occurrence = 0 };
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new ChangeTextOp { Target = anchor, With = "Globex", Mode = ChangeMode.Direct },
                new ChangeTextOp { Target = anchor, With = "Initech", Mode = ChangeMode.Direct }
            }
        };

        var report = office.Preview(Handle(bytes), plan);
        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, e => e.Code == ValidationErrorCodes.OperationConflict);
    }
}
