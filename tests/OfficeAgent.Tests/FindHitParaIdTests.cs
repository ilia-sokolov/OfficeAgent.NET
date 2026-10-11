using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Excel;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// <see cref="FindHit.ParaId"/> gives the paragraph a text hit sits in, without
/// pattern-matching the anchor (#26). Callers reached for it in the round-3 integration trial.
/// </summary>
public class FindHitParaIdTests
{
    [Fact]
    public void A_word_text_hit_carries_its_paragraph_id()
    {
        var client = new OfficeAgentClient(new WordModule());
        var document = Commit(client, new WordModule().CreateBlank(), new ChangeTextOp
        {
            Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = string.Empty },
            With = "Prepared for Northwind Labs.",
            Mode = ChangeMode.Direct
        });

        var hit = Assert.Single(client.Find(Handle(document), new FindQuery("Northwind")));

        var anchor = Assert.IsType<TextSpanAnchor>(hit.Anchor);
        Assert.False(string.IsNullOrEmpty(hit.ParaId));
        Assert.Equal(anchor.ParaId, hit.ParaId);
        Assert.Equal(client.Inspect(document).Paragraphs.Single().ParaId, hit.ParaId);
    }

    [Fact]
    public void A_deck_text_hit_carries_its_slide_paragraph_id()
    {
        var client = new OfficeAgentClient(new PowerPointModule());
        var deck = Commit(client, new PowerPointModule().CreateBlank(), new ChangeTextOp
        {
            Target = new TextSpanAnchor { ParaId = "slide256/shape2/p0", Expect = string.Empty },
            With = "Quarterly review",
            Mode = ChangeMode.Direct
        });

        var hit = Assert.Single(client.Find(Handle(deck), new FindQuery("Quarterly")));

        Assert.Equal("slide256/shape2/p0", hit.ParaId);
    }

    [Fact]
    public void A_cell_hit_has_no_paragraph_id()
    {
        var client = new OfficeAgentClient(new ExcelModule());

        var hit = Assert.Single(client.Find(Handle(XlsxFactory.WorkbookWithTable()), new FindQuery("EMEA")
        {
            Options = new MatchOptions { SpreadsheetValueView = SpreadsheetValueView.Displayed }
        }));

        Assert.IsType<CellAnchor>(hit.Anchor);
        Assert.Null(hit.ParaId);
    }

    private static StreamHandle Handle(byte[] bytes) => new(new MemoryStream(bytes));

    private static byte[] Commit(OfficeAgentClient client, byte[] document, PlanOperation operation)
    {
        using var applied = client.Commit(Handle(document), new DocumentPlan { Operations = new[] { operation } });
        Assert.True(applied.Committed, string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return applied.ToBytes();
    }
}
