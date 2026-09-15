using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// What a comparison says when it cannot say everything: which areas it examined, which
/// it could not represent, and the property that no plan is offered while any gap
/// remains.
/// </summary>
public sealed class ComparisonCoverageTests
{
    // ── Useful findings alongside an explicit gap ────────────────────────

    /// <summary>
    /// The case the coverage model exists for. A supported paragraph edit and an
    /// unsupported image change arrive together: the paragraph finding is still reported,
    /// the image area is named as blocked, and no plan is offered.
    /// </summary>
    [Fact]
    public void A_supported_paragraph_change_survives_an_unsupported_image_change()
    {
        var original = WithImage(Document("Old text", "Second paragraph."));
        var revised = Edit(original, document =>
        {
            document.MainDocumentPart!.Document.Descendants<Text>().First().Text = "New text";
            using var output = document.MainDocumentPart.ImageParts.Single()
                .GetStream(FileMode.Create, FileAccess.Write);
            output.Write(OtherPng());
        });

        var comparison = Client().CompareDocuments(original, revised);

        // The useful finding is still there.
        var difference = Assert.Single(comparison.Differences);
        Assert.Equal(DocumentDifferenceKind.Changed, difference.Kind);
        Assert.Equal("Old text", difference.Before);
        Assert.Equal("New text", difference.After);

        // The gap is named, not implied.
        Assert.Contains(comparison.Diagnostics, d => d.Code == "unsupported-image-change");
        Assert.Contains(comparison.Coverage, area =>
            area.Name == "images" && area.State == ComparisonAreaState.Blocked);

        // And no plan is offered while a gap remains.
        Assert.False(comparison.IsComplete);
        Assert.Null(comparison.Plan);
    }

    /// <summary>
    /// Each unsupported area is named on its own, so a caller can tell which one blocked
    /// the plan rather than being told something somewhere changed.
    /// </summary>
    [Theory]
    [InlineData("images", "unsupported-image-change")]
    [InlineData("notes", "unsupported-note-change")]
    [InlineData("headersAndFooters", "unsupported-header-footer-change")]
    public void Each_unsupported_area_is_named_separately(string area, string code)
    {
        var original = WithImage(Document("Text"));
        var revised = Edit(original, document => Mutate(document, area));

        var comparison = Client().CompareDocuments(original, revised);

        Assert.Contains(comparison.Diagnostics, d => d.Code == code);
        Assert.Contains(comparison.Coverage, entry =>
            entry.Name == area && entry.State == ComparisonAreaState.Blocked);
        Assert.Null(comparison.Plan);

        // Only the area that actually changed is blocked.
        Assert.All(comparison.Coverage.Where(entry => entry.Name != area && entry.Name != "otherParts"),
            entry => Assert.NotEqual(ComparisonAreaState.Blocked, entry.State));
    }

    // ── Unchanged is not the same as not compared ────────────────────────

    /// <summary>
    /// Two identical documents report every area unchanged and offer an empty plan.
    /// Unsupported content that is present but identical does not block anything.
    /// </summary>
    [Fact]
    public void Unchanged_unsupported_content_does_not_block_the_plan()
    {
        var original = WithImage(Document("Text", "More text."));
        var revised = Edit(original, document =>
            document.MainDocumentPart!.Document.Descendants<Text>().First().Text = "Changed text");

        var comparison = Client().CompareDocuments(original, revised);

        // The image is present in both and identical, so it is unchanged, not blocked.
        Assert.Contains(comparison.Coverage, area =>
            area.Name == "images" && area.State == ComparisonAreaState.Unchanged);
        Assert.True(comparison.IsComplete, string.Join("; ",
            comparison.Diagnostics.Select(d => $"{d.Code}: {d.Message}")));
        Assert.NotNull(comparison.Plan);
    }

    /// <summary>
    /// A comparison that stopped before examining anything reports every area as not
    /// compared. Reporting them unchanged would assert something nobody checked.
    /// </summary>
    [Fact]
    public void A_comparison_that_stopped_early_reports_nothing_as_unchanged()
    {
        var original = Document("Text");
        var revised = Document("Other text");

        var comparison = Client().CompareDocuments(original, revised, new DocumentComparisonOptions
        {
            MaximumDocumentBytes = 32
        });

        Assert.False(comparison.IsComplete);
        Assert.Null(comparison.Plan);
        Assert.NotEmpty(comparison.Coverage);
        Assert.All(comparison.Coverage, area =>
            Assert.Equal(ComparisonAreaState.NotCompared, area.State));
        Assert.DoesNotContain(comparison.Coverage, area => area.State == ComparisonAreaState.Unchanged);
    }

    /// <summary>Body paragraphs that match are reported unchanged, not absent.</summary>
    [Fact]
    public void Identical_documents_report_body_paragraphs_unchanged()
    {
        var document = Document("Text", "More text.");

        var comparison = Client().CompareDocuments(document, document);

        Assert.Empty(comparison.Differences);
        Assert.Contains(comparison.Coverage, area =>
            area.Name == "bodyParagraphs" && area.State == ComparisonAreaState.Unchanged);
        Assert.True(comparison.IsComplete);
    }

    // ── Truncation is visible and never complete ─────────────────────────

    /// <summary>
    /// Hitting the difference ceiling is reported, blocks the plan, and marks the body
    /// area blocked rather than changed: the finding list is known to be partial.
    /// </summary>
    [Fact]
    public void Reaching_the_difference_ceiling_is_visible_and_blocks_completeness()
    {
        var original = Document(Enumerable.Range(0, 12).Select(i => $"Paragraph {i}").ToArray());
        var revised = Document(Enumerable.Range(0, 12).Select(i => $"Changed {i}").ToArray());

        var comparison = Client().CompareDocuments(original, revised, new DocumentComparisonOptions
        {
            MaximumDifferences = 3
        });

        Assert.False(comparison.IsComplete);
        Assert.Null(comparison.Plan);
        Assert.Contains(comparison.Diagnostics, d => d.Code == "comparison-difference-limit-exceeded");
        Assert.Contains(comparison.Coverage, area =>
            area.Name == "bodyParagraphs" && area.State == ComparisonAreaState.Blocked);
        Assert.True(comparison.Differences.Count <= 3);
    }

    /// <summary>
    /// A resource ceiling cannot be traded for a plan. Every ceiling that fires leaves
    /// IsComplete false and Plan null.
    /// </summary>
    [Theory]
    [InlineData("bytes")]
    [InlineData("paragraphs")]
    [InlineData("differences")]
    public void No_resource_ceiling_can_produce_a_plan(string ceiling)
    {
        var original = Document(Enumerable.Range(0, 8).Select(i => $"Paragraph {i}").ToArray());
        var revised = Document(Enumerable.Range(0, 8).Select(i => $"Changed {i}").ToArray());

        var options = ceiling switch
        {
            "bytes" => new DocumentComparisonOptions { MaximumDocumentBytes = 64 },
            "paragraphs" => new DocumentComparisonOptions { MaximumParagraphs = 2 },
            _ => new DocumentComparisonOptions { MaximumDifferences = 1 }
        };

        var comparison = Client().CompareDocuments(original, revised, options);

        Assert.False(comparison.IsComplete);
        Assert.Null(comparison.Plan);
        Assert.NotEmpty(comparison.Diagnostics);
    }

    // ── Stable output ────────────────────────────────────────────────────

    /// <summary>
    /// Repeated runs over the same inputs produce the same findings, diagnostics and
    /// coverage in the same order, so a caller can diff two runs meaningfully.
    /// </summary>
    [Fact]
    public void Repeated_runs_produce_identical_ordered_results()
    {
        var original = WithImage(Document("Old text", "Second."));
        var revised = Edit(original, document =>
        {
            document.MainDocumentPart!.Document.Descendants<Text>().First().Text = "New text";
            using var output = document.MainDocumentPart.ImageParts.Single()
                .GetStream(FileMode.Create, FileAccess.Write);
            output.Write(OtherPng());
        });

        var client = Client();
        var first = client.CompareDocuments(original, revised);
        var second = client.CompareDocuments(original, revised);

        Assert.Equal(
            first.Diagnostics.Select(d => d.Code + "|" + d.Path),
            second.Diagnostics.Select(d => d.Code + "|" + d.Path));
        Assert.Equal(
            first.Coverage.Select(a => a.Name + "|" + a.State),
            second.Coverage.Select(a => a.Name + "|" + a.State));
        Assert.Equal(first.OriginalSha256, second.OriginalSha256);
        Assert.Equal(first.RevisedSha256, second.RevisedSha256);
    }

    /// <summary>The exact input hashes are reported and identify the compared bytes.</summary>
    [Fact]
    public void Source_hashes_identify_the_compared_bytes()
    {
        var original = Document("Text");
        var revised = Document("Other");

        var comparison = Client().CompareDocuments(original, revised);

        Assert.Equal(Hash(original), comparison.OriginalSha256);
        Assert.Equal(Hash(revised), comparison.RevisedSha256);
        Assert.NotEqual(comparison.OriginalSha256, comparison.RevisedSha256);
    }

    // ── The adapter cannot loosen any of this ────────────────────────────

    /// <summary>
    /// The agent tool reports the same incomplete state and carries no plan. An adapter
    /// that serialized one anyway would be a way around the whole property.
    /// </summary>
    [Fact]
    public async Task The_comparison_tool_never_serializes_a_plan_for_an_incomplete_result()
    {
        var store = new MemoryDocumentProvider("docs");
        var original = WithImage(Document("Old text"));
        var revised = Edit(original, document =>
        {
            document.MainDocumentPart!.Document.Descendants<Text>().First().Text = "New text";
            using var output = document.MainDocumentPart.ImageParts.Single()
                .GetStream(FileMode.Create, FileAccess.Write);
            output.Write(OtherPng());
        });

        var originalReference = store.Add("original.docx", original);
        var revisedReference = store.Add("revised.docx", revised);
        var tools = new OfficeAgentTools(
            new OfficeAgentClient(new DocumentProviderRegistry(new[] { store }), new WordModule()));

        var payload = await tools.CompareDocuments(
            "docs", originalReference.ItemId, "docs", revisedReference.ItemId);

        using var json = JsonDocument.Parse(payload);
        Assert.False(json.RootElement.GetProperty("isComplete").GetBoolean());

        // The plan is absent or null; there is no shape in which a caller receives one.
        if (json.RootElement.TryGetProperty("plan", out var plan))
            Assert.Equal(JsonValueKind.Null, plan.ValueKind);

        Assert.DoesNotContain("changeText", payload, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>A complete comparison through the adapter does carry its plan.</summary>
    [Fact]
    public async Task The_comparison_tool_carries_the_plan_when_coverage_is_complete()
    {
        var store = new MemoryDocumentProvider("docs");
        var original = Document("Old text");
        var revised = Document("New text");
        var originalReference = store.Add("original.docx", original);
        var revisedReference = store.Add("revised.docx", revised);
        var tools = new OfficeAgentTools(
            new OfficeAgentClient(new DocumentProviderRegistry(new[] { store }), new WordModule()));

        var payload = await tools.CompareDocuments(
            "docs", originalReference.ItemId, "docs", revisedReference.ItemId);

        using var json = JsonDocument.Parse(payload);
        Assert.True(json.RootElement.GetProperty("isComplete").GetBoolean());
        Assert.Contains("changeText", payload, StringComparison.OrdinalIgnoreCase);
    }

    // ── Helpers ──────────────────────────────────────────────────────────

    private static OfficeAgentClient Client() => new(new WordModule());

    private static string Hash(byte[] bytes) =>
        Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(bytes)).ToLowerInvariant();

    private static void Mutate(WordprocessingDocument document, string area)
    {
        var main = document.MainDocumentPart!;
        switch (area)
        {
            case "images":
                using (var output = main.ImageParts.Single().GetStream(FileMode.Create, FileAccess.Write))
                    output.Write(OtherPng());
                break;

            case "notes":
                var footnotes = main.AddNewPart<FootnotesPart>();
                footnotes.Footnotes = new Footnotes(
                    new Footnote(new Paragraph(new Run(new Text("A note.")))) { Id = 2 });
                footnotes.Footnotes.Save();
                break;

            case "headersAndFooters":
                var header = main.AddNewPart<HeaderPart>();
                header.Header = new Header(new Paragraph(new Run(new Text("A header."))));
                header.Header.Save();
                break;
        }
    }

    private static byte[] OtherPng() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAIAAACQd1PeAAAADElEQVR4nGP4z8AAAAMBAQDJ/pLvAAAAAElFTkSuQmCC");

    private static byte[] Png() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");

    private static byte[] Document(params string[] paragraphs)
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

    private static byte[] WithImage(byte[] source) => Edit(source, document =>
    {
        var main = document.MainDocumentPart!;
        var image = main.AddImagePart(ImagePartType.Png);
        using (var input = new MemoryStream(Png()))
            image.FeedData(input);
        main.Document.Save();
    });

    private static byte[] Edit(byte[] source, Action<WordprocessingDocument> mutate)
    {
        var copy = (byte[])source.Clone();
        using var ms = new MemoryStream();
        ms.Write(copy, 0, copy.Length);
        ms.Position = 0;
        using (var document = WordprocessingDocument.Open(ms, isEditable: true))
        {
            mutate(document);
            document.MainDocumentPart!.Document.Save();
        }

        return ms.ToArray();
    }
}
