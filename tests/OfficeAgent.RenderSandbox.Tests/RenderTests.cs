using System.Buffers.Binary;
using OfficeAgent.Abstractions;

namespace OfficeAgent.RenderSandbox.Tests;

/// <summary>Real renders through the reference image: the corpus, and every failure bound.</summary>
[Collection(SandboxCollection.Name)]
public sealed class RenderTests
{
    private static readonly byte[] PngSignature = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };

    public static IEnumerable<object[]> Corpus() => new[]
    {
        new object[] { "three-pages.docx", Documents.Word(3), 3 },
        new object[] { "deck.pptx", Documents.Deck(), 1 },
        new object[] { "book.xlsx", Documents.Workbook(), 1 },
    };

    [Theory]
    [MemberData(nameof(Corpus))]
    public async Task The_corpus_renders_to_the_expected_pages(string fileName, byte[] document, int pages)
    {
        var result = await Sandbox.Renderer().RenderAsync(new MemoryStream(document), new RenderOptions
        {
            FileName = fileName,
            Dpi = 72
        });

        Assert.True(result.Succeeded, $"{result.FailureCode}: {result.Message}");
        Assert.Equal(pages, result.PageCount);
        Assert.Equal(Enumerable.Range(1, pages), result.Pages.Select(page => page.PageNumber));
        foreach (var page in result.Pages)
        {
            Assert.Equal(PngSignature, page.Content.Take(8));
            var width = BinaryPrimitives.ReadInt32BigEndian(page.Content.AsSpan(16, 4));
            var height = BinaryPrimitives.ReadInt32BigEndian(page.Content.AsSpan(20, 4));
            Assert.InRange(width, 100, 2000);
            Assert.InRange(height, 100, 2000);
        }
    }

    [Fact]
    public async Task An_oversized_input_is_refused_before_any_container_starts()
    {
        var before = Sandbox.ContainersNamed("officeagent-render-");
        var result = await Sandbox.Renderer().RenderAsync(new MemoryStream(Documents.Word(1)), new RenderOptions
        {
            FileName = "small.docx",
            MaximumInputBytes = 16
        });
        Assert.Equal(RenderFailureCodes.InputLimitExceeded, result.FailureCode);
        Assert.Equal(before, Sandbox.ContainersNamed("officeagent-render-"));
    }

    [Fact]
    public async Task The_page_limit_holds_for_a_long_document()
    {
        var result = await Sandbox.Renderer().RenderAsync(new MemoryStream(Documents.Word(12)), new RenderOptions
        {
            FileName = "long.docx",
            Dpi = 72,
            MaximumPages = 5
        });
        Assert.Equal(RenderFailureCodes.PageLimitExceeded, result.FailureCode);
        Assert.Throws<RenderFailedException>(() => result.Pages);
    }

    [Fact]
    public async Task The_output_limit_holds()
    {
        var result = await Sandbox.Renderer().RenderAsync(new MemoryStream(Documents.Word(2)), new RenderOptions
        {
            FileName = "two.docx",
            Dpi = 150,
            MaximumOutputBytes = 1024
        });
        Assert.Equal(RenderFailureCodes.OutputLimitExceeded, result.FailureCode);
    }

    [Fact]
    public async Task Bytes_that_are_not_a_document_fail_explicitly()
    {
        var garbage = new byte[4096];
        new Random(7).NextBytes(garbage);
        var result = await Sandbox.Renderer().RenderAsync(new MemoryStream(garbage), new RenderOptions
        {
            FileName = "garbage.docx"
        });
        Assert.False(result.Succeeded);
        Assert.Contains(result.FailureCode, new[] { RenderFailureCodes.RendererFailed, RenderFailureCodes.RendererOutputMissing });
    }

    [Fact]
    public async Task A_missing_image_reports_the_renderer_unavailable()
    {
        var result = await Sandbox.Renderer(new OfficeAgent.Deploy.Renderer.SandboxLimits { Image = "officeagent-renderer:does-not-exist" })
            .RenderAsync(new MemoryStream(Documents.Word(1)), new RenderOptions { FileName = "one.docx" });
        Assert.Equal(RenderFailureCodes.RendererUnavailable, result.FailureCode);
    }

    [Fact]
    public async Task Failure_messages_carry_no_document_content()
    {
        const string secret = "SECRET-7f3a-do-not-leak";
        var result = await Sandbox.Renderer().RenderAsync(new MemoryStream(Documents.Word(8, secret)), new RenderOptions
        {
            FileName = "secret.docx",
            Dpi = 72,
            MaximumPages = 2
        });
        Assert.False(result.Succeeded);
        Assert.DoesNotContain(secret, result.Message ?? "");
        Assert.DoesNotContain("/tmp", result.Message ?? "");
    }
}
