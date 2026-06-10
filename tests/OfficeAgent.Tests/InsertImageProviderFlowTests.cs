using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// End-to-end coverage for the <see cref="InsertImageOp"/> opaque-document-id flow:
/// the host stages an image on disk under a provider connection (whose allow-list
/// permits image extensions), registers it to receive the id, and references it
/// from the plan. The client resolves the id to base64 upstream of the engine.
/// </summary>
public class InsertImageProviderFlowTests
{
    [Fact]
    public async Task InsertImage_by_document_id_resolves_through_the_provider()
    {
        using var workspace = new ImageWorkspace();
        var client = workspace.Client;

        // 1. Register the Word document.
        var doc = await client.RegisterBytesAsync("docs", workspace.DocsRoot, BuildSimpleDoc("paragraph"), "doc.docx");

        // 2. Register the image with its own connection (image extensions allowed).
        var image = await client.RegisterBytesAsync("images", workspace.ImagesRoot, MinimalPng(), "logo.png");

        // 3. Build a plan that references the image by id, not by bytes or path.
        var paraId = (await client.InspectAsync("docs", doc.ItemId)).Paragraphs.First().ParaId;
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertImageOp
                {
                    Target            = new TextSpanAnchor { ParaId = paraId, Expect = "" },
                    ImageConnectionId = "images",
                    ImageDocumentId   = image.ItemId,
                    ImageType         = "png",
                    WidthPx = 64, HeightPx = 64,
                    Position = InsertPosition.After,
                    AltText  = "Logo"
                }
            }
        };

        var result = await client.CommitAsync("docs", doc.ItemId, plan);
        Assert.True(result.Committed);

        // 4. Verify the saved document contains a drawing + image part.
        using var content = await client.OpenReadAsync(result.Document);
        using var saved = WordprocessingDocument.Open(content.Stream, false);
        Assert.Single(saved.MainDocumentPart!.ImageParts);
        Assert.Single(saved.MainDocumentPart!.Document.Body!.Descendants<Drawing>());
    }

    [Fact]
    public async Task InsertImage_unknown_image_id_fails_through_provider_with_not_found()
    {
        using var workspace = new ImageWorkspace();
        var client = workspace.Client;
        var doc = await client.RegisterBytesAsync("docs", workspace.DocsRoot, BuildSimpleDoc("paragraph"), "doc.docx");
        var paraId = (await client.InspectAsync("docs", doc.ItemId)).Paragraphs.First().ParaId;

        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertImageOp
                {
                    Target            = new TextSpanAnchor { ParaId = paraId, Expect = "" },
                    ImageConnectionId = "images",
                    ImageDocumentId   = "ffffffffffffffff",
                    ImageType         = "png",
                    WidthPx = 64, HeightPx = 64
                }
            }
        };

        var ex = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            client.PreviewAsync("docs", doc.ItemId, plan));
        Assert.Equal(ProviderErrorCode.NotFound, ex.Code);
        Assert.Equal("images", ex.ConnectionId);
    }

    private static byte[] BuildSimpleDoc(string text)
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var p = new Paragraph { ParagraphId = "BBBBBBBB" };
            p.AppendChild(new Run(new Text(text) { Space = SpaceProcessingModeValues.Preserve }));
            main.Document = new Document(new Body(p));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    private static byte[] MinimalPng() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4//8/AwAI/AL+XJ/PNwAAAABJRU5ErkJggg==");

    private sealed class ImageWorkspace : IDisposable
    {
        private readonly ServiceProvider _serviceProvider;

        public ImageWorkspace()
        {
            DocsRoot   = Path.Combine(Path.GetTempPath(), $"officeagent-docs-{Guid.NewGuid():N}");
            ImagesRoot = Path.Combine(Path.GetTempPath(), $"officeagent-images-{Guid.NewGuid():N}");
            Directory.CreateDirectory(DocsRoot);
            Directory.CreateDirectory(ImagesRoot);

            var services = new ServiceCollection();
            services.AddWordFormat();
            services.AddFileSystemDocumentProvider("docs", DocsRoot);
            services.AddFileSystemDocumentProvider("images", ImagesRoot, opts =>
            {
                opts.AllowedExtensions = new[] { ".png", ".jpeg", ".jpg", ".gif", ".bmp", ".tiff" };
            });
            services.AddOfficeAgent();
            _serviceProvider = services.BuildServiceProvider();
            Client = _serviceProvider.GetRequiredService<OfficeAgentClient>();
        }

        public string DocsRoot { get; }
        public string ImagesRoot { get; }
        public OfficeAgentClient Client { get; }

        public void Dispose()
        {
            _serviceProvider.Dispose();
            if (Directory.Exists(DocsRoot))   Directory.Delete(DocsRoot,   recursive: true);
            if (Directory.Exists(ImagesRoot)) Directory.Delete(ImagesRoot, recursive: true);
        }
    }
}
