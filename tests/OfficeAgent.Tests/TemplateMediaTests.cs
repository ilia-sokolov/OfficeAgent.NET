using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// Typed template bindings: an image or a native chart placed through the same preview,
/// authorization and budgets a scalar value goes through, and refused explicitly when
/// the slot cannot hold it.
/// </summary>
public sealed class TemplateMediaTests
{
    // ── Existing behavior is unchanged ───────────────────────────────────

    /// <summary>
    /// A scalar-only request behaves exactly as it did. Typed bindings are additive, so
    /// nothing that worked before needs rewriting.
    /// </summary>
    [Fact]
    public async Task A_scalar_only_request_is_unaffected()
    {
        using var workspace = new MediaWorkspace();
        var template = await workspace.RegisterAsync("quote.docx", Template());

        var result = await workspace.Client.PopulateTemplateBatchAsync(template, new TemplateBatchRequest
        {
            Items = new[] { Scalar("out.docx", "Fabrikam") }
        });

        Assert.True(result.Committed, Explain(result));
        Assert.Equal(TemplateItemOutcome.Committed, Assert.Single(result.Items).Outcome);
    }

    // ── Discovery reports media slots ────────────────────────────────────

    /// <summary>Discovery names the slots that can hold an image.</summary>
    [Fact]
    public async Task Discovery_reports_image_slots()
    {
        using var workspace = new MediaWorkspace();
        var template = await workspace.RegisterAsync("quote.docx", Template());

        var discovered = await workspace.Client.DiscoverTemplateAsync(template);

        Assert.Contains(discovered.MediaSlots, slot =>
            slot.Name == "Photo" && slot.MediaKind == "image");
    }

    /// <summary>
    /// A Word template never reports a chart slot, because this engine has no native
    /// Word chart handling. Absence here is the honest answer, not an omission.
    /// </summary>
    [Fact]
    public async Task A_word_template_reports_no_chart_slots()
    {
        using var workspace = new MediaWorkspace();
        var template = await workspace.RegisterAsync("quote.docx", Template());

        var discovered = await workspace.Client.DiscoverTemplateAsync(template);

        Assert.DoesNotContain(discovered.MediaSlots, slot => slot.MediaKind == "chart");
    }

    // ── Images bind, with alt text and a valid package ───────────────────

    /// <summary>
    /// A bound image reaches the output as a real picture carrying its alt text, and the
    /// package still validates.
    /// </summary>
    [Fact]
    public async Task A_bound_image_lands_with_its_alt_text_in_a_valid_package()
    {
        using var workspace = new MediaWorkspace();
        var template = await workspace.RegisterAsync("quote.docx", Template());

        var request = new TemplateBatchRequest { Items = new[] { WithPhoto("out.docx", Png()) } };
        var preview = await workspace.Client.PreviewTemplateBatchAsync(template, request);
        Assert.True(preview.IsValid, Explain(preview));
        Assert.Equal(1, preview.Items[0].ImageCount);
        Assert.True(preview.Items[0].ImageBytes > 0);

        var result = await workspace.Client.PopulateTemplateBatchAsync(template, request, preview.Token);
        Assert.True(result.Committed, Explain(result));

        var bytes = await File.ReadAllBytesAsync(Path.Combine(workspace.Root, "out.docx"));
        using var document = WordprocessingDocument.Open(new MemoryStream(bytes, writable: false), false);
        var main = document.MainDocumentPart!;

        Assert.NotEmpty(main.ImageParts);
        var descriptions = main.Document!.Body!
            .Descendants<DocumentFormat.OpenXml.Drawing.Wordprocessing.DocProperties>()
            .Select(properties => properties.Description?.Value)
            .ToList();
        Assert.Contains("A product photograph", descriptions);

        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document));
    }

    // ── Refusals are explicit ────────────────────────────────────────────

    /// <summary>An image aimed at a slot that cannot hold one is an explicit error.</summary>
    [Fact]
    public async Task An_image_bound_to_an_unknown_slot_is_refused()
    {
        using var workspace = new MediaWorkspace();
        var template = await workspace.RegisterAsync("quote.docx", Template());

        var preview = await workspace.Client.PreviewTemplateBatchAsync(template, new TemplateBatchRequest
        {
            Items = new[] { WithTyped("out.docx", "NoSuchSlot", Image(Png())) }
        });

        Assert.False(preview.IsValid);
        Assert.Contains(preview.Items[0].Diagnostics, d => d.Code == "wrong-slot-kind");
    }

    /// <summary>
    /// A chart binding against a Word template is refused with a message that says why,
    /// rather than being quietly dropped.
    /// </summary>
    [Fact]
    public async Task A_chart_bound_to_a_word_template_is_refused_explicitly()
    {
        using var workspace = new MediaWorkspace();
        var template = await workspace.RegisterAsync("quote.docx", Template());

        var preview = await workspace.Client.PreviewTemplateBatchAsync(template, new TemplateBatchRequest
        {
            Items = new[]
            {
                WithTyped("out.docx", "Photo", new TemplateChartValue
                {
                    Categories = new[] { "Q1" },
                    Series = new[] { new ChartSeries { Name = "Revenue", Values = new double?[] { 1 } } }
                })
            }
        });

        Assert.False(preview.IsValid);
        var diagnostic = Assert.Single(preview.Items[0].Diagnostics, d => d.Code == "unsupported-template-feature");
        Assert.Contains("PowerPoint only", diagnostic.Message, StringComparison.Ordinal);
    }

    /// <summary>An image with no alt text is refused rather than silently embedded.</summary>
    [Fact]
    public async Task An_image_without_alt_text_is_refused()
    {
        using var workspace = new MediaWorkspace();
        var template = await workspace.RegisterAsync("quote.docx", Template());

        var preview = await workspace.Client.PreviewTemplateBatchAsync(template, new TemplateBatchRequest
        {
            Items = new[] { WithTyped("out.docx", "Photo", new TemplateImageValue
            {
                Base64Bytes = Convert.ToBase64String(Png()),
                ImageType = "png",
                AltText = null
            }) }
        });

        Assert.False(preview.IsValid);
        Assert.Contains(preview.Items[0].Diagnostics, d => d.Code == "missing-alt-text");
    }

    /// <summary>
    /// Bytes that are not the declared image type are refused. A caller cannot label
    /// arbitrary content as a PNG and have it embedded under that name.
    /// </summary>
    [Fact]
    public async Task Bytes_that_do_not_match_the_declared_type_are_refused()
    {
        using var workspace = new MediaWorkspace();
        var template = await workspace.RegisterAsync("quote.docx", Template());

        var preview = await workspace.Client.PreviewTemplateBatchAsync(template, new TemplateBatchRequest
        {
            Items = new[] { WithTyped("out.docx", "Photo", new TemplateImageValue
            {
                Base64Bytes = Convert.ToBase64String(new byte[] { 1, 2, 3, 4, 5, 6, 7, 8 }),
                ImageType = "png",
                AltText = "Not really a picture"
            }) }
        });

        Assert.False(preview.IsValid);
        Assert.Contains(preview.Items[0].Diagnostics, d => d.Code == "image-type-mismatch");
    }

    /// <summary>An oversized image is refused and no output is produced.</summary>
    [Fact]
    public async Task An_oversized_image_is_refused_and_writes_nothing()
    {
        using var workspace = new MediaWorkspace(new TemplateBatchLimits
        {
            Media = new TemplateMediaLimits { MaximumImageBytes = 16 }
        });
        var template = await workspace.RegisterAsync("quote.docx", Template());
        int before = Directory.GetFiles(workspace.Root).Length;

        var request = new TemplateBatchRequest { Items = new[] { WithPhoto("out.docx", Png()) } };

        var preview = await workspace.Client.PreviewTemplateBatchAsync(template, request);
        Assert.False(preview.IsValid);
        Assert.Contains(preview.Items[0].Diagnostics, d => d.Code == "image-too-large");

        var result = await workspace.Client.PopulateTemplateBatchAsync(template, request);
        Assert.All(result.Items, item => Assert.Equal(TemplateItemOutcome.Failed, item.Outcome));
        Assert.Equal(before, Directory.GetFiles(workspace.Root).Length);
    }

    /// <summary>There is no URL form, and a binding with neither source is refused.</summary>
    [Fact]
    public async Task An_image_binding_needs_exactly_one_source()
    {
        using var workspace = new MediaWorkspace();
        var template = await workspace.RegisterAsync("quote.docx", Template());

        var preview = await workspace.Client.PreviewTemplateBatchAsync(template, new TemplateBatchRequest
        {
            Items = new[] { WithTyped("out.docx", "Photo", new TemplateImageValue
            {
                AltText = "Nothing to load"
            }) }
        });

        Assert.False(preview.IsValid);
        var diagnostic = Assert.Single(preview.Items[0].Diagnostics, d => d.Code == "invalid-image-binding");
        Assert.Contains("does not fetch images", diagnostic.Message, StringComparison.Ordinal);
    }

    /// <summary>A slot bound twice, once scalar and once typed, is refused.</summary>
    [Fact]
    public async Task A_slot_bound_twice_is_refused()
    {
        using var workspace = new MediaWorkspace();
        var template = await workspace.RegisterAsync("quote.docx", Template());

        var preview = await workspace.Client.PreviewTemplateBatchAsync(template, new TemplateBatchRequest
        {
            Items = new[]
            {
                new TemplateBatchItem
                {
                    OutputName = "out.docx",
                    Binding = new TemplateBinding
                    {
                        Values = new Dictionary<string, string?> { ["CustomerName"] = "Fabrikam" },
                        TypedValues = new Dictionary<string, TemplateValue>
                        {
                            ["CustomerName"] = new TemplateTextValue("Contoso")
                        },
                        MissingValueBehavior = MissingTemplateValueBehavior.Empty
                    }
                }
            }
        });

        Assert.False(preview.IsValid);
        Assert.Contains(preview.Items[0].Diagnostics, d => d.Code == "duplicate-template-binding");
    }

    // ── The token covers the resolved bytes ──────────────────────────────

    /// <summary>
    /// Two previews of the same image agree; a different image moves the media hash even
    /// though every other part of the batch is identical.
    /// </summary>
    [Fact]
    public async Task A_changed_image_moves_the_media_hash()
    {
        using var workspace = new MediaWorkspace();
        var template = await workspace.RegisterAsync("quote.docx", Template());

        var first = await workspace.Client.PreviewTemplateBatchAsync(
            template, new TemplateBatchRequest { Items = new[] { WithPhoto("out.docx", Png()) } });
        var same = await workspace.Client.PreviewTemplateBatchAsync(
            template, new TemplateBatchRequest { Items = new[] { WithPhoto("out.docx", Png()) } });
        var different = await workspace.Client.PreviewTemplateBatchAsync(
            template, new TemplateBatchRequest { Items = new[] { WithPhoto("out.docx", Png(variant: true)) } });

        Assert.Equal(first.Token.MediaSha256, same.Token.MediaSha256);
        Assert.NotEqual(first.Token.MediaSha256, different.Token.MediaSha256);
        Assert.NotEmpty(first.Token.MediaSha256);
    }

    /// <summary>
    /// A provider-held image resolved behind an unchanged reference still invalidates the
    /// preview when its bytes change. The batch hash cannot see that; the media hash can.
    /// </summary>
    [Fact]
    public async Task Replacing_a_referenced_image_invalidates_the_preview()
    {
        using var workspace = new MediaWorkspace();
        var template = await workspace.RegisterAsync("quote.docx", Template());
        var imagePath = Path.Combine(workspace.Root, "photo.png");
        await File.WriteAllBytesAsync(imagePath, Png());
        var image = await workspace.Client.RegisterAsync("workspace", imagePath);

        var request = new TemplateBatchRequest
        {
            Items = new[] { WithTyped("out.docx", "Photo", new TemplateImageValue
            {
                ImageConnectionId = "workspace",
                ImageDocumentId = image.ItemId,
                ImageType = "png",
                AltText = "A product photograph"
            }) }
        };

        var before = await workspace.Client.PreviewTemplateBatchAsync(template, request);
        Assert.True(before.IsValid, Explain(before));

        // Same reference, different bytes behind it.
        await File.WriteAllBytesAsync(imagePath, Png(variant: true));
        var after = await workspace.Client.PreviewTemplateBatchAsync(template, request);

        Assert.Equal(before.Token.BatchSha256, after.Token.BatchSha256);
        Assert.NotEqual(before.Token.MediaSha256, after.Token.MediaSha256);
    }

    // ── Native charts in a deck ──────────────────────────────────────────

    /// <summary>
    /// A deck that already holds a native chart reports it as a chart slot, and a chart
    /// binding rewrites its data while leaving it an editable chart with its embedded
    /// workbook intact.
    /// </summary>
    [Fact]
    public async Task A_chart_binding_updates_a_deck_chart_and_keeps_it_editable()
    {
        using var workspace = new MediaWorkspace();
        var template = await workspace.RegisterAsync("deck.pptx", DeckWithChart());

        var discovered = await workspace.DeckClient.DiscoverTemplateAsync(template);
        var slot = Assert.Single(discovered.MediaSlots, media => media.MediaKind == "chart");

        var request = new TemplateBatchRequest
        {
            Items = new[]
            {
                new TemplateBatchItem
                {
                    OutputName = "bound.pptx",
                    Binding = new TemplateBinding
                    {
                        TypedValues = new Dictionary<string, TemplateValue>
                        {
                            [slot.Name] = new TemplateChartValue
                            {
                                Kind = ChartKind.Bar,
                                Categories = new[] { "North", "South" },
                                Series = new[]
                                {
                                    new ChartSeries { Name = "Bookings", Values = new double?[] { 41, 58 } }
                                },
                                Title = "Bookings by region"
                            }
                        },
                        MissingValueBehavior = MissingTemplateValueBehavior.Empty,
                        Mode = ChangeMode.Direct
                    }
                }
            }
        };

        var preview = await workspace.DeckClient.PreviewTemplateBatchAsync(template, request);
        Assert.True(preview.IsValid, Explain(preview));

        var result = await workspace.DeckClient.PopulateTemplateBatchAsync(template, request, preview.Token);
        Assert.True(result.Committed, Explain(result));

        var bytes = await File.ReadAllBytesAsync(Path.Combine(workspace.Root, "bound.pptx"));
        using var document = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(
            new MemoryStream(bytes, writable: false), false);
        var slide = document.PresentationPart!.SlideParts.Single();
        var chart = Assert.Single(slide.ChartParts);

        // Still a native chart, still editable: the embedded workbook is what Word and
        // PowerPoint open when a reader clicks "Edit Data".
        Assert.NotNull(chart.ChartSpace);
        Assert.Single(chart.Parts.Select(part => part.OpenXmlPart)
            .OfType<DocumentFormat.OpenXml.Packaging.EmbeddedPackagePart>());

        var xml = chart.ChartSpace!.OuterXml;
        Assert.Contains("Bookings", xml, StringComparison.Ordinal);
        Assert.Contains("North", xml, StringComparison.Ordinal);

        Assert.Empty(new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document));
    }

    /// <summary>A chart binding carrying more points than the budget is refused.</summary>
    [Fact]
    public async Task An_oversized_chart_binding_is_refused()
    {
        using var workspace = new MediaWorkspace(new TemplateBatchLimits
        {
            Media = new TemplateMediaLimits { MaximumChartPoints = 2 }
        });
        var template = await workspace.RegisterAsync("deck.pptx", DeckWithChart());

        var discovered = await workspace.DeckClient.DiscoverTemplateAsync(template);
        var slot = Assert.Single(discovered.MediaSlots, media => media.MediaKind == "chart");

        var preview = await workspace.DeckClient.PreviewTemplateBatchAsync(template, new TemplateBatchRequest
        {
            Items = new[]
            {
                new TemplateBatchItem
                {
                    OutputName = "bound.pptx",
                    Binding = new TemplateBinding
                    {
                        TypedValues = new Dictionary<string, TemplateValue>
                        {
                            [slot.Name] = new TemplateChartValue
                            {
                                Categories = new[] { "A", "B", "C" },
                                Series = new[]
                                {
                                    new ChartSeries { Name = "S", Values = new double?[] { 1, 2, 3 } }
                                }
                            }
                        },
                        MissingValueBehavior = MissingTemplateValueBehavior.Empty,
                        Mode = ChangeMode.Direct
                    }
                }
            }
        });

        Assert.False(preview.IsValid);
        Assert.Contains(preview.Items[0].Diagnostics, d => d.Code == "chart-too-large");
    }

    /// <summary>A deck carrying one native chart, built through the engine's own verb.</summary>
    private static byte[] DeckWithChart()
    {
        var module = new OfficeAgent.PowerPoint.PowerPointModule();
        var client = new OfficeAgentClient(module);
        using var applied = client.Commit(
            new StreamHandle(new MemoryStream(module.CreateBlank(), writable: false)),
            new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new InsertChartOp
                    {
                        Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
                        Kind = ChartKind.ClusteredColumn,
                        Categories = new[] { "Q1", "Q2" },
                        Series = new[]
                        {
                            new ChartSeries { Name = "Revenue", Values = new double?[] { 10, 12 } }
                        },
                        Title = "Revenue",
                        Description = "Revenue by quarter"
                    }
                }
            });

        if (!applied.Committed)
            throw new InvalidOperationException(string.Join("; ",
                applied.Report.Errors.Select(error => error.Message)));
        return applied.ToBytes();
    }

    // ── Helpers ──────────────────────────────────────────────────────────

    private static string Explain(TemplateBatchPreview preview) => string.Join("; ",
        preview.Diagnostics.Concat(preview.Items.SelectMany(i => i.Diagnostics))
            .Select(d => $"{d.Code}: {d.Message}"));

    private static string Explain(TemplateBatchResult result) => string.Join("; ",
        result.Items.SelectMany(i => i.Diagnostics).Select(d => $"{d.Code}: {d.Message}"));

    private static TemplateImageValue Image(byte[] bytes) => new()
    {
        Base64Bytes = Convert.ToBase64String(bytes),
        ImageType = "png",
        AltText = "A product photograph"
    };

    private static TemplateBatchItem Scalar(string name, string customer) => new()
    {
        OutputName = name,
        Binding = new TemplateBinding
        {
            Values = new Dictionary<string, string?> { ["CustomerName"] = customer },
            MissingValueBehavior = MissingTemplateValueBehavior.Empty
        }
    };

    private static TemplateBatchItem WithPhoto(string name, byte[] png) =>
        WithTyped(name, "Photo", Image(png));

    private static TemplateBatchItem WithTyped(string name, string slot, TemplateValue value) => new()
    {
        OutputName = name,
        Binding = new TemplateBinding
        {
            Values = new Dictionary<string, string?> { ["CustomerName"] = "Fabrikam" },
            TypedValues = new Dictionary<string, TemplateValue> { [slot] = value },
            MissingValueBehavior = MissingTemplateValueBehavior.Empty
        }
    };

    /// <summary>A minimal but genuine PNG, so signature checks see a real header.</summary>
    private static byte[] Png(bool variant = false)
    {
        // 1x1 PNG. The variant flips a pixel byte so the two differ by content.
        var png = Convert.FromBase64String(
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==");
        if (variant) png[png.Length - 20] ^= 0x01;
        return png;
    }

    private static byte[] Template()
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Quote for ")), Control("CustomerName", "CUSTOMER")),
                new Paragraph(Control("Photo", "Photo")),
                new Paragraph()));
            main.Document.Save();
        }

        return ms.ToArray();
    }

    private static SdtRun Control(string tag, string text) => new(
        new SdtProperties(new Tag { Val = tag }, new SdtId { Val = Math.Abs(tag.GetHashCode() % 10000) + 10 }),
        new SdtContentRun(new Run(new Text(text))));

    private sealed class MediaWorkspace : IDisposable
    {
        public MediaWorkspace(TemplateBatchLimits? limits = null)
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-media-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);

            var provider = new FileSystemDocumentProvider(new FileSystemDocumentProviderOptions
            {
                ConnectionId = "workspace",
                RootPath = Root,
                DefaultChangeMode = ChangeMode.Direct,
                AllowedExtensions = new[] { ".docx", ".pptx", ".xlsx", ".png" }
            });

            Client = new OfficeAgentClient(
                new DocumentProviderRegistry(new[] { provider }), new WordModule());
            DeckClient = new OfficeAgentClient(
                new DocumentProviderRegistry(new[] { provider }),
                new OfficeAgent.PowerPoint.PowerPointModule());

            if (limits is not null)
            {
                Client.WithTemplateLimits(limits);
                DeckClient.WithTemplateLimits(limits);
            }
        }

        public string Root { get; }

        public OfficeAgentClient Client { get; }

        /// <summary>A client with the deck module registered, for the chart cases.</summary>
        public OfficeAgentClient DeckClient { get; }

        public async Task<DocumentReference> RegisterAsync(string name, byte[] content)
        {
            var path = Path.Combine(Root, name);
            await File.WriteAllBytesAsync(path, content);
            return await Client.RegisterAsync("workspace", path);
        }

        public void Dispose()
        {
            if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true);
        }
    }
}
