using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// Template preflight: what a template offers, what a whole batch would do before any
/// storage write, and what binds a reviewed intent to the commit that applies it.
/// </summary>
public sealed class TemplatePreflightTests
{
    // ── Discovery ────────────────────────────────────────────────────────

    /// <summary>
    /// Discovery reports the slots a template really has, and marks a duplicated tag
    /// unbindable rather than silently picking one of them.
    /// </summary>
    [Fact]
    public async Task Discovery_reports_bindable_slots_and_names_the_ambiguous_ones()
    {
        using var workspace = new TemplateWorkspace();
        var template = await workspace.RegisterAsync("quote.docx", QuoteTemplate(duplicateTag: true));

        var discovered = await workspace.Client.DiscoverTemplateAsync(template);

        var customer = Assert.Single(discovered.Slots, slot => slot.Name == "CustomerName");
        Assert.True(customer.Bindable);
        Assert.Equal(1, customer.Occurrences);

        var duplicated = Assert.Single(discovered.Slots, slot => slot.Name == "Duplicated");
        Assert.False(duplicated.Bindable);
        Assert.Equal(2, duplicated.Occurrences);
        Assert.Contains(discovered.Diagnostics, d => d.Code == "ambiguous-template-slot" && d.Path == "Duplicated");

        Assert.NotEmpty(discovered.TemplateSha256);
    }

    /// <summary>
    /// Repeating rows are reported as candidates, not as confirmed structures. Nothing in
    /// the file format declares a row repeatable; the fields are inferred from the
    /// placeholder convention, and the result says so.
    /// </summary>
    [Fact]
    public async Task Repeating_rows_are_reported_as_unconfirmed_candidates()
    {
        using var workspace = new TemplateWorkspace();
        var template = await workspace.RegisterAsync("quote.docx", QuoteTemplate());

        var discovered = await workspace.Client.DiscoverTemplateAsync(template);

        var candidate = Assert.Single(discovered.RepeatingRows);
        Assert.Equal("table#0", candidate.TablePath);
        Assert.False(candidate.Confirmed);
        Assert.Equal(new[] { "Description", "Price", "Quantity" }, candidate.Fields);
    }

    // ── Preview writes nothing ───────────────────────────────────────────

    /// <summary>
    /// Preview validates and produces no documents. The connection holds only the
    /// template afterwards.
    /// </summary>
    [Fact]
    public async Task Preview_validates_without_writing_anything()
    {
        using var workspace = new TemplateWorkspace();
        var template = await workspace.RegisterAsync("quote.docx", QuoteTemplate());
        int filesBefore = Directory.GetFiles(workspace.Root).Length;

        var preview = await workspace.Client.PreviewTemplateBatchAsync(template, Batch(
            ("first.docx", "Fabrikam"),
            ("second.docx", "Contoso")));

        Assert.True(preview.IsValid, Explain(preview));
        Assert.Equal(2, preview.Items.Count);
        Assert.All(preview.Items, item => Assert.True(item.OperationCount > 0));
        Assert.Equal(filesBefore, Directory.GetFiles(workspace.Root).Length);
    }

    /// <summary>
    /// Every item is reported, not just the first failure, so one call surfaces all the
    /// problems a caller has to fix.
    /// </summary>
    [Fact]
    public async Task Preview_reports_every_failing_item_not_only_the_first()
    {
        using var workspace = new TemplateWorkspace();
        var template = await workspace.RegisterAsync("quote.docx", QuoteTemplate());

        var request = new TemplateBatchRequest
        {
            Items = new[]
            {
                Item("first.docx", ("NoSuchSlot", "x")),
                Item("second.docx", ("AlsoMissing", "y")),
                Item("third.docx", ("CustomerName", "Fabrikam"))
            }
        };

        var preview = await workspace.Client.PreviewTemplateBatchAsync(template, request);

        Assert.False(preview.IsValid);
        Assert.False(preview.Items[0].IsValid);
        Assert.False(preview.Items[1].IsValid);
        Assert.NotEmpty(preview.Items[0].Diagnostics);
        Assert.NotEmpty(preview.Items[1].Diagnostics);
    }

    /// <summary>Per-item diagnostics are capped, and a truncated list says so.</summary>
    [Fact]
    public async Task Item_diagnostics_are_capped_and_marked_when_truncated()
    {
        using var workspace = new TemplateWorkspace();
        var template = await workspace.RegisterAsync("quote.docx", QuoteTemplate());

        var unknown = Enumerable.Range(0, 20)
            .ToDictionary(i => $"Unknown{i}", i => (string?)"value", StringComparer.Ordinal);

        var request = new TemplateBatchRequest
        {
            Limits = new TemplateBatchLimits { MaximumDiagnosticsPerItem = 3 },
            Items = new[]
            {
                new TemplateBatchItem
                {
                    OutputName = "first.docx",
                    Binding = new TemplateBinding
                    {
                        Values = unknown,
                        RejectUnknownValues = true,
                        MissingValueBehavior = MissingTemplateValueBehavior.Empty
                    }
                }
            }
        };

        var preview = await workspace.Client.PreviewTemplateBatchAsync(template, request);

        var item = Assert.Single(preview.Items);
        Assert.True(item.DiagnosticsTruncated);
        Assert.Equal(3, item.Diagnostics.Count);
    }

    /// <summary>Duplicate output names are a batch-level refusal.</summary>
    [Fact]
    public async Task Duplicate_output_names_are_refused()
    {
        using var workspace = new TemplateWorkspace();
        var template = await workspace.RegisterAsync("quote.docx", QuoteTemplate());

        var preview = await workspace.Client.PreviewTemplateBatchAsync(template, Batch(
            ("same.docx", "Fabrikam"),
            ("same.docx", "Contoso")));

        Assert.False(preview.IsValid);
        Assert.Contains(preview.Diagnostics, d => d.Code == "duplicate-output-name");
    }

    // ── The token binds intent to bytes ──────────────────────────────────

    /// <summary>The same template and the same batch produce the same token.</summary>
    [Fact]
    public async Task An_unchanged_batch_produces_a_stable_token()
    {
        using var workspace = new TemplateWorkspace();
        var template = await workspace.RegisterAsync("quote.docx", QuoteTemplate());

        var first = await workspace.Client.PreviewTemplateBatchAsync(template, Batch(("a.docx", "Fabrikam")));
        var second = await workspace.Client.PreviewTemplateBatchAsync(template, Batch(("a.docx", "Fabrikam")));

        Assert.Equal(first.Token.TemplateSha256, second.Token.TemplateSha256);
        Assert.Equal(first.Token.BatchSha256, second.Token.BatchSha256);
    }

    /// <summary>
    /// Each thing a reviewer would care about changes the token: the values, the output
    /// name, the item order and the template itself.
    /// </summary>
    [Fact]
    public async Task Changed_values_names_order_or_template_all_change_the_token()
    {
        using var workspace = new TemplateWorkspace();
        var template = await workspace.RegisterAsync("quote.docx", QuoteTemplate());
        var client = workspace.Client;

        var baseline = await client.PreviewTemplateBatchAsync(
            template, Batch(("a.docx", "Fabrikam"), ("b.docx", "Contoso")));

        var changedValue = await client.PreviewTemplateBatchAsync(
            template, Batch(("a.docx", "Fabrikam Services"), ("b.docx", "Contoso")));
        Assert.NotEqual(baseline.Token.BatchSha256, changedValue.Token.BatchSha256);

        var changedName = await client.PreviewTemplateBatchAsync(
            template, Batch(("renamed.docx", "Fabrikam"), ("b.docx", "Contoso")));
        Assert.NotEqual(baseline.Token.BatchSha256, changedName.Token.BatchSha256);

        var reordered = await client.PreviewTemplateBatchAsync(
            template, Batch(("b.docx", "Contoso"), ("a.docx", "Fabrikam")));
        Assert.NotEqual(baseline.Token.BatchSha256, reordered.Token.BatchSha256);

        // A different template, same batch: the template half moves, the batch half does not.
        var other = await workspace.RegisterAsync("other.docx", QuoteTemplate(extraParagraph: true));
        var changedTemplate = await client.PreviewTemplateBatchAsync(
            other, Batch(("a.docx", "Fabrikam"), ("b.docx", "Contoso")));
        Assert.NotEqual(baseline.Token.TemplateSha256, changedTemplate.Token.TemplateSha256);
        Assert.Equal(baseline.Token.BatchSha256, changedTemplate.Token.BatchSha256);
    }

    /// <summary>
    /// A commit carrying a token that no longer matches is refused, and writes nothing.
    /// </summary>
    [Fact]
    public async Task A_commit_with_a_stale_token_is_refused_and_writes_nothing()
    {
        using var workspace = new TemplateWorkspace();
        var template = await workspace.RegisterAsync("quote.docx", QuoteTemplate());

        var preview = await workspace.Client.PreviewTemplateBatchAsync(template, Batch(("a.docx", "Fabrikam")));
        Assert.True(preview.IsValid, Explain(preview));

        // The reviewed intent was "Fabrikam"; the commit asks for something else.
        int filesBefore = Directory.GetFiles(workspace.Root).Length;
        var result = await workspace.Client.PopulateTemplateBatchAsync(
            template, Batch(("a.docx", "Someone Else")), preview.Token);

        Assert.All(result.Items, item => Assert.Equal(TemplateItemOutcome.Failed, item.Outcome));
        Assert.Contains(result.Items.SelectMany(i => i.Diagnostics), d => d.Code == "stale-batch-preview");
        Assert.Equal(filesBefore, Directory.GetFiles(workspace.Root).Length);
    }

    /// <summary>A commit carrying the matching token proceeds and produces the outputs.</summary>
    [Fact]
    public async Task A_commit_with_the_matching_token_proceeds()
    {
        using var workspace = new TemplateWorkspace();
        var template = await workspace.RegisterAsync("quote.docx", QuoteTemplate());
        var batch = Batch(("a.docx", "Fabrikam"), ("b.docx", "Contoso"));

        var preview = await workspace.Client.PreviewTemplateBatchAsync(template, batch);
        var result = await workspace.Client.PopulateTemplateBatchAsync(template, batch, preview.Token);

        Assert.True(result.Committed, string.Join("; ",
            result.Items.SelectMany(i => i.Diagnostics).Select(d => $"{d.Code}: {d.Message}")));
        Assert.All(result.Items, item => Assert.Equal(TemplateItemOutcome.Committed, item.Outcome));
        Assert.Contains(Directory.GetFiles(workspace.Root), f => Path.GetFileName(f) == "a.docx");
        Assert.Contains(Directory.GetFiles(workspace.Root), f => Path.GetFileName(f) == "b.docx");
    }

    // ── Host budgets ─────────────────────────────────────────────────────

    /// <summary>
    /// A request cannot raise a host budget. Asking for a larger batch than the host
    /// allows fails preflight before any output is written.
    /// </summary>
    [Fact]
    public async Task A_request_cannot_inflate_the_host_budget()
    {
        using var workspace = new TemplateWorkspace(new TemplateBatchLimits { MaximumDocuments = 2 });
        var template = await workspace.RegisterAsync("quote.docx", QuoteTemplate());
        int filesBefore = Directory.GetFiles(workspace.Root).Length;

        var request = new TemplateBatchRequest
        {
            MaximumDocuments = 1000,
            Limits = new TemplateBatchLimits { MaximumDocuments = 1000 },
            Items = new[]
            {
                Item("a.docx", ("CustomerName", "One")),
                Item("b.docx", ("CustomerName", "Two")),
                Item("c.docx", ("CustomerName", "Three"))
            }
        };

        Assert.Equal(2, workspace.Client.EffectiveLimits(request).MaximumDocuments);

        var preview = await workspace.Client.PreviewTemplateBatchAsync(template, request);
        Assert.False(preview.IsValid);
        Assert.Contains(preview.Diagnostics, d => d.Code == "batch-too-large");

        var result = await workspace.Client.PopulateTemplateBatchAsync(template, request);
        Assert.All(result.Items, item => Assert.Equal(TemplateItemOutcome.Failed, item.Outcome));
        Assert.Equal(filesBefore, Directory.GetFiles(workspace.Root).Length);
    }

    /// <summary>A request may ask to be held to something stricter than the host budget.</summary>
    [Fact]
    public void A_request_may_lower_a_budget()
    {
        using var workspace = new TemplateWorkspace(new TemplateBatchLimits { MaximumDocuments = 50 });

        var effective = workspace.Client.EffectiveLimits(new TemplateBatchRequest
        {
            Limits = new TemplateBatchLimits { MaximumDocuments = 5 }
        });

        Assert.Equal(5, effective.MaximumDocuments);
    }

    /// <summary>Row budgets are enforced per output and across the whole batch.</summary>
    [Fact]
    public async Task Row_budgets_are_enforced_per_document_and_in_aggregate()
    {
        using var workspace = new TemplateWorkspace(new TemplateBatchLimits
        {
            MaximumRowsPerDocument = 3,
            MaximumTotalRows = 4
        });
        var template = await workspace.RegisterAsync("quote.docx", QuoteTemplate());

        var perDocument = await workspace.Client.PreviewTemplateBatchAsync(template, new TemplateBatchRequest
        {
            Items = new[] { ItemWithRows("a.docx", rows: 5) }
        });
        Assert.False(perDocument.IsValid);
        Assert.Contains(perDocument.Items[0].Diagnostics, d => d.Code == "too-many-rows");

        var aggregate = await workspace.Client.PreviewTemplateBatchAsync(template, new TemplateBatchRequest
        {
            Items = new[] { ItemWithRows("a.docx", rows: 3), ItemWithRows("b.docx", rows: 3) }
        });
        Assert.False(aggregate.IsValid);
        Assert.Contains(aggregate.Diagnostics, d => d.Code == "too-many-total-rows");
    }

    // ── Outcomes ─────────────────────────────────────────────────────────

    /// <summary>
    /// An item never attempted is reported as skipped, not as failed. A caller has to be
    /// able to tell "this was refused" from "this was never tried".
    /// </summary>
    [Fact]
    public async Task An_item_after_a_failure_is_skipped_not_failed()
    {
        using var workspace = new TemplateWorkspace();
        var template = await workspace.RegisterAsync("quote.docx", QuoteTemplate());

        var request = new TemplateBatchRequest
        {
            ContinueOnError = false,
            Items = new[]
            {
                Item("a.docx", ("CustomerName", "Fabrikam")),
                new TemplateBatchItem
                {
                    OutputName = "b.docx",
                    Binding = new TemplateBinding
                    {
                        Values = new Dictionary<string, string?> { ["NoSuchSlot"] = "x" },
                        RejectUnknownValues = true,
                        MissingValueBehavior = MissingTemplateValueBehavior.Empty
                    }
                },
                Item("c.docx", ("CustomerName", "Contoso"))
            }
        };

        var result = await workspace.Client.PopulateTemplateBatchAsync(template, request);

        Assert.Equal(TemplateItemOutcome.Committed, result.Items[0].Outcome);
        Assert.Equal(TemplateItemOutcome.Failed, result.Items[1].Outcome);
        Assert.Equal(TemplateItemOutcome.Skipped, result.Items[2].Outcome);
        Assert.DoesNotContain(Directory.GetFiles(workspace.Root), f => Path.GetFileName(f) == "c.docx");
    }

    /// <summary>
    /// A provider that accepts the bytes and then fails to confirm leaves an output that
    /// may or may not exist. That item is reported Uncertain, and its diagnostic tells the
    /// caller not to retry the name. Reporting it as Failed would send them to recreate a
    /// document that might already be there.
    /// </summary>
    [Fact]
    public async Task A_fault_after_an_accepted_write_reports_the_item_uncertain()
    {
        using var workspace = new TemplateWorkspace();
        var template = await workspace.RegisterAsync("quote.docx", QuoteTemplate());

        // The same client, with the provider wrapped so the second save faults after the
        // bytes have been taken.
        var faulting = new FaultAfterAcceptProvider(workspace.Provider, failFromCall: 0);
        var client = new OfficeAgentClient(
            new DocumentProviderRegistry(new IDocumentProvider[] { faulting }), new WordModule());

        var result = await client.PopulateTemplateBatchAsync(
            DocumentReference.ForFileSystem("workspace", template.ItemId),
            Batch(("a.docx", "Fabrikam")));

        var item = Assert.Single(result.Items);
        Assert.Equal(TemplateItemOutcome.Uncertain, item.Outcome);
        Assert.False(item.Committed);
        Assert.True(faulting.AcceptedWrite, "the provider took the bytes before faulting");
        Assert.Contains(item.Diagnostics, d =>
            d.Message.Contains("do not retry", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>
    /// A provider that refuses before taking any bytes is a plain failure, not an
    /// uncertain one: nothing was written and the caller may safely retry.
    /// </summary>
    [Fact]
    public async Task A_refusal_before_any_write_reports_the_item_failed()
    {
        using var workspace = new TemplateWorkspace();
        var template = await workspace.RegisterAsync("quote.docx", QuoteTemplate());

        var refusing = new RefuseBeforeWriteProvider(workspace.Provider);
        var client = new OfficeAgentClient(
            new DocumentProviderRegistry(new IDocumentProvider[] { refusing }), new WordModule());

        var result = await client.PopulateTemplateBatchAsync(
            DocumentReference.ForFileSystem("workspace", template.ItemId),
            Batch(("a.docx", "Fabrikam")));

        var item = Assert.Single(result.Items);
        Assert.Equal(TemplateItemOutcome.Failed, item.Outcome);
        Assert.DoesNotContain(item.Diagnostics, d =>
            d.Message.Contains("do not retry", StringComparison.OrdinalIgnoreCase));
    }

    /// <summary>Wraps a provider so a save takes the bytes and then fails to confirm.</summary>
    private sealed class FaultAfterAcceptProvider : IDocumentProvider
    {
        private readonly IDocumentProvider _inner;
        private readonly int _failFromCall;
        private int _saves;

        public FaultAfterAcceptProvider(IDocumentProvider inner, int failFromCall)
        {
            _inner = inner;
            _failFromCall = failFromCall;
        }

        public bool AcceptedWrite { get; private set; }

        public string Provider => _inner.Provider;

        public string ConnectionId => _inner.ConnectionId;

        public Task<DocumentReference> RegisterAsync(string source, CancellationToken cancellationToken = default) =>
            _inner.RegisterAsync(source, cancellationToken);

        public Task<DocumentContent> OpenReadAsync(
            DocumentReference reference, CancellationToken cancellationToken = default) =>
            _inner.OpenReadAsync(reference, cancellationToken);

        public async Task<DocumentReference> SaveAsync(
            DocumentReference source,
            Stream content,
            SaveDocumentOptions options,
            CancellationToken cancellationToken = default)
        {
            if (++_saves > _failFromCall)
            {
                // Drain the content first: the point is that storage has seen the bytes.
                using var buffer = new MemoryStream();
                await content.CopyToAsync(buffer, cancellationToken).ConfigureAwait(false);
                AcceptedWrite = true;
                throw new DocumentProviderException(
                    ProviderErrorCode.IO,
                    "The storage accepted the document but the result could not be confirmed.",
                    Provider, ConnectionId, source.ItemId);
            }

            return await _inner.SaveAsync(source, content, options, cancellationToken).ConfigureAwait(false);
        }

        public Task RemoveAsync(DocumentReference reference, CancellationToken cancellationToken = default) =>
            _inner.RemoveAsync(reference, cancellationToken);
    }

    /// <summary>Wraps a provider so a save is refused before any byte is taken.</summary>
    private sealed class RefuseBeforeWriteProvider : IDocumentProvider
    {
        private readonly IDocumentProvider _inner;

        public RefuseBeforeWriteProvider(IDocumentProvider inner) => _inner = inner;

        public string Provider => _inner.Provider;

        public string ConnectionId => _inner.ConnectionId;

        public Task<DocumentReference> RegisterAsync(string source, CancellationToken cancellationToken = default) =>
            _inner.RegisterAsync(source, cancellationToken);

        public Task<DocumentContent> OpenReadAsync(
            DocumentReference reference, CancellationToken cancellationToken = default) =>
            _inner.OpenReadAsync(reference, cancellationToken);

        public Task<DocumentReference> SaveAsync(
            DocumentReference source,
            Stream content,
            SaveDocumentOptions options,
            CancellationToken cancellationToken = default) =>
            throw new DocumentProviderException(
                ProviderErrorCode.AccessDenied,
                "The connection refused the write.",
                Provider, ConnectionId, source.ItemId);

        public Task RemoveAsync(DocumentReference reference, CancellationToken cancellationToken = default) =>
            _inner.RemoveAsync(reference, cancellationToken);
    }

    // ── Helpers ──────────────────────────────────────────────────────────

    private static string Explain(TemplateBatchPreview preview) => string.Join("; ",
        preview.Diagnostics.Concat(preview.Items.SelectMany(i => i.Diagnostics))
            .Select(d => $"{d.Code}: {d.Message}"));

    private static TemplateBatchItem Item(string name, params (string Key, string Value)[] values) => new()
    {
        OutputName = name,
        Binding = new TemplateBinding
        {
            Values = values.ToDictionary(v => v.Key, v => (string?)v.Value, StringComparer.Ordinal),
            MissingValueBehavior = MissingTemplateValueBehavior.Empty
        }
    };

    private static TemplateBatchItem ItemWithRows(string name, int rows) => new()
    {
        OutputName = name,
        Binding = new TemplateBinding
        {
            Values = new Dictionary<string, string?> { ["CustomerName"] = "Fabrikam" },
            MissingValueBehavior = MissingTemplateValueBehavior.Empty,
            RepeatingTables = new[]
            {
                new RepeatingTableBinding
                {
                    TablePath = "table#0",
                    TemplateRowIndex = 1,
                    Records = Enumerable.Range(0, rows).Select(i =>
                        (IReadOnlyDictionary<string, string?>)new Dictionary<string, string?>
                        {
                            ["Description"] = $"Line {i}",
                            ["Quantity"] = "1",
                            ["Price"] = "10.00"
                        }).ToArray()
                }
            }
        }
    };

    private static TemplateBatchRequest Batch(params (string Name, string Customer)[] items) => new()
    {
        Items = items.Select(entry => Item(entry.Name, ("CustomerName", entry.Customer))).ToArray()
    };

    private static byte[] QuoteTemplate(bool duplicateTag = false, bool extraParagraph = false)
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var body = new Body(
                new Paragraph(new Run(new Text("Quote for ")), Control("CustomerName", "CUSTOMER")),
                new Table(
                    new TableProperties(new TableStyle { Val = "TableGrid" }),
                    new TableGrid(
                        new GridColumn { Width = "3600" },
                        new GridColumn { Width = "1200" },
                        new GridColumn { Width = "1800" }),
                    Row("Description", "Quantity", "Price"),
                    Row("{{Description}}", "{{Quantity}}", "{{Price}}")),
                new Paragraph());

            if (duplicateTag)
            {
                body.AppendChild(new Paragraph(Control("Duplicated", "ONE")));
                body.AppendChild(new Paragraph(Control("Duplicated", "TWO")));
            }

            if (extraParagraph)
                body.AppendChild(new Paragraph(new Run(new Text("An extra clause."))));

            main.Document = new Document(body);
            main.Document.Save();
        }

        return ms.ToArray();
    }

    private static SdtRun Control(string tag, string text) => new(
        new SdtProperties(new Tag { Val = tag }, new SdtId { Val = Math.Abs(tag.GetHashCode() % 10000) + 10 }),
        new SdtContentRun(new Run(new Text(text))));

    private static TableRow Row(params string[] values) => new(values.Select(value =>
        new TableCell(new Paragraph(new Run(new Text(value))))));

    private sealed class TemplateWorkspace : IDisposable
    {
        public TemplateWorkspace(TemplateBatchLimits? limits = null)
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-template-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);

            Provider = new FileSystemDocumentProvider(new FileSystemDocumentProviderOptions
            {
                ConnectionId = "workspace",
                RootPath = Root,
                DefaultChangeMode = ChangeMode.Direct
            });

            Client = new OfficeAgentClient(
                new DocumentProviderRegistry(new[] { Provider }),
                new WordModule());

            if (limits is not null) Client.WithTemplateLimits(limits);
        }

        public string Root { get; }

        public OfficeAgentClient Client { get; }

        public FileSystemDocumentProvider Provider { get; }

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
