using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using OfficeAgent.Abstractions;
using OfficeAgent.Core.DocumentProviders;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Core;

public sealed partial class OfficeAgentClient
{
    /// <summary>Matches the <c>{{Field}}</c> placeholder convention used by repeating rows.</summary>
    private static readonly Regex PlaceholderPattern = new(
        @"\{\{\s*(?<name>[^{}]+?)\s*\}\}", RegexOptions.CultureInvariant, TimeSpan.FromSeconds(1));

    private TemplateBatchLimits _templateLimits = TemplateBatchLimits.Default;

    /// <summary>
    /// Sets the host budgets for template batches. A request may ask for something
    /// stricter; it can never raise these.
    /// </summary>
    public OfficeAgentClient WithTemplateLimits(TemplateBatchLimits limits)
    {
        if (limits is null) throw new ArgumentNullException(nameof(limits));
        limits.Validate();
        _templateLimits = limits;
        return this;
    }

    /// <summary>Gets the host budgets currently in force for template batches.</summary>
    public TemplateBatchLimits TemplateLimits => _templateLimits;

    /// <summary>
    /// Reports what a template offers to bind: its scalar slots, the rows that look like
    /// repeating templates, and anything that would make a binding ambiguous.
    /// </summary>
    /// <remarks>
    /// Repeating rows are reported as candidates. Nothing in the file format declares a
    /// row repeatable, so they are inferred from the <c>{{Field}}</c> placeholder
    /// convention and marked unconfirmed. Scalar slots are confirmed: a content control
    /// or a named shape really is there.
    /// </remarks>
    public async Task<TemplateDiscoveryResult> DiscoverTemplateAsync(
        DocumentReference template,
        CancellationToken cancellationToken = default)
    {
        if (template is null) throw new ArgumentNullException(nameof(template));

        byte[] bytes;
        using (var content = await OpenReadAsync(template, cancellationToken).ConfigureAwait(false))
        using (var buffer = new MemoryStream())
        {
            await content.Stream.CopyToAsync(buffer, 81920, cancellationToken).ConfigureAwait(false);
            bytes = buffer.ToArray();
        }

        var inspect = await InspectAsync(
            new StreamHandle(new MemoryStream(bytes, writable: false)),
            InspectOptions.Default,
            cancellationToken).ConfigureAwait(false);

        var diagnostics = new List<WorkflowDiagnostic>();

        var grouped = inspect.StructuralAnchors
            .Where(anchor => anchor.Kind is "contentControl" or "shapeName")
            .GroupBy(anchor => anchor.Tag, StringComparer.Ordinal)
            .ToList();

        var slots = new List<TemplateSlot>();
        foreach (var group in grouped)
        {
            var first = group.First();
            slots.Add(new TemplateSlot
            {
                Name = group.Key,
                Kind = first.Kind,
                Occurrences = group.Count(),
                CurrentText = null
            });

            if (group.Count() != 1)
                diagnostics.Add(new WorkflowDiagnostic
                {
                    Code = "ambiguous-template-slot",
                    Message = $"Template slot '{group.Key}' occurs {group.Count()} times; " +
                              "template tags must be unique before this slot can be bound.",
                    Path = group.Key
                });
        }

        slots.Sort((left, right) => string.CompareOrdinal(left.Name, right.Name));

        return new TemplateDiscoveryResult
        {
            Slots = slots,
            MediaSlots = MediaSlots(inspect),
            RepeatingRows = inspect.Format == DocFormat.Word
                ? RepeatingCandidates(inspect)
                : Array.Empty<RepeatingRowCandidate>(),
            Diagnostics = diagnostics,
            TemplateSha256 = Sha256(bytes)
        };
    }

    /// <summary>Discovers a template by opaque connection and document ids.</summary>
    public Task<TemplateDiscoveryResult> DiscoverTemplateAsync(
        string connectionId,
        string documentId,
        CancellationToken cancellationToken = default) =>
        DiscoverTemplateAsync(ReferenceFor(connectionId, documentId), cancellationToken);

    /// <summary>
    /// Validates an entire batch without writing anything, and returns a token binding
    /// that validation to the exact template bytes and the exact batch.
    /// </summary>
    /// <remarks>
    /// Every item is validated even when earlier ones fail, so one call reports all the
    /// problems rather than the first. Per-item diagnostics are capped by the effective
    /// budget and say so when they are truncated.
    /// </remarks>
    public async Task<TemplateBatchPreview> PreviewTemplateBatchAsync(
        DocumentReference template,
        TemplateBatchRequest request,
        CancellationToken cancellationToken = default)
    {
        if (template is null) throw new ArgumentNullException(nameof(template));
        if (request is null) throw new ArgumentNullException(nameof(request));

        var limits = EffectiveLimits(request);
        var batchDiagnostics = new List<WorkflowDiagnostic>();

        byte[] bytes;
        using (var content = await OpenReadAsync(template, cancellationToken).ConfigureAwait(false))
        using (var buffer = new MemoryStream())
        {
            await content.Stream.CopyToAsync(buffer, 81920, cancellationToken).ConfigureAwait(false);
            bytes = buffer.ToArray();
        }

        var inspection = await InspectAsync(
            new StreamHandle(new MemoryStream(bytes, writable: false)),
            InspectOptions.Default,
            cancellationToken).ConfigureAwait(false);

        if (request.Items.Count > limits.MaximumDocuments)
            batchDiagnostics.Add(new WorkflowDiagnostic
            {
                Code = "batch-too-large",
                Message = $"The batch requests {request.Items.Count} outputs but the effective limit is " +
                          $"{limits.MaximumDocuments}. The effective limit is the stricter of the host budget " +
                          "and the request, so raising MaximumDocuments in the request cannot raise it.",
                Path = "items"
            });

        foreach (var duplicate in request.Items
                     .GroupBy(item => item.OutputName, StringComparer.OrdinalIgnoreCase)
                     .Where(group => group.Count() > 1))
            batchDiagnostics.Add(new WorkflowDiagnostic
            {
                Code = "duplicate-output-name",
                Message = $"Output name '{duplicate.Key}' is requested {duplicate.Count()} times. " +
                          "Names must be unique; nothing is overwritten.",
                Path = duplicate.Key
            });

        var items = new List<TemplateBatchPreviewItem>();
        var mediaPayloads = new List<string>();
        long totalRows = 0;
        long totalImageBytes = 0;

        foreach (var item in request.Items)
        {
            cancellationToken.ThrowIfCancellationRequested();

            var diagnostics = new List<WorkflowDiagnostic>();
            if (string.IsNullOrWhiteSpace(item.OutputName))
                diagnostics.Add(new WorkflowDiagnostic
                {
                    Code = "missing-output-name",
                    Message = "Each batch item needs an output name.",
                    Path = null
                });

            int rows = item.Binding.RepeatingTables.Sum(table => table.Records.Count);
            totalRows += rows;

            if (rows > limits.MaximumRowsPerDocument)
                diagnostics.Add(new WorkflowDiagnostic
                {
                    Code = "too-many-rows",
                    Message = $"This output would create {rows} repeating rows; the effective limit is " +
                              $"{limits.MaximumRowsPerDocument}.",
                    Path = item.OutputName
                });

            if (item.Binding.Values.Count > limits.MaximumFieldsPerDocument)
                diagnostics.Add(new WorkflowDiagnostic
                {
                    Code = "too-many-fields",
                    Message = $"This output binds {item.Binding.Values.Count} scalar fields; the effective " +
                              $"limit is {limits.MaximumFieldsPerDocument}.",
                    Path = item.OutputName
                });

            foreach (var oversized in item.Binding.Values
                         .Where(entry => (entry.Value?.Length ?? 0) > limits.MaximumValueLength))
                diagnostics.Add(new WorkflowDiagnostic
                {
                    Code = "value-too-long",
                    Message = $"The value bound to '{oversized.Key}' is {oversized.Value!.Length} characters; " +
                              $"the effective limit is {limits.MaximumValueLength}.",
                    Path = oversized.Key
                });

            // Validated against the one inspection taken above, so the whole batch is
            // checked against the exact bytes the token hashes.
            var plan = BuildTemplatePlanCore(inspection, item.Binding);
            diagnostics.AddRange(plan.Diagnostics);

            // Media is resolved here, not at commit: the exact bytes have to be in hand
            // before the token can bind to them.
            var media = await ResolveMediaAsync(inspection, item.Binding, limits.Media, cancellationToken)
                .ConfigureAwait(false);
            diagnostics.AddRange(media.Diagnostics);
            mediaPayloads.AddRange(media.Payloads.Select(payload => item.OutputName + "|" + payload));
            totalImageBytes += media.ImageBytes;

            bool truncated = diagnostics.Count > limits.MaximumDiagnosticsPerItem;
            items.Add(new TemplateBatchPreviewItem
            {
                OutputName = item.OutputName,
                IsValid = diagnostics.Count == 0 && plan.Plan is not null,
                OperationCount = (plan.Plan?.Operations.Count ?? 0) + media.Operations.Count,
                RowCount = rows,
                ImageCount = media.ImageCount,
                ImageBytes = media.ImageBytes,
                Diagnostics = truncated
                    ? diagnostics.Take(limits.MaximumDiagnosticsPerItem).ToList()
                    : diagnostics,
                DiagnosticsTruncated = truncated
            });
        }

        if (totalImageBytes > limits.Media.MaximumTotalImageBytes)
            batchDiagnostics.Add(new WorkflowDiagnostic
            {
                Code = "too-many-total-image-bytes",
                Message = $"The batch carries {totalImageBytes} image bytes across all outputs; the " +
                          $"effective limit is {limits.Media.MaximumTotalImageBytes}.",
                Path = "items"
            });

        if (totalRows > limits.MaximumTotalRows)
            batchDiagnostics.Add(new WorkflowDiagnostic
            {
                Code = "too-many-total-rows",
                Message = $"The batch would create {totalRows} repeating rows across all outputs; the " +
                          $"effective limit is {limits.MaximumTotalRows}.",
                Path = "items"
            });

        return new TemplateBatchPreview
        {
            IsValid = batchDiagnostics.Count == 0 && items.Count > 0 && items.All(item => item.IsValid),
            Items = items,
            Diagnostics = batchDiagnostics,
            Limits = limits,
            Token = new TemplateBatchToken
            {
                TemplateSha256 = Sha256(bytes),
                BatchSha256 = Sha256(Encoding.UTF8.GetBytes(Normalize(request))),
                MediaSha256 = Sha256(Encoding.UTF8.GetBytes(string.Join("\n", mediaPayloads)))
            }
        };
    }

    /// <summary>Previews a template batch by opaque connection and document ids.</summary>
    public Task<TemplateBatchPreview> PreviewTemplateBatchAsync(
        string connectionId,
        string documentId,
        TemplateBatchRequest request,
        CancellationToken cancellationToken = default) =>
        PreviewTemplateBatchAsync(ReferenceFor(connectionId, documentId), request, cancellationToken);

    /// <summary>
    /// Builds a template plan including typed media, resolving images and chart data
    /// under the effective budgets.
    /// </summary>
    /// <remarks>
    /// The commit path uses this rather than the scalar-only builder, so the operations
    /// it applies are the same ones preflight validated. Media is resolved again here
    /// rather than carried over from the preview, because a preview is not an
    /// authorization to write and its bytes are not held anywhere in between; the token
    /// is what proves nothing changed.
    /// </remarks>
    internal async Task<TemplatePlanResult> BuildTemplatePlanWithMediaAsync(
        DocumentReference template,
        TemplateBinding binding,
        TemplateBatchLimits limits,
        CancellationToken cancellationToken)
    {
        var inspection = await InspectAsync(template, InspectOptions.Default, cancellationToken)
            .ConfigureAwait(false);

        var scalar = BuildTemplatePlanCore(inspection, binding);
        var media = await ResolveMediaAsync(inspection, binding, limits.Media, cancellationToken)
            .ConfigureAwait(false);

        var diagnostics = scalar.Diagnostics.Concat(media.Diagnostics).ToList();
        if (diagnostics.Count > 0)
            return new TemplatePlanResult { Diagnostics = diagnostics, Plan = null };

        var operations = (scalar.Plan?.Operations ?? Array.Empty<PlanOperation>())
            .Concat(media.Operations)
            .ToList();

        return new TemplatePlanResult
        {
            Diagnostics = Array.Empty<WorkflowDiagnostic>(),
            Plan = new DocumentPlan
            {
                Format = inspection.Format,
                Snapshot = inspection.Snapshot,
                Revision = binding.Revision,
                Operations = operations
            }
        };
    }

    /// <summary>
    /// The effective budgets for a request: the stricter of the host's and the request's
    /// own, including the legacy <see cref="TemplateBatchRequest.MaximumDocuments"/>.
    /// </summary>
    public TemplateBatchLimits EffectiveLimits(TemplateBatchRequest request)
    {
        var requested = request.Limits;

        // The pre-existing MaximumDocuments is honoured as a request, never as a ceiling.
        if (request.MaximumDocuments > 0)
        {
            requested = (requested ?? TemplateBatchLimits.Default) is var baseline && baseline is not null
                ? new TemplateBatchLimits
                {
                    MaximumDocuments = Math.Min(baseline.MaximumDocuments, request.MaximumDocuments),
                    MaximumRowsPerDocument = baseline.MaximumRowsPerDocument,
                    MaximumFieldsPerDocument = baseline.MaximumFieldsPerDocument,
                    MaximumTotalRows = baseline.MaximumTotalRows,
                    MaximumValueLength = baseline.MaximumValueLength,
                    MaximumDiagnosticsPerItem = baseline.MaximumDiagnosticsPerItem
                }
                : requested;
        }

        return _templateLimits.Restrict(requested);
    }

    /// <summary>
    /// Serializes a batch deterministically so its hash changes when the intent changes.
    /// </summary>
    /// <remarks>
    /// Item order is preserved rather than sorted: reordering outputs is a different
    /// batch, and a preview of one order must not authorize another. Within an item,
    /// dictionaries are ordinal-sorted so an unordered map does not produce two hashes
    /// for the same intent.
    /// </remarks>
    internal static string Normalize(TemplateBatchRequest request)
    {
        var builder = new StringBuilder();
        builder.Append("continueOnError=").Append(request.ContinueOnError ? '1' : '0').Append('\n');

        foreach (var item in request.Items)
        {
            builder.Append("item\n");
            builder.Append("  name=").Append(item.OutputName).Append('\n');
            builder.Append("  mode=").Append(item.Binding.Mode).Append('\n');
            builder.Append("  missing=").Append(item.Binding.MissingValueBehavior).Append('\n');
            builder.Append("  rejectUnknown=").Append(item.Binding.RejectUnknownValues ? '1' : '0').Append('\n');
            builder.Append("  revision=")
                .Append(item.Binding.Revision?.Author ?? string.Empty).Append('|')
                .Append(item.Binding.Revision?.TimestampUtc?.UtcDateTime
                    .ToString("O", CultureInfo.InvariantCulture) ?? string.Empty)
                .Append('\n');

            foreach (var value in item.Binding.Values.OrderBy(entry => entry.Key, StringComparer.Ordinal))
                builder.Append("  value=").Append(value.Key).Append('=')
                    .Append(value.Value ?? "\u0000null").Append('\n');

            // The typed binding as written. What a provider reference resolves to is
            // covered separately by the token's media hash.
            foreach (var typed in item.Binding.TypedValues.OrderBy(entry => entry.Key, StringComparer.Ordinal))
                builder.Append("  typed=").Append(typed.Key).Append('=')
                    .Append(Describe(typed.Value)).Append('\n');

            foreach (var table in item.Binding.RepeatingTables)
            {
                builder.Append("  table=").Append(table.TablePath)
                    .Append('#').Append(table.TemplateRowIndex).Append('\n');
                foreach (var record in table.Records)
                {
                    builder.Append("    record");
                    foreach (var field in record.OrderBy(entry => entry.Key, StringComparer.Ordinal))
                        builder.Append('|').Append(field.Key).Append('=').Append(field.Value ?? "\u0000null");
                    builder.Append('\n');
                }
            }
        }

        return builder.ToString();
    }

    /// <summary>A stable description of a typed binding as the caller wrote it.</summary>
    private static string Describe(TemplateValue value) => value switch
    {
        TemplateTextValue text => "text:" + (text.Text ?? "\u0000null"),
        TemplateImageValue image =>
            $"image:{image.ImageType}:{image.WidthPx}x{image.HeightPx}:{image.AltText}:" +
            $"{image.ImageConnectionId}/{image.ImageDocumentId}:" +
            (image.Base64Bytes is null ? string.Empty : Sha256(Convert.FromBase64String(SafeBase64(image.Base64Bytes)))),
        TemplateChartValue chart =>
            $"chart:{chart.Kind}:{chart.Title}:{chart.ShowLegend}:" +
            string.Join(",", chart.Categories) + ":" +
            string.Join(";", chart.Series.Select(series => series.Name + "=" + string.Join(",", series.Values))),
        _ => value.GetType().Name
    };

    /// <summary>
    /// Normalization must not throw on malformed input; preflight reports that as a
    /// diagnostic instead, so the hash simply covers the string as written.
    /// </summary>
    private static string SafeBase64(string candidate)
    {
        try
        {
            Convert.FromBase64String(candidate);
            return candidate;
        }
        catch (FormatException)
        {
            return Convert.ToBase64String(Encoding.UTF8.GetBytes(candidate));
        }
    }


    /// <summary>
    /// Slots that hold something other than text: paragraphs an image can be placed
    /// against, and native charts already in the deck.
    /// </summary>
    /// <remarks>
    /// Word has no native chart handler in this engine, so a Word template never reports
    /// a chart slot and a chart binding against one is refused rather than ignored.
    /// PowerPoint charts are reported only where one already exists, because the binding
    /// updates a chart rather than creating one.
    /// </remarks>
    private static IReadOnlyList<TemplateMediaSlot> MediaSlots(InspectResult inspect)
    {
        var slots = new List<TemplateMediaSlot>();

        foreach (var chart in inspect.Nodes.Where(node => node.Kind == "chart"))
            slots.Add(new TemplateMediaSlot
            {
                Name = chart.Path,
                MediaKind = "chart",
                Path = chart.Path,
                InRepeatingRow = false
            });

        // An image is placed relative to a paragraph, so a bindable image slot is a
        // content control whose paragraph can be anchored. A control inside a table is
        // flagged, because repeating that row would repeat one image across every row.
        foreach (var anchor in inspect.StructuralAnchors
                     .Where(anchor => anchor.Kind is "contentControl" or "shapeName"))
        {
            var owner = inspect.Paragraphs.FirstOrDefault(paragraph =>
                paragraph.Text.IndexOf(anchor.Tag, StringComparison.Ordinal) >= 0);

            slots.Add(new TemplateMediaSlot
            {
                Name = anchor.Tag,
                MediaKind = "image",
                Path = anchor.Tag,
                InRepeatingRow = owner?.In is { Length: > 0 }
            });
        }

        slots.Sort((left, right) =>
        {
            int kind = string.CompareOrdinal(left.MediaKind, right.MediaKind);
            return kind != 0 ? kind : string.CompareOrdinal(left.Name, right.Name);
        });
        return slots;
    }

    /// <summary>
    /// Resolves every typed value in one item: checks the slot can hold what is aimed at
    /// it, fetches provider-held image bytes, and reports the exact payload so the
    /// preview token can bind to it.
    /// </summary>
    private async Task<ResolvedMedia> ResolveMediaAsync(
        InspectResult inspection,
        TemplateBinding binding,
        TemplateMediaLimits limits,
        CancellationToken cancellationToken)
    {
        var diagnostics = new List<WorkflowDiagnostic>();
        var payloads = new List<string>();
        var operations = new List<PlanOperation>();
        long imageBytes = 0;
        int images = 0;

        var mediaSlots = MediaSlots(inspection).ToDictionary(slot => slot.Name, slot => slot, StringComparer.Ordinal);
        var textSlots = inspection.StructuralAnchors
            .Where(anchor => anchor.Kind is "contentControl" or "shapeName")
            .ToDictionary(anchor => anchor.Tag, anchor => anchor, StringComparer.Ordinal);

        foreach (var entry in binding.TypedValues.OrderBy(entry => entry.Key, StringComparer.Ordinal))
        {
            cancellationToken.ThrowIfCancellationRequested();
            var name = entry.Key;

            if (binding.Values.ContainsKey(name))
            {
                diagnostics.Add(Diagnostic("duplicate-template-binding",
                    $"Slot '{name}' is bound both as a scalar value and as a typed value. " +
                    "Bind it once.", name));
                continue;
            }

            switch (entry.Value)
            {
                case TemplateTextValue text:
                    if (!textSlots.TryGetValue(name, out var textAnchor))
                    {
                        diagnostics.Add(Diagnostic("unknown-template-value",
                            $"No template slot is tagged '{name}'.", name));
                        break;
                    }

                    operations.Add(new FillOp
                    {
                        Target = new StructuralAnchor
                        {
                            Id = textAnchor.Id, Kind = textAnchor.Kind, Tag = name
                        },
                        Value = text.Text ?? string.Empty,
                        Mode = binding.Mode
                    });
                    payloads.Add($"text|{name}|{text.Text ?? "\u0000null"}");
                    break;

                case TemplateImageValue image:
                {
                    if (!mediaSlots.TryGetValue(name, out var slot) || slot.MediaKind != "image")
                    {
                        diagnostics.Add(Diagnostic("wrong-slot-kind",
                            $"Slot '{name}' cannot hold an image. Bind an image only to a slot " +
                            "reported as an image slot by template discovery.", name));
                        break;
                    }

                    if (slot.InRepeatingRow)
                    {
                        diagnostics.Add(Diagnostic("media-in-repeating-row",
                            $"Slot '{name}' is inside a table row. Images in repeating rows are not " +
                            "supported, because every generated row would share one image. Move the " +
                            "image slot outside the table.", name));
                        break;
                    }

                    if (string.IsNullOrWhiteSpace(image.AltText))
                    {
                        diagnostics.Add(Diagnostic("missing-alt-text",
                            $"The image bound to '{name}' has no alt text. A template generates many " +
                            "documents, so a missing description is multiplied across all of them.",
                            name));
                        break;
                    }

                    var resolved = await ResolveImageBytesAsync(name, image, cancellationToken)
                        .ConfigureAwait(false);
                    if (resolved.Diagnostic is { } failure)
                    {
                        diagnostics.Add(failure);
                        break;
                    }

                    var bytes = resolved.Bytes!;
                    if (bytes.LongLength > limits.MaximumImageBytes)
                    {
                        diagnostics.Add(Diagnostic("image-too-large",
                            $"The image bound to '{name}' is {bytes.LongLength} bytes; the effective " +
                            $"limit is {limits.MaximumImageBytes}.", name));
                        break;
                    }

                    if (!SignatureMatches(bytes, image.ImageType))
                    {
                        diagnostics.Add(Diagnostic("image-type-mismatch",
                            $"The bytes bound to '{name}' are not a {image.ImageType} image. The declared " +
                            "type must match the actual content.", name));
                        break;
                    }

                    images++;
                    imageBytes += bytes.LongLength;

                    var host = inspection.Paragraphs.FirstOrDefault(paragraph =>
                        paragraph.Text.IndexOf(name, StringComparison.Ordinal) >= 0)
                        ?? inspection.Paragraphs.FirstOrDefault();
                    if (host is null)
                    {
                        diagnostics.Add(Diagnostic("template-anchor-not-found",
                            $"No paragraph could host the image bound to '{name}'.", name));
                        break;
                    }

                    operations.Add(new InsertImageOp
                    {
                        Target = new TextSpanAnchor { ParaId = host.ParaId, Expect = host.Text },
                        Base64Bytes = Convert.ToBase64String(bytes),
                        ImageType = image.ImageType,
                        WidthPx = image.WidthPx,
                        HeightPx = image.HeightPx,
                        AltText = image.AltText,
                        Position = InsertPosition.After,
                        Mode = binding.Mode
                    });
                    payloads.Add($"image|{name}|{Sha256(bytes)}|{image.ImageType}|{image.WidthPx}x{image.HeightPx}");
                    break;
                }

                case TemplateChartValue chart:
                {
                    if (inspection.Format != DocFormat.PowerPoint)
                    {
                        diagnostics.Add(Diagnostic("unsupported-template-feature",
                            $"Chart bindings are PowerPoint only. This engine has no native Word chart " +
                            $"support, so '{name}' cannot be bound here.", name));
                        break;
                    }

                    if (!mediaSlots.TryGetValue(name, out var slot) || slot.MediaKind != "chart")
                    {
                        diagnostics.Add(Diagnostic("wrong-slot-kind",
                            $"No native chart is at '{name}'. A chart binding updates a chart that is " +
                            "already in the template; it does not create one.", name));
                        break;
                    }

                    int points = chart.Series.Sum(series => series.Values.Count);
                    if (points > limits.MaximumChartPoints)
                    {
                        diagnostics.Add(Diagnostic("chart-too-large",
                            $"The chart bound to '{name}' carries {points} data points; the effective " +
                            $"limit is {limits.MaximumChartPoints}.", name));
                        break;
                    }

                    operations.Add(new UpdateChartOp
                    {
                        Target = new NodeAnchor { Kind = "chart", Path = slot.Path },
                        Kind = chart.Kind,
                        Categories = chart.Categories,
                        Series = chart.Series,
                        Title = chart.Title,
                        ShowLegend = chart.ShowLegend
                    });
                    payloads.Add($"chart|{name}|{chart.Kind}|{string.Join(",", chart.Categories)}|" +
                                 string.Join(";", chart.Series.Select(series =>
                                     series.Name + "=" + string.Join(",", series.Values))));
                    break;
                }

                default:
                    diagnostics.Add(Diagnostic("unsupported-template-feature",
                        $"The value bound to '{name}' is of an unsupported kind.", name));
                    break;
            }
        }

        if (images > limits.MaximumImagesPerDocument)
            diagnostics.Add(Diagnostic("too-many-images",
                $"This output binds {images} images; the effective limit is " +
                $"{limits.MaximumImagesPerDocument}.", null));

        return new ResolvedMedia(diagnostics, operations, payloads, images, imageBytes);
    }

    private async Task<(byte[]? Bytes, WorkflowDiagnostic? Diagnostic)> ResolveImageBytesAsync(
        string name, TemplateImageValue image, CancellationToken cancellationToken)
    {
        bool hasInline = !string.IsNullOrWhiteSpace(image.Base64Bytes);
        bool hasReference = !string.IsNullOrWhiteSpace(image.ImageDocumentId);

        if (hasInline == hasReference)
            return (null, Diagnostic("invalid-image-binding",
                $"The image bound to '{name}' needs exactly one of inline bytes or a provider " +
                "document id. There is no URL form: this engine does not fetch images.", name));

        if (hasInline)
        {
            try
            {
                return (Convert.FromBase64String(image.Base64Bytes!), null);
            }
            catch (FormatException)
            {
                return (null, Diagnostic("invalid-image-binding",
                    $"The inline bytes bound to '{name}' are not valid base64.", name));
            }
        }

        if (string.IsNullOrWhiteSpace(image.ImageConnectionId))
            return (null, Diagnostic("invalid-image-binding",
                $"The image bound to '{name}' names a document id without a connection id.", name));

        try
        {
            // Read through the provider, so its access rules and size limits apply to the
            // image exactly as they would to any other document the caller asked for.
            using var content = await OpenReadAsync(
                image.ImageConnectionId!, image.ImageDocumentId!, cancellationToken).ConfigureAwait(false);
            using var buffer = new MemoryStream();
            await content.Stream.CopyToAsync(buffer, 81920, cancellationToken).ConfigureAwait(false);
            return (buffer.ToArray(), null);
        }
        catch (DocumentProviderException ex)
        {
            return (null, Diagnostic("image-not-available",
                $"The image bound to '{name}' could not be read: {ex.Message}", name));
        }
    }

    /// <summary>
    /// Checks the declared type against the actual magic bytes, so a caller cannot label
    /// arbitrary content as an image and have it embedded under that name.
    /// </summary>
    private static bool SignatureMatches(byte[] bytes, string declaredType)
    {
        if (bytes.Length < 4) return false;

        return declaredType.ToLowerInvariant() switch
        {
            "png" => bytes[0] == 0x89 && bytes[1] == 0x50 && bytes[2] == 0x4E && bytes[3] == 0x47,
            "jpeg" or "jpg" => bytes[0] == 0xFF && bytes[1] == 0xD8 && bytes[2] == 0xFF,
            "gif" => bytes[0] == 0x47 && bytes[1] == 0x49 && bytes[2] == 0x46,
            "bmp" => bytes[0] == 0x42 && bytes[1] == 0x4D,
            "tiff" => (bytes[0] == 0x49 && bytes[1] == 0x49) || (bytes[0] == 0x4D && bytes[1] == 0x4D),
            _ => false
        };
    }

    private sealed record ResolvedMedia(
        IReadOnlyList<WorkflowDiagnostic> Diagnostics,
        IReadOnlyList<PlanOperation> Operations,
        IReadOnlyList<string> Payloads,
        int ImageCount,
        long ImageBytes);

    /// <summary>
    /// Finds rows that carry <c>{{Field}}</c> placeholders, which is the convention this
    /// engine's repeating-row support uses.
    /// </summary>
    private static IReadOnlyList<RepeatingRowCandidate> RepeatingCandidates(InspectResult inspect)
    {
        var candidates = new List<RepeatingRowCandidate>();

        foreach (var table in inspect.Nodes.Where(node => node.Kind == "table"))
        {
            var rows = inspect.Paragraphs
                .Where(paragraph => string.Equals(paragraph.In, table.Path, StringComparison.Ordinal))
                .ToList();
            if (rows.Count == 0) continue;

            // Paragraph rows are not indexed by table row here, so the candidate reports
            // the placeholder fields found anywhere in the table and the conventional
            // template row index of 1: a header row followed by one template row.
            var fields = rows
                .SelectMany(paragraph => PlaceholderPattern.Matches(paragraph.Text).Cast<Match>())
                .Select(match => match.Groups["name"].Value)
                .Distinct(StringComparer.Ordinal)
                .OrderBy(name => name, StringComparer.Ordinal)
                .ToList();

            if (fields.Count == 0) continue;

            candidates.Add(new RepeatingRowCandidate
            {
                TablePath = table.Path,
                TemplateRowIndex = 1,
                Fields = fields,
                Confirmed = false,
                TableRowCount = rows.Count
            });
        }

        return candidates;
    }
}
