using System.Globalization;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using OfficeAgent.Abstractions;
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
        long totalRows = 0;

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

            bool truncated = diagnostics.Count > limits.MaximumDiagnosticsPerItem;
            items.Add(new TemplateBatchPreviewItem
            {
                OutputName = item.OutputName,
                IsValid = diagnostics.Count == 0 && plan.Plan is not null,
                OperationCount = plan.Plan?.Operations.Count ?? 0,
                RowCount = rows,
                Diagnostics = truncated
                    ? diagnostics.Take(limits.MaximumDiagnosticsPerItem).ToList()
                    : diagnostics,
                DiagnosticsTruncated = truncated
            });
        }

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
                BatchSha256 = Sha256(Encoding.UTF8.GetBytes(Normalize(request)))
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
                    .Append(value.Value ?? " null").Append('\n');

            foreach (var table in item.Binding.RepeatingTables)
            {
                builder.Append("  table=").Append(table.TablePath)
                    .Append('#').Append(table.TemplateRowIndex).Append('\n');
                foreach (var record in table.Records)
                {
                    builder.Append("    record");
                    foreach (var field in record.OrderBy(entry => entry.Key, StringComparer.Ordinal))
                        builder.Append('|').Append(field.Key).Append('=').Append(field.Value ?? " null");
                    builder.Append('\n');
                }
            }
        }

        return builder.ToString();
    }

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
