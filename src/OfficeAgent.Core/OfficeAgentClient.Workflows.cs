using System.Security.Cryptography;
using System.IO.Compression;
using System.Text;
using System.Xml;
using System.Xml.Linq;
using OfficeAgent.Abstractions;
using OfficeAgent.Core.DocumentProviders;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Core;

public sealed partial class OfficeAgentClient
{
    /// <summary>Resolves scalar and repeating template bindings into a snapshot-bound plan.</summary>
    public async Task<TemplatePlanResult> BuildTemplatePlanAsync(
        DocumentReference template,
        TemplateBinding binding,
        CancellationToken cancellationToken = default)
    {
        if (template is null) throw new ArgumentNullException(nameof(template));
        if (binding is null) throw new ArgumentNullException(nameof(binding));

        var inspect = await InspectAsync(template, InspectOptions.Default, cancellationToken).ConfigureAwait(false);
        return BuildTemplatePlanCore(inspect, binding);
    }

    /// <summary>
    /// Resolves a binding against an inspection already in hand. Preflight uses this so a
    /// whole batch is validated against one reading of the template rather than one per
    /// item, which is also what makes the preview's template hash meaningful.
    /// </summary>
    internal static TemplatePlanResult BuildTemplatePlanCore(InspectResult inspect, TemplateBinding binding)
    {
        var diagnostics = new List<WorkflowDiagnostic>();
        var operations = new List<PlanOperation>();
        var slots = inspect.StructuralAnchors
            .Where(anchor => anchor.Kind is "contentControl" or "shapeName")
            .GroupBy(anchor => anchor.Tag, StringComparer.Ordinal)
            .ToDictionary(group => group.Key, group => group.ToList(), StringComparer.Ordinal);

        foreach (var slot in slots.Where(slot => slot.Value.Count != 1))
            diagnostics.Add(Diagnostic("ambiguous-template-slot",
                $"Template slot '{slot.Key}' occurs {slot.Value.Count} times; template tags must be unique.", slot.Key));

        foreach (var entry in binding.Values)
        {
            var name = entry.Key;
            var value = entry.Value;
            if (!slots.TryGetValue(name, out var matches))
            {
                if (binding.RejectUnknownValues)
                    diagnostics.Add(Diagnostic("unknown-template-value", $"No template slot is tagged '{name}'.", name));
                continue;
            }
            if (matches.Count != 1)
            {
                continue;
            }
            operations.Add(new FillOp
            {
                Target = new StructuralAnchor { Id = matches[0].Id, Kind = matches[0].Kind, Tag = name },
                Value = value ?? string.Empty,
                Mode = binding.Mode
            });
        }

        foreach (var entry in slots)
        {
            var name = entry.Key;
            var matches = entry.Value;
            if (matches.Count != 1) continue;
            if (binding.Values.ContainsKey(name)) continue;
            if (binding.MissingValueBehavior == MissingTemplateValueBehavior.Fail)
                diagnostics.Add(Diagnostic("missing-template-value", $"No value was supplied for template slot '{name}'.", name));
            else if (binding.MissingValueBehavior == MissingTemplateValueBehavior.Empty && matches.Count == 1)
                operations.Add(new FillOp
                {
                    Target = new StructuralAnchor { Id = matches[0].Id, Kind = matches[0].Kind, Tag = name },
                    Value = string.Empty,
                    Mode = binding.Mode
                });
        }

        if (binding.RepeatingTables.Count > 0 && inspect.Format != DocFormat.Word)
            diagnostics.Add(Diagnostic("unsupported-template-feature",
                "Repeating table rows are currently supported only for Word templates.", "repeatingTables"));

        var tablePaths = new HashSet<string>(inspect.Nodes
            .Where(node => node.Kind == "table")
            .Select(node => node.Path), StringComparer.Ordinal);
        foreach (var table in binding.RepeatingTables)
        {
            if (string.IsNullOrWhiteSpace(table.TablePath) || !tablePaths.Contains(table.TablePath))
            {
                diagnostics.Add(Diagnostic("template-table-not-found",
                    $"No inspected table has path '{table.TablePath}'.", table.TablePath));
                continue;
            }
            operations.Add(new RepeatTableRowOp
            {
                Target = new NodeAnchor { Kind = "table", Path = table.TablePath },
                TemplateRowIndex = table.TemplateRowIndex,
                Records = table.Records,
                MissingValueBehavior = binding.MissingValueBehavior,
                RejectUnknownValues = binding.RejectUnknownValues,
                Mode = binding.Mode
            });
        }

        return new TemplatePlanResult
        {
            Diagnostics = diagnostics,
            Plan = diagnostics.Count == 0
                ? new DocumentPlan
                {
                    Format = inspect.Format,
                    Snapshot = inspect.Snapshot,
                    Revision = binding.Revision,
                    Operations = operations
                }
                : null
        };
    }

    /// <summary>Builds a template plan by opaque connection and document ids.</summary>
    public Task<TemplatePlanResult> BuildTemplatePlanAsync(
        string connectionId,
        string documentId,
        TemplateBinding binding,
        CancellationToken cancellationToken = default) =>
        BuildTemplatePlanAsync(ReferenceFor(connectionId, documentId), binding, cancellationToken);

    /// <summary>
    /// Produces separate provider documents from one template. Each output is validated
    /// and saved atomically and has its own result and audit receipt.
    /// </summary>
    public Task<TemplateBatchResult> PopulateTemplateBatchAsync(
        DocumentReference template,
        TemplateBatchRequest request,
        CancellationToken cancellationToken = default) =>
        PopulateTemplateBatchAsync(template, request, expectedToken: null, cancellationToken);

    /// <summary>
    /// Populates a template batch, refusing the commit when the template or the batch has
    /// changed since the preview issued the supplied token.
    /// </summary>
    public async Task<TemplateBatchResult> PopulateTemplateBatchAsync(
        DocumentReference template,
        TemplateBatchRequest request,
        TemplateBatchToken? expectedToken,
        CancellationToken cancellationToken = default)
    {
        if (template is null) throw new ArgumentNullException(nameof(template));
        if (request is null) throw new ArgumentNullException(nameof(request));
        if (request.MaximumDocuments <= 0)
            throw new ArgumentOutOfRangeException(nameof(request), "MaximumDocuments must be positive.");

        // Every batch is preflighted, asked for or not. A batch that cannot validate is
        // refused before any output exists rather than leaving a prefix of it in storage.
        var limits = EffectiveLimits(request);
        var preview = await PreviewTemplateBatchAsync(template, request, cancellationToken)
            .ConfigureAwait(false);

        if (expectedToken is not null && !Matches(expectedToken, preview.Token))
            return Refused(request,
                "stale-batch-preview",
                "The template or the batch has changed since it was previewed, so this commit would " +
                "apply intent nobody reviewed. Preview again and commit the new token.");

        if (preview.Diagnostics.Count > 0)
            return new TemplateBatchResult
            {
                Items = request.Items.Select(item => new TemplateBatchItemResult
                {
                    OutputName = item.OutputName,
                    Outcome = TemplateItemOutcome.Failed,
                    Diagnostics = preview.Diagnostics
                }).ToArray()
            };

        var results = new List<TemplateBatchItemResult>();
        bool stopped = false;
        foreach (var item in request.Items)
        {
            if (stopped)
            {
                results.Add(new TemplateBatchItemResult
                {
                    OutputName = item.OutputName,
                    Outcome = TemplateItemOutcome.Skipped,
                    Diagnostics = new[]
                    {
                        Diagnostic("item-skipped",
                            "An earlier item failed and this batch stops on error, so this output was " +
                            "never attempted and does not exist.", item.OutputName)
                    }
                });
                continue;
            }

            cancellationToken.ThrowIfCancellationRequested();
            var plan = await BuildTemplatePlanWithMediaAsync(template, item.Binding, limits, cancellationToken)
                .ConfigureAwait(false);
            if (!plan.IsValid || plan.Plan is null)
            {
                results.Add(new TemplateBatchItemResult
                {
                    OutputName = item.OutputName,
                    Outcome = TemplateItemOutcome.Failed,
                    Diagnostics = plan.Diagnostics
                });
                if (!request.ContinueOnError) stopped = true;
                continue;
            }

            try
            {
                var applied = await CommitAsync(template, plan.Plan, new SaveDocumentOptions
                {
                    Mode = SaveMode.NewDocument,
                    NewName = item.OutputName
                }, cancellationToken).ConfigureAwait(false);
                results.Add(new TemplateBatchItemResult
                {
                    OutputName = item.OutputName,
                    Committed = applied.Committed,
                    Outcome = applied.Committed ? TemplateItemOutcome.Committed : TemplateItemOutcome.Failed,
                    Document = applied.Document,
                    Report = applied.Report,
                    Receipt = applied.Receipt,
                    Diagnostics = applied.Report.Errors.Select(error =>
                        Diagnostic(error.Code, error.Message, error.Target?.Id)).ToArray()
                });
                if (!applied.Committed && !request.ContinueOnError) stopped = true;
            }
            catch (DocumentProviderException ex)
            {
                // A provider that already accepted the bytes and then failed leaves an
                // output that may or may not exist. Reporting that as a plain failure
                // would be a guess, and a caller acting on it would either lose the file
                // or create a duplicate.
                bool decidedBeforeAnyWrite = ex.Code is ProviderErrorCode.AlreadyExists
                    or ProviderErrorCode.AccessDenied
                    or ProviderErrorCode.ExtensionNotAllowed
                    or ProviderErrorCode.InvalidArgument
                    or ProviderErrorCode.ContentTooLarge
                    or ProviderErrorCode.ConfigurationError
                    or ProviderErrorCode.NotFound;

                results.Add(new TemplateBatchItemResult
                {
                    OutputName = item.OutputName,
                    Outcome = decidedBeforeAnyWrite
                        ? TemplateItemOutcome.Failed
                        : TemplateItemOutcome.Uncertain,
                    Diagnostics = new[]
                    {
                        Diagnostic(ToKebabCase(ex.Code.ToString()),
                            decidedBeforeAnyWrite
                                ? ex.Message
                                : ex.Message + " Storage may already hold this output; do not retry the " +
                                  "same name blindly. Inspect the destination and reconcile it.",
                            item.OutputName)
                    }
                });
                if (!request.ContinueOnError) stopped = true;
            }
        }
        return new TemplateBatchResult { Items = results };
    }

    private static bool Matches(TemplateBatchToken expected, TemplateBatchToken actual) =>
        string.Equals(expected.TemplateSha256, actual.TemplateSha256, StringComparison.OrdinalIgnoreCase) &&
        string.Equals(expected.BatchSha256, actual.BatchSha256, StringComparison.OrdinalIgnoreCase);

    private static TemplateBatchResult Refused(TemplateBatchRequest request, string code, string message) =>
        new()
        {
            Items = request.Items.Select(item => new TemplateBatchItemResult
            {
                OutputName = item.OutputName,
                Outcome = TemplateItemOutcome.Failed,
                Diagnostics = new[] { Diagnostic(code, message, item.OutputName) }
            }).ToArray()
        };

    /// <summary>Populates a template batch by opaque connection and document ids.</summary>
    public Task<TemplateBatchResult> PopulateTemplateBatchAsync(
        string connectionId,
        string documentId,
        TemplateBatchRequest request,
        CancellationToken cancellationToken = default) =>
        PopulateTemplateBatchAsync(ReferenceFor(connectionId, documentId), request, cancellationToken);

    /// <summary>
    /// Populates a template batch by opaque ids, refusing the commit when the template or
    /// the batch has changed since the preview issued <paramref name="expectedToken"/>.
    /// </summary>
    /// <remarks>
    /// Without this overload the hash-bound commit was reachable only by callers holding a
    /// <see cref="DocumentReference"/>, which excluded both adapters: an agent could commit
    /// a batch but could not bind that commit to the preview it had reviewed.
    /// </remarks>
    public Task<TemplateBatchResult> PopulateTemplateBatchAsync(
        string connectionId,
        string documentId,
        TemplateBatchRequest request,
        TemplateBatchToken? expectedToken,
        CancellationToken cancellationToken = default) =>
        PopulateTemplateBatchAsync(
            ReferenceFor(connectionId, documentId), request, expectedToken, cancellationToken);

    /// <summary>
    /// Compares two provider-backed Word documents without writing either one. A complete
    /// result carries a tracked, original-snapshot-bound plan suitable for normal preview
    /// and commit against the original document.
    /// </summary>
    public async Task<DocumentComparisonResult> CompareDocumentsAsync(
        DocumentReference original,
        DocumentReference revised,
        DocumentComparisonOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        if (original is null) throw new ArgumentNullException(nameof(original));
        if (revised is null) throw new ArgumentNullException(nameof(revised));
        options ??= new DocumentComparisonOptions();
        if (options.MaximumParagraphs <= 0 || options.MaximumDifferences <= 0 || options.MaximumDocumentBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(options), "Comparison limits must be positive.");

        try
        {
            using var originalContent = await OpenWithTelemetryAsync(original, cancellationToken).ConfigureAwait(false);
            using var revisedContent = await OpenWithTelemetryAsync(revised, cancellationToken).ConfigureAwait(false);
            var originalBytes = await ReadBoundedAsync(originalContent.Stream, options.MaximumDocumentBytes, cancellationToken).ConfigureAwait(false);
            var revisedBytes = await ReadBoundedAsync(revisedContent.Stream, options.MaximumDocumentBytes, cancellationToken).ConfigureAwait(false);
            return CompareDocuments(originalBytes, revisedBytes, options, cancellationToken);
        }
        catch (InvalidDataException)
        {
            return new DocumentComparisonResult
            {
                Diagnostics = new[]
                {
                    Diagnostic("comparison-input-limit-exceeded",
                        $"Each comparison input must be at most {options.MaximumDocumentBytes} bytes.", "document")
                }
            };
        }
    }

    /// <summary>Compares two documents addressed by opaque connection and document ids.</summary>
    public Task<DocumentComparisonResult> CompareDocumentsAsync(
        string originalConnectionId,
        string originalDocumentId,
        string revisedConnectionId,
        string revisedDocumentId,
        DocumentComparisonOptions? options = null,
        CancellationToken cancellationToken = default) =>
        CompareDocumentsAsync(
            ReferenceFor(originalConnectionId, originalDocumentId),
            ReferenceFor(revisedConnectionId, revisedDocumentId),
            options,
            cancellationToken);

    /// <summary>Compares two in-memory Word packages without writing them.</summary>
    public DocumentComparisonResult CompareDocuments(
        byte[] original,
        byte[] revised,
        DocumentComparisonOptions? options = null,
        CancellationToken cancellationToken = default)
    {
        if (original is null) throw new ArgumentNullException(nameof(original));
        if (revised is null) throw new ArgumentNullException(nameof(revised));
        options ??= new DocumentComparisonOptions();
        if (options.MaximumParagraphs <= 0 || options.MaximumDifferences <= 0 || options.MaximumDocumentBytes <= 0)
            throw new ArgumentOutOfRangeException(nameof(options), "Comparison limits must be positive.");
        var diagnostics = new List<WorkflowDiagnostic>();
        if (original.LongLength > options.MaximumDocumentBytes || revised.LongLength > options.MaximumDocumentBytes)
        {
            diagnostics.Add(Diagnostic("comparison-input-limit-exceeded",
                $"Each comparison input must be at most {options.MaximumDocumentBytes} bytes.", "document"));
            return ComparisonResult(original, revised, diagnostics: diagnostics);
        }

        var before = Inspect(original, InspectOptions.Default);
        var after = Inspect(revised, InspectOptions.Default);
        if (before.Format != DocFormat.Word || after.Format != DocFormat.Word)
        {
            diagnostics.Add(Diagnostic("unsupported-comparison-format",
                "Two-document comparison currently requires two Word .docx packages.", "document"));
            return ComparisonResult(original, revised, diagnostics: diagnostics);
        }

        var beforeBody = before.Paragraphs.Where(IsFreeBodyParagraph).ToList();
        var afterBody = after.Paragraphs.Where(IsFreeBodyParagraph).ToList();
        if (beforeBody.Count == 0 && afterBody.Count > 0)
            diagnostics.Add(Diagnostic("comparison-anchor-unavailable",
                "The original has no free body paragraph that can anchor an inserted redline.", "body"));
        if (beforeBody.Count > options.MaximumParagraphs || afterBody.Count > options.MaximumParagraphs)
            diagnostics.Add(Diagnostic("comparison-paragraph-limit-exceeded",
                $"Each document may contain at most {options.MaximumParagraphs} body paragraphs.", "body"));

        if (before.Nodes.Any(node => node.Kind == "revision") || after.Nodes.Any(node => node.Kind == "revision"))
            diagnostics.Add(Diagnostic("unsupported-existing-revisions",
                "Comparison inputs must have existing tracked revisions accepted or rejected first.", "revisions"));

        if (!SequenceEqual(UnsupportedParagraphs(before), UnsupportedParagraphs(after)))
            diagnostics.Add(Diagnostic("unsupported-non-body-change",
                "Tables, headers, footers, footnotes, or endnotes changed; this comparison covers free body paragraphs only.", "non-body"));
        if (!SequenceEqual(NodeSignature(before, "image"), NodeSignature(after, "image")))
            diagnostics.Add(Diagnostic("unsupported-image-change",
                "Images changed; image comparison is outside this comparison's coverage.", "images"));
        if (!SequenceEqual(UnsupportedNodeSignature(before), UnsupportedNodeSignature(after)))
            diagnostics.Add(Diagnostic("unsupported-node-change",
                "Document nodes outside free body text changed, such as properties, fields, comments, sections, or table structure.", "nodes"));
        // Two different kinds of problem live in this list. A precondition means the diff
        // itself cannot run: there is no anchor, or there are more paragraphs than the
        // ceiling allows, or the inputs still carry revisions. Those stop here, and every
        // area is reported not compared rather than unchanged.
        //
        // An unsupported *area* is different: the body diff runs fine, it just cannot be
        // turned into a plan. Those no longer stop the walk. Withholding the plan is the
        // safety property; withholding the findings as well was only a side effect of
        // stopping early, and it left a caller with nothing to act on by hand.
        var blocking = diagnostics
            .Where(diagnostic => diagnostic.Code is "comparison-anchor-unavailable"
                or "comparison-paragraph-limit-exceeded"
                or "unsupported-existing-revisions")
            .ToList();
        if (blocking.Count > 0)
            return ComparisonResult(original, revised, diagnostics: diagnostics);

        var operations = new List<PlanOperation>();
        var differences = new List<DocumentDifference>();
        var beforeMarkup = FreeBodyParagraphMarkup(original);
        var afterMarkup = FreeBodyParagraphMarkup(revised);
        var matches = LongestCommonSubsequence(beforeBody, afterBody, cancellationToken);
        var oldStart = 0;
        var newStart = 0;
        var totalDifferenceCount = 0;
        foreach (var match in matches.Concat(new[] { (Old: beforeBody.Count, New: afterBody.Count) }))
        {
            totalDifferenceCount += Math.Max(match.Old - oldStart, match.New - newStart);
            var paired = Math.Min(match.Old - oldStart, match.New - newStart);
            for (var i = 0; i < paired; i++)
                CompareParagraphMarkup(beforeMarkup, afterMarkup, oldStart + i, newStart + i,
                    beforeBody[oldStart + i], diagnostics);
            if (match.New - newStart > paired)
            {
                var targetIndex = match.Old < beforeBody.Count ? match.Old : Math.Max(0, match.Old - 1);
                for (var i = newStart + paired; i < match.New; i++)
                {
                    var expected = AddedParagraphMarkup(beforeMarkup[targetIndex], afterBody[i].StyleId);
                    if (!string.Equals(expected, afterMarkup[i], StringComparison.Ordinal))
                        diagnostics.Add(Diagnostic("unsupported-paragraph-markup-change",
                            $"An added body paragraph at revised index {i} contains run structure or direct formatting the comparison plan cannot reproduce.",
                            afterBody[i].ParaId));
                }
            }
            ProcessHunk(beforeBody, afterBody, oldStart, match.Old, newStart, match.New,
                operations, differences, diagnostics, options.MaximumDifferences);
            if (match.Old < beforeBody.Count)
            {
                CompareParagraphMarkup(beforeMarkup, afterMarkup, match.Old, match.New,
                    beforeBody[match.Old], diagnostics);
                if (!string.Equals(beforeBody[match.Old].StyleId, afterBody[match.New].StyleId, StringComparison.Ordinal))
                    diagnostics.Add(Diagnostic("unsupported-style-change",
                        $"Paragraph style changed at original body paragraph {match.Old}.", beforeBody[match.Old].ParaId));
                oldStart = match.Old + 1;
                newStart = match.New + 1;
            }
        }

        if (totalDifferenceCount > options.MaximumDifferences)
            diagnostics.Add(Diagnostic("comparison-difference-limit-exceeded",
                $"The comparison reached the {options.MaximumDifferences}-difference limit.", "body"));

        var bodyDifferenceCount = differences.Count;

        // Cell text, but only once the table's shape is known to be identical. Geometry is
        // checked by the text-stripped table signature below, so an aligned cell pair
        // really is the same cell in both documents rather than two cells that happen to
        // share a position.
        CompareTableCells(original, revised, before, after, operations, differences, diagnostics, options);
        // Which areas outside free body paragraphs differ, named individually. One
        // blanket "something else changed" tells a caller nothing they can act on.
        var areaChanges = CompareAreas(original, revised);
        foreach (var area in areaChanges.Where(area => area.Value))
            diagnostics.Add(Diagnostic(AreaCodes[area.Key], AreaMessages[area.Key], area.Key));

        return ComparisonResult(
            original,
            revised,
            differences,
            diagnostics,
            diagnostics.Count == 0
                ? new DocumentPlan
                {
                    Format = DocFormat.Word,
                    Snapshot = before.Snapshot,
                    Revision = options.Revision,
                    Operations = operations
                }
                : null,
            BuildCoverage(
                bodyDifferenceCount,
                differences.Count - bodyDifferenceCount,
                diagnostics,
                areaChanges));
    }

    private static void ProcessHunk(
        IReadOnlyList<ParagraphInfo> before,
        IReadOnlyList<ParagraphInfo> after,
        int oldStart,
        int oldEnd,
        int newStart,
        int newEnd,
        List<PlanOperation> operations,
        List<DocumentDifference> differences,
        List<WorkflowDiagnostic> diagnostics,
        int maximumDifferences)
    {
        var oldCount = oldEnd - oldStart;
        var newCount = newEnd - newStart;
        var paired = Math.Min(oldCount, newCount);
        for (var i = 0; i < paired && differences.Count < maximumDifferences; i++)
        {
            var oldParagraph = before[oldStart + i];
            var newParagraph = after[newStart + i];
            if (!string.Equals(oldParagraph.StyleId, newParagraph.StyleId, StringComparison.Ordinal))
                diagnostics.Add(Diagnostic("unsupported-style-change",
                    $"Paragraph style changed at original body paragraph {oldStart + i}.", oldParagraph.ParaId));
            differences.Add(new DocumentDifference
            {
                Kind = DocumentDifferenceKind.Changed,
                OriginalIndex = oldStart + i,
                RevisedIndex = newStart + i,
                Before = oldParagraph.Text,
                After = newParagraph.Text
            });
            operations.Add(new ChangeTextOp
            {
                Target = Anchor(oldParagraph),
                With = newParagraph.Text,
                Mode = ChangeMode.Tracked
            });
        }

        if (newCount > paired && differences.Count < maximumDifferences)
        {
            var added = after.Skip(newStart + paired).Take(newCount - paired).ToList();
            for (var addedIndex = 0; addedIndex < added.Count && differences.Count < maximumDifferences; addedIndex++)
            {
                var paragraph = added[addedIndex];
                differences.Add(new DocumentDifference
                {
                    Kind = DocumentDifferenceKind.Added,
                    RevisedIndex = newStart + paired + addedIndex,
                    After = paragraph.Text
                });
            }
            var targetIndex = oldEnd < before.Count ? oldEnd : Math.Max(0, oldEnd - 1);
            operations.Add(new InsertParagraphsOp
            {
                Target = Anchor(before[targetIndex]),
                Position = oldEnd < before.Count ? InsertPosition.Before : InsertPosition.After,
                Paragraphs = added.Select(paragraph => new ParagraphData
                {
                    Text = paragraph.Text,
                    StyleId = paragraph.StyleId
                }).ToArray(),
                Mode = ChangeMode.Tracked
            });
        }

        for (var i = paired; i < oldCount && differences.Count < maximumDifferences; i++)
        {
            var paragraph = before[oldStart + i];
            differences.Add(new DocumentDifference
            {
                Kind = DocumentDifferenceKind.Removed,
                OriginalIndex = oldStart + i,
                Before = paragraph.Text
            });
            operations.Add(new RemoveParagraphOp { Target = Anchor(paragraph), Mode = ChangeMode.Tracked });
        }
    }

    private static IReadOnlyList<(int Old, int New)> LongestCommonSubsequence(
        IReadOnlyList<ParagraphInfo> before,
        IReadOnlyList<ParagraphInfo> after,
        CancellationToken cancellationToken)
    {
        var lengths = new int[before.Count + 1, after.Count + 1];
        for (var i = before.Count - 1; i >= 0; i--)
        {
            cancellationToken.ThrowIfCancellationRequested();
            for (var j = after.Count - 1; j >= 0; j--)
                lengths[i, j] = string.Equals(before[i].Text, after[j].Text, StringComparison.Ordinal)
                    ? lengths[i + 1, j + 1] + 1
                    : Math.Max(lengths[i + 1, j], lengths[i, j + 1]);
        }
        var matches = new List<(int Old, int New)>();
        var oldIndex = 0;
        var newIndex = 0;
        while (oldIndex < before.Count && newIndex < after.Count)
        {
            if (string.Equals(before[oldIndex].Text, after[newIndex].Text, StringComparison.Ordinal))
            {
                matches.Add((oldIndex++, newIndex++));
            }
            else if (lengths[oldIndex + 1, newIndex] >= lengths[oldIndex, newIndex + 1]) oldIndex++;
            else newIndex++;
        }
        return matches;
    }

    private static TextSpanAnchor Anchor(ParagraphInfo paragraph) => new()
    {
        Id = paragraph.ParaId,
        ParaId = paragraph.ParaId,
        Expect = paragraph.Text,
        Occurrence = 0
    };

    private static bool IsFreeBodyParagraph(ParagraphInfo paragraph) =>
        paragraph.Location == "body" && paragraph.In is null;

    /// <summary>
    /// Paragraphs outside the compared areas, keyed by where they live and what they say.
    /// </summary>
    /// <remarks>
    /// Table cells are excluded: their text is compared directly by the cell comparison
    /// once geometry is known to be unchanged. Including them here would report every
    /// supported cell edit as an unsupported non-body change.
    /// </remarks>
    private static IEnumerable<string> UnsupportedParagraphs(InspectResult result) =>
        result.Paragraphs
            .Where(paragraph => !IsFreeBodyParagraph(paragraph) && !IsTableCellParagraph(paragraph))
            .Select(paragraph => $"{paragraph.Location}|{paragraph.In}|{paragraph.Text}");

    /// <summary>Whether a paragraph lives inside a table cell in the document body.</summary>
    private static bool IsTableCellParagraph(ParagraphInfo paragraph) =>
        string.Equals(paragraph.Location, "body", StringComparison.Ordinal) &&
        paragraph.In is { Length: > 0 } node &&
        node.StartsWith("table#", StringComparison.Ordinal);

    private static IEnumerable<string> NodeSignature(InspectResult result, string kind) =>
        result.Nodes.Where(node => node.Kind == kind).Select(node => $"{node.Path}|{node.Summary}");

    private static IEnumerable<string> UnsupportedNodeSignature(InspectResult result) =>
        result.Nodes.Where(node => node.Kind is not "revision" and not "image" and not "table")
            .Select(node => $"{node.Kind}|{node.Path}|{node.Summary}");

    private static IReadOnlyList<string> FreeBodyParagraphMarkup(byte[] bytes)
    {
        XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        using var package = new ZipArchive(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);
        var entry = package.GetEntry("word/document.xml")
            ?? throw new InvalidDataException("The Word package has no main document part.");
        using var input = entry.Open();
        using var reader = XmlReader.Create(input, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = entry.Length + 1
        });
        var document = XDocument.Load(reader);
        var body = document.Root?.Element(word + "body");
        if (body is null) return Array.Empty<string>();
        return body.Elements(word + "p").Select(paragraph =>
        {
            var normalized = new XElement(paragraph);
            foreach (var text in normalized.Descendants().Where(element =>
                         element.Name == word + "t" || element.Name == word + "delText" ||
                         element.Name == word + "instrText"))
                text.RemoveNodes();
            foreach (var attribute in normalized.DescendantsAndSelf().Attributes().Where(attribute =>
                         attribute.Name.LocalName is "paraId" or "textId" ||
                         attribute.Name == XNamespace.Xml + "space" ||
                         (attribute.IsNamespaceDeclaration && attribute.Value != word.NamespaceName) ||
                         attribute.Name.LocalName.StartsWith("rsid", StringComparison.OrdinalIgnoreCase)).ToArray())
                attribute.Remove();
            MergeEquivalentRuns(normalized);
            return normalized.ToString(SaveOptions.DisableFormatting);
        }).ToArray();
    }


    /// <summary>
    /// Collapses adjacent runs that differ only in where the text was split.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Word re-segments runs constantly: typing in the middle of a sentence, a spell-check
    /// pass, or a round trip through another editor can turn one run into three carrying
    /// exactly the same formatting and exactly the same words. Comparing run structure
    /// directly reports that as a formatting change and withholds the plan, which is a
    /// false difference: no reader would call those documents different.
    /// </para>
    /// <para>
    /// Only text-carrying runs merge, and only into a neighbour with identical run
    /// properties. A run holding a break, a tab, a field, a drawing, a footnote reference
    /// or anything else is left exactly where it is, because those are content and moving
    /// them would change the document. Merging happens after text has been stripped, so
    /// this decides segmentation equivalence and never touches the words themselves.
    /// </para>
    /// </remarks>
    private static void MergeEquivalentRuns(XElement container)
    {
        XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

        // Hyperlinks, content controls and revision wrappers each hold their own run
        // sequence, and a run may only merge with a sibling inside the same wrapper.
        foreach (var parent in container.DescendantsAndSelf().ToList())
        {
            XElement? previous = null;
            foreach (var child in parent.Elements().ToList())
            {
                if (child.Name != word + "r" || !IsTextOnlyRun(child, word))
                {
                    previous = null;
                    continue;
                }

                if (previous is not null && SameRunProperties(previous, child, word))
                {
                    // The text is already stripped, so the merge is the removal itself.
                    child.Remove();
                    continue;
                }

                previous = child;
            }
        }
    }

    /// <summary>
    /// Whether a run carries only text, so its boundary is a segmentation detail rather
    /// than content.
    /// </summary>
    private static bool IsTextOnlyRun(XElement run, XNamespace word) =>
        run.Elements().All(element => element.Name == word + "rPr" || element.Name == word + "t");

    private static bool SameRunProperties(XElement left, XElement right, XNamespace word)
    {
        var leftProperties = left.Element(word + "rPr");
        var rightProperties = right.Element(word + "rPr");

        if (leftProperties is null && rightProperties is null) return true;
        if (leftProperties is null || rightProperties is null) return false;

        return string.Equals(
            leftProperties.ToString(SaveOptions.DisableFormatting),
            rightProperties.ToString(SaveOptions.DisableFormatting),
            StringComparison.Ordinal);
    }


    /// <summary>
    /// Compares the text of table cells that occupy the same position in tables whose
    /// geometry is unchanged.
    /// </summary>
    /// <remarks>
    /// <para>
    /// Alignment is positional, which is only sound because it runs after the table
    /// signature has established that the grids, spans and nesting are identical. If the
    /// shape moved, the cells in position three are not the same cell, and the signature
    /// difference blocks the plan before anything here is trusted.
    /// </para>
    /// <para>
    /// Only text changes are supported. A cell whose run structure carries something other
    /// than equivalently formatted text is reported unsupported rather than flattened:
    /// replacing rich content with plain text would lose the thing the caller was editing.
    /// </para>
    /// </remarks>
    private static void CompareTableCells(
        byte[] original,
        byte[] revised,
        InspectResult before,
        InspectResult after,
        List<PlanOperation> operations,
        List<DocumentDifference> differences,
        List<WorkflowDiagnostic> diagnostics,
        DocumentComparisonOptions options)
    {
        // Geometry first. A changed shape is reported by the area check, and cell text is
        // not comparable across it.
        if (!string.Equals(
                Convert.ToBase64String(TableSignature(DocumentPart(original))),
                Convert.ToBase64String(TableSignature(DocumentPart(revised))),
                StringComparison.Ordinal))
            return;

        var beforeCells = before.Paragraphs.Where(IsTableCellParagraph).ToList();
        var afterCells = after.Paragraphs.Where(IsTableCellParagraph).ToList();

        if (beforeCells.Count != afterCells.Count)
        {
            // The signature said the shape matched, so a differing cell-paragraph count
            // means something this comparison does not model. Refuse rather than guess.
            diagnostics.Add(Diagnostic("unsupported-table-change",
                "The tables have the same geometry but a different number of cell paragraphs, " +
                "so cells cannot be aligned safely.", "tables"));
            return;
        }

        var markupBefore = TableCellMarkup(original);
        var markupAfter = TableCellMarkup(revised);

        for (var i = 0; i < beforeCells.Count; i++)
        {
            var left = beforeCells[i];
            var right = afterCells[i];
            if (string.Equals(left.Text, right.Text, StringComparison.Ordinal)) continue;

            if (differences.Count >= options.MaximumDifferences)
            {
                diagnostics.Add(Diagnostic("comparison-difference-limit-exceeded",
                    $"The comparison reached the {options.MaximumDifferences}-difference limit.", "tables"));
                return;
            }

            if (i >= markupBefore.Count || i >= markupAfter.Count ||
                !string.Equals(markupBefore[i], markupAfter[i], StringComparison.Ordinal))
            {
                diagnostics.Add(Diagnostic("unsupported-table-markup-change",
                    $"Run structure or direct formatting changed in a cell of '{left.In}', so the " +
                    "text change cannot be reproduced without losing the formatting.", left.ParaId));
                continue;
            }

            differences.Add(new DocumentDifference
            {
                Kind = DocumentDifferenceKind.Changed,
                Before = left.Text,
                After = right.Text
            });

            operations.Add(new ChangeTextOp
            {
                Target = new TextSpanAnchor { ParaId = left.ParaId, Expect = left.Text },
                With = right.Text,
                Mode = ChangeMode.Tracked
            });
        }
    }

    /// <summary>The main document part of a Word package.</summary>
    private static byte[] DocumentPart(byte[] bytes)
    {
        using var package = new ZipArchive(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);
        var entry = package.GetEntry("word/document.xml")
            ?? throw new InvalidDataException("The Word package has no main document part.");
        using var input = entry.Open();
        using var copy = new MemoryStream();
        input.CopyTo(copy);
        return copy.ToArray();
    }

    /// <summary>
    /// The run structure of each table-cell paragraph, in document order, with text
    /// stripped and equivalent segmentation merged.
    /// </summary>
    private static IReadOnlyList<string> TableCellMarkup(byte[] bytes)
    {
        XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var part = DocumentPart(bytes);
        using var input = new MemoryStream(part, writable: false);
        using var reader = XmlReader.Create(input, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = part.LongLength + 1
        });

        var document = XDocument.Load(reader);
        var body = document.Root?.Element(word + "body");
        if (body is null) return Array.Empty<string>();

        return body.Elements(word + "tbl")
            .SelectMany(table => table.Descendants(word + "p"))
            .Select(paragraph =>
            {
                var normalized = new XElement(paragraph);
                foreach (var text in normalized.Descendants().Where(element =>
                             element.Name == word + "t" || element.Name == word + "delText" ||
                             element.Name == word + "instrText").ToList())
                    text.RemoveNodes();
                foreach (var attribute in normalized.DescendantsAndSelf().Attributes().Where(attribute =>
                             attribute.Name.LocalName is "paraId" or "textId" ||
                             attribute.Name == XNamespace.Xml + "space" ||
                             (attribute.IsNamespaceDeclaration && attribute.Value != word.NamespaceName) ||
                             attribute.Name.LocalName.StartsWith("rsid", StringComparison.OrdinalIgnoreCase))
                         .ToArray())
                    attribute.Remove();
                MergeEquivalentRuns(normalized);
                return normalized.ToString(SaveOptions.DisableFormatting);
            })
            .ToList();
    }

    private static void CompareParagraphMarkup(
        IReadOnlyList<string> before,
        IReadOnlyList<string> after,
        int oldIndex,
        int newIndex,
        ParagraphInfo original,
        List<WorkflowDiagnostic> diagnostics)
    {
        if (oldIndex >= before.Count || newIndex >= after.Count ||
            string.Equals(before[oldIndex], after[newIndex], StringComparison.Ordinal)) return;
        diagnostics.Add(Diagnostic("unsupported-paragraph-markup-change",
            $"Run structure or direct formatting changed at original body paragraph {oldIndex}.", original.ParaId));
    }

    private static string AddedParagraphMarkup(string neighborMarkup, string? styleId)
    {
        XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var neighbor = XElement.Parse(neighborMarkup);
        var paragraph = new XElement(word + "p");
        var properties = neighbor.Element(word + "pPr");
        if (properties is not null)
        {
            properties = new XElement(properties);
            properties.Elements(word + "pStyle").Remove();
        }
        if (styleId is not null)
        {
            properties ??= new XElement(word + "pPr");
            properties.AddFirst(new XElement(word + "pStyle", new XAttribute(word + "val", styleId)));
        }
        if (properties is not null) paragraph.Add(properties);
        paragraph.Add(new XElement(word + "r", new XElement(word + "t")));
        return paragraph.ToString(SaveOptions.DisableFormatting);
    }


    /// <summary>The package areas this comparison reports on, in a stable order.</summary>
    private static readonly string[] AreaNames =
    {
        "tables", "images", "notes", "headersAndFooters", "styles", "numbering", "otherParts"
    };

    private static readonly IReadOnlyDictionary<string, string> AreaCodes =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["tables"] = "unsupported-table-change",
            ["images"] = "unsupported-image-change",
            ["notes"] = "unsupported-note-change",
            ["headersAndFooters"] = "unsupported-header-footer-change",
            ["styles"] = "unsupported-style-definition-change",
            ["numbering"] = "unsupported-numbering-change",
            ["otherParts"] = "unsupported-package-change"
        };

    private static readonly IReadOnlyDictionary<string, string> AreaMessages =
        new Dictionary<string, string>(StringComparer.Ordinal)
        {
            ["tables"] = "Table geometry or unsupported table content differs. Cell text is " +
                         "compared only when geometry and supported markup remain unchanged.",
            ["images"] = "Embedded image bytes or drawings differ. Image content is not compared, " +
                         "so the difference is reported but not planned.",
            ["notes"] = "Footnote or endnote content differs. Note content is not compared, so the " +
                        "difference is reported but not planned.",
            ["headersAndFooters"] = "Header or footer content differs. Only the body is compared, so " +
                                    "the difference is reported but not planned.",
            ["styles"] = "The style definitions differ. Style definitions are not compared, so the " +
                         "difference is reported but not planned.",
            ["numbering"] = "List numbering definitions differ. Numbering is not compared, so the " +
                            "difference is reported but not planned.",
            ["otherParts"] = "A package part outside the compared areas differs, such as document " +
                             "properties, relationships, or custom XML."
        };

    /// <summary>
    /// Hashes each area of both packages separately and reports which ones differ.
    /// </summary>
    /// <remarks>
    /// The previous implementation hashed everything outside free body paragraphs into one
    /// value, so any difference anywhere produced the same message. Splitting it is what
    /// turns "something else changed" into "the images changed".
    /// </remarks>
    private static IReadOnlyDictionary<string, bool> CompareAreas(byte[] original, byte[] revised)
    {
        var before = AreaSignatures(original);
        var after = AreaSignatures(revised);

        var changed = new Dictionary<string, bool>(StringComparer.Ordinal);
        foreach (var area in AreaNames)
            changed[area] = !string.Equals(
                before.TryGetValue(area, out var left) ? left : string.Empty,
                after.TryGetValue(area, out var right) ? right : string.Empty,
                StringComparison.Ordinal);

        return changed;
    }

    /// <summary>One hash per reported area of a Word package.</summary>
    private static IReadOnlyDictionary<string, string> AreaSignatures(byte[] bytes)
    {
        var buckets = AreaNames.ToDictionary(
            area => area,
            _ => new SortedDictionary<string, byte[]>(StringComparer.Ordinal),
            StringComparer.Ordinal);

        using (var package = new ZipArchive(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read))
        {
            foreach (var entry in package.Entries)
            {
                var name = entry.FullName;
                byte[] content;
                using (var input = entry.Open())
                using (var copy = new MemoryStream())
                {
                    input.CopyTo(copy);
                    content = copy.ToArray();
                }

                if (string.Equals(name, "word/document.xml", StringComparison.Ordinal))
                {
                    // Free body paragraphs are compared paragraph by paragraph above, so
                    // only the table content of this part belongs to an area signature.
                    buckets["tables"][name] = TableSignature(content);
                    continue;
                }

                var area = AreaFor(name);
                if (area is null) continue;
                buckets[area][name] = content;
            }
        }

        var signatures = new Dictionary<string, string>(StringComparer.Ordinal);
        foreach (var bucket in buckets)
        {
            using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
            foreach (var entry in bucket.Value)
            {
                hash.AppendData(Encoding.UTF8.GetBytes(entry.Key));
                hash.AppendData(entry.Value);
            }

            signatures[bucket.Key] = BitConverter.ToString(hash.GetHashAndReset()).Replace("-", string.Empty);
        }

        return signatures;
    }

    private static readonly string[] ImageExtensions =
        { ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".tif", ".tiff", ".emf", ".wmf", ".svg" };

    private static string? AreaFor(string name)
    {
        // By media extension rather than by folder: an image part is not guaranteed to
        // live under word/media, and misfiling one would report it as a metadata change.
        if (ImageExtensions.Any(extension => name.EndsWith(extension, StringComparison.OrdinalIgnoreCase)))
            return "images";
        if (name is "word/footnotes.xml" or "word/endnotes.xml") return "notes";
        if (name.StartsWith("word/header", StringComparison.Ordinal) ||
            name.StartsWith("word/footer", StringComparison.Ordinal)) return "headersAndFooters";
        if (string.Equals(name, "word/styles.xml", StringComparison.Ordinal)) return "styles";
        if (string.Equals(name, "word/numbering.xml", StringComparison.Ordinal)) return "numbering";

        // Relationship parts move whenever anything they point at moves, so folding them
        // into otherParts would make every image change also look like a metadata change.
        if (name.EndsWith(".rels", StringComparison.Ordinal)) return null;
        return "otherParts";
    }

    /// <summary>
    /// The geometry and formatting of a part's tables, with cell words removed.
    /// </summary>
    /// <remarks>
    /// Text is stripped deliberately. This signature answers "did the table's shape
    /// change", and a cell whose words were edited is a supported difference handled by
    /// the cell comparison. Leaving text in would make every supported cell edit block
    /// itself as a geometry change.
    /// </remarks>
    private static byte[] TableSignature(byte[] documentPart)
    {
        XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        using var input = new MemoryStream(documentPart, writable: false);
        using var reader = XmlReader.Create(input, new XmlReaderSettings
        {
            DtdProcessing = DtdProcessing.Prohibit,
            XmlResolver = null,
            MaxCharactersInDocument = documentPart.LongLength + 1
        });

        var xml = XDocument.Load(reader);
        var tables = xml.Root?.Element(word + "body")?.Elements(word + "tbl")
            .Select(table => new XElement(table)).ToList();
        if (tables is null || tables.Count == 0) return Array.Empty<byte>();

        foreach (var table in tables)
        {
            foreach (var text in table.Descendants().Where(element =>
                         element.Name == word + "t" || element.Name == word + "delText" ||
                         element.Name == word + "instrText").ToList())
                text.RemoveNodes();

            foreach (var attribute in table.DescendantsAndSelf().Attributes().Where(attribute =>
                         attribute.Name.LocalName is "paraId" or "textId" ||
                         attribute.Name == XNamespace.Xml + "space" ||
                         attribute.Name.LocalName.StartsWith("rsid", StringComparison.OrdinalIgnoreCase))
                     .ToArray())
                attribute.Remove();

            MergeEquivalentRuns(table);
        }

        using var normalized = new MemoryStream();
        new XDocument(new XElement(word + "tables", tables)).Save(normalized, SaveOptions.DisableFormatting);
        return normalized.ToArray();
    }

    /// <summary>
    /// Describes every area, so a caller can tell an area that matched from one nobody
    /// looked at.
    /// </summary>
    private static IReadOnlyList<ComparisonArea> BuildCoverage(
        int bodyDifferenceCount,
        int tableDifferenceCount,
        IReadOnlyList<WorkflowDiagnostic> diagnostics,
        IReadOnlyDictionary<string, bool> areaChanges)
    {
        var bodyDiagnostic = diagnostics.FirstOrDefault(diagnostic =>
            (diagnostic.Code is "comparison-difference-limit-exceeded"
                or "comparison-paragraph-limit-exceeded"
                or "unsupported-existing-revisions") &&
            !string.Equals(diagnostic.Path, "tables", StringComparison.Ordinal));
        var tableDiagnostic = diagnostics.FirstOrDefault(diagnostic =>
            diagnostic.Code is "unsupported-table-change" or "unsupported-table-markup-change" ||
            (diagnostic.Code == "comparison-difference-limit-exceeded" &&
             string.Equals(diagnostic.Path, "tables", StringComparison.Ordinal)));

        var coverage = new List<ComparisonArea>
        {
            new()
            {
                Name = "bodyParagraphs",
                State = bodyDiagnostic is not null
                    ? ComparisonAreaState.Blocked
                    : bodyDifferenceCount > 0 ? ComparisonAreaState.Changed : ComparisonAreaState.Unchanged,
                Code = bodyDiagnostic?.Code
            },
            new()
            {
                Name = "paragraphFormatting",
                State = diagnostics.Any(diagnostic =>
                            diagnostic.Code is "unsupported-style-change"
                                or "unsupported-paragraph-markup-change")
                    ? ComparisonAreaState.Blocked
                    : ComparisonAreaState.Unchanged,
                Code = diagnostics.FirstOrDefault(diagnostic =>
                    diagnostic.Code is "unsupported-style-change"
                        or "unsupported-paragraph-markup-change")?.Code
            },
            new()
            {
                Name = "tables",
                State = tableDiagnostic is not null ||
                        areaChanges.TryGetValue("tables", out var tableChanged) && tableChanged
                    ? ComparisonAreaState.Blocked
                    : tableDifferenceCount > 0
                        ? ComparisonAreaState.Changed
                        : ComparisonAreaState.Unchanged,
                Code = tableDiagnostic?.Code ??
                       (areaChanges.TryGetValue("tables", out var tableAreaBlocked) && tableAreaBlocked
                           ? AreaCodes["tables"]
                           : null)
            }
        };

        foreach (var area in AreaNames.Where(area => area != "tables"))
            coverage.Add(new ComparisonArea
            {
                Name = area,
                State = areaChanges.TryGetValue(area, out var changed) && changed
                    ? ComparisonAreaState.Blocked
                    : ComparisonAreaState.Unchanged,
                Code = areaChanges.TryGetValue(area, out var blocked) && blocked ? AreaCodes[area] : null
            });

        return coverage;
    }

    private static string UnsupportedWordPackageSignature(byte[] bytes)
    {
        const string documentPart = "word/document.xml";
        XNamespace word = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        using var package = new ZipArchive(new MemoryStream(bytes, writable: false), ZipArchiveMode.Read);
        using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        foreach (var entry in package.Entries.OrderBy(entry => entry.FullName, StringComparer.Ordinal))
        {
            byte[] content;
            using (var input = entry.Open())
            using (var copy = new MemoryStream())
            {
                input.CopyTo(copy);
                content = copy.ToArray();
            }
            if (string.Equals(entry.FullName, documentPart, StringComparison.Ordinal))
            {
                using var input = new MemoryStream(content, writable: false);
                using var reader = XmlReader.Create(input, new XmlReaderSettings
                {
                    DtdProcessing = DtdProcessing.Prohibit,
                    XmlResolver = null,
                    MaxCharactersInDocument = content.LongLength + 1
                });
                var xml = XDocument.Load(reader);
                var body = xml.Root?.Element(word + "body");
                if (body is not null)
                {
                    foreach (var paragraph in body.Descendants(word + "p")
                                 .Where(paragraph => !paragraph.Ancestors(word + "tbl").Any()).ToArray())
                    {
                        var section = paragraph.Element(word + "pPr")?.Element(word + "sectPr");
                        if (section is null) paragraph.Remove();
                        else paragraph.ReplaceWith(new XElement(word + "p", new XElement(word + "pPr", new XElement(section))));
                    }
                }
                using var normalized = new MemoryStream();
                xml.Save(normalized, SaveOptions.DisableFormatting);
                content = normalized.ToArray();
            }
            else if (string.Equals(entry.FullName, "_rels/.rels", StringComparison.Ordinal))
            {
                using var input = new MemoryStream(content, writable: false);
                var xml = XDocument.Load(input);
                foreach (var relationship in xml.Root?.Elements() ?? Enumerable.Empty<XElement>())
                    relationship.Attribute("Id")?.Remove();
                xml.Root?.ReplaceNodes(xml.Root.Elements().OrderBy(element =>
                    (string?)element.Attribute("Type") + "|" + (string?)element.Attribute("Target") + "|" + (string?)element.Attribute("TargetMode"), StringComparer.Ordinal));
                using var normalized = new MemoryStream();
                xml.Save(normalized, SaveOptions.DisableFormatting);
                content = normalized.ToArray();
            }
            var name = Encoding.UTF8.GetBytes(entry.FullName);
            hash.AppendData(BitConverter.GetBytes(name.Length));
            hash.AppendData(name);
            hash.AppendData(BitConverter.GetBytes(content.Length));
            hash.AppendData(content);
        }
        return BitConverter.ToString(hash.GetHashAndReset()).Replace("-", string.Empty).ToLowerInvariant();
    }

    private static bool SequenceEqual(IEnumerable<string> left, IEnumerable<string> right) =>
        left.SequenceEqual(right, StringComparer.Ordinal);

    private static DocumentComparisonResult ComparisonResult(
        byte[] original,
        byte[] revised,
        IReadOnlyList<DocumentDifference>? differences = null,
        IReadOnlyList<WorkflowDiagnostic>? diagnostics = null,
        DocumentPlan? plan = null,
        IReadOnlyList<ComparisonArea>? coverage = null) => new()
        {
            OriginalSha256 = Sha256(original),
            RevisedSha256 = Sha256(revised),
            Differences = differences ?? Array.Empty<DocumentDifference>(),
            Diagnostics = diagnostics ?? Array.Empty<WorkflowDiagnostic>(),
            // A comparison that stopped early knows nothing about any area, and saying so
            // is the whole point: an empty coverage list would read as "all unchanged".
            Coverage = coverage ?? NothingCompared(),
            Plan = plan
        };

    /// <summary>Every area marked not compared, for a comparison that never ran.</summary>
    private static IReadOnlyList<ComparisonArea> NothingCompared() =>
        new[] { "bodyParagraphs", "paragraphFormatting" }
            .Concat(AreaNames)
            .Select(name => new ComparisonArea
            {
                Name = name,
                State = ComparisonAreaState.NotCompared,
                Detail = "The comparison stopped before this area was examined."
            })
            .ToList();

    private static async Task<byte[]> ReadBoundedAsync(
        Stream source,
        long maximumBytes,
        CancellationToken cancellationToken)
    {
        if (source.CanSeek && source.Length > maximumBytes)
            throw new InvalidDataException($"Document exceeds the {maximumBytes}-byte comparison limit.");
        using var output = new MemoryStream();
        var buffer = new byte[81920];
        while (true)
        {
            var read = await source.ReadAsync(buffer, 0, buffer.Length, cancellationToken).ConfigureAwait(false);
            if (read == 0) break;
            if (output.Length + read > maximumBytes)
                throw new InvalidDataException($"Document exceeds the {maximumBytes}-byte comparison limit.");
            output.Write(buffer, 0, read);
        }
        return output.ToArray();
    }

    private static string Sha256(byte[] bytes)
    {
        using var sha = SHA256.Create();
        return BitConverter.ToString(sha.ComputeHash(bytes)).Replace("-", string.Empty).ToLowerInvariant();
    }

    private static WorkflowDiagnostic Diagnostic(string code, string message, string? path) => new()
    {
        Code = code,
        Message = message,
        Path = path
    };

    private static string ToKebabCase(string value) =>
        string.Concat(value.Select((ch, index) => char.IsUpper(ch) && index > 0 ? $"-{char.ToLowerInvariant(ch)}" : char.ToLowerInvariant(ch).ToString()));
}
