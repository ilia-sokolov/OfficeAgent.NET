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
    public async Task<TemplateBatchResult> PopulateTemplateBatchAsync(
        DocumentReference template,
        TemplateBatchRequest request,
        CancellationToken cancellationToken = default)
    {
        if (template is null) throw new ArgumentNullException(nameof(template));
        if (request is null) throw new ArgumentNullException(nameof(request));
        if (request.MaximumDocuments <= 0 || request.Items.Count > request.MaximumDocuments)
            throw new ArgumentOutOfRangeException(nameof(request),
                $"Template batches must contain at most {request.MaximumDocuments} documents.");

        var results = new List<TemplateBatchItemResult>();
        foreach (var item in request.Items)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var plan = await BuildTemplatePlanAsync(template, item.Binding, cancellationToken).ConfigureAwait(false);
            if (!plan.IsValid || plan.Plan is null)
            {
                results.Add(new TemplateBatchItemResult
                {
                    OutputName = item.OutputName,
                    Diagnostics = plan.Diagnostics
                });
                if (!request.ContinueOnError) break;
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
                    Document = applied.Document,
                    Report = applied.Report,
                    Receipt = applied.Receipt,
                    Diagnostics = applied.Report.Errors.Select(error =>
                        Diagnostic(error.Code, error.Message, error.Target?.Id)).ToArray()
                });
                if (!applied.Committed && !request.ContinueOnError) break;
            }
            catch (DocumentProviderException ex)
            {
                results.Add(new TemplateBatchItemResult
                {
                    OutputName = item.OutputName,
                    Diagnostics = new[] { Diagnostic(ToKebabCase(ex.Code.ToString()), ex.Message, item.OutputName) }
                });
                if (!request.ContinueOnError) break;
            }
        }
        return new TemplateBatchResult { Items = results };
    }

    /// <summary>Populates a template batch by opaque connection and document ids.</summary>
    public Task<TemplateBatchResult> PopulateTemplateBatchAsync(
        string connectionId,
        string documentId,
        TemplateBatchRequest request,
        CancellationToken cancellationToken = default) =>
        PopulateTemplateBatchAsync(ReferenceFor(connectionId, documentId), request, cancellationToken);

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
        if (diagnostics.Count > 0)
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
        if (!string.Equals(UnsupportedWordPackageSignature(original), UnsupportedWordPackageSignature(revised), StringComparison.Ordinal))
            diagnostics.Add(Diagnostic("unsupported-package-change",
                "Content outside free body paragraphs changed, such as tables, images, relationships, properties, or package metadata.", "package"));
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
                : null);
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

    private static IEnumerable<string> UnsupportedParagraphs(InspectResult result) =>
        result.Paragraphs.Where(paragraph => !IsFreeBodyParagraph(paragraph))
            .Select(paragraph => $"{paragraph.Location}|{paragraph.In}|{paragraph.Text}");

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
            return normalized.ToString(SaveOptions.DisableFormatting);
        }).ToArray();
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
        DocumentPlan? plan = null) => new()
        {
            OriginalSha256 = Sha256(original),
            RevisedSha256 = Sha256(revised),
            Differences = differences ?? Array.Empty<DocumentDifference>(),
            Diagnostics = diagnostics ?? Array.Empty<WorkflowDiagnostic>(),
            Plan = plan
        };

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
