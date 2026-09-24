using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;

namespace OfficeAgent.Samples.RequirementsReview;

/// <summary>One anchored paragraph of evidence.</summary>
/// <param name="EvidenceId">Host-assigned id, e.g. <c>REQ-OPS-03:1</c>. Models cite this id, never the text.</param>
/// <param name="ParaId">The paragraph's <c>w14:paraId</c>.</param>
/// <param name="Text">The paragraph text as OfficeAgent inspection reports it.</param>
/// <param name="Container">The containing node, e.g. <c>table#0</c>, or null for body flow.</param>
/// <param name="HasPendingRevision">True when the paragraph carries somebody's unaccepted tracked change.</param>
/// <param name="HasExistingComment">True when an existing comment is anchored in the paragraph.</param>
/// <param name="ReviewStateKnown">False when the host could not read this paragraph's review state.</param>
public sealed record EvidenceItem(
    string EvidenceId,
    string ParaId,
    string Text,
    string? Container,
    bool HasPendingRevision,
    bool HasExistingComment,
    bool ReviewStateKnown = true);

/// <summary>The evidence gathered for one requirement.</summary>
/// <param name="RequirementId">The requirement.</param>
/// <param name="Section">The section heading the registry named.</param>
/// <param name="SectionMatches">How many headings matched. Anything but 1 is not usable evidence.</param>
/// <param name="HeadingParaId">The matched heading's paragraph id, when exactly one matched.</param>
/// <param name="HeadingText">The matched heading's text, when exactly one matched.</param>
/// <param name="Items">The non-empty paragraphs under the heading.</param>
public sealed record AnchoredDocumentEvidence(
    string RequirementId,
    string Section,
    int SectionMatches,
    string? HeadingParaId,
    string? HeadingText,
    IReadOnlyList<EvidenceItem> Items)
{
    /// <summary>Gets the evidence character count, the unit of the external-evaluation budget.</summary>
    public int Characters => Items.Sum(i => i.Text.Length);

    /// <summary>Gets the joined evidence text used by deterministic checks.</summary>
    public string JoinedText => string.Join("\n", Items.Select(i => i.Text));
}

/// <summary>Everything the run knows about the document, bound to one set of bytes.</summary>
/// <param name="ConnectionId">The OfficeAgent connection.</param>
/// <param name="DocumentId">The opaque document id.</param>
/// <param name="Name">The document name.</param>
/// <param name="SourceSha256">SHA-256 of the exact bytes inspected.</param>
/// <param name="Snapshot">The OfficeAgent snapshot of those bytes.</param>
/// <param name="ExistingCommentCount">Comments already in the document.</param>
/// <param name="ExistingRevisionCount">Tracked revisions already in the document.</param>
/// <param name="TableCount">Tables in the document.</param>
/// <param name="ByRequirement">Evidence keyed by requirement id.</param>
public sealed record DocumentEvidence(
    string ConnectionId,
    string DocumentId,
    string Name,
    string SourceSha256,
    SnapshotToken Snapshot,
    int ExistingCommentCount,
    int ExistingRevisionCount,
    int TableCount,
    IReadOnlyDictionary<string, AnchoredDocumentEvidence> ByRequirement);

/// <summary>
/// Reads the document once, hashes those bytes, inspects the same bytes with OfficeAgent,
/// and cuts one section of anchored evidence per requirement.
/// </summary>
public static class EvidenceCollector
{
    /// <summary>Collects evidence for every requirement in the set.</summary>
    /// <param name="client">The OfficeAgent client.</param>
    /// <param name="connectionId">The connection.</param>
    /// <param name="documentId">The document.</param>
    /// <param name="set">The requirement set.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>The document evidence.</returns>
    public static async Task<DocumentEvidence> CollectAsync(
        OfficeAgentClient client,
        string connectionId,
        string documentId,
        RequirementSet set,
        CancellationToken cancellationToken)
    {
        byte[] bytes;
        string name;
        using (var content = await client.OpenReadAsync(connectionId, documentId, cancellationToken).ConfigureAwait(false))
        {
            using var buffer = new MemoryStream();
            await content.Stream.CopyToAsync(buffer, cancellationToken).ConfigureAwait(false);
            bytes = buffer.ToArray();
            name = content.Reference.Name ?? documentId;
        }

        // Inspect the bytes that were hashed, not a second read of the provider. The
        // snapshot, the hash, and the review state below then describe one document.
        var inspected = await client
            .InspectAsync(new StreamHandle(new MemoryStream(bytes, writable: false), name), cancellationToken: cancellationToken)
            .ConfigureAwait(false);

        var state = ReviewStateReader.Read(bytes);
        var body = inspected.Paragraphs.Where(p => p.Location == "body").ToList();

        var byRequirement = new Dictionary<string, AnchoredDocumentEvidence>(StringComparer.Ordinal);
        foreach (var requirement in set.Requirements)
        {
            byRequirement[requirement.Id] = Cut(requirement, body, state);
        }

        return new DocumentEvidence(
            connectionId,
            documentId,
            name,
            Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant(),
            inspected.Snapshot,
            inspected.Nodes.Count(n => n.Kind == "comment"),
            inspected.Nodes.Count(n => n.Kind == "revision"),
            inspected.Nodes.Count(n => n.Kind == "table"),
            byRequirement);
    }

    private static AnchoredDocumentEvidence Cut(
        RegisteredRequirement requirement,
        IReadOnlyList<ParagraphInfo> body,
        ReviewState state)
    {
        var matches = body
            .Select((p, index) => (p, index))
            .Where(x => HeadingLevel(x.p) is not null
                        && string.Equals(x.p.Text.Trim(), requirement.Section, StringComparison.OrdinalIgnoreCase))
            .ToList();

        // Zero headings is missing evidence; two is ambiguous evidence. Neither is guessed.
        if (matches.Count != 1)
        {
            return new AnchoredDocumentEvidence(requirement.Id, requirement.Section, matches.Count, null, null, Array.Empty<EvidenceItem>());
        }

        var (heading, start) = matches[0];
        var level = HeadingLevel(heading)!.Value;
        var items = new List<EvidenceItem>();
        for (var i = start + 1; i < body.Count; i++)
        {
            var paragraph = body[i];
            if (HeadingLevel(paragraph) is { } next && next <= level)
            {
                break;
            }

            if (string.IsNullOrWhiteSpace(paragraph.Text))
            {
                continue;
            }

            items.Add(new EvidenceItem(
                $"{requirement.Id}:{items.Count + 1}",
                paragraph.ParaId,
                paragraph.Text,
                paragraph.In,
                state.PendingRevision.Contains(paragraph.ParaId),
                state.Commented.Contains(paragraph.ParaId),
                state.Known.Contains(paragraph.ParaId)));
        }

        return new AnchoredDocumentEvidence(requirement.Id, requirement.Section, 1, heading.ParaId, heading.Text, items);
    }

    private static int? HeadingLevel(ParagraphInfo paragraph) =>
        paragraph.StyleId is { Length: 8 } style
        && style.StartsWith("Heading", StringComparison.Ordinal)
        && char.IsDigit(style[7])
            ? style[7] - '0'
            : null;
}

/// <summary>Paragraph ids that carry pending revisions or existing comments.</summary>
/// <param name="Known">Every body paragraph id the reader saw, in OfficeAgent's id format.</param>
/// <param name="PendingRevision">Paragraphs containing an unaccepted tracked change.</param>
/// <param name="Commented">Paragraphs containing an existing comment anchor.</param>
public sealed record ReviewState(IReadOnlySet<string> Known, IReadOnlySet<string> PendingRevision, IReadOnlySet<string> Commented);

/// <summary>
/// Reads review state per paragraph. OfficeAgent 0.9 inspection reports paragraph text as if
/// every pending change were accepted, and lists revisions without their paragraph id, so the
/// host reads this one fact from the same bytes it hashed.
/// </summary>
public static class ReviewStateReader
{
    private const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    private static readonly HashSet<string> RevisionTags = new(StringComparer.Ordinal)
    {
        "ins", "del", "moveFrom", "moveTo", "pPrChange", "rPrChange",
        "moveFromRangeStart", "moveToRangeStart",
    };

    /// <summary>Reads the review state of a .docx.</summary>
    /// <param name="bytes">The document bytes.</param>
    /// <returns>The review state.</returns>
    public static ReviewState Read(byte[] bytes)
    {
        var known = new HashSet<string>(StringComparer.Ordinal);
        var pending = new HashSet<string>(StringComparer.Ordinal);
        var commented = new HashSet<string>(StringComparer.Ordinal);

        using var document = WordprocessingDocument.Open(new MemoryStream(bytes, writable: false), isEditable: false);
        var body = document.MainDocumentPart?.Document?.Body;
        if (body is null)
        {
            return new ReviewState(known, pending, commented);
        }

        // Same id format as OfficeAgent inspection: w14:paraId when present, otherwise the
        // positional id. An id the two disagree on is simply not found, and the workflow
        // treats a paragraph with unknown review state as unusable evidence.
        var index = 0;
        foreach (var paragraph in body.Descendants<Paragraph>())
        {
            var paraId = paragraph.ParagraphId?.Value is { Length: > 0 } w14 ? $"w14:{w14}" : $"auto-{index:D4}";
            index++;
            known.Add(paraId);

            if (paragraph.Descendants().Any(e => e.NamespaceUri == W && RevisionTags.Contains(e.LocalName)))
            {
                pending.Add(paraId);
            }

            if (paragraph.Descendants<CommentRangeStart>().Any() || paragraph.Descendants<CommentReference>().Any())
            {
                commented.Add(paraId);
            }
        }

        return new ReviewState(known, pending, commented);
    }
}
