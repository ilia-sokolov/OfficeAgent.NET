using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Accepts or rejects tracked revisions addressed by a revision <see cref="NodeAnchor"/> -
/// one by id (<c>ins#7</c>), everything by one author (<c>author:Jane Doe</c>), or the lot
/// (<c>all</c>).
/// </summary>
/// <remarks>
/// <para>
/// Every marker WordprocessingML defines is resolved, not only the run wrappers: paragraph
/// marks, table rows and cells, moves, and the <c>*PrChange</c> formatting revisions.
/// Accepting a deleted paragraph mark merges the paragraph with the next one, which is
/// what deleting a pilcrow means; rejecting an inserted one does the same, because the
/// split it recorded never happened.
/// </para>
/// <para>
/// Structural markers are applied after the run-level ones, and anything that has since
/// been detached - a run inside a row this same operation removed - is skipped rather than
/// edited in a subtree no longer in the document.
/// </para>
/// </remarks>
internal sealed class RevisionHandler : IOperationHandler
{
    private readonly RevisionNodeProvider _provider = new();

    public bool CanHandle(PlanOperation operation) =>
        operation is RevisionOp { Target: NodeAnchor { Kind: "revision" } };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (RevisionOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var node = _provider.Resolve(anchor, new WordObjectMap(context.Package));

        if (node is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No revision matches '{anchor.Path}'. Paths come from inspect_document " +
                "(for example 'ins#7' or 'paraFormat#3'), or use 'all', or 'author:<name>'.",
                anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "revision",
            Before = $"{node.Elements.Count} tracked revision(s)",
            After = op.Action == RevisionAction.Accept ? "accepted" : "rejected",
            Context = anchor.Path,
            BlastRadius = Math.Max(1, node.Elements.Count),
            Capability = Capability.Deterministic
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (RevisionOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var accept = op.Action == RevisionAction.Accept;

        var node = _provider.Resolve(anchor, new WordObjectMap(context.Package))
            ?? throw new InvalidOperationException($"Revision '{anchor.Path}' vanished before apply.");

        var main = WordModel.Main(context.Package);
        var notesBefore = NoteNodeProvider.ReferencedIds(main);

        // Runs first, then paragraph marks, then the structures that can carry both away
        // with them. Resolving a row before the runs inside it would leave those runs
        // detached and their revisions silently unresolved.
        foreach (var element in node.Elements.Where(IsRunLevel).ToList())
            Resolve(element, accept);

        foreach (var element in node.Elements.Where(e => !IsRunLevel(e) && !IsStructural(e)).ToList())
            Resolve(element, accept);

        foreach (var element in node.Elements.Where(IsStructural).ToList())
            Resolve(element, accept);

        // Accepting the deletion of a footnote reference deletes the footnote, the way Word
        // does: a note with nothing pointing at it renders nowhere, and leaving one behind
        // would offer an agent a note the reader cannot see.
        NoteNodeProvider.PruneOrphans(main, notesBefore);
    }

    private static bool IsRunLevel(OpenXmlElement element) =>
        element is InsertedRun or DeletedRun or MoveFromRun or MoveToRun;

    private static bool IsStructural(OpenXmlElement element) =>
        (element is Inserted or Deleted && element.Parent is TableRowProperties)
        || element is CellInsertion or CellDeletion;

    private static void Resolve(OpenXmlElement element, bool accept)
    {
        if (!IsLive(element)) return;

        switch (element)
        {
            // ── Runs ─────────────────────────────────────────────────────
            case InsertedRun ins:
                if (accept) Unwrap(ins); else ins.Remove();
                break;

            case MoveToRun moveTo:
                if (accept) Unwrap(moveTo); else moveTo.Remove();
                break;

            case DeletedRun del:
                if (accept) del.Remove();
                else { RestoreDeletedText(del); Unwrap(del); }
                break;

            case MoveFromRun moveFrom:
                if (accept) moveFrom.Remove();
                else { RestoreDeletedText(moveFrom); Unwrap(moveFrom); }
                break;

            // ── Paragraph marks ──────────────────────────────────────────
            case Inserted markIns when element.Parent is ParagraphMarkRunProperties:
                if (accept) RemoveMarker(markIns); else MergeWithNext(ParagraphOf(markIns));
                break;

            case Deleted markDel when element.Parent is ParagraphMarkRunProperties:
                if (accept) MergeWithNext(ParagraphOf(markDel)); else RemoveMarker(markDel);
                break;

            // ── Table rows and cells ─────────────────────────────────────
            case Inserted rowIns when element.Parent is TableRowProperties:
                if (accept) RemoveMarker(rowIns); else element.Ancestors<TableRow>().FirstOrDefault()?.Remove();
                break;

            case Deleted rowDel when element.Parent is TableRowProperties:
                if (accept) element.Ancestors<TableRow>().FirstOrDefault()?.Remove(); else RemoveMarker(rowDel);
                break;

            case CellInsertion cellIns:
                if (accept) RemoveMarker(cellIns); else element.Ancestors<TableCell>().FirstOrDefault()?.Remove();
                break;

            case CellDeletion cellDel:
                if (accept) element.Ancestors<TableCell>().FirstOrDefault()?.Remove(); else RemoveMarker(cellDel);
                break;

            // ── Formatting ───────────────────────────────────────────────
            case RunPropertiesChange runFormat:
                RestoreProperties<PreviousRunProperties>(runFormat, accept);
                break;

            case ParagraphPropertiesChange paraFormat:
                RestoreProperties<PreviousParagraphProperties>(paraFormat, accept);
                break;

            case TablePropertiesChange tableFormat:
                RestoreProperties<PreviousTableProperties>(tableFormat, accept);
                break;

            case TableRowPropertiesChange rowFormat:
                RestoreProperties<PreviousTableRowProperties>(rowFormat, accept);
                break;

            case TableCellPropertiesChange cellFormat:
                RestoreProperties<PreviousTableCellProperties>(cellFormat, accept);
                break;
        }
    }

    /// <summary>
    /// Whether the element is still part of the document. Removing a row detaches the
    /// revisions inside it without clearing their parent pointers, so reachability is what
    /// distinguishes a live marker from one already carried away.
    /// </summary>
    private static bool IsLive(OpenXmlElement element)
    {
        var root = element;
        while (root.Parent is not null) root = root.Parent;
        return root is Document or Footnotes or Endnotes or Header or Footer;
    }

    /// <summary>
    /// Removes a marker and the property container it leaves empty behind it. An empty
    /// <c>w:rPr</c> in a paragraph mark is harmless but noisy; an empty <c>w:trPr</c> is
    /// the same, and both are what Word itself removes when it accepts a change.
    /// </summary>
    private static void RemoveMarker(OpenXmlElement marker)
    {
        var container = marker.Parent;
        marker.Remove();

        if (container is null || container.HasChildren) return;
        if (container is ParagraphMarkRunProperties or TableRowProperties or TableCellProperties)
            container.Remove();
    }

    private static Paragraph? ParagraphOf(OpenXmlElement marker) =>
        marker.Ancestors<Paragraph>().FirstOrDefault();

    /// <summary>
    /// Joins a paragraph to the one after it, which is what resolving its paragraph mark
    /// amounts to: the mark is what ends the paragraph, so without it the two are one.
    /// </summary>
    private static void MergeWithNext(Paragraph? paragraph)
    {
        if (paragraph is null) return;

        var next = paragraph.NextSibling<Paragraph>();
        if (next is null)
        {
            // Nothing to merge into. A paragraph with content has to stay - there is
            // nowhere to put it - so the marker is dropped and the text kept. An empty one
            // is the tail of an inserted paragraph whose runs this same operation just
            // rejected, and removing it is what restores the document as it was.
            if (!HasContent(paragraph) && paragraph.Parent?.Elements<Paragraph>().Count() > 1)
            {
                paragraph.Remove();
                return;
            }

            if (paragraph.ParagraphProperties?.ParagraphMarkRunProperties is { } mark)
            {
                mark.RemoveAllChildren<Inserted>();
                mark.RemoveAllChildren<Deleted>();
                if (!mark.HasChildren) mark.Remove();
            }
            return;
        }

        // The surviving paragraph keeps its own properties: the merged text continues into
        // it, so it is that paragraph's style the reader ends up seeing.
        var content = paragraph.ChildElements
            .Where(child => child is not ParagraphProperties)
            .ToList();

        OpenXmlElement cursor = next.ParagraphProperties is { } properties ? properties : null!;
        foreach (var child in content)
        {
            child.Remove();
            if (cursor is null) next.InsertAt(child, 0);
            else cursor.InsertAfterSelf(child);
            cursor = child;
        }

        paragraph.Remove();
    }

    /// <summary>Whether the paragraph holds anything but its own properties.</summary>
    private static bool HasContent(Paragraph paragraph) =>
        paragraph.ChildElements.Any(child => child is not ParagraphProperties);

    /// <summary>
    /// Accepting a formatting revision drops the record of what came before; rejecting it
    /// puts those properties back exactly as they were.
    /// </summary>
    /// <remarks>
    /// The restore replaces only what the change could have captured. A <c>w:pPrChange</c>
    /// records <c>CT_PPrBase</c>, which is <c>w:pPr</c> <em>without</em> the paragraph
    /// mark's own run properties or the section properties - and <c>w:tcPrChange</c> and
    /// <c>w:trPrChange</c> exclude their revision markers the same way. Clearing those out
    /// along with the formatting would drop a pending paragraph-mark or row revision that
    /// this operation was never addressing.
    /// </remarks>
    private static void RestoreProperties<TPrevious>(OpenXmlElement change, bool accept)
        where TPrevious : OpenXmlElement
    {
        var properties = change.Parent;

        if (!accept && properties is not null)
        {
            var previous = change.GetFirstChild<TPrevious>();

            foreach (var child in properties.ChildElements.ToList())
                if (!ReferenceEquals(child, change) && !OutsideTheChange(child))
                    child.Remove();

            if (previous is not null)
                foreach (var child in previous.ChildElements.ToList())
                    properties.InsertBefore(child.CloneNode(deep: true), change);
        }

        change.Remove();

        if (properties is not null && !properties.HasChildren &&
            properties is ParagraphMarkRunProperties or TableRowProperties or TableCellProperties or RunProperties)
            properties.Remove();
    }

    /// <summary>Children a <c>*PrChange</c> never records, and so must never restore over.</summary>
    private static bool OutsideTheChange(OpenXmlElement child) =>
        child is ParagraphMarkRunProperties or SectionProperties
              or Inserted or Deleted or CellInsertion or CellDeletion;

    private static void Unwrap(OpenXmlElement wrapper)
    {
        var parent = wrapper.Parent ?? throw new InvalidOperationException("Revision has no parent.");
        foreach (var child in wrapper.ChildElements.ToList())
        {
            child.Remove();
            parent.InsertBefore(child, wrapper);
        }
        wrapper.Remove();
    }

    /// <summary>Turns <c>w:delText</c> back into <c>w:t</c>, restoring deleted wording.</summary>
    private static void RestoreDeletedText(OpenXmlElement deleted)
    {
        foreach (var delText in deleted.Descendants<DeletedText>().ToList())
        {
            var text = new Text(delText.Text ?? string.Empty) { Space = SpaceProcessingModeValues.Preserve };
            delText.InsertAfterSelf(text);
            delText.Remove();
        }
    }
}
