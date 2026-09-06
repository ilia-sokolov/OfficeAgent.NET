using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Creates, rewrites, and deletes footnotes and endnotes.
/// </summary>
/// <remarks>
/// <para>
/// The point of a real note over a superscript typed into the text is that Word owns the
/// number: add one in the middle of a contract and everything after it renumbers. That
/// costs two things this handler is careful about - the notes part has to exist with the
/// separator entries Word expects, and the styles the reference and the note body name
/// have to be defined, or the note renders as ordinary body text at the foot of the page.
/// </para>
/// <para>
/// The note body is also an ordinary paragraph in an ordinary text host, so
/// <c>changeText</c>, <c>format</c>, and <c>comment</c> all reach it once it exists; this
/// verb is for the parts of a note that are not text - its reference, its numbering, and
/// its existence.
/// </para>
/// </remarks>
internal sealed class WordNoteHandler : IOperationHandler
{
    private readonly TimeProvider _clock;

    public WordNoteHandler(TimeProvider clock) => _clock = clock;

    public bool CanHandle(PlanOperation operation) => operation switch
    {
        NoteOp { Target: TextSpanAnchor, Action: NoteAction.Add } => true,
        NoteOp { Target: NodeAnchor { Kind: "note" }, Action: NoteAction.Update } => true,
        NoteOp { Target: NodeAnchor { Kind: "note" }, Action: NoteAction.Remove } => true,
        _ => false
    };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (NoteOp)operation;
        return op.Action == NoteAction.Add
            ? PreviewAdd(context, op, (TextSpanAnchor)op.Target)
            : PreviewOnExisting(context, op, (NodeAnchor)op.Target);
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (NoteOp)operation;
        var main = WordModel.Main(context.Package);

        switch (op.Action)
        {
            case NoteAction.Add: ApplyAdd(context, main, op, (TextSpanAnchor)op.Target); break;
            case NoteAction.Update: ApplyUpdate(context, main, op, (NodeAnchor)op.Target); break;
            case NoteAction.Remove: ApplyRemove(context, main, op, (NodeAnchor)op.Target); break;
        }
    }

    // ── Add ──────────────────────────────────────────────────────────────

    private static OperationPreview PreviewAdd(ApplyContext context, NoteOp op, TextSpanAnchor anchor)
    {
        var paragraph = WordModel.ResolveParagraph(context, anchor.ParaId);
        if (paragraph is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No paragraph with id '{anchor.ParaId}'.", anchor));

        // A note reference inside a note is a document Word cannot lay out - the note it
        // points at would have to be printed inside itself.
        if (!WordModel.IsInMainBody(paragraph))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"Paragraph '{anchor.ParaId}' is not in the document body; a note reference belongs in body text. " +
                "Re-inspect and target a paragraph whose location is \"body\".", anchor));

        if (string.IsNullOrEmpty(op.Text))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"Adding a {NoteNodeProvider.Word(op.Kind)} needs 'text'.", anchor));

        if (!string.IsNullOrEmpty(anchor.Expect))
        {
            var text = WordModel.Text.GetLogicalText(paragraph);
            if (WordModel.Text.IndexOfOccurrence(text, anchor.Expect, anchor.Occurrence, WordModel.Comparison(true)) < 0)
                return OperationPreview.Fail(new ValidationError(
                    ValidationErrorCodes.ExpectMismatch,
                    $"Expected text '{anchor.Expect}' not found in paragraph '{anchor.ParaId}'.", anchor));
        }

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "note",
            Before = string.Empty,
            After = op.Text,
            Context = string.IsNullOrEmpty(anchor.Expect)
                ? $"{NoteNodeProvider.Word(op.Kind)} at the end of paragraph '{anchor.ParaId}'"
                : $"{NoteNodeProvider.Word(op.Kind)} after '{anchor.Expect}'",
            BlastRadius = 1,
            Capability = Capability.Deterministic
        });
    }

    private void ApplyAdd(ApplyContext context, MainDocumentPart main, NoteOp op, TextSpanAnchor anchor)
    {
        var paragraph = WordModel.ResolveParagraph(context, anchor.ParaId)
            ?? throw new InvalidOperationException($"Paragraph '{anchor.ParaId}' vanished before apply.");

        var notes = EnsureNotesRoot(main, op.Kind);
        var id = NextNoteId(notes, op.Kind);

        WordNoteStyles.Ensure(main, op.Kind);

        var body = BuildNoteBody(op, id);
        notes.AppendChild(body);

        var reference = ReferenceRun(op.Kind, id);
        PlaceReference(paragraph, anchor, reference);

        if (!WordRevisionMarker.IsTracked(op.Mode)) return;

        // The reference is the revision Word keys on - accepting or rejecting it is what
        // makes the note appear or vanish - and the note's own text is marked with it so a
        // rejected note does not leave its wording at the foot of the page.
        var marker = new WordRevisionMarker(context.Package, _clock);
        marker.WrapInserted(reference);
        foreach (var noteParagraph in body.Elements<Paragraph>())
            marker.MarkContentInserted(noteParagraph);
    }

    /// <summary>
    /// Places the reference immediately after the named text, or at the end of the
    /// paragraph when the anchor names none.
    /// </summary>
    private static void PlaceReference(Paragraph paragraph, TextSpanAnchor anchor, Run reference)
    {
        if (string.IsNullOrEmpty(anchor.Expect))
        {
            paragraph.AppendChild(reference);
            return;
        }

        var text = WordModel.Text.GetLogicalText(paragraph);
        int start = WordModel.Text.IndexOfOccurrence(
            text, anchor.Expect, anchor.Occurrence, WordModel.Comparison(true));
        if (start < 0)
            throw new InvalidOperationException($"Expected text '{anchor.Expect}' not found at apply time.");

        var covered = WordModel.Text.IsolateSpan(paragraph, start, anchor.Expect.Length);
        covered[covered.Count - 1].InsertAfterSelf(reference);
    }

    // ── Update and remove ────────────────────────────────────────────────

    private static OperationPreview PreviewOnExisting(ApplyContext context, NoteOp op, NodeAnchor anchor)
    {
        var main = WordModel.Main(context.Package);
        var located = NoteNodeProvider.Locate(main, anchor.Path);
        if (located is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"No note at path '{anchor.Path}'. Paths come from inspect_document, as 'footnote#<id>' or 'endnote#<id>'.",
                anchor));

        if (op.Action == NoteAction.Update && string.IsNullOrEmpty(op.Text))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "Updating a note needs 'text'. To delete it, use action 'Remove'.", anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "note",
            Before = NoteNodeProvider.TextOf(located.Value.Note),
            After = op.Action == NoteAction.Update ? op.Text : string.Empty,
            Context = anchor.Path,
            BlastRadius = 1,
            Capability = Capability.Deterministic
        });
    }

    private void ApplyUpdate(ApplyContext context, MainDocumentPart main, NoteOp op, NodeAnchor anchor)
    {
        var (note, _) = NoteNodeProvider.Locate(main, anchor.Path)
            ?? throw new InvalidOperationException($"Note '{anchor.Path}' vanished before apply.");

        var paragraph = note.Elements<Paragraph>().FirstOrDefault()
            ?? note.AppendChild(new Paragraph());

        var marker = WordRevisionMarker.IsTracked(op.Mode)
            ? new WordRevisionMarker(context.Package, _clock)
            : null;

        // The reference mark run is the note's number; it survives a rewrite of the text.
        var textRuns = paragraph.Elements<Run>().Where(r => !IsReferenceMark(r)).ToList();

        if (marker is null)
            foreach (var run in textRuns) run.Remove();
        else
            foreach (var run in textRuns) marker.MarkRunDeleted(run);

        var replacement = paragraph.AppendChild(
            new Run(new Text(" " + op.Text) { Space = SpaceProcessingModeValues.Preserve }));
        marker?.WrapInserted(replacement);
    }

    private void ApplyRemove(ApplyContext context, MainDocumentPart main, NoteOp op, NodeAnchor anchor)
    {
        var (note, kind) = NoteNodeProvider.Locate(main, anchor.Path)
            ?? throw new InvalidOperationException($"Note '{anchor.Path}' vanished before apply.");

        var id = NoteNodeProvider.IdOf(note)!;
        var references = ReferencesTo(main, kind, id).ToList();

        if (WordRevisionMarker.IsTracked(op.Mode))
        {
            // The note and its reference both stay until a reviewer accepts: a rejected
            // deletion has to put the note back, numbering and all.
            var marker = new WordRevisionMarker(context.Package, _clock);
            foreach (var reference in references)
                if (reference.Parent is Run run) marker.MarkRunDeleted(run);
            foreach (var paragraph in note.Elements<Paragraph>())
                marker.MarkContentDeleted(paragraph);
            return;
        }

        foreach (var reference in references)
        {
            var run = reference.Parent as Run;
            reference.Remove();
            if (run is not null && !run.HasChildren) run.Remove();
        }

        note.Remove();
    }

    /// <summary>Every reference in the document body pointing at the given note.</summary>
    private static IEnumerable<OpenXmlElement> ReferencesTo(MainDocumentPart main, NoteKind kind, string id)
    {
        if (main.Document?.Body is not { } body) yield break;

        if (kind == NoteKind.Footnote)
        {
            foreach (var reference in body.Descendants<FootnoteReference>().ToList())
                if (string.Equals(reference.Id?.Value.ToString(), id, StringComparison.Ordinal))
                    yield return reference;
        }
        else
        {
            foreach (var reference in body.Descendants<EndnoteReference>().ToList())
                if (string.Equals(reference.Id?.Value.ToString(), id, StringComparison.Ordinal))
                    yield return reference;
        }
    }

    // ── Markup ───────────────────────────────────────────────────────────

    private static Run ReferenceRun(NoteKind kind, long id)
    {
        var properties = new RunProperties(new RunStyle { Val = WordNoteStyles.ReferenceStyleId(kind) });

        return kind == NoteKind.Footnote
            ? new Run(properties, new FootnoteReference { Id = id })
            : new Run(properties, new EndnoteReference { Id = id });
    }

    private static OpenXmlElement BuildNoteBody(NoteOp op, long id)
    {
        var paragraph = new Paragraph(
            new ParagraphProperties(new ParagraphStyleId
            {
                Val = op.StyleId ?? WordNoteStyles.TextStyleId(op.Kind)
            }),
            new Run(
                new RunProperties(new RunStyle { Val = WordNoteStyles.ReferenceStyleId(op.Kind) }),
                op.Kind == NoteKind.Footnote
                    ? new FootnoteReferenceMark()
                    : (OpenXmlElement)new EndnoteReferenceMark()),
            new Run(new Text(" " + op.Text) { Space = SpaceProcessingModeValues.Preserve }));

        return op.Kind == NoteKind.Footnote
            ? new Footnote(paragraph) { Id = id }
            : new Endnote(paragraph) { Id = id };
    }

    private static bool IsReferenceMark(Run run) =>
        run.GetFirstChild<FootnoteReferenceMark>() is not null ||
        run.GetFirstChild<EndnoteReferenceMark>() is not null;

    /// <summary>
    /// The notes root, created with the separator entries when the document has no notes
    /// part yet. Word treats a notes part without them as corrupt.
    /// </summary>
    private static OpenXmlCompositeElement EnsureNotesRoot(MainDocumentPart main, NoteKind kind)
    {
        if (kind == NoteKind.Footnote)
        {
            var part = main.FootnotesPart ?? main.AddNewPart<FootnotesPart>();
            if (part.Footnotes is null)
            {
                part.Footnotes = new Footnotes(
                    Separator(new Footnote { Id = -1, Type = FootnoteEndnoteValues.Separator }, new SeparatorMark()),
                    Separator(new Footnote { Id = 0, Type = FootnoteEndnoteValues.ContinuationSeparator }, new ContinuationSeparatorMark()));
            }
            return part.Footnotes;
        }

        var endnotesPart = main.EndnotesPart ?? main.AddNewPart<EndnotesPart>();
        if (endnotesPart.Endnotes is null)
        {
            endnotesPart.Endnotes = new Endnotes(
                Separator(new Endnote { Id = -1, Type = FootnoteEndnoteValues.Separator }, new SeparatorMark()),
                Separator(new Endnote { Id = 0, Type = FootnoteEndnoteValues.ContinuationSeparator }, new ContinuationSeparatorMark()));
        }
        return endnotesPart.Endnotes;
    }

    private static T Separator<T>(T note, OpenXmlElement mark) where T : OpenXmlCompositeElement
    {
        note.AppendChild(new Paragraph(
            new ParagraphProperties(new SpacingBetweenLines
            {
                After = "0",
                Line = "240",
                LineRule = LineSpacingRuleValues.Auto
            }),
            new Run(mark)));
        return note;
    }

    /// <summary>
    /// The next free id. Word reserves the negative ids and zero for the separators, so a
    /// document's first real note is 1 however the part was built.
    /// </summary>
    private static long NextNoteId(OpenXmlCompositeElement notes, NoteKind kind)
    {
        var ids = kind == NoteKind.Footnote
            ? notes.Elements<Footnote>().Select(n => n.Id?.Value ?? 0)
            : notes.Elements<Endnote>().Select(n => n.Id?.Value ?? 0);

        return Math.Max(0, ids.DefaultIfEmpty(0).Max()) + 1;
    }
}

/// <summary>
/// The two styles a note names. Defining them matters for the same reason the blank
/// document ships a style table: <c>w:pStyle</c> and <c>w:rStyle</c> are references, and a
/// reference with no definition renders as ordinary body text - a note that reads as a
/// stray sentence at the foot of the page rather than as a note.
/// </summary>
internal static class WordNoteStyles
{
    public static string TextStyleId(NoteKind kind) =>
        kind == NoteKind.Footnote ? "FootnoteText" : "EndnoteText";

    public static string ReferenceStyleId(NoteKind kind) =>
        kind == NoteKind.Footnote ? "FootnoteReference" : "EndnoteReference";

    /// <summary>Adds the note styles the document does not already define. Idempotent.</summary>
    public static void Ensure(MainDocumentPart main, NoteKind kind)
    {
        var part = main.StyleDefinitionsPart ?? main.AddNewPart<StyleDefinitionsPart>();
        var styles = part.Styles ??= new Styles();

        var textId = TextStyleId(kind);
        if (!Defines(styles, textId))
            styles.AppendChild(new Style(
                new StyleName { Val = kind == NoteKind.Footnote ? "footnote text" : "endnote text" },
                new BasedOn { Val = "Normal" },
                new SemiHidden(),
                new UnhideWhenUsed(),
                new StyleParagraphProperties(
                    new SpacingBetweenLines { After = "0", Line = "240", LineRule = LineSpacingRuleValues.Auto }),
                new StyleRunProperties(new FontSize { Val = "20" }, new FontSizeComplexScript { Val = "20" }))
            {
                Type = StyleValues.Paragraph,
                StyleId = textId
            });

        var referenceId = ReferenceStyleId(kind);
        if (!Defines(styles, referenceId))
            styles.AppendChild(new Style(
                new StyleName { Val = kind == NoteKind.Footnote ? "footnote reference" : "endnote reference" },
                new BasedOn { Val = "DefaultParagraphFont" },
                new SemiHidden(),
                new UnhideWhenUsed(),
                new StyleRunProperties(new VerticalTextAlignment { Val = VerticalPositionValues.Superscript }))
            {
                Type = StyleValues.Character,
                StyleId = referenceId
            });
    }

    private static bool Defines(Styles styles, string styleId) =>
        styles.Elements<Style>().Any(s => string.Equals(s.StyleId?.Value, styleId, StringComparison.Ordinal));
}
