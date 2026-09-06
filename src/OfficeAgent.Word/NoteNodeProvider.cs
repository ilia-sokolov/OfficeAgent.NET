using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Surfaces each footnote and endnote as an addressable node, so a plan can rewrite or
/// remove one without guessing which paragraph in the notes part belongs to which
/// reference.
/// </summary>
/// <remarks>
/// The separator notes - the ids Word reserves for the rule drawn above the notes and its
/// continuation - are deliberately not surfaced. They are document furniture, they carry
/// no text, and an agent offered them as targets will eventually edit one.
/// </remarks>
internal sealed class NoteNodeProvider : IWordNodeProvider
{
    /// <inheritdoc />
    public string Kind => "note";

    /// <inheritdoc />
    public IEnumerable<NodeInfo> Enumerate(WordObjectMap map)
    {
        foreach (var (note, kind) in Notes(map.Main))
        {
            var id = IdOf(note);
            if (id is null) continue;

            var text = TextOf(note);
            var path = PathOf(kind, id);

            yield return new NodeInfo
            {
                Kind = Kind,
                Path = path,
                Summary = $"{Word(kind)} {id}: {Truncate(text)}",
                Anchor = new NodeAnchor { Id = $"note:{path}", Kind = Kind, Path = path, Expect = text }
            };
        }
    }

    /// <inheritdoc />
    public ResolvedNode? Resolve(NodeAnchor anchor, WordObjectMap map)
    {
        var located = Locate(map.Main, anchor.Path);
        return located is null
            ? null
            : new ResolvedNode
            {
                Kind = Kind,
                Elements = new OpenXmlElement[] { located.Value.Note },
                Value = TextOf(located.Value.Note)
            };
    }

    // ── Shared with the handler ──────────────────────────────────────────

    public static string PathOf(NoteKind kind, string id) => $"{Word(kind)}#{id}";

    /// <summary>Every real note in the document, in footnote-then-endnote order.</summary>
    public static IEnumerable<(OpenXmlElement Note, NoteKind Kind)> Notes(MainDocumentPart main)
    {
        if (main.FootnotesPart?.Footnotes is { } footnotes)
            foreach (var note in footnotes.Elements<Footnote>())
                if (!IsSeparator(note.Type)) yield return (note, NoteKind.Footnote);

        if (main.EndnotesPart?.Endnotes is { } endnotes)
            foreach (var note in endnotes.Elements<Endnote>())
                if (!IsSeparator(note.Type)) yield return (note, NoteKind.Endnote);
    }

    /// <summary>Finds the note a <c>footnote#N</c> or <c>endnote#N</c> path names.</summary>
    public static (OpenXmlElement Note, NoteKind Kind)? Locate(MainDocumentPart main, string path)
    {
        var separator = path.IndexOf('#');
        if (separator <= 0) return null;

        var word = path.Substring(0, separator);
        var id = path.Substring(separator + 1);

        NoteKind kind;
        if (string.Equals(word, "footnote", StringComparison.OrdinalIgnoreCase)) kind = NoteKind.Footnote;
        else if (string.Equals(word, "endnote", StringComparison.OrdinalIgnoreCase)) kind = NoteKind.Endnote;
        else return null;

        foreach (var (note, noteKind) in Notes(main))
            if (noteKind == kind && string.Equals(IdOf(note), id, StringComparison.Ordinal))
                return (note, kind);

        return null;
    }

    /// <summary>The ids of the notes the body still points at, by kind.</summary>
    public static HashSet<string> ReferencedIds(MainDocumentPart main)
    {
        var referenced = new HashSet<string>(StringComparer.Ordinal);
        if (main.Document?.Body is not { } body) return referenced;

        foreach (var reference in body.Descendants<FootnoteReference>())
            if (reference.Id?.Value is { } id) referenced.Add($"footnote#{id}");

        foreach (var reference in body.Descendants<EndnoteReference>())
            if (reference.Id?.Value is { } id) referenced.Add($"endnote#{id}");

        return referenced;
    }

    /// <summary>
    /// Deletes the notes that <paramref name="wereReferenced"/> covers and the body no
    /// longer points at.
    /// </summary>
    /// <remarks>
    /// A note is only reachable through its reference, so one whose reference has just been
    /// deleted is invisible in Word and a phantom entry in <c>inspect</c> - an agent would
    /// be offered a note the reader cannot see. Only notes that lost a reference in the
    /// caller's own edit are swept; a note that was already unreferenced when the operation
    /// started is somebody else's, and stays.
    /// </remarks>
    public static void PruneOrphans(MainDocumentPart main, HashSet<string> wereReferenced)
    {
        if (wereReferenced.Count == 0) return;

        var referenced = ReferencedIds(main);
        foreach (var (note, kind) in Notes(main).ToList())
        {
            if (IdOf(note) is not { } id) continue;

            var path = PathOf(kind, id);
            if (wereReferenced.Contains(path) && !referenced.Contains(path))
                note.Remove();
        }
    }

    public static string? IdOf(OpenXmlElement note) => note switch
    {
        Footnote f => f.Id?.Value.ToString(),
        Endnote e => e.Id?.Value.ToString(),
        _ => null
    };

    public static string TextOf(OpenXmlElement note) =>
        string.Concat(note.Descendants<Text>().Select(t => t.Text)).Trim();

    public static string Word(NoteKind kind) => kind == NoteKind.Footnote ? "footnote" : "endnote";

    private static bool IsSeparator(EnumValue<FootnoteEndnoteValues>? type) =>
        type is not null &&
        (type.Value == FootnoteEndnoteValues.Separator ||
         type.Value == FootnoteEndnoteValues.ContinuationSeparator);

    private static string Truncate(string text, int max = 80) =>
        text.Length <= max ? text : text.Substring(0, max) + "…";
}
