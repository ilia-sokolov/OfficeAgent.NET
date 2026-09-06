using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Word;

/// <summary>
/// Writes WordprocessingML revision markup: the <c>w:ins</c>, <c>w:del</c>,
/// <c>w:rPrChange</c> and friends that turn an edit into something a reviewer opens in
/// Word and accepts or rejects.
/// </summary>
/// <remarks>
/// <para>
/// One marker per apply, so every revision a plan writes shares an author and a timestamp
/// and Word groups them as one round of edits rather than as a scatter of unrelated ones.
/// </para>
/// <para>
/// A redline is not decoration. An inserted paragraph whose <em>mark</em> is not marked
/// inserted rejects to a blank paragraph rather than to nothing; a deleted row whose cells
/// still hold live runs rejects to a row of empty cells. Each helper here marks the whole
/// structure, which is why handlers call these rather than wrapping a run and moving on.
/// </para>
/// </remarks>
internal sealed class WordRevisionMarker
{
    /// <summary>The author recorded on revisions this engine writes.</summary>
    public const string DefaultAuthor = "OfficeAgent";

    private readonly WordRevisionIdAllocator _ids;
    private readonly DateTime _stamp;

    /// <summary>Gets the author name stamped on every revision this marker writes.</summary>
    public string Author { get; }

    /// <summary>Initializes a marker over an open package.</summary>
    public WordRevisionMarker(IOpenXmlPackage package, TimeProvider clock, string author = DefaultAuthor)
    {
        _ids = new WordRevisionIdAllocator(package);
        _stamp = clock.GetUtcNow().UtcDateTime;
        Author = author;
    }

    /// <summary>
    /// Resolves an operation's change mode for the Word module. Null means the caller left
    /// it to the module, and the module tracks - the same default
    /// <see cref="ChangeTextOp.Mode"/> has always carried, now applied to every verb that
    /// changes content.
    /// </summary>
    public static bool IsTracked(ChangeMode? mode) => mode is not ChangeMode.Direct;

    private T Stamp<T>(T element) where T : OpenXmlElement =>
        WordRevisions.Stamp(element, _ids.Next().ToString(), Author, _stamp);

    // ── Runs ─────────────────────────────────────────────────────────────

    /// <summary>Wraps a run in <c>w:ins</c> in place, and returns the wrapper.</summary>
    public InsertedRun WrapInserted(OpenXmlElement run) => Wrap(run, Stamp(new InsertedRun()));

    /// <summary>
    /// Marks every live run under <paramref name="container"/> deleted: each
    /// <c>w:t</c> becomes a <c>w:delText</c> and each run moves inside a <c>w:del</c>.
    /// </summary>
    /// <remarks>
    /// A run already inside a <c>w:del</c> is left alone - deleting a deletion twice is not
    /// a thing Word can represent. A run inside a <c>w:ins</c> is wrapped where it stands,
    /// producing the <c>w:ins/w:del</c> nesting Word writes when someone deletes text that
    /// an earlier, still-pending revision inserted.
    /// </remarks>
    public void MarkContentDeleted(OpenXmlElement container)
    {
        foreach (var run in container.Descendants<Run>().ToList())
            MarkRunDeleted(run);
    }

    /// <summary>Marks one run deleted, leaving everything beside it in the paragraph alone.</summary>
    public void MarkRunDeleted(Run run)
    {
        if (run.Ancestors<DeletedRun>().Any()) return;
        ToDeletedText(run);
        Wrap(run, Stamp(new DeletedRun()));
    }

    /// <summary>Marks every run under <paramref name="container"/> as an insertion.</summary>
    public void MarkContentInserted(OpenXmlElement container)
    {
        foreach (var run in container.Descendants<Run>().ToList())
        {
            if (run.Ancestors<InsertedRun>().Any() || run.Ancestors<DeletedRun>().Any()) continue;
            Wrap(run, Stamp(new InsertedRun()));
        }
    }

    // ── Paragraphs ───────────────────────────────────────────────────────

    /// <summary>
    /// Marks a whole paragraph inserted: its runs and its paragraph mark. Rejecting the
    /// revision then removes the paragraph rather than leaving an empty one behind.
    /// </summary>
    public void MarkParagraphInserted(Paragraph paragraph)
    {
        MarkContentInserted(paragraph);
        MarkParagraphMark(paragraph, inserted: true);
    }

    /// <summary>
    /// Marks a whole paragraph deleted: its runs and its paragraph mark. Accepting then
    /// removes the paragraph, rather than merging its emptied remains into the next one.
    /// </summary>
    public void MarkParagraphDeleted(Paragraph paragraph)
    {
        MarkContentDeleted(paragraph);
        MarkParagraphMark(paragraph, inserted: false);
    }

    /// <summary>
    /// Marks the paragraph mark - the pilcrow, which is what actually ends a paragraph -
    /// inserted or deleted, in <c>w:pPr/w:rPr</c>.
    /// </summary>
    public void MarkParagraphMark(Paragraph paragraph, bool inserted)
    {
        var properties = paragraph.ParagraphProperties ??= new ParagraphProperties();
        var mark = properties.ParagraphMarkRunProperties ??= new ParagraphMarkRunProperties();

        mark.RemoveAllChildren<Inserted>();
        mark.RemoveAllChildren<Deleted>();

        // CT_ParaRPr opens with ins/del/moveFrom/moveTo; everything that styles the mark
        // follows them.
        mark.InsertAt<OpenXmlElement>(
            inserted ? Stamp(new Inserted()) : Stamp(new Deleted()), 0);
    }

    // ── Tables ───────────────────────────────────────────────────────────

    /// <summary>Marks a table and every row in it inserted.</summary>
    public void MarkTableInserted(Table table)
    {
        foreach (var row in table.Elements<TableRow>().ToList())
            MarkRowInserted(row);
    }

    /// <summary>Marks a table and every row in it deleted.</summary>
    public void MarkTableDeleted(Table table)
    {
        foreach (var row in table.Elements<TableRow>().ToList())
            MarkRowDeleted(row);
    }

    /// <summary>Marks a table row inserted: the row itself and everything in its cells.</summary>
    public void MarkRowInserted(TableRow row)
    {
        MarkRowProperty(row, inserted: true);
        foreach (var paragraph in row.Descendants<Paragraph>().ToList())
            MarkParagraphInserted(paragraph);
    }

    /// <summary>Marks a table row deleted: the row itself and everything in its cells.</summary>
    public void MarkRowDeleted(TableRow row)
    {
        MarkRowProperty(row, inserted: false);
        foreach (var paragraph in row.Descendants<Paragraph>().ToList())
            MarkParagraphDeleted(paragraph);
    }

    /// <summary>Marks a cell as inserted by a tracked column insertion.</summary>
    public void MarkCellInserted(TableCell cell)
    {
        MarkCellProperty(cell, inserted: true);
        foreach (var paragraph in cell.Descendants<Paragraph>().ToList())
            MarkParagraphInserted(paragraph);
    }

    /// <summary>Marks a cell as deleted by a tracked column deletion.</summary>
    public void MarkCellDeleted(TableCell cell)
    {
        MarkCellProperty(cell, inserted: false);
        foreach (var paragraph in cell.Descendants<Paragraph>().ToList())
            MarkParagraphDeleted(paragraph);
    }

    private void MarkRowProperty(TableRow row, bool inserted)
    {
        var properties = row.TableRowProperties ??= new TableRowProperties();
        properties.RemoveAllChildren<Inserted>();
        properties.RemoveAllChildren<Deleted>();
        PlaceBeforeChange<TableRowPropertiesChange>(
            properties, inserted ? Stamp(new Inserted()) : Stamp(new Deleted()));
    }

    private void MarkCellProperty(TableCell cell, bool inserted)
    {
        var properties = cell.TableCellProperties ??= new TableCellProperties();
        properties.RemoveAllChildren<CellInsertion>();
        properties.RemoveAllChildren<CellDeletion>();
        PlaceBeforeChange<TableCellPropertiesChange>(
            properties, inserted ? Stamp(new CellInsertion()) : Stamp(new CellDeletion()));
    }

    // ── Formatting revisions ─────────────────────────────────────────────

    /// <summary>
    /// Records the run's properties as they were before the caller changed them, as
    /// <c>w:rPrChange</c>. Rejecting the revision restores exactly this.
    /// </summary>
    /// <param name="run">The run whose current <c>w:rPr</c> now holds the new formatting.</param>
    /// <param name="previous">A detached clone of the properties taken before the change, or null when the run had none.</param>
    public void RecordRunPropertiesChange(Run run, RunProperties? previous)
    {
        var properties = run.RunProperties ??= new RunProperties();

        // An rPrChange already there records an older original. Overwriting it would make
        // "reject" restore an intermediate state that never existed in the document.
        if (properties.GetFirstChild<RunPropertiesChange>() is not null) return;

        var change = Stamp(new RunPropertiesChange());
        change.AppendChild(CopyInto(new PreviousRunProperties(), previous));
        properties.AppendChild(change);
    }

    /// <summary>
    /// Records the paragraph's properties as they were before the caller changed them, as
    /// <c>w:pPrChange</c>.
    /// </summary>
    public void RecordParagraphPropertiesChange(Paragraph paragraph, ParagraphProperties? previous)
    {
        var properties = paragraph.ParagraphProperties ??= new ParagraphProperties();
        if (properties.GetFirstChild<ParagraphPropertiesChange>() is not null) return;

        // CT_PPrBase is w:pPr without the paragraph mark's own run properties, its section
        // properties, or a nested change. Cloning those in produces a file Word repairs.
        var change = Stamp(new ParagraphPropertiesChange());
        change.AppendChild(CopyInto(
            new PreviousParagraphProperties(),
            previous,
            child => child is not ParagraphMarkRunProperties
                  && child is not SectionProperties
                  && child is not ParagraphPropertiesChange));
        properties.AppendChild(change);
    }

    /// <summary>Records a table's previous properties as <c>w:tblPrChange</c>.</summary>
    public void RecordTablePropertiesChange(Table table, TableProperties? previous)
    {
        var properties = table.GetFirstChild<TableProperties>();
        if (properties is null || properties.GetFirstChild<TablePropertiesChange>() is not null) return;

        var change = Stamp(new TablePropertiesChange());
        change.AppendChild(CopyInto(new PreviousTableProperties(), previous));
        properties.AppendChild(change);
    }

    /// <summary>Records a row's previous properties as <c>w:trPrChange</c>.</summary>
    public void RecordRowPropertiesChange(TableRow row, TableRowProperties? previous)
    {
        var properties = row.TableRowProperties;
        if (properties is null || properties.GetFirstChild<TableRowPropertiesChange>() is not null) return;

        var change = Stamp(new TableRowPropertiesChange());
        change.AppendChild(CopyInto(new PreviousTableRowProperties(), previous));
        properties.AppendChild(change);
    }

    /// <summary>Records a cell's previous properties as <c>w:tcPrChange</c>.</summary>
    public void RecordCellPropertiesChange(TableCell cell, TableCellProperties? previous)
    {
        var properties = cell.TableCellProperties;
        if (properties is null || properties.GetFirstChild<TableCellPropertiesChange>() is not null) return;

        var change = Stamp(new TableCellPropertiesChange());
        change.AppendChild(CopyInto(new PreviousTableCellProperties(), previous));
        properties.AppendChild(change);
    }

    // ── Shared mechanics ─────────────────────────────────────────────────

    /// <summary>Moves <paramref name="element"/> inside <paramref name="wrapper"/>, in place.</summary>
    private static T Wrap<T>(OpenXmlElement element, T wrapper) where T : OpenXmlElement
    {
        var parent = element.Parent
            ?? throw new InvalidOperationException("Cannot mark a detached element as a revision.");
        parent.InsertBefore(wrapper, element);
        element.Remove();
        wrapper.AppendChild(element);
        return wrapper;
    }

    /// <summary>Rewrites a run's <c>w:t</c> as <c>w:delText</c>, which is how deleted text is stored.</summary>
    private static void ToDeletedText(Run run)
    {
        foreach (var text in run.Elements<Text>().ToList())
        {
            var deleted = new DeletedText(text.Text ?? string.Empty)
            {
                Space = SpaceProcessingModeValues.Preserve
            };
            text.InsertAfterSelf(deleted);
            text.Remove();
        }
    }

    /// <summary>
    /// Places a revision marker ahead of the properties container's trailing
    /// <c>*PrChange</c>, which the schema declares last.
    /// </summary>
    private static void PlaceBeforeChange<TChange>(OpenXmlElement properties, OpenXmlElement marker)
        where TChange : OpenXmlElement
    {
        if (properties.GetFirstChild<TChange>() is { } change)
            properties.InsertBefore(marker, change);
        else
            properties.AppendChild(marker);
    }

    private static T CopyInto<T>(
        T destination,
        OpenXmlElement? source,
        Func<OpenXmlElement, bool>? keep = null)
        where T : OpenXmlElement
    {
        if (source is null) return destination;
        foreach (var child in source.ChildElements)
            if (keep is null || keep(child))
                destination.AppendChild(child.CloneNode(deep: true));
        return destination;
    }
}

/// <summary>
/// The revision vocabulary itself: which elements are revision markers, what each one is
/// called, and how their shared <c>w:id</c>/<c>w:author</c>/<c>w:date</c> are read and
/// written.
/// </summary>
/// <remarks>
/// WordprocessingML gives every marker the same three attributes but no common base class:
/// the run wrappers derive from <c>RunTrackChangeType</c>, the property-level markers from
/// <c>TrackChangeType</c>, and each <c>*PrChange</c> declares its own. Going through the
/// attributes rather than a type hierarchy keeps one definition of what a revision is, for
/// the writer, the id allocator, and the reader alike.
/// </remarks>
internal static class WordRevisions
{
    /// <summary>The WordprocessingML main namespace, which every revision attribute lives in.</summary>
    public const string W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";

    /// <summary>
    /// The agent-facing name for a revision marker, or null when the element is not one.
    /// <c>w:ins</c> and <c>w:del</c> mean different things in different parents - a run, a
    /// paragraph mark, a table row - so the name is read from where the marker sits.
    /// </summary>
    public static string? TagOf(OpenXmlElement element) => element switch
    {
        InsertedRun => "ins",
        DeletedRun => "del",
        MoveFromRun => "moveFrom",
        MoveToRun => "moveTo",
        Inserted when element.Parent is ParagraphMarkRunProperties => "markIns",
        Deleted when element.Parent is ParagraphMarkRunProperties => "markDel",
        Inserted when element.Parent is TableRowProperties => "rowIns",
        Deleted when element.Parent is TableRowProperties => "rowDel",
        CellInsertion => "cellIns",
        CellDeletion => "cellDel",
        RunPropertiesChange => "runFormat",
        ParagraphPropertiesChange => "paraFormat",
        TablePropertiesChange => "tableFormat",
        TableRowPropertiesChange => "rowFormat",
        TableCellPropertiesChange => "cellFormat",
        _ => null
    };

    /// <summary>Writes the three attributes every revision marker carries.</summary>
    public static T Stamp<T>(T element, string id, string author, DateTime date) where T : OpenXmlElement
    {
        element.SetAttribute(new OpenXmlAttribute("w", "id", W, id));
        element.SetAttribute(new OpenXmlAttribute("w", "author", W, author));
        element.SetAttribute(new OpenXmlAttribute("w", "date", W, date.ToString("yyyy-MM-ddTHH:mm:ss'Z'")));
        return element;
    }

    public static string? IdOf(OpenXmlElement element) => Attribute(element, "id");

    public static string? AuthorOf(OpenXmlElement element) => Attribute(element, "author");

    public static string? DateOf(OpenXmlElement element) => Attribute(element, "date");

    private static string? Attribute(OpenXmlElement element, string localName)
    {
        foreach (var attribute in element.GetAttributes())
            if (string.Equals(attribute.LocalName, localName, StringComparison.Ordinal) &&
                string.Equals(attribute.NamespaceUri, W, StringComparison.Ordinal))
                return attribute.Value;
        return null;
    }
}
