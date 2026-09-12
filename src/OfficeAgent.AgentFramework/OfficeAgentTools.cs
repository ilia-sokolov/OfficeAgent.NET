using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using Microsoft.Extensions.AI;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.AgentFramework;

/// <summary>
/// Controls which tools <see cref="OfficeAgentTools.AsAIFunctions(OfficeAgentToolsOptions)"/>
/// exposes to the agent.
/// </summary>
public sealed class OfficeAgentToolsOptions
{
    /// <summary>
    /// Gets whether the agent may register documents with provider connections and
    /// remove registrations (<c>register_document</c>, <c>remove_document</c>,
    /// <c>open_document</c>, and <c>edit_document</c>). The
    /// default is <see langword="false"/>: the host pre-registers documents and the
    /// agent only ever sees opaque ids. Enabling this lets the agent hand
    /// provider-relative sources (a path under a filesystem root, or a SharePoint URL
    /// or <c>driveId/itemId</c> pair) to the configured connections; the connection boundary,
    /// extension allow-list, and size limits still apply, and removing a
    /// registration never deletes the underlying content. Creating new documents is a
    /// separate opt-in; see <see cref="AllowCreation"/>.
    /// </summary>
    public bool AllowRegistration { get; init; }

    /// <summary>
    /// Gets whether <c>create_document</c> is exposed. The default is <see langword="false"/>:
    /// authoring a brand-new, agent-named file at a connection root is a capability the
    /// host should choose deliberately, and upgrading the package must not grant it to a
    /// host that only ever opted into registration. Turning it on adds document creation
    /// only - registration and editing are unchanged.
    /// </summary>
    public bool AllowCreation { get; init; }

    /// <summary>
    /// Gets whether the inline-content tools are exposed: <c>create_document_content</c>,
    /// <c>inspect_document_content</c>, and <c>edit_document_content</c>, which take the
    /// document as base64 and hand the edited document straight back. They need no
    /// provider connection at all, which is the point - a host with no storage to offer,
    /// or one whose documents arrive as attachments, can still create and edit.
    /// </summary>
    /// <remarks>
    /// The default is <see langword="false"/>, and deliberately so. These tools move the
    /// whole package through the model's context in both directions: the document is
    /// spelled out in the request, the edited document in the response. That costs tokens
    /// in proportion to file size and puts the complete file - not just the text the
    /// inspect tools already return - in front of the model provider. A host that
    /// configured a storage connection precisely so that documents never travel that way
    /// should leave this off.
    /// </remarks>
    public bool AllowInlineContent { get; init; }

    /// <summary>
    /// Gets whether the tools that address documents by <c>(connectionId, documentId)</c>
    /// are exposed at all - the four core tools, and the registration and creation tools
    /// the two options above gate. The default is <see langword="true"/>.
    /// </summary>
    /// <remarks>
    /// A host with no provider connections configured sets this to false, so an agent is
    /// not handed <c>inspect_document</c> and friends when there is no connection any call
    /// to them could name. The same reasoning as the plan's verb map: a tool that can only
    /// fail is worse than no tool, because the agent spends a turn discovering that.
    /// </remarks>
    public bool AllowConnectionAddressing { get; init; } = true;

    /// <summary>
    /// Gets whether <c>import_document_content</c> and <c>export_document_content</c> are
    /// exposed: the two tools that move bytes into and out of a connection whose documents
    /// this process holds for the session.
    /// </summary>
    /// <remarks>
    /// They exist because an ephemeral connection has no path or URL to name a document
    /// by, so content has to enter and leave some other way. Everything between those two
    /// points is the ordinary connection-addressed surface, working on a short opaque id -
    /// which is what makes multi-step editing reliable, where passing the whole document
    /// back and forth is not.
    /// </remarks>
    public bool AllowEphemeralDocuments { get; init; }
}

/// <summary>
/// Projects <see cref="OfficeAgentClient"/> as Microsoft.Extensions.AI tools that
/// address documents by an opaque, provider-assigned id. The host registers documents
/// with a provider connection (<see cref="OfficeAgentClient.RegisterAsync"/>),
/// receives a <see cref="DocumentReference"/>, and the LLM drives inspect /
/// find / preview / apply by <c>(connectionId, documentId)</c>. The agent never
/// sees credentials or absolute storage locations; by default it cannot register
/// documents, escape the connection, or delete content the provider only
/// references. Hosts that want the agent to manage its own registrations opt in
/// via <see cref="OfficeAgentToolsOptions.AllowRegistration"/>; creation is a
/// separate opt-in through <see cref="OfficeAgentToolsOptions.AllowCreation"/>.
/// </summary>
public sealed class OfficeAgentTools
{
    private const string RegexTimeoutMessage =
        "The regular expression exceeded the search time limit. Use a simpler pattern or a literal search.";

    private static readonly JsonSerializerOptions Json = new() { WriteIndented = false };

    private static readonly JsonSerializerOptions PlanJson = new()
    {
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        PropertyNameCaseInsensitive = true,
        Converters = { new JsonStringEnumConverter() }
    };

    internal static readonly AIJsonSchemaCreateOptions StrictSchemaOptions = new()
    {
        TransformOptions = new AIJsonSchemaTransformOptions
        {
            RequireAllProperties = true,
            DisallowAdditionalProperties = true
        }
    };

    private readonly OfficeAgentClient _client;
    private readonly IConnectionAccessPolicy _connectionAccess;
    private readonly ITrustedPrincipalAccessor _principalAccessor;

    /// <summary>Initializes the tool projection over an <see cref="OfficeAgentClient"/>.</summary>
    public OfficeAgentTools(
        OfficeAgentClient client,
        IConnectionAccessPolicy? connectionAccess = null,
        ITrustedPrincipalAccessor? principalAccessor = null)
    {
        _client = client ?? throw new ArgumentNullException(nameof(client));
        _connectionAccess = connectionAccess ?? new AllowAllConnectionAccessPolicy();
        _principalAccessor = principalAccessor ?? new AnonymousPrincipalAccessor();
    }

    /// <summary>Tests a capability for connection discovery without disclosing a denial.</summary>
    public ValueTask<bool> CanAccessConnectionAsync(
        string connectionId,
        ConnectionCapability capability,
        CancellationToken cancellationToken = default) =>
        _connectionAccess.IsAllowedAsync(
            _principalAccessor.Principal, connectionId, capability, cancellationToken);

    /// <summary>
    /// System-prompt guidance to concatenate into the agent's instructions. Teaches
    /// the model the <c>(connectionId, documentId)</c> contract, the safety loop,
    /// and the structured-error vocabulary the tools surface.
    /// </summary>
    public const string SystemPromptGuidance = """
        You are editing Microsoft Office documents through the OfficeAgent tools - a Word document (.docx), PowerPoint deck (.pptx), or Excel workbook (.xlsx). inspect_document reports which format it found.

        Document addressing
        - Storage connections are host-configured and so is each document's registration. The host gives you an OPAQUE, provider-assigned documentId for every document you are allowed to work with; the connection-addressed document tools address it as (connectionId, documentId). Never invent a documentId, and never pass a filename or path as one. Do not ask the user to paste file bytes into the conversation; if a tool whose name ends in _content is available, that tool - and only that tool - takes the document as base64, and its own section below says how.
        - The connectionId and documentId are already in your instructions or in the conversation context. NEVER ask the user for them - the user does not know or manage these values. If a request mentions "the document", it means the current document you were given; start working with it immediately.
        - apply_plan returns outputDocumentId, outputName, and outputContentType for the saved revision. Use outputDocumentId as the next call's documentId if you keep editing. When the work is complete, tell the user the document is ready; the host retrieves its bytes and presents the download or attachment. Do not place document base64 in the final response.
        - Saving edits the document in place (saveMode "Replace", the default), so outputDocumentId is the same id you passed in. If the user wants the original kept, pass saveMode "NewVersion" to write a sibling revision instead, and tell them where the result landed.

        Plan shape, anchors, safety loop
        - Plan body is { "snapshot": { "eTag": "<snapshot from inspect_document>" }, "operations": [ ... ] }. Copy the scalar snapshot string returned by inspect_document into snapshot.eTag to detect drift in Word body/header/footer/footnote/endnote XML or PowerPoint slide/notes XML. It does not cover properties, comments, sections, media/image bytes, masters, or layouts; their anchors and provider version checks still apply. Omit snapshot only deliberately. Do not set contractVersion.
        - Available operations (the JSON shape of each is in the preview_plan description): Word and PowerPoint use the document operations below; decks additionally support insertChart/updateChart; workbooks support setCell, appendTableRows, and comment Add/Remove on a cell. These are plan operations inside preview_plan/apply_plan, not separate tools.
        - populate_template_batch resolves named slots and repeating Word rows into plans and saves one independent output per item. compare_documents reads two Word documents and returns a snapshot-bound redline plan only when every detected change is covered; preview that plan before applying it to the original.
        - preview_document_merge assembles ordered whole Word documents in memory and returns a separate merge plan plus compatibility diagnostics. Pass that complete plan to merge_documents when available. A merge creates a new document and never edits its sources. It does not reconcile independently edited versions. Respect unsupported-content diagnostics and do not substitute a text-only copy.
        - Call inspect_document or find_in_document before building a plan to obtain anchor ids; never invent paragraph ids, occurrence numbers, content-control tags, or node paths.
        - Tables and images only appear in inspect_document.nodes, never in the paragraphs list. Copy the path from there rather than composing one: Word uses "table#N"/"image#N", a deck uses "table#{slideId}/{shapeId}"/"image#{slideId}/{shapeId}". To recognise table content, look for paragraphs whose `in` field matches a table path.
        - Preview before you apply. If preview reports stale-snapshot, re-inspect and rebuild. If preview reports expect-mismatch, the document drifted - re-inspect/find that operation.
        - Change mode: in a Word document default to "Tracked" unless the user explicitly approves direct edits. "mode" is not only for changeText - it belongs on every verb that changes content (insert, insertParagraphs, removeParagraph, insertTable, removeTable, the table row/column verbs including repeatTableRow, insertImage, removeImage, format, fill, insertBreak, note), and Word tracks all of them when it is omitted. A tracked structural edit is a real redline: an added paragraph and its paragraph mark come back as insertions, a removed row stays in place struck through until someone accepts it, and a format records what it replaced. A PowerPoint deck has no tracked-changes representation and REFUSES mode "Tracked" on any of these, so use "Direct" there and say that edits to a deck cannot be redlined; add a comment if the change needs to be flagged for review.
        - Reviewing an existing Word redline: inspect_document.nodes lists every revision (kind "revision") with its author and date, under paths like "ins#7", "del#7", "markIns#7" (a paragraph split), "rowIns#7"/"rowDel#7", "cellIns#7"/"cellDel#7", and "runFormat#7"/"paraFormat#7" (formatting changes). Resolve them with { "op": "revision", "target": { "kind": "revision", "path": "…" }, "action": "Accept" | "Reject" } - "all" takes every revision in the document and "author:<name>" takes one person's.
        - Word comments are a conversation, not a write-only log: inspect_document.nodes lists each one (kind "comment") with its author, its text, whether it is resolved, and which comment it replies to. Answer one with action "Reply", close it with "Resolve", delete it with "Remove". Read the open comments before editing a document under review - they usually say what the edit should be.
        - Reject operations that need a renderer (pagination, field recalculation); explain the limitation instead.
        - An image behind the content is backgroundImage, not insertImage: { "op": "backgroundImage", "base64Bytes": "iVBORw0KGgo...", "imageType": "png", "opacity": 0.2 }. On a deck a slide target paints one slide and no target paints every slide; in Word there is no target and the image repeats on every page. ALWAYS set opacity for a photograph behind text - at full strength almost any image destroys the contrast the text needs, and 0.1-0.3 is the usable range. Supply no image at all to take an existing background away. A flat colour is format with fillColor instead, and a picture IN the text is still insertImage.
        - Real list numbering (Word only): { "op": "format", "target": {…}, "listStyle": "clause", "listLevel": 0, "listId": 0 }. listStyle is bullet, decimal (1./a./i.), clause (1./1.1/1.1.1) or none. listLevel is 0-8. NEVER type the number into the text as well - Word draws it, and a paragraph reading "1. Connect the drive" in a numbered list comes out as "1. 1. Connect the drive". Paragraphs sharing a listStyle AND a listId are one running sequence; a different listId starts a separate one, which is how a second chapter restarts its steps at 1. This is what makes an inserted clause renumber the rest.
        - Word page geometry: { "op": "pageSetup", "paperSize": "A4", "orientation": "Landscape", "marginTopTwips": 720 }. Measurements are twips - 1440 to the inch - matching the indents and spacing on format. paperSize is A3, A4, A5, Letter, Legal, Tabloid or Executive; give pageWidthTwips/pageHeightTwips instead for anything else, never both. No target sets the document's final section - the whole document until something splits it; a paragraph target sets the section that paragraph belongs to. Margins you do not name keep their current value.
        - Word breaks: { "op": "insertBreak", "target": {…}, "kind": "Page", "position": "After" }. kind is Page, Column, SectionNextPage, SectionContinuous, SectionEvenPage or SectionOddPage. The break lands in a paragraph of its own, so removing it later is removing one paragraph. A landscape run of pages in the middle of a portrait report is insertBreak twice and then pageSetup targeting a paragraph between them. Use format's pageBreakBefore instead when the page must keep starting at a given paragraph however the text above it changes.
        - Word footnotes and endnotes: { "op": "note", "target": { "paraId": "w14:…", "expect": "thirty days" }, "kind": "Footnote", "text": "Subject to clause 8.2." } puts the reference right after that text, or at the end of the paragraph when expect is empty. Word owns the numbering, so a note added in the middle renumbers the rest - never type a superscript number into the text yourself. Existing notes appear in inspect_document.nodes as kind "note" with paths "footnote#1"/"endnote#1"; rewrite one with action "Update" or delete it with action "Remove". A note's body is also an ordinary paragraph (location "footnote"/"endnote"), so changeText and format edit it in place.
        - Word running heads: { "op": "headerFooter", "header": "Northwind Traders", "footer": "Confidential", "showPageNumber": true, "alignment": "edges" }. alignment "edges" puts the text left and the page number right on one line; otherwise left/center/right. The page number is a field, so it stays right as the document grows. differentFirstPage:true gives the first page its own header and footer, which is how a cover page keeps the running head off it - then write that page's own with scope "firstPage" (scope is default, firstPage or evenPage). Clear either with an empty string. showSlideNumber/showFooter/dateTime are deck-only and are REFUSED here.

        Working with a PowerPoint deck
        - Each slide is one outline entry. Paragraph ids read "slide{slideId}/shape{shapeId}/p{n}", with "notes/..." for speaker notes and ".../r{row}c{col}/..." inside a table cell.
        - A slide has no text flow, so insertTable, insertImage, and an added comment target the SLIDE - { "kind": "slide", "path": "slide#256" } - not a paragraph. Resolve a comment with { "op": "comment", "action": "Resolve", "target": { "kind": "comment", "path": "comment#256/{id}" } }.
        - Only these verbs work on a deck: changeText, insert, format, fill, copyStyles, clearStyles, insertTable, removeTable, insertTableRows, removeTableRows, insertTableColumns, removeTableColumns, insertImage, removeImage, backgroundImage, insertShape, removeShape, insertMedia, insertChart, updateChart, comment, section, headerFooter, transition, animate, insertSlide, removeSlide, moveSlide, duplicateSlide.
        - Native charts: insertChart targets a slide and takes kind ClusteredColumn, Bar, Line, or Pie, categories, numeric series, title, showLegend, description, and pixel placement. It creates an editable chart with an embedded workbook. updateChart targets a chart node and only edits charts OfficeAgent created.
        - Slide transitions: { "op": "transition", "effect": "push", "direction": "up", "durationMs": 700 }. No target applies it to every slide, which is PowerPoint's "Apply To All"; a slide target sets just that one. effect "none" removes it. Set advanceAfterMs for a self-running deck and advanceOnClick:false to stop clicks skipping ahead.
        - Shape animations target a SHAPE node: { "op": "animate", "target": { "kind": "shape", "path": "shape#257/2" }, "effect": "fade", "kind": "Entrance", "trigger": "OnClick", "durationMs": 600 }. trigger is OnClick, WithPrevious or AfterPrevious and decides where the effect lands in the slide's sequence - a new click step, alongside the previous effect, or straight after it. Effects play in the order you send the operations. effect "none" removes that shape's animations.
        - Available animations: appear, fade, wipe, blinds, checkerboard, circle, diamond, dissolve, plus, randomBar, split, wedge, wheel, box. Fly-in, zoom, grow and motion paths are NOT available - they need interpolated properties rather than a filter - and are refused rather than approximated. Say so plainly if the user asks for one.
        - Footer, slide number and date: { "op": "headerFooter", "footer": "Confidential", "showSlideNumber": true, "showDateTime": true }. No target applies it to every slide, which is PowerPoint's "Apply to All"; a slide target changes just that one, which is how you keep the title slide clean. showFooter:false removes the placeholder rather than blanking it. Omitting dateTime gives an auto-updating date; supplying a string pins it. A slide has NO header - PresentationML puts headers on notes and handout pages only, which is why PowerPoint greys that box out on the Slide tab.
        - Embedded video and audio: { "op": "insertMedia", "target": { "kind": "slide", "path": "slide#257" }, "kind": "Video", "base64Bytes": "...", "mediaType": "mp4", "widthPx": 480, "heightPx": 270 }. mediaType is mp4/m4v/mov/wmv/avi for video and mp3/m4a/wav/wma for audio, and must agree with "kind" or the operation is refused. The bytes travel inside the deck so it still plays when mailed on. posterBase64 sets the frame shown before playback. Clips appear in inspect_document.nodes as kind "media"; remove one with removeShape on the matching shape path.
        - Template slots: a deck has no content controls, so a fillable slot is the SHAPE NAME a template sets (PowerPoint's Selection Pane shows them). They arrive in inspect_document.contentControls with kind "shapeName". Fill one by name: { "op": "fill", "target": { "tag": "ClientName" }, "value": "Northwind Traders" }. A name used on more than one slide comes back as ambiguous-anchor; qualify it as "slide256/ClientName".
        - copyStyles makes one line look like another and clearStyles strips direct formatting so the layout's own styling shows again: { "op": "clearStyles", "target": { "paraId": "slide257/shape3/p0", "expect": "" }, "scope": "all" }. scope is "run", "paragraph" or "all". Only direct a:pPr/a:rPr travel - the layout and master are never touched.
        - Sections are the named slide groups in the thumbnail pane: { "op": "section", "action": "Add", "name": "Financials", "target": { "kind": "slide", "path": "slide#257" } } starts one at that slide. "Rename" and "Remove" target a section node from inspect_document.nodes; removing a section keeps its slides. Sections follow the deck automatically as slides are added, moved, copied and removed - you never maintain them by hand.
        - Slides: insertSlide adds one, and several in one plan is how you author a whole deck. { "op": "insertSlide", "slide": { "layout": "titleAndContent", "title": "FY27 Priorities", "body": ["Finish the migration", "Rebuild the pipeline"], "notes": "Do not commit to a date." } }. Layouts are title, titleAndContent, sectionHeader, titleOnly, blank; omit "layout" and one is chosen from what you supply. Position defaults to the end of the deck - use "position": "Start"/"Before"/"After" with a slide target to place it elsewhere.
        - removeSlide, moveSlide and duplicateSlide take a slide target: { "op": "moveSlide", "target": { "kind": "slide", "path": "slide#259" }, "position": "After", "relativeTo": "slide#256" }. duplicateSlide defaults to landing right after the original. Removing the deck's only slide is refused, because PowerPoint cannot open a deck with none.
        - A slide added in one plan cannot be edited by a later operation in that SAME plan - its id does not exist until the plan is applied. Set its text through insertSlide's own title/body/notes, or apply, re-inspect, then edit.
        - format on a deck covers bold, italic, underline, sizeHalfPoints, fontFamily, color, highlight and alignment on text; widthPx/heightPx on an image; xPx/yPx/widthPx/heightPx, fillColor, lineColor, lineWidthPx and verticalAlignment on a shape; and fillColor on a slide. Word-only measures (styleId, indents, spacing, borders) are refused rather than ignored. Anchor the span you want styled, or use an empty "expect" to style a whole paragraph.
        - To write into an empty placeholder - the state a newly created deck's title is in - use changeText with an empty expect: { "op": "changeText", "target": { "paraId": "slide256/shape2/p0", "expect": "" }, "with": "Quarterly Review", "mode": "Direct" }. That still verifies the paragraph is blank, so it fails rather than overwriting text that drifted in.
        - Add a bullet or line to text that is already there with insert, targeting the paragraph it goes next to: { "op": "insert", "target": { "paraId": "slide257/shape3/p1", "expect": "Rebuild the pipeline" }, "position": "After", "text": "Hold headcount flat", "level": 1 }. It inherits the neighbour's bullet and run styling; "level" (0-8) makes it a sub-bullet. styleId is Word-only and refused here.
        - IMPORTANT: a slide paragraph id is positional, so inserting renumbers every later paragraph in the SAME shape. A plan that inserts and then addresses that shape at the same or a higher p-index is refused with operation-conflict. Apply the insert, re-inspect, then send the rest as a second plan. Earlier paragraphs, other shapes and other slides are unaffected.
        - Shapes: insertShape adds a free-standing text box to a slide - { "op": "insertShape", "target": { "kind": "slide", "path": "slide#257" }, "text": ["Draft"], "xPx": 40, "yPx": 620, "widthPx": 420, "heightPx": 50 }. Text belonging in the title or body should go through the placeholders instead. removeShape deletes any shape by its { "kind": "shape", "path": "shape#{slideId}/{shapeId}" } node; removing a placeholder is refused because the layout would re-offer it empty and the slide would look unchanged.
        - Move, resize or paint ANY shape - text box, table, picture - with format on its shape node: { "op": "format", "target": { "kind": "shape", "path": "shape#257/4" }, "xPx": 120, "yPx": 560, "widthPx": 700, "heightPx": 44, "fillColor": "FFF2CC", "lineColor": "7F6000" }. Shape formatting does not style the text inside it; target a paragraph for that.

        Working with an Excel workbook
        - inspect_document returns worksheets and up to maximumCells populated cells. Each cell anchor carries a durable sheetId plus an A1 address. Pass sheetId and range to inspect only the needed rectangle.
        - setCell writes a scalar or a formula: { "op": "setCell", "target": { "sheetId": 7, "address": "B2" }, "value": "42" } or use "formula": "SUM(B2:B8)". OfficeAgent clears the cached result and asks Excel to recalculate on open; it does not calculate formulas.
        - appendTableRows targets a spreadsheetTable node from inspection and requires one value per table column. It refuses to overwrite populated cells below the table and preserves other worksheet content.
        - Excel comments are legacy cell notes. Add one with a cell target and action Add; remove one with the cellComment node returned by inspection and action Remove.
        """;

    /// <summary>
    /// System-prompt guidance to append to <see cref="SystemPromptGuidance"/> when the
    /// host enables <see cref="OfficeAgentToolsOptions.AllowRegistration"/>.
    /// </summary>
    public const string RegistrationPromptGuidance = """

        Document registration
        - register_document(connectionId, source) registers an existing document with a host-configured connection and returns its opaque documentId. The source is connection-specific: a path under a filesystem connection's root, or - for a SharePoint connection - the document's SharePoint/OneDrive URL or a "driveId/itemId" pair. Never pass credentials.
        - remove_document(connectionId, documentId) removes the registration only - the underlying file is never deleted. Remove temporary registrations you made with register_document once the work is done, but keep the final document's output registration until the host has delivered it.
        - Register a document only when the user names a file the host has not already given you an id for; otherwise use the ids you were given.

        Working from a source in one call
        - open_document(connectionId, source) = register_document + inspect_document. Prefer it when the user names a file you have no id for and you need to see the document.
        - edit_document(connectionId, source, planJson) = register_document + find + apply_plan. Prefer it when you already know the text to change; it returns sourceDocumentId for follow-up work.
        - In edit_document a target may name text directly - { "find": "Acme Corp" } - instead of a paraId, so no lookup call is needed. Text matching more than once fails with "ambiguous-anchor" and lists the candidates: re-issue with { "find": "Acme Corp", "match": 2 } (zero-based) or use more surrounding text. Never guess a match index; use the one the error listed.
        - Reach for the single-purpose tools when the composites do not fit: an id you already hold, a plan you want to preview before applying, or targets that need regex or case-sensitive search (find_in_document, then paraId targets).
        """;

    /// <summary>
    /// System-prompt guidance to append when document-creation tools are exposed.
    /// </summary>
    public const string CreationPromptGuidance = """

        Creating a document
        - create_document(connectionId, name, planJson) creates and registers a new document and returns outputDocumentId. name is a bare file name, never a path; an existing name is not overwritten. The extension picks the format: .docx makes Word, .pptx PowerPoint, and .xlsx Excel.
        - populate_template_batch(connectionId, documentId, requestJson) creates bounded, independent outputs from one registered template. Use unique content-control tags or PowerPoint shape names for scalar values; repeating {{Field}} rows are Word-only. Each item returns its own receipt and diagnostics.
        - Pass "" for an empty document. An initial plan is applied in memory before storage. The empty starting anchor differs by format: a Word document has one empty paragraph at { "paraId": "auto-0000", "expect": "" }; a deck has one empty title placeholder at { "paraId": "slide256/shape2/p0", "expect": "" }. When unsure, create with planJson "" and then inspect_document.
        - planJson accepts a bare operations array [ … ] as well as { "operations": [ … ] }.
        - Plan-validation errors mean nothing was written. A provider or cancellation error can occur after storage accepted the file, so do not retry the same name; report the possibly unregistered file name to the host/operator for recovery.
        """;

    /// <summary>
    /// System-prompt guidance to append when the inline-content tools are exposed. It
    /// carries the exception to the document-addressing rules above, because those rules
    /// assume every document has a connection behind it and these tools have none.
    /// </summary>
    public const string InlineContentPromptGuidance = """

        Working on documents passed in as content
        - create_document_content(name, planJson), inspect_document_content(contentBase64), and edit_document_content(contentBase64, planJson) work on a document you hold the bytes of. They take NO connectionId and NO documentId, and nothing is stored: the base64 they return is the only copy of the result.
        - This is the exception to "never put document bytes in the conversation". It applies to THESE tools only: when a connection-addressed tool would do, use it instead, because it does not spend context on the file.
        - The loop is: create_document_content or the caller's own base64 -> (optional) inspect_document_content for anchors -> edit_document_content -> hand the returned contentBase64 to the host. Keep the newest contentBase64 and pass that one to the next edit; an earlier one is a stale document and editing it silently discards the work in between.
        - Batch the work. Every call spends the whole file twice - once going in, once coming back - so put the operations you know about into one plan rather than one call per edit.
        - Targets may name text directly - { "find": "Acme Corp" } - so inspecting first is optional. Text matching more than once fails with "ambiguous-anchor" and lists the candidates; re-issue with { "find": "Acme Corp", "match": 2 } (zero-based). On a large document prefer fidelity "outline" or "structure" when you do inspect.
        - preview=true on edit_document_content validates a plan and returns contentBase64 null. Use it when a plan is speculative; on a plan you are confident of, skip it, because a preview costs another full copy of the file.
        - contentBase64 comes back null whenever there is no document to hand back - a preview, or a plan that failed. Null means nothing was produced: report the errors rather than looking for a document.
        - Tell the user the document is ready and let the host deliver it. Do not paste contentBase64 into your reply to the user; it is for the host, not for reading.
        """;

    /// <summary>
    /// System-prompt guidance to append when a session connection is configured.
    /// </summary>
    public const string EphemeralPromptGuidance = """

        Session documents (a connection whose documents this server holds for you)
        - A session connection behaves like any other: create_document makes a document in it, and inspect_document / find_in_document / preview_plan / apply_plan address it by (connectionId, documentId). The difference is that the bytes never leave the server, so the id is all you ever pass.
        - import_document_content(connectionId, name, contentBase64) puts a document you were given into the session and returns its documentId. export_document_content(connectionId, documentId) hands the finished bytes back for the host to save.
        - PREFER this over the _content tools whenever more than one edit is coming, and whenever the document is more than a page or two. Passing a document back as base64 means reproducing every character of it exactly; on anything sizeable that fails, and it fails as content that is no longer a readable package rather than as an obvious mistake. An id cannot be got wrong.
        - Import once, edit as many times as you need, export once. Do not export between edits: each export spends the whole document in context for nothing.
        - Documents in a session connection are gone when the server stops, and are written to no storage. Export before you finish, or tell the user the result was not saved anywhere.
        """;

    /// <summary>Returns the four core AIFunctions the host registers with its agent.</summary>
    public AIFunction[] AsAIFunctions() => AsAIFunctions(new OfficeAgentToolsOptions());

    /// <summary>
    /// Returns the AIFunctions selected by <paramref name="options"/>: the five core
    /// inspect/find/preview/apply/comparison tools, plus the source-addressed tools
    /// (<c>register_document</c>, <c>remove_document</c>, <c>open_document</c>,
    /// <c>edit_document</c>) when registration is allowed, and independently
    /// <c>create_document</c> when creation is allowed. The composites are gated with
    /// registration because they take the same connection-relative source it does.
    /// </summary>
    public AIFunction[] AsAIFunctions(OfficeAgentToolsOptions options)
    {
        if (options is null) throw new ArgumentNullException(nameof(options));

        var functions = new List<AIFunction>();
        if (options.AllowConnectionAddressing) functions.AddRange(CoreFunctions());

        if (options.AllowRegistration && options.AllowConnectionAddressing)
        {
            functions.Add(AIFunctionFactory.Create(RegisterDocument, Opts(
                "register_document",
                "Register an existing document with a host-configured provider connection and return its opaque documentId. " +
                "source is connection-specific: for a filesystem connection, a path under its root; for a SharePoint connection, the document's SharePoint/OneDrive URL (e.g. 'https://contoso.sharepoint.com/:w:/s/…') or a 'driveId/itemId' pair (e.g. 'b!9a3f…/01ABCDEF'). " +
                "Never pass credentials. Returns {connectionId, documentId, name, contentType, version}.")));
            functions.Add(AIFunctionFactory.Create(RemoveDocument, Opts(
                "remove_document",
                "Remove a document registration from a provider connection by (connectionId, documentId). " +
                "Only the registration is removed - the underlying file is never deleted. Returns {removed, connectionId, documentId}.")));
            functions.Add(AIFunctionFactory.Create(OpenDocument, Opts(
                "open_document",
                "Open a document the user named by source: registers it and returns its inspection in one call - use this instead of register_document followed by inspect_document. " +
                "source is connection-specific, exactly as for register_document: a path under a filesystem connection's root, or a SharePoint/OneDrive URL or 'driveId/itemId' pair. " +
                "Returns {connectionId, documentId, name, contentType, version} followed by the inspect_document payload (snapshot, outline, paragraphs, contentControls, nodes, styles). " +
                "Keep documentId for follow-up calls; paging works as in inspect_document.")));
            functions.Add(AIFunctionFactory.Create(EditDocument, Opts(
                "edit_document",
                "Edit a document the user named by source, in one call: registers it, resolves targets, and applies the operations - use this instead of register_document + find_in_document + apply_plan. " +
                "planJson is an operations array [ … ] or { \"operations\": [ … ] }, the same operations preview_plan documents.\n" +
                "Targets may name text directly instead of a paragraph id, so no lookup call is needed first:\n" +
                "{ \"op\": \"changeText\", \"target\": { \"find\": \"Acme Corp\" }, \"with\": \"Globex Inc.\" }\n" +
                "If that text matches more than once the call fails with 'ambiguous-anchor' and lists each candidate with its context; re-issue with { \"find\": \"Acme Corp\", \"match\": 2 } (zero-based) or use more surrounding text. Text that matches nothing fails with 'anchor-not-found'. " +
                "Anchors resolved from inspect_document/find_in_document ({ \"paraId\": …, \"expect\": … }) work here too, and can be mixed in the same plan. " +
                "saveMode and newName behave as in apply_plan. Nothing is written unless every operation validates. " +
                "Returns the apply_plan shape plus sourceDocumentId - the id of the document that was opened, usable for follow-up calls even when the edit failed.")));
        }
        if (options.AllowCreation && options.AllowConnectionAddressing)
        {
            functions.Add(AIFunctionFactory.Create(CreateDocument, Opts(
                "create_document",
                "Create and register a new document in a host-configured connection, optionally applying an initial plan before writing. " +
                "name is a bare file name such as 'quarterly-report.docx'; an existing name is never overwritten. " +
                "The extension picks the format: '.docx' makes Word, '.pptx' PowerPoint, and '.xlsx' Excel. " +
                "Pass planJson \"\" for a minimal document. The starting anchor differs by format: a Word document has one empty paragraph at { \"paraId\": \"auto-0000\", \"expect\": \"\" }; a deck has one empty title placeholder at { \"paraId\": \"slide256/shape2/p0\", \"expect\": \"\" }, and its slide-targeted verbs use { \"kind\": \"slide\", \"path\": \"slide#256\" }. " +
                "Plan-validation errors guarantee no write. Provider and cancellation errors may occur after storage accepted the file, so do not retry the same name; report the possibly unregistered name to the host for recovery. " +
                "Returns {isValid, committed, receipt, sourceDocumentId, outputConnectionId, outputDocumentId, outputVersion, outputName, outputContentType, changes, errors}; non-applicable values are null.")));
            functions.Add(AIFunctionFactory.Create(PopulateTemplateBatch, Opts(
                "populate_template_batch",
                "Populate one Word or PowerPoint template into separate new documents. requestJson contains items with outputName and binding. Scalar values address content-control tags or PowerPoint shape names. Repeating rows are Word-only and replace {{Field}} placeholders in an explicitly selected row. Every output validates and saves atomically and returns its own receipt.")));
            functions.Add(AIFunctionFactory.Create(MergeDocuments, Opts(
                "merge_documents",
                "Commit a Word assembly plan from preview_document_merge. planJson is the complete returned plan, including all source hashes. Revalidates every input and creates one new .docx in destinationConnectionId. Requires read access to every source and create access to the destination. Provider errors can indicate an uncertain write; do not retry blindly.")));
        }
        if (options.AllowInlineContent)
        {
            functions.Add(AIFunctionFactory.Create(CreateDocumentContent, Opts(
                "create_document_content",
                "Create a new document from nothing and get its bytes back, with no storage connection involved. " +
                "name is a bare file name whose extension picks the format: '.docx' makes Word, '.pptx' PowerPoint, and '.xlsx' Excel. " +
                "planJson is optional and authors the document before it is returned; several operations in one call turn nothing into a finished document. " +
                "Pass planJson \"\" for an empty document. The starting anchor differs by format: a Word document has one empty paragraph at { \"paraId\": \"auto-0000\", \"expect\": \"\" }; a deck has one empty title placeholder at { \"paraId\": \"slide256/shape2/p0\", \"expect\": \"\" }. " +
                "Returns {isValid, committed, receipt, name, contentBase64, contentBytes, changes, errors}. contentBase64 is the finished document - hand it to the host to save, or pass it straight back to edit_document_content to keep working. It is null when the plan failed, and then nothing was created.\n\n" +
                PlanOperations)));
            functions.Add(AIFunctionFactory.Create(InspectDocumentContent, Opts(
                "inspect_document_content",
                "Inspect a document supplied inline as base64, with no storage connection involved. " +
                "Returns exactly what inspect_document returns - outline, paragraphs, content controls, nodes, styles, and a snapshot etag - for a document you hold the bytes of rather than one the host registered. " +
                "Use paragraphOffset/paragraphLimit to page; fidelity='outline'|'structure'|'content' to control payload size. " +
                "Prefer fidelity 'outline' or 'structure' on a large document: the bytes already cost you once on the way in.")));
            functions.Add(AIFunctionFactory.Create(EditDocumentContent, Opts(
                "edit_document_content",
                "Edit a document supplied inline as base64 and get the edited document back, with no storage connection involved. " +
                "planJson is an operations array [ … ] or { \"operations\": [ … ] }.\n" +
                "Targets may name text directly instead of a paragraph id, so inspecting first is optional:\n" +
                "{ \"op\": \"changeText\", \"target\": { \"find\": \"Acme Corp\" }, \"with\": \"Globex Inc.\" }\n" +
                "If that text matches more than once the call fails with 'ambiguous-anchor' and lists each candidate; re-issue with { \"find\": \"Acme Corp\", \"match\": 2 } (zero-based) or use more surrounding text. Anchors from inspect_document_content work here too, and can be mixed in the same plan. " +
                "Set preview=true to validate without producing a document - the report comes back with contentBase64 null and nothing is applied. " +
                "Returns {isValid, committed, receipt, name, contentBase64, contentBytes, changes, errors}. contentBase64 carries the edited document and is null on a preview or a failure. " +
                "Nothing is stored anywhere: the returned document is the only copy, so pass it on or hand it to the host before dropping it. " +
                "Pass back the contentBase64 you were given, complete and unchanged - content that arrives altered or truncated is refused as invalid-argument, because it is no longer a readable package.\n\n" +
                PlanOperations)));
        }
        if (options.AllowEphemeralDocuments)
        {
            functions.Add(AIFunctionFactory.Create(ImportDocumentContent, Opts(
                "import_document_content",
                "Put a document you hold the bytes of into a session connection and get back an opaque documentId. " +
                "From then on use that id with the ordinary document tools - inspect_document, find_in_document, preview_plan, apply_plan - and never send the bytes again. " +
                "Prefer this over edit_document_content whenever more than one edit is coming: passing a document back as base64 requires reproducing it exactly, and a long one will not survive that. " +
                "The document is held by the server for this session only and is not written to any storage. " +
                "Returns {connectionId, documentId, name, contentType, version}.")));
            functions.Add(AIFunctionFactory.Create(ExportDocumentContent, Opts(
                "export_document_content",
                "Return the current bytes of a document held in a session connection, as base64, so the host can save or deliver it. " +
                "Do this once, at the end - each call spends the whole document in context. Works only on session connections; a document in real storage is already saved where it belongs. " +
                "Returns {connectionId, documentId, name, contentType, contentBytes, contentBase64}.")));
        }
        return functions.ToArray();
    }

    /// <summary>
    /// The operation vocabulary, shared by every tool that accepts a plan.
    /// </summary>
    /// <remarks>
    /// Stated once because the tool that carries it is not always present. An
    /// inline-content deployment has no preview_plan to refer an agent to, and an agent
    /// that cannot see an operation's shape invents one - which is how a plan ends up
    /// with an insert whose text never arrives.
    /// </remarks>
    public const string PlanOperations =
        "Each operation is one object. Concrete examples:\n\n" +
            "// Replace text:\n" +
            "{ \"op\": \"changeText\", \"target\": { \"paraId\": \"w14:...\", \"expect\": \"Acme Corp\", \"occurrence\": 0 }, \"with\": \"Globex Inc.\", \"mode\": \"Tracked\" }\n\n" +
            "// Unified formatting (paragraph/run/table/row/cell/image):\n" +
            "{ \"op\": \"format\", \"target\": { \"paraId\": \"w14:...\", \"expect\": \"important\", \"occurrence\": 0 }, \"highlight\": \"yellow\", \"bold\": true, \"color\": \"FF0000\" }\n" +
            "{ \"op\": \"format\", \"target\": { \"kind\": \"table\",     \"path\": \"table#0\" }, \"styleId\": \"TableGrid\", \"borderStyle\": \"single\" }\n" +
            "{ \"op\": \"format\", \"target\": { \"kind\": \"image\",     \"path\": \"image#0\" }, \"widthPx\": 320, \"heightPx\": 200 }\n\n" +
            "// Fill / comment / insert paragraph / setProperty:\n" +
            "{ \"op\": \"fill\", \"target\": { \"tag\": \"ClientName\" }, \"value\": \"Globex\" }\n" +
            "{ \"op\": \"comment\", \"target\": { \"paraId\": \"w14:...\", \"expect\": \"...\" }, \"text\": \"Confirm this.\" }\n" +
            "{ \"op\": \"insert\", \"target\": { \"paraId\": \"w14:...\", \"expect\": \"...\" }, \"position\": \"After\", \"text\": \"New paragraph.\" }\n" +
            "{ \"op\": \"insertParagraphs\", \"target\": { \"paraId\": \"w14:...\", \"expect\": \"...\" }, \"position\": \"After\", \"paragraphs\": [{ \"text\": \"First\" }, { \"text\": \"Second\" }] }\n" +
            "{ \"op\": \"removeParagraph\", \"target\": { \"paraId\": \"w14:...\", \"expect\": \"Complete paragraph text\" } }\n" +
            "{ \"op\": \"setProperty\", \"target\": { \"kind\": \"docProperty\", \"path\": \"core/title\" }, \"value\": \"My Title\" }\n\n" +
            "// Every verb above that changes Word content also takes \"mode\": \"Tracked\" (the default - the edit lands as a redline a reviewer accepts or rejects) or \"Direct\". A deck refuses \"Tracked\": PresentationML has no revision markup.\n\n" +
            "// Review an existing redline. Revision paths come from inspect_document.nodes (kind \"revision\"): 'ins#7', 'del#7', 'markIns#7', 'rowIns#7', 'cellDel#7', 'runFormat#7', 'paraFormat#7'. 'all' takes every one; 'author:<name>' takes one person's:\n" +
            "{ \"op\": \"revision\", \"target\": { \"kind\": \"revision\", \"path\": \"all\" }, \"action\": \"Accept\" }\n" +
            "{ \"op\": \"revision\", \"target\": { \"kind\": \"revision\", \"path\": \"author:Jane Doe\" }, \"action\": \"Reject\" }\n\n" +
            "// Reply to, resolve, or delete an existing comment (comment paths from inspect_document.nodes, kind \"comment\"):\n" +
            "{ \"op\": \"comment\", \"target\": { \"kind\": \"comment\", \"path\": \"comment#1\" }, \"action\": \"Reply\", \"text\": \"Forty-five, per the MSA.\" }\n" +
            "{ \"op\": \"comment\", \"target\": { \"kind\": \"comment\", \"path\": \"comment#1\" }, \"action\": \"Resolve\" }\n" +
            "{ \"op\": \"comment\", \"target\": { \"kind\": \"comment\", \"path\": \"comment#1\" }, \"action\": \"Remove\" }\n\n" +
            "// Define a style once instead of repeating direct formatting on every paragraph. Word only.\n" +
            "// Define it first, then apply it with format's styleId - both can sit in the same plan:\n" +
            "{ \"op\": \"defineStyle\", \"styleId\": \"Quote\", \"name\": \"Pull Quote\", \"basedOn\": \"Normal\", \"next\": \"Normal\",\n" +
            "  \"fontFamily\": \"Georgia\", \"sizeHalfPoints\": 24, \"italic\": true, \"color\": \"444444\",\n" +
            "  \"alignment\": \"center\", \"indentLeftTwips\": 720, \"spacingBeforeTwips\": 240 }\n" +
            "{ \"op\": \"format\", \"target\": { \"paraId\": \"w14:...\", \"expect\": \"\" }, \"styleId\": \"Quote\" }\n" +
            "// type is paragraph (default), character or table. outlineLevel 1-9 puts a heading in the outline.\n" +
            "// Defining a style that exists updates it; properties you leave out keep their values. Styles are\n" +
            "// never deleted. A style cannot carry a highlight - w:highlight belongs to a run, so use color here.\n\n" +
            "// Word page geometry, breaks, and notes. All measurements are twips (1440 to the inch). Word only:\n" +
            "{ \"op\": \"pageSetup\", \"paperSize\": \"A4\", \"orientation\": \"Landscape\", \"marginTopTwips\": 720, \"marginLeftTwips\": 1080 }\n" +
            "{ \"op\": \"insertBreak\", \"target\": { \"paraId\": \"w14:...\", \"expect\": \"...\" }, \"kind\": \"Page\", \"position\": \"After\" }\n" +
            "{ \"op\": \"insertBreak\", \"target\": { \"paraId\": \"w14:...\", \"expect\": \"...\" }, \"kind\": \"SectionNextPage\" }   // then pageSetup with a target inside the new section\n" +
            "{ \"op\": \"note\", \"target\": { \"paraId\": \"w14:...\", \"expect\": \"thirty days\" }, \"kind\": \"Footnote\", \"text\": \"Subject to clause 8.2.\" }\n" +
            "{ \"op\": \"note\", \"target\": { \"kind\": \"note\", \"path\": \"footnote#1\" }, \"action\": \"Update\", \"text\": \"Revised wording.\" }\n" +
            "{ \"op\": \"note\", \"target\": { \"kind\": \"note\", \"path\": \"footnote#1\" }, \"action\": \"Remove\" }\n\n" +
            "// Insert a whole new table after a paragraph, or remove an entire table (table path from inspect_document.nodes):\n" +
            "{ \"op\": \"insertTable\", \"target\": { \"paraId\": \"w14:...\", \"expect\": \"...\" }, \"position\": \"After\", \"table\": { \"headers\": [\"Region\", \"Q1\"], \"rows\": [[\"NL\", \"41850\"]] } }\n" +
            "{ \"op\": \"removeTable\",  \"target\": { \"kind\": \"table\", \"path\": \"table#0\" } }\n\n" +
            "// Add or remove table rows / columns; insert or remove image; copy or clear styles. Paths come from inspect_document.nodes:\n" +
            "{ \"op\": \"insertTableRows\", \"target\": { \"kind\": \"table\", \"path\": \"table#0\" }, \"rows\": [[\"NL\",\"17\",\"41850\"]], \"position\": \"End\" }\n" +
            "{ \"op\": \"repeatTableRow\", \"target\": { \"kind\": \"table\", \"path\": \"table#0\" }, \"templateRowIndex\": 1, \"records\": [{ \"Description\": \"Consulting\", \"Amount\": \"1200.00\" }] }\n" +
            "{ \"op\": \"removeTableRows\", \"target\": { \"kind\": \"table\", \"path\": \"table#0\" }, \"onlyIfEmpty\": true }\n" +
            "{ \"op\": \"insertImage\", \"target\": { \"paraId\": \"w14:...\", \"expect\": \"...\" }, \"base64Bytes\": \"iVBORw0KGgo...\", \"imageType\": \"png\", \"widthPx\": 200, \"heightPx\": 80 }\n" +
            "{ \"op\": \"insertImage\", \"target\": { \"paraId\": \"w14:...\", \"expect\": \"...\" }, \"imageConnectionId\": \"images\", \"imageDocumentId\": \"<opaque id from a prior add>\", \"imageType\": \"png\", \"widthPx\": 200, \"heightPx\": 80 }\n" +
            "{ \"op\": \"removeImage\", \"target\": { \"kind\": \"image\", \"path\": \"image#0\" } }\n" +
            "{ \"op\": \"backgroundImage\", \"base64Bytes\": \"iVBORw0KGgo...\", \"imageType\": \"png\", \"opacity\": 0.2 }\n" +
            "{ \"op\": \"backgroundImage\", \"target\": { \"kind\": \"slide\", \"path\": \"slide#256\" }, \"base64Bytes\": \"iVBORw0KGgo...\", \"opacity\": 0.15 }\n" +
            "{ \"op\": \"headerFooter\", \"header\": \"Northwind Traders\", \"footer\": \"Confidential\", \"showPageNumber\": true, \"alignment\": \"edges\", \"differentFirstPage\": true }\n\n" +
            "// Native PowerPoint chart with an editable embedded workbook:\n" +
            "{ \"op\": \"insertChart\", \"target\": { \"kind\": \"slide\", \"path\": \"slide#256\" }, \"kind\": \"ClusteredColumn\", \"categories\": [\"Q1\",\"Q2\"], \"series\": [{ \"name\": \"Revenue\", \"values\": [10,12] }], \"title\": \"Revenue\", \"description\": \"Quarterly revenue\" }\n\n" +
            "// Excel cells and table rows; sheet ids and table paths come from inspection:\n" +
            "{ \"op\": \"setCell\", \"target\": { \"sheetId\": 7, \"address\": \"B2\" }, \"formula\": \"SUM(B3:B8)\" }\n" +
            "{ \"op\": \"appendTableRows\", \"target\": { \"kind\": \"spreadsheetTable\", \"path\": \"table#7/Sales\" }, \"rows\": [[\"APAC\",\"15\"]] }";

    private AIFunction[] CoreFunctions() => new[]
    {
        AIFunctionFactory.Create(InspectDocument, Opts(
            "inspect_document",
            "Inspect a Word, PowerPoint, or Excel document. Excel returns worksheets, tables, and a bounded cell list; use sheetId, range, and maximumCells to narrow it. Other formats return their outline, paragraphs, content controls, nodes, and styles. Copy anchors and node paths from this result.")),
        AIFunctionFactory.Create(FindInDocument, Opts(
            "find_in_document",
            "Find content in Word, PowerPoint, or Excel. Excel can search displayed, raw, or both cell representations and returns sheetId plus A1 address anchors.")),
        AIFunctionFactory.Create(PreviewPlan, Opts(
            "preview_plan",
            "Dry-run a DocumentPlan JSON against (connectionId, documentId). Returns {isValid, committed, receipt, sourceDocumentId, outputConnectionId, outputDocumentId, outputVersion, outputName, outputContentType, changes, errors}; the output fields are null and committed is false. " +
            "Plan shape: { \"snapshot\": { \"eTag\": \"<snapshot string from inspect_document>\" }, \"revision\": { \"author\": \"Review Bot\", \"timestampUtc\": \"2026-09-09T10:00:00Z\" }, \"operations\": [ ... ] }. revision controls Word's displayed revision identity; omit timestampUtc to use one engine timestamp for the whole apply. The snapshot detects drift in Word text-host XML or PowerPoint slide/notes XML; other parts rely on anchors and provider version checks. Omit it only intentionally. Do not set contractVersion. " +
            PlanOperations)),
        AIFunctionFactory.Create(ApplyPlan, Opts(
            "apply_plan",
            "Apply a DocumentPlan JSON to (connectionId, documentId) and save through the provider. Returns {isValid, committed, receipt, sourceDocumentId, outputConnectionId, outputDocumentId, outputVersion, outputName, outputContentType, changes, errors}; non-applicable values are null. The receipt hashes the effective plan and exact input/output bytes and keeps the host-authenticated actor separate from the plan's display revision author. saveMode: 'Replace' (default, overwrites the source after an optimistic version check), 'NewVersion' (keeps the source and mints a new id under the same connection), 'NewDocument' (mints a fresh id with an optional newName for display). On any failure nothing is written.")),
        AIFunctionFactory.Create(CompareDocuments, Opts(
            "compare_documents",
            "Read two Word documents and return paragraph differences, exact input SHA-256 hashes, coverage diagnostics, and a tracked-change plan bound to the original snapshot. The first version covers free body paragraphs. Unsupported changes make isComplete false and plan null. This tool writes nothing; preview and apply the returned plan against the original document.")),
        AIFunctionFactory.Create(PreviewDocumentMerge, Opts(
            "preview_document_merge",
            "Preview ordered whole-document Word assembly without saving. requestJson contains sources [{connectionId, documentId}] in output order and optional options {title, author}. Returns a merge plan bound to exact input hashes, source counts, identifier remapping decisions, and blocking diagnostics. Source formatting is preserved within the documented compatibility scope; each document starts on a new page. This is assembly, not reconciliation of edited versions."))
    };

    /// <summary>Previews assembly after authorizing every source before opening any source.</summary>
    public Task<string> PreviewDocumentMerge(string requestJson, CancellationToken cancellationToken = default) => SafeAsync(async () =>
    {
        var wireRequest = JsonSerializer.Deserialize<MergeToolRequest>(requestJson, PlanJson)
            ?? throw new JsonException("Merge request was null.");
        var request = new DocumentMergeRequest
        {
            Sources = wireRequest.Sources.Select(source =>
                DocumentReference.For(string.Empty, source.ConnectionId, source.DocumentId)).ToArray(),
            Options = wireRequest.Options
        };
        foreach (var source in request.Sources)
            await DemandAccessAsync(source.ConnectionId, ConnectionCapability.Read, cancellationToken).ConfigureAwait(false);
        var result = await _client.PreviewMergeAsync(request, cancellationToken).ConfigureAwait(false);
        return JsonSerializer.Serialize(result, PlanJson);
    });

    private sealed record MergeToolRequest
    {
        public IReadOnlyList<MergeToolSource> Sources { get; init; } = Array.Empty<MergeToolSource>();

        public DocumentMergeOptions Options { get; init; } = new();
    }

    private sealed record MergeToolSource
    {
        public string ConnectionId { get; init; } = string.Empty;

        public string DocumentId { get; init; } = string.Empty;
    }

    /// <summary>Creates an assembled document after authorizing all input and output connections.</summary>
    public Task<string> MergeDocuments(string planJson, string destinationConnectionId, string outputName,
        CancellationToken cancellationToken = default) => SafeAsync(async () =>
    {
        var plan = JsonSerializer.Deserialize<DocumentMergePlan>(planJson, PlanJson)
            ?? throw new JsonException("Merge plan was null.");
        foreach (var input in plan.Inputs)
            await DemandAccessAsync(input.Document?.ConnectionId ?? "", ConnectionCapability.Read, cancellationToken).ConfigureAwait(false);
        await DemandAccessAsync(destinationConnectionId, ConnectionCapability.Create, cancellationToken).ConfigureAwait(false);
        var result = await _client.CommitMergeAsync(plan, destinationConnectionId, outputName, cancellationToken).ConfigureAwait(false);
        return JsonSerializer.Serialize(result, PlanJson);
    });

    /// <summary>Populates one template into a bounded set of independent outputs.</summary>
    public Task<string> PopulateTemplateBatch(
        string connectionId,
        string documentId,
        string requestJson,
        CancellationToken cancellationToken = default) => SafeAsync(async () =>
        {
            await DemandAccessAsync(connectionId, ConnectionCapability.Read, cancellationToken).ConfigureAwait(false);
            await DemandAccessAsync(connectionId, ConnectionCapability.Create, cancellationToken).ConfigureAwait(false);
            var request = JsonSerializer.Deserialize<TemplateBatchRequest>(requestJson, PlanJson)
                ?? throw new JsonException("Template batch JSON was null.");
            var result = await _client.PopulateTemplateBatchAsync(
                connectionId, documentId, request, cancellationToken).ConfigureAwait(false);
            return JsonSerializer.Serialize(result, PlanJson);
        });

    /// <summary>Compares two Word documents without writing either document.</summary>
    public Task<string> CompareDocuments(
        string originalConnectionId,
        string originalDocumentId,
        string revisedConnectionId,
        string revisedDocumentId,
        string revisionAuthor = "OfficeAgent Compare",
        CancellationToken cancellationToken = default) => SafeAsync(async () =>
        {
            await DemandAccessAsync(originalConnectionId, ConnectionCapability.Read, cancellationToken).ConfigureAwait(false);
            await DemandAccessAsync(revisedConnectionId, ConnectionCapability.Read, cancellationToken).ConfigureAwait(false);
            var result = await _client.CompareDocumentsAsync(
                originalConnectionId,
                originalDocumentId,
                revisedConnectionId,
                revisedDocumentId,
                new DocumentComparisonOptions
                {
                    Revision = new RevisionMetadata { Author = revisionAuthor }
                },
                cancellationToken).ConfigureAwait(false);
            return JsonSerializer.Serialize(result, PlanJson);
        });

    /// <summary>Inspects a document and returns paginated JSON.</summary>
    public Task<string> InspectDocument(
        string connectionId,
        string documentId,
        string fidelity = "content",
        int paragraphOffset = 0,
        int paragraphLimit = 200,
        uint sheetId = 0,
        string range = "",
        int maximumCells = 1000,
        CancellationToken cancellationToken = default)
        => SafeAsync(async () =>
        {
            await DemandAccessAsync(connectionId, ConnectionCapability.Read, cancellationToken).ConfigureAwait(false);
            var options = new InspectOptions
            {
                Fidelity = ParseFidelity(fidelity),
                SheetId = sheetId == 0 ? null : sheetId,
                Range = string.IsNullOrWhiteSpace(range) ? null : range,
                MaximumCells = maximumCells
            };
            var result = await _client.InspectAsync(connectionId, documentId, options, cancellationToken).ConfigureAwait(false);
            return JsonSerializer.Serialize(
                InspectPayload(result, paragraphOffset, paragraphLimit), Json);
        });

    /// <summary>
    /// Shapes an inspection for the wire. Ordered so the fields an agent reads first come
    /// first, and shared with <c>open_document</c> so both return the identical structure.
    /// </summary>
    private static Dictionary<string, object?> InspectPayload(
        InspectResult result, int paragraphOffset, int paragraphLimit) => new()
    {
        ["format"] = result.Format.ToString(),
        ["snapshot"] = result.Snapshot.ETag,
        ["outline"] = result.Outline.Select(MapOutline),
        ["paragraphsTotal"] = result.Paragraphs.Count,
        ["paragraphOffset"] = paragraphOffset,
        ["paragraphLimit"] = paragraphLimit,
        // `location` says which part of the document a paragraph lives in - body, header,
        // footer, footnote, endnote, or a deck's speaker notes. Without it an agent
        // editing a Word document cannot tell a footnote from body text, and a caption in
        // a header reads identically to one in the flow.
        ["paragraphs"] = result.Paragraphs
            .Skip(Math.Max(0, paragraphOffset))
            .Take(Math.Max(0, paragraphLimit))
            .Select(p => new { p.ParaId, style = p.StyleId, p.Text, @in = p.In, location = p.Location }),
        ["contentControls"] = result.StructuralAnchors.Select(s => new { s.Tag, s.Kind }),
        ["nodes"] = result.Nodes.Select(n => new { n.Kind, n.Path, n.Summary }),
        ["styles"] = result.Styles.Styles.Select(s => new { s.Id, s.Name, s.InUseCount }),
        ["worksheets"] = result.Worksheets,
        ["cells"] = result.Cells
    };

    /// <summary>Finds text in a document and returns content-verified anchors.</summary>
    public Task<string> FindInDocument(
        string connectionId,
        string documentId,
        string pattern,
        bool regex = false,
        bool wholeWord = false,
        bool caseSensitive = false,
        string spreadsheetValueView = "both",
        CancellationToken cancellationToken = default)
        => SafeAsync(async () =>
        {
            await DemandAccessAsync(connectionId, ConnectionCapability.Read, cancellationToken).ConfigureAwait(false);
            var query = new FindQuery
            {
                Pattern = pattern,
                Options = new MatchOptions
                {
                    Regex = regex,
                    WholeWord = wholeWord,
                    CaseSensitive = caseSensitive,
                    SpreadsheetValueView = ParseSpreadsheetValueView(spreadsheetValueView)
                }
            };
            var hits = await _client.FindAsync(connectionId, documentId, query, cancellationToken).ConfigureAwait(false);
            return JsonSerializer.Serialize(hits.Select(h => new
            {
                paraId = (h.Anchor as TextSpanAnchor)?.ParaId,
                expect = h.Text,
                occurrence = (h.Anchor as TextSpanAnchor)?.Occurrence ?? 0,
                context = h.Context,
                location = h.Location,
                sheetId = (h.Anchor as CellAnchor)?.SheetId,
                address = (h.Anchor as CellAnchor)?.Address
            }), Json);
        });

    /// <summary>Dry-runs a plan against the document.</summary>
    public Task<string> PreviewPlan(
        string connectionId,
        string documentId,
        string planJson,
        CancellationToken cancellationToken = default)
        => SafeAsync(async () =>
        {
            await DemandAccessAsync(connectionId, ConnectionCapability.Read, cancellationToken).ConfigureAwait(false);
            var plan = DeserializePlan(planJson, connectionId);
            await DemandReferencedContentAccessAsync(plan, cancellationToken).ConfigureAwait(false);
            using var result = await _client.PreviewWithReceiptAsync(
                connectionId, documentId, plan, cancellationToken).ConfigureAwait(false);
            return SerializeReport(
                result.Report, committed: false, savedReference: null, receipt: result.Receipt);
        });

    /// <summary>Applies a plan and saves through the provider.</summary>
    public Task<string> ApplyPlan(
        string connectionId,
        string documentId,
        string planJson,
        string saveMode = "Replace",
        string newName = "",
        CancellationToken cancellationToken = default)
        => SafeAsync(async () =>
        {
            var parsedSaveMode = ParseSaveMode(saveMode);
            await DemandAccessAsync(connectionId, ConnectionCapability.Read, cancellationToken).ConfigureAwait(false);
            await DemandAccessAsync(
                connectionId,
                parsedSaveMode == SaveMode.Replace ? ConnectionCapability.Edit : ConnectionCapability.Create,
                cancellationToken).ConfigureAwait(false);
            var plan = DeserializePlan(planJson, connectionId);
            await DemandReferencedContentAccessAsync(plan, cancellationToken).ConfigureAwait(false);
            var options = new SaveDocumentOptions
            {
                Mode = parsedSaveMode,
                NewName = string.IsNullOrEmpty(newName) ? null : newName
            };
            var result = await _client.CommitAsync(connectionId, documentId, plan, options, cancellationToken).ConfigureAwait(false);
            return SerializeReport(
                result.Report, result.Committed, result.Committed ? result.Document : null,
                receipt: result.Receipt);
        });

    /// <summary>Registers a document with a provider connection and returns its opaque id.</summary>
    public Task<string> RegisterDocument(
        string connectionId,
        string source,
        CancellationToken cancellationToken = default)
        => SafeAsync(async () =>
        {
            await DemandAccessAsync(connectionId, ConnectionCapability.Register, cancellationToken).ConfigureAwait(false);
            var reference = await _client.RegisterAsync(connectionId, source, cancellationToken).ConfigureAwait(false);
            return JsonSerializer.Serialize(new
            {
                connectionId = reference.ConnectionId,
                documentId = reference.ItemId,
                name = reference.Name,
                contentType = reference.ContentType,
                version = reference.Version
            }, Json);
        });

    /// <summary>
    /// Creates a new document in a provider connection, optionally applying an initial
    /// plan to it before anything is written.
    /// </summary>
    public Task<string> CreateDocument(
        string connectionId,
        string name,
        string planJson = "",
        CancellationToken cancellationToken = default)
        => SafeAsync(async () =>
        {
            await DemandAccessAsync(connectionId, ConnectionCapability.Create, cancellationToken).ConfigureAwait(false);
            var plan = string.IsNullOrWhiteSpace(planJson) ? null : DeserializePlan(planJson, connectionId);
            if (plan is not null)
                await DemandReferencedContentAccessAsync(plan, cancellationToken).ConfigureAwait(false);
            var result = await _client.CreateAsync(connectionId, name, plan, cancellationToken).ConfigureAwait(false);
            return SerializeReport(
                result.Report, result.Committed, result.Committed ? result.Document : null,
                receipt: result.Receipt);
        });

    /// <summary>
    /// Registers a document by its connection-relative source and inspects it in one call,
    /// so the agent reaches a working documentId and the document's structure together.
    /// </summary>
    public Task<string> OpenDocument(
        string connectionId,
        string source,
        string fidelity = "content",
        int paragraphOffset = 0,
        int paragraphLimit = 200,
        uint sheetId = 0,
        string range = "",
        int maximumCells = 1000,
        CancellationToken cancellationToken = default)
        => SafeAsync(async () =>
        {
            await DemandAccessAsync(connectionId, ConnectionCapability.Register, cancellationToken).ConfigureAwait(false);
            await DemandAccessAsync(connectionId, ConnectionCapability.Read, cancellationToken).ConfigureAwait(false);
            var reference = await _client.RegisterAsync(connectionId, source, cancellationToken).ConfigureAwait(false);
            var options = new InspectOptions
            {
                Fidelity = ParseFidelity(fidelity), SheetId = sheetId == 0 ? null : sheetId,
                Range = string.IsNullOrWhiteSpace(range) ? null : range, MaximumCells = maximumCells
            };
            var result = await _client.InspectAsync(
                connectionId, reference.ItemId, options, cancellationToken).ConfigureAwait(false);

            // The registration fields lead, because the documentId is what every follow-up
            // call needs; the inspection then follows in its usual shape.
            var payload = new Dictionary<string, object?>
            {
                ["connectionId"] = reference.ConnectionId,
                ["documentId"] = reference.ItemId,
                ["name"] = reference.Name,
                ["contentType"] = reference.ContentType,
                ["version"] = reference.Version
            };
            foreach (var field in InspectPayload(result, paragraphOffset, paragraphLimit))
                payload[field.Key] = field.Value;

            return JsonSerializer.Serialize(payload, Json);
        });

    /// <summary>
    /// Registers a document by source, binds any <c>find</c> targets to live anchors, and
    /// applies the operations - the whole inspect/find/preview/apply loop in one call.
    /// </summary>
    public Task<string> EditDocument(
        string connectionId,
        string source,
        string planJson,
        string saveMode = "Replace",
        string newName = "",
        CancellationToken cancellationToken = default)
        => SafeAsync(async () =>
        {
            var parsedSaveMode = ParseSaveMode(saveMode);
            await DemandAccessAsync(connectionId, ConnectionCapability.Register, cancellationToken).ConfigureAwait(false);
            await DemandAccessAsync(connectionId, ConnectionCapability.Read, cancellationToken).ConfigureAwait(false);
            await DemandAccessAsync(
                connectionId,
                parsedSaveMode == SaveMode.Replace ? ConnectionCapability.Edit : ConnectionCapability.Create,
                cancellationToken).ConfigureAwait(false);
            var reference = await _client.RegisterAsync(connectionId, source, cancellationToken).ConfigureAwait(false);
            var planObject = ParsePlanObject(planJson);

            // Anchor binding happens against the freshly registered document, so the text
            // an operation names is verified against live content before anything is applied.
            var failures = await new FindTargetResolver(_client)
                .ResolveAsync(reference, planObject, cancellationToken).ConfigureAwait(false);
            if (failures.Count > 0)
                return SerializeErrors(
                    failures.Select(f => (f.Code, f.Message)).ToArray(),
                    connectionId, reference.ItemId);

            var options = new SaveDocumentOptions
            {
                Mode = parsedSaveMode,
                NewName = string.IsNullOrEmpty(newName) ? null : newName
            };
            var plan = DeserializePlan(planObject, connectionId);
            await DemandReferencedContentAccessAsync(plan, cancellationToken).ConfigureAwait(false);
            var result = await _client.CommitAsync(
                connectionId, reference.ItemId, plan, options, cancellationToken).ConfigureAwait(false);

            return SerializeReport(
                result.Report, result.Committed, result.Committed ? result.Document : null,
                sourceDocumentId: reference.ItemId, receipt: result.Receipt);
        });

    // ── Inline content: no connection, no stored document ────────────────

    /// <summary>
    /// Creates a document from nothing and returns its bytes, optionally applying an
    /// initial plan first. No provider is involved and nothing is stored.
    /// </summary>
    public Task<string> CreateDocumentContent(
        string name,
        string planJson = "",
        CancellationToken cancellationToken = default)
        => SafeContentAsync(name, async () =>
        {
            var bytes = _client.CreateBlank(name);
            if (string.IsNullOrWhiteSpace(planJson))
                return SerializeContent(new ChangeReport { IsValid = true }, committed: true, bytes, name);

            return await EditBytesAsync(bytes, name, planJson, preview: false, cancellationToken).ConfigureAwait(false);
        });

    /// <summary>Inspects a document supplied as base64, returning the usual inspection payload.</summary>
    public Task<string> InspectDocumentContent(
        string contentBase64,
        string fidelity = "content",
        int paragraphOffset = 0,
        int paragraphLimit = 200,
        uint sheetId = 0,
        string range = "",
        int maximumCells = 1000,
        CancellationToken cancellationToken = default)
        => SafeContentAsync(name: null, () =>
        {
            var bytes = DecodeContent(contentBase64);
            var handle = new StreamHandle(new MemoryStream(bytes, writable: false));

            var result = ReadingDocument(() =>
                _client.Inspect(handle, new InspectOptions
                {
                    Fidelity = ParseFidelity(fidelity), SheetId = sheetId == 0 ? null : sheetId,
                    Range = string.IsNullOrWhiteSpace(range) ? null : range, MaximumCells = maximumCells
                }));
            return Task.FromResult(JsonSerializer.Serialize(
                InspectPayload(result, paragraphOffset, paragraphLimit), Json));
        });

    /// <summary>
    /// Applies a plan to a document supplied as base64 and returns the edited document the
    /// same way. Nothing is stored; the caller owns the result.
    /// </summary>
    public Task<string> EditDocumentContent(
        string contentBase64,
        string planJson,
        bool preview = false,
        CancellationToken cancellationToken = default)
        => SafeContentAsync(name: null, () =>
            EditBytesAsync(DecodeContent(contentBase64), name: null, planJson, preview, cancellationToken));

    /// <summary>
    /// The shared body of the two inline verbs: bind any <c>find</c> targets against the
    /// content in hand, apply, and hand the bytes back.
    /// </summary>
    private async Task<string> EditBytesAsync(
        byte[] bytes, string? name, string planJson, bool preview, CancellationToken cancellationToken)
    {
        // Read the document once, before anything else touches it. It settles two things
        // at the same time: that the content arrived intact - which is what every later
        // step would otherwise discover as an opaque failure - and which format it is, so
        // the change mode can be resolved without opening the package a second time.
        var format = ReadingDocument(() => _client.Inspect(
            new StreamHandle(new MemoryStream(bytes, writable: false), name),
            new InspectOptions { Fidelity = Fidelity.Outline }).Format);

        var planObject = ParsePlanObject(planJson);

        var failures = new FindTargetResolver(_client).Resolve(bytes, name, planObject);
        if (failures.Count > 0)
            return SerializeContent(ReportOf(failures), committed: false, content: null, name);

        var plan = DeserializePlan(ApplyInlineChangeMode(planObject, format));
        await DemandReferencedContentAccessAsync(plan, cancellationToken).ConfigureAwait(false);
        var handle = new StreamHandle(new MemoryStream(bytes, writable: false), name);

        if (preview)
        {
            using var previewResult = await _client.ApplyAsync(
                handle, plan, ApplyOptions.Preview, cancellationToken).ConfigureAwait(false);
            return SerializeContent(
                previewResult.Report, committed: false, content: null, name, previewResult.Receipt);
        }

        using var applied = await _client.CommitAsync(handle, plan, cancellationToken).ConfigureAwait(false);
        return SerializeContent(
            applied.Report, applied.Committed, applied.Committed ? applied.ToBytes() : null,
            name, applied.Receipt);
    }

    /// <summary>
    /// Fills in the change mode for a document that has no connection to inherit one from.
    /// </summary>
    /// <remarks>
    /// The provider path takes this from the connection, because the host knows what kind
    /// of documents a connection serves. Inline content has no host to ask - but the bytes
    /// say what the document is, and PresentationML has no revision markup at all, so a
    /// deck's only workable default is <see cref="ChangeMode.Direct"/>. Word keeps the
    /// tracked default. A caller that explicitly asks a deck for <c>Tracked</c> is still
    /// refused, which is the answer they should get.
    /// </remarks>
    private static JsonObject ApplyInlineChangeMode(
        JsonObject planObject, OfficeAgent.Abstractions.DocumentFormat format) =>
        format == OfficeAgent.Abstractions.DocumentFormat.PowerPoint
            ? ApplyDefaultChangeMode(planObject, ChangeMode.Direct)
            : planObject;

    /// <summary>
    /// Reads inline content, turning "these bytes are not a document" into something the
    /// caller can act on.
    /// </summary>
    /// <remarks>
    /// This is the likeliest way an inline call fails, and it used to surface as
    /// <c>internal-error</c>, which tells an agent nothing. The content has to survive
    /// being reproduced in full on the way back in, and a single altered character, or a
    /// truncated tail, leaves valid base64 that is no longer a package. The underlying
    /// message is kept on the end so a genuine engine fault is still diagnosable rather
    /// than being reported as bad input.
    /// </remarks>
    private static T ReadingDocument<T>(Func<T> read)
    {
        try
        {
            return read();
        }
        catch (Exception ex) when (ex is not OperationCanceledException and not OutOfMemoryException)
        {
            throw new ArgumentException(
                "contentBase64 decoded, but the bytes are not a readable .docx, .pptx, or .xlsx package. " +
                "The usual cause is a copy of the content that differs from the original by a " +
                "character or two - it does not survive being reproduced by hand. Do NOT re-send " +
                "the same string: a second attempt reproduces the same copy and fails identically. " +
                "Take the contentBase64 from the tool result verbatim, or ask the host for the " +
                "document again. " +
                $"Underlying error: {ex.Message}");
        }
    }

    /// <summary>
    /// Runs an inline tool, reporting a failure in the same shape its success uses.
    /// </summary>
    /// <remarks>
    /// The connection-addressed wrapper answers with <c>outputDocumentId</c> and its
    /// neighbours, which name nothing here and, worse, leave out the one field an agent was
    /// told to read. An inline failure answers with <c>contentBase64: null</c> instead.
    /// </remarks>
    private static async Task<string> SafeContentAsync(string? name, Func<Task<string>> work)
    {
        try { return await work().ConfigureAwait(false); }
        catch (OperationCanceledException) { return ContentError(name, "cancelled", "Operation was cancelled."); }
        catch (ConnectionForbiddenException ex) { return ContentError(name, "connection-forbidden", ex.Message); }
        catch (RegexMatchTimeoutException) { return ContentError(name, "regex-timeout", RegexTimeoutMessage); }
        catch (JsonException ex) { return ContentError(name, "invalid-json", ex.Message); }
        catch (ArgumentException ex) { return ContentError(name, "invalid-argument", ex.Message); }
        catch (Exception) { return ContentError(name, "internal-error", "An unexpected internal error occurred."); }
    }

    private static string ContentError(string? name, string code, string message) =>
        SerializeContent(
            new ChangeReport
            {
                IsValid = false,
                Errors = new[] { new ValidationError(code, message, target: null) }
            },
            committed: false, content: null, name);

    // ── Ephemeral documents: a handle instead of the bytes ───────────────

    /// <summary>
    /// Puts a document the caller holds the bytes of into an ephemeral connection and
    /// returns its opaque id, so every later call names the id rather than the content.
    /// </summary>
    public Task<string> ImportDocumentContent(
        string connectionId,
        string name,
        string contentBase64,
        CancellationToken cancellationToken = default)
        => SafeAsync(async () =>
        {
            await DemandAccessAsync(connectionId, ConnectionCapability.Create, cancellationToken).ConfigureAwait(false);
            var store = Ephemeral(connectionId);
            var reference = store.Add(name, DecodeContent(contentBase64));

            return JsonSerializer.Serialize(new
            {
                connectionId = reference.ConnectionId,
                documentId = reference.ItemId,
                name = reference.Name,
                contentType = reference.ContentType,
                version = reference.Version
            }, Json);
        });

    /// <summary>Returns the bytes of a document held in an ephemeral connection.</summary>
    public Task<string> ExportDocumentContent(
        string connectionId,
        string documentId,
        CancellationToken cancellationToken = default)
        => SafeAsync(async () =>
        {
            await DemandAccessAsync(connectionId, ConnectionCapability.Read, cancellationToken).ConfigureAwait(false);
            // Restricted to ephemeral connections on purpose. Exporting from a filesystem
            // or SharePoint connection would turn every readable document into base64 an
            // agent can quote, which is a different capability than editing one in place.
            var store = Ephemeral(connectionId);
            var reference = store.Describe(documentId);
            var bytes = store.Read(documentId);

            return JsonSerializer.Serialize(new
            {
                connectionId,
                documentId,
                name = reference.Name,
                contentType = reference.ContentType,
                contentBytes = bytes.Length,
                contentBase64 = Convert.ToBase64String(bytes)
            }, Json);
        });

    /// <summary>
    /// The ephemeral store behind a connection id, or an error naming the ones that exist.
    /// </summary>
    private MemoryDocumentProvider Ephemeral(string connectionId) =>
        _client.EphemeralConnection(connectionId)
        ?? throw new ArgumentException(
            $"'{connectionId}' is not an in-memory connection. These tools work only on connections whose " +
            "documents this server holds for the session; a document in storage is read and written in place " +
            "by the connection-addressed tools instead.");

    /// <summary>Target-binding failures as a report, so they reach the caller in the shape every other result uses.</summary>
    private static ChangeReport ReportOf(IReadOnlyList<FindTargetResolver.Failure> failures) => new()
    {
        IsValid = false,
        Errors = failures.Select(f => new ValidationError(f.Code, f.Message, target: null)).ToArray()
    };

    /// <summary>Decodes the document a caller supplied inline, naming the fix when it is not base64.</summary>
    private static byte[] DecodeContent(string contentBase64)
    {
        if (string.IsNullOrWhiteSpace(contentBase64))
            throw new ArgumentException(
                "contentBase64 is required: pass the document's bytes, base64-encoded.", nameof(contentBase64));

        try
        {
            return Convert.FromBase64String(contentBase64);
        }
        catch (FormatException)
        {
            throw new ArgumentException(
                "contentBase64 is not valid base64. Pass the raw base64 of the .docx, .pptx, or .xlsx package, " +
                "with no data: prefix, quotes, or line wrapping.", nameof(contentBase64));
        }
    }

    /// <summary>Removes a document registration; the underlying content is left untouched.</summary>
    public Task<string> RemoveDocument(
        string connectionId,
        string documentId,
        CancellationToken cancellationToken = default)
        => SafeAsync(async () =>
        {
            await DemandAccessAsync(connectionId, ConnectionCapability.Delete, cancellationToken).ConfigureAwait(false);
            await _client.RemoveAsync(connectionId, documentId, cancellationToken).ConfigureAwait(false);
            return JsonSerializer.Serialize(new { removed = true, connectionId, documentId }, Json);
        });

    private static async Task<string> SafeAsync(Func<Task<string>> work)
    {
        try { return await work().ConfigureAwait(false); }
        catch (OperationCanceledException) { return SerializeError("cancelled", "Operation was cancelled."); }
        catch (ConnectionForbiddenException ex) { return SerializeError("connection-forbidden", ex.Message); }
        catch (RegexMatchTimeoutException) { return SerializeError("regex-timeout", RegexTimeoutMessage); }
        catch (JsonException ex) { return SerializeError("invalid-json", ex.Message); }
        catch (ArgumentException ex) { return SerializeError("invalid-argument", ex.Message); }
        catch (DocumentProviderException ex)
        {
            return SerializeError(
                ProviderCodeToWire(ex.Code),
                ProviderMessage(ex.Code),
                ex.Provider, ex.ConnectionId, ex.ItemId);
        }
        catch (Exception) { return SerializeError("internal-error", "An unexpected internal error occurred."); }
    }

    private async ValueTask DemandAccessAsync(
        string connectionId,
        ConnectionCapability capability,
        CancellationToken cancellationToken)
    {
        if (string.IsNullOrWhiteSpace(connectionId) ||
            !await CanAccessConnectionAsync(connectionId, capability, cancellationToken).ConfigureAwait(false))
            throw new ConnectionForbiddenException();
    }

    private async ValueTask DemandReferencedContentAccessAsync(
        DocumentPlan plan,
        CancellationToken cancellationToken)
    {
        var connections = plan.Operations.Select(operation => operation switch
            {
                InsertImageOp image when !string.IsNullOrWhiteSpace(image.ImageDocumentId) => image.ImageConnectionId,
                BackgroundImageOp background when !string.IsNullOrWhiteSpace(background.ImageDocumentId) => background.ImageConnectionId,
                InsertMediaOp media when !string.IsNullOrWhiteSpace(media.MediaDocumentId) => media.MediaConnectionId,
                _ => null
            })
            .Where(connectionId => !string.IsNullOrWhiteSpace(connectionId))
            .Distinct(StringComparer.Ordinal);

        foreach (var connectionId in connections)
            await DemandAccessAsync(connectionId!, ConnectionCapability.Read, cancellationToken).ConfigureAwait(false);
    }

    private static AIFunctionFactoryOptions Opts(string name, string description) => new()
    {
        Name = name,
        Description = description,
        JsonSchemaCreateOptions = StrictSchemaOptions
    };

    private static Fidelity ParseFidelity(string fidelity) => fidelity?.ToLowerInvariant() switch
    {
        "outline" => Fidelity.Outline,
        "structure" => Fidelity.Structure,
        _ => Fidelity.Content
    };

    private static SpreadsheetValueView ParseSpreadsheetValueView(string value) =>
        value?.Trim().ToLowerInvariant() switch
        {
            "displayed" => SpreadsheetValueView.Displayed,
            "raw" => SpreadsheetValueView.Raw,
            "both" or "" or null => SpreadsheetValueView.Both,
            _ => throw new ArgumentException(
                $"Unknown spreadsheetValueView '{value}'. Expected both, displayed, or raw.", nameof(value))
        };

    /// <summary>
    /// Omitted means the default, <see cref="SaveMode.Replace"/>. A value that is present
    /// but unrecognised is refused rather than defaulted: now that the default writes over
    /// the source, silently treating "NewVerison" as the default would destroy the very
    /// document the caller was trying to preserve.
    /// </summary>
    private static SaveMode ParseSaveMode(string mode)
    {
        var value = mode?.Trim();
        if (string.IsNullOrEmpty(value)) return SaveMode.Replace;

        return value switch
        {
            "NewVersion" => SaveMode.NewVersion,
            "NewDocument" => SaveMode.NewDocument,
            "Replace" => SaveMode.Replace,
            _ => throw new ArgumentException(
                $"Unknown saveMode '{mode}'. Expected NewVersion, NewDocument, or Replace.", nameof(mode))
        };
    }

    private static string ProviderCodeToWire(ProviderErrorCode code) => code switch
    {
        ProviderErrorCode.NotFound => "not-found",
        ProviderErrorCode.AccessDenied => "access-denied",
        ProviderErrorCode.ContentTooLarge => "content-too-large",
        ProviderErrorCode.ExtensionNotAllowed => "extension-not-allowed",
        ProviderErrorCode.VersionConflict => "version-conflict",
        ProviderErrorCode.InvalidArgument => "invalid-argument",
        ProviderErrorCode.ConfigurationError => "configuration-error",
        ProviderErrorCode.IO => "io-error",
        ProviderErrorCode.AlreadyExists => "already-exists",
        _ => "provider-error"
    };

    /// <summary>
    /// Provider exceptions are host diagnostics and may contain absolute paths,
    /// tenant details, or upstream response text. Tool callers receive a stable,
    /// actionable message keyed only by the public provider error code.
    /// </summary>
    private static string ProviderMessage(ProviderErrorCode code) => code switch
    {
        ProviderErrorCode.NotFound => "The document or registration was not found.",
        ProviderErrorCode.AccessDenied => "The source is outside the connection boundary or access was denied.",
        ProviderErrorCode.ContentTooLarge => "The document exceeds the connection's size limit.",
        ProviderErrorCode.ExtensionNotAllowed => "The document extension is not allowed by this connection.",
        ProviderErrorCode.VersionConflict => "The document changed after it was opened. Inspect it again before retrying.",
        ProviderErrorCode.InvalidArgument => "The document provider rejected an argument.",
        ProviderErrorCode.ConfigurationError => "The document connection is not configured for this operation.",
        ProviderErrorCode.IO => "The document provider operation failed.",
        ProviderErrorCode.AlreadyExists => "A document with that name already exists.",
        _ => "The document provider operation failed."
    };

    /// <summary>
    /// Parses the plan JSON a tool was given. Accepts either the full plan object,
    /// <c>{ "operations": [ … ] }</c>, or a bare operations array, <c>[ … ]</c> - models
    /// reach for the array form constantly, and rejecting it buys nothing.
    /// </summary>
    private static JsonObject ParsePlanObject(string planJson)
    {
        var node = JsonNode.Parse(planJson)
            ?? throw new JsonException("Plan JSON was null.");

        return node switch
        {
            JsonObject plan => plan,
            JsonArray operations => new JsonObject { ["operations"] = operations.DeepClone() },
            _ => throw new JsonException(
                "Plan JSON must be an operations array [ … ] or an object { \"operations\": [ … ] }.")
        };
    }

    /// <summary>
    /// The verbs that carry a <c>mode</c>, read off <see cref="ITrackedOperation"/> rather
    /// than listed here. A verb that grows a mode is covered the moment it implements the
    /// interface, and a verb that never had one cannot inherit this policy by accident.
    /// </summary>
    private static readonly HashSet<string> TrackedVerbs = new(
        PlanOperationJsonConverter.ByVerb
            .Where(pair => typeof(ITrackedOperation).IsAssignableFrom(pair.Value))
            .Select(pair => pair.Key),
        StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Writes the connection's default change mode into every operation that did not name
    /// one. This has to happen on the JSON, before deserializing: once a
    /// <see cref="ChangeTextOp"/> exists, its property initializer has already turned an
    /// absent <c>mode</c> into <see cref="ChangeMode.Tracked"/> and the distinction between
    /// "the agent omitted it" and "the agent asked for Tracked" is gone.
    /// </summary>
    private static JsonObject ApplyDefaultChangeMode(JsonObject planObject, ChangeMode connectionDefault)
    {
        if (connectionDefault == ChangeMode.Tracked) return planObject;
        if (planObject["operations"] is not JsonArray operations) return planObject;

        foreach (var operation in operations)
        {
            if (operation is not JsonObject op) continue;
            if (op["op"]?.GetValue<string>() is not { } verb || !TrackedVerbs.Contains(verb)) continue;
            if (op.ContainsKey("mode")) continue;

            op["mode"] = connectionDefault.ToString();
        }

        return planObject;
    }

    private DocumentPlan DeserializePlan(string planJson, string connectionId) =>
        DeserializePlan(ParsePlanObject(planJson), connectionId);

    private DocumentPlan DeserializePlan(JsonObject planObject, string connectionId) =>
        DeserializePlan(ApplyDefaultChangeMode(planObject, _client.DefaultChangeModeFor(connectionId)));

    private static DocumentPlan DeserializePlan(JsonObject planObject)
    {
        var plan = planObject.Deserialize<DocumentPlan>(PlanJson)
            ?? throw new JsonException("Plan JSON did not deserialize to a DocumentPlan.");

        // An explicit "operations": null overwrites the empty default, and every consumer
        // downstream enumerates the list. Name the fix rather than let it surface as an
        // internal error the model cannot act on.
        if (plan.Operations is null)
            throw new JsonException(
                "Plan \"operations\" was null; supply an array of operation objects, " +
                "for example { \"operations\": [ { \"op\": \"insert\", ... } ] }.");

        return plan;
    }

    /// <summary>
    /// The report for an inline edit, with the edited document in it. <c>contentBase64</c>
    /// is null whenever there is nothing to hand back - a preview, or a plan that failed -
    /// so an agent that finds it null knows not to look for a document.
    /// </summary>
    private static string SerializeContent(
        ChangeReport report,
        bool committed,
        byte[]? content,
        string? name,
        ApplyReceipt? receipt = null) =>
        JsonSerializer.Serialize(new
        {
            isValid = report.IsValid,
            committed,
            receipt = ReceiptPayload(receipt),
            name,
            contentBase64 = content is null ? null : Convert.ToBase64String(content),
            contentBytes = content?.Length,
            changes = report.Changes.Select(c => new
            {
                c.Verb,
                target = SummariseAnchor(c.Target),
                c.Before, c.After, c.Context, c.BlastRadius,
                capability = c.Capability.ToString()
            }),
            errors = report.Errors.Select(e => new
            {
                e.Code, e.Message,
                target = SummariseAnchor(e.Target)
            })
        }, Json);

    private static string SerializeReport(
        ChangeReport report,
        bool committed,
        DocumentReference? savedReference,
        string? sourceDocumentId = null,
        ApplyReceipt? receipt = null) =>
        JsonSerializer.Serialize(new
        {
            isValid = report.IsValid,
            committed,
            receipt = ReceiptPayload(receipt),
            // Present only for the composite tools, which mint the source id themselves:
            // it lets the agent keep working with the document it just named by path.
            sourceDocumentId,
            outputConnectionId = savedReference?.ConnectionId,
            outputDocumentId = savedReference?.ItemId,
            outputVersion = savedReference?.Version,
            outputName = savedReference?.Name,
            outputContentType = savedReference?.ContentType,
            changes = report.Changes.Select(c => new
            {
                c.Verb,
                target = SummariseAnchor(c.Target),
                c.Before, c.After, c.Context, c.BlastRadius,
                capability = c.Capability.ToString()
            }),
            errors = report.Errors.Select(e => new
            {
                e.Code, e.Message,
                target = SummariseAnchor(e.Target)
            })
        }, Json);

    private static string SerializeError(string code, string message, string? provider = null, string? connectionId = null, string? itemId = null) =>
        JsonSerializer.Serialize(new
        {
            isValid = false,
            committed = false,
            receipt = (ApplyReceipt?)null,
            sourceDocumentId = (string?)null,
            outputConnectionId = (string?)null,
            outputDocumentId = (string?)null,
            outputVersion = (string?)null,
            outputName = (string?)null,
            outputContentType = (string?)null,
            changes = Array.Empty<object>(),
            errors = new[] { new { Code = code, Message = message, target = (object?)null, provider, connectionId, itemId } }
        }, Json);

    /// <summary>
    /// Reports every anchor that could not be bound in one result. Carries the source
    /// documentId so a failed <c>edit_document</c> still leaves the agent with a usable
    /// handle - it can inspect, disambiguate, and retry without registering again.
    /// </summary>
    private static string SerializeErrors(
        IReadOnlyList<(string Code, string Message)> failures,
        string connectionId,
        string sourceDocumentId) =>
        JsonSerializer.Serialize(new
        {
            isValid = false,
            committed = false,
            receipt = (ApplyReceipt?)null,
            sourceDocumentId,
            outputConnectionId = (string?)null,
            outputDocumentId = (string?)null,
            outputVersion = (string?)null,
            outputName = (string?)null,
            outputContentType = (string?)null,
            changes = Array.Empty<object>(),
            errors = failures.Select(f => new
            {
                f.Code,
                f.Message,
                target = (object?)null,
                connectionId,
                itemId = sourceDocumentId
            })
        }, Json);

    private static object? SummariseAnchor(Anchor? anchor) => anchor switch
    {
        null => null,
        TextSpanAnchor t => new { kind = "textSpan", paraId = t.ParaId, expect = t.Expect, occurrence = t.Occurrence },
        StructuralAnchor s => new { kind = "structural", tag = s.Tag, structuralKind = s.Kind },
        NodeAnchor n => new { kind = "node", nodeKind = n.Kind, path = n.Path },
        StyleAnchor s => new { kind = "style", styleId = s.StyleId },
        _ => new { kind = anchor.GetType().Name, anchor.Id }
    };

    private static object? ReceiptPayload(ApplyReceipt? receipt) => receipt is null
        ? null
        : new
        {
            receiptVersion = receipt.ReceiptVersion,
            planSha256 = receipt.PlanSha256,
            inputSha256 = receipt.InputSha256,
            outputSha256 = receipt.OutputSha256,
            timestampUtc = receipt.TimestampUtc,
            outcome = receipt.Outcome.ToString(),
            revision = new
            {
                author = receipt.Revision.Author,
                timestampUtc = receipt.Revision.TimestampUtc
            },
            actor = receipt.Actor is null ? null : new
            {
                subject = receipt.Actor.Subject,
                issuer = receipt.Actor.Issuer,
                displayName = receipt.Actor.DisplayName
            },
            outputDocument = receipt.OutputDocument is null ? null : new
            {
                provider = receipt.OutputDocument.Provider,
                connectionId = receipt.OutputDocument.ConnectionId,
                itemId = receipt.OutputDocument.ItemId,
                version = receipt.OutputDocument.Version,
                name = receipt.OutputDocument.Name,
                contentType = receipt.OutputDocument.ContentType
            }
        };

    private static object MapOutline(OutlineNode node) => new
    {
        node.Level,
        node.Text,
        paraId = (node.Anchor as TextSpanAnchor)?.ParaId,
        children = node.Children.Select(MapOutline)
    };
}
