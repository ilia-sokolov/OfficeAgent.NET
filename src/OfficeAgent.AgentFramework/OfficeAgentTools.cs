using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
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
    /// remove registrations (<c>register_document</c> / <c>remove_document</c>). The
    /// default is <see langword="false"/>: the host pre-registers documents and the
    /// agent only ever sees opaque ids. Enabling this lets the agent hand
    /// provider-relative sources (a path under a filesystem root, a drive-relative
    /// SharePoint path) to the configured connections; the connection boundary,
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
/// via <see cref="OfficeAgentToolsOptions.AllowRegistration"/>.
/// </summary>
public sealed class OfficeAgentTools
{
    private static readonly JsonSerializerOptions Json = new() { WriteIndented = false };

    private static readonly JsonSerializerOptions PlanJson = new()
    {
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

    /// <summary>Initializes the tool projection over an <see cref="OfficeAgentClient"/>.</summary>
    public OfficeAgentTools(OfficeAgentClient client) => _client = client;

    /// <summary>
    /// System-prompt guidance to concatenate into the agent's instructions. Teaches
    /// the model the <c>(connectionId, documentId)</c> contract, the safety loop,
    /// and the structured-error vocabulary the tools surface.
    /// </summary>
    public const string SystemPromptGuidance = """
        You are editing Microsoft Office documents through the OfficeAgent tools - a Word document (.docx) or a PowerPoint deck (.pptx). inspect_document reports which: format "Word" or "PowerPoint". A few rules below differ by format, and each says so.

        Document addressing
        - Storage connections are host-configured and so is each document's registration. The host gives you an OPAQUE, provider-assigned documentId for every document you are allowed to work with; all document tools address it as (connectionId, documentId). Never invent a documentId, never pass a filename or path as one, and never ask the user to send raw file bytes through this conversation.
        - The connectionId and documentId are already in your instructions or in the conversation context. NEVER ask the user for them - the user does not know or manage these values. If a request mentions "the document", it means the current document you were given; start working with it immediately.
        - apply_plan returns outputDocumentId, outputName, and outputContentType for the saved revision. Use outputDocumentId as the next call's documentId if you keep editing. When the work is complete, tell the user the document is ready; the host retrieves its bytes and presents the download or attachment. Do not place document base64 in the final response.

        Plan shape, anchors, safety loop
        - Plan body is { "operations": [ ... ] }. Do NOT set contractVersion or snapshot - the engine fills them.
        - Available operations (the JSON shape of each is in the preview_plan description): changeText, insert (a paragraph), insertTable, removeTable, format, fill, comment, setProperty, revision, insertTableRows, removeTableRows, insertTableColumns, removeTableColumns, insertImage, removeImage, copyStyles, clearStyles. Create a table with insertTable; delete one with removeTable. These are plan operations inside preview_plan/apply_plan, not separate tools.
        - Call inspect_document or find_in_document before building a plan to obtain anchor ids; never invent paragraph ids, occurrence numbers, content-control tags, or node paths.
        - Tables and images only appear in inspect_document.nodes, never in the paragraphs list. Copy the path from there rather than composing one: Word uses "table#N"/"image#N", a deck uses "table#{slideId}/{shapeId}"/"image#{slideId}/{shapeId}". To recognise table content, look for paragraphs whose `in` field matches a table path.
        - Preview before you apply. If preview reports stale-snapshot, re-inspect and rebuild. If preview reports expect-mismatch, the document drifted - re-inspect/find that operation.
        - Change mode: in a Word document default to "Tracked" unless the user explicitly approves direct edits. A PowerPoint deck has no tracked-changes representation and REFUSES mode "Tracked", so use "Direct" there and say that edits to a deck cannot be redlined; add a comment if the change needs to be flagged for review.
        - Reject operations that need a renderer (pagination, field recalculation); explain the limitation instead.

        Working with a PowerPoint deck
        - Each slide is one outline entry. Paragraph ids read "slide{slideId}/shape{shapeId}/p{n}", with "notes/..." for speaker notes and ".../r{row}c{col}/..." inside a table cell.
        - A slide has no text flow, so insertTable, insertImage, and an added comment target the SLIDE - { "kind": "slide", "path": "slide#256" } - not a paragraph. Resolve a comment with { "op": "comment", "action": "Resolve", "target": { "kind": "comment", "path": "comment#256/{id}" } }.
        - Only these verbs work on a deck: changeText, format, insertTable, removeTable, insertTableRows, removeTableRows, insertTableColumns, removeTableColumns, insertImage, removeImage, comment. The others return "unsupported-operation" and nothing in the plan is applied.
        - format on a deck covers bold, italic, underline, sizeHalfPoints, fontFamily, color, highlight, alignment, and widthPx/heightPx on an image. Word-only measures (styleId, indents, spacing, borders) are refused rather than ignored. Anchor the span you want styled, or use an empty "expect" to style a whole paragraph.
        - A deck has no paragraph-inserting verb, so to write into an empty placeholder - the state a newly created deck's title is in - use changeText with an empty expect: { "op": "changeText", "target": { "paraId": "slide256/shape2/p0", "expect": "" }, "with": "Quarterly Review", "mode": "Direct" }. That still verifies the paragraph is blank, so it fails rather than overwriting text that drifted in.
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
    /// System-prompt guidance to append when <c>create_document</c> is exposed.
    /// </summary>
    public const string CreationPromptGuidance = """

        Creating a document
        - create_document(connectionId, name, planJson) creates and registers a new document and returns outputDocumentId. name is a bare file name, never a path; an existing name is not overwritten. The extension picks the format: .docx makes a Word document, .pptx makes a PowerPoint deck.
        - Pass "" for an empty document. An initial plan is applied in memory before storage. The empty starting anchor differs by format: a Word document has one empty paragraph at { "paraId": "auto-0000", "expect": "" }; a deck has one empty title placeholder at { "paraId": "slide256/shape2/p0", "expect": "" }. When unsure, create with planJson "" and then inspect_document.
        - planJson accepts a bare operations array [ … ] as well as { "operations": [ … ] }.
        - Plan-validation errors mean nothing was written. A provider or cancellation error can occur after storage accepted the file, so do not retry the same name; report the possibly unregistered file name to the host/operator for recovery.
        """;

    /// <summary>Returns the four core AIFunctions the host registers with its agent.</summary>
    public AIFunction[] AsAIFunctions() => AsAIFunctions(new OfficeAgentToolsOptions());

    /// <summary>
    /// Returns the AIFunctions selected by <paramref name="options"/>: the four core
    /// inspect/find/preview/apply tools, plus the source-addressed tools
    /// (<c>register_document</c>, <c>remove_document</c>, <c>open_document</c>,
    /// <c>edit_document</c>) when registration is allowed, and independently
    /// <c>create_document</c> when creation is allowed. The composites are gated with
    /// registration because they take the same connection-relative source it does.
    /// </summary>
    public AIFunction[] AsAIFunctions(OfficeAgentToolsOptions options)
    {
        if (options is null) throw new ArgumentNullException(nameof(options));
        var functions = new List<AIFunction>(CoreFunctions());
        if (options.AllowRegistration)
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
        if (options.AllowCreation)
        {
            functions.Add(AIFunctionFactory.Create(CreateDocument, Opts(
                "create_document",
                "Create and register a new document in a host-configured connection, optionally applying an initial plan before writing. " +
                "name is a bare file name such as 'quarterly-report.docx'; an existing name is never overwritten. " +
                "The extension picks the format: '.docx' makes a Word document, '.pptx' makes a PowerPoint deck. Use list_connections to see which connections accept which. " +
                "Pass planJson \"\" for a minimal document. The starting anchor differs by format: a Word document has one empty paragraph at { \"paraId\": \"auto-0000\", \"expect\": \"\" }; a deck has one empty title placeholder at { \"paraId\": \"slide256/shape2/p0\", \"expect\": \"\" }, and its slide-targeted verbs use { \"kind\": \"slide\", \"path\": \"slide#256\" }. " +
                "Plan-validation errors guarantee no write. Provider and cancellation errors may occur after storage accepted the file, so do not retry the same name; report the possibly unregistered name to the host for recovery. " +
                "Returns the apply_plan shape: {isValid, committed, outputConnectionId, outputDocumentId, outputVersion, outputName, outputContentType, changes, errors}.")));
        }
        return functions.ToArray();
    }

    private AIFunction[] CoreFunctions() => new[]
    {
        AIFunctionFactory.Create(InspectDocument, Opts(
            "inspect_document",
            "Inspect a document by (connectionId, documentId) - a Word document or a PowerPoint deck. Returns outline (headings, or one entry per slide), paragraphs (with their `in` containment - a table path in Word, a slide's shape or table cell in a deck), content controls, nodes (tables/images/docProperties/revisions in Word; slides/tables/images/comments in a deck - paths for node-targeted operations come from here), styles, and a snapshot etag for drift detection. Use paragraphOffset/paragraphLimit to page; fidelity='outline'|'structure'|'content' to control payload size.")),
        AIFunctionFactory.Create(FindInDocument, Opts(
            "find_in_document",
            "Find text in a document by (connectionId, documentId) - a Word document or a PowerPoint deck, including slide notes and table cells. Returns content-verified anchors (paragraphId + expected + occurrence) usable as plan targets.")),
        AIFunctionFactory.Create(PreviewPlan, Opts(
            "preview_plan",
            "Dry-run a DocumentPlan JSON against (connectionId, documentId). Returns {isValid, changes, errors} without writing. " +
            "Plan shape: { \"operations\": [ ... ] }. Do NOT set contractVersion or snapshot. Each operation is one object. Concrete examples:\n\n" +
            "// Replace text:\n" +
            "{ \"op\": \"changeText\", \"target\": { \"paraId\": \"w14:...\", \"expect\": \"Acme Corp\", \"occurrence\": 0 }, \"with\": \"Globex Inc.\", \"mode\": \"Tracked\" }\n\n" +
            "// Unified formatting (paragraph/run/table/row/cell/image):\n" +
            "{ \"op\": \"format\", \"target\": { \"paraId\": \"w14:...\", \"expect\": \"important\", \"occurrence\": 0 }, \"highlight\": \"yellow\", \"bold\": true, \"color\": \"FF0000\" }\n" +
            "{ \"op\": \"format\", \"target\": { \"kind\": \"table\",     \"path\": \"table#0\" }, \"styleId\": \"TableGrid\", \"borderStyle\": \"single\" }\n" +
            "{ \"op\": \"format\", \"target\": { \"kind\": \"image\",     \"path\": \"image#0\" }, \"widthPx\": 320, \"heightPx\": 200 }\n\n" +
            "// Fill / comment / insert paragraph / setProperty / revision:\n" +
            "{ \"op\": \"fill\", \"target\": { \"tag\": \"ClientName\" }, \"value\": \"Globex\" }\n" +
            "{ \"op\": \"comment\", \"target\": { \"paraId\": \"w14:...\", \"expect\": \"...\" }, \"text\": \"Confirm this.\" }\n" +
            "{ \"op\": \"insert\", \"target\": { \"paraId\": \"w14:...\", \"expect\": \"...\" }, \"position\": \"After\", \"text\": \"New paragraph.\" }\n" +
            "{ \"op\": \"setProperty\", \"target\": { \"kind\": \"docProperty\", \"path\": \"core/title\" }, \"value\": \"My Title\" }\n" +
            "{ \"op\": \"revision\",   \"target\": { \"kind\": \"revision\",    \"path\": \"all\" }, \"action\": \"Accept\" }\n\n" +
            "// Insert a whole new table after a paragraph, or remove an entire table (table path from inspect_document.nodes):\n" +
            "{ \"op\": \"insertTable\", \"target\": { \"paraId\": \"w14:...\", \"expect\": \"...\" }, \"position\": \"After\", \"table\": { \"headers\": [\"Region\", \"Q1\"], \"rows\": [[\"NL\", \"41850\"]] } }\n" +
            "{ \"op\": \"removeTable\",  \"target\": { \"kind\": \"table\", \"path\": \"table#0\" } }\n\n" +
            "// Add or remove table rows / columns; insert or remove image; copy or clear styles. Paths come from inspect_document.nodes:\n" +
            "{ \"op\": \"insertTableRows\", \"target\": { \"kind\": \"table\", \"path\": \"table#0\" }, \"rows\": [[\"NL\",\"17\",\"41850\"]], \"position\": \"End\" }\n" +
            "{ \"op\": \"removeTableRows\", \"target\": { \"kind\": \"table\", \"path\": \"table#0\" }, \"onlyIfEmpty\": true }\n" +
            "{ \"op\": \"insertImage\", \"target\": { \"paraId\": \"w14:...\", \"expect\": \"...\" }, \"base64Bytes\": \"iVBORw0KGgo...\", \"imageType\": \"png\", \"widthPx\": 200, \"heightPx\": 80 }\n" +
            "{ \"op\": \"insertImage\", \"target\": { \"paraId\": \"w14:...\", \"expect\": \"...\" }, \"imageConnectionId\": \"images\", \"imageDocumentId\": \"<opaque id from a prior add>\", \"imageType\": \"png\", \"widthPx\": 200, \"heightPx\": 80 }\n" +
            "{ \"op\": \"removeImage\", \"target\": { \"kind\": \"image\", \"path\": \"image#0\" } }")),
        AIFunctionFactory.Create(ApplyPlan, Opts(
            "apply_plan",
            "Apply a DocumentPlan JSON to (connectionId, documentId) and save through the provider. Returns {committed, outputConnectionId, outputDocumentId, outputVersion, outputName, outputContentType, changes, errors}. saveMode: 'NewVersion' (default, mints a new id under the same connection), 'NewDocument' (mints a fresh id with an optional newName for display), 'Replace' (overwrites the source after an optimistic version check). On any failure nothing is written."))
    };

    /// <summary>Inspects a document and returns paginated JSON.</summary>
    public Task<string> InspectDocument(
        string connectionId,
        string documentId,
        string fidelity = "content",
        int paragraphOffset = 0,
        int paragraphLimit = 200,
        CancellationToken cancellationToken = default)
        => SafeAsync(async () =>
        {
            var options = new InspectOptions { Fidelity = ParseFidelity(fidelity) };
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
        ["styles"] = result.Styles.Styles.Select(s => new { s.Id, s.Name, s.InUseCount })
    };

    /// <summary>Finds text in a document and returns content-verified anchors.</summary>
    public Task<string> FindInDocument(
        string connectionId,
        string documentId,
        string pattern,
        bool regex = false,
        bool wholeWord = false,
        bool caseSensitive = false,
        CancellationToken cancellationToken = default)
        => SafeAsync(async () =>
        {
            var query = new FindQuery
            {
                Pattern = pattern,
                Options = new MatchOptions { Regex = regex, WholeWord = wholeWord, CaseSensitive = caseSensitive }
            };
            var hits = await _client.FindAsync(connectionId, documentId, query, cancellationToken).ConfigureAwait(false);
            return JsonSerializer.Serialize(hits.Select(h => new
            {
                paraId = (h.Anchor as TextSpanAnchor)?.ParaId,
                expect = h.Text,
                occurrence = (h.Anchor as TextSpanAnchor)?.Occurrence ?? 0,
                context = h.Context
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
            var plan = DeserializePlan(planJson);
            var report = await _client.PreviewAsync(connectionId, documentId, plan, cancellationToken).ConfigureAwait(false);
            return SerializeReport(report, committed: false, savedReference: null);
        });

    /// <summary>Applies a plan and saves through the provider.</summary>
    public Task<string> ApplyPlan(
        string connectionId,
        string documentId,
        string planJson,
        string saveMode = "NewVersion",
        string newName = "",
        CancellationToken cancellationToken = default)
        => SafeAsync(async () =>
        {
            var plan = DeserializePlan(planJson);
            var options = new SaveDocumentOptions
            {
                Mode = ParseSaveMode(saveMode),
                NewName = string.IsNullOrEmpty(newName) ? null : newName
            };
            var result = await _client.CommitAsync(connectionId, documentId, plan, options, cancellationToken).ConfigureAwait(false);
            return SerializeReport(result.Report, result.Committed, result.Committed ? result.Document : null);
        });

    /// <summary>Registers a document with a provider connection and returns its opaque id.</summary>
    public Task<string> RegisterDocument(
        string connectionId,
        string source,
        CancellationToken cancellationToken = default)
        => SafeAsync(async () =>
        {
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
            var plan = string.IsNullOrWhiteSpace(planJson) ? null : DeserializePlan(planJson);
            var result = await _client.CreateAsync(connectionId, name, plan, cancellationToken).ConfigureAwait(false);
            return SerializeReport(result.Report, result.Committed, result.Committed ? result.Document : null);
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
        CancellationToken cancellationToken = default)
        => SafeAsync(async () =>
        {
            var reference = await _client.RegisterAsync(connectionId, source, cancellationToken).ConfigureAwait(false);
            var options = new InspectOptions { Fidelity = ParseFidelity(fidelity) };
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
        string saveMode = "NewVersion",
        string newName = "",
        CancellationToken cancellationToken = default)
        => SafeAsync(async () =>
        {
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
                Mode = ParseSaveMode(saveMode),
                NewName = string.IsNullOrEmpty(newName) ? null : newName
            };
            var result = await _client.CommitAsync(
                connectionId, reference.ItemId, DeserializePlan(planObject), options, cancellationToken).ConfigureAwait(false);

            return SerializeReport(
                result.Report, result.Committed, result.Committed ? result.Document : null,
                sourceDocumentId: reference.ItemId);
        });

    /// <summary>Removes a document registration; the underlying content is left untouched.</summary>
    public Task<string> RemoveDocument(
        string connectionId,
        string documentId,
        CancellationToken cancellationToken = default)
        => SafeAsync(async () =>
        {
            await _client.RemoveAsync(connectionId, documentId, cancellationToken).ConfigureAwait(false);
            return JsonSerializer.Serialize(new { removed = true, connectionId, documentId }, Json);
        });

    private static async Task<string> SafeAsync(Func<Task<string>> work)
    {
        try { return await work().ConfigureAwait(false); }
        catch (OperationCanceledException) { return SerializeError("cancelled", "Operation was cancelled."); }
        catch (JsonException ex) { return SerializeError("invalid-json", ex.Message); }
        catch (DocumentProviderException ex)
        {
            return SerializeError(
                ProviderCodeToWire(ex.Code),
                ex.Message,
                ex.Provider, ex.ConnectionId, ex.ItemId);
        }
        catch (Exception ex) { return SerializeError("internal-error", ex.Message); }
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

    private static SaveMode ParseSaveMode(string mode) => mode?.Trim() switch
    {
        "NewDocument" => SaveMode.NewDocument,
        "Replace" => SaveMode.Replace,
        _ => SaveMode.NewVersion
    };

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

    private static DocumentPlan DeserializePlan(string planJson) =>
        DeserializePlan(ParsePlanObject(planJson));

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

    private static string SerializeReport(
        ChangeReport report,
        bool committed,
        DocumentReference? savedReference,
        string? sourceDocumentId = null) =>
        JsonSerializer.Serialize(new
        {
            isValid = report.IsValid,
            committed,
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

    private static object MapOutline(OutlineNode node) => new
    {
        node.Level,
        node.Text,
        paraId = (node.Anchor as TextSpanAnchor)?.ParaId,
        children = node.Children.Select(MapOutline)
    };
}
