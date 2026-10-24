using ModelContextProtocol.Server;

namespace OfficeAgent.Mcp;

/// <summary>
/// The MCP annotations every tool this server publishes carries: a title people read, and the
/// hints clients use to tell a tool that only reads from one that writes. Some clients approve
/// read-only tools without asking and confirm the rest; a tool without hints counts as
/// destructive, so every tool states all of them.
/// </summary>
/// <remarks>
/// <para>
/// A tool is read-only when it changes nothing outside the call: no file, registration, or
/// session document. <c>create_document_content</c> and <c>edit_document_content</c> are
/// read-only for that reason; they hand the bytes back and store nothing.
/// </para>
/// <para>
/// A writing tool is destructive when it can replace content that already exists.
/// <c>apply_plan</c>, <c>edit_document</c>, <c>edit_documents</c> and <c>refresh_document</c> save over the
/// document in place by default.
/// The tools that create documents refuse to overwrite one, and <c>remove_document</c> removes
/// a registration but never the file, so none of them is destructive.
/// </para>
/// <para>
/// No tool is open-world: each one acts only on the connections the host configured.
/// </para>
/// </remarks>
internal static class ToolAnnotations
{
    /// <summary>One tool's annotations.</summary>
    internal sealed record Hints(string Title, bool ReadOnly, bool Destructive = false, bool Idempotent = false);

    private static Hints Reads(string title) => new(title, ReadOnly: true, Idempotent: true);

    private static Hints Writes(string title, bool destructive = false, bool idempotent = false) =>
        new(title, ReadOnly: false, destructive, idempotent);

    /// <summary>Every tool the server can publish, by name.</summary>
    internal static readonly IReadOnlyDictionary<string, Hints> ByName = new Dictionary<string, Hints>(StringComparer.Ordinal)
    {
        ["describe_capabilities"] = Reads("Describe capabilities"),
        ["list_connections"] = Reads("List connections"),
        ["inspect_document"] = Reads("Inspect a document"),
        ["find_in_document"] = Reads("Find in a document"),
        ["preview_plan"] = Reads("Preview an edit plan"),
        ["compare_documents"] = Reads("Compare two documents"),
        ["preview_document_merge"] = Reads("Preview a document merge"),
        ["discover_template"] = Reads("Discover a template's slots"),
        ["preview_template_batch"] = Reads("Preview a template batch"),
        ["inspect_document_content"] = Reads("Inspect document content"),
        ["create_document_content"] = Reads("Create document content"),
        ["edit_document_content"] = Reads("Edit document content"),
        ["export_document_content"] = Reads("Export a session document"),

        ["apply_plan"] = Writes("Apply an edit plan", destructive: true),
        ["edit_document"] = Writes("Edit a document", destructive: true),
        ["edit_documents"] = Writes("Edit several documents", destructive: true),
        ["refresh_document"] = Writes("Refresh a document from a workbook", destructive: true),
        ["create_document"] = Writes("Create a document"),
        ["merge_documents"] = Writes("Merge documents"),
        ["populate_template_batch"] = Writes("Populate a template batch"),
        ["register_document"] = Writes("Register a document"),
        ["open_document"] = Writes("Open a document"),
        ["import_document_content"] = Writes("Import document content"),
        ["remove_document"] = Writes("Remove a document registration", idempotent: true),
    };

    /// <summary>
    /// The create options that publish a tool with its annotations. A tool missing from
    /// <see cref="ByName"/> is refused, so a new tool cannot ship without them.
    /// </summary>
    internal static McpServerToolCreateOptions For(string name)
    {
        if (!ByName.TryGetValue(name, out var hints))
            throw new InvalidOperationException(
                $"Tool '{name}' has no MCP annotations. Add it to {nameof(ToolAnnotations)}.{nameof(ByName)}.");

        return new McpServerToolCreateOptions
        {
            Title = hints.Title,
            ReadOnly = hints.ReadOnly,
            Destructive = hints.Destructive,
            Idempotent = hints.Idempotent,
            OpenWorld = false
        };
    }
}
