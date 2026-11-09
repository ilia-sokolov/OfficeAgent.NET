namespace OfficeAgent.Core.DocumentProviders;

/// <summary>
/// Optional <see cref="IDocumentProvider"/> capability: list the documents a connection can
/// reach, so an agent can find a document or a template instead of being told its source.
/// Kept separate from <see cref="IDocumentProvider"/> so a custom provider keeps compiling and
/// is simply reported as not listable.
/// </summary>
public interface IDocumentListingProvider : IDocumentProvider
{
    /// <summary>
    /// Returns one page of the documents that match <paramref name="query"/>. Every entry's
    /// <see cref="DocumentListEntry.Source"/> is accepted by <see cref="IDocumentProvider.RegisterAsync"/>
    /// on the same connection, and nothing is registered by listing.
    /// </summary>
    Task<DocumentListPage> ListAsync(DocumentListQuery query, CancellationToken cancellationToken = default);
}

/// <summary>What to list in one <see cref="IDocumentListingProvider.ListAsync"/> call.</summary>
public sealed class DocumentListQuery
{
    /// <summary>
    /// Gets the provider-relative folder to list. <see langword="null"/> or empty means the
    /// connection root, or its templates folder when <see cref="Templates"/> is set.
    /// </summary>
    public string? Folder { get; init; }

    /// <summary>
    /// Gets whether to list the connection's configured templates folder instead of
    /// <see cref="Folder"/>. A connection with no templates folder returns an empty page.
    /// </summary>
    public bool Templates { get; init; }

    /// <summary>Gets a case-insensitive substring the file name must contain, or <see langword="null"/> for any name.</summary>
    public string? NameContains { get; init; }

    /// <summary>
    /// Gets the extensions to list, such as <c>.docx</c>. <see langword="null"/> means every
    /// extension the connection allows. It narrows the connection's allow-list and never widens it.
    /// </summary>
    public IReadOnlyList<string>? Extensions { get; init; }

    /// <summary>Gets whether documents in subfolders are listed too, within the provider's documented bounds.</summary>
    public bool Recursive { get; init; }

    /// <summary>Gets the largest page to return. Providers clamp it to 1..500.</summary>
    public int MaximumResults { get; init; } = 100;

    /// <summary>Gets the opaque token from a previous page, or <see langword="null"/> for the first page.</summary>
    public string? ContinuationToken { get; init; }
}

/// <summary>One document a listing found.</summary>
public sealed class DocumentListEntry
{
    /// <summary>
    /// Gets exactly what <see cref="IDocumentProvider.RegisterAsync"/> accepts for this
    /// connection, so the entry can be registered, opened or edited as listed.
    /// </summary>
    public string Source { get; init; } = "";

    /// <summary>Gets the file name, extension included.</summary>
    public string Name { get; init; } = "";

    /// <summary>
    /// Gets the folder the document is in, relative to the listing root, for display. It is
    /// never an absolute local path. <see langword="null"/> for a document at the root.
    /// </summary>
    public string? Folder { get; init; }

    /// <summary>Gets the document's media type.</summary>
    public string ContentType { get; init; } = "";

    /// <summary>Gets the size in bytes, when the provider reports it.</summary>
    public long? Length { get; init; }

    /// <summary>Gets when the document last changed, when the provider reports it.</summary>
    public DateTimeOffset? LastModified { get; init; }

    /// <summary>
    /// Gets whether the document is a template: a <c>.dotx</c>, <c>.potx</c> or <c>.xltx</c>
    /// file, or any document listed from a templates folder.
    /// </summary>
    public bool IsTemplate { get; init; }
}

/// <summary>One page of a listing.</summary>
public sealed class DocumentListPage
{
    /// <summary>Gets the documents on this page.</summary>
    public IReadOnlyList<DocumentListEntry> Items { get; init; } = Array.Empty<DocumentListEntry>();

    /// <summary>Gets the token that fetches the next page, or <see langword="null"/> when this is the last one.</summary>
    public string? ContinuationToken { get; init; }
}

/// <summary>Where a connection's templates come from, as agents are told it.</summary>
public enum TemplateLocationKind
{
    /// <summary>The connection has no templates.</summary>
    None,

    /// <summary>A templates folder the host configured inside an ordinary connection.</summary>
    Connection,

    /// <summary>The user's own templates, such as Office's personal templates folder.</summary>
    Personal,

    /// <summary>Templates shared with a team, such as Office's workgroup templates folder.</summary>
    Workgroup,

    /// <summary>The organization's template library, such as a SharePoint organization assets library.</summary>
    Organization
}

/// <summary>
/// Optional <see cref="IDocumentProvider"/> capability: the templates folder a host configured
/// for a connection. Separate from <see cref="IDocumentListingProvider"/> so a provider can list
/// without having templates.
/// </summary>
public interface IConnectionTemplateSource
{
    /// <summary>
    /// Gets the templates folder as agents are shown it: a relative folder, a URL or an id.
    /// <see langword="null"/> when the connection has none.
    /// </summary>
    string? TemplatesFolder { get; }

    /// <summary>Gets what kind of template location the connection is.</summary>
    TemplateLocationKind TemplateLocation { get; }
}
