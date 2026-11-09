using OfficeAgent.Abstractions;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.AzureBlob;

/// <summary>Configures one Azure Blob Storage provider connection: one container and an optional prefix.</summary>
public sealed class AzureBlobDocumentProviderOptions
{
    /// <summary>Gets or sets the provider connection identifier.</summary>
    public string ConnectionId { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets the Blob service endpoint, for example <c>https://contoso.blob.core.windows.net</c>.
    /// The provider authenticates to it with <c>DefaultAzureCredential</c>: a managed identity in
    /// Azure, a developer sign-in locally. Set this or <see cref="ConnectionString"/>, not both.
    /// </summary>
    public string ServiceUri { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets a storage connection string, the alternative to <see cref="ServiceUri"/>. Read it
    /// from a secret store or the environment; never write one into source or a checked-in file.
    /// </summary>
    public string ConnectionString { get; set; } = string.Empty;

    /// <summary>Gets or sets the container this connection is confined to.</summary>
    public string ContainerName { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets an optional blob-name prefix, such as <c>contracts/</c>. Documents are
    /// registered by names relative to it, and every name must stay inside it.
    /// </summary>
    public string Prefix { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets an optional folder of templates relative to <see cref="Prefix"/>, such as
    /// <c>templates</c>. A templates listing reads it, and every file listed from it counts as a
    /// template. Empty means the connection has no templates folder.
    /// </summary>
    public string TemplatesFolder { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets where <see cref="TemplatesFolder"/> sits, such as
    /// <see cref="TemplateLocationKind.Organization"/> for the organization's approved templates.
    /// Unset means <see cref="TemplateLocationKind.Connection"/> when a templates folder is set.
    /// </summary>
    public TemplateLocationKind? TemplateLocation { get; set; }

    /// <summary>
    /// Gets or sets the client id of a user-assigned managed identity for
    /// <c>DefaultAzureCredential</c>. Leave empty for a system-assigned identity.
    /// </summary>
    public string ManagedIdentityClientId { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets a JSON file that keeps registrations across restarts. Empty keeps them in
    /// memory. The file suits one host process; several instances need an
    /// <see cref="IAzureBlobRegistrationStore"/> over shared storage.
    /// </summary>
    public string RegistrationIndexPath { get; set; } = string.Empty;

    /// <summary>Gets or sets the maximum accepted document size in bytes.</summary>
    public long MaximumBytes { get; set; } = 100 * 1024 * 1024;

    /// <summary>Gets or sets the allowed document extensions.</summary>
    public IReadOnlyCollection<string> AllowedExtensions { get; set; } = new[] { ".docx" };

    /// <summary>
    /// Gets or sets the change mode used when an operation does not state one. The default is
    /// <see cref="ChangeMode.Tracked"/>, which suits documents a person reviews. Set
    /// <see cref="ChangeMode.Direct"/> for generated or machine-owned files, or for decks -
    /// PresentationML refuses Tracked.
    /// </summary>
    public ChangeMode DefaultChangeMode { get; set; } = ChangeMode.Tracked;
    /// <summary>
    /// Gets or sets how much of the matched text a tracked <c>changeText</c> that does not
    /// state a granularity marks. The default, <see cref="ChangeGranularity.Span"/>, marks the
    /// whole matched text; <see cref="ChangeGranularity.Minimal"/> marks only the words that
    /// differ.
    /// </summary>
    public ChangeGranularity DefaultChangeGranularity { get; set; } = ChangeGranularity.Span;

    /// <summary>
    /// Gets or sets whether a workbook previewed or committed through this connection has its
    /// formulas calculated. Takes effect only when a formula evaluator is registered
    /// (<c>AddExcelFormulaCalculation</c>). Defaults to <see langword="false"/>.
    /// </summary>
    public bool CalculateFormulas { get; set; }

    /// <summary>
    /// Gets or sets whether a plan whose calculation produces a new formula error or a circular
    /// reference is refused with <c>formula-error</c>. Applies only with
    /// <see cref="CalculateFormulas"/>. Defaults to <see langword="false"/>.
    /// </summary>
    public bool RejectFormulaErrors { get; set; }
}

/// <summary>
/// Persists the opaque registration id → blob name mapping for one Azure Blob connection. The
/// store holds only blob names - never bytes, URLs, or credentials. Implement it over shared
/// storage when several host instances serve the same connection.
/// </summary>
public interface IAzureBlobRegistrationStore
{
    /// <summary>Mints a new opaque registration id for a blob name and persists the mapping.</summary>
    Task<string> AddAsync(string blobName, CancellationToken cancellationToken = default);

    /// <summary>Returns the blob name for a registration, or <see langword="null"/> when unregistered.</summary>
    Task<string?> ResolveAsync(string registrationId, CancellationToken cancellationToken = default);

    /// <summary>Removes a registration. Returns <see langword="false"/> when it did not exist.</summary>
    Task<bool> RemoveAsync(string registrationId, CancellationToken cancellationToken = default);
}
