using OfficeAgent.Abstractions;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.AmazonS3;

/// <summary>Configures one Amazon S3 provider connection: one bucket and an optional key prefix.</summary>
public sealed class AmazonS3DocumentProviderOptions
{
    /// <summary>Gets or sets the provider connection identifier.</summary>
    public string ConnectionId { get; set; } = string.Empty;

    /// <summary>Gets or sets the bucket this connection is confined to.</summary>
    public string BucketName { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets an optional key prefix, such as <c>contracts/</c>. Documents are registered by
    /// keys relative to it, and every key must stay inside it.
    /// </summary>
    public string KeyPrefix { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets an optional folder of templates relative to <see cref="KeyPrefix"/>, such as
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
    /// Gets or sets the AWS region, such as <c>eu-west-1</c>. Empty takes the region from the AWS
    /// SDK's default chain (<c>AWS_REGION</c>, the shared config file, or instance metadata).
    /// </summary>
    public string Region { get; set; } = string.Empty;

    /// <summary>
    /// Gets or sets an endpoint for an S3-compatible store, such as MinIO or Cloudflare R2. Empty
    /// uses Amazon S3. The store must support conditional writes; see the provider documentation.
    /// </summary>
    public string ServiceUrl { get; set; } = string.Empty;

    /// <summary>Gets or sets whether requests use path-style addressing, which most S3-compatible stores need.</summary>
    public bool ForcePathStyle { get; set; }

    /// <summary>
    /// Gets or sets a JSON file that keeps registrations across restarts. Empty keeps them in
    /// memory. The file suits one host process; several instances need an
    /// <see cref="IAmazonS3RegistrationStore"/> over shared storage.
    /// </summary>
    public string RegistrationIndexPath { get; set; } = string.Empty;

    /// <summary>Gets or sets the maximum accepted document size in bytes. A single PUT accepts at most 5 GB.</summary>
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
/// Persists the opaque registration id → object key mapping for one S3 connection. The store
/// holds only keys - never bytes, URLs, or credentials. Implement it over shared storage when
/// several host instances serve the same connection.
/// </summary>
public interface IAmazonS3RegistrationStore
{
    /// <summary>Mints a new opaque registration id for an object key and persists the mapping.</summary>
    Task<string> AddAsync(string key, CancellationToken cancellationToken = default);

    /// <summary>Returns the object key for a registration, or <see langword="null"/> when unregistered.</summary>
    Task<string?> ResolveAsync(string registrationId, CancellationToken cancellationToken = default);

    /// <summary>Removes a registration. Returns <see langword="false"/> when it did not exist.</summary>
    Task<bool> RemoveAsync(string registrationId, CancellationToken cancellationToken = default);
}
