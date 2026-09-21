using System.Text.Json.Serialization;

namespace OfficeAgent.Abstractions;

/// <summary>Options for rendering an Office package to page images.</summary>
public sealed class RenderOptions
{
    /// <summary>Gets the input file name. Its extension selects the Office file format.</summary>
    public string FileName { get; init; } = "document.docx";

    /// <summary>Gets the raster resolution. Defaults to 144 DPI.</summary>
    public int Dpi { get; init; } = 144;

    /// <summary>Gets the total wall-clock limit for all renderer processes.</summary>
    public TimeSpan Timeout { get; init; } = TimeSpan.FromMinutes(2);

    /// <summary>Gets the maximum working set allowed for an individual renderer process.</summary>
    public long MaximumWorkingSetBytes { get; init; } = 512L * 1024 * 1024;

    /// <summary>Gets the maximum number of rendered pages.</summary>
    public int MaximumPages { get; init; } = 200;

    /// <summary>Gets the maximum combined size of rendered output.</summary>
    public long MaximumOutputBytes { get; init; } = 100L * 1024 * 1024;

    /// <summary>Gets the maximum accepted input package size.</summary>
    public long MaximumInputBytes { get; init; } = 100L * 1024 * 1024;
}

/// <summary>One rendered page image.</summary>
public sealed class RenderedPage
{
    /// <summary>Gets the one-based page number.</summary>
    public int PageNumber { get; init; }

    /// <summary>Gets the page image media type.</summary>
    public string ContentType { get; init; } = "image/png";

    /// <summary>Gets the encoded page image.</summary>
    public byte[] Content { get; init; } = Array.Empty<byte>();
}

/// <summary>
/// Thrown when a caller reads rendered pages from a render that did not succeed. Carries the
/// renderer's stable <see cref="FailureCode"/> so a caller that never checked
/// <see cref="RenderResult.Succeeded"/> fails closed with a diagnosable reason instead of
/// treating a failed render as a zero-page document.
/// </summary>
public sealed class RenderFailedException : Exception
{
    /// <summary>Initializes the exception with a renderer failure code and safe description.</summary>
    public RenderFailedException(string failureCode, string message)
        : base(message) => FailureCode = failureCode;

    /// <summary>Gets the stable renderer failure code.</summary>
    public string FailureCode { get; }
}

/// <summary>
/// The bounded result of an optional visual rendering operation. A result is either a success
/// carrying pages or a failure carrying a code; the two states cannot be mixed, and reading
/// pages from a failure throws rather than reporting an empty document.
/// </summary>
public sealed class RenderResult
{
    private readonly IReadOnlyList<RenderedPage> _pages;

    private RenderResult(
        bool succeeded, string? failureCode, string? message, IReadOnlyList<RenderedPage> pages)
    {
        Succeeded = succeeded;
        FailureCode = failureCode;
        Message = message;
        _pages = pages;
    }

    /// <summary>Creates a successful result carrying the rendered pages in document order.</summary>
    public static RenderResult Success(IReadOnlyList<RenderedPage> pages)
    {
        if (pages is null) throw new ArgumentNullException(nameof(pages));
        return new RenderResult(true, failureCode: null, message: null, pages);
    }

    /// <summary>Creates a failed result. A failed render never carries pages.</summary>
    public static RenderResult Failure(string failureCode, string message)
    {
        if (string.IsNullOrWhiteSpace(failureCode))
            throw new ArgumentException("A failure code is required.", nameof(failureCode));
        if (string.IsNullOrWhiteSpace(message))
            throw new ArgumentException("A failure message is required.", nameof(message));
        return new RenderResult(false, failureCode, message, Array.Empty<RenderedPage>());
    }

    /// <summary>Gets whether rendering completed within every configured limit.</summary>
    public bool Succeeded { get; }

    /// <summary>Gets a stable failure code, or null on success.</summary>
    public string? FailureCode { get; }

    /// <summary>Gets a safe failure description that excludes process output and document content.</summary>
    public string? Message { get; }

    /// <summary>
    /// Gets page images in document order. Throws <see cref="RenderFailedException"/> when
    /// rendering did not succeed. Check <see cref="Succeeded"/>, or use
    /// <see cref="TryGetPages"/>, before reading this on a result that may have failed.
    /// </summary>
    /// <exception cref="RenderFailedException">Rendering did not succeed.</exception>
    [JsonIgnore]
    public IReadOnlyList<RenderedPage> Pages
    {
        get
        {
            EnsureSucceeded();
            return _pages;
        }
    }

    /// <summary>
    /// Gets the number of rendered pages. Throws like <see cref="Pages"/> on a failed render, so
    /// a page-count gate cannot read a missing or crashed renderer as a zero-page document.
    /// </summary>
    /// <exception cref="RenderFailedException">Rendering did not succeed.</exception>
    [JsonIgnore]
    public int PageCount
    {
        get
        {
            EnsureSucceeded();
            return _pages.Count;
        }
    }

    /// <summary>Throws <see cref="RenderFailedException"/> unless rendering succeeded.</summary>
    /// <exception cref="RenderFailedException">Rendering did not succeed.</exception>
    public void EnsureSucceeded()
    {
        if (Succeeded) return;
        throw new RenderFailedException(
            FailureCode ?? RenderFailureCodes.RendererFailed, Message ?? "Rendering did not succeed.");
    }

    /// <summary>
    /// Gets the rendered pages without throwing. Returns false and an empty list when rendering
    /// did not succeed, so a caller can branch on the failure explicitly.
    /// </summary>
    public bool TryGetPages(out IReadOnlyList<RenderedPage> pages)
    {
        pages = Succeeded ? _pages : Array.Empty<RenderedPage>();
        return Succeeded;
    }
}

/// <summary>
/// Optional visual rendering boundary. Core inspection and structural preview do not depend
/// on a renderer and remain available when no implementation is registered.
/// </summary>
public interface IDocumentRenderer
{
    /// <summary>Renders an Office package into bounded page images.</summary>
    Task<RenderResult> RenderAsync(
        Stream document,
        RenderOptions options,
        CancellationToken cancellationToken = default);
}
