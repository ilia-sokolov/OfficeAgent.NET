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

/// <summary>The bounded result of an optional visual rendering operation.</summary>
public sealed class RenderResult
{
    /// <summary>Gets whether rendering completed within every configured limit.</summary>
    public bool Succeeded { get; init; }

    /// <summary>Gets a stable failure code, or null on success.</summary>
    public string? FailureCode { get; init; }

    /// <summary>Gets a safe failure description that excludes process output and document content.</summary>
    public string? Message { get; init; }

    /// <summary>Gets page images in document order.</summary>
    public IReadOnlyList<RenderedPage> Pages { get; init; } = Array.Empty<RenderedPage>();
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
