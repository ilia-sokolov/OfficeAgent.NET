namespace OfficeAgent.Abstractions;

/// <summary>
/// Represents a caller-supplied reference to a document.
/// </summary>
public abstract class DocumentHandle
{
}

/// <summary>
/// References a document stored on the local file system.
/// </summary>
public sealed class FileHandle : DocumentHandle
{
    /// <summary>
    /// Gets the local document path.
    /// </summary>
    public string Path { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="FileHandle"/> class.
    /// </summary>
    /// <param name="path">The local document path.</param>
    public FileHandle(string path) => Path = path;

    /// <inheritdoc/>
    public override string ToString() => $"file:{Path}";
}

/// <summary>
/// References a document provided as a stream.
/// </summary>
/// <remarks>
/// The engine copies the stream contents into an owned in-memory buffer and does
/// not dispose the caller's stream. If the stream is seekable, its position is
/// reset to 0 before reading; save and restore the position in the caller if that
/// matters.
/// </remarks>
public sealed class StreamHandle : DocumentHandle
{
    /// <summary>
    /// Gets the source stream.
    /// </summary>
    public Stream Stream { get; }

    /// <summary>
    /// Gets an optional display name for diagnostics.
    /// </summary>
    public string? Name { get; }

    /// <summary>
    /// Initializes a new instance of the <see cref="StreamHandle"/> class.
    /// </summary>
    /// <param name="stream">The source stream.</param>
    /// <param name="name">An optional display name.</param>
    public StreamHandle(Stream stream, string? name = null)
    {
        Stream = stream;
        Name = name;
    }

    /// <inheritdoc/>
    public override string ToString() => $"stream:{Name ?? "(unnamed)"}";
}

