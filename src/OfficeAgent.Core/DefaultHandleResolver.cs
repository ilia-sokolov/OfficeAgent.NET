using OfficeAgent.Abstractions;

namespace OfficeAgent.Core;

/// <summary>
/// Resolves file and in-memory handles to owned memory streams. Always returns a
/// fresh copy, so the engine never disposes or mutates a caller-owned stream.
/// </summary>
/// <remarks>
/// This copy is the earliest point a caller's bytes enter the engine, earlier than the
/// package open, so the host compressed-size ceiling is applied here as well. A bound
/// that lived only at the open would let an endless stream exhaust memory before
/// anything had a chance to refuse it.
/// </remarks>
internal sealed class DefaultHandleResolver : IHandleResolver
{
    private readonly long _maximumBytes;

    public DefaultHandleResolver() : this(OpenXmlIngestionLimits.Default.MaximumCompressedBytes)
    {
    }

    public DefaultHandleResolver(long maximumBytes) => _maximumBytes = maximumBytes;

    public static IReadOnlyList<IHandleResolver> All { get; } = new IHandleResolver[] { new DefaultHandleResolver() };

    /// <summary>The default resolvers bounded by a host's configured ceiling.</summary>
    public static IReadOnlyList<IHandleResolver> For(OpenXmlIngestionLimits limits) =>
        new IHandleResolver[] { new DefaultHandleResolver(limits.MaximumCompressedBytes) };

    public bool CanResolve(DocumentHandle handle) => handle is FileHandle or StreamHandle;

    public Stream Resolve(DocumentHandle handle)
    {
        switch (handle)
        {
            case FileHandle file:
                var length = new FileInfo(file.Path).Length;
                if (length > _maximumBytes)
                    throw new OpenXmlIngestionLimitException(
                        "MaximumCompressedBytes", _maximumBytes, length);
                return new MemoryStream(File.ReadAllBytes(file.Path));

            case StreamHandle streamHandle:
                var source = streamHandle.Stream;
                if (source.CanSeek)
                    source.Position = 0;
                return new MemoryStream(OpenXmlPackageService.ReadBounded(source, _maximumBytes));

            default:
                throw new NotSupportedException($"Cannot resolve handle of type {handle.GetType().Name}.");
        }
    }
}
