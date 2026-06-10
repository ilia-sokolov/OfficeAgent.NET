using OfficeAgent.Abstractions;

namespace OfficeAgent.Core;

/// <summary>
/// Resolves file and in-memory handles to owned memory streams. Always returns a
/// fresh copy, so the engine never disposes or mutates a caller-owned stream.
/// </summary>
internal sealed class DefaultHandleResolver : IHandleResolver
{
    public static IReadOnlyList<IHandleResolver> All { get; } = new IHandleResolver[] { new DefaultHandleResolver() };

    public bool CanResolve(DocumentHandle handle) => handle is FileHandle or StreamHandle;

    public Stream Resolve(DocumentHandle handle)
    {
        switch (handle)
        {
            case FileHandle file:
                return new MemoryStream(File.ReadAllBytes(file.Path));

            case StreamHandle streamHandle:
                var source = streamHandle.Stream;
                if (source.CanSeek)
                    source.Position = 0;
                var copy = new MemoryStream();
                source.CopyTo(copy);
                copy.Position = 0;
                return copy;

            default:
                throw new NotSupportedException($"Cannot resolve handle of type {handle.GetType().Name}.");
        }
    }
}
