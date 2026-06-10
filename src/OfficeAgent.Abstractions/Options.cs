namespace OfficeAgent.Abstractions;

/// <summary>
/// Specifies how much document detail inspection should return.
/// </summary>
public enum Fidelity
{
    /// <summary>Return outline information only.</summary>
    Outline,

    /// <summary>Return document structure such as headings, tables, and styles.</summary>
    Structure,

    /// <summary>Return the full anchored content model.</summary>
    Content
}

/// <summary>
/// Provides options for document inspection.
/// </summary>
public sealed class InspectOptions
{
    /// <summary>Gets the requested inspection fidelity.</summary>
    public Fidelity Fidelity { get; init; } = Fidelity.Content;

    /// <summary>Gets the default inspection options.</summary>
    public static InspectOptions Default { get; } = new();
}

/// <summary>
/// Provides options for applying a document plan.
/// </summary>
public sealed class ApplyOptions
{
    /// <summary>Gets a value indicating whether the plan should be previewed without writing.</summary>
    public bool DryRun { get; init; } = true;

    /// <summary>Gets options for previewing a plan.</summary>
    public static ApplyOptions Preview { get; } = new() { DryRun = true };

    /// <summary>Gets options for committing a plan.</summary>
    public static ApplyOptions Commit { get; } = new() { DryRun = false };
}

/// <summary>
/// Contains the result of applying or previewing a plan. When committed, the edited
/// document is available through <see cref="ToBytes"/>, <see cref="Save"/>, or
/// <see cref="SaveAsync"/>. Disposing the result releases the underlying output buffer.
/// </summary>
public sealed class ApplyResult : IDisposable
{
    /// <summary>Gets the output document handle when a commit writes to a handle.</summary>
    public DocumentHandle? Output { get; init; }

    /// <summary>Gets the validation and change report.</summary>
    public ChangeReport Report { get; init; } = new();

    /// <summary>Gets a value indicating whether the plan was committed.</summary>
    public bool Committed { get; init; }

    /// <summary>
    /// Returns the committed document bytes.
    /// </summary>
    /// <exception cref="InvalidOperationException">The result was a dry run or failed validation, so it has no output.</exception>
    public byte[] ToBytes()
    {
        if (Output is not StreamHandle handle)
            throw new InvalidOperationException(
                "ApplyResult has no output (the plan was a dry run or did not commit). Check Committed and Report.Errors first.");

        if (handle.Stream is MemoryStream ms)
            return ms.ToArray();

        using var copy = new MemoryStream();
        if (handle.Stream.CanSeek)
            handle.Stream.Position = 0;
        handle.Stream.CopyTo(copy);
        return copy.ToArray();
    }

    /// <summary>Writes the committed document to <paramref name="path"/>.</summary>
    public void Save(string path) => File.WriteAllBytes(path, ToBytes());

    /// <summary>Asynchronously writes the committed document to <paramref name="path"/>.</summary>
    public Task SaveAsync(string path, CancellationToken cancellationToken = default)
    {
        var bytes = ToBytes();
#if NET8_0_OR_GREATER
        return File.WriteAllBytesAsync(path, bytes, cancellationToken);
#else
        cancellationToken.ThrowIfCancellationRequested();
        return Task.Run(() => File.WriteAllBytes(path, bytes), cancellationToken);
#endif
    }

    /// <summary>Releases the underlying output buffer, if any.</summary>
    public void Dispose()
    {
        if (Output is StreamHandle handle)
            handle.Stream.Dispose();
    }
}
