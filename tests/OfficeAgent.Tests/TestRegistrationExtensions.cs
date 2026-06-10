using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.Tests;

/// <summary>
/// Test plumbing for the register-by-path provider model: stages a byte fixture under
/// a workspace root and hands the resulting path to RegisterAsync, returning the
/// canonical reference the producer code under test expects.
/// </summary>
internal static class TestRegistrationExtensions
{
    public static async Task<DocumentReference> RegisterBytesAsync(
        this OfficeAgentClient client,
        string connectionId,
        string root,
        byte[] bytes,
        string name,
        CancellationToken cancellationToken = default)
    {
        var path = StageFixture(root, bytes, name);
        return await client.RegisterAsync(connectionId, path, cancellationToken).ConfigureAwait(false);
    }

    public static Task<DocumentReference> RegisterBytesAsync(
        this IDocumentProvider provider,
        string root,
        byte[] bytes,
        string name,
        CancellationToken cancellationToken = default)
    {
        var path = StageFixture(root, bytes, name);
        return provider.RegisterAsync(path, cancellationToken);
    }

    public static string StageFixture(string root, byte[] bytes, string name)
    {
        Directory.CreateDirectory(root);
        // A per-call subdirectory keeps each fixture's display name intact while
        // sidestepping collisions when several tests register the same name.
        var directory = Path.Combine(root, Guid.NewGuid().ToString("N"));
        Directory.CreateDirectory(directory);
        var path = Path.Combine(directory, name);
        File.WriteAllBytes(path, bytes);
        return path;
    }
}
