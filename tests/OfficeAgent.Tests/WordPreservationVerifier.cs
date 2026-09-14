using System.IO.Compression;
using System.Security.Cryptography;

namespace OfficeAgent.Tests;

internal static class WordPreservationVerifier
{
    public static IReadOnlyDictionary<string, string> PartHashes(byte[] package)
    {
        using var stream = new MemoryStream(package, writable: false);
        using var archive = new ZipArchive(stream, ZipArchiveMode.Read);
        return archive.Entries.ToDictionary(
            entry => entry.FullName,
            entry =>
            {
                using var content = entry.Open();
                return Convert.ToHexString(SHA256.HashData(content)).ToLowerInvariant();
            },
            StringComparer.Ordinal);
    }

    public static IReadOnlyList<string> ChangedParts(byte[] before, byte[] after)
    {
        var beforeHashes = PartHashes(before);
        var afterHashes = PartHashes(after);
        return beforeHashes.Keys.Union(afterHashes.Keys, StringComparer.Ordinal)
            .Where(part => !beforeHashes.TryGetValue(part, out var beforeHash)
                || !afterHashes.TryGetValue(part, out var afterHash)
                || !string.Equals(beforeHash, afterHash, StringComparison.Ordinal))
            .OrderBy(part => part, StringComparer.Ordinal)
            .ToArray();
    }

    public static void Verify(
        byte[] before,
        byte[] after,
        IEnumerable<string> requiredChangedParts,
        IEnumerable<string> protectedParts,
        Action<byte[]> semanticAssertion)
    {
        var expectedChanged = requiredChangedParts.OrderBy(part => part, StringComparer.Ordinal).ToArray();
        var actualChanged = ChangedParts(before, after);
        if (!expectedChanged.SequenceEqual(actualChanged, StringComparer.Ordinal))
            throw new PreservationVerificationException(
                $"Changed parts differ. Expected [{string.Join(", ", expectedChanged)}]; " +
                $"actual [{string.Join(", ", actualChanged)}].");

        var beforeHashes = PartHashes(before);
        var afterHashes = PartHashes(after);
        foreach (var part in protectedParts)
        {
            if (!beforeHashes.TryGetValue(part, out var beforeHash))
                throw new PreservationVerificationException($"Missing protected input part: {part}.");
            if (!afterHashes.TryGetValue(part, out var afterHash))
                throw new PreservationVerificationException($"Missing protected output part: {part}.");
            if (!string.Equals(beforeHash, afterHash, StringComparison.Ordinal))
                throw new PreservationVerificationException($"Protected part changed: {part}.");
        }

        try
        {
            semanticAssertion(after);
        }
        catch (PreservationVerificationException)
        {
            throw;
        }
        catch (Exception error)
        {
            throw new PreservationVerificationException("Changed-part semantic assertion failed.", error);
        }
    }

    public static byte[] ReplacePart(byte[] package, string partName, Func<byte[], byte[]> transform)
    {
        using var stream = new MemoryStream();
        stream.Write(package, 0, package.Length);
        stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            var entry = archive.GetEntry(partName)
                ?? throw new InvalidOperationException($"Package part not found: {partName}.");
            byte[] original;
            using (var input = entry.Open())
            using (var copy = new MemoryStream())
            {
                input.CopyTo(copy);
                original = copy.ToArray();
            }

            var timestamp = entry.LastWriteTime;
            entry.Delete();
            var replacement = archive.CreateEntry(partName, CompressionLevel.Optimal);
            replacement.LastWriteTime = timestamp;
            using var output = replacement.Open();
            var changed = transform(original);
            output.Write(changed, 0, changed.Length);
        }

        return stream.ToArray();
    }
}

internal sealed class PreservationVerificationException : Exception
{
    public PreservationVerificationException(string message) : base(message) { }

    public PreservationVerificationException(string message, Exception innerException)
        : base(message, innerException) { }
}
