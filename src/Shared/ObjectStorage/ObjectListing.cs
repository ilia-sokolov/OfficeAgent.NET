// Compiled into OfficeAgent.AzureBlob and OfficeAgent.AmazonS3 as linked source, so both
// object-store providers list a container or bucket by the same rules.

using OfficeAgent.Core.DocumentProviders;

namespace OfficeAgent.ObjectStorage;

/// <summary>One object a delimited listing returned.</summary>
internal sealed class ListedObject
{
    public string Key { get; init; } = string.Empty;

    public long? Size { get; init; }

    public DateTimeOffset? LastModified { get; init; }
}

/// <summary>One page of a delimited listing: the objects directly under a prefix, and the next level of prefixes.</summary>
internal sealed class ListedObjectPage
{
    public IReadOnlyList<ListedObject> Objects { get; init; } = Array.Empty<ListedObject>();

    public IReadOnlyList<string> Prefixes { get; init; } = Array.Empty<string>();

    /// <summary>The service's continuation token, or <see langword="null"/> after the last page.</summary>
    public string? NextToken { get; init; }
}

/// <summary>Reads one delimited page under a full prefix, mapping the store's failures to provider errors.</summary>
internal delegate Task<ListedObjectPage> ListedObjectPageReader(string prefix, string? continuationToken, CancellationToken cancellationToken);

/// <summary>
/// Lists a container or bucket for <see cref="IDocumentListingProvider"/>. A listing reads one
/// prefix level per request with the <c>/</c> delimiter, so a recursive listing descends through
/// the returned prefixes and never reads below <see cref="CloudListing.MaximumDepth"/> levels.
/// </summary>
internal static class ObjectListing
{
    /// <summary>
    /// Lists the connection's prefix, or its templates prefix, below the query's folder. Every
    /// prefix read starts with the connection's prefix: the query's folder is checked for escapes,
    /// and a continuation token is encrypted, signed and checked again against the listed prefix.
    /// </summary>
    internal static async Task<DocumentListPage> ListAsync(
        DocumentListQuery query,
        string connectionId,
        string prefix,
        string? templatesPrefix,
        IEnumerable<string> allowedExtensions,
        long maximumBytes,
        Func<ProviderErrorCode, string, DocumentProviderException> error,
        ListedObjectPageReader read,
        CancellationToken cancellationToken)
    {
        if (query is null) throw new ArgumentNullException(nameof(query));
        var request = CloudListingRequest.Create(query, allowedExtensions, out var problem);
        if (request is null)
        {
            var (code, message) = CloudListing.FolderError(problem);
            throw error(code, message);
        }

        string root;
        if (request.Templates)
        {
            if (templatesPrefix is null) return new DocumentListPage();
            root = templatesPrefix;
        }
        else
        {
            root = prefix;
        }

        var listed = request.FolderPath is { } folder ? root + folder + "/" : root;
        if (listed.Length - prefix.Length > ObjectKeys.MaximumKeyLength)
            throw error(ProviderErrorCode.InvalidArgument, "The listing folder is longer than an object name can be.");

        CloudListingCursor cursor;
        if (string.IsNullOrEmpty(query.ContinuationToken))
            cursor = CloudListing.Start(listed, request);
        else if (!CloudListingCursor.TryUnprotect(query.ContinuationToken!, connectionId, request.Fingerprint, out cursor) ||
                 !StaysUnder(cursor, listed))
            throw error(ProviderErrorCode.InvalidArgument, CloudListing.InvalidTokenMessage);

        return await CloudListing.WalkAsync(
            request,
            cursor,
            async (folderPrefix, token, ct) => Children(
                await read(folderPrefix, token, ct).ConfigureAwait(false), folderPrefix, prefix, maximumBytes),
            CloudListing.MaximumServiceCalls,
            connectionId,
            cancellationToken).ConfigureAwait(false);
    }

    /// <summary>A signed token is trusted, but every prefix it names is still checked against the listed one.</summary>
    private static bool StaysUnder(CloudListingCursor cursor, string listed) =>
        (cursor.Current is null || cursor.Current.Id.StartsWith(listed, StringComparison.Ordinal)) &&
        cursor.Pending.All(folder => folder.Id.StartsWith(listed, StringComparison.Ordinal));

    /// <summary>
    /// Turns a page into children. A document's id is its name relative to the connection's
    /// prefix, which is exactly what registration accepts; a folder's id is the full prefix the
    /// next level is listed under.
    /// </summary>
    private static CloudListingPage Children(ListedObjectPage page, string folderPrefix, string prefix, long maximumBytes)
    {
        var children = new List<CloudListingChild>(page.Objects.Count + page.Prefixes.Count);
        foreach (var item in page.Objects)
        {
            // A zero-length "folder" placeholder is named after the prefix itself.
            if (!item.Key.StartsWith(folderPrefix, StringComparison.Ordinal) || item.Key.Length == folderPrefix.Length)
                continue;
            var name = item.Key.Substring(folderPrefix.Length);
            var relative = item.Key.Substring(prefix.Length);
            children.Add(new CloudListingChild
            {
                Id = relative,
                Name = name,
                Unusable = name.IndexOf('/') >= 0 ||
                           ObjectKeys.Check(relative) != ObjectKeyProblem.None ||
                           item.Size > maximumBytes,
                Length = item.Size,
                LastModified = item.LastModified
            });
        }
        foreach (var next in page.Prefixes)
        {
            if (!next.StartsWith(folderPrefix, StringComparison.Ordinal)) continue;
            var name = next.Substring(folderPrefix.Length).TrimEnd('/');
            if (name.Length == 0 || name.IndexOf('/') >= 0 || name == "." || name == "..") continue;
            children.Add(new CloudListingChild { Id = next, Name = name, IsFolder = true });
        }
        return new CloudListingPage { Items = children, NextToken = string.IsNullOrEmpty(page.NextToken) ? null : page.NextToken };
    }
}
