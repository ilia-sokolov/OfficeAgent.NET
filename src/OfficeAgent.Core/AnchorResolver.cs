using OfficeAgent.Abstractions;

namespace OfficeAgent.Core;

/// <summary>
/// Document-level staleness check. Span-level content verification (the
/// <see cref="TextSpanAnchor.Expect"/> guard) is enforced inside each format
/// module's handlers, where the live text is available.
/// </summary>
internal sealed class AnchorResolver
{
    public bool VerifyNotStale(SnapshotToken? planSnapshot, SnapshotToken current)
    {
        if (planSnapshot is null || string.IsNullOrEmpty(planSnapshot.ETag))
            return true;

        return string.Equals(planSnapshot.ETag, current.ETag, StringComparison.Ordinal);
    }
}
