namespace OfficeAgent.Core.DocumentProviders;

/// <summary>
/// What the SharePoint, iManage, Google Drive, Box, Azure Blob and Amazon S3 providers say when a
/// connection configured as read-only meets a write. Their registrations live in the host's own
/// store, not in the service, so registering and removing a registration stay allowed.
/// </summary>
internal static class ReadOnlyConnections
{
    /// <summary>The refusal a save or a creation meets before any request reaches the service.</summary>
    internal const string WriteRefused =
        "This connection is read-only: documents in it can be registered, listed, opened and used as " +
        "templates, but not saved or created. Nothing was changed.";

    /// <summary>The configuration error for a read-only connection that also names where to create documents.</summary>
    internal static string CreationDestinationNamed(string setting) =>
        $"A read-only connection cannot create documents, so it cannot also name {setting}. " +
        "Remove one or the other.";
}
