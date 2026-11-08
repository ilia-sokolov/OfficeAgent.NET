using OfficeAgent.Abstractions;
using OfficeAgent.Box;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.GoogleDrive;

namespace OfficeAgent.Tests;

/// <summary>
/// Opt-in tests against the real Google Drive and Box services. Each runs only when its
/// credentials are in the environment and is reported as skipped otherwise, so the ordinary
/// build never reaches the network.
/// </summary>
/// <remarks>
/// Each run creates uniquely named files in the configured folder (a document, a replaced
/// revision of it, and a versioned sibling) and leaves them there: providers never delete
/// content. Point the folder at a scratch location.
/// <list type="bullet">
/// <item>Google Drive: <c>OFFICEAGENT_GOOGLEDRIVE_KEY_FILE</c> (path to a service account key
/// JSON) and <c>OFFICEAGENT_GOOGLEDRIVE_FOLDER_ID</c> (a folder shared with the account as
/// editor, in My Drive or a shared drive). Optional: <c>OFFICEAGENT_GOOGLEDRIVE_SUBJECT</c>, a
/// user to impersonate through domain-wide delegation.</item>
/// <item>Box: <c>OFFICEAGENT_BOX_CLIENT_ID</c>, <c>OFFICEAGENT_BOX_CLIENT_SECRET</c>,
/// <c>OFFICEAGENT_BOX_SUBJECT_ID</c> (the enterprise id, or a user id with
/// <c>OFFICEAGENT_BOX_SUBJECT_TYPE=user</c>) and <c>OFFICEAGENT_BOX_FOLDER_ID</c> (a folder the
/// identity can upload to).</item>
/// </list>
/// </remarks>
public sealed class CloudProviderLiveTests
{
    [GoogleDriveLiveFact]
    public async Task Google_drive_round_trip_against_the_real_service()
    {
        using var http = new HttpClient();
        var tokens = new GoogleServiceAccountAccessTokenProvider(new GoogleServiceAccountOptions
        {
            KeyJson = File.ReadAllText(Environment.GetEnvironmentVariable("OFFICEAGENT_GOOGLEDRIVE_KEY_FILE")!),
            Subject = Environment.GetEnvironmentVariable("OFFICEAGENT_GOOGLEDRIVE_SUBJECT") ?? string.Empty
        }, http);
        var provider = new GoogleDriveDocumentProvider(new GoogleDriveDocumentProviderOptions
        {
            ConnectionId = "live-drive",
            CreationFolderId = Environment.GetEnvironmentVariable("OFFICEAGENT_GOOGLEDRIVE_FOLDER_ID")!
        }, http, tokens);

        await RoundTripAsync(provider);
    }

    [BoxLiveFact]
    public async Task Box_round_trip_against_the_real_service()
    {
        using var http = new HttpClient();
        var tokens = new BoxClientCredentialsAccessTokenProvider(new BoxClientCredentialsOptions
        {
            ClientId = Environment.GetEnvironmentVariable("OFFICEAGENT_BOX_CLIENT_ID")!,
            ClientSecret = Environment.GetEnvironmentVariable("OFFICEAGENT_BOX_CLIENT_SECRET")!,
            SubjectId = Environment.GetEnvironmentVariable("OFFICEAGENT_BOX_SUBJECT_ID")!,
            SubjectType = string.Equals(Environment.GetEnvironmentVariable("OFFICEAGENT_BOX_SUBJECT_TYPE"), "user",
                StringComparison.OrdinalIgnoreCase)
                ? BoxSubjectType.User
                : BoxSubjectType.Enterprise
        }, http);
        var provider = new BoxDocumentProvider(new BoxDocumentProviderOptions
        {
            ConnectionId = "live-box",
            CreationFolderId = Environment.GetEnvironmentVariable("OFFICEAGENT_BOX_FOLDER_ID")!
        }, http, tokens);

        await RoundTripAsync(provider);
    }

    /// <summary>
    /// Create → open → replace → stale replace (conflict) → new version → duplicate create
    /// (already exists) → remove.
    /// </summary>
    private static async Task RoundTripAsync(IDocumentCreatingProvider provider)
    {
        var name = $"officeagent-live-{DateTime.UtcNow:yyyyMMddHHmmss}-{Guid.NewGuid():N}".Substring(0, 50) + ".docx";
        var original = DocxFactory.Contract();

        var created = await provider.CreateAsync(name, new MemoryStream(original));
        Assert.Equal(name, created.Name);
        Assert.False(string.IsNullOrEmpty(created.Version));

        using (var opened = await provider.OpenReadAsync(created))
            Assert.Equal(original, ((MemoryStream)opened.Stream).ToArray());

        var changed = original.Concat(new byte[16]).ToArray();
        var replaced = await provider.SaveAsync(created, new MemoryStream(changed), new SaveDocumentOptions { Mode = SaveMode.Replace });
        Assert.Equal(created.ItemId, replaced.ItemId);
        Assert.NotEqual(created.Version, replaced.Version);

        await Assert.ThrowsAsync<DocumentVersionConflictException>(() =>
            provider.SaveAsync(created, new MemoryStream(original), new SaveDocumentOptions { Mode = SaveMode.Replace }));

        var sibling = await provider.SaveAsync(replaced, new MemoryStream(original), new SaveDocumentOptions { Mode = SaveMode.NewVersion });
        Assert.EndsWith(".v2.docx", sibling.Name);

        var duplicate = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.CreateAsync(name, new MemoryStream(original)));
        Assert.Equal(ProviderErrorCode.AlreadyExists, duplicate.Code);

        await provider.RemoveAsync(sibling);
        await Assert.ThrowsAsync<DocumentProviderException>(() => provider.OpenReadAsync(sibling));
    }
}

/// <summary>Runs only when Google Drive live-test credentials are configured.</summary>
public sealed class GoogleDriveLiveFactAttribute : FactAttribute
{
    /// <summary>Skips the test unless the key file and folder are set.</summary>
    public GoogleDriveLiveFactAttribute()
    {
        var keyFile = Environment.GetEnvironmentVariable("OFFICEAGENT_GOOGLEDRIVE_KEY_FILE");
        if (string.IsNullOrWhiteSpace(keyFile) || !File.Exists(keyFile) ||
            string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("OFFICEAGENT_GOOGLEDRIVE_FOLDER_ID")))
            Skip = "Set OFFICEAGENT_GOOGLEDRIVE_KEY_FILE and OFFICEAGENT_GOOGLEDRIVE_FOLDER_ID to run against Google Drive.";
    }
}

/// <summary>Runs only when Box live-test credentials are configured.</summary>
public sealed class BoxLiveFactAttribute : FactAttribute
{
    /// <summary>Skips the test unless the client credentials, subject and folder are set.</summary>
    public BoxLiveFactAttribute()
    {
        if (new[] { "OFFICEAGENT_BOX_CLIENT_ID", "OFFICEAGENT_BOX_CLIENT_SECRET", "OFFICEAGENT_BOX_SUBJECT_ID", "OFFICEAGENT_BOX_FOLDER_ID" }
            .Any(name => string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable(name))))
            Skip = "Set OFFICEAGENT_BOX_CLIENT_ID, OFFICEAGENT_BOX_CLIENT_SECRET, OFFICEAGENT_BOX_SUBJECT_ID and OFFICEAGENT_BOX_FOLDER_ID to run against Box.";
    }
}
