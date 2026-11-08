using OfficeAgent.Abstractions;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.IManage;

namespace OfficeAgent.Tests;

/// <summary>
/// Opt-in tests against a real iManage Work server. They are skipped unless the environment
/// names a server, an application, a service account, and a scratch document the account may
/// add versions to:
/// <c>OFFICEAGENT_IMANAGE_SERVER_URL</c>, <c>OFFICEAGENT_IMANAGE_CUSTOMER_ID</c>,
/// <c>OFFICEAGENT_IMANAGE_CLIENT_ID</c>, <c>OFFICEAGENT_IMANAGE_USERNAME</c>,
/// <c>OFFICEAGENT_IMANAGE_PASSWORD</c> and <c>OFFICEAGENT_IMANAGE_DOCUMENT_ID</c>
/// (a <c>.docx</c>), plus <c>OFFICEAGENT_IMANAGE_CLIENT_SECRET</c> for a confidential client.
/// <c>OFFICEAGENT_IMANAGE_FOLDER_ID</c> also runs the creation test, which files one new
/// document in that folder.
/// </summary>
public sealed class IManageIntegrationTests
{
    private const string Prefix = "OFFICEAGENT_IMANAGE_";

    [IManageIntegrationFact]
    public async Task Registers_opens_and_versions_a_real_document()
    {
        var provider = Provider();
        var source = await provider.RegisterAsync(Env("DOCUMENT_ID"));
        try
        {
            byte[] original;
            using (var opened = await provider.OpenReadAsync(source))
                original = ((MemoryStream)opened.Stream).ToArray();
            Assert.NotEmpty(original);

            // Saving the same bytes adds a version and changes nothing a reader would see.
            var saved = await provider.SaveAsync(source, new MemoryStream(original),
                new SaveDocumentOptions { Mode = SaveMode.Replace });
            Assert.Equal(source.ItemId, saved.ItemId);
            Assert.NotEqual(source.Version, saved.Version);

            using (var reopened = await provider.OpenReadAsync(saved))
                Assert.Equal(original, ((MemoryStream)reopened.Stream).ToArray());

            // The reference opened before that save is now stale.
            await Assert.ThrowsAsync<DocumentVersionConflictException>(() => provider.SaveAsync(
                source, new MemoryStream(original), new SaveDocumentOptions { Mode = SaveMode.Replace }));
        }
        finally
        {
            await provider.RemoveAsync(source);
        }
    }

    [IManageIntegrationFact("FOLDER_ID")]
    public async Task Creates_a_document_in_the_configured_folder()
    {
        var provider = Provider();
        var name = $"officeagent-integration-{Guid.NewGuid():N}.docx";

        var created = await provider.CreateAsync(name, new MemoryStream(DocxFactory.Contract()));

        Assert.Equal(name, created.Name);
        using var opened = await provider.OpenReadAsync(created);
        Assert.True(opened.Stream.Length > 0);
        await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.CreateAsync(name, new MemoryStream(DocxFactory.Contract())));
    }

    private static IManageDocumentProvider Provider()
    {
        var server = Env("SERVER_URL");
        var tokens = new IManageOAuthTokenProvider(new IManageOAuthOptions
        {
            ServerUrl = server,
            ClientId = Env("CLIENT_ID"),
            ClientSecret = Environment.GetEnvironmentVariable(Prefix + "CLIENT_SECRET") ?? string.Empty,
            Username = Env("USERNAME"),
            Password = Env("PASSWORD")
        }, IManageDocumentProvider.CreateHttpClient());
        return new IManageDocumentProvider(new IManageDocumentProviderOptions
        {
            ConnectionId = "imanage-integration",
            ServerUrl = server,
            CustomerId = Env("CUSTOMER_ID"),
            CreationFolderId = Environment.GetEnvironmentVariable(Prefix + "FOLDER_ID") ?? string.Empty
        }, IManageDocumentProvider.CreateHttpClient(), tokens);
    }

    private static string Env(string name) =>
        Environment.GetEnvironmentVariable(Prefix + name) ?? throw new InvalidOperationException($"{Prefix}{name} is not set.");
}

/// <summary>A fact that is skipped unless the iManage integration environment is configured.</summary>
public sealed class IManageIntegrationFactAttribute : FactAttribute
{
    private static readonly string[] Required =
        { "SERVER_URL", "CUSTOMER_ID", "CLIENT_ID", "USERNAME", "PASSWORD", "DOCUMENT_ID" };

    /// <summary>Skips the test when a required variable, or any of <paramref name="also"/>, is unset.</summary>
    public IManageIntegrationFactAttribute(params string[] also)
    {
        var missing = Required.Concat(also)
            .Where(name => string.IsNullOrWhiteSpace(Environment.GetEnvironmentVariable("OFFICEAGENT_IMANAGE_" + name)))
            .ToArray();
        if (missing.Length > 0)
            Skip = "Runs against a real iManage Work server; set " +
                   string.Join(", ", missing.Select(name => "OFFICEAGENT_IMANAGE_" + name)) + ".";
    }
}
