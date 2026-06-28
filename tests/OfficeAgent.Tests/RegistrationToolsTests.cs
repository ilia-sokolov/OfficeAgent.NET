using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// The v0.2 registration surface: agents may register documents with configured
/// connections and remove registrations through tools - but only when the host
/// opts in, and always inside the connection's boundary.
/// </summary>
public class RegistrationToolsTests
{
    [Fact]
    public void Registration_tools_are_absent_by_default_and_present_when_allowed()
    {
        using var workspace = new RegistrationWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);

        var defaults = tools.AsAIFunctions().Select(f => f.Name).ToArray();
        Assert.Equal(4, defaults.Length);
        Assert.DoesNotContain("register_document", defaults);
        Assert.DoesNotContain("remove_document", defaults);

        var opted = tools.AsAIFunctions(new OfficeAgentToolsOptions { AllowRegistration = true })
            .Select(f => f.Name).ToArray();
        Assert.Equal(6, opted.Length);
        Assert.Contains("register_document", opted);
        Assert.Contains("remove_document", opted);
    }

    [Fact]
    public async Task Agent_registers_edits_and_removes_a_document_through_tools()
    {
        using var workspace = new RegistrationWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        var relativePath = workspace.Stage(DocxFactory.Contract(), "contract.docx");

        // Register: the tool returns the opaque id, never echoes a path.
        var registered = JsonDocument.Parse(await tools.RegisterDocument("workspace", relativePath));
        var documentId = registered.RootElement.GetProperty("documentId").GetString()!;
        Assert.Equal("workspace", registered.RootElement.GetProperty("connectionId").GetString());
        Assert.Equal("contract.docx", registered.RootElement.GetProperty("name").GetString());
        Assert.DoesNotContain(relativePath, documentId);

        // The returned id drives the normal loop.
        using (var inspect = JsonDocument.Parse(await tools.InspectDocument("workspace", documentId)))
            Assert.Equal("Word", inspect.RootElement.GetProperty("format").GetString());

        // Remove: registration gone, file untouched.
        var removed = JsonDocument.Parse(await tools.RemoveDocument("workspace", documentId));
        Assert.True(removed.RootElement.GetProperty("removed").GetBoolean());
        Assert.True(File.Exists(Path.Combine(workspace.Root, relativePath)));

        var error = JsonDocument.Parse(await tools.InspectDocument("workspace", documentId));
        Assert.Equal("not-found", error.RootElement.GetProperty("errors")[0].GetProperty("Code").GetString());
    }

    [Fact]
    public async Task Register_tool_returns_structured_error_for_unknown_connection()
    {
        using var workspace = new RegistrationWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);

        var report = JsonDocument.Parse(await tools.RegisterDocument("nope", "contract.docx"));
        Assert.False(report.RootElement.GetProperty("isValid").GetBoolean());
        Assert.Equal("configuration-error", report.RootElement.GetProperty("errors")[0].GetProperty("Code").GetString());
    }

    [Fact]
    public async Task Register_tool_cannot_escape_the_connection_root()
    {
        using var workspace = new RegistrationWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);

        var report = JsonDocument.Parse(await tools.RegisterDocument("workspace", "../outside.docx"));
        Assert.False(report.RootElement.GetProperty("isValid").GetBoolean());
        Assert.Equal("access-denied", report.RootElement.GetProperty("errors")[0].GetProperty("Code").GetString());
    }

    [Fact]
    public void Connection_id_alone_resolves_the_provider_and_ambiguity_is_rejected()
    {
        var unique = new DocumentProviderRegistry(new IDocumentProvider[]
        {
            new FakeProvider("filesystem", "a"),
            new FakeProvider("sharepoint", "b")
        });
        Assert.Equal("sharepoint", unique.ResolveConnection("b").Provider);

        var duplicated = new DocumentProviderRegistry(new IDocumentProvider[]
        {
            new FakeProvider("filesystem", "a"),
            new FakeProvider("sharepoint", "a")
        });
        var ambiguous = Assert.Throws<DocumentProviderException>(() => duplicated.ResolveConnection("a"));
        Assert.Equal(ProviderErrorCode.ConfigurationError, ambiguous.Code);

        var missing = Assert.Throws<DocumentProviderException>(() => unique.ResolveConnection("zzz"));
        Assert.Equal(ProviderErrorCode.ConfigurationError, missing.Code);
    }

    private sealed class FakeProvider : IDocumentProvider
    {
        public FakeProvider(string provider, string connectionId)
        {
            Provider = provider;
            ConnectionId = connectionId;
        }

        public string Provider { get; }
        public string ConnectionId { get; }

        public Task<OfficeAgent.Abstractions.DocumentReference> RegisterAsync(string source, CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();

        public Task<DocumentContent> OpenReadAsync(OfficeAgent.Abstractions.DocumentReference reference, CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();

        public Task<OfficeAgent.Abstractions.DocumentReference> SaveAsync(OfficeAgent.Abstractions.DocumentReference source, Stream content, OfficeAgent.Abstractions.SaveDocumentOptions options, CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();

        public Task RemoveAsync(OfficeAgent.Abstractions.DocumentReference reference, CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();
    }

    private sealed class RegistrationWorkspace : IDisposable
    {
        private readonly ServiceProvider _serviceProvider;

        public RegistrationWorkspace()
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-regtools-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);

            var services = new ServiceCollection();
            services.AddWordFormat();
            services.AddFileSystemDocumentProvider("workspace", Root);
            services.AddOfficeAgent();
            _serviceProvider = services.BuildServiceProvider();
            Client = _serviceProvider.GetRequiredService<OfficeAgentClient>();
        }

        public string Root { get; }
        public OfficeAgentClient Client { get; }

        /// <summary>Stages a fixture under the root and returns its connection-relative path.</summary>
        public string Stage(byte[] bytes, string name)
        {
            File.WriteAllBytes(Path.Combine(Root, name), bytes);
            return name;
        }

        public void Dispose()
        {
            _serviceProvider.Dispose();
            if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true);
        }
    }
}
