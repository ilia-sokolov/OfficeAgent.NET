using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Box;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// Box provider tests over a fake Box API: files simulated by an in-memory
/// <see cref="HttpMessageHandler"/>, so register → open → save → remove, ETag preconditions,
/// name conflicts and the storage outcomes run without a Box account.
/// </summary>
public class BoxProviderTests
{
    // DocxFactory builds a fresh package each call, and its bytes need not repeat.
    private static readonly byte[] Original = DocxFactory.Contract();

    // ── Register ──────────────────────────────────────────────────────────────

    [Fact]
    public async Task Register_by_file_id_assigns_an_opaque_id_and_opens_by_it()
    {
        using var box = new FakeBox();
        var fileId = box.Seed("contract.docx", Original);
        var provider = box.Provider();

        var added = await provider.RegisterAsync(fileId);

        Assert.Equal("box", added.Provider);
        Assert.NotEqual(fileId, added.ItemId);
        Assert.Equal("contract.docx", added.Name);
        Assert.Equal(box.ETag(fileId), added.Version);

        using var content = await provider.OpenReadAsync(added);
        Assert.Equal(Original, ((MemoryStream)content.Stream).ToArray());
    }

    [Theory]
    [InlineData("https://app.box.com/file/{0}")]
    [InlineData("https://contoso.app.box.com/file/{0}?s=abc")]
    [InlineData("https://app.box.com/file/{0}/")]
    public async Task Register_accepts_box_file_urls(string template)
    {
        using var box = new FakeBox();
        var fileId = box.Seed("contract.docx", DocxFactory.Contract());

        var added = await box.Provider().RegisterAsync(string.Format(template, fileId));

        Assert.Equal("contract.docx", added.Name);
    }

    [Theory]
    [InlineData("contract.docx")]
    [InlineData("12/34")]
    [InlineData("https://app.box.com/s/abcdef0123456789")]
    [InlineData("https://app.box.com/folder/12345")]
    [InlineData("https://example.com/file/12345")]
    [InlineData(" ")]
    public async Task Register_rejects_names_shared_links_folders_and_foreign_urls(string source)
    {
        using var box = new FakeBox();

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() => box.Provider().RegisterAsync(source));

        Assert.Equal(ProviderErrorCode.InvalidArgument, error.Code);
        Assert.Empty(box.Requests);
    }

    [Fact]
    public async Task Register_rejects_a_disallowed_extension_and_a_trashed_file()
    {
        using var box = new FakeBox();
        var exe = box.Seed("contract.exe", DocxFactory.Contract());
        var trashed = box.Seed("old.docx", DocxFactory.Contract());
        box.Trash(trashed);
        var provider = box.Provider();

        Assert.Equal(ProviderErrorCode.ExtensionNotAllowed,
            (await Assert.ThrowsAsync<DocumentProviderException>(() => provider.RegisterAsync(exe))).Code);
        Assert.Equal(ProviderErrorCode.NotFound,
            (await Assert.ThrowsAsync<DocumentProviderException>(() => provider.RegisterAsync(trashed))).Code);
    }

    [Fact]
    public async Task Register_surfaces_not_found_with_the_box_code_only()
    {
        using var box = new FakeBox();

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() => box.Provider().RegisterAsync("999999"));

        Assert.Equal(ProviderErrorCode.NotFound, error.Code);
        Assert.Contains("Box code: not_found", error.Message);
        Assert.DoesNotContain("999999", error.Message);
    }

    // ── Open ──────────────────────────────────────────────────────────────────

    [Fact]
    public async Task The_bearer_token_never_travels_to_the_download_host()
    {
        using var box = new FakeBox();
        var provider = box.Provider();
        var added = await provider.RegisterAsync(box.Seed("contract.docx", DocxFactory.Contract()));

        using var _ = await provider.OpenReadAsync(added);

        Assert.True(box.SawDownload);
        Assert.False(box.SawAuthorizedDownload);
    }

    [Fact]
    public async Task A_file_box_is_still_processing_is_reported_as_retryable_io()
    {
        using var box = new FakeBox();
        var provider = box.Provider();
        var added = await provider.RegisterAsync(box.Seed("contract.docx", DocxFactory.Contract()));
        box.AnswerContentWithAccepted = true;

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.OpenReadAsync(added));

        Assert.Equal(ProviderErrorCode.IO, error.Code);
        Assert.Contains("try again", error.Message);
    }

    [Fact]
    public async Task A_shield_policy_block_is_access_denied_and_says_so()
    {
        using var box = new FakeBox();
        var provider = box.Provider();
        var added = await provider.RegisterAsync(box.Seed("contract.docx", DocxFactory.Contract()));
        box.ContentStatus = (HttpStatusCode.Forbidden, "forbidden_by_policy");

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.OpenReadAsync(added));

        Assert.Equal(ProviderErrorCode.AccessDenied, error.Code);
        Assert.Contains("Box Shield policy", error.Message);
    }

    // ── Limits ────────────────────────────────────────────────────────────────

    [Fact]
    public async Task An_oversized_file_is_refused_from_metadata_before_any_download()
    {
        using var box = new FakeBox();
        var fileId = box.Seed("contract.docx", DocxFactory.Contract());
        var store = new InMemoryBoxRegistrationStore();
        var added = await box.Provider(registrationStore: store).RegisterAsync(fileId);

        var small = box.Provider(maximumBytes: 16, registrationStore: store);
        var register = await Assert.ThrowsAsync<DocumentProviderException>(() => small.RegisterAsync(fileId));
        var open = await Assert.ThrowsAsync<DocumentProviderException>(() => small.OpenReadAsync(added));

        Assert.Equal(ProviderErrorCode.ContentTooLarge, register.Code);
        Assert.Equal(ProviderErrorCode.ContentTooLarge, open.Code);
        Assert.False(box.SawDownload);
    }

    [Fact]
    public async Task Save_rejects_oversized_output_before_uploading()
    {
        using var box = new FakeBox();
        var provider = box.Provider(maximumBytes: 64);
        var source = await provider.RegisterAsync(box.Seed("contract.docx", new byte[8]));

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.SaveAsync(source, new MemoryStream(new byte[65]), new SaveDocumentOptions()));

        Assert.Equal(ProviderErrorCode.ContentTooLarge, error.Code);
        Assert.Equal(0, box.Uploads);
    }

    // ── Save modes ────────────────────────────────────────────────────────────

    [Fact]
    public async Task Replace_uploads_a_new_version_with_if_match_on_the_opened_etag()
    {
        using var box = new FakeBox();
        var fileId = box.Seed("contract.docx", DocxFactory.Contract());
        var provider = box.Provider();
        var source = await provider.RegisterAsync(fileId);
        var output = DocxFactory.Contract().Concat(new byte[] { 1 }).ToArray();

        var replaced = await provider.SaveAsync(source, new MemoryStream(output), new SaveDocumentOptions { Mode = SaveMode.Replace });

        Assert.Equal(source.ItemId, replaced.ItemId);
        Assert.NotEqual(source.Version, replaced.Version);
        Assert.Equal(output, box.Bytes(fileId));
        Assert.Equal(source.Version, box.LastIfMatch);
        Assert.Equal(2, box.VersionCount(fileId));
    }

    [Fact]
    public async Task A_stale_reference_is_a_version_conflict_and_writes_nothing()
    {
        using var box = new FakeBox();
        var fileId = box.Seed("contract.docx", DocxFactory.Contract());
        var provider = box.Provider();
        var source = await provider.RegisterAsync(fileId);
        box.EditElsewhere(fileId, new byte[] { 9 });

        await Assert.ThrowsAsync<DocumentVersionConflictException>(() =>
            provider.SaveAsync(source, new MemoryStream(DocxFactory.Contract()), new SaveDocumentOptions()));
        await Assert.ThrowsAsync<DocumentVersionConflictException>(() => provider.OpenReadAsync(source));

        Assert.Equal(new byte[] { 9 }, box.Bytes(fileId));
        Assert.Equal(0, box.Uploads);
    }

    [Fact]
    public async Task A_change_between_the_check_and_the_upload_is_refused_by_box()
    {
        using var box = new FakeBox();
        var fileId = box.Seed("contract.docx", DocxFactory.Contract());
        var provider = box.Provider();
        var source = await provider.RegisterAsync(fileId);
        // Another writer lands after the provider's metadata check, before its upload.
        box.BeforeNextUpload = () => box.EditElsewhere(fileId, new byte[] { 9 });

        var error = await Assert.ThrowsAsync<DocumentVersionConflictException>(() =>
            provider.SaveAsync(source, new MemoryStream(DocxFactory.Contract()), new SaveDocumentOptions()));

        Assert.Equal(ProviderErrorCode.VersionConflict, error.Code);
        Assert.Equal(new byte[] { 9 }, box.Bytes(fileId));
    }

    [Fact]
    public async Task New_version_writes_a_versioned_sibling_with_a_fresh_id()
    {
        using var box = new FakeBox();
        var fileId = box.Seed("contract.docx", Original);
        box.Seed("contract.v2.docx", DocxFactory.Contract());
        var provider = box.Provider();
        var source = await provider.RegisterAsync(fileId);

        var saved = await provider.SaveAsync(source, new MemoryStream(DocxFactory.Contract()), new SaveDocumentOptions { Mode = SaveMode.NewVersion });

        Assert.Equal("contract.v3.docx", saved.Name);
        Assert.NotEqual(source.ItemId, saved.ItemId);
        Assert.Equal(Original, box.Bytes(fileId));
        Assert.Contains(box.Requests, r => r.StartsWith("OPTIONS https://box.fake/2.0/files/content"));
        using var reopened = await provider.OpenReadAsync(saved);
        Assert.Equal("contract.v3.docx", reopened.Reference.Name);
    }

    [Fact]
    public async Task New_document_honours_a_validated_name_and_never_takes_an_existing_one()
    {
        using var box = new FakeBox();
        var fileId = box.Seed("contract.docx", DocxFactory.Contract());
        box.Seed("taken.docx", new byte[] { 7 });
        var provider = box.Provider();
        var source = await provider.RegisterAsync(fileId);

        var saved = await provider.SaveAsync(source, new MemoryStream(DocxFactory.Contract()),
            new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "globex.docx" });
        var taken = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(source,
            new MemoryStream(DocxFactory.Contract()), new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "Taken.docx" }));
        var extension = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(source,
            new MemoryStream(DocxFactory.Contract()), new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "run.exe" }));

        Assert.Equal("globex.docx", saved.Name);
        Assert.Equal(ProviderErrorCode.AlreadyExists, taken.Code);
        Assert.Equal(new byte[] { 7 }, box.BytesNamed("taken.docx"));
        Assert.Equal(ProviderErrorCode.ExtensionNotAllowed, extension.Code);
    }

    [Fact]
    public async Task Save_rejects_caller_chosen_destinations_and_replace_renames()
    {
        using var box = new FakeBox();
        var provider = box.Provider();
        var source = await provider.RegisterAsync(box.Seed("contract.docx", DocxFactory.Contract()));

        var destination = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(
            source, new MemoryStream(DocxFactory.Contract()), new SaveDocumentOptions { DestinationItemId = "1" }));
        var rename = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(
            source, new MemoryStream(DocxFactory.Contract()), new SaveDocumentOptions { Mode = SaveMode.Replace, NewName = "r.docx" }));

        Assert.Equal(ProviderErrorCode.InvalidArgument, destination.Code);
        Assert.Equal(ProviderErrorCode.InvalidArgument, rename.Code);
    }

    [Fact]
    public async Task Remove_drops_the_registration_but_not_the_box_file()
    {
        using var box = new FakeBox();
        var fileId = box.Seed("contract.docx", DocxFactory.Contract());
        var provider = box.Provider();
        var added = await provider.RegisterAsync(fileId);

        await provider.RemoveAsync(added);

        Assert.Equal(ProviderErrorCode.NotFound,
            (await Assert.ThrowsAsync<DocumentProviderException>(() => provider.OpenReadAsync(added))).Code);
        Assert.True(box.HasName("contract.docx"));
        Assert.DoesNotContain(box.Requests, r => r.StartsWith("DELETE"));
    }

    // ── Creation ──────────────────────────────────────────────────────────────

    [Fact]
    public async Task Create_uploads_registers_and_applies_an_initial_plan()
    {
        using var box = new FakeBox();
        var client = new OfficeAgentClient(
            new DocumentProviderRegistry(new IDocumentProvider[] { box.Provider() }), new WordModule());

        var created = await client.CreateAsync("box", "brief.docx", new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertOp
                {
                    Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = string.Empty },
                    Position = InsertPosition.Before,
                    Text = "Box brief"
                }
            }
        });

        Assert.True(created.Committed);
        Assert.Equal("box", created.Document!.Provider);
        Assert.Equal(FakeBox.FolderId, box.ParentOf("brief.docx"));
        var inspected = await client.InspectAsync(created.Document);
        Assert.Equal("Box brief", inspected.Paragraphs[0].Text);
    }

    [Fact]
    public async Task Create_never_overwrites_an_existing_name()
    {
        using var box = new FakeBox();
        box.Seed("brief.docx", new byte[] { 7 });

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            box.Provider().CreateAsync("brief.docx", new MemoryStream(DocxFactory.Contract())));

        Assert.Equal(ProviderErrorCode.AlreadyExists, error.Code);
        Assert.Equal(new byte[] { 7 }, box.BytesNamed("brief.docx"));
    }

    [Fact]
    public async Task Create_requires_a_configured_folder()
    {
        using var box = new FakeBox();

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            box.Provider(configureCreation: false).CreateAsync("brief.docx", new MemoryStream(DocxFactory.Contract())));

        Assert.Equal(ProviderErrorCode.ConfigurationError, error.Code);
        Assert.Empty(box.Requests);
    }

    [Fact]
    public async Task Create_failures_do_not_expose_the_configured_folder_id()
    {
        using var box = new FakeBox();
        box.RemoveFolder();
        var tools = new OfficeAgentTools(new OfficeAgentClient(
            new DocumentProviderRegistry(new IDocumentProvider[] { box.Provider() }), new WordModule()));

        var payload = await tools.CreateDocument("box", "brief.docx");

        Assert.DoesNotContain(FakeBox.FolderId, payload);
        using var json = JsonDocument.Parse(payload);
        Assert.Equal("not-found", json.RootElement.GetProperty("errors")[0].GetProperty("code").GetString());
    }

    [Fact]
    public async Task Registration_failure_after_upload_reports_the_name_and_hash()
    {
        using var box = new FakeBox();
        var provider = box.Provider(registrationStore: new FailingStore());

        var error = await Assert.ThrowsAsync<DocumentRegistrationFailedException>(() =>
            provider.CreateAsync("brief.docx", new MemoryStream(DocxFactory.Contract())));

        Assert.Equal("brief.docx", error.OutputName);
        Assert.Equal(Sha256(box.BytesNamed("brief.docx")), error.OutputSha256);
    }

    // ── Storage outcomes ──────────────────────────────────────────────────────

    [Fact]
    public async Task A_read_only_connection_registers_and_opens_but_sends_no_write()
    {
        using var box = new FakeBox();
        var original = DocxFactory.Contract();
        var fileId = box.Seed("Invoice.docx", original);
        var provider = box.Provider(configureCreation: false, readOnly: true);

        Assert.True(provider.ReadOnly);
        var reference = await provider.RegisterAsync(fileId);
        using (var content = await provider.OpenReadAsync(reference))
            Assert.True(content.Stream.Length > 0);
        var before = box.Requests.Count;

        var errors = new List<DocumentProviderException>();
        foreach (var mode in new[] { SaveMode.Replace, SaveMode.NewVersion, SaveMode.NewDocument })
        {
            using var bytes = new MemoryStream(DocxFactory.Contract());
            errors.Add(await Assert.ThrowsAsync<DocumentProviderException>(() =>
                provider.SaveAsync(reference, bytes, new SaveDocumentOptions { Mode = mode })));
        }
        using (var bytes = new MemoryStream(DocxFactory.Contract()))
            errors.Add(await Assert.ThrowsAsync<DocumentProviderException>(() => provider.CreateAsync("new.docx", bytes)));

        Assert.All(errors, error =>
        {
            Assert.Equal(ProviderErrorCode.AccessDenied, error.Code);
            Assert.Contains("read-only", error.Message, StringComparison.Ordinal);
        });
        // Nothing reached Box: not even a lookup, let alone an upload.
        Assert.Equal(before, box.Requests.Count);
        await provider.RemoveAsync(reference);
    }

    [Fact]
    public void A_read_only_connection_cannot_also_name_a_creation_folder()
    {
        using var box = new FakeBox();

        var error = Assert.Throws<ArgumentException>(() => box.Provider(configureCreation: true, readOnly: true));

        Assert.Contains("CreationFolderId", error.Message, StringComparison.Ordinal);
    }

    private static OfficeAgentClient EngineOver(BoxDocumentProvider provider) =>
        new(new DocumentProviderRegistry(new IDocumentProvider[] { provider }), new WordModule());

    [Theory]
    [InlineData(HttpStatusCode.BadRequest, "bad_request")]
    [InlineData((HttpStatusCode)429, "rate_limit_exceeded")]
    [InlineData(HttpStatusCode.Forbidden, "storage_limit_exceeded")]
    [InlineData(HttpStatusCode.Conflict, "operation_blocked_temporary")]
    public async Task A_refused_upload_is_certain_and_leaves_nothing_behind(HttpStatusCode status, string code)
    {
        using var box = new FakeBox { UploadStatus = (status, code) };

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            EngineOver(box.Provider()).CreateAsync("box", "refused.docx"));

        Assert.Equal(ProviderErrorCode.WriteRejected, error.Code);
        Assert.False(box.HasName("refused.docx"));
    }

    [Fact]
    public async Task A_locked_file_says_why_the_save_was_refused_and_what_to_do()
    {
        using var box = new FakeBox();
        var fileId = box.Seed("contract.docx", Original);
        var provider = box.Provider();
        var source = await provider.RegisterAsync(fileId);
        box.Lock(fileId);

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.SaveAsync(source, new MemoryStream(new byte[] { 1 }), new SaveDocumentOptions()));

        Assert.Equal(ProviderErrorCode.WriteRejected, error.Code);
        Assert.Contains("locked in Box", error.Message);
        Assert.Contains("save as a new document", error.Message);
        Assert.Contains("Box code: access_denied_item_locked", error.Message);
        Assert.Equal(Original, box.Bytes(fileId));
    }

    [Fact]
    public async Task A_server_error_on_upload_is_an_unknown_outcome()
    {
        using var box = new FakeBox { UploadStatus = (HttpStatusCode.ServiceUnavailable, "unavailable") };

        var error = await Assert.ThrowsAsync<DocumentWriteOutcomeUnknownException>(() =>
            EngineOver(box.Provider()).CreateAsync("box", "maybe.docx"));

        Assert.Equal(ProviderErrorCode.OutcomeUnknown, error.Code);
    }

    [Fact]
    public async Task A_lost_response_after_a_stored_upload_is_unknown_and_reconcilable()
    {
        using var box = new FakeBox { StoreThenLoseResponse = true };

        var error = await Assert.ThrowsAsync<DocumentWriteOutcomeUnknownException>(() =>
            EngineOver(box.Provider()).CreateAsync("box", "lost.docx"));

        Assert.Equal("lost.docx", error.OutputName);
        Assert.True(box.HasName("lost.docx"));
        Assert.Equal(Sha256(box.BytesNamed("lost.docx")), error.OutputSha256);
    }

    [Fact]
    public async Task The_bearer_token_never_appears_in_urls_or_errors()
    {
        using var box = new FakeBox { UploadStatus = (HttpStatusCode.BadRequest, "bad_request") };

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            box.Provider().CreateAsync("brief.docx", new MemoryStream(DocxFactory.Contract())));

        Assert.All(box.Requests, r => Assert.DoesNotContain(FakeBox.Token, r));
        Assert.DoesNotContain(FakeBox.Token, error.ToString());
    }

    // ── Registration store and tokens ─────────────────────────────────────────

    [Fact]
    public async Task Json_file_registration_store_survives_reload()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"officeagent-boxindex-{Guid.NewGuid():N}");
        var path = Path.Combine(directory, "index.json");
        try
        {
            var id = await new JsonFileBoxRegistrationStore(path).AddAsync(new BoxFileRef("12345"));

            var reloaded = new JsonFileBoxRegistrationStore(path);
            Assert.Equal(new BoxFileRef("12345"), await reloaded.ResolveAsync(id));
            Assert.True(await reloaded.RemoveAsync(id));
            Assert.Null(await new JsonFileBoxRegistrationStore(path).ResolveAsync(id));
        }
        finally
        {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Theory]
    [InlineData(BoxSubjectType.Enterprise, "enterprise")]
    [InlineData(BoxSubjectType.User, "user")]
    public async Task Client_credentials_tokens_name_the_subject_and_are_cached(BoxSubjectType type, string wire)
    {
        var handler = new GoogleDriveProviderTests.TokenEndpoint("{\"access_token\":\"box-token\",\"expires_in\":3600}");
        using var http = new HttpClient(handler);
        var tokens = new BoxClientCredentialsAccessTokenProvider(new BoxClientCredentialsOptions
        {
            ClientId = "client-id",
            ClientSecret = "client-secret",
            SubjectType = type,
            SubjectId = "4242"
        }, http);

        Assert.Equal("box-token", await tokens.GetAccessTokenAsync());
        Assert.Equal("box-token", await tokens.GetAccessTokenAsync());

        var request = Assert.Single(handler.Requests);
        Assert.Equal("https://api.box.com/oauth2/token", request.Url);
        Assert.Equal("client_credentials", request.Form["grant_type"]);
        Assert.Equal("client-id", request.Form["client_id"]);
        Assert.Equal("client-secret", request.Form["client_secret"]);
        Assert.Equal(wire, request.Form["box_subject_type"]);
        Assert.Equal("4242", request.Form["box_subject_id"]);
    }

    [Fact]
    public async Task A_refused_token_request_names_the_error_but_not_the_secret()
    {
        var handler = new GoogleDriveProviderTests.TokenEndpoint(
            "{\"error\":\"invalid_client\",\"error_description\":\"The client credentials are invalid\"}", HttpStatusCode.BadRequest);
        using var http = new HttpClient(handler);
        var tokens = new BoxClientCredentialsAccessTokenProvider(new BoxClientCredentialsOptions
        {
            ClientId = "client-id",
            ClientSecret = "s3cr3t-value",
            SubjectId = "4242"
        }, http);

        var error = await Assert.ThrowsAsync<InvalidOperationException>(() => tokens.GetAccessTokenAsync());

        Assert.Contains("invalid_client", error.Message);
        Assert.DoesNotContain("s3cr3t-value", error.Message);
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private sealed class FailingStore : IBoxRegistrationStore
    {
        public Task<string> AddAsync(BoxFileRef file, CancellationToken cancellationToken = default) =>
            throw new IOException("Registration store unavailable.");

        public Task<BoxFileRef?> ResolveAsync(string registrationId, CancellationToken cancellationToken = default) =>
            Task.FromResult<BoxFileRef?>(null);

        public Task<bool> RemoveAsync(string registrationId, CancellationToken cancellationToken = default) =>
            Task.FromResult(false);
    }

    /// <summary>
    /// Simulates the Box endpoints the provider uses: file metadata, content download through
    /// a redirect to a separate download host, the upload preflight, new-file upload with
    /// name-conflict refusal, and new-version upload with <c>If-Match</c>.
    /// </summary>
    internal sealed class FakeBox : IDisposable
    {
        public const string Token = "box-fake-token";
        public const string FolderId = "1000";
        public const string Api = "https://box.fake/2.0";
        public const string Upload = "https://upload.box.fake/api/2.0";
        public const string TokenUrl = "https://api.box.com/oauth2/token";
        private const string Download = "https://dl.boxcloud.fake";

        private readonly Dictionary<string, Entry> _files = new(StringComparer.Ordinal);
        private readonly HttpClient _http;
        private bool _folderExists = true;
        private int _nextId = 2000;

        public FakeBox() => _http = new HttpClient(new Handler(this));

        /// <summary>The client that reaches this fake, for a provider built elsewhere.</summary>
        public HttpClient Http => _http;

        /// <summary>How many tokens were requested from the token endpoint.</summary>
        public int TokenRequests { get; private set; }

        public List<string> Requests { get; } = new();
        public int Uploads { get; private set; }
        public bool SawDownload { get; private set; }
        public bool SawAuthorizedDownload { get; private set; }
        public string? LastIfMatch { get; private set; }
        public (HttpStatusCode Status, string Code)? UploadStatus { get; set; }
        public (HttpStatusCode Status, string Code)? ContentStatus { get; set; }
        public bool StoreThenLoseResponse { get; set; }
        public bool AnswerContentWithAccepted { get; set; }
        public Action? BeforeNextUpload { get; set; }

        /// <summary>Every folder listing answers with this status and code.</summary>
        public (HttpStatusCode Status, string Code)? ListStatus { get; set; }

        public string Seed(string name, byte[] bytes, string parent = FolderId, string type = "file")
        {
            var entry = new Entry { Id = (++_nextId).ToString(), Name = name, Bytes = bytes, Parent = parent, Type = type };
            _files[entry.Id] = entry;
            return entry.Id;
        }

        /// <summary>Adds a folder under <paramref name="parent"/> and returns its id.</summary>
        public string SeedFolder(string name, string parent = FolderId)
        {
            var id = (++_nextId).ToString();
            _folders[id] = (name, parent);
            return id;
        }

        private readonly Dictionary<string, (string Name, string Parent)> _folders = new(StringComparer.Ordinal);

        private bool FolderExists(string id) => id == FolderId ? _folderExists : _folders.ContainsKey(id);

        public string ETag(string id) => _files[id].ETag.ToString();
        public byte[] Bytes(string id) => _files[id].Bytes;
        public byte[] BytesNamed(string name) => _files.Values.Single(f => f.Name == name).Bytes;
        public string ParentOf(string name) => _files.Values.Single(f => f.Name == name).Parent;
        public int VersionCount(string id) => _files[id].ETag + 1;
        public bool HasName(string name) => _files.Values.Any(f => f.Name == name);
        public void Trash(string id) => _files[id].Status = "trashed";
        public void Lock(string id) => _files[id].Locked = true;
        public void RemoveFolder() => _folderExists = false;

        public void EditElsewhere(string id, byte[] bytes)
        {
            _files[id].Bytes = bytes;
            _files[id].ETag++;
        }

        public BoxDocumentProvider Provider(
            long maximumBytes = 50 * 1024 * 1024,
            bool configureCreation = true,
            IBoxRegistrationStore? registrationStore = null,
            string? templatesFolder = null,
            IReadOnlyCollection<string>? allowedExtensions = null,
            bool readOnly = false) => new(
            new BoxDocumentProviderOptions
            {
                ConnectionId = "box",
                ReadOnly = readOnly,
                ApiBaseUrl = Api,
                UploadBaseUrl = Upload,
                MaximumBytes = maximumBytes,
                CreationFolderId = configureCreation ? FolderId : string.Empty,
                TemplatesFolder = templatesFolder ?? string.Empty,
                AllowedExtensions = allowedExtensions ?? new[] { ".docx" }
            },
            _http,
            new FixedTokens(),
            registrationStore);

        public void Dispose() => _http.Dispose();

        private sealed class FixedTokens : IBoxAccessTokenProvider
        {
            public Task<string> GetAccessTokenAsync(CancellationToken cancellationToken = default) => Task.FromResult(Token);
        }

        private sealed class Entry
        {
            public string Id = "";
            public string Name = "";
            public byte[] Bytes = Array.Empty<byte>();
            public string Parent = "";
            public int ETag;
            public string Status = "active";
            public bool Locked;
            public string Type = "file";
        }

        private sealed class Handler : HttpMessageHandler
        {
            private readonly FakeBox _box;

            public Handler(FakeBox box) => _box = box;

            protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                var url = request.RequestUri!.ToString();
                _box.Requests.Add($"{request.Method} {url}");

                // The OAuth token endpoint, so a connection the MCP server builds from configuration
                // signs in through its real token provider. The form is recorded, never echoed.
                if (request.Method == HttpMethod.Post && url == FakeBox.TokenUrl)
                {
                    _box.TokenRequests++;
                    return new HttpResponseMessage(HttpStatusCode.OK)
                    {
                        Content = new StringContent(
                            JsonSerializer.Serialize(new { access_token = Token, expires_in = 3600 }), Encoding.UTF8, "application/json")
                    };
                }
                var path = request.RequestUri!.GetLeftPart(UriPartial.Path);

                if (url.StartsWith(Download, StringComparison.Ordinal))
                {
                    _box.SawDownload = true;
                    if (request.Headers.Authorization is not null) _box.SawAuthorizedDownload = true;
                    var id = path.Substring(Download.Length + 1);
                    return new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(_box._files[id].Bytes) };
                }

                if (request.Headers.Authorization is not { Scheme: "Bearer", Parameter: Token })
                    return Error(HttpStatusCode.Unauthorized, "unauthorized");

                if (path == Upload + "/files/content")
                    return await UploadNewAsync(request);
                if (path.StartsWith(Upload + "/files/", StringComparison.Ordinal) && path.EndsWith("/content", StringComparison.Ordinal))
                    return await UploadVersionAsync(request, path.Substring((Upload + "/files/").Length).Replace("/content", ""));

                if (path == Api + "/files/content" && request.Method == HttpMethod.Options)
                {
                    using var body = JsonDocument.Parse(await request.Content!.ReadAsStringAsync(cancellationToken));
                    var name = body.RootElement.GetProperty("name").GetString()!;
                    var parent = body.RootElement.GetProperty("parent").GetProperty("id").GetString()!;
                    return Taken(parent, name)
                        ? Error(HttpStatusCode.Conflict, "item_name_in_use")
                        : Json(new { upload_url = Upload + "/files/content" });
                }

                if (path.StartsWith(Api + "/files/", StringComparison.Ordinal) && request.Method == HttpMethod.Get)
                {
                    var rest = path.Substring((Api + "/files/").Length);
                    var content = rest.EndsWith("/content", StringComparison.Ordinal);
                    var id = content ? rest.Substring(0, rest.Length - "/content".Length) : rest;
                    if (!_box._files.TryGetValue(id, out var entry))
                        return Error(HttpStatusCode.NotFound, "not_found", $"Item {id} not found");
                    if (!content) return Json(File(entry));
                    if (_box.ContentStatus is { } blocked) return Error(blocked.Status, blocked.Code);
                    if (_box.AnswerContentWithAccepted)
                        return new HttpResponseMessage(HttpStatusCode.Accepted) { Headers = { RetryAfter = new System.Net.Http.Headers.RetryConditionHeaderValue(TimeSpan.FromSeconds(5)) } };
                    var redirect = new HttpResponseMessage(HttpStatusCode.Found);
                    redirect.Headers.Location = new Uri($"{Download}/{id}");
                    return redirect;
                }

                if (path.StartsWith(Api + "/folders/", StringComparison.Ordinal) && path.EndsWith("/items", StringComparison.Ordinal) &&
                    request.Method == HttpMethod.Get)
                    return Items(path.Substring((Api + "/folders/").Length).Replace("/items", ""), request.RequestUri!);

                return Error(HttpStatusCode.NotFound, "not_found");
            }

            /// <summary>
            /// <c>GET /folders/{id}/items</c> with marker paging: folders first, then files, by id.
            /// The marker is the index of the next item.
            /// </summary>
            private HttpResponseMessage Items(string folderId, Uri uri)
            {
                if (_box.ListStatus is { } refused) return Error(refused.Status, refused.Code);
                if (!_box.FolderExists(folderId)) return Error(HttpStatusCode.NotFound, "not_found", $"Folder {folderId} not found");
                var query = uri.Query.TrimStart('?').Split('&', StringSplitOptions.RemoveEmptyEntries)
                    .Select(pair => pair.Split('=', 2))
                    .ToDictionary(pair => pair[0], pair => pair.Length > 1 ? Uri.UnescapeDataString(pair[1]) : string.Empty);
                Assert.Equal("true", query["usemarker"]);
                var limit = int.Parse(query["limit"]);
                var start = query.TryGetValue("marker", out var marker) ? int.Parse(marker) : 0;

                var items = _box._folders
                    .Where(pair => pair.Value.Parent == folderId)
                    .OrderBy(pair => long.Parse(pair.Key))
                    .Select(pair => (object)new Dictionary<string, object> { ["type"] = "folder", ["id"] = pair.Key, ["name"] = pair.Value.Name })
                    .Concat(_box._files.Values
                        .Where(f => f.Parent == folderId && f.Status == "active")
                        .OrderBy(f => long.Parse(f.Id))
                        .Select(f => (object)new Dictionary<string, object>
                        {
                            ["type"] = f.Type, ["id"] = f.Id, ["name"] = f.Name, ["size"] = f.Bytes.LongLength,
                            ["modified_at"] = "2026-11-01T09:00:00-08:00", ["item_status"] = f.Status
                        }))
                    .ToList();
                var payload = new Dictionary<string, object?>
                {
                    ["entries"] = items.Skip(start).Take(limit).ToList(),
                    ["limit"] = limit,
                    ["next_marker"] = start + limit < items.Count ? (start + limit).ToString() : null
                };
                return Json(payload);
            }

            private async Task<(JsonDocument Attributes, byte[] Bytes)> ReadFormAsync(HttpRequestMessage request)
            {
                var parts = ((MultipartFormDataContent)request.Content!).ToList();
                Assert.Equal(2, parts.Count);
                // Box requires the attributes part before the file part.
                Assert.Equal("attributes", parts[0].Headers.ContentDisposition!.Name!.Trim('"'));
                Assert.Equal("file", parts[1].Headers.ContentDisposition!.Name!.Trim('"'));
                return (JsonDocument.Parse(await parts[0].ReadAsStringAsync()), await parts[1].ReadAsByteArrayAsync());
            }

            private async Task<HttpResponseMessage> UploadNewAsync(HttpRequestMessage request)
            {
                Assert.Equal(HttpMethod.Post, request.Method);
                if (_box.UploadStatus is { } simulated)
                {
                    _box.UploadStatus = null;
                    return Error(simulated.Status, simulated.Code);
                }
                var (attributes, bytes) = await ReadFormAsync(request);
                using (attributes)
                {
                    var name = attributes.RootElement.GetProperty("name").GetString()!;
                    var parent = attributes.RootElement.GetProperty("parent").GetProperty("id").GetString()!;
                    if (parent != FolderId || !_box._folderExists)
                        return Error(HttpStatusCode.NotFound, "not_found", $"Folder {parent} not found");
                    if (Taken(parent, name))
                        return Error(HttpStatusCode.Conflict, "item_name_in_use");

                    _box.Uploads++;
                    var entry = new Entry { Id = (++_box._nextId).ToString(), Name = name, Bytes = bytes, Parent = parent };
                    _box._files[entry.Id] = entry;
                    if (_box.StoreThenLoseResponse)
                    {
                        _box.StoreThenLoseResponse = false;
                        throw new TaskCanceledException("The request timed out after the upload was stored.");
                    }
                    return Json(new { total_count = 1, entries = new[] { File(entry) } }, HttpStatusCode.Created);
                }
            }

            private async Task<HttpResponseMessage> UploadVersionAsync(HttpRequestMessage request, string id)
            {
                Assert.Equal(HttpMethod.Post, request.Method);
                _box.BeforeNextUpload?.Invoke();
                _box.BeforeNextUpload = null;
                if (_box.UploadStatus is { } simulated)
                {
                    _box.UploadStatus = null;
                    return Error(simulated.Status, simulated.Code);
                }
                if (!_box._files.TryGetValue(id, out var entry))
                    return Error(HttpStatusCode.NotFound, "not_found");
                if (entry.Locked)
                    return Error(HttpStatusCode.Forbidden, "access_denied_item_locked");
                _box.LastIfMatch = request.Headers.TryGetValues("If-Match", out var values) ? values.Single() : null;
                if (_box.LastIfMatch is not null && _box.LastIfMatch != entry.ETag.ToString())
                    return Error(HttpStatusCode.PreconditionFailed, "precondition_failed");

                var (attributes, bytes) = await ReadFormAsync(request);
                attributes.Dispose();
                _box.Uploads++;
                entry.Bytes = bytes;
                entry.ETag++;
                return Json(new { total_count = 1, entries = new[] { File(entry) } });
            }

            private bool Taken(string parent, string name) => _box._files.Values.Any(f =>
                f.Parent == parent && string.Equals(f.Name, name, StringComparison.OrdinalIgnoreCase));

            private static Dictionary<string, object> File(Entry entry) => new()
            {
                ["type"] = entry.Type,
                ["id"] = entry.Id,
                ["etag"] = entry.ETag.ToString(),
                ["name"] = entry.Name,
                ["size"] = entry.Bytes.LongLength,
                ["item_status"] = entry.Status,
                ["parent"] = new Dictionary<string, string> { ["type"] = "folder", ["id"] = entry.Parent }
            };

            private static HttpResponseMessage Json(object payload, HttpStatusCode status = HttpStatusCode.OK) => new(status)
            {
                Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
            };

            private static HttpResponseMessage Error(HttpStatusCode status, string code, string? message = null) => new(status)
            {
                Content = new StringContent(JsonSerializer.Serialize(new
                {
                    type = "error",
                    status = (int)status,
                    code,
                    message = message ?? code,
                    request_id = "req-1"
                }), Encoding.UTF8, "application/json")
            };
        }
    }
}
