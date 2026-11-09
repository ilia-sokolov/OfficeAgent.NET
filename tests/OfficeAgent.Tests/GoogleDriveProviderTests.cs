using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.GoogleDrive;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// Google Drive provider tests over a fake Drive API v3: files simulated by an in-memory
/// <see cref="HttpMessageHandler"/>, so register → open → save → remove, native-file refusal,
/// shared-drive parameters and the storage outcomes run without a Google account.
/// </summary>
public class GoogleDriveProviderTests
{
    // DocxFactory builds a fresh package each call, and its bytes need not repeat.
    private static readonly byte[] Original = DocxFactory.Contract();

    private const string DocxMime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    // ── Register ──────────────────────────────────────────────────────────────

    [Fact]
    public async Task Register_by_file_id_assigns_an_opaque_id_and_opens_by_it()
    {
        using var drive = new FakeDrive();
        var fileId = drive.Seed("contract.docx", Original);
        var provider = drive.Provider();

        var added = await provider.RegisterAsync(fileId);

        Assert.Equal("googledrive", added.Provider);
        Assert.NotEqual(fileId, added.ItemId);
        Assert.Equal("contract.docx", added.Name);
        Assert.Equal(DocxMime, added.ContentType);
        Assert.Equal(drive.Revision(fileId), added.Version);

        using var content = await provider.OpenReadAsync(added);
        Assert.Equal(added.ItemId, content.Reference.ItemId);
        Assert.Equal(Original.Length, content.Stream.Length);
    }

    [Theory]
    [InlineData("https://drive.google.com/file/d/{0}/view?usp=sharing")]
    [InlineData("https://drive.google.com/open?id={0}")]
    [InlineData("https://drive.google.com/uc?export=download&id={0}")]
    [InlineData("https://docs.google.com/document/d/{0}/edit?usp=sharing&ouid=1&rtpof=true&sd=true")]
    [InlineData("https://docs.google.com/document/u/1/d/{0}/edit")]
    public async Task Register_accepts_drive_and_docs_urls(string template)
    {
        using var drive = new FakeDrive();
        var fileId = drive.Seed("contract.docx", DocxFactory.Contract());

        var added = await drive.Provider().RegisterAsync(string.Format(template, fileId));

        Assert.Equal("contract.docx", added.Name);
    }

    [Theory]
    [InlineData("contract.docx")]
    [InlineData("folder/contract.docx")]
    [InlineData("https://example.com/file/d/abcdefghijk/view")]
    [InlineData("https://drive.google.com/drive/folders/abcdefghijk")]
    [InlineData("")]
    public async Task Register_rejects_names_paths_and_foreign_urls(string source)
    {
        using var drive = new FakeDrive();

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() => drive.Provider().RegisterAsync(source));

        Assert.Equal(ProviderErrorCode.InvalidArgument, error.Code);
        Assert.Empty(drive.Requests);
    }

    [Theory]
    [InlineData("application/vnd.google-apps.document", "Google Docs")]
    [InlineData("application/vnd.google-apps.spreadsheet", "Google Sheets")]
    [InlineData("application/vnd.google-apps.presentation", "Google Slides")]
    public async Task Native_google_files_are_refused_by_mime_type_and_never_downloaded(string mimeType, string kind)
    {
        using var drive = new FakeDrive();
        // A native file can even carry an Office-looking name; the MIME type decides.
        var fileId = drive.Seed("Quarterly report.docx", Array.Empty<byte>(), mimeType: mimeType, size: null);

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() => drive.Provider().RegisterAsync(fileId));

        Assert.Equal(ProviderErrorCode.ExtensionNotAllowed, error.Code);
        Assert.Contains($"native {kind} file, not an Office file", error.Message);
        Assert.Contains("does not convert", error.Message);
        Assert.DoesNotContain(drive.Requests, r => r.Contains("alt=media"));
    }

    [Fact]
    public async Task Folders_shortcuts_and_trashed_files_are_refused()
    {
        using var drive = new FakeDrive();
        var folder = drive.Seed("Contracts", Array.Empty<byte>(), mimeType: "application/vnd.google-apps.folder", size: null);
        var shortcut = drive.Seed("contract.docx", Array.Empty<byte>(), mimeType: "application/vnd.google-apps.shortcut", size: null);
        var trashed = drive.Seed("old.docx", DocxFactory.Contract());
        drive.Trash(trashed);
        var provider = drive.Provider();

        Assert.Equal(ProviderErrorCode.InvalidArgument,
            (await Assert.ThrowsAsync<DocumentProviderException>(() => provider.RegisterAsync(folder))).Code);
        var shortcutError = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.RegisterAsync(shortcut));
        Assert.Contains("shortcut", shortcutError.Message);
        Assert.Equal(ProviderErrorCode.NotFound,
            (await Assert.ThrowsAsync<DocumentProviderException>(() => provider.RegisterAsync(trashed))).Code);
    }

    [Fact]
    public async Task Register_rejects_a_disallowed_extension()
    {
        using var drive = new FakeDrive();
        var fileId = drive.Seed("budget.xlsx", DocxFactory.Contract(), mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() => drive.Provider().RegisterAsync(fileId));

        Assert.Equal(ProviderErrorCode.ExtensionNotAllowed, error.Code);
    }

    [Fact]
    public async Task Register_surfaces_not_found_without_the_upstream_message()
    {
        using var drive = new FakeDrive();

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            drive.Provider().RegisterAsync("missingFileId123"));

        Assert.Equal(ProviderErrorCode.NotFound, error.Code);
        Assert.Contains("Drive reason: notFound", error.Message);
        Assert.DoesNotContain("missingFileId123", error.Message);
    }

    // ── Limits ────────────────────────────────────────────────────────────────

    [Fact]
    public async Task An_oversized_file_is_refused_from_metadata_before_any_download()
    {
        using var drive = new FakeDrive();
        var fileId = drive.Seed("contract.docx", DocxFactory.Contract());
        var store = new InMemoryGoogleDriveRegistrationStore();
        var added = await drive.Provider(registrationStore: store).RegisterAsync(fileId);

        var small = drive.Provider(maximumBytes: 16, registrationStore: store);
        var register = await Assert.ThrowsAsync<DocumentProviderException>(() => small.RegisterAsync(fileId));
        var open = await Assert.ThrowsAsync<DocumentProviderException>(() => small.OpenReadAsync(added));

        Assert.Equal(ProviderErrorCode.ContentTooLarge, register.Code);
        Assert.Equal(ProviderErrorCode.ContentTooLarge, open.Code);
        Assert.DoesNotContain(drive.Requests, r => r.Contains("alt=media"));
    }

    [Fact]
    public async Task A_download_larger_than_its_metadata_is_cut_off_at_the_limit()
    {
        using var drive = new FakeDrive();
        var fileId = drive.Seed("contract.docx", DocxFactory.Contract());
        var store = new InMemoryGoogleDriveRegistrationStore();
        var added = await drive.Provider(registrationStore: store).RegisterAsync(fileId);
        drive.UnderReportSize(fileId);

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            drive.Provider(maximumBytes: 64, registrationStore: store).OpenReadAsync(added));

        Assert.Equal(ProviderErrorCode.ContentTooLarge, error.Code);
    }

    [Fact]
    public async Task Save_rejects_oversized_output_before_uploading()
    {
        using var drive = new FakeDrive();
        var fileId = drive.Seed("contract.docx", new byte[8]);
        var provider = drive.Provider(maximumBytes: 64);
        var source = await provider.RegisterAsync(fileId);

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.SaveAsync(source, new MemoryStream(new byte[65]), new SaveDocumentOptions()));

        Assert.Equal(ProviderErrorCode.ContentTooLarge, error.Code);
        Assert.Equal(0, drive.Uploads);
    }

    // ── Save modes ────────────────────────────────────────────────────────────

    [Fact]
    public async Task Replace_uploads_a_new_revision_in_place()
    {
        using var drive = new FakeDrive();
        var fileId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();
        var source = await provider.RegisterAsync(fileId);
        var output = DocxFactory.Contract().Concat(new byte[] { 1 }).ToArray();

        var replaced = await provider.SaveAsync(source, new MemoryStream(output), new SaveDocumentOptions { Mode = SaveMode.Replace });

        Assert.Equal(source.ItemId, replaced.ItemId);
        Assert.NotEqual(source.Version, replaced.Version);
        Assert.Equal(drive.Revision(fileId), replaced.Version);
        Assert.Equal(output, drive.Bytes(fileId));
        Assert.Contains(drive.Requests, r => r.StartsWith("PATCH https://drive.fake/upload/drive/v3/files/" + fileId + "?uploadType=media"));
    }

    [Fact]
    public async Task A_stale_reference_is_a_version_conflict_and_writes_nothing()
    {
        using var drive = new FakeDrive();
        var fileId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();
        var source = await provider.RegisterAsync(fileId);
        drive.EditElsewhere(fileId, new byte[] { 9, 9, 9 });

        var error = await Assert.ThrowsAsync<DocumentVersionConflictException>(() =>
            provider.SaveAsync(source, new MemoryStream(DocxFactory.Contract()), new SaveDocumentOptions()));

        Assert.Equal(ProviderErrorCode.VersionConflict, error.Code);
        Assert.Equal(new byte[] { 9, 9, 9 }, drive.Bytes(fileId));
        Assert.Equal(0, drive.Uploads);
        await Assert.ThrowsAsync<DocumentVersionConflictException>(() => provider.OpenReadAsync(source));
    }

    [Fact]
    public async Task The_version_ignores_renames_and_follows_content()
    {
        using var drive = new FakeDrive();
        var fileId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();
        var source = await provider.RegisterAsync(fileId);
        drive.RenameElsewhere(fileId, "contract (final).docx");

        // A rename bumps Drive's change counter but not the head revision, so the save proceeds.
        var replaced = await provider.SaveAsync(source, new MemoryStream(DocxFactory.Contract()), new SaveDocumentOptions());

        Assert.Equal("contract (final).docx", replaced.Name);
    }

    [Fact]
    public async Task New_version_writes_a_versioned_sibling_with_a_fresh_id()
    {
        using var drive = new FakeDrive();
        var fileId = drive.Seed("contract.docx", Original);
        drive.Seed("contract.v2.docx", DocxFactory.Contract());
        var provider = drive.Provider();
        var source = await provider.RegisterAsync(fileId);

        var saved = await provider.SaveAsync(source, new MemoryStream(DocxFactory.Contract()), new SaveDocumentOptions { Mode = SaveMode.NewVersion });

        Assert.Equal("contract.v3.docx", saved.Name);
        Assert.NotEqual(source.ItemId, saved.ItemId);
        Assert.Equal(Original, drive.Bytes(fileId));
        Assert.Equal(new[] { FakeDrive.FolderId }, drive.ParentsOf("contract.v3.docx"));
        Assert.Equal(DocxMime, drive.MimeTypeOf("contract.v3.docx"));
        using var reopened = await provider.OpenReadAsync(saved);
        Assert.Equal("contract.v3.docx", reopened.Reference.Name);
    }

    [Fact]
    public async Task New_document_honours_a_validated_name_and_never_takes_an_existing_one()
    {
        using var drive = new FakeDrive();
        var fileId = drive.Seed("contract.docx", DocxFactory.Contract());
        drive.Seed("taken.docx", new byte[] { 7 });
        var provider = drive.Provider();
        var source = await provider.RegisterAsync(fileId);

        var saved = await provider.SaveAsync(source, new MemoryStream(DocxFactory.Contract()),
            new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "globex.docx" });
        var taken = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(source,
            new MemoryStream(DocxFactory.Contract()), new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "taken.docx" }));
        var nested = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(source,
            new MemoryStream(DocxFactory.Contract()), new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "a/b.docx" }));
        var extension = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(source,
            new MemoryStream(DocxFactory.Contract()), new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "run.exe" }));

        Assert.Equal("globex.docx", saved.Name);
        Assert.Equal(ProviderErrorCode.AlreadyExists, taken.Code);
        Assert.Equal(new byte[] { 7 }, drive.BytesNamed("taken.docx"));
        Assert.Equal(ProviderErrorCode.InvalidArgument, nested.Code);
        Assert.Equal(ProviderErrorCode.ExtensionNotAllowed, extension.Code);
    }

    [Fact]
    public async Task Save_rejects_caller_chosen_destinations_and_replace_renames()
    {
        using var drive = new FakeDrive();
        var provider = drive.Provider();
        var source = await provider.RegisterAsync(drive.Seed("contract.docx", DocxFactory.Contract()));

        var destination = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(
            source, new MemoryStream(DocxFactory.Contract()), new SaveDocumentOptions { DestinationItemId = "x" }));
        var rename = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(
            source, new MemoryStream(DocxFactory.Contract()), new SaveDocumentOptions { Mode = SaveMode.Replace, NewName = "r.docx" }));

        Assert.Equal(ProviderErrorCode.InvalidArgument, destination.Code);
        Assert.Equal(ProviderErrorCode.InvalidArgument, rename.Code);
    }

    [Fact]
    public async Task Remove_drops_the_registration_but_not_the_drive_file()
    {
        using var drive = new FakeDrive();
        var fileId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();
        var added = await provider.RegisterAsync(fileId);

        await provider.RemoveAsync(added);

        Assert.Equal(ProviderErrorCode.NotFound,
            (await Assert.ThrowsAsync<DocumentProviderException>(() => provider.OpenReadAsync(added))).Code);
        Assert.Equal(ProviderErrorCode.NotFound,
            (await Assert.ThrowsAsync<DocumentProviderException>(() => provider.RemoveAsync(added))).Code);
        Assert.True(drive.HasName("contract.docx"));
        Assert.DoesNotContain(drive.Requests, r => r.StartsWith("DELETE"));
    }

    // ── Creation ──────────────────────────────────────────────────────────────

    [Fact]
    public async Task Create_uploads_registers_and_applies_an_initial_plan()
    {
        using var drive = new FakeDrive();
        var client = new OfficeAgentClient(
            new DocumentProviderRegistry(new IDocumentProvider[] { drive.Provider() }), new WordModule());

        var created = await client.CreateAsync("drive", "brief.docx", new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertOp
                {
                    Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = string.Empty },
                    Position = InsertPosition.Before,
                    Text = "Drive brief"
                }
            }
        });

        Assert.True(created.Committed);
        Assert.Equal("googledrive", created.Document!.Provider);
        Assert.Equal(new[] { FakeDrive.FolderId }, drive.ParentsOf("brief.docx"));
        Assert.Equal(DocxMime, drive.MimeTypeOf("brief.docx"));
        var inspected = await client.InspectAsync(created.Document);
        Assert.Equal("Drive brief", inspected.Paragraphs[0].Text);
    }

    [Fact]
    public async Task Create_never_overwrites_an_existing_name()
    {
        using var drive = new FakeDrive();
        drive.Seed("brief.docx", new byte[] { 7 });

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            drive.Provider().CreateAsync("brief.docx", new MemoryStream(DocxFactory.Contract())));

        Assert.Equal(ProviderErrorCode.AlreadyExists, error.Code);
        Assert.Equal(new byte[] { 7 }, drive.BytesNamed("brief.docx"));
        Assert.Equal(0, drive.Uploads);
    }

    [Fact]
    public async Task Create_requires_a_configured_folder()
    {
        using var drive = new FakeDrive();

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            drive.Provider(configureCreation: false).CreateAsync("brief.docx", new MemoryStream(DocxFactory.Contract())));

        Assert.Equal(ProviderErrorCode.ConfigurationError, error.Code);
        Assert.Empty(drive.Requests);
    }

    [Theory]
    [InlineData("CON.docx")]
    [InlineData("a:b.docx")]
    [InlineData(" brief.docx")]
    public async Task Create_rejects_unportable_names(string name)
    {
        using var drive = new FakeDrive();

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            drive.Provider().CreateAsync(name, new MemoryStream(DocxFactory.Contract())));

        Assert.Equal(ProviderErrorCode.InvalidArgument, error.Code);
    }

    [Fact]
    public async Task Create_failures_do_not_expose_the_configured_folder_id()
    {
        using var drive = new FakeDrive();
        drive.RemoveFolder();
        var tools = new OfficeAgentTools(new OfficeAgentClient(
            new DocumentProviderRegistry(new IDocumentProvider[] { drive.Provider() }), new WordModule()));

        var payload = await tools.CreateDocument("drive", "brief.docx");

        Assert.DoesNotContain(FakeDrive.FolderId, payload);
        using var json = JsonDocument.Parse(payload);
        Assert.Equal("not-found", json.RootElement.GetProperty("errors")[0].GetProperty("code").GetString());
    }

    [Fact]
    public async Task Registration_failure_after_upload_reports_the_name_and_hash()
    {
        using var drive = new FakeDrive();
        var provider = drive.Provider(registrationStore: new FailingStore());

        var error = await Assert.ThrowsAsync<DocumentRegistrationFailedException>(() =>
            provider.CreateAsync("brief.docx", new MemoryStream(DocxFactory.Contract())));

        Assert.Equal("brief.docx", error.OutputName);
        Assert.Equal(Sha256(drive.BytesNamed("brief.docx")), error.OutputSha256);
        Assert.DoesNotContain(FakeDrive.FolderId, error.Message);
    }

    // ── Shared drives ─────────────────────────────────────────────────────────

    [Fact]
    public async Task Every_request_supports_shared_drives_and_name_checks_search_the_files_drive()
    {
        using var drive = new FakeDrive { SharedDriveId = "0AsharedDriveId" };
        var fileId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();
        var source = await provider.RegisterAsync(fileId);

        await provider.SaveAsync(source, new MemoryStream(DocxFactory.Contract()), new SaveDocumentOptions { Mode = SaveMode.NewVersion });

        Assert.All(drive.Requests, r => Assert.Contains("supportsAllDrives=true", r));
        Assert.Contains(drive.Requests, r =>
            r.Contains("/files?q=") && r.Contains("corpora=drive") && r.Contains("driveId=0AsharedDriveId") &&
            r.Contains("includeItemsFromAllDrives=true"));
    }

    // ── Storage outcomes ──────────────────────────────────────────────────────

    private static OfficeAgentClient EngineOver(GoogleDriveDocumentProvider provider) =>
        new(new DocumentProviderRegistry(new IDocumentProvider[] { provider }), new WordModule());

    [Theory]
    [InlineData(HttpStatusCode.BadRequest, "badRequest")]
    [InlineData(HttpStatusCode.Forbidden, "userRateLimitExceeded")]
    [InlineData(HttpStatusCode.Forbidden, "storageQuotaExceeded")]
    [InlineData((HttpStatusCode)429, "rateLimitExceeded")]
    public async Task A_refused_upload_is_certain_and_leaves_nothing_behind(HttpStatusCode status, string reason)
    {
        using var drive = new FakeDrive { UploadStatus = status, UploadReason = reason };

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            EngineOver(drive.Provider()).CreateAsync("drive", "refused.docx"));

        Assert.Equal(ProviderErrorCode.WriteRejected, error.Code);
        Assert.False(drive.HasName("refused.docx"));
    }

    [Fact]
    public async Task A_refused_replace_is_write_rejected()
    {
        using var drive = new FakeDrive();
        var fileId = drive.Seed("contract.docx", Original);
        var provider = drive.Provider();
        var source = await provider.RegisterAsync(fileId);
        drive.UploadStatus = HttpStatusCode.BadRequest;

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.SaveAsync(source, new MemoryStream(new byte[] { 1 }), new SaveDocumentOptions()));

        Assert.Equal(ProviderErrorCode.WriteRejected, error.Code);
        Assert.Equal(Original, drive.Bytes(fileId));
    }

    [Fact]
    public async Task A_locked_file_is_refused_before_upload_with_what_to_do()
    {
        using var drive = new FakeDrive();
        var fileId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();
        var source = await provider.RegisterAsync(fileId);
        drive.Lock(fileId);

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.SaveAsync(source, new MemoryStream(new byte[] { 1 }), new SaveDocumentOptions()));

        Assert.Equal(ProviderErrorCode.WriteRejected, error.Code);
        Assert.Contains("locked in Google Drive", error.Message);
        Assert.Contains("save as a new document", error.Message);
        Assert.Equal(0, drive.Uploads);
    }

    [Fact]
    public async Task A_read_only_identity_is_refused_before_upload()
    {
        using var drive = new FakeDrive();
        var fileId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();
        var source = await provider.RegisterAsync(fileId);
        drive.MakeReadOnlyForCaller(fileId);

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.SaveAsync(source, new MemoryStream(new byte[] { 1 }), new SaveDocumentOptions()));

        Assert.Equal(ProviderErrorCode.AccessDenied, error.Code);
        Assert.Equal(0, drive.Uploads);
    }

    [Fact]
    public async Task A_server_error_on_upload_is_an_unknown_outcome()
    {
        using var drive = new FakeDrive { UploadStatus = HttpStatusCode.ServiceUnavailable, UploadReason = "backendError" };

        var error = await Assert.ThrowsAsync<DocumentWriteOutcomeUnknownException>(() =>
            EngineOver(drive.Provider()).CreateAsync("drive", "maybe.docx"));

        Assert.Equal(ProviderErrorCode.OutcomeUnknown, error.Code);
    }

    [Fact]
    public async Task A_lost_response_after_a_stored_upload_is_unknown_and_reconcilable()
    {
        using var drive = new FakeDrive { StoreThenLoseResponse = true };

        var error = await Assert.ThrowsAsync<DocumentWriteOutcomeUnknownException>(() =>
            EngineOver(drive.Provider()).CreateAsync("drive", "lost.docx"));

        Assert.Equal("lost.docx", error.OutputName);
        Assert.True(drive.HasName("lost.docx"));
        Assert.Equal(Sha256(drive.BytesNamed("lost.docx")), error.OutputSha256);
    }

    [Fact]
    public async Task A_lost_replace_response_is_unknown_not_failed()
    {
        using var drive = new FakeDrive();
        var fileId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();
        var source = await provider.RegisterAsync(fileId);
        drive.StoreThenLoseResponse = true;
        var output = new byte[] { 4, 2 };

        // Directly, the provider lets the transport failure through; the engine classifies it.
        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            provider.SaveAsync(source, new MemoryStream(output), new SaveDocumentOptions()));
        Assert.Equal(output, drive.Bytes(fileId));
    }

    [Fact]
    public async Task The_bearer_token_never_appears_in_urls_or_errors()
    {
        using var drive = new FakeDrive { UploadStatus = HttpStatusCode.BadRequest, UploadReason = "badRequest" };
        var provider = drive.Provider();

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.CreateAsync("brief.docx", new MemoryStream(DocxFactory.Contract())));

        Assert.All(drive.Requests, r => Assert.DoesNotContain(FakeDrive.Token, r));
        Assert.DoesNotContain(FakeDrive.Token, error.ToString());
    }

    // ── Registration store ────────────────────────────────────────────────────

    [Fact]
    public async Task Json_file_registration_store_survives_reload()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"officeagent-gdindex-{Guid.NewGuid():N}");
        var path = Path.Combine(directory, "index.json");
        try
        {
            var id = await new JsonFileGoogleDriveRegistrationStore(path).AddAsync(new GoogleDriveFileRef("file-1"));

            var reloaded = new JsonFileGoogleDriveRegistrationStore(path);
            Assert.Equal(new GoogleDriveFileRef("file-1"), await reloaded.ResolveAsync(id));
            Assert.True(await reloaded.RemoveAsync(id));
            Assert.Null(await new JsonFileGoogleDriveRegistrationStore(path).ResolveAsync(id));
        }
        finally
        {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    // ── Service-account tokens ────────────────────────────────────────────────

    [Fact]
    public async Task Service_account_tokens_use_a_signed_jwt_bearer_assertion_and_are_cached()
    {
        using var rsa = RSA.Create(2048);
        var keyJson = JsonSerializer.Serialize(new Dictionary<string, string>
        {
            ["type"] = "service_account",
            ["client_email"] = "officeagent@project.iam.gserviceaccount.com",
            ["private_key_id"] = "key-1",
            ["private_key"] = rsa.ExportPkcs8PrivateKeyPem(),
            // Never used: the assertion goes only to the configured token endpoint.
            ["token_uri"] = "https://attacker.example/token"
        });
        var handler = new TokenEndpoint("{\"access_token\":\"ya29.fake\",\"expires_in\":3599,\"token_type\":\"Bearer\"}");
        using var http = new HttpClient(handler);
        var tokens = new GoogleServiceAccountAccessTokenProvider(new GoogleServiceAccountOptions
        {
            KeyJson = keyJson,
            Subject = "editor@contoso.example"
        }, http);

        Assert.Equal("ya29.fake", await tokens.GetAccessTokenAsync());
        Assert.Equal("ya29.fake", await tokens.GetAccessTokenAsync());

        var request = Assert.Single(handler.Requests);
        Assert.Equal("https://oauth2.googleapis.com/token", request.Url);
        Assert.Equal("urn:ietf:params:oauth:grant-type:jwt-bearer", request.Form["grant_type"]);
        var parts = request.Form["assertion"].Split('.');
        Assert.Equal(3, parts.Length);
        using var header = JsonDocument.Parse(FromBase64Url(parts[0]));
        Assert.Equal("RS256", header.RootElement.GetProperty("alg").GetString());
        Assert.Equal("key-1", header.RootElement.GetProperty("kid").GetString());
        using var claims = JsonDocument.Parse(FromBase64Url(parts[1]));
        Assert.Equal("officeagent@project.iam.gserviceaccount.com", claims.RootElement.GetProperty("iss").GetString());
        Assert.Equal(GoogleDriveScopes.Drive, claims.RootElement.GetProperty("scope").GetString());
        Assert.Equal("https://oauth2.googleapis.com/token", claims.RootElement.GetProperty("aud").GetString());
        Assert.Equal("editor@contoso.example", claims.RootElement.GetProperty("sub").GetString());
        Assert.Equal(3600, claims.RootElement.GetProperty("exp").GetInt64() - claims.RootElement.GetProperty("iat").GetInt64());
        Assert.True(rsa.VerifyData(
            Encoding.ASCII.GetBytes(parts[0] + "." + parts[1]), FromBase64Url(parts[2]),
            HashAlgorithmName.SHA256, RSASignaturePadding.Pkcs1));
    }

    [Fact]
    public async Task A_refused_token_request_names_the_error_but_not_the_key_or_assertion()
    {
        using var rsa = RSA.Create(2048);
        var pem = rsa.ExportPkcs8PrivateKeyPem();
        var handler = new TokenEndpoint("{\"error\":\"invalid_grant\",\"error_description\":\"Invalid JWT Signature.\"}", HttpStatusCode.BadRequest);
        using var http = new HttpClient(handler);
        var tokens = new GoogleServiceAccountAccessTokenProvider(new GoogleServiceAccountOptions
        {
            ClientEmail = "officeagent@project.iam.gserviceaccount.com",
            PrivateKey = pem
        }, http);

        var error = await Assert.ThrowsAsync<InvalidOperationException>(() => tokens.GetAccessTokenAsync());

        Assert.Contains("invalid_grant", error.Message);
        Assert.DoesNotContain(handler.Requests[0].Form["assertion"], error.Message);
        Assert.DoesNotContain("PRIVATE KEY", error.Message);
        Assert.False(handler.Requests[0].Form.ContainsKey("sub"));
    }

    [Fact]
    public void An_unreadable_key_fails_at_startup_without_echoing_it()
    {
        using var http = new HttpClient();
        const string secret = "-----BEGIN PRIVATE KEY-----\nbm90IGEga2V5\n-----END PRIVATE KEY-----\n";

        var error = Assert.Throws<ArgumentException>(() => new GoogleServiceAccountAccessTokenProvider(
            new GoogleServiceAccountOptions { ClientEmail = "a@b.example", PrivateKey = secret }, http));
        var badJson = Assert.Throws<ArgumentException>(() => new GoogleServiceAccountAccessTokenProvider(
            new GoogleServiceAccountOptions { KeyJson = "{\"private_key\": \"" + "abc" }, http));

        Assert.DoesNotContain("bm90IGEga2V5", error.Message);
        Assert.DoesNotContain("abc", badJson.Message);
    }

    [Theory]
    [InlineData(1024)]
    [InlineData(2048)]
    [InlineData(3072)]
    public void The_netstandard_key_reader_matches_the_platform_importer(int bits)
    {
        // Several keys per size, so components with leading zero bytes are exercised too.
        for (var i = 0; i < 8; i++)
        {
            using var rsa = RSA.Create(bits);
            var expected = rsa.ExportParameters(includePrivateParameters: true);

            foreach (var pem in new[] { rsa.ExportPkcs8PrivateKeyPem(), rsa.ExportRSAPrivateKeyPem() })
            {
                var actual = RsaPrivateKey.ReadParameters(pem);
                Assert.Equal(expected.Modulus, actual.Modulus);
                Assert.Equal(expected.Exponent, actual.Exponent);
                Assert.Equal(expected.D, actual.D);
                Assert.Equal(expected.P, actual.P);
                Assert.Equal(expected.Q, actual.Q);
                Assert.Equal(expected.DP, actual.DP);
                Assert.Equal(expected.DQ, actual.DQ);
                Assert.Equal(expected.InverseQ, actual.InverseQ);
            }
        }
    }

    // ── Helpers ───────────────────────────────────────────────────────────────

    private static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();

    private static byte[] FromBase64Url(string value)
    {
        var base64 = value.Replace('-', '+').Replace('_', '/');
        return Convert.FromBase64String(base64.PadRight(base64.Length + (4 - base64.Length % 4) % 4, '='));
    }

    private sealed class FailingStore : IGoogleDriveRegistrationStore
    {
        public Task<string> AddAsync(GoogleDriveFileRef file, CancellationToken cancellationToken = default) =>
            throw new IOException("Registration store unavailable.");

        public Task<GoogleDriveFileRef?> ResolveAsync(string registrationId, CancellationToken cancellationToken = default) =>
            Task.FromResult<GoogleDriveFileRef?>(null);

        public Task<bool> RemoveAsync(string registrationId, CancellationToken cancellationToken = default) =>
            Task.FromResult(false);
    }

    /// <summary>Records token requests and answers each with one fixed response.</summary>
    internal sealed class TokenEndpoint : HttpMessageHandler
    {
        private readonly string _response;
        private readonly HttpStatusCode _status;

        public TokenEndpoint(string response, HttpStatusCode status = HttpStatusCode.OK)
        {
            _response = response;
            _status = status;
        }

        public List<(string Url, Dictionary<string, string> Form)> Requests { get; } = new();

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            var body = await request.Content!.ReadAsStringAsync(cancellationToken);
            var form = body.Split('&', StringSplitOptions.RemoveEmptyEntries)
                .Select(pair => pair.Split('=', 2))
                .ToDictionary(pair => Uri.UnescapeDataString(pair[0]), pair => Uri.UnescapeDataString(pair[1].Replace('+', ' ')));
            Requests.Add((request.RequestUri!.ToString(), form));
            return new HttpResponseMessage(_status)
            {
                Content = new StringContent(_response, Encoding.UTF8, "application/json")
            };
        }
    }

    /// <summary>
    /// Simulates the Drive API v3 file endpoints the provider uses: <c>files.get</c> (metadata
    /// and <c>alt=media</c>), <c>files.list</c> name queries, media <c>files.update</c>, and
    /// multipart <c>files.create</c>. Each file has a head revision that changes only with
    /// content, and a change counter that changes with everything.
    /// </summary>
    internal sealed class FakeDrive : IDisposable
    {
        public const string Token = "ya29.fake-drive-token";
        public const string FolderId = "folderContracts01";
        public const string Api = "https://drive.fake/drive/v3";
        public const string Upload = "https://drive.fake/upload/drive/v3";
        public const string TokenUrl = "https://oauth2.googleapis.com/token";

        private readonly Dictionary<string, Entry> _files = new(StringComparer.Ordinal);
        private readonly HttpClient _http;
        private int _nextId;
        private int _nextRevision;

        /// <summary>The client that reaches this fake, for a provider built elsewhere.</summary>
        public HttpClient Http => _http;

        /// <summary>How many tokens were requested from the token endpoint.</summary>
        public int TokenRequests { get; private set; }

        public FakeDrive()
        {
            _http = new HttpClient(new Handler(this));
            _files[FolderId] = new Entry
            {
                Id = FolderId, Name = "Contracts", MimeType = "application/vnd.google-apps.folder", Size = null
            };
        }

        public string? SharedDriveId { get; init; }
        public List<string> Requests { get; } = new();
        public int Uploads { get; private set; }
        public HttpStatusCode? UploadStatus { get; set; }
        public string UploadReason { get; set; } = "badRequest";
        public bool StoreThenLoseResponse { get; set; }

        /// <summary>Every folder listing answers with this status.</summary>
        public HttpStatusCode? ListStatus { get; set; }

        public string Seed(string name, byte[] bytes, string mimeType = DocxMime, long? size = -1, string parent = FolderId)
        {
            var entry = new Entry
            {
                Id = $"fileId{++_nextId:D6}abcdef",
                Name = name,
                MimeType = mimeType,
                Bytes = bytes,
                Size = size == -1 ? bytes.LongLength : size,
                Parents = new List<string> { parent },
                HeadRevision = mimeType.StartsWith("application/vnd.google-apps.", StringComparison.Ordinal) ? null : NextRevision(),
                Version = 1
            };
            _files[entry.Id] = entry;
            return entry.Id;
        }

        /// <summary>Adds a folder under <paramref name="parent"/> and returns its id.</summary>
        public string SeedFolder(string name, string parent = FolderId) =>
            Seed(name, Array.Empty<byte>(), "application/vnd.google-apps.folder", size: null, parent);

        public string Revision(string fileId) => _files[fileId].HeadRevision!;
        public byte[] Bytes(string fileId) => _files[fileId].Bytes;
        public byte[] BytesNamed(string name) => _files.Values.Single(f => f.Name == name).Bytes;
        public string MimeTypeOf(string name) => _files.Values.Single(f => f.Name == name).MimeType;
        public IReadOnlyList<string> ParentsOf(string name) => _files.Values.Single(f => f.Name == name).Parents;
        public bool HasName(string name) => _files.Values.Any(f => f.Name == name && !f.Trashed);
        public void Trash(string fileId) => _files[fileId].Trashed = true;
        public void Lock(string fileId) => _files[fileId].ReadOnly = true;
        public void MakeReadOnlyForCaller(string fileId) => _files[fileId].CanModifyContent = false;
        public void UnderReportSize(string fileId) => _files[fileId].Size = 1;
        public void RemoveFolder() => _files.Remove(FolderId);

        public void EditElsewhere(string fileId, byte[] bytes)
        {
            var entry = _files[fileId];
            entry.Bytes = bytes;
            entry.Size = bytes.LongLength;
            entry.HeadRevision = NextRevision();
            entry.Version++;
        }

        public void RenameElsewhere(string fileId, string name)
        {
            _files[fileId].Name = name;
            _files[fileId].Version++;
        }

        public GoogleDriveDocumentProvider Provider(
            long maximumBytes = 100 * 1024 * 1024,
            bool configureCreation = true,
            IGoogleDriveRegistrationStore? registrationStore = null,
            string? templatesFolder = null,
            IReadOnlyCollection<string>? allowedExtensions = null) => new(
            new GoogleDriveDocumentProviderOptions
            {
                ConnectionId = "drive",
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

        private string NextRevision() => $"0Brev{++_nextRevision:D4}";

        private sealed class FixedTokens : IGoogleDriveAccessTokenProvider
        {
            public Task<string> GetAccessTokenAsync(CancellationToken cancellationToken = default) => Task.FromResult(Token);
        }

        private sealed class Entry
        {
            public string Id = "";
            public string Name = "";
            public string MimeType = "";
            public byte[] Bytes = Array.Empty<byte>();
            public long? Size;
            public List<string> Parents = new();
            public string? HeadRevision;
            public long Version;
            public bool Trashed;
            public bool ReadOnly;
            public bool CanModifyContent = true;
        }

        private sealed class Handler : HttpMessageHandler
        {
            private readonly FakeDrive _drive;

            public Handler(FakeDrive drive) => _drive = drive;

            protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                var url = request.RequestUri!.ToString();
                _drive.Requests.Add($"{request.Method} {url}");

                // The OAuth token endpoint, so a connection the MCP server builds from configuration
                // signs in through its real token provider. The form is recorded, never echoed.
                if (request.Method == HttpMethod.Post && url == FakeDrive.TokenUrl)
                {
                    _drive.TokenRequests++;
                    return new HttpResponseMessage(HttpStatusCode.OK)
                    {
                        Content = new StringContent(
                            JsonSerializer.Serialize(new { access_token = Token, expires_in = 3600 }), Encoding.UTF8, "application/json")
                    };
                }
                if (request.Headers.Authorization is not { Scheme: "Bearer", Parameter: Token })
                    return Error(HttpStatusCode.Unauthorized, "authError");

                var query = Query(request.RequestUri!);
                var path = request.RequestUri!.GetLeftPart(UriPartial.Path);

                if (path.StartsWith(Upload + "/files", StringComparison.Ordinal))
                {
                    if (_drive.UploadStatus is { } simulated)
                    {
                        _drive.UploadStatus = null;
                        return Error(simulated, _drive.UploadReason);
                    }
                    return path == Upload + "/files"
                        ? await CreateAsync(request, query)
                        : await UpdateAsync(request, path.Substring((Upload + "/files/").Length), query);
                }

                if (path == Api + "/files" && request.Method == HttpMethod.Get)
                    return List(query);

                if (path.StartsWith(Api + "/files/", StringComparison.Ordinal) && request.Method == HttpMethod.Get)
                {
                    var id = Uri.UnescapeDataString(path.Substring((Api + "/files/").Length));
                    if (!_drive._files.TryGetValue(id, out var entry))
                        return Error(HttpStatusCode.NotFound, "notFound", $"File not found: {id}.");
                    if (query.TryGetValue("alt", out var alt) && alt == "media")
                        return new HttpResponseMessage(HttpStatusCode.OK) { Content = new ByteArrayContent(entry.Bytes) };
                    return File(entry);
                }

                return Error(HttpStatusCode.NotFound, "notFound");
            }

            private async Task<HttpResponseMessage> UpdateAsync(HttpRequestMessage request, string id, Dictionary<string, string> query)
            {
                Assert.Equal("PATCH", request.Method.Method);
                Assert.Equal("media", query["uploadType"]);
                if (!_drive._files.TryGetValue(id, out var entry))
                    return Error(HttpStatusCode.NotFound, "notFound");
                _drive.Uploads++;
                entry.Bytes = await request.Content!.ReadAsByteArrayAsync();
                entry.Size = entry.Bytes.LongLength;
                entry.HeadRevision = _drive.NextRevision();
                entry.Version++;
                return StoredOrLost(entry);
            }

            private async Task<HttpResponseMessage> CreateAsync(HttpRequestMessage request, Dictionary<string, string> query)
            {
                Assert.Equal(HttpMethod.Post, request.Method);
                Assert.Equal("multipart", query["uploadType"]);
                Assert.Equal("related", request.Content!.Headers.ContentType!.MediaType!.Split('/')[1]);
                var parts = ((MultipartContent)request.Content).ToList();
                Assert.Equal(2, parts.Count);
                Assert.Equal("application/json", parts[0].Headers.ContentType!.MediaType);
                using var metadata = JsonDocument.Parse(await parts[0].ReadAsStringAsync());
                var parents = metadata.RootElement.GetProperty("parents").EnumerateArray().Select(p => p.GetString()!).ToList();
                if (!parents.All(p => _drive._files.ContainsKey(p)))
                    return Error(HttpStatusCode.NotFound, "notFound");

                _drive.Uploads++;
                // Drive never overwrites on create: a new file id every time, even for a taken name.
                var bytes = await parts[1].ReadAsByteArrayAsync();
                var entry = new Entry
                {
                    Id = $"fileId{++_drive._nextId:D6}abcdef",
                    Name = metadata.RootElement.GetProperty("name").GetString()!,
                    MimeType = metadata.RootElement.GetProperty("mimeType").GetString()!,
                    Bytes = bytes,
                    Size = bytes.LongLength,
                    Parents = parents,
                    HeadRevision = _drive.NextRevision(),
                    Version = 1
                };
                _drive._files[entry.Id] = entry;
                return StoredOrLost(entry);
            }

            private HttpResponseMessage StoredOrLost(Entry entry)
            {
                if (_drive.StoreThenLoseResponse)
                {
                    _drive.StoreThenLoseResponse = false;
                    throw new TaskCanceledException("The request timed out after the upload was stored.");
                }
                return File(entry);
            }

            private HttpResponseMessage List(Dictionary<string, string> query)
            {
                Assert.Equal("true", query["includeItemsFromAllDrives"]);
                if (_drive.SharedDriveId is not null)
                {
                    Assert.Equal("drive", query["corpora"]);
                    Assert.Equal(_drive.SharedDriveId, query["driveId"]);
                }
                var children = Regex.Match(query["q"], @"^'([^']*)' in parents and trashed = false$");
                if (children.Success)
                    return Children(children.Groups[1].Value, query);

                var match = Regex.Match(query["q"],
                    @"^name = '((?:[^'\\]|\\.)*)' and '([^']*)' in parents (?:and mimeType = '([^']*)' )?and trashed = false$");
                Assert.True(match.Success, query["q"]);
                var name = Regex.Unescape(match.Groups[1].Value.Replace("\\'", "'"));
                var parent = match.Groups[2].Value;
                var mimeType = match.Groups[3].Success ? match.Groups[3].Value : null;
                var files = _drive._files.Values
                    .Where(f => !f.Trashed && f.Parents.Contains(parent) && string.Equals(f.Name, name, StringComparison.OrdinalIgnoreCase))
                    .Where(f => mimeType is null || f.MimeType == mimeType)
                    .Select(f => new Dictionary<string, object> { ["id"] = f.Id, ["name"] = f.Name })
                    .ToList();
                return Json(new Dictionary<string, object> { ["files"] = files });
            }

            /// <summary>files.list over one folder: the page token is the index of the next file.</summary>
            private HttpResponseMessage Children(string parent, Dictionary<string, string> query)
            {
                if (_drive.ListStatus is { } status) return Error(status, "insufficientFilePermissions");
                if (!_drive._files.ContainsKey(parent)) return Error(HttpStatusCode.NotFound, "notFound");
                var size = int.Parse(query["pageSize"]);
                var start = query.TryGetValue("pageToken", out var token) ? int.Parse(token) : 0;
                var all = _drive._files.Values
                    .Where(f => !f.Trashed && f.Parents.Contains(parent))
                    .OrderBy(f => f.Id, StringComparer.Ordinal)
                    .ToList();
                var files = all.Skip(start).Take(size).Select(f =>
                {
                    var file = new Dictionary<string, object?>
                    {
                        ["id"] = f.Id, ["name"] = f.Name, ["mimeType"] = f.MimeType, ["trashed"] = f.Trashed,
                        ["modifiedTime"] = "2026-11-01T09:00:00.000Z"
                    };
                    if (f.Size is { } bytes) file["size"] = bytes.ToString();
                    return file;
                }).ToList();
                var payload = new Dictionary<string, object> { ["files"] = files };
                if (start + size < all.Count) payload["nextPageToken"] = (start + size).ToString();
                return Json(payload);
            }

            private HttpResponseMessage File(Entry entry)
            {
                var file = new Dictionary<string, object?>
                {
                    ["id"] = entry.Id,
                    ["name"] = entry.Name,
                    ["mimeType"] = entry.MimeType,
                    ["version"] = entry.Version.ToString(),
                    ["parents"] = entry.Parents,
                    ["trashed"] = entry.Trashed,
                    ["capabilities"] = new Dictionary<string, object> { ["canModifyContent"] = entry.CanModifyContent }
                };
                if (entry.Size is { } size) file["size"] = size.ToString();
                if (entry.HeadRevision is not null) file["headRevisionId"] = entry.HeadRevision;
                if (_drive.SharedDriveId is not null) file["driveId"] = _drive.SharedDriveId;
                if (entry.ReadOnly)
                    file["contentRestrictions"] = new[] { new Dictionary<string, object> { ["readOnly"] = true, ["type"] = "globalContentRestriction" } };
                return Json(file);
            }

            private static Dictionary<string, string> Query(Uri uri) => uri.Query.TrimStart('?')
                .Split('&', StringSplitOptions.RemoveEmptyEntries)
                .Select(pair => pair.Split('=', 2))
                .ToDictionary(pair => pair[0], pair => pair.Length > 1 ? Uri.UnescapeDataString(pair[1]) : string.Empty);

            private static HttpResponseMessage Json(object payload) => new(HttpStatusCode.OK)
            {
                Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
            };

            private static HttpResponseMessage Error(HttpStatusCode status, string reason, string? message = null) => new(status)
            {
                Content = new StringContent(JsonSerializer.Serialize(new
                {
                    error = new
                    {
                        code = (int)status,
                        message = message ?? reason,
                        errors = new[] { new { domain = "global", reason, message = message ?? reason } }
                    }
                }), Encoding.UTF8, "application/json")
            };
        }
    }
}
