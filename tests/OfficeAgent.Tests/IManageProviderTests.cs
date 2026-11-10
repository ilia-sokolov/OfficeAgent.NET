using System.Net;
using System.Text;
using System.Text.Json;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.IManage;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// iManage Work provider tests over a fake Work API v2 server: documents with numbered versions,
/// locks, folders, and multipart uploads simulated by an in-memory <see cref="HttpMessageHandler"/>,
/// so the register → open → save → remove loop runs without an iManage customer.
/// </summary>
public class IManageProviderTests
{
    [Fact]
    public async Task A_read_only_connection_registers_and_opens_but_sends_no_write()
    {
        using var server = new FakeIManage();
        var original = DocxFactory.Contract();
        var number = server.Seed("Invoice.docx", original);
        var provider = server.Provider(configureCreation: false, readOnly: true);

        Assert.True(provider.ReadOnly);
        var reference = await provider.RegisterAsync(server.IdOf(number));
        using (var content = await provider.OpenReadAsync(reference))
            Assert.True(content.Stream.Length > 0);
        var before = server.Requests.Count;

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
        // Nothing reached iManage: not even a lookup, let alone an upload.
        Assert.Equal(before, server.Requests.Count);
        await provider.RemoveAsync(reference);
    }

    [Fact]
    public void A_read_only_connection_cannot_also_name_a_creation_folder()
    {
        using var server = new FakeIManage();

        var error = Assert.Throws<ArgumentException>(() => server.Provider(configureCreation: true, readOnly: true));

        Assert.Contains("CreationFolderId", error.Message, StringComparison.Ordinal);
    }

    private static OfficeAgentClient EngineOver(IManageDocumentProvider provider) =>
        new(new DocumentProviderRegistry(new IDocumentProvider[] { provider }), new WordModule());

    // Built once, so byte comparisons compare the same bytes.
    private static readonly byte[] Original = DocxFactory.Contract();
    private static readonly byte[] EditedBytes = Original.Concat(new byte[] { 0 }).ToArray();

    private static MemoryStream Docx() => new(Original);

    private static byte[] Edited() => EditedBytes;

    /// <summary>Runs a save the way the engine does, so the asserted outcome is the caller's.</summary>
    private static Task<DocumentReference> EngineSave(
        IManageDocumentProvider provider, DocumentReference source, SaveDocumentOptions options) =>
        StorageWrite.RunAsync(provider, source.ItemId, outputName: null, EditedBytes,
            token => provider.SaveAsync(source, new MemoryStream(EditedBytes), options, token), CancellationToken.None);

    // ── Register ──────────────────────────────────────────────────────────────

    [Fact]
    public async Task Register_by_document_id_assigns_an_opaque_id_and_opens_by_it()
    {
        using var server = new FakeIManage();
        var number = server.Seed("contract.docx", Original);
        var provider = server.Provider();

        var added = await provider.RegisterAsync(server.IdOf(number));

        Assert.Equal("imanage", added.Provider);
        Assert.Equal("legal", added.ConnectionId);
        Assert.Equal("contract.docx", added.Name);
        Assert.DoesNotContain("ACTIVE", added.ItemId);
        Assert.DoesNotContain(number.ToString(), added.ItemId);
        Assert.DoesNotContain("ACTIVE", added.Version);
        Assert.False(string.IsNullOrEmpty(added.Version));

        using var content = await provider.OpenReadAsync(added);
        Assert.Equal(added.ItemId, content.Reference.ItemId);
        Assert.Equal(Original, content.Stream is MemoryStream m ? m.ToArray() : null);
    }

    [Theory]
    [InlineData("https://imanage.fake/work/web/r/libraries/ACTIVE/documents/{id}")]
    [InlineData("https://imanage.fake/work/link/d/{escaped}")]
    [InlineData("https://imanage.fake/work/web/?lib=ACTIVE&num={num}&ver=1")]
    public async Task Register_by_iManage_work_link(string template)
    {
        using var server = new FakeIManage();
        var number = server.Seed("contract.docx", Original);
        var provider = server.Provider();
        var id = server.IdOf(number);
        var url = template.Replace("{id}", id).Replace("{escaped}", Uri.EscapeDataString(id))
            .Replace("{num}", number.ToString());

        var added = await provider.RegisterAsync(url);

        Assert.Equal("contract.docx", added.Name);
    }

    [Theory]
    [InlineData("contract.docx")]
    [InlineData("ACTIVE!1000")]
    [InlineData("ACTIVE/1000/1")]
    [InlineData("https://elsewhere.fake/work/link/d/ACTIVE!1000.1")]
    [InlineData("http://imanage.fake/work/link/d/ACTIVE!1000.1")]
    [InlineData("https://imanage.fake/work/web/")]
    public async Task Register_rejects_sources_that_are_not_a_document_on_this_server(string source)
    {
        using var server = new FakeIManage();
        server.Seed("contract.docx", Original);

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() => server.Provider().RegisterAsync(source));

        Assert.Equal(ProviderErrorCode.InvalidArgument, error.Code);
        Assert.Empty(server.Requests);
    }

    [Fact]
    public async Task Register_refuses_an_older_version_of_the_document()
    {
        using var server = new FakeIManage();
        var number = server.Seed("contract.docx", Original);
        server.AddVersion(number, Edited());

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            server.Provider().RegisterAsync($"ACTIVE!{number}.1"));

        Assert.Equal(ProviderErrorCode.InvalidArgument, error.Code);
        Assert.Contains("latest version is 2", error.Message);
    }

    [Fact]
    public async Task Register_rejects_a_disallowed_extension_from_the_profile_without_downloading()
    {
        using var server = new FakeIManage();
        var number = server.Seed("memo.pdf", Original);

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            server.Provider().RegisterAsync(server.IdOf(number)));

        Assert.Equal(ProviderErrorCode.ExtensionNotAllowed, error.Code);
        Assert.Equal(0, server.Downloads);
    }

    [Fact]
    public async Task Register_rejects_an_oversized_document_from_the_profile_without_downloading()
    {
        using var server = new FakeIManage();
        var number = server.Seed("contract.docx", Original);

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            server.Provider(maximumBytes: 16).RegisterAsync(server.IdOf(number)));

        Assert.Equal(ProviderErrorCode.ContentTooLarge, error.Code);
        Assert.Equal(0, server.Downloads);
    }

    [Fact]
    public async Task Open_rejects_a_document_that_grew_past_the_limit_before_downloading()
    {
        using var server = new FakeIManage();
        var number = server.Seed("contract.docx", new byte[10]);
        var provider = server.Provider(maximumBytes: 100);
        var added = await provider.RegisterAsync(server.IdOf(number));
        server.AddVersion(number, new byte[500]);

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.OpenReadAsync(DocumentReference.For("imanage", "legal", added.ItemId)));

        Assert.Equal(ProviderErrorCode.ContentTooLarge, error.Code);
        Assert.Equal(0, server.Downloads);
    }

    [Fact]
    public async Task Register_reports_a_missing_document_as_not_found_without_upstream_detail()
    {
        using var server = new FakeIManage { ErrorMessage = "Document ACTIVE!999 in matter 'Project Falcon' denied." };

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            server.Provider().RegisterAsync("ACTIVE!999.1"));

        Assert.Equal(ProviderErrorCode.NotFound, error.Code);
        Assert.Contains("iManage code: NRC_NO_RECORD", error.Message);
        Assert.DoesNotContain("Project Falcon", error.Message);
    }

    // ── Save modes ────────────────────────────────────────────────────────────

    [Fact]
    public async Task Replace_adds_a_new_iManage_version_and_keeps_the_id_and_every_earlier_version()
    {
        using var server = new FakeIManage();
        var number = server.Seed("contract.docx", Original);
        var provider = server.Provider();
        var source = await provider.RegisterAsync(server.IdOf(number));

        using var edited = new MemoryStream(Edited());
        var saved = await provider.SaveAsync(source, edited, new SaveDocumentOptions { Mode = SaveMode.Replace });

        Assert.Equal(source.ItemId, saved.ItemId);
        Assert.NotEqual(source.Version, saved.Version);
        Assert.Equal(2, server.VersionCount(number));
        Assert.Equal(Original, server.BytesOf(number, 1));
        Assert.Equal(Edited(), server.BytesOf(number, 2));
        Assert.False(server.IsLocked(number));
        Assert.Equal("Saved by OfficeAgent.", server.CommentOf(number, 2));

        using var reopened = await provider.OpenReadAsync(saved);
        Assert.Equal(saved.Version, reopened.Reference.Version);
    }

    [Fact]
    public async Task Replace_with_a_stale_version_is_a_conflict_and_writes_nothing()
    {
        using var server = new FakeIManage();
        var number = server.Seed("contract.docx", Original);
        var provider = server.Provider();
        var source = await provider.RegisterAsync(server.IdOf(number));
        server.AddVersion(number, Edited());

        await Assert.ThrowsAsync<DocumentVersionConflictException>(() =>
            provider.SaveAsync(source, Docx(), new SaveDocumentOptions { Mode = SaveMode.Replace }));

        Assert.Equal(2, server.VersionCount(number));
        Assert.Equal(0, server.Uploads);
    }

    [Fact]
    public async Task Replace_after_an_in_place_edit_of_the_latest_version_is_a_conflict()
    {
        using var server = new FakeIManage();
        var number = server.Seed("contract.docx", Original);
        var provider = server.Provider();
        var source = await provider.RegisterAsync(server.IdOf(number));
        server.ReplaceLatestInPlace(number, Edited());

        await Assert.ThrowsAsync<DocumentVersionConflictException>(() =>
            provider.SaveAsync(source, Docx(), new SaveDocumentOptions { Mode = SaveMode.Replace }));
        Assert.Equal(0, server.Uploads);
    }

    [Fact]
    public async Task A_version_added_while_the_save_takes_its_lock_is_a_conflict_and_the_lock_is_released()
    {
        using var server = new FakeIManage();
        var number = server.Seed("contract.docx", Original);
        var provider = server.Provider();
        var source = await provider.RegisterAsync(server.IdOf(number));
        server.OnLock = () => server.AddVersion(number, Edited());

        await Assert.ThrowsAsync<DocumentVersionConflictException>(() =>
            provider.SaveAsync(source, Docx(), new SaveDocumentOptions { Mode = SaveMode.Replace }));

        Assert.Equal(0, server.Uploads);
        Assert.False(server.IsLocked(number));
    }

    [Fact]
    public async Task Without_locking_a_save_still_checks_the_opened_version()
    {
        using var server = new FakeIManage();
        var number = server.Seed("contract.docx", Original);
        var provider = server.Provider(lockDuringSave: false);
        var source = await provider.RegisterAsync(server.IdOf(number));

        var saved = await provider.SaveAsync(source, new MemoryStream(Edited()), new SaveDocumentOptions());
        Assert.Equal(0, server.Locks);
        Assert.Equal(2, server.VersionCount(number));

        await Assert.ThrowsAsync<DocumentVersionConflictException>(() =>
            provider.SaveAsync(source, Docx(), new SaveDocumentOptions()));
        Assert.Equal(2, server.VersionCount(number));
        Assert.NotEqual(source.Version, saved.Version);
    }

    [Fact]
    public async Task New_version_adds_an_iManage_version_and_mints_a_fresh_id()
    {
        using var server = new FakeIManage();
        var number = server.Seed("contract.docx", Original);
        var provider = server.Provider();
        var source = await provider.RegisterAsync(server.IdOf(number));

        var saved = await provider.SaveAsync(source, new MemoryStream(Edited()),
            new SaveDocumentOptions { Mode = SaveMode.NewVersion });

        Assert.NotEqual(source.ItemId, saved.ItemId);
        Assert.Equal("contract.docx", saved.Name);
        Assert.Equal(2, server.VersionCount(number));
        Assert.Equal(Original, server.BytesOf(number, 1));
        using var reopened = await provider.OpenReadAsync(saved);
        Assert.Equal(Edited(), ((MemoryStream)reopened.Stream).ToArray());
    }

    [Fact]
    public async Task New_document_files_a_versioned_sibling_in_the_creation_folder()
    {
        using var server = new FakeIManage();
        var number = server.Seed("contract.docx", Original, inFolder: true);
        var provider = server.Provider();
        var source = await provider.RegisterAsync(server.IdOf(number));

        var saved = await provider.SaveAsync(source, new MemoryStream(Edited()),
            new SaveDocumentOptions { Mode = SaveMode.NewDocument });

        Assert.Equal("contract.v2.docx", saved.Name);
        Assert.NotEqual(source.ItemId, saved.ItemId);
        Assert.Equal(1, server.VersionCount(number));
        Assert.True(server.FolderHas("contract.v2.docx"));
        Assert.Equal("legal-class", server.ClassOf("contract.v2.docx"));
    }

    [Fact]
    public async Task New_document_honours_a_validated_new_name_and_refuses_a_taken_one()
    {
        using var server = new FakeIManage();
        var number = server.Seed("contract.docx", Original);
        server.Seed("taken.docx", Original, inFolder: true);
        var provider = server.Provider();
        var source = await provider.RegisterAsync(server.IdOf(number));

        var saved = await provider.SaveAsync(source, Docx(),
            new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "globex.docx" });
        Assert.Equal("globex.docx", saved.Name);

        var taken = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(source, Docx(),
            new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "TAKEN.docx" }));
        Assert.Equal(ProviderErrorCode.AlreadyExists, taken.Code);

        var nested = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(source, Docx(),
            new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "nested/name.docx" }));
        Assert.Equal(ProviderErrorCode.InvalidArgument, nested.Code);
    }

    [Fact]
    public async Task Save_rejects_caller_chosen_destinations_and_renamed_versions()
    {
        using var server = new FakeIManage();
        var number = server.Seed("contract.docx", Original);
        var provider = server.Provider();
        var source = await provider.RegisterAsync(server.IdOf(number));

        var destination = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(
            source, Docx(), new SaveDocumentOptions { DestinationItemId = "chosen" }));
        Assert.Equal(ProviderErrorCode.InvalidArgument, destination.Code);

        foreach (var mode in new[] { SaveMode.Replace, SaveMode.NewVersion })
        {
            var rename = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(
                source, Docx(), new SaveDocumentOptions { Mode = mode, NewName = "renamed.docx" }));
            Assert.Equal(ProviderErrorCode.InvalidArgument, rename.Code);
        }
        Assert.Equal(0, server.Uploads);
    }

    [Fact]
    public async Task New_document_without_a_creation_folder_is_a_configuration_error()
    {
        using var server = new FakeIManage();
        var number = server.Seed("contract.docx", Original);
        var provider = server.Provider(configureCreation: false);
        var source = await provider.RegisterAsync(server.IdOf(number));

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(
            source, Docx(), new SaveDocumentOptions { Mode = SaveMode.NewDocument }));

        Assert.Equal(ProviderErrorCode.ConfigurationError, error.Code);
        Assert.Equal(0, server.Uploads);
    }

    [Fact]
    public async Task Remove_drops_the_registration_but_not_the_iManage_document()
    {
        using var server = new FakeIManage();
        var number = server.Seed("contract.docx", Original);
        var provider = server.Provider();
        var added = await provider.RegisterAsync(server.IdOf(number));

        await provider.RemoveAsync(added);

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.OpenReadAsync(added));
        Assert.Equal(ProviderErrorCode.NotFound, error.Code);
        Assert.Equal(1, server.VersionCount(number));
        var again = await provider.RegisterAsync(server.IdOf(number));
        Assert.NotEqual(added.ItemId, again.ItemId);
    }

    // ── Creation ──────────────────────────────────────────────────────────────

    [Fact]
    public async Task Create_files_registers_and_applies_an_initial_plan()
    {
        using var server = new FakeIManage();
        var client = EngineOver(server.Provider());
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertOp
                {
                    Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = string.Empty },
                    Position = InsertPosition.Before,
                    Text = "iManage brief"
                }
            }
        };

        var created = await client.CreateAsync("legal", "brief.docx", plan);

        Assert.True(created.Committed);
        Assert.Equal("imanage", created.Document!.Provider);
        Assert.True(server.FolderHas("brief.docx"));
        var inspected = await client.InspectAsync(created.Document);
        Assert.Equal("iManage brief", inspected.Paragraphs[0].Text);
    }

    [Fact]
    public async Task Create_never_duplicates_an_existing_name_in_the_folder()
    {
        using var server = new FakeIManage();
        var provider = server.Provider();

        await provider.CreateAsync("brief.docx", Docx());
        var error = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.CreateAsync("brief.docx", Docx()));

        Assert.Equal(ProviderErrorCode.AlreadyExists, error.Code);
        Assert.Equal(1, server.Uploads);
    }

    [Fact]
    public async Task Create_requires_a_configured_folder_and_rejects_bad_names()
    {
        using var server = new FakeIManage();

        var unconfigured = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            server.Provider(configureCreation: false).CreateAsync("brief.docx", Docx()));
        Assert.Equal(ProviderErrorCode.ConfigurationError, unconfigured.Code);

        var extension = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            server.Provider().CreateAsync("brief.exe", Docx()));
        Assert.Equal(ProviderErrorCode.ExtensionNotAllowed, extension.Code);

        var oversized = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            server.Provider(maximumBytes: 16).CreateAsync("brief.docx", Docx()));
        Assert.Equal(ProviderErrorCode.ContentTooLarge, oversized.Code);
        Assert.Equal(0, server.Uploads);
    }

    [Fact]
    public async Task Create_tool_does_not_expose_the_configured_folder()
    {
        using var server = new FakeIManage { MissingFolder = true };
        var tools = new OfficeAgentTools(EngineOver(server.Provider()));

        var payload = await tools.CreateDocument("legal", "brief.docx");

        Assert.DoesNotContain("773", payload);
        Assert.DoesNotContain("imanage.fake", payload);
        using var json = JsonDocument.Parse(payload);
        Assert.Equal("not-found", json.RootElement.GetProperty("errors")[0].GetProperty("code").GetString());
    }

    [Fact]
    public async Task A_registration_failure_after_filing_leaves_the_document_and_reports_a_locator()
    {
        using var server = new FakeIManage();
        var provider = server.Provider(registrationStore: new FailingStore());

        var error = await Assert.ThrowsAsync<DocumentRegistrationFailedException>(() =>
            provider.CreateAsync("brief.docx", Docx()));

        Assert.Equal(ProviderErrorCode.RegistrationFailed, error.Code);
        Assert.True(server.FolderHas("brief.docx"));
        Assert.Equal("brief.docx", error.OutputName);
        Assert.Equal(Sha256(Original), error.OutputSha256);
        Assert.DoesNotContain("ACTIVE", error.Message);
        Assert.DoesNotContain("imanage.fake", error.Message);
    }

    // ── Write outcomes ────────────────────────────────────────────────────────

    [Fact]
    public async Task A_document_checked_out_by_someone_else_is_refused_with_what_to_do()
    {
        using var server = new FakeIManage();
        var number = server.Seed("contract.docx", Original);
        var provider = server.Provider();
        var source = await provider.RegisterAsync(server.IdOf(number));
        server.CheckOut(number);

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.SaveAsync(source, Docx(), new SaveDocumentOptions { Mode = SaveMode.Replace }));

        Assert.Equal(ProviderErrorCode.WriteRejected, error.Code);
        Assert.Contains("checked out or locked in iManage", error.Message);
        Assert.Contains("check it in", error.Message);
        Assert.Contains("save as a new document", error.Message);
        Assert.Equal(0, server.Uploads);
        Assert.Equal(0, server.Locks);
    }

    [Theory]
    [InlineData((HttpStatusCode)423, "simulated")]
    [InlineData(HttpStatusCode.Conflict, "DOCUMENT_CHECKED_OUT")]
    public async Task A_lock_taken_by_someone_else_is_refused_and_nothing_is_written(HttpStatusCode status, string code)
    {
        using var server = new FakeIManage();
        var number = server.Seed("contract.docx", Original);
        var provider = server.Provider();
        var source = await provider.RegisterAsync(server.IdOf(number));
        server.LockStatus = status;
        server.LockErrorCode = code;

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.SaveAsync(source, Docx(), new SaveDocumentOptions { Mode = SaveMode.Replace }));

        Assert.Equal(ProviderErrorCode.WriteRejected, error.Code);
        Assert.Contains("checked out or locked in iManage", error.Message);
        Assert.Equal(0, server.Uploads);
    }

    [Fact]
    public async Task A_server_without_document_locks_says_how_to_turn_locking_off()
    {
        using var server = new FakeIManage { LockStatus = HttpStatusCode.NotFound };
        var number = server.Seed("contract.docx", Original);
        var provider = server.Provider();
        var source = await provider.RegisterAsync(server.IdOf(number));

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.SaveAsync(source, Docx(), new SaveDocumentOptions()));

        Assert.Equal(ProviderErrorCode.ConfigurationError, error.Code);
        Assert.Contains("LockDuringSave", error.Message);
        Assert.Equal(0, server.Uploads);
    }

    [Theory]
    [InlineData(HttpStatusCode.BadRequest)]
    [InlineData(HttpStatusCode.TooManyRequests)]
    public async Task A_refused_version_upload_is_certain_and_releases_the_lock(HttpStatusCode status)
    {
        using var server = new FakeIManage();
        var number = server.Seed("contract.docx", Original);
        var provider = server.Provider();
        var source = await provider.RegisterAsync(server.IdOf(number));
        server.UploadStatus = status;

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            EngineSave(provider, source, new SaveDocumentOptions { Mode = SaveMode.Replace }));

        Assert.Equal(ProviderErrorCode.WriteRejected, error.Code);
        Assert.Equal(1, server.VersionCount(number));
        Assert.False(server.IsLocked(number));
    }

    [Fact]
    public async Task A_refused_filing_is_certain_and_leaves_nothing_behind()
    {
        using var server = new FakeIManage { UploadStatus = HttpStatusCode.BadRequest };

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            EngineOver(server.Provider()).CreateAsync("legal", "refused.docx"));

        Assert.Equal(ProviderErrorCode.WriteRejected, error.Code);
        Assert.False(server.FolderHas("refused.docx"));
    }

    [Fact]
    public async Task A_server_error_on_a_version_upload_is_an_unknown_outcome()
    {
        using var server = new FakeIManage();
        var number = server.Seed("contract.docx", Original);
        var provider = server.Provider();
        var source = await provider.RegisterAsync(server.IdOf(number));
        server.UploadStatus = HttpStatusCode.ServiceUnavailable;

        var error = await Assert.ThrowsAsync<DocumentWriteOutcomeUnknownException>(() =>
            EngineSave(provider, source, new SaveDocumentOptions { Mode = SaveMode.Replace }));

        Assert.Equal(ProviderErrorCode.OutcomeUnknown, error.Code);
        Assert.False(server.IsLocked(number));
    }

    [Fact]
    public async Task A_lost_response_after_a_stored_filing_is_unknown_and_reconcilable()
    {
        using var server = new FakeIManage { StoreThenLoseResponse = true };

        var error = await Assert.ThrowsAsync<DocumentWriteOutcomeUnknownException>(() =>
            EngineOver(server.Provider()).CreateAsync("legal", "lost.docx"));

        Assert.Equal(ProviderErrorCode.OutcomeUnknown, error.Code);
        Assert.Equal("lost.docx", error.OutputName);
        Assert.True(server.FolderHas("lost.docx"));
        Assert.Equal(Sha256(server.BytesNamed("lost.docx")), error.OutputSha256);
    }

    // ── Credentials ───────────────────────────────────────────────────────────

    [Fact]
    public async Task The_token_travels_only_to_the_iManage_server_and_never_into_results()
    {
        using var server = new FakeIManage { RedirectDownloads = true };
        var number = server.Seed("contract.docx", Original);
        var provider = server.Provider();
        var added = await provider.RegisterAsync(server.IdOf(number));

        using var opened = await provider.OpenReadAsync(added);
        var saved = await provider.SaveAsync(added, new MemoryStream(Edited()), new SaveDocumentOptions());

        Assert.Equal(Original, ((MemoryStream)opened.Stream).ToArray());
        Assert.True(server.SawStorageDownload);
        Assert.False(server.SawTokenOffServer);
        Assert.False(server.SawAuthorizationHeader);
        foreach (var reference in new[] { added, opened.Reference, saved })
            Assert.DoesNotContain(FakeIManage.Token, JsonSerializer.Serialize(reference));
    }

    [Fact]
    public async Task A_client_that_followed_a_redirect_off_the_server_fails_loudly()
    {
        using var server = new FakeIManage { AnswerFromAnotherHost = true };

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            server.Provider().RegisterAsync("ACTIVE!1000.1"));

        Assert.Equal(ProviderErrorCode.IO, error.Code);
        Assert.Contains("CreateHttpClient", error.Message);
    }


    // ── OAuth ─────────────────────────────────────────────────────────────────

    [Fact]
    public async Task OAuth_password_grant_posts_the_documented_form_and_caches_the_token()
    {
        var handler = new TokenEndpoint();
        var provider = new IManageOAuthTokenProvider(new IManageOAuthOptions
        {
            ServerUrl = "https://imanage.fake",
            ClientId = "client-1",
            ClientSecret = "secret-1",
            Username = "svc-officeagent",
            Password = "pw-1"
        }, new HttpClient(handler));

        var first = await provider.GetAccessTokenAsync();
        var second = await provider.GetAccessTokenAsync();

        Assert.Equal("access-1", first);
        Assert.Equal(first, second);
        Assert.Single(handler.Forms);
        Assert.Equal("https://imanage.fake/auth/oauth2/token", handler.Urls[0]);
        var form = handler.Forms[0];
        Assert.Equal("password", form["grant_type"]);
        Assert.Equal("svc-officeagent", form["username"]);
        Assert.Equal("client-1", form["client_id"]);
        Assert.Equal("user", form["scope"]);
    }

    [Fact]
    public async Task OAuth_refresh_grant_keeps_the_rotated_refresh_token()
    {
        var handler = new TokenEndpoint { ExpiresIn = 0, RotateRefreshToken = true };
        var provider = new IManageOAuthTokenProvider(new IManageOAuthOptions
        {
            ServerUrl = "https://imanage.fake",
            ClientId = "client-1",
            GrantType = IManageGrantType.RefreshToken,
            RefreshToken = "refresh-0"
        }, new HttpClient(handler));

        await provider.GetAccessTokenAsync();
        // expires_in 0 is floored to a minute; force the next call to refresh.
        typeof(IManageOAuthTokenProvider).GetField("_expiresAt",
                System.Reflection.BindingFlags.Instance | System.Reflection.BindingFlags.NonPublic)!
            .SetValue(provider, DateTimeOffset.MinValue);
        await provider.GetAccessTokenAsync();

        Assert.Equal("refresh-0", handler.Forms[0]["refresh_token"]);
        Assert.Equal("refresh-1", handler.Forms[1]["refresh_token"]);
        Assert.False(handler.Forms[0].ContainsKey("client_secret"));
    }

    [Fact]
    public async Task OAuth_failure_names_the_error_code_but_never_the_secret()
    {
        var handler = new TokenEndpoint { Fail = true };
        var provider = new IManageOAuthTokenProvider(new IManageOAuthOptions
        {
            ServerUrl = "https://imanage.fake",
            ClientId = "client-1",
            ClientSecret = "secret-1",
            Username = "svc-officeagent",
            Password = "pw-1"
        }, new HttpClient(handler));

        var error = await Assert.ThrowsAsync<InvalidOperationException>(() => provider.GetAccessTokenAsync());

        Assert.Contains("invalid_grant", error.Message);
        Assert.DoesNotContain("secret-1", error.Message);
        Assert.DoesNotContain("pw-1", error.Message);
    }

    [Fact]
    public void OAuth_options_are_validated_per_grant()
    {
        Assert.Throws<ArgumentException>(() => new IManageOAuthTokenProvider(
            new IManageOAuthOptions { ServerUrl = "http://imanage.fake", ClientId = "c", Username = "u", Password = "p" },
            new HttpClient()));
        Assert.Throws<ArgumentException>(() => new IManageOAuthTokenProvider(
            new IManageOAuthOptions { ServerUrl = "https://imanage.fake", ClientId = "c" }, new HttpClient()));
        Assert.Throws<ArgumentException>(() => new IManageOAuthTokenProvider(
            new IManageOAuthOptions
            {
                ServerUrl = "https://imanage.fake", ClientId = "c", GrantType = IManageGrantType.ClientCredentials
            }, new HttpClient()));
    }

    // ── Registration store ────────────────────────────────────────────────────

    [Fact]
    public async Task Json_file_registration_store_survives_reload()
    {
        var directory = Path.Combine(Path.GetTempPath(), $"officeagent-imindex-{Guid.NewGuid():N}");
        var path = Path.Combine(directory, "index.json");
        try
        {
            var store = new JsonFileIManageRegistrationStore(path);
            var id = await store.AddAsync(new IManageDocumentRef("ACTIVE", 1234, 3));

            var reloaded = new JsonFileIManageRegistrationStore(path);
            var resolved = await reloaded.ResolveAsync(id);
            Assert.Equal(new IManageDocumentRef("ACTIVE", 1234, 3), resolved);
            Assert.Equal("ACTIVE!1234.3", resolved!.Value.DocumentId);

            Assert.True(await reloaded.RemoveAsync(id));
            Assert.Null(await new JsonFileIManageRegistrationStore(path).ResolveAsync(id));
            Assert.DoesNotContain("token", File.ReadAllText(path), StringComparison.OrdinalIgnoreCase);
        }
        finally
        {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void Provider_options_are_validated()
    {
        static IManageDocumentProvider Build(Action<IManageDocumentProviderOptions> change)
        {
            var options = new IManageDocumentProviderOptions
            {
                ConnectionId = "legal", ServerUrl = "https://imanage.fake", CustomerId = "42"
            };
            change(options);
            return new IManageDocumentProvider(options, new HttpClient(), new StaticTokens());
        }

        Build(_ => { });
        Assert.Throws<ArgumentException>(() => Build(o => o.ServerUrl = "http://imanage.fake"));
        Assert.Throws<ArgumentException>(() => Build(o => o.CustomerId = ""));
        Assert.Throws<ArgumentException>(() => Build(o => o.CreationFolderId = "ACTIVE!773.1"));
        Assert.Throws<ArgumentOutOfRangeException>(() => Build(o => o.MaximumBytes = 0));
    }

    // ── Helpers ───────────────────────────────────────────────────────────────


    private static string Sha256(byte[] bytes) =>
        Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(bytes)).ToLowerInvariant();

    private sealed class StaticTokens : IIManageAccessTokenProvider
    {
        public Task<string> GetAccessTokenAsync(CancellationToken cancellationToken = default) =>
            Task.FromResult(FakeIManage.Token);
    }

    private sealed class FailingStore : IIManageRegistrationStore
    {
        public Task<string> AddAsync(IManageDocumentRef document, CancellationToken cancellationToken = default) =>
            throw new IOException("Registration store unavailable.");

        public Task<IManageDocumentRef?> ResolveAsync(string registrationId, CancellationToken cancellationToken = default) =>
            Task.FromResult<IManageDocumentRef?>(null);

        public Task<bool> RemoveAsync(string registrationId, CancellationToken cancellationToken = default) =>
            Task.FromResult(false);
    }

    private sealed class TokenEndpoint : HttpMessageHandler
    {
        public List<Dictionary<string, string>> Forms { get; } = new();
        public List<string> Urls { get; } = new();
        public int ExpiresIn { get; init; } = 3600;
        public bool RotateRefreshToken { get; init; }
        public bool Fail { get; init; }

        protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
        {
            Urls.Add(request.RequestUri!.ToString());
            var body = await request.Content!.ReadAsStringAsync(cancellationToken);
            Forms.Add(body.Split('&').Select(pair => pair.Split('=', 2))
                .ToDictionary(p => Uri.UnescapeDataString(p[0]), p => Uri.UnescapeDataString(p[1].Replace('+', ' '))));
            if (Fail)
                return new HttpResponseMessage(HttpStatusCode.BadRequest)
                {
                    Content = new StringContent(
                        "{\"error\":\"invalid_grant\",\"error_description\":\"bad password pw-1 for client secret-1\"}")
                };
            var payload = new Dictionary<string, object>
            {
                ["access_token"] = $"access-{Forms.Count}",
                ["expires_in"] = ExpiresIn
            };
            if (RotateRefreshToken) payload["refresh_token"] = $"refresh-{Forms.Count}";
            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
            };
        }
    }

    /// <summary>
    /// Simulates one iManage customer's Work API v2: library <c>ACTIVE</c>, documents with
    /// numbered versions, locks, one creation folder, multipart uploads, and an optional storage
    /// host that downloads redirect to.
    /// </summary>
    internal sealed class FakeIManage : IDisposable
    {
        public const string Token = "fake-imanage-token";
        public const string Server = "https://imanage.fake";
        public const string TokenUrl = Server + "/auth/oauth2/token";
        private const string Api = Server + "/work/api/v2/customers/42/libraries/ACTIVE/";
        private const string Storage = "https://storage.fake/";

        private readonly Dictionary<long, Doc> _docs = new();
        private readonly List<long> _folder = new();
        private readonly HttpClient _http;
        private long _nextNumber = 1000;
        private int _clock;

        public FakeIManage() => _http = new HttpClient(new Handler(this));

        /// <summary>The client that reaches this fake, for a provider built elsewhere.</summary>
        public HttpClient Http => _http;

        /// <summary>How many tokens were requested from the token endpoint.</summary>
        public int TokenRequests { get; private set; }

        public List<string> Requests { get; } = new();
        public int Downloads { get; private set; }
        public int Uploads { get; private set; }
        public int Locks { get; private set; }
        public bool SawTokenOffServer { get; private set; }
        public bool SawAuthorizationHeader { get; private set; }
        public bool SawStorageDownload { get; private set; }

        public string? ErrorMessage { get; init; }
        public bool MissingFolder { get; init; }
        public bool RedirectDownloads { get; init; }
        public bool AnswerFromAnotherHost { get; init; }
        public HttpStatusCode? UploadStatus { get; set; }
        public HttpStatusCode? LockStatus { get; set; }
        public string LockErrorCode { get; set; } = "simulated";
        public bool StoreThenLoseResponse { get; set; }
        public Action? OnLock { get; set; }

        public long Seed(string fileName, byte[] bytes, bool inFolder = false)
        {
            var number = _nextNumber++;
            var doc = new Doc { Number = number };
            doc.Versions.Add(NewVersion(1, Path.GetFileNameWithoutExtension(fileName),
                Path.GetExtension(fileName).TrimStart('.').ToUpperInvariant(), bytes, comment: null, cls: null));
            _docs[number] = doc;
            if (inFolder) _folder.Add(number);
            return number;
        }

        /// <summary>Every folder listing answers with this status.</summary>
        public HttpStatusCode? ListStatus { get; set; }

        // Subfolders by id, with their name and parent. ACTIVE!774 is the folder the creation
        // folder has always held, named like a document to prove name checks pass folders over.
        private readonly Dictionary<string, (string Name, string Parent)> _subfolders = new()
        {
            ["ACTIVE!774"] = ("brief", "ACTIVE!773")
        };
        private readonly Dictionary<string, List<long>> _filed = new();

        /// <summary>Adds a folder under <paramref name="parent"/> and returns its id.</summary>
        public string AddFolder(string name, string parent = "ACTIVE!773")
        {
            var id = $"ACTIVE!{_nextNumber++}";
            _subfolders[id] = (name, parent);
            return id;
        }

        /// <summary>Files a new document in a folder other than the creation folder.</summary>
        public long SeedIn(string folderId, string fileName, byte[] bytes)
        {
            var number = Seed(fileName, bytes);
            if (!_filed.TryGetValue(folderId, out var numbers)) _filed[folderId] = numbers = new List<long>();
            numbers.Add(number);
            return number;
        }

        public string IdOf(long number) => $"ACTIVE!{number}.{_docs[number].Latest.Version}";
        public void AddVersion(long number, byte[] bytes)
        {
            var latest = _docs[number].Latest;
            _docs[number].Versions.Add(NewVersion(latest.Version + 1, latest.Name, latest.Extension, bytes, null, null));
        }

        public void ReplaceLatestInPlace(long number, byte[] bytes)
        {
            var latest = _docs[number].Latest;
            latest.Bytes = bytes;
            latest.EditDate = NextDate();
        }

        public void CheckOut(long number) => _docs[number].CheckedOut = true;
        public bool IsLocked(long number) => _docs[number].Locked;
        public int VersionCount(long number) => _docs[number].Versions.Count;
        public byte[] BytesOf(long number, int version) => _docs[number].Versions.Single(v => v.Version == version).Bytes;
        public string? CommentOf(long number, int version) => _docs[number].Versions.Single(v => v.Version == version).Comment;
        public bool FolderHas(string fileName) => FolderDoc(fileName) is not null;
        public string? ClassOf(string fileName) => FolderDoc(fileName)!.Latest.Class;
        public byte[] BytesNamed(string fileName) => FolderDoc(fileName)!.Latest.Bytes;

        private Doc? FolderDoc(string fileName) => _folder.Select(n => _docs[n]).FirstOrDefault(d =>
            string.Equals($"{d.Latest.Name}.{d.Latest.Extension}", fileName, StringComparison.OrdinalIgnoreCase));

        public IManageDocumentProvider Provider(
            long maximumBytes = 100 * 1024 * 1024,
            bool configureCreation = true,
            IIManageRegistrationStore? registrationStore = null,
            bool lockDuringSave = true,
            string? templatesFolder = null,
            IReadOnlyCollection<string>? allowedExtensions = null,
            bool readOnly = false) => new(
            new IManageDocumentProviderOptions
            {
                ConnectionId = "legal",
                ReadOnly = readOnly,
                ServerUrl = Server,
                CustomerId = "42",
                MaximumBytes = maximumBytes,
                CreationFolderId = configureCreation ? "ACTIVE!773" : string.Empty,
                CreationDocumentClass = "legal-class",
                LockDuringSave = lockDuringSave,
                TemplatesFolder = templatesFolder ?? string.Empty,
                AllowedExtensions = allowedExtensions ?? new[] { ".docx" }
            },
            _http,
            new StaticTokens(),
            registrationStore);

        public void Dispose() => _http.Dispose();

        private Ver NewVersion(int version, string name, string extension, byte[] bytes, string? comment, string? cls) => new()
        {
            Version = version, Name = name, Extension = extension, Bytes = bytes, EditDate = NextDate(),
            Comment = comment, Class = cls
        };

        private string NextDate() =>
            new DateTime(2026, 11, 1, 9, 0, 0, DateTimeKind.Utc).AddMinutes(++_clock).ToString("yyyy-MM-ddTHH:mm:ssZ");

        private sealed class Doc
        {
            public long Number;
            public readonly List<Ver> Versions = new();
            public bool CheckedOut;
            public bool Locked;
            public Ver Latest => Versions.OrderBy(v => v.Version).Last();
        }

        private sealed class Ver
        {
            public int Version;
            public string Name = "";
            public string Extension = "";
            public byte[] Bytes = Array.Empty<byte>();
            public string EditDate = "";
            public string? Comment;
            public string? Class;
        }

        private Dictionary<string, object?> ProfileOf(Doc doc, Ver version) => new()
        {
            ["id"] = $"ACTIVE!{doc.Number}.{version.Version}",
            ["database"] = "ACTIVE",
            ["document_number"] = doc.Number,
            ["version"] = version.Version,
            ["name"] = version.Name,
            ["extension"] = version.Extension,
            ["size"] = version.Bytes.LongLength,
            ["edit_date"] = version.EditDate,
            ["is_checked_out"] = doc.CheckedOut || doc.Locked,
            ["wstype"] = "document"
        };

        private sealed class Handler : HttpMessageHandler
        {
            private readonly FakeIManage _server;

            public Handler(FakeIManage server) => _server = server;

            protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                var url = request.RequestUri!.ToString();
                _server.Requests.Add($"{request.Method} {url}");

                // The OAuth token endpoint, so a connection the MCP server builds from configuration
                // signs in through its real token provider. The form is recorded, never echoed.
                if (request.Method == HttpMethod.Post && url == FakeIManage.TokenUrl)
                {
                    _server.TokenRequests++;
                    return new HttpResponseMessage(HttpStatusCode.OK)
                    {
                        Content = new StringContent(
                            JsonSerializer.Serialize(new { access_token = Token, expires_in = 3600 }), Encoding.UTF8, "application/json")
                    };
                }
                var token = request.Headers.TryGetValues("X-Auth-Token", out var values) ? values.Single() : null;
                if (request.Headers.Authorization is not null) _server.SawAuthorizationHeader = true;

                if (url.StartsWith(Storage, StringComparison.Ordinal))
                {
                    _server.SawStorageDownload = true;
                    if (token is not null) _server.SawTokenOffServer = true;
                    var parts = url.Substring(Storage.Length).Split('/');
                    return Bytes(_server._docs[long.Parse(parts[0])].Versions.Single(v => v.Version == int.Parse(parts[1])).Bytes);
                }

                if (!url.StartsWith(Server, StringComparison.Ordinal))
                {
                    if (token is not null) _server.SawTokenOffServer = true;
                    return Status(HttpStatusCode.NotFound);
                }
                if (token != Token)
                    return Error(HttpStatusCode.Unauthorized, "NRC_UNAUTHORIZED");

                var response = await RouteAsync(request, url);
                if (_server.AnswerFromAnotherHost)
                    response.RequestMessage = new HttpRequestMessage(request.Method, "https://elsewhere.fake/");
                return response;
            }

            private async Task<HttpResponseMessage> RouteAsync(HttpRequestMessage request, string url)
            {
                if (!url.StartsWith(Api, StringComparison.Ordinal))
                    return Status(HttpStatusCode.NotFound);
                var segments = url.Substring(Api.Length).Split('?')[0].Split('/');
                var collection = segments[0];
                var id = Uri.UnescapeDataString(segments[1]);
                var action = segments.Length > 2 ? segments[2] : string.Empty;

                if (collection == "folders")
                {
                    var known = id == "ACTIVE!773" ? !_server.MissingFolder : _server._subfolders.ContainsKey(id);
                    if (!known || (id != "ACTIVE!773" && request.Method != HttpMethod.Get))
                        return Error(HttpStatusCode.NotFound, "NRC_NO_RECORD");
                    if (request.Method == HttpMethod.Get && action == "children")
                    {
                        if (_server.ListStatus is { } listStatus) return Error(listStatus, "NRC_ACCESS_DENIED");
                        var numbers = id == "ACTIVE!773" ? _server._folder : _server._filed.TryGetValue(id, out var filed) ? filed : new List<long>();
                        var children = numbers.Select(n => _server._docs[n])
                            .Select(d => (object)_server.ProfileOf(d, d.Latest))
                            .Concat(_server._subfolders.Where(pair => pair.Value.Parent == id).Select(pair => (object)new Dictionary<string, object?>
                            {
                                ["id"] = pair.Key, ["name"] = pair.Value.Name, ["extension"] = "docx", ["wstype"] = "folder"
                            }))
                            .ToList();
                        // limit and offset page the children when the caller asks for them.
                        var query = url.Contains('?') ? url.Substring(url.IndexOf('?') + 1).Split('&')
                            .Select(pair => pair.Split('=', 2)).ToDictionary(pair => pair[0], pair => pair.Length > 1 ? pair[1] : string.Empty)
                            : new Dictionary<string, string>();
                        if (query.TryGetValue("limit", out var limit))
                            children = children.Skip(query.TryGetValue("offset", out var offset) ? int.Parse(offset) : 0)
                                .Take(int.Parse(limit)).ToList();
                        return Json(new { data = children });
                    }
                    if (request.Method == HttpMethod.Post && action == "documents")
                    {
                        if (TakeUploadFailure() is { } refused) return refused;
                        var (profile, bytes) = await ReadMultipartAsync(request);
                        var number = _server._nextNumber++;
                        var doc = new Doc { Number = number };
                        doc.Versions.Add(_server.NewVersion(1, profile.GetProperty("name").GetString()!,
                            profile.GetProperty("extension").GetString()!, bytes,
                            profile.TryGetProperty("comment", out var c) ? c.GetString() : null,
                            profile.TryGetProperty("class", out var k) ? k.GetString() : null));
                        _server._docs[number] = doc;
                        _server._folder.Add(number);
                        _server.Uploads++;
                        if (_server.StoreThenLoseResponse)
                        {
                            _server.StoreThenLoseResponse = false;
                            throw new TaskCanceledException("The request timed out after the upload was stored.");
                        }
                        return Json(new { data = _server.ProfileOf(doc, doc.Latest) });
                    }
                    return Status(HttpStatusCode.MethodNotAllowed);
                }

                if (collection != "documents" || !TryParse(id, out var docNumber, out var versionNumber) ||
                    !_server._docs.TryGetValue(docNumber, out var document) ||
                    document.Versions.All(v => v.Version != versionNumber))
                    return Error(HttpStatusCode.NotFound, "NRC_NO_RECORD", _server.ErrorMessage);
                var version = document.Versions.Single(v => v.Version == versionNumber);

                switch (action)
                {
                    case "versions" when request.Method == HttpMethod.Get:
                        return Json(new { data = document.Versions.Select(v => _server.ProfileOf(document, v)).ToList() });

                    case "download" when request.Method == HttpMethod.Get:
                        _server.Downloads++;
                        if (_server.RedirectDownloads)
                        {
                            var redirect = new HttpResponseMessage(HttpStatusCode.Found);
                            redirect.Headers.Location = new Uri($"{Storage}{document.Number}/{version.Version}");
                            return redirect;
                        }
                        return Bytes(version.Bytes);

                    case "lock" when request.Method == HttpMethod.Post:
                        if (_server.LockStatus is { } lockStatus)
                            return Error(lockStatus, _server.LockErrorCode);
                        if (document.CheckedOut || document.Locked)
                            return Error((HttpStatusCode)423, "DOCUMENT_LOCKED");
                        document.Locked = true;
                        _server.Locks++;
                        _server.OnLock?.Invoke();
                        return Json(new { data = new { } });

                    case "lock" when request.Method == HttpMethod.Delete:
                        if (!document.Locked) return Error(HttpStatusCode.BadRequest, "NOT_LOCKED");
                        document.Locked = false;
                        return Json(new { data = new { } });

                    case "versions" when request.Method == HttpMethod.Post:
                    {
                        if (document.CheckedOut)
                            return Error((HttpStatusCode)423, "DOCUMENT_CHECKED_OUT");
                        if (TakeUploadFailure() is { } refused) return refused;
                        var (profile, bytes) = await ReadMultipartAsync(request);
                        var latest = document.Latest;
                        var added = _server.NewVersion(latest.Version + 1, profile.GetProperty("name").GetString()!,
                            profile.GetProperty("extension").GetString()!, bytes,
                            profile.TryGetProperty("comment", out var c) ? c.GetString() : null, null);
                        document.Versions.Add(added);
                        _server.Uploads++;
                        return Json(new { data = _server.ProfileOf(document, added) });
                    }
                }
                return Status(HttpStatusCode.MethodNotAllowed);
            }

            private HttpResponseMessage? TakeUploadFailure()
            {
                if (_server.UploadStatus is not { } status) return null;
                _server.UploadStatus = null;
                return Error(status, "simulated");
            }

            private static bool TryParse(string id, out long number, out int version)
            {
                number = 0;
                version = 0;
                var bang = id.IndexOf('!');
                var dot = id.LastIndexOf('.');
                return id.StartsWith("ACTIVE!", StringComparison.Ordinal) && dot > bang &&
                       long.TryParse(id.Substring(bang + 1, dot - bang - 1), out number) &&
                       int.TryParse(id.Substring(dot + 1), out version);
            }

            private static async Task<(JsonElement Profile, byte[] Bytes)> ReadMultipartAsync(HttpRequestMessage request)
            {
                var multipart = Assert.IsType<MultipartFormDataContent>(request.Content);
                JsonElement? profile = null;
                byte[]? bytes = null;
                foreach (var part in multipart)
                {
                    var name = part.Headers.ContentDisposition!.Name!.Trim('"');
                    if (name == "profile")
                    {
                        using var json = JsonDocument.Parse(await part.ReadAsStringAsync());
                        Assert.True(json.RootElement.GetProperty("warnings_for_required_and_disabled_fields").GetBoolean());
                        profile = json.RootElement.GetProperty("doc_profile").Clone();
                    }
                    else if (name == "file")
                    {
                        Assert.False(string.IsNullOrEmpty(part.Headers.ContentDisposition.FileName));
                        bytes = await part.ReadAsByteArrayAsync();
                    }
                }
                Assert.Equal(bytes!.LongLength, profile!.Value.GetProperty("size").GetInt64());
                return (profile.Value, bytes);
            }

            private static HttpResponseMessage Json(object payload) => new(HttpStatusCode.OK)
            {
                Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
            };

            private static HttpResponseMessage Bytes(byte[] bytes) => new(HttpStatusCode.OK)
            {
                Content = new ByteArrayContent(bytes)
            };

            private static HttpResponseMessage Status(HttpStatusCode code) => new(code);

            private static HttpResponseMessage Error(HttpStatusCode status, string code, string? message = null) => new(status)
            {
                Content = new StringContent(
                    JsonSerializer.Serialize(new { error = new { code, message = message ?? code } }),
                    Encoding.UTF8, "application/json")
            };
        }
    }
}
