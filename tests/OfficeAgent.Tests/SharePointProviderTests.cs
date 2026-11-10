using System.Net;
using System.Text;
using System.Text.Json;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.SharePoint;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// SharePoint provider tests over a fake Graph drive: drive items simulated by an
/// in-memory <see cref="HttpMessageHandler"/>, so the full register → open → save →
/// remove loop (including registration by URL and by driveId/itemId, ETag
/// concurrency, and the no-credentials-to-storage rule) runs without a tenant.
/// </summary>
public class SharePointProviderTests
{
    [Fact]
    public async Task Create_uploads_registers_and_applies_an_initial_plan()
    {
        using var drive = new FakeGraphDrive();
        var provider = drive.Provider();
        var client = new OfficeAgentClient(
            new DocumentProviderRegistry(new IDocumentProvider[] { provider }),
            new WordModule());
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertOp
                {
                    Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = string.Empty },
                    Position = InsertPosition.Before,
                    Text = "SharePoint brief"
                }
            }
        };

        var created = await client.CreateAsync("legal", "brief.docx", plan);

        Assert.True(created.Committed);
        Assert.Equal("sharepoint", created.Document!.Provider);
        Assert.True(drive.HasName("brief.docx"));
        var inspected = await client.InspectAsync(created.Document);
        Assert.Equal("SharePoint brief", inspected.Paragraphs[0].Text);
    }

    [Fact]
    public async Task Create_never_overwrites_an_existing_sharepoint_name()
    {
        using var drive = new FakeGraphDrive();
        var provider = drive.Provider();
        using var first = new MemoryStream(DocxFactory.Contract());
        using var second = new MemoryStream(DocxFactory.Contract());

        await provider.CreateAsync("brief.docx", first);
        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.CreateAsync("brief.docx", second));

        Assert.Equal(ProviderErrorCode.AlreadyExists, error.Code);
    }

    [Theory]
    [InlineData("CON.docx")]
    [InlineData("nul.docx")]
    [InlineData("LPT1.docx")]
    public async Task Create_rejects_reserved_windows_device_names(string name)
    {
        using var drive = new FakeGraphDrive();
        var provider = drive.Provider();
        using var content = new MemoryStream(DocxFactory.Contract());

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.CreateAsync(name, content));

        Assert.Equal(ProviderErrorCode.InvalidArgument, error.Code);
        Assert.Contains("reserved Windows device name", error.Message, StringComparison.OrdinalIgnoreCase);
        Assert.False(drive.HasName(name));
    }

    [Fact]
    public async Task Create_requires_a_configured_sharepoint_destination()
    {
        using var drive = new FakeGraphDrive();
        var provider = drive.Provider(configureCreation: false);
        using var content = new MemoryStream(DocxFactory.Contract());

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.CreateAsync("brief.docx", content));

        Assert.Equal(ProviderErrorCode.ConfigurationError, error.Code);
        Assert.False(drive.HasName("brief.docx"));
    }

    [Fact]
    public async Task Registration_failure_leaves_the_uploaded_sharepoint_file_intact()
    {
        using var drive = new FakeGraphDrive();
        var provider = drive.Provider(registrationStore: new FailingRegistrationStore());
        using var content = new MemoryStream(DocxFactory.Contract());

        var error = await Assert.ThrowsAsync<DocumentRegistrationFailedException>(() =>
            provider.CreateAsync("brief.docx", content));

        Assert.Equal(ProviderErrorCode.RegistrationFailed, error.Code);
        Assert.Contains("unregistered", error.Message);
        Assert.True(drive.HasName("brief.docx"));
        AssertLocated(drive, error, "brief.docx", sourceItemId: null);
    }

    [Fact]
    public async Task Create_tool_does_not_expose_sharepoint_target_details()
    {
        using var drive = new FakeGraphDrive
        {
            CreationFailureMessage =
                $"Target {FakeGraphDrive.DriveId}/{FakeGraphDrive.FolderId} was not found."
        };
        var provider = drive.Provider();
        var client = new OfficeAgentClient(
            new DocumentProviderRegistry(new IDocumentProvider[] { provider }),
            new WordModule());
        var tools = new OfficeAgentTools(client);

        var payload = await tools.CreateDocument("legal", "brief.docx");

        Assert.DoesNotContain(FakeGraphDrive.DriveId, payload);
        Assert.DoesNotContain(FakeGraphDrive.FolderId, payload);
        using var json = JsonDocument.Parse(payload);
        Assert.Equal("not-found",
            json.RootElement.GetProperty("errors")[0].GetProperty("code").GetString());
    }

    [Fact]
    public async Task Graph_error_message_is_not_exposed_by_registration()
    {
        const string upstreamDetail = "Resource https://tenant.example/sites/legal/folder-secret was denied.";
        using var drive = new FakeGraphDrive { ResolutionFailureMessage = upstreamDetail };
        var provider = drive.Provider();

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.RegisterAsync(drive.UrlFor("missing-item")));

        Assert.Equal(ProviderErrorCode.NotFound, error.Code);
        Assert.Contains("itemNotFound", error.Message);
        Assert.DoesNotContain(upstreamDetail, error.Message);
        Assert.DoesNotContain("folder-secret", error.Message);
    }

    [Fact]
    public async Task Register_by_drive_and_item_id_assigns_opaque_id_and_opens_by_it()
    {
        using var drive = new FakeGraphDrive();
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();

        var added = await provider.RegisterAsync(drive.SourceById(itemId));

        Assert.Equal("sharepoint", added.Provider);
        Assert.NotEqual(itemId, added.ItemId);
        Assert.DoesNotContain('/', added.ItemId);
        Assert.Equal("contract.docx", added.Name);
        Assert.False(string.IsNullOrEmpty(added.Version));

        using var content = await provider.OpenReadAsync(added);
        Assert.Equal(added.ItemId, content.Reference.ItemId);
        Assert.True(content.Stream.Length > 0);
    }

    [Fact]
    public async Task Register_by_url_resolves_the_drive_item()
    {
        using var drive = new FakeGraphDrive();
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();

        var added = await provider.RegisterAsync(drive.UrlFor(itemId));

        Assert.Equal("contract.docx", added.Name);
        using var content = await provider.OpenReadAsync(added);
        Assert.True(content.Stream.Length > 0);
    }

    [Fact]
    public async Task Register_accepts_the_graph_resource_path_form()
    {
        using var drive = new FakeGraphDrive();
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();

        var added = await provider.RegisterAsync($"drives/{FakeGraphDrive.DriveId}/items/{itemId}");

        Assert.Equal("contract.docx", added.Name);
    }

    [Fact]
    public async Task Register_rejects_bare_names_and_disallowed_extensions()
    {
        using var drive = new FakeGraphDrive();
        drive.Seed("contract.docx", DocxFactory.Contract());
        var exeId = drive.Seed("contract.exe", DocxFactory.Contract());
        var provider = drive.Provider();

        // A bare name with no driveId/itemId delimiter is not a valid source.
        var bare = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.RegisterAsync("contract.docx"));
        Assert.Equal(ProviderErrorCode.InvalidArgument, bare.Code);

        // An empty item id half is rejected too.
        var empty = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.RegisterAsync($"{FakeGraphDrive.DriveId}/"));
        Assert.Equal(ProviderErrorCode.InvalidArgument, empty.Code);

        // The extension is enforced against the resolved item name.
        var extension = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.RegisterAsync(drive.SourceById(exeId)));
        Assert.Equal(ProviderErrorCode.ExtensionNotAllowed, extension.Code);
    }

    [Fact]
    public async Task Register_surfaces_not_found_for_missing_items()
    {
        using var drive = new FakeGraphDrive();
        var provider = drive.Provider();

        var ex = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.RegisterAsync($"{FakeGraphDrive.DriveId}/item-missing"));
        Assert.Equal(ProviderErrorCode.NotFound, ex.Code);
    }

    [Fact]
    public async Task Save_new_version_mints_sibling_with_versioned_name_and_fresh_id()
    {
        using var drive = new FakeGraphDrive();
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();
        var source = await provider.RegisterAsync(drive.SourceById(itemId));

        using var output = new MemoryStream(DocxFactory.Contract());
        var saved = await provider.SaveAsync(
            source, output, new SaveDocumentOptions { Mode = SaveMode.NewVersion });

        Assert.Equal("contract.v2.docx", saved.Name);
        Assert.NotEqual(source.ItemId, saved.ItemId);

        using var roundTrip = await provider.OpenReadAsync(saved);
        Assert.Equal("contract.v2.docx", roundTrip.Reference.Name);
    }

    [Fact]
    public async Task Save_replace_overwrites_in_place_and_rejects_stale_etag()
    {
        using var drive = new FakeGraphDrive();
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();
        var source = await provider.RegisterAsync(drive.SourceById(itemId));

        using (var fresh = new MemoryStream(DocxFactory.Contract().Concat(new byte[] { 0 }).ToArray()))
        {
            var replaced = await provider.SaveAsync(source, fresh, new SaveDocumentOptions { Mode = SaveMode.Replace });
            Assert.Equal(source.ItemId, replaced.ItemId);
            Assert.NotEqual(source.Version, replaced.Version);
        }

        // The reference still carries the original (now stale) ETag.
        using var output = new MemoryStream(DocxFactory.Contract());
        await Assert.ThrowsAsync<DocumentVersionConflictException>(() =>
            provider.SaveAsync(source, output, new SaveDocumentOptions { Mode = SaveMode.Replace }));
    }

    [Fact]
    public async Task Save_rejects_caller_chosen_destinations_and_replace_renames()
    {
        using var drive = new FakeGraphDrive();
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();
        var source = await provider.RegisterAsync(drive.SourceById(itemId));

        using var output = new MemoryStream(DocxFactory.Contract());
        var destination = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(
            source, output, new SaveDocumentOptions { DestinationItemId = "chosen-id" }));
        Assert.Equal(ProviderErrorCode.InvalidArgument, destination.Code);

        var rename = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(
            source, output, new SaveDocumentOptions { Mode = SaveMode.Replace, NewName = "renamed.docx" }));
        Assert.Equal(ProviderErrorCode.InvalidArgument, rename.Code);
    }

    [Fact]
    public async Task Save_new_document_honours_validated_new_name()
    {
        using var drive = new FakeGraphDrive();
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();
        var source = await provider.RegisterAsync(drive.SourceById(itemId));

        using var output = new MemoryStream(DocxFactory.Contract());
        var saved = await provider.SaveAsync(source, output, new SaveDocumentOptions
        {
            Mode = SaveMode.NewDocument,
            NewName = "globex-contract.docx"
        });
        Assert.Equal("globex-contract.docx", saved.Name);

        using var bad = new MemoryStream(DocxFactory.Contract());
        var ex = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(
            source, bad, new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "nested/name.docx" }));
        Assert.Equal(ProviderErrorCode.InvalidArgument, ex.Code);
    }

    [Fact]
    public async Task Remove_drops_registration_but_not_sharepoint_content()
    {
        using var drive = new FakeGraphDrive();
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();
        var added = await provider.RegisterAsync(drive.SourceById(itemId));

        await provider.RemoveAsync(added);

        var ex = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.OpenReadAsync(added));
        Assert.Equal(ProviderErrorCode.NotFound, ex.Code);
        // The drive item survives: a new registration of the same item still works.
        var again = await provider.RegisterAsync(drive.SourceById(itemId));
        Assert.NotEqual(added.ItemId, again.ItemId);
    }

    [Fact]
    public async Task Open_rejects_oversized_content()
    {
        using var drive = new FakeGraphDrive();
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider(maximumBytes: 16);

        var ex = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.RegisterAsync(drive.SourceById(itemId)));
        Assert.Equal(ProviderErrorCode.ContentTooLarge, ex.Code);
    }

    [Fact]
    public async Task Bearer_token_never_travels_to_the_download_host()
    {
        using var drive = new FakeGraphDrive();
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();
        var added = await provider.RegisterAsync(drive.SourceById(itemId));

        using var _ = await provider.OpenReadAsync(added);

        Assert.True(drive.SawAuthorizedGraphCall);
        Assert.False(drive.SawAuthorizedDownloadCall);
    }

    [Fact]
    public async Task Json_file_registration_store_survives_reload()
    {
        var path = Path.Combine(Path.GetTempPath(), $"officeagent-spindex-{Guid.NewGuid():N}", "index.json");
        try
        {
            var store = new JsonFileRegistrationStore(path);
            var id = await store.AddAsync(new SharePointItemRef("drive-1", "item-1"));

            var reloaded = new JsonFileRegistrationStore(path);
            Assert.Equal(new SharePointItemRef("drive-1", "item-1"), await reloaded.ResolveAsync(id));

            Assert.True(await reloaded.RemoveAsync(id));
            var reloadedAgain = new JsonFileRegistrationStore(path);
            Assert.Null(await reloadedAgain.ResolveAsync(id));
        }
        finally
        {
            var directory = Path.GetDirectoryName(path)!;
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public async Task Json_registration_store_rolls_back_an_add_when_persistence_fails()
    {
        var directory = Path.Combine(
            Path.GetTempPath(), $"officeagent-spindex-failure-{Guid.NewGuid():N}");
        var path = Path.Combine(directory, "index.json");
        Directory.CreateDirectory(path);
        try
        {
            var store = new JsonFileRegistrationStore(path);
            await Assert.ThrowsAsync<IOException>(() =>
                store.AddAsync(new SharePointItemRef("drive-ghost", "item-ghost")));

            Directory.Delete(path);
            await store.AddAsync(new SharePointItemRef("drive-real", "item-real"));

            using var json = JsonDocument.Parse(File.ReadAllText(path));
            var items = json.RootElement.GetProperty("Items");
            Assert.Single(items.EnumerateObject());
            Assert.Equal("drive-real",
                items.EnumerateObject().Single().Value.GetProperty("DriveId").GetString());
        }
        finally
        {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    // ── Fake Graph drive ──────────────────────────────────────────────────────

    /// <summary>
    /// Simulates one Graph drive (`/drives/{id}`) addressed by item id, plus the
    /// `/shares/{id}/driveItem` URL-resolution endpoint, with ETags,
    /// conflictBehavior=fail uploads, and pre-authenticated download URLs on a
    /// separate host. Items report their parent drive id, so registration by URL or
    /// by driveId/itemId both resolve a fully-addressed item.
    /// </summary>
    // ── Storage outcome simulations (V10-02) ─────────────────────────────────────────────
    // Controlled simulations against the fake Graph drive, not live SharePoint. Each drives a
    // create through the engine, so the outcome asserted is the one a caller receives.

    [Fact]
    public async Task A_read_only_connection_registers_and_opens_but_sends_no_write()
    {
        using var drive = new FakeGraphDrive();
        var original = DocxFactory.Contract();
        var itemId = drive.Seed("Invoice.docx", original);
        var provider = drive.Provider(configureCreation: false, readOnly: true);

        Assert.True(provider.ReadOnly);
        var reference = await provider.RegisterAsync(drive.SourceById(itemId));
        using (var content = await provider.OpenReadAsync(reference))
            Assert.True(content.Stream.Length > 0);

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
        Assert.Equal(0, drive.WriteRequests);
        Assert.Equal(original, drive.BytesNamed("Invoice.docx"));

        // Registrations live in the host's store, not in SharePoint, so removing one still works.
        await provider.RemoveAsync(reference);
    }

    [Fact]
    public void A_read_only_connection_cannot_also_name_a_creation_folder()
    {
        using var drive = new FakeGraphDrive();

        var error = Assert.Throws<ArgumentException>(() => drive.Provider(configureCreation: true, readOnly: true));

        Assert.Contains("CreationDriveId", error.Message, StringComparison.Ordinal);
    }

    [Fact]
    public async Task A_template_on_a_read_only_connection_seeds_a_document_on_a_writable_one()
    {
        using var drive = new FakeGraphDrive();
        var original = DocxFactory.Contract();
        var templateId = drive.Seed("Invoice.docx", original);
        var templates = drive.Provider(configureCreation: false, readOnly: true, connectionId: "templates");
        var documents = drive.Provider(connectionId: "legal");
        var client = new OfficeAgentClient(
            new DocumentProviderRegistry(new IDocumentProvider[] { templates, documents }), new WordModule());
        var template = await templates.RegisterAsync(drive.SourceById(templateId));

        var created = await client.CreateFromTemplateAsync("legal", "Invoice 2026-11.docx", template, plan: null);

        Assert.True(created.Committed);
        Assert.True(drive.HasName("Invoice 2026-11.docx"));
        Assert.Equal(original, drive.BytesNamed("Invoice.docx"));
    }

    [Fact]
    public async Task Editing_a_template_on_a_read_only_connection_writes_nothing()
    {
        using var drive = new FakeGraphDrive();
        var itemId = drive.Seed("Invoice.docx", DocxFactory.Contract());
        var provider = drive.Provider(configureCreation: false, readOnly: true);
        var tools = new OfficeAgentTools(EngineOver(provider));

        using var json = JsonDocument.Parse(await tools.EditDocument("legal", drive.SourceById(itemId),
            "[{\"op\":\"changeText\",\"target\":{\"find\":\"Client:\"},\"with\":\"Fabrikam\",\"mode\":\"Direct\"}]"));

        Assert.False(json.RootElement.GetProperty("committed").GetBoolean());
        Assert.Equal("notWritten", json.RootElement.GetProperty("writeOutcome").GetString());
        Assert.Equal("access-denied", json.RootElement.GetProperty("errors")[0].GetProperty("code").GetString());
        Assert.Equal(0, drive.WriteRequests);
    }

    private static OfficeAgentClient EngineOver(SharePointDocumentProvider provider) =>
        new(new DocumentProviderRegistry(new IDocumentProvider[] { provider }), new WordModule());

    [Theory]
    [InlineData(HttpStatusCode.BadRequest)]
    [InlineData((HttpStatusCode)423)] // Locked
    [InlineData(HttpStatusCode.TooManyRequests)]
    public async Task A_refused_upload_is_certain_and_leaves_nothing_behind(HttpStatusCode status)
    {
        using var drive = new FakeGraphDrive { UploadStatus = status };
        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            EngineOver(drive.Provider()).CreateAsync("legal", "refused.docx"));

        Assert.Equal(ProviderErrorCode.WriteRejected, error.Code);
        Assert.False(drive.HasName("refused.docx"));
    }

    [Theory]
    [InlineData((HttpStatusCode)423)]
    [InlineData(HttpStatusCode.BadRequest)] // Graph can report the lock in its code alone
    public async Task A_locked_file_says_why_the_save_was_refused_and_what_to_do(HttpStatusCode status)
    {
        using var drive = new FakeGraphDrive();
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();
        var source = await provider.RegisterAsync(drive.SourceById(itemId));

        drive.UploadStatus = status;
        drive.UploadErrorCode = "resourceLocked";
        var error = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(
            source, new MemoryStream(DocxFactory.Contract()), new SaveDocumentOptions { Mode = SaveMode.Replace }));

        Assert.Equal(ProviderErrorCode.WriteRejected, error.Code);
        Assert.Contains("locked in SharePoint", error.Message);
        Assert.Contains("check the file in", error.Message);
        Assert.Contains("save as a new document", error.Message);
        Assert.Contains("Graph code: resourceLocked", error.Message);
    }

    [Fact]
    public async Task A_server_error_on_upload_is_an_unknown_outcome()
    {
        // A 503 is not a refusal: Graph may have processed the request before failing.
        using var drive = new FakeGraphDrive { UploadStatus = HttpStatusCode.ServiceUnavailable };
        var error = await Assert.ThrowsAsync<DocumentWriteOutcomeUnknownException>(() =>
            EngineOver(drive.Provider()).CreateAsync("legal", "maybe.docx"));

        Assert.Equal(ProviderErrorCode.OutcomeUnknown, error.Code);
    }

    [Fact]
    public async Task A_lost_response_after_a_stored_upload_is_unknown_and_reconcilable()
    {
        using var drive = new FakeGraphDrive { StoreThenLoseResponse = true };
        var error = await Assert.ThrowsAsync<DocumentWriteOutcomeUnknownException>(() =>
            EngineOver(drive.Provider()).CreateAsync("legal", "lost.docx"));

        // The upload landed although the caller was never told. The exception says so, and
        // its hash is exactly what the destination now holds, which is how a caller reconciles.
        Assert.Equal(ProviderErrorCode.OutcomeUnknown, error.Code);
        Assert.Equal("lost.docx", error.OutputName);
        Assert.True(drive.HasName("lost.docx"));
        Assert.Equal(Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(drive.BytesNamed("lost.docx"))),
            error.OutputSha256, ignoreCase: true);
    }

    [Fact]
    public async Task A_taken_name_on_a_new_version_is_refused_before_upload()
    {
        using var drive = new FakeGraphDrive();
        var provider = drive.Provider();
        var source = drive.Seed("contract.docx", DocxFactory.Contract());
        drive.Seed("taken.docx", DocxFactory.Contract());
        var reference = await provider.RegisterAsync(drive.SourceById(source));

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.SaveAsync(reference, new MemoryStream(DocxFactory.Contract()),
                new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "taken.docx" }));

        // Was reported as IO, which a template batch then called uncertain.
        Assert.Equal(ProviderErrorCode.AlreadyExists, error.Code);
    }

    [Fact]
    public async Task A_new_version_whose_registration_fails_is_written_not_registered()
    {
        using var drive = new FakeGraphDrive();
        var provider = drive.Provider(registrationStore: new RegisterOnceStore());
        var source = drive.Seed("contract.docx", DocxFactory.Contract());
        var reference = await provider.RegisterAsync(drive.SourceById(source));

        var error = await Assert.ThrowsAsync<DocumentRegistrationFailedException>(() =>
            provider.SaveAsync(reference, new MemoryStream(DocxFactory.Contract()),
                new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "copy.docx" }));

        Assert.Equal(ProviderErrorCode.RegistrationFailed, error.Code);
        Assert.True(drive.HasName("copy.docx"));
        AssertLocated(drive, error, "copy.docx", reference.ItemId);
    }

    /// <summary>
    /// Through the tools, the SharePoint outcome carries the same locator, and nothing names the
    /// drive, folder or Graph item the upload landed in.
    /// </summary>
    [Fact]
    public async Task A_sharepoint_registration_failure_reports_a_path_free_locator_through_the_tools()
    {
        using var drive = new FakeGraphDrive();
        var provider = drive.Provider(registrationStore: new RegisterOnceStore());
        var source = drive.Seed("contract.docx", DocxFactory.Contract());
        var reference = await provider.RegisterAsync(drive.SourceById(source));
        var tools = new OfficeAgentTools(new OfficeAgentClient(new DocumentProviderRegistry(new IDocumentProvider[] { provider }), new WordModule()));
        using var hits = JsonDocument.Parse(await tools.FindInDocument("legal", reference.ItemId, "Acme Corp"));
        var hit = hits.RootElement[0];
        var plan = "{\"operations\":[{\"op\":\"changeText\",\"target\":{\"paraId\":\"" + hit.GetProperty("paraId").GetString() +
                   "\",\"expect\":\"Acme Corp\",\"occurrence\":" + hit.GetProperty("occurrence").GetInt32() +
                   "},\"mode\":\"Direct\",\"with\":\"Contoso\"}]}";

        using var result = JsonDocument.Parse(await tools.ApplyPlan("legal", reference.ItemId, plan, saveMode: "NewDocument", newName: "copy.docx"));

        Assert.Equal("registration-failed", result.RootElement.GetProperty("errors")[0].GetProperty("code").GetString());
        Assert.Equal("writtenNotRegistered", result.RootElement.GetProperty("writeOutcome").GetString());
        var locator = result.RootElement.GetProperty("possibleOutput");
        Assert.Equal("legal", locator.GetProperty("connectionId").GetString());
        Assert.Equal(reference.ItemId, locator.GetProperty("sourceDocumentId").GetString());
        Assert.Equal("copy.docx", locator.GetProperty("outputName").GetString());
        Assert.Equal(Sha256(drive.BytesNamed("copy.docx")), locator.GetProperty("expectedSha256").GetString());
        var json = result.RootElement.ToString();
        Assert.DoesNotContain(FakeGraphDrive.DriveId, json);
        Assert.DoesNotContain("graph.fake", json);
    }

    // ── Upload sessions (#57) ────────────────────────────────────────────────────────────
    // Saves at or above the threshold go through a Graph upload session. The fake keeps each
    // session's bytes and stores the file only when the last byte arrives, as Graph does.

    private const int ChunkUnit = 320 * 1024;

    private static byte[] LargeContent()
    {
        // Three full 320 KiB chunks and a short last one.
        var bytes = new byte[3 * ChunkUnit + 16_960];
        new Random(57).NextBytes(bytes);
        return bytes;
    }

    private static SharePointDocumentProvider SessionProvider(FakeGraphDrive drive) =>
        drive.Provider(uploadSessionThresholdBytes: 512 * 1024, uploadSessionChunkBytes: ChunkUnit);

    [Fact]
    public async Task A_large_replace_uploads_in_320_KiB_chunks_without_the_bearer_token()
    {
        using var drive = new FakeGraphDrive();
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = SessionProvider(drive);
        var source = await provider.RegisterAsync(drive.SourceById(itemId));
        var content = LargeContent();

        var replaced = await provider.SaveAsync(source, new MemoryStream(content), new SaveDocumentOptions { Mode = SaveMode.Replace });

        Assert.Equal(source.ItemId, replaced.ItemId);
        Assert.NotEqual(source.Version, replaced.Version);
        Assert.Equal(content, drive.BytesNamed("contract.docx"));
        Assert.Equal(1, drive.SessionsCreated);
        Assert.Equal(new[] { ChunkUnit, ChunkUnit, ChunkUnit, 16_960 }, drive.ChunkSizes);
        Assert.All(drive.ChunkSizes.Take(drive.ChunkSizes.Count - 1), size => Assert.Equal(0, size % ChunkUnit));
        // Replace keeps its precondition: the session request carries If-Match with the ETag read.
        Assert.Equal(source.Version, drive.SessionIfMatch);
        Assert.Equal("replace", drive.SessionConflictBehavior);
        // The token goes to Graph, never to the pre-authenticated upload URL.
        Assert.True(drive.SawAuthorizedGraphCall);
        Assert.False(drive.SawAuthorizedUploadCall);
    }

    [Fact]
    public async Task A_save_below_the_threshold_is_still_one_request()
    {
        using var drive = new FakeGraphDrive();
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = drive.Provider();
        var source = await provider.RegisterAsync(drive.SourceById(itemId));

        await provider.SaveAsync(source, new MemoryStream(DocxFactory.Contract()), new SaveDocumentOptions { Mode = SaveMode.NewVersion });

        Assert.True(drive.HasName("contract.v2.docx"));
        Assert.Equal(0, drive.SessionsCreated);
        Assert.Empty(drive.ChunkSizes);
    }

    [Theory]
    [InlineData(ChunkFault.DropBeforeStore, 5)]   // the chunk never landed: it is sent again
    [InlineData(ChunkFault.ServerError, 5)]       // a 503: sent again after Retry-After
    [InlineData(ChunkFault.StoreThenDrop, 4)]     // it landed but the reply was lost: resume after it
    public async Task A_failed_chunk_is_retried_from_the_range_the_session_expects(ChunkFault fault, int puts)
    {
        using var drive = new FakeGraphDrive();
        drive.ChunkFaults[1] = fault;
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = SessionProvider(drive);
        var source = await provider.RegisterAsync(drive.SourceById(itemId));
        var content = LargeContent();

        var saved = await provider.SaveAsync(source, new MemoryStream(content),
            new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "large.docx" });

        Assert.Equal("large.docx", saved.Name);
        Assert.Equal(content, drive.BytesNamed("large.docx"));
        Assert.Equal(puts, drive.ChunkSizes.Count);
        Assert.Equal("fail", drive.SessionConflictBehavior);
        Assert.Equal(0, drive.SessionsCancelled);
    }

    [Fact]
    public async Task A_chunk_that_keeps_failing_is_rejected_and_its_session_cancelled()
    {
        using var drive = new FakeGraphDrive();
        for (var put = 1; put <= 4; put++) drive.ChunkFaults[put] = ChunkFault.ServerError;
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = SessionProvider(drive);
        var source = await provider.RegisterAsync(drive.SourceById(itemId));

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(source,
            new MemoryStream(LargeContent()), new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "large.docx" }));

        Assert.Equal(ProviderErrorCode.WriteRejected, error.Code);
        Assert.Contains("nothing was stored", error.Message);
        Assert.False(drive.HasName("large.docx"));
        Assert.Equal(1, drive.SessionsCancelled);
    }

    [Fact]
    public async Task A_replace_whose_item_changed_is_refused_when_the_session_opens()
    {
        using var drive = new FakeGraphDrive { ChangeItemAtSessionStart = true };
        var original = DocxFactory.Contract();
        var itemId = drive.Seed("contract.docx", original);
        var provider = SessionProvider(drive);
        var source = await provider.RegisterAsync(drive.SourceById(itemId));

        var error = await Assert.ThrowsAsync<DocumentVersionConflictException>(() => provider.SaveAsync(
            source, new MemoryStream(LargeContent()), new SaveDocumentOptions { Mode = SaveMode.Replace }));

        Assert.Equal(ProviderErrorCode.VersionConflict, error.Code);
        Assert.Equal(original, drive.BytesNamed("contract.docx"));
        Assert.Empty(drive.ChunkSizes);
    }

    [Fact]
    public async Task A_replace_whose_item_changed_during_the_session_is_refused_at_the_last_chunk()
    {
        // Graph's documentation does not say whether If-Match is checked again at commit. If it is,
        // the 412 arrives on the last chunk: a refusal, so nothing was stored and the session goes.
        using var drive = new FakeGraphDrive { ChangeItemDuringSession = true };
        var original = DocxFactory.Contract();
        var itemId = drive.Seed("contract.docx", original);
        var provider = SessionProvider(drive);
        var source = await provider.RegisterAsync(drive.SourceById(itemId));

        await Assert.ThrowsAsync<DocumentVersionConflictException>(() => provider.SaveAsync(
            source, new MemoryStream(LargeContent()), new SaveDocumentOptions { Mode = SaveMode.Replace }));

        Assert.Equal(original, drive.BytesNamed("contract.docx"));
        Assert.Equal(1, drive.SessionsCancelled);
    }

    [Fact]
    public async Task A_taken_name_reported_on_the_last_chunk_is_refused_and_its_session_cancelled()
    {
        using var drive = new FakeGraphDrive();
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var taken = DocxFactory.Contract();
        drive.Seed("taken.docx", taken);
        var provider = SessionProvider(drive);
        var source = await provider.RegisterAsync(drive.SourceById(itemId));

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(source,
            new MemoryStream(LargeContent()), new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "taken.docx" }));

        Assert.Equal(ProviderErrorCode.AlreadyExists, error.Code);
        Assert.Equal(taken, drive.BytesNamed("taken.docx"));
        Assert.Equal(1, drive.SessionsCancelled);
    }

    [Fact]
    public async Task A_session_expiring_mid_upload_is_rejected_with_nothing_stored()
    {
        using var drive = new FakeGraphDrive();
        drive.ChunkFaults[2] = ChunkFault.ExpireSession;
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = SessionProvider(drive);
        var source = await provider.RegisterAsync(drive.SourceById(itemId));

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(source,
            new MemoryStream(LargeContent()), new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "large.docx" }));

        Assert.Equal(ProviderErrorCode.WriteRejected, error.Code);
        Assert.Contains("expired", error.Message);
        Assert.False(drive.HasName("large.docx"));
    }

    [Fact]
    public async Task A_lost_response_to_the_last_chunk_is_unknown()
    {
        using var drive = new FakeGraphDrive();
        drive.ChunkFaults[3] = ChunkFault.StoreThenLoseResponse;
        var itemId = drive.Seed("contract.docx", DocxFactory.Contract());
        var provider = SessionProvider(drive);
        var source = await provider.RegisterAsync(drive.SourceById(itemId));
        var content = LargeContent();

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(source,
            new MemoryStream(content), new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "large.docx" }));

        Assert.Equal(ProviderErrorCode.OutcomeUnknown, error.Code);
        Assert.Equal(content, drive.BytesNamed("large.docx"));
    }

    [Fact]
    public async Task A_lost_response_to_the_last_chunk_is_unknown_and_reconcilable_through_the_engine()
    {
        using var drive = new FakeGraphDrive();
        drive.ChunkFaults[0] = ChunkFault.StoreThenLoseResponse;
        var provider = drive.Provider(uploadSessionThresholdBytes: 1, uploadSessionChunkBytes: ChunkUnit);

        var error = await Assert.ThrowsAsync<DocumentWriteOutcomeUnknownException>(() =>
            EngineOver(provider).CreateAsync("legal", "lost.docx"));

        Assert.Equal(1, drive.SessionsCreated);
        Assert.Equal("lost.docx", error.OutputName);
        Assert.Equal(Sha256(drive.BytesNamed("lost.docx")), error.OutputSha256);
    }

    [Theory]
    [InlineData(0L, ChunkUnit)]
    [InlineData(250L * 1000 * 1000 + 1, ChunkUnit)]
    [InlineData(4L * 1024 * 1024, 100_000)]
    [InlineData(4L * 1024 * 1024, 192 * ChunkUnit)] // 60 MiB: Graph wants less
    public void Upload_session_options_are_validated(long threshold, int chunk)
    {
        using var drive = new FakeGraphDrive();
        Assert.Throws<ArgumentOutOfRangeException>(() =>
            drive.Provider(uploadSessionThresholdBytes: threshold, uploadSessionChunkBytes: chunk));
    }

    private static void AssertLocated(FakeGraphDrive drive, DocumentRegistrationFailedException error, string name, string? sourceItemId)
    {
        Assert.IsAssignableFrom<DocumentWriteRecoveryException>(error);
        Assert.Equal("legal", error.ConnectionId);
        Assert.Equal(sourceItemId, error.ItemId);
        Assert.Equal(name, error.OutputName);
        // Graph confirmed the upload: the locator hashes exactly the bytes the drive now holds.
        Assert.Equal(Sha256(drive.BytesNamed(name)), error.OutputSha256);
        Assert.DoesNotContain(FakeGraphDrive.DriveId, error.Message);
        Assert.DoesNotContain("graph.fake", error.Message);
    }

    private static string Sha256(byte[] bytes) =>
        Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(bytes)).ToLowerInvariant();

    /// <summary>Accepts the first registration, then fails every later one.</summary>
    private sealed class RegisterOnceStore : ISharePointRegistrationStore
    {
        private readonly Dictionary<string, SharePointItemRef> _items = new();

        public Task<string> AddAsync(SharePointItemRef item, CancellationToken cancellationToken = default)
        {
            if (_items.Count > 0) throw new IOException("Registration store unavailable.");
            var id = "reg-" + _items.Count;
            _items[id] = item;
            return Task.FromResult(id);
        }

        public Task<SharePointItemRef?> ResolveAsync(string registrationId, CancellationToken cancellationToken = default) =>
            Task.FromResult<SharePointItemRef?>(_items.TryGetValue(registrationId, out var item) ? item : null);

        public Task<bool> RemoveAsync(string registrationId, CancellationToken cancellationToken = default) =>
            Task.FromResult(_items.Remove(registrationId));
    }

    private sealed class FailingRegistrationStore : ISharePointRegistrationStore
    {
        public Task<string> AddAsync(
            SharePointItemRef item,
            CancellationToken cancellationToken = default) =>
            throw new IOException("Registration store unavailable.");

        public Task<SharePointItemRef?> ResolveAsync(
            string registrationId,
            CancellationToken cancellationToken = default) =>
            Task.FromResult<SharePointItemRef?>(null);

        public Task<bool> RemoveAsync(
            string registrationId,
            CancellationToken cancellationToken = default) =>
            Task.FromResult(false);
    }

    /// <summary>A simulated failure of one chunk PUT to an upload session.</summary>
    public enum ChunkFault
    {
        /// <summary>The connection drops before the chunk is stored.</summary>
        DropBeforeStore,

        /// <summary>The chunk is stored, then the connection drops before the reply.</summary>
        StoreThenDrop,

        /// <summary>A 503 with <c>Retry-After: 0</c>; the chunk is not stored.</summary>
        ServerError,

        /// <summary>The session has expired: it is gone and the chunk gets a 404.</summary>
        ExpireSession,

        /// <summary>The chunk is stored (committing the file when it is the last), then the reply is lost.</summary>
        StoreThenLoseResponse
    }

    private sealed class FakeGraphDrive : IDisposable
    {
        private const string GraphBase = "https://graph.fake/v1.0";
        public const string DriveId = "fakedrive";
        private const string DownloadHost = "https://download.fake";
        private const string UploadHost = "https://upload.fake/session/";
        private const string ShareHost = "https://contoso.fake/doc";
        public const string FolderId = "folder-contracts";

        private readonly Dictionary<string, Entry> _byName = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, Entry> _byId = new(StringComparer.Ordinal);
        private readonly HttpClient _http;
        private int _nextId;
        private int _nextETag;

        public FakeGraphDrive() => _http = new HttpClient(new Handler(this));

        public bool SawAuthorizedGraphCall { get; private set; }
        public bool SawAuthorizedDownloadCall { get; private set; }
        public string? CreationFailureMessage { get; init; }
        public string? ResolutionFailureMessage { get; init; }

        /// <summary>Simulation: answer the next upload with this status without storing it.</summary>
        public HttpStatusCode? UploadStatus { get; set; }

        /// <summary>The Graph error code sent with <see cref="UploadStatus"/>.</summary>
        public string UploadErrorCode { get; set; } = "simulated";

        /// <summary>
        /// Simulation: store the next upload, then lose the response the way a timeout or a
        /// dropped connection does. The item exists; the caller never hears so.
        /// </summary>
        public bool StoreThenLoseResponse { get; set; }

        private readonly Dictionary<string, Session> _sessions = new(StringComparer.Ordinal);
        private int _nextSession;
        private int _chunkPuts;

        /// <summary>Simulated failures of chunk PUTs, keyed by the PUT's 0-based position across the drive.</summary>
        public Dictionary<int, ChunkFault> ChunkFaults { get; } = new();

        /// <summary>Simulation: the item changes after its ETag was read and before its session opens.</summary>
        public bool ChangeItemAtSessionStart { get; init; }

        /// <summary>Simulation: the item changes after the session's first chunk.</summary>
        public bool ChangeItemDuringSession { get; init; }

        public int SessionsCreated { get; private set; }
        public int SessionsCancelled { get; private set; }
        public string? SessionIfMatch { get; private set; }
        public string? SessionConflictBehavior { get; private set; }
        public bool SawAuthorizedUploadCall { get; private set; }

        /// <summary>The size of every chunk PUT received, retries included, in order.</summary>
        public List<int> ChunkSizes { get; } = new();

        /// <summary>The bytes stored under a name, for reconciliation assertions.</summary>
        public byte[] BytesNamed(string name) => _byName[name].Bytes;

        /// <summary>Seeds a drive item and returns its Graph item id.</summary>
        public string Seed(string name, byte[] bytes)
        {
            var entry = new Entry
            {
                Id = $"item-{++_nextId}",
                Name = name,
                Bytes = bytes,
                ETag = NextETag()
            };
            _byName[name] = entry;
            _byId[entry.Id] = entry;
            return entry.Id;
        }

        /// <summary>A driveId/itemId registration source for a seeded item.</summary>
        public string SourceById(string itemId) => $"{DriveId}/{itemId}";

        /// <summary>A sharing-style URL registration source for a seeded item.</summary>
        public string UrlFor(string itemId) => $"{ShareHost}/{itemId}";

        public bool HasName(string name) => _byName.ContainsKey(name);

        /// <summary>Every request that is not a GET, to Graph or to an upload session.</summary>
        public int WriteRequests { get; private set; }

        public SharePointDocumentProvider Provider(
            long maximumBytes = 100 * 1024 * 1024,
            bool configureCreation = true,
            ISharePointRegistrationStore? registrationStore = null,
            long uploadSessionThresholdBytes = 4 * 1024 * 1024,
            int uploadSessionChunkBytes = 10 * 1024 * 1024,
            bool readOnly = false,
            string connectionId = "legal") => new(
            new SharePointDocumentProviderOptions
            {
                ConnectionId = connectionId,
                ReadOnly = readOnly,
                GraphBaseUrl = GraphBase,
                MaximumBytes = maximumBytes,
                UploadSessionThresholdBytes = uploadSessionThresholdBytes,
                UploadSessionChunkBytes = uploadSessionChunkBytes,
                CreationDriveId = configureCreation ? DriveId : string.Empty,
                CreationFolderItemId = configureCreation ? FolderId : string.Empty
            },
            _http,
            new FakeTokens(),
            registrationStore);

        public void Dispose() => _http.Dispose();

        /// <summary>A fixed Graph token for tests; production has no static-token path.</summary>
        private sealed class FakeTokens : IAccessTokenProvider
        {
            public Task<string> GetAccessTokenAsync(CancellationToken cancellationToken = default) =>
                Task.FromResult("fake-token");
        }

        private string NextETag() => $"\"etag-{++_nextETag}\"";

        private sealed class Entry
        {
            public string Id = "";
            public string Name = "";
            public byte[] Bytes = Array.Empty<byte>();
            public string ETag = "";
        }

        /// <summary>One upload session: what it targets and the bytes received so far.</summary>
        private sealed class Session
        {
            public string? ReplaceItemId;
            public string? Name;
            public string? IfMatch;
            public readonly MemoryStream Received = new();
            public bool Changed;
        }

        /// <summary>POST createUploadSession, for an item by id (replace) or a folder and name (new).</summary>
        private HttpResponseMessage OpenSession(HttpRequestMessage request, string rest)
        {
            var session = new Session();
            using (var body = JsonDocument.Parse(request.Content!.ReadAsStringAsync().GetAwaiter().GetResult()))
            {
                var item = body.RootElement.GetProperty("item");
                SessionConflictBehavior = item.GetProperty("@microsoft.graph.conflictBehavior").GetString();
                session.Name = item.TryGetProperty("name", out var name) ? name.GetString() : null;
            }
            SessionIfMatch = request.Headers.TryGetValues("If-Match", out var ifMatch) ? ifMatch.Single() : null;
            session.IfMatch = SessionIfMatch;

            var path = rest.Substring(0, rest.Length - "/createUploadSession".Length);
            if (path.EndsWith(":", StringComparison.Ordinal))
            {
                // {folderId}:/{name}: - a new file. The name must match the body's.
                if (!path.StartsWith(FolderId + ":/", StringComparison.Ordinal))
                    return Handler.GraphError(HttpStatusCode.NotFound, "itemNotFound");
                var name = Uri.UnescapeDataString(path.Substring(FolderId.Length + 2, path.Length - FolderId.Length - 3));
                if (!string.Equals(name, session.Name, StringComparison.Ordinal))
                    return Handler.GraphError(HttpStatusCode.BadRequest, "invalidRequest");
            }
            else
            {
                if (!_byId.TryGetValue(path, out var entry))
                    return Handler.GraphError(HttpStatusCode.NotFound, "itemNotFound");
                if (ChangeItemAtSessionStart) entry.ETag = NextETag();
                if (session.IfMatch is not null && session.IfMatch != entry.ETag)
                    return Handler.GraphError(HttpStatusCode.PreconditionFailed, "resourceModified");
                session.ReplaceItemId = path;
            }

            var id = $"s{++_nextSession}";
            _sessions[id] = session;
            SessionsCreated++;
            return Handler.Json(new Dictionary<string, object?>
            {
                ["uploadUrl"] = UploadHost + id,
                ["expirationDateTime"] = "2026-11-03T00:00:00Z"
            });
        }

        /// <summary>PUT, GET and DELETE on a session's pre-authenticated upload URL.</summary>
        private HttpResponseMessage Upload(HttpRequestMessage request, string id)
        {
            if (request.Headers.Authorization is not null) SawAuthorizedUploadCall = true;
            ChunkFault? fault = null;
            byte[]? chunk = null;
            if (request.Method == HttpMethod.Put)
            {
                chunk = request.Content!.ReadAsByteArrayAsync().GetAwaiter().GetResult();
                ChunkSizes.Add(chunk.Length);
                if (ChunkFaults.TryGetValue(_chunkPuts++, out var simulated)) fault = simulated;
                if (fault == ChunkFault.ExpireSession) _sessions.Remove(id);
            }

            if (!_sessions.TryGetValue(id, out var session))
                return Handler.GraphError(HttpStatusCode.NotFound, "itemNotFound");

            if (request.Method == HttpMethod.Delete)
            {
                _sessions.Remove(id);
                SessionsCancelled++;
                return new HttpResponseMessage(HttpStatusCode.NoContent);
            }
            if (request.Method == HttpMethod.Get)
                return Expecting(session, HttpStatusCode.OK);

            if (fault == ChunkFault.DropBeforeStore)
                throw new HttpRequestException("The connection was reset.");
            if (fault == ChunkFault.ServerError)
            {
                var busy = Handler.GraphError(HttpStatusCode.ServiceUnavailable, "serviceNotAvailable");
                busy.Headers.RetryAfter = new System.Net.Http.Headers.RetryConditionHeaderValue(TimeSpan.Zero);
                return busy;
            }

            var range = request.Content!.Headers.ContentRange!;
            if (range.From != session.Received.Length || range.To != range.From + chunk!.Length - 1)
                return Handler.GraphError(HttpStatusCode.RequestedRangeNotSatisfiable, "invalidRange");
            session.Received.Write(chunk, 0, chunk.Length);

            if (ChangeItemDuringSession && !session.Changed && session.ReplaceItemId is { } changing)
            {
                session.Changed = true;
                _byId[changing].ETag = NextETag();
            }

            if (session.Received.Length < range.Length)
            {
                if (fault == ChunkFault.StoreThenDrop)
                    throw new HttpRequestException("The connection was reset after the chunk was stored.");
                return Expecting(session, HttpStatusCode.Accepted);
            }

            // The last byte arrived: commit. A refused commit leaves the session open, as Graph does.
            HttpResponseMessage committed;
            if (session.ReplaceItemId is { } replaceId)
            {
                var entry = _byId[replaceId];
                if (session.IfMatch is not null && session.IfMatch != entry.ETag)
                    return Handler.GraphError(HttpStatusCode.PreconditionFailed, "resourceModified");
                entry.Bytes = session.Received.ToArray();
                entry.ETag = NextETag();
                committed = Handler.Item(entry);
            }
            else
            {
                if (_byName.ContainsKey(session.Name!))
                    return Handler.GraphError(HttpStatusCode.Conflict, "nameAlreadyExists");
                committed = Handler.Item(_byId[Seed(session.Name!, session.Received.ToArray())]);
                committed.StatusCode = HttpStatusCode.Created;
            }
            _sessions.Remove(id);
            if (fault == ChunkFault.StoreThenLoseResponse)
                throw new TaskCanceledException("The request timed out after the upload was stored.");
            return committed;
        }

        private static HttpResponseMessage Expecting(Session session, HttpStatusCode status)
        {
            var response = Handler.Json(new Dictionary<string, object?>
            {
                ["expirationDateTime"] = "2026-11-03T00:00:00Z",
                ["nextExpectedRanges"] = new[] { $"{session.Received.Length}-" }
            });
            response.StatusCode = status;
            return response;
        }

        private sealed class Handler : HttpMessageHandler
        {
            private readonly FakeGraphDrive _drive;

            public Handler(FakeGraphDrive drive) => _drive = drive;

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                var url = request.RequestUri!.ToString();
                var authorized = request.Headers.Authorization is { Scheme: "Bearer", Parameter: "fake-token" };
                if (request.Method != HttpMethod.Get) _drive.WriteRequests++;

                if (url.StartsWith(DownloadHost, StringComparison.Ordinal))
                {
                    if (request.Headers.Authorization is not null) _drive.SawAuthorizedDownloadCall = true;
                    var id = url.Substring(url.LastIndexOf('/') + 1);
                    return Task.FromResult(_drive._byId.TryGetValue(id, out var file)
                        ? Bytes(file.Bytes)
                        : Status(HttpStatusCode.NotFound));
                }

                // The pre-authenticated upload URL of a session: no bearer token expected.
                if (url.StartsWith(UploadHost, StringComparison.Ordinal))
                    return Task.FromResult(_drive.Upload(request, url.Substring(UploadHost.Length)));

                if (!authorized)
                    return Task.FromResult(Status(HttpStatusCode.Unauthorized));
                _drive.SawAuthorizedGraphCall = true;

                // GET /shares/{shareId}/driveItem - resolve a URL to its drive item.
                var sharesPrefix = $"{GraphBase}/shares/";
                if (url.StartsWith(sharesPrefix, StringComparison.Ordinal))
                {
                    var rest = url.Substring(sharesPrefix.Length);
                    var shareId = rest.Substring(0, rest.IndexOf('/'));
                    var decoded = DecodeShareId(shareId);
                    var id = decoded.Substring(decoded.LastIndexOf('/') + 1);
                    return Task.FromResult(_drive._byId.TryGetValue(id, out var entry)
                        ? Item(entry)
                        : GraphError(
                            HttpStatusCode.NotFound,
                            "itemNotFound",
                            _drive.ResolutionFailureMessage));
                }

                var prefix = $"{GraphBase}/drives/{Uri.EscapeDataString(DriveId)}";
                if (!url.StartsWith(prefix, StringComparison.Ordinal))
                    return Task.FromResult(Status(HttpStatusCode.NotFound));

                return Task.FromResult(Route(request, url.Substring(prefix.Length)));
            }

            private HttpResponseMessage Route(HttpRequestMessage request, string rest)
            {
                if (!rest.StartsWith("/items/", StringComparison.Ordinal))
                    return Status(HttpStatusCode.NotFound);
                rest = rest.Substring("/items/".Length);

                if (request.Method == HttpMethod.Post && rest.EndsWith("/createUploadSession", StringComparison.Ordinal))
                    return _drive.OpenSession(request, rest);

                if (request.Method == HttpMethod.Put && _drive.UploadStatus is { } simulated)
                {
                    _drive.UploadStatus = null;
                    return GraphError(simulated, _drive.UploadErrorCode);
                }

                // PUT /items/{folderId}:/{name}:/content. Graph replaces by default;
                // conflictBehavior=fail is what gives creation its no-overwrite guarantee.
                if (request.Method == HttpMethod.Put && rest.Contains(":/"))
                {
                    if (!rest.StartsWith(FolderId + ":/", StringComparison.Ordinal))
                        return GraphError(HttpStatusCode.NotFound, "itemNotFound");
                    if (_drive.CreationFailureMessage is not null)
                        return GraphError(
                            HttpStatusCode.NotFound,
                            "itemNotFound",
                            _drive.CreationFailureMessage);

                    var afterFolder = rest.Substring(rest.IndexOf(":/", StringComparison.Ordinal) + 2);
                    var contentMarker = afterFolder.IndexOf(":/content", StringComparison.Ordinal);
                    var name = Uri.UnescapeDataString(afterFolder.Substring(0, contentMarker));
                    var failOnConflict = rest.Contains(
                        "@microsoft.graph.conflictBehavior=fail",
                        StringComparison.Ordinal);
                    if (_drive._byName.TryGetValue(name, out var existing))
                    {
                        if (failOnConflict)
                            return GraphError(HttpStatusCode.Conflict, "nameAlreadyExists");
                        existing.Bytes = request.Content!.ReadAsByteArrayAsync().GetAwaiter().GetResult();
                        existing.ETag = _drive.NextETag();
                        return Item(existing);
                    }
                    var bytes = request.Content!.ReadAsByteArrayAsync().GetAwaiter().GetResult();
                    var id = _drive.Seed(name, bytes);
                    if (_drive.StoreThenLoseResponse)
                    {
                        _drive.StoreThenLoseResponse = false;
                        // HttpClient reports a timeout as a TaskCanceledException the caller did not ask for.
                        throw new TaskCanceledException("The request timed out after the upload was stored.");
                    }
                    return Item(_drive._byId[id]);
                }

                // GET /items/{folderId}:/{name} - existence probe for versioned names
                if (request.Method == HttpMethod.Get && rest.Contains(":/"))
                {
                    var name = Uri.UnescapeDataString(rest.Substring(rest.IndexOf(":/", StringComparison.Ordinal) + 2));
                    return _drive._byName.TryGetValue(name, out var probe)
                        ? Item(probe)
                        : GraphError(HttpStatusCode.NotFound, "itemNotFound");
                }

                // PUT /items/{id}/content - replace
                if (request.Method == HttpMethod.Put && rest.EndsWith("/content", StringComparison.Ordinal))
                {
                    var id = rest.Substring(0, rest.Length - "/content".Length);
                    if (!_drive._byId.TryGetValue(id, out var entry))
                        return GraphError(HttpStatusCode.NotFound, "itemNotFound");
                    if (request.Headers.TryGetValues("If-Match", out var values) &&
                        !values.Contains(entry.ETag, StringComparer.Ordinal))
                        return GraphError(HttpStatusCode.PreconditionFailed, "resourceModified");
                    entry.Bytes = request.Content!.ReadAsByteArrayAsync().GetAwaiter().GetResult();
                    entry.ETag = _drive.NextETag();
                    return Item(entry);
                }

                // GET /items/{id}
                if (request.Method == HttpMethod.Get)
                {
                    return _drive._byId.TryGetValue(rest, out var entry)
                        ? Item(entry)
                        : GraphError(HttpStatusCode.NotFound, "itemNotFound");
                }

                return Status(HttpStatusCode.MethodNotAllowed);
            }

            private static string DecodeShareId(string shareId)
            {
                var b64 = shareId.Substring("u!".Length).Replace('_', '/').Replace('-', '+');
                switch (b64.Length % 4)
                {
                    case 2: b64 += "=="; break;
                    case 3: b64 += "="; break;
                }
                return Encoding.UTF8.GetString(Convert.FromBase64String(b64));
            }

            internal static HttpResponseMessage Item(Entry entry) => Json(new Dictionary<string, object?>
            {
                ["id"] = entry.Id,
                ["name"] = entry.Name,
                ["eTag"] = entry.ETag,
                ["size"] = entry.Bytes.LongLength,
                ["parentReference"] = new Dictionary<string, object?>
                {
                    ["id"] = FolderId,
                    ["driveId"] = DriveId
                },
                ["file"] = new Dictionary<string, object?>
                {
                    ["mimeType"] = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                },
                ["@microsoft.graph.downloadUrl"] = $"{DownloadHost}/{entry.Id}"
            });

            internal static HttpResponseMessage Json(object payload) => new(HttpStatusCode.OK)
            {
                Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json")
            };

            private static HttpResponseMessage Bytes(byte[] bytes) => new(HttpStatusCode.OK)
            {
                Content = new ByteArrayContent(bytes)
            };

            private static HttpResponseMessage Status(HttpStatusCode code) => new(code);

            internal static HttpResponseMessage GraphError(
                HttpStatusCode status,
                string code,
                string? message = null) => new(status)
            {
                Content = new StringContent(
                    JsonSerializer.Serialize(new { error = new { code, message = message ?? code } }),
                    Encoding.UTF8, "application/json")
            };
        }
    }
}
