using System.Text.Json;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// The provider contract (docs/document-providers.md) for the object-store providers, run once per
/// store against an in-memory fake behind the real SDK: registration and the prefix boundary,
/// every save mode, version conflicts, size and extension limits, and how failed writes are
/// classified. Store-specific behaviour lives in the derived classes.
/// </summary>
public abstract class ObjectStoreProviderTests
{
    protected const string Connection = "store";
    protected const string Prefix = "contracts/";

    private protected abstract FakeObjectStore NewStore();

    /// <summary>Builds the provider under test over a fake store.</summary>
    private protected abstract IDocumentCreatingProvider Provider(
        FakeObjectStore store,
        string prefix = Prefix,
        long maximumBytes = 100 * 1024 * 1024,
        string? registrationIndexPath = null,
        bool registerOnce = false,
        string? templatesFolder = null,
        IReadOnlyCollection<string>? allowedExtensions = null,
        bool readOnly = false);

    protected abstract string ProviderName { get; }

    /// <summary>A URL (or URI) naming an object in the configured container or bucket.</summary>
    protected abstract string LocatorFor(string fullName);

    /// <summary>A URL (or URI) naming an object in a different container or bucket.</summary>
    protected abstract string ForeignLocatorFor(string fullName);

    /// <summary>Text that would reveal the account, endpoint or bucket if it appeared in a message.</summary>
    protected abstract string StorageIdentity { get; }

    [Fact]
    public async Task A_read_only_connection_registers_and_opens_but_writes_nothing()
    {
        var store = NewStore();
        store.Seed("contracts/Invoice.docx", Contract);
        var provider = Provider(store, readOnly: true);

        var reference = await provider.RegisterAsync("Invoice.docx");
        using (var content = await provider.OpenReadAsync(reference))
            Assert.Equal(Contract, ReadAll(content.Stream));
        var before = store.Requests.Count;

        var errors = new List<DocumentProviderException>();
        foreach (var mode in new[] { SaveMode.Replace, SaveMode.NewVersion, SaveMode.NewDocument })
        {
            using var bytes = Docx();
            errors.Add(await Assert.ThrowsAsync<DocumentProviderException>(() =>
                provider.SaveAsync(reference, bytes, new SaveDocumentOptions { Mode = mode })));
        }
        using (var bytes = Docx())
            errors.Add(await Assert.ThrowsAsync<DocumentProviderException>(() => provider.CreateAsync("new.docx", bytes)));

        Assert.All(errors, error =>
        {
            Assert.Equal(ProviderErrorCode.AccessDenied, error.Code);
            Assert.Contains("read-only", error.Message, StringComparison.Ordinal);
        });
        Assert.Equal(before, store.Requests.Count);
        await provider.RemoveAsync(reference);
    }

    private static OfficeAgentClient EngineOver(IDocumentProvider provider) =>
        new(new DocumentProviderRegistry(new[] { provider }), new WordModule());

    /// <summary>One contract per test: the factory's bytes differ from call to call.</summary>
    private protected byte[] Contract { get; } = DocxFactory.Contract();

    private MemoryStream Docx() => new(Contract);

    [Fact]
    public async Task Register_mints_an_opaque_id_and_open_returns_the_bytes_and_version()
    {
        var store = NewStore();
        var etag = store.Seed("contracts/2026/acme.docx", Contract);
        var provider = Provider(store);

        var reference = await provider.RegisterAsync("2026/acme.docx");

        Assert.Equal(ProviderName, reference.Provider);
        Assert.Equal(Connection, reference.ConnectionId);
        Assert.DoesNotContain('/', reference.ItemId);
        Assert.DoesNotContain("acme", reference.ItemId);
        Assert.Equal("acme.docx", reference.Name);
        Assert.Equal(OfficeContentTypesForDocx, reference.ContentType);
        Assert.False(string.IsNullOrEmpty(reference.Version));

        using var content = await provider.OpenReadAsync(reference);
        Assert.Equal(Contract, ReadAll(content.Stream));
        Assert.Equal(reference.Version, content.Reference.Version);
        Assert.Equal(etag, reference.Version);
    }

    [Fact]
    public async Task Register_accepts_a_locator_inside_the_container_and_prefix()
    {
        var store = NewStore();
        store.Seed("contracts/acme.docx", Contract);
        var provider = Provider(store);

        var reference = await provider.RegisterAsync(LocatorFor("contracts/acme.docx"));

        Assert.Equal("acme.docx", reference.Name);
    }

    [Theory]
    [InlineData("../acme.docx")]
    [InlineData("2026/../../acme.docx")]
    [InlineData("./acme.docx")]
    [InlineData("/contracts/acme.docx")]
    [InlineData("2026\\acme.docx")]
    [InlineData("https://elsewhere.example/acme.docx")]
    [InlineData("ftp://elsewhere.example/acme.docx")]
    public async Task Names_that_escape_the_prefix_are_access_denied(string source)
    {
        var store = NewStore();
        store.Seed("acme.docx", Contract);
        store.Seed("contracts/acme.docx", Contract);
        var provider = Provider(store);

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.RegisterAsync(source));

        Assert.Equal(ProviderErrorCode.AccessDenied, error.Code);
        Assert.Empty(store.Requests);
    }

    [Fact]
    public async Task Locators_for_another_container_or_outside_the_prefix_are_access_denied()
    {
        var store = NewStore();
        store.Seed("acme.docx", Contract);
        var provider = Provider(store);

        var foreign = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.RegisterAsync(ForeignLocatorFor("contracts/acme.docx")));
        var outsidePrefix = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.RegisterAsync(LocatorFor("acme.docx")));

        Assert.Equal(ProviderErrorCode.AccessDenied, foreign.Code);
        Assert.Equal(ProviderErrorCode.AccessDenied, outsidePrefix.Code);
        Assert.Empty(store.Requests);
    }

    [Theory]
    [InlineData("")]
    [InlineData("2026//acme.docx")]
    [InlineData("2026/")]
    public async Task Malformed_names_are_invalid_arguments(string source)
    {
        var store = NewStore();
        var provider = Provider(store);

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.RegisterAsync(source));

        Assert.Equal(ProviderErrorCode.InvalidArgument, error.Code);
    }

    [Fact]
    public async Task A_disallowed_extension_is_refused_before_the_store_is_asked()
    {
        var store = NewStore();
        store.Seed("contracts/payload.exe", new byte[] { 1, 2, 3 });
        var provider = Provider(store);

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.RegisterAsync("payload.exe"));

        Assert.Equal(ProviderErrorCode.ExtensionNotAllowed, error.Code);
        Assert.Empty(store.Requests);
    }

    [Fact]
    public async Task A_missing_object_is_not_found()
    {
        var store = NewStore();
        var provider = Provider(store);

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.RegisterAsync("missing.docx"));

        Assert.Equal(ProviderErrorCode.NotFound, error.Code);
    }

    [Fact]
    public async Task An_oversized_object_is_refused_without_downloading_it()
    {
        var store = NewStore();
        store.Seed("contracts/big.docx", new byte[2048]);
        store.Seed("contracts/small.docx", new byte[512]);
        var provider = Provider(store, maximumBytes: 1024);

        var onRegister = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.RegisterAsync("big.docx"));
        Assert.Equal(ProviderErrorCode.ContentTooLarge, onRegister.Code);

        // An object that grows after registration is refused on open, still before any download.
        var reference = await provider.RegisterAsync("small.docx");
        store.Touch("contracts/small.docx", new byte[4096]);
        var onOpen = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.OpenReadAsync(reference));

        Assert.Equal(ProviderErrorCode.ContentTooLarge, onOpen.Code);
        Assert.Equal(0, store.Count("GET"));
    }

    [Fact]
    public async Task Oversized_save_content_is_refused_before_any_write()
    {
        var store = NewStore();
        store.Seed("contracts/acme.docx", new byte[16]);
        var provider = Provider(store, maximumBytes: 1024);
        var reference = await provider.RegisterAsync("acme.docx");

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.SaveAsync(reference, new MemoryStream(new byte[4096]), new SaveDocumentOptions()));

        Assert.Equal(ProviderErrorCode.ContentTooLarge, error.Code);
        Assert.Equal(0, store.Count("PUT"));
    }

    [Fact]
    public async Task Replace_overwrites_in_place_under_if_match_and_keeps_the_id()
    {
        var store = NewStore();
        store.Seed("contracts/acme.docx", new byte[] { 1 }, metadata: new Dictionary<string, string> { ["owner"] = "legal" });
        var provider = Provider(store);
        var reference = await provider.RegisterAsync("acme.docx");
        var before = store.Objects["contracts/acme.docx"].ETag;

        var saved = await provider.SaveAsync(reference, Docx(), new SaveDocumentOptions { Mode = SaveMode.Replace });

        Assert.Equal(reference.ItemId, saved.ItemId);
        Assert.NotEqual(reference.Version, saved.Version);
        Assert.Equal(Contract, store.BytesOf("contracts/acme.docx"));
        var put = store.Requests.Single(request => request.Method == "PUT");
        Assert.Equal(before, put.Headers["If-Match"]);
        // A replace keeps what the host attached to the object.
        Assert.Equal("legal", store.Objects["contracts/acme.docx"].Metadata["owner"]);

        using var reopened = await provider.OpenReadAsync(saved);
        Assert.Equal(saved.Version, reopened.Reference.Version);
    }

    [Fact]
    public async Task A_stale_version_is_a_conflict_and_nothing_is_written()
    {
        var store = NewStore();
        store.Seed("contracts/acme.docx", Contract);
        var provider = Provider(store);
        var reference = await provider.RegisterAsync("acme.docx");
        store.Touch("contracts/acme.docx");
        var current = store.BytesOf("contracts/acme.docx");

        var onSave = await Assert.ThrowsAsync<DocumentVersionConflictException>(() =>
            provider.SaveAsync(reference, Docx(), new SaveDocumentOptions { Mode = SaveMode.Replace }));
        var onOpen = await Assert.ThrowsAsync<DocumentVersionConflictException>(() => provider.OpenReadAsync(reference));

        Assert.Equal(ProviderErrorCode.VersionConflict, onSave.Code);
        Assert.Equal(ProviderErrorCode.VersionConflict, onOpen.Code);
        Assert.Equal(current, store.BytesOf("contracts/acme.docx"));
        Assert.Equal(0, store.Count("PUT"));
    }

    [Fact]
    public async Task A_change_between_the_check_and_the_write_is_a_conflict_not_an_overwrite()
    {
        var store = NewStore();
        store.Seed("contracts/acme.docx", Contract);
        var provider = Provider(store);
        var reference = await provider.RegisterAsync("acme.docx");
        var intruder = new byte[] { 9, 9, 9 };
        store.BeforeWrite = name => { store.BeforeWrite = null; store.Touch(name, intruder); };

        var error = await Assert.ThrowsAsync<DocumentVersionConflictException>(() =>
            provider.SaveAsync(reference, Docx(), new SaveDocumentOptions { Mode = SaveMode.Replace }));

        Assert.Equal(ProviderErrorCode.VersionConflict, error.Code);
        Assert.Equal(intruder, store.BytesOf("contracts/acme.docx"));
    }

    [Fact]
    public async Task New_version_writes_a_versioned_sibling_create_only_and_mints_a_new_id()
    {
        var store = NewStore();
        store.Seed("contracts/2026/acme.docx", Contract);
        var provider = Provider(store);
        var reference = await provider.RegisterAsync("2026/acme.docx");
        var original = store.BytesOf("contracts/2026/acme.docx");

        var v2 = await provider.SaveAsync(reference, Docx(), new SaveDocumentOptions { Mode = SaveMode.NewVersion });
        var v3 = await provider.SaveAsync(reference, Docx(), new SaveDocumentOptions { Mode = SaveMode.NewVersion });

        Assert.Equal("acme.v2.docx", v2.Name);
        Assert.Equal("acme.v3.docx", v3.Name);
        Assert.NotEqual(reference.ItemId, v2.ItemId);
        Assert.NotEqual(v2.ItemId, v3.ItemId);
        Assert.True(store.Has("contracts/2026/acme.v2.docx"));
        Assert.True(store.Has("contracts/2026/acme.v3.docx"));
        Assert.Equal(original, store.BytesOf("contracts/2026/acme.docx"));
        Assert.All(store.Requests.Where(request => request.Method == "PUT"),
            put => Assert.Equal("*", put.Headers["If-None-Match"]));

        using var opened = await provider.OpenReadAsync(v2);
        Assert.Equal(Contract, ReadAll(opened.Stream));
    }

    [Fact]
    public async Task New_version_takes_the_next_name_when_another_writer_wins_the_race()
    {
        var store = NewStore();
        store.Seed("contracts/acme.docx", Contract);
        var provider = Provider(store);
        var reference = await provider.RegisterAsync("acme.docx");
        // Another writer creates acme.v2.docx after the probe and before this upload.
        store.BeforeWrite = name =>
        {
            store.BeforeWrite = null;
            store.Seed(name, new byte[] { 7 });
        };

        var saved = await provider.SaveAsync(reference, Docx(), new SaveDocumentOptions { Mode = SaveMode.NewVersion });

        Assert.Equal("acme.v3.docx", saved.Name);
        Assert.Equal(new byte[] { 7 }, store.BytesOf("contracts/acme.v2.docx"));
    }

    [Fact]
    public async Task New_document_uses_the_validated_name_and_never_overwrites()
    {
        var store = NewStore();
        store.Seed("contracts/acme.docx", Contract);
        store.Seed("contracts/taken.docx", new byte[] { 5 });
        var provider = Provider(store);
        var reference = await provider.RegisterAsync("acme.docx");

        var copy = await provider.SaveAsync(reference, Docx(),
            new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "copy.docx" });
        var taken = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(reference, Docx(),
            new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "taken.docx" }));
        var path = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(reference, Docx(),
            new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "../escape.docx" }));
        var destination = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.SaveAsync(reference, Docx(),
            new SaveDocumentOptions { Mode = SaveMode.NewDocument, DestinationItemId = "chosen" }));

        Assert.Equal("copy.docx", copy.Name);
        Assert.True(store.Has("contracts/copy.docx"));
        Assert.Equal(ProviderErrorCode.AlreadyExists, taken.Code);
        Assert.Equal(new byte[] { 5 }, store.BytesOf("contracts/taken.docx"));
        Assert.Equal(ProviderErrorCode.InvalidArgument, path.Code);
        Assert.Equal(ProviderErrorCode.InvalidArgument, destination.Code);
    }

    [Fact]
    public async Task Create_writes_under_the_prefix_and_never_overwrites()
    {
        var store = NewStore();
        var provider = Provider(store);

        var created = await provider.CreateAsync("brief.docx", Docx());
        var again = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.CreateAsync("brief.docx", new MemoryStream(new byte[] { 1 })));

        Assert.Equal("brief.docx", created.Name);
        Assert.True(store.Has("contracts/brief.docx"));
        Assert.Equal(OfficeContentTypesForDocx, store.Objects["contracts/brief.docx"].ContentType);
        Assert.Equal(ProviderErrorCode.AlreadyExists, again.Code);
        Assert.Equal(Contract, store.BytesOf("contracts/brief.docx"));
    }

    [Theory]
    [InlineData("nested/brief.docx")]
    [InlineData("CON.docx")]
    [InlineData("brief.exe")]
    public async Task Create_refuses_paths_reserved_names_and_disallowed_extensions(string name)
    {
        var store = NewStore();
        var provider = Provider(store);

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.CreateAsync(name, Docx()));

        Assert.Contains(error.Code, new[] { ProviderErrorCode.InvalidArgument, ProviderErrorCode.ExtensionNotAllowed });
        Assert.Equal(0, store.Count("PUT"));
    }

    [Fact]
    public async Task Create_through_the_engine_applies_an_initial_plan()
    {
        var store = NewStore();
        var client = EngineOver(Provider(store));
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertOp
                {
                    Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = string.Empty },
                    Position = InsertPosition.Before,
                    Text = "Object store brief"
                }
            }
        };

        var created = await client.CreateAsync(Connection, "brief.docx", plan);

        Assert.True(created.Committed);
        var inspected = await client.InspectAsync(created.Document!);
        Assert.Equal("Object store brief", inspected.Paragraphs[0].Text);
    }

    [Fact]
    public async Task Remove_drops_the_registration_but_never_the_object()
    {
        var store = NewStore();
        store.Seed("contracts/acme.docx", Contract);
        var provider = Provider(store);
        var reference = await provider.RegisterAsync("acme.docx");

        await provider.RemoveAsync(reference);

        Assert.True(store.Has("contracts/acme.docx"));
        var open = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.OpenReadAsync(reference));
        var again = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.RemoveAsync(reference));
        Assert.Equal(ProviderErrorCode.NotFound, open.Code);
        Assert.Equal(ProviderErrorCode.NotFound, again.Code);
        Assert.DoesNotContain(store.Requests, request => request.Method == "DELETE");
    }

    [Fact]
    public async Task Path_like_ids_are_access_denied()
    {
        var store = NewStore();
        var provider = Provider(store);

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.OpenReadAsync(DocumentReference.For(ProviderName, Connection, "../contracts/acme.docx")));

        Assert.Equal(ProviderErrorCode.AccessDenied, error.Code);
    }

    [Fact]
    public async Task A_registration_outside_a_narrowed_prefix_is_access_denied()
    {
        var index = Path.Combine(Path.GetTempPath(), $"officeagent-objectstore-{Guid.NewGuid():N}.json");
        try
        {
            var store = NewStore();
            store.Seed("contracts/acme.docx", Contract);
            var wide = Provider(store, prefix: "", registrationIndexPath: index);
            var reference = await wide.RegisterAsync("contracts/acme.docx");

            // The same index, reloaded by a connection confined to another prefix.
            var narrow = Provider(store, prefix: "other/", registrationIndexPath: index);
            var error = await Assert.ThrowsAsync<DocumentProviderException>(() => narrow.OpenReadAsync(reference));

            Assert.Equal(ProviderErrorCode.AccessDenied, error.Code);
        }
        finally
        {
            File.Delete(index);
        }
    }

    [Fact]
    public async Task A_json_registration_index_survives_a_restart()
    {
        var index = Path.Combine(Path.GetTempPath(), $"officeagent-objectstore-{Guid.NewGuid():N}", "index.json");
        try
        {
            var store = NewStore();
            store.Seed("contracts/acme.docx", Contract);
            var reference = await Provider(store, registrationIndexPath: index).RegisterAsync("acme.docx");

            using var reopened = await Provider(store, registrationIndexPath: index).OpenReadAsync(reference);

            Assert.Equal(reference.ItemId, reopened.Reference.ItemId);
            Assert.DoesNotContain(StorageIdentity, File.ReadAllText(index));
        }
        finally
        {
            Directory.Delete(Path.GetDirectoryName(index)!, recursive: true);
        }
    }

    [Theory]
    [InlineData(400)]
    [InlineData(409)]
    public async Task A_refused_write_is_certain_and_stores_nothing(int status)
    {
        var store = NewStore();
        store.WriteStatus = status;

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            EngineOver(Provider(store)).CreateAsync(Connection, "refused.docx"));

        Assert.Equal(ProviderErrorCode.WriteRejected, error.Code);
        Assert.IsNotAssignableFrom<DocumentWriteRecoveryException>(error);
        Assert.False(store.Has("contracts/refused.docx"));
    }

    [Fact]
    public async Task A_forbidden_write_is_access_denied()
    {
        var store = NewStore();
        store.Seed("contracts/acme.docx", Contract);
        var provider = Provider(store);
        var reference = await provider.RegisterAsync("acme.docx");
        store.WriteStatus = 403;
        store.WriteErrorCode = "AccessDenied";

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.SaveAsync(reference, Docx(), new SaveDocumentOptions { Mode = SaveMode.Replace }));

        Assert.Equal(ProviderErrorCode.AccessDenied, error.Code);
    }

    [Fact]
    public async Task A_server_error_on_a_write_is_an_unknown_outcome()
    {
        var store = NewStore();
        store.WriteStatus = 503;
        store.WriteErrorCode = "ServerBusy";

        var error = await Assert.ThrowsAsync<DocumentWriteOutcomeUnknownException>(() =>
            EngineOver(Provider(store)).CreateAsync(Connection, "maybe.docx"));

        Assert.Equal(ProviderErrorCode.OutcomeUnknown, error.Code);
        Assert.Equal("maybe.docx", error.OutputName);
    }

    [Fact]
    public async Task A_lost_response_after_a_stored_write_is_unknown_and_reconcilable()
    {
        var store = NewStore();
        store.Seed("contracts/acme.docx", Contract);
        var provider = Provider(store);
        var reference = await provider.RegisterAsync("acme.docx");
        store.StoreThenLoseResponse = true;
        var bytes = Contract.Concat(new byte[] { 1 }).ToArray();

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.SaveAsync(reference, new MemoryStream(bytes), new SaveDocumentOptions { Mode = SaveMode.Replace }));

        // The write landed although the caller was never told: it must not read as nothing written.
        Assert.Equal(ProviderErrorCode.OutcomeUnknown, error.Code);
        Assert.Equal(bytes, store.BytesOf("contracts/acme.docx"));
    }

    [Theory]
    [InlineData(SaveMode.Replace)]
    [InlineData(SaveMode.NewVersion)]
    public async Task A_retried_write_that_already_landed_is_recognised_as_this_write(SaveMode mode)
    {
        // The SDK retries a timed-out request. If the first attempt stored the bytes, the retry
        // fails its own precondition; that must not be reported as a conflict or a taken name.
        var store = NewStore();
        store.Seed("contracts/acme.docx", Contract);
        var provider = Provider(store);
        var reference = await provider.RegisterAsync("acme.docx");
        store.StoreThenStatus = mode == SaveMode.Replace ? 412 : PreconditionStatusForCreate;
        store.StoreThenErrorCode = mode == SaveMode.Replace ? PreconditionCodeForReplace : PreconditionCodeForCreate;

        var saved = await provider.SaveAsync(reference, Docx(), new SaveDocumentOptions { Mode = mode });

        var name = mode == SaveMode.Replace ? "contracts/acme.docx" : "contracts/acme.v2.docx";
        Assert.Equal(Contract, store.BytesOf(name));
        using var opened = await provider.OpenReadAsync(saved);
        Assert.Equal(saved.Version, opened.Reference.Version);
    }

    [Fact]
    public async Task A_registration_failure_after_a_stored_write_carries_a_path_free_locator()
    {
        var store = NewStore();
        store.Seed("contracts/acme.docx", Contract);
        var provider = Provider(store, registerOnce: true);
        var reference = await provider.RegisterAsync("acme.docx");

        var error = await Assert.ThrowsAsync<DocumentRegistrationFailedException>(() => provider.SaveAsync(
            reference, Docx(), new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "copy.docx" }));

        Assert.Equal(ProviderErrorCode.RegistrationFailed, error.Code);
        Assert.Equal("copy.docx", error.OutputName);
        Assert.Equal(reference.ItemId, error.ItemId);
        Assert.True(store.Has("contracts/copy.docx"));
        Assert.Equal(Sha256(store.BytesOf("contracts/copy.docx")), error.OutputSha256);
        Assert.DoesNotContain(StorageIdentity, error.Message);
    }

    [Fact]
    public async Task Tool_results_never_name_the_account_endpoint_or_bucket()
    {
        var store = NewStore();
        store.Seed("contracts/acme.docx", Contract);
        var provider = Provider(store);
        var reference = await provider.RegisterAsync("acme.docx");
        var tools = new OfficeAgentTools(EngineOver(provider));
        store.WriteStatus = 403;
        store.WriteErrorCode = "AccessDenied";

        using var hits = JsonDocument.Parse(await tools.FindInDocument(Connection, reference.ItemId, "Acme Corp"));
        var hit = hits.RootElement[0];
        var plan = "{\"operations\":[{\"op\":\"changeText\",\"target\":{\"paraId\":\"" + hit.GetProperty("paraId").GetString() +
                   "\",\"expect\":\"Acme Corp\",\"occurrence\":" + hit.GetProperty("occurrence").GetInt32() +
                   "},\"mode\":\"Direct\",\"with\":\"Contoso\"}]}";
        var result = await tools.ApplyPlan(Connection, reference.ItemId, plan);

        Assert.Contains("access-denied", result);
        Assert.DoesNotContain(StorageIdentity, result);
        Assert.DoesNotContain(Prefix, result);
    }

    [Fact]
    public async Task A_plan_commits_through_the_engine_with_replace()
    {
        var store = NewStore();
        store.Seed("contracts/acme.docx", Contract);
        var provider = Provider(store);
        var reference = await provider.RegisterAsync("acme.docx");
        var client = EngineOver(provider);
        var hit = (await client.FindAsync(reference, new FindQuery("Acme Corp"))).First();

        var result = await client.CommitAsync(reference, new DocumentPlan
        {
            Operations = new PlanOperation[] { new ChangeTextOp { Target = hit.Anchor, With = "Contoso", Mode = ChangeMode.Direct } }
        });

        Assert.True(result.Committed);
        var inspected = await client.InspectAsync(result.Document!);
        Assert.Contains(inspected.Paragraphs, paragraph => paragraph.Text.Contains("Contoso"));
    }

    // ── Listing ───────────────────────────────────────────────────────────────

    private IDocumentListingProvider Lister(
        FakeObjectStore store, string prefix = Prefix, string? templatesFolder = null,
        IReadOnlyCollection<string>? allowedExtensions = null, long maximumBytes = 100 * 1024 * 1024) =>
        Assert.IsAssignableFrom<IDocumentListingProvider>(Provider(
            store, prefix, maximumBytes, templatesFolder: templatesFolder, allowedExtensions: allowedExtensions));

    [Fact]
    public async Task Listing_pages_across_service_pages_and_resumes_inside_one()
    {
        var store = NewStore();
        for (var i = 0; i < 205; i++) store.Seed($"contracts/doc{i:000}.docx", new byte[] { 1 });
        store.Seed("outside.docx", new byte[] { 1 });
        var provider = Lister(store);

        var first = await provider.ListAsync(new DocumentListQuery { MaximumResults = 150 });
        var second = await provider.ListAsync(new DocumentListQuery { MaximumResults = 150, ContinuationToken = first.ContinuationToken });

        Assert.Equal(150, first.Items.Count);
        Assert.NotNull(first.ContinuationToken);
        Assert.Equal(55, second.Items.Count);
        Assert.Null(second.ContinuationToken);
        var names = first.Items.Concat(second.Items).Select(item => item.Name).ToList();
        Assert.Equal(Enumerable.Range(0, 205).Select(i => $"doc{i:000}.docx"), names);
        // The first page stopped inside the store's first page of 200; the second read it again
        // from where it stopped, then the store's second page.
        Assert.Equal(3, store.Count("LIST"));
        Assert.All(store.Requests.Where(r => r.Method == "LIST"), r => Assert.Equal("contracts/", r.Name));
        // The token is encrypted: it names neither the prefix nor the store.
        Assert.DoesNotContain("contracts", first.ContinuationToken);
        Assert.DoesNotContain(StorageIdentity, first.ContinuationToken);
    }

    [Fact]
    public async Task Listing_filters_by_name_extension_and_size_and_never_lists_macro_files()
    {
        var store = NewStore();
        store.Seed("contracts/Acme-NDA.docx", new byte[] { 1 });
        store.Seed("contracts/acme-deck.pptx", new byte[] { 1 });
        store.Seed("contracts/acme-budget.xlsx", new byte[] { 1 });
        store.Seed("contracts/acme-macro.docm", new byte[] { 1 });
        store.Seed("contracts/acme-big.docx", new byte[4096]);
        store.Seed("contracts/other.docx", new byte[] { 1 });
        var provider = Lister(store, allowedExtensions: new[] { ".docx", ".pptx", ".docm" }, maximumBytes: 1024);

        var all = await provider.ListAsync(new DocumentListQuery());
        var named = await provider.ListAsync(new DocumentListQuery { NameContains = "ACME" });
        var narrowed = await provider.ListAsync(new DocumentListQuery { Extensions = new[] { "pptx", ".xlsx" } });
        var macros = await provider.ListAsync(new DocumentListQuery { Extensions = new[] { ".docm" } });

        Assert.Equal(new[] { "Acme-NDA.docx", "acme-deck.pptx", "other.docx" }, all.Items.Select(item => item.Name));
        Assert.Equal(new[] { "Acme-NDA.docx", "acme-deck.pptx" }, named.Items.Select(item => item.Name));
        Assert.Equal("acme-deck.pptx", Assert.Single(narrowed.Items).Name);
        Assert.Empty(macros.Items);
        Assert.Null(macros.ContinuationToken);
        var entry = all.Items[0];
        Assert.Equal("application/vnd.openxmlformats-officedocument.wordprocessingml.document", entry.ContentType);
        Assert.Equal(1, entry.Length);
        Assert.NotNull(entry.LastModified);
        Assert.False(entry.IsTemplate);
        Assert.Null(entry.Folder);
    }

    [Fact]
    public async Task Listing_reads_one_level_unless_recursive_and_names_each_folder()
    {
        var store = NewStore();
        store.Seed("contracts/top.docx", new byte[] { 1 });
        store.Seed("contracts/2026/mid.docx", new byte[] { 1 });
        store.Seed("contracts/2026/q1/deep.docx", new byte[] { 1 });
        store.Seed("contracts/2026/", Array.Empty<byte>());
        store.Seed("contractsX/beside.docx", new byte[] { 1 });
        var provider = Lister(store);

        var flat = await provider.ListAsync(new DocumentListQuery());
        var recursive = await provider.ListAsync(new DocumentListQuery { Recursive = true });
        var folder = await provider.ListAsync(new DocumentListQuery { Folder = "/2026/" });

        Assert.Equal("top.docx", Assert.Single(flat.Items).Name);
        Assert.Equal(
            new[] { ("top.docx", (string?)null), ("mid.docx", "2026"), ("deep.docx", "2026/q1") },
            recursive.Items.Select(item => (item.Name, item.Folder)));
        Assert.Equal(new[] { "top.docx", "2026/mid.docx", "2026/q1/deep.docx" }, recursive.Items.Select(item => item.Source));
        var mid = Assert.Single(folder.Items);
        Assert.Equal(("mid.docx", "2026", "2026/mid.docx"), (mid.Name, mid.Folder, mid.Source));
    }

    [Fact]
    public async Task Every_listed_source_registers_and_opens()
    {
        var store = NewStore();
        store.Seed("contracts/acme.docx", Contract);
        store.Seed("contracts/2026/beta.docx", Contract);
        var provider = Lister(store);

        var listed = await provider.ListAsync(new DocumentListQuery { Recursive = true });

        Assert.Equal(2, listed.Items.Count);
        foreach (var entry in listed.Items)
        {
            var reference = await provider.RegisterAsync(entry.Source);
            Assert.Equal(entry.Name, reference.Name);
            using var opened = await provider.OpenReadAsync(reference);
            Assert.Equal(Contract, ReadAll(opened.Stream));
        }
    }

    [Fact]
    public async Task A_templates_listing_reads_the_templates_folder_and_flags_every_entry()
    {
        var store = NewStore();
        store.Seed("contracts/templates/letter.dotx", Contract);
        store.Seed("contracts/templates/memo.docx", new byte[] { 1 });
        store.Seed("contracts/plan.dotx", new byte[] { 1 });
        store.Seed("contracts/a.docx", new byte[] { 1 });
        var provider = Lister(store, templatesFolder: "/templates/", allowedExtensions: new[] { ".docx", ".dotx" });
        var source = Assert.IsAssignableFrom<IConnectionTemplateSource>(provider);

        var templates = await provider.ListAsync(new DocumentListQuery { Templates = true });
        var documents = await provider.ListAsync(new DocumentListQuery());

        Assert.Equal("templates", source.TemplatesFolder);
        Assert.Equal(TemplateLocationKind.Connection, source.TemplateLocation);
        Assert.Equal(new[] { "letter.dotx", "memo.docx" }, templates.Items.Select(item => item.Name));
        Assert.All(templates.Items, item => Assert.True(item.IsTemplate));
        Assert.All(templates.Items, item => Assert.Null(item.Folder));
        Assert.Equal("templates/letter.dotx", templates.Items[0].Source);
        Assert.Equal("letter.dotx", (await provider.RegisterAsync(templates.Items[0].Source)).Name);
        Assert.Equal(new[] { ("a.docx", false), ("plan.dotx", true) }, documents.Items.Select(item => (item.Name, item.IsTemplate)));
    }

    [Fact]
    public async Task A_connection_without_templates_lists_none_without_asking_the_store()
    {
        var store = NewStore();
        store.Seed("contracts/templates/letter.docx", new byte[] { 1 });
        var provider = Lister(store);
        var source = Assert.IsAssignableFrom<IConnectionTemplateSource>(provider);

        var templates = await provider.ListAsync(new DocumentListQuery { Templates = true });

        Assert.Null(source.TemplatesFolder);
        Assert.Equal(TemplateLocationKind.None, source.TemplateLocation);
        Assert.Empty(templates.Items);
        Assert.Null(templates.ContinuationToken);
        Assert.Equal(0, store.Count("LIST"));
    }

    [Theory]
    [InlineData("..", ProviderErrorCode.AccessDenied)]
    [InlineData("2026/../..", ProviderErrorCode.AccessDenied)]
    [InlineData("./2026", ProviderErrorCode.AccessDenied)]
    [InlineData("2026\\q1", ProviderErrorCode.AccessDenied)]
    [InlineData("2026//q1", ProviderErrorCode.InvalidArgument)]
    public async Task A_listing_folder_that_leaves_the_prefix_is_refused(string folder, ProviderErrorCode expected)
    {
        var store = NewStore();
        store.Seed("secret.docx", new byte[] { 1 });
        var provider = Lister(store);

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
            provider.ListAsync(new DocumentListQuery { Folder = folder }));

        Assert.Equal(expected, error.Code);
        Assert.Equal(0, store.Count("LIST"));
    }

    [Fact]
    public async Task A_continuation_token_cannot_be_edited_reused_elsewhere_or_moved_out_of_the_prefix()
    {
        var store = NewStore();
        for (var i = 0; i < 3; i++) store.Seed($"contracts/doc{i}.docx", new byte[] { 1 });
        store.Seed("other/doc.docx", new byte[] { 1 });
        var provider = Lister(store);
        var token = (await provider.ListAsync(new DocumentListQuery { MaximumResults = 1 })).ContinuationToken!;
        var edited = token.Substring(0, 10) + (token[10] == 'A' ? 'B' : 'A') + token.Substring(11);
        // The same connection id confined to another prefix, as after a configuration change.
        var narrowed = Lister(store, prefix: "other/");

        foreach (var attempt in new Func<Task>[]
                 {
                     () => provider.ListAsync(new DocumentListQuery { MaximumResults = 1, ContinuationToken = edited }),
                     () => provider.ListAsync(new DocumentListQuery { MaximumResults = 1, ContinuationToken = "not-a-token" }),
                     () => provider.ListAsync(new DocumentListQuery { MaximumResults = 1, NameContains = "doc", ContinuationToken = token }),
                     () => provider.ListAsync(new DocumentListQuery { MaximumResults = 1, Recursive = true, ContinuationToken = token }),
                     () => narrowed.ListAsync(new DocumentListQuery { MaximumResults = 1, ContinuationToken = token })
                 })
        {
            var error = await Assert.ThrowsAsync<DocumentProviderException>(attempt);
            Assert.Equal(ProviderErrorCode.InvalidArgument, error.Code);
        }
        Assert.Equal(1, store.Count("LIST"));
    }

    [Theory]
    [InlineData(403, "AuthorizationFailure", ProviderErrorCode.AccessDenied)]
    [InlineData(404, "NotFound", ProviderErrorCode.NotFound)]
    [InlineData(500, "InternalError", ProviderErrorCode.IO)]
    public async Task Listing_failures_map_to_provider_errors(int status, string code, ProviderErrorCode expected)
    {
        var store = NewStore();
        store.ListStatus = status;
        store.ListErrorCode = code;
        var provider = Lister(store);

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() => provider.ListAsync(new DocumentListQuery()));

        Assert.Equal(expected, error.Code);
        Assert.Equal(ProviderName, error.Provider);
        Assert.DoesNotContain(StorageIdentity, error.Message);
    }

    protected abstract int PreconditionStatusForCreate { get; }
    protected abstract string PreconditionCodeForCreate { get; }
    protected abstract string PreconditionCodeForReplace { get; }

    private const string OfficeContentTypesForDocx = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";

    private static byte[] ReadAll(Stream stream)
    {
        using var copy = new MemoryStream();
        stream.CopyTo(copy);
        return copy.ToArray();
    }

    private protected static string Sha256(byte[] bytes) =>
        Convert.ToHexString(System.Security.Cryptography.SHA256.HashData(bytes)).ToLowerInvariant();
}
