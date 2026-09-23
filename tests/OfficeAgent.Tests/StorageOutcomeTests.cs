using System.Security.Claims;
using System.Security.Cryptography;
using System.Text.Json;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.Tests;

/// <summary>
/// The fault-injection matrix: one boundary at a time, what storage then holds, and what the
/// caller is told. Every case asserts the observable storage state as well as the report,
/// because a report that disagrees with storage is the defect this matrix exists to catch.
/// </summary>
/// <remarks>
/// Four outcomes exist, and only one lets a caller conclude nothing was written:
/// <c>notWritten</c> (validation, access, a stale source, or a refusal the provider proves),
/// <c>committed</c>, <c>writtenNotRegistered</c>, and <c>unknown</c> (anything after a write
/// began that the provider did not classify, including cancellation). Filesystem and memory are
/// real providers here; <see cref="Injector"/> reaches boundaries they cannot be made to fail
/// at on demand. SharePoint is covered by controlled simulations in SharePointProviderTests.
/// </remarks>
public sealed class StorageOutcomeTests : IDisposable
{
    private readonly string _root = Path.Combine(Path.GetTempPath(), $"officeagent-outcomes-{Guid.NewGuid():N}");

    public StorageOutcomeTests() => Directory.CreateDirectory(_root);

    // ── before validation ──────────────────────────────────────────────────────────────

    [Fact]
    public async Task An_invalid_plan_writes_nothing_and_says_so()
    {
        var (tools, provider, id) = Tools(new Injector("mem"));
        using var result = await Json(tools.ApplyPlan("mem", id, BadPlan));
        Assert.Equal("notWritten", result.RootElement.GetProperty("writeOutcome").GetString());
        Assert.Equal(0, provider.Writes);
    }

    // ── before provider open ───────────────────────────────────────────────────────────

    [Fact]
    public async Task An_unknown_document_is_refused_before_any_write()
    {
        var (tools, provider, _) = Tools(new Injector("mem"));
        using var result = await Json(tools.ApplyPlan("mem", "no-such-document", EmptyPlan));
        Assert.Equal(ToolErrorCodes.NotFound, Code(result));
        Assert.Equal("notWritten", result.RootElement.GetProperty("writeOutcome").GetString());
        Assert.Equal(0, provider.Writes);
    }

    // ── between preview and commit ─────────────────────────────────────────────────────

    [Fact]
    public async Task A_source_changed_after_preview_is_refused_before_the_write()
    {
        var provider = FileSystem();
        var path = Path.Combine(_root, "contract.docx");
        await File.WriteAllBytesAsync(path, DocxFactory.Contract());
        var client = Client(provider);
        var reference = await client.RegisterAsync("fs", path);
        var plan = await PlanFor(client, reference);
        await client.PreviewAsync(reference, plan);

        // Another writer changes the file between preview and commit.
        var external = DocxFactory.Contract().Concat(new byte[] { 0 }).ToArray();
        await File.WriteAllBytesAsync(path, external);

        await Assert.ThrowsAsync<DocumentVersionConflictException>(() =>
            client.CommitAsync(reference, plan, new SaveDocumentOptions { Mode = SaveMode.Replace }));
        Assert.Equal(external, await File.ReadAllBytesAsync(path));
    }

    [Fact]
    public async Task Access_revoked_after_preview_is_refused_before_the_write()
    {
        var provider = new Injector("mem");
        var policy = new TogglePolicy();
        var (tools, _, id) = Tools(provider, policy);
        var plan = await EditPlanAsync(tools, id);
        using (var preview = await Json(tools.PreviewPlan("mem", id, plan)))
            Assert.True(preview.RootElement.GetProperty("isValid").GetBoolean());

        policy.Allowed = false;
        using var result = await Json(tools.ApplyPlan("mem", id, plan));
        Assert.Equal(ToolErrorCodes.ConnectionForbidden, Code(result));
        Assert.Equal("notWritten", result.RootElement.GetProperty("writeOutcome").GetString());
        Assert.Equal(0, provider.Writes);
    }

    // ── before the write ───────────────────────────────────────────────────────────────

    [Fact]
    public async Task A_taken_output_name_is_refused_as_already_exists()
    {
        var provider = FileSystem();
        var path = Path.Combine(_root, "contract.docx");
        var taken = DocxFactory.Contract();
        await File.WriteAllBytesAsync(path, DocxFactory.Contract());
        await File.WriteAllBytesAsync(Path.Combine(_root, "taken.docx"), taken);
        var client = Client(provider);
        var reference = await client.RegisterAsync("fs", path);

        var error = await Assert.ThrowsAsync<DocumentProviderException>(async () =>
            await client.CommitAsync(reference, await PlanFor(client, reference),
                new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "taken.docx" }));

        // Once IO, which callers had to treat as uncertain; the name was refused before any write.
        Assert.Equal(ProviderErrorCode.AlreadyExists, error.Code);
        Assert.Equal(taken, await File.ReadAllBytesAsync(Path.Combine(_root, "taken.docx")));
    }

    [Fact]
    public async Task A_memory_version_conflict_is_refused_before_the_write()
    {
        var memory = new MemoryDocumentProvider("mem");
        var reference = memory.Add("contract.docx", DocxFactory.Contract());
        var client = Client(memory);
        var plan = await PlanFor(client, reference);
        await memory.SaveAsync(reference, new MemoryStream(DocxFactory.Contract().Concat(new byte[] { 1 }).ToArray()),
            new SaveDocumentOptions { Mode = SaveMode.Replace });

        await Assert.ThrowsAsync<DocumentVersionConflictException>(() =>
            client.CommitAsync(reference, plan, new SaveDocumentOptions { Mode = SaveMode.Replace }));
    }

    // ── during the write ───────────────────────────────────────────────────────────────

    [Fact]
    public async Task A_filesystem_publish_that_fails_is_certain_and_changes_nothing()
    {
        var provider = FileSystem();
        var folder = Path.Combine(_root, "locked");
        Directory.CreateDirectory(folder);
        var path = Path.Combine(folder, "contract.docx");
        var original = DocxFactory.Contract();
        await File.WriteAllBytesAsync(path, original);
        var client = Client(provider);
        var reference = await client.RegisterAsync("fs", path);
        var plan = await PlanFor(client, reference);

        // Make the destination refuse the publish: a read-only file on Windows, a read-only
        // directory elsewhere, where the rename needs directory write access.
        if (OperatingSystem.IsWindows())
            File.SetAttributes(path, FileAttributes.ReadOnly);
        else
            File.SetUnixFileMode(folder, UnixFileMode.UserRead | UnixFileMode.UserExecute);
        try
        {
            var error = await Assert.ThrowsAsync<DocumentProviderException>(() =>
                client.CommitAsync(reference, plan, new SaveDocumentOptions { Mode = SaveMode.Replace }));
            Assert.Equal(ProviderErrorCode.WriteRejected, error.Code);
        }
        finally
        {
            if (OperatingSystem.IsWindows())
                File.SetAttributes(path, FileAttributes.Normal);
            else
                File.SetUnixFileMode(folder, UnixFileMode.UserRead | UnixFileMode.UserWrite | UnixFileMode.UserExecute);
        }
        Assert.Equal(original, await File.ReadAllBytesAsync(path));
    }

    [Fact]
    public async Task A_failure_mid_write_that_the_provider_does_not_classify_is_unknown()
    {
        var provider = new Injector("mem") { Fault = Fault.FailDuringWrite };
        var (tools, _, id) = Tools(provider);
        using var result = await Json(tools.ApplyPlan("mem", id, await EditPlanAsync(tools, id)));
        Assert.Equal(ToolErrorCodes.OutcomeUnknown, Code(result));
        Assert.Equal("unknown", result.RootElement.GetProperty("writeOutcome").GetString());
    }

    // ── immediately after the write is accepted ────────────────────────────────────────

    [Fact]
    public async Task A_failure_after_an_accepted_write_is_unknown_with_a_reconcilable_locator()
    {
        var provider = new Injector("mem") { Fault = Fault.FailAfterAccept };
        var (tools, _, id) = Tools(provider);
        using var result = await Json(tools.ApplyPlan("mem", id, await EditPlanAsync(tools, id)));

        Assert.Equal("unknown", result.RootElement.GetProperty("writeOutcome").GetString());
        Assert.False(result.RootElement.GetProperty("committed").GetBoolean()); // not confirmed, not "nothing"
        var locator = result.RootElement.GetProperty("possibleOutput");
        Assert.Equal("mem", locator.GetProperty("connectionId").GetString());
        Assert.Equal(id, locator.GetProperty("sourceDocumentId").GetString());
        // The destination holds exactly the bytes the locator names: the write landed.
        Assert.Equal(provider.HashOf(id), locator.GetProperty("expectedSha256").GetString(), ignoreCase: true);
        Assert.DoesNotContain(_root, result.RootElement.ToString());
    }

    // ── during registration ────────────────────────────────────────────────────────────

    [Fact]
    public async Task A_new_version_whose_registration_fails_is_written_not_registered()
    {
        var provider = FileSystem();
        var path = Path.Combine(_root, "contract.docx");
        await File.WriteAllBytesAsync(path, DocxFactory.Contract());
        var tools = new OfficeAgentTools(Client(provider));
        using var registered = await Json(tools.RegisterDocument("fs", "contract.docx"));
        var id = registered.RootElement.GetProperty("documentId").GetString()!;
        var plan = await EditPlanAsync(tools, id, "fs");

        // Replace the registration index with a directory, so persisting it fails.
        var index = Path.Combine(_root, ".officeagent", "index.json");
        File.Delete(index);
        Directory.CreateDirectory(index);

        using var result = await Json(tools.ApplyPlan("fs", id, plan, saveMode: "NewDocument", newName: "copy.docx"));
        Assert.Equal(ToolErrorCodes.RegistrationFailed, Code(result));
        Assert.Equal("writtenNotRegistered", result.RootElement.GetProperty("writeOutcome").GetString());
        var stored = Path.Combine(_root, "copy.docx");
        Assert.True(File.Exists(stored));

        // The locator finds the stored file: its name, and the hash of the exact bytes on disk.
        // It was null for this outcome until 1.0, so a caller could not recover the file.
        var locator = result.RootElement.GetProperty("possibleOutput");
        Assert.Equal("fs", locator.GetProperty("connectionId").GetString());
        Assert.Equal(id, locator.GetProperty("sourceDocumentId").GetString());
        Assert.Equal("copy.docx", locator.GetProperty("outputName").GetString());
        Assert.Equal(Sha256Of(stored), locator.GetProperty("expectedSha256").GetString());
        Assert.Contains("register it by name", result.RootElement.GetProperty("errors")[0].GetProperty("message").GetString());
        Assert.DoesNotContain(_root, result.RootElement.ToString());
    }

    /// <summary>
    /// The same outcome through a direct .NET call: the exception is the located type, names the
    /// file the caller asked for, and hashes exactly the bytes that were stored.
    /// </summary>
    [Fact]
    public async Task A_filesystem_new_version_registration_failure_throws_the_located_exception()
    {
        var provider = FileSystem();
        await File.WriteAllBytesAsync(Path.Combine(_root, "contract.docx"), DocxFactory.Contract());
        var reference = await provider.RegisterAsync("contract.docx");
        BreakRegistrationIndex();
        var bytes = DocxFactory.Contract();

        var error = await Assert.ThrowsAsync<DocumentRegistrationFailedException>(() => provider.SaveAsync(
            reference, new MemoryStream(bytes), new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "copy.docx" }));

        AssertLocated(error, "copy.docx", bytes, reference.ItemId);
    }

    /// <summary>A created document whose registration fails carries the same locator.</summary>
    [Fact]
    public async Task A_filesystem_create_registration_failure_throws_the_located_exception()
    {
        var provider = FileSystem();
        BreakRegistrationIndex();
        var bytes = DocxFactory.Contract();

        var error = await Assert.ThrowsAsync<DocumentRegistrationFailedException>(() =>
            provider.CreateAsync("new.docx", new MemoryStream(bytes)));

        AssertLocated(error, "new.docx", bytes, sourceItemId: null);
        Assert.Contains("Do not retry creation", error.Message);
    }

    /// <summary>The create_document tool reports the created file's locator.</summary>
    [Fact]
    public async Task The_create_tool_reports_the_locator_when_registration_fails()
    {
        var tools = new OfficeAgentTools(Client(FileSystem()));
        BreakRegistrationIndex();

        using var result = await Json(tools.CreateDocument("fs", "new.docx"));

        Assert.Equal(ToolErrorCodes.RegistrationFailed, Code(result));
        Assert.Equal("writtenNotRegistered", result.RootElement.GetProperty("writeOutcome").GetString());
        var locator = result.RootElement.GetProperty("possibleOutput");
        Assert.Equal("new.docx", locator.GetProperty("outputName").GetString());
        Assert.Equal(Sha256Of(Path.Combine(_root, "new.docx")), locator.GetProperty("expectedSha256").GetString());
        Assert.DoesNotContain(_root, result.RootElement.ToString());
    }

    /// <summary>
    /// A provider that reports a registration failure without the locator still yields one: the
    /// engine knows the name it asked for and the bytes it sent.
    /// </summary>
    [Fact]
    public async Task The_engine_locates_a_registration_failure_a_provider_reported_bare()
    {
        var provider = new Injector("mem") { ThrowCode = ProviderErrorCode.RegistrationFailed };
        var (tools, _, id) = Tools(provider);

        using var result = await Json(tools.ApplyPlan("mem", id, await EditPlanAsync(tools, id), saveMode: "NewDocument", newName: "copy.docx"));

        Assert.Equal(ToolErrorCodes.RegistrationFailed, Code(result));
        Assert.Equal("writtenNotRegistered", result.RootElement.GetProperty("writeOutcome").GetString());
        var locator = result.RootElement.GetProperty("possibleOutput");
        Assert.Equal("copy.docx", locator.GetProperty("outputName").GetString());
        Assert.Matches("^[0-9a-f]{64}$", locator.GetProperty("expectedSha256").GetString()!);
    }

    /// <summary>
    /// Giving the unknown-outcome exception a base class must not change how its frozen public
    /// constructor behaves: it never validated its arguments, so a provider that builds one keeps
    /// raising that exception rather than an <see cref="ArgumentNullException"/>. Only the new
    /// registration type, which has no earlier behaviour to keep, requires its locator.
    /// </summary>
    [Fact]
    public void The_new_base_leaves_the_frozen_constructor_unchanged_and_the_new_type_requires_its_locator()
    {
        var unknown = new DocumentWriteOutcomeUnknownException("m", "p", "c", itemId: null, outputName: null, outputSha256: null!);
        Assert.Null(unknown.OutputSha256);
        Assert.Equal(ProviderErrorCode.OutcomeUnknown, unknown.Code);

        Assert.Throws<ArgumentNullException>(() =>
            new DocumentRegistrationFailedException("m", "p", "c", null, outputName: null!, outputSha256: new string('0', 64)));
        Assert.Throws<ArgumentNullException>(() =>
            new DocumentRegistrationFailedException("m", "p", "c", null, outputName: "a.docx", outputSha256: null!));
    }

    private void BreakRegistrationIndex()
    {
        // Replace the registration index with a directory, so persisting it fails.
        var index = Path.Combine(_root, ".officeagent", "index.json");
        if (File.Exists(index)) File.Delete(index);
        Directory.CreateDirectory(index);
    }

    private void AssertLocated(DocumentRegistrationFailedException error, string name, byte[] bytes, string? sourceItemId)
    {
        Assert.Equal(ProviderErrorCode.RegistrationFailed, error.Code);
        Assert.IsAssignableFrom<DocumentWriteRecoveryException>(error);
        Assert.Equal("fs", error.ConnectionId);
        Assert.Equal(sourceItemId, error.ItemId);
        Assert.Equal(name, error.OutputName);
        var stored = Path.Combine(_root, name);
        Assert.True(File.Exists(stored));
        Assert.Equal(Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant(), error.OutputSha256);
        Assert.Equal(Sha256Of(stored), error.OutputSha256);
        Assert.DoesNotContain(_root, error.Message);
    }

    private static string Sha256Of(string path) =>
        Convert.ToHexString(SHA256.HashData(File.ReadAllBytes(path))).ToLowerInvariant();

    // ── during receipt creation ────────────────────────────────────────────────────────

    [Fact]
    public async Task A_failure_resolving_the_audit_actor_happens_before_any_write()
    {
        // The receipt, actor included, is built before the save, so a failure there can never
        // follow an accepted write. Persisting the receipt is the host's, after the call returns.
        var provider = new Injector("mem");
        var services = new ServiceCollection();
        services.AddWordFormat();
        services.AddSingleton<IDocumentProvider>(provider);
        services.AddSingleton<IAuditActorProvider>(new ThrowingActor());
        services.AddOfficeAgent();
        using var container = services.BuildServiceProvider();
        var client = container.GetRequiredService<OfficeAgentClient>();
        var reference = provider.Seed();
        await Assert.ThrowsAnyAsync<Exception>(async () =>
            await client.CommitAsync(reference, await PlanFor(client, reference), new SaveDocumentOptions { Mode = SaveMode.Replace }));
        Assert.Equal(0, provider.Writes);
    }

    // ── cancellation and timeout ───────────────────────────────────────────────────────

    [Fact]
    public async Task Cancellation_before_the_write_is_plain_and_writes_nothing()
    {
        var provider = new Injector("mem");
        var client = Client(provider);
        var reference = provider.Seed();
        var plan = await PlanFor(client, reference);
        using var cancelled = new CancellationTokenSource();
        cancelled.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(() =>
            client.CommitAsync(reference, plan, new SaveDocumentOptions { Mode = SaveMode.Replace }, cancelled.Token));
        Assert.Equal(0, provider.Writes);
    }

    [Fact]
    public async Task Cancellation_after_the_write_is_accepted_is_never_reported_as_nothing_written()
    {
        var provider = new Injector("mem") { Fault = Fault.CancelAfterAccept };
        var client = Client(provider);
        var reference = provider.Seed();
        var plan = await PlanFor(client, reference);
        using var cancellation = new CancellationTokenSource();
        provider.Cancel = cancellation;

        var error = await Assert.ThrowsAsync<DocumentWriteOutcomeUnknownException>(() =>
            client.CommitAsync(reference, plan, new SaveDocumentOptions { Mode = SaveMode.Replace }, cancellation.Token));
        Assert.IsAssignableFrom<OperationCanceledException>(error.InnerException);
        Assert.Equal(1, provider.Writes);
        Assert.Equal(provider.HashOf(reference.ItemId), error.OutputSha256, ignoreCase: true);
    }

    [Fact]
    public async Task Cancelling_a_batch_after_its_first_item_keeps_the_record_of_what_was_written()
    {
        var provider = new Injector("mem");
        var client = Client(provider);
        var template = provider.Seed(Quote());
        using var cancellation = new CancellationTokenSource();
        provider.AfterEachWrite = () => cancellation.Cancel(); // cancel once the first output lands

        var result = await client.PopulateTemplateBatchAsync(template, new TemplateBatchRequest
        {
            Items = new[] { "a.docx", "b.docx", "c.docx" }.Select(name => new TemplateBatchItem
            {
                OutputName = name,
                Binding = new TemplateBinding { Values = new Dictionary<string, string?> { ["CustomerName"] = "Fabrikam" } }
            }).ToArray()
        }, cancellationToken: cancellation.Token);

        Assert.Equal(TemplateItemOutcome.Committed, result.Items[0].Outcome);
        Assert.All(result.Items.Skip(1), item => Assert.Equal(TemplateItemOutcome.Skipped, item.Outcome));
        Assert.Equal(1, provider.Writes);
    }

    // ── fixtures ───────────────────────────────────────────────────────────────────────

    private const string BadPlan = "{\"operations\":[{\"op\":\"changeText\",\"target\":{\"paraId\":\"00000000\",\"expect\":\"absent\"},\"with\":\"x\"}]}";
    private const string EmptyPlan = "{\"operations\":[]}";

    private FileSystemDocumentProvider FileSystem() => new(new FileSystemDocumentProviderOptions
    {
        ConnectionId = "fs", RootPath = _root, DefaultChangeMode = ChangeMode.Direct
    });

    private static OfficeAgentClient Client(IDocumentProvider provider) =>
        new(new DocumentProviderRegistry(new[] { provider }), new WordModule());

    private static (OfficeAgentTools Tools, Injector Provider, string Id) Tools(Injector provider, IConnectionAccessPolicy? policy = null)
    {
        var tools = new OfficeAgentTools(Client(provider), policy);
        return (tools, provider, provider.Seed().ItemId);
    }

    private static async Task<DocumentPlan> PlanFor(OfficeAgentClient client, DocumentReference reference)
    {
        var hit = (await client.FindAsync(reference, new FindQuery("Acme Corp"))).First();
        return new DocumentPlan
        {
            Operations = new PlanOperation[] { new ChangeTextOp { Target = hit.Anchor, With = "Contoso", Mode = ChangeMode.Direct } }
        };
    }

    private static async Task<string> EditPlanAsync(OfficeAgentTools tools, string id, string connection = "mem")
    {
        using var hits = await Json(tools.FindInDocument(connection, id, "Acme Corp"));
        var hit = hits.RootElement[0];
        return "{\"operations\":[{\"op\":\"changeText\",\"target\":{\"paraId\":\"" + hit.GetProperty("paraId").GetString() +
               "\",\"expect\":\"Acme Corp\",\"occurrence\":" + hit.GetProperty("occurrence").GetInt32() +
               "},\"mode\":\"Direct\",\"with\":\"Contoso\"}]}";
    }

    private static async Task<JsonDocument> Json(Task<string> call) => JsonDocument.Parse(await call);

    /// <summary>A template with one plain-text content control, tagged CustomerName.</summary>
    private static byte[] Quote()
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            document.AddMainDocumentPart().Document = new W.Document(new W.Body(new W.Paragraph(
                new W.Run(new W.Text("Quote for ")),
                new W.SdtRun(
                    new W.SdtProperties(new W.Tag { Val = "CustomerName" }, new W.SdtId { Val = 11 }, new W.SdtContentText()),
                    new W.SdtContentRun(new W.Run(new W.Text("CUSTOMER")))))));
        }
        return stream.ToArray();
    }

    private static string? Code(JsonDocument result) =>
        result.RootElement.GetProperty("errors")[0].GetProperty("code").GetString();

    [Fact]
    public async Task A_batch_item_reports_the_same_code_a_tool_would()
    {
        // Exercises the code a batch actually emits: comparing the two mapping functions alone
        // missed a batch that stopped using its mapping. A code the provider cannot vouch for is
        // reclassified as outcome-unknown by the engine, on both surfaces.
        foreach (ProviderErrorCode code in Enum.GetValues(typeof(ProviderErrorCode)))
        {
            var provider = new Injector("mem") { ThrowCode = code };
            var client = Client(provider);
            var template = provider.Seed(Quote());
            var result = await client.PopulateTemplateBatchAsync(template, new TemplateBatchRequest
            {
                Items = new[] { new TemplateBatchItem
                {
                    OutputName = "out.docx",
                    Binding = new TemplateBinding { Values = new Dictionary<string, string?> { ["CustomerName"] = "Fabrikam" } }
                } }
            });
            var certain = StorageWrite.ProvesNothingWritten(code) ||
                          code is ProviderErrorCode.RegistrationFailed or ProviderErrorCode.OutcomeUnknown;
            var expected = OfficeAgentTools.ProviderCodeToWire(certain ? code : ProviderErrorCode.OutcomeUnknown);
            Assert.True(result.Items[0].Diagnostics.Any(d => d.Code == expected),
                $"{code}: expected {expected}, got {string.Join(", ", result.Items[0].Diagnostics.Select(d => d.Code))}");
        }
    }

    public void Dispose()
    {
        if (Directory.Exists(_root)) Directory.Delete(_root, recursive: true);
    }

    private enum Fault { None, FailDuringWrite, FailAfterAccept, CancelAfterAccept }

    /// <summary>
    /// A memory-backed provider that fails at a chosen boundary. It counts accepted writes, so a
    /// test can compare what storage did with what the caller was told.
    /// </summary>
    private sealed class Injector : IDocumentCreatingProvider
    {
        private readonly Dictionary<string, (string Name, byte[] Bytes)> _items = new(StringComparer.Ordinal);
        private int _next;

        public Injector(string connectionId) => ConnectionId = connectionId;

        public string Provider => "injector";
        public string ConnectionId { get; }
        public Fault Fault { get; init; }
        public CancellationTokenSource? Cancel { get; set; }
        public Action? AfterEachWrite { get; set; }
        public ProviderErrorCode? ThrowCode { get; set; }
        public int Writes { get; private set; }

        public DocumentReference Seed(byte[]? content = null)
        {
            var id = $"item-{++_next}";
            _items[id] = ("contract.docx", content ?? DocxFactory.Contract());
            return Reference(id);
        }

        public string HashOf(string id) => Convert.ToHexString(SHA256.HashData(_items[id].Bytes));

        public Task<DocumentReference> RegisterAsync(string source, CancellationToken cancellationToken = default) =>
            throw new DocumentProviderException(ProviderErrorCode.InvalidArgument, "Not supported.", Provider, ConnectionId);

        public Task<DocumentContent> OpenReadAsync(DocumentReference reference, CancellationToken cancellationToken = default) =>
            _items.TryGetValue(reference.ItemId, out var item)
                ? Task.FromResult(new DocumentContent(Reference(reference.ItemId), new MemoryStream(item.Bytes, writable: false)))
                : throw new DocumentProviderException(ProviderErrorCode.NotFound, "Not found.", Provider, ConnectionId, reference.ItemId);

        public async Task<DocumentReference> SaveAsync(DocumentReference source, Stream content, SaveDocumentOptions options,
            CancellationToken cancellationToken = default)
        {
            var bytes = await Read(content);
            if (ThrowCode is { } code)
                throw new DocumentProviderException(code, "Injected provider failure.", Provider, ConnectionId, source.ItemId);
            if (Fault == Fault.FailDuringWrite)
                throw new IOException("The storage connection dropped while writing.");
            var id = options.Mode == SaveMode.Replace ? source.ItemId : $"item-{++_next}";
            _items[id] = (options.NewName ?? _items[source.ItemId].Name, bytes);
            return After(id, cancellationToken);
        }

        public async Task<DocumentReference> CreateAsync(string name, Stream content, CancellationToken cancellationToken = default)
        {
            var bytes = await Read(content);
            var id = $"item-{++_next}";
            _items[id] = (name, bytes);
            return After(id, cancellationToken);
        }

        public Task RemoveAsync(DocumentReference reference, CancellationToken cancellationToken = default) => Task.CompletedTask;

        private DocumentReference After(string id, CancellationToken cancellationToken)
        {
            Writes++;
            AfterEachWrite?.Invoke();
            if (Fault == Fault.FailAfterAccept)
                throw new IOException("The response was lost after the write was stored.");
            if (Fault == Fault.CancelAfterAccept)
            {
                Cancel!.Cancel();
                cancellationToken.ThrowIfCancellationRequested();
            }
            return Reference(id);
        }

        private static async Task<byte[]> Read(Stream content)
        {
            using var buffer = new MemoryStream();
            await content.CopyToAsync(buffer);
            return buffer.ToArray();
        }

        private DocumentReference Reference(string id) => new()
        {
            Provider = Provider, ConnectionId = ConnectionId, ItemId = id,
            Version = HashOf(id), Name = _items[id].Name
        };
    }

    private sealed class TogglePolicy : IConnectionAccessPolicy
    {
        public bool Allowed { get; set; } = true;

        public ValueTask<bool> IsAllowedAsync(ClaimsPrincipal principal, string connectionId,
            ConnectionCapability capability, CancellationToken cancellationToken = default) => new(Allowed);
    }

    private sealed class ThrowingActor : IAuditActorProvider
    {
        public AuditActor? GetCurrentActor() => throw new InvalidOperationException("The identity service is unavailable.");
    }
}

/// <summary>A provider failure carries one code on every surface.</summary>
public sealed class ProviderCodeMappingTests
{
    [Fact]
    public void Tools_and_template_batches_report_every_provider_code_identically()
    {
        // Batches once derived their code by kebab-casing the enum name, so an IO failure was
        // "i-o" in a batch and "io-error" from every tool, and "i-o" was in no catalogue.
        foreach (ProviderErrorCode code in Enum.GetValues(typeof(ProviderErrorCode)))
            Assert.Equal(OfficeAgentTools.ProviderCodeToWire(code), StorageWrite.WireCode(code));
    }
}

/// <summary>Every provider refuses to save under a name it already holds.</summary>
public sealed class TakenNameTests
{
    [Fact]
    public async Task The_memory_provider_refuses_a_new_version_under_a_taken_name()
    {
        // CreateAsync already refused a taken name; a save with a new name did not, so a batch
        // into a session connection could leave two documents sharing a name, unlike storage.
        var memory = new MemoryDocumentProvider("mem");
        var source = memory.Add("contract.docx", DocxFactory.Contract());
        memory.Add("taken.docx", DocxFactory.Contract());

        var error = await Assert.ThrowsAsync<DocumentProviderException>(() => memory.SaveAsync(source,
            new MemoryStream(DocxFactory.Contract()), new SaveDocumentOptions { Mode = SaveMode.NewDocument, NewName = "taken.docx" }));
        Assert.Equal(ProviderErrorCode.AlreadyExists, error.Code);
    }
}
