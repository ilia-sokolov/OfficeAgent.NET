using System.Security.Claims;
using System.Security.Cryptography;
using System.Text.Json;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// What happens when two callers arrive at once: who wins a contested write, whether
/// unrelated work serializes behind them, whether one caller can reach another's
/// connection, whose identity lands on each receipt, and what a caller can conclude
/// about storage after a provider fault.
/// </summary>
/// <remarks>
/// Every race here is driven by an explicit barrier rather than by timing, so a failure
/// is a real ordering defect and not a slow machine. Provider behavior is exercised
/// through controlled fakes and the local filesystem provider; no live SharePoint
/// tenant is involved and none of these results describe one.
/// </remarks>
public sealed class ConcurrencyAndIsolationTests
{
    // ── Contested writes ─────────────────────────────────────────────────

    /// <summary>
    /// Two callers read the same version and both commit against it. Exactly one wins;
    /// the loser is told its version is stale rather than silently replacing the winner.
    /// </summary>
    [Fact]
    public async Task Two_callers_committing_from_one_version_produce_one_winner_and_one_conflict()
    {
        using var workspace = new TemporaryWorkspace();
        var path = Path.Combine(workspace.Root, "contract.docx");
        await File.WriteAllBytesAsync(path, DocxFactory.Contract());

        var client = workspace.Client();
        var reference = await client.RegisterAsync("workspace", path);

        string sharedVersion;
        using (var content = await client.OpenReadAsync(reference))
            sharedVersion = content.Reference.Version!;

        var unpinned = DocumentReference.ForFileSystem("workspace", reference.ItemId);

        // Both callers are held until both have their plan ready, so neither can finish
        // before the other starts. This rendezvous is awaited rather than blocked on: a
        // blocking barrier inside async work deadlocks under thread-pool pressure.
        var bothReady = new AsyncRendezvous(2);

        async Task<Outcome> Commit(string replacement)
        {
            var inspection = await client.InspectAsync(unpinned);
            var hit = (await client.FindAsync(unpinned, new FindQuery("Acme Corp"))).First();
            var plan = new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new ChangeTextOp { Target = hit.Anchor, With = replacement, Mode = ChangeMode.Direct }
                }
            };

            await bothReady.ArriveAsync();

            try
            {
                var result = await client.CommitAsync(unpinned, plan, new SaveDocumentOptions
                {
                    Mode = SaveMode.Replace,
                    ExpectedVersion = sharedVersion
                });
                return result.Committed ? Outcome.Won : Outcome.Rejected;
            }
            catch (DocumentVersionConflictException)
            {
                return Outcome.Conflicted;
            }
        }

        var results = await Task.WhenAll(Commit("Contoso Research"), Commit("Fabrikam Services"));

        Assert.Equal(1, results.Count(outcome => outcome == Outcome.Won));
        Assert.Equal(1, results.Count(outcome => outcome == Outcome.Conflicted));

        // The winner's text is what is on disk, whole. A lost update would leave the
        // file at the loser's content or at a mix of both.
        var final = await File.ReadAllBytesAsync(path);
        var text = string.Concat(client.Inspect(final).Paragraphs.Select(p => p.Text));
        Assert.True(
            text.Contains("Contoso Research", StringComparison.Ordinal) ^
            text.Contains("Fabrikam Services", StringComparison.Ordinal),
            $"exactly one winner expected, got: {text}");
    }

    /// <summary>
    /// The session provider has the same optimistic-concurrency contract as persistent
    /// providers. Both reads are held until they have observed the shared version, making
    /// this a deterministic lost-update test rather than a timing-dependent stress test.
    /// </summary>
    [Fact]
    public async Task Two_memory_saves_from_one_version_produce_one_winner_and_one_conflict()
    {
        var provider = new MemoryDocumentProvider("session");
        var original = provider.Add("contract.docx", new byte[] { 0 });
        var bothReading = new AsyncRendezvous(2);

        async Task<(bool Won, byte Value)> Save(byte value)
        {
            await using var content = new RendezvousReadStream(new[] { value }, bothReading);
            try
            {
                await provider.SaveAsync(original, content, new SaveDocumentOptions
                {
                    Mode = SaveMode.Replace,
                    ExpectedVersion = original.Version
                });
                return (true, value);
            }
            catch (DocumentVersionConflictException)
            {
                return (false, value);
            }
        }

        var outcomes = await Task.WhenAll(Save(1), Save(2));
        var winner = Assert.Single(outcomes, outcome => outcome.Won);
        Assert.Single(outcomes, outcome => !outcome.Won);
        Assert.Equal(new[] { winner.Value }, provider.Read(original.ItemId));
    }

    /// <summary>The provider owns stored bytes and never exposes a mutable backing array.</summary>
    [Fact]
    public void Memory_documents_cannot_be_changed_without_a_versioned_save()
    {
        var provider = new MemoryDocumentProvider("session");
        var supplied = new byte[] { 1, 2, 3 };
        var reference = provider.Add("contract.docx", supplied);

        supplied[0] = 9;
        var returned = provider.Read(reference.ItemId);
        returned[1] = 9;

        Assert.Equal(new byte[] { 1, 2, 3 }, provider.Read(reference.ItemId));
        Assert.Equal(reference.Version, provider.Describe(reference.ItemId).Version);
    }

    /// <summary>
    /// Edits to unrelated documents run at the same time. The barrier only releases once
    /// both callers are inside their commit, so a global lock would deadlock here rather
    /// than merely run slowly.
    /// </summary>
    [Fact]
    public async Task Unrelated_documents_are_not_serialized_behind_one_another()
    {
        using var workspace = new TemporaryWorkspace();
        var first = Path.Combine(workspace.Root, "first.docx");
        var second = Path.Combine(workspace.Root, "second.docx");
        await File.WriteAllBytesAsync(first, DocxFactory.Contract());
        await File.WriteAllBytesAsync(second, DocxFactory.Contract());

        var gate = new ObservableGate(participants: 2);
        var client = workspace.Client(gate);

        var a = await client.RegisterAsync("workspace", first);
        var b = await client.RegisterAsync("workspace", second);

        async Task Edit(DocumentReference reference, string replacement)
        {
            var hit = (await client.FindAsync(reference, new FindQuery("Acme Corp"))).First();
            await client.CommitAsync(reference, new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new ChangeTextOp { Target = hit.Anchor, With = replacement, Mode = ChangeMode.Direct }
                }
            }, new SaveDocumentOptions { Mode = SaveMode.Replace });
        }

        var work = Task.WhenAll(Edit(a, "Contoso Research"), Edit(b, "Fabrikam Services"));
        var finished = await Task.WhenAny(work, Task.Delay(TimeSpan.FromSeconds(30)));

        Assert.True(ReferenceEquals(finished, work),
            "two edits to different documents did not overlap; they are serialized behind a shared lock");
        await work;
        Assert.Equal(2, gate.PeakConcurrency);
    }

    private enum Outcome { Won, Rejected, Conflicted }

    /// <summary>
    /// An awaitable meeting point. Every participant completes only once all of them
    /// have arrived, without blocking a thread while it waits.
    /// </summary>
    private sealed class AsyncRendezvous
    {
        private readonly TaskCompletionSource<bool> _all =
            new(TaskCreationOptions.RunContinuationsAsynchronously);
        private readonly int _participants;
        private int _arrived;

        public AsyncRendezvous(int participants) => _participants = participants;

        public Task ArriveAsync()
        {
            if (Interlocked.Increment(ref _arrived) >= _participants)
                _all.TrySetResult(true);
            return _all.Task;
        }
    }

    private sealed class RendezvousReadStream : Stream
    {
        private readonly MemoryStream _content;
        private readonly AsyncRendezvous _rendezvous;
        private bool _arrived;

        public RendezvousReadStream(byte[] content, AsyncRendezvous rendezvous)
        {
            _content = new MemoryStream(content, writable: false);
            _rendezvous = rendezvous;
        }

        public override bool CanRead => true;
        public override bool CanSeek => false;
        public override bool CanWrite => false;
        public override long Length => throw new NotSupportedException();
        public override long Position
        {
            get => throw new NotSupportedException();
            set => throw new NotSupportedException();
        }

        public override int Read(byte[] buffer, int offset, int count) =>
            throw new NotSupportedException("The provider must use asynchronous reads.");

        public override async Task<int> ReadAsync(
            byte[] buffer,
            int offset,
            int count,
            CancellationToken cancellationToken)
        {
            if (!_arrived)
            {
                _arrived = true;
                await _rendezvous.ArriveAsync().WaitAsync(cancellationToken);
            }

            return await _content.ReadAsync(buffer, offset, count, cancellationToken);
        }

        public override void Flush() { }
        public override long Seek(long offset, SeekOrigin origin) => throw new NotSupportedException();
        public override void SetLength(long value) => throw new NotSupportedException();
        public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();

        protected override void Dispose(bool disposing)
        {
            if (disposing) _content.Dispose();
            base.Dispose(disposing);
        }

        public override async ValueTask DisposeAsync()
        {
            await _content.DisposeAsync();
            await base.DisposeAsync();
        }
    }

    // ── Caller isolation ─────────────────────────────────────────────────

    /// <summary>
    /// A caller who may use connection A cannot reach connection B by naming it, even
    /// with a document id that really exists there. The denial is the same whether the
    /// id exists or not, so the refusal itself reveals nothing.
    /// </summary>
    [Fact]
    public async Task A_guessed_id_in_a_forbidden_connection_is_denied_without_revealing_it()
    {
        var alice = new MemoryDocumentProvider("alice");
        var bob = new MemoryDocumentProvider("bob");
        var real = bob.Add("secret-plan.docx", DocxFactory.Contract());

        var tools = Tools(new[] { alice, bob }, allowed: "alice");

        var existing = await tools.InspectDocument("bob", real.ItemId);
        var invented = await tools.InspectDocument("bob", "there-is-no-such-document");

        var existingCode = ErrorCode(existing);
        Assert.Equal("connection-forbidden", existingCode);
        Assert.Equal(existingCode, ErrorCode(invented));
        Assert.Equal(ErrorMessage(existing), ErrorMessage(invented));

        // Nothing about the document, its name, or its contents escapes.
        Assert.DoesNotContain("secret-plan", ErrorMessage(existing), StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain(real.ItemId, ErrorMessage(existing), StringComparison.Ordinal);
    }

    /// <summary>
    /// Connection discovery lists only what the caller may actually use. The listing
    /// filters on this capability check, so proving the check is the point.
    /// </summary>
    [Fact]
    public async Task Connection_discovery_reports_only_permitted_connections()
    {
        var alice = new MemoryDocumentProvider("alice");
        var bob = new MemoryDocumentProvider("bob");
        var tools = Tools(new[] { alice, bob }, allowed: "alice");

        foreach (var capability in Enum.GetValues<ConnectionCapability>())
        {
            Assert.True(await tools.CanAccessConnectionAsync("alice", capability));
            Assert.False(await tools.CanAccessConnectionAsync("bob", capability));
        }
    }

    /// <summary>
    /// A denial never carries the principal's claims. An error that echoed the caller's
    /// identity back into a model-visible payload would leak it by construction.
    /// </summary>
    [Fact]
    public async Task A_denial_does_not_carry_the_callers_claims()
    {
        var bob = new MemoryDocumentProvider("bob");
        var reference = bob.Add("secret.docx", DocxFactory.Contract());

        var principal = new ClaimsPrincipal(new ClaimsIdentity(new[]
        {
            new Claim(ClaimTypes.Name, "alice@contoso.example"),
            new Claim("access_token", "sensitive-token-value"),
            new Claim("oid", "11111111-2222-3333-4444-555555555555")
        }, "test"));

        var tools = Tools(new[] { bob }, allowed: "nothing", principal: principal);
        var denied = await tools.InspectDocument("bob", reference.ItemId);

        foreach (var secret in new[] { "sensitive-token-value", "11111111-2222", "alice@contoso.example" })
            Assert.DoesNotContain(secret, denied, StringComparison.OrdinalIgnoreCase);
    }

    // ── Audit identity ───────────────────────────────────────────────────

    /// <summary>
    /// Concurrent commits each record their own trusted actor. A shared per-request field
    /// would let one caller's identity land on another caller's receipt.
    /// </summary>
    [Fact]
    public async Task Concurrent_commits_each_record_their_own_actor()
    {
        using var workspace = new TemporaryWorkspace();
        var client = workspace.Client();

        var references = new List<DocumentReference>();
        for (int i = 0; i < 6; i++)
        {
            var path = Path.Combine(workspace.Root, $"doc{i}.docx");
            await File.WriteAllBytesAsync(path, DocxFactory.Contract());
            references.Add(await client.RegisterAsync("workspace", path));
        }

        async Task<(string Expected, string? Recorded)> Commit(DocumentReference reference, int index)
        {
            var actor = new AuditActor { Subject = $"user-{index}", DisplayName = $"User {index}" };
            var hit = (await client.FindAsync(reference, new FindQuery("Acme Corp"))).First();

            var result = await client.CommitAsync(reference, new DocumentPlan
            {
                // The displayed Word author is deliberately the same for every caller:
                // it is document metadata and must not influence the audited actor.
                Revision = new RevisionMetadata { Author = "Shared Review Bot" },
                Operations = new PlanOperation[]
                {
                    new ChangeTextOp { Target = hit.Anchor, With = $"Client {index}", Mode = ChangeMode.Tracked }
                }
            }, new SaveDocumentOptions { Mode = SaveMode.Replace, Actor = actor });

            return ($"user-{index}", result.Receipt?.Actor?.Subject);
        }

        var results = await Task.WhenAll(references.Select((reference, i) => Commit(reference, i)));

        foreach (var (expected, recorded) in results)
            Assert.Equal(expected, recorded);

        // Every receipt is distinct: no identity was reused across requests.
        Assert.Equal(results.Length, results.Select(r => r.Recorded).Distinct().Count());
    }

    /// <summary>
    /// The displayed Word revision author is document metadata and is not the audited
    /// actor. They are recorded separately and are allowed to disagree.
    /// </summary>
    [Fact]
    public async Task The_displayed_author_never_becomes_the_audited_actor()
    {
        using var workspace = new TemporaryWorkspace();
        var path = Path.Combine(workspace.Root, "contract.docx");
        await File.WriteAllBytesAsync(path, DocxFactory.Contract());

        var client = workspace.Client();
        var reference = await client.RegisterAsync("workspace", path);
        var hit = (await client.FindAsync(reference, new FindQuery("Acme Corp"))).First();

        var result = await client.CommitAsync(reference, new DocumentPlan
        {
            Revision = new RevisionMetadata { Author = "Totally Legitimate Auditor" },
            Operations = new PlanOperation[]
            {
                new ChangeTextOp { Target = hit.Anchor, With = "Contoso Research", Mode = ChangeMode.Tracked }
            }
        }, new SaveDocumentOptions { Mode = SaveMode.Replace, Actor = new AuditActor { Subject = "host-verified-user" } });

        Assert.True(result.Committed);
        Assert.Equal("host-verified-user", result.Receipt?.Actor?.Subject);
        Assert.NotEqual("Totally Legitimate Auditor", result.Receipt?.Actor?.Subject);
    }

    // ── Fault classification ─────────────────────────────────────────────

    /// <summary>
    /// A fault before the provider accepts the bytes leaves storage untouched, and the
    /// caller can say so.
    /// </summary>
    [Fact]
    public async Task A_fault_before_save_leaves_nothing_written()
    {
        var provider = new FaultingProvider("faulty", FaultPoint.BeforeSave);
        var reference = provider.Add("contract.docx", DocxFactory.Contract());
        var before = provider.ContentHash(reference.ItemId);
        var client = new OfficeAgentClient(new DocumentProviderRegistry(new[] { provider }), new WordModule());

        var hit = (await client.FindAsync(reference, new FindQuery("Acme Corp"))).First();
        await Assert.ThrowsAsync<DocumentProviderException>(() => client.CommitAsync(reference, new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new ChangeTextOp { Target = hit.Anchor, With = "Contoso Research", Mode = ChangeMode.Direct }
            }
        }, new SaveDocumentOptions { Mode = SaveMode.Replace }));

        Assert.Equal(before, provider.ContentHash(reference.ItemId));
        Assert.False(provider.AcceptedWrite);
    }

    /// <summary>
    /// A fault after the provider has accepted the bytes is the uncertain case. The
    /// write did land, so a caller must not be told nothing happened, and must not
    /// blindly retry. This test fixes that the bytes are observably present even though
    /// the call failed, which is exactly what makes the outcome unknown to the caller.
    /// </summary>
    [Fact]
    public async Task A_fault_after_an_accepted_save_leaves_a_write_the_caller_cannot_disown()
    {
        var provider = new FaultingProvider("faulty", FaultPoint.AfterSaveAccepted);
        var reference = provider.Add("contract.docx", DocxFactory.Contract());
        var before = provider.ContentHash(reference.ItemId);
        var client = new OfficeAgentClient(new DocumentProviderRegistry(new[] { provider }), new WordModule());

        var hit = (await client.FindAsync(reference, new FindQuery("Acme Corp"))).First();
        var failure = await Assert.ThrowsAsync<DocumentProviderException>(
            () => client.CommitAsync(reference, new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new ChangeTextOp { Target = hit.Anchor, With = "Contoso Research", Mode = ChangeMode.Direct }
                }
            }, new SaveDocumentOptions { Mode = SaveMode.Replace }));

        Assert.True(provider.AcceptedWrite, "the provider accepted the bytes before faulting");
        Assert.NotEqual(before, provider.ContentHash(reference.ItemId));

        // The failure must not assert that nothing was written, because something was.
        Assert.DoesNotContain("nothing was changed", failure.Message, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("no write", failure.Message, StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// A plan that fails validation never reaches the provider at all. This is the one
    /// case where a caller may safely conclude nothing was written.
    /// </summary>
    [Fact]
    public async Task An_invalid_plan_never_reaches_the_provider()
    {
        var provider = new FaultingProvider("faulty", FaultPoint.None);
        var reference = provider.Add("contract.docx", DocxFactory.Contract());
        var before = provider.ContentHash(reference.ItemId);
        var client = new OfficeAgentClient(new DocumentProviderRegistry(new[] { provider }), new WordModule());

        var result = await client.CommitAsync(reference, new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new ChangeTextOp
                {
                    Target = new TextSpanAnchor { ParaId = "no-such-paragraph", Expect = "whatever" },
                    With = "Contoso Research"
                }
            }
        }, new SaveDocumentOptions { Mode = SaveMode.Replace });

        Assert.False(result.Committed);
        Assert.False(provider.AcceptedWrite);
        Assert.Equal(before, provider.ContentHash(reference.ItemId));
        Assert.Contains(result.Report.Errors, error => error.Code == ValidationErrorCodes.AnchorNotFound);
    }

    /// <summary>A cancelled commit is reported as cancellation, not as success.</summary>
    [Fact]
    public async Task A_cancelled_commit_does_not_report_success()
    {
        using var workspace = new TemporaryWorkspace();
        var path = Path.Combine(workspace.Root, "contract.docx");
        await File.WriteAllBytesAsync(path, DocxFactory.Contract());

        var client = workspace.Client();
        var reference = await client.RegisterAsync("workspace", path);
        var hit = (await client.FindAsync(reference, new FindQuery("Acme Corp"))).First();

        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();

        await Assert.ThrowsAnyAsync<OperationCanceledException>(
            () => client.CommitAsync(reference, new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new ChangeTextOp { Target = hit.Anchor, With = "Contoso Research", Mode = ChangeMode.Direct }
                }
            }, new SaveDocumentOptions { Mode = SaveMode.Replace }, cancellation.Token));
    }

    // ── Helpers ──────────────────────────────────────────────────────────

    private static OfficeAgentTools Tools(
        IEnumerable<IDocumentProvider> providers, string allowed, ClaimsPrincipal? principal = null) =>
        new(new OfficeAgentClient(new DocumentProviderRegistry(providers), new WordModule()),
            new SingleConnectionPolicy(allowed),
            new FixedPrincipal(principal ?? new ClaimsPrincipal(new ClaimsIdentity("test"))));

    private static string? ErrorCode(string payload)
    {
        using var document = JsonDocument.Parse(payload);
        return document.RootElement.GetProperty("errors")[0].GetProperty("code").GetString();
    }

    private static string ErrorMessage(string payload)
    {
        using var document = JsonDocument.Parse(payload);
        return document.RootElement.GetProperty("errors")[0].GetProperty("message").GetString() ?? string.Empty;
    }

    private sealed class SingleConnectionPolicy : IConnectionAccessPolicy
    {
        private readonly string _allowed;

        public SingleConnectionPolicy(string allowed) => _allowed = allowed;

        public ValueTask<bool> IsAllowedAsync(
            ClaimsPrincipal principal,
            string connectionId,
            ConnectionCapability capability,
            CancellationToken cancellationToken = default) =>
            new(string.Equals(connectionId, _allowed, StringComparison.Ordinal));
    }

    private sealed class FixedPrincipal : ITrustedPrincipalAccessor
    {
        public FixedPrincipal(ClaimsPrincipal principal) => Principal = principal;

        public ClaimsPrincipal Principal { get; }
    }

    /// <summary>Records how many callers were inside a save at the same moment.</summary>
    private sealed class ObservableGate
    {
        private readonly int _participants;
        private readonly object _lock = new();
        private readonly TaskCompletionSource<bool> _all =
            new(TaskCreationOptions.RunContinuationsAsynchronously);
        private int _inside;

        public ObservableGate(int participants) => _participants = participants;

        public int PeakConcurrency { get; private set; }

        public async Task EnterAsync(CancellationToken cancellationToken)
        {
            lock (_lock)
            {
                _inside++;
                PeakConcurrency = Math.Max(PeakConcurrency, _inside);
                if (_inside >= _participants) _all.TrySetResult(true);
            }

            // Wait for the others, but never forever: if the provider serialized these
            // saves, the wait expires and the test's assertion reports the serialization
            // instead of the run hanging.
            await Task.WhenAny(_all.Task, Task.Delay(TimeSpan.FromSeconds(10), cancellationToken))
                .ConfigureAwait(false);

            lock (_lock) { _inside--; }
        }
    }

    private enum FaultPoint { None, BeforeSave, AfterSaveAccepted }

    /// <summary>
    /// A provider that can fail at a chosen point, so a test can tell the difference
    /// between a write that never happened and one that did.
    /// </summary>
    private sealed class FaultingProvider : IDocumentProvider
    {
        private readonly FaultPoint _fault;
        private readonly Dictionary<string, byte[]> _items = new(StringComparer.Ordinal);
        private int _next;

        public FaultingProvider(string connectionId, FaultPoint fault)
        {
            ConnectionId = connectionId;
            _fault = fault;
        }

        public string Provider => "faulting";

        public string ConnectionId { get; }

        public bool AcceptedWrite { get; private set; }

        public DocumentReference Add(string name, byte[] content)
        {
            var id = $"item-{++_next}";
            _items[id] = content;
            return DocumentReference.For(Provider, ConnectionId, id, Hash(content));
        }

        public string ContentHash(string itemId) => Hash(_items[itemId]);

        public Task<DocumentReference> RegisterAsync(string source, CancellationToken cancellationToken = default) =>
            throw new NotSupportedException();

        public Task<DocumentContent> OpenReadAsync(
            DocumentReference reference, CancellationToken cancellationToken = default)
        {
            var bytes = _items[reference.ItemId];
            return Task.FromResult(new DocumentContent(
                DocumentReference.For(Provider, ConnectionId, reference.ItemId, Hash(bytes)),
                new MemoryStream(bytes, writable: false)));
        }

        public async Task<DocumentReference> SaveAsync(
            DocumentReference source,
            Stream content,
            SaveDocumentOptions options,
            CancellationToken cancellationToken = default)
        {
            if (_fault == FaultPoint.BeforeSave)
                throw new DocumentProviderException(
                    ProviderErrorCode.IO, "The storage refused the write.", Provider, ConnectionId, source.ItemId);

            using var buffer = new MemoryStream();
            await content.CopyToAsync(buffer, cancellationToken).ConfigureAwait(false);
            _items[source.ItemId] = buffer.ToArray();
            AcceptedWrite = true;

            if (_fault == FaultPoint.AfterSaveAccepted)
                throw new DocumentProviderException(
                    ProviderErrorCode.IO,
                    "The storage accepted the document but the result could not be confirmed.",
                    Provider, ConnectionId, source.ItemId);

            return DocumentReference.For(Provider, ConnectionId, source.ItemId, Hash(_items[source.ItemId]));
        }

        public Task RemoveAsync(DocumentReference reference, CancellationToken cancellationToken = default)
        {
            _items.Remove(reference.ItemId);
            return Task.CompletedTask;
        }

        private static string Hash(byte[] content) => Convert.ToHexString(SHA256.HashData(content));
    }

    private sealed class TemporaryWorkspace : IDisposable
    {
        private readonly ObservableGate? _gate;

        public TemporaryWorkspace()
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-concurrency-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);
        }

        public string Root { get; }

        public OfficeAgentClient Client(ObservableGate? gate = null)
        {
            IDocumentProvider provider = new FileSystemDocumentProvider(new FileSystemDocumentProviderOptions
            {
                ConnectionId = "workspace",
                RootPath = Root,
                DefaultChangeMode = ChangeMode.Direct
            });

            if (gate is not null) provider = new GatedProvider(provider, gate);

            return new OfficeAgentClient(new DocumentProviderRegistry(new[] { provider }), new WordModule());
        }

        public void Dispose()
        {
            if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true);
        }
    }

    /// <summary>Wraps a provider so a test can observe how many saves overlap.</summary>
    private sealed class GatedProvider : IDocumentProvider
    {
        private readonly IDocumentProvider _inner;
        private readonly ObservableGate _gate;

        public GatedProvider(IDocumentProvider inner, ObservableGate gate)
        {
            _inner = inner;
            _gate = gate;
        }

        public string Provider => _inner.Provider;

        public string ConnectionId => _inner.ConnectionId;

        public Task<DocumentReference> RegisterAsync(string source, CancellationToken cancellationToken = default) =>
            _inner.RegisterAsync(source, cancellationToken);

        public Task<DocumentContent> OpenReadAsync(
            DocumentReference reference, CancellationToken cancellationToken = default) =>
            _inner.OpenReadAsync(reference, cancellationToken);

        public async Task<DocumentReference> SaveAsync(
            DocumentReference source,
            Stream content,
            SaveDocumentOptions options,
            CancellationToken cancellationToken = default)
        {
            await _gate.EnterAsync(cancellationToken).ConfigureAwait(false);
            return await _inner.SaveAsync(source, content, options, cancellationToken).ConfigureAwait(false);
        }

        public Task RemoveAsync(DocumentReference reference, CancellationToken cancellationToken = default) =>
            _inner.RemoveAsync(reference, cancellationToken);
    }
}
