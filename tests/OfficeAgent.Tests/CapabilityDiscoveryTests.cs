using System.Security.Claims;
using System.Text.Json;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Excel;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Tests;

/// <summary>
/// Discovery has to agree with the engine it describes. These tests reconcile what is
/// advertised against what validation actually does, so a claim cannot drift into
/// something an agent would act on and then be refused for.
/// </summary>
public sealed class CapabilityDiscoveryTests
{
    // ── Reconciliation with real validation ──────────────────────────────

    /// <summary>
    /// Every verb advertised for Word is accepted by a Word handler, and every verb in
    /// the wire vocabulary that Word does not advertise is refused with
    /// <c>unsupported-operation</c> when actually sent. Advertising and refusing are the
    /// same decision, taken once.
    /// </summary>
    [Fact]
    public void Advertised_word_verbs_are_exactly_the_verbs_word_accepts()
    {
        var module = new WordModule();
        var advertised = Describe(module).Operations.ToHashSet(StringComparer.Ordinal);

        Assert.NotEmpty(advertised);

        foreach (var verb in PlanOperationJsonConverter.ByVerb.Keys)
        {
            bool accepted = AnyHandlerAccepts(module, verb);
            Assert.True(
                advertised.Contains(verb) == accepted,
                $"'{verb}' is advertised={advertised.Contains(verb)} but accepted={accepted}");
        }
    }

    /// <summary>The same reconciliation for the deck and the workbook.</summary>
    [Theory]
    [InlineData("powerpoint")]
    [InlineData("excel")]
    public void Advertised_verbs_match_accepted_verbs_for_every_format(string format)
    {
        IFormatModule module = format == "powerpoint" ? new PowerPointModule() : new ExcelModule();
        var advertised = Describe(module).Operations.ToHashSet(StringComparer.Ordinal);

        Assert.NotEmpty(advertised);

        foreach (var verb in PlanOperationJsonConverter.ByVerb.Keys)
            Assert.Equal(advertised.Contains(verb), AnyHandlerAccepts(module, verb));
    }

    /// <summary>
    /// A verb a format does not advertise really is refused end to end, not merely absent
    /// from a list. This is the claim an agent relies on when it reads discovery.
    /// </summary>
    [Fact]
    public void A_verb_absent_from_a_formats_list_is_refused_when_sent()
    {
        var client = new OfficeAgentClient(new WordModule());
        var word = Describe(new WordModule());

        Assert.DoesNotContain("setCell", word.Operations);

        var report = client.Preview(Handle(DocxFactory.Contract()), new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new SetCellOp { Target = new CellAnchor { SheetId = 1, Address = "A1" }, Value = "nope" }
            }
        });

        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, error => error.Code == ValidationErrorCodes.UnsupportedOperation);
    }

    /// <summary>
    /// The declared change modes match what plan validation enforces. PowerPoint declares
    /// Direct only, and a tracked deck plan is genuinely refused; Word declares both, and
    /// a tracked Word plan is genuinely accepted.
    /// </summary>
    [Fact]
    public void Declared_change_modes_match_what_validation_enforces()
    {
        var deck = Describe(new PowerPointModule());
        Assert.Equal(new[] { ChangeMode.Direct }, deck.ChangeModes);

        var deckClient = new OfficeAgentClient(new PowerPointModule());
        var deckReport = deckClient.Preview(Handle(PptxFactory.Deck()), new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new ChangeTextOp
                {
                    Target = new TextSpanAnchor { ParaId = "slide256/shape2/p0", Expect = "" },
                    With = "Tracked is not available here",
                    Mode = ChangeMode.Tracked
                }
            }
        });
        Assert.False(deckReport.IsValid);
        Assert.Contains(deckReport.Errors, error =>
            error.Message.Contains("Tracked", StringComparison.Ordinal));

        var word = Describe(new WordModule());
        Assert.Contains(ChangeMode.Tracked, word.ChangeModes);

        var wordClient = new OfficeAgentClient(new WordModule());
        var hit = wordClient.Find(Handle(DocxFactory.Contract()), new FindQuery("Acme Corp")).First();
        var wordReport = wordClient.Preview(Handle(DocxFactory.Contract()), new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new ChangeTextOp { Target = hit.Anchor, With = "Contoso", Mode = ChangeMode.Tracked }
            }
        });
        Assert.True(wordReport.IsValid);
    }

    /// <summary>The advertised contract version is the one enforcement accepts.</summary>
    [Fact]
    public void The_advertised_contract_version_is_the_one_enforced()
    {
        var capabilities = CapabilityDiscovery.Describe(new[] { new WordModule() });
        Assert.Equal(DocumentPlan.CurrentContractVersion, capabilities.Contracts.EditPlan);

        var client = new OfficeAgentClient(new WordModule());

        var matching = new DocumentPlan { ContractVersion = capabilities.Contracts.EditPlan };
        Assert.True(client.Preview(Handle(DocxFactory.Contract()), matching).IsValid);

        var mismatched = new DocumentPlan { ContractVersion = "9.9" };
        var report = client.Preview(Handle(DocxFactory.Contract()), mismatched);
        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, error => error.Code == ValidationErrorCodes.ContractMismatch);
    }

    // ── Registration-derived, not hard-coded ─────────────────────────────

    /// <summary>An unregistered format contributes nothing to discovery.</summary>
    [Fact]
    public void Only_registered_formats_appear()
    {
        var wordOnly = CapabilityDiscovery.Describe(new[] { new WordModule() });
        Assert.Equal(new[] { DocFormat.Word }, wordOnly.Formats.Select(f => f.Format));

        var all = CapabilityDiscovery.Describe(
            new IFormatModule[] { new WordModule(), new PowerPointModule(), new ExcelModule() });
        Assert.Equal(3, all.Formats.Count);
        Assert.Contains(all.Formats, f => f.Format == DocFormat.PowerPoint);
    }

    /// <summary>The reported ceilings are the ones the engine was configured with.</summary>
    [Fact]
    public void Reported_limits_are_the_configured_ceilings()
    {
        var limits = new OpenXmlIngestionLimits { MaximumCompressedBytes = 12_345, MaximumParts = 7 };
        var client = new OfficeAgentClient(
            new OfficeAgentEngine(new[] { new WordModule() }, limits: limits));

        var reported = client.DescribeCapabilities().Limits;

        Assert.Equal(12_345, reported.MaximumCompressedBytes);
        Assert.Equal(7, reported.MaximumParts);
    }

    /// <summary>Node kinds come from the module's own providers.</summary>
    [Fact]
    public void Node_kinds_come_from_the_modules_providers()
    {
        var word = Describe(new WordModule());

        Assert.Contains("revision", word.NodeKinds);
        Assert.Contains("table", word.NodeKinds);
        Assert.DoesNotContain("slide", word.NodeKinds);

        var deck = Describe(new PowerPointModule());
        Assert.Contains("slide", deck.NodeKinds);
        Assert.DoesNotContain("revision", deck.NodeKinds);
    }

    /// <summary>
    /// Discovery says out loud what it cannot answer, so silence is not read as support.
    /// </summary>
    [Fact]
    public void Discovery_names_what_requires_inspection()
    {
        var capabilities = CapabilityDiscovery.Describe(new[] { new WordModule() });

        Assert.NotEmpty(capabilities.RequiresInspection);
        Assert.Contains(capabilities.RequiresInspection, note =>
            note.Contains("Anchor", StringComparison.OrdinalIgnoreCase));
    }

    // ── Policy is not bypassed ───────────────────────────────────────────

    /// <summary>
    /// Discovery through the adapter shows only connections the caller may use, with only
    /// the capabilities they hold. A connection they cannot touch is absent, not denied,
    /// so its existence is not disclosed.
    /// </summary>
    [Fact]
    public async Task Discovery_shows_only_the_connections_this_caller_may_use()
    {
        var tools = Tools(allowed: "alice");

        var capabilities = await tools.DescribeCapabilitiesAsync();

        var connection = Assert.Single(capabilities.Connections);
        Assert.Equal("alice", connection.ConnectionId);
        Assert.NotEmpty(connection.Allowed);

        var payload = await tools.DescribeCapabilities();
        Assert.DoesNotContain("bob", payload, StringComparison.Ordinal);
    }

    /// <summary>A caller allowed nothing sees no connections at all, and no error.</summary>
    [Fact]
    public async Task A_caller_with_no_access_sees_no_connections()
    {
        var tools = Tools(allowed: "nothing-at-all");

        var capabilities = await tools.DescribeCapabilitiesAsync();

        Assert.Empty(capabilities.Connections);
        // The engine half is still described: what the server supports is not a secret.
        Assert.NotEmpty(capabilities.Formats);
    }

    /// <summary>
    /// Partial access is reported partially. A caller with read but not edit sees the
    /// connection with exactly the capability they hold.
    /// </summary>
    [Fact]
    public async Task Partial_access_is_reported_as_partial()
    {
        var tools = Tools(new ReadOnlyOnAlice());

        var connection = Assert.Single((await tools.DescribeCapabilitiesAsync()).Connections);

        Assert.Equal("alice", connection.ConnectionId);
        Assert.Equal(new[] { ConnectionCapability.Read.ToString() }, connection.Allowed);
    }

    /// <summary>The tool is part of the model-facing surface and serializes as JSON.</summary>
    [Fact]
    public async Task The_discovery_tool_is_offered_and_returns_json()
    {
        var tools = Tools(allowed: "alice");

        var names = tools.AsAIFunctions(new OfficeAgentToolsOptions { AllowConnectionAddressing = true })
            .Select(function => function.Name)
            .ToArray();
        Assert.Contains("describe_capabilities", names);

        using var payload = JsonDocument.Parse(await tools.DescribeCapabilities());
        Assert.True(payload.RootElement.TryGetProperty("contracts", out _));
        Assert.True(payload.RootElement.TryGetProperty("formats", out _));
        Assert.True(payload.RootElement.TryGetProperty("limits", out _));
    }

    /// <summary>
    /// Every node kind a module's inspection emits is declared, so discovery offers it.
    /// </summary>
    /// <remarks>
    /// The verb reconciliation above takes node kinds from the declaration it checks, so a kind
    /// missing from the declaration blinds both sides at once: Excel declared none, and
    /// appendTableRows, which targets a spreadsheetTable node, was never advertised. This side
    /// takes the kinds from what inspection actually returns for a representative document.
    /// </remarks>
    [Theory]
    [InlineData("word")]
    [InlineData("powerpoint")]
    [InlineData("excel")]
    public void Every_node_kind_inspection_emits_is_declared(string format)
    {
        var (module, bytes) = format switch
        {
            "word" => ((IFormatModule)new WordModule(), DocxFactory.Contract()),
            "powerpoint" => (new PowerPointModule(), PptxFactory.DeckWithTable()),
            _ => (new ExcelModule(), XlsxFactory.WorkbookWithTable())
        };
        var inspected = new OfficeAgentClient(module).Inspect(Handle(bytes), new InspectOptions { Fidelity = Fidelity.Content });
        var emitted = inspected.Nodes.Select(node => node.Kind).Distinct().ToList();
        var declared = ((ICapabilityDeclaringModule)module).NodeKinds;

        Assert.NotEmpty(emitted);
        Assert.Empty(emitted.Except(declared));
    }

    /// <summary>
    /// copyStyles has two anchors, and discovery once probed only Target, so it was omitted
    /// for both formats that implement it. Advertised, and a real one previews cleanly.
    /// </summary>
    [Fact]
    public async Task Word_and_PowerPoint_advertise_copyStyles_and_a_real_one_previews()
    {
        Assert.Contains("copyStyles", Describe(new WordModule()).Operations);
        Assert.Contains("copyStyles", Describe(new PowerPointModule()).Operations);

        var client = new OfficeAgentClient(new WordModule());
        var hits = await client.FindAsync(Handle(DocxFactory.Contract()), new FindQuery("Acme Corp"));
        var paragraphs = client.Inspect(Handle(DocxFactory.Contract()), new InspectOptions()).Paragraphs
            .Where(p => p.Text.Length > 0).Take(2).ToArray();
        var report = client.Preview(Handle(DocxFactory.Contract()), new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new CopyStylesOp
                {
                    Source = new TextSpanAnchor { ParaId = paragraphs[0].ParaId, Expect = "" },
                    Target = new TextSpanAnchor { ParaId = paragraphs[1].ParaId, Expect = "" },
                    Scope = "all"
                }
            }
        });
        Assert.True(report.IsValid, string.Join("; ", report.Errors.Select(e => e.Code + " " + e.Message)));
    }

    [Fact]
    public void Excel_advertises_appending_rows_to_a_table()
    {
        Assert.Contains("appendTableRows", Describe(new ExcelModule()).Operations);
    }

    // ── Helpers ──────────────────────────────────────────────────────────

    private static StreamHandle Handle(byte[] bytes) => new(new MemoryStream(bytes, writable: false));

    internal static FormatCapabilities Describe(IFormatModule module) =>
        CapabilityDiscovery.Describe(new[] { module }).Formats.Single();

    /// <summary>
    /// Asks the module's handlers directly, using the same anchor shapes discovery offers.
    /// This is the independent side of the reconciliation.
    /// </summary>
    private static bool AnyHandlerAccepts(IFormatModule module, string verb)
    {
        var type = PlanOperationJsonConverter.ByVerb[verb];
        var kinds = (module as ICapabilityDeclaringModule)?.NodeKinds ?? Array.Empty<string>();

        var anchors = new List<Anchor?>
        {
            null,
            new TextSpanAnchor(),
            new StructuralAnchor(),
            new StyleAnchor(),
            new CellAnchor(),
            new ShapeAnchor()
        };
        foreach (var kind in kinds) anchors.Add(new NodeAnchor { Kind = kind });

        foreach (var anchor in anchors)
        {
            if (Activator.CreateInstance(type) is not PlanOperation candidate) continue;
            if (anchor is not null)
            {
                var property = type.GetProperty(nameof(PlanOperation.Target));
                if (property is null || !property.PropertyType.IsInstanceOfType(anchor)) continue;
                property.SetValue(candidate, anchor);
                foreach (var slot in type.GetProperties())
                    if (slot != property && slot.CanWrite && slot.PropertyType.IsInstanceOfType(anchor))
                        slot.SetValue(candidate, anchor);
            }

            if (module.Handlers.Any(handler => handler.CanHandle(candidate))) return true;
        }

        return false;
    }

    private static OfficeAgentTools Tools(string allowed) => Tools(new SingleConnectionPolicy(allowed));

    private static OfficeAgentTools Tools(IConnectionAccessPolicy policy) =>
        new(new OfficeAgentClient(
                new DocumentProviderRegistry(new[]
                {
                    new MemoryDocumentProvider("alice"),
                    new MemoryDocumentProvider("bob")
                }),
                new WordModule()),
            policy,
            new FixedPrincipal(new ClaimsPrincipal(new ClaimsIdentity("test"))));

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

    private sealed class ReadOnlyOnAlice : IConnectionAccessPolicy
    {
        public ValueTask<bool> IsAllowedAsync(
            ClaimsPrincipal principal,
            string connectionId,
            ConnectionCapability capability,
            CancellationToken cancellationToken = default) =>
            new(connectionId == "alice" && capability == ConnectionCapability.Read);
    }

    private sealed class FixedPrincipal : ITrustedPrincipalAccessor
    {
        public FixedPrincipal(ClaimsPrincipal principal) => Principal = principal;

        public ClaimsPrincipal Principal { get; }
    }
}
