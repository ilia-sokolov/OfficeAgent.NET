using System.Text.Json;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

public sealed class PlanCompatibilityTests
{
    [Fact]
    public void Omitted_and_explicit_current_versions_have_the_same_typed_behavior()
    {
        var omitted = JsonSerializer.Deserialize<DocumentPlan>("""{ "operations": [] }""")!;
        var explicitCurrent = JsonSerializer.Deserialize<DocumentPlan>("""{ "contractVersion": "0.2", "operations": [] }""")!;
        var client = new OfficeAgentClient(new WordModule());
        var input = DocxFactory.Contract();

        var omittedReport = client.Preview(Handle(input), omitted);
        var explicitReport = client.Preview(Handle(input), explicitCurrent);
        using var omittedApply = client.Commit(Handle(input), omitted);
        using var explicitApply = client.Commit(Handle(input), explicitCurrent);

        Assert.Equal(DocumentPlan.CurrentContractVersion, omitted.ContractVersion);
        Assert.True(omittedReport.IsValid);
        Assert.Equal(omittedReport.IsValid, explicitReport.IsValid);
        Assert.Equal(omittedReport.Changes.Count, explicitReport.Changes.Count);
        Assert.True(omittedApply.Committed);
        Assert.True(explicitApply.Committed);
        Assert.Equal(omittedApply.Receipt!.PlanSha256, explicitApply.Receipt!.PlanSha256);
    }

    [Theory]
    [InlineData("[]")]
    [InlineData("{ \"operations\": [] }")]
    [InlineData("{ \"contractVersion\": \"0.2\", \"operations\": [] }")]
    public async Task Agent_tool_accepts_bare_legacy_and_explicit_current_plans(string planJson)
    {
        var store = new MemoryDocumentProvider("docs");
        var reference = store.Add("contract.docx", DocxFactory.Contract());
        var tools = new OfficeAgentTools(
            new OfficeAgentClient(new DocumentProviderRegistry(new[] { store }), new WordModule()));

        using var result = JsonDocument.Parse(
            await tools.PreviewPlan("docs", reference.ItemId, planJson));

        Assert.True(result.RootElement.GetProperty("isValid").GetBoolean(), result.RootElement.ToString());
        Assert.False(result.RootElement.GetProperty("committed").GetBoolean());
        Assert.Equal(1, store.Count);
    }

    [Theory]
    [InlineData(null)]
    [InlineData("")]
    [InlineData(" ")]
    [InlineData("0.2.0")]
    [InlineData("1.0")]
    public async Task Typed_entry_points_reject_incompatible_versions_before_service_access(string? version)
    {
        var client = new OfficeAgentClient(new ThrowingDocumentService());
        var plan = new DocumentPlan { ContractVersion = version!, Operations = Array.Empty<PlanOperation>() };

        var preview = client.Preview(null!, plan);
        using var apply = client.Apply(null!, plan, ApplyOptions.Commit);
        var asyncPreview = await client.PreviewAsync((DocumentHandle)null!, plan);
        using var asyncApply = await client.ApplyAsync((DocumentHandle)null!, plan, ApplyOptions.Commit);

        AssertContractMismatch(preview);
        AssertContractMismatch(apply.Report);
        AssertContractMismatch(asyncPreview);
        AssertContractMismatch(asyncApply.Report);
        Assert.False(apply.Committed);
        Assert.Null(apply.Output);
        Assert.Null(apply.Receipt);
        Assert.False(asyncApply.Committed);
        Assert.Null(asyncApply.Output);
        Assert.Null(asyncApply.Receipt);
    }

    [Fact]
    public async Task Provider_commit_and_create_reject_unknown_version_without_writing()
    {
        var input = DocxFactory.Contract();
        var store = new MemoryDocumentProvider("docs");
        var reference = store.Add("contract.docx", input);
        var client = new OfficeAgentClient(new DocumentProviderRegistry(new[] { store }), new WordModule());
        var plan = new DocumentPlan { ContractVersion = "1.0" };

        var committed = await client.CommitAsync(reference, plan);
        var created = await client.CreateAsync("docs", "new.docx", plan);

        Assert.False(committed.Committed);
        Assert.False(created.Committed);
        AssertContractMismatch(committed.Report);
        AssertContractMismatch(created.Report);
        Assert.Null(committed.Document);
        Assert.Null(committed.Receipt);
        Assert.Null(created.Document);
        Assert.Null(created.Receipt);
        Assert.Equal(1, store.Count);
        Assert.Equal(input, store.Read(reference.ItemId));
        Assert.Equal(reference.Version, store.Describe(reference.ItemId).Version);
    }

    [Theory]
    [InlineData("null")]
    [InlineData("\"\"")]
    [InlineData("\"1.0\"")]
    public async Task Agent_tool_reports_contract_mismatch_and_does_not_save(string wireVersion)
    {
        var input = DocxFactory.Contract();
        var store = new MemoryDocumentProvider("docs");
        var reference = store.Add("contract.docx", input);
        var tools = new OfficeAgentTools(
            new OfficeAgentClient(new DocumentProviderRegistry(new[] { store }), new WordModule()));

        using var result = JsonDocument.Parse(await tools.ApplyPlan(
            "docs", reference.ItemId,
            $"{{ \"contractVersion\": {wireVersion}, \"operations\": [] }}"));

        Assert.False(result.RootElement.GetProperty("isValid").GetBoolean());
        Assert.False(result.RootElement.GetProperty("committed").GetBoolean());
        Assert.Equal(
            ValidationErrorCodes.ContractMismatch,
            result.RootElement.GetProperty("errors")[0].GetProperty("code").GetString());
        Assert.Equal(1, store.Count);
        Assert.Equal(input, store.Read(reference.ItemId));
        Assert.Equal(reference.Version, store.Describe(reference.ItemId).Version);
    }

    [Theory]
    [InlineData("{ \"contractVersion\": 2, \"operations\": [] }")]
    [InlineData("{ \"operations\": [], \"futureProperty\": true }")]
    [InlineData("{ \"operations\": [ { \"op\": \"deleteEverything\" } ] }")]
    [InlineData("{ \"operations\": [ { \"op\": \"changeText\", \"target\": { \"paraId\": \"p\", \"expect\": \"x\" }, \"with\": \"y\", \"mode\": \"FutureMode\" } ] }")]
    [InlineData("{ \"operations\": [ { \"op\": \"changeText\", \"target\": { \"paraId\": \"p\", \"expect\": \"x\" }, \"with\": \"y\", \"mode\": 1 } ] }")]
    [InlineData("{ \"operations\": [ { \"op\": \"changeText\", \"target\": { \"paraId\": \"p\", \"expect\": \"x\", \"futureProperty\": true }, \"with\": \"y\" } ] }")]
    [InlineData("{ \"operations\": [ { \"op\": \"changeText\", \"target\": { \"paraId\": \"p\", \"expect\": \"x\" }, \"with\": \"y\", \"futureProperty\": true } ] }")]
    public async Task Agent_tool_rejects_malformed_or_unknown_wire_members(string planJson)
    {
        var store = new MemoryDocumentProvider("docs");
        var reference = store.Add("contract.docx", DocxFactory.Contract());
        var tools = new OfficeAgentTools(
            new OfficeAgentClient(new DocumentProviderRegistry(new[] { store }), new WordModule()));

        using var result = JsonDocument.Parse(
            await tools.PreviewPlan("docs", reference.ItemId, planJson));

        Assert.False(result.RootElement.GetProperty("isValid").GetBoolean());
        Assert.Equal(
            "invalid-json",
            result.RootElement.GetProperty("errors")[0].GetProperty("code").GetString());
    }

    private static StreamHandle Handle(byte[] bytes) =>
        new(new MemoryStream(bytes, writable: false), "contract.docx");

    private static void AssertContractMismatch(ChangeReport report)
    {
        Assert.False(report.IsValid);
        var error = Assert.Single(report.Errors);
        Assert.Equal(ValidationErrorCodes.ContractMismatch, error.Code);
        Assert.Contains(DocumentPlan.CurrentContractVersion, error.Message);
    }

    private sealed class ThrowingDocumentService : IDocumentService
    {
        private static Exception Accessed() => new InvalidOperationException("Document service must not be accessed.");

        public InspectResult Inspect(DocumentHandle handle, InspectOptions options) => throw Accessed();
        public IReadOnlyList<FindHit> Find(DocumentHandle handle, FindQuery query) => throw Accessed();
        public ChangeReport Validate(DocumentHandle handle, DocumentPlan plan) => throw Accessed();
        public ApplyResult Apply(DocumentHandle handle, DocumentPlan plan, ApplyOptions options) => throw Accessed();
        public Task<InspectResult> InspectAsync(DocumentHandle handle, InspectOptions options, CancellationToken cancellationToken = default) => throw Accessed();
        public Task<IReadOnlyList<FindHit>> FindAsync(DocumentHandle handle, FindQuery query, CancellationToken cancellationToken = default) => throw Accessed();
        public Task<ChangeReport> ValidateAsync(DocumentHandle handle, DocumentPlan plan, CancellationToken cancellationToken = default) => throw Accessed();
        public Task<ApplyResult> ApplyAsync(DocumentHandle handle, DocumentPlan plan, ApplyOptions options, CancellationToken cancellationToken = default) => throw Accessed();
    }
}
