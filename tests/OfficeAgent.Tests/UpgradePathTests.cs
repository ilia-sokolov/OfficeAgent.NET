using System.Text.Json;
using System.Text.Json.Serialization;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// The upgrade paths docs/upgrading-to-1.0.md promises, each run against 1.0.
/// </summary>
/// <remarks>
/// Stored plans are covered by the v0.8 corpus tests, which run the committed 0.8 plan through
/// the direct, provider and tool paths. These cover what those do not: a stored receipt read
/// back into the typed record, and a host that deserialises a tool response into the typed
/// record it describes, which is the one caller the camelCase change can break silently.
/// </remarks>
public sealed class UpgradePathTests
{
    private static readonly string CorpusRoot = Path.Combine(
        RepositoryRoot(), "tests", "OfficeAgent.Tests", "Corpus", "v0.8.0");

    /// <summary>What the guide tells a host to deserialise with.</summary>
    private static readonly JsonSerializerOptions HostOptions = new(JsonSerializerDefaults.Web)
    {
        Converters = { new JsonStringEnumConverter() }
    };

    [Fact]
    public void A_stored_08_receipt_reads_into_the_10_receipt_type()
    {
        var receipt = JsonSerializer.Deserialize<ApplyReceipt>(
            File.ReadAllText(Path.Combine(CorpusRoot, "apply-receipt.json")), HostOptions)!;

        Assert.Equal(ApplyOutcome.Committed, receipt.Outcome);
        Assert.Equal(new string('1', 64), receipt.PlanSha256);
        Assert.Equal(new string('0', 64), receipt.InputSha256);
        Assert.Equal(new string('2', 64), receipt.OutputSha256);
        Assert.Equal("Corpus Review", receipt.Revision.Author);
        Assert.Equal("fixture-user", receipt.Actor!.Subject);
        // Every receipt emitted since 0.8 is schema "1", so a stored receipt that omits the
        // field reads as the schema it was written in.
        Assert.Equal("1", receipt.ReceiptVersion);
        Assert.Equal(ApplyReceipt.CurrentReceiptVersion, receipt.ReceiptVersion);
    }

    [Fact]
    public async Task A_host_deserialising_a_response_with_web_defaults_reads_10_unchanged()
    {
        using var services = Services();
        var tools = new OfficeAgentTools(services.GetRequiredService<OfficeAgentClient>());
        var json = await tools.DescribeCapabilities();

        var capabilities = JsonSerializer.Deserialize<EngineCapabilities>(json, HostOptions)!;

        Assert.Equal(DocumentPlan.CurrentContractVersion, capabilities.Contracts.EditPlan);
        // Every member has a non-empty default, so only a list the response fills proves the
        // names matched.
        var word = Assert.Single(capabilities.Formats);
        Assert.Equal(OfficeAgent.Abstractions.DocumentFormat.Word, word.Format);
        Assert.Contains("changeText", word.Operations);
    }

    [Fact]
    public async Task A_host_deserialising_case_sensitively_with_pascal_names_does_not()
    {
        // The break the guide warns about. System.Text.Json's plain defaults match names
        // exactly, so a host that read 0.9's PascalCase capabilities this way now gets each
        // member's default, not an exception. The defaults look plausible (the edit-plan
        // version is even correct), which is why the guide calls this out. Pinned so the
        // warning stays true.
        using var services = Services();
        var tools = new OfficeAgentTools(services.GetRequiredService<OfficeAgentClient>());
        var json = await tools.DescribeCapabilities();

        var capabilities = JsonSerializer.Deserialize<EngineCapabilities>(
            json, new JsonSerializerOptions { Converters = { new JsonStringEnumConverter() } })!;

        Assert.Empty(capabilities.Formats);
        Assert.Equal(DocumentPlan.CurrentContractVersion, capabilities.Contracts.EditPlan);
    }

    private static ServiceProvider Services()
    {
        var services = new ServiceCollection();
        services.AddWordFormat();
        services.AddOfficeAgent();
        return services.BuildServiceProvider();
    }

    private static string RepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "OfficeAgent.NET.sln")))
            directory = directory.Parent;
        return directory?.FullName ?? throw new DirectoryNotFoundException("repository root not found");
    }
}
