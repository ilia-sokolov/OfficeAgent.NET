using System.Text.Json;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// The path-addressed composite tools: open_document and edit_document collapse the
/// register → inspect → find → apply loop into one call each, and edit_document binds
/// "find" targets to live anchors so the agent never has to look a paragraph id up first.
/// </summary>
public class CompositeToolsTests
{
    [Fact]
    public async Task Open_document_registers_and_inspects_in_one_call()
    {
        using var workspace = new CompositeWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        workspace.Stage("contract.docx", DocxFactory.Contract());

        using var opened = JsonDocument.Parse(await tools.OpenDocument("workspace", "contract.docx"));
        var root = opened.RootElement;

        // The registration half.
        var documentId = root.GetProperty("documentId").GetString()!;
        Assert.Equal("workspace", root.GetProperty("connectionId").GetString());
        Assert.Equal("contract.docx", root.GetProperty("name").GetString());
        Assert.False(string.IsNullOrEmpty(root.GetProperty("version").GetString()));

        // …and the inspection half, in the same shape inspect_document returns.
        Assert.Equal("Word", root.GetProperty("format").GetString());
        Assert.False(string.IsNullOrEmpty(root.GetProperty("snapshot").GetString()));
        Assert.True(root.GetProperty("paragraphsTotal").GetInt32() > 0);
        Assert.NotEmpty(root.GetProperty("paragraphs").EnumerateArray());

        // The id it hands back drives the ordinary loop.
        using var inspected = JsonDocument.Parse(await tools.InspectDocument("workspace", documentId));
        Assert.Equal(
            root.GetProperty("snapshot").GetString(),
            inspected.RootElement.GetProperty("snapshot").GetString());
    }

    [Fact]
    public async Task Open_document_reports_a_bad_source_without_leaking_the_root()
    {
        using var workspace = new CompositeWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);

        using var escape = JsonDocument.Parse(await tools.OpenDocument("workspace", "../outside.docx"));

        Assert.False(escape.RootElement.GetProperty("isValid").GetBoolean());
        Assert.Equal("access-denied",
            escape.RootElement.GetProperty("errors")[0].GetProperty("Code").GetString());
    }

    [Fact]
    public async Task Edit_document_resolves_a_find_target_and_applies_it()
    {
        using var workspace = new CompositeWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        workspace.Stage("contract.docx", DocxFactory.Contract());

        // "Service Agreement" is unique in the fixture, so no match index is needed.
        var plan = """
            [ { "op": "changeText",
                "target": { "find": "Service Agreement" },
                "with": "Master Agreement",
                "mode": "Direct" } ]
            """;

        using var edited = JsonDocument.Parse(
            await tools.EditDocument("workspace", "contract.docx", plan));
        var root = edited.RootElement;

        Assert.True(root.GetProperty("committed").GetBoolean(),
            root.GetProperty("errors").ToString());
        Assert.False(string.IsNullOrEmpty(root.GetProperty("sourceDocumentId").GetString()));

        // The edit really landed in the saved revision, addressed by the returned id.
        var outputId = root.GetProperty("outputDocumentId").GetString()!;
        Assert.Contains("Master Agreement", await TextOf(tools, outputId));
    }

    [Fact]
    public async Task Ambiguous_find_text_is_refused_and_lists_the_candidates()
    {
        using var workspace = new CompositeWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        workspace.Stage("contract.docx", DocxFactory.Contract());

        // The fixture says "Acme Corp shall provide services to Acme Corp." - editing the
        // first of the two silently would be undetectable from the result.
        var plan = """
            [ { "op": "changeText", "target": { "find": "Acme Corp" }, "with": "Globex Inc." } ]
            """;

        using var refused = JsonDocument.Parse(
            await tools.EditDocument("workspace", "contract.docx", plan));
        var error = refused.RootElement.GetProperty("errors")[0];

        Assert.False(refused.RootElement.GetProperty("committed").GetBoolean());
        Assert.Equal("ambiguous-anchor", error.GetProperty("Code").GetString());
        Assert.Contains("match 0:", error.GetProperty("Message").GetString());
        Assert.Contains("\"match\"", error.GetProperty("Message").GetString());

        // A failed edit still leaves a usable handle rather than making the agent re-register.
        Assert.False(string.IsNullOrEmpty(
            refused.RootElement.GetProperty("sourceDocumentId").GetString()));
    }

    [Fact]
    public async Task A_match_index_picks_the_occurrence_the_agent_meant()
    {
        using var workspace = new CompositeWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        workspace.Stage("contract.docx", DocxFactory.Contract());

        // Both matches sit in the same paragraph, so this also pins down that a
        // document-wide match index maps to the right in-paragraph occurrence.
        var plan = """
            [ { "op": "changeText",
                "target": { "find": "Acme Corp", "match": 1 },
                "with": "Globex Inc.",
                "mode": "Direct" } ]
            """;

        using var edited = JsonDocument.Parse(
            await tools.EditDocument("workspace", "contract.docx", plan));

        Assert.True(edited.RootElement.GetProperty("committed").GetBoolean(),
            edited.RootElement.GetProperty("errors").ToString());
        Assert.Single(edited.RootElement.GetProperty("changes").EnumerateArray());

        // The second occurrence changed and the first did not.
        var text = await TextOf(tools, edited.RootElement.GetProperty("outputDocumentId").GetString()!);
        Assert.Contains("Acme Corp shall provide services to Globex Inc..", text);
    }

    [Fact]
    public async Task An_out_of_range_match_and_missing_text_are_both_actionable()
    {
        using var workspace = new CompositeWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        workspace.Stage("contract.docx", DocxFactory.Contract());

        using var missing = JsonDocument.Parse(await tools.EditDocument("workspace", "contract.docx",
            """[ { "op": "changeText", "target": { "find": "Nowhere In This File" }, "with": "x" } ]"""));
        using var outOfRange = JsonDocument.Parse(await tools.EditDocument("workspace", "contract.docx",
            """[ { "op": "changeText", "target": { "find": "Acme Corp", "match": 99 }, "with": "x" } ]"""));

        Assert.Equal("anchor-not-found",
            missing.RootElement.GetProperty("errors")[0].GetProperty("Code").GetString());
        Assert.Equal("anchor-not-found",
            outOfRange.RootElement.GetProperty("errors")[0].GetProperty("Code").GetString());
        Assert.Contains("out of range",
            outOfRange.RootElement.GetProperty("errors")[0].GetProperty("Message").GetString());
    }

    [Fact]
    public async Task Every_unresolvable_target_is_reported_in_one_result()
    {
        using var workspace = new CompositeWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        workspace.Stage("contract.docx", DocxFactory.Contract());

        // Two bad targets: the agent should learn about both from a single call rather
        // than fixing one, retrying, and discovering the next.
        var plan = """
            [ { "op": "changeText", "target": { "find": "Absent One" }, "with": "x" },
              { "op": "changeText", "target": { "find": "Absent Two" }, "with": "y" } ]
            """;

        using var refused = JsonDocument.Parse(
            await tools.EditDocument("workspace", "contract.docx", plan));

        Assert.Equal(2, refused.RootElement.GetProperty("errors").GetArrayLength());
    }

    [Fact]
    public async Task Paragraph_id_targets_still_work_and_mix_with_find_targets()
    {
        using var workspace = new CompositeWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        workspace.Stage("contract.docx", DocxFactory.Contract());

        var documentId = await IdOf(tools, "contract.docx");
        using var hits = JsonDocument.Parse(
            await tools.FindInDocument("workspace", documentId, "Acme Corp"));
        var paraId = hits.RootElement[0].GetProperty("paraId").GetString()!;

        var plan = $$"""
            [ { "op": "changeText",
                "target": { "paraId": "{{paraId}}", "expect": "Acme Corp", "occurrence": 0 },
                "with": "Globex Inc.",
                "mode": "Direct" },
              { "op": "comment",
                "target": { "find": "Effective date" },
                "text": "Confirm the term." } ]
            """;

        using var edited = JsonDocument.Parse(
            await tools.EditDocument("workspace", "contract.docx", plan));

        Assert.True(edited.RootElement.GetProperty("committed").GetBoolean(),
            edited.RootElement.GetProperty("errors").ToString());
        Assert.Equal(2, edited.RootElement.GetProperty("changes").GetArrayLength());
    }

    [Fact]
    public async Task Nothing_is_written_when_one_operation_fails_to_resolve()
    {
        using var workspace = new CompositeWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        workspace.Stage("contract.docx", DocxFactory.Contract());
        var before = File.ReadAllBytes(Path.Combine(workspace.Root, "contract.docx"));

        var plan = """
            [ { "op": "changeText", "target": { "find": "Acme Corp" }, "with": "Globex Inc." },
              { "op": "changeText", "target": { "find": "Absent" }, "with": "x" } ]
            """;

        using var refused = JsonDocument.Parse(
            await tools.EditDocument("workspace", "contract.docx", plan));

        Assert.False(refused.RootElement.GetProperty("committed").GetBoolean());
        Assert.Equal(before, File.ReadAllBytes(Path.Combine(workspace.Root, "contract.docx")));
        Assert.Empty(Directory.GetFiles(workspace.Root, "contract.v*.docx"));
    }

    [Fact]
    public async Task A_bare_operations_array_is_accepted_everywhere_a_plan_is()
    {
        using var workspace = new CompositeWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);
        workspace.Stage("contract.docx", DocxFactory.Contract());
        var documentId = await IdOf(tools, "contract.docx");

        const string bareArray = """
            [ { "op": "comment", "target": { "find": "Service Agreement" }, "text": "Check." } ]
            """;
        const string planObject = """
            { "operations": [ { "op": "comment", "target": { "find": "Service Agreement" }, "text": "Check." } ] }
            """;

        using var fromArray = JsonDocument.Parse(
            await tools.EditDocument("workspace", "contract.docx", bareArray));
        using var fromObject = JsonDocument.Parse(
            await tools.EditDocument("workspace", "contract.docx", planObject));

        Assert.True(fromArray.RootElement.GetProperty("committed").GetBoolean());
        Assert.True(fromObject.RootElement.GetProperty("committed").GetBoolean());

        // preview_plan takes the array form too, so one habit works across every tool.
        using var previewed = JsonDocument.Parse(await tools.PreviewPlan("workspace", documentId,
            """[ { "op": "comment", "target": { "paraId": "auto-0000", "expect": "" }, "text": "x" } ]"""));
        Assert.False(string.IsNullOrEmpty(previewed.RootElement.GetProperty("isValid").ToString()));
    }

    [Fact]
    public void Composite_tools_follow_the_registration_switch()
    {
        using var workspace = new CompositeWorkspace();
        var tools = new OfficeAgentTools(workspace.Client);

        var defaults = tools.AsAIFunctions().Select(f => f.Name).ToArray();
        Assert.DoesNotContain("open_document", defaults);
        Assert.DoesNotContain("edit_document", defaults);

        // They address documents by the same connection-relative source register_document
        // takes, so they belong to that opt-in rather than to a new one.
        var registering = tools.AsAIFunctions(new OfficeAgentToolsOptions { AllowRegistration = true })
            .Select(f => f.Name).ToArray();
        Assert.Contains("open_document", registering);
        Assert.Contains("edit_document", registering);
    }

    /// <summary>Registers an already-staged fixture and returns its id, for lookups.</summary>
    private static async Task<string> IdOf(OfficeAgentTools tools, string name)
    {
        using var registered = JsonDocument.Parse(await tools.RegisterDocument("workspace", name));
        return registered.RootElement.GetProperty("documentId").GetString()!;
    }

    /// <summary>The document's body text, joined, for asserting on what an edit produced.</summary>
    private static async Task<string> TextOf(OfficeAgentTools tools, string documentId)
    {
        using var inspected = JsonDocument.Parse(await tools.InspectDocument("workspace", documentId));
        return string.Join(" ", inspected.RootElement.GetProperty("paragraphs").EnumerateArray()
            .Select(p => p.GetProperty("Text").GetString()));
    }

    private sealed class CompositeWorkspace : IDisposable
    {
        private readonly ServiceProvider _services;

        public CompositeWorkspace()
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-composite-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);
            var services = new ServiceCollection();
            services.AddWordFormat();
            services.AddFileSystemDocumentProvider("workspace", Root);
            services.AddOfficeAgent();
            _services = services.BuildServiceProvider();
            Client = _services.GetRequiredService<OfficeAgentClient>();
        }

        public string Root { get; }
        public OfficeAgentClient Client { get; }

        public void Stage(string name, byte[] bytes) =>
            File.WriteAllBytes(Path.Combine(Root, name), bytes);

        public void Dispose()
        {
            _services.Dispose();
            if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true);
        }
    }
}
