using System.Security.Cryptography;
using System.Text.Json;
using Azure.AI.OpenAI;
using Azure.Identity;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.AI;
using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

string endpoint = Environment.GetEnvironmentVariable("AZURE_OPENAI_ENDPOINT")
    ?? throw new InvalidOperationException("Set AZURE_OPENAI_ENDPOINT.");
string deployment = Environment.GetEnvironmentVariable("AZURE_OPENAI_DEPLOYMENT")
    ?? throw new InvalidOperationException("Set AZURE_OPENAI_DEPLOYMENT.");
string outputPath = Path.GetFullPath(
    args.Length > 0 ? args[0] : "reviewed-contract.docx");
Directory.CreateDirectory(
    Path.GetDirectoryName(outputPath)
        ?? throw new InvalidOperationException("Output path has no directory."));

string storageRoot = Path.Combine(
    Path.GetTempPath(),
    $"officeagent-ichatclient-{Guid.NewGuid():N}");
Directory.CreateDirectory(storageRoot);

try
{
    string sourcePath = Path.Combine(storageRoot, "contract.docx");
    CreateFixture(sourcePath);
    byte[] sourceHashBefore = SHA256.HashData(
        await File.ReadAllBytesAsync(sourcePath));

    using ServiceProvider services = new ServiceCollection()
        .AddWordFormat()
        .AddFileSystemDocumentProvider("contracts", storageRoot)
        .AddOfficeAgent()
        .BuildServiceProvider();

    OfficeAgentClient office =
        services.GetRequiredService<OfficeAgentClient>();
    DocumentReference source = await office.RegisterAsync(
        "contracts",
        sourcePath);

    AIFunction[] tools = new OfficeAgentTools(office).AsAIFunctions();
    string? outputDocumentId = null;
    string? lastApplyResult = null;

    IChatClient chat = new AzureOpenAIClient(
            new Uri(endpoint),
            new DefaultAzureCredential())
        .GetChatClient(deployment)
        .AsIChatClient()
        .AsBuilder()
        .UseFunctionInvocation(configure: invoker =>
        {
            invoker.FunctionInvoker = async (context, cancellationToken) =>
            {
                Console.WriteLine($"tool: {context.Function.Name}");

                object? rawResult = await context.Function.InvokeAsync(
                    context.Arguments,
                    cancellationToken);
                object? result = NormalizeToolResult(rawResult);

                if (context.Function.Name == "apply_plan"
                    && result is JsonElement resultRoot)
                {
                    lastApplyResult = resultRoot.GetRawText();

                    if (resultRoot.TryGetProperty(
                            "committed",
                            out JsonElement committed)
                        && committed.GetBoolean()
                        && resultRoot.TryGetProperty(
                            "outputDocumentId",
                            out JsonElement outputId))
                    {
                        outputDocumentId = outputId.GetString();
                    }
                }

                return result;
            };
        })
        .Build();

    string instructions = $"""
        You are editing connectionId=contracts,
        documentId={source.ItemId}.

        {OfficeAgentTools.SystemPromptGuidance}

        Call inspect_document before planning. Use find_in_document when needed
        to locate the exact target. Preview before applying.
        When applying, use saveMode NewVersion.
        """;

    var messages = new List<ChatMessage>
    {
        new(ChatRole.System, instructions),
        new(
            ChatRole.User,
            "Change '60 days' to '30 days' as a tracked change. " +
            "Make no other edits.")
    };

    var options = new ChatOptions
    {
        Tools = tools.Cast<AITool>().ToList(),
        AllowMultipleToolCalls = false,
        MaxOutputTokens = 1200
    };

    ChatResponse response = await chat.GetResponseAsync(messages, options);
    Console.WriteLine(response.Text);

    if (outputDocumentId is null)
    {
        throw new InvalidOperationException(
            lastApplyResult is null
                ? "The model did not call apply_plan."
                : $"apply_plan did not commit: {lastApplyResult}");
    }

    using (var saved = await office.OpenReadAsync(
        DocumentReference.ForFileSystem("contracts", outputDocumentId)))
    await using (var output = File.Create(outputPath))
    {
        await saved.Stream.CopyToAsync(output);
    }

    byte[] sourceHashAfter = SHA256.HashData(
        await File.ReadAllBytesAsync(sourcePath));
    if (!sourceHashBefore.SequenceEqual(sourceHashAfter))
        throw new InvalidOperationException("Source document changed in place.");

    Console.WriteLine($"Saved {outputPath}");
    return 0;
}
finally
{
    if (Directory.Exists(storageRoot))
        Directory.Delete(storageRoot, recursive: true);
}

static object? NormalizeToolResult(object? result)
{
    if (result is JsonElement element)
    {
        if (element.ValueKind == JsonValueKind.String
            && element.GetString() is string json)
        {
            using JsonDocument parsed = JsonDocument.Parse(json);
            return parsed.RootElement.Clone();
        }

        return element;
    }

    if (result is string jsonText)
    {
        using JsonDocument parsed = JsonDocument.Parse(jsonText);
        return parsed.RootElement.Clone();
    }

    return result;
}

static void CreateFixture(string path)
{
    using WordprocessingDocument document =
        WordprocessingDocument.Create(
            path,
            WordprocessingDocumentType.Document);
    MainDocumentPart main = document.AddMainDocumentPart();
    main.Document = new Document(
        new Body(
            new Paragraph(
                new Run(
                    new Text(
                        "Invoices are payable within 60 days of receipt.")))));
    main.Document.Save();
}
