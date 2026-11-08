using System.ClientModel;
using System.ClientModel.Primitives;
using System.Security.Cryptography;
using System.Text.Json;
using Azure.Identity;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.Extensions.AI;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.SemanticKernel;
using Microsoft.SemanticKernel.ChatCompletion;
using Microsoft.SemanticKernel.Connectors.OpenAI;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;
using OpenAI;

const string Instruction =
    "Change the initial term in clause 5 from twelve months to twenty-four months as a " +
    "tracked change. Clause 4 also says twelve months; leave it as it is. Make no other edits.";

bool listFunctions = args.Contains("--list-functions");
string outputPath = Path.GetFullPath(
    args.FirstOrDefault(a => !a.StartsWith("--", StringComparison.Ordinal))
        ?? "reviewed-services-agreement.docx");

string storageRoot = Path.Combine(
    Path.GetTempPath(),
    $"officeagent-semantickernel-{Guid.NewGuid():N}");
Directory.CreateDirectory(storageRoot);

try
{
    string sourcePath = Path.Combine(storageRoot, "services-agreement.docx");
    File.Copy(
        Path.Combine(AppContext.BaseDirectory, "services-agreement.docx"),
        sourcePath);
    byte[] sourceHashBefore = SHA256.HashData(await File.ReadAllBytesAsync(sourcePath));

    using ServiceProvider services = new ServiceCollection()
        .AddWordFormat()
        .AddFileSystemDocumentProvider("contracts", storageRoot)
        .AddOfficeAgent()
        .BuildServiceProvider();

    OfficeAgentClient office = services.GetRequiredService<OfficeAgentClient>();
    DocumentReference source = await office.RegisterAsync("contracts", sourcePath);

    // OfficeAgent's tools are Microsoft.Extensions.AI functions; Semantic Kernel takes each
    // one as a kernel function, and the set becomes one plugin.
    AIFunction[] tools = new OfficeAgentTools(office).AsAIFunctions();
    IKernelBuilder builder = Kernel.CreateBuilder();
    builder.Plugins.AddFromFunctions(
        "officeagent",
        tools.Select(tool => tool.AsKernelFunction()));

    if (listFunctions)
    {
        Kernel plugins = builder.Build();
        foreach (KernelFunctionMetadata function in plugins.Plugins.GetFunctionsMetadata())
            Console.WriteLine($"{function.PluginName}.{function.Name}");

        // Call one function through the kernel, as the model would, to check the wiring.
        FunctionResult inspection = await plugins.InvokeAsync(
            "officeagent",
            "inspect_document",
            new KernelArguments
            {
                ["connectionId"] = "contracts",
                ["documentId"] = source.ItemId,
                ["fidelity"] = "outline",
                ["paragraphOffset"] = 0,
                ["paragraphLimit"] = 20,
                ["sheetId"] = 0,
                ["range"] = "",
                ["maximumCells"] = 0
            });
        string text = inspection.GetValue<object>()?.ToString() ?? "";
        Console.WriteLine($"inspect_document: {text[..Math.Min(text.Length, 200)]}...");
        return 0;
    }

    (OpenAIClient client, string model) = CreateOpenAIClient();
    builder.AddOpenAIChatCompletion(model, client);

    Kernel kernel = builder.Build();
    var trace = new ToolTrace();
    kernel.AutoFunctionInvocationFilters.Add(trace);

    var history = new ChatHistory($"""
        You are editing connectionId=contracts,
        documentId={source.ItemId}.

        {OfficeAgentTools.SystemPromptGuidance}

        When applying, use saveMode NewVersion.
        """);
    history.AddUserMessage(Instruction);

    var settings = new OpenAIPromptExecutionSettings
    {
        FunctionChoiceBehavior = FunctionChoiceBehavior.Auto(
            options: new FunctionChoiceBehaviorOptions { AllowParallelCalls = false }),
        MaxTokens = 1200
    };

    IChatCompletionService chat = kernel.GetRequiredService<IChatCompletionService>();
    ChatMessageContent reply = await chat.GetChatMessageContentAsync(history, settings, kernel);
    Console.WriteLine(reply.Content);

    if (trace.OutputDocumentId is null)
    {
        throw new InvalidOperationException(
            trace.LastApplyResult is null
                ? "The model did not call apply_plan."
                : $"apply_plan did not commit: {trace.LastApplyResult}");
    }

    using (var saved = await office.OpenReadAsync(
        DocumentReference.ForFileSystem("contracts", trace.OutputDocumentId)))
    await using (var output = File.Create(outputPath))
    {
        await saved.Stream.CopyToAsync(output);
    }

    byte[] sourceHashAfter = SHA256.HashData(await File.ReadAllBytesAsync(sourcePath));
    if (!sourceHashBefore.SequenceEqual(sourceHashAfter))
        throw new InvalidOperationException("Source document changed in place.");

    Console.WriteLine($"Saved {outputPath}");
    ReportRevisions(outputPath);
    return 0;
}
finally
{
    if (Directory.Exists(storageRoot))
        Directory.Delete(storageRoot, recursive: true);
}

// Azure OpenAI when AZURE_OPENAI_ENDPOINT is set, otherwise OpenAI. Nothing is read from code.
static (OpenAIClient Client, string Model) CreateOpenAIClient()
{
    string? endpoint = Environment.GetEnvironmentVariable("AZURE_OPENAI_ENDPOINT");
    if (!string.IsNullOrEmpty(endpoint))
    {
        string deployment = Environment.GetEnvironmentVariable("AZURE_OPENAI_DEPLOYMENT")
            ?? throw new InvalidOperationException("Set AZURE_OPENAI_DEPLOYMENT.");
        // Azure OpenAI's v1 API takes the OpenAI client at <endpoint>/openai/v1/.
        var options = new OpenAIClientOptions
        {
            Endpoint = new Uri(new Uri(endpoint), "openai/v1/")
        };
        string? key = Environment.GetEnvironmentVariable("AZURE_OPENAI_API_KEY");
        if (!string.IsNullOrEmpty(key))
            return (new OpenAIClient(new ApiKeyCredential(key), options), deployment);

        // Without a key, sign in with Microsoft Entra ID (az login). The OpenAI client marks its
        // token-policy constructor as evaluation-only (OPENAI001); it is the documented way to
        // use Entra ID with Azure OpenAI's v1 API.
#pragma warning disable OPENAI001
        var entra = new OpenAIClient(
            new BearerTokenPolicy(
                new DefaultAzureCredential(),
                "https://cognitiveservices.azure.com/.default"),
            options);
#pragma warning restore OPENAI001
        return (entra, deployment);
    }

    string apiKey = Environment.GetEnvironmentVariable("OPENAI_API_KEY")
        ?? throw new InvalidOperationException(
            "Set AZURE_OPENAI_ENDPOINT and AZURE_OPENAI_DEPLOYMENT, or OPENAI_API_KEY.");
    string model = Environment.GetEnvironmentVariable("OPENAI_MODEL") ?? "gpt-5.6-luna";
    return (new OpenAIClient(new ApiKeyCredential(apiKey)), model);
}

static void ReportRevisions(string path)
{
    using WordprocessingDocument document = WordprocessingDocument.Open(path, false);
    Body body = document.MainDocumentPart!.Document!.Body!;
    bool inserted = body.Descendants<InsertedRun>()
        .Any(run => run.InnerText.Contains("twenty-four months", StringComparison.Ordinal));
    bool deleted = body.Descendants<DeletedRun>()
        .Any(run => string.Concat(run.Descendants<DeletedText>().Select(t => t.Text))
            .Contains("twelve months", StringComparison.Ordinal));
    Console.WriteLine($"Tracked insertion of 'twenty-four months': {inserted}");
    Console.WriteLine($"Tracked deletion of 'twelve months':       {deleted}");
}

/// <summary>Prints each tool call and captures <c>outputDocumentId</c> from apply_plan.</summary>
sealed class ToolTrace : IAutoFunctionInvocationFilter
{
    public string? OutputDocumentId { get; private set; }

    public string? LastApplyResult { get; private set; }

    public async Task OnAutoFunctionInvocationAsync(
        AutoFunctionInvocationContext context,
        Func<AutoFunctionInvocationContext, Task> next)
    {
        Console.WriteLine($"tool: {context.Function.Name}");
        await next(context);

        if (context.Function.Name != "apply_plan"
            || Normalize(context.Result.GetValue<object>()) is not JsonElement result)
        {
            return;
        }

        LastApplyResult = result.GetRawText();
        if (result.TryGetProperty("committed", out JsonElement committed)
            && committed.ValueKind == JsonValueKind.True
            && result.TryGetProperty("outputDocumentId", out JsonElement outputId))
        {
            OutputDocumentId = outputId.GetString();
            if (result.TryGetProperty("changes", out JsonElement changes))
                Console.WriteLine($"changes: {changes.GetRawText()}");
        }
    }

    // An OfficeAgent result can arrive as a JSON object or as JSON text inside a string.
    private static JsonElement? Normalize(object? value) => value switch
    {
        JsonElement { ValueKind: JsonValueKind.String } element
            when element.GetString() is string json => Parse(json),
        JsonElement element => element,
        string json => Parse(json),
        _ => null
    };

    private static JsonElement? Parse(string json)
    {
        try
        {
            using JsonDocument document = JsonDocument.Parse(json);
            return document.RootElement.Clone();
        }
        catch (JsonException)
        {
            return null;
        }
    }
}
