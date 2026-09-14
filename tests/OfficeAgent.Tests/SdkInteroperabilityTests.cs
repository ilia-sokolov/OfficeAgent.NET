using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Validation;
using DocumentFormat.OpenXml;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// Boundary tests for the documented Open XML SDK interoperability recipe: an
/// authorized snapshot is copied, the copy is edited with the SDK, and the edited
/// output is reinspected before any further OfficeAgent plan. These tests fix the
/// guarantees that actually hold at each stage and the ones that do not.
/// </summary>
public sealed class SdkInteroperabilityTests
{
    private const string SourceClause = "Invoices are due in 30 days.";
    private const string SdkAppendedClause = "Appendix A is incorporated by reference.";

    /// <summary>
    /// Stage 1 to 4 of the recipe. The registered source keeps its exact bytes and
    /// provider version, the SDK mutation lands only in the separate output, and the
    /// output is a valid package that OfficeAgent can inspect and then edit under a
    /// freshly authored plan.
    /// </summary>
    [Fact]
    public async Task Sdk_copy_edit_and_reinspect_leaves_the_source_intact()
    {
        using var workspace = new TemporaryWorkspace();
        var sourcePath = Path.Combine(workspace.Root, "agreement.docx");
        await File.WriteAllBytesAsync(sourcePath, Agreement());
        var sourceHashBefore = Sha256(await File.ReadAllBytesAsync(sourcePath));

        var client = workspace.Client();
        var source = await client.RegisterAsync("workspace", sourcePath);

        // Stage 1: take an authorized snapshot through the provider. The canonical
        // reference carries the provider version the copy was taken at.
        byte[] snapshotBytes;
        string snapshotVersion;
        using (var content = await client.OpenReadAsync(source))
        {
            snapshotVersion = content.Reference.Version!;
            using var buffer = new MemoryStream();
            await content.Stream.CopyToAsync(buffer);
            snapshotBytes = buffer.ToArray();
        }

        // Stage 2: edit a separate output with the SDK. Nothing is handed a live
        // OfficeAgent package and the source file is never opened for writing.
        var outputPath = Path.Combine(workspace.Root, "agreement-sdk.docx");
        await File.WriteAllBytesAsync(outputPath, snapshotBytes);
        using (var document = WordprocessingDocument.Open(outputPath, isEditable: true))
        {
            var body = document.MainDocumentPart!.Document!.Body!;
            body.AppendChild(new Paragraph(new Run(new Text(SdkAppendedClause))));
            document.MainDocumentPart.Document.Save();
        }

        // Stage 3: validate the SDK output before OfficeAgent is asked to trust it.
        using (var document = WordprocessingDocument.Open(outputPath, isEditable: false))
        {
            var errors = new OpenXmlValidator(FileFormatVersions.Office2019).Validate(document).ToList();
            Assert.Empty(errors);
        }

        Assert.Equal(sourceHashBefore, Sha256(await File.ReadAllBytesAsync(sourcePath)));
        using (var reopened = await client.OpenReadAsync(source))
            Assert.Equal(snapshotVersion, reopened.Reference.Version);

        // Stage 4: register and reinspect the SDK output, then author a fresh plan
        // against the new inspection and commit it.
        var output = await client.RegisterAsync("workspace", outputPath);
        var inspection = await client.InspectAsync(output);
        Assert.Contains(inspection.Paragraphs, paragraph => paragraph.Text == SdkAppendedClause);

        var hit = Assert.Single(await client.FindAsync(output, new FindQuery(SourceClause)));
        var plan = new DocumentPlan
        {
            Snapshot = inspection.Snapshot,
            Operations = new PlanOperation[]
            {
                new ChangeTextOp
                {
                    Target = hit.Anchor,
                    With = "Invoices are due in 45 days.",
                    Mode = ChangeMode.Tracked
                }
            }
        };

        var preview = await client.PreviewAsync(output, plan);
        Assert.True(preview.IsValid);

        string outputVersion;
        using (var current = await client.OpenReadAsync(output))
            outputVersion = current.Reference.Version!;

        var commit = await client.CommitAsync(output, plan, new SaveDocumentOptions
        {
            Mode = SaveMode.Replace,
            ExpectedVersion = outputVersion
        });
        Assert.True(commit.Committed);

        // The commit is scoped to the SDK output. The original registered source is
        // still byte-for-byte what it was before any of this ran.
        Assert.Equal(sourceHashBefore, Sha256(await File.ReadAllBytesAsync(sourcePath)));
    }

    /// <summary>
    /// A plan authored before an SDK text-host edit is refused as stale rather than
    /// silently reapplied against different content.
    /// </summary>
    [Fact]
    public async Task A_plan_authored_before_an_sdk_text_edit_is_refused_as_stale()
    {
        using var workspace = new TemporaryWorkspace();
        var path = Path.Combine(workspace.Root, "agreement.docx");
        await File.WriteAllBytesAsync(path, Agreement());

        var client = workspace.Client();
        var reference = await client.RegisterAsync("workspace", path);
        var inspection = await client.InspectAsync(reference);
        var hit = Assert.Single(await client.FindAsync(reference, new FindQuery(SourceClause)));
        var stalePlan = new DocumentPlan
        {
            Snapshot = inspection.Snapshot,
            Operations = new PlanOperation[]
            {
                new ChangeTextOp { Target = hit.Anchor, With = "Invoices are due in 45 days." }
            }
        };

        // An external SDK edit to a text host, outside any OfficeAgent operation.
        using (var document = WordprocessingDocument.Open(path, isEditable: true))
        {
            document.MainDocumentPart!.Document!.Body!
                .AppendChild(new Paragraph(new Run(new Text(SdkAppendedClause))));
            document.MainDocumentPart.Document.Save();
        }

        var afterHash = Sha256(await File.ReadAllBytesAsync(path));

        // The registered reference is pinned to the version it was registered at, so the
        // provider refuses to open the changed document at all. This is the first guard.
        await Assert.ThrowsAsync<DocumentVersionConflictException>(
            () => client.PreviewAsync(reference, stalePlan));

        // Deliberately dropping the version to reach the plan-snapshot guard: the plan is
        // still refused, now because it was authored against different content.
        var unpinned = Unpinned(reference);
        var report = await client.PreviewAsync(unpinned, stalePlan);

        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, error => error.Code == ValidationErrorCodes.StaleSnapshot);

        var rejected = await client.CommitAsync(unpinned, stalePlan);
        Assert.False(rejected.Committed);
        Assert.Equal(afterHash, Sha256(await File.ReadAllBytesAsync(path)));
    }

    /// <summary>
    /// The provider version, not the plan snapshot, is what stops a commit from
    /// overwriting a document an external writer has changed. It is checked when the
    /// reference carries the stale version and again at save time.
    /// </summary>
    [Fact]
    public async Task A_stale_provider_version_cannot_overwrite_an_sdk_changed_document()
    {
        using var workspace = new TemporaryWorkspace();
        var path = Path.Combine(workspace.Root, "agreement.docx");
        await File.WriteAllBytesAsync(path, Agreement());

        var client = workspace.Client();
        var registered = await client.RegisterAsync("workspace", path);

        string staleVersion;
        InspectResult inspection;
        using (var content = await client.OpenReadAsync(registered))
        {
            staleVersion = content.Reference.Version!;
            inspection = await client.InspectAsync(registered);
        }

        var hit = Assert.Single(await client.FindAsync(registered, new FindQuery(SourceClause)));

        // An external SDK writer changes the same document.
        using (var document = WordprocessingDocument.Open(path, isEditable: true))
        {
            document.MainDocumentPart!.Document!.Body!
                .AppendChild(new Paragraph(new Run(new Text(SdkAppendedClause))));
            document.MainDocumentPart.Document.Save();
        }

        var afterSdkHash = Sha256(await File.ReadAllBytesAsync(path));

        // A reference pinned to the stale version cannot even be opened.
        var pinned = DocumentReference.ForFileSystem("workspace", registered.ItemId, staleVersion);
        await Assert.ThrowsAsync<DocumentVersionConflictException>(
            () => client.OpenReadAsync(pinned));

        // With an unpinned reference the plan still reaches the save, where the
        // expected version is checked. Nothing is written.
        var unpinned = Unpinned(registered);
        var freshInspection = await client.InspectAsync(unpinned);
        var conflicting = new DocumentPlan
        {
            Snapshot = freshInspection.Snapshot,
            Operations = new PlanOperation[]
            {
                new ChangeTextOp { Target = hit.Anchor, With = "Invoices are due in 45 days." }
            }
        };

        await Assert.ThrowsAsync<DocumentVersionConflictException>(
            () => client.CommitAsync(unpinned, conflicting, new SaveDocumentOptions
            {
                Mode = SaveMode.Replace,
                ExpectedVersion = staleVersion
            }));

        Assert.Equal(afterSdkHash, Sha256(await File.ReadAllBytesAsync(path)));
        Assert.NotEqual(staleVersion, inspection.Snapshot.ETag);
    }

    /// <summary>
    /// The documented limit of the plan snapshot. A Word snapshot etag covers the text
    /// hosts only, so an SDK edit confined to another part leaves it unchanged and a
    /// plan authored before that edit is not refused. The provider version does change,
    /// which is why an external-change guard must be built on the provider version.
    /// </summary>
    [Fact]
    public async Task A_plan_snapshot_does_not_detect_sdk_edits_outside_the_text_hosts()
    {
        using var workspace = new TemporaryWorkspace();
        var path = Path.Combine(workspace.Root, "agreement.docx");
        await File.WriteAllBytesAsync(path, Agreement());

        var client = workspace.Client();
        var reference = await client.RegisterAsync("workspace", path);

        var before = await client.InspectAsync(reference);
        string versionBefore;
        using (var content = await client.OpenReadAsync(reference))
            versionBefore = content.Reference.Version!;

        // An SDK edit outside every text host: package metadata, not document content.
        using (var document = WordprocessingDocument.Open(path, isEditable: true))
            document.PackageProperties.Subject = "Externally edited by the Open XML SDK";

        // The pinned reference is refused because the bytes changed, which is the guard
        // that actually detects this edit.
        await Assert.ThrowsAsync<DocumentVersionConflictException>(
            () => client.InspectAsync(reference));

        var unpinned = Unpinned(reference);
        var after = await client.InspectAsync(unpinned);
        string versionAfter;
        using (var content = await client.OpenReadAsync(unpinned))
            versionAfter = content.Reference.Version!;

        // The plan snapshot is unchanged, so it cannot be used to prove no external
        // mutation happened.
        Assert.Equal(before.Snapshot.ETag, after.Snapshot.ETag);

        // The provider version is changed, so it can.
        Assert.NotEqual(versionBefore, versionAfter);
    }

    /// <summary>
    /// The agent and MCP surface is a fixed list of document operations. No tool
    /// accepts SDK code, a package callback, or an arbitrary expression, so no external
    /// mutation path bypasses provider authorization or host limits.
    /// </summary>
    [Fact]
    public void The_agent_tool_surface_exposes_no_arbitrary_execution_tool()
    {
        var store = new MemoryDocumentProvider("docs");
        var tools = new OfficeAgentTools(
            new OfficeAgentClient(new DocumentProviderRegistry(new[] { store }), new WordModule()));

        var names = tools
            .AsAIFunctions(new OfficeAgentToolsOptions
            {
                AllowRegistration = true,
                AllowCreation = true,
                AllowInlineContent = true,
                AllowEphemeralDocuments = true,
                AllowConnectionAddressing = true
            })
            .Select(function => function.Name)
            .ToArray();

        Assert.NotEmpty(names);
        foreach (var forbidden in new[]
                 {
                     "execute", "eval", "script", "code", "shell", "command",
                     "openxml", "sdk", "part", "xml", "package", "callback", "invoke"
                 })
        {
            Assert.DoesNotContain(names, name => name.Contains(forbidden, StringComparison.OrdinalIgnoreCase));
        }

        // The public client exposes no live-package callback either: the only package
        // handle is the engine's internal IOpenXmlPackage.
        Assert.DoesNotContain(
            typeof(OfficeAgentClient).GetMethods(),
            method => method.GetParameters().Any(parameter =>
                typeof(IOpenXmlPackage).IsAssignableFrom(parameter.ParameterType)) ||
                typeof(IOpenXmlPackage).IsAssignableFrom(method.ReturnType));
    }

    /// <summary>
    /// Drops the pinned provider version from a reference. Only a caller that has
    /// deliberately accepted external drift should do this.
    /// </summary>
    private static DocumentReference Unpinned(DocumentReference reference) =>
        DocumentReference.ForFileSystem(reference.ConnectionId, reference.ItemId);

    private static byte[] Agreement()
    {
        using var ms = new MemoryStream();
        using (var document = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Services agreement"))),
                new Paragraph(new Run(new Text(SourceClause))),
                new Paragraph(new Run(new Text("Signed for Adventure Works.")))));
            main.Document.Save();
        }

        return ms.ToArray();
    }

    private static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes));

    private sealed class TemporaryWorkspace : IDisposable
    {
        public TemporaryWorkspace()
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-sdk-interop-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);
        }

        public string Root { get; }

        public OfficeAgentClient Client() => new(
            new DocumentProviderRegistry(new[]
            {
                new FileSystemDocumentProvider(new FileSystemDocumentProviderOptions
                {
                    ConnectionId = "workspace",
                    RootPath = Root,
                    DefaultChangeMode = ChangeMode.Direct
                })
            }),
            new WordModule());

        public void Dispose()
        {
            if (Directory.Exists(Root)) Directory.Delete(Root, recursive: true);
        }
    }
}
