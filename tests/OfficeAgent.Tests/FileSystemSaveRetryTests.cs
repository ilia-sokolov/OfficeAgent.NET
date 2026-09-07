using Microsoft.Extensions.DependencyInjection;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// Saving the same document repeatedly, which is what an editing session actually does.
/// </summary>
/// <remarks>
/// On Windows the atomic write is where a save fails for a reason that has nothing to do
/// with the caller: a scanner or indexer opens the destination for a moment and
/// <c>File.Replace</c> answers "Unable to remove the file to be replaced". It surfaced as
/// an unexplained internal error roughly once in thirty saves on a developer machine, and
/// as a test suite that failed somewhere different every third run.
/// </remarks>
public class FileSystemSaveRetryTests
{
    [Fact]
    public async Task Repeated_saves_of_one_document_all_succeed()
    {
        using var workspace = new Workspace();
        var document = await workspace.Register();

        for (var step = 0; step < 40; step++)
        {
            var result = await workspace.Client.CommitAsync("workspace", document.ItemId, new DocumentPlan
            {
                Operations = new PlanOperation[]
                {
                    new InsertOp
                    {
                        Target = new TextSpanAnchor { ParaId = "w14:00000004", Expect = DocxFactory.DateText },
                        Position = InsertPosition.After,
                        Text = $"Clause {step}.",
                        Mode = ChangeMode.Direct
                    }
                }
            });

            Assert.True(result.Committed,
                $"save {step}: " + string.Join("; ", result.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        }
    }

    [Fact]
    public async Task Many_documents_saved_at_once_all_succeed()
    {
        using var workspace = new Workspace();

        // Concurrency is what makes the interference likely enough to catch: several
        // writes landing in one directory while the scanner is still following the last.
        var saves = Enumerable.Range(0, 12).Select(async index =>
        {
            var document = await workspace.Client.RegisterBytesAsync(
                "workspace", workspace.Root, DocxFactory.Contract(), $"contract-{index}.docx");

            for (var step = 0; step < 4; step++)
            {
                var result = await workspace.Client.CommitAsync("workspace", document.ItemId, new DocumentPlan
                {
                    Operations = new PlanOperation[]
                    {
                        // Appending leaves the anchor paragraph as it was, so every
                        // iteration targets text that is still there.
                        new InsertOp
                        {
                            Target = new TextSpanAnchor { ParaId = "w14:00000004", Expect = DocxFactory.DateText },
                            Position = InsertPosition.After,
                            Text = $"Clause {index}.{step}",
                            Mode = ChangeMode.Direct
                        }
                    }
                });
                Assert.True(result.Committed,
                    $"document {index} save {step}: " +
                    string.Join("; ", result.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));

                document = result.Document!;
            }
        });

        await Task.WhenAll(saves);
    }

    [Fact]
    public async Task A_destination_held_open_is_reported_as_a_provider_error()
    {
        using var workspace = new Workspace();
        var document = await workspace.Register();

        // The fixture is staged in a per-registration subdirectory, so find it rather than
        // assuming it sits at the root.
        var path = Directory.GetFiles(workspace.Root, "contract.docx", SearchOption.AllDirectories).Single();

        // A lock nothing is going to release - a document open in Word - must be reported,
        // and promptly, rather than retried until the request times out.
        using (File.Open(path, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
        {
            var failure = await Assert.ThrowsAsync<DocumentProviderException>(() =>
                workspace.Client.CommitAsync("workspace", document.ItemId, new DocumentPlan
                {
                    Operations = new PlanOperation[]
                    {
                        new ChangeTextOp
                        {
                            Target = new TextSpanAnchor { ParaId = "w14:00000001", Expect = DocxFactory.HeadingText },
                            With = "Blocked",
                            Mode = ChangeMode.Direct
                        }
                    }
                }));

            // Named as an IO problem with the file, not as an unexplained internal error.
            Assert.Equal(ProviderErrorCode.IO, failure.Code);
            Assert.Contains("contract.docx", failure.Message);
            Assert.Contains("another process", failure.Message);
        }

        // With the lock gone the same edit goes through, which is what makes it transient.
        var recovered = await workspace.Client.CommitAsync("workspace", document.ItemId, new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new ChangeTextOp
                {
                    Target = new TextSpanAnchor { ParaId = "w14:00000001", Expect = DocxFactory.HeadingText },
                    With = "Unblocked",
                    Mode = ChangeMode.Direct
                }
            }
        });
        Assert.True(recovered.Committed);
    }

    private sealed class Workspace : IDisposable
    {
        private readonly ServiceProvider _services;

        public string Root { get; }
        public OfficeAgentClient Client { get; }

        public Workspace()
        {
            Root = Path.Combine(Path.GetTempPath(), $"officeagent-save-{Guid.NewGuid():N}");
            Directory.CreateDirectory(Root);

            var services = new ServiceCollection();
            services.AddWordFormat();
            services.AddFileSystemDocumentProvider("workspace", Root);
            services.AddOfficeAgent();
            _services = services.BuildServiceProvider();
            Client = _services.GetRequiredService<OfficeAgentClient>();
        }

        public Task<DocumentReference> Register() =>
            Client.RegisterBytesAsync("workspace", Root, DocxFactory.Contract(), "contract.docx");

        public void Dispose()
        {
            _services.Dispose();
            try { Directory.Delete(Root, recursive: true); } catch { }
        }
    }
}
