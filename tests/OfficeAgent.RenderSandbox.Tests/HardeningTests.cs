using OfficeAgent.Abstractions;
using Xunit.Abstractions;

namespace OfficeAgent.RenderSandbox.Tests;

/// <summary>
/// LibreOffice's own behaviour inside the sandbox, and the refusal that keeps it from seeing
/// anything but the OOXML package it was promised.
/// </summary>
/// <remarks>
/// The macro and linked-file samples are also run through the permissive control image, which
/// sets every hardening key to its most permissive value. On LibreOffice 7.4.7 (Debian
/// 4:7.4.7-1+deb12u14) neither sample triggers even there: headless conversion ran no
/// document-event macro and followed no linked section. The control is therefore reported
/// rather than required, and the registry hardening stays defense in depth that these tests do
/// not prove effective. What they do prove is the property itself: on the reference image no
/// macro ran and no local file content reached the output.
/// </remarks>
[Collection(SandboxCollection.Name)]
public sealed class HardeningTests
{
    private readonly ITestOutputHelper _output;

    public HardeningTests(ITestOutputHelper output) => _output = output;

    [Fact]
    public async Task Content_that_is_not_the_declared_package_never_reaches_LibreOffice()
    {
        // LibreOffice detects format by content, so a .docx name alone would let any of its
        // import filters parse attacker input. The renderer refuses before the backend runs.
        foreach (var disguised in new[] { Documents.MacroDocumentDisguisedAsDocx(), "plain text"u8.ToArray() })
        {
            var result = await Sandbox.Renderer().RenderAsync(new MemoryStream(disguised),
                new RenderOptions { FileName = "disguised.docx" });
            Assert.False(result.Succeeded);
            Assert.Equal(RenderFailureCodes.RendererFailed, result.FailureCode);
            Assert.Contains("not a valid .docx package", result.Message);
        }
    }

    [Fact]
    public void A_document_event_macro_does_not_run()
    {
        var document = Documents.MacroDocumentDisguisedAsDocx();

        var hardened = Sandbox.Worker(Sandbox.Reference, document, "evil.odt", audit: true);
        Assert.False(hardened.GetProperty("audit").GetProperty("macroRan").GetBoolean());

        var control = Sandbox.Worker(Sandbox.PermissiveControl, document, "evil.odt", audit: true);
        _output.WriteLine("permissive control ran the macro: " +
                          control.GetProperty("audit").GetProperty("macroRan").GetBoolean());
    }

    [Fact]
    public void A_linked_local_file_is_not_pulled_into_the_output()
    {
        var document = Documents.LinkedLocalFile("/etc/passwd");

        var hardened = Sandbox.Worker(Sandbox.Reference, document, "linked.docx", audit: true);
        var text = hardened.GetProperty("audit").GetProperty("pdfText").GetString()!;
        Assert.Contains("Before the link.", text); // the document itself did render
        Assert.DoesNotContain("root:", text);

        var control = Sandbox.Worker(Sandbox.PermissiveControl, document, "linked.docx", audit: true);
        _output.WriteLine("permissive control included the linked file: " +
                          control.GetProperty("audit").GetProperty("pdfText").GetString()!.Contains("root:"));
    }

    [Fact]
    public void The_backend_is_the_pinned_version()
    {
        var versions = Sandbox.Probe("soffice --version; pdftoppm -v 2>&1 | head -1");
        Assert.Contains("LibreOffice 7.4.7.2", versions.Stdout);
        Assert.Contains("pdftoppm version 22.12.0", versions.Stdout);
    }
}
