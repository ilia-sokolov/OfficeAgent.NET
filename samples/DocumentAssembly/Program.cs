using System.Text.Json;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Rendering;
using OfficeAgent.Word;

var render = args.Contains("--render", StringComparer.Ordinal);
var arguments = args.Where(a => a != "--render").ToArray();
if (arguments.Length < 2 || (arguments[0] != "--demo" && arguments.Length < 3))
{
    Console.Error.WriteLine("Usage: DocumentAssembly --demo output-directory [--render]\n       DocumentAssembly output.docx first.docx second.docx [more.docx] [--render]");
    return 2;
}

try
{
    string output;
    string[] paths;
    if (arguments[0] == "--demo")
    {
        var folder = Path.GetFullPath(arguments[1]);
        Directory.CreateDirectory(folder);
        paths = new[] { "proposal.docx", "statement-of-work.docx", "appendix.docx" }.Select(name => Path.Combine(folder, name)).ToArray();
        WriteNew(paths[0], Document("Northwind proposal", "17365D", "Proposal team", false,
            "A document automation engagement for Northwind Traders.", "We will implement an editable quote and document assembly workflow."));
        WriteNew(paths[1], Document("Statement of work", "336633", "Delivery team", false,
            "Scope: configure templates, validate outputs, and integrate provider storage.", "Acceptance: approved examples, preserved source content, and repeatable checks."));
        WriteNew(paths[2], Document("Appendix: delivery schedule", "703070", null, true,
            "The appendix uses landscape page geometry and intentionally has no running header.", "Dates and organizations in this sample are fictional."));
        output = Path.Combine(folder, "proposal-packet.docx");
    }
    else { output = Path.GetFullPath(arguments[0]); paths = arguments.Skip(1).Select(Path.GetFullPath).ToArray(); }
    if (!output.EndsWith(".docx", StringComparison.OrdinalIgnoreCase)) throw new ArgumentException("Output must be a .docx file.");
    var client = new OfficeAgentClient(new WordModule());
    var inputs = paths.Select(File.ReadAllBytes).ToArray();
    var preview = client.PreviewMerge(inputs, new() { Title = "Northwind proposal packet", Author = "DocumentAssembly sample" });
    Console.WriteLine(JsonSerializer.Serialize(preview, new JsonSerializerOptions { WriteIndented = true }));
    if (!preview.IsValid) return 1;
    // Re-read every source so a change between preview and commit fails as stale-merge-source.
    var result = client.CommitMerge(preview.Plan!, paths.Select(File.ReadAllBytes).ToArray());
    if (!result.Committed)
    {
        Console.Error.WriteLine(JsonSerializer.Serialize(result.Diagnostics));
        return 1;
    }
    WriteNew(output, result.Content!);
    WriteNew(output + ".receipt.json", JsonSerializer.SerializeToUtf8Bytes(result.Receipt, new JsonSerializerOptions { WriteIndented = true }));
    Console.WriteLine($"Created {output}");
    if (render)
    {
        var renderer = new LibreOfficeDocumentRenderer(new LibreOfficeRendererOptions
        {
            LibreOfficeExecutable = Environment.GetEnvironmentVariable("SOFFICE") ?? "soffice",
            PdfToPpmExecutable = Environment.GetEnvironmentVariable("PDFTOPPM") ?? "pdftoppm"
        });
        using var content = new MemoryStream(result.Content!, writable: false);
        var rendered = await renderer.RenderAsync(content, new RenderOptions { FileName = "packet.docx", MaximumPages = 20 });
        if (!rendered.Succeeded)
        {
            Console.Error.WriteLine(JsonSerializer.Serialize(rendered));
            return 1;
        }
        foreach (var page in rendered.Pages) WriteNew(output + $".page-{page.PageNumber}.png", page.Content);
        Console.WriteLine($"Rendered {rendered.Pages.Count} pages; inspect these images for layout acceptance.");
    }
    return 0;
}
catch (Exception error) when (error is IOException or ArgumentException or InvalidOperationException)
{
    Console.Error.WriteLine(error.Message);
    return 1;
}

static void WriteNew(string path, byte[] bytes)
{
    using var file = new FileStream(path, FileMode.CreateNew, FileAccess.Write);
    file.Write(bytes);
}

static byte[] Document(string heading, string color, string? header, bool landscape, params string[] paragraphs)
{
    using var output = new MemoryStream();
    using (var doc = WordprocessingDocument.Create(output, WordprocessingDocumentType.Document))
    {
        var main = doc.AddMainDocumentPart();
        main.AddNewPart<StyleDefinitionsPart>().Styles = new Styles(
            new DocDefaults(new RunPropertiesDefault(new RunPropertiesBaseStyle(new RunFonts { Ascii = "Calibri", HighAnsi = "Calibri" }, new FontSize { Val = "22" }))),
            new Style(new StyleName { Val = "Normal" }) { Type = StyleValues.Paragraph, StyleId = "Normal", Default = true },
            new Style(new StyleName { Val = "Heading" }, new BasedOn { Val = "Normal" },
                new StyleParagraphProperties(new SpacingBetweenLines { After = "240" }),
                new StyleRunProperties(new Bold(), new Color { Val = color }, new FontSize { Val = "36" }))
            { Type = StyleValues.Paragraph, StyleId = "Heading" });
        var body = new Body(new Paragraph(new ParagraphProperties(new ParagraphStyleId { Val = "Heading" }), new Run(new Text(heading))));
        foreach (var text in paragraphs) body.Append(new Paragraph(new ParagraphProperties(new SpacingBetweenLines { After = "160" }), new Run(new Text(text))));
        body.Append(new Table(new TableProperties(new TableWidth { Width = "8000", Type = TableWidthUnitValues.Dxa }),
            new TableGrid(new GridColumn { Width = "4000" }, new GridColumn { Width = "4000" }),
            new TableRow(new TableCell(new Paragraph(new Run(new Text("Deliverable")))), new TableCell(new Paragraph(new Run(new Text("Acceptance"))))),
            new TableRow(new TableCell(new Paragraph(new Run(new Text("Document workflow")))), new TableCell(new Paragraph(new Run(new Text("Validated examples")))))));
        body.Append(new Paragraph(new Run(new Text("Prepared for review."))));
        var section = new SectionProperties();
        if (header is not null)
        {
            var part = main.AddNewPart<HeaderPart>();
            part.Header = new Header(new Paragraph(new Run(new Text(header))));
            section.Append(new HeaderReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(part) });
        }
        var footer = main.AddNewPart<FooterPart>();
        footer.Footer = new Footer(new Paragraph(new SimpleField(new Run(new Text("1"))) { Instruction = "PAGE" }));
        section.Append(new FooterReference { Type = HeaderFooterValues.Default, Id = main.GetIdOfPart(footer) });
        section.Append(new PageSize { Width = landscape ? 16838U : 11906U, Height = landscape ? 11906U : 16838U, Orient = landscape ? PageOrientationValues.Landscape : PageOrientationValues.Portrait },
            new PageMargin { Top = 1080, Right = 1080, Bottom = 1080, Left = 1080, Header = 540, Footer = 540, Gutter = 0 });
        body.Append(section);
        main.Document = new Document(body);
    }
    return output.ToArray();
}
