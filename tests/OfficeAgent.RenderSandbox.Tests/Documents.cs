using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeAgent.Core;
using OfficeAgent.PowerPoint;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeAgent.RenderSandbox.Tests;

/// <summary>
/// The render corpus and the adversarial samples, generated in code so every input is
/// reproducible and nothing confidential is ever committed.
/// </summary>
public static class Documents
{
    /// <summary>A Word document of <paramref name="pages"/> pages, each ending in a page break.</summary>
    public static byte[] Word(int pages, string text = "OfficeAgent renderer reference page")
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var body = new W.Body();
            for (var page = 1; page <= pages; page++)
            {
                body.Append(new W.Paragraph(new W.Run(new W.Text($"{text} {page}"))));
                if (page < pages)
                    body.Append(new W.Paragraph(new W.Run(new W.Break { Type = W.BreakValues.Page })));
            }
            document.AddMainDocumentPart().Document = new W.Document(body);
        }
        return stream.ToArray();
    }

    /// <summary>The blank deck OfficeAgent itself creates.</summary>
    public static byte[] Deck() => new OfficeAgentClient(new PowerPointModule()).CreateBlank("deck.pptx");

    /// <summary>A one-sheet workbook with a few values.</summary>
    public static byte[] Workbook()
    {
        using var stream = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
        {
            var workbook = document.AddWorkbookPart();
            workbook.Workbook = new Workbook();
            var sheet = workbook.AddNewPart<WorksheetPart>();
            var data = new SheetData();
            for (uint row = 1; row <= 5; row++)
                data.Append(new Row(new Cell
                {
                    CellReference = $"A{row}",
                    DataType = CellValues.String,
                    CellValue = new CellValue($"Row {row}")
                }) { RowIndex = row });
            sheet.Worksheet = new Worksheet(data);
            workbook.Workbook.AppendChild(new Sheets(new Sheet
            {
                Id = workbook.GetIdOfPart(sheet), SheetId = 1, Name = "Data"
            }));
        }
        return stream.ToArray();
    }

    /// <summary>
    /// An OpenDocument text file with a Basic macro bound to the document-load event, which
    /// writes a marker file. The renderer is given it under a .docx name: LibreOffice detects the
    /// format by content, so a file name is no protection, and this is how an attacker would try.
    /// </summary>
    public static byte[] MacroDocumentDisguisedAsDocx()
    {
        using var stream = new MemoryStream();
        using (var zip = new ZipArchive(stream, ZipArchiveMode.Create, leaveOpen: true))
        {
            // The mimetype entry must be first and stored uncompressed.
            Entry(zip, "mimetype", "application/vnd.oasis.opendocument.text", CompressionLevel.NoCompression);
            Entry(zip, "META-INF/manifest.xml", """
                <?xml version="1.0" encoding="UTF-8"?>
                <manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.2">
                 <manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>
                 <manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/>
                 <manifest:file-entry manifest:full-path="Basic/script-lc.xml" manifest:media-type="text/xml"/>
                 <manifest:file-entry manifest:full-path="Basic/Standard/script-lb.xml" manifest:media-type="text/xml"/>
                 <manifest:file-entry manifest:full-path="Basic/Standard/Module1.xml" manifest:media-type="text/xml"/>
                </manifest:manifest>
                """);
            Entry(zip, "content.xml", """
                <?xml version="1.0" encoding="UTF-8"?>
                <office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"
                  xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0"
                  xmlns:script="urn:oasis:names:tc:opendocument:xmlns:script:1.0"
                  xmlns:xlink="http://www.w3.org/1999/xlink" office:version="1.2">
                 <office:scripts>
                  <office:event-listeners>
                   <script:event-listener script:language="ooo:script" script:event-name="office:load"
                     xlink:href="vnd.sun.star.script:Standard.Module1.Main?language=Basic&amp;location=document"/>
                  </office:event-listeners>
                 </office:scripts>
                 <office:body><office:text><text:p>Macro sample</text:p></office:text></office:body>
                </office:document-content>
                """);
            Entry(zip, "Basic/script-lc.xml", """
                <?xml version="1.0" encoding="UTF-8"?>
                <!DOCTYPE library:libraries PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "libraries.dtd">
                <library:libraries xmlns:library="http://openoffice.org/2000/library" xmlns:xlink="http://www.w3.org/1999/xlink">
                 <library:library library:name="Standard" library:link="false"/>
                </library:libraries>
                """);
            Entry(zip, "Basic/Standard/script-lb.xml", """
                <?xml version="1.0" encoding="UTF-8"?>
                <!DOCTYPE library:library PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "library.dtd">
                <library:library xmlns:library="http://openoffice.org/2000/library" library:name="Standard" library:readonly="false" library:passwordprotected="false">
                 <library:element library:name="Module1"/>
                </library:library>
                """);
            Entry(zip, "Basic/Standard/Module1.xml", """
                <?xml version="1.0" encoding="UTF-8"?>
                <!DOCTYPE script:module PUBLIC "-//OpenOffice.org//DTD OfficeDocument 1.0//EN" "module.dtd">
                <script:module xmlns:script="http://openoffice.org/2000/script" script:name="Module1" script:language="StarBasic">Sub Main
                  Open "/tmp/officeagent-macro-marker" For Output As #1
                  Print #1, "ran"
                  Close #1
                End Sub
                </script:module>
                """);
        }
        return stream.ToArray();
    }

    /// <summary>
    /// A Word document whose INCLUDETEXT field names a file inside the container. If links were
    /// updated on load, that file's text would appear in the rendered output.
    /// </summary>
    public static byte[] LinkedLocalFile(string path = "/etc/passwd")
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            document.AddMainDocumentPart().Document = new W.Document(new W.Body(
                new W.Paragraph(new W.Run(new W.Text("Before the link."))),
                new W.Paragraph(
                    new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin }),
                    new W.Run(new W.FieldCode($" INCLUDETEXT \"{path}\" ") { Space = SpaceProcessingModeValues.Preserve }),
                    new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }),
                    new W.Run(new W.Text("placeholder")),
                    new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.End })),
                new W.Paragraph(new W.Run(new W.Text("After the link.")))));
        }
        return stream.ToArray();
    }

    private static void Entry(ZipArchive zip, string name, string content,
        CompressionLevel level = CompressionLevel.Optimal)
    {
        var entry = zip.CreateEntry(name, level);
        using var writer = new StreamWriter(entry.Open(), new UTF8Encoding(false));
        writer.Write(content.Trim());
    }
}
