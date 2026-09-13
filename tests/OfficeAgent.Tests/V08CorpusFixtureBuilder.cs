using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// Deterministic generation recipes for the checked-in v0.8 acceptance corpus.
/// Generated documents are fictional engineering fixtures, not customer documents.
/// </summary>
internal static class V08CorpusFixtureBuilder
{
    private const string Pixel =
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII=";

    private static readonly DateTimeOffset ZipTimestamp =
        new(2000, 1, 1, 0, 0, 0, TimeSpan.Zero);

    public static IReadOnlyDictionary<string, byte[]> BuildAll(string repositoryRoot) =>
        new Dictionary<string, byte[]>(StringComparer.Ordinal)
        {
            ["complex-contract.docx"] = NormalizePackage(ComplexContract(repositoryRoot)),
            ["assembly-source-a.docx"] = NormalizePackage(DocumentMergeTests.Fixture(
                "Proposal source", "17365D", header: "Proposal team")),
            ["assembly-source-b.docx"] = NormalizePackage(DocumentMergeTests.Fixture(
                "Landscape appendix", "703070", landscape: true)),
            ["assembly-source-incompatible.docx"] = NormalizePackage(IncompatibleAssemblySource()),
            ["template.docx"] = NormalizePackage(Template(duplicateTags: false)),
            ["template-duplicate-tags.docx"] = NormalizePackage(Template(duplicateTags: true)),
            ["deck-with-chart.pptx"] = NormalizePackage(DeckWithChart()),
            ["workbook-styles.xlsx"] = NormalizePackage(XlsxFactory.WorkbookWithTable()),
            ["workbook-shared-strings.xlsx"] = NormalizePackage(XlsxFactory.WorkbookWithSharedString("Northwind")),
            ["workbook-shared-formula.xlsx"] = NormalizePackage(XlsxFactory.WorkbookWithSharedFormula()),
            ["workbook-array-formula.xlsx"] = NormalizePackage(XlsxFactory.WorkbookWithArrayFormula()),
            ["comparison-original.docx"] = NormalizePackage(ComparisonDocument("Payment is due in 30 days.")),
            ["comparison-revised.docx"] = NormalizePackage(ComparisonDocument("Payment is due in 45 days.")),
            ["comparison-table-changed.docx"] = NormalizePackage(ComparisonDocument(
                "Payment is due in 45 days.", "Changed protected table value")),
            ["word-edit-plan.json"] = Utf8(
                """
                {
                  "contractVersion": "0.2",
                  "format": "Word",
                  "revision": {
                    "author": "Corpus Review",
                    "timestampUtc": "2026-09-12T10:00:00Z"
                  },
                  "operations": [
                    {
                      "op": "changeText",
                      "target": {
                        "paraId": "w14:00000002",
                        "expect": "Acme Corp",
                        "occurrence": 0
                      },
                      "with": "Globex Inc.",
                      "mode": "Tracked"
                    }
                  ]
                }
                """),
            ["mcp-config.json"] = Utf8(
                """
                {
                  "OfficeAgent": {
                    "AllowCreation": true,
                    "FileSystemConnections": [
                      {
                        "ConnectionId": "documents",
                        "RootPath": "/absolute/path/to/documents",
                        "AllowedExtensions": [ ".docx", ".pptx" ],
                        "DefaultChangeMode": "Direct"
                      }
                    ]
                  }
                }
                """),
            ["validation-error.json"] = Utf8(
                """
                {
                  "committed": false,
                  "errors": [
                    {
                      "code": "stale-snapshot",
                      "message": "The document changed after inspection."
                    }
                  ]
                }
                """),
            ["apply-receipt.json"] = Utf8(
                """
                {
                  "outcome": "Committed",
                  "timestampUtc": "2026-09-12T10:00:00Z",
                  "inputSha256": "0000000000000000000000000000000000000000000000000000000000000000",
                  "planSha256": "1111111111111111111111111111111111111111111111111111111111111111",
                  "outputSha256": "2222222222222222222222222222222222222222222222222222222222222222",
                  "revision": {
                    "author": "Corpus Review",
                    "timestampUtc": "2026-09-12T10:00:00Z"
                  },
                  "actor": {
                    "subject": "fixture-user",
                    "issuer": "fixture-host",
                    "displayName": "Fixture User"
                  }
                }
                """)
        };

    private static byte[] ComplexContract(string repositoryRoot)
    {
        var source = File.ReadAllBytes(Path.Combine(
            repositoryRoot, "samples", "documents", "services-agreement.docx"));
        var client = new OfficeAgentClient(new WordModule(new FixedTimeProvider(
            new DateTimeOffset(2026, 9, 12, 10, 0, 0, TimeSpan.Zero))));
        var inspection = client.Inspect(source);
        var payment = inspection.Paragraphs.Single(paragraph =>
            paragraph.Text.Contains("thirty days", StringComparison.Ordinal));
        var heading = inspection.Paragraphs.First();
        using var applied = client.Commit(
            new StreamHandle(new MemoryStream(source, writable: false)),
            new DocumentPlan
            {
                Revision = new RevisionMetadata
                {
                    Author = "Corpus Review",
                    TimestampUtc = new DateTimeOffset(2026, 9, 12, 10, 0, 0, TimeSpan.Zero)
                },
                Operations = new PlanOperation[]
                {
                    new NoteOp
                    {
                        Target = new TextSpanAnchor
                        {
                            ParaId = payment.ParaId,
                            Expect = "thirty days"
                        },
                        Kind = NoteKind.Footnote,
                        Text = "Generated corpus footnote.",
                        Mode = ChangeMode.Direct
                    },
                    new InsertImageOp
                    {
                        Target = new TextSpanAnchor
                        {
                            ParaId = heading.ParaId,
                            Expect = heading.Text
                        },
                        Base64Bytes = Pixel,
                        ImageType = "png",
                        WidthPx = 16,
                        HeightPx = 16,
                        Position = InsertPosition.After,
                        Mode = ChangeMode.Direct
                    }
                }
            });
        if (!applied.Committed)
            throw new InvalidOperationException(string.Join("; ",
                applied.Report.Errors.Select(error => $"{error.Code}: {error.Message}")));

        var bytes = applied.ToBytes();
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            var body = document.MainDocumentPart!.Document.Body!;
            var imageParagraph = body.Elements<Paragraph>()
                .Single(paragraph => paragraph.Descendants<Drawing>().Any());
            imageParagraph.ParagraphId = "0BADBEEF";
            var imagePart = document.MainDocumentPart.ImageParts.Single();
            var previousImageId = document.MainDocumentPart.GetIdOfPart(imagePart);
            document.MainDocumentPart.ChangeIdOfPart(imagePart, "image1");
            foreach (var blip in body.Descendants<DocumentFormat.OpenXml.Drawing.Blip>()
                .Where(blip => blip.Embed?.Value == previousImageId))
                blip.Embed = "image1";
            document.MainDocumentPart.ChangeIdOfPart(
                document.MainDocumentPart.FootnotesPart!, "footnotes1");
            var noteParagraphs = document.MainDocumentPart.FootnotesPart!.Footnotes!
                .Descendants<Paragraph>()
                .ToArray();
            for (var index = 0; index < noteParagraphs.Length; index++)
                noteParagraphs[index].ParagraphId = (0x0F000000 + index).ToString("X8");
            document.MainDocumentPart.FootnotesPart.Footnotes.Save();
            var finalSection = body.GetFirstChild<SectionProperties>();
            if (finalSection is null)
            {
                finalSection = new SectionProperties(new PageSize
                {
                    Width = 11906U,
                    Height = 16838U
                });
                body.Append(finalSection);
            }

            var sectionParagraph = new Paragraph(
                new ParagraphProperties(new SectionProperties(
                    new SectionType { Val = SectionMarkValues.NextPage },
                    new PageSize { Width = 11906U, Height = 16838U })),
                new Run(new Text("Generated schedule section.")),
                new SimpleField(new Run(new Text("1"))) { Instruction = "PAGE" });
            body.InsertBefore(sectionParagraph, finalSection);
            document.MainDocumentPart.Document.Save();
        }

        return stream.ToArray();
    }

    private static byte[] IncompatibleAssemblySource()
    {
        var bytes = DocumentMergeTests.Fixture("Source with a pending revision", "336633");
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var document = WordprocessingDocument.Open(stream, true))
        {
            document.MainDocumentPart!.Document.Body!.Elements<Paragraph>().First().Append(
                new InsertedRun(new Run(new Text("Pending")))
                {
                    Id = "9",
                    Author = "Corpus Review",
                    Date = new DateTimeValue(new DateTime(2026, 9, 12, 10, 0, 0, DateTimeKind.Utc))
                });
            document.MainDocumentPart.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] Template(bool duplicateTags)
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            var body = new Body(
                TaggedParagraph("CustomerName", 10, "CUSTOMER"),
                new Table(
                    new TableProperties(new TableStyle { Val = "TableGrid" }),
                    new TableGrid(
                        new GridColumn { Width = "3200" },
                        new GridColumn { Width = "1600" },
                        new GridColumn { Width = "1600" }),
                    Row("Description", "Quantity", "Price"),
                    Row("{{Description}}", "{{Quantity}}", "{{Price}}")),
                new Paragraph());
            if (duplicateTags)
                body.InsertAfter(TaggedParagraph("CustomerName", 11, "DUPLICATE"), body.FirstChild);
            main.Document = new Document(body);
            main.Document.Save();
        }
        return stream.ToArray();
    }

    private static byte[] DeckWithChart()
    {
        var module = new PowerPointModule();
        var client = new OfficeAgentClient(module);
        using var applied = client.Commit(
            new StreamHandle(new MemoryStream(PptxFactory.Deck(), writable: false)),
            new DocumentPlan
            {
                Format = OfficeAgent.Abstractions.DocumentFormat.PowerPoint,
                Operations = new PlanOperation[]
                {
                    new InsertChartOp
                    {
                        Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
                        Kind = ChartKind.ClusteredColumn,
                        Categories = new[] { "Q1", "Q2" },
                        Series = new[]
                        {
                            new ChartSeries
                            {
                                Name = "Revenue",
                                Values = new double?[] { 10, 12 }
                            }
                        },
                        Title = "Revenue",
                        Description = "Revenue by quarter"
                    }
                }
            });
        if (!applied.Committed)
            throw new InvalidOperationException(string.Join("; ",
                applied.Report.Errors.Select(error => $"{error.Code}: {error.Message}")));
        return applied.ToBytes();
    }

    private static byte[] ComparisonDocument(string paragraphText, string tableText = "Protected table value")
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddMainDocumentPart();
            main.Document = new Document(new Body(
                new Paragraph(new Run(new Text("Agreement"))) { ParagraphId = "00000001" },
                new Paragraph(new Run(new Text(paragraphText))) { ParagraphId = "00000002" },
                new Table(
                    new TableProperties(new TableWidth { Width = "4800", Type = TableWidthUnitValues.Dxa }),
                    new TableGrid(new GridColumn { Width = "4800" }),
                    new TableRow(new TableCell(new Paragraph(new Run(new Text(tableText))))))));
            main.Document.Save();
        }
        return stream.ToArray();
    }

    private static Paragraph TaggedParagraph(string tag, int id, string text) => new(
        new SdtRun(
            new SdtProperties(new Tag { Val = tag }, new SdtId { Val = id }),
            new SdtContentRun(new Run(new Text(text)))));

    private static TableRow Row(params string[] values) => new(values.Select(value =>
        new TableCell(new Paragraph(new Run(new Text(value))))));

    private static byte[] NormalizePackage(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes);
        stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            foreach (var entry in archive.Entries)
                entry.LastWriteTime = ZipTimestamp;
        }
        return stream.ToArray();
    }

    private static byte[] Utf8(string text) =>
        new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(text.Replace("\r\n", "\n") + "\n");

    private sealed class FixedTimeProvider : TimeProvider
    {
        private readonly DateTimeOffset _now;

        public FixedTimeProvider(DateTimeOffset now) => _now = now;

        public override DateTimeOffset GetUtcNow() => _now;
    }
}
