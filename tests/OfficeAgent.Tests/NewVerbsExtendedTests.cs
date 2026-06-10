using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

public class NewVerbsExtendedTests
{
    private static OfficeAgentClient Office() => new(new WordModule());
    private static StreamHandle Handle(byte[] bytes) => new(new MemoryStream(bytes), "doc.docx");

    [Fact]
    public void InsertTableRows_at_Before_index_0_prepends_new_header()
    {
        var bytes = BuildCountriesTable();
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertTableRowsOp
                {
                    Target = new NodeAnchor { Kind = "table", Path = "table#0" },
                    Rows = new[] { new[] { "Title", "Title2", "Title3" } },
                    Position = TablePosition.Before,
                    RowIndex = 0
                }
            }
        };

        var result = Office().Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var rows = doc.MainDocumentPart!.Document.Body!.Descendants<Table>().Single().Elements<TableRow>().ToList();
        Assert.Equal("Title", string.Concat(rows[0].Elements<TableCell>().First().Descendants<Text>().Select(t => t.Text)));
    }

    [Fact]
    public void InsertTableRows_at_After_negative_index_appends_after_last_row()
    {
        var bytes = BuildCountriesTable();
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertTableRowsOp
                {
                    Target = new NodeAnchor { Kind = "table", Path = "table#0" },
                    Rows = new[] { new[] { "NL", "17", "41850" } },
                    Position = TablePosition.After,
                    RowIndex = -1
                }
            }
        };

        var result = Office().Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var rows = doc.MainDocumentPart!.Document.Body!.Descendants<Table>().Single().Elements<TableRow>().ToList();
        Assert.Contains("NL", rows[^1].InnerText);
    }

    [Fact]
    public void InsertTableColumns_at_End_appends_column_values_per_row()
    {
        var bytes = BuildCountriesTable();
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertTableColumnsOp
                {
                    Target = new NodeAnchor { Kind = "table", Path = "table#0" },
                    Columns = new[] { new[] { "Capital", "Moscow", "Washington", "London" } },
                    Position = TablePosition.End
                }
            }
        };

        var result = Office().Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var rows = doc.MainDocumentPart!.Document.Body!.Descendants<Table>().Single().Elements<TableRow>().ToList();
        foreach (var row in rows)
            Assert.Equal(4, row.Elements<TableCell>().Count());
        Assert.Contains("Moscow", rows[1].InnerText);
    }

    [Fact]
    public void RemoveTableColumns_drops_specified_indices_in_every_row()
    {
        var bytes = BuildCountriesTable();
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new RemoveTableColumnsOp
                {
                    Target = new NodeAnchor { Kind = "table", Path = "table#0" },
                    ColumnIndices = new[] { 1 } // remove "Population mil" column
                }
            }
        };

        var result = Office().Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var rows = doc.MainDocumentPart!.Document.Body!.Descendants<Table>().Single().Elements<TableRow>().ToList();
        Assert.All(rows, r => Assert.Equal(2, r.Elements<TableCell>().Count()));
        Assert.DoesNotContain("Population", string.Join(" ", rows.SelectMany(r => r.Descendants<Text>().Select(t => t.Text))));
    }

    [Fact]
    public void CopyStyles_copies_run_properties_to_target_span()
    {
        var bytes = BuildSimpleDocWithFormatting();
        var inspect = Office().Inspect(Handle(bytes));

        var boldParaId = inspect.Paragraphs.First(p => p.Text == "I am bold.").ParaId;
        var plainParaId = inspect.Paragraphs.First(p => p.Text == "I am plain.").ParaId;

        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new CopyStylesOp
                {
                    Source = new TextSpanAnchor { ParaId = boldParaId, Expect = "" },
                    Target = new TextSpanAnchor { ParaId = plainParaId, Expect = "" },
                    Scope = "run"
                }
            }
        };

        var result = Office().Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var plainParagraph = doc.MainDocumentPart!.Document.Body!.Descendants<Paragraph>()
            .Single(p => p.InnerText == "I am plain.");
        Assert.NotNull(plainParagraph.Elements<Run>().First().RunProperties?.GetFirstChild<Bold>());
    }

    [Fact]
    public void CopyStyles_paragraph_scope_propagates_style_id()
    {
        var bytes = BuildSimpleDocWithFormatting();
        var inspect = Office().Inspect(Handle(bytes));

        var headingParaId = inspect.Paragraphs.First(p => p.StyleId == "Heading1").ParaId;
        var plainParaId = inspect.Paragraphs.First(p => p.Text == "I am plain.").ParaId;

        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new CopyStylesOp
                {
                    Source = new TextSpanAnchor { ParaId = headingParaId, Expect = "" },
                    Target = new TextSpanAnchor { ParaId = plainParaId, Expect = "" },
                    Scope = "paragraph"
                }
            }
        };

        var result = Office().Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var plainParagraph = doc.MainDocumentPart!.Document.Body!.Descendants<Paragraph>()
            .Single(p => p.InnerText == "I am plain.");
        Assert.Equal("Heading1", plainParagraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value);
    }

    [Fact]
    public void ClearStyles_run_scope_drops_direct_formatting()
    {
        var bytes = BuildSimpleDocWithFormatting();
        var inspect = Office().Inspect(Handle(bytes));
        var boldParaId = inspect.Paragraphs.First(p => p.Text == "I am bold.").ParaId;

        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new ClearStylesOp
                {
                    Target = new TextSpanAnchor { ParaId = boldParaId, Expect = "" },
                    Scope = "run"
                }
            }
        };

        var result = Office().Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        var paragraph = doc.MainDocumentPart!.Document.Body!.Descendants<Paragraph>()
            .Single(p => p.InnerText == "I am bold.");
        Assert.Null(paragraph.Elements<Run>().First().RunProperties);
    }

    [Fact]
    public void InsertImage_with_base64_adds_inline_drawing_and_image_part()
    {
        var bytes = BuildSimpleDocWithFormatting();
        var paraId = Office().Inspect(Handle(bytes)).Paragraphs.First().ParaId;

        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertImageOp
                {
                    Target = new TextSpanAnchor { ParaId = paraId, Expect = "" },
                    Base64Bytes = Convert.ToBase64String(MinimalPng()),
                    ImageType = "png",
                    WidthPx = 50, HeightPx = 50,
                    Position = InsertPosition.After,
                    AltText = "Logo"
                }
            }
        };

        var result = Office().Commit(Handle(bytes), plan);
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        Assert.Single(doc.MainDocumentPart!.ImageParts);
        Assert.Single(doc.MainDocumentPart!.Document.Body!.Descendants<Drawing>());
    }

    [Fact]
    public void RemoveImage_removes_the_drawing_addressed_by_image_node()
    {
        var bytes = BuildSimpleDocWithFormatting();
        var paraId = Office().Inspect(Handle(bytes)).Paragraphs.First().ParaId;
        var insert = Office().Commit(Handle(bytes), new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertImageOp
                {
                    Target = new TextSpanAnchor { ParaId = paraId, Expect = "" },
                    Base64Bytes = Convert.ToBase64String(MinimalPng()),
                    ImageType = "png",
                    WidthPx = 50, HeightPx = 50,
                    Position = InsertPosition.After
                }
            }
        });
        Assert.True(insert.Committed);

        var afterInsertBytes = OfficeAgentClient.ToBytes(insert);
        var inspect = Office().Inspect(Handle(afterInsertBytes));
        Assert.Contains(inspect.Nodes, n => n.Kind == "image" && n.Path == "image#0");

        var remove = Office().Commit(Handle(afterInsertBytes), new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new RemoveImageOp { Target = new NodeAnchor { Kind = "image", Path = "image#0" } }
            }
        });
        Assert.True(remove.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(remove)), false);
        Assert.Empty(doc.MainDocumentPart!.Document.Body!.Descendants<Drawing>());
        // Removing the only image releases its underlying resource - no orphaned ImagePart.
        Assert.Empty(doc.MainDocumentPart!.ImageParts);
    }

    [Fact]
    public void InsertImage_twice_with_identical_bytes_references_a_single_image_part()
    {
        var bytes = BuildSimpleDocWithFormatting();
        var paraId = Office().Inspect(Handle(bytes)).Paragraphs.First().ParaId;
        var png = Convert.ToBase64String(MinimalPng());

        var result = Office().Commit(Handle(bytes), new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertImageOp
                {
                    Target = new TextSpanAnchor { ParaId = paraId, Expect = "" },
                    Base64Bytes = png, ImageType = "png", WidthPx = 50, HeightPx = 50,
                    Position = InsertPosition.After
                },
                new InsertImageOp
                {
                    Target = new TextSpanAnchor { ParaId = paraId, Expect = "" },
                    Base64Bytes = png, ImageType = "png", WidthPx = 50, HeightPx = 50,
                    Position = InsertPosition.Before
                }
            }
        });
        Assert.True(result.Committed);

        using var doc = WordprocessingDocument.Open(new MemoryStream(OfficeAgentClient.ToBytes(result)), false);
        // Two drawings, but the identical bytes are stored once and shared by reference.
        Assert.Equal(2, doc.MainDocumentPart!.Document.Body!.Descendants<Drawing>().Count());
        Assert.Single(doc.MainDocumentPart!.ImageParts);
    }

    [Fact]
    public void RemoveImage_keeps_a_shared_image_part_until_the_last_reference_is_gone()
    {
        var bytes = BuildSimpleDocWithFormatting();
        var paraId = Office().Inspect(Handle(bytes)).Paragraphs.First().ParaId;
        var png = Convert.ToBase64String(MinimalPng());

        // Two drawings sharing one deduplicated image part.
        var inserted = OfficeAgentClient.ToBytes(Office().Commit(Handle(bytes), new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertImageOp { Target = new TextSpanAnchor { ParaId = paraId, Expect = "" }, Base64Bytes = png, ImageType = "png", Position = InsertPosition.After },
                new InsertImageOp { Target = new TextSpanAnchor { ParaId = paraId, Expect = "" }, Base64Bytes = png, ImageType = "png", Position = InsertPosition.Before }
            }
        }));

        // Removing one drawing must NOT drop the part the other still references.
        var afterFirst = OfficeAgentClient.ToBytes(Office().Commit(Handle(inserted), new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new RemoveImageOp { Target = new NodeAnchor { Kind = "image", Path = "image#0" } }
            }
        }));
        using (var doc = WordprocessingDocument.Open(new MemoryStream(afterFirst), false))
        {
            Assert.Single(doc.MainDocumentPart!.Document.Body!.Descendants<Drawing>());
            Assert.Single(doc.MainDocumentPart!.ImageParts);
        }

        // Removing the last drawing releases the now-orphaned part.
        var afterSecond = OfficeAgentClient.ToBytes(Office().Commit(Handle(afterFirst), new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new RemoveImageOp { Target = new NodeAnchor { Kind = "image", Path = "image#0" } }
            }
        }));
        using (var doc = WordprocessingDocument.Open(new MemoryStream(afterSecond), false))
        {
            Assert.Empty(doc.MainDocumentPart!.Document.Body!.Descendants<Drawing>());
            Assert.Empty(doc.MainDocumentPart!.ImageParts);
        }
    }

    [Fact]
    public void InsertImage_with_no_source_is_rejected()
    {
        var bytes = BuildSimpleDocWithFormatting();
        var paraId = Office().Inspect(Handle(bytes)).Paragraphs.First().ParaId;

        var report = Office().Preview(Handle(bytes), new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertImageOp
                {
                    Target = new TextSpanAnchor { ParaId = paraId, Expect = "" },
                    ImageType = "png"
                }
            }
        });

        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, e => e.Code == ValidationErrorCodes.InvalidOperation
            && e.Message.Contains("Base64Bytes", StringComparison.Ordinal));
    }

    [Fact]
    public void InsertImage_with_both_sources_is_rejected()
    {
        var bytes = BuildSimpleDocWithFormatting();
        var paraId = Office().Inspect(Handle(bytes)).Paragraphs.First().ParaId;

        var report = Office().Preview(Handle(bytes), new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new InsertImageOp
                {
                    Target = new TextSpanAnchor { ParaId = paraId, Expect = "" },
                    Base64Bytes = Convert.ToBase64String(MinimalPng()),
                    ImageConnectionId = "images",
                    ImageDocumentId = "abc123",
                    ImageType = "png"
                }
            }
        });

        Assert.False(report.IsValid);
        Assert.Contains(report.Errors, e => e.Code == ValidationErrorCodes.InvalidOperation
            && e.Message.Contains("cannot mix", StringComparison.Ordinal));
    }

    private static byte[] BuildCountriesTable()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();
            var table = new Table(
                new TableProperties(new TableStyle { Val = "TableGrid" }),
                Row("Country", "Population mil", "Surface sq km"),
                Row("RU", "146", "17098246"),
                Row("US", "332", "9833520"),
                Row("UK", "68",  "243610"));
            main.Document = new Document(new Body(table));
            main.Document.Save();
        }
        return ms.ToArray();
    }

    private static byte[] BuildSimpleDocWithFormatting()
    {
        using var ms = new MemoryStream();
        using (var doc = WordprocessingDocument.Create(ms, WordprocessingDocumentType.Document))
        {
            var main = doc.AddMainDocumentPart();

            var heading = new Paragraph { ParagraphId = "00000001" };
            heading.ParagraphProperties = new ParagraphProperties(new ParagraphStyleId { Val = "Heading1" });
            heading.AppendChild(new Run(new Text("Heading")));

            var boldRun = new Run(new Text("I am bold."));
            boldRun.RunProperties = new RunProperties(new Bold(), new DocumentFormat.OpenXml.Wordprocessing.Color { Val = "FF0000" });
            var boldPara = new Paragraph { ParagraphId = "00000002" };
            boldPara.AppendChild(boldRun);

            var plainPara = new Paragraph { ParagraphId = "00000003" };
            plainPara.AppendChild(new Run(new Text("I am plain.")));

            main.Document = new Document(new Body(heading, boldPara, plainPara));

            var stylePart = main.AddNewPart<StyleDefinitionsPart>();
            stylePart.Styles = new Styles(
                new Style { Type = StyleValues.Paragraph, StyleId = "Heading1", StyleName = new StyleName { Val = "heading 1" } });
            stylePart.Styles.Save();
            main.Document.Save();
        }
        return ms.ToArray();
    }

    private static TableRow Row(params string[] cells)
    {
        var row = new TableRow();
        foreach (var c in cells)
            row.AppendChild(new TableCell(new Paragraph(new Run(new Text(c) { Space = SpaceProcessingModeValues.Preserve }))));
        return row;
    }

    /// <summary>The smallest valid PNG: 1x1 transparent pixel (67 bytes).</summary>
    private static byte[] MinimalPng() => Convert.FromBase64String(
        "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR4nGP4//8/AwAI/AL+XJ/PNwAAAABJRU5ErkJggg==");
}
