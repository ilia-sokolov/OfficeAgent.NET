using System.IO.Compression;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using W15 = DocumentFormat.OpenXml.Office2013.Word;
using W16Cid = DocumentFormat.OpenXml.Office2019.Word.Cid;

namespace OfficeAgent.Tests;

/// <summary>
/// Creates the fictional, versioned Word fixture used by the v0.9 preservation matrix.
/// The fixed identifiers make comment, revision, numbering, field and section assertions
/// reproducible without presenting the fixture as an independently authored document.
/// </summary>
internal static class WordPreservationFixtureBuilder
{
    private static readonly DateTimeOffset ZipTimestamp =
        new(2000, 1, 1, 0, 0, 0, TimeSpan.Zero);

    public static byte[] Build()
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(
                   stream, WordprocessingDocumentType.Document))
        {
            var main = document.AddNewPart<MainDocumentPart>(
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml",
                "rId1");
            main.Document = new Document(new Body(
                MixedRunParagraph(),
                PriorRevisionParagraph(),
                CommentedParagraph(),
                NestedControl(),
                NumberedParagraph(),
                FieldParagraph(),
                SectionParagraph(),
                new SectionProperties(
                    new PageSize { Width = 11906U, Height = 16838U },
                    new PageMargin
                    {
                        Top = 1440,
                        Right = 1440U,
                        Bottom = 1440,
                        Left = 1440U
                    })));
            main.Document.AddNamespaceDeclaration(
                "w14", "http://schemas.microsoft.com/office/word/2010/wordml");

            AddNumbering(main);
            AddCommentMetadata(main);
            main.Document.Save();
        }

        return NormalizePackage(stream.ToArray());
    }

    private static Paragraph MixedRunParagraph() => new(
        new Run(new Text("Quarterly ")),
        new Run(new RunProperties(new Bold()), new Text("tar")),
        new Run(new RunProperties(new Italic()), new Text("get")),
        new Run(new Text(" remains within plan.")))
    {
        ParagraphId = "10000001"
    };

    private static Paragraph PriorRevisionParagraph() => new(
        new Run(new Text("Prior review: ")),
        new InsertedRun(new Run(new Text("approved")))
        {
            Id = "100",
            Author = "Prior Author",
            Date = new DateTimeValue(new DateTime(2026, 8, 1, 9, 0, 0, DateTimeKind.Utc))
        },
        new Run(new Text(" / ")),
        new DeletedRun(new Run(new DeletedText("draft")))
        {
            Id = "101",
            Author = "Prior Author",
            Date = new DateTimeValue(new DateTime(2026, 8, 1, 9, 1, 0, DateTimeKind.Utc))
        })
    {
        ParagraphId = "10000002"
    };

    private static Paragraph CommentedParagraph() => new(
        new CommentRangeStart { Id = "1" },
        new Run(new Text("Commented clause")),
        new CommentRangeEnd { Id = "1" },
        new Run(new CommentReference { Id = "1" }))
    {
        ParagraphId = "10000003"
    };

    private static SdtBlock NestedControl() => new(
        new SdtProperties(
            new Tag { Val = "OuterClause" },
            new SdtId { Val = 2001 }),
        new SdtContentBlock(
            new Paragraph(
                new Run(new Text("Outer start ")),
                new SdtRun(
                    new SdtProperties(
                        new Tag { Val = "InnerParty" },
                        new SdtId { Val = 2002 }),
                    new SdtContentRun(new Run(new Text("Nested value")))),
                new Run(new Text(" outer end")))
            {
                ParagraphId = "10000004"
            }));

    private static Paragraph NumberedParagraph() => new(
        new ParagraphProperties(
            new NumberingProperties(
                new NumberingLevelReference { Val = 0 },
                new NumberingId { Val = 42 })),
        new Run(new Text("Numbered obligation")))
    {
        ParagraphId = "10000005"
    };

    private static Paragraph FieldParagraph() => new(
        new Run(new Text("Page ")),
        new SimpleField(new Run(new Text("7"))) { Instruction = "PAGE" })
    {
        ParagraphId = "10000006"
    };

    private static Paragraph SectionParagraph() => new(
        new ParagraphProperties(
            new SectionProperties(
                new SectionType { Val = SectionMarkValues.NextPage },
                new PageSize
                {
                    Width = 16838U,
                    Height = 11906U,
                    Orient = PageOrientationValues.Landscape
                })),
        new Run(new Text("Landscape appendix")))
    {
        ParagraphId = "10000007"
    };

    private static void AddNumbering(MainDocumentPart main)
    {
        var part = main.AddNewPart<NumberingDefinitionsPart>("numbering");
        part.Numbering = new Numbering(
            new AbstractNum(
                new Level(
                    new StartNumberingValue { Val = 1 },
                    new NumberingFormat { Val = NumberFormatValues.Decimal },
                    new LevelText { Val = "%1." })
                {
                    LevelIndex = 0
                })
            {
                AbstractNumberId = 10
            },
            new NumberingInstance(new AbstractNumId { Val = 10 })
            {
                NumberID = 42
            });
        part.Numbering.Save();
    }

    private static void AddCommentMetadata(MainDocumentPart main)
    {
        const string commentParaId = "20000001";

        var comments = main.AddNewPart<WordprocessingCommentsPart>("comments");
        comments.Comments = new Comments(
            new Comment(
                new Paragraph(new Run(new Text("Preserve this review thread.")))
                {
                    ParagraphId = commentParaId
                })
            {
                Id = "1",
                Author = "Reviewer One",
                Initials = "R1",
                Date = new DateTimeValue(new DateTime(2026, 8, 2, 10, 0, 0, DateTimeKind.Utc))
            });
        comments.Comments.AddNamespaceDeclaration(
            "w14", "http://schemas.microsoft.com/office/word/2010/wordml");
        comments.Comments.Save();

        var extended = main.AddNewPart<WordprocessingCommentsExPart>("commentsExtended");
        extended.CommentsEx = new W15.CommentsEx(
            new W15.CommentEx
            {
                ParaId = commentParaId,
                Done = false
            });
        extended.CommentsEx.Save();

        var ids = main.AddNewPart<WordprocessingCommentsIdsPart>("commentsIds");
        ids.CommentsIds = new W16Cid.CommentsIds(
            new W16Cid.CommentId
            {
                ParaId = commentParaId,
                DurableId = "D0000001"
            });
        ids.CommentsIds.Save();

        var people = main.AddNewPart<WordprocessingPeoplePart>("people");
        people.People = new W15.People(
            new W15.Person
            {
                Author = "Reviewer One"
            });
        people.People.Save();
    }

    public static byte[] NormalizePackage(byte[] bytes)
    {
        using var stream = new MemoryStream();
        stream.Write(bytes, 0, bytes.Length);
        stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true))
        {
            foreach (var entry in archive.Entries)
                entry.LastWriteTime = ZipTimestamp;
        }

        return stream.ToArray();
    }
}
