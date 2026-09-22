using System.IO.Compression;
using System.Security.Cryptography;
using DocumentFormat.OpenXml.Packaging;
using W = DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.Core.DocumentProviders;
using OfficeAgent.Excel;
using OfficeAgent.PowerPoint;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// One output per advertised operation and format, plus the workflows, each with what Office
/// itself must observe when it opens the file.
/// </summary>
/// <remarks>
/// The corpus is the source of all three evidence layers. <see cref="NativeCorpusTests"/>
/// checks package validity and meaning on every build. The native layer is run separately:
/// the generation test writes the files and a manifest of expectations, and
/// <c>scripts/native/office_check.ps1</c> opens each one in Word, PowerPoint or Excel.
/// Inputs are generated in code, so nothing here is a real customer document.
/// </remarks>
internal static class NativeCorpus
{
    /// <summary>One corpus entry: what was done, to what, and what Office must then see.</summary>
    internal sealed record Case(
        string Id,
        string Format,
        string Family,
        string Description,
        byte[] Input,
        byte[] Output,
        IReadOnlyDictionary<string, object> Expect,
        IReadOnlyList<string> AllowedChangedParts,
        bool ExportForVisualReview)
    {
        public string FileName => Id + Format switch { "word" => ".docx", "powerpoint" => ".pptx", _ => ".xlsx" };
    }

    public static IReadOnlyList<Case> Build()
    {
        var cases = new List<Case>();
        cases.AddRange(Word());
        cases.AddRange(PowerPoint());
        cases.AddRange(Excel());
        return cases;
    }

    // Every part is allowed to change unless a case narrows it; narrowed only where the input is
    // the rich fixture, whose untouched parts must survive byte for byte.
    private static readonly IReadOnlyList<string> AnyPart = new[] { "*" };

    // ── Word ──────────────────────────────────────────────────────────────────────────────

    private static IEnumerable<Case> Word()
    {
        var client = new OfficeAgentClient(new WordModule());
        var contract = DocxFactory.Contract();
        var rich = WordPreservationFixtureBuilder.Build();
        var hit = Find(client, contract, "Acme Corp");

        Case W(string id, string family, string description, byte[] input, PlanOperation[] ops,
            Dictionary<string, object> expect, IReadOnlyList<string>? allowed = null, bool visual = false) =>
            new(id, "word", family, description, input, Commit(client, input, "input.docx", ops), expect, allowed ?? AnyPart, visual);

        // Revisions: the priority family. Accept and reject are checked in Word itself.
        yield return W("w-changetext-tracked", "changeText", "Tracked replacement of a phrase", contract,
            new PlanOperation[] { new ChangeTextOp { Target = hit, With = "Globex Inc.", Mode = ChangeMode.Tracked } },
            new() { ["revisionsAtLeast"] = 2, ["acceptContains"] = new[] { "Globex Inc." }, ["rejectContains"] = new[] { "Acme Corp" },
                    ["rejectExcludes"] = new[] { "Globex" } }, visual: true);

        var firstRound = Commit(client, contract, "input.docx", new PlanOperation[]
        {
            new ChangeTextOp { Target = hit, With = "Globex Inc.", Mode = ChangeMode.Tracked }
        }, author: "Alice Reviewer");
        var client2 = Paragraphs(client, firstRound).First(p => p.Text.Contains("Client:"));
        yield return new Case("w-successive-authors", "word", "revision",
            "Two authors' tracked edits in successive commits", contract,
            Commit(client, firstRound, "input.docx", new PlanOperation[]
            {
                new InsertOp { Target = new TextSpanAnchor { ParaId = client2.ParaId, Expect = "Client:" },
                    Position = InsertPosition.After, Text = "Northwind Traders joins as guarantor.", Mode = ChangeMode.Tracked }
            }, author: "Bob Approver"),
            new Dictionary<string, object>
            {
                ["revisionAuthors"] = new[] { "Alice Reviewer", "Bob Approver" },
                ["acceptContains"] = new[] { "Globex Inc.", "Northwind Traders" },
                ["rejectContains"] = new[] { "Acme Corp" }, ["rejectExcludes"] = new[] { "Northwind", "Globex" }
            }, AnyPart, true);

        var target = Paragraphs(client, contract).First(p => p.Text.Length > 0);
        var anchor = new TextSpanAnchor { ParaId = target.ParaId, Expect = "" };
        yield return W("w-insertparagraphs-tracked", "insertParagraphs", "Tracked paragraph insertion", contract,
            new PlanOperation[] { new InsertParagraphsOp { Target = anchor, Position = InsertPosition.After, Mode = ChangeMode.Tracked,
                Paragraphs = new[] { new ParagraphData { Text = "Inserted obligation paragraph." } } } },
            new() { ["revisionsAtLeast"] = 1, ["acceptContains"] = new[] { "Inserted obligation paragraph." },
                    ["rejectExcludes"] = new[] { "Inserted obligation paragraph." } });
        yield return W("w-removeparagraph-tracked", "removeParagraph", "Tracked paragraph removal", contract,
            new PlanOperation[] { new RemoveParagraphOp { Target = new TextSpanAnchor { ParaId = target.ParaId, Expect = target.Text }, Mode = ChangeMode.Tracked } },
            new() { ["revisionsAtLeast"] = 1, ["acceptExcludes"] = new[] { target.Text }, ["rejectContains"] = new[] { target.Text } });
        yield return W("w-format-tracked", "format", "Tracked formatting change", contract,
            new PlanOperation[] { new FormatOp { Target = hit, Bold = true, Color = "C00000", Mode = ChangeMode.Tracked } },
            new() { ["revisionsAtLeast"] = 1, ["acceptContains"] = new[] { "Acme Corp" } });
        yield return W("w-insert-direct", "insert", "Direct paragraph insertion after a found phrase", contract,
            new PlanOperation[] { new InsertOp { Target = hit, Position = InsertPosition.After, Text = "The Supplier is Acme Corp.", Mode = ChangeMode.Direct } },
            new() { ["revisions"] = 0, ["textContains"] = new[] { "The Supplier is Acme Corp." } });

        // Existing revisions in the rich fixture, resolved with the revision verb.
        yield return W("w-revision-accept-all", "revision", "Accept every pending revision", rich,
            new PlanOperation[] { new RevisionOp { Target = new NodeAnchor { Kind = "revision", Path = "all" }, Action = RevisionAction.Accept } },
            new() { ["revisions"] = 0, ["textContains"] = new[] { "approved" }, ["textExcludes"] = new[] { "draft" } });
        yield return W("w-revision-reject-all", "revision", "Reject every pending revision", rich,
            new PlanOperation[] { new RevisionOp { Target = new NodeAnchor { Kind = "revision", Path = "all" }, Action = RevisionAction.Reject } },
            new() { ["revisions"] = 0, ["textContains"] = new[] { "draft" }, ["textExcludes"] = new[] { "approved" } });

        // Comments, notes and content controls on the rich fixture, whose other parts must survive.
        var commented = Find(client, rich, "Numbered obligation");
        yield return W("w-comment-add", "comment", "Comment beside existing review threads", rich,
            new PlanOperation[] { new CommentOp { Target = commented, Text = "Confirm the obligation owner.", Author = "Legal" } },
            new() { ["commentsAtLeast"] = 2, ["textContains"] = new[] { "Numbered obligation" } },
            new[] { "word/document.xml", "word/comments.xml", "word/commentsExtended.xml", "word/commentsIds.xml",
                    "word/commentsExtensible.xml", "word/people.xml", "[Content_Types].xml", "word/_rels/document.xml.rels" }, true);
        yield return W("w-note-footnote", "note", "Footnote", contract,
            new PlanOperation[] { new NoteOp { Target = hit, Kind = NoteKind.Footnote, Text = "Registered in England." } },
            new() { ["footnotes"] = 1, ["footnoteText"] = "Registered in England." }, visual: true);
        yield return W("w-note-endnote", "note", "Endnote", contract,
            new PlanOperation[] { new NoteOp { Target = hit, Kind = NoteKind.Endnote, Text = "See schedule 2." } },
            new() { ["endnotes"] = 1, ["endnoteText"] = "See schedule 2." });
        yield return W("w-fill-control", "fill", "Fill a nested content control", rich,
            new PlanOperation[] { new FillOp { Target = new StructuralAnchor { Tag = "InnerParty" }, Value = "Contoso Holdings", Mode = ChangeMode.Direct } },
            new() { ["textContains"] = new[] { "Contoso Holdings" }, ["contentControlsAtLeast"] = 2 },
            new[] { "word/document.xml" });

        // Tables, built by the engine and then edited.
        var tabled = Commit(client, contract, "input.docx", new PlanOperation[]
        {
            new InsertTableOp { Target = anchor, Position = InsertPosition.After, Mode = ChangeMode.Direct,
                Table = new TableData { Headers = new[] { "Item", "Qty", "Price" }, Rows = new[] { new[] { "Widget", "2", "10" }, new[] { "Gadget", "1", "25" } } } }
        });
        var table = new NodeAnchor { Kind = "table", Path = "table#0" };
        yield return new Case("w-inserttable", "word", "insertTable", "Insert a 3x3 table", contract, tabled,
            new Dictionary<string, object> { ["tables"] = 1, ["tableRows"] = 3, ["tableColumns"] = 3, ["textContains"] = new[] { "Widget" } }, AnyPart, true);
        yield return W("w-inserttablerows", "insertTableRows", "Append table rows", tabled,
            new PlanOperation[] { new InsertTableRowsOp { Target = table, Rows = new[] { new[] { "Sprocket", "5", "3" } }, Position = TablePosition.End, Mode = ChangeMode.Direct } },
            new() { ["tables"] = 1, ["tableRows"] = 4, ["textContains"] = new[] { "Sprocket" } });
        yield return W("w-removetablerows", "removeTableRows", "Remove a table row", tabled,
            new PlanOperation[] { new RemoveTableRowsOp { Target = table, RowIndices = new[] { 2 }, Mode = ChangeMode.Direct } },
            new() { ["tables"] = 1, ["tableRows"] = 2, ["textExcludes"] = new[] { "Gadget" } });
        yield return W("w-inserttablecolumns", "insertTableColumns", "Insert a table column", tabled,
            new PlanOperation[] { new InsertTableColumnsOp { Target = table, Columns = new[] { new[] { "Total", "20", "25" } } } },
            new() { ["tables"] = 1, ["tableColumns"] = 4, ["textContains"] = new[] { "Total" } });
        yield return W("w-removetablecolumns", "removeTableColumns", "Remove a table column", tabled,
            new PlanOperation[] { new RemoveTableColumnsOp { Target = table, ColumnIndices = new[] { 2 }, Mode = ChangeMode.Direct } },
            new() { ["tables"] = 1, ["tableColumns"] = 2, ["textExcludes"] = new[] { "Price" } });
        yield return W("w-removetablecolumns-tracked", "removeTableColumns", "Tracked column removal, resolved in Word", tabled,
            new PlanOperation[] { new RemoveTableColumnsOp { Target = table, ColumnIndices = new[] { 2 }, Mode = ChangeMode.Tracked } },
            new() { ["revisionsAtLeast"] = 1, ["tableColumns"] = 3, ["acceptTableColumns"] = 2, ["rejectTableColumns"] = 3,
                    ["acceptExcludes"] = new[] { "Price" }, ["rejectContains"] = new[] { "Price" } });
        yield return W("w-repeattablerow", "repeatTableRow", "Repeat a template row per record", TemplateWithRow(),
            new PlanOperation[] { new RepeatTableRowOp { Target = new NodeAnchor { Kind = "table", Path = "table#0" }, TemplateRowIndex = 1, Mode = ChangeMode.Direct,
                Records = new[] { Record("Widget", "2"), Record("Gadget", "1"), Record("Sprocket", "5") } } },
            new() { ["tables"] = 1, ["tableRows"] = 4, ["textContains"] = new[] { "Widget", "Gadget", "Sprocket" }, ["textExcludes"] = new[] { "{{" } }, visual: true);
        yield return W("w-removetable", "removeTable", "Remove a table", tabled,
            new PlanOperation[] { new RemoveTableOp { Target = table, Mode = ChangeMode.Direct } },
            new() { ["tables"] = 0, ["textExcludes"] = new[] { "Widget" } });

        // Images and page-level features.
        var imaged = Commit(client, contract, "input.docx", new PlanOperation[]
        {
            new InsertImageOp { Target = anchor, Position = InsertPosition.After, Base64Bytes = Png(160, 80, 0x2E, 0x75, 0xB6),
                WidthPx = 160, HeightPx = 80, AltText = "Blue banner", Mode = ChangeMode.Direct }
        });
        yield return new Case("w-insertimage", "word", "insertImage", "Insert an inline picture", contract, imaged,
            new Dictionary<string, object> { ["inlineShapes"] = 1 }, AnyPart, true);
        var image = new NodeAnchor { Kind = "image", Path = Paths(client, imaged, "image").First() };
        yield return W("w-removeimage", "removeImage", "Remove the picture", imaged,
            new PlanOperation[] { new RemoveImageOp { Target = image, Mode = ChangeMode.Direct } },
            new() { ["inlineShapes"] = 0 });
        yield return W("w-backgroundimage", "backgroundImage", "Page background picture", contract,
            new PlanOperation[] { new BackgroundImageOp { Base64Bytes = Png(64, 64, 0xF2, 0xE8, 0xD5), ImageType = "png" } },
            new() { ["headerPictureBehindText"] = true }, visual: true);
        yield return W("w-headerfooter", "headerFooter", "Header, footer and page number", contract,
            new PlanOperation[] { new HeaderFooterOp { Header = "Contoso — Confidential", Footer = "Services agreement", ShowPageNumber = true } },
            new() { ["headerContains"] = "Contoso", ["footerContains"] = "Services agreement" }, visual: true);
        yield return W("w-pagesetup", "pageSetup", "Landscape A4 with wide margins", contract,
            new PlanOperation[] { new PageSetupOp { Orientation = PageOrientation.Landscape, PaperSize = "A4", MarginLeftTwips = 2160, MarginRightTwips = 2160 } },
            new() { ["landscape"] = true }, visual: true);
        yield return W("w-insertbreak", "insertBreak", "Page break", contract,
            new PlanOperation[] { new InsertBreakOp { Target = anchor, Kind = BreakKind.Page, Position = InsertPosition.After, Mode = ChangeMode.Direct } },
            new() { ["pagesAtLeast"] = 2 }, visual: true);
        yield return W("w-setproperty", "setProperty", "Document title property", contract,
            new PlanOperation[] { new SetPropertyOp { Target = new NodeAnchor { Kind = "docProperty", Path = "core/title" }, Value = "Master services agreement" } },
            new() { ["title"] = "Master services agreement" });

        // Styles.
        yield return W("w-definestyle", "defineStyle", "Define and apply a heading style", contract,
            new PlanOperation[] { new DefineStyleOp { StyleId = "ContractHeading", Name = "Contract Heading", Bold = true, SizeHalfPoints = 32, OutlineLevel = 1, Color = "1F3864" } },
            new() { ["styleExists"] = "Contract Heading" });
        yield return W("w-clearstyles", "clearStyles", "Clear direct formatting on a bolded phrase", Commit(client, contract, "input.docx",
                new PlanOperation[] { new FormatOp { Target = hit, Bold = true, Color = "C00000", Mode = ChangeMode.Direct } }),
            new PlanOperation[] { new ClearStylesOp { Target = hit, Scope = "all" } },
            new() { ["wordTextSameAsInput"] = true, ["revisions"] = 0 });
        // Known edge: the fixture's runs, like some non-Word producers', omit xml:space="preserve",
        // so Word hides their edge spaces. Rewritten runs carry it, and Word then shows the space.
        yield return W("w-clearstyles-unpreserved-space", "clearStyles", "Clear formatting on runs written without xml:space", rich,
            new PlanOperation[] { new ClearStylesOp { Target = Find(client, rich, "Quarterly target"), Scope = "all" } },
            new() { ["wordTextContains"] = new[] { "Quarterly target" } }, new[] { "word/document.xml" });
        var mixed = Paragraphs(client, rich).First(p => p.Text.StartsWith("Quarterly"));
        var plain = Paragraphs(client, rich).First(p => p.Text.StartsWith("Numbered"));
        yield return W("w-copystyles", "copyStyles", "Copy direct formatting between paragraphs", rich,
            new PlanOperation[] { new CopyStylesOp { Source = new TextSpanAnchor { ParaId = mixed.ParaId, Expect = "" },
                Target = new TextSpanAnchor { ParaId = plain.ParaId, Expect = "" }, Scope = "paragraph" } },
            new() { ["wordTextSameAsInput"] = true }, new[] { "word/document.xml" });

        // Workflows.
        var created = client.CreateBlank("new.docx");
        yield return W("w-create-with-plan", "create", "Blank document from OfficeAgent, then written", created,
            new PlanOperation[] { new ChangeTextOp { Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = "" }, With = "Created by OfficeAgent.", Mode = ChangeMode.Direct } },
            new() { ["textContains"] = new[] { "Created by OfficeAgent." }, ["revisions"] = 0, ["wordCompatibilityMode"] = 15 });

        var template = Commit(client, TemplateWithRow(), "input.docx", Array.Empty<PlanOperation>());
        var populated = Commit(client, template, "input.docx", new PlanOperation[]
        {
            new FillOp { Target = new StructuralAnchor { Tag = "CustomerName" }, Value = "Fabrikam Ltd", Mode = ChangeMode.Direct },
            new RepeatTableRowOp { Target = new NodeAnchor { Kind = "table", Path = "table#0" }, TemplateRowIndex = 1, Mode = ChangeMode.Direct,
                Records = new[] { Record("Consulting", "3"), Record("Support", "12") } }
        });
        yield return new Case("w-template-populated", "word", "template", "Template with a slot and a repeating row, populated",
            template, populated, new Dictionary<string, object> { ["textContains"] = new[] { "Fabrikam Ltd", "Consulting", "Support" },
                ["textExcludes"] = new[] { "{{", "CUSTOMER" }, ["tableRows"] = 3 }, AnyPart, true);

        var photoTemplate = PhotoTemplate();
        yield return new Case("w-template-image", "word", "template", "Template image slot bound through a template batch on the filesystem provider",
            photoTemplate, TemplateBatch(photoTemplate, ".docx", new TemplateBinding
            {
                Values = new Dictionary<string, string?> { ["CustomerName"] = "Fabrikam Ltd" },
                TypedValues = new Dictionary<string, TemplateValue>
                {
                    ["Photo"] = new TemplateImageValue { Base64Bytes = Png(240, 120, 0x2E, 0x75, 0xB6), ImageType = "png", AltText = "Product photograph", WidthPx = 240, HeightPx = 120 }
                },
                MissingValueBehavior = MissingTemplateValueBehavior.Empty,
                Mode = ChangeMode.Direct
            }),
            new Dictionary<string, object> { ["inlineShapes"] = 1, ["inlineShapeSizePt"] = new[] { 180.0, 90.0 }, ["textContains"] = new[] { "Quote for Fabrikam Ltd" }, ["textExcludes"] = new[] { "CUSTOMER" } },
            AnyPart, true);

        var cellRevised = Commit(client, tabled, "input.docx", new PlanOperation[]
        {
            new ChangeTextOp { Target = Find(client, tabled, "Gadget"), With = "Gizmo", Mode = ChangeMode.Direct }
        });
        var cellComparison = client.CompareDocuments(tabled, cellRevised, new DocumentComparisonOptions());
        yield return new Case("w-comparison-table-cell", "word", "comparison", "Comparison redline of a table-cell edit", tabled,
            CommitPlan(client, tabled, "input.docx", cellComparison.Plan ?? throw new InvalidOperationException("The cell comparison proposed no plan.")),
            new Dictionary<string, object> { ["revisionsAtLeast"] = 1, ["tables"] = 1, ["tableRows"] = 3,
                ["acceptContains"] = new[] { "Gizmo" }, ["acceptExcludes"] = new[] { "Gadget" },
                ["rejectContains"] = new[] { "Gadget" }, ["rejectExcludes"] = new[] { "Gizmo" } }, AnyPart, true);

        var revised = Commit(client, contract, "input.docx", new PlanOperation[]
        {
            new ChangeTextOp { Target = hit, With = "Globex Inc.", Mode = ChangeMode.Direct }
        });
        var comparison = client.CompareDocuments(contract, revised, new DocumentComparisonOptions());
        yield return new Case("w-comparison", "word", "comparison", "Comparison redline between two versions", contract,
            CommitPlan(client, contract, "input.docx", comparison.Plan ?? throw new InvalidOperationException("The comparison proposed no redline plan.")),
            new Dictionary<string, object> { ["revisionsAtLeast"] = 1, ["acceptContains"] = new[] { "Globex Inc." },
                ["rejectContains"] = new[] { "Acme Corp" }, ["rejectExcludes"] = new[] { "Globex" } }, AnyPart, true);

        var appendix = Commit(client, client.CreateBlank("appendix.docx"), "input.docx", new PlanOperation[]
        {
            new ChangeTextOp { Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = "" }, With = "Appendix A: rate card.", Mode = ChangeMode.Direct },
            new HeaderFooterOp { Header = "Appendix header" }
        });
        // The contract fixture holds rich content controls, which assembly refuses by design,
        // so the main document is built the same way as the appendix.
        var main = Commit(client, client.CreateBlank("main.docx"), "input.docx", new PlanOperation[]
        {
            new ChangeTextOp { Target = new TextSpanAnchor { ParaId = "auto-0000", Expect = "" }, With = "Main agreement between Acme Corp and Contoso.", Mode = ChangeMode.Direct },
            new HeaderFooterOp { Header = "Main header" }
        });
        var sources = new[] { main, appendix };
        var preview = client.PreviewMerge(sources);
        if (!preview.IsValid) throw new InvalidOperationException("Assembly preview failed: " + string.Join("; ", preview.Diagnostics.Select(d => d.Message)));
        var merged = client.CommitMerge(preview.Plan!, sources);
        yield return new Case("w-assembly", "word", "assembly", "Two documents assembled, each keeping its header", main,
            merged.Content ?? throw new InvalidOperationException("The assembly produced no content."),
            new Dictionary<string, object> { ["textContains"] = new[] { "Acme Corp", "Appendix A: rate card." }, ["sectionsAtLeast"] = 2,
                ["revisions"] = 0, ["wordCompatibilityMode"] = 15 }, AnyPart, true);
    }

    // ── PowerPoint ────────────────────────────────────────────────────────────────────────

    private static IEnumerable<Case> PowerPoint()
    {
        var module = new PowerPointModule();
        var client = new OfficeAgentClient(module);
        var deck = PptxFactory.Deck();
        var tabled = PptxFactory.DeckWithTable();
        var three = Commit(client, module.CreateBlank(), "input.pptx", new PlanOperation[]
        {
            new InsertSlideOp { Slide = new SlideData { Layout = "titleAndContent", Title = "Second", Body = new[] { "A point", "Another point" } } },
            new InsertSlideOp { Slide = new SlideData { Layout = "titleAndContent", Title = "Third", Body = new[] { "Closing point" } } }
        });
        var slides = Paths(client, three, "slide", ".pptx");
        var slide1 = new NodeAnchor { Kind = "slide", Path = slides[0] };

        Case P(string id, string family, string description, byte[] input, PlanOperation[] ops,
            Dictionary<string, object> expect, bool visual = false) =>
            new(id, "powerpoint", family, description, input, Commit(client, input, "input.pptx", ops), expect, AnyPart, visual);

        yield return new Case("p-insertslide", "powerpoint", "insertSlide", "Two slides added to a blank deck",
            module.CreateBlank(), three, new Dictionary<string, object> { ["slides"] = 3, ["textContains"] = new[] { "Second", "Closing point" } }, AnyPart, true);
        yield return P("p-removeslide", "removeSlide", "Remove a slide", three,
            new PlanOperation[] { new RemoveSlideOp { Target = new NodeAnchor { Kind = "slide", Path = slides[1] } } },
            new() { ["slides"] = 2, ["textExcludes"] = new[] { "Another point" } });
        yield return P("p-moveslide", "moveSlide", "Move the last slide first", three,
            new PlanOperation[] { new MoveSlideOp { Target = new NodeAnchor { Kind = "slide", Path = slides[2] }, Position = SlidePosition.Start } },
            new() { ["slides"] = 3, ["firstSlideContains"] = "Third" });
        yield return P("p-duplicateslide", "duplicateSlide", "Duplicate a slide", three,
            new PlanOperation[] { new DuplicateSlideOp { Target = new NodeAnchor { Kind = "slide", Path = slides[1] } } },
            new() { ["slides"] = 4 });
        yield return P("p-section", "section", "Add a section", three,
            new PlanOperation[] { new SectionOp { Action = SectionAction.Add, Name = "Financials", Target = new NodeAnchor { Kind = "slide", Path = slides[1] } } },
            new() { ["sectionsAtLeast"] = 2, ["sectionNames"] = new[] { "Financials" } });
        yield return P("p-transition", "transition", "Fade transition on every slide", three,
            new PlanOperation[] { new TransitionOp { Effect = "fade", DurationMs = 700 } },
            new() { ["slides"] = 3, ["transitionsOnAll"] = true });
        var shapePath = Paths(client, three, "shape", ".pptx").First();
        yield return P("p-animate", "animate", "Fade-in animation on a shape", three,
            new PlanOperation[] { new AnimateOp { Target = new NodeAnchor { Kind = "shape", Path = shapePath }, Effect = "fade" } },
            new() { ["animationsAtLeast"] = 1 });
        yield return P("p-headerfooter", "headerFooter", "Footer and slide numbers", three,
            new PlanOperation[] { new HeaderFooterOp { Footer = "Confidential — internal only", ShowSlideNumber = true } },
            new() { ["footerContains"] = "Confidential" }, visual: true);

        var withShape = Commit(client, deck, "input.pptx", new PlanOperation[]
        {
            new InsertShapeOp { Target = new NodeAnchor { Kind = "slide", Path = "slide#256" }, Text = new[] { "Accent" }, XPx = 40, YPx = 400, WidthPx = 300, HeightPx = 60 }
        });
        yield return new Case("p-insertshape", "powerpoint", "insertShape", "Insert a text box", deck, withShape,
            new Dictionary<string, object> { ["textContains"] = new[] { "Accent" } }, AnyPart, true);
        var box = client.Inspect(new StreamHandle(new MemoryStream(withShape), "d.pptx"), new InspectOptions()).Nodes
            .First(n => n.Kind == "shape" && n.Summary.Contains("Accent")).Path;
        yield return P("p-removeshape", "removeShape", "Remove the text box", withShape,
            new PlanOperation[] { new RemoveShapeOp { Target = new NodeAnchor { Kind = "shape", Path = box } } },
            new() { ["textExcludes"] = new[] { "Accent" } });

        var charted = Commit(client, module.CreateBlank(), "input.pptx", new PlanOperation[]
        {
            new InsertChartOp { Target = new NodeAnchor { Kind = "slide", Path = "slide#256" }, Kind = ChartKind.ClusteredColumn, Title = "Revenue",
                Categories = new[] { "Q1", "Q2", "Q3" }, Series = new[] { new ChartSeries { Name = "Revenue", Values = new double?[] { 10, 12, 15 } } },
                Description = "Revenue by quarter" }
        });
        yield return new Case("p-insertchart", "powerpoint", "insertChart", "Native column chart with embedded data",
            module.CreateBlank(), charted, new Dictionary<string, object> { ["charts"] = 1, ["chartEditable"] = true, ["chartSeriesValues"] = new[] { 10.0, 12.0, 15.0 } }, AnyPart, true);
        var chart = new NodeAnchor { Kind = "chart", Path = Paths(client, charted, "chart", ".pptx").First() };
        yield return P("p-updatechart", "updateChart", "Replace the chart's data", charted,
            new PlanOperation[] { new UpdateChartOp { Target = chart, Kind = ChartKind.Line, Categories = new[] { "A", "B" },
                Series = new[] { new ChartSeries { Name = "New", Values = new double?[] { 2, 3 } } }, Description = "Updated chart" } },
            new() { ["charts"] = 1, ["chartEditable"] = true, ["chartSeriesValues"] = new[] { 2.0, 3.0 } }, visual: true);

        yield return new Case("p-template-chart", "powerpoint", "template", "Template chart slot bound through a template batch on the filesystem provider",
            charted, TemplateBatch(charted, ".pptx", null, chartBinding: new TemplateChartValue
            {
                Kind = ChartKind.Bar, Title = "Bookings by region", Categories = new[] { "North", "South" },
                Series = new[] { new ChartSeries { Name = "Bookings", Values = new double?[] { 41, 58 } } }
            }),
            new Dictionary<string, object> { ["charts"] = 1, ["chartEditable"] = true, ["chartSeriesValues"] = new[] { 41.0, 58.0 } }, AnyPart, true);

        var pictured = Commit(client, deck, "input.pptx", new PlanOperation[]
        {
            new InsertImageOp { Target = new NodeAnchor { Kind = "slide", Path = "slide#256" }, Base64Bytes = Png(200, 100, 0x2E, 0x75, 0xB6), WidthPx = 200, HeightPx = 100 }
        });
        yield return new Case("p-insertimage", "powerpoint", "insertImage", "Insert a picture", deck, pictured,
            new Dictionary<string, object> { ["pictures"] = 1 }, AnyPart, true);
        yield return P("p-removeimage", "removeImage", "Remove the picture", pictured,
            new PlanOperation[] { new RemoveImageOp { Target = new NodeAnchor { Kind = "image", Path = Paths(client, pictured, "image", ".pptx").First() } } },
            new() { ["pictures"] = 0 });
        yield return P("p-backgroundimage", "backgroundImage", "Slide background picture", deck,
            new PlanOperation[] { new BackgroundImageOp { Target = new NodeAnchor { Kind = "slide", Path = "slide#256" }, Base64Bytes = Png(64, 64, 0xF2, 0xE8, 0xD5), ImageType = "png" } },
            new() { ["backgroundPicture"] = true }, visual: true);
        yield return P("p-insertmedia", "insertMedia", "Embedded video with a poster frame", module.CreateBlank(),
            new PlanOperation[] { new InsertMediaOp { Target = new NodeAnchor { Kind = "slide", Path = "slide#256" }, Kind = MediaKind.Video,
                Base64Bytes = Convert.ToBase64String(new byte[] { 0, 0, 0, 0x18, 0x66, 0x74, 0x79, 0x70 }), MediaType = "mp4",
                PosterBase64 = Png(48, 27, 0x40, 0x40, 0x40), XPx = 100, YPx = 120, WidthPx = 480, HeightPx = 270, AltText = "Walkthrough" } },
            new() { ["media"] = 1 });

        var tablePath = Paths(client, tabled, "table", ".pptx").First();
        var pptTable = new NodeAnchor { Kind = "table", Path = tablePath };
        yield return P("p-inserttable", "insertTable", "Insert a table", module.CreateBlank(),
            new PlanOperation[] { new InsertTableOp { Target = new NodeAnchor { Kind = "slide", Path = "slide#256" },
                Table = new TableData { Headers = new[] { "Region", "Q1" }, Rows = new[] { new[] { "EMEA", "41850" } } } } },
            new() { ["tables"] = 1, ["textContains"] = new[] { "EMEA" } }, visual: true);
        yield return P("p-inserttablerows", "insertTableRows", "Append a table row", tabled,
            new PlanOperation[] { new InsertTableRowsOp { Target = pptTable, Rows = new[] { new[] { "APAC", "12" } }, Position = TablePosition.End } },
            new() { ["tables"] = 1, ["textContains"] = new[] { "APAC" } });
        yield return P("p-removetablerows", "removeTableRows", "Remove a table row", tabled,
            new PlanOperation[] { new RemoveTableRowsOp { Target = pptTable, RowIndices = new[] { 1 } } },
            new() { ["tables"] = 1 });
        yield return P("p-inserttablecolumns", "insertTableColumns", "Insert a table column", tabled,
            new PlanOperation[] { new InsertTableColumnsOp { Target = pptTable, Columns = new[] { new[] { "Q2", "9" } } } },
            new() { ["tables"] = 1, ["textContains"] = new[] { "Q2" } });
        yield return P("p-removetablecolumns", "removeTableColumns", "Remove a table column", tabled,
            new PlanOperation[] { new RemoveTableColumnsOp { Target = pptTable, ColumnIndices = new[] { 1 } } },
            new() { ["tables"] = 1 });
        yield return P("p-removetable", "removeTable", "Remove the table", tabled,
            new PlanOperation[] { new RemoveTableOp { Target = pptTable } },
            new() { ["tables"] = 0 });

        // Text verbs on a deck with real layouts: the minimal factory deck's shapes carry no
        // geometry, so PowerPoint lays them out at zero size and a PDF of them shows nothing.
        var title = Find(client, three, "Second", ".pptx");
        yield return P("p-changetext", "changeText", "Replace the title text", three,
            new PlanOperation[] { new ChangeTextOp { Target = title, With = "Revised title", Mode = ChangeMode.Direct } },
            new() { ["textContains"] = new[] { "Revised title" } });
        yield return P("p-insert", "insert", "Insert a paragraph after the title", three,
            new PlanOperation[] { new InsertOp { Target = title, Position = InsertPosition.After, Text = "Draft for discussion", Mode = ChangeMode.Direct } },
            new() { ["textContains"] = new[] { "Draft for discussion" } });
        yield return P("p-format", "format", "Bold, coloured title", three,
            new PlanOperation[] { new FormatOp { Target = title, Bold = true, Color = "C00000" } },
            new() { ["textContains"] = new[] { "Second" } }, visual: true);
        yield return P("p-clearstyles", "clearStyles", "Clear direct formatting on the title", Commit(client, three, "input.pptx", new PlanOperation[] { new FormatOp { Target = title, Bold = true } }),
            new PlanOperation[] { new ClearStylesOp { Target = title, Scope = "all" } },
            new() { ["textContains"] = new[] { "Second" } });
        var body = Paragraphs(client, three, ".pptx").Where(p => p.Text.Length > 0).ToArray();
        yield return P("p-copystyles", "copyStyles", "Copy paragraph formatting between lines", three,
            new PlanOperation[] { new CopyStylesOp { Source = new TextSpanAnchor { ParaId = body[0].ParaId, Expect = "" },
                Target = new TextSpanAnchor { ParaId = body[1].ParaId, Expect = "" }, Scope = "paragraph" } },
            new() { ["slides"] = 3 });
        yield return P("p-comment", "comment", "Modern comment on a slide", deck,
            new PlanOperation[] { new CommentOp { Target = new NodeAnchor { Kind = "slide", Path = "slide#256" }, Text = "Check the figures.", Author = "Reviewer" } },
            new() { ["commentsAtLeast"] = 1 });
        var template = SlideTemplate(client, module);
        yield return P("p-fill", "fill", "Fill a named template shape", template,
            new PlanOperation[] { new FillOp { Target = new StructuralAnchor { Tag = "ClientName" }, Value = "Northwind Traders Limited", Mode = ChangeMode.Direct } },
            new() { ["textContains"] = new[] { "Northwind Traders Limited" }, ["textExcludes"] = new[] { "[CLIENT]" } }, visual: true);
    }

    // ── Excel ─────────────────────────────────────────────────────────────────────────────

    private static IEnumerable<Case> Excel()
    {
        var module = new ExcelModule();
        var client = new OfficeAgentClient(module);
        var blank = module.CreateBlank();

        Case X(string id, string family, string description, byte[] input, PlanOperation[] ops,
            Dictionary<string, object> expect, bool visual = false) =>
            new(id, "excel", family, description, input, Commit(client, input, "input.xlsx", ops), expect, AnyPart, visual);

        yield return X("x-setcell-values-and-formula", "setCell", "Text, numbers and a SUM formula", blank,
            new PlanOperation[]
            {
                new SetCellOp { Target = new CellAnchor { SheetId = 1, Address = "A1" }, Value = "Verified by OfficeAgent", ValueKind = SpreadsheetCellValueKind.String },
                new SetCellOp { Target = new CellAnchor { SheetId = 1, Address = "B2" }, Value = "12.5", ValueKind = SpreadsheetCellValueKind.Number },
                new SetCellOp { Target = new CellAnchor { SheetId = 1, Address = "B3" }, Value = "7.5", ValueKind = SpreadsheetCellValueKind.Number },
                new SetCellOp { Target = new CellAnchor { SheetId = 1, Address = "B4" }, Formula = "SUM(B2:B3)" }
            },
            new() { ["cells"] = new[]
            {
                Cell("A1", value: "Verified by OfficeAgent"), Cell("B2", value: "12.5"), Cell("B4", value: "20", formula: "=SUM(B2:B3)")
            }, ["recalculate"] = true }, visual: true);
        yield return X("x-comment", "comment", "Cell comment", blank,
            new PlanOperation[] { new CommentOp { Target = new CellAnchor { SheetId = 1, Address = "A1" }, Text = "Source: finance export.", Author = "Analyst" } },
            new() { ["commentsAtLeast"] = 1 });
        var workbook = XlsxFactory.WorkbookWithTable();
        var tablePath = Paths(client, workbook, "spreadsheetTable", ".xlsx").First();
        yield return X("x-appendtablerows", "appendTableRows", "Append rows to an Excel table", workbook,
            new PlanOperation[] { new AppendTableRowsOp { Target = new NodeAnchor { Kind = "spreadsheetTable", Path = tablePath }, Rows = new[] { new[] { "APAC", "15" } } } },
            new() { ["tables"] = 1, ["tableRowsGrewBy"] = 1, ["textContains"] = new[] { "APAC" } }, visual: true);
    }

    // ── refusals ──────────────────────────────────────────────────────────────────────────

    /// <summary>A plan the engine must refuse before writing, and the code it must refuse with.</summary>
    internal sealed record Refusal(string Id, string Format, string Description, byte[] Input, PlanOperation[] Operations, string Code);

    /// <summary>
    /// The published unsupported cases, each proved to be refused rather than written wrongly.
    /// </summary>
    public static IReadOnlyList<Refusal> Refusals()
    {
        var word = new OfficeAgentClient(new WordModule());
        var contract = DocxFactory.Contract();
        var main = word.CreateBlank("x.docx");
        return new[]
        {
            new Refusal("x-refuse-shared-formula-member", "excel", "setCell on one member of a shared-formula group",
                XlsxFactory.WorkbookWithSharedFormula(),
                new PlanOperation[] { new SetCellOp { Target = new CellAnchor { SheetId = 1, Address = "A1" }, Value = "100", ValueKind = SpreadsheetCellValueKind.Number } },
                "invalid-operation"),
            new Refusal("w-refuse-stale-expect", "word", "changeText whose expected text is no longer there", contract,
                new PlanOperation[] { new ChangeTextOp { Target = new TextSpanAnchor { ParaId = Paragraphs(word, contract).First(p => p.Text.Length > 0).ParaId, Expect = "text that is not there" }, With = "x", Mode = ChangeMode.Tracked } },
                "expect-mismatch"),
            new Refusal("p-refuse-tracked", "powerpoint", "a tracked edit on a deck, which has no tracked-change vocabulary", PptxFactory.Deck(),
                new PlanOperation[] { new ChangeTextOp { Target = Find(new OfficeAgentClient(new PowerPointModule()), PptxFactory.Deck(), PptxFactory.TitleText, ".pptx"), With = "x", Mode = ChangeMode.Tracked } },
                "invalid-operation"),
            new Refusal("p-refuse-missing-slot", "powerpoint", "fill naming a shape the deck does not have", PptxFactory.Deck(),
                new PlanOperation[] { new FillOp { Target = new StructuralAnchor { Tag = "NoSuchShape" }, Value = "x", Mode = ChangeMode.Direct } },
                "anchor-not-found"),
            new Refusal("w-refuse-missing-paragraph", "word", "removeParagraph on a paragraph id that does not exist", main,
                new PlanOperation[] { new RemoveParagraphOp { Target = new TextSpanAnchor { ParaId = "7FFFFFFF", Expect = "" }, Mode = ChangeMode.Tracked } },
                "anchor-not-found"),
        };
    }

    // ── helpers ───────────────────────────────────────────────────────────────────────────

    private static byte[] Commit(OfficeAgentClient client, byte[] input, string fileName, PlanOperation[] ops, string? author = null)
    {
        using var result = client.Commit(new StreamHandle(new MemoryStream(input, writable: false), fileName), new DocumentPlan
        {
            Operations = ops,
            Revision = new RevisionMetadata { Author = author ?? "OfficeAgent", TimestampUtc = new DateTimeOffset(2026, 9, 22, 12, 0, 0, TimeSpan.Zero) }
        });
        if (!result.Committed)
            throw new InvalidOperationException($"{string.Join(", ", ops.Select(op => op.GetType().Name))} did not commit: " +
                string.Join("; ", result.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return result.ToBytes();
    }

    private static byte[] CommitPlan(OfficeAgentClient client, byte[] input, string fileName, DocumentPlan plan)
    {
        using var result = client.Commit(new StreamHandle(new MemoryStream(input, writable: false), fileName), plan);
        if (!result.Committed)
            throw new InvalidOperationException("The plan did not commit: " + string.Join("; ", result.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return result.ToBytes();
    }

    /// <summary>
    /// Populates one output through the real template batch path, on the filesystem provider,
    /// and returns the bytes it wrote. A chart binding goes to the deck's only chart slot.
    /// </summary>
    private static byte[] TemplateBatch(byte[] template, string extension, TemplateBinding? binding, TemplateChartValue? chartBinding = null)
    {
        var root = Path.Combine(Path.GetTempPath(), $"officeagent-native-template-{Guid.NewGuid():N}");
        Directory.CreateDirectory(root);
        try
        {
            var client = new OfficeAgentClient(
                new DocumentProviderRegistry(new[]
                {
                    new FileSystemDocumentProvider(new FileSystemDocumentProviderOptions
                    {
                        ConnectionId = "workspace", RootPath = root, DefaultChangeMode = ChangeMode.Direct,
                        AllowedExtensions = new[] { ".docx", ".pptx", ".xlsx" }
                    })
                }),
                new WordModule(), new PowerPointModule());
            var path = Path.Combine(root, "template" + extension);
            File.WriteAllBytes(path, template);
            var reference = client.RegisterAsync("workspace", path).GetAwaiter().GetResult();
            if (chartBinding is not null)
            {
                var slot = client.DiscoverTemplateAsync(reference).GetAwaiter().GetResult().MediaSlots.Single(m => m.MediaKind == "chart");
                binding = new TemplateBinding
                {
                    TypedValues = new Dictionary<string, TemplateValue> { [slot.Name] = chartBinding },
                    MissingValueBehavior = MissingTemplateValueBehavior.Empty,
                    Mode = ChangeMode.Direct
                };
            }
            var request = new TemplateBatchRequest { Items = new[] { new TemplateBatchItem { OutputName = "out" + extension, Binding = binding! } } };
            var preview = client.PreviewTemplateBatchAsync(reference, request).GetAwaiter().GetResult();
            if (!preview.IsValid)
                throw new InvalidOperationException("Template preview failed: " + string.Join("; ",
                    preview.Diagnostics.Concat(preview.Items.SelectMany(i => i.Diagnostics)).Select(d => $"{d.Code}: {d.Message}")));
            var result = client.PopulateTemplateBatchAsync(reference, request, preview.Token).GetAwaiter().GetResult();
            if (!result.Committed)
                throw new InvalidOperationException("Template batch failed: " + string.Join("; ",
                    result.Items.SelectMany(i => i.Diagnostics).Select(d => $"{d.Code}: {d.Message}")));
            return File.ReadAllBytes(Path.Combine(root, "out" + extension));
        }
        finally
        {
            Directory.Delete(root, recursive: true);
        }
    }

    /// <summary>A quote template with a CustomerName text slot and a Photo slot, as a template author writes them.</summary>
    private static byte[] PhotoTemplate()
    {
        static W.SdtRun Control(string tag, int id, string text) => new(
            new W.SdtProperties(new W.Tag { Val = tag }, new W.SdtId { Val = id }),
            new W.SdtContentRun(new W.Run(new W.Text(text))));
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
            document.AddMainDocumentPart().Document = new W.Document(new W.Body(
                new W.Paragraph(new W.Run(new W.Text("Quote for ") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }), Control("CustomerName", 21, "CUSTOMER")),
                new W.Paragraph(Control("Photo", 22, "Photo")),
                new W.Paragraph()));
        return stream.ToArray();
    }

    private static InspectResult Inspect(OfficeAgentClient client, byte[] bytes, string extension) =>
        client.Inspect(new StreamHandle(new MemoryStream(bytes, writable: false), "doc" + extension), new InspectOptions { Fidelity = Fidelity.Content });

    private static IReadOnlyList<ParagraphInfo> Paragraphs(OfficeAgentClient client, byte[] bytes, string extension = ".docx") =>
        Inspect(client, bytes, extension).Paragraphs;

    private static IReadOnlyList<string> Paths(OfficeAgentClient client, byte[] bytes, string kind, string extension = ".docx") =>
        Inspect(client, bytes, extension).Nodes.Where(n => n.Kind == kind).Select(n => n.Path).ToArray();

    private static Anchor Find(OfficeAgentClient client, byte[] bytes, string text, string extension = ".docx") =>
        client.Find(new StreamHandle(new MemoryStream(bytes, writable: false), "doc" + extension), new FindQuery(text)).First().Anchor;

    private static IReadOnlyDictionary<string, string?> Record(string item, string quantity) =>
        new Dictionary<string, string?> { ["Description"] = item, ["Quantity"] = quantity };

    private static Dictionary<string, object> Cell(string address, string? value = null, string? formula = null)
    {
        var cell = new Dictionary<string, object> { ["address"] = address };
        if (value is not null) cell["value"] = value;
        if (formula is not null) cell["formula"] = formula;
        return cell;
    }

    /// <summary>A quote template: a CustomerName slot and a table with a {{placeholder}} row.</summary>
    private static byte[] TemplateWithRow()
    {
        using var stream = new MemoryStream();
        using (var document = WordprocessingDocument.Create(stream, DocumentFormat.OpenXml.WordprocessingDocumentType.Document))
        {
            static W.TableCell Cell(string text) => new(new W.Paragraph(new W.Run(new W.Text(text))));
            document.AddMainDocumentPart().Document = new W.Document(new W.Body(
                new W.Paragraph(new W.Run(new W.Text("Quote for ") { Space = DocumentFormat.OpenXml.SpaceProcessingModeValues.Preserve }),
                    new W.SdtRun(new W.SdtProperties(new W.Tag { Val = "CustomerName" }, new W.SdtId { Val = 11 }, new W.SdtContentText()),
                        new W.SdtContentRun(new W.Run(new W.Text("CUSTOMER"))))),
                new W.Table(
                    new W.TableProperties(new W.TableBorders(
                        new W.TopBorder { Val = W.BorderValues.Single, Size = 4 }, new W.BottomBorder { Val = W.BorderValues.Single, Size = 4 },
                        new W.InsideHorizontalBorder { Val = W.BorderValues.Single, Size = 4 })),
                    new W.TableGrid(new W.GridColumn { Width = "4000" }, new W.GridColumn { Width = "1500" }),
                    new W.TableRow(Cell("Description"), Cell("Quantity")),
                    new W.TableRow(Cell("{{Description}}"), Cell("{{Quantity}}"))),
                new W.Paragraph()));
        }
        return stream.ToArray();
    }

    /// <summary>A deck with a text box a template author named ClientName in the Selection Pane.</summary>
    private static byte[] SlideTemplate(OfficeAgentClient client, PowerPointModule module)
    {
        var deck = Commit(client, module.CreateBlank(), "input.pptx", new PlanOperation[]
        {
            new InsertShapeOp { Target = new NodeAnchor { Kind = "slide", Path = "slide#256" }, Text = new[] { "[CLIENT]" }, XPx = 80, YPx = 200, WidthPx = 600, HeightPx = 80 }
        });
        using var stream = new MemoryStream();
        stream.Write(deck);
        using (var document = PresentationDocument.Open(stream, true))
        {
            var shape = document.PresentationPart!.SlideParts.First().Slide!.Descendants<DocumentFormat.OpenXml.Presentation.Shape>()
                .First(s => s.InnerText.Contains("[CLIENT]"));
            shape.NonVisualShapeProperties!.NonVisualDrawingProperties!.Name = "ClientName";
        }
        return stream.ToArray();
    }

    /// <summary>A solid-colour PNG, base64, so pictures are visible when the output is rendered.</summary>
    internal static string Png(int width, int height, byte r, byte g, byte b)
    {
        var raw = new byte[height * (width * 3 + 1)];
        for (var y = 0; y < height; y++)
        for (var x = 0; x < width; x++)
        {
            var at = y * (width * 3 + 1) + 1 + x * 3;
            raw[at] = r; raw[at + 1] = g; raw[at + 2] = b;
        }
        using var png = new MemoryStream();
        png.Write(new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A });
        void Chunk(string type, byte[] data)
        {
            var length = BitConverter.GetBytes(data.Length); Array.Reverse(length); png.Write(length);
            var typed = System.Text.Encoding.ASCII.GetBytes(type).Concat(data).ToArray();
            png.Write(typed);
            var crc = BitConverter.GetBytes(Crc32(typed)); Array.Reverse(crc); png.Write(crc);
        }
        var header = new byte[13];
        BitConverter.GetBytes(width).Reverse().ToArray().CopyTo(header, 0);
        BitConverter.GetBytes(height).Reverse().ToArray().CopyTo(header, 4);
        header[8] = 8; header[9] = 2;
        Chunk("IHDR", header);
        using (var compressed = new MemoryStream())
        {
            using (var z = new ZLibStream(compressed, CompressionLevel.Optimal, leaveOpen: true)) z.Write(raw);
            Chunk("IDAT", compressed.ToArray());
        }
        Chunk("IEND", Array.Empty<byte>());
        return Convert.ToBase64String(png.ToArray());
    }

    private static uint Crc32(byte[] data)
    {
        var crc = 0xFFFFFFFFu;
        foreach (var d in data)
        {
            crc ^= d;
            for (var k = 0; k < 8; k++) crc = (crc & 1) != 0 ? (crc >> 1) ^ 0xEDB88320u : crc >> 1;
        }
        return ~crc;
    }

    public static string Sha256(byte[] bytes) => Convert.ToHexString(SHA256.HashData(bytes)).ToLowerInvariant();
}
