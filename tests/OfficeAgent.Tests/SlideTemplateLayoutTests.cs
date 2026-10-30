using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Validation;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using OfficeAgent.PowerPoint;
using A = DocumentFormat.OpenXml.Drawing;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Tests;

/// <summary>
/// A company template's own layouts. The built-in layout names cover ordinary authoring,
/// but a template's author designs layouts such as "Two Column With Logo" that no generic
/// name reaches; an agent building a deck on that template has to be able to ask for them
/// by the name PowerPoint shows, and to see which ones exist.
/// </summary>
public class SlideTemplateLayoutTests
{
    private static OfficeAgentClient Client() => new(new PowerPointModule());

    [Fact]
    public void Inspection_lists_the_deck_own_layouts_with_their_placeholders()
    {
        var layouts = Client().Inspect(CompanyTemplate()).Layouts;

        // The five a created deck ships with, then the company's own, in master order.
        Assert.Equal(
            new[] { "Title Slide", "Title and Content", "Section Header", "Title Only", "Blank",
                    "Two Column With Logo", "Quote", "Picture Only", "Agenda", "Agenda" },
            layouts.Select(l => l.Name));

        var twoColumn = layouts.Single(l => l.Name == "Two Column With Logo");
        // The footer is a running item, not content, so it is left out.
        Assert.Equal(
            new[] { ("Title 1", "title"), ("Left Column", "content"), ("Right Column", "content"), ("Logo", "picture") },
            twoColumn.Placeholders.Select(p => (p.Name, p.Kind)));
        Assert.Equal("Office Theme", twoColumn.Master);
    }

    [Fact]
    public void A_slide_goes_on_a_layout_named_by_the_template()
    {
        var client = Client();

        var deck = Apply(client, CompanyTemplate(), new InsertSlideOp
        {
            Slide = new SlideData
            {
                Layout = "two column with logo",
                Title = "Why Northwind",
                Body = new[] { "Twelve years in logistics", "Offices in four countries" }
            }
        });

        var (slide, layoutName) = LastSlide(deck);
        Assert.Equal("Two Column With Logo", layoutName);

        // Every content placeholder arrives, as PowerPoint's New Slide makes it, under the
        // layout's own names, so the second column and the logo can be filled next.
        var shapes = slide.CommonSlideData!.ShapeTree!.Elements<Shape>().ToList();
        Assert.Equal(
            new[] { "Title 1", "Left Column", "Right Column", "Logo" },
            shapes.Select(s => s.NonVisualShapeProperties!.NonVisualDrawingProperties!.Name!.Value));
        Assert.Equal("Why Northwind", TextOfShape(shapes[0]));
        Assert.Equal("Twelve years in logisticsOffices in four countries", TextOfShape(shapes[1]));
        Assert.Equal(string.Empty, TextOfShape(shapes[2]));

        // Geometry is inherited, never pinned, and a picture frame carries no text body.
        Assert.All(shapes, s => Assert.Null(s.ShapeProperties?.Transform2D));
        Assert.Null(shapes[3].TextBody);
        AssertValid(deck);
    }

    [Fact]
    public void The_empty_column_can_be_filled_once_the_slide_exists()
    {
        var client = Client();
        var deck = Apply(client, CompanyTemplate(), new InsertSlideOp
        {
            Slide = new SlideData { Layout = "Two Column With Logo", Title = "Why Northwind", Body = new[] { "Left" } }
        });

        var right = client.Inspect(deck).Paragraphs
            .Where(p => p.Text.Length == 0)
            .Select(p => p.ParaId)
            .Last();

        var filled = Apply(client, deck, new ChangeTextOp
        {
            Target = new TextSpanAnchor { ParaId = right, Expect = "" },
            With = "Right",
            Mode = ChangeMode.Direct
        });

        var shapes = LastSlide(filled).Slide.CommonSlideData!.ShapeTree!.Elements<Shape>().ToList();
        Assert.Equal("Right", TextOfShape(shapes[2]));
        AssertValid(filled);
    }

    [Fact]
    public void The_built_in_names_still_win_and_keep_their_two_shape_slide()
    {
        var client = Client();

        var deck = Apply(client, CompanyTemplate(), new InsertSlideOp
        {
            Slide = new SlideData { Layout = "titleOnly", Title = "Plain" }
        });

        var (slide, layoutName) = LastSlide(deck);
        Assert.Equal("Title Only", layoutName);
        Assert.Single(slide.CommonSlideData!.ShapeTree!.Elements<Shape>());
    }

    [Fact]
    public void A_layout_name_shared_by_two_masters_is_refused_until_qualified()
    {
        var client = Client();

        var refused = Preview(client, CompanyTemplate(), new InsertSlideOp
        {
            Slide = new SlideData { Layout = "Agenda", Title = "Agenda" }
        });

        var error = Assert.Single(refused.Errors);
        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        Assert.Contains("'Office Theme/Agenda'", error.Message);
        Assert.Contains("'Brand Two/Agenda'", error.Message);

        var deck = Apply(client, CompanyTemplate(), new InsertSlideOp
        {
            Slide = new SlideData { Layout = "Brand Two/Agenda", Title = "Agenda" }
        });

        var (_, layoutName, master) = LastSlideWithMaster(deck);
        Assert.Equal("Agenda", layoutName);
        Assert.Equal("Brand Two", master);
        AssertValid(deck);
    }

    [Fact]
    public void An_unknown_name_is_refused_with_the_deck_own_layouts()
    {
        var report = Preview(Client(), CompanyTemplate(), new InsertSlideOp
        {
            Slide = new SlideData { Layout = "Three Column", Title = "Nope" }
        });

        var error = Assert.Single(report.Errors);
        Assert.Equal(ValidationErrorCodes.InvalidOperation, error.Code);
        Assert.Contains("titleAndContent", error.Message);
        Assert.Contains("'Two Column With Logo'", error.Message);
    }

    [Fact]
    public void Text_the_layout_has_no_place_for_is_refused_rather_than_dropped()
    {
        var client = Client();

        var noTitle = Preview(client, CompanyTemplate(), new InsertSlideOp
        {
            Slide = new SlideData { Layout = "Quote", Title = "A title", Body = new[] { "Words" } }
        });
        Assert.Contains("no title placeholder", Assert.Single(noTitle.Errors).Message);

        var noBody = Preview(client, CompanyTemplate(), new InsertSlideOp
        {
            Slide = new SlideData { Layout = "Picture Only", Body = new[] { "Words" } }
        });
        Assert.Contains("no text placeholder", Assert.Single(noBody.Errors).Message);

        // The same rule holds for a built-in name: a title-only slide has nowhere for bullets.
        var builtIn = Preview(client, CompanyTemplate(), new InsertSlideOp
        {
            Slide = new SlideData { Layout = "titleOnly", Title = "T", Body = new[] { "Lost" } }
        });
        Assert.Contains("no text placeholder", Assert.Single(builtIn.Errors).Message);
    }

    // ── the template ──────────────────────────────────────────────────────────

    /// <summary>
    /// A created deck with a company's own layouts added to its master, and a second master
    /// that repeats one layout name, as multi-brand templates do.
    /// </summary>
    internal static byte[] CompanyTemplate()
    {
        var buffer = new MemoryStream();
        var blank = new PowerPointModule().CreateBlank();
        buffer.Write(blank, 0, blank.Length);

        using (var document = PresentationDocument.Open(buffer, isEditable: true))
        {
            var main = document.PresentationPart!;
            var master = main.SlideMasterParts.Single();
            uint id = 2147483700U;

            AddLayout(master, "Two Column With Logo", ref id,
                Placeholder(2, "Title 1", PlaceholderValues.Title, null),
                Placeholder(3, "Left Column", null, 1U),
                Placeholder(4, "Right Column", null, 2U),
                Placeholder(5, "Logo", PlaceholderValues.Picture, 13U),
                Placeholder(6, "Footer Placeholder 5", PlaceholderValues.Footer, 11U));
            AddLayout(master, "Quote", ref id,
                Placeholder(2, "Quote Text", PlaceholderValues.Body, 1U));
            AddLayout(master, "Picture Only", ref id,
                Placeholder(2, "Picture", PlaceholderValues.Picture, 1U));
            AddLayout(master, "Agenda", ref id,
                Placeholder(2, "Title 1", PlaceholderValues.Title, null),
                Placeholder(3, "Items", null, 1U));

            // A second brand: its own master and theme, repeating the "Agenda" name.
            var second = main.AddNewPart<SlideMasterPart>();
            var theme = second.AddNewPart<ThemePart>();
            theme.Theme = (A.Theme)master.ThemePart!.Theme!.CloneNode(true);
            theme.Theme.Name = "Brand Two";
            second.SlideMaster = (SlideMaster)master.SlideMaster!.CloneNode(true);
            second.SlideMaster.SlideLayoutIdList = new SlideLayoutIdList();
            main.Presentation!.SlideMasterIdList!.Append(new SlideMasterId
            {
                Id = id++,
                RelationshipId = main.GetIdOfPart(second)
            });
            AddLayout(second, "Agenda", ref id,
                Placeholder(2, "Title 1", PlaceholderValues.Title, null),
                Placeholder(3, "Items", null, 1U));
        }

        return buffer.ToArray();
    }

    private static void AddLayout(SlideMasterPart master, string name, ref uint id, params Shape[] placeholders)
    {
        var layout = master.AddNewPart<SlideLayoutPart>();
        layout.AddPart(master);

        var tree = new ShapeTree(
            new NonVisualGroupShapeProperties(
                new NonVisualDrawingProperties { Id = 1U, Name = string.Empty },
                new NonVisualGroupShapeDrawingProperties(),
                new ApplicationNonVisualDrawingProperties()),
            new GroupShapeProperties());
        foreach (var placeholder in placeholders) tree.Append(placeholder);

        layout.SlideLayout = new SlideLayout(
            new CommonSlideData(tree) { Name = name },
            new ColorMapOverride(new A.MasterColorMapping()))
        { Type = SlideLayoutValues.Custom, Preserve = true };

        master.SlideMaster!.SlideLayoutIdList!.Append(new SlideLayoutId
        {
            Id = id++,
            RelationshipId = master.GetIdOfPart(layout)
        });
    }

    private static Shape Placeholder(uint shapeId, string name, PlaceholderValues? type, uint? index)
    {
        var placeholder = new PlaceholderShape();
        if (type is { } t) placeholder.Type = t;
        if (index is { } i) placeholder.Index = i;

        return new Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = shapeId, Name = name },
                new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(placeholder)),
            new ShapeProperties(new A.Transform2D(
                new A.Offset { X = 838200L * shapeId, Y = 1825625L },
                new A.Extents { Cx = 3000000L, Cy = 2000000L })),
            new TextBody(
                new A.BodyProperties(),
                new A.ListStyle(),
                new A.Paragraph(new A.EndParagraphRunProperties { Language = "en-US" })));
    }

    // ── helpers ───────────────────────────────────────────────────────────────

    private static (Slide Slide, string LayoutName) LastSlide(byte[] deck)
    {
        var (slide, layoutName, _) = LastSlideWithMaster(deck);
        return (slide, layoutName);
    }

    private static (Slide Slide, string LayoutName, string Master) LastSlideWithMaster(byte[] deck)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var main = document.PresentationPart!;
        var last = main.Presentation!.SlideIdList!.Elements<SlideId>().Last();
        var part = (SlidePart)main.GetPartById(last.RelationshipId!);
        var layout = part.SlideLayoutPart!;
        return (part.Slide!,
            layout.SlideLayout!.CommonSlideData!.Name!.Value!,
            layout.SlideMasterPart!.ThemePart!.Theme!.Name!.Value!);
    }

    private static string TextOfShape(Shape shape) =>
        string.Concat(shape.Descendants<A.Text>().Select(t => t.Text));

    private static byte[] Apply(OfficeAgentClient client, byte[] deck, params PlanOperation[] operations)
    {
        using var applied = client.Commit(new StreamHandle(new MemoryStream(deck)), Plan(operations));
        Assert.True(applied.Committed,
            string.Join("; ", applied.Report.Errors.Select(e => $"{e.Code}: {e.Message}")));
        return applied.ToBytes();
    }

    private static ChangeReport Preview(OfficeAgentClient client, byte[] deck, params PlanOperation[] operations) =>
        client.Preview(new StreamHandle(new MemoryStream(deck)), Plan(operations));

    private static DocumentPlan Plan(PlanOperation[] operations) => new()
    {
        Format = DocFormat.PowerPoint,
        Operations = operations
    };

    private static void AssertValid(byte[] deck)
    {
        using var stream = new MemoryStream(deck);
        using var document = PresentationDocument.Open(stream, isEditable: false);
        var problems = new OpenXmlValidator(FileFormatVersions.Office2019)
            .Validate(document)
            .Select(e => $"{e.Path?.XPath}: {e.Description}")
            .ToList();
        Assert.True(problems.Count == 0, string.Join("; ", problems.Take(3)));
    }
}
