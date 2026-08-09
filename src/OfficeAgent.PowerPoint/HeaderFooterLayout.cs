using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Presentation;
using OfficeAgent.Core;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// The master- and layout-side scaffolding the running items need.
/// </summary>
/// <remarks>
/// A slide's footer, date, and slide-number shapes carry no geometry of their own - like
/// any placeholder they inherit it. If the master and the slide's layout do not declare
/// them, PowerPoint has nothing to inherit from and drops each one in the top-left corner,
/// stacked. This adds them where they are missing, at the positions PowerPoint's own
/// masters use, and leaves any it finds alone so a corporate template keeps its placement.
/// </remarks>
internal static class HeaderFooterLayout
{
    // A 16:9 slide is 12192000 x 6858000 EMU; these are the standard footer-row boxes.
    private const long Row = 6356350L;
    private const long Height = 365125L;

    /// <summary>
    /// The three running items, with the placeholder indices PowerPoint's own masters use.
    /// </summary>
    /// <remarks>
    /// The index is not decoration: a slide's placeholder is matched to the layout's by
    /// <em>type and index together</em>. Omitting it makes PowerPoint fall back to the
    /// first placeholder on the layout - the title - so the footer, date and slide number
    /// all inherit the title's box and render stacked on top of it.
    /// </remarks>
    private static readonly (PlaceholderValues Type, uint Index, string Name, long X, long Cx, A.TextAlignmentTypeValues Align)[] Items =
    {
        (PlaceholderValues.DateAndTime, 10U, "Date Placeholder", 838200L, 2743200L, A.TextAlignmentTypeValues.Left),
        (PlaceholderValues.Footer, 11U, "Footer Placeholder", 4038600L, 4114800L, A.TextAlignmentTypeValues.Center),
        (PlaceholderValues.SlideNumber, 12U, "Slide Number Placeholder", 8610600L, 2743200L, A.TextAlignmentTypeValues.Right)
    };

    /// <summary>Declares the three placeholders on the master and every layout that lacks them.</summary>
    public static void EnsureOnMasterAndLayouts(IOpenXmlPackage package)
    {
        var main = PowerPointModel.Main(package);

        foreach (var master in main.SlideMasterParts)
        {
            EnsureIn(master.SlideMaster?.CommonSlideData?.ShapeTree, master.SlideMaster);

            foreach (var layout in master.SlideLayoutParts)
                EnsureIn(layout.SlideLayout?.CommonSlideData?.ShapeTree, layout.SlideLayout);
        }
    }

    private static void EnsureIn(ShapeTree? tree, OpenXmlElement? owner)
    {
        if (tree is null || owner is null) return;

        // A layout that opts out of the footer row - PowerPoint's Title Slide does by
        // default - still needs the placeholders present for a slide using it to inherit.
        foreach (var item in Items)
        {
            var present = tree.Elements<Shape>().Any(s =>
                s.NonVisualShapeProperties?.ApplicationNonVisualDrawingProperties?
                    .PlaceholderShape?.Type?.Value == item.Type);
            if (present) continue;

            tree.Append(Placeholder(NextId(tree), item));
        }
    }

    private static uint NextId(ShapeTree tree)
    {
        uint highest = 1;
        foreach (var properties in tree.Descendants<NonVisualDrawingProperties>())
            if (properties.Id?.Value is { } id && id > highest) highest = id;
        return highest + 1;
    }

    private static Shape Placeholder(uint shapeId, (PlaceholderValues Type, uint Index, string Name, long X, long Cx, A.TextAlignmentTypeValues Align) item) =>
        new(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = shapeId, Name = item.Name },
                new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                new ApplicationNonVisualDrawingProperties(
                    new PlaceholderShape { Type = item.Type, Index = item.Index, Size = PlaceholderSizeValues.Quarter })),
            new ShapeProperties(new A.Transform2D(
                new A.Offset { X = item.X, Y = Row },
                new A.Extents { Cx = item.Cx, Cy = Height })),
            new TextBody(
                new A.BodyProperties { Anchor = A.TextAnchoringTypeValues.Center },
                new A.ListStyle(
                    new A.Level1ParagraphProperties(
                        new A.DefaultRunProperties { FontSize = 1200 })
                    { Alignment = item.Align }),
                new A.Paragraph(new A.EndParagraphRunProperties { Language = "en-US" })));

    /// <summary>The slide-side shape, which carries only the placeholder reference.</summary>
    public static Shape SlideShape(uint shapeId, PlaceholderValues type)
    {
        var item = Items.First(i => i.Type == type);

        return new Shape(
            new NonVisualShapeProperties(
                new NonVisualDrawingProperties { Id = shapeId, Name = item.Name },
                new NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true }),
                // The index must match the layout's, or PowerPoint matches on type alone,
                // finds the title first, and the item inherits the title's box.
                new ApplicationNonVisualDrawingProperties(
                    new PlaceholderShape { Type = type, Index = item.Index, Size = PlaceholderSizeValues.Quarter })),
            // Empty on purpose: geometry comes from the layout.
            new ShapeProperties(),
            new TextBody(new A.BodyProperties(), new A.ListStyle()));
    }
}
