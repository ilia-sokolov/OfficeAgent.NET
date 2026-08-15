using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// Solid colour on a shape and behind a slide.
/// </summary>
/// <remarks>
/// These are what turn a deck from "text on white" into something designed: a full-bleed
/// background, an accent bar, a card behind a statistic. DrawingML expresses all of them
/// as a solid fill, but in two different places - <c>p:spPr</c> for a shape and
/// <c>p:bg</c> for the slide - so both live here rather than in the format handler.
/// </remarks>
internal static class SlidePaint
{
    /// <summary>The literal a caller passes to mean "no fill" rather than a colour.</summary>
    public const string None = "none";

    public static bool IsNone(string? value) =>
        string.Equals(value, None, StringComparison.OrdinalIgnoreCase);

    /// <summary>Normalises <c>#1F3A5F</c> or <c>1f3a5f</c> to the <c>1F3A5F</c> DrawingML wants.</summary>
    public static string Hex(string value) => value.TrimStart('#').ToUpperInvariant();

    /// <summary>Whether the value is a usable colour: six hex digits, or the "none" literal.</summary>
    public static bool IsColour(string? value) =>
        value is not null &&
        (IsNone(value) ||
         (Hex(value).Length == 6 && Hex(value).All(Uri.IsHexDigit)));

    /// <summary>
    /// Paints the slide's background. <c>p:bg</c> is the first child of
    /// <c>p:cSld</c>, before the shape tree, so it is inserted rather than appended.
    /// </summary>
    public static void SetBackground(SlideRef slide, string color)
    {
        var common = slide.Part.Slide.CommonSlideData
            ?? throw new InvalidOperationException("Slide has no common slide data.");

        common.Background?.Remove();
        if (IsNone(color)) return;

        common.InsertAt(
            new Background(
                new BackgroundProperties(
                    new A.SolidFill(new A.RgbColorModelHex { Val = Hex(color) }),
                    new A.EffectList())),
            0);
    }

    /// <summary>
    /// Puts an image behind the slide, stretched to fill it.
    /// </summary>
    /// <remarks>
    /// The alpha lives on the <c>a:blip</c> as <c>a:alphaModFix</c> rather than on a shape
    /// over the top, so the slide has one background rather than a background and a scrim -
    /// and PowerPoint's own "Transparency" slider reads and writes the same attribute.
    /// </remarks>
    public static void SetBackgroundImage(SlideRef slide, string relationshipId, double? opacity)
    {
        var common = slide.Part.Slide.CommonSlideData
            ?? throw new InvalidOperationException("Slide has no common slide data.");

        common.Background?.Remove();

        var blip = new A.Blip { Embed = relationshipId };
        if (Alpha(opacity) is { } amount)
            blip.Append(new A.AlphaModulationFixed { Amount = amount });

        common.InsertAt(
            new Background(
                new BackgroundProperties(
                    new A.BlipFill(blip, new A.Stretch(new A.FillRectangle())),
                    new A.EffectList())),
            0);
    }

    /// <summary>Removes whatever background the slide has, colour or image.</summary>
    public static void ClearBackground(SlideRef slide) =>
        slide.Part.Slide.CommonSlideData?.Background?.Remove();

    /// <summary>Whether the value is a usable opacity.</summary>
    public static bool IsOpacity(double? value) => value is null or (>= 0 and <= 1);

    /// <summary>
    /// Converts 0-1 to the thousandths of a percent DrawingML counts in. Full strength
    /// returns null: writing <c>100%</c> is the same as writing nothing, and leaving it out
    /// keeps the markup the size PowerPoint would have written.
    /// </summary>
    public static int? Alpha(double? opacity) =>
        opacity is null or >= 1 ? null : (int)Math.Round(Math.Max(0, opacity!.Value) * 100000);

    /// <summary>
    /// Sets a shape's fill and outline. Values left null are not touched, so a plan can
    /// move a shape without repainting it, or repaint it without moving it.
    /// </summary>
    public static void SetShapeFill(OpenXmlElement element, string? fill, string? line, int? lineWidthPx)
    {
        if (fill is null && line is null && lineWidthPx is null) return;

        var properties = element switch
        {
            Shape s => s.ShapeProperties ??= new ShapeProperties(),
            Picture p => p.ShapeProperties ??= new ShapeProperties(),
            _ => throw new InvalidOperationException(
                $"Shape kind '{element.LocalName}' has no fill to set.")
        };

        // A shape with no preset geometry renders nothing however it is filled. The shapes
        // this module inserts carry a rectangle already, but a *placeholder* does not: it
        // takes its position from the layout and has no geometry of its own, so a fill set
        // on one is written to the file, survives a round trip, and is simply never drawn.
        // Giving it a rectangle is what PowerPoint itself does the moment you fill a
        // placeholder by hand.
        if ((fill is not null && !IsNone(fill)) || (line is not null && !IsNone(line)))
            EnsureGeometry(properties);

        if (fill is not null)
        {
            foreach (var existing in properties.Elements<A.SolidFill>().ToList()) existing.Remove();
            foreach (var existing in properties.Elements<A.NoFill>().ToList()) existing.Remove();

            // a:solidFill and a:noFill follow the geometry in the p:spPr sequence.
            OpenXmlElement paint = IsNone(fill)
                ? new A.NoFill()
                : new A.SolidFill(new A.RgbColorModelHex { Val = Hex(fill) });
            InsertAfterGeometry(properties, paint);
        }

        if (line is null && lineWidthPx is null) return;

        var outline = properties.GetFirstChild<A.Outline>();
        if (outline is null)
        {
            outline = new A.Outline();
            // a:ln is the last of the fill/line group in p:spPr.
            properties.Append(outline);
        }

        if (lineWidthPx is { } width) outline.Width = (int)Emu.FromPixels(width);
        if (line is not null)
        {
            foreach (var existing in outline.Elements<A.SolidFill>().ToList()) existing.Remove();
            foreach (var existing in outline.Elements<A.NoFill>().ToList()) existing.Remove();
            outline.InsertAt(
                IsNone(line)
                    ? new A.NoFill()
                    : (OpenXmlElement)new A.SolidFill(new A.RgbColorModelHex { Val = Hex(line) }),
                0);
        }
    }

    /// <summary>Whether the value names a vertical anchor this module understands.</summary>
    public static bool IsAnchor(string? value) =>
        value is "top" or "middle" or "bottom";

    /// <summary>
    /// Sets where text sits in a shape's box. The anchor lives on <c>a:bodyPr</c>, which is
    /// the first child of the text body, so a shape with no body properties gets one.
    /// </summary>
    public static void SetVerticalAlignment(OpenXmlElement element, string alignment)
    {
        var body = element switch
        {
            Shape s => s.TextBody,
            _ => null
        };
        if (body is null) return;

        var properties = body.GetFirstChild<A.BodyProperties>();
        if (properties is null)
        {
            properties = new A.BodyProperties();
            body.InsertAt(properties, 0);
        }

        properties.Anchor = alignment switch
        {
            "middle" => A.TextAnchoringTypeValues.Center,
            "bottom" => A.TextAnchoringTypeValues.Bottom,
            _ => A.TextAnchoringTypeValues.Top
        };
    }

    /// <summary>
    /// Gives a shape a rectangle to be filled, when it has no geometry of its own. The
    /// geometry follows <c>a:xfrm</c> and precedes the fill in the <c>p:spPr</c> sequence.
    /// </summary>
    private static void EnsureGeometry(ShapeProperties properties)
    {
        if (properties.GetFirstChild<A.PresetGeometry>() is not null) return;
        if (properties.GetFirstChild<A.CustomGeometry>() is not null) return;

        var geometry = new A.PresetGeometry(new A.AdjustValueList())
        {
            Preset = A.ShapeTypeValues.Rectangle
        };

        var transform = properties.GetFirstChild<A.Transform2D>();
        if (transform is null) properties.InsertAt(geometry, 0);
        else properties.InsertAfter(geometry, transform);
    }

    /// <summary>
    /// Places a fill after the transform and geometry, which precede it in the
    /// <c>p:spPr</c> sequence. Appending would put it after <c>a:ln</c> and invalidate the
    /// shape.
    /// </summary>
    private static void InsertAfterGeometry(ShapeProperties properties, OpenXmlElement paint)
    {
        OpenXmlElement? last = null;
        foreach (var child in properties.ChildElements)
        {
            if (child is A.Transform2D or A.PresetGeometry or A.CustomGeometry) last = child;
            else break;
        }

        if (last is null) properties.InsertAt(paint, 0);
        else properties.InsertAfter(paint, last);
    }
}
