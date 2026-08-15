using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace OfficeAgent.Word;

/// <summary>
/// Puts an image behind every page of the document.
/// </summary>
/// <remarks>
/// <para>
/// Word has a <c>w:background</c> element, but it holds a colour and - for a picture - a
/// lump of VML that Word writes and few other producers read, and it does not print by
/// default. What Word's own designed templates do instead, and what this does, is anchor a
/// page-sized picture in the header with <c>behindDoc</c> set: it repeats on every page,
/// prints, survives a round trip through Word, and cannot be selected while editing the
/// body.
/// </para>
/// <para>
/// The picture goes into every header the section uses, because a section with a distinct
/// first-page header would otherwise show the background on page two onwards only - which
/// looks like a bug on the very page most likely to be a cover.
/// </para>
/// </remarks>
internal sealed class WordBackgroundImageHandler : IOperationHandler
{
    /// <summary>The name every background drawing carries, so it can be found again to replace it.</summary>
    private const string DrawingName = "OfficeAgent Background";

    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) =>
        operation is BackgroundImageOp { Target: null };

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (BackgroundImageOp)operation;

        if (op.Opacity is { } opacity && (opacity < 0 || opacity > 1))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"opacity must be between 0 and 1; got {opacity}.", null));

        if (WordImages.Unsupported(op) is { } unsupported)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation, unsupported, null));

        if (op.Scope is not null && ScopesOf(op.Scope) is null)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"'{op.Scope}' is not a background scope. Expected all, firstPage, default, or evenPage.",
                null));

        var clearing = string.IsNullOrEmpty(op.Base64Bytes);
        return OperationPreview.Ok(new ProposedChange
        {
            Target = new NodeAnchor { Kind = "page", Path = "page#background" },
            Verb = "backgroundImage",
            Before = string.Empty,
            After = clearing
                ? "[page background cleared]"
                : $"[page background image{(op.Opacity is { } o && o < 1 ? $" at {o:P0}" : string.Empty)}]",
            Context = "every page",
            BlastRadius = 1
        });
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (BackgroundImageOp)operation;
        var main = WordModel.Doc(context.Package).MainDocumentPart!;
        var section = WordSections.Require(main);

        var scopes = ScopesOf(op.Scope) ?? throw new InvalidOperationException($"Bad scope '{op.Scope}'.");

        // Whatever was there before goes first, so setting a background twice replaces it
        // rather than stacking two images nobody can tell apart. Only the scopes being
        // written are cleared, or setting a cover background would strip the body's.
        RemoveExisting(main, section, scopes);
        if (string.IsNullOrEmpty(op.Base64Bytes)) return;

        var bytes = Convert.FromBase64String(op.Base64Bytes!);
        var partType = WordImages.PartTypeFor(op.ImageType)
            ?? throw new InvalidOperationException($"Unsupported imageType '{op.ImageType}'.");

        var (width, height) = WordSections.PageSizeEmu(section);

        var id = NextDrawingId(context.Package);
        foreach (var kind in Resolve(main, section, scopes))
        {
            var header = WordSections.HeaderFor(main, section, kind);
            var part = (HeaderPart)header.OpenXmlPart!;

            var imagePart = part.AddImagePart(partType);
            using (var stream = new MemoryStream(bytes))
                imagePart.FeedData(stream);

            var drawing = BuildPageAnchor(
                part.GetIdOfPart(imagePart), width, height, op.Opacity, id++);

            // The background opens the header, so header text sits over it.
            header.InsertAt(new Paragraph(new Run(drawing)), 0);
        }
    }

    /// <summary>
    /// Builds the anchored drawing: pinned to the top-left of the page, page-sized, behind
    /// the text, and out of the text flow entirely.
    /// </summary>
    private static Drawing BuildPageAnchor(
        string relationshipId, long width, long height, double? opacity, uint id)
    {
        var blip = new A.Blip { Embed = relationshipId };
        if (WordImages.Alpha(opacity) is { } amount)
            blip.Append(new A.AlphaModulationFixed { Amount = amount });

        return new Drawing(
            new DW.Anchor(
                new DW.SimplePosition { X = 0L, Y = 0L },
                new DW.HorizontalPosition(new DW.PositionOffset("0"))
                {
                    RelativeFrom = DW.HorizontalRelativePositionValues.Page
                },
                new DW.VerticalPosition(new DW.PositionOffset("0"))
                {
                    RelativeFrom = DW.VerticalRelativePositionValues.Page
                },
                new DW.Extent { Cx = width, Cy = height },
                new DW.EffectExtent { LeftEdge = 0L, TopEdge = 0L, RightEdge = 0L, BottomEdge = 0L },
                // Out of the flow: text runs over the picture rather than around it.
                new DW.WrapNone(),
                new DW.DocProperties { Id = id, Name = DrawingName },
                new DW.NonVisualGraphicFrameDrawingProperties(
                    new A.GraphicFrameLocks { NoChangeAspect = false }),
                new A.Graphic(
                    new A.GraphicData(
                        new PIC.Picture(
                            new PIC.NonVisualPictureProperties(
                                new PIC.NonVisualDrawingProperties { Id = 0U, Name = DrawingName },
                                new PIC.NonVisualPictureDrawingProperties()),
                            new PIC.BlipFill(blip, new A.Stretch(new A.FillRectangle())),
                            new PIC.ShapeProperties(
                                new A.Transform2D(
                                    new A.Offset { X = 0L, Y = 0L },
                                    new A.Extents { Cx = width, Cy = height }),
                                new A.PresetGeometry(new A.AdjustValueList())
                                {
                                    Preset = A.ShapeTypeValues.Rectangle
                                })))
                    {
                        Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture"
                    }))
            {
                DistanceFromTop = 0U,
                DistanceFromBottom = 0U,
                DistanceFromLeft = 0U,
                DistanceFromRight = 0U,
                SimplePos = false,
                RelativeHeight = 0U,
                BehindDoc = true,
                Locked = false,
                LayoutInCell = true,
                AllowOverlap = true
            });
    }

    /// <summary>
    /// The scope names a caller may use, mapped to the header kinds each covers.
    /// <c>all</c> is null-safe: it is what an unset scope means.
    /// </summary>
    private static HeaderFooterValues[]? ScopesOf(string? scope) =>
        scope?.Trim().ToLowerInvariant() switch
        {
            null or "" or "all" => new[]
            {
                HeaderFooterValues.Default, HeaderFooterValues.First, HeaderFooterValues.Even
            },
            "firstpage" or "first" => new[] { HeaderFooterValues.First },
            "default" => new[] { HeaderFooterValues.Default },
            "evenpage" or "even" => new[] { HeaderFooterValues.Even },
            _ => null
        };

    /// <summary>
    /// Narrows the requested scopes to the ones this document actually shows. Writing a
    /// first-page background into a section with no distinct first page creates a header
    /// Word never displays, which looks to the caller like the operation did nothing.
    /// </summary>
    private static IEnumerable<HeaderFooterValues> Resolve(
        MainDocumentPart main, SectionProperties section, HeaderFooterValues[] scopes)
    {
        foreach (var kind in scopes)
        {
            if (kind == HeaderFooterValues.First && !WordSections.HasTitlePage(section)) continue;
            if (kind == HeaderFooterValues.Even && !UsesEvenPages(main)) continue;
            yield return kind;
        }
    }

    /// <summary>
    /// Takes out the background this handler put in, and the now-orphaned image with it.
    /// Only drawings it named are touched, so a picture the author placed in a header by
    /// hand survives - and only in the scopes being rewritten.
    /// </summary>
    private static void RemoveExisting(
        MainDocumentPart main, SectionProperties section, HeaderFooterValues[] scopes)
    {
        var wanted = new HashSet<string>(
            section.Elements<HeaderReference>()
                .Where(r => r.Type is not null && scopes.Contains(r.Type.Value))
                .Select(r => r.Id?.Value ?? string.Empty),
            StringComparer.Ordinal);

        foreach (var part in main.HeaderParts.ToList())
        {
            if (part.Header is null) continue;
            if (!wanted.Contains(main.GetIdOfPart(part))) continue;

            foreach (var drawing in part.Header.Descendants<Drawing>().ToList())
            {
                var name = drawing.Descendants<DW.DocProperties>().FirstOrDefault()?.Name?.Value;
                if (name != DrawingName) continue;

                foreach (var blip in drawing.Descendants<A.Blip>())
                    if (blip.Embed?.Value is { Length: > 0 } id &&
                        part.GetPartById(id) is ImagePart image)
                        part.DeletePart(image);

                // The run and paragraph that held it go too, or the header keeps an empty
                // line that pushes the real header text down the page.
                var paragraph = drawing.Ancestors<Paragraph>().FirstOrDefault();
                drawing.Ancestors<Run>().FirstOrDefault()?.Remove();
                if (paragraph is not null && paragraph.InnerText.Length == 0 &&
                    !paragraph.Descendants<Drawing>().Any())
                    paragraph.Remove();
            }
        }
    }

    /// <summary>Whether the document sets different odd and even headers.</summary>
    private static bool UsesEvenPages(MainDocumentPart main) =>
        main.DocumentSettingsPart?.Settings?.GetFirstChild<EvenAndOddHeaders>() is { } setting &&
        (setting.Val is null || setting.Val.Value);

    private static uint NextDrawingId(IOpenXmlPackage package)
    {
        uint max = 0;
        foreach (var (drawing, _) in ImageNodeProvider.EnumerateDrawingsWithHost(package))
        {
            var id = drawing.Descendants<DW.DocProperties>().FirstOrDefault()?.Id?.Value ?? 0U;
            if (id > max) max = id;
        }
        return max + 1U;
    }
}
