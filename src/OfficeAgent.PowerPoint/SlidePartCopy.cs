using System.Globalization;
using System.Security.Cryptography;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using PC = DocumentFormat.OpenXml.Office2021.PowerPoint.Comment;
using PCMD = DocumentFormat.OpenXml.Office2016.Presentation.Command;

namespace OfficeAgent.PowerPoint;

/// <summary>
/// Relates a duplicated slide to the parts its XML points at: its own copy of every part that is
/// per-slide content, and the source's own part for everything slides legitimately share.
/// </summary>
/// <remarks>
/// Layouts, images, media and slides reached by a jump-to-slide link are shared, as PowerPoint
/// shares them. Charts, SmartArt, embedded objects and packages, controls, tags, comments and
/// notes are copied: they are edited in place, so a shared part would carry one slide's edit onto
/// the other, and PowerPoint refuses a deck in which two slides point at one chart part. Every
/// copied part keeps the relationship id it had, so the r:id attributes in the cloned XML resolve
/// unchanged.
/// </remarks>
internal static class SlidePartCopy
{
    private const string P14Namespace = "http://schemas.microsoft.com/office/powerpoint/2010/main";

    /// <summary>
    /// Relates <paramref name="copy"/>, whose XML is already a clone of <paramref name="source"/>'s,
    /// to its own or shared parts, and gives it and its comments a slide identity of its own:
    /// <paramref name="copySlideId"/> and a fresh creation id.
    /// </summary>
    public static void Relate(SlidePart source, SlidePart copy, uint copySlideId)
    {
        // The source maps to the copy, so a part that points back at its slide - a notes slide -
        // points at the copy rather than at the slide it was copied from.
        var copied = new Dictionary<OpenXmlPart, OpenXmlPart> { [source] = copy };
        CopyRelationships(source, copy, copied);

        // A slide without a creation id keeps none, and its comments keep the zero PowerPoint
        // accepts in that case.
        var creationIds = CreationIds(copy.Slide!).ToList();
        uint? creationId = null;
        if (creationIds.Count > 0)
        {
            var fresh = FreshCreationId(ValueOf(creationIds[0]));
            foreach (var element in creationIds)
                element.SetAttribute(new OpenXmlAttribute(string.Empty, "val", string.Empty, fresh.ToString(CultureInfo.InvariantCulture)));
            creationId = fresh;
        }

        foreach (var comments in copy.GetPartsOfType<PowerPointCommentPart>())
            Readdress(comments, copySlideId, creationId);
    }

    /// <summary>
    /// The slide's <c>p14:creationId</c> elements. A slide read from a file holds them as unknown
    /// elements inside <c>p:ext</c>, not as typed ones, so they are matched by name.
    /// </summary>
    private static IEnumerable<OpenXmlElement> CreationIds(OpenXmlElement slide) =>
        slide.Descendants().Where(e => e.LocalName == "creationId" && e.NamespaceUri == P14Namespace);

    private static uint ValueOf(OpenXmlElement creationId) =>
        uint.TryParse(creationId.GetAttribute("val", string.Empty).Value, NumberStyles.None, CultureInfo.InvariantCulture, out var value)
            ? value : 0U;

    /// <summary>
    /// A random creation id other than the source's, as PowerPoint gives a slide it duplicates.
    /// Reading every other slide to rule out a clash would load and so rewrite parts the plan
    /// never touched, for a one-in-billions chance.
    /// </summary>
    private static uint FreshCreationId(uint source)
    {
        var bytes = new byte[4];
        uint candidate;
        do
        {
            using (var random = RandomNumberGenerator.Create()) random.GetBytes(bytes);
            candidate = BitConverter.ToUInt32(bytes, 0);
        }
        while (candidate == 0U || candidate == source);
        return candidate;
    }

    private static void CopyRelationships(
        OpenXmlPartContainer from, OpenXmlPartContainer to, Dictionary<OpenXmlPart, OpenXmlPart> copied)
    {
        foreach (var pair in from.Parts)
        {
            var child = pair.OpenXmlPart;
            if (copied.TryGetValue(child, out var already))
            {
                to.AddPart(already, pair.RelationshipId);
                continue;
            }

            if (NewLike(to, child, pair.RelationshipId) is not { } clone)
            {
                to.AddPart(child, pair.RelationshipId);
                continue;
            }

            copied[child] = clone;
            CopyContent(child, clone);
            CopyRelationships(child, clone, copied);
        }

        foreach (var external in from.ExternalRelationships)
            to.AddExternalRelationship(external.RelationshipType, external.Uri, external.Id);
        foreach (var hyperlink in from.HyperlinkRelationships)
            to.AddHyperlinkRelationship(hyperlink.Uri, hyperlink.IsExternal, hyperlink.Id);
        // Video and audio point at media data parts, which are shared like images. Only a slide
        // carries them.
        if (to is SlidePart slide)
            foreach (var data in from.DataPartReferenceRelationships)
            {
                if (data.DataPart is not MediaDataPart media) continue;
                _ = data switch
                {
                    VideoReferenceRelationship => (DataPartReferenceRelationship)slide.AddVideoReferenceRelationship(media, data.Id),
                    AudioReferenceRelationship => slide.AddAudioReferenceRelationship(media, data.Id),
                    _ => slide.AddMediaReferenceRelationship(media, data.Id)
                };
            }
    }

    /// <summary>
    /// A new, empty part of the same kind under the same relationship id when the part is per-slide
    /// content, or null when the copy should share it.
    /// </summary>
    private static OpenXmlPart? NewLike(OpenXmlPartContainer parent, OpenXmlPart part, string id)
    {
        var type = part.ContentType;
        return part switch
        {
            // A chart and everything beneath it: the chart XML, its embedded workbook, and its
            // style, colour, theme-override and user-shape parts.
            ChartPart => parent.AddNewPart<ChartPart>(type, id),
            ExtendedChartPart => parent.AddNewPart<ExtendedChartPart>(type, id),
            ChartStylePart => parent.AddNewPart<ChartStylePart>(type, id),
            ChartColorStylePart => parent.AddNewPart<ChartColorStylePart>(type, id),
            ThemeOverridePart => parent.AddNewPart<ThemeOverridePart>(type, id),
            ChartDrawingPart => parent.AddNewPart<ChartDrawingPart>(type, id),
            EmbeddedPackagePart => parent.AddNewPart<EmbeddedPackagePart>(type, id),
            EmbeddedObjectPart => parent.AddNewPart<EmbeddedObjectPart>(type, id),

            // SmartArt: PowerPoint writes all five parts per diagram, and the data part is the
            // diagram's text.
            DiagramDataPart => parent.AddNewPart<DiagramDataPart>(type, id),
            DiagramLayoutDefinitionPart => parent.AddNewPart<DiagramLayoutDefinitionPart>(type, id),
            DiagramStylePart => parent.AddNewPart<DiagramStylePart>(type, id),
            DiagramColorsPart => parent.AddNewPart<DiagramColorsPart>(type, id),
            DiagramPersistLayoutPart => parent.AddNewPart<DiagramPersistLayoutPart>(type, id),

            // Controls, legacy drawings and tags hold the state of one slide's objects.
            EmbeddedControlPersistencePart => parent.AddNewPart<EmbeddedControlPersistencePart>(type, id),
            EmbeddedControlPersistenceBinaryDataPart => parent.AddNewPart<EmbeddedControlPersistenceBinaryDataPart>(type, id),
            VmlDrawingPart => parent.AddNewPart<VmlDrawingPart>(type, id),
            UserDefinedTagsPart => parent.AddNewPart<UserDefinedTagsPart>(type, id),

            // PowerPoint copies a slide's comments and notes with it.
            PowerPointCommentPart => parent.AddNewPart<PowerPointCommentPart>(type, id),
            SlideCommentsPart => parent.AddNewPart<SlideCommentsPart>(type, id),
            NotesSlidePart => parent.AddNewPart<NotesSlidePart>(type, id),

            _ => null
        };
    }

    /// <summary>
    /// Copies a part's content, from its loaded XML when an earlier operation in the plan has
    /// loaded and perhaps changed it, since its stream holds only what was last saved.
    /// </summary>
    private static void CopyContent(OpenXmlPart from, OpenXmlPart to)
    {
        if (from.IsRootElementLoaded && from.RootElement is { } root)
        {
            using var output = to.GetStream(FileMode.Create, FileAccess.Write);
            root.Save(output);
            return;
        }

        using var input = from.GetStream(FileMode.Open, FileAccess.Read);
        to.FeedData(input);
    }

    /// <summary>
    /// Points copied comments at the copy. PowerPoint ties a modern comment to its slide through
    /// the slide and creation ids in its moniker, and the comment ids are given anew so that no
    /// two comments in the deck share one.
    /// </summary>
    private static void Readdress(PowerPointCommentPart part, uint slideId, uint? creationId)
    {
        if (part.CommentList is not { } list) return;

        foreach (var moniker in list.Descendants<PCMD.SlideMoniker>())
        {
            moniker.SldId = slideId;
            if (creationId is { } value) moniker.CId = value;
        }

        foreach (var comment in list.Descendants<PC.Comment>())
            comment.Id = NewCommentId();
        foreach (var reply in list.Descendants<PC.CommentReply>())
            reply.Id = NewCommentId();
    }

    private static string NewCommentId() => "{" + Guid.NewGuid().ToString().ToUpperInvariant() + "}";
}
