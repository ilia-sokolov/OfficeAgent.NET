using OfficeAgent.Abstractions;
using OfficeAgent.Core;
using P14 = DocumentFormat.OpenXml.Office2010.PowerPoint;

namespace OfficeAgent.PowerPoint;

/// <summary>Surfaces each section as an addressable node.</summary>
internal sealed class SectionNodeProvider : IPowerPointNodeProvider
{
    /// <inheritdoc />
    public string Kind => "section";

    /// <inheritdoc />
    public IEnumerable<NodeInfo> Enumerate(PowerPointObjectMap map)
    {
        foreach (var section in Sections.All(map.Package))
        {
            var path = $"section#{section.Id}";
            var slides = section.SlideIds.Count;
            yield return new NodeInfo
            {
                Kind = Kind,
                Path = path,
                Summary = $"section {section.Number}: \"{section.Name}\" — {slides} slide(s)",
                Anchor = new NodeAnchor { Id = path, Kind = Kind, Path = path }
            };
        }
    }

    /// <inheritdoc />
    public ResolvedNode? Resolve(NodeAnchor anchor, PowerPointObjectMap map)
    {
        if (!Sections.TryParseId(anchor.Path, out var id)) return null;

        var section = Sections.Find(map.Package, id);
        return section is null
            ? null
            : new ResolvedNode
            {
                Kind = Kind,
                Elements = new[] { (DocumentFormat.OpenXml.OpenXmlElement)section.Section },
                Value = section.Name
            };
    }
}

/// <summary>
/// Adds, renames, and removes the named slide groups PowerPoint shows in the thumbnail
/// pane.
/// </summary>
internal sealed class SectionHandler : IOperationHandler
{
    /// <inheritdoc />
    public bool CanHandle(PlanOperation operation) => operation is SectionOp;

    /// <inheritdoc />
    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (SectionOp)operation;

        return op.Action switch
        {
            SectionAction.Add => PreviewAdd(op, context),
            SectionAction.Rename => PreviewRename(op, context),
            _ => PreviewRemove(op, context)
        };
    }

    /// <inheritdoc />
    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (SectionOp)operation;

        switch (op.Action)
        {
            case SectionAction.Add:
                ApplyAdd(op, context);
                break;
            case SectionAction.Rename:
                Resolve(op, context)!.Section.Name = op.Name;
                break;
            default:
                ApplyRemove(op, context);
                break;
        }

        Sections.Reconcile(context.Package);
    }

    // ── add ───────────────────────────────────────────────────────────────────

    private static OperationPreview PreviewAdd(SectionOp op, ApplyContext context)
    {
        if (string.IsNullOrWhiteSpace(op.Name))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "A section needs a name.", op.Target));

        if (op.Target is not NodeAnchor { Kind: "slide" } anchor)
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "Adding a section targets the slide it starts at: " +
                "{ \"kind\": \"slide\", \"path\": \"slide#256\" }.", op.Target));

        var slide = SlideList.Target(context, anchor);
        if (slide is null) return OperationPreview.Fail(SlideList.NoSuchSlide(anchor));

        // Two sections cannot start at one slide: the second would own no slides at all,
        // and PowerPoint shows an empty section as a grouping the user cannot fill.
        if (Sections.All(context.Package).Any(s => s.SlideIds.FirstOrDefault() == slide.SlideId))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                $"A section already starts at slide {slide.Number}. Rename that one, or start " +
                "this section at a different slide.", anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "section",
            Before = string.Empty,
            After = $"\"{op.Name}\"",
            Context = $"starting at slide {slide.Number}",
            BlastRadius = 1
        });
    }

    private static void ApplyAdd(SectionOp op, ApplyContext context)
    {
        var anchor = (NodeAnchor)op.Target;
        var slide = SlideList.Target(context, anchor)
            ?? throw new InvalidOperationException($"Slide '{anchor.Path}' vanished before apply.");

        var list = Sections.Require(context.Package);
        var section = new P14.Section(new P14.SectionSlideIdList())
        {
            Name = op.Name,
            Id = Sections.NewId()
        };

        // A deck that had no sections needs one covering the slides before this one, or
        // they would belong to nothing and PowerPoint would report the file as damaged.
        var existing = list.Elements<P14.Section>().ToList();
        if (existing.Count == 0)
        {
            var precede = PowerPointModel.Slides(context.Package)
                .TakeWhile(s => s.SlideId != slide.SlideId)
                .ToList();

            if (precede.Count > 0)
            {
                var head = new P14.Section(new P14.SectionSlideIdList())
                {
                    Name = "Default Section",
                    Id = Sections.NewId()
                };
                list.Append(head);
            }

            list.Append(section);
        }
        else
        {
            // Place it before the first section that starts after this slide, so section
            // order follows slide order the way the thumbnail pane shows it.
            var order = PowerPointModel.Slides(context.Package).Select(s => s.SlideId).ToList();
            var position = order.IndexOf(slide.SlideId);

            var following = existing.FirstOrDefault(s =>
            {
                var first = s.SectionSlideIdList?.Elements<P14.SectionSlideIdListEntry>()
                    .FirstOrDefault()?.Id?.Value;
                return first is { } id && order.IndexOf(id) >= position;
            });

            if (following is null) list.Append(section);
            else list.InsertBefore(section, following);
        }

        // Claim this slide, so Reconcile gives it and its followers to the new section.
        section.SectionSlideIdList!.Append(new P14.SectionSlideIdListEntry { Id = slide.SlideId });
        foreach (var other in list.Elements<P14.Section>())
        {
            if (ReferenceEquals(other, section)) continue;
            other.SectionSlideIdList?.Elements<P14.SectionSlideIdListEntry>()
                .FirstOrDefault(e => e.Id?.Value == slide.SlideId)?.Remove();
        }
    }

    // ── rename / remove ───────────────────────────────────────────────────────

    private static OperationPreview PreviewRename(SectionOp op, ApplyContext context)
    {
        if (string.IsNullOrWhiteSpace(op.Name))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.InvalidOperation,
                "A section needs a name.", op.Target));

        var section = Resolve(op, context);
        if (section is null) return NoSuchSection(op);

        return OperationPreview.Ok(new ProposedChange
        {
            Target = op.Target,
            Verb = "section",
            Before = $"\"{section.Name}\"",
            After = $"\"{op.Name}\"",
            Context = $"section {section.Number}",
            BlastRadius = 1
        });
    }

    private static OperationPreview PreviewRemove(SectionOp op, ApplyContext context)
    {
        var section = Resolve(op, context);
        if (section is null) return NoSuchSection(op);

        return OperationPreview.Ok(new ProposedChange
        {
            Target = op.Target,
            Verb = "section",
            Before = $"\"{section.Name}\" ({section.SlideIds.Count} slide(s))",
            After = string.Empty,
            Context = $"section {section.Number}; its slides are kept",
            BlastRadius = 1
        });
    }

    private static void ApplyRemove(SectionOp op, ApplyContext context)
    {
        var section = Resolve(op, context)
            ?? throw new InvalidOperationException("Section vanished before apply.");

        var list = Sections.List(context.Package)!;
        section.Section.Remove();

        // Removing the last section leaves the deck unsectioned, which means removing the
        // now-empty container too: an empty p14:sectionLst is not valid.
        if (!list.Elements<P14.Section>().Any())
            list.Parent?.Remove();
    }

    private static SectionRef? Resolve(SectionOp op, ApplyContext context) =>
        op.Target is NodeAnchor { Kind: "section" } anchor && Sections.TryParseId(anchor.Path, out var id)
            ? Sections.Find(context.Package, id)
            : null;

    private static OperationPreview NoSuchSection(SectionOp op) =>
        OperationPreview.Fail(new ValidationError(
            ValidationErrorCodes.AnchorNotFound,
            $"No section with path '{(op.Target as NodeAnchor)?.Path}'. " +
            "Section paths come from inspect_document.nodes.",
            op.Target));
}
