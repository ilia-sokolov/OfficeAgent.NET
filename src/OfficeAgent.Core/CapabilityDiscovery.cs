using System.Diagnostics.CodeAnalysis;
using OfficeAgent.Abstractions;
using DocFormat = OfficeAgent.Abstractions.DocumentFormat;

namespace OfficeAgent.Core;

/// <summary>
/// Optional <see cref="IFormatModule"/> capability: declare what the module's own
/// validation enforces where probing a handler cannot reveal it.
/// </summary>
/// <remarks>
/// A handler's <c>CanHandle</c> answers which verbs exist, but not which change modes a
/// module will honour: that rule lives in plan-wide validation and needs an open package
/// to run. A module states it here, and a reconciliation test proves the statement
/// matches what validation actually does, so the declaration cannot drift into a claim.
/// </remarks>
[Experimental(EngineExtensibility.DiagnosticId, UrlFormat = EngineExtensibility.UrlFormat)]
public interface ICapabilityDeclaringModule
{
    /// <summary>Gets the change modes this format can honour.</summary>
    IReadOnlyList<ChangeMode> SupportedChangeModes { get; }

    /// <summary>Gets the node kinds this module's inspection surfaces.</summary>
    IReadOnlyList<string> NodeKinds { get; }
}

/// <summary>
/// Optional <see cref="IDocumentService"/> capability: describe the engine as configured.
/// Kept separate from <see cref="IDocumentService"/> so an existing implementation of that
/// interface keeps compiling.
/// </summary>
[Experimental(EngineExtensibility.DiagnosticId, UrlFormat = EngineExtensibility.UrlFormat)]
public interface ICapabilityReportingService
{
    /// <summary>Describes the engine's registered formats, versions and ceilings.</summary>
    EngineCapabilities Describe();
}

/// <summary>
/// Builds <see cref="EngineCapabilities"/> from what is actually registered.
/// </summary>
/// <remarks>
/// Nothing here is a maintained list. Formats come from the registered modules, verbs
/// come from asking those modules' handlers, versions come from the contract constants,
/// and ceilings come from the host's configured limits. A capability that cannot be
/// derived is declared by the module that owns the rule and reconciled by a test.
/// </remarks>
[Experimental(EngineExtensibility.DiagnosticId, UrlFormat = EngineExtensibility.UrlFormat)]
public static class CapabilityDiscovery
{
    /// <summary>
    /// Describes the engine as configured. Connections are not included: they depend on
    /// the caller, and <see cref="OfficeAgentClient"/> adds only the ones that caller may
    /// actually use.
    /// </summary>
    public static EngineCapabilities Describe(
        IEnumerable<IFormatModule> modules,
        OpenXmlIngestionLimits? limits = null,
        bool renderingAvailable = false)
    {
        if (modules is null) throw new ArgumentNullException(nameof(modules));

        return new EngineCapabilities
        {
            Contracts = new ContractVersions(),
            Formats = modules
                .Select(Describe)
                .OrderBy(format => format.Format.ToString(), StringComparer.Ordinal)
                .ToList(),
            Limits = limits ?? OpenXmlIngestionLimits.Default,
            RenderingAvailable = renderingAvailable,
            RequiresInspection = new[]
            {
                "Anchor identifiers. Paragraph, table, slide and shape ids come from inspect_document " +
                "for the document in hand; they are not predictable and are not listed here.",
                "Content-dependent refusals. Whether a specific edit is legal can depend on the " +
                "document, for example an edit that spans a pending tracked revision.",
                "Template bindings. Which tags and repeating rows a template exposes is a property " +
                "of that template, discoverable by inspecting it."
            }
        };
    }

    private static FormatCapabilities Describe(IFormatModule module)
    {
        var declared = module as ICapabilityDeclaringModule;

        return new FormatCapabilities
        {
            Format = module.Format,
            Operations = SupportedVerbs(module, declared?.NodeKinds ?? Array.Empty<string>()),
            ChangeModes = declared?.SupportedChangeModes
                ?? new[] { ChangeMode.Direct, ChangeMode.Tracked },
            NodeKinds = (declared?.NodeKinds ?? Array.Empty<string>())
                .OrderBy(kind => kind, StringComparer.Ordinal)
                .ToList()
        };
    }

    /// <summary>
    /// Asks the module's handlers which verbs they accept, by offering each verb with
    /// every anchor shape the module could address.
    /// </summary>
    /// <remarks>
    /// <c>CanHandle</c> is the predicate validation itself uses, so a verb reported here
    /// is one the module really takes. A verb no handler accepts under any anchor is
    /// absent rather than advertised and then refused at apply time.
    /// </remarks>
    private static IReadOnlyList<string> SupportedVerbs(IFormatModule module, IReadOnlyList<string> nodeKinds)
    {
        var anchors = CandidateAnchors(nodeKinds);
        var supported = new List<string>();

        foreach (var entry in PlanOperationJsonConverter.ByVerb)
        {
            var verb = entry.Key;
            var type = entry.Value;
            if (Activator.CreateInstance(type) is not PlanOperation prototype) continue;

            foreach (var anchor in anchors)
            {
                var candidate = WithTarget(prototype, type, anchor);
                if (candidate is null) continue;
                if (module.Handlers.Any(handler => handler.CanHandle(candidate)))
                {
                    supported.Add(verb);
                    break;
                }
            }
        }

        supported.Sort(StringComparer.Ordinal);
        return supported;
    }

    private static IReadOnlyList<Anchor?> CandidateAnchors(IReadOnlyList<string> nodeKinds)
    {
        var anchors = new List<Anchor?>
        {
            null,
            new TextSpanAnchor(),
            new StructuralAnchor(),
            new StyleAnchor(),
            new CellAnchor(),
            new ShapeAnchor()
        };

        // A NodeAnchor's Kind is part of the predicate, so every kind the module surfaces
        // is offered. Kinds come from the module's own providers, not from a list here.
        foreach (var kind in nodeKinds)
            anchors.Add(new NodeAnchor { Kind = kind });

        return anchors;
    }

    /// <summary>
    /// Returns a copy of the prototype carrying the candidate anchor, or null when the
    /// operation type has no settable target.
    /// </summary>
    private static PlanOperation? WithTarget(PlanOperation prototype, Type type, Anchor? anchor)
    {
        if (anchor is null) return prototype;

        var property = type.GetProperty(nameof(PlanOperation.Target));
        if (property is null) return null;

        // Operations are init-only records of a sort; a fresh instance per candidate keeps
        // probing free of shared state.
        if (Activator.CreateInstance(type) is not PlanOperation candidate) return null;
        if (!property.PropertyType.IsInstanceOfType(anchor)) return null;

        try
        {
            property.SetValue(candidate, anchor);
        }
        catch (Exception ex) when (ex is ArgumentException or MethodAccessException or InvalidOperationException)
        {
            return null;
        }

        return candidate;
    }
}
