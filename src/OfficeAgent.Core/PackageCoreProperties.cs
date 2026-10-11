using DocumentFormat.OpenXml.Packaging;
using OfficeAgent.Abstractions;

namespace OfficeAgent.Core;

/// <summary>
/// The OPC core properties every Office package carries in <c>docProps/core.xml</c> -
/// creator, title, subject and the rest - exposed as <c>docProperty</c> nodes with
/// <c>core/*</c> paths. Shared by the format modules, so a deck and a workbook answer the
/// same paths a Word document does.
/// </summary>
internal static class PackageCoreProperties
{
    /// <summary>The node kind core properties are listed under.</summary>
    public const string Kind = "docProperty";

    private const string PathPrefix = "core/";

    /// <summary>The core property names, in the order inspection lists them.</summary>
    public static readonly IReadOnlyList<string> Names = new[]
    {
        "creator", "title", "subject", "keywords", "lastModifiedBy", "revision", "category", "contentStatus"
    };

    /// <summary>Whether the name is a core property.</summary>
    public static bool IsCore(string name) => Names.Contains(name, StringComparer.Ordinal);

    /// <summary>The property name in a <c>core/*</c> path, or the path itself when it has no prefix.</summary>
    public static string NameOf(string path) =>
        path.StartsWith(PathPrefix, StringComparison.Ordinal) ? path.Substring(PathPrefix.Length) : path;

    /// <summary>Reads a core property, or returns <see langword="null"/> when it is unset or unknown.</summary>
    public static string? Read(OpenXmlPackage package, string name)
    {
        var p = package.PackageProperties;
        return name switch
        {
            "creator" => p.Creator,
            "title" => p.Title,
            "subject" => p.Subject,
            "keywords" => p.Keywords,
            "lastModifiedBy" => p.LastModifiedBy,
            "revision" => p.Revision,
            "category" => p.Category,
            "contentStatus" => p.ContentStatus,
            _ => null
        };
    }

    /// <summary>Writes a core property. Unknown names are ignored; preview refuses them first.</summary>
    public static void Write(OpenXmlPackage package, string name, string? value)
    {
        var p = package.PackageProperties;
        switch (name)
        {
            case "creator": p.Creator = value; break;
            case "title": p.Title = value; break;
            case "subject": p.Subject = value; break;
            case "keywords": p.Keywords = value; break;
            case "lastModifiedBy": p.LastModifiedBy = value; break;
            case "revision": p.Revision = value; break;
            case "category": p.Category = value; break;
            case "contentStatus": p.ContentStatus = value; break;
        }
    }

    /// <summary>One <c>docProperty</c> node per core property, with its current value.</summary>
    public static IEnumerable<NodeInfo> Nodes(OpenXmlPackage package)
    {
        foreach (var name in Names)
        {
            var value = Read(package, name);
            yield return new NodeInfo
            {
                Kind = Kind,
                Path = PathPrefix + name,
                Summary = value ?? string.Empty,
                Anchor = new NodeAnchor { Id = $"docprop:{name}", Kind = Kind, Path = PathPrefix + name, Expect = value }
            };
        }
    }
}

/// <summary>
/// <c>setProperty</c> on a <c>core/*</c> <c>docProperty</c> node, for formats whose only
/// document properties are the package's core properties. Word has its own handler, which
/// also covers its settings and fields.
/// </summary>
internal sealed class CorePropertyHandler : IOperationHandler
{
    public bool CanHandle(PlanOperation operation) =>
        operation is SetPropertyOp { Target: NodeAnchor { Kind: PackageCoreProperties.Kind } };

    public OperationPreview Preview(ApplyContext context, PlanOperation operation)
    {
        var op = (SetPropertyOp)operation;
        var anchor = (NodeAnchor)op.Target;
        var name = PackageCoreProperties.NameOf(anchor.Path);
        if (!PackageCoreProperties.IsCore(name))
            return OperationPreview.Fail(new ValidationError(
                ValidationErrorCodes.AnchorNotFound,
                $"Unknown document property '{anchor.Path}'. Core properties are " +
                string.Join(", ", PackageCoreProperties.Names.Select(n => "core/" + n)) + ".",
                anchor));

        return OperationPreview.Ok(new ProposedChange
        {
            Target = anchor,
            Verb = "setProperty",
            Before = PackageCoreProperties.Read(context.Package.Package, name) ?? string.Empty,
            After = op.Value ?? string.Empty,
            Context = anchor.Path,
            BlastRadius = 1,
            Capability = Capability.Deterministic
        });
    }

    public void Apply(ApplyContext context, PlanOperation operation)
    {
        var op = (SetPropertyOp)operation;
        var anchor = (NodeAnchor)op.Target;
        PackageCoreProperties.Write(context.Package.Package, PackageCoreProperties.NameOf(anchor.Path), op.Value);
    }
}
