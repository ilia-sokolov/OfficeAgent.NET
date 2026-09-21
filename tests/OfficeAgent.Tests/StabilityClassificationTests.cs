using System.Diagnostics.CodeAnalysis;
using System.Reflection;
using OfficeAgent.Abstractions;
using OfficeAgent.AgentFramework;
using OfficeAgent.Core;
using OfficeAgent.Excel;
using OfficeAgent.PowerPoint;
using OfficeAgent.Rendering;
using OfficeAgent.SharePoint;
using OfficeAgent.Word;

namespace OfficeAgent.Tests;

/// <summary>
/// Every public interface is deliberately placed inside or outside the 1.x stability promise.
/// </summary>
/// <remarks>
/// On netstandard2.0 an interface cannot gain a member without breaking every implementer, so
/// each public interface is a decision: freeze it, or mark it <c>[Experimental]</c> so the
/// engine can still grow. 0.9 made no such decision, and fifteen engine seams were public with
/// nothing saying whether hosts could build on them. A new public interface fails here until its
/// author places it in one of the three lists, and docs/compatibility.md says what each means.
/// </remarks>
public sealed class StabilityClassificationTests
{
    private static readonly Assembly[] Shipped =
    {
        typeof(DocumentPlan).Assembly,
        typeof(OfficeAgentTools).Assembly,
        typeof(OfficeAgentClient).Assembly,
        typeof(ExcelModule).Assembly,
        typeof(PowerPointModule).Assembly,
        typeof(LibreOfficeDocumentRenderer).Assembly,
        typeof(SharePointDocumentProvider).Assembly,
        typeof(WordModule).Assembly,
    };

    /// <summary>Hosts implement these; they are frozen for 1.x.</summary>
    private static readonly string[] HostExtensionPoints =
    {
        "OfficeAgent.Abstractions.IDocumentRenderer",
        "OfficeAgent.AgentFramework.IConnectionAccessPolicy",
        "OfficeAgent.AgentFramework.ITrustedPrincipalAccessor",
        "OfficeAgent.Core.DocumentProviders.IConnectionEditingDefaults",
        "OfficeAgent.Core.DocumentProviders.IDocumentCreatingProvider",
        "OfficeAgent.Core.DocumentProviders.IDocumentProvider",
        "OfficeAgent.Core.IAuditActorProvider",
        "OfficeAgent.SharePoint.IAccessTokenProvider",
        "OfficeAgent.SharePoint.ISharePointRegistrationStore",
    };

    /// <summary>
    /// Implemented only by the library's own plan records. Hosts read them; 1.x may add members,
    /// which breaks no host because no host implements them.
    /// </summary>
    private static readonly string[] ReadOnlyContracts =
    {
        "OfficeAgent.Abstractions.ITextFormat",
        "OfficeAgent.Abstractions.ITrackedOperation",
    };

    private static IEnumerable<Type> PublicInterfaces() =>
        Shipped.SelectMany(assembly => assembly.ExportedTypes).Where(type => type.IsInterface);

    private static bool IsExperimental(Type type) =>
        type.GetCustomAttributesData().Any(attribute =>
            attribute.AttributeType.FullName == typeof(ExperimentalAttribute).FullName &&
            (string?)attribute.ConstructorArguments[0].Value == "OFFICEAGENT001");

    [Fact]
    public void Every_public_interface_is_classified()
    {
        var unclassified = PublicInterfaces()
            .Where(type => !IsExperimental(type))
            .Select(type => type.FullName!)
            .Except(HostExtensionPoints.Concat(ReadOnlyContracts))
            .OrderBy(name => name, StringComparer.Ordinal);
        Assert.Empty(unclassified);
    }

    [Fact]
    public void No_host_extension_point_is_experimental()
    {
        // Marking one would make every host that implements it fail to compile.
        var shipped = PublicInterfaces().ToDictionary(type => type.FullName!);
        foreach (var name in HostExtensionPoints.Concat(ReadOnlyContracts))
        {
            Assert.True(shipped.ContainsKey(name), $"{name} is no longer a public interface");
            Assert.False(IsExperimental(shipped[name]), $"{name} is a stable contract but is marked experimental");
        }
    }

    [Fact]
    public void Host_setup_needs_no_experimental_type()
    {
        // The documented one-liner builds a client from a format module. The client and module
        // are stable even though IFormatModule, which the constructor names, is not: the
        // compiler reports a type only where the caller names it.
        Assert.False(IsExperimental(typeof(OfficeAgentClient)));
        Assert.False(IsExperimental(typeof(WordModule)));
        Assert.False(IsExperimental(typeof(PowerPointModule)));
        Assert.False(IsExperimental(typeof(ExcelModule)));
        Assert.True(IsExperimental(typeof(IFormatModule)));
    }
}
