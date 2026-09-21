#if NETSTANDARD2_0
namespace System.Diagnostics.CodeAnalysis;

/// <summary>
/// The compiler recognises this attribute by name, so the netstandard2.0 build reports the same
/// diagnostic as the net8.0 build, where the type ships in the runtime.
/// </summary>
[AttributeUsage(AttributeTargets.Assembly | AttributeTargets.Module | AttributeTargets.Class |
                AttributeTargets.Struct | AttributeTargets.Enum | AttributeTargets.Constructor |
                AttributeTargets.Method | AttributeTargets.Property | AttributeTargets.Field |
                AttributeTargets.Event | AttributeTargets.Interface | AttributeTargets.Delegate,
                Inherited = false)]
internal sealed class ExperimentalAttribute : Attribute
{
    public ExperimentalAttribute(string diagnosticId) => DiagnosticId = diagnosticId;

    public string DiagnosticId { get; }

    public string? UrlFormat { get; set; }
}
#endif
