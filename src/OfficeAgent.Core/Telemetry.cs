using System.Diagnostics;
using System.Reflection;

namespace OfficeAgent.Core;

/// <summary>
/// Telemetry primitives shared by the engine. The activity source name is
/// <c>OfficeAgent</c> so OpenTelemetry collectors can subscribe with one
/// <c>AddSource("OfficeAgent")</c> call.
/// </summary>
public static class OfficeAgentTelemetry
{
    /// <summary>The <see cref="ActivitySource"/> name used by every engine span.</summary>
    public const string ActivitySourceName = "OfficeAgent";

    /// <summary>The shared <see cref="ActivitySource"/>.</summary>
    public static readonly ActivitySource ActivitySource =
        new(ActivitySourceName, typeof(OfficeAgentTelemetry).Assembly
            .GetCustomAttribute<AssemblyInformationalVersionAttribute>()?.InformationalVersion ?? "0.0.0");

    /// <summary>The category name used by every engine logger.</summary>
    public const string LogCategory = "OfficeAgent";
}
