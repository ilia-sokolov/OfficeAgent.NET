using System.Reflection;
using System.Runtime.Loader;
using System.Text.Json;

/// <summary>
/// Checks that the netstandard2.0 build of each library exposes exactly the public surface the
/// net8.0 baseline records.
/// </summary>
/// <remarks>
/// The baseline is rendered from the net8.0 assemblies, so on its own it says nothing about a
/// .NET Framework or Unity consumer, who compiles against the netstandard2.0 build. The two can
/// differ: a member inside <c>#if NET8_0_OR_GREATER</c>, or a polyfill made public by mistake,
/// changes one target's surface and not the other's. This loads the netstandard2.0 binaries into
/// their own load context and renders them with the same renderer, so any difference is a line
/// the check reports.
/// </remarks>
internal static class TargetParity
{
    private const string NetStandard = "netstandard2.0";

    public static (int Types, IReadOnlyList<string> Differences) Compare()
    {
        var root = RepositoryRoot();
        var configuration = new DirectoryInfo(AppContext.BaseDirectory).Parent!.Name;
        var differences = new List<string>();
        var context = new NetStandardContext(root, configuration);
        var compared = 0;

        foreach (var assembly in ApiSurface.Shipped)
        {
            var name = assembly.GetName().Name!;
            var project = Path.Combine(root, "src", name, name + ".csproj");
            if (!File.ReadAllText(project).Contains(NetStandard, StringComparison.Ordinal))
                continue;

            var path = context.PathOf(name)
                ?? throw new FileNotFoundException(
                    $"No {NetStandard} build of {name} under {configuration}. Build the solution, not just this tool.");
            var netStandard = context.LoadFromAssemblyPath(path);

            var expected = Surface(assembly);
            var actual = Surface(netStandard);
            foreach (var type in expected.Keys.Union(actual.Keys).OrderBy(type => type, StringComparer.Ordinal))
            {
                compared++;
                if (!actual.TryGetValue(type, out var nsLines))
                {
                    differences.Add($"{name}: {type} is public in net8.0 only");
                    continue;
                }
                if (!expected.TryGetValue(type, out var netLines))
                {
                    differences.Add($"{name}: {type} is public in {NetStandard} only");
                    continue;
                }
                differences.AddRange(netLines.Except(nsLines).Select(line => $"{name}: {type}: net8.0 only: {line}"));
                differences.AddRange(nsLines.Except(netLines).Select(line => $"{name}: {type}: {NetStandard} only: {line}"));
            }
        }
        return (compared, differences);
    }

    private static Dictionary<string, IReadOnlyList<string>> Surface(Assembly assembly) =>
        assembly.ExportedTypes
            .Where(type => type.Namespace?.StartsWith("OfficeAgent", StringComparison.Ordinal) == true)
            .ToDictionary(type => type.FullName!, ApiSurface.RenderType, StringComparer.Ordinal);

    private static string RepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "OfficeAgent.NET.sln")))
            directory = directory.Parent;
        return directory?.FullName ?? throw new DirectoryNotFoundException("repository root not found");
    }

    /// <summary>
    /// Loads OfficeAgent assemblies from their netstandard2.0 output, and their package
    /// dependencies from the paths the netstandard2.0 builds were resolved against.
    /// </summary>
    private sealed class NetStandardContext : AssemblyLoadContext
    {
        private readonly Dictionary<string, string> _paths = new(StringComparer.OrdinalIgnoreCase);

        public NetStandardContext(string root, string configuration) : base(NetStandard, isCollectible: true)
        {
            var packages = Environment.GetEnvironmentVariable("NUGET_PACKAGES") ??
                           Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.UserProfile), ".nuget", "packages");
            foreach (var project in Directory.GetDirectories(Path.Combine(root, "src")))
            {
                var output = Path.Combine(project, "bin", configuration, NetStandard);
                if (!Directory.Exists(output)) continue;
                // Only the project's own assembly. Its output folder also holds copies of the
                // OfficeAgent assemblies it references, which are stale whenever only a
                // dependency was rebuilt; taking those once hid a real divergence.
                var own = Path.Combine(output, Path.GetFileName(project) + ".dll");
                if (File.Exists(own))
                    _paths[Path.GetFileName(project)] = own;
                foreach (var deps in Directory.GetFiles(output, "*.deps.json"))
                    AddPackageAssemblies(deps, packages);
            }
            Resolving += (_, name) => _paths.TryGetValue(name.Name!, out var path) ? LoadFromAssemblyPath(path) : null;
        }

        public string? PathOf(string name) => _paths.TryGetValue(name, out var path) ? path : null;

        // OfficeAgent assemblies must come from here even though the tool's own context already
        // holds their net8.0 builds under the same names. Everything else falls through to the
        // runtime, and then to the Resolving handler for packages only netstandard2.0 needs.
        protected override Assembly? Load(AssemblyName name) =>
            name.Name!.StartsWith("OfficeAgent.", StringComparison.Ordinal) && _paths.TryGetValue(name.Name, out var path)
                ? LoadFromAssemblyPath(path)
                : null;

        private void AddPackageAssemblies(string depsPath, string packages)
        {
            using var deps = JsonDocument.Parse(File.ReadAllText(depsPath));
            if (!deps.RootElement.TryGetProperty("targets", out var targets)) return;
            foreach (var target in targets.EnumerateObject())
            foreach (var library in target.Value.EnumerateObject())
            {
                if (!library.Value.TryGetProperty("runtime", out var runtime)) continue;
                foreach (var file in runtime.EnumerateObject())
                {
                    var path = Path.Combine(packages, library.Name.ToLowerInvariant(), file.Name);
                    if (!library.Name.StartsWith("OfficeAgent.", StringComparison.Ordinal) && File.Exists(path))
                        _paths.TryAdd(Path.GetFileNameWithoutExtension(path), path);
                }
            }
        }
    }
}
