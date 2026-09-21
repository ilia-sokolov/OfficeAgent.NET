using System.Reflection;
using System.Text.RegularExpressions;
using OfficeAgent.Abstractions;

namespace OfficeAgent.Tests;

/// <summary>
/// Every stable code lives in a public catalogue, where the API baseline freezes its value.
/// </summary>
/// <remarks>
/// Before 1.0, 81 codes existed only as string literals across eight files, so nothing froze
/// them and a typo at one emission site would have shipped a new code silently. They were
/// promoted to constants with their values unchanged. These tests keep it that way: a new
/// kebab-case literal in the source fails here, which points its author at the catalogue.
/// </remarks>
public sealed class ErrorCodeCatalogTests
{
    private static readonly Type[] Catalogues =
    {
        typeof(ValidationErrorCodes),
        typeof(ToolErrorCodes),
        typeof(TemplateDiagnosticCodes),
        typeof(ComparisonDiagnosticCodes),
        typeof(AssemblyDiagnosticCodes),
        typeof(RenderFailureCodes),
    };

    /// <summary>
    /// Kebab-case strings in the source that are identifiers or configuration values rather
    /// than codes. Adding to this list is a claim that the new string is not a code.
    /// </summary>
    private static readonly HashSet<string> NotCodes = new(StringComparer.Ordinal)
    {
        "auto-0000",            // the starting paragraph id of a blank Word document
        "core-properties",      // package part names used by assembly
        "custom-properties",
        "extended-properties",
        "non-body",             // a comparison coverage area name
        "on-behalf-of",         // configuration values for SharePoint authentication
        "app-only",
    };

    private static IEnumerable<(string Catalogue, string Name, string Value)> Codes() =>
        Catalogues.SelectMany(catalogue => catalogue
            .GetFields(BindingFlags.Public | BindingFlags.Static)
            .Where(field => field.IsLiteral && field.FieldType == typeof(string))
            .Select(field => (catalogue.Name, field.Name, (string)field.GetRawConstantValue()!)));

    [Fact]
    public void Every_code_is_lowercase_kebab_case()
    {
        // A single word is a valid code. Requiring a hyphen once kept "cancelled" out of the
        // catalogue even though every tool emitted it.
        var malformed = Codes()
            .Where(code => !Regex.IsMatch(code.Value, "^[a-z]+(-[a-z0-9]+)*$"))
            .Select(code => $"{code.Catalogue}.{code.Name} = {code.Value}");
        Assert.Empty(malformed);
    }

    [Fact]
    public void No_code_value_is_catalogued_twice()
    {
        // Two names for one code would let a host switch on either and miss the other.
        var duplicates = Codes().GroupBy(code => code.Value)
            .Where(group => group.Count() > 1)
            .Select(group => $"{group.Key}: {string.Join(", ", group.Select(c => $"{c.Catalogue}.{c.Name}"))}");
        Assert.Empty(duplicates);
    }

    [Fact]
    public void No_code_is_emitted_as_a_bare_string_literal()
    {
        var catalogued = Codes().Select(code => code.Value).ToHashSet(StringComparer.Ordinal);
        var catalogueFiles = new[] { "ChangeReport.cs", "DiagnosticCodes.cs", "OpenXmlIngestionLimits.cs" };

        var escaped = new List<string>();
        var source = Path.Combine(RepositoryRoot(), "src");
        foreach (var path in Directory.EnumerateFiles(source, "*.cs", SearchOption.AllDirectories))
        {
            if (path.Contains($"{Path.DirectorySeparatorChar}bin{Path.DirectorySeparatorChar}") ||
                path.Contains($"{Path.DirectorySeparatorChar}obj{Path.DirectorySeparatorChar}") ||
                catalogueFiles.Contains(Path.GetFileName(path)))
                continue;

            var lines = File.ReadAllLines(path);
            var insideRawString = false;
            for (var index = 0; index < lines.Length; index++)
            {
                // Raw string literals hold prose for a model, which quotes codes on purpose:
                // "fails with \"ambiguous-anchor\"" must stay readable text, not become a
                // constant name. Replacing one there once changed the prompt an agent reads.
                var wasInside = insideRawString;
                if (Regex.Matches(lines[index], "\"\"\"").Count % 2 == 1)
                    insideRawString = !insideRawString;
                if (wasInside || insideRawString) continue;

                var line = lines[index].TrimStart();
                if (line.StartsWith("//", StringComparison.Ordinal)) continue;
                foreach (Match match in Regex.Matches(lines[index], "\"([a-z]+(?:-[a-z0-9]+)+)\""))
                {
                    var value = match.Groups[1].Value;
                    if (NotCodes.Contains(value)) continue;
                    var relative = Path.GetRelativePath(RepositoryRoot(), path);
                    escaped.Add(catalogued.Contains(value)
                        ? $"{relative}:{index + 1} uses \"{value}\" instead of its catalogue constant"
                        : $"{relative}:{index + 1} emits \"{value}\", which no catalogue defines");
                }
            }
        }

        Assert.True(escaped.Count == 0, string.Join(Environment.NewLine, escaped));
    }

    private static string RepositoryRoot()
    {
        var directory = new DirectoryInfo(AppContext.BaseDirectory);
        while (directory is not null && !File.Exists(Path.Combine(directory.FullName, "OfficeAgent.NET.sln")))
            directory = directory.Parent;
        return directory?.FullName ?? throw new DirectoryNotFoundException("repository root not found");
    }
}
