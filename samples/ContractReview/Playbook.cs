using System.Text.RegularExpressions;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeAgent.Samples.ContractReview;

/// <summary>How serious a departure from the playbook is.</summary>
public enum Severity
{
    /// <summary>Worth noting; no change proposed.</summary>
    Info,

    /// <summary>Minor departure.</summary>
    Low,

    /// <summary>Departure a reviewer should decide on.</summary>
    Medium,

    /// <summary>Departure that normally blocks signature.</summary>
    High,
}

/// <summary>What the reviewer is allowed to do when a rule is breached.</summary>
public enum FindingAction
{
    /// <summary>Attach a comment only; never alter the wording.</summary>
    CommentOnly,

    /// <summary>Propose replacement wording as a tracked change, plus a comment.</summary>
    Redline,
}

/// <summary>
/// One house position, expressed as data rather than as prompt text. The patterns
/// locate candidate wording deterministically; the model only decides whether a
/// candidate actually breaches the rule.
/// </summary>
public sealed record PlaybookRule
{
    /// <summary>Gets the stable rule identifier, used in comments and the report.</summary>
    public string Id { get; init; } = string.Empty;

    /// <summary>Gets the short human title.</summary>
    public string Title { get; init; } = string.Empty;

    /// <summary>Gets the severity applied to findings from this rule.</summary>
    public Severity Severity { get; init; } = Severity.Medium;

    /// <summary>Gets the literal or regular-expression patterns that locate candidates.</summary>
    public IReadOnlyList<string> Patterns { get; init; } = Array.Empty<string>();

    /// <summary>Gets a value indicating whether <see cref="Patterns"/> are regular expressions.</summary>
    public bool Regex { get; init; }

    /// <summary>Gets the house position, given to the model as the test to apply.</summary>
    public string Rationale { get; init; } = string.Empty;

    /// <summary>Gets optional preferred wording the model should move towards.</summary>
    public string? PreferredWording { get; init; }

    /// <summary>Gets what may be done when the rule is breached.</summary>
    public FindingAction Action { get; init; } = FindingAction.CommentOnly;
}

/// <summary>A named set of rules.</summary>
public sealed record Playbook
{
    /// <summary>Gets the playbook name, shown in the report header.</summary>
    public string Name { get; init; } = "Playbook";

    /// <summary>Gets the rules.</summary>
    public IReadOnlyList<PlaybookRule> Rules { get; init; } = Array.Empty<PlaybookRule>();

    private static readonly JsonSerializerOptions Options = new()
    {
        PropertyNameCaseInsensitive = true,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase, allowIntegerValues: false) },
    };

    /// <summary>Loads a playbook from JSON, failing loudly on a malformed file.</summary>
    /// <param name="json">The playbook JSON.</param>
    /// <returns>The parsed playbook.</returns>
    /// <exception cref="InvalidDataException">The JSON is not a usable playbook.</exception>
    public static Playbook FromJson(string json)
    {
        Playbook? parsed;
        try
        {
            parsed = JsonSerializer.Deserialize<Playbook>(json, Options);
        }
        catch (JsonException ex)
        {
            throw new InvalidDataException($"The playbook is not valid JSON: {ex.Message}", ex);
        }

        if (parsed is null || parsed.Rules.Count == 0)
        {
            throw new InvalidDataException("The playbook contains no rules.");
        }

        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var rule in parsed.Rules)
        {
            if (string.IsNullOrWhiteSpace(rule.Id))
            {
                throw new InvalidDataException("Every rule needs an id.");
            }

            if (!seen.Add(rule.Id))
            {
                throw new InvalidDataException($"Duplicate rule id '{rule.Id}'.");
            }

            if (rule.Patterns.Count == 0 || rule.Patterns.Any(string.IsNullOrWhiteSpace))
            {
                throw new InvalidDataException($"Rule '{rule.Id}' needs at least one non-empty pattern.");
            }

            if (rule.Regex)
            {
                foreach (var pattern in rule.Patterns)
                {
                    try
                    {
                        // Compiling here turns a playbook typo into a startup
                        // failure naming the rule, instead of a rule that
                        // silently reports "no match" at review time.
                        _ = Regex.Match(string.Empty, pattern, RegexOptions.None, TimeSpan.FromSeconds(1));
                    }
                    catch (ArgumentException ex)
                    {
                        throw new InvalidDataException(
                            $"Rule '{rule.Id}' has an invalid regular expression '{pattern}': {ex.Message}", ex);
                    }
                }
            }

            if (string.IsNullOrWhiteSpace(rule.Rationale))
            {
                throw new InvalidDataException($"Rule '{rule.Id}' needs a rationale; it is the test the model applies.");
            }
        }

        return parsed;
    }

    /// <summary>Loads a playbook from a file.</summary>
    /// <param name="path">Path to the playbook JSON.</param>
    /// <returns>The parsed playbook.</returns>
    public static Playbook FromFile(string path) => FromJson(File.ReadAllText(path));
}
