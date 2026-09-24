using System.Security.Cryptography;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;

namespace OfficeAgent.Samples.RequirementsReview;

/// <summary>How a requirement is decided.</summary>
public enum RequirementKind
{
    /// <summary>Decided by host code alone. Never sent to an external service.</summary>
    Deterministic,

    /// <summary>Needs a semantic judgement over anchored evidence.</summary>
    Semantic,
}

/// <summary>What the host may propose in Word when a semantic requirement is violated.</summary>
public enum ProposalMode
{
    /// <summary>A comment only; the reviewer writes any fix.</summary>
    CommentOnly,

    /// <summary>A comment plus one tracked change the reviewer can accept or reject.</summary>
    TrackedChange,
}

/// <summary>What a failed deterministic rule does to the run.</summary>
public enum HardRuleFailure
{
    /// <summary>Stop the run before any content leaves the host.</summary>
    Stop,

    /// <summary>Record the failure and route the requirement to a person.</summary>
    HumanReview,
}

/// <summary>One requirement, exactly as registered. The agent never sees or edits this type.</summary>
public sealed record RegisteredRequirement
{
    /// <summary>Gets the requirement id.</summary>
    public required string Id { get; init; }

    /// <summary>Gets the requirement's own version.</summary>
    public required string Version { get; init; }

    /// <summary>Gets how the requirement is decided.</summary>
    public required RequirementKind Kind { get; init; }

    /// <summary>Gets a short title.</summary>
    public required string Title { get; init; }

    /// <summary>Gets the requirement text that is judged.</summary>
    public required string Text { get; init; }

    /// <summary>Gets the heading of the document section that holds the evidence.</summary>
    public required string Section { get; init; }

    /// <summary>Gets deterministic patterns that must all match (deterministic requirements).</summary>
    public IReadOnlyList<Regex> Patterns { get; init; } = Array.Empty<Regex>();

    /// <summary>
    /// Gets patterns that must all match before a semantic "satisfied" is accepted. They can
    /// block a pass; they can never produce one.
    /// </summary>
    public IReadOnlyList<Regex> PassRequires { get; init; } = Array.Empty<Regex>();

    /// <summary>Gets what a failed deterministic rule does.</summary>
    public HardRuleFailure OnFail { get; init; } = HardRuleFailure.HumanReview;

    /// <summary>Gets what may be proposed when a semantic requirement is violated.</summary>
    public ProposalMode Proposal { get; init; } = ProposalMode.CommentOnly;
}

/// <summary>A versioned requirement set, loaded once and never modified.</summary>
public sealed class RequirementSet
{
    /// <summary>Regex timeout for every requirement pattern. A timeout fails the check.</summary>
    public static readonly TimeSpan PatternTimeout = TimeSpan.FromMilliseconds(250);

    private RequirementSet(string id, string title, string version, string sha256, IReadOnlyList<RegisteredRequirement> requirements)
    {
        Id = id;
        Title = title;
        Version = version;
        Sha256 = sha256;
        Requirements = requirements;
    }

    /// <summary>Gets the set id.</summary>
    public string Id { get; }

    /// <summary>Gets the set title.</summary>
    public string Title { get; }

    /// <summary>Gets the set version.</summary>
    public string Version { get; }

    /// <summary>Gets the SHA-256 of the exact JSON the set was loaded from.</summary>
    public string Sha256 { get; }

    /// <summary>Gets the requirements in registration order.</summary>
    public IReadOnlyList<RegisteredRequirement> Requirements { get; }

    /// <summary>Loads a requirement set from a file.</summary>
    /// <param name="path">Path to the JSON file.</param>
    /// <returns>The set.</returns>
    public static RequirementSet FromFile(string path) => FromJson(File.ReadAllText(path));

    /// <summary>Loads and validates a requirement set. Any defect throws; nothing is guessed.</summary>
    /// <param name="json">The JSON text.</param>
    /// <returns>The set.</returns>
    public static RequirementSet FromJson(string json)
    {
        var dto = JsonSerializer.Deserialize<SetDto>(json, JsonOptions)
                  ?? throw new InvalidDataException("Requirement set is empty.");

        Require(dto.Id, "id");
        Require(dto.Version, "version");
        if (dto.Requirements is not { Count: > 0 })
        {
            throw new InvalidDataException("Requirement set has no requirements.");
        }

        var seen = new HashSet<string>(StringComparer.Ordinal);
        var requirements = new List<RegisteredRequirement>();
        foreach (var r in dto.Requirements)
        {
            Require(r.Id, "requirements[].id");
            Require(r.Version, $"{r.Id}.version");
            Require(r.Text, $"{r.Id}.text");
            Require(r.Section, $"{r.Id}.section");
            if (!seen.Add(r.Id!))
            {
                throw new InvalidDataException($"Duplicate requirement id '{r.Id}'.");
            }

            var kind = r.Kind switch
            {
                "deterministic" => RequirementKind.Deterministic,
                "semantic" => RequirementKind.Semantic,
                _ => throw new InvalidDataException($"{r.Id}: unknown kind '{r.Kind}'."),
            };

            var patterns = Compile(r.Id!, r.Patterns);
            if (kind == RequirementKind.Deterministic && patterns.Count == 0)
            {
                throw new InvalidDataException($"{r.Id}: a deterministic requirement needs at least one pattern.");
            }

            requirements.Add(new RegisteredRequirement
            {
                Id = r.Id!,
                Version = r.Version!,
                Kind = kind,
                Title = r.Title ?? r.Id!,
                Text = r.Text!,
                Section = r.Section!,
                Patterns = patterns,
                PassRequires = Compile(r.Id!, r.PassRequires),
                OnFail = r.OnFail switch
                {
                    null or "humanReview" => HardRuleFailure.HumanReview,
                    "stop" => HardRuleFailure.Stop,
                    _ => throw new InvalidDataException($"{r.Id}: unknown onFail '{r.OnFail}'."),
                },
                Proposal = r.Proposal switch
                {
                    null or "commentOnly" => ProposalMode.CommentOnly,
                    "trackedChange" => ProposalMode.TrackedChange,
                    _ => throw new InvalidDataException($"{r.Id}: unknown proposal '{r.Proposal}'."),
                },
            });
        }

        return new RequirementSet(dto.Id!, dto.Title ?? dto.Id!, dto.Version!, Hash(json), requirements);
    }

    internal static string Hash(string text) =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(text))).ToLowerInvariant();

    private static IReadOnlyList<Regex> Compile(string id, List<string>? patterns)
    {
        var compiled = new List<Regex>();
        foreach (var pattern in patterns ?? new List<string>())
        {
            try
            {
                compiled.Add(new Regex(pattern, RegexOptions.CultureInvariant | RegexOptions.IgnoreCase, PatternTimeout));
            }
            catch (ArgumentException ex)
            {
                throw new InvalidDataException($"{id}: invalid pattern '{pattern}': {ex.Message}", ex);
            }
        }

        return compiled;
    }

    private static void Require(string? value, string name)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            throw new InvalidDataException($"Requirement set field '{name}' is required.");
        }
    }

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        PropertyNameCaseInsensitive = false,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        UnmappedMemberHandling = JsonUnmappedMemberHandling.Disallow,
    };

    private sealed class SetDto
    {
        public string? Id { get; set; }
        public string? Title { get; set; }
        public string? Version { get; set; }
        public List<RequirementDto>? Requirements { get; set; }
    }

    private sealed class RequirementDto
    {
        public string? Id { get; set; }
        public string? Version { get; set; }
        public string? Kind { get; set; }
        public string? Title { get; set; }
        public string? Text { get; set; }
        public string? Section { get; set; }
        public List<string>? Patterns { get; set; }
        public List<string>? PassRequires { get; set; }
        public string? OnFail { get; set; }
        public string? Proposal { get; set; }
    }
}
