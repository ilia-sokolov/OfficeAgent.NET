namespace OfficeAgent.Abstractions;

/// <summary>
/// Defines a text lookup request.
/// </summary>
public sealed class FindQuery
{
    /// <summary>Gets the text or regular expression pattern to find.</summary>
    public string Pattern { get; init; } = string.Empty;

    /// <summary>Gets matching options for the query.</summary>
    public MatchOptions Options { get; init; } = new();

    /// <summary>Initializes a new instance of the <see cref="FindQuery"/> class.</summary>
    public FindQuery()
    {
    }

    /// <summary>Initializes a new instance of the <see cref="FindQuery"/> class.</summary>
    /// <param name="pattern">The text pattern to find.</param>
    public FindQuery(string pattern) => Pattern = pattern;
}

/// <summary>
/// Defines options for text matching.
/// </summary>
public sealed class MatchOptions
{
    /// <summary>Gets a value indicating whether matches must align to word boundaries.</summary>
    public bool WholeWord { get; init; }

    /// <summary>Gets a value indicating whether matching is case-sensitive.</summary>
    public bool CaseSensitive { get; init; }

    /// <summary>Gets a value indicating whether <see cref="FindQuery.Pattern"/> is interpreted as a regular expression.</summary>
    public bool Regex { get; init; }
}

/// <summary>
/// Represents one text match and the anchor that can target it in a plan.
/// </summary>
public sealed class FindHit
{
    /// <summary>Gets the content-verified anchor for the match.</summary>
    public Anchor Anchor { get; init; } = null!;

    /// <summary>Gets the matched text.</summary>
    public string Text { get; init; } = string.Empty;

    /// <summary>Gets surrounding text for preview display.</summary>
    public string Context { get; init; } = string.Empty;
}
