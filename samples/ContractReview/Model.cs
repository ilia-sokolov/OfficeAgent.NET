namespace OfficeAgent.Samples.ContractReview;

/// <summary>
/// A location the screener found for one rule. The host mints the id; the model
/// refers to a candidate by that id and never retypes anchor data, so a model
/// slip cannot produce an unresolvable anchor.
/// </summary>
public sealed record Candidate
{
    /// <summary>Gets the host-assigned id the model quotes back.</summary>
    public string Id { get; init; } = string.Empty;

    /// <summary>Gets the rule that produced this candidate.</summary>
    public string RuleId { get; init; } = string.Empty;

    /// <summary>Gets the paragraph containing the match.</summary>
    public string ParaId { get; init; } = string.Empty;

    /// <summary>Gets the exact matched text, engine-sourced.</summary>
    public string MatchedText { get; init; } = string.Empty;

    /// <summary>Gets the occurrence index of the match within its paragraph.</summary>
    public int Occurrence { get; init; }

    /// <summary>Gets surrounding text, so the model can judge in context.</summary>
    public string Context { get; init; } = string.Empty;
}

/// <summary>The model's judgement about one candidate.</summary>
public sealed record Finding
{
    /// <summary>Gets the candidate this judgement is about.</summary>
    public string CandidateId { get; init; } = string.Empty;

    /// <summary>Gets a value indicating whether the rule is actually breached here.</summary>
    public bool Violation { get; init; }

    /// <summary>Gets the one-sentence reason, written for a lawyer.</summary>
    public string Rationale { get; init; } = string.Empty;

    /// <summary>Gets proposed replacement wording, when the rule allows a redline.</summary>
    public string? ReplacementText { get; init; }
}

/// <summary>Why a finding did not become a tracked change.</summary>
public enum RedlineOutcome
{
    /// <summary>No redline was requested by the rule.</summary>
    NotRequested,

    /// <summary>Replacement wording was applied as a tracked change.</summary>
    Applied,

    /// <summary>The rule allowed a redline but the model proposed no wording.</summary>
    NoWordingProposed,

    /// <summary>Another finding on the same span carried the redline.</summary>
    SupersededOnSpan,
}

/// <summary>One row of the review report.</summary>
public sealed record ReviewItem
{
    /// <summary>Gets the rule id.</summary>
    public string RuleId { get; init; } = string.Empty;

    /// <summary>Gets the rule title.</summary>
    public string Title { get; init; } = string.Empty;

    /// <summary>Gets the severity.</summary>
    public Severity Severity { get; init; }

    /// <summary>Gets the matched text, or an empty string when nothing was found.</summary>
    public string MatchedText { get; init; } = string.Empty;

    /// <summary>Gets a value indicating whether the model judged this a breach.</summary>
    public bool Violation { get; init; }

    /// <summary>Gets the model's reason.</summary>
    public string Rationale { get; init; } = string.Empty;

    /// <summary>Gets what happened to the wording.</summary>
    public RedlineOutcome Redline { get; init; }
}

/// <summary>Rules that produced no candidates at all.</summary>
public sealed record UndetectedRule
{
    /// <summary>Gets the rule id.</summary>
    public string RuleId { get; init; } = string.Empty;

    /// <summary>Gets the rule title.</summary>
    public string Title { get; init; } = string.Empty;
}
