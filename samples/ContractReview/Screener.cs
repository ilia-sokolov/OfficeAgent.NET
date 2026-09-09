using OfficeAgent.Abstractions;
using OfficeAgent.Core;

namespace OfficeAgent.Samples.ContractReview;

/// <summary>
/// Deterministic first pass. Every candidate anchor comes from the engine's own
/// find, so the model is never asked to invent a location.
/// </summary>
public sealed class Screener
{
    private readonly OfficeAgentClient _client;

    /// <summary>Initializes a new instance of the <see cref="Screener"/> class.</summary>
    /// <param name="client">The OfficeAgent client.</param>
    public Screener(OfficeAgentClient client) => _client = client;

    /// <summary>Screens a document against every rule in the playbook.</summary>
    /// <param name="connectionId">Connection holding the document.</param>
    /// <param name="documentId">Opaque document id.</param>
    /// <param name="playbook">The playbook to apply.</param>
    /// <param name="cancellationToken">Cancellation token.</param>
    /// <returns>Candidates and the rules that matched nothing.</returns>
    public async Task<ScreeningResult> ScreenAsync(
        string connectionId,
        string documentId,
        Playbook playbook,
        CancellationToken cancellationToken = default)
    {
        var candidates = new List<Candidate>();
        var undetected = new List<UndetectedRule>();
        var seenSpans = new HashSet<string>(StringComparer.Ordinal);

        foreach (var rule in playbook.Rules)
        {
            var before = candidates.Count;

            foreach (var pattern in rule.Patterns)
            {
                // Patterns are validated when the playbook loads, so a failure
                // here is an engine problem and must not be mistaken for "this
                // rule found nothing".
                IReadOnlyList<FindHit> hits = await _client.FindAsync(
                        connectionId,
                        documentId,
                        new FindQuery
                        {
                            Pattern = pattern,
                            Options = new MatchOptions { Regex = rule.Regex, CaseSensitive = false },
                        },
                        cancellationToken).ConfigureAwait(false);

                foreach (var hit in hits)
                {
                    if (hit.Anchor is not TextSpanAnchor span)
                    {
                        continue;
                    }

                    // One rule can carry several patterns that match the same span.
                    var key = $"{rule.Id}|{span.ParaId}|{span.Occurrence}|{hit.Text}";
                    if (!seenSpans.Add(key))
                    {
                        continue;
                    }

                    candidates.Add(new Candidate
                    {
                        Id = $"c{candidates.Count + 1:D3}",
                        RuleId = rule.Id,
                        ParaId = span.ParaId,
                        MatchedText = hit.Text,
                        Occurrence = span.Occurrence,
                        Context = hit.Context,
                    });
                }
            }

            if (candidates.Count == before)
            {
                undetected.Add(new UndetectedRule { RuleId = rule.Id, Title = rule.Title });
            }
        }

        return new ScreeningResult(candidates, undetected);
    }
}

/// <summary>The outcome of screening.</summary>
/// <param name="Candidates">Locations for the model to judge.</param>
/// <param name="Undetected">
/// Rules whose patterns matched nothing. Reported separately, because "no match"
/// means the patterns did not fire, which is not the same as "the contract complies".
/// </param>
public sealed record ScreeningResult(
    IReadOnlyList<Candidate> Candidates,
    IReadOnlyList<UndetectedRule> Undetected);
