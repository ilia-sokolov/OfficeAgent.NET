using System.Text.Json.Nodes;
using OfficeAgent.Abstractions;
using OfficeAgent.Core;

namespace OfficeAgent.AgentFramework;

/// <summary>
/// Binds <c>find</c> targets to the content-verified anchors the engine works in.
/// </summary>
/// <remarks>
/// <para>
/// An operation may address text it has not looked up yet - <c>"target": { "find": "Acme
/// Corp" }</c> - and this rewrites it into the <c>paraId</c> / <c>expect</c> /
/// <c>occurrence</c> triple a plan actually carries. That saves the agent a
/// <c>find_in_document</c> round trip per operation without weakening the rule that
/// anchors come from the document rather than from the model: the lookup still runs
/// against live content, and text the engine cannot locate is refused.
/// </para>
/// <para>
/// Resolution is deliberately strict about ambiguity. Text matching more than once is an
/// error unless the operation says which match it means, because silently editing the
/// first of several is the one outcome an agent cannot detect from the result. Saying
/// <c>"match": "all"</c> means every match, so the operation is repeated once per match - and
/// once per match may be not at all, which is what lets one plan run across many documents
/// that do not all contain the text.
/// </para>
/// <para>
/// This is an agent-surface convenience, not an engine concept: the rewritten plan that
/// reaches <see cref="OfficeAgentClient"/> is an ordinary plan, and the engine never sees
/// a <c>find</c> target.
/// </para>
/// </remarks>
internal sealed class FindTargetResolver
{
    private readonly OfficeAgentClient _client;

    internal FindTargetResolver(OfficeAgentClient client) => _client = client;

    /// <summary>One operation's failure to bind, in the wire vocabulary the tools return.</summary>
    internal readonly struct Failure
    {
        internal Failure(string code, string message)
        {
            Code = code;
            Message = message;
        }

        internal string Code { get; }
        internal string Message { get; }
    }

    /// <summary>
    /// Rewrites every <c>find</c> target in <paramref name="plan"/> in place. Returns the
    /// failures, so a caller can report all of them at once rather than making the agent
    /// discover them one call at a time. An empty result means the plan is ready to apply.
    /// </summary>
    internal async Task<IReadOnlyList<Failure>> ResolveAsync(
        DocumentReference reference,
        JsonObject plan,
        CancellationToken cancellationToken)
    {
        if (!NeedsResolving(plan)) return Array.Empty<Failure>();

        // One storage read for the whole plan; repeated patterns then cost nothing.
        using var content = await _client.OpenReadAsync(reference, cancellationToken).ConfigureAwait(false);
        var bytes = await ReadAllAsync(content.Stream, cancellationToken).ConfigureAwait(false);

        return Resolve(bytes, content.Reference.Name, plan);
    }

    /// <summary>
    /// The same binding against content the caller already holds, for the inline-content
    /// tools: there is no storage to read from, and the document arrived with the plan.
    /// </summary>
    internal IReadOnlyList<Failure> Resolve(byte[] document, string? name, JsonObject plan)
    {
        if (!NeedsResolving(plan)) return Array.Empty<Failure>();
        return BindAll(document, name, plan);
    }

    /// <summary>Whether any operation names its target by text rather than by anchor.</summary>
    private static bool NeedsResolving(JsonObject plan) =>
        plan["operations"] is JsonArray operations &&
        operations.Any(node => node is JsonObject op && op["target"] is JsonObject target && HasFind(target));

    private IReadOnlyList<Failure> BindAll(byte[] bytes, string? name, JsonObject plan)
    {
        var failures = new List<Failure>();
        var operations = (JsonArray)plan["operations"]!;
        var matchesByPattern = new Dictionary<string, IReadOnlyList<FindHit>>(StringComparer.Ordinal);

        // An operation saying "match": "all" is repeated once per match, which can be no
        // times at all; every other operation is bound in place. Null means "in place".
        var expansions = new List<JsonObject>?[operations.Count];

        for (var index = 0; index < operations.Count; index++)
        {
            if (operations[index] is not JsonObject node ||
                node["target"] is not JsonObject target || !HasFind(target))
                continue;

            if (!TryGetString(target, "find", out var pattern) || pattern.Length == 0)
            {
                failures.Add(new Failure(ToolErrorCodes.InvalidArgument,
                    $"Operation {index}: \"find\" must be a non-empty string."));
                continue;
            }

            var caseSensitive = TryGetBool(target, "caseSensitive");
            var wholeWord = TryGetBool(target, "wholeWord");
            var key = $"{(caseSensitive ? 'C' : 'c')}{(wholeWord ? 'W' : 'w')}{pattern}";
            if (!matchesByPattern.TryGetValue(key, out var hits))
            {
                hits = _client.Find(
                    new StreamHandle(new MemoryStream(bytes, writable: false), name),
                    new FindQuery
                    {
                        Pattern = pattern,
                        Options = new MatchOptions { CaseSensitive = caseSensitive, WholeWord = wholeWord }
                    });
                matchesByPattern[key] = hits;
            }

            if (MatchesAll(target, index, failures, out var requestedMatch))
            {
                // Last match first. Edits to one paragraph apply in plan order, and an edit
                // shifts the occurrence numbers of the matches after it but never of those
                // before, so working backwards keeps every remaining anchor true.
                var copies = new List<JsonObject>();
                foreach (var hit in hits.Reverse())
                {
                    var copy = (JsonObject)node.DeepClone();
                    Bind((JsonObject)copy["target"]!, hit);
                    copies.Add(copy);
                }

                expansions[index] = copies;
                continue;
            }

            var chosen = Choose(hits, pattern, requestedMatch, index, failures);
            if (chosen is null) continue;

            Bind(target, chosen);
        }

        if (failures.Count == 0 && expansions.Any(copies => copies is not null))
        {
            var rebuilt = new JsonArray();
            for (var index = 0; index < operations.Count; index++)
            {
                if (expansions[index] is { } copies)
                    foreach (var copy in copies) rebuilt.Add(copy);
                else
                    rebuilt.Add(operations[index]?.DeepClone());
            }

            plan["operations"] = rebuilt;
        }

        return failures;
    }

    /// <summary>
    /// Whether the target asks for every match, reading <c>match</c> as an index otherwise.
    /// Any other string is refused rather than read as "no index", which would quietly turn
    /// an intended choice into an ambiguity or a guess.
    /// </summary>
    private static bool MatchesAll(JsonObject target, int index, List<Failure> failures, out int? requestedMatch)
    {
        requestedMatch = null;
        if (!target.TryGetPropertyValue("match", out var node) || node is not JsonValue value)
            return false;

        if (value.TryGetValue<string>(out var text))
        {
            if (string.Equals(text?.Trim(), "all", StringComparison.OrdinalIgnoreCase)) return true;
            failures.Add(new Failure(ToolErrorCodes.InvalidArgument,
                $"Operation {index}: \"match\" must be a zero-based index or \"all\", not \"{text}\"."));
            return false;
        }

        if (value.TryGetValue<int>(out var match)) requestedMatch = match;
        return false;
    }

    /// <summary>
    /// Picks the hit an operation meant, or records why it cannot be picked. Ambiguity is
    /// reported with each candidate's context so the agent can re-issue with a "match"
    /// index instead of guessing.
    /// </summary>
    private static FindHit? Choose(
        IReadOnlyList<FindHit> hits,
        string pattern,
        int? requestedMatch,
        int operationIndex,
        List<Failure> failures)
    {
        if (hits.Count == 0)
        {
            failures.Add(new Failure(ValidationErrorCodes.AnchorNotFound,
                $"Operation {operationIndex}: no text matching \"{pattern}\" is in the document. " +
                "Check the wording against inspect_document rather than retrying the same text."));
            return null;
        }

        if (requestedMatch is { } index)
        {
            if (index < 0 || index >= hits.Count)
            {
                failures.Add(new Failure(ValidationErrorCodes.AnchorNotFound,
                    $"Operation {operationIndex}: \"match\": {index} is out of range - " +
                    $"\"{pattern}\" matches {hits.Count} time(s), so valid values are 0 to {hits.Count - 1}."));
                return null;
            }
            return hits[index];
        }

        if (hits.Count > 1)
        {
            var candidates = string.Join("; ", hits.Select((hit, i) =>
                hit.Location is { Length: > 0 } location
                    ? $"match {i} ({location}): …{hit.Context}…"
                    : $"match {i}: …{hit.Context}…"));
            failures.Add(new Failure(ValidationErrorCodes.AmbiguousAnchor,
                $"Operation {operationIndex}: \"{pattern}\" matches {hits.Count} times, so it is unclear which one to edit. " +
                $"Re-issue the operation with \"match\": <index>, or use more surrounding text. Candidates - {candidates}"));
            return null;
        }

        return hits[0];
    }

    /// <summary>Replaces the lookup with the anchor it resolved to, leaving the rest of the operation untouched.</summary>
    private static void Bind(JsonObject target, FindHit hit)
    {
        target.Remove("find");
        target.Remove("match");
        target.Remove("caseSensitive");
        target.Remove("wholeWord");

        var anchor = hit.Anchor as TextSpanAnchor;
        target["paraId"] = anchor?.ParaId ?? string.Empty;
        target["expect"] = hit.Text;
        target["occurrence"] = anchor?.Occurrence ?? 0;
    }

    /// <summary>An optional boolean on a find target; absent means false, matching find_in_document's defaults.</summary>
    private static bool TryGetBool(JsonObject target, string name) =>
        target.TryGetPropertyValue(name, out var node) &&
        node is JsonValue value && value.TryGetValue<bool>(out var flag) && flag;

    private static bool HasFind(JsonObject target) =>
        target.TryGetPropertyValue("find", out var node) && node is not null;

    private static bool TryGetString(JsonObject target, string name, out string value)
    {
        value = string.Empty;
        if (!target.TryGetPropertyValue(name, out var node) || node is not JsonValue jsonValue)
            return false;
        if (!jsonValue.TryGetValue<string>(out var parsed) || parsed is null) return false;
        value = parsed;
        return true;
    }

    private static async Task<byte[]> ReadAllAsync(Stream stream, CancellationToken cancellationToken)
    {
        if (stream is MemoryStream memory) return memory.ToArray();
        using var copy = new MemoryStream();
        await stream.CopyToAsync(copy, 81920, cancellationToken).ConfigureAwait(false);
        return copy.ToArray();
    }
}
