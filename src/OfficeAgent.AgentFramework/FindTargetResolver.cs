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
/// first of several is the one outcome an agent cannot detect from the result.
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
        var failures = new List<Failure>();
        if (plan["operations"] is not JsonArray operations || operations.Count == 0)
            return failures;

        var pending = operations
            .Select((node, index) => (Node: node as JsonObject, Index: index))
            .Where(entry => entry.Node?["target"] is JsonObject target && HasFind(target))
            .ToList();
        if (pending.Count == 0) return failures;

        // One storage read for the whole plan; repeated patterns then cost nothing.
        byte[] bytes;
        string? name;
        using (var content = await _client.OpenReadAsync(reference, cancellationToken).ConfigureAwait(false))
        {
            name = content.Reference.Name;
            bytes = await ReadAllAsync(content.Stream, cancellationToken).ConfigureAwait(false);
        }

        var matchesByPattern = new Dictionary<string, IReadOnlyList<FindHit>>(StringComparer.Ordinal);

        foreach (var (node, index) in pending)
        {
            var target = (JsonObject)node!["target"]!;
            if (!TryGetString(target, "find", out var pattern) || pattern.Length == 0)
            {
                failures.Add(new Failure("invalid-argument",
                    $"Operation {index}: \"find\" must be a non-empty string."));
                continue;
            }

            if (!matchesByPattern.TryGetValue(pattern, out var hits))
            {
                hits = _client.Find(
                    new StreamHandle(new MemoryStream(bytes, writable: false), name),
                    new FindQuery { Pattern = pattern });
                matchesByPattern[pattern] = hits;
            }

            var requestedMatch = TryGetInt(target, "match", out var match) ? match : (int?)null;
            var chosen = Choose(hits, pattern, requestedMatch, index, failures);
            if (chosen is null) continue;

            Bind(target, chosen);
        }

        return failures;
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
            failures.Add(new Failure("anchor-not-found",
                $"Operation {operationIndex}: no text matching \"{pattern}\" is in the document. " +
                "Check the wording against inspect_document rather than retrying the same text."));
            return null;
        }

        if (requestedMatch is { } index)
        {
            if (index < 0 || index >= hits.Count)
            {
                failures.Add(new Failure("anchor-not-found",
                    $"Operation {operationIndex}: \"match\": {index} is out of range - " +
                    $"\"{pattern}\" matches {hits.Count} time(s), so valid values are 0 to {hits.Count - 1}."));
                return null;
            }
            return hits[index];
        }

        if (hits.Count > 1)
        {
            var candidates = string.Join("; ", hits.Select((hit, i) => $"match {i}: …{hit.Context}…"));
            failures.Add(new Failure("ambiguous-anchor",
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

        var anchor = hit.Anchor as TextSpanAnchor;
        target["paraId"] = anchor?.ParaId ?? string.Empty;
        target["expect"] = hit.Text;
        target["occurrence"] = anchor?.Occurrence ?? 0;
    }

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

    private static bool TryGetInt(JsonObject target, string name, out int value)
    {
        value = 0;
        if (!target.TryGetPropertyValue(name, out var node) || node is not JsonValue jsonValue)
            return false;
        return jsonValue.TryGetValue<int>(out value);
    }

    private static async Task<byte[]> ReadAllAsync(Stream stream, CancellationToken cancellationToken)
    {
        if (stream is MemoryStream memory) return memory.ToArray();
        using var copy = new MemoryStream();
        await stream.CopyToAsync(copy, 81920, cancellationToken).ConfigureAwait(false);
        return copy.ToArray();
    }
}
