using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using OfficeAgent.Abstractions;

namespace OfficeAgent.Samples.RequirementsReview;

/// <summary>
/// The machine-readable record of one run. It holds enough to recompute every route from the
/// recorded answers and policy, and nothing secret: no API keys, no tokens, no headers.
/// Evidence excerpts are included because the record must show what was judged; store it
/// with the same protection as the document.
/// </summary>
public static class AuditRecord
{
    /// <summary>The record schema id.</summary>
    public const string Schema = "requirements-review-audit/1";

    private static readonly JsonSerializerOptions Options = new()
    {
        WriteIndented = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        Converters = { new JsonStringEnumConverter(JsonNamingPolicy.CamelCase) },
    };

    /// <summary>Builds the record.</summary>
    /// <param name="proposal">The proposal phase.</param>
    /// <param name="evaluator">Evaluator description: provider and endpoint.</param>
    /// <param name="commit">The commit phase, when it ran.</param>
    /// <returns>The record as a JSON object.</returns>
    public static JsonObject Build(ReviewProposal proposal, EvaluatorInfo evaluator, CommitOutcome? commit = null)
    {
        var e = proposal.Evidence;
        var record = new JsonObject
        {
            ["schema"] = Schema,
            ["runId"] = proposal.RunId,
            ["startedAtUtc"] = proposal.StartedAtUtc.ToString("O"),
            ["document"] = new JsonObject
            {
                ["connectionId"] = e.ConnectionId,
                ["documentId"] = e.DocumentId,
                ["name"] = e.Name,
                ["sha256"] = e.SourceSha256,
                ["snapshotETag"] = e.Snapshot.ETag,
                ["existingComments"] = e.ExistingCommentCount,
                ["existingRevisions"] = e.ExistingRevisionCount,
                ["tables"] = e.TableCount,
            },
            ["requirementSet"] = new JsonObject
            {
                ["id"] = proposal.Requirements.Id,
                ["version"] = proposal.Requirements.Version,
                ["sha256"] = proposal.Requirements.Sha256,
            },
            ["policy"] = new JsonObject
            {
                ["version"] = proposal.Policy.Version,
                ["sha256"] = proposal.Policy.Sha256,
                ["thresholds"] = JsonSerializer.SerializeToNode(proposal.Policy.Thresholds, Options),
                ["maxEvidenceCharacters"] = proposal.Policy.MaxEvidenceCharacters,
            },
            ["evaluator"] = new JsonObject
            {
                ["provider"] = evaluator.Provider,
                ["endpoint"] = evaluator.Endpoint,
                ["requestedModel"] = proposal.Policy.EvaluatorModel,
                ["questionSetVersion"] = QuestionSet.Version,
                ["questionSetSha256"] = RequirementSet.Hash(QuestionSet.ToJson()),
                ["questions"] = QuestionSet.Build(),
            },
            ["status"] = new JsonObject
            {
                ["proposal"] = proposal.Status.ToString(),
                ["reason"] = proposal.StatusReason,
            },
            ["requirements"] = new JsonArray(proposal.Items.Select(Item).ToArray()),
        };

        if (proposal.Built is { } built)
        {
            record["plan"] = new JsonObject
            {
                ["sha256"] = proposal.PreviewReceipt?.PlanSha256,
                ["snapshotETag"] = built.Plan.Snapshot?.ETag,
                ["operations"] = JsonSerializer.SerializeToNode(built.Plan.Operations.Select(Describe).ToList(), Options),
            };
        }

        if (proposal.Preview is { } preview)
        {
            record["preview"] = new JsonObject
            {
                ["valid"] = preview.IsValid,
                ["errors"] = new JsonArray(preview.Errors.Select(x => (JsonNode?)JsonValue.Create($"{x.Code}: {x.Message}")).ToArray()),
                ["receipt"] = Receipt(proposal.PreviewReceipt),
            };
        }

        if (commit is not null)
        {
            if (commit.Approval is { } approval)
            {
                record["humanDecision"] = new JsonObject
                {
                    ["reviewerId"] = approval.ReviewerId,
                    ["decision"] = approval.Decision.ToString(),
                    ["decidedAtUtc"] = approval.DecidedAtUtc.ToString("O"),
                    ["planSha256"] = approval.PlanSha256,
                    ["inputSha256"] = approval.InputSha256,
                    ["snapshotETag"] = approval.SnapshotETag,
                };
            }

            record["commit"] = new JsonObject
            {
                ["status"] = commit.Status.ToString(),
                ["reason"] = commit.Reason,
                ["revalidationReceipt"] = Receipt(commit.Revalidation),
                ["receipt"] = Receipt(commit.Receipt),
                ["outputDocumentId"] = commit.Output?.ItemId,
                ["outputName"] = commit.Output?.Name,
            };
        }

        return record;
    }

    /// <summary>Serializes a record.</summary>
    /// <param name="record">The record.</param>
    /// <returns>Indented JSON.</returns>
    public static string ToJson(JsonObject record) => record.ToJsonString(Options);

    private static JsonNode Item(ReviewItem item)
    {
        var node = new JsonObject
        {
            ["requirementId"] = item.Requirement.Id,
            ["version"] = item.Requirement.Version,
            ["kind"] = item.Requirement.Kind.ToString(),
            ["text"] = item.Requirement.Text,
            ["section"] = item.Evidence.Section,
            ["sectionMatches"] = item.Evidence.SectionMatches,
            ["headingParaId"] = item.Evidence.HeadingParaId,
            ["evidence"] = new JsonArray(item.Evidence.Items.Select(i => (JsonNode?)new JsonObject
            {
                ["evidenceId"] = i.EvidenceId,
                ["paraId"] = i.ParaId,
                ["container"] = i.Container,
                ["excerpt"] = i.Text,
                ["pendingRevision"] = i.HasPendingRevision,
                ["existingComment"] = i.HasExistingComment,
            }).ToArray()),
            ["deterministic"] = JsonSerializer.SerializeToNode(item.Deterministic, Options),
            ["route"] = item.Routing.Route.ToString(),
            ["outcome"] = item.Routing.Outcome.ToString(),
            ["wordAction"] = item.Routing.Action.ToString(),
            ["routingRule"] = item.Routing.Rule,
            ["reason"] = item.Routing.Reason,
            ["note"] = item.Routing.Note,
        };

        if (item.Decision is { } d)
        {
            node["evaluation"] = new JsonObject
            {
                ["provider"] = d.Provider,
                ["model"] = d.Model,
                ["requestId"] = d.RequestId,
                ["questionSetVersion"] = d.QuestionSetVersion,
                ["stateSha256"] = d.StateSha256,
                ["evidenceIdsSent"] = new JsonArray(d.EvidenceIdsSent.Select(x => (JsonNode?)JsonValue.Create(x)).ToArray()),
                ["answers"] = d.IsComplete ? new JsonObject
                {
                    [QuestionSet.EvidenceStatusName] = JsonSerializer.SerializeToNode(d.EvidenceStatus, Options),
                    [QuestionSet.RequirementResultName] = JsonSerializer.SerializeToNode(d.Result, Options),
                    [QuestionSet.MaterialityName] = JsonSerializer.SerializeToNode(d.Materiality, Options),
                } : null,
                ["failure"] = d.Failure is null ? null : JsonSerializer.SerializeToNode(d.Failure, Options),
                ["inputTokens"] = d.InputTokens,
                ["outputTokens"] = d.OutputTokens,
                ["elapsedMs"] = (long)d.Elapsed.TotalMilliseconds,
                ["attempts"] = d.Attempts,
            };
        }

        if (item.Draft is { } draft)
        {
            node["draft"] = new JsonObject
            {
                ["drafter"] = "microsoft-agent-framework",
                ["model"] = draft.Model,
                ["proposal"] = draft.Proposal is null ? null : JsonSerializer.SerializeToNode(draft.Proposal, Options),
                ["rejections"] = new JsonArray(draft.Rejections.Select(x => (JsonNode?)JsonValue.Create(x)).ToArray()),
                ["error"] = draft.Error,
            };
        }

        return node;
    }

    private static object Describe(PlanOperation op) => op switch
    {
        CommentOp c when c.Target is TextSpanAnchor a => new { op = "comment", paraId = a.ParaId, expect = a.Expect, text = c.Text, author = c.Author },
        ChangeTextOp t when t.Target is TextSpanAnchor a => new { op = "changeText", mode = t.Mode.ToString(), paraId = a.ParaId, expect = a.Expect, with = t.With },
        _ => new { op = op.GetType().Name },
    };

    private static JsonNode? Receipt(ApplyReceipt? receipt) => receipt is null ? null : new JsonObject
    {
        ["receiptVersion"] = receipt.ReceiptVersion,
        ["outcome"] = receipt.Outcome.ToString(),
        ["planSha256"] = receipt.PlanSha256,
        ["inputSha256"] = receipt.InputSha256,
        ["outputSha256"] = receipt.OutputSha256,
        ["timestampUtc"] = receipt.TimestampUtc.ToString("O"),
        ["revisionAuthor"] = receipt.Revision.Author,
        ["actor"] = receipt.Actor?.Subject,
        ["outputDocumentId"] = receipt.OutputDocument?.ItemId,
    };
}

/// <summary>Describes the evaluator for the audit record.</summary>
/// <param name="Provider">Provider name, e.g. <c>typesafe-jev</c> or <c>scripted-jev-wire</c>.</param>
/// <param name="Endpoint">The endpoint called, or a description of the scripted source.</param>
public sealed record EvaluatorInfo(string Provider, string Endpoint);
