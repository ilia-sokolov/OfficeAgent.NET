using System.Text.Json;
using OfficeAgent.Samples.RequirementsReview;
using Xunit;

namespace RequirementsReview.Tests;

/// <summary>Requirement sets and policies load strictly or not at all.</summary>
public sealed class ConfigurationTests
{
    private const string Minimal =
        """
        {
          "id": "S", "title": "t", "version": "1",
          "requirements": [
            { "id": "R1", "version": "1", "kind": "semantic", "title": "t", "text": "x", "section": "s" }
          ]
        }
        """;

    [Fact]
    public void The_bundled_set_and_policy_load()
    {
        var set = RequirementSet.FromFile(Path.Combine(AppContext.BaseDirectory, "requirements.json"));
        var policy = ReviewPolicy.FromFile(Path.Combine(AppContext.BaseDirectory, "review-policy.json"));

        Assert.Equal(("MS-REQ", "2026.4", 4), (set.Id, set.Version, set.Requirements.Count));
        Assert.Single(set.Requirements, r => r.Kind == RequirementKind.Deterministic);
        Assert.Equal(64, set.Sha256.Length);
        Assert.Equal("rp-2026.09.1", policy.Version);
        Assert.Contains(Harness.Reviewer, policy.AuthorizedReviewers);
    }

    [Theory]
    [InlineData("\"section\": \"s\"", "\"section\": \"s\", \"patterns\": [\"(\"]")]
    [InlineData("\"kind\": \"semantic\"", "\"kind\": \"judgement\"")]
    [InlineData("\"kind\": \"semantic\"", "\"kind\": \"deterministic\"")]
    [InlineData("\"text\": \"x\", ", "")]
    [InlineData("\"section\": \"s\"", "\"section\": \"s\", \"proposal\": \"rewrite\"")]
    public void A_defective_requirement_set_is_refused(string find, string replace)
    {
        Assert.Throws<InvalidDataException>(() => RequirementSet.FromJson(Minimal.Replace(find, replace, StringComparison.Ordinal)));
    }

    [Fact]
    public void Duplicate_ids_are_refused()
    {
        var json = Minimal.Replace(
            "\"section\": \"s\" }",
            "\"section\": \"s\" }, { \"id\": \"R1\", \"version\": \"2\", \"kind\": \"semantic\", \"text\": \"y\", \"section\": \"s\" }",
            StringComparison.Ordinal);
        Assert.Throws<InvalidDataException>(() => RequirementSet.FromJson(json));
    }

    [Fact]
    public void Unknown_fields_are_refused_rather_than_ignored()
    {
        Assert.Throws<JsonException>(() => RequirementSet.FromJson(Minimal.Replace("\"title\": \"t\", \"text\"", "\"title\": \"t\", \"threshold\": 0.1, \"text\"", StringComparison.Ordinal)));
    }

    [Theory]
    [InlineData("\"minResultConfidence\": 0.80", "\"minResultConfidence\": 0")]
    [InlineData("\"minResultConfidence\": 0.80", "\"minResultConfidence\": 1.5")]
    [InlineData("\"materialityProposalScore\": 1.5", "\"materialityProposalScore\": 4")]
    [InlineData("\"timeoutSeconds\": 30", "\"timeoutSeconds\": 0")]
    public void A_policy_out_of_range_is_refused(string find, string replace)
    {
        var json = File.ReadAllText(Path.Combine(AppContext.BaseDirectory, "review-policy.json"));
        Assert.Contains(find, json);
        Assert.Throws<InvalidDataException>(() => ReviewPolicy.FromJson(json.Replace(find, replace, StringComparison.Ordinal)));
    }

    [Fact]
    public void Any_change_to_the_set_changes_its_hash()
    {
        var a = RequirementSet.FromJson(Minimal);
        var b = RequirementSet.FromJson(Minimal.Replace("\"text\": \"x\"", "\"text\": \"x.\"", StringComparison.Ordinal));
        Assert.NotEqual(a.Sha256, b.Sha256);
    }
}
