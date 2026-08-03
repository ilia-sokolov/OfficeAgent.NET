using System.Text.Json;
using System.Text.Json.Serialization;
using OfficeAgent.Abstractions;

namespace OfficeAgent.Tests;

/// <summary>
/// The wire contract for plan operations. A model writes JSON in whatever order it likes,
/// so the verb has to be readable wherever it lands, and a payload it gets wrong has to
/// come back with something it can act on.
/// </summary>
public class PlanOperationJsonConverterTests
{
    private static readonly JsonSerializerOptions Options = new()
    {
        PropertyNameCaseInsensitive = true,
        Converters = { new JsonStringEnumConverter() }
    };

    [Fact]
    public void The_verb_is_read_wherever_it_appears_in_the_object()
    {
        // System.Text.Json's own polymorphism requires the discriminator first and reports
        // a payload like this as having no discriminator at all.
        const string json = """
            { "target": { "paraId": "w14:00000002", "expect": "Acme Corp" },
              "op": "changeText",
              "with": "Globex Inc.",
              "mode": "Direct" }
            """;

        var operation = Assert.IsType<ChangeTextOp>(
            JsonSerializer.Deserialize<PlanOperation>(json, Options));

        Assert.Equal("Globex Inc.", operation.With);
        Assert.Equal(ChangeMode.Direct, operation.Mode);
        Assert.Equal("Acme Corp", Assert.IsType<TextSpanAnchor>(operation.Target).Expect);
    }

    [Fact]
    public void A_plan_round_trips_through_serialization()
    {
        var plan = new DocumentPlan
        {
            Operations = new PlanOperation[]
            {
                new ChangeTextOp
                {
                    Target = new TextSpanAnchor { ParaId = "w14:00000002", Expect = "Acme Corp" },
                    With = "Globex Inc."
                },
                new InsertTableRowsOp
                {
                    Target = new NodeAnchor { Kind = "table", Path = "table#0" },
                    Rows = new[] { new[] { "UK", "68" } }
                }
            }
        };

        var json = JsonSerializer.Serialize(plan, Options);
        var restored = JsonSerializer.Deserialize<DocumentPlan>(json, Options)!;

        Assert.Collection(restored.Operations,
            first => Assert.Equal("Globex Inc.", Assert.IsType<ChangeTextOp>(first).With),
            second => Assert.Equal("UK", Assert.IsType<InsertTableRowsOp>(second).Rows[0][0]));
    }

    [Fact]
    public void A_missing_or_unknown_verb_is_reported_with_the_ones_that_exist()
    {
        var missing = Assert.Throws<JsonException>(() =>
            JsonSerializer.Deserialize<PlanOperation>("""{ "with": "Globex Inc." }""", Options));
        var unknown = Assert.Throws<JsonException>(() =>
            JsonSerializer.Deserialize<PlanOperation>("""{ "op": "deleteEverything" }""", Options));

        // "must specify a type discriminator" tells an agent nothing it can act on.
        Assert.Contains("changeText", missing.Message);
        Assert.Contains("deleteEverything", unknown.Message);
        Assert.Contains("insertTableRows", unknown.Message);
    }

    [Fact]
    public void Every_operation_type_in_the_contract_is_reachable_from_a_verb()
    {
        var declared = typeof(PlanOperation).Assembly.GetTypes()
            .Where(t => t.IsSubclassOf(typeof(PlanOperation)) && !t.IsAbstract)
            .ToList();

        // Losing the [JsonDerivedType] attributes means nothing fails at compile time when
        // a verb is added, so the gap has to fail here instead.
        var unreachable = declared.Except(PlanOperationJsonConverter.ByVerb.Values).ToList();

        Assert.True(unreachable.Count == 0,
            $"Not in the verb map: {string.Join(", ", unreachable.Select(t => t.Name))}");
        Assert.NotEmpty(declared);
    }

    [Fact]
    public void The_verb_written_out_is_the_one_that_reads_back()
    {
        foreach (var (verb, type) in PlanOperationJsonConverter.ByVerb)
        {
            var instance = (PlanOperation)Activator.CreateInstance(type)!;
            using var document = JsonDocument.Parse(JsonSerializer.Serialize(instance, Options));

            Assert.Equal(verb, document.RootElement.GetProperty("op").GetString());
        }
    }
}
