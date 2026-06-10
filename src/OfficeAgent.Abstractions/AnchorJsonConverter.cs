using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeAgent.Abstractions;

/// <summary>
/// Tolerant <see cref="Anchor"/> JSON converter wired up by
/// <see cref="JsonConverterAttribute"/> on <see cref="Anchor"/>. LLMs reliably forget
/// the <c>$anchor</c> discriminator inside <c>target</c>; without help System.Text.Json
/// throws <c>The JSON payload for polymorphic ... must specify a type discriminator</c>
/// and the agent stalls. This converter:
/// <list type="number">
/// <item>Uses the <c>$anchor</c> discriminator on read when the model does provide it.</item>
/// <item>Otherwise infers the concrete <see cref="Anchor"/> type from the property names
///   that are present - e.g. <c>paraId</c>/<c>expect</c>/<c>occurrence</c> ⇒
///   <see cref="TextSpanAnchor"/>; <c>tag</c> ⇒ <see cref="StructuralAnchor"/>;
///   <c>kind</c>/<c>path</c> ⇒ <see cref="NodeAnchor"/>; <c>styleId</c> alone ⇒
///   <see cref="StyleAnchor"/>.
/// </item>
/// <item>On write, always emits a <c>$anchor</c> discriminator so the contract round-trips.</item>
/// </list>
/// </summary>
public sealed class AnchorJsonConverter : JsonConverter<Anchor>
{
    /// <inheritdoc/>
    public override Anchor? Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        using var doc = JsonDocument.ParseValue(ref reader);
        var root = doc.RootElement;

        if (root.ValueKind == JsonValueKind.Null) return null;
        if (root.ValueKind != JsonValueKind.Object)
            throw new JsonException("Anchor must be a JSON object.");

        // 1. Explicit discriminator wins.
        if (TryFindProperty(root, "$anchor", out var disc) && disc.ValueKind == JsonValueKind.String)
        {
            var concrete = disc.GetString() switch
            {
                "textSpan" => typeof(TextSpanAnchor),
                "structural" => typeof(StructuralAnchor),
                "node" => typeof(NodeAnchor),
                "style" => typeof(StyleAnchor),
                "cell" => typeof(CellAnchor),
                "shape" => typeof(ShapeAnchor),
                _ => null
            };
            if (concrete is not null)
                return (Anchor?)root.Deserialize(concrete, options);
        }

        // 2. Infer from property shape. Order matters: NodeAnchor's kind/path is also
        //    used by StructuralAnchor, so check the more-specific signals first.
        var names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var prop in root.EnumerateObject())
            names.Add(prop.Name);

        if (names.Contains("paraId") || names.Contains("expect") || names.Contains("occurrence"))
            return root.Deserialize<TextSpanAnchor>(options);

        if (names.Contains("tag"))
            return root.Deserialize<StructuralAnchor>(options);

        if (names.Contains("path") || (names.Contains("kind") && !names.Contains("styleId")))
            return root.Deserialize<NodeAnchor>(options);

        if (names.Contains("styleId"))
            return root.Deserialize<StyleAnchor>(options);

        if (names.Contains("sheet") || names.Contains("ref"))
            return root.Deserialize<CellAnchor>(options);

        if (names.Contains("slideId") || names.Contains("shapeId"))
            return root.Deserialize<ShapeAnchor>(options);

        throw new JsonException(
            "Cannot infer anchor type from JSON properties. Include a \"$anchor\" discriminator " +
            "(textSpan, structural, node, style, cell, or shape).");
    }

    /// <inheritdoc/>
    public override void Write(Utf8JsonWriter writer, Anchor value, JsonSerializerOptions options)
    {
        var discriminator = value switch
        {
            TextSpanAnchor => "textSpan",
            StructuralAnchor => "structural",
            NodeAnchor => "node",
            StyleAnchor => "style",
            CellAnchor => "cell",
            ShapeAnchor => "shape",
            _ => throw new JsonException($"Unknown Anchor subtype: {value.GetType().FullName}")
        };

        writer.WriteStartObject();
        writer.WriteString("$anchor", discriminator);
        using var doc = JsonSerializer.SerializeToDocument(value, value.GetType(), options);
        foreach (var prop in doc.RootElement.EnumerateObject())
            prop.WriteTo(writer);
        writer.WriteEndObject();
    }

    private static bool TryFindProperty(JsonElement obj, string name, out JsonElement value)
    {
        if (obj.TryGetProperty(name, out value)) return true;
        foreach (var prop in obj.EnumerateObject())
        {
            if (string.Equals(prop.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                value = prop.Value;
                return true;
            }
        }
        value = default;
        return false;
    }
}
