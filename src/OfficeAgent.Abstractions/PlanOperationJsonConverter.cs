using System.Text.Json;
using System.Text.Json.Serialization;

namespace OfficeAgent.Abstractions;

/// <summary>
/// Tolerant <see cref="PlanOperation"/> converter wired up by
/// <see cref="JsonConverterAttribute"/> on <see cref="PlanOperation"/>.
/// <para>
/// System.Text.Json requires a polymorphic type discriminator to be the <em>first</em>
/// property of the object. A model writing a plan has no reason to order it that way, and
/// when it does not the built-in reader fails with <c>must specify a type discriminator</c>
/// - which reads as "you left <c>op</c> out" even though <c>op</c> is right there. The
/// agent then rewrites the payload it already had correct.
/// </para>
/// <para>
/// This converter buffers the object and looks <c>op</c> up wherever it appears, and names
/// the accepted verbs when it is genuinely missing or unrecognised.
/// </para>
/// </summary>
public sealed class PlanOperationJsonConverter : JsonConverter<PlanOperation>
{
    /// <summary>
    /// The wire vocabulary: verb to CLR type. Only verbs a registered module implements
    /// belong here, so an agent never sees one that always fails.
    /// </summary>
    public static readonly IReadOnlyDictionary<string, Type> ByVerb =
        new Dictionary<string, Type>(StringComparer.OrdinalIgnoreCase)
        {
            ["fill"] = typeof(FillOp),
            ["changeText"] = typeof(ChangeTextOp),
            ["insert"] = typeof(InsertOp),
            ["comment"] = typeof(CommentOp),
            ["format"] = typeof(FormatOp),
            ["setProperty"] = typeof(SetPropertyOp),
            ["revision"] = typeof(RevisionOp),
            ["insertTable"] = typeof(InsertTableOp),
            ["removeTable"] = typeof(RemoveTableOp),
            ["insertTableRows"] = typeof(InsertTableRowsOp),
            ["removeTableRows"] = typeof(RemoveTableRowsOp),
            ["insertTableColumns"] = typeof(InsertTableColumnsOp),
            ["removeTableColumns"] = typeof(RemoveTableColumnsOp),
            ["copyStyles"] = typeof(CopyStylesOp),
            ["clearStyles"] = typeof(ClearStylesOp),
            ["insertImage"] = typeof(InsertImageOp),
            ["removeImage"] = typeof(RemoveImageOp),
            ["insertSlide"] = typeof(InsertSlideOp),
            ["removeSlide"] = typeof(RemoveSlideOp),
            ["moveSlide"] = typeof(MoveSlideOp),
            ["duplicateSlide"] = typeof(DuplicateSlideOp),
            ["insertShape"] = typeof(InsertShapeOp),
            ["removeShape"] = typeof(RemoveShapeOp),
            ["section"] = typeof(SectionOp),
            ["headerFooter"] = typeof(HeaderFooterOp),
            ["insertMedia"] = typeof(InsertMediaOp),
            ["transition"] = typeof(TransitionOp),
            ["animate"] = typeof(AnimateOp),
            ["backgroundImage"] = typeof(BackgroundImageOp)
        };

    private static string KnownVerbs => string.Join(", ", ByVerb.Keys.OrderBy(v => v, StringComparer.Ordinal));

    /// <inheritdoc/>
    public override PlanOperation? Read(ref Utf8JsonReader reader, Type typeToConvert, JsonSerializerOptions options)
    {
        using var doc = JsonDocument.ParseValue(ref reader);
        var root = doc.RootElement;

        if (root.ValueKind == JsonValueKind.Null) return null;
        if (root.ValueKind != JsonValueKind.Object)
            throw new JsonException("A plan operation must be a JSON object.");

        if (!TryFindProperty(root, "op", out var discriminator) ||
            discriminator.ValueKind != JsonValueKind.String)
        {
            throw new JsonException(
                $"A plan operation needs an \"op\" property naming the verb. Expected one of: {KnownVerbs}.");
        }

        var verb = discriminator.GetString()!;
        if (!ByVerb.TryGetValue(verb, out var concrete))
        {
            throw new JsonException(
                $"Unknown plan operation \"{verb}\". Expected one of: {KnownVerbs}.");
        }

        // Deserializing the concrete type does not re-enter this converter: a
        // JsonConverter<PlanOperation> only converts PlanOperation itself.
        return (PlanOperation?)root.Deserialize(concrete, options);
    }

    /// <inheritdoc/>
    public override void Write(Utf8JsonWriter writer, PlanOperation value, JsonSerializerOptions options)
    {
        var type = value.GetType();
        var verb = ByVerb.FirstOrDefault(pair => pair.Value == type).Key
            ?? throw new JsonException($"Unknown plan operation subtype: {type.FullName}");

        // Written first so the output stays readable by the built-in polymorphic reader too.
        writer.WriteStartObject();
        writer.WriteString("op", verb);
        using var doc = JsonSerializer.SerializeToDocument(value, type, options);
        foreach (var property in doc.RootElement.EnumerateObject())
            property.WriteTo(writer);
        writer.WriteEndObject();
    }

    private static bool TryFindProperty(JsonElement obj, string name, out JsonElement value)
    {
        if (obj.TryGetProperty(name, out value)) return true;

        foreach (var property in obj.EnumerateObject())
        {
            if (string.Equals(property.Name, name, StringComparison.OrdinalIgnoreCase))
            {
                value = property.Value;
                return true;
            }
        }

        value = default;
        return false;
    }
}
