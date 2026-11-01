namespace OfficeAgent.Core;

/// <summary>What a revision author may be, wherever the name comes from.</summary>
internal static class RevisionAuthorRules
{
    /// <summary>The longest author name a plan or a host may set.</summary>
    public const int MaximumLength = 255;

    /// <summary>Whether a name has 1 to 255 visible characters and no control characters.</summary>
    public static bool IsValid(string? author) =>
        !string.IsNullOrWhiteSpace(author) &&
        author!.Length <= MaximumLength &&
        !author.Any(char.IsControl);

    /// <summary>The name when it is valid, otherwise <see langword="null"/>.</summary>
    public static string? Usable(string? author) => IsValid(author) ? author : null;
}
