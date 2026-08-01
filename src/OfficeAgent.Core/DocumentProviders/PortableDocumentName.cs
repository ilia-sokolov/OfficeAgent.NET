namespace OfficeAgent.Core.DocumentProviders;

internal static class PortableDocumentName
{
    private static readonly char[] InvalidCharacters =
        { '<', '>', ':', '"', '|', '?', '*', '/', '\\', '\0' };

    internal static bool ContainsInvalidCharacter(string name) =>
        name.IndexOfAny(InvalidCharacters) >= 0 || name.Any(char.IsControl);

    internal static bool IsWindowsDeviceName(string name)
    {
        var dot = name.IndexOf('.');
        var stem = (dot < 0 ? name : name.Substring(0, dot)).TrimEnd(' ', '.');

        if (stem.Equals("CON", StringComparison.OrdinalIgnoreCase) ||
            stem.Equals("PRN", StringComparison.OrdinalIgnoreCase) ||
            stem.Equals("AUX", StringComparison.OrdinalIgnoreCase) ||
            stem.Equals("NUL", StringComparison.OrdinalIgnoreCase))
            return true;

        if (stem.Length != 4 ||
            (!stem.StartsWith("COM", StringComparison.OrdinalIgnoreCase) &&
             !stem.StartsWith("LPT", StringComparison.OrdinalIgnoreCase)))
            return false;

        return stem[3] is >= '1' and <= '9' or '\u00b9' or '\u00b2' or '\u00b3';
    }
}
