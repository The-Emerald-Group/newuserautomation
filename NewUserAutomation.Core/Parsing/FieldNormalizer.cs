namespace NewUserAutomation.Core.Parsing;

public static class FieldNormalizer
{
    public static string NormalizeScalar(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        return value
            .Replace('\u2018', '\'')
            .Replace('\u2019', '\'')
            .Replace('\u201C', '"')
            .Replace('\u201D', '"')
            .Replace('\u2013', '-')
            .Replace('\u2014', '-')
            .Trim();
    }

    public static List<string> NormalizeList(string? value)
    {
        var normalized = NormalizeScalar(value);
        if (string.IsNullOrWhiteSpace(normalized))
        {
            return [];
        }

        return normalized
            .Split([',', ';', '|', '\n'], StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries)
            .Select(item => item.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }
}
