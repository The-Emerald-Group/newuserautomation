using System.IO;
using System.Text.Json;

namespace NewUserAutomation.App.Services;

public sealed class CustomerDirectoryStore
{
    private readonly string _root = ResolveDataRoot("settings");

    private string FilePath => Path.Combine(_root, "customers.json");

    private static string ResolveDataRoot(string subfolder)
    {
        var sharedRoot = Environment.GetEnvironmentVariable("NEWUSERAUTOMATION_DATA_ROOT")?.Trim();
        if (!string.IsNullOrWhiteSpace(sharedRoot))
        {
            return Path.Combine(sharedRoot, subfolder);
        }

        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "NewUserAutomation",
            subfolder);
    }

    public List<CustomerDirectoryEntry> LoadAll()
    {
        if (!File.Exists(FilePath))
        {
            return [];
        }

        var json = File.ReadAllText(FilePath);
        return JsonSerializer.Deserialize<List<CustomerDirectoryEntry>>(json) ?? [];
    }

    public void SaveAll(IEnumerable<CustomerDirectoryEntry> entries)
    {
        Directory.CreateDirectory(_root);
        var normalized = entries
            .GroupBy(e => e.Name.Trim(), StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .OrderBy(e => e.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();
        var json = JsonSerializer.Serialize(normalized, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(FilePath, json);
    }
}

public sealed record CustomerDirectoryEntry(
    string Name,
    string TenantDomain,
    string SiteUrl,
    string PnPClientId,
    string PnPThumbprint)
{
    public override string ToString() => Name;
}
