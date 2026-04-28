using System.IO;
using System.Text.Json;

namespace NewUserAutomation.App.Services;

public sealed class CustomerProfileStore
{
    private readonly string _root = ResolveDataRoot("customers");

    public CustomerProfile? Load(string tenantDomain)
    {
        var path = GetProfilePath(tenantDomain);
        if (!File.Exists(path))
        {
            return null;
        }

        var json = File.ReadAllText(path);
        return JsonSerializer.Deserialize<CustomerProfile>(json);
    }

    public void Save(CustomerProfile profile)
    {
        var dir = GetCustomerDir(profile.TenantDomain);
        Directory.CreateDirectory(dir);
        var path = Path.Combine(dir, "profile.json");
        profile = profile with { UpdatedUtc = DateTimeOffset.UtcNow };
        var json = JsonSerializer.Serialize(profile, new JsonSerializerOptions { WriteIndented = true });
        File.WriteAllText(path, json);
    }

    private string GetProfilePath(string tenantDomain) => Path.Combine(GetCustomerDir(tenantDomain), "profile.json");

    private string GetCustomerDir(string tenantDomain)
    {
        var key = tenantDomain.Trim().ToLowerInvariant().Replace(".", "_");
        return Path.Combine(_root, key);
    }

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
}

public sealed record CustomerProfile(
    string TenantDomain,
    string SiteUrl,
    string ClientId,
    string Thumbprint,
    DateTimeOffset UpdatedUtc);
