using System.Diagnostics;
using System.IO;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;

namespace NewUserAutomation.App.Services;

internal sealed class AppUpdateService
{
    private const string GitHubOwner = "The-Emerald-Group";
    private const string GitHubRepo = "newuserautomation";
    private const string LatestReleaseApiUrl = $"https://api.github.com/repos/{GitHubOwner}/{GitHubRepo}/releases/latest";
    private const string InstallerAssetName = "NewUserAutomationInstaller.exe";

    private readonly HttpClient _httpClient = new();

    public async Task<UpdateCheckResult> CheckForUpdateAsync(string currentVersionRaw)
    {
        var currentVersion = TryParseVersion(currentVersionRaw);
        if (currentVersion is null)
        {
            return UpdateCheckResult.NoUpdate;
        }

        using var request = new HttpRequestMessage(HttpMethod.Get, LatestReleaseApiUrl);
        request.Headers.Add("User-Agent", "NewUserAutomationApp");
        request.Headers.Add("Accept", "application/vnd.github+json");
        using var response = await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead);
        response.EnsureSuccessStatusCode();

        await using var stream = await response.Content.ReadAsStreamAsync();
        var release = await JsonSerializer.DeserializeAsync<GitHubRelease>(stream)
            ?? throw new InvalidOperationException("Could not parse GitHub release metadata.");

        var latestVersion = TryParseVersion(release.TagName);
        if (latestVersion is null || latestVersion <= currentVersion)
        {
            return UpdateCheckResult.NoUpdate;
        }

        var installerAsset = release.Assets.FirstOrDefault(a =>
            string.Equals(a.Name, InstallerAssetName, StringComparison.OrdinalIgnoreCase));
        if (installerAsset is null || string.IsNullOrWhiteSpace(installerAsset.BrowserDownloadUrl))
        {
            return UpdateCheckResult.NoUpdate;
        }

        return new UpdateCheckResult(
            true,
            currentVersion.ToString(3),
            latestVersion.ToString(3),
            installerAsset.BrowserDownloadUrl);
    }

    public async Task<string> DownloadInstallerAsync(string installerDownloadUrl)
    {
        var tempPath = Path.Combine(Path.GetTempPath(), $"{Guid.NewGuid():N}-{InstallerAssetName}");
        using var response = await _httpClient.GetAsync(installerDownloadUrl, HttpCompletionOption.ResponseHeadersRead);
        response.EnsureSuccessStatusCode();
        await using var source = await response.Content.ReadAsStreamAsync();
        await using var destination = File.Create(tempPath);
        await source.CopyToAsync(destination);
        return tempPath;
    }

    public static bool TryLaunchInstaller(string installerPath, bool relaunchAfterInstall)
    {
        if (!File.Exists(installerPath))
        {
            return false;
        }

        var args = relaunchAfterInstall ? "--relaunch" : string.Empty;
        Process.Start(new ProcessStartInfo
        {
            FileName = installerPath,
            Arguments = args,
            UseShellExecute = true
        });
        return true;
    }

    private static Version? TryParseVersion(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw))
        {
            return null;
        }

        var cleaned = raw.Trim();
        if (cleaned.StartsWith("v", StringComparison.OrdinalIgnoreCase))
        {
            cleaned = cleaned[1..];
        }

        var match = Regex.Match(cleaned, @"^\d+(\.\d+){0,3}");
        if (!match.Success)
        {
            return null;
        }

        return Version.TryParse(match.Value, out var version) ? version : null;
    }

    private sealed record GitHubRelease(
        [property: JsonPropertyName("tag_name")] string? TagName,
        [property: JsonPropertyName("assets")] IReadOnlyList<GitHubAsset> Assets);

    private sealed record GitHubAsset(
        [property: JsonPropertyName("name")] string? Name,
        [property: JsonPropertyName("browser_download_url")] string? BrowserDownloadUrl);
}

internal sealed record UpdateCheckResult(
    bool IsUpdateAvailable,
    string CurrentVersion,
    string LatestVersion,
    string InstallerDownloadUrl)
{
    public static readonly UpdateCheckResult NoUpdate = new(false, string.Empty, string.Empty, string.Empty);
}
