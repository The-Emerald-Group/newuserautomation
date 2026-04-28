using System.Diagnostics;
using System.IO.Compression;
using System.Net.Http;
using System.Security.Principal;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Windows.Forms;

internal static class Program
{
    [STAThread]
    private static void Main()
    {
        ApplicationConfiguration.Initialize();
        Application.Run(new InstallerForm());
    }
}

internal sealed class InstallerForm : Form
{
    private const string AppName = "NewUserAutomation";
    private const string AppExeName = "NewUserAutomation.App.exe";
    private const string GitHubOwner = "The-Emerald-Group";
    private const string GitHubRepo = "newuserautomation";
    private const string ZipAssetName = "NewUserAutomation-win-x64.zip";
    private const string LatestReleaseApiUrl = $"https://api.github.com/repos/{GitHubOwner}/{GitHubRepo}/releases/latest";

    private readonly ProgressBar _progress = new() { Minimum = 0, Maximum = 100, Height = 22, Dock = DockStyle.Bottom };
    private readonly Label _status = new() { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft, Padding = new Padding(12), AutoEllipsis = true };
    private readonly HttpClient _httpClient = new();

    public InstallerForm()
    {
        Text = "Installing NewUserAutomation";
        Width = 560;
        Height = 140;
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        Controls.Add(_status);
        Controls.Add(_progress);
        Shown += async (_, _) => await InstallAsync();
    }

    private async Task InstallAsync()
    {
        var installRoot = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), AppName);
        if (NeedsElevation(installRoot) && !IsAdministrator())
        {
            RestartElevated();
            Close();
            return;
        }

        var currentDir = Path.Combine(installRoot, "current");
        var tempRoot = Path.Combine(Path.GetTempPath(), $"{AppName}-install-{Guid.NewGuid():N}");
        var zipPath = Path.Combine(tempRoot, $"{AppName}.zip");
        var extractPath = Path.Combine(tempRoot, "extracted");

        try
        {
            UpdateStatus("Preparing workspace...", 5);
            Directory.CreateDirectory(tempRoot);
            Directory.CreateDirectory(extractPath);

            UpdateStatus("Checking latest release manifest...", 8);
            var package = await ResolvePackageAsync();
            var installedExe = Path.Combine(currentDir, AppExeName);
            var installedVersion = GetProductVersionSafe(installedExe);
            if (!string.IsNullOrWhiteSpace(package.Version)
                && string.Equals(installedVersion, package.Version, StringComparison.OrdinalIgnoreCase))
            {
                UpdateStatus("Already up to date.", 100);
                MessageBox.Show(
                    $"NewUserAutomation is already on the latest version.\n\nVersion: {installedVersion}",
                    "No Update Needed",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                Close();
                return;
            }

            await DownloadWithProgressAsync(package.ZipUrl, zipPath);

            UpdateStatus("Extracting package...", 65);
            ZipFile.ExtractToDirectory(zipPath, extractPath, overwriteFiles: true);

            UpdateStatus("Locating app executable...", 72);
            var exePathInExtract = Directory.GetFiles(extractPath, AppExeName, SearchOption.AllDirectories).FirstOrDefault();
            if (string.IsNullOrWhiteSpace(exePathInExtract))
            {
                throw new InvalidOperationException($"Could not find {AppExeName} in downloaded package.");
            }

            var sourceRoot = Path.GetDirectoryName(exePathInExtract)!;
            var stagedExe = Path.Combine(sourceRoot, AppExeName);

            if (File.Exists(installedExe) && IsSameProductVersion(installedExe, stagedExe))
            {
                UpdateStatus("Already up to date.", 100);
                MessageBox.Show(
                    $"NewUserAutomation is already on the latest version.\n\nVersion: {GetProductVersionSafe(stagedExe)}",
                    "No Update Needed",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                Close();
                return;
            }

            UpdateStatus("Stopping running app (if open)...", 78);
            foreach (var process in Process.GetProcessesByName(Path.GetFileNameWithoutExtension(AppExeName)))
            {
                try { process.Kill(entireProcessTree: true); } catch { }
            }

            UpdateStatus("Copying updated files...", 84);
            Directory.CreateDirectory(currentDir);
            var changedFiles = CopyChangedFiles(sourceRoot, currentDir);

            if (!File.Exists(installedExe))
            {
                throw new InvalidOperationException($"Install completed but executable not found at {installedExe}.");
            }

            UpdateStatus("Creating Start Menu shortcut...", 92);
            var startMenuDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), @"Microsoft\Windows\Start Menu\Programs");
            CreateShortcut(Path.Combine(startMenuDir, $"{AppName}.lnk"), installedExe, currentDir);

            UpdateStatus("Creating Desktop shortcut...", 96);
            var desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
            CreateShortcut(Path.Combine(desktop, $"{AppName}.lnk"), installedExe, currentDir);

            UpdateStatus("Installation complete.", 100);
            MessageBox.Show(
                $"NewUserAutomation has been installed/updated successfully.\n\nVersion: {GetProductVersionSafe(installedExe)}\nChanged files: {changedFiles}\n\nLocation:\n{installedExe}",
                "Install Complete",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            Close();
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Installation failed.\n\n{ex.Message}",
                "Install Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            Close();
        }
        finally
        {
            try
            {
                if (Directory.Exists(tempRoot))
                {
                    Directory.Delete(tempRoot, recursive: true);
                }
            }
            catch
            {
            }
        }
    }

    private async Task DownloadWithProgressAsync(string url, string destinationPath)
    {
        UpdateStatus("Downloading package...", 10);

        using var response = await _httpClient.GetAsync(url, HttpCompletionOption.ResponseHeadersRead);
        response.EnsureSuccessStatusCode();

        var totalBytes = response.Content.Headers.ContentLength;
        await using var stream = await response.Content.ReadAsStreamAsync();
        await using var file = File.Create(destinationPath);
        var buffer = new byte[81920];
        long readTotal = 0;
        int read;
        while ((read = await stream.ReadAsync(buffer)) > 0)
        {
            await file.WriteAsync(buffer.AsMemory(0, read));
            readTotal += read;

            if (totalBytes is > 0)
            {
                var pct = (int)Math.Clamp((readTotal * 100.0) / totalBytes.Value, 0, 100);
                var mapped = 10 + (int)Math.Round(pct * 0.50); // map download into 10-60
                UpdateStatus($"Downloading package... {readTotal / 1024d / 1024d:N1} MB / {totalBytes.Value / 1024d / 1024d:N1} MB", mapped);
            }
            else
            {
                UpdateStatus("Downloading package...", 35);
            }
        }
    }

    private async Task<ReleasePackage> ResolvePackageAsync()
    {
        using var request = new HttpRequestMessage(HttpMethod.Get, LatestReleaseApiUrl);
        request.Headers.Add("User-Agent", "NewUserAutomationInstaller");
        request.Headers.Add("Accept", "application/vnd.github+json");
        using var response = await _httpClient.SendAsync(request, HttpCompletionOption.ResponseHeadersRead);
        response.EnsureSuccessStatusCode();

        await using var stream = await response.Content.ReadAsStreamAsync();
        var release = await JsonSerializer.DeserializeAsync<GitHubRelease>(stream)
            ?? throw new InvalidOperationException("Could not parse GitHub release metadata.");
        var zipAsset = release.Assets.FirstOrDefault(a =>
            string.Equals(a.Name, ZipAssetName, StringComparison.OrdinalIgnoreCase));
        if (zipAsset is null || string.IsNullOrWhiteSpace(zipAsset.BrowserDownloadUrl))
        {
            throw new InvalidOperationException($"Latest release is missing required asset '{ZipAssetName}'.");
        }

        var version = release.TagName?.TrimStart('v', 'V') ?? string.Empty;
        return new ReleasePackage(zipAsset.BrowserDownloadUrl, version);
    }

    private void UpdateStatus(string text, int percent)
    {
        if (InvokeRequired)
        {
            Invoke(() => UpdateStatus(text, percent));
            return;
        }

        _status.Text = text;
        _progress.Value = Math.Clamp(percent, 0, 100);
    }

    private static int CopyChangedFiles(string sourceDir, string destinationDir)
    {
        var changed = 0;
        foreach (var directory in Directory.GetDirectories(sourceDir, "*", SearchOption.AllDirectories))
        {
            var targetDir = directory.Replace(sourceDir, destinationDir, StringComparison.OrdinalIgnoreCase);
            Directory.CreateDirectory(targetDir);
        }

        foreach (var file in Directory.GetFiles(sourceDir, "*", SearchOption.AllDirectories))
        {
            var targetFile = file.Replace(sourceDir, destinationDir, StringComparison.OrdinalIgnoreCase);
            Directory.CreateDirectory(Path.GetDirectoryName(targetFile)!);
            if (!NeedsCopy(file, targetFile))
            {
                continue;
            }

            File.Copy(file, targetFile, overwrite: true);
            changed++;
        }

        return changed;
    }

    private static bool NeedsCopy(string sourceFile, string destinationFile)
    {
        if (!File.Exists(destinationFile))
        {
            return true;
        }

        var srcInfo = new FileInfo(sourceFile);
        var dstInfo = new FileInfo(destinationFile);
        if (srcInfo.Length != dstInfo.Length)
        {
            return true;
        }

        // 2-second tolerance for zip timestamp granularity.
        var delta = (srcInfo.LastWriteTimeUtc - dstInfo.LastWriteTimeUtc).Duration();
        return delta > TimeSpan.FromSeconds(2);
    }

    private static bool IsSameProductVersion(string installedExe, string stagedExe)
    {
        var a = GetProductVersionSafe(installedExe);
        var b = GetProductVersionSafe(stagedExe);
        if (string.IsNullOrWhiteSpace(a) || string.IsNullOrWhiteSpace(b))
        {
            return false;
        }

        return string.Equals(a, b, StringComparison.OrdinalIgnoreCase);
    }

    private static string GetProductVersionSafe(string exePath)
    {
        try
        {
            return FileVersionInfo.GetVersionInfo(exePath).ProductVersion ?? string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static void CreateShortcut(string shortcutPath, string targetPath, string workingDirectory)
    {
        var shellType = Type.GetTypeFromProgID("WScript.Shell") ?? throw new InvalidOperationException("WScript.Shell COM type unavailable.");
        dynamic shell = Activator.CreateInstance(shellType)!;
        dynamic shortcut = shell.CreateShortcut(shortcutPath);
        shortcut.TargetPath = targetPath;
        shortcut.WorkingDirectory = workingDirectory;
        shortcut.IconLocation = targetPath + ",0";
        shortcut.Save();
    }

    private static bool IsAdministrator()
    {
        var identity = WindowsIdentity.GetCurrent();
        var principal = new WindowsPrincipal(identity);
        return principal.IsInRole(WindowsBuiltInRole.Administrator);
    }

    private static bool NeedsElevation(string targetRoot)
    {
        var fullTarget = Path.GetFullPath(targetRoot).TrimEnd('\\');
        var programFiles = Path.GetFullPath(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)).TrimEnd('\\');
        var programFilesX86 = Environment.GetEnvironmentVariable("ProgramFiles(x86)");
        var isUnderProgramFiles = fullTarget.StartsWith(programFiles, StringComparison.OrdinalIgnoreCase);
        if (!string.IsNullOrWhiteSpace(programFilesX86))
        {
            var normalizedX86 = Path.GetFullPath(programFilesX86).TrimEnd('\\');
            isUnderProgramFiles |= fullTarget.StartsWith(normalizedX86, StringComparison.OrdinalIgnoreCase);
        }

        return isUnderProgramFiles;
    }

    private static void RestartElevated()
    {
        var exePath = Environment.ProcessPath ?? throw new InvalidOperationException("Could not resolve launcher path.");
        var psi = new ProcessStartInfo
        {
            FileName = exePath,
            UseShellExecute = true,
            Verb = "runas"
        };
        Process.Start(psi);
    }
}

internal sealed record ReleasePackage(string ZipUrl, string Version);

internal sealed record GitHubRelease(
    [property: JsonPropertyName("tag_name")] string? TagName,
    [property: JsonPropertyName("assets")] IReadOnlyList<GitHubAsset> Assets);

internal sealed record GitHubAsset(
    [property: JsonPropertyName("name")] string? Name,
    [property: JsonPropertyName("browser_download_url")] string? BrowserDownloadUrl);
