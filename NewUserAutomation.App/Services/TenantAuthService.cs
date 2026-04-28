using System.Diagnostics;
using System.IO;
using System.Text;
using System.Text.Json;

namespace NewUserAutomation.App.Services;

public sealed class TenantAuthService
{
    public async Task<AuthResult> SignInAsync(IProgress<string>? progress = null, CancellationToken cancellationToken = default)
    {
        var resultFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-auth-{Guid.NewGuid():N}.json");

        var script = $$"""
        $ErrorActionPreference = 'Stop'
        if ($PSStyle) { $PSStyle.OutputRendering = 'PlainText' }
        Write-Output '__STEP__:Checking Microsoft Graph module'
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
            Write-Output '__STEP__:Installing Microsoft Graph module'
            try {
                Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue
            } catch {
            }
            Install-Module Microsoft.Graph -Scope CurrentUser -Force -AllowClobber
            if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
                throw 'Microsoft.Graph PowerShell module install failed. Please run: Install-Module Microsoft.Graph -Scope CurrentUser'
            }
        }
        Write-Output '__STEP__:Importing Microsoft Graph module'
        Import-Module Microsoft.Graph.Authentication
        Write-Output '__STEP__:Launching browser for tenant sign-in'
        try {
            Connect-MgGraph -Scopes 'User.Read','Organization.Read.All' -NoWelcome | Out-Null
        } catch {
            Write-Output '__STEP__:Browser sign-in unavailable, waiting for device code sign-in'
            Connect-MgGraph -Scopes 'User.Read','Organization.Read.All' -UseDeviceCode -NoWelcome | Out-Null
        }
        Write-Output '__STEP__:Validating tenant context'
        $ctx = Get-MgContext
        if (-not $ctx) { throw 'Graph context was not created.' }
        $payload = [PSCustomObject]@{
            TenantId = $ctx.TenantId
            Account = $ctx.Account
            Scopes = ($ctx.Scopes -join ', ')
        } | ConvertTo-Json -Compress
        Set-Content -Path '{{resultFile}}' -Value $payload -Encoding UTF8
        """;

        var stderrFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-auth-{Guid.NewGuid():N}.err.txt");
        var wrappedScript = $$"""
        try {
            {{script}}
            exit 0
        } catch {
            $_ | Out-File -FilePath '{{stderrFile}}' -Encoding utf8 -Force
            exit 1
        }
        """;
        var encodedScript = Convert.ToBase64String(Encoding.Unicode.GetBytes(wrappedScript));

        var startInfo = new ProcessStartInfo
        {
            FileName = "pwsh",
            Arguments = $"-NoProfile -ExecutionPolicy Bypass -EncodedCommand {encodedScript}",
            UseShellExecute = true,
            CreateNoWindow = false
        };

        progress?.Report("Opening PowerShell sign-in window...");
        using var process = Process.Start(startInfo) ?? throw new InvalidOperationException("Failed to start PowerShell.");
        progress?.Report("Launching browser sign-in (or device code fallback)...");
        await process.WaitForExitAsync(cancellationToken);

        if (process.ExitCode != 0)
        {
            var stdErr = File.Exists(stderrFile) ? await File.ReadAllTextAsync(stderrFile, cancellationToken) : string.Empty;
            var message = string.IsNullOrWhiteSpace(stdErr) ? "Sign-in failed in PowerShell window." : SimplifyPowerShellError(stdErr);
            return new AuthResult(false, string.Empty, string.Empty, string.Empty, message);
        }

        try
        {
            if (!File.Exists(resultFile))
            {
                return new AuthResult(false, string.Empty, string.Empty, string.Empty, "Sign-in completed but no auth payload was returned.");
            }

            var json = await File.ReadAllTextAsync(resultFile, cancellationToken);
            var doc = JsonDocument.Parse(json);
            var tenantId = doc.RootElement.GetProperty("TenantId").GetString() ?? string.Empty;
            var account = doc.RootElement.GetProperty("Account").GetString() ?? string.Empty;
            var scopes = doc.RootElement.GetProperty("Scopes").GetString() ?? string.Empty;
            return new AuthResult(true, tenantId, account, scopes, string.Empty);
        }
        catch (Exception ex)
        {
            return new AuthResult(false, string.Empty, string.Empty, string.Empty, $"Could not parse sign-in output: {ex.Message}");
        }
        finally
        {
            TryDelete(resultFile);
            TryDelete(stderrFile);
        }
    }

    private static string SimplifyPowerShellError(string rawError)
    {
        var text = rawError
            .Replace("_x000D__x000A_", "\n", StringComparison.OrdinalIgnoreCase)
            .Replace("_x001B_[31;1m", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("_x001B_[36;1m", string.Empty, StringComparison.OrdinalIgnoreCase)
            .Replace("_x001B_[0m", string.Empty, StringComparison.OrdinalIgnoreCase);

        if (text.Contains("Microsoft.Graph PowerShell module is not installed", StringComparison.OrdinalIgnoreCase))
        {
            return "Microsoft Graph PowerShell module is not installed. Run: Install-Module Microsoft.Graph -Scope CurrentUser";
        }
        if (text.Contains("module install failed", StringComparison.OrdinalIgnoreCase))
        {
            return "Automatic dependency install failed. Run: Install-Module Microsoft.Graph -Scope CurrentUser";
        }

        var lines = text.Split('\n', StringSplitOptions.RemoveEmptyEntries)
            .Select(line => line.Trim())
            .Where(line => !line.StartsWith("#< CLIXML", StringComparison.OrdinalIgnoreCase))
            .Where(line => !line.StartsWith("<Objs", StringComparison.OrdinalIgnoreCase))
            .Where(line => !line.StartsWith("<Obj", StringComparison.OrdinalIgnoreCase))
            .Where(line => !line.StartsWith("<S S=", StringComparison.OrdinalIgnoreCase))
            .ToList();

        return lines.Count == 0 ? "Sign-in failed." : string.Join(" ", lines);
    }

    private static void TryDelete(string path)
    {
        try
        {
            if (File.Exists(path))
            {
                File.Delete(path);
            }
        }
        catch
        {
        }
    }
}

public sealed record AuthResult(
    bool Success,
    string TenantId,
    string Account,
    string Scopes,
    string ErrorMessage);
