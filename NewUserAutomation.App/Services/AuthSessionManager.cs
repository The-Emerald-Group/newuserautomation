using System.Diagnostics;
using System.IO;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Text.Json;
using Azure.Identity;
using Microsoft.Graph;
using NewUserAutomation.Core.Models;

namespace NewUserAutomation.App.Services;

public sealed class AuthSessionManager
{
    private static readonly string[] GraphScopes = ["https://graph.microsoft.com/.default"];
    private GraphServiceClient? _graphClient;
    private string _graphTenantId = string.Empty;
    private string _graphAccount = string.Empty;
    private Process? _exchangeHostProcess;
    private StreamWriter? _exchangeHostInput;
    private readonly SemaphoreSlim _exchangeHostLock = new(1, 1);
    private string _exchangeAppId = string.Empty;
    private string _exchangeOrganization = string.Empty;
    private string _exchangeThumbprint = string.Empty;
    public string GraphTenantId => _graphTenantId;
    public string GraphAccount => _graphAccount;

    public async Task<AuthResult> ConnectGraphAsync(
        string appId,
        string tenantDomain,
        string thumbprint,
        IProgress<string>? progress = null,
        CancellationToken cancellationToken = default)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(appId) || string.IsNullOrWhiteSpace(tenantDomain) || string.IsNullOrWhiteSpace(thumbprint))
            {
                return new AuthResult(false, string.Empty, string.Empty, string.Empty, "Graph app auth requires App ID, Tenant Domain, and certificate thumbprint.");
            }

            progress?.Report("Connecting Graph (app registration)...");
            var cert = LoadCertificateFromCurrentUserStore(thumbprint);
            var credential = new ClientCertificateCredential(tenantDomain, appId, cert);
            _ = await credential.GetTokenAsync(new Azure.Core.TokenRequestContext(GraphScopes), cancellationToken);
            var client = new GraphServiceClient(credential, GraphScopes);

            var tenantInfo = await client.Organization.GetAsync(config =>
            {
                config.QueryParameters.Select = ["id", "displayName"];
                config.QueryParameters.Top = 1;
            }, cancellationToken);

            var tenantId = tenantInfo?.Value?.FirstOrDefault()?.Id ?? string.Empty;
            var account = $"App: {appId}";
            var scopesLabel = "Graph app-only (.default)";

            _graphClient = client;
            _graphTenantId = tenantId;
            _graphAccount = account;

            return new AuthResult(true, tenantId, account, scopesLabel, string.Empty);
        }
        catch (Exception ex)
        {
            return new AuthResult(false, string.Empty, string.Empty, string.Empty, Simplify(ex.Message));
        }
    }

    public async Task<Dictionary<string, IReadOnlyList<DirectoryUserMatch>>> FindDirectoryUserMatchesAsync(
        IReadOnlyList<string> userIdentities,
        string? expectedTenantId = null,
        IProgress<string>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var result = new Dictionary<string, IReadOnlyList<DirectoryUserMatch>>(StringComparer.OrdinalIgnoreCase);
        if (userIdentities.Count == 0)
        {
            return result;
        }

        progress?.Report("Checking user identities in Entra ID...");
        if (_graphClient is null)
        {
            return result;
        }

        if (!string.IsNullOrWhiteSpace(expectedTenantId)
            && !string.IsNullOrWhiteSpace(_graphTenantId)
            && !string.Equals(expectedTenantId, _graphTenantId, StringComparison.OrdinalIgnoreCase))
        {
            return result;
        }

        foreach (var rawIdentity in userIdentities)
        {
            cancellationToken.ThrowIfCancellationRequested();
            var identity = rawIdentity.Trim();
            if (string.IsNullOrWhiteSpace(identity))
            {
                continue;
            }

            var matches = new List<DirectoryUserMatch>();
            var dedupe = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            await TryAddByUserId(identity, matches, dedupe, cancellationToken);
            if (matches.Count == 0)
            {
                if (identity.Contains('@', StringComparison.Ordinal))
                {
                    await AddFromFilter($"userPrincipalName eq '{EscapeGraphLiteral(identity)}' or mail eq '{EscapeGraphLiteral(identity)}'", 10, matches, dedupe, cancellationToken);
                    await AddFromSearch($"userPrincipalName:{identity}", 20, matches, dedupe, cancellationToken);
                    await AddFromSearch($"mail:{identity}", 20, matches, dedupe, cancellationToken);
                    await AddFromSearch($"proxyAddresses:SMTP:{identity}", 20, matches, dedupe, cancellationToken);
                    await AddFromSearch($"proxyAddresses:smtp:{identity}", 20, matches, dedupe, cancellationToken);
                }
                else
                {
                    await AddFromFilter($"startswith(displayName,'{EscapeGraphLiteral(identity)}') or displayName eq '{EscapeGraphLiteral(identity)}' or userPrincipalName eq '{EscapeGraphLiteral(identity)}' or mail eq '{EscapeGraphLiteral(identity)}'", 10, matches, dedupe, cancellationToken);
                    var first = identity.Split(' ', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault();
                    if (!string.IsNullOrWhiteSpace(first))
                    {
                        await AddFromFilter($"startswith(displayName,'{EscapeGraphLiteral(first)}')", 20, matches, dedupe, cancellationToken);
                        await AddFromSearch($"displayName:{first}", 20, matches, dedupe, cancellationToken);
                    }

                    if (matches.Count == 0)
                    {
                        await AddFromSearch($"displayName:{identity}", 20, matches, dedupe, cancellationToken);
                    }
                }
            }

            result[identity] = matches
                .GroupBy(x => string.IsNullOrWhiteSpace(x.UserPrincipalName) ? x.DisplayName : x.UserPrincipalName, StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .ToList();
        }

        return result;
    }

    public async Task<LiveExecutionResult> EnsureUserExistsAsync(NewUserRequest request, CancellationToken cancellationToken = default)
    {
        if (_graphClient is null)
        {
            return new LiveExecutionResult(false, "Graph client is not connected.");
        }

        try
        {
            var existing = await _graphClient.Users[request.Upn].GetAsync(config =>
            {
                config.QueryParameters.Select = ["id", "displayName", "userPrincipalName"];
            }, cancellationToken);
            if (existing is not null)
            {
                return new LiveExecutionResult(true, $"User already exists: {existing.UserPrincipalName ?? request.Upn}");
            }
        }
        catch
        {
        }

        try
        {
            var user = new Microsoft.Graph.Models.User
            {
                AccountEnabled = true,
                DisplayName = request.DisplayName,
                GivenName = request.FirstName,
                Surname = request.LastName,
                JobTitle = request.JobTitle,
                UserPrincipalName = request.Upn,
                MailNickname = BuildMailNickname(request.PreferredUsername, request.FirstName, request.LastName),
                UsageLocation = "GB",
                PasswordProfile = new Microsoft.Graph.Models.PasswordProfile
                {
                    Password = request.TemporaryPassword,
                    ForceChangePasswordNextSignIn = true
                }
            };
            var created = await _graphClient.Users.PostAsync(user, cancellationToken: cancellationToken);
            var createdUpn = created?.UserPrincipalName ?? request.Upn;
            return new LiveExecutionResult(true, $"Created user {createdUpn}.");
        }
        catch (Exception ex)
        {
            return new LiveExecutionResult(false, $"Create user failed: {Simplify(ex.Message)}");
        }
    }

    public async Task<LiveExecutionResult> EnsureSecurityGroupAsync(string groupName, CancellationToken cancellationToken = default)
    {
        if (_graphClient is null)
        {
            return new LiveExecutionResult(false, "Graph client is not connected.");
        }

        var existing = await FindGroupByDisplayNameAsync(groupName, cancellationToken);
        if (existing is not null)
        {
            return new LiveExecutionResult(true, $"Group already exists: {groupName}");
        }

        try
        {
            var group = new Microsoft.Graph.Models.Group
            {
                DisplayName = groupName,
                Description = $"Created by NewUserAutomation for {groupName}",
                MailEnabled = false,
                MailNickname = BuildMailNickname(groupName),
                SecurityEnabled = true
            };
            await _graphClient.Groups.PostAsync(group, cancellationToken: cancellationToken);
            return new LiveExecutionResult(true, $"Created security group: {groupName}");
        }
        catch (Exception ex)
        {
            return new LiveExecutionResult(false, $"Create security group failed: {Simplify(ex.Message)}");
        }
    }

    public async Task<LiveExecutionResult> AddUserToGroupByNameAsync(string userUpn, string groupName, CancellationToken cancellationToken = default)
    {
        if (_graphClient is null)
        {
            return new LiveExecutionResult(false, "Graph client is not connected.");
        }

        try
        {
            var user = await _graphClient.Users[userUpn].GetAsync(config =>
            {
                config.QueryParameters.Select = ["id", "userPrincipalName"];
            }, cancellationToken);
            if (string.IsNullOrWhiteSpace(user?.Id))
            {
                return new LiveExecutionResult(false, $"User not found in Graph: {userUpn}");
            }

            var group = await FindGroupByDisplayNameAsync(groupName, cancellationToken);
            if (group is null || string.IsNullOrWhiteSpace(group.Id))
            {
                return new LiveExecutionResult(false, $"Group not found: {groupName}");
            }

            var refBody = new Microsoft.Graph.Models.ReferenceCreate
            {
                OdataId = $"https://graph.microsoft.com/v1.0/directoryObjects/{user.Id}"
            };
            await _graphClient.Groups[group.Id].Members.Ref.PostAsync(refBody, cancellationToken: cancellationToken);
            return new LiveExecutionResult(true, $"Added {userUpn} to {groupName}");
        }
        catch (Exception ex)
        {
            var message = Simplify(ex.Message);
            if (message.Contains("added object references already exist", StringComparison.OrdinalIgnoreCase) ||
                message.Contains("One or more added object references already exist", StringComparison.OrdinalIgnoreCase))
            {
                return new LiveExecutionResult(true, $"Membership already present: {userUpn} in {groupName}");
            }

            return new LiveExecutionResult(false, $"Add to group failed: {message}");
        }
    }

    public async Task<LiveExecutionResult> AssignLicenseBySkuPartNumberAsync(string userUpn, string skuPartNumber, CancellationToken cancellationToken = default)
    {
        if (_graphClient is null)
        {
            return new LiveExecutionResult(false, "Graph client is not connected.");
        }

        try
        {
            var user = await _graphClient.Users[userUpn].GetAsync(config =>
            {
                config.QueryParameters.Select = ["id", "userPrincipalName"];
            }, cancellationToken);
            if (string.IsNullOrWhiteSpace(user?.Id))
            {
                return new LiveExecutionResult(false, $"User not found in Graph: {userUpn}");
            }

            var skus = (await _graphClient.SubscribedSkus.GetAsync(cancellationToken: cancellationToken))?.Value ?? [];
            var targetSku = ResolveSkuByAliases(skus, GetSkuAliases(skuPartNumber));
            var substitutedFrom = string.Empty;
            if (targetSku?.SkuId is null && string.Equals(NormalizeSkuKey(skuPartNumber), "BUSINESS_STANDARD", StringComparison.Ordinal))
            {
                var premiumFallback = ResolveSkuByAliases(skus, GetSkuAliases("BUSINESS_PREMIUM"));
                if (premiumFallback?.SkuId is not null)
                {
                    targetSku = premiumFallback;
                    substitutedFrom = skuPartNumber;
                }
            }
            if (targetSku?.SkuId is null)
            {
                var available = string.Join(", ", skus
                    .Select(s => s.SkuPartNumber)
                    .Where(x => !string.IsNullOrWhiteSpace(x))
                    .OrderBy(x => x, StringComparer.OrdinalIgnoreCase));
                return new LiveExecutionResult(false, $"License SKU not found in tenant: {skuPartNumber}. Available SKUs: {available}");
            }

            var details = await _graphClient.Users[user.Id].LicenseDetails.GetAsync(cancellationToken: cancellationToken);
            var alreadyAssigned = details?.Value?.Any(d => d.SkuId == targetSku.SkuId) == true;
            if (alreadyAssigned)
            {
                var assignedSku = targetSku.SkuPartNumber ?? skuPartNumber;
                if (!string.IsNullOrWhiteSpace(substitutedFrom))
                {
                    return new LiveExecutionResult(true, $"License already assigned: {assignedSku} (substituted from {substitutedFrom}).");
                }

                return new LiveExecutionResult(true, $"License already assigned: {assignedSku}");
            }

            var requestBody = new Microsoft.Graph.Users.Item.AssignLicense.AssignLicensePostRequestBody
            {
                AddLicenses =
                [
                    new Microsoft.Graph.Models.AssignedLicense { SkuId = targetSku.SkuId }
                ],
                RemoveLicenses = []
            };
            await _graphClient.Users[user.Id].AssignLicense.PostAsync(requestBody, cancellationToken: cancellationToken);
            var appliedSku = targetSku.SkuPartNumber ?? skuPartNumber;
            if (!string.IsNullOrWhiteSpace(substitutedFrom))
            {
                return new LiveExecutionResult(true, $"Assigned license {appliedSku} (substituted from {substitutedFrom}).");
            }

            return new LiveExecutionResult(true, $"Assigned license {appliedSku}");
        }
        catch (Exception ex)
        {
            return new LiveExecutionResult(false, $"Assign license failed: {Simplify(ex.Message)}");
        }
    }

    private static string NormalizeSkuKey(string value)
        => value.Trim().Replace("-", "_", StringComparison.Ordinal).ToUpperInvariant();

    private static IReadOnlyList<string> GetSkuAliases(string requested)
    {
        var normalized = NormalizeSkuKey(requested);
        return normalized switch
        {
            "BUSINESS_PREMIUM" => ["BUSINESS_PREMIUM", "SPB"],
            "BUSINESS_STANDARD" => ["BUSINESS_STANDARD", "O365_BUSINESS_PREMIUM"],
            "BUSINESS_BASIC" => ["BUSINESS_BASIC", "O365_BUSINESS_ESSENTIALS"],
            "EXCHANGE_ONLINE_MAILBOX" => ["EXCHANGE_ONLINE_MAILBOX", "EXCHANGESTANDARD", "STANDARDPACK_EXCHANGE"],
            _ => [normalized]
        };
    }

    private static Microsoft.Graph.Models.SubscribedSku? ResolveSkuByAliases(
        IEnumerable<Microsoft.Graph.Models.SubscribedSku> skus,
        IReadOnlyList<string> aliases)
    {
        var aliasSet = new HashSet<string>(aliases.Select(NormalizeSkuKey), StringComparer.OrdinalIgnoreCase);
        return skus.FirstOrDefault(s =>
            !string.IsNullOrWhiteSpace(s.SkuPartNumber)
            && aliasSet.Contains(NormalizeSkuKey(s.SkuPartNumber)));
    }

    public async Task<LiveExecutionResult> GrantExchangeTargetAccessAsync(string userUpn, string target, CancellationToken cancellationToken = default)
    {
        var resultFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-exec-{Guid.NewGuid():N}.json");
        var stderrFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-exec-{Guid.NewGuid():N}.err.txt");
        var safeUser = userUpn.Replace("'", "''", StringComparison.Ordinal);
        var safeTarget = target.Replace("'", "''", StringComparison.Ordinal);
        var script = $$"""
        $ErrorActionPreference = 'Stop'
        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue } catch {}
            Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
        }
        Import-Module ExchangeOnlineManagement
        $conn = $null
        try { $conn = Get-ConnectionInformation | Select-Object -First 1 } catch { $conn = $null }
        if (-not $conn) {
            Connect-ExchangeOnline -AppId '{{EscapePsLiteral(_exchangeAppId)}}' -Organization '{{EscapePsLiteral(_exchangeOrganization)}}' -CertificateThumbprint '{{EscapePsLiteral(_exchangeThumbprint)}}' -ShowBanner:$false -SkipLoadingFormatData | Out-Null
            $conn = Get-ConnectionInformation | Select-Object -First 1
        }
        if (-not $conn) {
            throw "No active Exchange session in this process. Reconnect Exchange from the Sign In page."
        }
        $user = '{{safeUser}}'
        $target = '{{safeTarget}}'
        $applied = $false
        $detail = ''
        try {
            $mb = Get-EXOMailbox -Identity $target -ResultSize 1 -ErrorAction Stop
            if ($mb) {
                try {
                    Add-MailboxPermission -Identity $target -User $user -AccessRights FullAccess -InheritanceType All -AutoMapping:$false -ErrorAction Stop | Out-Null
                    $applied = $true
                    $detail = "Applied mailbox permission for '$target'."
                } catch {
                    $msg = $_.Exception.Message
                    if ($msg -match 'already' -or $msg -match 'exists') {
                        $applied = $true
                        $detail = "Mailbox permission already present for '$target'."
                    }
                }
            }
        } catch {
        }
        if (-not $applied) {
            try {
                Add-DistributionGroupMember -Identity $target -Member $user -ErrorAction Stop
                $applied = $true
                $detail = "Added member to distribution group '$target'."
            } catch {
                $msg = $_.Exception.Message
                if ($msg -match 'already' -or $msg -match 'exists') {
                    $applied = $true
                    $detail = "Distribution group membership already present for '$target'."
                }
            }
        }
        if (-not $applied) {
            throw "Could not apply Exchange access for target '$target'."
        }
        if ([string]::IsNullOrWhiteSpace($detail)) { $detail = "Applied Exchange access for '$target'." }
        [PSCustomObject]@{ Success = $true; Detail = $detail } | ConvertTo-Json -Compress | Set-Content -Path '{{resultFile}}' -Encoding UTF8
        """;
        return await RunJsonScriptAsync(script, resultFile, stderrFile, cancellationToken);
    }

    public async Task<LiveExecutionResult> GrantExchangeTargetsAccessAsync(string userUpn, IReadOnlyList<string> targets, CancellationToken cancellationToken = default)
    {
        if (targets.Count == 0)
        {
            return new LiveExecutionResult(true, "No Exchange targets to process.");
        }

        var resultFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-exec-{Guid.NewGuid():N}.json");
        var stderrFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-exec-{Guid.NewGuid():N}.err.txt");
        var safeUser = userUpn.Replace("'", "''", StringComparison.Ordinal);
        var joinedTargets = string.Join("|", targets).Replace("'", "''", StringComparison.Ordinal);
        var script = $$"""
        $ErrorActionPreference = 'Stop'
        $conn = $null
        try { $conn = Get-ConnectionInformation | Select-Object -First 1 } catch { $conn = $null }
        if (-not $conn) {
            Connect-ExchangeOnline -AppId '{{EscapePsLiteral(_exchangeAppId)}}' -Organization '{{EscapePsLiteral(_exchangeOrganization)}}' -CertificateThumbprint '{{EscapePsLiteral(_exchangeThumbprint)}}' -ShowBanner:$false -SkipLoadingFormatData | Out-Null
            $conn = Get-ConnectionInformation | Select-Object -First 1
        }
        if (-not $conn) {
            throw "No active Exchange session in persistent host. Reconnect Exchange from the Sign In page."
        }
        $user = '{{safeUser}}'
        $targets = '{{joinedTargets}}' -split '\|'
        $applied = 0
        $already = 0
        $failed = @()
        foreach ($target in $targets) {
            if ([string]::IsNullOrWhiteSpace($target)) { continue }
            $done = $false
            try {
                $mb = Get-EXOMailbox -Identity $target -ResultSize 1 -ErrorAction Stop
                if ($mb) {
                    try {
                        Add-MailboxPermission -Identity $target -User $user -AccessRights FullAccess -InheritanceType All -AutoMapping:$false -ErrorAction Stop | Out-Null
                        $applied++
                        $done = $true
                    } catch {
                        $msg = $_.Exception.Message
                        if ($msg -match 'already' -or $msg -match 'exists') {
                            $already++
                            $done = $true
                        }
                    }
                }
            } catch {
            }
            if (-not $done) {
                try {
                    Add-DistributionGroupMember -Identity $target -Member $user -ErrorAction Stop
                    $applied++
                    $done = $true
                } catch {
                    $msg = $_.Exception.Message
                    if ($msg -match 'already' -or $msg -match 'exists') {
                        $already++
                        $done = $true
                    }
                }
            }
            if (-not $done) {
                $failed += $target
            }
        }
        if ($failed.Count -gt 0) {
            throw "Exchange targets failed: $($failed -join ', ')"
        }
        [PSCustomObject]@{
            Success = $true
            Detail = "Exchange access processed for $($targets.Count) target(s): $applied applied, $already already present."
        } | ConvertTo-Json -Compress | Set-Content -Path '{{resultFile}}' -Encoding UTF8
        """;
        return await RunJsonScriptAsync(script, resultFile, stderrFile, cancellationToken);
    }

    public async Task<LiveExecutionResult> AddMailboxAliasAsync(string mailboxIdentity, string aliasEmail, CancellationToken cancellationToken = default)
    {
        var resultFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-exec-{Guid.NewGuid():N}.json");
        var stderrFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-exec-{Guid.NewGuid():N}.err.txt");
        var safeMailbox = mailboxIdentity.Replace("'", "''", StringComparison.Ordinal);
        var safeAlias = aliasEmail.Replace("'", "''", StringComparison.Ordinal);
        var script = $$"""
        $ErrorActionPreference = 'Stop'
        $conn = $null
        try { $conn = Get-ConnectionInformation | Select-Object -First 1 } catch { $conn = $null }
        if (-not $conn) {
            Connect-ExchangeOnline -AppId '{{EscapePsLiteral(_exchangeAppId)}}' -Organization '{{EscapePsLiteral(_exchangeOrganization)}}' -CertificateThumbprint '{{EscapePsLiteral(_exchangeThumbprint)}}' -ShowBanner:$false -SkipLoadingFormatData | Out-Null
            $conn = Get-ConnectionInformation | Select-Object -First 1
        }
        if (-not $conn) {
            throw "No active Exchange session in persistent host. Reconnect Exchange from the Sign In page."
        }
        $mailbox = '{{safeMailbox}}'
        $alias = '{{safeAlias}}'
        Set-Mailbox -Identity $mailbox -EmailAddresses @{Add="smtp:$alias"} -ErrorAction Stop | Out-Null
        [PSCustomObject]@{ Success = $true; Detail = "Alias '$alias' added to '$mailbox'." } | ConvertTo-Json -Compress | Set-Content -Path '{{resultFile}}' -Encoding UTF8
        """;
        return await RunJsonScriptAsync(script, resultFile, stderrFile, cancellationToken);
    }

    public async Task<LiveExecutionResult> EnsureSecondaryMailboxUserAsync(NewUserRequest primaryRequest, string secondaryEmail, CancellationToken cancellationToken = default)
    {
        if (_graphClient is null)
        {
            return new LiveExecutionResult(false, "Graph client is not connected.");
        }

        try
        {
            var existing = await _graphClient.Users[secondaryEmail].GetAsync(config =>
            {
                config.QueryParameters.Select = ["id", "userPrincipalName"];
            }, cancellationToken);
            if (existing is not null)
            {
                return new LiveExecutionResult(true, $"Secondary mailbox user already exists: {existing.UserPrincipalName ?? secondaryEmail}");
            }
        }
        catch
        {
        }

        try
        {
            var display = string.IsNullOrWhiteSpace(primaryRequest.DisplayName)
                ? secondaryEmail
                : $"{primaryRequest.DisplayName} (Secondary Mailbox)";
            var fallbackPassword = string.IsNullOrWhiteSpace(primaryRequest.TemporaryPassword)
                ? $"Temp!{Guid.NewGuid():N}".Substring(0, 14)
                : primaryRequest.TemporaryPassword;
            var mailboxUser = new Microsoft.Graph.Models.User
            {
                AccountEnabled = true,
                DisplayName = display,
                GivenName = primaryRequest.FirstName,
                Surname = primaryRequest.LastName,
                UserPrincipalName = secondaryEmail,
                MailNickname = BuildMailNickname(secondaryEmail),
                UsageLocation = "GB",
                PasswordProfile = new Microsoft.Graph.Models.PasswordProfile
                {
                    Password = fallbackPassword,
                    ForceChangePasswordNextSignIn = false
                }
            };
            var created = await _graphClient.Users.PostAsync(mailboxUser, cancellationToken: cancellationToken);
            return new LiveExecutionResult(true, $"Created secondary mailbox user {created?.UserPrincipalName ?? secondaryEmail}.");
        }
        catch (Exception ex)
        {
            return new LiveExecutionResult(false, $"Create secondary mailbox user failed: {Simplify(ex.Message)}");
        }
    }

    public async Task<LiveExecutionResult> GrantMailboxDelegateAccessAsync(string delegateUserUpn, string mailboxIdentity, CancellationToken cancellationToken = default)
    {
        var resultFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-exec-{Guid.NewGuid():N}.json");
        var stderrFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-exec-{Guid.NewGuid():N}.err.txt");
        var safeDelegate = delegateUserUpn.Replace("'", "''", StringComparison.Ordinal);
        var safeMailbox = mailboxIdentity.Replace("'", "''", StringComparison.Ordinal);
        var script = $$"""
        $ErrorActionPreference = 'Stop'
        $conn = $null
        try { $conn = Get-ConnectionInformation | Select-Object -First 1 } catch { $conn = $null }
        if (-not $conn) {
            Connect-ExchangeOnline -AppId '{{EscapePsLiteral(_exchangeAppId)}}' -Organization '{{EscapePsLiteral(_exchangeOrganization)}}' -CertificateThumbprint '{{EscapePsLiteral(_exchangeThumbprint)}}' -ShowBanner:$false -SkipLoadingFormatData | Out-Null
            $conn = Get-ConnectionInformation | Select-Object -First 1
        }
        if (-not $conn) {
            throw "No active Exchange session in persistent host. Reconnect Exchange from the Sign In page."
        }
        $delegateUser = '{{safeDelegate}}'
        $mailbox = '{{safeMailbox}}'
        try {
            Add-MailboxPermission -Identity $mailbox -User $delegateUser -AccessRights FullAccess -InheritanceType All -AutoMapping:$false -ErrorAction Stop | Out-Null
        } catch {
            $msg = $_.Exception.Message
            if ($msg -notmatch 'already' -and $msg -notmatch 'exists') { throw }
        }
        try {
            Add-RecipientPermission -Identity $mailbox -Trustee $delegateUser -AccessRights SendAs -Confirm:$false -ErrorAction Stop | Out-Null
        } catch {
            $msg = $_.Exception.Message
            if ($msg -notmatch 'already' -and $msg -notmatch 'exists') { throw }
        }
        [PSCustomObject]@{ Success = $true; Detail = "Granted FullAccess + SendAs on '$mailbox' to '$delegateUser'." } | ConvertTo-Json -Compress | Set-Content -Path '{{resultFile}}' -Encoding UTF8
        """;
        return await RunJsonScriptAsync(script, resultFile, stderrFile, cancellationToken);
    }

    public async Task<LiveExecutionResult> EnsurePersonalFolderAtSiteRootAsync(
        string siteUrl,
        string folderName,
        string clientId,
        string tenantDomain,
        string thumbprint,
        CancellationToken cancellationToken = default)
    {
        var resultFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-exec-{Guid.NewGuid():N}.json");
        var stderrFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-exec-{Guid.NewGuid():N}.err.txt");
        var safeUrl = siteUrl.Replace("'", "''", StringComparison.Ordinal);
        var safeFolder = folderName.Replace("'", "''", StringComparison.Ordinal);
        var safeClientId = clientId.Replace("'", "''", StringComparison.Ordinal);
        var safeTenant = tenantDomain.Replace("'", "''", StringComparison.Ordinal);
        var safeThumb = thumbprint.Replace("'", "''", StringComparison.Ordinal);
        var script = $$"""
        $ErrorActionPreference = 'Stop'
        if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
            try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue } catch {}
            Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
        }
        Import-Module PnP.PowerShell
        Connect-PnPOnline -Url '{{safeUrl}}' -ClientId '{{safeClientId}}' -Tenant '{{safeTenant}}' -Thumbprint '{{safeThumb}}'
        $libraryName = '{{safeFolder}}'
        $existing = $null
        try { $existing = Get-PnPList -Identity $libraryName -Includes Title,BaseTemplate -ErrorAction Stop } catch { $existing = $null }
        if ($existing) {
            [PSCustomObject]@{ Success = $true; Detail = "Document library '$libraryName' already exists." } | ConvertTo-Json -Compress | Set-Content -Path '{{resultFile}}' -Encoding UTF8
            return
        }

        try {
            New-PnPList -Title $libraryName -Template DocumentLibrary -OnQuickLaunch:$false -ErrorAction Stop | Out-Null
            [PSCustomObject]@{ Success = $true; Detail = "Created document library '$libraryName'." } | ConvertTo-Json -Compress | Set-Content -Path '{{resultFile}}' -Encoding UTF8
        } catch {
            $msg = $_.Exception.Message
            # If creation is denied, allow manual pre-creation and retry.
            if ($msg -match 'Access denied') {
                $existingAfterDenied = $null
                try { $existingAfterDenied = Get-PnPList -Identity $libraryName -Includes Title -ErrorAction Stop } catch { $existingAfterDenied = $null }
                if ($existingAfterDenied) {
                    [PSCustomObject]@{
                        Success = $true
                        Detail = "Document library '$libraryName' already exists."
                    } | ConvertTo-Json -Compress | Set-Content -Path '{{resultFile}}' -Encoding UTF8
                } else {
                    throw "Access denied creating document library '$libraryName'. Create the library manually, then click Run Next Live Step to continue."
                }
            } else {
                throw
            }
        }
        """;
        var createResult = await RunJsonScriptAsync(script, resultFile, stderrFile, cancellationToken);
        if (createResult.Success || !createResult.Detail.Contains("Access denied", StringComparison.OrdinalIgnoreCase))
        {
            return createResult;
        }

        // Retry once after ensuring app has explicit FullControl on the target site.
        var grantResult = await EnsureSharePointSiteFullControlForAppAsync(clientId, siteUrl, tenantDomain, cancellationToken: cancellationToken);
        if (!grantResult.Success)
        {
            return new LiveExecutionResult(false, $"{createResult.Detail} | Site permission grant failed: {grantResult.Detail}");
        }

        var retryResult = await RunJsonScriptAsync(script, resultFile, stderrFile, cancellationToken);
        if (retryResult.Success)
        {
            return new LiveExecutionResult(true, $"{retryResult.Detail} (After site FullControl grant)");
        }

        return new LiveExecutionResult(false, $"{retryResult.Detail} | Site grant attempted: {grantResult.Detail}");
    }

    public async Task<LiveExecutionResult> ApplyPersonalFolderPermissionsAsync(
        string siteUrl,
        string folderName,
        string permissionGroupName,
        string clientId,
        string tenantDomain,
        string thumbprint,
        CancellationToken cancellationToken = default)
    {
        var resultFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-exec-{Guid.NewGuid():N}.json");
        var stderrFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-exec-{Guid.NewGuid():N}.err.txt");
        var safeUrl = siteUrl.Replace("'", "''", StringComparison.Ordinal);
        var safeFolder = folderName.Replace("'", "''", StringComparison.Ordinal);
        var safeGroup = permissionGroupName.Replace("'", "''", StringComparison.Ordinal);
        var safeClientId = clientId.Replace("'", "''", StringComparison.Ordinal);
        var safeTenant = tenantDomain.Replace("'", "''", StringComparison.Ordinal);
        var safeThumb = thumbprint.Replace("'", "''", StringComparison.Ordinal);
        var script = $$"""
        $ErrorActionPreference = 'Stop'
        if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
            try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue } catch {}
            Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
        }
        Import-Module PnP.PowerShell
        Connect-PnPOnline -Url '{{safeUrl}}' -ClientId '{{safeClientId}}' -Tenant '{{safeTenant}}' -Thumbprint '{{safeThumb}}'

        $libraryName = '{{safeFolder}}'
        $principal = '{{safeGroup}}'
        $targetList = $null
        try { $targetList = Get-PnPList -Identity $libraryName -Includes Title -ErrorAction Stop } catch { $targetList = $null }
        if (-not $targetList) {
            throw "Document library '$libraryName' was not found."
        }

        Set-PnPList -Identity $libraryName -BreakRoleInheritance -CopyRoleAssignments:$false -ClearSubScopes:$true -ErrorAction Stop
        try {
            Set-PnPListPermission -Identity $libraryName -User $principal -AddRole "Edit" -ErrorAction Stop
        } catch {
            Set-PnPListPermission -Identity $libraryName -Group $principal -AddRole "Edit" -ErrorAction Stop
        }

        [PSCustomObject]@{
            Success = $true
            Detail = "Applied library permissions: '$principal' -> Edit on '$libraryName'."
        } | ConvertTo-Json -Compress | Set-Content -Path '{{resultFile}}' -Encoding UTF8
        """;

        return await RunJsonScriptAsync(script, resultFile, stderrFile, cancellationToken);
    }

    private async Task TryAddByUserId(string identity, List<DirectoryUserMatch> matches, HashSet<string> dedupe, CancellationToken cancellationToken)
    {
        if (_graphClient is null)
        {
            return;
        }

        try
        {
            var user = await _graphClient.Users[identity].GetAsync(config =>
            {
                config.QueryParameters.Select = ["id", "displayName", "userPrincipalName", "mail"];
            }, cancellationToken);

            if (user is not null)
            {
                AddCandidate(user, matches, dedupe);
            }
        }
        catch
        {
        }
    }

    private async Task AddFromFilter(string filter, int top, List<DirectoryUserMatch> matches, HashSet<string> dedupe, CancellationToken cancellationToken)
    {
        if (_graphClient is null)
        {
            return;
        }

        try
        {
            var users = await _graphClient.Users.GetAsync(config =>
            {
                config.QueryParameters.Filter = filter;
                config.QueryParameters.Top = top;
                config.QueryParameters.Select = ["id", "displayName", "userPrincipalName", "mail"];
            }, cancellationToken);

            if (users?.Value is null)
            {
                return;
            }

            foreach (var user in users.Value)
            {
                AddCandidate(user, matches, dedupe);
            }
        }
        catch
        {
        }
    }

    private async Task AddFromSearch(string searchExpression, int top, List<DirectoryUserMatch> matches, HashSet<string> dedupe, CancellationToken cancellationToken)
    {
        if (_graphClient is null)
        {
            return;
        }

        try
        {
            var users = await _graphClient.Users.GetAsync(config =>
            {
                config.Headers.Add("ConsistencyLevel", "eventual");
                config.QueryParameters.Search = $"\"{searchExpression}\"";
                config.QueryParameters.Top = top;
                config.QueryParameters.Select = ["id", "displayName", "userPrincipalName", "mail"];
            }, cancellationToken);

            if (users?.Value is null)
            {
                return;
            }

            foreach (var user in users.Value)
            {
                AddCandidate(user, matches, dedupe);
            }
        }
        catch
        {
        }
    }

    private static void AddCandidate(Microsoft.Graph.Models.User user, List<DirectoryUserMatch> matches, HashSet<string> dedupe)
    {
        var id = user.Id ?? string.Empty;
        var displayName = user.DisplayName ?? string.Empty;
        var upn = user.UserPrincipalName ?? string.Empty;
        var mail = user.Mail ?? string.Empty;
        if (string.IsNullOrWhiteSpace(upn) && string.IsNullOrWhiteSpace(displayName))
        {
            return;
        }

        var key = FirstNonEmpty(id, upn, mail, displayName);
        if (dedupe.Add(key))
        {
            matches.Add(new DirectoryUserMatch(id, displayName, upn, mail));
        }
    }

    private static string EscapeGraphLiteral(string value)
        => value.Replace("'", "''", StringComparison.Ordinal);

    private static string EscapePsLiteral(string value)
        => value.Replace("'", "''", StringComparison.Ordinal);

    private static X509Certificate2 LoadCertificateFromCurrentUserStore(string thumbprint)
    {
        var normalized = thumbprint.Replace(" ", string.Empty, StringComparison.Ordinal).Trim();
        using var store = new X509Store(StoreName.My, StoreLocation.CurrentUser);
        store.Open(OpenFlags.ReadOnly);
        var cert = store.Certificates
            .Find(X509FindType.FindByThumbprint, normalized, validOnly: false)
            .OfType<X509Certificate2>()
            .FirstOrDefault(c => c.HasPrivateKey);
        if (cert is null)
        {
            throw new InvalidOperationException($"Certificate with thumbprint '{thumbprint}' was not found in CurrentUser\\My or has no private key.");
        }
        return cert;
    }

    private static string FirstNonEmpty(params string?[] values)
        => values.FirstOrDefault(v => !string.IsNullOrWhiteSpace(v))?.Trim() ?? string.Empty;

    public async Task<AuthResult> ConnectExchangeAsync(
        string? appId = null,
        string? organization = null,
        string? thumbprint = null,
        IProgress<string>? progress = null,
        CancellationToken cancellationToken = default)
    {
        _exchangeAppId = appId?.Trim() ?? string.Empty;
        _exchangeOrganization = organization?.Trim() ?? string.Empty;
        _exchangeThumbprint = thumbprint?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(_exchangeAppId)
            || string.IsNullOrWhiteSpace(_exchangeOrganization)
            || string.IsNullOrWhiteSpace(_exchangeThumbprint))
        {
            return new AuthResult(false, string.Empty, string.Empty, string.Empty, "Exchange app auth requires App ID, Tenant Domain, and certificate thumbprint.");
        }

        var script = """
        $ErrorActionPreference = 'Stop'
        if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) {
            try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue } catch {}
            Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber
        }
        Import-Module ExchangeOnlineManagement
        Connect-ExchangeOnline -AppId '__APP_ID__' -Organization '__ORG__' -CertificateThumbprint '__THUMB__' -ShowBanner:$false -SkipLoadingFormatData | Out-Null
        $info = $null
        try { $info = Get-ConnectionInformation | Select-Object -First 1 } catch { $info = $null }
        $tenantId = ''
        $account = ''
        if ($info) {
            $tenantId = ($info.Organization -as [string])
            $account = ($info.UserPrincipalName -as [string])
        }
        $payload = [PSCustomObject]@{
            Module='ExchangeOnline'
            TenantId=$tenantId
            Account=$account
            Details='Connected (app auth)'
        } | ConvertTo-Json -Compress
        Set-Content -Path '__RESULT_FILE__' -Value $payload -Encoding UTF8
        """;
        script = script.Replace("__APP_ID__", _exchangeAppId.Replace("'", "''", StringComparison.Ordinal), StringComparison.Ordinal);
        script = script.Replace("__ORG__", _exchangeOrganization.Replace("'", "''", StringComparison.Ordinal), StringComparison.Ordinal);
        script = script.Replace("__THUMB__", _exchangeThumbprint.Replace("'", "''", StringComparison.Ordinal), StringComparison.Ordinal);
        var authResult = await RunInteractiveAuthScriptAsync(
            script,
            "Connecting Exchange Online...",
            progress,
            cancellationToken,
            visibleWindow: false,
            timeoutSeconds: 120);
        if (!authResult.Success)
        {
            return authResult;
        }

        return authResult;
    }

    public Task<AuthResult> ConnectPnPAsync(string siteUrl, string clientId, string tenantDomain, string thumbprint, IProgress<string>? progress = null, CancellationToken cancellationToken = default)
    {
        var safeUrl = siteUrl.Replace("'", "''", StringComparison.Ordinal);
        var safeClientId = clientId.Replace("'", "''", StringComparison.Ordinal);
        var safeTenantDomain = tenantDomain.Replace("'", "''", StringComparison.Ordinal);
        var safeThumbprint = thumbprint.Replace("'", "''", StringComparison.Ordinal);
        var script = """
        $ErrorActionPreference = 'Stop'
        if ([string]::IsNullOrWhiteSpace('__PNP_CLIENT_ID__')) { throw 'PnP Client ID is required.' }
        if ([string]::IsNullOrWhiteSpace('__TENANT_DOMAIN__')) { throw 'Tenant domain is required (example: contoso.onmicrosoft.com).' }
        if ([string]::IsNullOrWhiteSpace('__PNP_THUMBPRINT__')) { throw 'PnP certificate thumbprint is required. Run Set Up SharePoint App first.' }
        if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
            try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue } catch {}
            Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
        }
        Import-Module PnP.PowerShell
        Connect-PnPOnline -Url '__SITEURL__' -ClientId '__PNP_CLIENT_ID__' -Tenant '__TENANT_DOMAIN__' -Thumbprint '__PNP_THUMBPRINT__'
        $conn = Get-PnPConnection
        if (-not $conn) { throw 'PnP connection context was not returned.' }
        $payload = [PSCustomObject]@{
            Module='PnP'
            TenantId=($conn.Url -as [string])
            Account=($conn.ClientId -as [string])
            Details='Connected'
        } | ConvertTo-Json -Compress
        Set-Content -Path '__RESULT_FILE__' -Value $payload -Encoding UTF8
        """;
        script = script.Replace("__SITEURL__", safeUrl, StringComparison.Ordinal);
        script = script.Replace("__PNP_CLIENT_ID__", safeClientId, StringComparison.Ordinal);
        script = script.Replace("__TENANT_DOMAIN__", safeTenantDomain, StringComparison.Ordinal);
        script = script.Replace("__PNP_THUMBPRINT__", safeThumbprint, StringComparison.Ordinal);
        return RunInteractiveAuthScriptAsync(script, $"Connecting SharePoint PnP ({siteUrl})...", progress, cancellationToken);
    }

    public async Task<PnPAppSetupResult> EnsurePnPAppRegistrationAsync(string tenantDomain, IProgress<string>? progress = null, CancellationToken cancellationToken = default)
    {
        var resultFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-pnpapp-{Guid.NewGuid():N}.json");
        var stderrFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-pnpapp-{Guid.NewGuid():N}.err.txt");
        var safeTenantDomain = tenantDomain.Replace("'", "''", StringComparison.Ordinal);

        var script = $$"""
        $ErrorActionPreference = 'Stop'
        if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
            try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue } catch {}
            Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
        }
        Import-Module PnP.PowerShell
        $tenant = '{{safeTenantDomain}}'
        if ([string]::IsNullOrWhiteSpace($tenant)) {
            throw 'Tenant domain is required for PnP app registration setup.'
        }
        $graphPerms = @("Sites.Read.All", "Files.Read.All", "User.Read.All")
        $spoPerms = @("Sites.FullControl.All", "Sites.Manage.All")
        try {
            $app = Register-PnPEntraIDApp -ApplicationName "NewUserAutomation.PnP $(Get-Random)" -Tenant $tenant -Store CurrentUser -GraphApplicationPermissions $graphPerms -SharePointApplicationPermissions $spoPerms -ErrorAction Stop
        } catch {
            $app = Register-PnPEntraIDApp -ApplicationName "NewUserAutomation.PnP $(Get-Random)" -Tenant $tenant -Store CurrentUser -ErrorAction Stop
        }
        $raw = $app | ConvertTo-Json -Depth 20 -Compress
        $idCandidates = @(
            $app.AppId, $app.ClientId, $app.appId, $app.clientId,
            $app.ApplicationId, $app.'Application ID', $app.'Application (client) ID',
            $app.Id
        ) | Where-Object { $null -ne $_ } | ForEach-Object { [string]$_ }
        $id = ($idCandidates | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -ne 'null' } | Select-Object -First 1)
        if ([string]::IsNullOrWhiteSpace($id) -and $raw -match '[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}') {
            $id = $Matches[0]
        }
        $thumbCandidates = @(
            $app.Thumbprint, $app.thumbprint, $app.'Certificate Thumbprint', $app.CertificateThumbprint
        ) | Where-Object { $null -ne $_ } | ForEach-Object { [string]$_ }
        $thumb = ($thumbCandidates | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -ne 'null' } | Select-Object -First 1)
        if ([string]::IsNullOrWhiteSpace($thumb) -and $raw -match '(?im)"Thumbprint"\s*:\s*"([^"]+)"') {
            $thumb = $Matches[1]
        }
        if (-not [string]::IsNullOrWhiteSpace($thumb)) { $thumb = $thumb.Trim().Trim("'").Trim('"') }
        if ([string]::IsNullOrWhiteSpace($id) -or [string]::IsNullOrWhiteSpace($thumb)) { throw "Could not extract AppId/Thumbprint from Register-PnPEntraIDApp output." }

        $consentUrl = "https://login.microsoftonline.com/$tenant/adminconsent?client_id=$id&redirect_uri=http%3A%2F%2Flocalhost"
        [PSCustomObject]@{
            AppId = $id
            Thumbprint = $thumb
            TenantDomain = $tenant
            ConsentUrl = $consentUrl
        } | ConvertTo-Json -Compress | Set-Content -Path '{{resultFile}}' -Encoding UTF8
        """;

        var wrapped = $$"""
        try {
            {{script}}
            exit 0
        } catch {
            $_ | Out-File -FilePath '{{stderrFile}}' -Encoding utf8 -Force
            exit 1
        }
        """;

        var encoded = Convert.ToBase64String(Encoding.Unicode.GetBytes(wrapped));
        var psi = new ProcessStartInfo
        {
            FileName = "pwsh",
            Arguments = $"-NoProfile -ExecutionPolicy Bypass -EncodedCommand {encoded}",
            UseShellExecute = false,
            CreateNoWindow = true
        };

        progress?.Report("Creating or locating SharePoint app registration...");
        using var process = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start PowerShell.");
        await process.WaitForExitAsync(cancellationToken);

        try
        {
            if (process.ExitCode != 0)
            {
                var err = File.Exists(stderrFile) ? await File.ReadAllTextAsync(stderrFile, cancellationToken) : "PnP app setup failed.";
                return new PnPAppSetupResult(false, string.Empty, string.Empty, string.Empty, string.Empty, Simplify(err));
            }

            if (!File.Exists(resultFile))
            {
                return new PnPAppSetupResult(false, string.Empty, string.Empty, string.Empty, string.Empty, "PnP app setup completed but no payload was returned.");
            }

            var json = await File.ReadAllTextAsync(resultFile, cancellationToken);
            var doc = JsonDocument.Parse(json);
            var appId = doc.RootElement.TryGetProperty("AppId", out var appIdElement) ? appIdElement.GetString() ?? string.Empty : string.Empty;
            var tenant = doc.RootElement.TryGetProperty("TenantDomain", out var tenantElement) ? tenantElement.GetString() ?? string.Empty : string.Empty;
            var thumb = doc.RootElement.TryGetProperty("Thumbprint", out var thumbElement) ? thumbElement.GetString() ?? string.Empty : string.Empty;
            var consentUrl = doc.RootElement.TryGetProperty("ConsentUrl", out var consentElement) ? consentElement.GetString() ?? string.Empty : string.Empty;
            return new PnPAppSetupResult(true, appId, thumb, tenant, consentUrl, string.Empty);
        }
        catch (Exception ex)
        {
            return new PnPAppSetupResult(false, string.Empty, string.Empty, string.Empty, string.Empty, ex.Message);
        }
        finally
        {
            TryDelete(resultFile);
            TryDelete(stderrFile);
        }
    }

    public async Task<LiveExecutionResult> EnsureSharePointSiteFullControlForAppAsync(
        string appId,
        string siteUrl,
        string? tenantDomain = null,
        IProgress<string>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var safeAppId = appId.Replace("'", "''", StringComparison.Ordinal).Trim();
        var safeSiteUrl = siteUrl.Replace("'", "''", StringComparison.Ordinal).Trim();
        var safeTenant = (tenantDomain ?? string.Empty).Replace("'", "''", StringComparison.Ordinal).Trim();
        var resultFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-spositegrant-{Guid.NewGuid():N}.json");
        var stderrFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-spositegrant-{Guid.NewGuid():N}.err.txt");
        var script = $$"""
        $ErrorActionPreference = 'Stop'
        if ([string]::IsNullOrWhiteSpace('{{safeAppId}}')) { throw 'App ID is required.' }
        if ([string]::IsNullOrWhiteSpace('{{safeSiteUrl}}')) { throw 'SharePoint site URL is required.' }
        if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
            try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue } catch {}
            Install-Module PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
        }
        Import-Module PnP.PowerShell

        $siteUrl = '{{safeSiteUrl}}'
        $appId = '{{safeAppId}}'
        $tenantDomain = '{{safeTenant}}'
        $siteUri = [Uri]$siteUrl
        $adminHost = $siteUri.Host -replace '\.sharepoint\.com$', '-admin.sharepoint.com'
        $adminUrl = "$($siteUri.Scheme)://$adminHost"

        $authErrors = @()
        $connected = $false
        # PnP interactive/device auth requires a public client app id.
        # Prefer env override; otherwise provision/reuse an interactive-login app in this tenant.
        $adminAuthClientId = [Environment]::GetEnvironmentVariable("PNP_INTERACTIVE_CLIENT_ID")
        if ([string]::IsNullOrWhiteSpace($adminAuthClientId) -and -not [string]::IsNullOrWhiteSpace($tenantDomain)) {
            try {
                $interactiveApp = Register-PnPEntraIDAppForInteractiveLogin -ApplicationName "NewUserAutomation.PnP.Interactive $(Get-Random)" -Tenant $tenantDomain -ErrorAction Stop
                $raw = $interactiveApp | ConvertTo-Json -Depth 20 -Compress
                $idCandidates = @(
                    $interactiveApp.AppId, $interactiveApp.ClientId, $interactiveApp.appId, $interactiveApp.clientId,
                    $interactiveApp.ApplicationId, $interactiveApp.'Application ID', $interactiveApp.'Application (client) ID',
                    $interactiveApp.Id
                ) | Where-Object { $null -ne $_ } | ForEach-Object { [string]$_ }
                $adminAuthClientId = ($idCandidates | Where-Object { -not [string]::IsNullOrWhiteSpace($_) -and $_ -ne 'null' } | Select-Object -First 1)
                if ([string]::IsNullOrWhiteSpace($adminAuthClientId) -and $raw -match '[0-9a-fA-F]{8}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{4}\-[0-9a-fA-F]{12}') {
                    $adminAuthClientId = $Matches[0]
                }
            } catch {
                $authErrors += "RegisterInteractiveApp: $($_.Exception.Message)"
            }
        }
        if ([string]::IsNullOrWhiteSpace($adminAuthClientId)) {
            throw "Could not resolve an interactive login Client ID for SharePoint admin auth. Set PNP_INTERACTIVE_CLIENT_ID environment variable, or provide Tenant Domain and retry setup. Details: $($authErrors -join ' | ')"
        }
        # Keep retries minimal to avoid spawning multiple browser windows.
        $attempts = @(
            @{ Mode = "DeviceLogin+Tenant"; UseTenant = $true; Interactive = $false; DeviceLogin = $true },
            @{ Mode = "DeviceLogin"; UseTenant = $false; Interactive = $false; DeviceLogin = $true }
        )
        foreach ($attempt in $attempts) {
            if ($connected) { break }
            if ($attempt.UseTenant -and [string]::IsNullOrWhiteSpace($tenantDomain)) { continue }
            try {
                if ($attempt.Interactive) {
                    if ($attempt.UseTenant) {
                        Connect-PnPOnline -Url $adminUrl -Interactive -ClientId $adminAuthClientId -Tenant $tenantDomain -ErrorAction Stop
                    } else {
                        Connect-PnPOnline -Url $adminUrl -Interactive -ClientId $adminAuthClientId -ErrorAction Stop
                    }
                } else {
                    if ($attempt.UseTenant) {
                        Connect-PnPOnline -Url $adminUrl -DeviceLogin -ClientId $adminAuthClientId -Tenant $tenantDomain -ErrorAction Stop
                    } else {
                        Connect-PnPOnline -Url $adminUrl -DeviceLogin -ClientId $adminAuthClientId -ErrorAction Stop
                    }
                }
                $connected = $true
            } catch {
                $authErrors += "$($attempt.Mode): $($_.Exception.Message)"
            }
        }
        if (-not $connected) {
            throw "SharePoint admin authentication failed. Tried: Interactive and DeviceLogin (with/without tenant hint). Complete the sign-in prompt, or grant site permission manually in SharePoint admin. Details: $($authErrors -join ' | ')"
        }
        $existing = @(Get-PnPAzureADAppSitePermission -Site $siteUrl -ErrorAction SilentlyContinue | Where-Object { [string]$_.AppId -eq $appId })
        $grant = $existing | Select-Object -First 1
        if ($grant) {
            $hasFull = $false
            foreach ($p in @($grant.Permissions)) {
                if ([string]$p -eq 'FullControl') { $hasFull = $true; break }
            }
            if (-not $hasFull) {
                Set-PnPAzureADAppSitePermission -Site $siteUrl -PermissionId $grant.Id -Permissions FullControl | Out-Null
                [PSCustomObject]@{ Success = $true; Detail = "Updated site app permission to FullControl on $siteUrl." } | ConvertTo-Json -Compress | Set-Content -Path '{{resultFile}}' -Encoding UTF8
            } else {
                [PSCustomObject]@{ Success = $true; Detail = "Site app permission already FullControl on $siteUrl." } | ConvertTo-Json -Compress | Set-Content -Path '{{resultFile}}' -Encoding UTF8
            }
        } else {
            try {
                Grant-PnPAzureADAppSitePermission -AppId $appId -DisplayName "NewUserAutomation" -Site $siteUrl -Permissions FullControl | Out-Null
                [PSCustomObject]@{ Success = $true; Detail = "Granted FullControl site app permission on $siteUrl." } | ConvertTo-Json -Compress | Set-Content -Path '{{resultFile}}' -Encoding UTF8
            } catch {
                $msg = $_.Exception.Message
                if ($msg -match 'already' -or $msg -match 'same as the one sent' -or $msg -match 'mismatcherror') {
                    $existingAfter = @(Get-PnPAzureADAppSitePermission -Site $siteUrl -ErrorAction SilentlyContinue | Where-Object { [string]$_.AppId -eq $appId })
                    $grantAfter = $existingAfter | Select-Object -First 1
                    if ($grantAfter) {
                        $hasFullAfter = $false
                        foreach ($p in @($grantAfter.Permissions)) {
                            if ([string]$p -eq 'FullControl') { $hasFullAfter = $true; break }
                        }
                        if (-not $hasFullAfter) {
                            Set-PnPAzureADAppSitePermission -Site $siteUrl -PermissionId $grantAfter.Id -Permissions FullControl | Out-Null
                        }
                        [PSCustomObject]@{ Success = $true; Detail = "Site app permission already existed; ensured FullControl on $siteUrl." } | ConvertTo-Json -Compress | Set-Content -Path '{{resultFile}}' -Encoding UTF8
                    } else {
                        throw
                    }
                } else {
                    throw
                }
            }
        }
        """;

        var wrapped = $$"""
        try {
            {{script}}
            exit 0
        } catch {
            $_ | Out-File -FilePath '{{stderrFile}}' -Encoding utf8 -Force
            exit 1
        }
        """;
        var encoded = Convert.ToBase64String(Encoding.Unicode.GetBytes(wrapped));
        var psi = new ProcessStartInfo
        {
            FileName = "pwsh",
            Arguments = $"-NoProfile -ExecutionPolicy Bypass -EncodedCommand {encoded}",
            UseShellExecute = true,
            CreateNoWindow = false
        };

        progress?.Report("Granting SharePoint site permission for app (admin sign-in may prompt)...");
        using var process = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start PowerShell.");
        await process.WaitForExitAsync(cancellationToken);
        if (process.ExitCode != 0 || !File.Exists(resultFile))
        {
            var err = File.Exists(stderrFile)
                ? await File.ReadAllTextAsync(stderrFile, cancellationToken)
                : "SharePoint site permission grant failed.";
            return new LiveExecutionResult(false, Simplify(err));
        }

        try
        {
            var json = await File.ReadAllTextAsync(resultFile, cancellationToken);
            var doc = JsonDocument.Parse(json);
            var success = doc.RootElement.TryGetProperty("Success", out var s) && s.ValueKind == JsonValueKind.True;
            var detail = doc.RootElement.TryGetProperty("Detail", out var d) ? d.GetString() ?? string.Empty : string.Empty;
            return new LiveExecutionResult(success, string.IsNullOrWhiteSpace(detail) ? (success ? "Completed." : "Failed.") : detail);
        }
        catch (Exception ex)
        {
            return new LiveExecutionResult(false, Simplify(ex.Message));
        }
        finally
        {
            TryDelete(resultFile);
            TryDelete(stderrFile);
        }
    }

    public async Task<bool> CertificateThumbprintExistsAsync(string thumbprint, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(thumbprint))
        {
            return false;
        }

        var safeThumb = thumbprint.Replace("'", "''", StringComparison.Ordinal).Trim();
        var resultFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-certexists-{Guid.NewGuid():N}.json");
        var stderrFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-certexists-{Guid.NewGuid():N}.err.txt");
        var script = $$"""
        $ErrorActionPreference = 'Stop'
        $thumb = '{{safeThumb}}'
        $cert = Get-ChildItem Cert:\CurrentUser\My | Where-Object { $_.Thumbprint -eq $thumb } | Select-Object -First 1
        [PSCustomObject]@{ Exists = ($null -ne $cert) } | ConvertTo-Json -Compress | Set-Content -Path '{{resultFile}}' -Encoding UTF8
        """;

        var exec = await RunJsonScriptAsync(script, resultFile, stderrFile, cancellationToken, deleteResultFile: false);
        if (!exec.Success || !File.Exists(resultFile))
        {
            return false;
        }

        try
        {
            var json = await File.ReadAllTextAsync(resultFile, cancellationToken);
            var doc = JsonDocument.Parse(json);
            return doc.RootElement.TryGetProperty("Exists", out var exists) && exists.ValueKind == JsonValueKind.True;
        }
        catch
        {
            return false;
        }
        finally
        {
            TryDelete(resultFile);
            TryDelete(stderrFile);
        }
    }

    public async Task<string> FindNewestCertificateThumbprintBySubjectAsync(string subjectContains, CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(subjectContains))
        {
            return string.Empty;
        }

        var safeSubject = subjectContains.Replace("'", "''", StringComparison.Ordinal).Trim();
        var resultFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-certfind-{Guid.NewGuid():N}.json");
        var stderrFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-certfind-{Guid.NewGuid():N}.err.txt");
        var script = $$"""
        $ErrorActionPreference = 'Stop'
        $hint = '{{safeSubject}}'
        $cert = Get-ChildItem Cert:\CurrentUser\My `
            | Where-Object { $_.Subject -like "*$hint*" -and $_.HasPrivateKey } `
            | Sort-Object NotAfter -Descending `
            | Select-Object -First 1
        [PSCustomObject]@{
            Thumbprint = (if ($cert) { [string]$cert.Thumbprint } else { '' })
        } | ConvertTo-Json -Compress | Set-Content -Path '{{resultFile}}' -Encoding UTF8
        """;

        var exec = await RunJsonScriptAsync(script, resultFile, stderrFile, cancellationToken, deleteResultFile: false);
        if (!exec.Success || !File.Exists(resultFile))
        {
            return string.Empty;
        }

        try
        {
            var json = await File.ReadAllTextAsync(resultFile, cancellationToken);
            var doc = JsonDocument.Parse(json);
            return doc.RootElement.TryGetProperty("Thumbprint", out var thumb)
                ? thumb.GetString() ?? string.Empty
                : string.Empty;
        }
        catch
        {
            return string.Empty;
        }
        finally
        {
            TryDelete(resultFile);
            TryDelete(stderrFile);
        }
    }

    public async Task<ExchangeCertSetupResult> EnsureExchangeCertificateAsync(
        string subjectName,
        string exportDirectory,
        IProgress<string>? progress = null,
        CancellationToken cancellationToken = default)
    {
        var exportDir = string.IsNullOrWhiteSpace(exportDirectory)
            ? Path.Combine(AppContext.BaseDirectory, "certs")
            : exportDirectory;
        var resultFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-exchangecert-{Guid.NewGuid():N}.json");
        var stderrFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-exchangecert-{Guid.NewGuid():N}.err.txt");
        var safeSubject = string.IsNullOrWhiteSpace(subjectName)
            ? "NewUserAutomation.Exchange"
            : subjectName.Replace("'", "''", StringComparison.Ordinal).Trim();
        var safeExportDir = exportDir.Replace("'", "''", StringComparison.Ordinal);

        var script = $$"""
        $ErrorActionPreference = 'Stop'
        $subjectCore = '{{safeSubject}}'
        if ([string]::IsNullOrWhiteSpace($subjectCore)) { $subjectCore = 'NewUserAutomation.Exchange' }
        $fullSubject = if ($subjectCore.StartsWith('CN=', [System.StringComparison]::OrdinalIgnoreCase)) { $subjectCore } else { "CN=$subjectCore" }

        $cert = Get-ChildItem Cert:\CurrentUser\My `
            | Where-Object { $_.Subject -eq $fullSubject -and $_.HasPrivateKey } `
            | Sort-Object NotAfter -Descending `
            | Select-Object -First 1

        if (-not $cert -or $cert.NotAfter -lt (Get-Date).AddDays(30)) {
            $cert = New-SelfSignedCertificate `
                -Subject $fullSubject `
                -CertStoreLocation "Cert:\CurrentUser\My" `
                -KeyExportPolicy Exportable `
                -KeySpec Signature `
                -NotAfter (Get-Date).AddYears(2)
        }

        $exportDir = '{{safeExportDir}}'
        New-Item -ItemType Directory -Path $exportDir -Force | Out-Null
        $sanitized = ($subjectCore -replace '[^a-zA-Z0-9\.\-_]', '_')
        if ([string]::IsNullOrWhiteSpace($sanitized)) { $sanitized = 'NewUserAutomation.Exchange' }
        $cerPath = Join-Path $exportDir "$sanitized.cer"
        Export-Certificate -Cert $cert -FilePath $cerPath -Force | Out-Null

        [PSCustomObject]@{
            Subject = $fullSubject
            Thumbprint = ($cert.Thumbprint -as [string])
            CerPath = $cerPath
            Store = 'CurrentUser\My'
        } | ConvertTo-Json -Compress | Set-Content -Path '{{resultFile}}' -Encoding UTF8
        """;

        var wrapped = $$"""
        try {
            {{script}}
            exit 0
        } catch {
            $_ | Out-File -FilePath '{{stderrFile}}' -Encoding utf8 -Force
            exit 1
        }
        """;

        var encoded = Convert.ToBase64String(Encoding.Unicode.GetBytes(wrapped));
        var psi = new ProcessStartInfo
        {
            FileName = "pwsh",
            Arguments = $"-NoProfile -ExecutionPolicy Bypass -EncodedCommand {encoded}",
            UseShellExecute = false,
            CreateNoWindow = true
        };

        progress?.Report("Generating Exchange certificate...");
        using var process = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start PowerShell.");
        await process.WaitForExitAsync(cancellationToken);

        try
        {
            if (process.ExitCode != 0)
            {
                var err = File.Exists(stderrFile) ? await File.ReadAllTextAsync(stderrFile, cancellationToken) : "Exchange certificate setup failed.";
                return new ExchangeCertSetupResult(false, string.Empty, string.Empty, string.Empty, Simplify(err));
            }

            if (!File.Exists(resultFile))
            {
                return new ExchangeCertSetupResult(false, string.Empty, string.Empty, string.Empty, "Exchange certificate setup completed but no payload was returned.");
            }

            var json = await File.ReadAllTextAsync(resultFile, cancellationToken);
            var doc = JsonDocument.Parse(json);
            var subject = doc.RootElement.TryGetProperty("Subject", out var subjectElement) ? subjectElement.GetString() ?? string.Empty : string.Empty;
            var thumbprint = doc.RootElement.TryGetProperty("Thumbprint", out var thumbElement) ? thumbElement.GetString() ?? string.Empty : string.Empty;
            var cerPath = doc.RootElement.TryGetProperty("CerPath", out var cerElement) ? cerElement.GetString() ?? string.Empty : string.Empty;
            return new ExchangeCertSetupResult(true, subject, thumbprint, cerPath, string.Empty);
        }
        catch (Exception ex)
        {
            return new ExchangeCertSetupResult(false, string.Empty, string.Empty, string.Empty, ex.Message);
        }
        finally
        {
            TryDelete(resultFile);
            TryDelete(stderrFile);
        }
    }

    public async Task<CustomerEnterpriseAppSetupResult> EnsureCustomerEnterpriseAppSetupAsync(
        string customerName,
        string tenantDomain,
        string certificateCerPath,
        string certificateThumbprint,
        IProgress<string>? progress = null,
        CancellationToken cancellationToken = default)
    {
        if (string.IsNullOrWhiteSpace(customerName))
        {
            return new CustomerEnterpriseAppSetupResult(false, string.Empty, string.Empty, string.Empty, string.Empty, "Customer name is required.");
        }
        if (string.IsNullOrWhiteSpace(tenantDomain))
        {
            return new CustomerEnterpriseAppSetupResult(false, string.Empty, string.Empty, string.Empty, string.Empty, "Tenant domain is required.");
        }
        if (string.IsNullOrWhiteSpace(certificateCerPath) || !File.Exists(certificateCerPath))
        {
            return new CustomerEnterpriseAppSetupResult(false, string.Empty, string.Empty, string.Empty, string.Empty, "Certificate .cer path is missing or not found.");
        }
        if (string.IsNullOrWhiteSpace(certificateThumbprint))
        {
            return new CustomerEnterpriseAppSetupResult(false, string.Empty, string.Empty, string.Empty, string.Empty, "Certificate thumbprint is required.");
        }

        var resultFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-customerapp-{Guid.NewGuid():N}.json");
        var stderrFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-customerapp-{Guid.NewGuid():N}.err.txt");
        var safeCustomer = customerName.Replace("'", "''", StringComparison.Ordinal).Trim();
        var safeTenant = tenantDomain.Replace("'", "''", StringComparison.Ordinal).Trim();
        var safeCerPath = certificateCerPath.Replace("'", "''", StringComparison.Ordinal).Trim();
        var safeThumb = certificateThumbprint.Replace("'", "''", StringComparison.Ordinal).Trim();

        var script = $$"""
        $ErrorActionPreference = 'Stop'
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Authentication)) {
            try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue } catch {}
            Install-Module Microsoft.Graph.Authentication -Scope CurrentUser -Force -AllowClobber
        }
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Applications)) {
            try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue } catch {}
            Install-Module Microsoft.Graph.Applications -Scope CurrentUser -Force -AllowClobber
        }
        if (-not (Get-Module -ListAvailable -Name Microsoft.Graph.Identity.DirectoryManagement)) {
            try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue } catch {}
            Install-Module Microsoft.Graph.Identity.DirectoryManagement -Scope CurrentUser -Force -AllowClobber
        }
        Import-Module Microsoft.Graph.Authentication
        Import-Module Microsoft.Graph.Applications
        Import-Module Microsoft.Graph.Identity.DirectoryManagement

        $tenant = '{{safeTenant}}'
        $display = "NewUserAutomation {{safeCustomer}}"
        $certPath = '{{safeCerPath}}'
        $thumb = '{{safeThumb}}'

        Connect-MgGraph -TenantId $tenant -Scopes "Application.ReadWrite.All","AppRoleAssignment.ReadWrite.All","Directory.Read.All","RoleManagement.ReadWrite.Directory" -NoWelcome
        $app = Get-MgApplication -Filter "displayName eq '$display'" | Select-Object -First 1
        if (-not $app) {
            $app = New-MgApplication -DisplayName $display -SignInAudience "AzureADMyOrg"
        }
        $clientId = [string]$app.AppId
        $appObjectId = [string]$app.Id
        $nativeRedirect = "https://login.microsoftonline.com/common/oauth2/nativeclient"
        $loopbackRedirect = "http://localhost"
        $webRedirect = "https://localhost"
        $existingRedirects = @()
        if ($app.PublicClient -and $app.PublicClient.RedirectUris) {
            $existingRedirects = @($app.PublicClient.RedirectUris | ForEach-Object { [string]$_ })
        }
        if (-not ($existingRedirects -contains $nativeRedirect) -or -not ($existingRedirects -contains $loopbackRedirect)) {
            $mergedRedirects = @($existingRedirects + $nativeRedirect + $loopbackRedirect | Select-Object -Unique)
            Update-MgApplication -ApplicationId $appObjectId -PublicClient @{ RedirectUris = $mergedRedirects }
            $app = Get-MgApplication -ApplicationId $appObjectId
        }
        try {
            # Required for device-login/public client auth fallback flows.
            Update-MgApplication -ApplicationId $appObjectId -IsFallbackPublicClient $true | Out-Null
            $app = Get-MgApplication -ApplicationId $appObjectId
        } catch {
        }
        $existingWebRedirects = @()
        if ($app.Web -and $app.Web.RedirectUris) {
            $existingWebRedirects = @($app.Web.RedirectUris | ForEach-Object { [string]$_ })
        }
        if (-not ($existingWebRedirects -contains $webRedirect)) {
            $mergedWebRedirects = @($existingWebRedirects + $webRedirect | Select-Object -Unique)
            Update-MgApplication -ApplicationId $appObjectId -Web @{ RedirectUris = $mergedWebRedirects }
            $app = Get-MgApplication -ApplicationId $appObjectId
        }
        $confirmedPublicRedirects = @()
        if ($app.PublicClient -and $app.PublicClient.RedirectUris) {
            $confirmedPublicRedirects = @($app.PublicClient.RedirectUris | ForEach-Object { [string]$_ })
        }
        $confirmedWebRedirects = @()
        if ($app.Web -and $app.Web.RedirectUris) {
            $confirmedWebRedirects = @($app.Web.RedirectUris | ForEach-Object { [string]$_ })
        }

        $graphSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0000-c000-000000000000'" | Select-Object -First 1
        $spoSp = Get-MgServicePrincipal -Filter "appId eq '00000003-0000-0ff1-ce00-000000000000'" | Select-Object -First 1
        $exoSp = Get-MgServicePrincipal -Filter "appId eq '00000002-0000-0ff1-ce00-000000000000'" | Select-Object -First 1
        if (-not $graphSp -or -not $spoSp -or -not $exoSp) {
            throw "Could not locate required Microsoft service principals (Graph/SharePoint/Exchange) in this tenant."
        }

        function Get-AppRoleId([object]$sp, [string]$value) {
            $role = $sp.AppRoles | Where-Object {
                $_.Value -eq $value -and $_.AllowedMemberTypes -contains "Application" -and $_.IsEnabled
            } | Select-Object -First 1
            if (-not $role) { return $null }
            return [Guid]$role.Id
        }

        $graphRoleValues = @("User.Read.All", "Group.ReadWrite.All", "Directory.ReadWrite.All", "Organization.Read.All")
        $graphRoleAccess = @()
        foreach ($v in $graphRoleValues) {
            $id = Get-AppRoleId $graphSp $v
            if ($id) {
                $graphRoleAccess += @{ id = $id; type = "Role" }
            }
        }
        if ($graphRoleAccess.Count -eq 0) {
            throw "Could not resolve Graph application roles. Check tenant service principal availability."
        }

        $spoRoleValues = @("Sites.FullControl.All", "Sites.Manage.All")
        $spoRoleAccess = @()
        foreach ($v in $spoRoleValues) {
            $id = Get-AppRoleId $spoSp $v
            if ($id) {
                $spoRoleAccess += @{ id = $id; type = "Role" }
            }
        }
        if ($spoRoleAccess.Count -eq 0) {
            throw "Could not resolve SharePoint application roles Sites.FullControl.All / Sites.Manage.All."
        }

        $exoRoleId = Get-AppRoleId $exoSp "Exchange.ManageAsApp"
        if (-not $exoRoleId) {
            # Well-known fallback for Exchange.ManageAsApp app role ID
            $exoRoleId = [Guid]"dc50a0fb-09a3-484d-be87-e023b12c6440"
        }
        if (-not $exoRoleId) {
            throw "Could not resolve Exchange.ManageAsApp application role."
        }

        $requiredResourceAccess = @(
            @{
                resourceAppId = [string]$graphSp.AppId
                resourceAccess = $graphRoleAccess
            },
            @{
                resourceAppId = [string]$spoSp.AppId
                resourceAccess = $spoRoleAccess
            },
            @{
                resourceAppId = [string]$exoSp.AppId
                resourceAccess = @(@{ id = $exoRoleId; type = "Role" })
            }
        )
        Update-MgApplication -ApplicationId $appObjectId -RequiredResourceAccess $requiredResourceAccess
        $app = Get-MgApplication -ApplicationId $appObjectId
        $hasExchangeManageAsAppInManifest = $false
        foreach ($r in @($app.RequiredResourceAccess)) {
            if ($r.ResourceAppId -eq [string]$exoSp.AppId) {
                foreach ($ra in @($r.ResourceAccess)) {
                    if ([string]$ra.Type -eq "Role" -and [Guid]$ra.Id -eq $exoRoleId) {
                        $hasExchangeManageAsAppInManifest = $true
                        break
                    }
                }
            }
            if ($hasExchangeManageAsAppInManifest) { break }
        }
        if (-not $hasExchangeManageAsAppInManifest) {
            throw "Exchange.ManageAsApp was not persisted to RequiredResourceAccess. Tenant rejected application permission manifest update."
        }

        $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2($certPath)
        $existing = @($app.KeyCredentials)
        $certB64 = [Convert]::ToBase64String($cert.RawData)
        $hasCert = $false
        foreach ($k in $existing) {
            if ($k.Key) {
                $existingB64 = [Convert]::ToBase64String($k.Key)
                if ($existingB64 -eq $certB64) { $hasCert = $true; break }
            }
        }
        if (-not $hasCert) {
            $newKey = @{
                type = "AsymmetricX509Cert"
                usage = "Verify"
                key = $cert.RawData
                displayName = "NewUserAutomation cert"
                startDateTime = $cert.NotBefore.ToUniversalTime()
                endDateTime = $cert.NotAfter.ToUniversalTime()
            }
            Update-MgApplication -ApplicationId $appObjectId -KeyCredentials ($existing + $newKey)
        }

        $sp = Get-MgServicePrincipal -Filter "appId eq '$clientId'" | Select-Object -First 1
        if (-not $sp) { $sp = New-MgServicePrincipal -AppId $clientId }
        $spId = [string]$sp.Id

        if ($exoSp) {
            $assignments = Get-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $spId -All -ErrorAction SilentlyContinue
            $hasRole = $false
            foreach ($a in $assignments) {
                if ($a.ResourceId -eq $exoSp.Id -and $a.AppRoleId -eq $exoRoleId) { $hasRole = $true; break }
            }
            if (-not $hasRole) {
                New-MgServicePrincipalAppRoleAssignment -ServicePrincipalId $spId -PrincipalId $spId -ResourceId $exoSp.Id -AppRoleId $exoRoleId | Out-Null
            }
        }

        # Ensure app service principal has Exchange Administrator directory role when possible.
        # Some operators do not have enough directory-role privileges; do not hard-fail setup in that case.
        $exchangeRoleWarning = ''
        $exchangeRole = Get-MgDirectoryRole -Filter "displayName eq 'Exchange Administrator'" | Select-Object -First 1
        if (-not $exchangeRole) {
            $template = Get-MgDirectoryRoleTemplate -Filter "displayName eq 'Exchange Administrator'" | Select-Object -First 1
            if ($template) {
                try {
                    Enable-MgDirectoryRole -DirectoryRoleTemplateId $template.Id | Out-Null
                } catch {
                    $exchangeRoleWarning = "Could not enable Exchange Administrator directory role template automatically. Details: $($_.Exception.Message)"
                }
                $exchangeRole = Get-MgDirectoryRole -Filter "displayName eq 'Exchange Administrator'" | Select-Object -First 1
            }
        }
        if ($exchangeRole) {
            $roleMembers = Get-MgDirectoryRoleMember -DirectoryRoleId $exchangeRole.Id -All -ErrorAction SilentlyContinue
            $alreadyInRole = $false
            foreach ($m in @($roleMembers)) {
                if ([string]$m.Id -eq $spId) { $alreadyInRole = $true; break }
            }
            if (-not $alreadyInRole) {
                try {
                    New-MgDirectoryRoleMemberByRef -DirectoryRoleId $exchangeRole.Id -BodyParameter @{
                        "@odata.id" = "https://graph.microsoft.com/v1.0/directoryObjects/$spId"
                    } | Out-Null
                } catch {
                    $exchangeRoleWarning = "Could not assign Exchange Administrator role to app service principal automatically. Ensure you are signed in with Privileged Role Administrator or Global Administrator and assign this role manually if Exchange app-only actions fail. Details: $($_.Exception.Message)"
                }
            }
        } elseif ([string]::IsNullOrWhiteSpace($exchangeRoleWarning)) {
            $exchangeRoleWarning = "Exchange Administrator directory role not available during setup. Assign it manually to the app service principal if Exchange app-only actions fail."
        }

        Disconnect-MgGraph | Out-Null

        $consentRedirect = $null
        if ($confirmedWebRedirects -contains $webRedirect) {
            $consentRedirect = $webRedirect
        } elseif ($confirmedPublicRedirects -contains $nativeRedirect) {
            $consentRedirect = $nativeRedirect
        }
        if (-not $consentRedirect) {
            $consentRedirect = $nativeRedirect
        }
        $encodedRedirect = [System.Uri]::EscapeDataString($consentRedirect)
        $consentUrl = "https://login.microsoftonline.com/$tenant/adminconsent?client_id=$clientId&redirect_uri=$encodedRedirect"
        [PSCustomObject]@{
            Success = $true
            ClientId = $clientId
            TenantDomain = $tenant
            Thumbprint = $thumb
            ConsentUrl = $consentUrl
            Warning = $exchangeRoleWarning
        } | ConvertTo-Json -Compress | Set-Content -Path '{{resultFile}}' -Encoding UTF8
        """;

        var wrapped = $$"""
        try {
            {{script}}
            exit 0
        } catch {
            $_ | Out-File -FilePath '{{stderrFile}}' -Encoding utf8 -Force
            exit 1
        }
        """;
        var encoded = Convert.ToBase64String(Encoding.Unicode.GetBytes(wrapped));
        var psi = new ProcessStartInfo
        {
            FileName = "pwsh",
            Arguments = $"-NoProfile -ExecutionPolicy Bypass -EncodedCommand {encoded}",
            UseShellExecute = true,
            CreateNoWindow = false
        };

        progress?.Report("Setting up customer app registration and permissions...");
        using var process = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start PowerShell.");
        await process.WaitForExitAsync(cancellationToken);

        try
        {
            if (process.ExitCode != 0)
            {
                var err = File.Exists(stderrFile) ? await File.ReadAllTextAsync(stderrFile, cancellationToken) : "Customer app setup failed.";
                return new CustomerEnterpriseAppSetupResult(false, string.Empty, string.Empty, string.Empty, string.Empty, Simplify(err));
            }
            if (!File.Exists(resultFile))
            {
                return new CustomerEnterpriseAppSetupResult(false, string.Empty, string.Empty, string.Empty, string.Empty, "Customer app setup completed but no payload was returned.");
            }

            var json = await File.ReadAllTextAsync(resultFile, cancellationToken);
            var doc = JsonDocument.Parse(json);
            var clientId = doc.RootElement.TryGetProperty("ClientId", out var c) ? c.GetString() ?? string.Empty : string.Empty;
            var tenant = doc.RootElement.TryGetProperty("TenantDomain", out var t) ? t.GetString() ?? string.Empty : string.Empty;
            var thumb = doc.RootElement.TryGetProperty("Thumbprint", out var th) ? th.GetString() ?? string.Empty : string.Empty;
            var consent = doc.RootElement.TryGetProperty("ConsentUrl", out var u) ? u.GetString() ?? string.Empty : string.Empty;
            var warning = doc.RootElement.TryGetProperty("Warning", out var w) ? w.GetString() ?? string.Empty : string.Empty;
            return new CustomerEnterpriseAppSetupResult(true, clientId, tenant, thumb, consent, warning);
        }
        catch (Exception ex)
        {
            return new CustomerEnterpriseAppSetupResult(false, string.Empty, string.Empty, string.Empty, string.Empty, ex.Message);
        }
        finally
        {
            TryDelete(resultFile);
            TryDelete(stderrFile);
        }
    }

    private static async Task<AuthResult> RunInteractiveAuthScriptAsync(
        string script,
        string phaseMessage,
        IProgress<string>? progress,
        CancellationToken cancellationToken,
        bool visibleWindow = false,
        int timeoutSeconds = 300)
    {
        var resultFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-auth-{Guid.NewGuid():N}.json");
        var stderrFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-auth-{Guid.NewGuid():N}.err.txt");
        var phaseFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-auth-{Guid.NewGuid():N}.phase.txt");
        script = script.Replace("__RESULT_FILE__", resultFile.Replace("'", "''", StringComparison.Ordinal), StringComparison.Ordinal);
        script = script.Replace("__PHASE_FILE__", phaseFile.Replace("'", "''", StringComparison.Ordinal), StringComparison.Ordinal);
        var wrappedScript = $$"""
        try {
            {{script}}
            exit 0
        } catch {
            $_ | Out-File -FilePath '{{stderrFile}}' -Encoding utf8 -Force
            exit 1
        }
        """;
        var encoded = Convert.ToBase64String(Encoding.Unicode.GetBytes(wrappedScript));
        var psi = new ProcessStartInfo
        {
            FileName = "pwsh",
            Arguments = $"-NoProfile -ExecutionPolicy Bypass -EncodedCommand {encoded}",
            UseShellExecute = visibleWindow,
            CreateNoWindow = !visibleWindow
        };

        progress?.Report(phaseMessage);
        using var process = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start PowerShell.");

        // Some interactive auth flows can successfully write the result file but keep pwsh running.
        // Treat the result file as completion and avoid blocking indefinitely on process exit.
        var start = DateTime.UtcNow;
        var lastPhase = string.Empty;
        while (true)
        {
            cancellationToken.ThrowIfCancellationRequested();

            if ((DateTime.UtcNow - start).TotalSeconds > timeoutSeconds)
            {
                try { process.Kill(entireProcessTree: true); } catch { }
                var timeoutDetail = string.IsNullOrWhiteSpace(lastPhase)
                    ? $"Authentication timed out after {timeoutSeconds} seconds."
                    : $"Authentication timed out after {timeoutSeconds} seconds. Last phase: {lastPhase}";
                return new AuthResult(false, string.Empty, string.Empty, string.Empty, timeoutDetail);
            }

            if (File.Exists(phaseFile))
            {
                try
                {
                    var lines = await File.ReadAllLinesAsync(phaseFile, cancellationToken);
                    var phase = lines.LastOrDefault(x => !string.IsNullOrWhiteSpace(x))?.Trim() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(phase) && !string.Equals(phase, lastPhase, StringComparison.Ordinal))
                    {
                        lastPhase = phase;
                        progress?.Report(phase);
                    }
                }
                catch
                {
                }
            }

            if (File.Exists(resultFile))
            {
                break;
            }

            if (process.HasExited)
            {
                break;
            }

            await Task.Delay(250, cancellationToken);
        }

        try
        {
            if (!File.Exists(resultFile))
            {
                if (!process.HasExited)
                {
                    try { process.Kill(entireProcessTree: true); } catch { }
                }

                var err = File.Exists(stderrFile) ? await File.ReadAllTextAsync(stderrFile, cancellationToken) : "Authentication failed.";
                return new AuthResult(false, string.Empty, string.Empty, string.Empty, Simplify(err));
            }

            if (!process.HasExited && !visibleWindow)
            {
                try { process.Kill(entireProcessTree: true); } catch { }
            }

            var json = await File.ReadAllTextAsync(resultFile, cancellationToken);
            var doc = JsonDocument.Parse(json);
            return new AuthResult(
                true,
                doc.RootElement.TryGetProperty("TenantId", out var tenant) ? tenant.GetString() ?? string.Empty : string.Empty,
                doc.RootElement.TryGetProperty("Account", out var account) ? account.GetString() ?? string.Empty : string.Empty,
                doc.RootElement.TryGetProperty("Details", out var details) ? details.GetString() ?? string.Empty : string.Empty,
                string.Empty);
        }
        catch (Exception ex)
        {
            return new AuthResult(false, string.Empty, string.Empty, string.Empty, ex.Message);
        }
        finally
        {
            TryDelete(resultFile);
            TryDelete(stderrFile);
            TryDelete(phaseFile);
        }
    }

    public async Task<Dictionary<string, ExchangeAccessTargetResolution>> CheckExchangeAccessTargetsAsync(IReadOnlyList<string> targets, IProgress<string>? progress = null, CancellationToken cancellationToken = default)
    {
        var result = new Dictionary<string, ExchangeAccessTargetResolution>(StringComparer.OrdinalIgnoreCase);
        if (targets.Count == 0)
        {
            return result;
        }

        var resultFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-exchange-targets-{Guid.NewGuid():N}.json");
        var stderrFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-exchange-targets-{Guid.NewGuid():N}.err.txt");
        var joined = string.Join("|", targets).Replace("'", "''", StringComparison.Ordinal);

        var script = $$"""
        $ErrorActionPreference = 'Stop'
        $conn = $null
        try {
            $conn = Get-ConnectionInformation | Select-Object -First 1
        } catch {
            $conn = $null
        }
        if (-not $conn) {
            Connect-ExchangeOnline -AppId '{{EscapePsLiteral(_exchangeAppId)}}' -Organization '{{EscapePsLiteral(_exchangeOrganization)}}' -CertificateThumbprint '{{EscapePsLiteral(_exchangeThumbprint)}}' -ShowBanner:$false -SkipLoadingFormatData | Out-Null
            $conn = Get-ConnectionInformation | Select-Object -First 1
        }
        if (-not $conn) {
            throw "No active Exchange session in persistent host. Reconnect Exchange from the Sign In page."
        }
        $targets = '{{joined}}' -split '\|'
        $out = @()
        foreach ($target in $targets) {
            if ([string]::IsNullOrWhiteSpace($target)) { continue }
            $exists = $false
            $kind = 'NotFound'
            $action = 'Target not found. Check spelling or use primary SMTP address.'
            $details = ''
            $lookupErrors = @()
            try {
                $x = Get-EXOMailbox -Identity $target -ResultSize 1 -ErrorAction Stop
                if ($x) {
                    $exists = $true
                    $rt = [string]$x.RecipientTypeDetails
                    $details = $rt
                    if ($rt -eq 'SharedMailbox') {
                        $kind = 'SharedMailbox'
                        $action = 'Grant mailbox access permissions.'
                    } elseif ($rt -eq 'UserMailbox') {
                        $kind = 'UserMailbox'
                        $action = 'Review request; this is a user mailbox, not a shared mailbox.'
                    } else {
                        $kind = 'Mailbox'
                        $action = "Grant mailbox access permissions (recipient type: $rt)."
                    }
                }
            } catch {
                $lookupErrors += "Get-EXOMailbox: $($_.Exception.Message)"
                try {
                    $dg = Get-DistributionGroup -Identity $target -ErrorAction Stop
                    if ($dg) {
                        $exists = $true
                        $rt = [string]$dg.RecipientTypeDetails
                        $details = $rt
                        if ($rt -eq 'MailUniversalSecurityGroup') {
                            $kind = 'MailEnabledSecurityGroup'
                            $action = 'Add user as member of this mail-enabled security group.'
                        } else {
                            $kind = 'DistributionGroup'
                            $action = 'Add user as member of this distribution group.'
                        }
                    }
                } catch {
                    $lookupErrors += "Get-DistributionGroup: $($_.Exception.Message)"
                    try {
                        $rec = Get-EXORecipient -Identity $target -ResultSize 1 -ErrorAction Stop
                        if ($rec) {
                            $exists = $true
                            $rt = [string]$rec.RecipientTypeDetails
                            $details = $rt
                            if ($rt -eq 'GroupMailbox') {
                                $kind = 'UnifiedGroup'
                                $action = 'Add user as member of this Microsoft 365 group.'
                            } else {
                                $kind = 'OtherRecipient'
                                $action = "Recipient exists (type: $rt). Review required action manually."
                            }
                        }
                    } catch {
                        $lookupErrors += "Get-EXORecipient: $($_.Exception.Message)"
                    }
                }
            }
            if (-not $exists -and $lookupErrors.Count -gt 0) {
                $details = ($lookupErrors -join ' | ')
            }
            $out += [PSCustomObject]@{
                Address = $target
                Exists = $exists
                Kind = $kind
                Action = $action
                Details = $details
            }
        }
        $out | ConvertTo-Json -Compress | Set-Content -Path '{{resultFile}}' -Encoding UTF8
        """;

        progress?.Report("Checking Exchange targets (mailbox/group/distribution list)...");
        var exec = await RunJsonScriptAsync(script, resultFile, stderrFile, cancellationToken, deleteResultFile: false);

        try
        {
            if (!exec.Success || !File.Exists(resultFile))
            {
                var failureDetail = string.IsNullOrWhiteSpace(exec.Detail)
                    ? "Exchange target lookup failed before results were returned."
                    : exec.Detail;
                foreach (var target in targets.Where(t => !string.IsNullOrWhiteSpace(t)))
                {
                    result[target] = new ExchangeAccessTargetResolution(
                        target,
                        false,
                        "LookupFailed",
                        "Reconnect Exchange and verify app permissions/RBAC scope.",
                        failureDetail);
                }
                return result;
            }

            var json = await File.ReadAllTextAsync(resultFile, cancellationToken);
            var doc = JsonDocument.Parse(json);
            if (doc.RootElement.ValueKind == JsonValueKind.Array)
            {
                foreach (var item in doc.RootElement.EnumerateArray())
                {
                    var address = item.TryGetProperty("Address", out var a) ? a.GetString() ?? string.Empty : string.Empty;
                    var exists = item.TryGetProperty("Exists", out var e) && e.ValueKind == JsonValueKind.True;
                    var kind = item.TryGetProperty("Kind", out var k) ? k.GetString() ?? "NotFound" : "NotFound";
                    var action = item.TryGetProperty("Action", out var ac) ? ac.GetString() ?? string.Empty : string.Empty;
                    var details = item.TryGetProperty("Details", out var d) ? d.GetString() ?? string.Empty : string.Empty;
                    if (!string.IsNullOrWhiteSpace(address))
                    {
                        result[address] = new ExchangeAccessTargetResolution(address, exists, kind, action, details);
                    }
                }
            }
            else if (doc.RootElement.ValueKind == JsonValueKind.Object)
            {
                var item = doc.RootElement;
                var address = item.TryGetProperty("Address", out var a) ? a.GetString() ?? string.Empty : string.Empty;
                var exists = item.TryGetProperty("Exists", out var e) && e.ValueKind == JsonValueKind.True;
                var kind = item.TryGetProperty("Kind", out var k) ? k.GetString() ?? "NotFound" : "NotFound";
                var action = item.TryGetProperty("Action", out var ac) ? ac.GetString() ?? string.Empty : string.Empty;
                var details = item.TryGetProperty("Details", out var d) ? d.GetString() ?? string.Empty : string.Empty;
                if (!string.IsNullOrWhiteSpace(address))
                {
                    result[address] = new ExchangeAccessTargetResolution(address, exists, kind, action, details);
                }
            }
        }
        catch
        {
            return result;
        }
        finally
        {
            TryDelete(resultFile);
        }

        return result;
    }

    private async Task<Microsoft.Graph.Models.Group?> FindGroupByDisplayNameAsync(string groupName, CancellationToken cancellationToken)
    {
        if (_graphClient is null)
        {
            return null;
        }

        try
        {
            var escaped = EscapeGraphLiteral(groupName);
            var groups = await _graphClient.Groups.GetAsync(config =>
            {
                config.QueryParameters.Filter = $"displayName eq '{escaped}'";
                config.QueryParameters.Top = 5;
                config.QueryParameters.Select = ["id", "displayName"];
            }, cancellationToken);
            return groups?.Value?.FirstOrDefault(g => string.Equals(g.DisplayName, groupName, StringComparison.OrdinalIgnoreCase));
        }
        catch
        {
            return null;
        }
    }

    private static string BuildMailNickname(params string[] rawValues)
    {
        var source = string.Join("-", rawValues.Where(x => !string.IsNullOrWhiteSpace(x)));
        var cleaned = new string(source.Where(ch => char.IsLetterOrDigit(ch) || ch == '.').ToArray());
        if (string.IsNullOrWhiteSpace(cleaned))
        {
            cleaned = $"newuser{DateTimeOffset.UtcNow.ToUnixTimeSeconds()}";
        }

        return cleaned.Length > 60 ? cleaned[..60] : cleaned;
    }

    private async Task<LiveExecutionResult> RunJsonScriptAsync(string script, string resultFile, string stderrFile, CancellationToken cancellationToken, bool deleteResultFile = true)
    {
        var wrapped = $$"""
        try {
            {{script}}
            exit 0
        } catch {
            $_ | Out-File -FilePath '{{stderrFile}}' -Encoding utf8 -Force
            exit 1
        }
        """;
        var encoded = Convert.ToBase64String(Encoding.Unicode.GetBytes(wrapped));
        var psi = new ProcessStartInfo
        {
            FileName = "pwsh",
            Arguments = $"-NoProfile -ExecutionPolicy Bypass -EncodedCommand {encoded}",
            UseShellExecute = false,
            CreateNoWindow = true
        };

        using var process = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start PowerShell.");
        await process.WaitForExitAsync(cancellationToken);
        try
        {
            if (process.ExitCode != 0 || !File.Exists(resultFile))
            {
                var err = File.Exists(stderrFile)
                    ? await File.ReadAllTextAsync(stderrFile, cancellationToken)
                    : "Execution script failed.";
                return new LiveExecutionResult(false, Simplify(err));
            }

            var json = await File.ReadAllTextAsync(resultFile, cancellationToken);
            var doc = JsonDocument.Parse(json);
            if (doc.RootElement.ValueKind == JsonValueKind.Array)
            {
                // Some helper scripts return raw arrays; treat that as successful execution.
                return new LiveExecutionResult(true, "Completed.");
            }

            if (doc.RootElement.ValueKind != JsonValueKind.Object)
            {
                return new LiveExecutionResult(false, "Execution script returned an unexpected payload.");
            }

            var success = doc.RootElement.TryGetProperty("Success", out var s) && s.ValueKind == JsonValueKind.True;
            var detail = doc.RootElement.TryGetProperty("Detail", out var d) ? d.GetString() ?? string.Empty : string.Empty;
            return new LiveExecutionResult(success, string.IsNullOrWhiteSpace(detail) ? (success ? "Completed." : "Failed.") : detail);
        }
        catch (Exception ex)
        {
            return new LiveExecutionResult(false, Simplify(ex.Message));
        }
        finally
        {
            if (deleteResultFile)
            {
                TryDelete(resultFile);
            }
            TryDelete(stderrFile);
        }
    }

    private async Task EnsureExchangeHostStartedAsync(CancellationToken cancellationToken, bool interactiveWindow)
    {
        await _exchangeHostLock.WaitAsync(cancellationToken);
        try
        {
            if (_exchangeHostProcess is not null && !_exchangeHostProcess.HasExited && _exchangeHostInput is not null)
            {
                return;
            }

            var psi = new ProcessStartInfo
            {
                FileName = "pwsh",
                Arguments = "-NoLogo -NoProfile -ExecutionPolicy Bypass -NoExit",
                UseShellExecute = false,
                CreateNoWindow = !interactiveWindow,
                RedirectStandardInput = true,
                RedirectStandardOutput = true,
                RedirectStandardError = true
            };
            _exchangeHostProcess = Process.Start(psi) ?? throw new InvalidOperationException("Failed to start persistent Exchange host.");
            _exchangeHostInput = _exchangeHostProcess.StandardInput;
            await _exchangeHostInput.WriteLineAsync("$ErrorActionPreference = 'Stop'");
            await _exchangeHostInput.WriteLineAsync("if (-not (Get-Module -ListAvailable -Name ExchangeOnlineManagement)) { try { Set-PSRepository -Name PSGallery -InstallationPolicy Trusted -ErrorAction SilentlyContinue } catch {}; Install-Module ExchangeOnlineManagement -Scope CurrentUser -Force -AllowClobber }");
            await _exchangeHostInput.WriteLineAsync("Import-Module ExchangeOnlineManagement");
            await _exchangeHostInput.FlushAsync();
        }
        finally
        {
            _exchangeHostLock.Release();
        }
    }

    private async Task ResetExchangeHostAsync(CancellationToken cancellationToken)
    {
        await _exchangeHostLock.WaitAsync(cancellationToken);
        try
        {
            if (_exchangeHostProcess is not null && !_exchangeHostProcess.HasExited)
            {
                try { _exchangeHostProcess.Kill(entireProcessTree: true); } catch { }
            }

            _exchangeHostProcess = null;
            _exchangeHostInput = null;
        }
        finally
        {
            _exchangeHostLock.Release();
        }
    }

    private async Task<LiveExecutionResult> ExecuteInExchangeHostAsync(string script, string resultFile, CancellationToken cancellationToken, int timeoutSeconds = 180)
    {
        await EnsureExchangeHostStartedAsync(cancellationToken, interactiveWindow: false);
        await _exchangeHostLock.WaitAsync(cancellationToken);
        try
        {
            if (_exchangeHostProcess is null || _exchangeHostInput is null || _exchangeHostProcess.HasExited)
            {
                return new LiveExecutionResult(false, "Persistent Exchange host is not running.");
            }

            var stderrFile = Path.Combine(Path.GetTempPath(), $"newuserautomation-exhost-{Guid.NewGuid():N}.err.txt");
            var wrapped = $$"""
            try {
                {{script}}
            } catch {
                $_ | Out-File -FilePath '{{stderrFile}}' -Encoding utf8 -Force
            }
            """;
            await _exchangeHostInput.WriteLineAsync(wrapped);
            await _exchangeHostInput.FlushAsync();

            var start = DateTime.UtcNow;
            while ((DateTime.UtcNow - start).TotalSeconds < timeoutSeconds)
            {
                cancellationToken.ThrowIfCancellationRequested();
                if (File.Exists(resultFile))
                {
                    TryDelete(stderrFile);
                    return new LiveExecutionResult(true, string.Empty);
                }

                if (File.Exists(stderrFile))
                {
                    var err = await File.ReadAllTextAsync(stderrFile, cancellationToken);
                    TryDelete(stderrFile);
                    return new LiveExecutionResult(false, Simplify(err));
                }

                if (_exchangeHostProcess.HasExited)
                {
                    return new LiveExecutionResult(false, "Persistent Exchange host exited unexpectedly.");
                }

                await Task.Delay(200, cancellationToken);
            }

            if (_exchangeHostProcess is not null && !_exchangeHostProcess.HasExited)
            {
                try { _exchangeHostProcess.Kill(entireProcessTree: true); } catch { }
            }
            _exchangeHostProcess = null;
            _exchangeHostInput = null;
            return new LiveExecutionResult(false, "Timed out waiting for Exchange host response.");
        }
        finally
        {
            _exchangeHostLock.Release();
        }
    }

    private static string Simplify(string raw)
    {
        var cleaned = raw.Replace("_x000D__x000A_", "\n", StringComparison.OrdinalIgnoreCase);
        var lines = cleaned.Split('\n', StringSplitOptions.RemoveEmptyEntries)
            .Select(x => x.Trim())
            .Where(x => !x.StartsWith("#< CLIXML", StringComparison.OrdinalIgnoreCase))
            .Where(x => !x.StartsWith("<Objs", StringComparison.OrdinalIgnoreCase))
            .Where(x => !x.StartsWith("<Obj", StringComparison.OrdinalIgnoreCase))
            .Where(x => !x.StartsWith("<S S=", StringComparison.OrdinalIgnoreCase));
        var merged = string.Join(" ", lines);
        if (merged.Contains("Insufficient privileges", StringComparison.OrdinalIgnoreCase))
        {
            return "Missing admin rights for app registration. Sign in with an Entra admin account and try again.";
        }
        if (merged.Contains("AADSTS53003", StringComparison.OrdinalIgnoreCase) ||
            merged.Contains("Error Code: 53003", StringComparison.OrdinalIgnoreCase))
        {
            return "Sign-in blocked by Conditional Access (AADSTS53003). This tenant requires additional controls (for example a compliant/hybrid-joined device, approved network/location, or approved client app). Ask your Entra admin to review Conditional Access sign-in logs for this request and either allow this device/user flow or provide an approved access path.";
        }
        if (merged.Contains("that is restricted by your admin", StringComparison.OrdinalIgnoreCase))
        {
            return "Sign-in blocked by tenant policy. Ask your Entra admin to allow Microsoft Graph PowerShell sign-in for this user/device or provide an approved admin workstation.";
        }
        if (merged.Contains("OperationStopped: Unauthorized", StringComparison.OrdinalIgnoreCase)
            || merged.Contains("Unauthorized", StringComparison.OrdinalIgnoreCase))
        {
            return "Exchange app authentication is not authorized for this app/certificate yet. Ensure the app has Exchange.ManageAsApp (Application) permission with admin consent, and that the service principal has Exchange admin RBAC (for example Exchange Administrator role) in this tenant.";
        }
        return merged;
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

public sealed record PnPAppSetupResult(bool Success, string AppId, string Thumbprint, string TenantDomain, string ConsentUrl, string ErrorMessage);
public sealed record ExchangeCertSetupResult(bool Success, string Subject, string Thumbprint, string CerPath, string ErrorMessage);
public sealed record CustomerEnterpriseAppSetupResult(bool Success, string ClientId, string TenantDomain, string Thumbprint, string ConsentUrl, string ErrorMessage);
public sealed record DirectoryUserMatch(string Id, string DisplayName, string UserPrincipalName, string Mail);
public sealed record ExchangeAccessTargetResolution(string Address, bool Exists, string Kind, string RequiredAction, string Details);
