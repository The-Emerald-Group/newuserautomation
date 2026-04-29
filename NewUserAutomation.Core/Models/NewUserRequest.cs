using System.Text.Json.Serialization;

namespace NewUserAutomation.Core.Models;

public sealed class NewUserRequest
{
    public string FirstName { get; init; } = string.Empty;
    public string LastName { get; init; } = string.Empty;
    public string DisplayName { get; init; } = string.Empty;
    public string PreferredUsername { get; init; } = string.Empty;
    public string TemporaryPassword { get; init; } = string.Empty;
    public string JobTitle { get; init; } = string.Empty;
    public string PrimaryEmail { get; init; } = string.Empty;
    public string SecondaryEmail { get; init; } = string.Empty;
    public SecondaryEmailHandlingMode SecondaryEmailMode { get; init; } = SecondaryEmailHandlingMode.AliasOnPrimaryUser;
    public List<string> LicenseSkus { get; init; } = [];
    public List<string> GroupAccess { get; init; } = [];
    public List<string> SharedMailboxAccess { get; init; } = [];
    public List<string> SharePointAccess { get; init; } = [];
    public bool RequiresPersonalSharePointFolder { get; init; }
    public string PersonalSharePointFolderName { get; init; } = string.Empty;
    public string PersonalSharePointPermissionGroup { get; init; } = string.Empty;
    public List<string> PersonalSharePointAdditionalMembers { get; init; } = [];
    public string SpecialRequirements { get; init; } = string.Empty;
    public string RequestApprovedBy { get; init; } = string.Empty;
    public IReadOnlyList<string> ParseDiagnostics { get; init; } = [];

    [JsonIgnore]
    public string Upn
    {
        get
        {
            if (string.IsNullOrWhiteSpace(PreferredUsername))
            {
                return string.Empty;
            }

            if (PreferredUsername.Contains('@'))
            {
                return PreferredUsername.Trim().ToLowerInvariant();
            }

            var domain = PrimaryEmail.Split('@').LastOrDefault();
            if (string.IsNullOrWhiteSpace(domain))
            {
                domain = "emerald-group.example";
            }

            return $"{PreferredUsername.Trim().ToLowerInvariant()}@{domain.Trim().ToLowerInvariant()}";
        }
    }
}

public enum SecondaryEmailHandlingMode
{
    AliasOnPrimaryUser = 0,
    SeparateMailboxWithDelegation = 1
}
