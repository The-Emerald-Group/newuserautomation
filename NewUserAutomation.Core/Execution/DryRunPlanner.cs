using NewUserAutomation.Core.Models;

namespace NewUserAutomation.Core.Execution;

public sealed class DryRunPlanner
{
    public IReadOnlyList<ExecutionStep> BuildPlan(NewUserRequest request)
    {
        var identityDetails = new List<string> { $"Create user {request.DisplayName} with UPN {request.Upn}" };
        if (!string.IsNullOrWhiteSpace(request.TemporaryPassword))
        {
            identityDetails.Add("set mailbox password from form");
        }
        if (!string.IsNullOrWhiteSpace(request.JobTitle))
        {
            identityDetails.Add($"set Job Title to '{request.JobTitle}' in contact info");
        }

        var steps = new List<ExecutionStep>
        {
            new("Identity.CreateUser", string.Join("; ", identityDetails)),
        };

        foreach (var license in request.LicenseSkus.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
        {
            steps.Add(new("Licenses.AssignLicense", $"Assign license: {license}"));
        }

        foreach (var group in request.GroupAccess.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
        {
            steps.Add(new("Groups.AddMembership", $"Add to group: {group}"));
        }

        foreach (var mailbox in request.SharedMailboxAccess.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
        {
            steps.Add(new("Mailboxes.GrantAccess", $"Grant access to Exchange target (shared mailbox/group/distribution list): {mailbox}"));
        }

        foreach (var access in request.SharePointAccess.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
        {
            steps.Add(new("SharePoint.GrantAccess", $"Grant SharePoint access: {access}"));
        }

        if (request.RequiresPersonalSharePointFolder)
        {
            var groupText = string.IsNullOrWhiteSpace(request.PersonalSharePointPermissionGroup)
                ? "Create personal SharePoint group (name unresolved)"
                : $"Create SharePoint group: {request.PersonalSharePointPermissionGroup}";
            steps.Add(new("SharePoint.CreatePersonalFolder", $"Create personal folder '{request.PersonalSharePointFolderName}' at site root level (same level as 'Documents'); {groupText}"));
            steps.Add(new("SharePoint.GrantAccess", $"Add new user to personal SharePoint group: {request.PersonalSharePointPermissionGroup}"));

            foreach (var member in request.PersonalSharePointAdditionalMembers.OrderBy(x => x, StringComparer.OrdinalIgnoreCase))
            {
                steps.Add(new("SharePoint.GrantAccess", $"Add additional member '{member}' to personal SharePoint group '{request.PersonalSharePointPermissionGroup}'"));
            }
        }

        return steps;
    }
}

public sealed record ExecutionStep(string Action, string Description);
