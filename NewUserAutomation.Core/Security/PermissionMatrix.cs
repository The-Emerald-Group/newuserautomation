namespace NewUserAutomation.Core.Security;

public sealed record PermissionRequirement(
    string Action,
    IReadOnlyList<string> GraphScopes,
    IReadOnlyList<string> ExchangeRoles,
    IReadOnlyList<string> PnPRights);

public sealed class PermissionMatrix
{
    public static readonly IReadOnlyList<PermissionRequirement> DefaultRequirements =
    [
        new("CreateUser", ["User.ReadWrite.All"], [], []),
        new("AssignLicense", ["User.ReadWrite.All"], [], []),
        new("AddGroupMembership", ["Group.ReadWrite.All"], [], []),
        new("GrantMailboxAccess", [], ["Mailbox.Permission.Assign"], []),
        new("GrantSharePointAccess", [], [], ["Site.Member.Write"]),
        new("CreatePersonalFolder", [], [], ["Site.Member.Write"])
    ];

    public PermissionCheckResult Evaluate(
        IReadOnlyCollection<string> selectedActions,
        IReadOnlyCollection<string> grantedGraphScopes,
        IReadOnlyCollection<string> grantedExchangeRoles,
        IReadOnlyCollection<string> grantedPnPRights)
    {
        var normalizedActions = selectedActions
            .Select(static action => action.Contains('.') ? action.Split('.', 2)[1] : action)
            .ToHashSet(StringComparer.OrdinalIgnoreCase);

        var missing = new List<string>();
        foreach (var req in DefaultRequirements.Where(r => normalizedActions.Contains(r.Action)))
        {
            missing.AddRange(req.GraphScopes.Where(scope => !grantedGraphScopes.Contains(scope)).Select(scope => $"{req.Action}: missing Graph scope {scope}"));
            missing.AddRange(req.ExchangeRoles.Where(role => !grantedExchangeRoles.Contains(role)).Select(role => $"{req.Action}: missing EXO role {role}"));
            missing.AddRange(req.PnPRights.Where(right => !grantedPnPRights.Contains(right)).Select(right => $"{req.Action}: missing PnP right {right}"));
        }

        return new PermissionCheckResult(missing);
    }
}

public sealed record PermissionCheckResult(IReadOnlyList<string> MissingPermissions)
{
    public bool CanExecute => MissingPermissions.Count == 0;
}
