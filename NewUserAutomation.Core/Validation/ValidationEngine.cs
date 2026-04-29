using System.Text.RegularExpressions;
using NewUserAutomation.Core.Models;

namespace NewUserAutomation.Core.Validation;

public sealed class ValidationEngine
{
    private static readonly Regex EmailRegex = new(@"^[^@\s]+@[^@\s]+\.[^@\s]+$", RegexOptions.Compiled);

    public ValidationReport Validate(NewUserRequest request)
    {
        var errors = new List<string>();
        var warnings = new List<string>();

        if (string.IsNullOrWhiteSpace(request.FirstName) || string.IsNullOrWhiteSpace(request.LastName))
        {
            errors.Add("FirstName and LastName are required.");
        }

        if (!EmailRegex.IsMatch(request.PrimaryEmail))
        {
            errors.Add("PrimaryEmail must be a valid email address.");
        }
        if (!string.IsNullOrWhiteSpace(request.SecondaryEmail) && !EmailRegex.IsMatch(request.SecondaryEmail))
        {
            errors.Add("SecondaryEmail must be a valid email address when provided.");
        }

        if (request.LicenseSkus.Count == 0)
        {
            warnings.Add("No license selected; user will be created without license assignment.");
        }

        if (string.IsNullOrWhiteSpace(request.TemporaryPassword))
        {
            warnings.Add("TemporaryPassword is empty; mailbox password will not be set from the form.");
        }

        if (string.IsNullOrWhiteSpace(request.JobTitle))
        {
            warnings.Add("JobTitle is empty; Microsoft 365 contact info job title will be skipped.");
        }

        if (string.IsNullOrWhiteSpace(request.RequestApprovedBy))
        {
            errors.Add("ApprovedBy is required before execution.");
        }

        if (request.RequiresPersonalSharePointFolder)
        {
            if (string.IsNullOrWhiteSpace(request.PersonalSharePointFolderName))
            {
                errors.Add("Personal SharePoint folder was requested, but folder name could not be resolved.");
            }

            if (string.IsNullOrWhiteSpace(request.PersonalSharePointPermissionGroup))
            {
                errors.Add("Personal SharePoint folder was requested, but permission group name could not be resolved.");
            }
        }

        return new ValidationReport(errors, warnings);
    }
}

public sealed record ValidationReport(IReadOnlyList<string> Errors, IReadOnlyList<string> Warnings)
{
    public bool IsValid => Errors.Count == 0;
}
