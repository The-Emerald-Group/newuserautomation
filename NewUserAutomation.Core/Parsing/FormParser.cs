using NewUserAutomation.Core.Models;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace NewUserAutomation.Core.Parsing;

public sealed class FormParser
{
    private static readonly string[] KnownSharePointSelections =
    [
        "HR ADMIN",
        "HR SHARED",
        "SIT ADMIN",
        "SIT SHARED",
        "OWN PERSONAL FOLDER",
        "PAYROLL (ACCOUNTS ONLY)"
    ];

    private static readonly Dictionary<string, string> SharePointGroupAliasMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["HR ADMIN"] = "SP-Documents-DL-HRAdmin-RW",
        ["HR SHARED"] = "SP-Documents-DL-HRShared-RW",
        ["SIT ADMIN"] = "SP-Documents-DL-HE Admin-RW",
        ["SIT SHARED"] = "SP-Documents-DL-HE Shared-RW"
    };

    private sealed record FieldSpec(string Canonical, bool Required, params string[] Aliases);

    private static readonly FieldSpec[] ParseContract =
    [
        new("FirstName", true, "first name", "firstname"),
        new("LastName", true, "last name", "surname", "lastname"),
        new("DisplayName", false, "display name", "employee name"),
        new("PreferredUsername", true, "preferred username", "preferred server login username", "username"),
        new("TemporaryPassword", false, "temporary password", "preferred server password"),
        new("JobTitle", false, "job title", "jobtitle", "position", "role title"),
        new("PrimaryEmail", true, "primary email", "preferred email address", "email"),
        new("SecondaryEmail", false, "secondary email"),
        new("SecondaryEmailMode", false, "secondary email mode", "secondary mailbox mode"),
        new("LicenseSkus", false, "licenses", "licences", "license"),
        new("GroupAccess", false, "group access", "groups"),
        new("SharedMailboxAccess", false, "shared mailbox access", "shared mailbox"),
        new("SharePointAccess", false, "sharepoint access", "sharepoint folder access", "share point access"),
        new("SpecialRequirements", false, "special requirements", "any other requirements"),
        new("RequestApprovedBy", true, "approved by", "approved by supervisor", "request approved by")
    ];

    private static readonly Dictionary<string, string[]> SectionAliasMap = new(StringComparer.OrdinalIgnoreCase)
    {
        ["SharedMailboxAccessStart"] = ["Shared Mailbox access", "Shared mailbox access", "Shared Mailbox Access"],
        ["SharedMailboxAccessEnd"] = ["Phone", "Phone Number", "Phone Setup"],
        ["SharePointAccessStart"] = ["Sharepoint Folder Access", "SharePoint Folder Access", "Sharepoint online access links", "Sharepoint"],
        ["SharePointAccessEnd"] = ["Access over Users SharePoint Folder", "Access over user SharePoint folder", "Phone", "Any Other Requirements"],
        ["SpecialRequirementsStart"] = ["Any Other Requirements", "Special Requirements", "Access to Apps/Programs"],
        ["SpecialRequirementsEnd"] = ["Sharepoint online access links", "Internet Setup", "Phone"]
    };

    private static readonly Dictionary<string, string> CanonicalMap = BuildCanonicalMap();

    public NewUserRequest ParseFromFile(string filePath)
    {
        var extension = Path.GetExtension(filePath);
        if (extension.Equals(".docx", StringComparison.OrdinalIgnoreCase))
        {
            return ParseFromDocx(filePath);
        }

        return ParseFromKeyValueText(File.ReadAllText(filePath));
    }

    public NewUserRequest ParseFromKeyValueText(string rawText)
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        var diagnostics = new List<string>();
        foreach (var line in rawText.Split('\n', StringSplitOptions.RemoveEmptyEntries))
        {
            var idx = line.IndexOf(':');
            if (idx <= 0)
            {
                continue;
            }

            var key = FieldNormalizer.NormalizeScalar(line[..idx]);
            var value = FieldNormalizer.NormalizeScalar(line[(idx + 1)..]);
            if (!CanonicalMap.TryGetValue(key, out var canonical))
            {
                diagnostics.Add($"Unknown field ignored: {key}");
                continue;
            }

            diagnostics.Add($"Mapped field '{key}' -> '{canonical}'.");
            map[canonical] = value;
        }

        var firstName = Read(map, "FirstName");
        var lastName = Read(map, "LastName");
        var displayName = Read(map, "DisplayName");
        var preferredUsername = Read(map, "PreferredUsername").ToLowerInvariant();
        var primaryEmail = Read(map, "PrimaryEmail").ToLowerInvariant();
        var licenseSkus = ResolveLicenseSkus(FieldNormalizer.NormalizeList(Read(map, "LicenseSkus")), diagnostics);
        var approvedBy = Read(map, "RequestApprovedBy");
        diagnostics.AddRange(BuildContractDiagnostics(firstName, lastName, displayName, preferredUsername, primaryEmail, approvedBy));

        return new NewUserRequest
        {
            FirstName = firstName,
            LastName = lastName,
            DisplayName = displayName,
            PreferredUsername = preferredUsername,
            TemporaryPassword = Read(map, "TemporaryPassword"),
            JobTitle = Read(map, "JobTitle"),
            PrimaryEmail = primaryEmail,
            SecondaryEmail = Read(map, "SecondaryEmail").ToLowerInvariant(),
            SecondaryEmailMode = ParseSecondaryEmailMode(Read(map, "SecondaryEmailMode")),
            LicenseSkus = licenseSkus,
            GroupAccess = FieldNormalizer.NormalizeList(Read(map, "GroupAccess")),
            SharedMailboxAccess = FieldNormalizer.NormalizeList(Read(map, "SharedMailboxAccess")),
            SharePointAccess = FieldNormalizer.NormalizeList(Read(map, "SharePointAccess")),
            SpecialRequirements = Read(map, "SpecialRequirements"),
            RequestApprovedBy = approvedBy,
            ParseDiagnostics = diagnostics
        };
    }

    public NewUserRequest ParseFromDocx(string docxPath)
    {
        var lines = ExtractDocxLines(docxPath);
        var fullText = string.Join("\n", lines);
        var diagnostics = new List<string>();

        var employeeName = MatchValue(fullText, @"Employee\s*Name\s*\(First\s*Name\/Surname\)\s*(?<v>[^\r\n]+)", diagnostics, "DisplayName", "Employee Name (First Name/Surname)");
        if (string.IsNullOrWhiteSpace(employeeName))
        {
            employeeName = MatchValue(fullText, @"Employee\s*Name\s*(?<v>[^\r\n]+)", diagnostics, "DisplayName", "Employee Name");
        }
        var nameParts = employeeName.Split(' ', StringSplitOptions.RemoveEmptyEntries);
        var firstName = nameParts.Length > 0 ? nameParts[0] : string.Empty;
        var lastName = nameParts.Length > 1 ? string.Join(" ", nameParts.Skip(1)) : string.Empty;

        var preferredUsername = MatchValue(fullText, @"Preferred\s*server\s*login\s*username\s*(?<v>\S+)", diagnostics, "PreferredUsername", "Preferred server login username");
        if (string.IsNullOrWhiteSpace(preferredUsername))
        {
            preferredUsername = MatchValue(fullText, @"Preferred\s*username\s*(?<v>\S+)", diagnostics, "PreferredUsername", "Preferred username");
        }

        var temporaryPassword = MatchValue(fullText, @"Preferred\s*Server\s*Password\s*(?<v>\S+)", diagnostics, "TemporaryPassword", "Preferred Server Password");
        var primaryEmail = MatchValue(fullText, @"Preferred\s*Email\s*Address\s*(?<v>[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})", diagnostics, "PrimaryEmail", "Preferred Email Address");
        var secondaryEmail = MatchValue(fullText, @"Secondary\s*email.*?(?<v>[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,})", diagnostics, "SecondaryEmail", "Secondary email");
        var approvedBy = MatchValue(fullText, @"Approved\s*By\s*Supervisor\s*(?<v>[^\r\n]+)", diagnostics, "RequestApprovedBy", "Approved By Supervisor");

        var licenseSkus = new List<string>();
        if (Regex.IsMatch(fullText, @"365\s*Business\s*Standard\s*Licence.*?\bYes\b", RegexOptions.IgnoreCase | RegexOptions.Singleline))
        {
            licenseSkus.Add("BUSINESS_STANDARD");
            diagnostics.Add("LicenseSkus: found BUSINESS_STANDARD from '365 Business Standard Licence = Yes'");
        }
        if (Regex.IsMatch(fullText, @"365\s*Exchange\s*online\s*mailbox\s*licence.*?\bYes\b", RegexOptions.IgnoreCase | RegexOptions.Singleline))
        {
            licenseSkus.Add("EXCHANGE_ONLINE_MAILBOX");
            diagnostics.Add("LicenseSkus: found EXCHANGE_ONLINE_MAILBOX from '365 Exchange online mailbox licence = Yes'");
        }
        licenseSkus = ResolveLicenseSkus(licenseSkus, diagnostics);
        diagnostics.Add(licenseSkus.Count == 0 ? "LicenseSkus: no explicit license toggles detected." : $"LicenseSkus: resolved {licenseSkus.Count} item(s).");

        var sharedMailboxes = ExtractEmailsFromSection(lines, "SharedMailboxAccessStart", "SharedMailboxAccessEnd", diagnostics);
        var sharePointSection = ExtractSectionItems(lines, "SharePointAccessStart", "SharePointAccessEnd", diagnostics);
        if (sharePointSection.Count == 0)
        {
            sharePointSection = ExtractSectionItems(lines, "SharePointAccessStart", "Phone", diagnostics);
        }

        var sharePointAccess = NormalizeSharePointAccess(sharePointSection);
        sharePointAccess = RecoverSharePointSelections(sharePointAccess, sharePointSection, fullText, diagnostics);
        diagnostics.Add($"SharePointAccess: resolved {sharePointAccess.Count} item(s) after normalization.");
        var personalFolderMembers = ExtractUserIdentitiesFromSection(lines, diagnostics);

        var specialRequirements = ExtractSectionText(fullText, "SpecialRequirementsStart", "SpecialRequirementsEnd", diagnostics);
        if (string.IsNullOrWhiteSpace(specialRequirements))
        {
            specialRequirements = ExtractSectionText(fullText, "Access to Apps/Programs", "Internet Setup", diagnostics);
        }
        diagnostics.Add($"SpecialRequirements: {(string.IsNullOrWhiteSpace(specialRequirements) ? "empty" : "captured")}");
        var jobTitle = MatchValue(fullText, @"Job\s*Title\s*(?<v>[^\r\n]+)", diagnostics, "JobTitle", "Job Title");
        diagnostics.AddRange(BuildContractDiagnostics(firstName, lastName, employeeName, preferredUsername, primaryEmail, approvedBy));
        var requiresPersonalFolder =
            sharePointSection.Any(static line => line.Contains("own personal folder", StringComparison.OrdinalIgnoreCase))
            || sharePointAccess.Any(static item => item.Contains("own personal folder", StringComparison.OrdinalIgnoreCase));
        var resolvedSharePointGroups = ResolveSharePointGroups(sharePointAccess, diagnostics);
        var personalFolderName = BuildPersonalFolderName(employeeName, firstName, lastName);
        var personalGroupName = requiresPersonalFolder && !string.IsNullOrWhiteSpace(personalFolderName)
            ? $"SP-Documents-DL-{personalFolderName}-RW"
            : string.Empty;

        return new NewUserRequest
        {
            FirstName = FieldNormalizer.NormalizeScalar(firstName),
            LastName = FieldNormalizer.NormalizeScalar(lastName),
            DisplayName = FieldNormalizer.NormalizeScalar(employeeName),
            PreferredUsername = FieldNormalizer.NormalizeScalar(preferredUsername).ToLowerInvariant(),
            TemporaryPassword = FieldNormalizer.NormalizeScalar(temporaryPassword),
            JobTitle = FieldNormalizer.NormalizeScalar(jobTitle),
            PrimaryEmail = FieldNormalizer.NormalizeScalar(primaryEmail).ToLowerInvariant(),
            SecondaryEmail = FieldNormalizer.NormalizeScalar(secondaryEmail).ToLowerInvariant(),
            SecondaryEmailMode = SecondaryEmailHandlingMode.AliasOnPrimaryUser,
            LicenseSkus = licenseSkus,
            GroupAccess = [],
            SharedMailboxAccess = sharedMailboxes,
            SharePointAccess = resolvedSharePointGroups,
            RequiresPersonalSharePointFolder = requiresPersonalFolder,
            PersonalSharePointFolderName = personalFolderName,
            PersonalSharePointPermissionGroup = personalGroupName,
            PersonalSharePointAdditionalMembers = personalFolderMembers,
            SpecialRequirements = FieldNormalizer.NormalizeScalar(specialRequirements),
            RequestApprovedBy = FieldNormalizer.NormalizeScalar(approvedBy),
            ParseDiagnostics = diagnostics
        };
    }

    private static List<string> ExtractDocxLines(string docxPath)
    {
        using var archive = ZipFile.OpenRead(docxPath);
        var entry = archive.GetEntry("word/document.xml")
                    ?? throw new InvalidDataException("The DOCX file is missing word/document.xml.");

        using var stream = entry.Open();
        var doc = XDocument.Load(stream);
        XNamespace w = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
        var lines = new List<string>();

        foreach (var p in doc.Descendants(w + "p"))
        {
            var line = string.Concat(p.Descendants(w + "t").Select(t => t.Value));
            line = FieldNormalizer.NormalizeScalar(line);
            if (!string.IsNullOrWhiteSpace(line))
            {
                lines.Add(line);
            }
        }

        return lines;
    }

    private static List<string> ExtractEmailsFromSection(IReadOnlyList<string> lines, string startKeyOrLabel, string endKeyOrLabel, List<string>? diagnostics = null)
    {
        var list = ExtractSectionItems(lines, startKeyOrLabel, endKeyOrLabel, diagnostics)
            .SelectMany(line => Regex.Matches(line, @"[A-Za-z0-9._%+\-]+@[A-Za-z0-9.\-]+\.[A-Za-z]{2,}").Select(m => m.Value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        diagnostics?.Add($"SharedMailboxAccess: found {list.Count} email target(s).");
        return list;
    }

    private static List<string> ExtractUserIdentitiesFromSection(IReadOnlyList<string> lines, List<string>? diagnostics = null)
    {
        var sections = new List<string>();
        var endCandidates = new[]
        {
            "Any Other Requirements",
            "Access to Apps/Programs",
            "Internet Setup",
            "Phone",
            "Phone Setup"
        };

        foreach (var endLabel in endCandidates)
        {
            var segment = ExtractSectionItems(lines, "Access over Users SharePoint Folder", endLabel, diagnostics);
            if (segment.Count == 0)
            {
                continue;
            }

            sections.AddRange(segment);
            break;
        }

        var noisyIdentityKeywords = new[]
        {
            "access to apps/programs", "access to apps", "apps/programs", "apps", "programs", "please state",
            "breathe", "calendar", "internet setup",
            "bookmarks", "outlook", "http://", "https://", "please add", "please state", "call staff"
        };

        var identities = sections
            .SelectMany(line => line.Split([',', ';', '|', '/'], StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries))
            .Select(FieldNormalizer.NormalizeScalar)
            .Where(item => !string.IsNullOrWhiteSpace(item))
            .Where(item => !item.Equals("NA", StringComparison.OrdinalIgnoreCase))
            .Where(item => !item.Equals("N/A", StringComparison.OrdinalIgnoreCase))
            .Where(item => !item.Equals("YES", StringComparison.OrdinalIgnoreCase))
            .Where(item => !item.Equals("NO", StringComparison.OrdinalIgnoreCase))
            .Where(item => !item.Contains("access over users sharepoint folder", StringComparison.OrdinalIgnoreCase))
            .Where(item => !noisyIdentityKeywords.Any(k => item.Contains(k, StringComparison.OrdinalIgnoreCase)))
            .Where(item => IsLikelyUserIdentity(item))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        diagnostics?.Add($"PersonalSharePointAdditionalMembers: found {identities.Count} user target(s).");
        return identities;
    }

    private static bool IsLikelyUserIdentity(string item)
    {
        if (string.IsNullOrWhiteSpace(item))
        {
            return false;
        }

        if (item.Contains(":", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (item.Contains("http", StringComparison.OrdinalIgnoreCase)
            || item.Contains("www.", StringComparison.OrdinalIgnoreCase))
        {
            return false;
        }

        if (item.Contains("@", StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        var hasLetter = item.Any(char.IsLetter);
        if (!hasLetter)
        {
            return false;
        }

        // Reject URL/token-like fragments (for example long ids from calendar links).
        if (item.Contains('_') || item.Length >= 20 || item.Any(char.IsDigit))
        {
            return false;
        }

        return true;
    }

    private static List<string> ExtractSectionItems(IReadOnlyList<string> lines, string startKeyOrLabel, string endKeyOrLabel, List<string>? diagnostics = null)
    {
        var startLabels = ResolveAliases(startKeyOrLabel);
        var endLabels = ResolveAliases(endKeyOrLabel);

        var start = lines.ToList().FindIndex(line => startLabels.Any(label => line.Contains(label, StringComparison.OrdinalIgnoreCase)));
        if (start < 0)
        {
            diagnostics?.Add($"Section '{startKeyOrLabel}': start label not found.");
            return [];
        }

        var end = lines.Count;
        for (var i = start + 1; i < lines.Count; i++)
        {
            if (endLabels.Any(label => lines[i].Contains(label, StringComparison.OrdinalIgnoreCase)))
            {
                end = i;
                break;
            }
        }

        var section = new List<string>();
        for (var i = start + 1; i < end; i++)
        {
            var candidate = FieldNormalizer.NormalizeScalar(lines[i]);
            if (!string.IsNullOrWhiteSpace(candidate))
            {
                section.Add(candidate);
            }
        }

        diagnostics?.Add($"Section '{startKeyOrLabel}': captured {section.Count} raw line(s).");
        return section;
    }

    private static string ExtractSectionText(string fullText, string startKeyOrLabel, string endKeyOrLabel, List<string>? diagnostics = null)
    {
        foreach (var startLabel in ResolveAliases(startKeyOrLabel))
        {
            foreach (var endLabel in ResolveAliases(endKeyOrLabel))
            {
                var pattern = $"{Regex.Escape(startLabel)}\\s*(?<v>[\\s\\S]*?){Regex.Escape(endLabel)}";
                var match = Regex.Match(fullText, pattern, RegexOptions.IgnoreCase);
                if (!match.Success)
                {
                    continue;
                }

                var value = match.Groups["v"].Value;
                var lines = value.Split('\n', StringSplitOptions.RemoveEmptyEntries)
                    .Select(FieldNormalizer.NormalizeScalar)
                    .Where(line => !string.IsNullOrWhiteSpace(line));
                diagnostics?.Add($"SectionText '{startKeyOrLabel}': matched between '{startLabel}' and '{endLabel}'.");
                return string.Join(" ", lines);
            }
        }

        diagnostics?.Add($"SectionText '{startKeyOrLabel}': no match.");
        return string.Empty;
    }

    private static string MatchValue(string source, string pattern, List<string>? diagnostics = null, string? field = null, string? strategy = null)
    {
        var match = Regex.Match(source, pattern, RegexOptions.IgnoreCase | RegexOptions.Singleline);
        if (match.Success)
        {
            diagnostics?.Add($"{field ?? "Field"}: matched using {strategy ?? pattern}.");
            return FieldNormalizer.NormalizeScalar(match.Groups["v"].Value);
        }

        diagnostics?.Add($"{field ?? "Field"}: no match for {strategy ?? pattern}.");
        return string.Empty;
    }

    private static List<string> NormalizeSharePointAccess(IReadOnlyList<string> sectionLines)
    {
        var noisyKeywords = new[]
        {
            "license", "mailbox", "phone", "equipment", "yes/no", "restrictions", "email",
            "apps", "requirements", "call queue", "machine", "service tag", "internet", "bookmarks",
            "na yes", "n/a yes", "yes", "no"
        };

        return sectionLines
            .SelectMany(line => line.Split([',', ';', '|', '/'], StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries))
            .Select(FieldNormalizer.NormalizeScalar)
            .Where(item => !string.IsNullOrWhiteSpace(item))
            .Where(static item => !IsPlaceholderSharePointValue(item))
            .Where(item => !item.Contains("@"))
            .Where(item => item.Length <= 50)
            .Where(item => item.Any(char.IsLetter))
            .Where(item => !noisyKeywords.Any(k => item.Contains(k, StringComparison.OrdinalIgnoreCase)))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(static x => x, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static bool IsPlaceholderSharePointValue(string item)
    {
        return item.Equals("NA", StringComparison.OrdinalIgnoreCase)
            || item.Equals("N/A", StringComparison.OrdinalIgnoreCase)
            || item.Equals("YES", StringComparison.OrdinalIgnoreCase)
            || item.Equals("NO", StringComparison.OrdinalIgnoreCase)
            || item.Equals("N", StringComparison.OrdinalIgnoreCase)
            || item.Equals("Y", StringComparison.OrdinalIgnoreCase);
    }

    private static List<string> RecoverSharePointSelections(
        IReadOnlyList<string> normalizedItems,
        IReadOnlyList<string> sectionLines,
        string fullText,
        List<string> diagnostics)
    {
        var recovered = normalizedItems.ToList();
        var sectionText = string.Join(" ", sectionLines);
        var probeText = $"{sectionText} {fullText}";

        foreach (var known in KnownSharePointSelections)
        {
            var pattern = Regex.Escape(known).Replace("\\ ", "\\s+");
            if (!Regex.IsMatch(probeText, pattern, RegexOptions.IgnoreCase))
            {
                continue;
            }

            if (recovered.Any(x => string.Equals(x, known, StringComparison.OrdinalIgnoreCase)))
            {
                continue;
            }

            recovered.Add(known);
            diagnostics.Add($"SharePointAccess: recovered known selection '{known}' from DOCX content.");
        }

        if (recovered.Count == 0)
        {
            diagnostics.Add("SharePointAccess: no selections detected in SharePoint section.");
        }

        return recovered
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(static x => x, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static List<string> ResolveSharePointGroups(IEnumerable<string> rawItems, List<string> diagnostics)
    {
        var resolved = new List<string>();
        foreach (var item in rawItems)
        {
            if (item.Contains("own personal folder", StringComparison.OrdinalIgnoreCase))
            {
                diagnostics.Add("SharePointAccess: detected personal folder request from 'Own Personal Folder'.");
                continue;
            }

            var key = item.Replace('_', ' ').Trim().ToUpperInvariant();
            if (SharePointGroupAliasMap.TryGetValue(key, out var mapped))
            {
                diagnostics.Add($"SharePointAccess: mapped '{item}' to '{mapped}'.");
                resolved.Add(mapped);
            }
            else
            {
                resolved.Add(item);
            }
        }

        return resolved
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(static x => x, StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static string BuildPersonalFolderName(string employeeName, string firstName, string lastName)
    {
        var fullName = FieldNormalizer.NormalizeScalar(employeeName);
        if (!string.IsNullOrWhiteSpace(fullName))
        {
            return fullName;
        }

        var first = FieldNormalizer.NormalizeScalar(firstName);
        var last = FieldNormalizer.NormalizeScalar(lastName);
        return FieldNormalizer.NormalizeScalar($"{first} {last}");
    }

    private static IReadOnlyList<string> ResolveAliases(string keyOrLabel)
    {
        if (SectionAliasMap.TryGetValue(keyOrLabel, out var aliases))
        {
            return aliases;
        }

        return [keyOrLabel];
    }

    private static List<string> ResolveLicenseSkus(IEnumerable<string> rawSkus, List<string> diagnostics)
    {
        var normalized = rawSkus
            .Select(sku => FieldNormalizer.NormalizeScalar(sku).Replace(' ', '_').ToUpperInvariant())
            .Where(sku => !string.IsNullOrWhiteSpace(sku))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        if (normalized.Contains("BUSINESS_STANDARD", StringComparer.OrdinalIgnoreCase)
            && normalized.Contains("EXCHANGE_ONLINE_MAILBOX", StringComparer.OrdinalIgnoreCase))
        {
            normalized.RemoveAll(static sku => sku.Equals("EXCHANGE_ONLINE_MAILBOX", StringComparison.OrdinalIgnoreCase));
            diagnostics.Add("LicenseSkus: removed EXCHANGE_ONLINE_MAILBOX because BUSINESS_STANDARD already includes Exchange Online mailbox.");
        }

        return normalized;
    }

    private static SecondaryEmailHandlingMode ParseSecondaryEmailMode(string raw)
    {
        var normalized = FieldNormalizer.NormalizeScalar(raw).Replace(" ", string.Empty, StringComparison.Ordinal).ToUpperInvariant();
        return normalized switch
        {
            "SEPARATEMAILBOXWITHDELEGATION" => SecondaryEmailHandlingMode.SeparateMailboxWithDelegation,
            "SEPARATEMAILBOX" => SecondaryEmailHandlingMode.SeparateMailboxWithDelegation,
            "SEPARATE" => SecondaryEmailHandlingMode.SeparateMailboxWithDelegation,
            "MAILBOX" => SecondaryEmailHandlingMode.SeparateMailboxWithDelegation,
            _ => SecondaryEmailHandlingMode.AliasOnPrimaryUser
        };
    }

    private static IEnumerable<string> BuildContractDiagnostics(
        string firstName,
        string lastName,
        string displayName,
        string preferredUsername,
        string primaryEmail,
        string approvedBy)
    {
        var values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            ["FirstName"] = firstName,
            ["LastName"] = lastName,
            ["DisplayName"] = displayName,
            ["PreferredUsername"] = preferredUsername,
            ["PrimaryEmail"] = primaryEmail,
            ["RequestApprovedBy"] = approvedBy
        };

        foreach (var spec in ParseContract)
        {
            if (!spec.Required)
            {
                continue;
            }

            values.TryGetValue(spec.Canonical, out var value);
            yield return string.IsNullOrWhiteSpace(value)
                ? $"Contract: required field '{spec.Canonical}' missing."
                : $"Contract: required field '{spec.Canonical}' populated.";
        }
    }

    private static string Read(IDictionary<string, string> map, string key) => map.TryGetValue(key, out var value) ? value : string.Empty;

    private static Dictionary<string, string> BuildCanonicalMap()
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var field in ParseContract)
        {
            foreach (var alias in field.Aliases)
            {
                map[alias] = field.Canonical;
            }
        }

        return map;
    }
}
