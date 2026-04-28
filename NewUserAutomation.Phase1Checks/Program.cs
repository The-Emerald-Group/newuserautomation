using System.Security.Cryptography;
using System.Text;
using NewUserAutomation.Core.Audit;
using NewUserAutomation.Core.Execution;
using NewUserAutomation.Core.Parsing;
using NewUserAutomation.Core.Security;
using NewUserAutomation.Core.Validation;

var root = ResolveRoot();
var auditPath = Path.Combine(root, "artifacts", "audit", "phase1-audit.jsonl");

var parser = new FormParser();
var fixturePaths = new[]
{
    Path.Combine(root, "fixtures", "forms", "happy_path_form.txt"),
    Path.Combine(root, "fixtures", "forms", "variant_alias_a.txt"),
    Path.Combine(root, "fixtures", "forms", "variant_alias_b.txt")
};

NewUserAutomation.Core.Models.NewUserRequest? request = null;
IReadOnlyList<ExecutionStep>? steps = null;

foreach (var fixturePath in fixturePaths)
{
    var parsed = parser.ParseFromKeyValueText(await File.ReadAllTextAsync(fixturePath));
    var validationReport = new ValidationEngine().Validate(parsed);
    Assert(validationReport.IsValid, $"{Path.GetFileName(fixturePath)} validation failed: {string.Join(" | ", validationReport.Errors)}");

    var planned = new DryRunPlanner().BuildPlan(parsed);
    Assert(planned.Count >= 3, $"{Path.GetFileName(fixturePath)} expected at least three dry-run steps.");
    Assert(parsed.ParseDiagnostics.Count > 0, $"{Path.GetFileName(fixturePath)} expected parser diagnostics.");
    Assert(planned.Select(x => x.Action).Distinct(StringComparer.OrdinalIgnoreCase).Any(x => x.StartsWith("Identity.", StringComparison.OrdinalIgnoreCase)),
        $"{Path.GetFileName(fixturePath)} expected grouped action names.");

    request = parsed;
    steps = planned;
}

Assert(request is not null && steps is not null, "No fixtures were processed.");
var finalSteps = steps!;

var permissions = new PermissionMatrix().Evaluate(
    selectedActions: finalSteps.Select(s => s.Action).ToArray(),
    grantedGraphScopes: ["User.ReadWrite.All", "Group.ReadWrite.All"],
    grantedExchangeRoles: ["Mailbox.Permission.Assign"],
    grantedPnPRights: ["Site.Member.Write"]);

Assert(permissions.CanExecute, $"Missing permissions: {string.Join(", ", permissions.MissingPermissions)}");

var record = new AuditRecord(
    RunId: Guid.NewGuid().ToString("N"),
    TimestampUtc: DateTimeOffset.UtcNow,
    OperatorUpn: "operator@emerald-group.example",
    InputHash: Sha256(await File.ReadAllTextAsync(fixturePaths[0])),
    PlannedActions: finalSteps.Select(s => s.Action).ToArray(),
    Outcomes: finalSteps.Select(s => $"{s.Action}:DryRun").ToArray());

await new JsonLineAuditStore().AppendAsync(auditPath, record);

Console.WriteLine("Phase 1 checks passed.");
Console.WriteLine($"Dry-run action count: {finalSteps.Count}");
Console.WriteLine($"Audit output: {auditPath}");

static void Assert(bool condition, string message)
{
    if (!condition)
    {
        throw new InvalidOperationException(message);
    }
}

static string Sha256(string input)
{
    var hash = SHA256.HashData(Encoding.UTF8.GetBytes(input));
    return Convert.ToHexString(hash);
}

static string ResolveRoot()
{
    var current = AppContext.BaseDirectory;
    for (var i = 0; i < 8; i++)
    {
        var candidate = current;
        for (var up = 0; up < i; up++)
        {
            candidate = Path.GetFullPath(Path.Combine(candidate, ".."));
        }
        if (Directory.Exists(Path.Combine(candidate, "docs")) && Directory.Exists(Path.Combine(candidate, "NewUserAutomation.Core")))
        {
            return candidate;
        }
    }

    throw new DirectoryNotFoundException("Could not resolve workspace root.");
}
