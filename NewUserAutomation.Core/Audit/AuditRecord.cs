namespace NewUserAutomation.Core.Audit;

public sealed record AuditRecord(
    string RunId,
    DateTimeOffset TimestampUtc,
    string OperatorUpn,
    string InputHash,
    IReadOnlyList<string> PlannedActions,
    IReadOnlyList<string> Outcomes);
