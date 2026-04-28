using System.Text.Json.Serialization;

namespace NewUserAutomation.App.Services;

public sealed class LiveRunState
{
    public string RunId { get; set; } = Guid.NewGuid().ToString("N");
    public string UserUpn { get; set; } = string.Empty;
    public string OperatorAccount { get; set; } = string.Empty;
    public string TenantId { get; set; } = string.Empty;
    public DateTimeOffset StartedUtc { get; set; } = DateTimeOffset.UtcNow;
    public DateTimeOffset UpdatedUtc { get; set; } = DateTimeOffset.UtcNow;
    public List<LiveRunStepState> Steps { get; set; } = [];
}

public sealed class LiveRunStepState
{
    public string Id { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string Command { get; set; } = string.Empty;
    public string Argument { get; set; } = string.Empty;
    public string Status { get; set; } = "Pending";
    public string Detail { get; set; } = string.Empty;
    public DateTimeOffset UpdatedUtc { get; set; } = DateTimeOffset.UtcNow;

    [JsonIgnore]
    public bool IsCompleted =>
        string.Equals(Status, "Completed", StringComparison.OrdinalIgnoreCase)
        || string.Equals(Status, "Skipped", StringComparison.OrdinalIgnoreCase);
}

public sealed record LiveExecutionResult(bool Success, string Detail);
