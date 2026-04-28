using System.Text.Json;
using System.IO;

namespace NewUserAutomation.App.Services;

public sealed class LiveRunLogger
{
    public void Append(string filePath, LiveRunLogEvent logEvent)
    {
        var folder = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrWhiteSpace(folder))
        {
            Directory.CreateDirectory(folder);
        }

        var line = JsonSerializer.Serialize(logEvent);
        File.AppendAllText(filePath, line + Environment.NewLine);
    }
}

public sealed class LiveRunLogEvent
{
    public string RunId { get; init; } = string.Empty;
    public DateTimeOffset TimestampUtc { get; init; } = DateTimeOffset.UtcNow;
    public string EventType { get; init; } = string.Empty;
    public string UserUpn { get; init; } = string.Empty;
    public string TenantId { get; init; } = string.Empty;
    public string OperatorAccount { get; init; } = string.Empty;
    public string StepId { get; init; } = string.Empty;
    public string StepDescription { get; init; } = string.Empty;
    public string Command { get; init; } = string.Empty;
    public string Argument { get; init; } = string.Empty;
    public string Status { get; init; } = string.Empty;
    public string Detail { get; init; } = string.Empty;
    public string CheckpointPath { get; init; } = string.Empty;
    public string MachineName { get; init; } = Environment.MachineName;
}
