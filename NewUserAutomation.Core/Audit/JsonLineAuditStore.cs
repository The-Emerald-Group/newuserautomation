using System.Text.Json;

namespace NewUserAutomation.Core.Audit;

public sealed class JsonLineAuditStore
{
    public async Task AppendAsync(string filePath, AuditRecord record, CancellationToken cancellationToken = default)
    {
        var folder = Path.GetDirectoryName(filePath);
        if (!string.IsNullOrWhiteSpace(folder))
        {
            Directory.CreateDirectory(folder);
        }

        var line = JsonSerializer.Serialize(record);
        await File.AppendAllTextAsync(filePath, line + Environment.NewLine, cancellationToken);
    }
}
