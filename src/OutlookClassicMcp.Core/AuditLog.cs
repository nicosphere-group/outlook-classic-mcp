using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;

namespace OutlookClassicMcp;

/// <summary>Append-only record of send decisions. Never contains the message body.</summary>
public sealed class AuditLog(string directory)
{
    private static readonly JsonSerializerOptions Json = new()
    {
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
    };

    public string FilePath { get; } = Path.Combine(directory, "audit.log");

    public void Record(string outcome, DraftSnapshot snapshot)
    {
        var entry = new Dictionary<string, object>
        {
            ["time"] = DateTimeOffset.Now.ToString("yyyy-MM-ddTHH:mm:sszzz"),
            ["event"] = outcome,
            ["from"] = snapshot.SenderAddress.Length > 0 ? snapshot.SenderAddress : snapshot.SenderName,
            ["to"] = snapshot.To.Select(p => p.Address).ToArray(),
            ["cc"] = snapshot.Cc.Select(p => p.Address).ToArray(),
            ["bcc"] = snapshot.Bcc.Select(p => p.Address).ToArray(),
            ["subject"] = snapshot.Subject,
            ["attachments"] = snapshot.Attachments.Select(a => a.Filename).ToArray(),
        };
        try
        {
            Directory.CreateDirectory(directory);
            File.AppendAllText(FilePath, JsonSerializer.Serialize(entry, Json) + "\n", new UTF8Encoding(false));
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            // A logging failure must not turn a cancelled send into an error, and a
            // completed send cannot be undone, so recording is best effort.
        }
    }
}
