using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace OutlookClassicMcp;

public sealed record Person(string Name, string Address)
{
    public string Display() =>
        Name.Length > 0 && Address.Length > 0 && Name != Address
            ? $"{Name} <{Address}>"
            : Address.Length > 0 ? Address : Name;
}

/// <param name="Index">1-based, as in Outlook.</param>
/// <param name="Size">Bytes.</param>
public sealed record AttachmentInfo(int Index, string Filename, long Size);

/// <summary>Everything a human needs to see before a draft is sent.</summary>
public sealed record DraftSnapshot(
    string SenderName,
    string SenderAddress,
    string AccountType, // "exchange" | "imap" | "pop3" | "other"
    IReadOnlyList<Person> To,
    IReadOnlyList<Person> Cc,
    IReadOnlyList<Person> Bcc,
    string Subject,
    string Body,
    IReadOnlyList<AttachmentInfo> Attachments)
{
    public IReadOnlyList<Person> Recipients() => [.. To, .. Cc, .. Bcc];

    /// <summary>Stable hash of what would be sent; detects edits made during confirmation.</summary>
    public string Fingerprint() =>
        Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(JsonSerializer.Serialize(this))));
}

public sealed record AccountInfo(string Name, string Address, string Type, bool IsDefault);

public sealed record AccountList(IReadOnlyList<AccountInfo> Accounts, IReadOnlyList<string> DataFiles);

public sealed record FolderInfo(string Path, string Name, int Total, int Unread);

public sealed record EmailSummary(
    string Id,
    string Subject,
    Person From,
    string? Received, // ISO 8601 with offset
    bool Unread,
    int AttachmentCount);

/// <param name="TimedOut">The scan stopped early because it took too long.</param>
public sealed record EmailPage(IReadOnlyList<EmailSummary> Emails, bool HasMore, bool TimedOut = false);

public sealed record EmailDetails(
    EmailSummary Summary,
    IReadOnlyList<string> To,
    IReadOnlyList<string> Cc,
    IReadOnlyList<string> Bcc,
    bool IsDraft,
    string Body,
    IReadOnlyList<AttachmentInfo> Attachments);

public sealed record SearchTerms(string? Query, string? Subject, string? Sender, string? Body);

/// <summary>A null field means "leave as it is".</summary>
public sealed record DraftChanges(
    IReadOnlyList<string>? To,
    IReadOnlyList<string>? Cc,
    IReadOnlyList<string>? Bcc,
    string? Subject,
    string? Body);

public sealed record Warning(string Code, int Count = 0, IReadOnlyList<string>? Domains = null);

public enum Decision
{
    Send,
    Cancel,
    Open,
    Timeout,
    Failed, // the window could not be shown
}

// ---- what the tools return ----

public sealed record FolderListResult(IReadOnlyList<FolderInfo> Folders);

public sealed record EmailListResult(
    string Notice, IReadOnlyList<EmailSummary> Emails, bool HasMore, bool TimedOut);

public sealed record ReadEmailResult(
    string Notice,
    string Id,
    string Subject,
    Person From,
    string? Received,
    bool Unread,
    IReadOnlyList<string> To,
    IReadOnlyList<string> Cc,
    IReadOnlyList<string> Bcc,
    bool IsDraft,
    string Body,
    int BodyTotalChars,
    bool BodyTruncated,
    IReadOnlyList<AttachmentInfo> Attachments);

public sealed record AttachmentListResult(string Notice, IReadOnlyList<AttachmentInfo> Attachments);

public sealed record SaveAttachmentResult(string SavedTo, string Note);

public sealed record DraftPreview(
    string From,
    IReadOnlyList<string> To,
    IReadOnlyList<string> Cc,
    IReadOnlyList<string> Bcc,
    string Subject,
    string Body,
    IReadOnlyList<AttachmentInfo> Attachments);

public sealed record DraftResult(string DraftId, string Status, DraftPreview Draft);

public sealed record SendResult(bool Sent, string Outcome, string Message);
