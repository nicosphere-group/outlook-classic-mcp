using System.ComponentModel;
using System.Text.Encodings.Web;
using System.Text.Json;
using ModelContextProtocol.Server;

namespace OutlookClassicMcp;

/// <summary>Runs a service call on the Outlook thread and returns its result as JSON text.</summary>
public sealed class MailRunner(MailService service, IWorkQueue queue)
{
    // Keep Japanese readable instead of \uXXXX escapes.
    public static readonly JsonSerializerOptions Json = new()
    {
        Encoder = JavaScriptEncoder.UnsafeRelaxedJsonEscaping,
        PropertyNamingPolicy = JsonNamingPolicy.SnakeCaseLower,
    };

    public async Task<string> Run<T>(Func<MailService, T> call) =>
        JsonSerializer.Serialize(await queue.RunAsync(() => call(service)), Json);
}

// Tool descriptions are what the model reads. Reading and drafting never send anything.

[McpServerToolType]
public sealed class MailTools
{
    [McpServerTool(Name = "list_accounts", ReadOnly = true, Destructive = false, OpenWorld = false)]
    [Description("List the email accounts and data files (PST etc.) set up in Outlook.")]
    public static Task<string> ListAccounts(MailRunner mail) => mail.Run(s => s.ListAccounts());

    [McpServerTool(Name = "list_folders", ReadOnly = true, Destructive = false, OpenWorld = false)]
    [Description("List mail folders with total and unread counts.")]
    public static Task<string> ListFolders(
        MailRunner mail,
        [Description("Account name or address from list_accounts; omit for all accounts.")] string? account = null) =>
        mail.Run(s => s.ListFolders(account));

    [McpServerTool(Name = "list_emails", ReadOnly = true, Destructive = false, OpenWorld = false)]
    [Description("List emails in a folder, newest first. Use offset to page when has_more is true.")]
    public static Task<string> ListEmails(
        MailRunner mail,
        [Description("inbox, sent, drafts, deleted, outbox, junk, or a path from list_folders.")] string folder = "inbox",
        [Description("Account name or address; omit for the default account.")] string? account = null,
        [Description("Only emails on or after this ISO 8601 date or date-time (local time), e.g. 2026-01-31.")] string? since = null,
        [Description("Only emails on or before this ISO 8601 date or date-time (local time).")] string? until = null,
        [Description("Only unread emails.")] bool unread_only = false,
        [Description("How many to return (1-100).")] int limit = 20,
        [Description("How many to skip, for paging.")] int offset = 0) =>
        mail.Run(s => s.ListEmails(folder, account, since, until, unread_only, limit, offset));

    [McpServerTool(Name = "search_emails", ReadOnly = true, Destructive = false, OpenWorld = false)]
    [Description(
        "Search one folder for emails, newest first. All given conditions must match. " +
        "Searches a single folder; call again with another folder (e.g. sent) if needed. " +
        "Body search can be slow; if timed_out is true the result is partial - narrow the dates.")]
    public static Task<string> SearchEmails(
        MailRunner mail,
        [Description("Text to find in subject, sender or body.")] string? query = null,
        [Description("Text to find in the subject.")] string? subject = null,
        [Description("Sender name or address.")] string? sender = null,
        [Description("Text to find in the body.")] string? body = null,
        [Description("ISO 8601 date or date-time (local time).")] string? since = null,
        [Description("ISO 8601 date or date-time (local time).")] string? until = null,
        [Description("inbox, sent, drafts, deleted, outbox, junk, or a path from list_folders.")] string folder = "inbox",
        [Description("Account name or address; omit for the default account.")] string? account = null,
        [Description("How many to return (1-100).")] int limit = 20) =>
        mail.Run(s => s.SearchEmails(query, subject, sender, body, since, until, folder, account, limit));

    [McpServerTool(Name = "read_email", ReadOnly = true, Destructive = false, OpenWorld = false)]
    [Description(
        "Read one email: sender, recipients, body text and attachment names. Does not mark it as read. " +
        "If body_truncated is true, call again with a larger max_chars only if the rest is needed.")]
    public static Task<string> ReadEmail(
        MailRunner mail,
        [Description("The id from list_emails or search_emails.")] string email_id,
        [Description("Maximum body characters to return (500-50000).")] int max_chars = 8000) =>
        mail.Run(s => s.ReadEmail(email_id, max_chars));

    [McpServerTool(Name = "list_attachments", ReadOnly = true, Destructive = false, OpenWorld = false)]
    [Description("List the attachments of an email (number, file name, size).")]
    public static Task<string> ListAttachments(MailRunner mail, string email_id) =>
        mail.Run(s => s.ListAttachments(email_id));

    [McpServerTool(Name = "save_attachment", Destructive = false, OpenWorld = false)]
    [Description(
        "Save one attachment into the tool's download folder and return its path. " +
        "Existing files are never overwritten. The file is only saved, never opened.")]
    public static Task<string> SaveAttachment(
        MailRunner mail,
        string email_id,
        [Description("The attachment number from list_attachments.")] int index) =>
        mail.Run(s => s.SaveAttachment(email_id, index));

    [McpServerTool(Name = "create_draft", Destructive = false, OpenWorld = false)]
    [Description(
        "Create a new email as a draft in Outlook. Does NOT send. " +
        "After creating, show the user who it is addressed to, the subject and the text.")]
    public static Task<string> CreateDraft(
        MailRunner mail,
        [Description("Full email addresses, one per entry.")] string[] to,
        string subject,
        string body,
        [Description("Full email addresses, one per entry.")] string[]? cc = null,
        [Description("Full email addresses, one per entry.")] string[]? bcc = null,
        [Description("Which account to send from (name or address); omit for the default.")] string? account = null,
        [Description("Full paths of files in the user's Documents, Desktop or Downloads.")] string[]? attachments = null) =>
        mail.Run(s => s.CreateDraft(to, subject, body, cc, bcc, account, attachments));

    [McpServerTool(Name = "create_reply_draft", Destructive = false, OpenWorld = false)]
    [Description("Create a reply to an email as a draft. Does NOT send.")]
    public static Task<string> CreateReplyDraft(
        MailRunner mail,
        string email_id,
        [Description("Only the new text; the original message is quoted below it automatically.")] string body,
        [Description("True to include everyone on the original email. Only when the user asked to reply to everyone.")] bool reply_all = false) =>
        mail.Run(s => s.CreateReplyDraft(email_id, body, reply_all));

    [McpServerTool(Name = "create_forward_draft", Destructive = false, OpenWorld = false)]
    [Description("Create a forward of an email as a draft. Does NOT send.")]
    public static Task<string> CreateForwardDraft(
        MailRunner mail,
        string email_id,
        [Description("Full email addresses, one per entry.")] string[] to,
        [Description("Optional text placed above the forwarded message.")] string body = "") =>
        mail.Run(s => s.CreateForwardDraft(email_id, to, body));

    [McpServerTool(Name = "update_draft", Destructive = false, OpenWorld = false)]
    [Description(
        "Change an unsent draft. Does NOT send. Omitted fields stay as they are. " +
        "to / cc / bcc replace the whole list. body replaces the whole text, including any quoted " +
        "original in a reply or forward - include the quoted part if it must stay.")]
    public static Task<string> UpdateDraft(
        MailRunner mail,
        string draft_id,
        string[]? to = null,
        string[]? cc = null,
        string[]? bcc = null,
        string? subject = null,
        string? body = null,
        [Description("Full paths of files to add, from the user's Documents, Desktop or Downloads.")] string[]? add_attachments = null) =>
        mail.Run(s => s.UpdateDraft(draft_id, to, cc, bcc, subject, body, add_attachments));
}

/// <summary>Registered only when sending is allowed.</summary>
[McpServerToolType]
public sealed class SendTools
{
    [McpServerTool(Name = "send_draft")]
    [Description(
        "Ask the user to confirm, then send an existing draft. Call this only when the user has " +
        "clearly asked to send this draft. A window appears on the user's screen showing the " +
        "recipients, subject, text and attachments; the email is sent only if the user presses Send " +
        "there. This call waits for their answer (up to a few minutes). Tell the user the outcome.")]
    public static Task<string> SendDraft(MailRunner mail, string draft_id) =>
        mail.Run(s => s.SendDraft(draft_id));
}

/// <summary>
/// Registered only when diagnostics are switched on. Used to find out how long Claude waits
/// for a tool to answer, which decides how long the confirmation window may stay open.
/// </summary>
[McpServerToolType]
public sealed class DiagnosticTools
{
    private const int MaxSeconds = 900;

    [McpServerTool(Name = "wait_test", ReadOnly = true, Destructive = false, OpenWorld = false)]
    [Description("Diagnostic: waits the given number of seconds, then answers. Touches nothing.")]
    public static async Task<string> WaitTest(
        [Description("Seconds to wait (1-900).")] int seconds,
        CancellationToken cancellationToken)
    {
        seconds = Math.Clamp(seconds, 1, MaxSeconds);
        await Task.Delay(TimeSpan.FromSeconds(seconds), cancellationToken);
        return $"Answered after {seconds} seconds.";
    }
}
