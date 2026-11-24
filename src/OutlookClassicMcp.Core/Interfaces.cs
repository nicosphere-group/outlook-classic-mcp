namespace OutlookClassicMcp;

/// <summary>The boundary to Outlook. Tests supply an in-memory fake.</summary>
public interface IMailBackend
{
    AccountList ListAccounts();

    IReadOnlyList<FolderInfo> ListFolders(string? account);

    EmailPage ListEmails(
        string folder, string? account, DateTimeOffset? since, DateTimeOffset? until,
        bool unreadOnly, int limit, int offset);

    EmailPage SearchEmails(
        string folder, string? account, SearchTerms terms, DateTimeOffset? since,
        DateTimeOffset? until, int limit);

    /// <summary>Full details; Body is the complete plain text (the service truncates it).</summary>
    EmailDetails ReadEmail(string emailId);

    IReadOnlyList<AttachmentInfo> ListAttachments(string emailId);

    void SaveAttachment(string emailId, int index, string destinationPath);

    string CreateDraft(
        IReadOnlyList<string> to, IReadOnlyList<string> cc, IReadOnlyList<string> bcc,
        string subject, string body, string? account, IReadOnlyList<string> attachments);

    string CreateReplyDraft(string emailId, string body, bool replyAll);

    string CreateForwardDraft(string emailId, IReadOnlyList<string> to, string body);

    /// <summary>Must throw <see cref="NotADraftException"/> unless the item is an unsent draft.</summary>
    void UpdateDraft(string draftId, DraftChanges changes, IReadOnlyList<string> addAttachments);

    /// <summary>Must throw <see cref="NotADraftException"/> unless the item is an unsent draft.</summary>
    DraftSnapshot GetDraft(string draftId);

    void Send(string draftId);

    /// <summary>Opens the draft in an Outlook window. Does not send.</summary>
    void Display(string draftId);
}

/// <summary>
/// Asks a human whether to send. The model has no way to answer on their behalf.
/// Anything other than <see cref="Decision.Send"/> means: do not send.
/// </summary>
public interface IConfirmer
{
    Decision Ask(DraftSnapshot snapshot, IReadOnlyList<Warning> warnings, int timeoutSeconds);
}

/// <summary>Runs work on the one thread that is allowed to touch Outlook, in order.</summary>
public interface IWorkQueue
{
    Task<T> RunAsync<T>(Func<T> work);
}
