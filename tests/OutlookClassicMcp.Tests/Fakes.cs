namespace OutlookClassicMcp.Tests;

/// <summary>In-memory stand-in for Outlook.</summary>
public sealed class FakeBackend(string accountType = "exchange") : IMailBackend
{
    private int _next;

    public Dictionary<string, DraftSnapshot> Drafts { get; } = [];
    public List<DraftSnapshot> Sent { get; } = [];
    public HashSet<string> SentIds { get; } = [];
    public List<string> Displayed { get; } = [];
    public Dictionary<string, EmailDetails> Inbox { get; } = [];
    public (DateTimeOffset? Since, DateTimeOffset? Until, int Limit, int Offset) LastList { get; private set; }

    public void AddInbox(
        string id, string subject = "件名", string body = "本文",
        IReadOnlyList<string>? cc = null, IReadOnlyList<AttachmentInfo>? attachments = null) =>
        Inbox[id] = new EmailDetails(
            new EmailSummary(id, subject, new Person("田中 太郎", "tanaka@example.co.jp"),
                "2026-01-10T09:00:00+09:00", true, attachments?.Count ?? 0),
            [], cc ?? [], [], false, body, attachments ?? []);

    private static List<Person> People(IEnumerable<string> addresses) =>
        addresses.Select(a => new Person(a, a)).ToList();

    private string NewDraft(
        IReadOnlyList<Person> to, IReadOnlyList<Person> cc, IReadOnlyList<Person> bcc,
        string subject, string body, IReadOnlyList<AttachmentInfo>? attachments = null)
    {
        var id = $"store.draft{++_next}";
        Drafts[id] = new DraftSnapshot(
            "自分", "me@example.com", accountType, to, cc, bcc, subject, body, attachments ?? []);
        return id;
    }

    private DraftSnapshot Draft(string id)
    {
        if (SentIds.Contains(id) || Inbox.ContainsKey(id))
        {
            throw new NotADraftException("not a draft");
        }
        return Drafts.TryGetValue(id, out var draft) ? draft : throw new NotFoundException("not found");
    }

    public AccountList ListAccounts() =>
        new([new AccountInfo("自分", "me@example.com", accountType, true)], ["me@example.com"]);

    public IReadOnlyList<FolderInfo> ListFolders(string? account) =>
        [new FolderInfo("me@example.com\\受信トレイ", "受信トレイ", Inbox.Count, Inbox.Count)];

    public EmailPage ListEmails(
        string folder, string? account, DateTimeOffset? since, DateTimeOffset? until,
        bool unreadOnly, int limit, int offset)
    {
        LastList = (since, until, limit, offset);
        return new EmailPage(Inbox.Values.Select(d => d.Summary).Skip(offset).Take(limit).ToList(), false);
    }

    public EmailPage SearchEmails(
        string folder, string? account, SearchTerms terms, DateTimeOffset? since,
        DateTimeOffset? until, int limit) =>
        new(Inbox.Values.Select(d => d.Summary).Take(limit).ToList(), false);

    public EmailDetails ReadEmail(string emailId) =>
        Inbox.TryGetValue(emailId, out var details) ? details : throw new NotFoundException("not found");

    public IReadOnlyList<AttachmentInfo> ListAttachments(string emailId) => ReadEmail(emailId).Attachments;

    public void SaveAttachment(string emailId, int index, string destinationPath) =>
        File.WriteAllText(destinationPath, "data");

    public string CreateDraft(
        IReadOnlyList<string> to, IReadOnlyList<string> cc, IReadOnlyList<string> bcc,
        string subject, string body, string? account, IReadOnlyList<string> attachments) =>
        NewDraft(People(to), People(cc), People(bcc), subject, body,
            attachments.Select((p, i) => new AttachmentInfo(i + 1, Path.GetFileName(p), new FileInfo(p).Length)).ToList());

    public string CreateReplyDraft(string emailId, string body, bool replyAll)
    {
        var original = ReadEmail(emailId);
        return NewDraft(
            [original.Summary.From], replyAll ? People(original.Cc) : [], [],
            "RE: " + original.Summary.Subject, body + "\n> " + original.Body);
    }

    public string CreateForwardDraft(string emailId, IReadOnlyList<string> to, string body) =>
        NewDraft(People(to), [], [], "FW: " + ReadEmail(emailId).Summary.Subject, body);

    public void UpdateDraft(string draftId, DraftChanges changes, IReadOnlyList<string> addAttachments)
    {
        var d = Draft(draftId);
        Drafts[draftId] = d with
        {
            To = changes.To is null ? d.To : People(changes.To),
            Cc = changes.Cc is null ? d.Cc : People(changes.Cc),
            Bcc = changes.Bcc is null ? d.Bcc : People(changes.Bcc),
            Subject = changes.Subject ?? d.Subject,
            Body = changes.Body ?? d.Body,
        };
    }

    public DraftSnapshot GetDraft(string draftId) => Draft(draftId);

    public void Send(string draftId)
    {
        Sent.Add(Draft(draftId));
        SentIds.Add(draftId);
        Drafts.Remove(draftId);
    }

    public void Display(string draftId)
    {
        Draft(draftId);
        Displayed.Add(draftId);
    }
}

/// <summary>Stands in for the human at the confirmation window.</summary>
public sealed class FakeConfirmer(Decision decision) : IConfirmer
{
    public Action? During { get; set; }
    public List<(DraftSnapshot Snapshot, IReadOnlyList<Warning> Warnings, int Timeout)> Asked { get; } = [];

    public Decision Ask(DraftSnapshot snapshot, IReadOnlyList<Warning> warnings, int timeoutSeconds)
    {
        Asked.Add((snapshot, warnings, timeoutSeconds));
        During?.Invoke();
        return decision;
    }
}

/// <summary>A service wired to fakes, with its own temporary folders.</summary>
public sealed class Rig : IDisposable
{
    public string Root { get; } = Path.Combine(Path.GetTempPath(), "ocm-tests-" + Guid.NewGuid().ToString("N"));
    public Config Config { get; }
    public FakeBackend Backend { get; }
    public FakeConfirmer Confirmer { get; }
    public MailService Service { get; }
    public string AttachDir => Config.AttachRoots[0];
    public string AuditFile => Path.Combine(Config.AuditDir, "audit.log");

    public Rig(Decision decision = Decision.Send, string accountType = "exchange", bool allowSend = true)
    {
        Config = new Config
        {
            AllowSend = allowSend,
            ConfirmTimeout = 30,
            SaveDir = Path.Combine(Root, "saved"),
            AttachRoots = [Path.Combine(Root, "Documents")],
            AuditDir = Path.Combine(Root, "audit"),
        };
        Directory.CreateDirectory(AttachDir);
        Backend = new FakeBackend(accountType);
        Confirmer = new FakeConfirmer(decision);
        Service = new MailService(Backend, Confirmer, new AuditLog(Config.AuditDir), Config);
    }

    public void Dispose()
    {
        try
        {
            Directory.Delete(Root, recursive: true);
        }
        catch (IOException)
        {
            // Leftover temp files are harmless.
        }
    }
}
