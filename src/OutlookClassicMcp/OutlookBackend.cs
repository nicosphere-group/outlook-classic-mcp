using System.Diagnostics;
using System.Runtime.InteropServices;

namespace OutlookClassicMcp;

/// <summary>
/// Classic Outlook through COM, late-bound so no Office libraries are needed at build time.
/// Every method must be called on the <see cref="StaWorkQueue"/> thread.
/// Mail only; calendar support would be a separate backend next to this one.
/// </summary>
public sealed class OutlookBackend : IMailBackend
{
    private const int OlMail = 43;
    private const int OlFormatHtml = 2;
    private const int OlMailItemType = 0;
    private const int MaxScan = 5000;
    private const int MaxFolders = 500;
    private const int MaxFolderDepth = 6;
    private const string SmtpAddressProperty = "http://schemas.microsoft.com/mapi/proptag/0x39FE001F";

    /// <summary>A scan (list or search) gives up after this long and reports a partial result.</summary>
    private static readonly TimeSpan ScanBudget = TimeSpan.FromSeconds(20);

    private static readonly Dictionary<string, int> WellKnownFolders = new(StringComparer.OrdinalIgnoreCase)
    {
        ["inbox"] = 6,
        ["sent"] = 5,
        ["drafts"] = 16,
        ["deleted"] = 3,
        ["outbox"] = 4,
        ["junk"] = 23,
    };

    private dynamic? _application;

    // ---- connection and lookup ----

    private dynamic Application()
    {
        if (_application is null)
        {
            var type = Type.GetTypeFromProgID("Outlook.Application") ?? throw OutlookErrors.NotInstalled();
            _application = Activator.CreateInstance(type) ?? throw OutlookErrors.NotInstalled();
        }
        return _application;
    }

    private dynamic Namespace() => Application().GetNamespace("MAPI");

    private T Guard<T>(Func<T> work)
    {
        try
        {
            return work();
        }
        catch (COMException ex)
        {
            _application = null; // reconnect on the next call
            throw OutlookErrors.Translate(ex);
        }
    }

    private void Guard(Action work) => Guard<object?>(() =>
    {
        work();
        return null;
    });

    private static dynamic FindStore(dynamic ns, string storeKey)
    {
        dynamic stores = ns.Stores;
        int count = stores.Count;
        for (var i = 1; i <= count; i++)
        {
            dynamic store = stores.Item(i);
            if (OutlookHelpers.StoreKey((string)store.StoreID) == storeKey)
            {
                return store;
            }
        }
        throw new NotFoundException("That email's mailbox is no longer available. List the emails again.");
    }

    private dynamic Item(string itemId)
    {
        var (storeKey, entryId) = OutlookHelpers.SplitId(itemId);
        dynamic ns = Namespace();
        dynamic store = FindStore(ns, storeKey);
        string storeId = store.StoreID;
        try
        {
            return ns.GetItemFromID(entryId, storeId);
        }
        catch (COMException)
        {
            throw new NotFoundException(
                "That email was not found. It may have been moved or deleted; list the emails again.");
        }
    }

    private dynamic Draft(string draftId)
    {
        dynamic item = Item(draftId);
        if ((int)item.Class != OlMail || (bool)item.Sent)
        {
            throw new NotADraftException("That item is not an unsent draft, so it cannot be changed or sent.");
        }
        return item;
    }

    private static dynamic? AccountNamed(dynamic ns, string name)
    {
        var wanted = name.Trim();
        dynamic accounts = ns.Accounts;
        int count = accounts.Count;
        for (var i = 1; i <= count; i++)
        {
            dynamic account = accounts.Item(i);
            if (wanted.Equals((string?)account.DisplayName, StringComparison.OrdinalIgnoreCase) ||
                wanted.Equals((string?)account.SmtpAddress, StringComparison.OrdinalIgnoreCase))
            {
                return account;
            }
        }
        return null;
    }

    private static dynamic StoreFor(dynamic ns, string? account)
    {
        if (account is null)
        {
            return ns.DefaultStore;
        }
        dynamic? found = AccountNamed(ns, account);
        if (found is not null)
        {
            return found.DeliveryStore;
        }
        dynamic stores = ns.Stores;
        int count = stores.Count;
        for (var i = 1; i <= count; i++)
        {
            dynamic store = stores.Item(i);
            if (account.Trim().Equals((string?)store.DisplayName, StringComparison.OrdinalIgnoreCase))
            {
                return store;
            }
        }
        throw new NotFoundException($"No account or data file named '{account}' was found.");
    }

    private static dynamic ResolveFolder(dynamic ns, string? folder, string? account)
    {
        var name = string.IsNullOrWhiteSpace(folder) ? "inbox" : folder.Trim();
        if (WellKnownFolders.TryGetValue(name, out var folderType))
        {
            dynamic store = StoreFor(ns, account);
            try
            {
                return store.GetDefaultFolder(folderType);
            }
            catch (COMException)
            {
                throw new NotFoundException($"This account has no '{name}' folder.");
            }
        }

        var parts = name.Split('\\', StringSplitOptions.RemoveEmptyEntries).ToList();
        if (parts.Count == 0)
        {
            throw new InvalidInputException("The folder name is empty.");
        }
        dynamic? current = null;
        dynamic stores = ns.Stores;
        int count = stores.Count;
        for (var i = 1; i <= count && current is null; i++)
        {
            dynamic store = stores.Item(i);
            if (parts[0].Equals((string?)store.DisplayName, StringComparison.OrdinalIgnoreCase))
            {
                current = store.GetRootFolder();
                parts.RemoveAt(0);
            }
        }
        current ??= StoreFor(ns, account).GetRootFolder();
        foreach (var part in parts)
        {
            try
            {
                current = current!.Folders.Item(part);
            }
            catch (COMException)
            {
                throw new NotFoundException($"The folder '{folder}' was not found. Use a path from list_folders.");
            }
        }
        return current!;
    }

    // ---- converting Outlook items ----

    private static Person Sender(dynamic item)
    {
        string name = (string?)item.SenderName ?? "";
        string address = (string?)item.SenderEmailAddress ?? "";
        if ((string?)item.SenderEmailType == "EX")
        {
            // Exchange gives an internal directory name; ask for the normal email address.
            try
            {
                string? smtp = item.Sender.GetExchangeUser().PrimarySmtpAddress;
                if (!string.IsNullOrEmpty(smtp))
                {
                    address = smtp;
                }
            }
            catch (Exception)
            {
                // Keep what Outlook gave us.
            }
        }
        return new Person(name, address);
    }

    private static (List<Person> To, List<Person> Cc, List<Person> Bcc) Recipients(dynamic item)
    {
        var to = new List<Person>();
        var cc = new List<Person>();
        var bcc = new List<Person>();
        dynamic recipients = item.Recipients;
        int count = recipients.Count;
        for (var i = 1; i <= count; i++)
        {
            dynamic r = recipients.Item(i);
            string address = (string?)r.Address ?? "";
            if (!address.Contains('@'))
            {
                try
                {
                    string? smtp = r.PropertyAccessor.GetProperty(SmtpAddressProperty);
                    if (!string.IsNullOrEmpty(smtp))
                    {
                        address = smtp;
                    }
                }
                catch (Exception)
                {
                    try
                    {
                        string? smtp = r.AddressEntry.GetExchangeUser().PrimarySmtpAddress;
                        if (!string.IsNullOrEmpty(smtp))
                        {
                            address = smtp;
                        }
                    }
                    catch (Exception)
                    {
                        // Keep what Outlook gave us.
                    }
                }
            }
            var person = new Person((string?)r.Name ?? "", address);
            switch ((int)r.Type)
            {
                case 2: cc.Add(person); break;
                case 3: bcc.Add(person); break;
                default: to.Add(person); break;
            }
        }
        return (to, cc, bcc);
    }

    private static DateTimeOffset? When(dynamic item)
    {
        try
        {
            if (OutlookHelpers.ToLocal((DateTime)item.ReceivedTime) is { } received)
            {
                return received;
            }
        }
        catch (Exception)
        {
            // Fall through to the modification time.
        }
        try
        {
            return OutlookHelpers.ToLocal((DateTime)item.LastModificationTime);
        }
        catch (Exception)
        {
            return null;
        }
    }

    private static EmailSummary Summary(dynamic item, DateTimeOffset? when)
    {
        string storeId = item.Parent.StoreID;
        string entryId = item.EntryID;
        Person sender = Sender(item);
        return new EmailSummary(
            OutlookHelpers.MakeId(storeId, entryId),
            (string?)item.Subject ?? "",
            sender,
            when is null ? null : OutlookHelpers.Iso(when),
            (bool)item.UnRead,
            (int)item.Attachments.Count);
    }

    private static List<AttachmentInfo> Attachments(dynamic item)
    {
        var result = new List<AttachmentInfo>();
        dynamic attachments = item.Attachments;
        int count = attachments.Count;
        for (var i = 1; i <= count; i++)
        {
            dynamic a = attachments.Item(i);
            string name = "";
            try
            {
                name = (string?)a.FileName ?? "";
            }
            catch (Exception)
            {
                // Embedded items have no file name.
            }
            if (name.Length == 0)
            {
                try
                {
                    name = (string?)a.DisplayName ?? "";
                }
                catch (Exception)
                {
                    // Use the fallback below.
                }
            }
            if (name.Length == 0)
            {
                name = $"attachment-{i}";
            }
            result.Add(new AttachmentInfo(i, name, (long)(int)a.Size));
        }
        return result;
    }

    /// <summary>
    /// Walks items newest-first and filters dates in code, so the result does not depend on
    /// the PC's regional date format. Stops at the item limit or the time budget.
    /// </summary>
    private static EmailPage Collect(
        dynamic items, DateTimeOffset? since, DateTimeOffset? until, int limit, int offset)
    {
        var clock = Stopwatch.StartNew();
        items.Sort("[ReceivedTime]", true);
        var emails = new List<EmailSummary>();
        var skipped = 0;
        var scanned = 0;
        var hasMore = false;
        var timedOut = false;

        dynamic? item = items.GetFirst();
        while (item is not null)
        {
            if (++scanned > MaxScan)
            {
                hasMore = true;
                break;
            }
            if (clock.Elapsed > ScanBudget)
            {
                hasMore = true;
                timedOut = true;
                break;
            }
            if ((int)item.Class == OlMail)
            {
                DateTimeOffset? when = When(item);
                if (when is not null && since is not null && when < since)
                {
                    break; // sorted newest first, so nothing older can match
                }
                if (!(when is not null && until is not null && when > until))
                {
                    if (skipped < offset)
                    {
                        skipped++;
                    }
                    else if (emails.Count >= limit)
                    {
                        hasMore = true;
                        break;
                    }
                    else
                    {
                        EmailSummary summary = Summary(item, when);
                        emails.Add(summary);
                    }
                }
            }
            item = items.GetNext();
        }
        return new EmailPage(emails, hasMore, timedOut);
    }

    // ---- reading ----

    public AccountList ListAccounts() => Guard(() =>
    {
        dynamic ns = Namespace();
        var accounts = new List<AccountInfo>();
        dynamic all = ns.Accounts;
        int count = all.Count;
        for (var i = 1; i <= count; i++)
        {
            dynamic a = all.Item(i);
            accounts.Add(new AccountInfo(
                (string?)a.DisplayName ?? "",
                (string?)a.SmtpAddress ?? "",
                AccountType((int)a.AccountType),
                i == 1));
        }
        var dataFiles = new List<string>();
        dynamic stores = ns.Stores;
        int storeCount = stores.Count;
        for (var i = 1; i <= storeCount; i++)
        {
            dataFiles.Add((string?)stores.Item(i).DisplayName ?? "");
        }
        return new AccountList(accounts, dataFiles);
    });

    private static string AccountType(int type) => type switch
    {
        0 => "exchange",
        1 => "imap",
        2 => "pop3",
        _ => "other",
    };

    public IReadOnlyList<FolderInfo> ListFolders(string? account) => Guard<IReadOnlyList<FolderInfo>>(() =>
    {
        dynamic ns = Namespace();
        var result = new List<FolderInfo>();
        if (account is not null)
        {
            dynamic store = StoreFor(ns, account);
            Walk(store.GetRootFolder(), 1, result);
            return result;
        }
        dynamic stores = ns.Stores;
        int count = stores.Count;
        for (var i = 1; i <= count; i++)
        {
            Walk(stores.Item(i).GetRootFolder(), 1, result);
        }
        return result;
    });

    private static void Walk(dynamic parent, int depth, List<FolderInfo> result)
    {
        if (depth > MaxFolderDepth)
        {
            return;
        }
        dynamic children = parent.Folders;
        int count = children.Count;
        for (var i = 1; i <= count && result.Count < MaxFolders; i++)
        {
            try
            {
                dynamic folder = children.Item(i);
                if ((int)folder.DefaultItemType == OlMailItemType)
                {
                    string path = folder.FolderPath;
                    result.Add(new FolderInfo(
                        path.TrimStart('\\'),
                        (string)folder.Name,
                        (int)folder.Items.Count,
                        (int)folder.UnReadItemCount));
                }
                Walk(folder, depth + 1, result);
            }
            catch (COMException)
            {
                // A folder that cannot be read (offline, no permission) is skipped, not fatal.
            }
        }
    }

    public EmailPage ListEmails(
        string folder, string? account, DateTimeOffset? since, DateTimeOffset? until,
        bool unreadOnly, int limit, int offset) => Guard(() =>
    {
        dynamic items = ResolveFolder(Namespace(), folder, account).Items;
        if (unreadOnly)
        {
            items = items.Restrict("[UnRead] = True");
        }
        EmailPage page = Collect(items, since, until, limit, offset);
        return page;
    });

    public EmailPage SearchEmails(
        string folder, string? account, SearchTerms terms, DateTimeOffset? since,
        DateTimeOffset? until, int limit) => Guard(() =>
    {
        dynamic items = ResolveFolder(Namespace(), folder, account).Items;
        if (OutlookHelpers.BuildSearchFilter(terms) is { } filter)
        {
            items = items.Restrict(filter);
        }
        EmailPage page = Collect(items, since, until, limit, 0);
        return page;
    });

    public EmailDetails ReadEmail(string emailId) => Guard(() =>
    {
        dynamic item = Item(emailId);
        if ((int)item.Class != OlMail)
        {
            throw new NotFoundException("That item is not an email (for example, a meeting request).");
        }
        (List<Person> To, List<Person> Cc, List<Person> Bcc) people = Recipients(item);
        var (to, cc, bcc) = people;
        EmailSummary summary = Summary(item, When(item));
        List<AttachmentInfo> attachments = Attachments(item);
        return new EmailDetails(
            summary,
            to.Select(p => p.Display()).ToList(),
            cc.Select(p => p.Display()).ToList(),
            bcc.Select(p => p.Display()).ToList(),
            !(bool)item.Sent,
            (string?)item.Body ?? "",
            attachments);
    });

    public IReadOnlyList<AttachmentInfo> ListAttachments(string emailId) =>
        Guard<IReadOnlyList<AttachmentInfo>>(() =>
        {
            List<AttachmentInfo> attachments = Attachments(Item(emailId));
            return attachments;
        });

    public void SaveAttachment(string emailId, int index, string destinationPath) => Guard(() =>
    {
        dynamic attachments = Item(emailId).Attachments;
        if (index < 1 || index > (int)attachments.Count)
        {
            throw new NotFoundException("That attachment number does not exist on this email.");
        }
        attachments.Item(index).SaveAsFile(destinationPath);
    });

    // ---- drafting ----

    private static void Apply(dynamic mail, DraftChanges changes)
    {
        if (changes.To is not null)
        {
            mail.To = string.Join("; ", changes.To);
        }
        if (changes.Cc is not null)
        {
            mail.CC = string.Join("; ", changes.Cc);
        }
        if (changes.Bcc is not null)
        {
            mail.BCC = string.Join("; ", changes.Bcc);
        }
        if (changes.Subject is not null)
        {
            mail.Subject = changes.Subject;
        }
        if (changes.Body is not null)
        {
            mail.Body = changes.Body;
        }
        if (changes.To is not null || changes.Cc is not null || changes.Bcc is not null)
        {
            mail.Recipients.ResolveAll();
        }
    }

    private static void Prepend(dynamic mail, string text)
    {
        if (text.Length == 0)
        {
            return;
        }
        if ((int)mail.BodyFormat == OlFormatHtml)
        {
            mail.HTMLBody = OutlookHelpers.InsertHtmlReply((string?)mail.HTMLBody, text);
        }
        else
        {
            mail.Body = text + "\r\n\r\n" + ((string?)mail.Body ?? "");
        }
    }

    private static string IdOf(dynamic mail)
    {
        string storeId = mail.Parent.StoreID;
        string entryId = mail.EntryID;
        return OutlookHelpers.MakeId(storeId, entryId);
    }

    public string CreateDraft(
        IReadOnlyList<string> to, IReadOnlyList<string> cc, IReadOnlyList<string> bcc,
        string subject, string body, string? account, IReadOnlyList<string> attachments) => Guard(() =>
    {
        dynamic ns = Namespace();
        dynamic mail = Application().CreateItem(0);
        if (account is not null)
        {
            dynamic found = AccountNamed(ns, account)
                ?? throw new NotFoundException($"No account named '{account}' was found.");
            mail.SendUsingAccount = found;
        }
        Apply(mail, new DraftChanges(to, cc, bcc, subject, body));
        foreach (var path in attachments)
        {
            mail.Attachments.Add(path);
        }
        mail.Save();
        string id = IdOf(mail);
        return id;
    });

    public string CreateReplyDraft(string emailId, string body, bool replyAll) => Guard(() =>
    {
        dynamic original = Item(emailId);
        dynamic reply = replyAll ? original.ReplyAll() : original.Reply();
        Prepend(reply, body);
        reply.Save();
        string id = IdOf(reply);
        return id;
    });

    public string CreateForwardDraft(string emailId, IReadOnlyList<string> to, string body) => Guard(() =>
    {
        dynamic forward = Item(emailId).Forward();
        forward.To = string.Join("; ", to);
        forward.Recipients.ResolveAll();
        Prepend(forward, body);
        forward.Save();
        string id = IdOf(forward);
        return id;
    });

    public void UpdateDraft(string draftId, DraftChanges changes, IReadOnlyList<string> addAttachments) =>
        Guard(() =>
        {
            dynamic mail = Draft(draftId);
            Apply(mail, changes);
            foreach (var path in addAttachments)
            {
                mail.Attachments.Add(path);
            }
            mail.Save();
        });

    public DraftSnapshot GetDraft(string draftId) => Guard(() =>
    {
        dynamic mail = Draft(draftId);
        dynamic? account = null;
        try
        {
            account = mail.SendUsingAccount;
        }
        catch (Exception)
        {
            // Fall back to the default account below.
        }
        if (account is null)
        {
            dynamic accounts = Namespace().Accounts;
            if ((int)accounts.Count > 0)
            {
                account = accounts.Item(1);
            }
        }
        (List<Person> To, List<Person> Cc, List<Person> Bcc) people = Recipients(mail);
        var (to, cc, bcc) = people;
        List<AttachmentInfo> attachments = Attachments(mail);
        return new DraftSnapshot(
            account is null ? "" : (string?)account.DisplayName ?? "",
            account is null ? "" : (string?)account.SmtpAddress ?? "",
            account is null ? "other" : AccountType((int)account.AccountType),
            to, cc, bcc,
            (string?)mail.Subject ?? "",
            (string?)mail.Body ?? "",
            attachments);
    });

    public void Send(string draftId) => Guard(() => Draft(draftId).Send());

    public void Display(string draftId) => Guard(() => Draft(draftId).Display(false));
}
