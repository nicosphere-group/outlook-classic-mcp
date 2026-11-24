namespace OutlookClassicMcp.Tests;

public class ServiceAndSafetyTests
{
    // ---- file names and paths ----

    [Theory]
    [InlineData("見積書.pdf", "見積書.pdf")]
    [InlineData("..\\..\\Windows\\evil.exe", "evil.exe")]
    [InlineData("../../etc/passwd", "passwd")]
    [InlineData("C:\\Users\\x\\a.txt", "a.txt")]
    [InlineData("a<b>:\"c|d?.txt", "a_b___c_d_.txt")]
    [InlineData("CON.txt", "_CON.txt")]
    [InlineData("...", "attachment")]
    [InlineData("", "attachment")]
    public void Sanitize_filename(string raw, string expected) =>
        Assert.Equal(expected, Safety.SanitizeFilename(raw));

    [Fact]
    public void Long_names_keep_their_extension()
    {
        var name = Safety.SanitizeFilename(new string('あ', 300) + ".xlsx");
        Assert.True(name.Length <= 150);
        Assert.EndsWith(".xlsx", name);
    }

    [Fact]
    public void Saved_attachments_stay_in_the_folder_and_never_overwrite()
    {
        using var rig = new Rig();
        var dir = rig.Config.SaveDir;
        var first = Safety.SafeSavePath(dir, "..\\..\\report.pdf");
        File.WriteAllText(first, "x");
        var second = Safety.SafeSavePath(dir, "report.pdf");
        Assert.Equal(Path.GetFullPath(Path.Combine(dir, "report.pdf")), first);
        Assert.Equal(Path.GetFullPath(Path.Combine(dir, "report (1).pdf")), second);
        Assert.True(Safety.IsWithin(first, [dir]));
        Assert.True(Safety.IsWithin(second, [dir]));
    }

    [Fact]
    public void Save_attachment_uses_a_safe_path()
    {
        using var rig = new Rig();
        rig.Backend.AddInbox(
            "store.mail1", attachments: [new AttachmentInfo(1, "..\\..\\startup\\run.bat", 10)]);
        var saved = rig.Service.SaveAttachment("store.mail1", 1).SavedTo;
        Assert.Equal(Path.GetFullPath(rig.Config.SaveDir), Path.GetDirectoryName(saved));
        Assert.Equal("run.bat", Path.GetFileName(saved));
        Assert.True(File.Exists(saved));
        Assert.Throws<InvalidInputException>(() => rig.Service.SaveAttachment("store.mail1", 2));
    }

    [Fact]
    public void Attachments_only_from_allowed_folders()
    {
        using var rig = new Rig();
        var allowed = Path.Combine(rig.AttachDir, "資料.txt");
        File.WriteAllText(allowed, "ok");
        var outside = Path.Combine(rig.Root, "secret.txt");
        File.WriteAllText(outside, "secret");
        string[] to = ["a@example.com"];

        var result = rig.Service.CreateDraft(to, "件名", "本文", attachments: [allowed]);
        Assert.Equal("資料.txt", result.Draft.Attachments[0].Filename);

        Assert.Throws<NotAllowedException>(
            () => rig.Service.CreateDraft(to, "件名", "本文", attachments: [outside]));
        var sneaky = Path.Combine(rig.AttachDir, "..", "secret.txt");
        Assert.Throws<NotAllowedException>(
            () => rig.Service.CreateDraft(to, "件名", "本文", attachments: [sneaky]));
        Assert.Throws<NotAllowedException>(
            () => rig.Service.CreateDraft(to, "件名", "本文", attachments: ["relative.txt"]));
        Assert.Throws<NotFoundException>(
            () => rig.Service.CreateDraft(to, "件名", "本文", attachments: [Path.Combine(rig.AttachDir, "none")]));
        Assert.Single(rig.Backend.Drafts);
    }

    [Fact]
    public void A_link_inside_an_allowed_folder_cannot_reach_outside()
    {
        using var rig = new Rig();
        var outside = Path.Combine(rig.Root, "secret.txt");
        File.WriteAllText(outside, "secret");
        var link = Path.Combine(rig.AttachDir, "innocent.txt");
        try
        {
            File.CreateSymbolicLink(link, outside);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            return; // creating links needs a privilege on some Windows setups
        }
        Assert.Throws<NotAllowedException>(
            () => rig.Service.CreateDraft(["a@example.com"], "件名", "本文", attachments: [link]));
    }

    // ---- input handling ----

    [Theory]
    [InlineData("田中")]
    [InlineData("a@example.com; b@example.com")]
    [InlineData("a@example.com,b@example.com")]
    public void Addresses_must_be_complete(string bad)
    {
        using var rig = new Rig();
        Assert.Throws<InvalidInputException>(() => rig.Service.CreateDraft([bad], "件名", "本文"));
        Assert.Empty(rig.Backend.Drafts);
    }

    [Fact]
    public void Parse_when()
    {
        Assert.Null(MailService.ParseWhen(null));
        Assert.Null(MailService.ParseWhen("  "));
        var start = MailService.ParseWhen("2026-01-31")!.Value;
        var end = MailService.ParseWhen("2026-01-31", endOfDay: true)!.Value;
        Assert.Equal((0, 0), (start.Hour, start.Minute));
        Assert.Equal(TimeZoneInfo.Local.GetUtcOffset(start.DateTime), start.Offset);
        Assert.Equal((31, 23, 59, 59), (end.Day, end.Hour, end.Minute, end.Second));
        Assert.Equal(TimeSpan.Zero, MailService.ParseWhen("2026-01-31T09:30:00Z")!.Value.Offset);
        Assert.Equal(9, MailService.ParseWhen("2026-01-31T09:30")!.Value.Hour);
        Assert.Throws<InvalidInputException>(() => MailService.ParseWhen("昨日"));
        Assert.Throws<InvalidInputException>(() => MailService.ParseWhen("1/2/2026")); // ambiguous by region
    }

    [Fact]
    public void List_limits_are_clamped()
    {
        using var rig = new Rig();
        rig.Service.ListEmails(limit: 100000, offset: -5, until: "2026-01-31");
        Assert.Equal(100, rig.Backend.LastList.Limit);
        Assert.Equal(0, rig.Backend.LastList.Offset);
        Assert.Equal(23, rig.Backend.LastList.Until!.Value.Hour);
    }

    [Fact]
    public void Search_needs_something_to_search_for()
    {
        using var rig = new Rig();
        Assert.Throws<InvalidInputException>(() => rig.Service.SearchEmails());
        Assert.Throws<InvalidInputException>(() => rig.Service.SearchEmails(query: "   "));
        Assert.NotNull(rig.Service.SearchEmails(sender: "田中").Emails);
    }

    // ---- reading ----

    [Fact]
    public void Read_email_truncates_and_marks_content_untrusted()
    {
        using var rig = new Rig();
        rig.Backend.AddInbox("store.mail1", subject: "【重要】ご確認ください", body: new string('あ', 2000));
        var result = rig.Service.ReadEmail("store.mail1", maxChars: 500);
        Assert.Equal(MailService.UntrustedNotice, result.Notice);
        Assert.Equal("【重要】ご確認ください", result.Subject);
        Assert.Equal(500, result.Body.Length);
        Assert.True(result.BodyTruncated);
        Assert.Equal(2000, result.BodyTotalChars);
        Assert.Equal(MailService.UntrustedNotice, rig.Service.ListEmails().Notice);
    }

    // ---- warnings ----

    private static DraftSnapshot Snapshot(
        IReadOnlyList<Person>? to = null, IReadOnlyList<Person>? bcc = null, string subject = "",
        string body = "", IReadOnlyList<AttachmentInfo>? attachments = null) =>
        new("自分", "me@example.com", "exchange", to ?? [], [], bcc ?? [], subject, body, attachments ?? []);

    [Fact]
    public void No_warnings_for_a_plain_internal_email() =>
        Assert.Empty(Safety.ComputeWarnings(Snapshot(to: [new Person("同僚", "colleague@example.com")])));

    [Fact]
    public void Warnings()
    {
        var many = Enumerable.Range(0, 5).Select(i => new Person("", $"user{i}@example.com")).ToList();
        var snapshot = Snapshot(
            to: many, bcc: [new Person("", "x@Other.example")], attachments: [new AttachmentInfo(1, "a.pdf", 100)]);
        var warnings = Safety.ComputeWarnings(snapshot, replyAll: true).ToDictionary(w => w.Code);
        Assert.Equal(6, warnings["many_recipients"].Count);
        Assert.Equal(["other.example"], warnings["external"].Domains);
        Assert.Equal(1, warnings["attachments"].Count);
        Assert.True(warnings.ContainsKey("reply_all"));
    }

    [Fact]
    public void Fingerprint_changes_with_any_visible_field()
    {
        IReadOnlyList<Person> a = [new Person("A", "a@example.com")];
        var same = Snapshot(to: a, subject: "s", body: "b");
        DraftSnapshot[] variants =
        [
            Snapshot(to: [new Person("A", "b@example.com")], subject: "s", body: "b"),
            Snapshot(to: a, subject: "s2", body: "b"),
            Snapshot(to: a, subject: "s", body: "b2"),
            Snapshot(to: a, bcc: [new Person("", "c@x.com")], subject: "s", body: "b"),
            Snapshot(to: a, subject: "s", body: "b", attachments: [new AttachmentInfo(1, "f", 1)]),
        ];
        Assert.Equal(same.Fingerprint(), Snapshot(to: [new Person("A", "a@example.com")], subject: "s", body: "b").Fingerprint());
        var all = variants.Select(v => v.Fingerprint()).Append(same.Fingerprint()).ToHashSet();
        Assert.Equal(6, all.Count);
    }

    // ---- config ----

    [Fact]
    public void Config_defaults_are_safe_and_env_is_respected()
    {
        var none = Config.From(_ => null);
        Assert.True(none.AllowSend);
        Assert.False(none.Diagnostics);
        Assert.Equal(180, none.ConfirmTimeout);

        var env = new Dictionary<string, string>
        {
            ["OUTLOOK_MCP_ALLOW_SEND"] = "0",
            ["OUTLOOK_MCP_CONFIRM_TIMEOUT"] = "5",
            ["OUTLOOK_MCP_SAVE_DIR"] = "/somewhere",
            ["OUTLOOK_MCP_LANG"] = "EN",
            ["OUTLOOK_MCP_DIAGNOSTICS"] = "true",
        };
        var custom = Config.From(env.GetValueOrDefault);
        Assert.False(custom.AllowSend);
        Assert.Equal(15, custom.ConfirmTimeout); // clamped to the minimum
        Assert.Equal("/somewhere", custom.SaveDir);
        Assert.Equal("en", custom.Lang);
        Assert.True(custom.Diagnostics);
    }

    [Theory]
    [InlineData("false", false)]
    [InlineData("FALSE", false)]
    [InlineData("0", false)]
    [InlineData("true", true)]
    [InlineData("1", true)]
    [InlineData("", true)]
    [InlineData("${user_config.allow_send}", true)] // a placeholder the host did not fill in
    [InlineData("nonsense", true)]
    public void Send_is_off_only_when_switched_off_explicitly(string value, bool expected) =>
        Assert.Equal(expected, Config.From(n => n == "OUTLOOK_MCP_ALLOW_SEND" ? value : null).AllowSend);

    [Fact]
    public void Unfilled_placeholders_and_bad_numbers_fall_back_to_defaults()
    {
        var config = Config.From(n => n switch
        {
            "OUTLOOK_MCP_SAVE_DIR" => "${user_config.save_dir}",
            "OUTLOOK_MCP_CONFIRM_TIMEOUT" => "abc",
            _ => null,
        });
        Assert.Equal(new Config().SaveDir, config.SaveDir);
        Assert.Equal(180, config.ConfirmTimeout);
    }
}
