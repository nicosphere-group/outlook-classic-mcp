using System.Text.Json;

namespace OutlookClassicMcp.Tests;

public class HelperAndDialogTests
{
    // ---- confirmation window wording ----

    private static DraftSnapshot Snapshot() => new(
        "自分", "me@example.com", "exchange",
        [new Person("田中 太郎", "tanaka@example.co.jp")], [], [],
        "Re: お見積りの件", "田中様\nお世話になっております。",
        [new AttachmentInfo(1, "見積書.pdf", 123_000)]);

    [Fact]
    public void Window_shows_everything_that_will_be_sent()
    {
        var view = DialogText.Build(Snapshot(), [new Warning("attachments", Count: 1)], "ja");
        var rows = view.Rows.ToDictionary(r => r.Label, r => r.Value);
        Assert.Equal("自分 <me@example.com>", rows["差出人"]);
        Assert.Equal("田中 太郎 <tanaka@example.co.jp>", rows["宛先"]);
        Assert.Equal("なし", rows["CC"]);
        Assert.Equal("なし", rows["BCC"]);
        Assert.Equal("Re: お見積りの件", rows["件名"]);
        Assert.Equal("見積書.pdf（120 KB）", rows["添付"]);
        Assert.Equal(["添付ファイルが 1 件あります。"], view.WarningLines);
        Assert.StartsWith("田中様", view.Body);
        Assert.Equal("送信する", view.SendText);
        Assert.Equal("Outlook で開いて確認", view.OpenText);
        Assert.Equal("送信しない", view.CancelText);
    }

    [Fact]
    public void Warning_lines_are_plain_sentences()
    {
        IReadOnlyList<Warning> warnings =
        [
            new("many_recipients", Count: 12),
            new("reply_all"),
            new("external", Domains: ["example.org", "example.net"]),
        ];
        var ja = DialogText.Build(Snapshot(), warnings, "ja").WarningLines;
        Assert.Equal("宛先が 12 件あります。", ja[0]);
        Assert.Contains("全員に返信", ja[1]);
        Assert.Equal("別の組織のアドレスが含まれています：example.org, example.net", ja[2]);
        var en = DialogText.Build(Snapshot(), warnings, "en").WarningLines;
        Assert.Equal("There are 12 recipients.", en[0]);
    }

    [Fact]
    public void Languages_have_the_same_keys_and_empty_subject_is_labelled()
    {
        Assert.Equal(DialogText.Keys("ja").Order(), DialogText.Keys("en").Order());
        var view = DialogText.Build(Snapshot() with { Subject = "" }, [], "en");
        Assert.Contains(view.Rows, r => r.Label == "Subject" && r.Value == "(no subject)");
    }

    [Fact]
    public void Format_size()
    {
        Assert.Equal("1 KB", DialogText.FormatSize(10));
        Assert.Equal("5.0 MB", DialogText.FormatSize(5 * 1024 * 1024));
    }

    // ---- Outlook helpers that need no Outlook ----

    [Fact]
    public void Ids_round_trip_and_bad_ids_are_rejected()
    {
        var id = OutlookHelpers.MakeId("STORE-ID-XYZ", "00000000ABCDEF");
        var (key, entry) = OutlookHelpers.SplitId(id);
        Assert.Equal("00000000ABCDEF", entry);
        Assert.Equal(10, key.Length);
        Assert.Equal(OutlookHelpers.StoreKey("STORE-ID-XYZ"), key);
        Assert.NotEqual(OutlookHelpers.StoreKey("OTHER-STORE"), key);
        foreach (var bad in new[] { "", "nodot", ".x", "x.", null })
        {
            Assert.Throws<InvalidInputException>(() => OutlookHelpers.SplitId(bad));
        }
    }

    [Fact]
    public void Outlook_time_is_treated_as_local_wall_clock()
    {
        var local = OutlookHelpers.ToLocal(new DateTime(2026, 1, 10, 9, 30, 0))!.Value;
        Assert.Equal((9, 30), (local.Hour, local.Minute));
        Assert.Equal(TimeZoneInfo.Local.GetUtcOffset(new DateTime(2026, 1, 10, 9, 30, 0)), local.Offset);
        Assert.Null(OutlookHelpers.ToLocal(new DateTime(4501, 1, 1))); // Outlook's "no date"
        Assert.Matches(@"^2026-01-10T09:30:00[+-]\d\d:\d\d$", OutlookHelpers.Iso(local));
    }

    [Fact]
    public void Search_filter_escapes_quotes_and_has_no_dates()
    {
        Assert.Null(OutlookHelpers.BuildSearchFilter(new SearchTerms(null, null, null, null)));
        var filter = OutlookHelpers.BuildSearchFilter(new SearchTerms(null, "O'Brien の件", "田中", null))!;
        Assert.StartsWith("@SQL=", filter);
        Assert.Contains("like '%O''Brien の件%'", filter);
        Assert.Contains("\"urn:schemas:httpmail:fromname\" like '%田中%'", filter);
        Assert.Contains(" AND ", filter);
        Assert.DoesNotContain("date", filter, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Reply_text_is_inserted_above_the_quote_and_escaped()
    {
        const string original = "<html><body lang=\"ja\"><p>元のメール</p></body></html>";
        var result = OutlookHelpers.InsertHtmlReply(original, "承知しました。\n<script>");
        Assert.True(result.IndexOf("承知しました。", StringComparison.Ordinal) < result.IndexOf("元のメール", StringComparison.Ordinal));
        Assert.Contains("&lt;script&gt;", result);
        Assert.DoesNotContain("<script>", result);
        Assert.StartsWith("<html><body lang=\"ja\"><div>", result);
        Assert.StartsWith("<div>x</div>", OutlookHelpers.InsertHtmlReply("no body tag", "x"));
    }

    // ---- what the tools hand to Claude ----

    [Fact]
    public async Task Tool_output_keeps_japanese_readable_and_uses_snake_case()
    {
        using var rig = new Rig();
        rig.Backend.AddInbox("store.mail1", subject: "【至急】お見積りの件①", body: "髙橋様　～よろしく～");
        var runner = new MailRunner(rig.Service, new InlineQueue());
        var json = await runner.Run(s => s.ReadEmail("store.mail1"));
        Assert.Contains("【至急】お見積りの件①", json); // not turned into \\uXXXX escapes
        var root = JsonDocument.Parse(json).RootElement;
        Assert.Equal("髙橋様　～よろしく～", root.GetProperty("body").GetString());
        Assert.False(root.GetProperty("body_truncated").GetBoolean());
        Assert.Equal("tanaka@example.co.jp", root.GetProperty("from").GetProperty("address").GetString());
        Assert.Equal(MailService.UntrustedNotice, root.GetProperty("notice").GetString());
    }

    private sealed class InlineQueue : IWorkQueue
    {
        public Task<T> RunAsync<T>(Func<T> work) => Task.FromResult(work());
    }
}
