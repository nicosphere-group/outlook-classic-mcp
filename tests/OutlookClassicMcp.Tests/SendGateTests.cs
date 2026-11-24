using System.Text.Json;

namespace OutlookClassicMcp.Tests;

/// <summary>The most important tests: nothing is sent unless a human pressed Send.</summary>
public class SendGateTests
{
    private const string SecretBody = "これは本文です。記録に残ってはいけません。";

    private static string Draft(Rig rig, string[]? to = null, string[]? cc = null) =>
        rig.Service.CreateDraft(to ?? ["tanaka@example.co.jp"], "お見積りの件", SecretBody, cc).DraftId;

    private static JsonElement LastAudit(Rig rig) =>
        JsonDocument.Parse(File.ReadAllLines(rig.AuditFile)[^1]).RootElement;

    [Fact]
    public void Creating_a_draft_never_sends_or_asks()
    {
        using var rig = new Rig();
        var result = rig.Service.CreateDraft(["a@example.com"], "件名", "本文");
        Assert.Contains("NOT sent", result.Status);
        Assert.Empty(rig.Backend.Sent);
        Assert.Empty(rig.Confirmer.Asked);
    }

    [Fact]
    public void Sends_when_the_user_presses_send()
    {
        using var rig = new Rig(Decision.Send);
        var result = rig.Service.SendDraft(Draft(rig));
        Assert.True(result.Sent);
        Assert.Equal("sent", result.Outcome);
        Assert.Single(rig.Backend.Sent);
        var (shown, _, timeout) = rig.Confirmer.Asked[0];
        Assert.Equal("お見積りの件", shown.Subject);
        Assert.Equal(SecretBody, shown.Body);
        Assert.Equal(rig.Config.ConfirmTimeout, timeout);
    }

    [Theory]
    [InlineData(Decision.Cancel, "declined")]
    [InlineData(Decision.Timeout, "timeout")]
    [InlineData(Decision.Failed, "failed")]
    [InlineData(Decision.Open, "opened")]
    public void Anything_other_than_send_does_not_send(Decision decision, string outcome)
    {
        using var rig = new Rig(decision);
        var id = Draft(rig);
        var result = rig.Service.SendDraft(id);
        Assert.False(result.Sent);
        Assert.Equal(outcome, result.Outcome);
        Assert.Empty(rig.Backend.Sent);
        Assert.True(rig.Backend.Drafts.ContainsKey(id)); // the draft is kept
        Assert.Equal(outcome, LastAudit(rig).GetProperty("event").GetString());
    }

    [Fact]
    public void Open_shows_the_draft_in_outlook_without_sending()
    {
        using var rig = new Rig(Decision.Open);
        var id = Draft(rig);
        rig.Service.SendDraft(id);
        Assert.Equal([id], rig.Backend.Displayed);
        Assert.Empty(rig.Backend.Sent);
    }

    [Fact]
    public void Draft_changed_during_confirmation_is_not_sent()
    {
        using var rig = new Rig(Decision.Send);
        var id = Draft(rig);
        rig.Confirmer.During = () =>
            rig.Backend.Drafts[id] = rig.Backend.Drafts[id] with { Body = "書き換えられた本文" };
        var result = rig.Service.SendDraft(id);
        Assert.False(result.Sent);
        Assert.Equal("changed", result.Outcome);
        Assert.Empty(rig.Backend.Sent);
    }

    [Fact]
    public void Recipient_added_during_confirmation_is_not_sent()
    {
        using var rig = new Rig(Decision.Send);
        var id = Draft(rig);
        rig.Confirmer.During = () => rig.Service.UpdateDraft(id, bcc: ["evil@example.net"]);
        Assert.Equal("changed", rig.Service.SendDraft(id).Outcome);
        Assert.Empty(rig.Backend.Sent);
    }

    [Fact]
    public void Every_send_asks_again()
    {
        using var rig = new Rig(Decision.Send);
        rig.Service.SendDraft(Draft(rig));
        rig.Service.SendDraft(Draft(rig));
        Assert.Equal(2, rig.Confirmer.Asked.Count);
        Assert.Equal(2, rig.Backend.Sent.Count);
    }

    [Fact]
    public void An_already_sent_email_cannot_be_sent_or_edited()
    {
        using var rig = new Rig(Decision.Send);
        var id = Draft(rig);
        rig.Service.SendDraft(id);
        Assert.Throws<NotADraftException>(() => rig.Service.SendDraft(id));
        Assert.Throws<NotADraftException>(() => rig.Service.UpdateDraft(id, subject: "変更"));
        Assert.Single(rig.Backend.Sent);
        Assert.Single(rig.Confirmer.Asked);
    }

    [Fact]
    public void A_received_email_cannot_be_sent()
    {
        using var rig = new Rig(Decision.Send);
        rig.Backend.AddInbox("store.mail1");
        Assert.Throws<NotADraftException>(() => rig.Service.SendDraft("store.mail1"));
        Assert.Empty(rig.Confirmer.Asked);
    }

    [Fact]
    public void Draft_without_recipients_is_rejected_before_asking()
    {
        using var rig = new Rig(Decision.Send);
        var id = Draft(rig, to: []);
        Assert.Throws<InvalidInputException>(() => rig.Service.SendDraft(id));
        Assert.Empty(rig.Confirmer.Asked);
        Assert.Empty(rig.Backend.Sent);
    }

    [Fact]
    public void Send_can_be_turned_off()
    {
        using var rig = new Rig(Decision.Send, allowSend: false);
        var id = Draft(rig);
        Assert.Throws<NotAllowedException>(() => rig.Service.SendDraft(id));
        Assert.Empty(rig.Confirmer.Asked);
        Assert.Empty(rig.Backend.Sent);
    }

    [Fact]
    public void Audit_log_records_recipients_but_never_the_body()
    {
        using var rig = new Rig(Decision.Send);
        rig.Service.SendDraft(Draft(rig, cc: ["cc@example.com"]));
        var raw = File.ReadAllText(rig.AuditFile);
        var entry = LastAudit(rig);
        Assert.Equal("sent", entry.GetProperty("event").GetString());
        Assert.Equal("tanaka@example.co.jp", entry.GetProperty("to")[0].GetString());
        Assert.Equal("cc@example.com", entry.GetProperty("cc")[0].GetString());
        Assert.Equal("お見積りの件", entry.GetProperty("subject").GetString());
        Assert.Contains("お見積りの件", raw); // Japanese is stored readably
        Assert.DoesNotContain(SecretBody, raw);
        Assert.False(entry.TryGetProperty("body", out _));
    }

    [Theory]
    [InlineData("pop3")]
    [InlineData("imap")]
    public void Pop3_and_imap_report_queued(string accountType)
    {
        using var rig = new Rig(Decision.Send, accountType);
        var result = rig.Service.SendDraft(Draft(rig));
        Assert.True(result.Sent);
        Assert.Equal("queued", result.Outcome);
    }

    [Fact]
    public void Reply_all_and_external_recipients_are_highlighted()
    {
        using var rig = new Rig(Decision.Cancel);
        rig.Backend.AddInbox("store.mail1", cc: ["other@example.org"]);
        var id = rig.Service.CreateReplyDraft("store.mail1", "承知しました。", replyAll: true).DraftId;
        rig.Service.SendDraft(id);
        var codes = rig.Confirmer.Asked[0].Warnings.Select(w => w.Code).ToHashSet();
        Assert.Contains("reply_all", codes);
        Assert.Contains("external", codes);
    }
}
