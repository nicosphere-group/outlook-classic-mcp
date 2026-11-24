using System.Globalization;

namespace OutlookClassicMcp;

/// <summary>Everything the confirmation window displays, already localized.</summary>
public sealed record DialogView(
    string Title,
    string Heading,
    IReadOnlyList<string> WarningLines,
    IReadOnlyList<(string Label, string Value)> Rows,
    string BodyLabel,
    string Body,
    string SendText,
    string OpenText,
    string CancelText,
    string CountdownFormat);

/// <summary>Wording of the confirmation window. Plain language, Japanese and English.</summary>
public static class DialogText
{
    private static readonly Dictionary<string, Dictionary<string, string>> Text = new()
    {
        ["ja"] = new()
        {
            ["title"] = "メールを送信しますか？",
            ["heading"] = "このメールを送信しますか？",
            ["from"] = "差出人",
            ["to"] = "宛先",
            ["cc"] = "CC",
            ["bcc"] = "BCC",
            ["subject"] = "件名",
            ["attachments"] = "添付",
            ["none"] = "なし",
            ["no_subject"] = "（件名なし）",
            ["body"] = "本文",
            ["send"] = "送信する",
            ["open"] = "Outlook で開いて確認",
            ["cancel"] = "送信しない",
            ["countdown"] = "あと {0} 秒で自動的に閉じます（その場合は送信されません）",
            ["w_many_recipients"] = "宛先が {0} 件あります。",
            ["w_reply_all"] = "「全員に返信」です。全員に送ってよいか確認してください。",
            ["w_external"] = "別の組織のアドレスが含まれています：{1}",
            ["w_attachments"] = "添付ファイルが {0} 件あります。",
        },
        ["en"] = new()
        {
            ["title"] = "Send this email?",
            ["heading"] = "Do you want to send this email?",
            ["from"] = "From",
            ["to"] = "To",
            ["cc"] = "CC",
            ["bcc"] = "BCC",
            ["subject"] = "Subject",
            ["attachments"] = "Attachments",
            ["none"] = "None",
            ["no_subject"] = "(no subject)",
            ["body"] = "Message",
            ["send"] = "Send",
            ["open"] = "Open in Outlook to check",
            ["cancel"] = "Don't send",
            ["countdown"] = "Closes automatically in {0} seconds (the email will not be sent)",
            ["w_many_recipients"] = "There are {0} recipients.",
            ["w_reply_all"] = "This is a Reply All. Check that everyone should receive it.",
            ["w_external"] = "Includes addresses outside your organization: {1}",
            ["w_attachments"] = "{0} file(s) attached.",
        },
    };

    public static IReadOnlyCollection<string> Languages => Text.Keys;

    public static IReadOnlyCollection<string> Keys(string lang) => Text[lang].Keys;

    public static string PickLanguage(string requested) =>
        Text.ContainsKey(requested)
            ? requested
            : CultureInfo.CurrentUICulture.TwoLetterISOLanguageName == "ja" ? "ja" : "en";

    public static string FormatSize(long bytes) =>
        bytes >= 1024 * 1024
            ? string.Create(CultureInfo.InvariantCulture, $"{bytes / (1024.0 * 1024.0):0.0} MB")
            : $"{Math.Max(1, (long)Math.Round(bytes / 1024.0))} KB";

    public static DialogView Build(DraftSnapshot snapshot, IReadOnlyList<Warning> warnings, string lang)
    {
        var t = Text[PickLanguage(lang)];

        string People(IReadOnlyList<Person> people) =>
            people.Count == 0 ? t["none"] : string.Join("\n", people.Select(p => p.Display()));

        var from = new Person(snapshot.SenderName, snapshot.SenderAddress).Display();
        var attachments = snapshot.Attachments.Count == 0
            ? t["none"]
            : string.Join("\n", snapshot.Attachments.Select(a => $"{a.Filename}（{FormatSize(a.Size)}）"));

        var warningLines = warnings
            .Where(w => t.ContainsKey("w_" + w.Code))
            .Select(w => string.Format(
                CultureInfo.InvariantCulture, t["w_" + w.Code], w.Count, string.Join(", ", w.Domains ?? [])))
            .ToList();

        return new DialogView(
            t["title"],
            t["heading"],
            warningLines,
            [
                (t["from"], from),
                (t["to"], People(snapshot.To)),
                (t["cc"], People(snapshot.Cc)),
                (t["bcc"], People(snapshot.Bcc)),
                (t["subject"], snapshot.Subject.Length > 0 ? snapshot.Subject : t["no_subject"]),
                (t["attachments"], attachments),
            ],
            t["body"],
            snapshot.Body,
            t["send"],
            t["open"],
            t["cancel"],
            t["countdown"]);
    }
}
