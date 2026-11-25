// Wording of the confirmation window. Plain language, Japanese and English.

export const TEXT = {
  ja: {
    title: "メールを送信しますか？",
    heading: "このメールを送信しますか？",
    from: "差出人",
    to: "宛先",
    cc: "CC",
    bcc: "BCC",
    subject: "件名",
    attachments: "添付",
    none: "なし",
    no_subject: "（件名なし）",
    body: "本文",
    send: "送信する",
    open: "Outlook で開いて確認",
    cancel: "送信しない",
    countdown: "あと {0} 秒で自動的に閉じます（その場合は送信されません）",
    w_many_recipients: "宛先が {count} 件あります。",
    w_reply_all: "「全員に返信」です。全員に送ってよいか確認してください。",
    w_external: "別の組織のアドレスが含まれています：{domains}",
    w_attachments: "添付ファイルが {count} 件あります。",
  },
  en: {
    title: "Send this email?",
    heading: "Do you want to send this email?",
    from: "From",
    to: "To",
    cc: "CC",
    bcc: "BCC",
    subject: "Subject",
    attachments: "Attachments",
    none: "None",
    no_subject: "(no subject)",
    body: "Message",
    send: "Send",
    open: "Open in Outlook to check",
    cancel: "Don't send",
    countdown: "Closes automatically in {0} seconds (the email will not be sent)",
    w_many_recipients: "There are {count} recipients.",
    w_reply_all: "This is a Reply All. Check that everyone should receive it.",
    w_external: "Includes addresses outside your organization: {domains}",
    w_attachments: "{count} file(s) attached.",
  },
};

export function pickLanguage(requested, systemLocale = Intl.DateTimeFormat().resolvedOptions().locale) {
  if (requested in TEXT) return requested;
  return String(systemLocale).toLowerCase().startsWith("ja") ? "ja" : "en";
}

export function formatSize(bytes) {
  return bytes >= 1024 * 1024
    ? `${(bytes / (1024 * 1024)).toFixed(1)} MB`
    : `${Math.max(1, Math.round(bytes / 1024))} KB`;
}

const display = (p) => (p.name && p.address && p.name !== p.address ? `${p.name} <${p.address}>` : p.address || p.name);

/** Everything the confirmation window displays, already localized. */
export function buildView(snapshot, warnings, lang) {
  const t = TEXT[pickLanguage(lang)];
  const people = (list) => (list.length === 0 ? t.none : list.map(display).join("\n"));
  const attachments =
    snapshot.attachments.length === 0
      ? t.none
      : snapshot.attachments.map((a) => `${a.filename}（${formatSize(a.size)}）`).join("\n");
  const warningLines = warnings
    .filter((w) => t["w_" + w.code])
    .map((w) =>
      t["w_" + w.code].replace("{count}", String(w.count ?? "")).replace("{domains}", (w.domains ?? []).join(", ")),
    );
  return {
    title: t.title,
    heading: t.heading,
    warning_lines: warningLines,
    rows: [
      { label: t.from, value: display({ name: snapshot.sender_name, address: snapshot.sender_address }) },
      { label: t.to, value: people(snapshot.to) },
      { label: t.cc, value: people(snapshot.cc) },
      { label: t.bcc, value: people(snapshot.bcc) },
      { label: t.subject, value: snapshot.subject || t.no_subject },
      { label: t.attachments, value: attachments },
    ],
    body_label: t.body,
    body: snapshot.body,
    send_text: t.send,
    open_text: t.open,
    cancel_text: t.cancel,
    countdown_format: t.countdown,
  };
}
