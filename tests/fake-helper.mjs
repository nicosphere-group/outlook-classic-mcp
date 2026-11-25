// Pretends to be the PowerShell helper, so the Node side of the line protocol can be
// tested on any OS. Speaks the same one-JSON-object-per-line protocol.

import readline from "node:readline";

const drafts = new Map();
let next = 0;
const reply = (message) => process.stdout.write(JSON.stringify(message) + "\n");

const ops = {
  ping: (a) => ({ pong: true, echo: a.echo }),
  list_accounts: () => ({ accounts: { name: "自分", address: "me@example.com", type: "pop3", is_default: true }, data_files: "個人用フォルダー" }),
  list_folders: () => ({ folders: [{ path: "me\\受信トレイ", name: "受信トレイ", total: 3, unread: 1 }] }),
  list_emails: (a) => ({
    emails: { id: "k.1", subject: "【至急】お見積りの件①", from: { name: "髙橋", address: "t@example.co.jp" }, received: a.since, unread: true, attachment_count: 0 },
    has_more: false,
    timed_out: false,
  }),
  search_emails: (a) => ({ emails: [], has_more: false, timed_out: false, echo_filter: a.filter }),
  create_draft: (a) => {
    const id = `k.d${++next}`;
    drafts.set(id, a);
    return { id };
  },
  get_draft: (a) => {
    const d = drafts.get(a.id);
    if (!d) throw { kind: "not_draft", message: "That item is not an unsent draft, so it cannot be changed or sent." };
    return {
      sender_name: "自分", sender_address: "me@example.com", account_type: "pop3",
      to: d.to.map((x) => ({ name: x, address: x })), cc: [], bcc: [],
      subject: d.subject, body: d.body, attachments: [],
    };
  },
  confirm: (a) => ({ decision: process.env.FAKE_DECISION ?? "cancel", saw_title: a.view.title }),
  read_email: () => {
    throw { kind: "not_found", message: "That email was not found. It may have been moved or deleted; list the emails again." };
  },
  send: () => {
    throw { kind: "unavailable", message: "The connection to Outlook was lost. Make sure Outlook is open and try again." };
  },
  display: () => "not an object", // a malformed answer
  save_attachment: () => new Promise(() => {}), // never answers
  list_attachments: () => process.exit(3), // crashes
};

readline.createInterface({ input: process.stdin }).on("line", async (line) => {
  if (!line.trim()) return;
  const { id, op, args } = JSON.parse(line);
  // The real helper must never see raw non-ASCII bytes on its input.
  if (/[^\x00-\x7f]/.test(line)) return reply({ id, ok: false, kind: "error", message: "non-ASCII request" });
  process.stdout.write("stray output that is not JSON\n");
  try {
    reply({ id, ok: true, result: await ops[op](args ?? {}) });
  } catch (e) {
    reply({ id, ok: false, kind: e.kind ?? "error", message: e.message ?? String(e) });
  }
});
