// Talks to Outlook through one long-lived Windows PowerShell process.
//
// Why PowerShell: it ships with Windows and is signed by Microsoft, so nothing unsigned
// has to be started on the user's PC. The helper script does the Outlook (COM) calls and
// shows the confirmation window; all rules stay on this side.
//
// Protocol: one JSON object per line in each direction, UTF-8.

import { spawn } from "node:child_process";
import fs from "node:fs";
import path from "node:path";

import { buildView } from "./dialogText.js";
import {
  InvalidInputError,
  NotADraftError,
  NotFoundError,
  OutlookMcpError,
  OutlookUnavailableError,
} from "./errors.js";

const DEFAULT_TIMEOUT_MS = 90_000;
const CONFIRM_GRACE_MS = 20_000;

const ERROR_KINDS = {
  unavailable: OutlookUnavailableError,
  not_found: NotFoundError,
  not_draft: NotADraftError,
  invalid: InvalidInputError,
};

const SEARCH_PROPERTIES = [
  "urn:schemas:httpmail:subject",
  "urn:schemas:httpmail:fromname",
  "urn:schemas:httpmail:fromemail",
  "urn:schemas:httpmail:textdescription",
];

export const daslLike = (property, text) => `"${property}" like '%${text.replaceAll("'", "''")}%'`;

/**
 * Text conditions only. Dates are deliberately not part of the filter: Outlook's date
 * filter syntax depends on the PC's regional settings, so dates are compared by the helper.
 */
export function buildSearchFilter({ query, subject, sender, body }) {
  const clauses = [];
  if (query) clauses.push("(" + SEARCH_PROPERTIES.map((p) => daslLike(p, query)).join(" OR ") + ")");
  if (subject) clauses.push(daslLike(SEARCH_PROPERTIES[0], subject));
  if (sender) clauses.push(`(${daslLike(SEARCH_PROPERTIES[1], sender)} OR ${daslLike(SEARCH_PROPERTIES[2], sender)})`);
  if (body) clauses.push(daslLike(SEARCH_PROPERTIES[3], body));
  return clauses.length === 0 ? null : "@SQL=" + clauses.join(" AND ");
}

const escapeHtml = (s) =>
  s.replaceAll("&", "&amp;").replaceAll("<", "&lt;").replaceAll(">", "&gt;").replaceAll('"', "&quot;");

/** The new text as HTML, to be placed above the quoted original. */
export function htmlReplyBlock(text) {
  if (!text) return "";
  return "<div>" + text.split(/\r?\n/).map(escapeHtml).join("<br>") + "</div><br>";
}

/** JSON with every non-ASCII character escaped, so no code page can damage it in transit. */
export function asciiJson(value) {
  return JSON.stringify(value).replace(/[\u007f-\uffff]/g, (c) => "\\u" + c.charCodeAt(0).toString(16).padStart(4, "0"));
}

// Windows PowerShell 5.1 can turn a one-element list into a bare value; undo that here.
const list = (value) => (Array.isArray(value) ? value : value === null || value === undefined ? [] : [value]);

const person = (p) => ({ name: String(p?.name ?? ""), address: String(p?.address ?? "") });
const attachment = (a) => ({ index: Number(a.index), filename: String(a.filename ?? ""), size: Number(a.size ?? 0) });
const summary = (e) => ({
  id: String(e.id),
  subject: String(e.subject ?? ""),
  from: person(e.from),
  received: e.received || null,
  unread: Boolean(e.unread),
  attachment_count: Number(e.attachment_count ?? 0),
});
const page = (r) => ({
  emails: list(r.emails).map(summary),
  has_more: Boolean(r.has_more),
  timed_out: Boolean(r.timed_out),
});

export function helperPath() {
  return path.join(path.dirname(fs.realpathSync(process.argv[1])), "outlook-helper.ps1");
}

export function powershellPath(env = process.env) {
  // Full path, so a different powershell.exe earlier in PATH can never be picked up.
  return path.join(env.SystemRoot || env.windir || "C:\\Windows", "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
}

export class PowerShellBackend {
  /**
   * @param {{ command?: string, args?: string[], spawnImpl?: typeof spawn }} [options]
   */
  constructor(options = {}) {
    this.command = options.command ?? powershellPath();
    this.args = options.args ?? [
      "-NoLogo",
      "-NoProfile",
      "-NonInteractive",
      "-ExecutionPolicy",
      "Bypass",
      "-File",
      helperPath(),
    ];
    this.spawnImpl = options.spawnImpl ?? spawn;
    this.child = null;
    this.buffer = "";
    this.stderrTail = "";
    this.pending = null;
    this.nextId = 1;
    this.queue = Promise.resolve();
  }

  // ---- process handling ----

  #start() {
    const child = this.spawnImpl(this.command, this.args, { windowsHide: true, stdio: ["pipe", "pipe", "pipe"] });
    this.child = child;
    this.buffer = "";
    this.stderrTail = "";
    child.stdout.setEncoding("utf8");
    child.stderr.setEncoding("utf8");
    child.stdout.on("data", (chunk) => this.#onData(chunk));
    child.stderr.on("data", (chunk) => {
      this.stderrTail = (this.stderrTail + chunk).slice(-1500);
    });
    child.stdin.on("error", () => {});
    const gone = (why) => {
      if (this.child === child) this.child = null;
      // Only the request that was sent to this process can be failed by its exit;
      // a later request already belongs to a replacement process.
      if (this.pending?.child !== child) return;
      const detail = this.stderrTail.trim();
      this.#settle(
        new OutlookUnavailableError(
          "The Outlook helper stopped unexpectedly, so the operation was not completed. Try again." +
            (detail ? ` (${why}: ${detail.slice(-400)})` : ` (${why})`),
        ),
      );
    };
    child.on("error", (err) => gone(`could not start Windows PowerShell: ${err.message}`));
    child.on("exit", (code) => gone(`exit code ${code}`));
  }

  #onData(chunk) {
    this.buffer += chunk;
    let newline;
    while ((newline = this.buffer.indexOf("\n")) >= 0) {
      const line = this.buffer.slice(0, newline).trim();
      this.buffer = this.buffer.slice(newline + 1);
      if (!line) continue;
      let message;
      try {
        message = JSON.parse(line.replace(/^\ufeff/, ""));
      } catch {
        continue; // not ours (a stray line of output); ignore
      }
      if (this.pending && message.id === this.pending.id) this.#settle(null, message);
    }
  }

  #settle(error, message) {
    const pending = this.pending;
    if (!pending) return;
    this.pending = null;
    clearTimeout(pending.timer);
    if (error) pending.reject(error);
    else pending.resolve(message);
  }

  #stop() {
    const child = this.child;
    this.child = null;
    if (child) {
      try {
        child.kill();
      } catch {
        // already gone
      }
    }
  }

  close() {
    this.#stop();
  }

  /** Runs one operation. Operations never overlap: Outlook is used strictly in order. */
  call(op, args = {}, timeoutMs = DEFAULT_TIMEOUT_MS) {
    const run = () => this.#exchange(op, args, timeoutMs);
    const result = this.queue.then(run, run);
    this.queue = result.catch(() => {});
    return result;
  }

  async #exchange(op, args, timeoutMs) {
    if (!this.child) this.#start();
    const id = this.nextId++;
    const message = await new Promise((resolve, reject) => {
      const timer = setTimeout(() => {
        this.#stop(); // a stuck helper is replaced on the next call
        this.#settle(
          new OutlookUnavailableError(
            "Outlook did not answer in time. Make sure Outlook is open and not showing a dialog, then try again.",
          ),
        );
      }, timeoutMs);
      this.pending = { id, resolve, reject, timer, child: this.child };
      try {
        this.child.stdin.write(asciiJson({ id, op, args }) + "\n");
      } catch (err) {
        this.#settle(new OutlookUnavailableError(`The Outlook helper could not be reached: ${err.message}`));
      }
    });
    if (message.ok) {
      const result = message.result;
      if (result === null || typeof result !== "object" || Array.isArray(result)) {
        throw new OutlookMcpError("The Outlook helper returned an unexpected answer. The operation may not have completed.");
      }
      return result;
    }
    const ErrorClass = ERROR_KINDS[message.kind] ?? OutlookMcpError;
    throw new ErrorClass(String(message.message || "Outlook reported a problem and the operation was not completed."));
  }

  // ---- reading ----

  async listAccounts() {
    const r = await this.call("list_accounts");
    return {
      accounts: list(r.accounts).map((a) => ({
        name: String(a.name ?? ""),
        address: String(a.address ?? ""),
        type: String(a.type ?? "other"),
        is_default: Boolean(a.is_default),
      })),
      data_files: list(r.data_files).map(String),
    };
  }

  async listFolders(account) {
    const r = await this.call("list_folders", { account });
    return list(r.folders).map((f) => ({
      path: String(f.path ?? ""),
      name: String(f.name ?? ""),
      total: Number(f.total ?? 0),
      unread: Number(f.unread ?? 0),
    }));
  }

  async listEmails({ folder, account, since, until, unreadOnly, limit, offset }) {
    return page(await this.call("list_emails", { folder, account, since, until, unread_only: unreadOnly, limit, offset }));
  }

  async searchEmails({ folder, account, terms, since, until, limit }) {
    return page(
      await this.call("search_emails", { folder, account, filter: buildSearchFilter(terms), since, until, limit }),
    );
  }

  async readEmail(emailId) {
    const r = await this.call("read_email", { id: emailId });
    return {
      ...summary(r),
      to: list(r.to).map(String),
      cc: list(r.cc).map(String),
      bcc: list(r.bcc).map(String),
      is_draft: Boolean(r.is_draft),
      body: String(r.body ?? ""),
      attachments: list(r.attachments).map(attachment),
    };
  }

  async listAttachments(emailId) {
    const r = await this.call("list_attachments", { id: emailId });
    return list(r.attachments).map(attachment);
  }

  async saveAttachment(emailId, index, destination) {
    await this.call("save_attachment", { id: emailId, index, dest: destination });
  }

  // ---- drafting ----

  async createDraft({ to, cc, bcc, subject, body, account, attachments }) {
    return String((await this.call("create_draft", { to, cc, bcc, subject, body, account, attachments })).id);
  }

  async createReplyDraft(emailId, body, replyAll) {
    const r = await this.call("create_reply_draft", {
      id: emailId,
      text: body,
      html_block: htmlReplyBlock(body),
      reply_all: replyAll,
    });
    return String(r.id);
  }

  async createForwardDraft(emailId, to, body) {
    const r = await this.call("create_forward_draft", { id: emailId, to, text: body, html_block: htmlReplyBlock(body) });
    return String(r.id);
  }

  async updateDraft(draftId, { to, cc, bcc, subject, body, addAttachments }) {
    await this.call("update_draft", {
      id: draftId,
      set_to: to !== null,
      to: to ?? [],
      set_cc: cc !== null,
      cc: cc ?? [],
      set_bcc: bcc !== null,
      bcc: bcc ?? [],
      set_subject: subject !== null,
      subject: subject ?? "",
      set_body: body !== null,
      body: body ?? "",
      add_attachments: addAttachments ?? [],
    });
  }

  async getDraft(draftId) {
    const r = await this.call("get_draft", { id: draftId });
    return {
      sender_name: String(r.sender_name ?? ""),
      sender_address: String(r.sender_address ?? ""),
      account_type: String(r.account_type ?? "other"),
      to: list(r.to).map(person),
      cc: list(r.cc).map(person),
      bcc: list(r.bcc).map(person),
      subject: String(r.subject ?? ""),
      body: String(r.body ?? ""),
      attachments: list(r.attachments).map(attachment),
    };
  }

  async send(draftId) {
    await this.call("send", { id: draftId });
  }

  async display(draftId) {
    await this.call("display", { id: draftId });
  }

  /** Shows the confirmation window and waits for the human. */
  async confirm(view, timeoutSeconds) {
    const r = await this.call("confirm", { view, timeout: timeoutSeconds }, timeoutSeconds * 1000 + CONFIRM_GRACE_MS);
    return String(r.decision ?? "failed");
  }
}

/**
 * Asks a human whether to send. The model has no way to answer on their behalf.
 * Anything other than "send" means: do not send.
 */
export class DialogConfirmer {
  constructor(backend, lang) {
    this.backend = backend;
    this.lang = lang;
  }

  async ask(snapshot, warnings, timeoutSeconds) {
    try {
      const decision = await this.backend.confirm(buildView(snapshot, warnings, this.lang), timeoutSeconds);
      return ["send", "cancel", "open", "timeout"].includes(decision) ? decision : "failed";
    } catch {
      return "failed"; // never send when the window could not be shown
    }
  }
}
