// In-memory stand-ins for Outlook and for the human at the confirmation window.

import fs from "node:fs";
import os from "node:os";
import path from "node:path";

import { AuditLog } from "../src/audit.js";
import { NotADraftError, NotFoundError } from "../src/errors.js";
import { MailService } from "../src/service.js";

const people = (addresses) => addresses.map((a) => ({ name: a, address: a }));

export class FakeBackend {
  constructor(accountType = "exchange") {
    this.accountType = accountType;
    this.drafts = new Map();
    this.sent = [];
    this.sentIds = new Set();
    this.displayed = [];
    this.inbox = new Map();
    this.next = 0;
    this.lastList = null;
    this.lastSearch = null;
  }

  addInbox(id, { subject = "件名", body = "本文", cc = [], attachments = [] } = {}) {
    this.inbox.set(id, {
      id,
      subject,
      from: { name: "田中 太郎", address: "tanaka@example.co.jp" },
      received: "2026-01-10T09:00:00+09:00",
      unread: true,
      attachment_count: attachments.length,
      to: [],
      cc,
      bcc: [],
      is_draft: false,
      body,
      attachments,
    });
  }

  #newDraft(fields) {
    const id = `store.draft${++this.next}`;
    this.drafts.set(id, {
      sender_name: "自分",
      sender_address: "me@example.com",
      account_type: this.accountType,
      to: [],
      cc: [],
      bcc: [],
      subject: "",
      body: "",
      attachments: [],
      ...fields,
    });
    return id;
  }

  #draft(id) {
    if (this.sentIds.has(id) || this.inbox.has(id)) throw new NotADraftError("not a draft");
    if (!this.drafts.has(id)) throw new NotFoundError("not found");
    return this.drafts.get(id);
  }

  #summaries() {
    return [...this.inbox.values()].map(({ id, subject, from, received, unread, attachment_count }) => ({
      id, subject, from, received, unread, attachment_count,
    }));
  }

  async listAccounts() {
    return {
      accounts: [{ name: "自分", address: "me@example.com", type: this.accountType, is_default: true }],
      data_files: ["me@example.com"],
    };
  }

  async listFolders() {
    return [{ path: "me@example.com\\受信トレイ", name: "受信トレイ", total: this.inbox.size, unread: this.inbox.size }];
  }

  async listEmails(args) {
    this.lastList = args;
    return { emails: this.#summaries().slice(args.offset, args.offset + args.limit), has_more: false, timed_out: false };
  }

  async searchEmails(args) {
    this.lastSearch = args;
    return { emails: this.#summaries().slice(0, args.limit), has_more: false, timed_out: false };
  }

  async readEmail(id) {
    if (!this.inbox.has(id)) throw new NotFoundError("not found");
    return structuredClone(this.inbox.get(id));
  }

  async listAttachments(id) {
    return (await this.readEmail(id)).attachments;
  }

  async saveAttachment(_id, _index, destination) {
    fs.writeFileSync(destination, "data");
  }

  async createDraft({ to, cc, bcc, subject, body, attachments }) {
    return this.#newDraft({
      to: people(to),
      cc: people(cc),
      bcc: people(bcc),
      subject,
      body,
      attachments: attachments.map((p, i) => ({ index: i + 1, filename: path.basename(p), size: fs.statSync(p).size })),
    });
  }

  async createReplyDraft(id, body, replyAll) {
    const original = await this.readEmail(id);
    return this.#newDraft({
      to: [original.from],
      cc: replyAll ? people(original.cc) : [],
      subject: "RE: " + original.subject,
      body: body + "\n> " + original.body,
    });
  }

  async createForwardDraft(id, to, body) {
    return this.#newDraft({ to: people(to), subject: "FW: " + (await this.readEmail(id)).subject, body });
  }

  async updateDraft(id, { to, cc, bcc, subject, body }) {
    const d = this.#draft(id);
    this.drafts.set(id, {
      ...d,
      to: to === null ? d.to : people(to),
      cc: cc === null ? d.cc : people(cc),
      bcc: bcc === null ? d.bcc : people(bcc),
      subject: subject ?? d.subject,
      body: body ?? d.body,
    });
  }

  async getDraft(id) {
    return structuredClone(this.#draft(id));
  }

  async send(id) {
    this.sent.push(this.#draft(id));
    this.sentIds.add(id);
    this.drafts.delete(id);
  }

  async display(id) {
    this.#draft(id);
    this.displayed.push(id);
  }
}

/** Stands in for the human at the confirmation window. */
export class FakeConfirmer {
  constructor(decision) {
    this.decision = decision;
    this.during = null;
    this.asked = [];
  }

  async ask(snapshot, warnings, timeout) {
    this.asked.push({ snapshot, warnings, timeout });
    if (this.during) await this.during();
    return this.decision;
  }
}

/** A service wired to fakes, with its own temporary folders. */
export function makeRig({ decision = "send", accountType = "exchange", allowSend = true } = {}) {
  const root = fs.mkdtempSync(path.join(os.tmpdir(), "ocm-tests-"));
  const config = {
    allowSend,
    diagnostics: false,
    confirmTimeout: 30,
    saveDir: path.join(root, "saved"),
    attachRoots: [path.join(root, "Documents")],
    auditDir: path.join(root, "audit"),
    lang: "ja",
  };
  fs.mkdirSync(config.attachRoots[0], { recursive: true });
  const backend = new FakeBackend(accountType);
  const confirmer = new FakeConfirmer(decision);
  const service = new MailService(backend, confirmer, new AuditLog(config.auditDir), config);
  return {
    root,
    config,
    backend,
    confirmer,
    service,
    attachDir: config.attachRoots[0],
    auditFile: path.join(config.auditDir, "audit.log"),
    cleanup: () => fs.rmSync(root, { recursive: true, force: true }),
  };
}
