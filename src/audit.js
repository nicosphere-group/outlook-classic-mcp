import fs from "node:fs";
import path from "node:path";

/** ISO 8601 with the local offset, e.g. 2026-01-31T09:00:00+09:00. */
export function isoLocal(date = new Date()) {
  const pad = (n) => String(Math.trunc(Math.abs(n))).padStart(2, "0");
  const offset = -date.getTimezoneOffset();
  return (
    `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}` +
    `T${pad(date.getHours())}:${pad(date.getMinutes())}:${pad(date.getSeconds())}` +
    `${offset >= 0 ? "+" : "-"}${pad(offset / 60)}:${pad(offset % 60)}`
  );
}

/** Append-only record of send decisions. Never contains the message body. */
export class AuditLog {
  constructor(directory) {
    this.directory = directory;
    this.filePath = path.join(directory, "audit.log");
  }

  record(outcome, snapshot) {
    const entry = {
      time: isoLocal(),
      event: outcome,
      from: snapshot.sender_address || snapshot.sender_name,
      to: snapshot.to.map((p) => p.address),
      cc: snapshot.cc.map((p) => p.address),
      bcc: snapshot.bcc.map((p) => p.address),
      subject: snapshot.subject,
      attachments: snapshot.attachments.map((a) => a.filename),
    };
    try {
      fs.mkdirSync(this.directory, { recursive: true });
      fs.appendFileSync(this.filePath, JSON.stringify(entry) + "\n", "utf8");
    } catch {
      // A logging failure must not turn a cancelled send into an error, and a completed
      // send cannot be undone, so recording is best effort.
    }
  }
}
