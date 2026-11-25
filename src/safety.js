// File-path limits and pre-send warnings. No Outlook access here.

import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";

import { NotAllowedError, NotFoundError } from "./errors.js";

export const MANY_RECIPIENTS = 5;
const MAX_NAME_LENGTH = 150;
const BAD_CHARS = /[<>:"/\\|?*\x00-\x1f]/g;
const RESERVED = new Set([
  "CON", "PRN", "AUX", "NUL",
  ...Array.from({ length: 9 }, (_, i) => `COM${i + 1}`),
  ...Array.from({ length: 9 }, (_, i) => `LPT${i + 1}`),
]);

/** Turns an attacker-chosen attachment name into a harmless file name. */
export function sanitizeFilename(name) {
  let clean = String(name ?? "").replaceAll("\\", "/").split("/").pop();
  clean = clean.replace(BAD_CHARS, "_").replace(/^[ .]+|[ .]+$/g, "");
  if (clean === "") clean = "attachment";
  if (RESERVED.has(clean.split(".")[0].toUpperCase())) clean = "_" + clean;
  if (clean.length > MAX_NAME_LENGTH) {
    const dot = clean.lastIndexOf(".");
    const ext = dot >= 0 ? clean.slice(dot + 1) : "";
    clean =
      dot >= 0 && ext.length <= 10
        ? clean.slice(0, MAX_NAME_LENGTH - ext.length - 1) + "." + ext
        : clean.slice(0, MAX_NAME_LENGTH);
  }
  return clean;
}

/** Follows links and junctions when the path exists, so a link cannot reach outside. */
function realOrResolved(p) {
  const resolved = path.resolve(p);
  try {
    return fs.realpathSync.native(resolved);
  } catch {
    // Not there yet: resolve the nearest existing parent instead.
    const parent = path.dirname(resolved);
    return parent === resolved ? resolved : path.join(realOrResolved(parent), path.basename(resolved));
  }
}

export function isWithin(target, roots) {
  const full = realOrResolved(target).toLowerCase();
  return roots.some((root) => {
    const r = realOrResolved(root).toLowerCase().replace(/[\\/]+$/, "");
    return full === r || full.startsWith(r + path.sep);
  });
}

/** A path inside saveDir that does not exist yet. Never overwrites. */
export function safeSavePath(saveDir, filename) {
  fs.mkdirSync(saveDir, { recursive: true });
  const clean = sanitizeFilename(filename);
  const dot = clean.lastIndexOf(".");
  const stem = dot > 0 ? clean.slice(0, dot) : clean;
  const ext = dot > 0 ? clean.slice(dot) : "";
  let candidate = path.join(saveDir, clean);
  for (let n = 1; fs.existsSync(candidate); n++) {
    candidate = path.join(saveDir, `${stem} (${n})${ext}`);
  }
  if (!isWithin(candidate, [saveDir])) {
    throw new NotAllowedError("The attachment could not be saved inside the allowed folder.");
  }
  return path.resolve(candidate);
}

/** Files may be attached only from the allowed folders (limits data exfiltration). */
export function checkAttachSources(paths, roots) {
  const result = [];
  for (const raw of paths ?? []) {
    if (typeof raw !== "string" || !path.isAbsolute(raw)) {
      throw new NotAllowedError(`Give the full path of the file to attach: ${raw}`);
    }
    if (!isWithin(raw, roots)) {
      throw new NotAllowedError(
        `Files can be attached only from these folders: ${roots.join(", ")}. Not allowed: ${raw}`,
      );
    }
    const full = realOrResolved(raw);
    if (!fs.existsSync(full) || !fs.statSync(full).isFile()) {
      throw new NotFoundError(`The file to attach was not found: ${raw}`);
    }
    result.push(full);
  }
  return result;
}

function domain(address) {
  const at = String(address ?? "").lastIndexOf("@");
  return at < 0 ? "" : address.slice(at + 1).trim().toLowerCase();
}

export const recipientsOf = (snapshot) => [...snapshot.to, ...snapshot.cc, ...snapshot.bcc];

/** Things the confirmation window highlights. Codes are localized by the window. */
export function computeWarnings(snapshot, replyAll = false) {
  const warnings = [];
  const recipients = recipientsOf(snapshot);
  if (recipients.length >= MANY_RECIPIENTS) {
    warnings.push({ code: "many_recipients", count: recipients.length });
  }
  if (replyAll) warnings.push({ code: "reply_all" });
  const own = domain(snapshot.sender_address);
  if (own) {
    const external = [...new Set(recipients.map((r) => domain(r.address)).filter((d) => d && d !== own))].sort();
    if (external.length > 0) warnings.push({ code: "external", domains: external });
  }
  if (snapshot.attachments.length > 0) {
    warnings.push({ code: "attachments", count: snapshot.attachments.length });
  }
  return warnings;
}

/** Stable hash of what would be sent; detects edits made during confirmation. */
export function fingerprint(snapshot) {
  const people = (list) => list.map((p) => [p.name, p.address]);
  const canonical = JSON.stringify([
    snapshot.sender_name,
    snapshot.sender_address,
    snapshot.account_type,
    people(snapshot.to),
    people(snapshot.cc),
    people(snapshot.bcc),
    snapshot.subject,
    snapshot.body,
    snapshot.attachments.map((a) => [a.index, a.filename, a.size]),
  ]);
  return crypto.createHash("sha256").update(canonical, "utf8").digest("hex");
}
