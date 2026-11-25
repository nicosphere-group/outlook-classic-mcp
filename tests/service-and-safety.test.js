import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { test } from "node:test";

import { loadConfig } from "../src/config.js";
import { buildView, formatSize, pickLanguage, TEXT } from "../src/dialogText.js";
import { InvalidInputError, NotAllowedError, NotFoundError } from "../src/errors.js";
import { computeWarnings, fingerprint, isWithin, safeSavePath, sanitizeFilename } from "../src/safety.js";
import { parseWhen, UNTRUSTED_NOTICE } from "../src/service.js";
import { makeRig } from "./fakes.js";

function withRig(options, fn) {
  return async () => {
    const rig = makeRig(options);
    try {
      await fn(rig);
    } finally {
      rig.cleanup();
    }
  };
}

// ---- file names and paths ----

for (const [raw, expected] of [
  ["見積書.pdf", "見積書.pdf"],
  ["..\\..\\Windows\\evil.exe", "evil.exe"],
  ["../../etc/passwd", "passwd"],
  ["C:\\Users\\x\\a.txt", "a.txt"],
  ['a<b>:"c|d?.txt', "a_b___c_d_.txt"],
  ["CON.txt", "_CON.txt"],
  ["...", "attachment"],
  ["", "attachment"],
  [null, "attachment"],
]) {
  test(`sanitizeFilename(${JSON.stringify(raw)})`, () => assert.equal(sanitizeFilename(raw), expected));
}

test("long names keep their extension", () => {
  const name = sanitizeFilename("あ".repeat(300) + ".xlsx");
  assert.ok(name.length <= 150);
  assert.ok(name.endsWith(".xlsx"));
});

test("saved attachments stay in the folder and never overwrite", withRig({}, async (rig) => {
  const dir = rig.config.saveDir;
  const first = safeSavePath(dir, "..\\..\\report.pdf");
  fs.writeFileSync(first, "x");
  const second = safeSavePath(dir, "report.pdf");
  assert.equal(first, path.resolve(dir, "report.pdf"));
  assert.equal(second, path.resolve(dir, "report (1).pdf"));
  assert.ok(isWithin(first, [dir]) && isWithin(second, [dir]));
}));

test("save_attachment uses a safe path", withRig({}, async (rig) => {
  rig.backend.addInbox("store.mail1", { attachments: [{ index: 1, filename: "..\\..\\startup\\run.bat", size: 10 }] });
  const saved = (await rig.service.saveAttachment({ email_id: "store.mail1", index: 1 })).saved_to;
  assert.equal(path.dirname(saved), path.resolve(rig.config.saveDir));
  assert.equal(path.basename(saved), "run.bat");
  assert.ok(fs.existsSync(saved));
  await assert.rejects(rig.service.saveAttachment({ email_id: "store.mail1", index: 2 }), InvalidInputError);
}));

test("attachments only from allowed folders", withRig({}, async (rig) => {
  const allowed = path.join(rig.attachDir, "資料.txt");
  fs.writeFileSync(allowed, "ok");
  const outside = path.join(rig.root, "secret.txt");
  fs.writeFileSync(outside, "secret");
  const base = { to: ["a@example.com"], subject: "件名", body: "本文" };

  const result = await rig.service.createDraft({ ...base, attachments: [allowed] });
  assert.equal(result.draft.attachments[0].filename, "資料.txt");

  await assert.rejects(rig.service.createDraft({ ...base, attachments: [outside] }), NotAllowedError);
  const sneaky = path.join(rig.attachDir, "..", "secret.txt");
  await assert.rejects(rig.service.createDraft({ ...base, attachments: [sneaky] }), NotAllowedError);
  await assert.rejects(rig.service.createDraft({ ...base, attachments: ["relative.txt"] }), NotAllowedError);
  await assert.rejects(rig.service.createDraft({ ...base, attachments: [path.join(rig.attachDir, "none")] }), NotFoundError);
  await assert.rejects(rig.service.createDraft({ ...base, attachments: [rig.attachDir] }), NotFoundError); // a folder
  assert.equal(rig.backend.drafts.size, 1);
}));

test("a link inside an allowed folder cannot reach outside", withRig({}, async (rig) => {
  const outside = path.join(rig.root, "secret.txt");
  fs.writeFileSync(outside, "secret");
  const link = path.join(rig.attachDir, "innocent.txt");
  try {
    fs.symlinkSync(outside, link);
  } catch {
    return; // creating links needs a privilege on some Windows setups
  }
  await assert.rejects(
    rig.service.createDraft({ to: ["a@example.com"], subject: "件名", body: "本文", attachments: [link] }),
    NotAllowedError,
  );
}));

// ---- input handling ----

for (const bad of ["田中", "a@example.com; b@example.com", "a@example.com,b@example.com"]) {
  test(`address ${JSON.stringify(bad)} is rejected`, withRig({}, async (rig) => {
    await assert.rejects(rig.service.createDraft({ to: [bad], subject: "件名", body: "本文" }), InvalidInputError);
    assert.equal(rig.backend.drafts.size, 0);
  }));
}

test("parseWhen", () => {
  assert.equal(parseWhen(null), null);
  assert.equal(parseWhen("  "), null);
  assert.match(parseWhen("2026-01-31"), /^2026-01-31T00:00:00[+-]\d\d:\d\d$/);
  assert.match(parseWhen("2026-01-31", true), /^2026-01-31T23:59:59[+-]\d\d:\d\d$/);
  assert.match(parseWhen("2026-01-31T09:30"), /^2026-01-31T09:30:00[+-]\d\d:\d\d$/);
  assert.equal(new Date(parseWhen("2026-01-31T09:30:00Z")).toISOString(), "2026-01-31T09:30:00.000Z");
  for (const bad of ["昨日", "1/2/2026", "2026-02-31", "2026-13-01", "20260131"]) {
    assert.throws(() => parseWhen(bad), InvalidInputError, bad);
  }
});

test("list limits are clamped", withRig({}, async (rig) => {
  await rig.service.listEmails({ limit: 100000, offset: -5, until: "2026-01-31" });
  assert.equal(rig.backend.lastList.limit, 100);
  assert.equal(rig.backend.lastList.offset, 0);
  assert.match(rig.backend.lastList.until, /T23:59:59/);
  assert.equal(rig.backend.lastList.folder, "inbox");
}));

test("search needs something to search for", withRig({}, async (rig) => {
  await assert.rejects(rig.service.searchEmails({}), InvalidInputError);
  await assert.rejects(rig.service.searchEmails({ query: "   " }), InvalidInputError);
  assert.ok(Array.isArray((await rig.service.searchEmails({ sender: "田中" })).emails));
  assert.deepEqual(rig.backend.lastSearch.terms, { query: null, subject: null, sender: "田中", body: null });
}));

// ---- reading ----

test("read_email truncates and marks content untrusted", withRig({}, async (rig) => {
  rig.backend.addInbox("store.mail1", { subject: "【重要】ご確認ください", body: "あ".repeat(2000) });
  const result = await rig.service.readEmail({ email_id: "store.mail1", max_chars: 500 });
  assert.equal(result.notice, UNTRUSTED_NOTICE);
  assert.equal(result.subject, "【重要】ご確認ください");
  assert.equal(result.body.length, 500);
  assert.equal(result.body_truncated, true);
  assert.equal(result.body_total_chars, 2000);
  assert.equal((await rig.service.listEmails({})).notice, UNTRUSTED_NOTICE);
}));

// ---- warnings and fingerprint ----

const snapshot = (fields = {}) => ({
  sender_name: "自分",
  sender_address: "me@example.com",
  account_type: "exchange",
  to: [],
  cc: [],
  bcc: [],
  subject: "",
  body: "",
  attachments: [],
  ...fields,
});

test("no warnings for a plain internal email", () => {
  assert.deepEqual(computeWarnings(snapshot({ to: [{ name: "同僚", address: "colleague@example.com" }] })), []);
});

test("warnings", () => {
  const many = Array.from({ length: 5 }, (_, i) => ({ name: "", address: `user${i}@example.com` }));
  const s = snapshot({
    to: many,
    bcc: [{ name: "", address: "x@Other.example" }],
    attachments: [{ index: 1, filename: "a.pdf", size: 100 }],
  });
  const warnings = Object.fromEntries(computeWarnings(s, true).map((w) => [w.code, w]));
  assert.equal(warnings.many_recipients.count, 6);
  assert.deepEqual(warnings.external.domains, ["other.example"]);
  assert.equal(warnings.attachments.count, 1);
  assert.ok(warnings.reply_all);
});

test("fingerprint changes with any visible field", () => {
  const a = [{ name: "A", address: "a@example.com" }];
  const base = snapshot({ to: a, subject: "s", body: "b" });
  const variants = [
    snapshot({ to: [{ name: "A", address: "b@example.com" }], subject: "s", body: "b" }),
    snapshot({ to: a, subject: "s2", body: "b" }),
    snapshot({ to: a, subject: "s", body: "b2" }),
    snapshot({ to: a, bcc: [{ name: "", address: "c@x.com" }], subject: "s", body: "b" }),
    snapshot({ to: [], cc: a, subject: "s", body: "b" }), // moved from To to CC
    snapshot({ to: a, subject: "s", body: "b", attachments: [{ index: 1, filename: "f", size: 1 }] }),
    snapshot({ to: a, subject: "s", body: "b", sender_address: "other@example.com" }),
  ];
  assert.equal(fingerprint(base), fingerprint(structuredClone(base)));
  assert.equal(new Set([base, ...variants].map(fingerprint)).size, variants.length + 1);
});

// ---- confirmation window wording ----

const dialogSnapshot = () =>
  snapshot({
    to: [{ name: "田中 太郎", address: "tanaka@example.co.jp" }],
    subject: "Re: お見積りの件",
    body: "田中様\nお世話になっております。",
    attachments: [{ index: 1, filename: "見積書.pdf", size: 123_000 }],
  });

test("window shows everything that will be sent", () => {
  const view = buildView(dialogSnapshot(), [{ code: "attachments", count: 1 }], "ja");
  const rows = Object.fromEntries(view.rows.map((r) => [r.label, r.value]));
  assert.equal(rows["差出人"], "自分 <me@example.com>");
  assert.equal(rows["宛先"], "田中 太郎 <tanaka@example.co.jp>");
  assert.equal(rows["CC"], "なし");
  assert.equal(rows["BCC"], "なし");
  assert.equal(rows["件名"], "Re: お見積りの件");
  assert.equal(rows["添付"], "見積書.pdf（120 KB）");
  assert.deepEqual(view.warning_lines, ["添付ファイルが 1 件あります。"]);
  assert.ok(view.body.startsWith("田中様"));
  assert.equal(view.send_text, "送信する");
  assert.equal(view.open_text, "Outlook で開いて確認");
  assert.equal(view.cancel_text, "送信しない");
  assert.ok(view.countdown_format.includes("{0}"));
});

test("warning lines are plain sentences", () => {
  const warnings = [
    { code: "many_recipients", count: 12 },
    { code: "reply_all" },
    { code: "external", domains: ["example.org", "example.net"] },
  ];
  const ja = buildView(dialogSnapshot(), warnings, "ja").warning_lines;
  assert.equal(ja[0], "宛先が 12 件あります。");
  assert.match(ja[1], /全員に返信/);
  assert.equal(ja[2], "別の組織のアドレスが含まれています：example.org, example.net");
  assert.equal(buildView(dialogSnapshot(), warnings, "en").warning_lines[0], "There are 12 recipients.");
});

test("languages have the same keys; empty subject is labelled; language choice", () => {
  assert.deepEqual(Object.keys(TEXT.ja).sort(), Object.keys(TEXT.en).sort());
  const view = buildView({ ...dialogSnapshot(), subject: "" }, [], "en");
  assert.ok(view.rows.some((r) => r.label === "Subject" && r.value === "(no subject)"));
  assert.equal(pickLanguage("auto", "ja-JP"), "ja");
  assert.equal(pickLanguage("auto", "en-US"), "en");
  assert.equal(pickLanguage("en", "ja-JP"), "en");
  assert.equal(formatSize(10), "1 KB");
  assert.equal(formatSize(5 * 1024 * 1024), "5.0 MB");
});

// ---- config ----

test("config defaults are safe and env is respected", () => {
  const none = loadConfig({});
  assert.equal(none.allowSend, true);
  assert.equal(none.diagnostics, false);
  assert.equal(none.confirmTimeout, 180);
  assert.ok(none.attachRoots.length >= 3);

  const custom = loadConfig({
    OUTLOOK_MCP_ALLOW_SEND: "0",
    OUTLOOK_MCP_CONFIRM_TIMEOUT: "5",
    OUTLOOK_MCP_SAVE_DIR: "/somewhere",
    OUTLOOK_MCP_LANG: "EN",
    OUTLOOK_MCP_DIAGNOSTICS: "true",
    OUTLOOK_MCP_ATTACH_DIRS: "C:\\a; C:\\b",
  });
  assert.equal(custom.allowSend, false);
  assert.equal(custom.confirmTimeout, 15); // clamped to the minimum
  assert.equal(custom.saveDir, "/somewhere");
  assert.equal(custom.lang, "en");
  assert.equal(custom.diagnostics, true);
  assert.deepEqual(custom.attachRoots, ["C:\\a", "C:\\b"]);
});

for (const [value, expected] of [
  ["false", false],
  ["FALSE", false],
  ["0", false],
  ["true", true],
  ["1", true],
  ["", true],
  ["${user_config.allow_send}", true], // a placeholder the host did not fill in
  ["nonsense", true],
]) {
  test(`OUTLOOK_MCP_ALLOW_SEND=${JSON.stringify(value)} -> send ${expected ? "on" : "off"}`, () => {
    assert.equal(loadConfig({ OUTLOOK_MCP_ALLOW_SEND: value }).allowSend, expected);
  });
}

test("unfilled placeholders and bad numbers fall back to defaults; OneDrive folders are allowed", () => {
  const config = loadConfig({ OUTLOOK_MCP_SAVE_DIR: "${user_config.save_dir}", OUTLOOK_MCP_CONFIRM_TIMEOUT: "abc" });
  assert.equal(config.saveDir, loadConfig({}).saveDir);
  assert.equal(config.confirmTimeout, 180);
  const od = loadConfig({ USERPROFILE: "/home/u", OneDrive: "/home/u/OneDrive" });
  assert.ok(od.attachRoots.includes(path.join("/home/u/OneDrive", "ドキュメント")));
  assert.ok(od.attachRoots.includes(path.join("/home/u", "Documents")));
});
