// The most important tests: nothing is sent unless a human pressed Send.

import assert from "node:assert/strict";
import fs from "node:fs";
import { test } from "node:test";

import { InvalidInputError, NotADraftError, NotAllowedError } from "../src/errors.js";
import { makeRig } from "./fakes.js";

const SECRET_BODY = "これは本文です。記録に残ってはいけません。";

const draft = async (rig, extra = {}) =>
  (await rig.service.createDraft({ to: ["tanaka@example.co.jp"], subject: "お見積りの件", body: SECRET_BODY, ...extra }))
    .draft_id;

const lastAudit = (rig) => JSON.parse(fs.readFileSync(rig.auditFile, "utf8").trim().split("\n").pop());

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

test("creating a draft never sends or asks", withRig({}, async (rig) => {
  const result = await rig.service.createDraft({ to: ["a@example.com"], subject: "件名", body: "本文" });
  assert.match(result.status, /NOT sent/);
  assert.equal(rig.backend.sent.length, 0);
  assert.equal(rig.confirmer.asked.length, 0);
}));

test("sends when the user presses send", withRig({ decision: "send" }, async (rig) => {
  const result = await rig.service.sendDraft({ draft_id: await draft(rig) });
  assert.equal(result.sent, true);
  assert.equal(result.outcome, "sent");
  assert.equal(rig.backend.sent.length, 1);
  const { snapshot, timeout } = rig.confirmer.asked[0];
  assert.equal(snapshot.subject, "お見積りの件");
  assert.equal(snapshot.body, SECRET_BODY);
  assert.equal(timeout, rig.config.confirmTimeout);
}));

for (const [decision, outcome] of [
  ["cancel", "declined"],
  ["timeout", "timeout"],
  ["failed", "failed"],
  ["open", "opened"],
  // Anything unexpected is treated as "no": only the exact word "send" sends.
  ["yes", "declined"],
  ["SEND", "declined"],
  [undefined, "declined"],
  [true, "declined"],
]) {
  test(`decision ${JSON.stringify(decision)} does not send`, withRig({}, async (rig) => {
    rig.confirmer.decision = decision;
    const id = await draft(rig);
    const result = await rig.service.sendDraft({ draft_id: id });
    assert.equal(result.sent, false);
    assert.equal(result.outcome, outcome);
    assert.equal(rig.backend.sent.length, 0);
    assert.ok(rig.backend.drafts.has(id)); // the draft is kept
    assert.equal(lastAudit(rig).event, outcome);
  }));
}

test("open shows the draft in Outlook without sending", withRig({ decision: "open" }, async (rig) => {
  const id = await draft(rig);
  await rig.service.sendDraft({ draft_id: id });
  assert.deepEqual(rig.backend.displayed, [id]);
  assert.equal(rig.backend.sent.length, 0);
}));

test("draft changed during confirmation is not sent", withRig({ decision: "send" }, async (rig) => {
  const id = await draft(rig);
  rig.confirmer.during = () => {
    rig.backend.drafts.set(id, { ...rig.backend.drafts.get(id), body: "書き換えられた本文" });
  };
  const result = await rig.service.sendDraft({ draft_id: id });
  assert.equal(result.sent, false);
  assert.equal(result.outcome, "changed");
  assert.equal(rig.backend.sent.length, 0);
}));

test("recipient added during confirmation is not sent", withRig({ decision: "send" }, async (rig) => {
  const id = await draft(rig);
  rig.confirmer.during = () => rig.service.updateDraft({ draft_id: id, bcc: ["evil@example.net"] });
  assert.equal((await rig.service.sendDraft({ draft_id: id })).outcome, "changed");
  assert.equal(rig.backend.sent.length, 0);
}));

test("every send asks again", withRig({ decision: "send" }, async (rig) => {
  await rig.service.sendDraft({ draft_id: await draft(rig) });
  await rig.service.sendDraft({ draft_id: await draft(rig) });
  assert.equal(rig.confirmer.asked.length, 2);
  assert.equal(rig.backend.sent.length, 2);
}));

test("an already sent email cannot be sent or edited", withRig({ decision: "send" }, async (rig) => {
  const id = await draft(rig);
  await rig.service.sendDraft({ draft_id: id });
  await assert.rejects(rig.service.sendDraft({ draft_id: id }), NotADraftError);
  await assert.rejects(rig.service.updateDraft({ draft_id: id, subject: "変更" }), NotADraftError);
  assert.equal(rig.backend.sent.length, 1);
  assert.equal(rig.confirmer.asked.length, 1);
}));

test("a received email cannot be sent", withRig({ decision: "send" }, async (rig) => {
  rig.backend.addInbox("store.mail1");
  await assert.rejects(rig.service.sendDraft({ draft_id: "store.mail1" }), NotADraftError);
  assert.equal(rig.confirmer.asked.length, 0);
}));

test("draft without recipients is rejected before asking", withRig({ decision: "send" }, async (rig) => {
  const id = await draft(rig, { to: [] });
  await assert.rejects(rig.service.sendDraft({ draft_id: id }), InvalidInputError);
  assert.equal(rig.confirmer.asked.length, 0);
  assert.equal(rig.backend.sent.length, 0);
}));

test("send can be turned off", withRig({ decision: "send", allowSend: false }, async (rig) => {
  const id = await draft(rig);
  await assert.rejects(rig.service.sendDraft({ draft_id: id }), NotAllowedError);
  assert.equal(rig.confirmer.asked.length, 0);
  assert.equal(rig.backend.sent.length, 0);
}));

test("audit log records recipients but never the body", withRig({ decision: "send" }, async (rig) => {
  await rig.service.sendDraft({ draft_id: await draft(rig, { cc: ["cc@example.com"] }) });
  const raw = fs.readFileSync(rig.auditFile, "utf8");
  const entry = lastAudit(rig);
  assert.equal(entry.event, "sent");
  assert.deepEqual(entry.to, ["tanaka@example.co.jp"]);
  assert.deepEqual(entry.cc, ["cc@example.com"]);
  assert.equal(entry.subject, "お見積りの件");
  assert.match(entry.time, /^\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d[+-]\d\d:\d\d$/);
  assert.ok(raw.includes("お見積りの件")); // Japanese is stored readably
  assert.ok(!raw.includes(SECRET_BODY));
  assert.ok(!("body" in entry));
}));

for (const accountType of ["pop3", "imap"]) {
  test(`${accountType} reports queued`, withRig({ decision: "send", accountType }, async (rig) => {
    const result = await rig.service.sendDraft({ draft_id: await draft(rig) });
    assert.equal(result.sent, true);
    assert.equal(result.outcome, "queued");
  }));
}

test("reply-all and external recipients are highlighted", withRig({ decision: "cancel" }, async (rig) => {
  rig.backend.addInbox("store.mail1", { cc: ["other@example.org"] });
  const { draft_id } = await rig.service.createReplyDraft({ email_id: "store.mail1", body: "承知しました。", reply_all: true });
  await rig.service.sendDraft({ draft_id });
  const codes = new Set(rig.confirmer.asked[0].warnings.map((w) => w.code));
  assert.ok(codes.has("reply_all"));
  assert.ok(codes.has("external"));
}));
