import assert from "node:assert/strict";
import fs from "node:fs";
import path from "node:path";
import { test } from "node:test";
import { fileURLToPath } from "node:url";

import { NotADraftError, NotFoundError, OutlookMcpError, OutlookUnavailableError } from "../src/errors.js";
import {
  asciiJson,
  buildSearchFilter,
  DialogConfirmer,
  htmlReplyBlock,
  PowerShellBackend,
} from "../src/powershellBackend.js";

const here = path.dirname(fileURLToPath(import.meta.url));
const fakeHelper = path.join(here, "fake-helper.mjs");

function withBackend(fn, env = {}) {
  return async () => {
    const saved = { ...process.env };
    Object.assign(process.env, env);
    const backend = new PowerShellBackend({ command: process.execPath, args: [fakeHelper] });
    try {
      await fn(backend);
    } finally {
      backend.close();
      process.env = saved;
    }
  };
}

test("Japanese survives the round trip and requests are pure ASCII", withBackend(async (backend) => {
  const text = "髙橋様　～よろしく～ ①㈱ 🙂";
  const r = await backend.call("ping", { echo: text });
  assert.equal(r.echo, text);
  assert.ok(/^[\x00-\x7f]*$/.test(asciiJson({ text })));
  assert.deepEqual(JSON.parse(asciiJson({ text })), { text });
}));

test("a single item where a list is expected is repaired", withBackend(async (backend) => {
  const accounts = await backend.listAccounts();
  assert.equal(accounts.accounts.length, 1);
  assert.equal(accounts.accounts[0].name, "自分");
  assert.deepEqual(accounts.data_files, ["個人用フォルダー"]);
  const page = await backend.listEmails({ folder: "inbox", account: null, since: "2026-01-01T00:00:00+09:00", until: null, unreadOnly: false, limit: 5, offset: 0 });
  assert.equal(page.emails.length, 1);
  assert.equal(page.emails[0].subject, "【至急】お見積りの件①");
  assert.equal(page.emails[0].received, "2026-01-01T00:00:00+09:00");
}));

test("helper errors become the right kind with the helper's wording", withBackend(async (backend) => {
  await assert.rejects(backend.readEmail("k.x"), (e) => e instanceof NotFoundError && /moved or deleted/.test(e.message));
  await assert.rejects(backend.getDraft("k.none"), NotADraftError);
  await assert.rejects(backend.send("k.x"), OutlookUnavailableError);
  await assert.rejects(backend.display("k.x"), (e) => e instanceof OutlookMcpError && /unexpected answer/.test(e.message));
}));

test("calls run strictly one at a time, in order", withBackend(async (backend) => {
  const results = await Promise.all([1, 2, 3, 4, 5].map((n) => backend.call("ping", { echo: n })));
  assert.deepEqual(results.map((r) => r.echo), [1, 2, 3, 4, 5]);
}));

test("a helper that never answers times out and is replaced", withBackend(async (backend) => {
  await assert.rejects(backend.call("save_attachment", {}, 300), (e) => e instanceof OutlookUnavailableError && /in time/.test(e.message));
  assert.equal((await backend.call("ping", { echo: "again" })).echo, "again");
}));

test("a helper that crashes is reported and restarted", withBackend(async (backend) => {
  await assert.rejects(backend.listAttachments("k.x"), (e) => e instanceof OutlookUnavailableError && /stopped unexpectedly/.test(e.message));
  assert.equal((await backend.call("ping", { echo: "back" })).echo, "back");
}));

test("PowerShell that cannot be started is a clear error, not a hang", async () => {
  const backend = new PowerShellBackend({ command: path.join(here, "no-such-program"), args: [] });
  await assert.rejects(backend.call("ping"), OutlookUnavailableError);
});

const snapshot = {
  sender_name: "自分", sender_address: "me@example.com", account_type: "exchange",
  to: [{ name: "田中", address: "t@example.com" }], cc: [], bcc: [], subject: "件名", body: "本文", attachments: [],
};

test("confirmer passes the human's answer through", withBackend(async (backend) => {
  assert.equal(await new DialogConfirmer(backend, "ja").ask(snapshot, [], 5), "send");
}, { FAKE_DECISION: "send" }));

for (const answer of ["yes", "SEND", "", "ok"]) {
  test(`confirmer turns the unknown answer ${JSON.stringify(answer)} into "failed"`, withBackend(async (backend) => {
    assert.equal(await new DialogConfirmer(backend, "ja").ask(snapshot, [], 5), "failed");
  }, { FAKE_DECISION: answer }));
}

test("confirmer says failed when the helper cannot be reached", async () => {
  const backend = new PowerShellBackend({ command: path.join(here, "no-such-program"), args: [] });
  assert.equal(await new DialogConfirmer(backend, "ja").ask(snapshot, [], 5), "failed");
});

test("search filter escapes quotes and has no dates", () => {
  assert.equal(buildSearchFilter({ query: null, subject: null, sender: null, body: null }), null);
  const filter = buildSearchFilter({ query: null, subject: "O'Brien の件", sender: "田中", body: null });
  assert.ok(filter.startsWith("@SQL="));
  assert.ok(filter.includes("like '%O''Brien の件%'"));
  assert.ok(filter.includes(`"urn:schemas:httpmail:fromname" like '%田中%'`));
  assert.ok(filter.includes(" AND "));
  assert.ok(!/date/i.test(filter));
});

test("reply text becomes escaped HTML", () => {
  const block = htmlReplyBlock("承知しました。\n<script>alert(1)</script> & \"x\"");
  assert.equal(block, "<div>承知しました。<br>&lt;script&gt;alert(1)&lt;/script&gt; &amp; &quot;x&quot;</div><br>");
  assert.equal(htmlReplyBlock(""), "");
});

test("the PowerShell helper is plain ASCII (5.1 misreads anything else) and avoids 7-only syntax", () => {
  const source = fs.readFileSync(path.join(here, "..", "src", "outlook-helper.ps1"), "utf8");
  const bad = [...source].filter((c) => c.charCodeAt(0) > 0x7e || (c.charCodeAt(0) < 0x20 && !"\r\n\t".includes(c)));
  assert.deepEqual(bad, []);
  assert.ok(!/\?\?|\?\.|\s\?\s[^\n]*\s:\s/.test(source.replace(/#.*$/gm, "")), "PowerShell 7 operators found");
});
