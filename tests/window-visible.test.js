// Windows only: starts the real helper exactly as the product does (a hidden PowerShell)
// and checks that the confirmation window is actually visible on screen.
//
// A process started hidden has its first window hidden by Windows unless the program
// takes care; this test exists because that made the confirmation window invisible.

import assert from "node:assert/strict";
import { execFile } from "node:child_process";
import path from "node:path";
import { test } from "node:test";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

import { PowerShellBackend, powershellPath } from "../src/powershellBackend.js";

const here = path.dirname(fileURLToPath(import.meta.url));
const helper = path.join(here, "..", "src", "outlook-helper.ps1");
const TITLE = "OCM-WINDOW-TEST";

const view = {
  title: TITLE,
  heading: "Do you want to send this email?",
  warning_lines: [],
  rows: [{ label: "To", value: "a@example.com" }],
  body_label: "Message",
  body: "hello",
  send_text: "Send",
  open_text: "Open in Outlook to check",
  cancel_text: "Don't send",
  countdown_format: "Closes automatically in {0} seconds",
};

/** Number of processes whose visible main window carries the test title. */
async function visibleWindows() {
  const { stdout } = await promisify(execFile)(
    powershellPath(),
    ["-NoProfile", "-Command", `@(Get-Process | Where-Object { $_.MainWindowTitle -eq '${TITLE}' }).Count`],
    { windowsHide: true },
  );
  return Number(stdout.trim());
}

test("the confirmation window is visible when the helper is started hidden", { skip: process.platform !== "win32" }, async () => {
  const backend = new PowerShellBackend({
    args: ["-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", helper],
  });
  try {
    for (const round of [1, 2]) {
      const answer = backend.confirm(view, 8);
      answer.catch(() => {}); // reported through the awaits below
      await new Promise((resolve) => setTimeout(resolve, 4000));
      assert.equal(await visibleWindows(), 1, `window not visible (round ${round})`);
      assert.equal(await answer, "timeout");
      assert.equal(await visibleWindows(), 0, `window still there after closing (round ${round})`);
    }
  } finally {
    backend.close();
  }
});
