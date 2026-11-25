// Windows only: starts the real helper exactly as the product does (a hidden PowerShell)
// and checks that the confirmation window is actually visible on screen.
//
// A process started hidden has its first window hidden by Windows unless the program
// takes care; this test exists because that made the confirmation window invisible.

import assert from "node:assert/strict";
import { execFile, spawn } from "node:child_process";
import path from "node:path";
import { test } from "node:test";
import { fileURLToPath } from "node:url";
import { promisify } from "node:util";

import { PowerShellBackend, powershellPath } from "../src/powershellBackend.js";

const here = path.dirname(fileURLToPath(import.meta.url));
const helper = path.join(here, "..", "src", "outlook-helper.ps1");
const args = ["-NoLogo", "-NoProfile", "-NonInteractive", "-ExecutionPolicy", "Bypass", "-File", helper];

const view = (title) => ({
  title,
  heading: "Do you want to send this email?",
  warning_lines: [],
  rows: [{ label: "To", value: "a@example.com" }],
  body_label: "Message",
  body: "hello",
  send_text: "Send",
  open_text: "Open in Outlook to check",
  cancel_text: "Don't send",
  countdown_format: "Closes automatically in {0} seconds",
});

/** Asks Windows directly: does a window with this title exist, and is it visible? */
async function windowState(title) {
  const script =
    `Add-Type -Namespace W -Name U -MemberDefinition '[DllImport("user32.dll", CharSet=CharSet.Unicode)] public static extern IntPtr FindWindow(string c, string t); [DllImport("user32.dll")] public static extern bool IsWindowVisible(IntPtr h); [DllImport("user32.dll")] public static extern bool IsIconic(IntPtr h);'; ` +
    `$h = [W.U]::FindWindow([NullString]::Value, '${title}'); ` +
    `if ($h -eq [IntPtr]::Zero) { 'none' } elseif ([W.U]::IsIconic($h)) { 'minimized' } elseif ([W.U]::IsWindowVisible($h)) { 'visible' } else { 'hidden' }`;
  const { stdout } = await promisify(execFile)(powershellPath(), ["-NoProfile", "-Command", script], { windowsHide: true });
  return stdout.trim();
}

async function observe(backend, title) {
  const answer = backend.confirm(view(title), 8);
  answer.catch(() => {});
  await new Promise((resolve) => setTimeout(resolve, 4000));
  const during = await windowState(title);
  const decision = await answer;
  const after = await windowState(title);
  return { during, decision, after };
}

const skip = process.platform !== "win32";

test("control: the check itself can see a window when the helper is started normally", { skip }, async () => {
  const backend = new PowerShellBackend({ args, spawnImpl: (c, a, o) => spawn(c, a, { ...o, windowsHide: false }) });
  try {
    const r = await observe(backend, "OCM-WINDOW-CONTROL");
    console.log("control:", JSON.stringify(r));
    assert.deepEqual(r, { during: "visible", decision: "timeout", after: "none" });
  } finally {
    backend.close();
  }
});

test("the confirmation window is visible when the helper is started hidden", { skip }, async () => {
  const backend = new PowerShellBackend({ args });
  try {
    for (const round of [1, 2]) {
      const r = await observe(backend, `OCM-WINDOW-HIDDEN-${round}`);
      console.log(`hidden start, round ${round}:`, JSON.stringify(r));
      assert.deepEqual(r, { during: "visible", decision: "timeout", after: "none" });
    }
  } finally {
    backend.close();
  }
});
