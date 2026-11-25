// Builds extension/server: the whole Node side bundled into one file, plus the helper.
// The result runs on the Node that Claude Desktop ships, so nothing else is installed.

import { build } from "esbuild";
import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";

const root = path.dirname(path.dirname(fileURLToPath(import.meta.url)));
const out = path.join(root, "extension", "server");

fs.rmSync(out, { recursive: true, force: true });
fs.mkdirSync(out, { recursive: true });
fs.mkdirSync(path.join(root, "dist"), { recursive: true });

await build({
  entryPoints: [path.join(root, "src", "index.js")],
  outfile: path.join(out, "index.js"),
  bundle: true,
  platform: "node",
  target: "node18",
  format: "cjs",
  legalComments: "none",
  logLevel: "warning",
});

// Windows PowerShell 5.1 reads a script without a byte-order mark as the local code page,
// so the helper must stay plain ASCII.
const helper = fs.readFileSync(path.join(root, "src", "outlook-helper.ps1"), "utf8");
if (/[^\x00-\x7e]/.test(helper.replace(/[\r\n\t]/g, ""))) {
  throw new Error("outlook-helper.ps1 must contain only ASCII characters");
}
fs.writeFileSync(path.join(out, "outlook-helper.ps1"), helper.replace(/\r?\n/g, "\r\n"), "ascii");

for (const file of fs.readdirSync(out)) {
  console.log(`${file}: ${(fs.statSync(path.join(out, file)).size / 1024).toFixed(0)} KB`);
}
